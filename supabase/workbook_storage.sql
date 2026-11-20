begin;

create schema requests_private;
revoke all on schema requests_private from public, anon, authenticated;

create table requests_private.workbook (
    singleton boolean primary key default true check (singleton),
    token_hash text check (token_hash ~ '^[0-9a-f]{64}$'),
    revision uuid
);
insert into requests_private.workbook (singleton) values (true);

create table requests_private.write_receipts (
    operation_id uuid primary key,
    fingerprint text not null check (fingerprint ~ '^[0-9a-f]{64}$'),
    revision uuid not null unique,
    write_at text not null check (write_at ~ '^[AD][0-9]{1,4}$'),
    ticket_id integer,
    expires_at timestamptz not null default now() + interval '7 days'
);

create table requests_private.status_plans (
    session_id uuid primary key,
    plan_id uuid not null,
    data jsonb not null,
    expires_at timestamptz not null
);

create table requests_private.chat_budget (
    singleton boolean primary key default true check (singleton),
    window_start timestamptz not null,
    used integer not null
);

create function requests_private.has_workbook_write_token()
returns boolean
language sql stable security definer set search_path = ''
as $$
    select coalesce((
        select token_hash = encode(extensions.digest(
            coalesce(nullif(current_setting('request.headers', true), ''), '{}')::jsonb ->> 'x-workbook-write',
            'sha256'
        ), 'hex')
        from requests_private.workbook
    ), false);
$$;
revoke all on function requests_private.has_workbook_write_token() from public;
grant usage on schema requests_private to anon;
grant execute on function requests_private.has_workbook_write_token() to anon;

create policy requests_workbook_versions_read on storage.objects for select to anon
using (
    bucket_id = 'workbooks' and name ~ '^versions/[0-9a-f-]{36}\.xlsx$'
    and (select requests_private.has_workbook_write_token())
);

create policy requests_workbook_versions_upload on storage.objects for insert to anon
with check (
    bucket_id = 'workbooks' and name ~ '^versions/[0-9a-f-]{36}\.xlsx$'
    and (select requests_private.has_workbook_write_token())
);

create function public.requests_workbook_state(
    operation uuid default null,
    request_hash text default null,
    expected uuid default null,
    candidate uuid default null,
    write_at text default null,
    ticket_id integer default null
)
returns jsonb
language plpgsql security definer set search_path = ''
as $$
declare
    latest uuid;
    existing requests_private.write_receipts%rowtype;
begin
    if not requests_private.has_workbook_write_token() then
        raise sqlstate '42501' using message = 'Workbook access denied';
    end if;

    if operation is null then
        select revision into latest from requests_private.workbook;
        return jsonb_build_object('revision', latest);
    end if;
    if request_hash is null or request_hash !~ '^[0-9a-f]{64}$' then
        raise sqlstate '22023' using message = 'Invalid request fingerprint';
    end if;

    -- Serialize publication of versions and their receipts across Vercel instances.
    select revision into latest from requests_private.workbook for update;
    select * into existing from requests_private.write_receipts where operation_id = operation;
    if found then
        if existing.fingerprint <> request_hash then
            return jsonb_build_object('conflict', true);
        end if;
        return jsonb_build_object('revision', existing.revision, 'replayed', true,
                                  'at', existing.write_at, 'ticketId', existing.ticket_id);
    end if;
    if candidate is null then
        return jsonb_build_object('pending', true);
    end if;
    if latest is distinct from expected then
        return jsonb_build_object('conflict', true);
    end if;
    if not exists (
        select 1 from storage.objects
        where bucket_id = 'workbooks' and name = 'versions/' || candidate::text || '.xlsx'
    ) then
        raise sqlstate '22023' using message = 'Workbook version was not uploaded';
    end if;

    update requests_private.workbook set revision = candidate;
    insert into requests_private.write_receipts (operation_id, fingerprint, revision, write_at, ticket_id)
    values (operation, request_hash, candidate, write_at, ticket_id);
    delete from requests_private.write_receipts where expires_at < now();
    return jsonb_build_object('revision', candidate, 'replayed', false, 'at', write_at, 'ticketId', ticket_id);
end;
$$;
revoke all on function public.requests_workbook_state(uuid, text, uuid, uuid, text, integer) from public, authenticated;
grant execute on function public.requests_workbook_state(uuid, text, uuid, uuid, text, integer) to anon;

create function public.requests_status_plan(session uuid, plan uuid, payload jsonb default null)
returns jsonb
language plpgsql security definer set search_path = ''
as $$
declare
    saved jsonb;
begin
    if not requests_private.has_workbook_write_token() then
        raise sqlstate '42501' using message = 'Workbook access denied';
    end if;
    if payload is not null then
        delete from requests_private.status_plans where expires_at <= now();
        insert into requests_private.status_plans (session_id, plan_id, data, expires_at)
        values (session, plan, payload, now() + interval '5 minutes')
        on conflict (session_id) do update
        set plan_id = excluded.plan_id, data = excluded.data, expires_at = excluded.expires_at;
        return payload;
    end if;
    delete from requests_private.status_plans
    where session_id = session and plan_id = plan and expires_at > now()
    returning data into saved;
    return saved;
end;
$$;
revoke all on function public.requests_status_plan(uuid, uuid, jsonb) from public, authenticated;
grant execute on function public.requests_status_plan(uuid, uuid, jsonb) to anon;

create function public.requests_chat_allowed()
returns boolean
language plpgsql security definer set search_path = ''
as $$
declare
    calls integer;
begin
    if not requests_private.has_workbook_write_token() then
        raise sqlstate '42501' using message = 'Workbook access denied';
    end if;
    insert into requests_private.chat_budget (singleton, window_start, used)
    values (true, date_trunc('minute', now()), 1)
    on conflict (singleton) do update set
        used = case when requests_private.chat_budget.window_start < excluded.window_start
                    then 1 else least(21, requests_private.chat_budget.used + 1) end,
        window_start = greatest(requests_private.chat_budget.window_start, excluded.window_start)
    returning used into calls;
    return calls <= 20;
end;
$$;
revoke all on function public.requests_chat_allowed() from public, authenticated;
grant execute on function public.requests_chat_allowed() to anon;

commit;
