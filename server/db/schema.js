export async function ensureSchema(pool) {
  await pool.query(`
    create table if not exists auto_bid_users (
      id text primary key,
      first_name text not null,
      last_name text not null,
      email text not null unique,
      password text not null,
      timezone text not null default 'UTC',
      active boolean not null default true,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );

    create table if not exists auto_bid_profiles (
      id text primary key,
      user_id text not null references auto_bid_users(id) on delete cascade,
      name text not null,
      static_fields jsonb not null default '{}'::jsonb,
      resume_text text not null default '',
      preferences jsonb not null default '{}'::jsonb,
      profile_version integer not null default 1,
      active boolean not null default true,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );

    create table if not exists auto_bid_questions (
      id text primary key,
      question_hash text not null unique,
      domain text not null,
      url_pattern text,
      normalized_label text not null,
      field_type text not null,
      options_json jsonb,
      required boolean not null default false,
      cache_scope text not null default 'profile_job',
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now(),
      check (cache_scope in ('global', 'profile', 'profile_job'))
    );

    create table if not exists auto_bid_answer_cache (
      id text primary key,
      question_hash text not null references auto_bid_questions(question_hash) on delete cascade,
      cache_scope text not null,
      profile_id text references auto_bid_profiles(id) on delete cascade,
      profile_version integer,
      job_hash text,
      answer text not null,
      confidence numeric(5, 4),
      source text not null default 'ai',
      expires_at timestamptz,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now(),
      check (cache_scope in ('global', 'profile', 'profile_job')),
      check (source in ('ai', 'user', 'static'))
    );

    create table if not exists auto_bid_application_drafts (
      id text primary key,
      user_id text not null references auto_bid_users(id) on delete cascade,
      profile_id text references auto_bid_profiles(id) on delete set null,
      domain text not null,
      url text not null,
      normalized_url text,
      job_hash text,
      form_hash text,
      field_snapshot jsonb not null,
      answers_json jsonb not null,
      status text not null default 'draft',
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now(),
      check (status in ('draft', 'filled', 'submitted'))
    );

    create table if not exists auto_bid_outlook_connections (
      user_id text primary key references auto_bid_users(id) on delete cascade,
      microsoft_user_id text not null,
      tenant_id text,
      email text,
      display_name text,
      access_token_encrypted text not null,
      refresh_token_encrypted text not null,
      token_expires_at timestamptz not null,
      scopes text[] not null default '{}',
      active boolean not null default true,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );

    create table if not exists auto_bid_outlook_mailboxes (
      id text primary key,
      user_id text not null references auto_bid_users(id) on delete cascade,
      profile_id text references auto_bid_profiles(id) on delete set null,
      microsoft_user_id text not null,
      tenant_id text,
      email text,
      display_name text,
      access_token_encrypted text not null,
      refresh_token_encrypted text not null,
      token_expires_at timestamptz not null,
      scopes text[] not null default '{}',
      active boolean not null default true,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    );

    insert into auto_bid_outlook_mailboxes
      (id, user_id, profile_id, microsoft_user_id, tenant_id, email, display_name,
       access_token_encrypted, refresh_token_encrypted, token_expires_at, scopes,
       active, created_at, updated_at)
    select
      'abo_legacy_' || substr(md5(user_id || ':' || microsoft_user_id), 1, 24),
      user_id, null, microsoft_user_id, tenant_id, email, display_name,
      access_token_encrypted, refresh_token_encrypted, token_expires_at, scopes,
      active, created_at, updated_at
    from auto_bid_outlook_connections
    on conflict (id) do nothing;

    create index if not exists auto_bid_profiles_user_active_idx on auto_bid_profiles(user_id, active);
    create index if not exists auto_bid_questions_domain_scope_idx on auto_bid_questions(domain, cache_scope);
    create index if not exists auto_bid_answer_cache_question_scope_idx on auto_bid_answer_cache(question_hash, cache_scope, created_at desc);
    create index if not exists auto_bid_answer_cache_profile_idx on auto_bid_answer_cache(profile_id, profile_version);
    create index if not exists auto_bid_drafts_user_created_idx on auto_bid_application_drafts(user_id, created_at desc);
    create index if not exists auto_bid_drafts_profile_created_idx on auto_bid_application_drafts(profile_id, created_at desc);
    create index if not exists auto_bid_outlook_connections_active_idx on auto_bid_outlook_connections(active, updated_at desc);
    create index if not exists auto_bid_outlook_mailboxes_user_active_idx on auto_bid_outlook_mailboxes(user_id, active, updated_at desc);
    create index if not exists auto_bid_outlook_mailboxes_profile_idx on auto_bid_outlook_mailboxes(user_id, profile_id, active);
    create unique index if not exists auto_bid_outlook_mailboxes_profile_account_idx
      on auto_bid_outlook_mailboxes(user_id, microsoft_user_id, coalesce(profile_id, ''));
  `);
}
