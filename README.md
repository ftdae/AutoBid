# AutoBid

Standalone auto-bid assistant with:

- PostgreSQL-backed signup/login
- multiple user profiles
- static profile autofill fields
- GPT-generated answers for dynamic required fields
- scoped answer cache
- Chrome extension hotkey autofill

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the codebase structure and rules for adding new modules.

## Run Server

```bash
cd ~/Documents/AutoBid
cp .env.example .env
nano .env
npm install
npm start
```

The server defaults to:

```text
http://localhost:7003
```

Set `APP_SECRET`, `DATABASE_URL`, and `OPENAI_API_KEY` in `.env`.

If you do not already have PostgreSQL running, start the included local database:

```bash
docker compose up -d db
```

The included database publishes PostgreSQL on host port `5433` to avoid conflicts with other local Postgres/Supabase services.

If you use your own PostgreSQL server, create the database first if it does not exist:

```bash
createdb autobid
```

The app creates these tables automatically on startup:

- `auto_bid_users`
- `auto_bid_profiles`
- `auto_bid_questions`
- `auto_bid_answer_cache`
- `auto_bid_application_drafts`

## Load Extension

1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Click Load unpacked.
4. Select:

```text
~/Documents/AutoBid/extension
```

Default hotkey:

- Windows/Linux: `Ctrl+Shift+Y`
- macOS: `Command+Shift+Y`

Press the hotkey on a job application page to detect and fill fields immediately. Custom dropdowns use Chrome's debugger permission to send native mouse clicks; keep DevTools closed on the application tab while autofill runs. The extension does not submit applications.

## Storage

Users, profiles, cached answers, detected questions, and drafts are stored in PostgreSQL using `DATABASE_URL`.
