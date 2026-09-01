# AutoBid

Standalone auto-bid assistant with:

- PostgreSQL-backed signup/login
- multiple user profiles
- static profile autofill fields
- ChatGPT extension routing for unresolved required complex fields; OpenAI routing disabled by default
- ATS-aware form adapters with a common fallback engine
- encrypted Outlook verification-message integration
- resizable right-side workspace and persistent execution logs
- dynamic answer handoff through Google Sheets
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

### Backend terminal logs

The backend prints structured, timestamped logs for every HTTP request and response and background assist job. The request or job ID in brackets correlates parallel work. When OpenAI routing is disabled, assist logs report `OPENAI_SKIPPED` with `route-disabled`.

```text
2026-08-27T10:15:30.120Z [AutoBid] [abaj_...] OPENAI_SKIPPED {"reason":"route-disabled",...}
2026-08-27T10:15:31.410Z [AutoBid] [abaj_...] ASSIST_JOB_COMPLETED {...}
```

Secrets are redacted and long page/resume text is truncated. Restart `npm start` after code changes and keep that terminal visible to monitor all parallel autofill requests.

Set `APP_SECRET` and `DATABASE_URL` in `.env`.

Complex required questions go to the ChatGPT browser workers. Fields that remain unresolved are left unfilled; they are not sent to OpenAI. To restore the backend fallback later, set `OPENAI_ROUTE_ENABLED=true` and enable the extension route deliberately.

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
- `auto_bid_outlook_connections`

## Load Extension

1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Click Load unpacked.
4. Select:

```text
~/Documents/AutoBid/extension
```

Default hotkey:

- Autofill on the job page: `Ctrl+Q`
- Autofill command on macOS: `Ctrl+Q`

Press `Ctrl+Q` on a job application page to detect and fill fields immediately. Chrome may preserve a previously assigned extension shortcut, so confirm the command in `chrome://extensions/shortcuts` after reloading the extension.

Clicking the Auto Bid toolbar icon opens a resizable right-side workspace on normal web pages. Chrome-owned pages such as `chrome://extensions` cannot accept injected panels, so Auto Bid opens its detached window there instead. The workspace contains Dashboard, Profiles, Outlook, Log, and Settings views. Use Chrome's Extension options command to open the full-width settings dashboard.

The common field engine is extended by adapters for Ashby, BambooHR, Gem, GoHire, Greenhouse, HiBob, iCIMS, JazzHR, Jobvite, Lever, Oracle Recruiting, SAP SuccessFactors, Personio, Rippling, Recruitee, SmartRecruiters, Sourceflow, Teamtailor, Wellfound, Workable, Workday, and Work at a Startup. Unknown application systems continue through the common semantic collector.

Custom dropdowns use Chrome's debugger permission to send native mouse clicks; keep DevTools closed on the application tab while autofill runs. After runtime GPT answers, direct AI fallbacks, and resume upload are applied, AutoBid can click the final submit/apply button when required fields are complete.

## Outlook Verification Messages

Outlook integration uses delegated Microsoft Graph access. Access and refresh tokens are encrypted with a key derived from `APP_SECRET` before PostgreSQL storage; they are never exposed to job-page scripts.

1. Create a Microsoft Entra app registration.
2. Add a **Web** redirect URI in this exact format, using the ID shown for Auto Bid in `chrome://extensions`:

```text
https://YOUR_EXTENSION_ID.chromiumapp.org/outlook
```

3. Add delegated Microsoft Graph permissions `User.Read` and `Mail.ReadWrite`.
4. Create a client secret and configure `.env`:

```text
MICROSOFT_CLIENT_ID=your-application-client-id
MICROSOFT_CLIENT_SECRET=your-client-secret
MICROSOFT_TENANT_ID=common
MICROSOFT_OUTLOOK_SCOPES=openid,profile,email,offline_access,User.Read,Mail.ReadWrite
```

5. Restart the API, reload the unpacked extension, open **Outlook**, and click **Connect Outlook**.

Auto Bid scans recent Inbox and Junk messages for application/account verification content. It displays likely messages, extracts verification codes and links, supports marking messages read, and can fill a required verification-code field before any AI request.

## Google Sheets Context And Runtime GPT

Use PostgreSQL for users and profiles, and Google Sheets for job rows, JD/resume context, and tailored resume file links. Complex answer exchange now uses Chrome runtime messaging instead of Google Sheet columns.

### Option A: Apps Script Web App

This is the recommended local-development path if you already use an Apps Script Web App URL and extension secret.

1. Paste [docs/apps-script-autobid-bridge.gs](docs/apps-script-autobid-bridge.gs) into your existing Apps Script project.
2. Add the `doPost` branches shown at the bottom of that file.
3. Deploy a new Web App version.
4. Add these values to `.env`:

```text
GOOGLE_APPS_SCRIPT_WEB_APP_URL=https://script.google.com/macros/s/.../exec
GOOGLE_APPS_SCRIPT_SECRET=your-existing-extension-secret
```

With this option, you do not need `GOOGLE_SERVICE_ACCOUNT_EMAIL` or `GOOGLE_PRIVATE_KEY`.

### Option B: Direct Google Sheets API

Use this if you do not want Apps Script in the middle.

1. Create a Google service account and share the target spreadsheet with the service account email as Editor.
2. Add these values to `.env`:

```text
GOOGLE_SERVICE_ACCOUNT_EMAIL=your-service-account@project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
GOOGLE_SHEETS_SPREADSHEET_ID=optional-default-spreadsheet-id
```

### Sheet Format

In the sheet header row, include a job URL column named one of: `job_url`, `apply_url`, `application_url`, `url`, or `link`.

AutoBid sends these fixed columns as first-class GPT context when present:

- Column G: tailored resume/profile content
- Column M: job description

In the extension popup, enter the sheet tab name, start row, and end row. If you use Option B, also enter the spreadsheet ID or full spreadsheet URL. Click `Open` to open those job application URLs in Chrome.

AutoBid can still create these columns for fallback/debugging:

- `autobid_questions`
- `autobid_answers`
- `autobid_status`
- `autobid_updated_at`

When you press `Ctrl+Q` on an opened job page, AutoBid fills profile/static fields locally first. Basic profile fields such as name, email, phone, city, country, location, LinkedIn, GitHub, portfolio, salary, notice period, work authorization, sponsorship, languages, and years/skill selectors are not sent to GPT or API AI.

After local filling, AutoBid sends only the remaining required complex fields through this route:

```text
saved profile/local fill -> one ChatGPT extension attempt -> one OpenAI request for unresolved fields -> stop
```

Optional fields never make this AI round trip. Basic profile fields continue to use pre-saved profile data locally. Resume upload runs in parallel. OpenAI is skipped entirely when ChatGPT fills every requested field.

```json
{
  "requests": [
    {
      "request_id": "abg_exact_request_id",
      "answers": [
        { "field_id": "ab_12_example", "value": "Short answer to fill" }
      ]
    }
  ]
}
```

## Storage

Users, profiles, cached answers, detected questions, and drafts are stored in PostgreSQL using `DATABASE_URL`.
