# AutoBid Architecture

AutoBid is split into a local API server and a Chrome extension.

## Repository Layout

```text
AutoBid/
  extension/
    background.js                 Chrome service worker, durable GPT batch queue, and API bridge
    content.js                    content-script runner and legacy helpers
    content-modules/              feature modules injected before content.js
      ats-adapters.js             ATS detection and semantic selectors
      deterministic-defaults.js   local non-AI defaults such as dates and sliders
    gpt-answer-worker.js          disposable ChatGPT-tab batch prompt and response worker
    hotkey-listener.js            page-level Ctrl+Q shortcut
    page-helper.js                main-world DOM/React bridge
    panel-host.js                 resizable right-side iframe host
    popup.*                       dashboard, profiles, Outlook, settings, and logs
    options.*                     full-width options entry point
    manifest.json
  server/
    server.js                     API bootstrap and route wiring
    config.js                     environment and runtime constants
    auth/                         token, password, and email helpers
    assist/                       field policy, answer cache, and OpenAI fallback call
    db/                           PostgreSQL pool and schema creation
    http/                         JSON and CORS helpers
    outlook/                      encrypted Microsoft OAuth and Graph mailbox access
    profiles/                     static profile field matching
    sheets/                       Google Sheets job queue and answer handoff
    users/                        response serializers
    utils/                        IDs, hashes, text normalization
  docs/
    ARCHITECTURE.md
    apps-script-autobid-bridge.gs
```

## Extension Rules

- Keep `content.js` focused on orchestration: collect fields, fill saved profile data locally, send required complex fields to the ChatGPT extension first, send only unresolved fields to OpenAI second, and trace fields that remain empty.
- Put reusable autofill logic in `extension/content-modules/`.
- Keep every ChatGPT runtime request durable and independently owned. Up to 50 batch workers may run concurrently, each worker leases at most five requests, and delayed retries must not block newer runnable requests. A worker closes only after every leased request is complete, cancelled, or durably requeued.
- Treat `request_id` as the routing boundary. A batched response must never be applied to a different job tab or frame.
- Content modules must expose a single namespace on `window.AutoBid...` and avoid top-level globals that can break when the extension is injected multiple times.
- A module should receive dependencies from `content.js` through a `create(helpers)` factory. This keeps browser helpers in one place and makes modules easier to test later.
- `background.js` controls injection order. Any new content module must be injected before `content.js`.

## Backend Rules

Keep `server/server.js` focused on HTTP routing, request composition, and calling domain modules. Do not add new profile matching, cache policy, OpenAI, schema, token, or serialization code directly to `server.js`.

Current backend structure:

```text
server/
  server.js
  config.js
  db/
    pool.js
    schema.js
  auth/
    security.js
  profiles/
    static-fields.js
  sheets/
    google-sheets.js            Apps Script bridge first, direct Sheets API fallback
  assist/
    ai.js                        one OpenAI attempt after the ChatGPT browser worker
    cache.js
    field-policy.js
  http/
    json.js
  users/
    serializers.js
  utils/
    id.js
    text.js
```

New backend features should go into the closest module. Create a new module when a feature has separate ownership, state, or policy.
