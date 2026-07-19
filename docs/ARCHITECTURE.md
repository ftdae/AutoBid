# AutoBid Architecture

AutoBid is split into a local API server and a Chrome extension.

## Repository Layout

```text
AutoBid/
  extension/
    background.js                 Chrome service worker and API bridge
    content.js                    content-script runner and legacy helpers
    content-modules/              feature modules injected before content.js
      deterministic-defaults.js   local non-AI defaults such as dates and sliders
    page-helper.js                main-world DOM/React bridge
    popup.*                       profile/settings UI
    manifest.json
  server/
    server.js                     API bootstrap and route wiring
    config.js                     environment and runtime constants
    auth/                         token, password, and email helpers
    assist/                       field policy, answer cache, OpenAI calls
    db/                           PostgreSQL pool and schema creation
    http/                         JSON and CORS helpers
    profiles/                     static profile field matching
    users/                        response serializers
    utils/                        IDs, hashes, text normalization
  docs/
    ARCHITECTURE.md
```

## Extension Rules

- Keep `content.js` focused on orchestration: collect fields, call local modules, call the server, apply answers, and trace results.
- Put reusable autofill logic in `extension/content-modules/`.
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
  assist/
    openai.js
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
