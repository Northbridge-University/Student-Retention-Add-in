# CLAUDE.md

Project-specific guidance for Claude. Read this before making changes.

## What this project is

Microsoft Excel Office Add-in for student retention tracking. Lives in
two Office runtimes plus shared code:

- `react/` — React + Vite task pane UI (visible)
- `commands/` — hidden long-lived Office "commands runtime" (HTML loaded
  by `commands/commands.html` → `background-service.js`); registers
  ribbon handlers, dispatches Chrome extension messages, polls a custom
  doc property for Power Automate-initiated highlights
- `shared/` — code used by both runtimes: `chromeExtensionService.js`
  singleton, sheet/batch constants, Excel helpers (`findColumnIndex`,
  `parseHyperlinkFormula`, `normalizeHeader`, `sheetNameDateVariants`),
  `highlightQueue.js` (pure coalescing for the highlight request queue),
  and `columnAliases.js`. Highlight requests from the extension are
  executed ONLY by the commands runtime (the taskpane calls
  `setHighlightExecutionEnabled(false)`), serialized through the queue.

## Deploy model

```
main    → GitHub Pages → prod (manifest.xml)
staging → Vercel branch → dev (manifest.staging.xml)
```

Vercel only builds the `staging` branch (Ignored Build Step in
project settings); other pushes skip Vercel entirely. GitHub Pages
serves the committed `react/dist/` on `main` — so promoting to prod
requires a local `npm run build` then commit + push.

Branch off `staging` for work, merge into `staging` to test on
Vercel, PR `staging` → `main` to ship.

## Critical conventions

### Column aliases — single source of truth

`shared/columnAliases.js` exports ~28 alias arrays (one per concept:
`STUDENT_NAME_ALIASES`, `GRADE_BOOK_ALIASES`, etc.). Both runtimes'
COLUMN_MAPPINGS-style maps compose from these — NEVER inline an
alias array. Add or change aliases there and they propagate.

Each alias is one canonical lowercase-and-whitespace-stripped form.
DO NOT enumerate case or whitespace variants ("Grade", "GRADE",
"grade book", "gradebook") — `normalizeHeader` handles all of those
via NFKC + lowercase + whitespace strip.

### Header lookup contract

Callers must pre-normalize headers via `normalizeHeader` before
passing to `findColumnIndex`. Pattern:

```js
import { findColumnIndex, normalizeHeader } from '<shared>/excel-helpers.js';
const normalized = headers.map(normalizeHeader);
const idx = findColumnIndex(normalized, CONSTANTS.COLUMN_MAPPINGS.grade);
```

`findColumnIndex` normalizes the alias internally, but expects headers
already-normalized for performance (called many times against the
same headers).

### Layout rule for `commands/`

- Top level = files Office loads directly (HTML pages, the script
  referenced by `commands.html`)
- `commands/src/` = JS modules imported by other JS
- DO NOT put module code at top level of `commands/`

## Tests

```
cd commands && npm test     # 205 tests (constants, excel-helpers, columnAliases)
cd react   && npm test      # 48 tests (helpers, allowlist)
```

Both runtimes have vitest set up. Tests live in `__tests__/`
directories (top-level for commands, colocated for react). The
`columnAliases.test.js` matrix has 113 named tests verifying every
alias either runtime has ever matched is still reachable — if you
touch alias data and these fail, you've dropped a previously-matched
form.

## What NOT to do

- **Don't unit-test components** that wrap `Excel.run` or
  `Office.context.*`. Those globals only exist inside an Office host;
  mocking them is a tarpit and gives false confidence. Verify those
  flows by sideloading a manifest in Excel.

- **Don't add backwards-compat shims** for code we just removed
  (`forMSGraphAccess` flags, consent-dialog HTMLs, the Azure Function
  token exchange path). All of that is dead — the flow is now
  Office SSO + Power Automate webhook for emails.

- **Don't commit secrets**, `.env` files, or `azure-function/local.settings.json`-style
  templates. Office Add-in SSO uses an Azure AD app with a published
  client ID (`71f37f39-a330-413a-be61-0baa5ce03ea3`); no client
  secret runs in this codebase.

- **Don't change `vite.config.js`'s `base`** to fix Vercel asset
  paths. Vercel's build override (`--base=/react/dist/` in
  `vercel.json`) handles that; the config-file value is correct for
  GitHub Pages prod and shouldn't change.

- **Don't add wildcards to Azure AD redirect URIs** — Microsoft
  doesn't support them for SPA platform. Each new host needs an
  explicit Application ID URI + redirect URI added to the app
  registration.

## Useful files

- `README.md` — human-facing setup/architecture/deploy
- `shared/columnAliases.js` — all column aliases
- `shared/excel-helpers.js` — `findColumnIndex`, `parseHyperlinkFormula`, `normalizeHeader`
- `commands/src/master-list-import.js` — biggest file (~830 lines); the
  Master List ingestion pipeline. Has natural seams (validate → read →
  build merge plan → write → format) if it ever needs splitting.
- `manifest.xml` / `manifest.staging.xml` — Office Add-in manifests
- `vercel.json` — Vercel build config
