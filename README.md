# Student Retention Add-in

A Microsoft Excel Office Add-in that helps educators track, analyze, and follow
up with at-risk students. Pulls Master List data from Excel, generates
personalized outreach emails (via Power Automate), and integrates with the
Student Retention Kit Chrome extension.

## Architecture

The add-in has two Office runtimes plus shared code:

- **`react/`** — the visible task pane UI (React + Vite). Renders all the
  feature pages: Student View, Personalized Email, Create LDA, Settings,
  Reports, Welcome.
- **`commands/`** — the hidden long-lived "commands runtime" Office loads
  separately (`commands.html` → `background-service.js`). Handles ribbon
  button clicks, document-open events, Master List import/transfer to the
  Chrome extension, conditional formatting, and the Power Automate
  document-property poller.
- **`shared/`** — code used by both runtimes. The Chrome extension client
  singleton, sheet-name and batch-size constants, Excel helpers
  (`findColumnIndex`, `parseHyperlinkFormula`, `normalizeHeader`), and the
  single source of truth for all column aliases.

Office loads each runtime independently, so they don't share memory at
runtime — they each instantiate their own copy of the shared code.

## Layout

```
.
├── manifest.xml              Office Add-in manifest (registered with Excel)
├── assets/                   Icons referenced by the manifest
├── commands/                 Hidden background runtime
│   ├── commands.html         Office runtime entry
│   ├── background-service.js Script entry (registers ribbon handlers, dispatches extension messages)
│   ├── missing-masterlist-dialog.html
│   ├── __tests__/            Vitest tests
│   └── src/                  Internal modules
│       ├── constants.js              Commands-specific constants
│       ├── ribbon-actions.js         Ribbon button handlers
│       ├── master-list-import.js     Chrome ext → Master List sheet
│       ├── master-list-transfer.js   Master List sheet → Chrome ext
│       ├── conditional-formatting.js Color scales / highlights
│       ├── chrome-extension-messaging.js
│       └── power-automate-poller.js  Custom-property → highlight bridge
├── react/                    Task pane UI
│   ├── index.html
│   ├── package.json
│   ├── config/               vite + vitest + eslint configs
│   ├── dist/                 Built output (served from GitHub Pages)
│   └── src/
│       ├── App.jsx
│       └── components/       Feature pages + reusable utilities
├── shared/                   Cross-runtime code
│   ├── chromeExtensionService.js   Singleton client for the Chrome extension
│   ├── constants.js                Sheet names, batch size
│   ├── columnAliases.js            Single source of truth for column header aliases
│   └── excel-helpers.js            findColumnIndex / parseHyperlinkFormula / normalizeHeader
├── documentation/            Integration guides (Chrome extension, SSO setup, etc.)
└── README.md
```

## Setup

```bash
# Install React deps
cd react
npm install

# Install commands deps (for tests)
cd ../commands
npm install
```

## Run (development)

```bash
cd react
npm run dev          # Vite dev server for the task pane
npm run build        # Production build → react/dist/
npm run lint
```

The commands runtime has no build step — Office loads `background-service.js`
directly via ES modules from `commands.html`.

## Test

```bash
cd commands && npm test    # 205 tests covering shared utilities, column aliases, helpers
cd react   && npm test     # 48 tests covering the pure-function utility surface
```

Components themselves aren't unit-tested (they wrap `Excel.run` /
`Office.context.*` which only exist inside an Office host). Verify those
flows by sideloading the manifest into Excel.

## Deploy

Two environments:

- **Prod — GitHub Pages.** `https://Northbridge-University.github.io/Student-Retention-Add-in/`
  Serves the `react/dist/` committed on `main`. Sideload `manifest.xml`.
- **Staging — `staging` branch via Vercel.**
  `https://student-retention-kit-git-staging-Northbridge-University.vercel.app/`
  Auto-deployed by Vercel on every push to `staging` (build config in
  `vercel.json`). Sideload `manifest.staging.xml`. Use this for full-deploy
  testing before promoting to prod.

Both manifests serve identical add-in functionality — only the host URLs
differ.

### Branch model

```
feature-branch → staging (test on Vercel) → main (prod on GitHub Pages)
```

Feature branches pushed to GitHub also get automatic Vercel preview URLs
(pattern: `student-retention-kit-git-<branch>-Northbridge-University.vercel.app`), but those
hosts aren't registered in Azure AD by default — for full SSO-enabled testing,
merge into `staging` and use `manifest.staging.xml`.

### Promoting to prod

Prod is whatever's committed in `react/dist/` on `main`. To ship:

```bash
cd react
npm run build           # rebuilds dist/ with the GitHub Pages base path
git add react/dist
git commit -m "chore: rebuild dist for prod"
git checkout main
git merge staging       # or open a PR
git push
```

GitHub Pages picks up the new `dist/` within a few minutes.

### Azure AD configuration

Each environment's host needs **two** entries on the Azure AD app registration
(client id `71f37f39-a330-413a-be61-0baa5ce03ea3`):

1. An **Application ID URI** (`api://<host>/<client-id>`) under "Expose an API"
2. A **redirect URI** (`https://<host>/react/dist/index.html`) of type SPA under
   "Authentication"

Currently registered: `Northbridge-University.github.io` (prod) and
`student-retention-kit-git-staging-Northbridge-University.vercel.app` (staging). New
environments require the same two-entry add.
