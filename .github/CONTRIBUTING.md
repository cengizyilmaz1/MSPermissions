# Contributing to Graph Permissions Explorer

Thanks for your interest in improving this project! This repository builds a fully
static reference site for Microsoft Graph permissions and Microsoft first-party
application IDs. The public data is refreshed from upstream sources, normalized into
a deterministic snapshot, validated, and only then rendered into a GitHub Pages
artifact.

## Prerequisites

- [Node.js](https://nodejs.org/) `>=20` (see [`.nvmrc`](.nvmrc))
- npm `>=10`
- For a full upstream refresh: PowerShell 7+, Azure CLI, and Microsoft Graph
  PowerShell modules

## Getting started

```bash
npm install
npm run check      # lint + format check + tests + fixture build
```

## Local development loop

The fastest inner loop uses the checked-in fixture dataset and never touches the network:

```bash
npm run build:fixture
npm run serve
```

To work against the canonical raw inputs in `data/`:

```bash
npm run normalize:data -- --raw-dir data --output .generated/local-real
npm run validate:data -- --input .generated/local-real/site-data.json
npm run build:site -- --input .generated/local-real/site-data.json --output docs
```

## Project layout

| Path                  | Purpose                                                 |
| --------------------- | ------------------------------------------------------- |
| `scripts/node/`       | Refresh, normalize, validate, and build entry points    |
| `scripts/powershell/` | Microsoft Graph / OpenAPI fetch and parse scripts       |
| `src/lib/`            | Normalization, rendering helpers, public data contracts |
| `src/templates/`      | HTML, CSS, and browser JavaScript                       |
| `src/config/`         | Validation thresholds and resource mapping              |
| `data/`               | Canonical raw inputs for local normalize/build runs     |
| `fixtures/raw/`       | CI-safe fixture dataset used by tests                   |
| `docs/`               | Generated output only (never edit by hand)              |

## Code style

- Formatting is enforced with [Prettier](https://prettier.io/) (`npm run format`).
- Linting is enforced with [ESLint](https://eslint.org/) (`npm run lint`).
- Please run `npm run check` before opening a pull request.

## Commit and pull request guidelines

1. Create a feature branch off `main`.
2. Keep changes focused and scoped to a single concern.
3. Ensure `npm run check` passes locally.
4. Never commit generated `docs/` or `.generated/` output.
5. Fill out the pull request template.

## Reporting bugs and requesting features

Please use the [issue templates](https://github.com/cengizyilmaz1/Permissions/issues/new/choose).
For security-related reports, follow [SECURITY.md](SECURITY.md) instead of opening a
public issue.
