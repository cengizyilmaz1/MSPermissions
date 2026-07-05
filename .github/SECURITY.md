# Security Policy

## Supported versions

This project publishes a single, continuously deployed static site. Only the latest
`main` branch and the currently published site are supported.

| Version         | Supported |
| --------------- | --------- |
| `main` (latest) | ✅        |
| older commits   | ❌        |

## Reporting a vulnerability

Please **do not** open a public GitHub issue for security vulnerabilities.

Instead, report privately through one of the following channels:

- GitHub private vulnerability reporting:
  [Report a vulnerability](https://github.com/cengizyilmaz1/MSPermissions/security/advisories/new)
- Email: `contact@cengizyilmaz.net`

When reporting, please include:

- A description of the vulnerability and its impact
- Steps to reproduce or a proof of concept
- Any relevant logs, URLs, or affected files

## What to expect

- Acknowledgement of your report within a reasonable timeframe.
- An assessment of the issue and, where valid, a remediation plan.
- Public disclosure only after a fix is available, credited to you if desired.

## Scope

This is a static, data-only site with no authenticated backend and no user data
collection beyond anonymized analytics. Relevant concerns include:

- Cross-site scripting via rendered permission or app metadata
- Supply-chain issues in build dependencies
- Data integrity of the published permission and app catalogs
