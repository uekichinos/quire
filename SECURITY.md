# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| 0.x     | Yes       |

## Reporting a Vulnerability

If you discover a security vulnerability, please **do not** open a public GitHub issue.

Instead, report it privately via GitHub:
[https://github.com/uekichinos/quire/security/advisories/new](https://github.com/uekichinos/quire/security/advisories/new)

Please include:
- A description of the vulnerability
- Steps to reproduce
- Potential impact

You can expect a response within **72 hours**. If the vulnerability is confirmed, a fix will be released as soon as possible and credited to you (unless you prefer to remain anonymous).

## Threat model notes

`@uekichinos/quire` **writes** `.xlsx` files; it does not parse untrusted spreadsheets, which removes the whole class of parser vulnerabilities (zip bombs, XXE, path traversal, prototype pollution via merge) that affect reader libraries.

Callers remain responsible for:
- **CSV/formula injection** — if you export user-supplied text and the file may be re-opened as CSV, sanitise leading `= + - @` characters yourself, or write values as text.
- **Resource use** — a very large workbook still allocates memory proportional to its size.
