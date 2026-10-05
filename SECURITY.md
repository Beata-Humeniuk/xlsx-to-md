# Security Policy

## Supported versions

Security fixes are released for the latest published version of the
extension. Older versions are not patched separately — please update to the
newest release.

## Reporting a vulnerability

Please use
[GitHub private vulnerability reporting](https://github.com/Beata-Humeniuk/xlsx-to-md/security/advisories/new)
so the issue is not public before a fix exists. If that is not possible,
open a regular issue **without** the sensitive details and ask for a private
channel.

## What not to post

Workbooks often hold internal or personal information. In any report —
public or private — do **not** attach real workbooks. A minimal **synthetic**
workbook with made-up content that reproduces the problem is all that is
needed.

## Scope notes

The extension makes no network requests, sends no telemetry and starts no
programs. It reads the workbook or CSV file you open, and writes only the Markdown file
you choose in the save dialog. Saved filters are kept in VS Code's global
extension storage (and synced by Settings Sync when you use it); they hold the
sheet name, range addresses, header names, conditions and options, and the
values last given for conditions asked when used — no other cell contents. Formulas and macros in the
workbook are never run: the values Excel saved are read as they are. Anything
contradicting that — a network request, a file written elsewhere, workbook
content executed — is a security bug; please report it.
