# Auto Contacts

Outlook add-in that adds the people you email to your contacts — with phone,
title, company, and website pulled from their email signatures.

Outlook's autocomplete cache remembers addresses but never creates real
contact cards. Auto Contacts fixes that:

- **Add to Contacts** (one click, message ribbon — works on desktop, web, and
  mobile): adds the sender of the open message (or, for a message you sent,
  its recipients). Their signature is parsed for phone numbers, job title,
  company, website, and LinkedIn.
- **Contact Sweep** (task pane — desktop/web): scans your sent mail (default
  30 days) for people who aren't in your contacts, previews the parsed card
  for each, and adds the ones you select.

## Privacy by design

No publisher server, no data collection. Static files are hosted on GitHub
Pages; everything else is Microsoft Graph called directly from your Outlook
session with your own delegated token (`Contacts.ReadWrite` + `Mail.Read`,
your own mailbox only). Signature parsing is deterministic pattern-matching
that runs inside your Outlook session — email content never leaves the
Microsoft 365 boundary. Existing contacts are only ever *enriched* (blank
fields filled); nothing you've entered is overwritten.

## Project layout

- `manifest.xml` — add-in manifest (desktop + web + mobile one-click)
- `src/graph.js` — MSAL nested-app-auth + Graph data layer
- `src/parser.js` — pure signature parser (offline unit tests in `test/`)
- `src/commands/` — one-click Add to Contacts command
- `src/taskpane/` — Contact Sweep pane
- `SETUP.md` — one-time Entra + hosting setup

`npm run validate` checks the manifest; `npm test` runs the parser tests.
