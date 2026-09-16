# Job Lead Engine

A Google Sheet plus Apps Script that runs an executive job search end to end: it finds roles,
works out who to email, spots warm connections, drafts the intro, sends it from Gmail when you
tick a box, and logs the outreach to Insightly.

**Version 1.0.0 · Built 2026-09-15 · Alexander Rea · alexander@alexanderrea.com**

## What it does

1. Searches job listings (LinkedIn, Indeed, Glassdoor, company sites) through the JSearch API. LinkedIn is never scraped.
2. Filters on title, location and pay, then has Claude score each job against your brief.
3. Pulls the full listing, and a pay estimate when the listing shows none.
4. Finds likely hiring contacts and their emails with Hunter.io.
5. Cross-references your LinkedIn connections export and your Gmail history for warm paths.
6. Has Claude pick the best contact and draft a short intro email.
7. You review in the sheet and tick Send, immediately or into a scheduled send window.
8. Logs each send to Insightly and to a Sent Log tab, and watches for replies.

## Files

| File | What it is |
|---|---|
| `Code.gs` | The whole script. Paste into Extensions > Apps Script. |
| `appsscript.json` | Manifest. Sets the V8 runtime and the New York time zone. |
| `SETUP.md` | Install steps, every setting, the menu, and troubleshooting. |

## Install

Short version: new Google Sheet > Extensions > Apps Script > paste `Code.gs` and `appsscript.json` >
reload the sheet > **Job Leads > 1. Set up sheet**, then Set API keys, Test API keys, Check setup,
Install schedule. Full detail, including the LinkedIn connections import, is in [SETUP.md](SETUP.md).

## Services used

- **JSearch** (RapidAPI or OpenWeb Ninja) — listings, job details, salary estimates
- **Hunter.io** — hiring contacts and email addresses
- **Anthropic (Claude)** — job scoring, contact choice, email drafting
- **Insightly** — leads, notes and tags for each send
- **Google Apps Script** — Sheets, Gmail, Drive and triggers

API keys live on the sheet's API Keys tab, not in this repository.

## Note

Built with Claude (Anthropic) in Cowork.

