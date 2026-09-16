# Job Lead Engine: Setup Guide

A Google Sheet that finds executive roles, figures out who to email, spots warm connections, drafts the intro, and sends it from your Gmail when you tick a box. Every send is logged to Insightly.

## How it works

1. **Search** (Mon + Thu, 7am): pulls listings from LinkedIn, Indeed, Glassdoor and company career pages through the JSearch API. LinkedIn is never scraped, so your account is safe.
2. **Filter**: checks titles against your include/exclude words and locations (remote, or NY/CT/NJ).
3. **Score**: Claude rates each job 1–10 against your brief. Anything under 7 is dropped.
4. **Find the hirer**: Hunter.io looks up senior people and their emails at the company. Claude picks the most likely hiring manager.
5. **Warm paths**: checks the contact and company against your LinkedIn connections and your Gmail history.
6. **Draft**: Claude writes a 90–150 word email, plus a LinkedIn intro request if you know someone there.
7. **You send**: review the row, edit if you want, tick **Send**. The email goes out from Gmail, and Insightly gets a Lead, a tag and a note.
8. **Replies**: every 4 hours it checks for replies, marks the row **Replied** and adds the reply to Insightly.

## What you need

| Service | Why | Cost |
|---|---|---|
| Google account (Workspace or Gmail) | The sheet, the script, and sending | Free |
| [JSearch on RapidAPI](https://rapidapi.com/letscrape-6bRBa3QguO5/api/jsearch) | Job listings, job details, salary estimates | Free: 200 requests/mo. Searches use about 120, and details plus salary estimates add 1–3 per new lead. That leaves room for roughly 25–40 new leads a month. Pro is $25/mo |
| [Hunter.io](https://hunter.io/api-keys) | Hiring contact emails | Free tier for light use. Starter plan from about $34/mo |
| [Anthropic Console](https://console.anthropic.com) | Scoring and drafting | Pay as you go. Roughly a few cents per lead |
| Insightly | CRM logging | API works on every plan. Key is under **User Settings > API** |

## Install (about 15 minutes)

1. Create a new Google Sheet and name it **Job Leads**.
2. Open **Extensions > Apps Script**. Delete the starter code and paste in all of `Code.gs`.
3. In Apps Script, go to **Project Settings**, turn on *Show "appsscript.json" manifest file*, then paste in `appsscript.json`. This sets the time zone to New York. Save.
4. Reload the sheet. A **Job Leads** menu appears.
5. **Job Leads > 1. Set up sheet**. Google will ask you to authorize the script. It needs Sheets, Gmail, external requests, and triggers.
6. **Job Leads > 2. Set API keys**, or type them straight into the **API Keys** tab. Either way they live in that tab, so they travel with the sheet if you copy it to another account. Anyone you share the spreadsheet with can read them, so keep sharing tight.
7. **Job Leads > 3. Test API keys**. You should see four OKs.
8. Go through the **Settings** tab, especially `SEARCH_QUERIES`, `TARGET_ROLE_BRIEF`, `MY_PITCH`, `MY_NAME`, `EMAIL_SIGNATURE` and `INSIGHTLY_POD`. The defaults are placeholders.
9. Import your LinkedIn connections (steps below).
10. **Job Leads > Run job search now** for a first batch, then **4. Install schedule**.

### Import your LinkedIn connections

1. On LinkedIn, go to **Settings & Privacy > Data privacy > Get a copy of your data**.
2. Choose **Connections** and request the archive. LinkedIn emails it, usually within a few minutes.
3. In the sheet, open the **Connections** tab, then **File > Import > Upload** `Connections.csv` and pick **Replace current sheet**.
4. Do this again every month or two.

Your connections live only in your sheet. The script sends Claude just the names and titles of people at the company in each lead, never the whole list.

## JSearch options (Settings tab)

| Setting | Default | What it does |
|---|---|---|
| `WORK_FROM_HOME` | `mixed` | `mixed` searches each location plus a remote-only search from the "Remote" line in `LOCATIONS`. `only` keeps remote jobs only. `none` skips remote jobs |
| `RADIUS_MILES` | `50` | How far around each location to search. The script converts it to km for the API. It doesn't apply to remote searches, and Google treats it as a guide rather than a hard cutoff. Leave blank for no radius |
| `JOB_DETAILS` | `yes` | Pulls the full listing for each new lead. It fills **Apply Options** (including direct company links) and **Role Details** (seniority, experience, industry, whether it manages people, employer rating). It can also supply the company website, which helps Hunter find contacts, and gives Claude more to write from. Costs 1 request per lead |
| `MIN_SALARY` | blank | Annual USD. Drops jobs whose listed pay tops out below this, with hourly and monthly pay annualized. Jobs with no pay listed are kept. Leave blank to turn it off |
| `SALARY_ESTIMATE` | `company` | Runs only when a listing has no pay. `company` asks for that company's pay for the title and falls back to `market` if there's no data. `market` uses title and location. `off` skips it. The result goes in **Salary Estimate**, and Notes flags estimates below `MIN_SALARY` |
| `SALARY_EXPERIENCE` | `ABOVE_FIFTEEN` | The experience band used for estimates |

Every run writes the JSearch requests it used to the **Log** tab, so you can watch your monthly quota.

## Why a job didn't make it into Leads

Every job a search skips is listed in the **Filtered** tab with the exact reason, for example:

- `Title contains excluded word "junior"`
- `Title has none of the TITLE_INCLUDE words`
- `Location "Austin, TX" is outside ALLOWED_STATES`
- `Listed pay $45-$60 / hour (about $125K/yr) is below MIN_SALARY`
- `Fit 4 is below MIN_FIT_SCORE 7`, with Claude's reasoning next to it

Tick **Add to Leads** on any row to move that job into Leads. It gets contacts and a draft a minute or two later.

Each Log summary also counts what was skipped and why, including duplicates (the same job found by several searches) and jobs already in Leads.

**Title matching:** `TITLE_EXCLUDE` only matches whole words, so "intern" skips "Summer Intern" but not "International". `TITLE_INCLUDE` matches the start of a word, so "technolog" matches "Technologist". Words of 3 letters or fewer, like "ai", have to match exactly. Add `*` to match a word start on purpose, e.g. `develop*`.

**Changing MIN_FIT_SCORE:** Claude only scores a job once. If you lower the score, jobs that scored at or above the new minimum on earlier runs get added the next time a search finds them, with no re-scoring. To have everything scored fresh, for example after rewriting `TARGET_ROLE_BRIEF`, use **Job Leads > Forget scored jobs**.

## Daily use

| Status | What to do |
|---|---|
| **Queued** | Ticked and waiting for the next send window. Notes shows when it goes |
| **Ready** | Read the draft, edit Subject or Email Body right in the cell, then tick **Send**. In `draft` mode the email appears in Gmail > Drafts, not Sent |
| **No Email** | Try the **Warm Intro Ask** on LinkedIn, or paste an email into Contact Email and tick Send |
| **Enriching** | Waiting for contacts and a draft. Each run handles a few leads, best fit first, and the next one starts about a minute later until none are left. To start it yourself, use **Find contacts + draft emails now** |
| **Sent / Drafted** | Done. Set it to **Meeting** or **Passed** yourself as things move |
| **Replied** | The reply preview is in Notes. Follow up from Gmail |
| **Error** | The reason is in Notes and the **Log** tab |

- **Rewrite a draft:** type over the Subject or Email Body cell, or select the rows and use **Rewrite drafts for selected rows**, which asks what to change ("shorter and warmer", "lead with the Cannes work") and has Claude write them again. Rewriting only costs a Claude call, no Hunter or JSearch credits, and it skips anything already sent. To change how every future draft reads, edit `EMAIL_STYLE` or `MY_PITCH` in Settings.
- **Send on a schedule:** set `SEND_TIMING` to `window` (it starts as `now`, which sends the moment you tick). Ticking Send then marks the row **Queued** with the date it will go, and the emails leave on the `SEND_WINDOW_DAYS` at `SEND_WINDOW_HOUR` (Tuesday to Thursday at 9am by default). Run **Install schedule** after changing any of those, since that is what sets the timer, and **Check setup** warns you if the settings have moved on since. Google runs the timer at some point inside the chosen hour rather than exactly on the hour, so 9 means roughly 9 to 10. To send a queued batch early, use **Send queued leads now**. To pull one back, set its status to **Ready**.
- **Draft instead of send:** set `SEND_MODE` to `draft` and ticking Send creates a Gmail draft instead. This is a good way to run the first week.
- **Attach your resume:** put the file's name (or its Drive link) in `RESUME_FILE` and set `ATTACH_RESUME` to `yes`. The file has to be in your Google Drive, and the first send after you switch this on asks Google for permission to read it. Claude only mentions the attachment in the draft when it's actually being attached, and a lead won't send at all if the file can't be found. Cold emails with attachments are more likely to be filtered as spam, so linking your resume in `EMAIL_SIGNATURE` is the safer default.
- **Get emails into Insightly:** put your Insightly email dropbox address in `BCC_ADDRESS` so the full email attaches to the record.
- **Nothing sends on its own.** Only ticking the box (or *Send all ticked leads*) sends anything. Ticking a row that's already been sent does nothing.

## Sent Log

Every email that goes out gets a line on the **Sent Log** tab: a number, the date, the company, the job title, a link to the posting, the current status, and a link that jumps to that lead's row on the Leads tab. Status follows the lead, so a row becomes **Replied** there as soon as a reply lands. **Job Leads > Refresh Sent Log** rebuilds it from the Leads tab, which is also how older sends get added.

## When a lead sits on Enriching

Each lead needs a minute of contact lookup and drafting, and runs handle a few at a time, so a short wait is normal. If one is still there after 10 minutes:

1. Run **Job Leads > Find contacts + draft emails now**.
2. Check the **Log** tab. A failure there names the cause, usually a Hunter credit limit or an API key problem.
3. Run **Check setup**, which counts anything stuck or errored.
4. Run **Job Leads > Retry stuck leads**. It resets every stuck or errored lead and starts again.

Each lead is tried three times. Notes shows which attempt is running, and after the third the lead is marked **Error** so it stops holding up the queue. Retry stuck leads gives it another three.

To deal with one by hand, type a contact email into the row, write a Subject and Email Body, and set the status to **Ready**. Or set it to **Passed** to drop it.

## When a send doesn't happen

Run **Job Leads > Check setup**. It reports whether the Send tick is wired up, whether you're in draft or send mode, how many leads are ready, and when the last email went out. The two usual causes:

- **The schedule was never installed.** Ticking Send only does something once **Install schedule** has been run, since that is what connects the checkbox to the script.
- **The row's status blocks it.** A row marked Sent, Drafted, Replied or Meeting will not send again. If you set one of those by hand, put it back to **Ready**.

Every attempt that stops, and why, is written to the **Log** tab.

## Tuning

- **Too much noise:** raise `MIN_FIT_SCORE` to 8, tighten `TITLE_INCLUDE`, or add words to `TITLE_EXCLUDE`.
- **Too few leads:** add queries (one per line), set `DATE_POSTED` to `month`, or add another city to `LOCATIONS`.
- **Budget:** each query runs once per location per search, so 7 queries × 2 locations × 2 runs a week comes to about 120 requests a month. `WORK_FROM_HOME` set to `only` or `none` halves that. `MAX_NEW_LEADS_PER_RUN` caps how much Hunter, Claude and JSearch details get used. To stay on the free tier, set `JOB_DETAILS` to `no` or `SALARY_ESTIMATE` to `off`.
- **Voice:** edit `EMAIL_STYLE` and `MY_PITCH`. Claude only uses facts you've put there.

## Worth knowing

- **Updating from an earlier version:** paste in the new `Code.gs` and that's it. On its next run the script adds any new settings and moves your existing Leads columns into the new layout.

- Hunter emails come with a confidence score. Treat anything under 80 as a guess.
- A warm connection means someone in your LinkedIn export lists that company. It doesn't tell you how well you know them, so check before you mention anyone.
- Gmail's daily sending limit is 100 recipients on a personal account and 1,500 on Workspace.
- Google stops any single run after 6 minutes. Each lead takes up to about a minute (job details, salary, Hunter, Claude), so the script handles a few per run and starts the next run about a minute later. Every Log summary shows how long each step took.
- Rejected jobs are saved to a hidden **Seen** tab so they aren't scored again. Delete a row there to re-score that job.
