# Shift Call-Out Management System

A Google Apps Script web app plus Google Sheets back office that replaces a long call-out form for a university admissions ambassador program. An ambassador picks their shift, and the system works out the timing tier, whether it counts against them, who must be told, what they still need to do, and where it gets logged.

> All names, emails and phone numbers in this repo are sample data. The roster, recipients and phone numbers in the real deployment live in a private Config sheet and are not included.

## What it does

1. Ambassador picks their name from a type-to-search roster, then a shift category (Regular, Tour Training, Marquee Event, Invited), then the shift and time.
2. The backend classifies the call-out in **New York time** into one of four tiers: 72h+ (proper), under 72h, after 10pm the night before, after 9am day-of.
3. It logs the row to `All Callouts`, which feeds 11 auto-filtered tabs (Proper, Improper, Mon to Sun, Tour Training, Invited).
4. It bumps the ambassador's Proper/Improper counter and highlights rows orange (more than 5 improper) or yellow (more than 10 proper).
5. Proper Regular and Marquee call-outs are appended to a separate Coverage sheet so other ambassadors can pick the shift up.
6. The ambassador gets a confirmation, and the right staff are emailed per the routing matrix (recipients come from the Config tab, so changing staff needs no redeploy).

| Category | Proper | Improper (under 72h / after 10pm) | After 9am day-of |
|---|---|---|---|
| Regular | Coverage sheet | DTS (after 10pm: + SOD + GA/XME) | Same, plus "call the SOD" box |
| Tour Training | Tour Supervisors | DTS + Tour Supervisors (after 10pm: + SOD + GA/XME) | Same, plus "call the SOD" box |
| Invited | Ambassador told to message the inviter | DTS | Same |
| Marquee group A | Coverage sheet | Scheduling + GA/XME | Same |
| Marquee group B | Coverage sheet | DTS (after 10pm: + SOD) | Same, plus "call the SOD" box |

## Repo layout

```
src/Code.gs          Apps Script backend (classification, logging, counters, routing, admin menu)
src/Index.html       Mobile-first web UI (state machine, NYC-time preview, validation)
templates/           Master and Coverage sheet templates with sample data
docs/SETUP.md        Step-by-step setup and deployment
```

## Design decisions worth reading

- **The server decides the tier, not the browser.** The page previews it with a DST-safe NYC conversion (`Intl.DateTimeFormat`), but `classifyServer_` recomputes it under the project time zone and ignores the client value.
- **Self-declared identity.** The app runs from a dedicated Gmail account, so it cannot use university-domain login. Ambassadors pick their name from a roster, and the backend looks up their email. This stops typos and bad data but is not authentication. Every record is visible to staff, which is the compensating control.
- **Config-driven routing.** Every recipient and phone number is a row in the Config sheet.
- **Conditional formatting cannot reference another sheet,** so hidden helper columns pull each ambassador's counters in with an `ARRAYFORMULA` + `VLOOKUP`.
- **Email is best effort.** `safeMail_` catches failures so logging and counters always complete.
- **`appendRow` skips past phantom blank rows** left by copying a sheet, so there is a `trimEmptyRows` admin tool.
- **Deploy gotcha:** Manage deployments, then a new version, keeps the same `/exec` URL. A "New deployment" creates a new URL.

## Known limits and roadmap

- Identity is self-declared (see above).
- Gmail senders can land in spam at university domains. Allowlisting the sender or sending from an organisation group address fixes it.
- In progress: Humanity API integration to remove the ambassador from the shift automatically after a call-out (OAuth2, non-blocking, with a dry-run flag). Not in this code yet.

## Stack

Google Apps Script (JavaScript), HtmlService, Google Sheets (QUERY, ARRAYFORMULA, VLOOKUP), MailApp, Google Forms.

Built by Umer Ali.
