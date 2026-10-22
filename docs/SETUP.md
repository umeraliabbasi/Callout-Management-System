# Setup Guide

Everything runs under **your dedicated Google account**. Do all of this while signed into that account.
(I can't create these in your Drive directly — no account access — so this is the short list of
in-account steps. All the files are pre-built; you're mostly uploading and pasting.)

Files (in `src/` and `templates/`): `Master_Callout_Sheet_TEMPLATE.xlsx`, `Coverage_Sheet_TEMPLATE.xlsx`, `Code.gs`, `Index.html`.

---

## 1. Upload the two sheets
1. In the dedicated account's Google Drive, upload **Master_Callout_Sheet_TEMPLATE.xlsx**, then right-click →
   Open with → Google Sheets. Once it opens as a Google Sheet, delete the leftover `.xlsx`.
2. Same for **Coverage_Sheet_TEMPLATE.xlsx**.
3. Copy the Coverage sheet's **ID** from its URL (the long string between `/d/` and `/edit`).

The master already has everything: a sample roster (10 fake names, replace with your own, counters at 0), the Config tab (every recipient set to a test address you own for testing),
the All Callouts intake, and the Proper / Improper / day / Tour Training / Invited / Archived tabs
(headers ready; the script fills the filter formulas in step 4).

---

## 2. Add the Apps Script
1. From the master sheet: **Extensions → Apps Script**.
2. Paste **Code.gs** into the default file. Add a file named **Index** (HTML) and paste **Index.html**.
3. Set `COVERAGE_SPREADSHEET_ID` to the Coverage ID from step 1.
4. **Project Settings (gear) → Time zone → America/New_York.**  ← required for the 9am/72h logic.
5. Save.

---

## 3. Wire it up (Callout Tools menu)
Reload the master sheet; a **Callout Tools** menu appears. Run these once (approve permissions when asked):
1. **Set up / rebuild filter tabs** — fills the Proper/Improper/day/Tour Training/Invited QUERYs.
2. **Apply row highlighting** — orange when an ambassador's improper counter passes 5, yellow when
   proper passes 10 (orange wins).

(**Reset ALL counters to 0** is the semester reset — run it when you want everyone back to zero.)

---

## 4. Deploy the web app
1. **Deploy → New deployment → Web app.**
2. **Execute as:** Me. **Who has access:** Anyone with the link.
3. Copy the `/exec` URL — that's the link ambassadors open on any device.

---

## 5. Test (everything comes to you)
Open the URL on your phone and submit a few:
- a **proper** callout → row lands in All Callouts + Proper, a Coverage row appears, Proper counter +1,
  a confirmation email hits your inbox.
- an **after-9am day-of** callout → lands in Improper, Improper counter **+2**, you get the confirmation
  (with a "call the SOD" note) plus the DTS + SOD emails — all to your test address.

Because every Config value and every roster email is your address, **all mail routes to you** right now.
When you go live for real, just edit those cells (Config recipients + roster emails) — no code change,
no redeploy.

---

## Notes
- **Self-contained:** unlike the old sheet, this master has no IMPORTRANGE — the filter tabs read the
  local All Callouts, so there's nothing to re-authorize.
- **Counters are submit-only** (rows won't be deleted), so they only go up until a reset.
- The old COUNTIFS analytics tabs (Performance Tracker, Callouts by Shift/Day) are intentionally gone;
  the roster counters replace them.
