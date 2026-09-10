# Matrix Price Finder — Outlook Add-in

A task-pane add-in that reads Texas ERCOT electricity matrix pricing
straight out of supplier emails. Give it a zip code, customer class, usage
tier, start month, and the contract terms you want (12/24/36, etc.) — it
scans every matrix workbook in the email (attached, or linked via an Office
Online viewer link) and returns the price for each term, plus any cheaper
"sweet spot" terms nearby.

This version was built and tested directly against three real supplier
files (NRG, an ERCOT broker matrix, and Chariot Energy), which turned out
to use three different layouts — the engine auto-detects all three.

## What it actually does with your example

> zip 77020, residential, low, starting October 2026, terms 12/24/36

1. Looks up 77020 in whichever zip-mapping sheet it finds → resolves to
   **CenterPoint** (matched across "CNP" / "CNPT" / "CenterPoint" spellings)
   and the **Houston** zone.
2. Filters every matrix sheet down to CenterPoint + Houston + residential +
   low usage + October 2026 start.
3. Returns the 12/24/36-month prices, and separately flags any term (e.g.
   18 or 30 months) that's cheaper than its neighbors — a "sweet spot" —
   even if you didn't ask for that term specifically.

Of the three sample files, only the ERCOT broker matrix had residential
rows (NRG's and Chariot's were commercial-only in the sheets provided) —
the add-in handles both, and a query like this one will only surface cards
from whichever files actually contain matching data.

## Files

```
manifest.xml      Add-in manifest (registers the task pane + ribbon button)
taskpane.html      Search form UI
taskpane.css       Styling
taskpane.js        Layout detection, extraction, filtering, sweet-spot logic
commands.html/js   Required function-file stub for the ribbon button
assets/            Icons
```

## How the layout detection works

Real supplier matrices don't share one shape. This engine scans the first
~40 rows of every sheet for a header row containing recognizable columns
(utility, zone, load profile, start date, term, usage) and classifies the
sheet as one of:

- **WIDE** — one row per (utility, zone, profile, usage cap, start date),
  with contract terms as columns: `Term - 12`, `12 Mth`, or bare numbers
  `1, 2, 3 … 60`. (This is NRG's and Chariot's layout.)
- **LONG** — one row per (utility, zone, profile, **term**, start date),
  with annual-usage kWh brackets as columns, e.g. `0-200,000`. (This is the
  ERCOT broker matrix's layout.)
- **Zip map** — any sheet with a `zip` column plus a `zone` and/or
  `utility`/`TDSP` column. All zip maps found across every attachment are
  merged, so a zip code can resolve even if the utility mapping and the
  price data live in different files.

If neither shape is detected (no term-like columns and no explicit `Term`
column), the sheet is skipped — it's shown in "Files scanned" so you can
see what was and wasn't picked up.

### Utility & zone name matching

Suppliers abbreviate the same utility differently — CenterPoint shows up as
`CNP`, `CNPT`, and `CenterPoint` across just these three files. The engine
normalizes through a synonym table:

```
CENTERPOINT → cnp, cnpt, cpt, centerpoint
AEP_CENTRAL → cpl, aeptcc, aep central          (formerly "CPL" / Central Power & Light)
AEP_NORTH   → wtu, aeptnc, aep north            (formerly "WTU" / West Texas Utilities)
ONCOR       → onc, oncor
TNMP        → tnp, tnmp
LPL         → lpl
```

Zones normalize similarly (`Houston`, `Houston LZ`, `HOUSTON` all match).
**If a new supplier uses an abbreviation not in this list**, add it to
`UTILITY_SYNONYMS` / `ZONE_SYNONYMS` near the top of `taskpane.js` — that's
the only place this needs to change.

### Customer class & usage tier

Classified from the load-profile code:
- Contains `RES` → residential; `HI`/`LO`/`MED` within it → tier.
- Starts with `BUS`, or is bare `HI`/`MED`/`LO` → commercial; same tier logic.
- Anything else → unclassified, and is excluded whenever a specific tier is
  requested (so, e.g., a `BUSNODEM` "no demand" profile never quietly shows
  up under a "low" query).

## Sweet-spot detection

For every matched price grid, terms are sorted and checked for local
minima (a term cheaper than both neighbors) plus the single cheapest term
overall. These are shown as extra pills alongside your requested terms and
highlighted green, and the full term-by-term table (in "All available
terms") highlights the same rows.

## Performance note (important if you extend this)

Some of these workbooks are large — one sample file's "matrix prices_all"
sheet has ~7,000 real data rows × 60 term columns, and the ERCOT sheet has
~140,000 real data rows × 5 usage-bracket columns. Extracting everything
into memory before filtering reliably crashed a Node test with an
out-of-memory error, and would do the same in a browser tab. So filtering
by utility/zone/class/tier/start-month happens **during** extraction, not
after — see the comments in `extractPriceRows` and `collectAllData` in
`taskpane.js` before changing this flow.

Relatedly: one of the sample files had a hidden `UserControl` sheet (a
common side effect of VBA userforms) whose *declared* size was Excel's
absolute maximum — 1,048,561 rows × 16,383 columns — despite being empty.
Converting that to a grid is what actually caused the crash in testing,
more than the real data did. The add-in now skips any sheet reporting more
than 200,000 rows or 500 columns before trying to read it, and lists it as
skipped in "Files scanned" rather than choking on it.

## Linked files (not just attachments)

Some suppliers paste a link to Office's online viewer instead of attaching
the file, e.g.:

```
https://view.officeapps.live.com/op/view.aspx?src=https%3A%2F%2Fcnst.blob.core.windows.net%2F...%2FALL_ELEC_UDCS_PowerBroker_2026_08_31.xlsx&wdOrigin=BROWSELINK
```

The add-in scans the email body for links like this, decodes the real file
URL out of the `src=` parameter (also handling plain direct links to
`.xlsx`/`.xlsm`/`.xls`/`.csv`), and tries to `fetch()` it directly.

**This only works if the hosting server allows cross-origin requests
(CORS) from the add-in's origin.** Many broker/blob-storage links used only
for the Office viewer are *not* CORS-enabled for arbitrary origins — the
viewer itself fetches the file server-side, which doesn't require CORS, but
a client-side `fetch()` from the task pane does. If a link fails, it shows
up in "Files scanned" with the reason (usually a network/CORS error). Two
ways to handle that if it comes up often:
1. Ask the sender to attach the file directly instead of linking it.
2. Stand up a small serverless proxy (an Azure Function or similar) that
   fetches the file server-side and returns it with permissive CORS headers
   — point the add-in's fetch at your proxy instead of the raw blob URL.

## Running it locally (development / testing)

1. **Install Node.js** (18+).
2. Serve the folder over HTTPS:
   ```bash
   npm install -g office-addin-dev-certs http-server
   office-addin-dev-certs install
   http-server . -p 3000 --ssl --cert ~/.office-addin-dev-certs/localhost.crt --key ~/.office-addin-dev-certs/localhost.key
   ```
3. **Sideload the manifest:**
   - **Outlook on the web / new Outlook:** Settings → *Manage add-ins* → *My
     add-ins* → *Add a custom add-in* → *Add from file* → `manifest.xml`.
   - **Outlook desktop (Windows):** Home tab → *Get Add-ins* → *My add-ins*
     → *Add a custom add-in* → *Add from file*.
   - **Outlook on Mac:** Home tab → *Get Add-ins* → *My add-ins* → gear/`+`
     → *Add from file*.
4. Open a supplier email, click **Find Price** in the ribbon, fill in the
   form, and click **Scan email & find prices**.

## Moving to production

- Host the files over HTTPS somewhere real (Azure Static Web Apps, an
  internal server, etc.) and update every `https://localhost:3000` URL in
  `manifest.xml`.
- Roll out org-wide via Microsoft 365 admin center → *Integrated apps*
  instead of manual sideloading.
- Consider bundling `xlsx.full.min.js` locally instead of loading it from
  the CDN, if your organization restricts external scripts.

## Known limitations

- **Units aren't normalized.** The ERCOT file's instructions note prices
  are in cents/kWh; other files use $/kWh decimals. The add-in shows
  numbers exactly as stored — check the source file's header/instructions
  for units before trusting a number in isolation. Normalizing this
  automatically would need a per-file unit hint, which isn't built yet.
- **Multi-level or merged headers** (a header spanning several sub-columns)
  aren't handled — `detectSheetLayout` expects one flat header row.
- **Scanned/image PDFs** aren't supported at all in this version (dropped
  in favor of the spreadsheet-focused engine, since all three sample files
  were Excel). If suppliers regularly send PDF matrices, PDF.js-based
  text-position reconstruction can be added back in — ask and it can be
  layered onto this engine's grid format.
- **Cloud attachments** (a link rather than an actual uploaded file, e.g.
  from OneDrive within Outlook itself) can't be downloaded client-side due
  to CORS; they're flagged in "Files scanned" rather than silently skipped.
