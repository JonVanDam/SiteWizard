# SiteWizard

An Electron desktop app for triaging a list of sites from an Excel column: browse each
site in a real embedded browser, then either generate a Word screenshot report, mark it
False Positive / Not Sure, or skip it.

## Setup

Requires [Node.js](https://nodejs.org/) (LTS). Clone the repo, then from its folder:

```
npm install
npm start
```

### Launching it later

Once the first `npm install` has run, you can just **double-click `SiteWizard.bat`** in
the project folder instead of using the terminal. It starts the app directly, and if
`node_modules` is missing it runs `npm install` for you first. Right-click it and choose
*Send to > Desktop (create shortcut)* if you want it somewhere handier.

## Using it

Setup lives in the collapsible sidebar on the left. The **☰** button hides it to give the
browser more room; the one in the toolbar brings it back.

1. **Load Excel File…** — pick the source `.xlsx`/`.xls`, then pick the **Sheet** from the
   dropdown if it isn't the first one. The workbook, sheet and docked columns are
   remembered, so the next launch reopens them and jumps straight to the first
   unprocessed row. Anything that has since moved or been renamed is reported in the log
   and skipped rather than treated as an error.
2. **Dock the columns.** The sheet's header row is read automatically and each column name
   appears as a draggable chip. Drag a chip onto a slot:

   | Slot | Required | What it's used for |
   | --- | --- | --- |
   | URL | yes | the site to browse |
   | Status | yes | where False Positive / Not Sure is written |
   | Comment | no | the comment box above the browser |
   | Gevolg | no | the follow-up measures chosen when a report is generated |
   | Reference | no | the report reference (`I085-2026`) |
   | Infractions | no | the articles ticked when a report is generated |

   A slot whose name isn't in the header row is created in the next free column. Clear a
   slot with the **×**. **Apply Columns** stays disabled until URL and Status are docked.
3. **Gevolg options sheet** — which sheet holds the list of follow-up measures. It lists
   the workbook's sheets; if `Gevolg opties` doesn't exist yet it's offered anyway and
   created on demand.
4. **Apply Columns** — scans down from the header row and jumps to the first row that has
   a URL but no status yet (so re-launching the app resumes where you left off). The log
   reports how many rows are in play.

   **Filters are honoured.** If the sheet has an AutoFilter applied, rows the filter
   excludes are skipped entirely — SiteWizard never navigates to them and never writes to
   them. Because Excel implements filtering by hiding rows, any row hidden by hand is
   skipped the same way. Change the filter in Excel, save, then re-load the file and press
   **Apply Columns** again to pick up the new selection.

   **Cells listing more than one site.** Where a URL cell holds a site plus its mirrors
   (`winbeast.com + winbeast1.com`), only the first is opened and reported on. Separators
   recognised are `+`, `,`, `;`, line breaks and runs of spaces. The cell itself is never
   rewritten, so the other URLs stay in the sheet.
5. Browse the site normally in the embedded pane. A spinner beside the toolbar shows
   while a page is loading, and a site that refuses to load is reported **in the pane
   itself** with a plain-language reason rather than only in the activity log, along with
   three buttons: **Open in your browser**, **Mark as No access** and **Try again**. The
   **↗** button in the toolbar does the same as the first, for a site that did load. (it's a real Chromium view, not an
   iframe, so it isn't blocked by sites that refuse to be framed — you can click links,
   scroll, log in, etc.).
6. **Comment** field — free text, saved to the comment column as soon as you click away
   from it (or immediately before any button below moves you to a different entry).
7. Buttons:
   - **Undo** / **Redo** — step backward/forward through your last status marks, comments,
     gevolg selections and reference writes (not Skips, which never touch the file).
   - **Generate Report** — opens the capture window (see below). Nothing is written until
     you press Generate in there. Relabels itself **Regenerate Report** once a report
     exists for the entry.
   - **Delete Report** — deletes the current entry's report file, after a confirmation
     prompt.
   - **False Positive** / **Not Sure** / **No access** — writes that status and saves the
     workbook immediately. Whether it then advances is controlled by the **Behaviour**
     toggle in the sidebar; turn it off to stay on the entry and add a comment or queue a
     capture before moving on.
   - **←** — jumps back to the entry you were on immediately before this one.
   - **→** — advances without touching the Excel file at all.

   The triage verbs sit on their own row, and the layout tightens below 1000px, because
   agents typically run this side by side with the report window.

## Generating a report

**Generate Report** queues a capture and returns immediately — you keep triaging while it
runs. A **Captures** strip appears above the buttons showing each job's progress, with
**Cancel** while it runs and **Review** once it's ready. Two captures run at a time;
the rest wait their turn.

The crawl loads the site's homepage at A4 width so the page lays out like a printed sheet,
then collects the links on it, keeps the same-domain ones, and loads each in turn — one
level deep, 20 pages maximum. Tall pages are captured as several stacked slices (up to
four). Screenshots are written to a temp folder rather than held in memory, so a queue of
sites doesn't grow the app's footprint.

**Review** opens the report window, docked to the right of the main window with the two
sharing the work area. If several captures are ready it shows **a tab per capture**, so
they can be dealt with as they land; generating one moves straight to the next and the
window only closes when none are left. A capture that finishes while the window is open
adds its tab without disturbing what you are working on.

For the capture being reviewed:

- Every screenshot is listed with a thumbnail, its page URL and which slice it is. All are
  kept by default; untick the ones you don't want, or use **Select all** / **Select none**.
- **Click or right-click a thumbnail** to open it fullscreen at native resolution. Escape
  closes it.
- **Reference** is filled in at the top. If the row already has one it is reused as-is;
  otherwise the next free `I###-YYYY` is worked out by scanning the whole workbook. It is
  editable, and the badge says whether it's new or existing.
- **Agent Name**, **Datum onderzoek**, **Bron**, the **Infractions** articles and the
  **Gevolg** measures are chosen in the right-hand panel. The agent name is remembered between runs; Bron is pre-filled from
  the row's own Bron column when the sheet has one.

Only when you press **Generate Report** in that window is the `.docx` written and the
sheet updated — the gevolg measures, the infractions and the reference go in then, and
never before. Cancel leaves everything untouched, and the job's temp files are cleared
once the report is written.

Because a job holds its own row, a capture queued earlier still writes to the right entry
even if you have moved on several rows by the time you review it.

## Report template requirements

Select any `.docx` as the template via **Select Report Template…**. In that document:

| Tag | Filled with |
| --- | --- |
| `{site}` | the URL being reported on |
| `{refnr}` | the report reference |
| `{datum}` | the investigation date |
| `{controleur}` | the agent name |
| `{bron}` | the source |

For screenshots, use a `docxtemplater` loop. Each iteration carries the page the shot came
from, so the URL can be printed above the image:

```
{#screenshots}
{pageUrl}
{%image}
{/screenshots}
```

Each selected screenshot becomes one iteration. Earlier forms of this tag — a single
`{%screenshots}`, or a loop body of `{%.}` — no longer work; update the template to the
form above.

For the follow-up measures, three tag shapes are available — use whichever suits the
document:

```
{#gevolgAll}{mark} {label}{/gevolgAll}    every option, ☒ when ticked and ☐ when not
{#gevolg}{.}{/gevolg}                     only the ticked measures, one per line
{gevolgText}                              only the ticked measures, on one line
```

The infraction articles have the same three shapes: `{#inbreukAll}{mark} {label}{/inbreukAll}`,
`{#inbreuk}{.}{/inbreuk}` and `{inbreukText}`.

`{#gevolgAll}` and `{#inbreukAll}` reproduce a paper tick-box list; the other forms only
mention what was actually selected. All tags are always available, so a template can use
any subset.

You must also pick an **Output Folder** before generating reports; each report is saved
there as `report_<sanitized-url>.docx` — one file per entry.

## Gevolg (follow-up measures)

**Where the options come from.** They're read from a separate sheet in the loaded
workbook (`Gevolg opties` unless you pick another), one option per row under a header row
in the first column. If that sheet doesn't exist, SiteWizard creates it and seeds it with
a default list. After that the sheet is the source of truth — edit it in Excel to change
the list, and re-click **Apply Columns** to reload.

An option whose text ends in `:` gets a free-text box next to it, so
`Doorsturen naar andere dienst :` is stored as `Doorsturen naar andere dienst : Dienst X`.
Word's `Klik of tik om tekst in te voeren.` prompt is stripped automatically if you paste
options straight out of the report template.

**What gets written.** The ticked measures go into the gevolg column of the row as one
cell, separated by ` | `, at the moment the report is generated. Changes are undoable.

**Carry-over.** Whatever you ticked last time is pre-selected the next time the capture
window opens, so a run of sites getting the same treatment is quick to mark. A row that
already has measures recorded shows its own values instead.

## Remembered settings

The workbook path, sheet, docked columns, options sheet name, agent name, auto-advance
toggle, report template path and output folder are all remembered between launches, in
`sitewizard-settings.json` under Electron's per-user app data folder. On startup the
workbook is reopened and the columns reapplied automatically.

## Known limitations

- Screenshots are as tall as the display allows, not a true A4 sheet. A window cannot be
  taller than the screen work area, and neither offscreen rendering nor the DevTools
  protocol lifts that limit, so pages are rendered at A4 *width* (which is what makes them
  lay out like a printed page) and captured in viewport-height slices.
- Captures run two at a time. More would thrash memory, since each job is a full Chromium
  window.
- Crawling is one level deep and same-domain only. Pages reachable only through a menu
  that needs JavaScript interaction, or on a different host, are not visited.
- The error page's buttons work by navigating to a reserved `.invalid` host that the main
  process intercepts. They deliberately do **not** go through an IPC bridge: the view that
  shows the error page is the same one that loads the sites under investigation, and a
  preload there would expose that bridge to them. Two guards apply — the view must be
  showing the error document, and that document must be the local file — so a remote page
  cannot trigger one.
- Docking resizes both windows to split the work area. It is skipped when the display is
  too narrow to leave the main window at least 640px.
- **Undo/Redo** and **Previous** history are per-session (in memory only); they reset when
  you reload the sheet or restart the app. The Excel file itself is always safe either way
  since every write is saved immediately.
- Excel writing goes cell-by-cell directly against the loaded workbook (rather than
  rebuilding the whole sheet) so formulas/formatting on untouched cells and other sheets
  are left alone, but SheetJS's free/community engine still has limited fidelity for some
  formatting (e.g. conditional formatting, some styles) on save. Keep a backup of the
  source file if that matters.
- Dropdown lists whose source range sits on another sheet are stored by Excel as x14 data
  validations inside a worksheet's `<extLst>`, which SheetJS's writer drops. SiteWizard
  captures those blocks when the file is opened and puts them back after every save, so
  the dropdowns survive. Parts it does **not** restore, because they are metadata rather
  than sheet content: `customXml/*` (Office/SharePoint document properties),
  `xl/printerSettings` (page setup) and `xl/featurePropertyBag`. If a workbook depends on
  those, work on a copy.
- Filter state is read when the file is loaded, not watched live. If you re-filter the
  sheet in Excel while SiteWizard has it open, reload the file to pick up the change.
- The top row of the sheet's used range is assumed to be the header row.
