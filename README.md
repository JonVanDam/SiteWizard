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

1. **Load Excel File…** — pick the source `.xlsx`/`.xls`. Pick a sheet from the dropdown
   if it's not the first one.
2. **URL column** / **Status column** / **Comment column** / **Gevolg column** — type a
   header name (matched against the sheet's top row, case-insensitive), a column letter
   (`A`, `B`, …), or a 1-based column number (`1`, `2`, …). Status, comment and gevolg
   columns are created automatically in the next free column if their header doesn't
   already exist. Comment and gevolg columns are optional — leave them blank to disable
   the matching field. **Gevolg options sheet** names the sheet holding the list of
   follow-up measures (see below); it defaults to `Gevolg opties`.
3. **Apply Columns** — scans down from the header row and jumps to the first row that has
   a URL but no status yet (so re-launching the app resumes where you left off). The log
   reports how many rows are in play.

   **Filters are honoured.** If the sheet has an AutoFilter applied, rows the filter
   excludes are skipped entirely — SiteWizard never navigates to them and never writes to
   them. Filter the sheet in Excel first, save it, then load it here to work through just
   that selection. Because Excel implements filtering by hiding rows, any row hidden by
   hand is skipped the same way. Change the filter in Excel, save, then re-load the file
   and press **Apply Columns** again to pick up the new selection.
   **Cells listing more than one site.** Where a URL cell holds a site plus its
   mirrors (`winbeast.com + winbeast1.com`), only the first is opened and reported on.
   Separators recognised are `+`, `,`, `;`, line breaks and runs of spaces. The cell
   itself is never rewritten, so the other URLs stay in the sheet.

4. Browse the site normally in the embedded pane (it's a real Chromium view, not an
   iframe, so it isn't blocked by sites that refuse to be framed — you can click links,
   scroll, log in, etc.).
5. **Comment** field — free text, saved to the comment column as soon as you click away
   from it (or immediately before any button below moves you to a different entry).
   Pre-filled with whatever is already in that row's comment cell when you arrive at it.
6. Buttons:
   - **Undo** / **Redo** — step backward/forward through your last status marks and
     comment edits (not Skips, which never touch the file). Undoing jumps you back to the
     row that change was made on and reverts it; Redo reapplies it.
   - **Generate Report** — captures a screenshot of the page as currently shown and fills
     out the selected template (see below). Does *not* advance to the next row. Once a
     report exists for this entry, the button relabels itself **Add Screenshot** — click
     it again (after scrolling/navigating) to append another screenshot to the same
     report.
   - **Delete Report** — deletes the current entry's report file, after a confirmation
     prompt. Enabled only when a report exists for the entry on screen.
   - **False Positive** — writes `False Positive` into the status column for the current
     row, saves the workbook immediately, and advances.
   - **Not Sure** — writes `Not Sure` into the status column and advances. *(Your
     original spec had this button also writing "False Positive" — that looked like a
     copy/paste slip since it'd make the two buttons identical, so it writes "Not Sure"
     instead.)*
   - **← (previous entry)** — jumps back to the entry you were on immediately before this
     one, purely for reviewing/re-editing it; it does not undo anything by itself.
   - **→ Skip** — advances without touching the Excel file at all.

## Report template requirements

Select any `.docx` as the template via **Select Report Template…**. In that document:

- Use `{site}` as a normal text placeholder — it's replaced with the current URL.
- For screenshots, since a report can now hold more than one, use a `docxtemplater` loop
  rather than a single image tag:
  ```
  {#screenshots}
  {%.}
  {/screenshots}
  ```
  Each captured screenshot becomes one iteration of that loop (one image). This replaces
  the older single `{%screenshots}` tag — if you already built a template with that,
  change it to the loop form above.

For the follow-up measures, three tag shapes are available — use whichever suits the
document:

```
{#gevolgAll}{mark} {label}{/gevolgAll}    every option, ☒ when ticked and ☐ when not
{#gevolg}{.}{/gevolg}                     only the ticked measures, one per line
{gevolgText}                              only the ticked measures, on one line
```

`{#gevolgAll}` reproduces a paper tick-box list; the other two only mention what was
actually selected. All three are always available, so a template can use none, one or
several of them.

You must also pick an **Output Folder** before generating reports; each report is saved
there as `report_<sanitized-url>.docx` — one file per entry, so **Add Screenshot**/
**Delete Report** always know which file to update.

## Gevolg (follow-up measures)

If you fill in a **Gevolg column**, a collapsible **Gevolg** panel appears above the
browser with a tick box per measure — several can be ticked at once.

**Where the options come from.** They're read from a separate sheet in the loaded
workbook (`Gevolg opties` unless you name another), one option per row under a header
row in the first column. If that sheet doesn't exist, SiteWizard creates it and seeds it
with a default list, so a workbook that has never been used with SiteWizard still works
on first run. After that the sheet is the source of truth — edit it in Excel to change
the list, and re-click **Apply Columns** to reload.

An option whose text ends in `:` gets a free-text box next to it, so
`Doorsturen naar andere dienst :` is stored as
`Doorsturen naar andere dienst : Dienst X`. Word's `Klik of tik om tekst in te voeren.`
prompt is stripped automatically if you paste options straight out of the report
template.

**What gets written.** The ticked measures go into the gevolg column of the current row
as one cell, separated by ` | `. Changes are undoable with the normal Undo/Redo buttons.

**Carry-over.** Whatever you tick stays ticked when you move to the next entry, so a run
of sites getting the same treatment is quick to mark. The panel shows a carried-over
selection in blue italics. A row that already has measures recorded always shows its own
values instead. Carried-over measures are written to the sheet when you press
**False Positive**, **Not Sure** or **Generate Report**; **Skip** and **←** leave the
file untouched unless you actually changed the panel on that entry.

## Remembered settings

The URL/status/comment/gevolg columns, options sheet name, report template path, and
output folder are remembered
between launches (stored in `sitewizard-settings.json` under Electron's per-user app data
folder). The source Excel file itself is not reopened automatically — pick it again each
run via **Load Excel File…**.

## Known limitations

- The screenshot captured is the current visible viewport of the embedded browser, not a
  stitched full-page capture — scroll to what you want captured before clicking Generate
  Report / Add Screenshot.
- **Add Screenshot** across an app restart re-extracts the previously embedded images
  straight out of the existing `.docx` (they're never re-downsized further), so quality
  doesn't degrade across sessions, but this means the report file itself is the source of
  truth for what's in it — don't hand-edit it outside SiteWizard if you plan to add more
  screenshots later.
- **Undo/Redo** and **Previous** history are per-session (in memory only); they reset
  when you reload the sheet or restart the app. The Excel file itself is always safe
  either way since every write is saved immediately.
- Excel writing goes cell-by-cell directly against the loaded workbook (rather than
  rebuilding the whole sheet) so formulas/formatting on untouched cells and other sheets
  are left alone, but SheetJS's free/community engine still has limited fidelity for
  some formatting (e.g. conditional formatting, some styles) on save. Keep a backup of
  the source file if that matters.
- Dropdown lists whose source range sits on another sheet are stored by Excel as x14
  data validations inside a worksheet's `<extLst>`, which SheetJS's writer drops.
  SiteWizard captures those blocks when the file is opened and puts them back after
  every save, so the dropdowns survive. Parts it does **not** restore, because they are
  metadata rather than sheet content: `customXml/*` (Office/SharePoint document
  properties), `xl/printerSettings` (page setup) and `xl/featurePropertyBag`. If a
  workbook depends on those, work on a copy.
- The top row of the sheet's used range is assumed to be the header row.
- Filter state is read when the file is loaded, not watched live. If you re-filter the
  sheet in Excel while SiteWizard has it open, reload the file to pick up the change.
