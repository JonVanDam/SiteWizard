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

## Using it

1. **Load Excel File…** — pick the source `.xlsx`/`.xls`. Pick a sheet from the dropdown
   if it's not the first one.
2. **URL column** / **Status column** / **Comment column** — type a header name (matched
   against the sheet's top row, case-insensitive), a column letter (`A`, `B`, …), or a
   1-based column number (`1`, `2`, …). Status and comment columns are created
   automatically in the next free column if their header doesn't already exist. Comment
   column is optional — leave it blank to disable the comment field.
3. **Apply Columns** — scans down from the header row and jumps to the first row that has
   a URL but no status yet (so re-launching the app resumes where you left off).
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

You must also pick an **Output Folder** before generating reports; each report is saved
there as `report_<sanitized-url>.docx` — one file per entry, so **Add Screenshot**/
**Delete Report** always know which file to update.

## Remembered settings

The URL/status/comment columns, report template path, and output folder are remembered
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
- The top row of the sheet's used range is assumed to be the header row.
