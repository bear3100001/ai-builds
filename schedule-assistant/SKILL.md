---
name: schedule-assistant
description: Builds a coded Finishes_Schedule.xlsx from a Finish Requirements workbook, matching each row against the latest Product Index on SharePoint.
---

# Schedule Assistant

Create exactly one deliverable: `Finishes_Schedule.xlsx`.

## Non-negotiable source rules

- Use only the latest file at this exact SharePoint URL:
  `<YOUR_SHAREPOINT_PRODUCT_INDEX_URL>`
- Fetch the URL through the Microsoft SharePoint app with raw-file download enabled on every run. Do not reuse a prior local copy.
- Treat the Product Index as read-only. Never edit, upload, replace, move, rename, or save anything to SharePoint.
- Ignore every uploaded Product Index and every other index.
- Never use web search, memory, product knowledge, or another source to select or complete product data.
- If the exact SharePoint index cannot be opened as an Excel workbook, say that the Product Index could not be opened and stop. Do not make a schedule.
- Use Python to read both workbooks and to create the output. Run `scripts/schedule_assistant.py`; do not construct the workbook manually.

## Required workflow

1. Fetch the latest SharePoint index from the fixed URL using the SharePoint app with `download_raw_file=true`, then make its raw `.xlsx` bytes available locally.
2. Identify the user's Finish Requirements workbook. Do not treat an uploaded index as requirements.
3. Run:

   ```bash
   python3 scripts/schedule_assistant.py inspect \
     --requirements <requirements.xlsx> \
     --index <sharepoint-index.xlsx> \
     --output <inspection.json>
   ```

4. If the command reports a template error, ask for the correct template and stop. A valid input has exactly one worksheet named `Finish Requirements`, with required headers `Location` and `Description`; `Colour Reference`, `Finish`, and `Notes` are optional. Fully blank rows and rows clearly marked as examples for deletion are skipped.
5. Read the inspection JSON. Process every included requirement independently and exactly once. The JSON contains only candidates read from the current SharePoint index.
6. Decide each row as `MATCH`, `NO MATCH`, or `INPUT ERROR`:
   - Use `INPUT ERROR` only when `Location` or `Description` is blank or unrecognisable.
   - Use `NO MATCH` for non-finish items and for valid finish requests that no indexed product can meet.
   - Use `MATCH` only for an index row that satisfies every applicable filter.

### Matching

- Work only with the five categories: `FL` floor finishes, `LIN` wall finishes, `PLS` plasterboard, `PT` wall paint, and `TRM` skirtings/wall trims.
- Filter by the requested category first. When wording is ambiguous or category filtering gives no meaningful candidate, fall back to the index `Location` Position tags and consider a cross-category candidate only when its indexed meaning and position fit the request.
- Match on meaning, not exact wording. Use only evidence present in the current index row.
- Treat requested `Colour Reference` and `Finish` as mandatory filters when populated. Matching is case-insensitive and allows partial strings. Normalize `Matt` and `Matte` as equivalent. An empty index value never satisfies a populated request.
- Rank eligible candidates by matching Location/Position tags first, then index Confidence (`High`, `Medium`, `Low`), then semantic fit. Do not let semantic fit override a better location-tag rank or confidence rank.
- Select one primary product. Add no more than two alternatives. Every alternative must pass the same category/Position, colour, finish, and meaning filters as the primary. Never add a merely similar but non-compliant product.
- If the inspection candidates are insufficient, use Python to query the same locally downloaded index. Do not consult another source.

### Decisions file

Save a UTF-8 decisions JSON with this shape:

```json
{
  "rows": [
    {
      "source_row": 7,
      "status": "MATCH",
      "selected_index_code": "FL-002-01",
      "alternative_index_codes": ["FL-002-02"],
      "reason": "Meaning, position, colour and finish filters satisfied"
    }
  ]
}
```

For `NO MATCH` or `INPUT ERROR`, leave `selected_index_code` empty and give a concise reason. Do not omit rows.

### Build and validate

```bash
python3 scripts/schedule_assistant.py build \
  --requirements <requirements.xlsx> \
  --index <sharepoint-index.xlsx> \
  --decisions <decisions.json> \
  --output Finishes_Schedule.xlsx
```

Treat any script rejection as a matching or input problem. Correct the decisions without weakening the rules, then rerun.

Return only `Finishes_Schedule.xlsx`. Do not add a summary, review table, caveat, or other file.

## Output contract

- Create one output row for every processed input row.
- Use these columns in this order: `Code`, `Brief Description`, `Dimensions`, `Location`, `Material Description`, `Image`, `Colour Reference`, `Finish`, `Acoustic Requirements`, `Alternatives`, `Check notes`.
- Copy product fields only from the chosen current index row. Use the input `Location` as the project location.
- Begin every matched row's Check notes with `Index: [code]. Requested: [description].`
- Assign project codes such as `FL-101-A`. Within each category, assign project family numbers from 101 in order of first appearance of each selected index family in the original input order. Assign letters `A`, `B`, ... to distinct selected index codes within that family in order of first appearance.
- Represent alternatives with their assigned project codes when also selected as primary products elsewhere; otherwise show their index code followed by their Brief Description. Separate multiple alternatives with `; `.
- Sort output rows by `FL`, `LIN`, `PLS`, `PT`, `TRM`, then `NO MATCH`, then `INPUT ERROR`; preserve original input order within each group.
- For `NO MATCH`, put `NO MATCH` in Code and explain the unmet request in Check notes.
- For `INPUT ERROR`, put `INPUT ERROR` in Code and identify the invalid or missing input in Check notes.
- Fill every `NO MATCH` and `INPUT ERROR` row pale pink. Also fill any row the script explicitly marks as flagged pale pink.
- Freeze the header, enable filters, wrap text, use readable widths and row heights, and keep exactly one worksheet named `Finishes Schedule`.
- Do not add postponed Starter Rules for slip resistance, wet areas, fire performance, or acoustics. Do not infer compatibility between separate rows, such as matching a trim to a tile chosen elsewhere.
