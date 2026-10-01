---
name: consultant-door-sync
description: Overlays a consultant's door plan (PDF or DXF) on the project door plan, matches each consultant marker to the correct Revit door, and writes the consultant data into an Ideate BIMLink door export ready to import back into Revit.
---

# Consultant Door Sync

Create exactly one deliverable: `<export name>_BIMLink.xlsx`, an updated copy of the user's Ideate BIMLink door export. Also show the overlay image so the user can check the matching.

## Non-negotiable source rules

- Use only the three files the user provides:
  1. **BIMLink door export** (`.xlsx`) from Ideate BIMLink, with an element ID column (`Id` or `Element ID`), a `Mark` column, and the target parameter columns already included in the link.
  2. **Project door plan** (`.pdf`) exported from Revit, with door tags showing each door's Mark as text.
  3. **Consultant plan** (`.pdf` or `.dxf`) with the consultant's markers and a legend.
- If the consultant plan is `.dwg`, ask the user to save it as PDF or DXF and stop. Do not try to convert it.
- Never invent requirements. Every value written must come from the consultant's legend or notes, applied to a door the consultant marked.
- Never change the element ID column, the Mark column, the header row, the sheet name, or the column order. Never add or delete columns, rows or sheets; BIMLink import depends on the structure staying the same.
- Leave doors the consultant did not mark exactly as they are. Do not clear existing values.
- Use Python for all reading, aligning, matching and writing. Run `scripts/consultant_door_sync.py`; do not edit the workbook manually.

## Required workflow

1. Identify the three inputs. If any is missing, ask for it and stop. If the plan PDF has several sheets, ask which page holds the relevant level.
2. Run:

   ```bash
   python3 scripts/consultant_door_sync.py inspect \
     --export <bimlink_export.xlsx> \
     --plan <project_door_plan.pdf> \
     --consultant <consultant_plan.pdf|dxf> \
     --output inspection.json \
     --overlay overlay.png
   ```

   Add `--plan-page N` / `--consultant-page N` for multi-page PDFs.
3. If the command reports an error, explain it in one sentence and ask for the corrected input. If it reports too few shared labels, follow **Manual alignment** below.
4. Open `overlay.png` and check the alignment. The consultant's linework should sit on the project plan's walls and grid. If it does not, use **Manual alignment**.
5. Read `inspection.json` and the consultant plan itself, then write the decisions file (see **Decisions**).
6. Run:

   ```bash
   python3 scripts/consultant_door_sync.py build \
     --export <bimlink_export.xlsx> \
     --inspection inspection.json \
     --decisions decisions.json \
     --output <export name>_BIMLink.xlsx
   ```

7. If the script rejects the decisions, fix the decisions without weakening the rules and rerun.

Return the workbook and `overlay.png`, then list any conflicts and unresolved markers, one line each. Do not add other files or a long summary.

## How matching works

- **Alignment:** the script finds labels that appear exactly once in both drawings, such as grid bubbles, room or zone names and level titles, and fits a scale, rotation and offset. Labels that disagree are rejected automatically. Check `alignment.rms_error`: anything above about 2% of the drawing width means the overlay is unreliable.
- **Markers:** coloured symbols in a PDF, blocks, circles and hatches in a DXF (grouped by layer, block and colour), or, if there are no symbols, the consultant's own text tags (grouped by letter prefix). Each **marker group** is one legend category.
- **Assignment:** each marker snaps to the nearest door tag on the project plan within `tolerance`.
  - `ASSIGNED`: one clear nearest door.
  - `AMBIGUOUS`: two door tags are similarly close, for example a double door pair.
  - `UNMATCHED`: no door tag nearby, usually a legend swatch, note or title-block symbol.

### Manual alignment

When labels are insufficient, pick at least three points that are clearly the same in both drawings, such as grid intersections or building corners, far apart and not in a line. Read their coordinates from the PDFs with Python (PyMuPDF words or drawings) or the DXF. Save them as:

```json
{"pairs": [
  {"name": "A/1", "consultant": [120.5, 88.0], "plan": [150.0, 120.0]},
  {"name": "E/1", "consultant": [701.3, 88.2], "plan": [1040.0, 120.0]},
  {"name": "A/4", "consultant": [120.4, 520.1], "plan": [150.0, 720.0]}
]}
```

Rerun `inspect` with `--points points.json`. Use PDF points for PDFs and drawing units for DXF.

## Decisions

Save a UTF-8 `decisions.json`:

```json
{
  "legend": {
    "colour #D91A1A": {"Security Level": "SL3", "Access Control": "Card reader"}
  },
  "ignore_markers": ["M11"],
  "overrides": [{"marker": "M10", "mark": "D-111", "reason": "Symbol sits on the D-111 leaf"}],
  "doors": [{"mark": "D-112", "values": {"Security Level": "SL3"}, "reason": "Second leaf of pair"}]
}
```

- **legend:** map every marker group that is not ignored to parameter values. Read the meaning from the consultant's legend; `possible_legend_text` in the inspection is a hint, not a source. Parameter names must exactly match columns in the export's `writable_columns`. If the consultant's data needs a parameter that is not in the export, tell the user to add it to the BIMLink link and re-export, then stop.
- **ignore_markers:** legend swatches, north arrows, title-block symbols and anything else that is not a door requirement. Never ignore a marker just to make the build pass.
- **overrides:** resolve `AMBIGUOUS` markers, or correct a wrong assignment visible in the overlay, only when the drawing makes the intended door clear. Otherwise leave the marker unresolved so the build flags it.
- **doors:** only for requirements the consultant states for a specific door that are not carried by a marker, such as a note, a schedule on the drawing, or the second leaf of a pair covered by one symbol. Give the reason.

## Output contract

- The workbook is the original export with consultant values written into the mapped parameter columns of the matched door rows. Nothing else changes.
- Conflicting markers on one door (different values for the same parameter) are not written; that door's row is filled pale pink.
- Doors near an unresolved (`AMBIGUOUS`) marker are filled pale pink for review. Pink fill does not affect BIMLink import.
- Tell the user to import the workbook through Ideate BIMLink using the same link, then spot-check the pink rows in Revit.
