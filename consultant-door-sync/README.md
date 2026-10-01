# Consultant Door Sync

An OpenAI skill that overlays a consultant's door plan on the project's Revit door plan, matches each consultant marker to the correct door, and writes the consultant data into an **Ideate BIMLink** door export, ready to import back into Revit.

Built for security consultant data on a stadium, where the doors between quads need specific security requirements. It works for any consultant who marks doors on a plan, such as fire, access or acoustic consultants.

**Platforms:** ChatGPT · Codex · API · Atlas
**Status:** Beta · **Version:** 0.1.0

## How it works

1. **Overlay:** aligns the consultant plan with the project plan using labels both drawings share, such as grid bubbles and zone names. If too few labels match, it uses reference points instead.
2. **Match:** snaps each consultant marker (coloured symbol, CAD block or tag) to the nearest door tag. Markers near two doors, or near none, are flagged for review.
3. **Legend:** reads the consultant's legend and maps each marker type to door parameters, for example red = `Security Level: SL3`, `Access Control: Card reader`.
4. **Write back:** fills those values into a copy of the BIMLink export. Only the mapped columns of matched doors change, and rows needing review are highlighted pink.

## Inputs

| Input | Format | Notes |
|---|---|---|
| BIMLink door export | `.xlsx` | Must include `Id`, `Mark`, and the parameters to fill. Add those parameters to the BIMLink link before exporting. |
| Project door plan | `.pdf` | Vector PDF exported from Revit, with door tags visible. |
| Consultant plan | `.pdf` or `.dxf` | Save DWG files as PDF or DXF first. |

## Output

- `<export name>_BIMLink.xlsx`: the original export with consultant values filled in. IDs, headers, columns and sheets are unchanged, so it imports cleanly through Ideate BIMLink.
- An overlay image for checking the matches visually.

## Rules

- Values come only from the consultant's legend or notes. Nothing is inferred.
- Doors the consultant didn't mark are left as they are.
- Conflicting or ambiguous markers are flagged, never guessed.

## Contents

| Path | Purpose |
|---|---|
| `SKILL.md` | Full instructions the model follows |
| `scripts/consultant_door_sync.py` | `inspect` (align and match) and `build` (write the BIMLink file) |
| `agents/openai.yaml` | Display name, prompt, icon and product availability |
| `assets/icon.svg` | Skill icon |

## Install

Copy this folder into your skills directory, or upload it as a skill in ChatGPT. No connectors are needed.

The script needs `openpyxl`, `pymupdf`, `opencv-python-headless`, `numpy` and, for DXF files, `ezdxf`. ChatGPT's code environment usually has these; to run it locally:

```bash
pip install openpyxl pymupdf opencv-python-headless numpy ezdxf
```

## Notes

- After importing into Revit, spot-check the pink rows and a few matched doors.
- Never commit real project drawings or exports to this repository.
