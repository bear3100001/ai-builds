# Schedule Assistant

An OpenAI skill that turns a **Finish Requirements** workbook into a coded **Finishes_Schedule.xlsx**, matching every row against the live Product Index on SharePoint.

**Platforms:** ChatGPT · Codex · API · Atlas
**Status:** Live · **Version:** 1.0.0

## How it works

1. Downloads the latest Product Index from SharePoint (read-only).
2. `schedule_assistant.py inspect` validates the requirements template and lists candidates.
3. The model decides `MATCH` / `NO MATCH` / `INPUT ERROR` per row using fixed filter and ranking rules.
4. `schedule_assistant.py build` validates those decisions and writes the schedule.

## Contents

| Path | Purpose |
|---|---|
| `SKILL.md` | Full instructions the model follows |
| `agents/openai.yaml` | Display name, prompt, icon and product availability |
| `scripts/schedule_assistant.py` | Inspect and build engine |
| `assets/icon.svg` | Skill icon |
| `examples/` | Sample input, output and screenshots (demo data only) |

## Try the demo

The `examples/` folder runs end to end without SharePoint:

| File | What it is |
|---|---|
| `Finish_Requirements_Demo.xlsx` | Sample input: 11 requests from a stadium project |
| `Product_Index_Demo.xlsx` | Small stand-in Product Index (5 products) |
| `decisions_demo.json` | The model's MATCH / NO MATCH decisions |
| `Finishes_Schedule_Demo.xlsx` | Resulting schedule |

```bash
pip install openpyxl
python3 scripts/schedule_assistant.py build \
  --requirements examples/Finish_Requirements_Demo.xlsx \
  --index examples/Product_Index_Demo.xlsx \
  --decisions examples/decisions_demo.json \
  --output Finishes_Schedule.xlsx
```

## Input template

One worksheet named `Finish Requirements`. `Location` and `Description` are required; `Colour Reference`, `Finish` and `Notes` are optional. No other columns are allowed.

## Install

Copy this folder into your skills directory, or upload it as a skill in ChatGPT. Requires the Microsoft SharePoint app connected, with access to the Product Index.

## Notes

- Before use, replace `<YOUR_SHAREPOINT_PRODUCT_INDEX_URL>` in `SKILL.md` with the link to your own Product Index.
- Only demo data belongs in `examples/`. Never commit real project workbooks.
