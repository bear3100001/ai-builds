#!/usr/bin/env python3
"""Overlay a consultant plan on the project door plan, match consultant markers
to Revit doors, and write the results into an Ideate BIMLink door export."""
from __future__ import annotations

import argparse
import colorsys
import json
import math
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

ID_HEADERS = {"id", "element id", "elementid", "element_id"}
MARK_HEADERS = {"mark", "door mark", "door number", "number"}
PINK = "F4CCCC"
TAG_PATTERN = re.compile(r"^(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9._/-]{2,10}$")


class InputProblem(ValueError):
    pass


def text(value) -> str:
    return "" if value is None else str(value).strip()


def key(value) -> str:
    return re.sub(r"\s+", " ", text(value)).casefold()


# ---------------------------------------------------------------- BIMLink export

def load_export(path: Path, sheet: str | None):
    wb = load_workbook(path)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    for row_no in range(1, min(ws.max_row, 20) + 1):
        headers = {key(ws.cell(row_no, c).value): c for c in range(1, ws.max_column + 1)
                   if text(ws.cell(row_no, c).value)}
        id_col = next((headers[h] for h in ID_HEADERS if h in headers), None)
        mark_col = next((headers[h] for h in MARK_HEADERS if h in headers), None)
        if id_col and mark_col:
            break
    else:
        raise InputProblem("BIMLink export needs an element ID column (Id / Element ID) and a Mark column in the first 20 rows.")
    names = {c: text(ws.cell(row_no, c).value) for c in range(1, ws.max_column + 1) if text(ws.cell(row_no, c).value)}
    doors, seen = [], Counter()
    for r in range(row_no + 1, ws.max_row + 1):
        eid, mark = text(ws.cell(r, id_col).value), text(ws.cell(r, mark_col).value)
        if not eid and not mark:
            continue
        if not eid:
            raise InputProblem(f"Row {r} has Mark '{mark}' but no element ID.")
        doors.append({"row": r, "element_id": eid, "mark": mark})
        seen[key(mark)] += 1
    if not doors:
        raise InputProblem("BIMLink export contains no door rows.")
    duplicates = sorted(m for m, n in seen.items() if m and n > 1)
    return {
        "workbook": wb, "sheet": ws.title, "header_row": row_no, "id_column": names[id_col],
        "mark_column": names[mark_col], "columns": names, "doors": doors, "duplicate_marks": duplicates,
    }


# ---------------------------------------------------------------- drawings

def chromatic(rgb) -> str | None:
    if not rgb:
        return None
    h, l, s = colorsys.rgb_to_hls(*rgb[:3])
    if s < 0.3 or l < 0.12 or l > 0.92:
        return None
    return "#%02X%02X%02X" % tuple(round(v * 255) for v in rgb[:3])


def load_pdf(path: Path, page_no: int):
    import pymupdf as fitz
    doc = fitz.open(path)
    if not 1 <= page_no <= len(doc):
        raise InputProblem(f"{path.name} has {len(doc)} page(s); page {page_no} does not exist.")
    page = doc[page_no - 1]
    words = [{"text": w[4], "x": (w[0] + w[2]) / 2, "y": (w[1] + w[3]) / 2, "h": w[3] - w[1]}
             for w in page.get_text("words")]
    diag = math.hypot(page.rect.width, page.rect.height)
    shapes = []
    for d in page.get_drawings():
        colour = chromatic(d.get("fill")) or chromatic(d.get("color"))
        r = d["rect"]
        if not colour or max(r.width, r.height) > diag * 0.04 or max(r.width, r.height) <= 0:
            continue
        shapes.append({"group": f"colour {colour}", "colour": colour, "x": (r.x0 + r.x1) / 2,
                       "y": (r.y0 + r.y1) / 2, "size": max(r.width, r.height)})
    return {"kind": "pdf", "path": path, "page": page_no, "words": words, "shapes": shapes,
            "width": page.rect.width, "height": page.rect.height}


ACI_HEX = {1: "#FF0000", 2: "#FFFF00", 3: "#00FF00", 4: "#00FFFF", 5: "#0000FF", 6: "#FF00FF"}


def load_dxf(path: Path):
    import ezdxf
    from ezdxf import bbox as ezbbox
    doc = ezdxf.readfile(path)
    words, shapes = [], []
    for e in doc.modelspace():
        t = e.dxftype()
        if t in {"TEXT", "MTEXT"}:
            content = e.plain_text() if t == "MTEXT" else e.dxf.text
            p = e.dxf.insert
            height = float((e.dxf.get("char_height") if t == "MTEXT" else e.dxf.get("height")) or 1)
            for part in text(content).split():
                words.append({"text": part, "x": p.x, "y": -p.y, "h": height})
        elif t in {"INSERT", "CIRCLE", "HATCH", "LWPOLYLINE", "POINT"}:
            try:
                box = ezbbox.extents([e])
                cx, cy = box.center.x, box.center.y
                size = max(box.size.x, box.size.y)
            except Exception:
                continue
            colour = ACI_HEX.get(e.dxf.get("color", 256))
            label = f"block {e.dxf.name}" if t == "INSERT" else f"{t.lower()}"
            group = f"layer {e.dxf.layer} | {label}" + (f" | {colour}" if colour else "")
            shapes.append({"group": group, "colour": colour, "x": cx, "y": -cy, "size": size})
    if not words and not shapes:
        raise InputProblem("DXF model space has no text or markers.")
    return {"kind": "dxf", "path": path, "words": words, "shapes": shapes}


def load_consultant(path: Path, page_no: int):
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return load_pdf(path, page_no)
    if suffix == ".dxf":
        return load_dxf(path)
    if suffix == ".dwg":
        raise InputProblem("DWG cannot be read directly. Save it as PDF or DXF and upload that.")
    raise InputProblem(f"Unsupported consultant plan type: {path.suffix}")


def merge_shapes(shapes):
    """Combine overlapping parts of one symbol (fill + outline + hatch) into one marker."""
    merged = []
    for s in sorted(shapes, key=lambda s: -s["size"]):
        for m in merged:
            if m["group"] == s["group"] and math.hypot(m["x"] - s["x"], m["y"] - s["y"]) <= max(m["size"], s["size"]) * 0.6:
                break
        else:
            merged.append(dict(s))
    return merged


# ---------------------------------------------------------------- alignment

def unique_labels(words):
    counts = Counter(key(w["text"]) for w in words)
    return {key(w["text"]): (w["x"], w["y"]) for w in words if counts[key(w["text"])] == 1 and len(key(w["text"])) <= 20}


def similarity(src, dst):
    import cv2
    src, dst = np.float32(src), np.float32(dst)
    if len(src) < 2:
        raise InputProblem("Alignment needs at least two reference points.")
    diag = float(np.linalg.norm(dst.max(0) - dst.min(0))) or 1.0
    matrix, inliers = cv2.estimateAffinePartial2D(src, dst, method=cv2.RANSAC, ransacReprojThreshold=diag * 0.01)
    if matrix is None:
        raise InputProblem("Could not compute an alignment from the reference points.")
    inliers = inliers.ravel().astype(bool)
    moved = src @ matrix[:, :2].T + matrix[:, 2]
    rms = float(np.sqrt(np.mean(np.sum((moved[inliers] - dst[inliers]) ** 2, axis=1))))
    return matrix, inliers, rms


def align(plan, consultant, points_file: Path | None):
    if points_file:
        pairs = json.loads(points_file.read_text(encoding="utf-8"))["pairs"]
        src = [p["consultant"] for p in pairs]
        if consultant["kind"] == "dxf":
            src = [(x, -y) for x, y in src]
        dst = [p["plan"] for p in pairs]
        names = [p.get("name", f"point {i + 1}") for i, p in enumerate(pairs)]
        method = "manual points"
    else:
        a, b = unique_labels(plan["words"]), unique_labels(consultant["words"])
        names = sorted(set(a) & set(b))
        if len(names) < 3:
            raise InputProblem(
                f"Only {len(names)} label(s) appear once in both drawings ({', '.join(names) or 'none'}). "
                "Supply --points with at least three matching reference points.")
        src, dst = [b[n] for n in names], [a[n] for n in names]
        method = "shared labels"
    matrix, inliers, rms = similarity(src, dst)
    scale = float(math.hypot(matrix[0, 0], matrix[1, 0]))
    rotation = float(math.degrees(math.atan2(matrix[1, 0], matrix[0, 0])))
    used = [n for n, ok in zip(names, inliers) if ok]
    rejected = [n for n, ok in zip(names, inliers) if not ok]
    if len(used) < 2:
        raise InputProblem("Reference points disagree; alignment is unreliable. Supply better --points.")
    return matrix, {"method": method, "labels_used": used, "labels_rejected": rejected,
                    "scale": round(scale, 6), "rotation_deg": round(rotation, 3), "rms_error": round(rms, 3)}


def transform(matrix, x, y):
    return (float(matrix[0, 0] * x + matrix[0, 1] * y + matrix[0, 2]),
            float(matrix[1, 0] * x + matrix[1, 1] * y + matrix[1, 2]))


# ---------------------------------------------------------------- matching

def door_tags(plan, export):
    by_mark = {key(d["mark"]): d for d in export["doors"] if d["mark"]}
    tags = defaultdict(list)
    for w in plan["words"]:
        if key(w["text"]) in by_mark:
            tags[key(w["text"])].append((w["x"], w["y"]))
    return by_mark, tags


def inspect_command(args):
    export = load_export(Path(args.export), args.sheet)
    plan = load_pdf(Path(args.plan), args.plan_page)
    consultant = load_consultant(Path(args.consultant), args.consultant_page)
    matrix, alignment = align(plan, consultant, Path(args.points) if args.points else None)
    by_mark, tags = door_tags(plan, export)
    if not tags:
        raise InputProblem("None of the export's door Marks were found as text on the project plan PDF.")
    tag_points = [(m, x, y) for m, pts in tags.items() for x, y in pts]

    shared = set(unique_labels(plan["words"])) & set(unique_labels(consultant["words"]))
    markers = merge_shapes(consultant["shapes"])
    if not markers:
        # No coloured symbols or CAD blocks: fall back to the consultant's own text tags (e.g. S01, AC-3),
        # grouped by their letter prefix so the legend can map each tag type.
        for w in consultant["words"]:
            k = key(w["text"])
            if TAG_PATTERN.match(w["text"]) and k not in by_mark and k not in shared:
                prefix = re.match(r"[A-Za-z]+", w["text"])
                markers.append({"group": f"tag {prefix.group(0).upper() if prefix else w['text']}", "colour": None,
                                "x": w["x"], "y": w["y"], "size": w["h"]})

    pts = np.array([(x, y) for _, x, y in tag_points])
    if len(pts) > 1:
        d = np.sqrt(((pts[:, None] - pts[None]) ** 2).sum(-1))
        np.fill_diagonal(d, np.inf)
        spacing = float(np.median(d.min(1)))
    else:
        spacing = 50.0
    tolerance = args.tolerance or spacing * 0.6

    out_markers = []
    for i, m in enumerate(markers, 1):
        x, y = transform(matrix, m["x"], m["y"])
        rec = {"id": f"M{i}", "group": m["group"], "colour": m["colour"], "plan_xy": [round(x, 2), round(y, 2)]}
        ranked = sorted(((math.hypot(tx - x, ty - y), mk) for mk, tx, ty in tag_points))
        best = {}
        for dist, mk in ranked:
            best.setdefault(mk, dist)
        near = sorted((dist, mk) for mk, dist in best.items() if dist <= tolerance * 2)
        rec["candidates"] = [{"mark": by_mark[mk]["mark"], "distance": round(dist, 2)} for dist, mk in near[:3]]
        # Door tags sit beside doors, not on them, so a rival tag up to twice as far away still competes.
        if not near or near[0][0] > tolerance:
            rec.update(status="UNMATCHED")
        elif len(near) > 1 and near[1][0] <= near[0][0] * 2:
            rec.update(status="AMBIGUOUS")
        else:
            mk = near[0][1]
            rec.update(status="ASSIGNED", method="nearest door tag", mark=by_mark[mk]["mark"],
                       element_id=by_mark[mk]["element_id"], distance=round(near[0][0], 2))
        out_markers.append(rec)

    groups = Counter(m["group"] for m in out_markers)
    group_info = []
    for g, n in sorted(groups.items()):
        unmatched = [m for m in out_markers if m["group"] == g and m["status"] == "UNMATCHED"]
        group_info.append({"group": g, "count": n, "unmatched": len(unmatched),
                           "possible_legend_text": legend_text(consultant, markers, g)})

    per_door = defaultdict(set)
    for m in out_markers:
        if m["status"] == "ASSIGNED":
            per_door[m["mark"]].add(m["group"])
    result = {
        "export": {"file": Path(args.export).name, "sheet": export["sheet"], "id_column": export["id_column"],
                   "mark_column": export["mark_column"], "door_count": len(export["doors"]),
                   "writable_columns": [c for c in export["columns"].values()
                                        if c not in (export["id_column"], export["mark_column"])],
                   "duplicate_marks": export["duplicate_marks"]},
        "plan": {"file": Path(args.plan).name, "page": args.plan_page, "door_tags_found": len(tags),
                 "marks_not_on_plan": sorted(d["mark"] for d in export["doors"] if d["mark"] and key(d["mark"]) not in tags),
                 "marks_tagged_more_than_once": sorted(by_mark[m]["mark"] for m, p in tags.items() if len(p) > 1)},
        "consultant": {"file": Path(args.consultant).name, "kind": consultant["kind"]},
        "alignment": alignment,
        "tolerance": round(tolerance, 2),
        "marker_groups": group_info,
        "markers": out_markers,
        "doors_with_several_groups": {mk: sorted(g) for mk, g in per_door.items() if len(g) > 1},
        "summary": dict(Counter(m["status"] for m in out_markers)),
    }
    Path(args.output).write_text(json.dumps(result, indent=2), encoding="utf-8")
    if args.overlay:
        draw_overlay(plan, consultant, matrix, out_markers, tag_points, Path(args.overlay))
    print(json.dumps({"ok": True, "alignment": alignment, "summary": result["summary"],
                      "groups": [g["group"] for g in group_info]}, indent=2))


def legend_text(consultant, markers, group):
    """Words on the same line just right of an unmatched marker of this group (legend swatches)."""
    found = []
    for m in markers:
        if m["group"] != group or m.get("label"):
            continue
        line = sorted((w["x"], w["text"]) for w in consultant["words"]
                      if abs(w["y"] - m["y"]) <= max(m["size"], w["h"]) * 0.8 and 0 < w["x"] - m["x"] <= m["size"] * 40)
        if line:
            found.append(" ".join(t for _, t in line[:12]))
    return sorted(set(found))[:3]


def draw_overlay(plan, consultant, matrix, markers, tag_points, out: Path):
    import cv2
    import pymupdf as fitz
    zoom = 2.0
    page = fitz.open(plan["path"])[plan["page"] - 1]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    base = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3].copy()
    base = (255 - (255 - base) * 0.45).astype(np.uint8)
    if consultant["kind"] == "pdf":
        cpage = fitz.open(consultant["path"])[consultant["page"] - 1]
        cpix = cpage.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        cimg = np.frombuffer(cpix.samples, dtype=np.uint8).reshape(cpix.height, cpix.width, cpix.n)[:, :, :3]
        m = matrix.copy()
        m[:, 2] *= zoom
        warped = cv2.warpAffine(cimg, m, (pix.width, pix.height), borderValue=(255, 255, 255))
        ink = (warped.min(2) < 230)[:, :, None]
        base = np.where(ink, (base * 0.35 + warped * 0.65).astype(np.uint8), base)
    img = cv2.cvtColor(base, cv2.COLOR_RGB2BGR)
    tags = {}
    for mk, x, y in tag_points:
        tags.setdefault(mk, (x, y))
    for m in markers:
        x, y = (int(v * zoom) for v in m["plan_xy"])
        if m["status"] == "ASSIGNED":
            tx, ty = tags.get(key(m["mark"]), m["plan_xy"])
            cv2.line(img, (x, y), (int(tx * zoom), int(ty * zoom)), (60, 160, 60), 2)
            cv2.circle(img, (x, y), 9, (60, 160, 60), 2)
        else:
            cv2.drawMarker(img, (x, y), (0, 0, 220), cv2.MARKER_TILTED_CROSS, 22, 3)
        cv2.putText(img, m["id"], (x + 10, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (40, 40, 40), 1, cv2.LINE_AA)
    cv2.imwrite(str(out), img)


# ---------------------------------------------------------------- build

def build_command(args):
    export = load_export(Path(args.export), args.sheet)
    inspection = json.loads(Path(args.inspection).read_text(encoding="utf-8"))
    decisions = json.loads(Path(args.decisions).read_text(encoding="utf-8"))
    legend = decisions.get("legend", {})
    ignore = set(decisions.get("ignore_markers", []))
    overrides = {o["marker"]: o["mark"] for o in decisions.get("overrides", [])}
    columns = {v: c for c, v in export["columns"].items()}
    protected = {export["id_column"], export["mark_column"]}
    by_mark = {key(d["mark"]): d for d in export["doors"] if d["mark"]}
    markers = {m["id"]: m for m in inspection["markers"]}
    problems = []

    for group, values in legend.items():
        for param in values:
            if param not in columns:
                problems.append(f"Legend parameter '{param}' is not a column in the BIMLink export. Add it to the BIMLink link and re-export.")
            elif param in protected:
                problems.append(f"Legend may not write to '{param}'.")
    for mid in list(ignore) + list(overrides):
        if mid not in markers:
            problems.append(f"Marker {mid} does not exist in the inspection file.")
    for mid, mark in overrides.items():
        if key(mark) not in by_mark:
            problems.append(f"Override for {mid} names door '{mark}', which is not in the export.")
    used_groups = {m["group"] for mid, m in markers.items() if mid not in ignore}
    for g in sorted(used_groups - set(legend)):
        problems.append(f"Marker group '{g}' has no legend entry. Map it, or list its markers in ignore_markers.")
    if problems:
        raise InputProblem("\n".join(problems))

    per_door, unresolved = defaultdict(list), []
    for mid, m in markers.items():
        if mid in ignore:
            continue
        mark = overrides.get(mid) or (m.get("mark") if m["status"] == "ASSIGNED" else None)
        if not mark:
            unresolved.append({"marker": mid, "status": m["status"], "group": m["group"], "candidates": m.get("candidates", [])})
            continue
        per_door[key(mark)].append(m["group"])

    for d in decisions.get("doors", []):
        if key(d["mark"]) not in by_mark:
            raise InputProblem(f"Door decision names '{d['mark']}', which is not in the export.")

    ws = export["workbook"][export["sheet"]]
    pink = PatternFill("solid", start_color=PINK, end_color=PINK)
    written, conflicts, changed = [], [], 0
    explicit = {key(d["mark"]): d for d in decisions.get("doors", [])}
    for mk in sorted(set(per_door) | set(explicit)):
        door = by_mark[mk]
        values, clash = {}, set()
        for g in per_door.get(mk, []):
            for param, value in legend[g].items():
                if param in values and values[param] != value:
                    clash.add(param)
                values.setdefault(param, value)
        if mk in explicit:
            values.update(explicit[mk].get("values", {}))
            clash -= set(explicit[mk].get("values", {}))
        for param in values:
            if param not in columns or param in protected:
                raise InputProblem(f"Door '{door['mark']}' writes unknown or protected column '{param}'.")
        if clash:
            conflicts.append({"mark": door["mark"], "parameters": sorted(clash), "groups": sorted(set(per_door[mk]))})
            for c in range(1, ws.max_column + 1):
                ws.cell(door["row"], c).fill = pink
            continue
        for param, value in values.items():
            cell = ws.cell(door["row"], columns[param])
            if text(cell.value) != text(value):
                changed += 1
            cell.value = value
        written.append(door["mark"])

    for u in unresolved:
        for cand in u["candidates"]:
            row = by_mark[key(cand["mark"])]["row"]
            for c in range(1, ws.max_column + 1):
                ws.cell(row, c).fill = pink

    export["workbook"].save(args.output)
    print(json.dumps({"ok": True, "output": args.output, "doors_written": len(written), "cells_changed": changed,
                      "conflicts": conflicts, "unresolved_markers": unresolved}, indent=2))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("inspect", help="align drawings and match consultant markers to doors")
    p.add_argument("--export", required=True, help="Ideate BIMLink door export (.xlsx)")
    p.add_argument("--sheet", help="worksheet name (default: first sheet)")
    p.add_argument("--plan", required=True, help="project door plan PDF exported from Revit")
    p.add_argument("--plan-page", type=int, default=1)
    p.add_argument("--consultant", required=True, help="consultant plan (.pdf or .dxf)")
    p.add_argument("--consultant-page", type=int, default=1)
    p.add_argument("--points", help="JSON of manual reference point pairs, used instead of shared labels")
    p.add_argument("--tolerance", type=float, help="max marker-to-door-tag distance in plan units")
    p.add_argument("--output", required=True, help="inspection JSON to write")
    p.add_argument("--overlay", help="overlay PNG to write")
    p.set_defaults(func=inspect_command)

    b = sub.add_parser("build", help="write matched consultant data into the BIMLink export")
    b.add_argument("--export", required=True)
    b.add_argument("--sheet")
    b.add_argument("--inspection", required=True)
    b.add_argument("--decisions", required=True)
    b.add_argument("--output", required=True)
    b.set_defaults(func=build_command)

    args = parser.parse_args()
    try:
        args.func(args)
    except InputProblem as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, indent=2))
        sys.exit(2)


if __name__ == "__main__":
    main()
