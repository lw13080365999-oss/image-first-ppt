#!/usr/bin/env python3
"""Static PPTX package/content/editability audit; this is not visual approval."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET

from project import project_path, read_json, sha256_file, validate_project, write_json

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
EMU = 914400
LIMITATIONS = [
    "Static XML checks cannot prove visual fidelity, glyph coverage, readable equations, or semantic/factual correctness.",
    "Render every final page and review it against the selected design and source; inspect the intended PowerPoint/WPS environment when available.",
    "Bounds cover translated/scaled groups and ordinary shape rotation. Rotated/flipped groups, shadows, clipping, effects and text-box internal overflow need rendered inspection.",
    "Text retention normalizes whitespace. Native-object presence does not establish full editability, table/chart semantics, equation equivalence or usability.",
    "This is not a full OOXML/XSD validator and does not prove that Office opens without a repair prompt or that embedded workbooks/media render correctly.",
]


def compact(value) -> str:
    return re.sub(r"\s+", "", str(value))


def _issue(items, code, message, **where):
    items.append({"code": code, "message": message, **where})


def _relationship_source(name: str) -> str:
    if name == "_rels/.rels":
        return ""
    folder, filename = posixpath.split(name)
    if not folder.endswith("/_rels"):
        raise ValueError("Unexpected relationship-part path")
    return posixpath.join(folder[:-6], filename[:-5])


def _target(source: str, value: str) -> str:
    parsed = urlsplit(value)
    if parsed.scheme or parsed.netloc or "\\" in value:
        raise ValueError("Internal relationship contains an external/Windows path")
    path = unquote(parsed.path)
    resolved = posixpath.normpath(path.lstrip("/") if path.startswith("/") else posixpath.join(posixpath.dirname(source), path))
    if resolved in {"", ".", ".."} or resolved.startswith("../"):
        raise ValueError("Relationship escapes package root")
    return resolved


def _xfrm(element):
    if element.tag == P + "graphicFrame":
        return element.find(P + "xfrm")
    if element.tag == P + "grpSp":
        return element.find(P + "grpSpPr/" + A + "xfrm")
    return element.find(P + "spPr/" + A + "xfrm")


def _pair(xfrm, name, first, second, default=(0, 0)):
    node = xfrm.find(A + name) if xfrm is not None else None
    if node is None:
        return default
    return float(node.get(first, default[0])), float(node.get(second, default[1]))


def _shape_info(tree, canvas, report, slide_index):
    """Read IDs/types and compute group translation/scale-aware bounds."""
    objects = []
    seen_ids = set()
    for metadata in tree.iter(P + "cNvPr"):
        shape_id = metadata.get("id")
        if shape_id in seen_ids:
            _issue(report["errors"], "duplicate_shape_id", "Shape/group IDs must be unique within a slide", slide=slide_index, shape_id=shape_id)
        seen_ids.add(shape_id)

    def walk(parent, transform=(0.0, 0.0, 1.0, 1.0), uncertain=False, inherited_hidden=False):
        tx, ty, sx, sy = transform
        for element in parent:
            if element.tag not in {P + "sp", P + "pic", P + "graphicFrame", P + "cxnSp", P + "grpSp"}:
                continue
            metadata = element.find(".//" + P + "cNvPr")
            hidden = inherited_hidden or (metadata is not None and metadata.get("hidden") in {"1", "true"})
            shape_id = metadata.get("id") if metadata is not None else None
            xfrm = _xfrm(element)
            x, y = _pair(xfrm, "off", "x", "y")
            width, height = _pair(xfrm, "ext", "cx", "cy")
            if element.tag == P + "grpSp":
                cx, cy = _pair(xfrm, "chOff", "x", "y")
                cw, ch = _pair(xfrm, "chExt", "cx", "cy", (width, height))
                rotation = float(xfrm.get("rot", "0")) if xfrm is not None else 0
                flip = xfrm is not None and any(xfrm.get(k) in {"1", "true"} for k in ("flipH", "flipV"))
                unclear = uncertain or rotation != 0 or flip or cw == 0 or ch == 0
                if unclear and not uncertain:
                    _issue(report["warnings"], "group_bounds_limited", "Group rotation/flip or incomplete group transform requires visual bounds review", slide=slide_index, shape_id=shape_id)
                gsx, gsy = width / cw if cw else 1, height / ch if ch else 1
                walk(element, (tx + sx * (x - cx * gsx), ty + sy * (y - cy * gsy), sx * gsx, sy * gsy), unclear, hidden)
                continue
            table = element.find(".//" + A + "tbl")
            chart = element.find(".//" + C + "chart")
            equations = element.findall(".//" + M + "oMath")
            text = "".join(node.text or "" for node in element.iter(A + "t"))
            math_text = "".join(node.text or "" for node in element.iter(M + "t"))
            kind = "image" if element.tag == P + "pic" else "shape"
            if table is not None:
                kind = "table"
            elif chart is not None:
                kind = "chart"
            elif equations:
                kind = "formula"
            elif text:
                kind = "text"
            elif element.tag == P + "cxnSp":
                kind = "connector"
            sizes = [int(node.get("sz")) for node in element.iter() if node.get("sz", "").isdigit() and node.tag in {A + "rPr", A + "defRPr", A + "endParaRPr"}]
            negligible_text = bool(sizes) and max(sizes) < 100
            if negligible_text and text:
                _issue(report["warnings"], "negligible_text", "Text is explicitly below 1 pt and does not count as useful native-text coverage", slide=slide_index, shape_id=shape_id)
            bounds = None
            if xfrm is not None and not uncertain:
                rotation = math.radians(float(xfrm.get("rot", "0")) / 60000)
                cosine, sine = math.cos(rotation), math.sin(rotation)
                midx, midy = x + width / 2, y + height / 2
                points = [(tx + sx * (midx + dx * cosine - dy * sine), ty + sy * (midy + dx * sine + dy * cosine))
                          for dx, dy in ((-width / 2, -height / 2), (-width / 2, height / 2), (width / 2, -height / 2), (width / 2, height / 2))]
                left, top = min(v[0] for v in points), min(v[1] for v in points)
                right, bottom = max(v[0] for v in points), max(v[1] for v in points)
                bounds = [left / EMU, top / EMU, (right - left) / EMU, (bottom - top) / EMU]
                tolerance = 0.02 * EMU
                if not hidden and (left < -tolerance or top < -tolerance or right > canvas[0] + tolerance or bottom > canvas[1] + tolerance):
                    _issue(report["errors"], "out_of_bounds", "Object exceeds the slide canvas", slide=slide_index, shape_id=shape_id, bbox_in=bounds)
            if xfrm is None:
                _issue(report["warnings"], "inherited_geometry", "Object geometry may be inherited from a layout/master; static bounds were not established", slide=slide_index, shape_id=shape_id)
            obj = {"id": shape_id, "kind": kind, "text": text, "math_text": math_text,
                   "formula_count": len(equations), "hidden": hidden, "negligible_text": negligible_text,
                   "bbox_in": bounds, "element": element, "table": table,
                   "chart_rid": chart.get(R + "id") if chart is not None else None}
            if shape_id is None:
                _issue(report["errors"], "missing_shape_id", "Slide object has no nonvisual shape ID", slide=slide_index)
            objects.append(obj)

    walk(tree)
    return objects


def _native_table_cells(element):
    return [["".join(node.text or "" for node in cell.iter(A + "t")) for cell in row.findall(A + "tc")]
            for row in element.findall(A + "tr")]


def _chart_series(tree):
    def values(node):
        if node is None:
            return []
        pts = list(node.iter(C + "pt"))
        pts.sort(key=lambda point: int(point.get("idx", "0")))
        return [point.findtext(C + "v", "") for point in pts]
    result = []
    for series in tree.iter(C + "ser"):
        val = series.find(C + "val")
        if val is None:
            val = series.find(C + "yVal")
        tx = series.find(C + "tx")
        name = ""
        if tx is not None:
            direct_name = tx.find(C + "v")
            cached_names = values(tx)
            if direct_name is not None:
                name = direct_name.text or ""
            elif cached_names:
                name = "".join(cached_names)
            else:
                name = "".join(node.text or "" for node in tx.iter(A + "t"))
        result.append({"name": name, "values": values(val), "categories": values(series.find(C + "cat"))})
    return result


def audit(pptx: Path, root: Path | None = None, expected_slides: int | None = None) -> dict:
    report = {"schema_version": 1, "pptx": str(pptx.resolve()), "success": False, "errors": [], "warnings": [],
              "slides": [], "limitations": LIMITATIONS, "visual_review_verified": False}
    deck = None
    if root is not None:
        root = root.resolve()
        planning = validate_project(root, require_maps=True)
        report["errors"].extend({**i, "scope": "project"} for i in planning["errors"])
        report["warnings"].extend({**i, "scope": "project"} for i in planning["warnings"])
        try:
            deck = read_json(root / "deck.json")
            report["deck_sha256"] = sha256_file(root / "deck.json")
            if isinstance(deck, dict) and isinstance(deck.get("slides"), list):
                expected_slides = len(deck["slides"])
                report["object_maps_sha256"] = {}
                for slide in deck["slides"]:
                    if isinstance(slide, dict) and isinstance(slide.get("id"), str):
                        try:
                            path = project_path(root, f"slides/{slide['id']}/objects.json")
                            if path.is_file():
                                report["object_maps_sha256"][slide["id"]] = sha256_file(path)
                        except ValueError:
                            pass
            else:
                deck = None
            if planning["errors"]:
                # Still inspect the package, but do not interpret malformed or
                # incomplete source/mapping records as an acceptance contract.
                deck = None
        except (OSError, ValueError):
            pass
    try:
        report["pptx_sha256"] = sha256_file(pptx)
        with zipfile.ZipFile(pptx) as archive:
            infos = archive.infolist()
            if len(infos) > 10000 or sum(info.file_size for info in infos) > 512 * 1024 * 1024 or any(info.file_size > 128 * 1024 * 1024 for info in infos):
                _issue(report["errors"], "package_limit", "Package exceeds static-audit resource limits")
                return report
            names = archive.namelist()
            if len(names) != len(set(names)):
                _issue(report["errors"], "duplicate_parts", "ZIP contains duplicate part names")
            for name in names:
                if name.startswith("/") or "\\" in name or ".." in name.split("/"):
                    _issue(report["errors"], "unsafe_zip_name", "Unsafe package part path", part=name)
            bad = archive.testzip()
            if bad:
                _issue(report["errors"], "zip_crc", "Corrupt ZIP member", part=bad)
            parts = set(names)
            for required in ("[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"):
                if required not in parts:
                    _issue(report["errors"], "missing_package_part", "Required package part is absent", part=required)
            xml = {}
            for name in names:
                if name.endswith((".xml", ".rels")):
                    try:
                        content = archive.read(name)
                        if b"<!DOCTYPE" in content.upper():
                            raise ValueError("DOCTYPE is not permitted in generated Office XML")
                        xml[name] = ET.fromstring(content)
                    except (ET.ParseError, ValueError) as exc:
                        _issue(report["errors"], "invalid_xml", str(exc), part=name)
            types = xml.get("[Content_Types].xml")
            if types is not None:
                defaults = {node.get("Extension", "").lower(): node.get("ContentType") for node in types if node.tag.endswith("}Default")}
                overrides = {node.get("PartName", "").lstrip("/"): node.get("ContentType") for node in types if node.tag.endswith("}Override")}
                for name in names:
                    if name.endswith("/") or name == "[Content_Types].xml":
                        continue
                    extension = name.rsplit(".", 1)[-1].lower()
                    if not overrides.get(name) and not defaults.get(extension):
                        _issue(report["errors"], "missing_content_type", "Package part has no content-type registration", part=name)
            relations = {}
            for name, tree in xml.items():
                if not name.endswith(".rels"):
                    continue
                try:
                    source = _relationship_source(name)
                except ValueError as exc:
                    _issue(report["errors"], "relationship_part", str(exc), part=name)
                    continue
                links = {}
                for rel in tree:
                    rid, value = rel.get("Id"), rel.get("Target")
                    if not rid or not value or rid in links:
                        _issue(report["errors"], "relationship_identity", "Missing/duplicate relationship ID or missing target", part=name)
                        continue
                    if rel.get("TargetMode", "").lower() == "external":
                        links[rid] = {"external": True, "target": value, "type": rel.get("Type", "")}
                        continue
                    try:
                        target = _target(source, value)
                    except ValueError as exc:
                        _issue(report["errors"], "unsafe_relationship", str(exc), part=name, relationship=rid)
                        continue
                    links[rid] = {"external": False, "target": target, "type": rel.get("Type", "")}
                    if target not in parts:
                        _issue(report["errors"], "missing_relationship_target", "Relationship points to a missing package resource", part=name, relationship=rid, target=target)
                relations[source] = links
            offices = [link for link in relations.get("", {}).values() if link["type"].endswith("/officeDocument")]
            if len(offices) != 1 or offices[0].get("external") or offices[0].get("target") != "ppt/presentation.xml":
                _issue(report["errors"], "office_document_relationship", "Package must declare one internal officeDocument relationship to its presentation")
            # A reference attribute can be broken even when every declared relationship resolves.
            for name, tree in xml.items():
                if name.endswith(".rels"):
                    continue
                links = relations.get(name, {})
                for element in tree.iter():
                    for key in (R + "id", R + "embed", R + "link"):
                        if key in element.attrib and element.get(key) not in links:
                            _issue(report["errors"], "undeclared_relationship", "XML references an undeclared relationship", part=name, relationship=element.get(key))
            presentation = xml.get("ppt/presentation.xml")
            if presentation is None:
                return report
            size = presentation.find(P + "sldSz")
            if size is None:
                _issue(report["errors"], "missing_canvas", "Presentation has no slide size")
                return report
            try:
                canvas = (float(size.get("cx", "0")), float(size.get("cy", "0")))
                if min(canvas) <= 0:
                    raise ValueError("Nonpositive slide size")
            except ValueError as exc:
                _issue(report["errors"], "invalid_canvas", str(exc))
                return report
            report["canvas_in"] = {"width": canvas[0] / EMU, "height": canvas[1] / EMU}
            if deck is not None:
                expected_canvas = deck.get("canvas", {})
                if isinstance(expected_canvas, dict) and all(isinstance(expected_canvas.get(k), (int, float)) for k in ("width_in", "height_in")):
                    if any(abs(expected_canvas[k] - canvas[i] / EMU) > 0.02 for i, k in enumerate(("width_in", "height_in"))):
                        _issue(report["errors"], "canvas_mismatch", "PPTX canvas differs from the planned size")
            pages = presentation.findall(P + "sldIdLst/" + P + "sldId")
            report["slide_count"] = len(pages)
            if not pages:
                _issue(report["errors"], "empty_deck", "Presentation contains no slides")
            if expected_slides is not None and len(pages) != expected_slides:
                _issue(report["errors"], "slide_count", "Slide count differs from the planned count", expected=expected_slides, actual=len(pages))
            for index, page in enumerate(pages, 1):
                link = relations.get("ppt/presentation.xml", {}).get(page.get(R + "id"), {})
                part = link.get("target")
                tree = xml.get(part)
                if tree is None or link.get("external"):
                    _issue(report["errors"], "slide_unreadable", "Presentation slide cannot be resolved", slide=index)
                    continue
                shape_tree = tree.find(P + "cSld/" + P + "spTree")
                if shape_tree is None:
                    _issue(report["errors"], "missing_shape_tree", "Slide has no shape tree", slide=index)
                    continue
                objects = _shape_info(shape_tree, canvas, report, index)
                counts = {key: sum(o["kind"] == key for o in objects) for key in ("text", "table", "chart", "formula", "image", "shape", "connector")}
                counts["math_expressions"] = sum(o["formula_count"] for o in objects)
                counts["visible_text"] = sum(o["kind"] == "text" and not o["hidden"] and not o["negligible_text"] for o in objects)
                info = {"index": index, "part": part, "counts": counts,
                        "objects": [{k: o[k] for k in ("id", "kind", "bbox_in", "hidden")} for o in objects]}
                report["slides"].append(info)
                for obj in objects:
                    if any(char in obj["text"] + obj["math_text"] for char in ("□", "�", "⬜")):
                        _issue(report["warnings"], "suspicious_glyph", "Text contains a box/replacement character; it may be intentional and needs inspection", slide=index, shape_id=obj["id"])
                full_images = [o for o in objects if o["kind"] == "image" and not o["hidden"] and o["bbox_in"] and
                               o["bbox_in"][2] * o["bbox_in"][3] >= (canvas[0] / EMU) * (canvas[1] / EMU) * 0.95]
                if full_images and not counts["visible_text"] and not any(counts[k] for k in ("table", "chart", "formula")):
                    _issue(report["errors"] if deck is not None and deck.get("requirements", {}).get("native_text", True) else report["warnings"],
                           "image_only_slide", "Full-page raster image has no useful native content objects", slide=index)
                if deck is not None and index <= len(deck["slides"]) and isinstance(deck["slides"][index - 1], dict):
                    planned = deck["slides"][index - 1]
                    info["slide_id"] = planned.get("id")
                    _check_content(report, index, planned, deck, objects, root, relations.get(part, {}), xml)
    except (OSError, zipfile.BadZipFile, RuntimeError, ValueError, OverflowError) as exc:
        _issue(report["errors"], "package_unreadable", str(exc))
    report["success"] = not report["errors"]
    report["hard_error_count"] = len(report["errors"])
    return report


def _check_content(report, index, planned, deck, objects, root, page_rels, xml):
    by_id = {o["id"]: o for o in objects}
    refs = {}
    try:
        mapping = read_json(project_path(root, f"slides/{planned['id']}/objects.json"))
        if isinstance(mapping, dict) and isinstance(mapping.get("objects"), list):
            for record in mapping["objects"]:
                if not isinstance(record, dict):
                    continue
                ref, shape_id = record.get("content_ref"), record.get("pptx_shape_id")
                if not ref:
                    continue
                if not shape_id:
                    _issue(report["warnings"], "shape_binding_missing", "Object map has no exported shape ID; only content/type fallback can be checked", slide=index, content_ref=ref)
                elif shape_id not in by_id:
                    _issue(report["errors"], "mapped_shape_missing", "Object map refers to a missing exported shape", slide=index, content_ref=ref, shape_id=shape_id)
                else:
                    refs.setdefault(ref, []).append(by_id[shape_id])
    except (OSError, ValueError, KeyError, TypeError):
        pass  # Project validation already reports the malformed/missing map.
    useful = [o for o in objects if not o["hidden"] and not o["negligible_text"]]
    text = compact("".join(o["text"] for o in useful if o["kind"] == "text"))
    texts = [("title", planned.get("title", ""))]
    for item in planned.get("texts", []):
        if isinstance(item, dict) and item.get("required", True):
            texts.append(("text:" + str(item.get("id")), item.get("text", "")))
    for ref, expected in texts:
        expected = compact(expected)
        if not expected:
            continue
        bound = refs.get(ref)
        actual = compact("".join(o["text"] for o in bound if o["kind"] == "text" and not o["hidden"] and not o["negligible_text"])) if bound else text
        if expected not in actual:
            _issue(report["errors"] if deck.get("requirements", {}).get("native_text", True) else report["warnings"],
                   "text_missing", "Required text was not retained as useful native text", slide=index, content_ref=ref, expected=expected)
    req = deck.get("requirements", {})
    for kind, field, flag in (("table", "tables", "editable_tables"), ("chart", "charts", "editable_charts"), ("formula", "formulas", "editable_formulas")):
        candidates = [o for o in useful if o["kind"] == kind]
        expected_native = [i for i in planned.get(field, []) if isinstance(i, dict) and i.get("required", True) and i.get("editable", req.get(flag, True))]
        actual_count = sum(o["formula_count"] for o in candidates) if kind == "formula" else len(candidates)
        if actual_count < len(expected_native):
            _issue(report["errors"], "native_object_missing", f"Required editable {kind} object count is too small", slide=index, expected=len(expected_native), actual=actual_count)
        for order, item in enumerate(expected_native):
            ref = kind + ":" + str(item.get("id"))
            bound = refs.get(ref)
            selected = [o for o in bound if o["kind"] == kind and not o["hidden"]] if bound else candidates[order:order + 1]
            if bound and not selected:
                _issue(report["errors"], "editability_mismatch", f"Mapped {kind} is not a visible native {kind} object", slide=index, content_ref=ref)
            if not selected:
                continue
            obj = selected[0]
            if kind == "table":
                actual = _native_table_cells(obj["table"])
                expected = item.get("cells", [])
                if [[compact(v) for v in row] for row in actual] != [[compact(v) for v in row] for row in expected]:
                    _issue(report["errors"], "table_content_mismatch", "Native table cells differ from the planned source cells", slide=index, content_ref=ref)
            elif kind == "chart":
                link = page_rels.get(obj["chart_rid"], {})
                chart = xml.get(link.get("target"))
                if chart is None or link.get("external"):
                    _issue(report["errors"], "chart_unreadable", "Native chart cannot be resolved to an internal chart XML part", slide=index, content_ref=ref)
                    continue
                actual = _chart_series(chart)
                expected = item.get("series", [])
                if len(actual) != len(expected):
                    _issue(report["errors"], "chart_series_mismatch", "Native chart series count differs from the source", slide=index, content_ref=ref)
                for si, series in enumerate(expected):
                    if si >= len(actual) or not isinstance(series, dict):
                        continue
                    if "name" in series and compact(series["name"]) != compact(actual[si]["name"]):
                        _issue(report["errors"], "chart_series_name_mismatch", "Native chart series name differs from the source", slide=index, content_ref=ref, series=si,
                               expected=series["name"], actual=actual[si]["name"])
                    values = series.get("values", [])
                    try:
                        same = len(values) == len(actual[si]["values"]) and all(math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-9) for a, b in zip(values, actual[si]["values"]))
                    except (ValueError, TypeError):
                        same = False
                    if not same:
                        _issue(report["errors"], "chart_values_mismatch", "Native chart cached values differ from the source", slide=index, content_ref=ref, series=si)
                    cats = actual[si]["categories"]
                    if cats and [compact(v) for v in cats] != [compact(v) for v in item.get("categories", [])]:
                        _issue(report["errors"], "chart_categories_mismatch", "Native chart categories differ from the source", slide=index, content_ref=ref, series=si)
                    elif not cats:
                        _issue(report["warnings"], "chart_categories_unverified", "No cached chart categories were found", slide=index, content_ref=ref)
            elif kind == "formula":
                display = item.get("display_text")
                if display and compact(display) not in compact("".join(o["math_text"] for o in selected)):
                    _issue(report["errors"], "formula_text_mismatch", "Native formula tokens differ from the provided display_text", slide=index, content_ref=ref)
                elif not display:
                    _issue(report["warnings"], "formula_equivalence_unverified", "Native equation exists; LaTeX equivalence still requires mathematical and visual review", slide=index, content_ref=ref)
    for asset_id in planned.get("required_assets", []):
        bound = refs.get("asset:" + str(asset_id))
        if bound and not any(o["kind"] == "image" and not o["hidden"] for o in bound):
            _issue(report["errors"], "asset_object_missing", "Required image asset is not mapped to a visible picture object", slide=index, content_ref="asset:" + str(asset_id))
        elif not bound:
            _issue(report["warnings"] if any(o["kind"] == "image" for o in useful) else report["errors"],
                   "asset_binding_unverified", "Required asset has no actual exported picture binding", slide=index, content_ref="asset:" + str(asset_id))


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--project", type=Path, help="Task directory containing deck.json and per-page object maps")
    parser.add_argument("--expected-slides", type=int)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args(argv)
    try:
        report = audit(args.pptx, args.project, args.expected_slides)
        if args.out:
            write_json(args.out, report)
    except (OSError, ValueError, TypeError, KeyError) as exc:
        report = {"success": False, "errors": [{"code": "audit_error", "message": str(exc)}], "limitations": LIMITATIONS}
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report.get("success") else 1


if __name__ == "__main__":
    sys.exit(main())
