#!/usr/bin/env python3
"""Bind object maps to exported PPTX node IDs without changing the PPTX."""
from __future__ import annotations

import argparse
import copy
import json
import math
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import audit_pptx as package_audit
from project import project_path, read_json, sha256_file, validate_project, write_json

P, A, R = package_audit.P, package_audit.A, package_audit.R


def _error(report, code, message, **where):
    report["errors"].append({"code": code, "message": message, **where})


def _compatible(expected, actual):
    # A decorative shape can contain a label while remaining a native p:sp.
    if expected == "shape":
        return actual["element"].tag == P + "sp" and actual["kind"] in {"shape", "text"}
    return expected == actual["kind"]


def _bbox_matches(expected, actual, tolerance):
    return (isinstance(expected, list) and isinstance(actual, list) and len(expected) == len(actual) == 4
            and all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in expected + actual)
            and all(abs(a - b) <= tolerance for a, b in zip(expected, actual)))


def _endpoints(node):
    start = node["element"].find(".//" + A + "stCxn")
    end = node["element"].find(".//" + A + "endCxn")
    return (start.get("id") if start is not None else None, end.get("id") if end is not None else None)


def bind_ids(pptx: Path, root: Path, tolerance=0.03, dry_run=False) -> dict:
    root, pptx = root.resolve(), pptx.resolve()
    report = {"schema_version": 1, "project": str(root), "pptx": str(pptx), "success": False,
              "dry_run": bool(dry_run), "errors": [], "warnings": [], "bindings": [], "changed_maps": [],
              "pptx_modified": False,
              "note": "Only object maps are updated. Changed maps invalidate previous static audits and hash-bound reviews; rerun audit and review on the final artifact."}
    if not isinstance(tolerance, (float, int)) or not math.isfinite(tolerance) or not 0 <= tolerance <= 0.1:
        _error(report, "invalid_tolerance", "Use a finite bbox tolerance between 0 and 0.1 inches")
        return report
    planning = validate_project(root, require_maps=True)
    report["errors"].extend({**value, "scope": "project"} for value in planning["errors"])
    if report["errors"]:
        return report
    deck_path = project_path(root, "deck.json")
    deck = read_json(deck_path)
    deck_hash, pptx_hash = sha256_file(deck_path), sha256_file(pptx)
    report["pptx_sha256"] = pptx_hash
    preflight = package_audit.audit(pptx, expected_slides=len(deck["slides"]))
    # A caller can still bind an out-of-bounds item and then repair its layout.
    # Structural package defects cannot be safely used to identify objects.
    for value in preflight["errors"]:
        if value["code"] == "out_of_bounds":
            report["warnings"].append(value)
        else:
            report["errors"].append({**value, "scope": "pptx"})
    if report["errors"]:
        return report
    snapshots, replacements = {}, {}
    with zipfile.ZipFile(pptx) as archive:
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        relationships = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
        links = {rel.get("Id"): rel for rel in relationships}
        size = presentation.find(P + "sldSz")
        canvas = (float(size.get("cx")), float(size.get("cy")))
        page_ids = presentation.findall(P + "sldIdLst/" + P + "sldId")
        for index, (slide, page) in enumerate(zip(deck["slides"], page_ids), 1):
            slide_id = slide["id"]
            map_path = project_path(root, f"slides/{slide_id}/objects.json")
            snapshots[map_path] = map_path.read_bytes()
            mapping = json.loads(snapshots[map_path])
            updated = copy.deepcopy(mapping)
            link = links[page.get(R + "id")]
            part = package_audit._target("ppt/presentation.xml", link.get("Target"))
            tree = ET.fromstring(archive.read(part))
            shape_tree = tree.find(P + "cSld/" + P + "spTree")
            diagnostics = {"errors": [], "warnings": []}
            nodes = package_audit._shape_info(shape_tree, canvas, diagnostics, index)
            by_name = {}
            known_map_ids = {obj["id"] for obj in mapping["objects"]}
            for node in nodes:
                metadata = node["element"].find(".//" + P + "cNvPr")
                node["name"] = metadata.get("name", "") if metadata is not None else ""
                if not isinstance(node["id"], str) or not node["id"].isdigit():
                    _error(report, "invalid_exported_id", "Actual exported shape ID is not numeric", slide_id=slide_id)
                    continue
                by_name.setdefault(node["name"], []).append(node)
                if node["name"].startswith("ifppt:"):
                    prefix = f"ifppt:{slide_id}:"
                    if not node["name"].startswith(prefix):
                        _error(report, "slide_identity_mismatch", "Managed object name belongs to a different slide; verify presentation order", slide_id=slide_id, name=node["name"], part=part)
                    elif node["name"][len(prefix):] not in known_map_ids:
                        _error(report, "unexpected_named_object", "Export contains a managed object missing from this map", slide_id=slide_id, name=node["name"])
            bound, used = {}, set()

            def assign(obj, node, method):
                if node["hidden"]:
                    _error(report, "hidden_candidate", "Cannot bind to a hidden object", slide_id=slide_id, object_id=obj["id"])
                    return
                if not _compatible(obj["kind"], node):
                    _error(report, "binding_kind_mismatch", "Exported object type differs from the map", slide_id=slide_id, object_id=obj["id"], expected=obj["kind"], actual=node["kind"])
                    return
                if node["id"] in used:
                    _error(report, "duplicate_binding", "Two map objects would bind to the same exported object", slide_id=slide_id, object_id=obj["id"], shape_id=node["id"])
                    return
                bound[obj["id"]] = node
                used.add(node["id"])
                report["bindings"].append({"slide_id": slide_id, "object_id": obj["id"], "shape_id": node["id"], "method": method, "part": part})

            # Reserve exact managed names before anonymous geometry matching.
            for obj in mapping["objects"]:
                candidates = by_name.get(f"ifppt:{slide_id}:{obj['id']}", [])
                if len(candidates) > 1:
                    _error(report, "ambiguous_name", "More than one exported object has the exact managed name", slide_id=slide_id, object_id=obj["id"])
                elif len(candidates) == 1:
                    assign(obj, candidates[0], "name")
            for obj in mapping["objects"]:
                if obj["id"] in bound or obj["kind"] == "connector":
                    continue
                candidates = [node for node in nodes if node["id"] not in used and not node["hidden"]
                              and not node["name"].startswith("ifppt:") and _compatible(obj["kind"], node)
                              and _bbox_matches(obj["bbox"], node["bbox_in"], tolerance)]
                if len(candidates) != 1:
                    _error(report, "ambiguous_geometry" if candidates else "no_geometry_match",
                           "Type and bbox must identify exactly one unbound exported object; no IDs were guessed",
                           slide_id=slide_id, object_id=obj["id"], candidates=[node["id"] for node in candidates])
                else:
                    assign(obj, candidates[0], "type_and_bbox")
            # Connectors are directional. Geometry alone is unsafe after rerouting.
            for obj in mapping["objects"]:
                if obj["kind"] != "connector":
                    continue
                start, end = bound.get(obj.get("from")), bound.get(obj.get("to"))
                if start is None or end is None:
                    _error(report, "connector_endpoints_unbound", "Connector requires from/to object IDs bound to actual nonconnector nodes", slide_id=slide_id, object_id=obj["id"])
                    continue
                endpoints = (start["id"], end["id"])
                if obj["id"] in bound:
                    if _endpoints(bound[obj["id"]]) != endpoints:
                        _error(report, "connector_endpoint_mismatch", "Named connector endpoint IDs do not match the map's from/to objects", slide_id=slide_id, object_id=obj["id"])
                    continue
                candidates = [node for node in nodes if node["id"] not in used and not node["hidden"] and node["kind"] == "connector"
                              and not node["name"].startswith("ifppt:") and _endpoints(node) == endpoints]
                if len(candidates) != 1:
                    _error(report, "ambiguous_connector" if candidates else "connector_not_found",
                           "Exactly one connector must have the directed native endpoint IDs; bbox or old IDs will not resolve ambiguity",
                           slide_id=slide_id, object_id=obj["id"], from_shape_id=endpoints[0], to_shape_id=endpoints[1], candidates=[node["id"] for node in candidates])
                else:
                    assign(obj, candidates[0], "native_endpoints")
            for obj in updated["objects"]:
                if obj["id"] in bound:
                    obj["pptx_shape_id"] = bound[obj["id"]]["id"]
            if mapping != updated:
                replacements[map_path] = updated
    if report["errors"]:
        report["bindings_planned_only"] = True
        return report
    if sha256_file(pptx) != pptx_hash or sha256_file(deck_path) != deck_hash or any(path.read_bytes() != value for path, value in snapshots.items()):
        _error(report, "inputs_changed", "PPTX, deck or maps changed during binding; no files were modified")
        return report
    if not dry_run:
        written = []
        try:
            for path, value in replacements.items():
                write_json(path, value)
                written.append(path)
        except OSError:
            for path in reversed(written):
                temporary = path.with_name(path.name + ".restore.tmp")
                temporary.write_bytes(snapshots[path])
                temporary.replace(path)
            raise
    report["changed_maps"] = [str(path.relative_to(root)) for path in replacements]
    report["map_hashes_after"] = {str(path.relative_to(root)): sha256_file(path) for path in snapshots} if not dry_run else None
    report["binding_count"] = len(report["bindings"])
    report["success"] = True
    return report


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--tolerance-in", type=float, default=0.03, help="Per-coordinate bbox tolerance in inches (0 to 0.1)")
    parser.add_argument("--dry-run", action="store_true", help="Report all bindings without changing maps")
    args = parser.parse_args(argv)
    try:
        report = bind_ids(args.pptx, args.project, args.tolerance_in, args.dry_run)
    except (OSError, ValueError, TypeError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        report = {"success": False, "errors": [{"code": "binding_error", "message": str(exc)}], "pptx_modified": False}
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report.get("success") else 1


if __name__ == "__main__":
    sys.exit(main())
