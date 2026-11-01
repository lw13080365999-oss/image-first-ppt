#!/usr/bin/env python3
"""Create and inspect an image-first PPT workspace. Python standard library only."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import shutil
import sys
from pathlib import Path, PurePosixPath

SCHEMA_VERSION = 1
ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,79}$")
CONTENT_KINDS = {"text": "texts", "formula": "formulas", "table": "tables", "chart": "charts"}
OBJECT_KINDS = {*CONTENT_KINDS, "image", "shape", "connector"}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def project_path(root: Path, value: str) -> Path:
    """Reject traversal, absolute/Windows paths and symlinks escaping the root."""
    if not isinstance(value, str) or not value or "\\" in value or ":" in value or "\x00" in value:
        raise ValueError("Expected a nonempty POSIX relative project path")
    pure = PurePosixPath(value)
    if pure.is_absolute() or any(part in {"", ".", ".."} for part in value.split("/")):
        raise ValueError("Absolute paths, empty segments and traversal are not allowed")
    path = root.joinpath(*pure.parts)
    try:
        path.resolve().relative_to(root.resolve())
    except ValueError as exc:
        raise ValueError("Path or symlink leaves the project root") from exc
    return path


def issue(items: list, code: str, location: str, message: str) -> None:
    items.append({"code": code, "path": location, "message": message})


def _number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def validate_project(root: Path, require_maps: bool = False) -> dict:
    root = root.resolve()
    errors, warnings = [], []
    report = {"schema_version": 1, "project": str(root), "errors": errors, "warnings": warnings,
              "success": False, "slides": []}
    try:
        deck = read_json(root / "deck.json")
    except (OSError, ValueError) as exc:
        issue(errors, "deck_unreadable", "deck.json", str(exc))
        return report
    if not isinstance(deck, dict):
        issue(errors, "deck_type", "deck.json", "Expected an object")
        return report
    if deck.get("schema_version") != SCHEMA_VERSION:
        issue(errors, "schema_version", "schema_version", "Expected schema_version 1")
    if not isinstance(deck.get("title"), str) or not deck["title"].strip():
        issue(errors, "title", "title", "A nonempty deck title is required")
    if "style" in deck and not isinstance(deck["style"], dict):
        issue(errors, "style_type", "style", "Expected a style object")
    canvas = deck.get("canvas", {})
    if not isinstance(canvas, dict) or any(not _number(canvas.get(k)) or canvas[k] <= 0 for k in ("width_in", "height_in")):
        issue(errors, "canvas", "canvas", "Positive finite width_in and height_in are required")
    requirements = deck.get("requirements", {})
    if not isinstance(requirements, dict):
        issue(errors, "requirements_type", "requirements", "Expected an object")
    else:
        for key in ("native_text", "editable_tables", "editable_charts", "editable_formulas"):
            if key in requirements and not isinstance(requirements[key], bool):
                issue(errors, "requirement_type", f"requirements.{key}", "Expected a boolean")

    def entries(name: str, container=deck):
        value = container.get(name, [])
        if not isinstance(value, list):
            issue(errors, "array_type", name, "Expected an array")
            return []
        return value

    def indexed(items: list, location: str) -> dict:
        result = {}
        for index, item in enumerate(items):
            loc = f"{location}[{index}]"
            if not isinstance(item, dict):
                issue(errors, "entry_type", loc, "Expected an object")
                continue
            item_id = item.get("id")
            if not isinstance(item_id, str) or not ID_RE.fullmatch(item_id):
                issue(errors, "invalid_id", loc + ".id", "Use 1-80 letters, digits, underscores or hyphens; start with a letter or digit")
            elif item_id in result:
                issue(errors, "duplicate_id", loc + ".id", f"Duplicate id: {item_id}")
            else:
                result[item_id] = item
        return result

    def refs(value, known: dict, location: str) -> None:
        if not isinstance(value, list) or any(not isinstance(v, str) for v in value):
            issue(errors, "reference_type", location, "Expected an array of string IDs")
            return
        for ref in value:
            if ref not in known:
                issue(errors, "unknown_reference", location, f"Unknown id: {ref}")

    def check_file(item: dict, location: str, mandatory=True):
        if "path" not in item:
            if mandatory:
                issue(errors, "missing_path", location, "A project-relative file path is required")
            return None
        try:
            path = project_path(root, item["path"])
        except ValueError as exc:
            issue(errors, "unsafe_path", location + ".path", str(exc))
            return None
        if not path.is_file():
            issue(errors, "missing_file", location + ".path", f"File does not exist: {item['path']}")
            return None
        expected = item.get("sha256")
        if expected is not None:
            if not isinstance(expected, str) or not re.fullmatch(r"[0-9a-f]{64}", expected):
                issue(errors, "sha256_format", location + ".sha256", "Expected lowercase SHA-256 hex")
            elif sha256_file(path) != expected:
                issue(errors, "file_changed", location + ".sha256", "File hash differs from the recorded artifact")
        return path

    sources = indexed(entries("sources"), "sources")
    assets = indexed(entries("assets"), "assets")
    for key, value in sources.items():
        check_file(value, f"sources.{key}")
    for key, value in assets.items():
        check_file(value, f"assets.{key}")
        refs(value.get("source_ids", []), sources, f"assets.{key}.source_ids")
    slides = indexed(entries("slides"), "slides")
    if not slides:
        issue(errors, "no_slides", "slides", "No content slides have been planned")
    for slide_id, slide in slides.items():
        loc = f"slides.{slide_id}"
        if not isinstance(slide.get("title"), str) or not slide["title"].strip():
            issue(errors, "slide_title", loc + ".title", "A nonempty page title is required")
        refs(slide.get("source_ids", []), sources, loc + ".source_ids")
        refs(slide.get("required_assets", []), assets, loc + ".required_assets")
        known = {"title": "text"}
        required = {"title"}
        for kind, field in CONTENT_KINDS.items():
            records = indexed(entries(field, slide), loc + "." + field)
            for content_id, content in records.items():
                ref = f"{kind}:{content_id}"
                known[ref] = kind
                if content.get("required", True):
                    required.add(ref)
                for flag in ("required", "editable"):
                    if flag in content and not isinstance(content[flag], bool):
                        issue(errors, "content_flag", loc + f".{field}.{content_id}.{flag}", "Expected a boolean")
                refs(content.get("source_ids", []), sources, loc + f".{field}.{content_id}.source_ids")
                if kind in {"text", "formula"}:
                    key = "text" if kind == "text" else "latex"
                    if not isinstance(content.get(key), str) or not content[key].strip():
                        issue(errors, "empty_content", loc + f".{field}.{content_id}.{key}", "Expected nonempty source content")
                if kind == "table":
                    cells = content.get("cells")
                    if not isinstance(cells, list) or not cells or any(not isinstance(row, list) or not row for row in cells):
                        issue(errors, "table_cells", loc + f".tables.{content_id}", "Expected a nonempty rectangular array")
                    elif len({len(row) for row in cells}) != 1:
                        issue(errors, "table_shape", loc + f".tables.{content_id}", "Table rows must have equal lengths")
                if kind == "chart":
                    categories, series = content.get("categories"), content.get("series")
                    if not isinstance(categories, list) or not categories or not isinstance(series, list) or not series:
                        issue(errors, "chart_data", loc + f".charts.{content_id}", "Nonempty categories and series are required")
                    else:
                        for si, ser in enumerate(series):
                            values = ser.get("values") if isinstance(ser, dict) else None
                            if not isinstance(values, list) or len(values) != len(categories) or any(not _number(v) for v in values):
                                issue(errors, "chart_values", loc + f".charts.{content_id}.series[{si}]", "Finite values must match the category count")
        known.update({f"asset:{key}": "image" for key in assets})
        asset_refs = slide.get("required_assets", [])
        if isinstance(asset_refs, list):
            required.update(f"asset:{key}" for key in asset_refs if isinstance(key, str))
        try:
            object_path = project_path(root, f"slides/{slide_id}/objects.json")
        except ValueError as exc:
            issue(errors, "unsafe_map_path", loc + ".objects", str(exc))
            continue
        state = {"id": slide_id, "objects_present": object_path.is_file(), "design_present": False,
                 "mapping_count": 0, "mapped_references": []}
        report["slides"].append(state)
        if not object_path.is_file():
            if require_maps:
                issue(errors, "missing_map", str(object_path.relative_to(root)), "Page object map is required")
            continue
        try:
            mapping = read_json(object_path)
        except (OSError, ValueError) as exc:
            issue(errors, "map_unreadable", loc + ".objects", str(exc))
            continue
        if not isinstance(mapping, dict):
            issue(errors, "map_type", loc + ".objects", "Expected an object")
            continue
        if mapping.get("schema_version") != 1 or mapping.get("slide_id") != slide_id:
            issue(errors, "map_identity", loc + ".objects", "Map schema/version and slide_id must match the page")
        design = mapping.get("design")
        if design:
            if not isinstance(design, dict):
                issue(errors, "design_type", loc + ".objects.design", "Expected path and optional sha256")
            else:
                state["design_present"] = check_file(design, loc + ".objects.design") is not None
        elif require_maps:
            issue(errors, "missing_design", loc + ".objects.design", "A selected design image is required")
        objects = indexed(entries("objects", mapping), loc + ".objects.objects")
        mapped = set()
        for oid, obj in objects.items():
            oloc = loc + f".objects.{oid}"
            kind = obj.get("kind")
            if kind not in OBJECT_KINDS:
                issue(errors, "object_kind", oloc + ".kind", "Unknown object kind")
            ref = obj.get("content_ref")
            if ref is not None:
                if not isinstance(ref, str) or ref not in known:
                    issue(errors, "unknown_content_ref", oloc + ".content_ref", f"Unknown content reference: {ref}")
                else:
                    mapped.add(ref)
                    if known[ref] != kind:
                        issue(errors, "mapping_kind", oloc + ".kind", f"{ref} requires object kind {known[ref]}")
            elif kind in {*CONTENT_KINDS, "image"}:
                issue(warnings, "unbound_object", oloc, "Object has no source-content binding")
            bbox = obj.get("bbox")
            if not isinstance(bbox, list) or len(bbox) != 4 or any(not _number(v) for v in bbox) or any(v < 0 for v in bbox[2:]) or (bbox[2] == 0 and bbox[3] == 0):
                issue(errors, "bbox", oloc + ".bbox", "Expected finite [x, y, width, height] in inches; positive area or connector extent")
            if "editable" in obj and not isinstance(obj["editable"], bool):
                issue(errors, "editable_type", oloc + ".editable", "Expected a boolean")
            if "pptx_shape_id" in obj and (not isinstance(obj["pptx_shape_id"], str) or not obj["pptx_shape_id"].isdigit()):
                issue(errors, "shape_id", oloc + ".pptx_shape_id", "Expected the exported numeric shape ID as a string")
        state["mapping_count"] = len(objects)
        state["mapped_references"] = sorted(mapped)
        if require_maps:
            for ref in sorted(required - mapped):
                issue(errors, "unmapped_content", loc + ".objects", f"Required content is not mapped: {ref}")
    outputs = deck.get("outputs", {})
    if not isinstance(outputs, dict):
        issue(errors, "outputs_type", "outputs", "Expected an object of artifact paths")
    else:
        for key, value in outputs.items():
            try:
                project_path(root, value)
            except ValueError as exc:
                issue(errors, "unsafe_output_path", f"outputs.{key}", str(exc))
    report["success"] = not errors
    report["counts"] = {"sources": len(sources), "assets": len(assets), "slides": len(slides)}
    return report


def init_project(root: Path, title: str, inputs: list[str], count: int) -> dict:
    root = root.resolve()
    if (root / "deck.json").exists():
        raise ValueError("deck.json already exists; initialization never overwrites a project")
    files = [Path(value).expanduser().resolve() for value in inputs]
    if any(not path.is_file() for path in files):
        raise ValueError("Every --source must be an existing regular file")
    for folder in ("source", "assets", "slides", "output"):
        project_path(root, folder).mkdir(parents=True, exist_ok=True)
    sources = []
    for number, path in enumerate(files, 1):
        source_id = f"src-{number:03d}"
        name = re.sub(r"[^\w.() -]", "_", path.name)
        relative = f"source/{source_id}-{name}"
        destination = project_path(root, relative)
        if destination.exists():
            raise ValueError(f"Refusing to replace an existing source: {relative}")
        shutil.copy2(path, destination)
        sources.append({"id": source_id, "path": relative, "kind": path.suffix.lstrip(".").lower() or "file",
                        "sha256": sha256_file(destination)})
    slides = []
    for number in range(1, count + 1):
        slide_id = f"slide-{number:03d}"
        slides.append({"id": slide_id, "title": "", "source_ids": [], "texts": [], "formulas": [],
                       "tables": [], "charts": [], "required_assets": [], "notes": ""})
        for folder in ("design", "assets", "renders"):
            project_path(root, f"slides/{slide_id}/{folder}").mkdir(parents=True, exist_ok=True)
    deck = {"schema_version": 1, "title": title, "canvas": {"width_in": 13.333333, "height_in": 7.5},
            "style": {"id": "auto"},
            "requirements": {"native_text": True, "editable_tables": True, "editable_charts": True, "editable_formulas": True},
            "sources": sources, "assets": [], "slides": slides,
            "outputs": {"pptx": "output/deck.pptx", "audit": "output/static-audit.json"}}
    write_json(project_path(root, "source/index.json"), {"schema_version": 1, "sources": sources})
    write_json(root / "deck.json", deck)
    return {"success": True, "project": str(root), "stage": "initialized", "source_count": len(sources),
            "slide_count": count, "note": "Content is not inferred by this script; fill deck.json before validation/build."}


def project_status(root: Path) -> dict:
    result = validate_project(root)
    result["stage"] = "needs_content_or_repair" if result["errors"] else "content_planned"
    try:
        deck = read_json(root / "deck.json")
        output = project_path(root, deck.get("outputs", {}).get("pptx", "output/deck.pptx"))
        result["pptx_present"] = output.is_file()
        result["static_audit_current"] = False
        if output.is_file():
            result["stage"] = "built_static_check_pending"
            digest = sha256_file(output)
            result["pptx_sha256"] = digest
            audit_path = project_path(root, deck.get("outputs", {}).get("audit", "output/static-audit.json"))
            if audit_path.is_file():
                audit = read_json(audit_path)
                recorded_maps = audit.get("object_maps_sha256", {})
                current_maps = {slide["id"]: sha256_file(project_path(root, f"slides/{slide['id']}/objects.json"))
                                for slide in deck.get("slides", []) if isinstance(slide, dict)
                                and isinstance(slide.get("id"), str) and ID_RE.fullmatch(slide["id"])
                                and project_path(root, f"slides/{slide['id']}/objects.json").is_file()}
                current = (audit.get("pptx_sha256") == digest and audit.get("success") is True
                           and not audit.get("errors") and audit.get("deck_sha256") == sha256_file(root / "deck.json")
                           and recorded_maps == current_maps)
                result["static_audit_current"] = current
                if current:
                    result["stage"] = "static_checked_visual_review_pending"
        elif result["success"] and result["slides"] and all(s["design_present"] for s in result["slides"]):
            result["stage"] = "designs_present"
        if result["errors"]:
            result["stage"] = "needs_repair"
        result["visual_review_verified"] = False
        result["note"] = "Status is derived from files and hashes. Static checks do not establish visual approval or completion."
    except (OSError, ValueError, TypeError, AttributeError) as exc:
        issue(result["errors"], "status_unreadable", "outputs", str(exc))
        result["success"] = False
    return result


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    init = commands.add_parser("init", help="Create task-local source index and content manifest")
    init.add_argument("project", type=Path)
    init.add_argument("--title", required=True)
    init.add_argument("--source", action="append", default=[])
    init.add_argument("--slides", type=int, default=0, help="Create N empty planning records, without inventing content")
    for name in ("validate", "status"):
        command = commands.add_parser(name)
        command.add_argument("project", type=Path)
        if name == "validate":
            command.add_argument("--require-maps", action="store_true")
        command.add_argument("--out", type=Path, help="Optional report file")
    args = parser.parse_args(argv)
    try:
        if args.command == "init":
            if args.slides < 0 or not args.title.strip():
                raise ValueError("Use a nonempty title and a nonnegative slide count")
            result = init_project(args.project, args.title, args.source, args.slides)
        elif args.command == "validate":
            result = validate_project(args.project, args.require_maps)
        else:
            result = project_status(args.project)
        if getattr(args, "out", None):
            write_json(args.out, result)
    except (OSError, ValueError, TypeError) as exc:
        result = {"success": False, "errors": [{"code": "command_error", "message": str(exc)}]}
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    sys.exit(main())
