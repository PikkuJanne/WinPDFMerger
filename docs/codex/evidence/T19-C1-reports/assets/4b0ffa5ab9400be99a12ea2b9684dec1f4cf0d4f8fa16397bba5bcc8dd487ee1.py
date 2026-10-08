"""Read-only structural/PDFium oracle for original T19 synthetic feature PDFs.

Strict pypdf parsing observes raw canonical fields/widgets, links, outlines,
annotations, embedded payloads and minimal tag relationships. Native PDFium
independently reads pages/identifiers/geometry and optionally renders appearances.
This development tool never repairs, fills, flattens or rewrites an input PDF,
follows a URI, or certifies signatures, accessibility, PDF/A or universal fidelity.
"""
from __future__ import annotations

import argparse
from collections import Counter
from contextlib import closing
import hashlib
import io
import json
import logging
from pathlib import Path
import re
import sys
import warnings

import PIL
import pypdf
from pypdf import PdfReader
from pypdf.generic import (ArrayObject, BooleanObject, ContentStream,
                          DictionaryObject, IndirectObject, NullObject,
                          StreamObject)
import pypdfium2 as pdfium
import pypdfium2_raw

VERSIONS = {"python": "3.12.14", "pypdf": "6.10.0", "pypdfium2": "5.13.0",
            "pdfium": "153.0.7999.0", "pillow": "12.3.0"}
PDFIUM_DLL_SHA256 = "958e5342ed7e2e20fb914adde238bbae0ac8ad4a3267aa49d0b9dd266c7667f2"
IDENTIFIER = re.compile(r"T19-[AB]-P[12]")
REPO = Path(__file__).resolve().parents[2]


def require_versions() -> dict:
    actual = {"python": sys.version.split()[0], "pypdf": pypdf.__version__,
              "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO),
              "pillow": PIL.__version__}
    if actual != VERSIONS:
        raise RuntimeError(f"Feature oracle requires approved development pins: {actual}")
    library = Path(pypdfium2_raw.__file__).parent / "pdfium.dll"
    actual["pdfium_dll_sha256"] = hashlib.sha256(library.read_bytes()).hexdigest()
    if actual["pdfium_dll_sha256"] != PDFIUM_DLL_SHA256:
        raise RuntimeError("Feature oracle requires the approved PDFium DLL bytes.")
    return actual


def resolve(value):
    return value.get_object() if isinstance(value, IndirectObject) else value


def reference(value) -> str | None:
    if isinstance(value, IndirectObject):
        return f"{value.idnum}:{value.generation}"
    ref = getattr(value, "indirect_reference", None)
    return reference(ref) if ref is not None else None


def plain(value):
    value = resolve(value)
    if value is None or isinstance(value, NullObject): return None
    if isinstance(value, BooleanObject): return bool(value.value)
    if isinstance(value, bytes): return {"bytes": len(value), "sha256": hashlib.sha256(value).hexdigest()}
    if isinstance(value, (str, bool, int, float)): return value
    if isinstance(value, (list, ArrayObject)): return [plain(item) for item in value]
    return {"object_ref": reference(value), "kind": type(value).__name__}


def raw_get(value, key, default=None):
    value = resolve(value)
    return value.get(key, default) if isinstance(value, DictionaryObject) else default


def page_index(value, page_refs) -> int | None:
    value_ref = reference(value)
    if value_ref in page_refs: return page_refs[value_ref]
    value = resolve(value)
    if isinstance(value, int) and not isinstance(value, bool) and 0 <= value < len(page_refs): return int(value)
    return None


def walk_pairs(tree, key: str, issues: list[str], ancestry=None) -> list[tuple]:
    """Walk raw name/number trees without collapsing duplicate keys."""
    ancestry = set() if ancestry is None else ancestry
    tree = resolve(tree)
    if tree is None: return []
    if not isinstance(tree, DictionaryObject):
        issues.append(f"{key} tree node is not a dictionary"); return []
    identity = id(tree)
    if identity in ancestry:
        issues.append(f"Cycle in {key} tree"); return []
    ancestry = ancestry | {identity}
    result = []
    values = resolve(tree.get(key))
    if values is not None:
        if not isinstance(values, ArrayObject) or len(values) % 2:
            issues.append(f"{key} tree leaf must have key/value pairs")
        else:
            result.extend((values[index], values[index + 1]) for index in range(0, len(values), 2))
    kids = resolve(tree.get("/Kids", ArrayObject()))
    if not isinstance(kids, ArrayObject): issues.append(f"{key} tree Kids is not an array")
    else:
        for child in kids: result.extend(walk_pairs(child, key, issues, ancestry))
    return result


def stream_snapshot(value) -> dict:
    value = resolve(value)
    if not isinstance(value, StreamObject): return {"kind": type(value).__name__, "object_ref": reference(value)}
    raw = value.get_data()
    encoded = getattr(value, "_data", b"")
    return {"kind": "stream", "object_ref": reference(value), "decoded_bytes": len(raw),
            "decoded_sha256": hashlib.sha256(raw).hexdigest(),
            "encoded_bytes": len(encoded), "encoded_sha256": hashlib.sha256(encoded).hexdigest(),
            "text_literals": re.findall(r"\(([^()]*)\)\s*Tj", raw.decode("latin-1"))}


def appearance_snapshot(annotation) -> dict:
    appearance = resolve(raw_get(annotation, "/AP"))
    if not isinstance(appearance, DictionaryObject): return {"present": False, "normal": None}
    normal = resolve(appearance.get("/N"))
    if isinstance(normal, StreamObject): normalized = stream_snapshot(normal)
    elif isinstance(normal, DictionaryObject):
        normalized = {"kind": "states", "states": {str(key): stream_snapshot(value) for key, value in normal.items()}}
    else: normalized = None
    return {"present": normalized is not None, "normal": normalized}


def file_spec_snapshot(value) -> dict:
    spec = resolve(value)
    if not isinstance(spec, DictionaryObject):
        return {"file_spec_ref": reference(value), "filename": None, "streams": {}, "issue": "Filespec is not a dictionary"}
    embedded = resolve(spec.get("/EF"))
    streams = {}
    if isinstance(embedded, DictionaryObject):
        for key, stream in embedded.items(): streams[str(key)] = stream_snapshot(stream)
    return {"file_spec_ref": reference(value), "filename": plain(spec.get("/UF", spec.get("/F"))),
            "legacy_filename": plain(spec.get("/F")), "description": plain(spec.get("/Desc")), "streams": streams}


def inherited(field, key):
    visited = set()
    while isinstance(resolve(field), DictionaryObject):
        field = resolve(field)
        if id(field) in visited: return None
        visited.add(id(field))
        if key in field: return field.get(key)
        field = field.get("/Parent")
    return None


def inspect_forms(catalog, pages, page_refs) -> dict:
    issues = []; nodes = []; fields = []; widgets = []; relations = []
    acroform = resolve(raw_get(catalog, "/AcroForm"))
    root_fields = resolve(raw_get(acroform, "/Fields", ArrayObject()))
    visited = set()

    def field_tree(raw, parent_name="", parent_ref=None):
        node = resolve(raw)
        if not isinstance(node, DictionaryObject): issues.append("Field-tree node is not a dictionary"); return
        if id(node) in visited: issues.append("Repeated object/cycle in canonical Fields tree"); return
        visited.add(id(node)); ref = reference(raw)
        local_name = plain(node.get("/T")); full_name = ".".join(part for part in [parent_name, local_name] if part)
        kids = resolve(node.get("/Kids", ArrayObject()))
        if not isinstance(kids, ArrayObject): issues.append("Field Kids is not an array"); kids = ArrayObject()
        widget_kids = [reference(child) for child in kids if raw_get(child, "/Subtype") == "/Widget"]
        record = {"object_ref": ref, "local_name": local_name, "full_name": full_name,
                  "field_type": plain(inherited(node, "/FT")), "value": plain(inherited(node, "/V")),
                  "default_value": plain(inherited(node, "/DV")), "flags": plain(inherited(node, "/Ff")),
                  "parent_ref": reference(node.get("/Parent")), "traversal_parent_ref": parent_ref,
                  "kids_refs": [reference(child) for child in kids], "widget_refs": widget_kids,
                  "combined_widget": node.get("/Subtype") == "/Widget"}
        nodes.append(record)
        child_fields = [child for child in kids if raw_get(child, "/Subtype") != "/Widget" or raw_get(child, "/T") is not None]
        if record["field_type"] is not None and not child_fields:
            canonical = dict(record)
            if record["combined_widget"]: canonical["widget_refs"] = [ref]
            fields.append(canonical)
        for child in child_fields: field_tree(child, full_name, ref)

    if not isinstance(root_fields, ArrayObject): issues.append("AcroForm Fields is not an array")
    else:
        for root in root_fields: field_tree(root)
    canonical_refs = {row["object_ref"]: row for row in fields if row["object_ref"] is not None}
    for index, page in enumerate(pages):
        for raw in resolve(page.get("/Annots", ArrayObject())):
            widget = resolve(raw)
            if not isinstance(widget, DictionaryObject) or widget.get("/Subtype") != "/Widget": continue
            ref = reference(raw); chain = []; ancestor = widget.get("/Parent"); seen = set()
            while isinstance(resolve(ancestor), DictionaryObject):
                parent = resolve(ancestor)
                if id(parent) in seen: issues.append("Cycle in widget Parent chain"); break
                seen.add(id(parent)); chain.append(reference(ancestor)); ancestor = parent.get("/Parent")
            canonical_ref = next((item for item in [ref] + chain if item in canonical_refs), None)
            name_parts = []; cursor = widget; seen = set()
            while isinstance(resolve(cursor), DictionaryObject):
                cursor = resolve(cursor)
                if id(cursor) in seen: break
                seen.add(id(cursor))
                if cursor.get("/T") is not None: name_parts.append(str(cursor.get("/T")))
                cursor = cursor.get("/Parent")
            record = {"object_ref": ref, "page_index": index, "declared_page_index": page_index(widget.get("/P"), page_refs),
                      "full_name": ".".join(reversed(name_parts)), "direct_name": plain(widget.get("/T")),
                      "field_type": plain(inherited(widget, "/FT")), "effective_value": plain(inherited(widget, "/V")),
                      "direct_value": plain(widget.get("/V")), "canonical_field_ref": canonical_ref,
                      "parent_chain": chain, "rectangle": plain(widget.get("/Rect")),
                      "appearance_state": plain(widget.get("/AS")), "appearance": appearance_snapshot(widget)}
            widgets.append(record)
            canonical = canonical_refs.get(canonical_ref)
            relations.append({"widget_ref": ref, "canonical_field_ref": canonical_ref,
                "canonical_in_field_tree": canonical is not None,
                "widget_in_canonical_kids": canonical is not None and ref in canonical["widget_refs"],
                "parent_chain_reaches_canonical": canonical_ref is not None and (canonical_ref in chain or canonical_ref == ref),
                "page_association_matches": record["declared_page_index"] == index,
                "effective_value_matches_canonical": canonical is not None and record["effective_value"] == canonical["value"]})
    name_counts = Counter(row["full_name"] for row in fields)
    return {"acroform_present": isinstance(acroform, DictionaryObject),
            "need_appearances": plain(raw_get(acroform, "/NeedAppearances")),
            "canonical_fields": fields, "field_tree_nodes": nodes, "widgets": widgets,
            "relations": relations, "duplicate_canonical_names": sorted(name for name, count in name_counts.items() if count > 1),
            "issues": issues}


def inspect_destinations(reader, catalog, page_refs, identifiers) -> tuple[list, dict, list]:
    issues = []; named = []
    names = resolve(raw_get(catalog, "/Names")); pairs = walk_pairs(raw_get(names, "/Dests"), "/Names", issues)
    legacy = resolve(raw_get(catalog, "/Dests"))
    if isinstance(legacy, DictionaryObject): pairs.extend(legacy.items())

    def target(value, lookup=None, visited=None):
        visited = set() if visited is None else visited
        value = resolve(value)
        if isinstance(value, str):
            name = str(value).lstrip("/")
            if name in visited: return {"named": name, "page_index": None, "page_identifier": None, "issue": "Named destination cycle"}
            if lookup and name in lookup:
                result = target(lookup[name], lookup, visited | {name}); return {"named": name, **result}
            return {"named": name, "page_index": None, "page_identifier": None, "issue": "Named target absent"}
        if isinstance(value, DictionaryObject): value = resolve(value.get("/D"))
        if not isinstance(value, ArrayObject) or not value:
            return {"page_index": None, "page_identifier": None, "issue": "Destination is not a page array"}
        index = page_index(value[0], page_refs)
        return {"page_index": index, "page_identifier": identifiers[index] if index is not None else None,
                "mode": plain(value[1]) if len(value) > 1 else None, "operands": [plain(item) for item in value[2:]]}

    lookup = {str(name).lstrip("/"): destination for name, destination in pairs}
    for name, destination in pairs: named.append({"name": str(name).lstrip("/"), "target": target(destination, lookup)})
    outlines = []

    def outline(items, depth=0):
        for item in items:
            if isinstance(item, list): outline(item, depth + 1)
            else:
                index = reader.get_destination_page_number(item)
                if index is not None and not 0 <= index < len(identifiers): index = None
                outlines.append({"title": str(item.get("/Title", "")), "depth": depth,
                                 "page_index": index, "page_identifier": identifiers[index] if index is not None else None})
    outline(reader.outline)
    return named, {"resolve": lambda value: target(value, lookup)}, outlines


def content_mcids(page, reader) -> list[dict]:
    if page.get("/Contents") is None: return []
    result = []
    for operands, operator in ContentStream(page.get_contents(), reader).operations:
        if operator != b"BDC" or len(operands) < 2: continue
        properties = resolve(operands[1])
        if isinstance(properties, str):
            property_map = resolve(raw_get(raw_get(page, "/Resources"), "/Properties"))
            properties = resolve(raw_get(property_map, properties))
        if isinstance(properties, DictionaryObject) and "/MCID" in properties:
            result.append({"tag": str(operands[0]), "mcid": int(properties["/MCID"])})
    return result


def inspect_tags(catalog, pages, page_refs, mcids) -> dict:
    root_raw = raw_get(catalog, "/StructTreeRoot"); root = resolve(root_raw)
    issues = []; nodes = []; seen = set()

    def node(raw, inherited_page=None):
        item = resolve(raw)
        if isinstance(item, ArrayObject):
            for child in item: node(child, inherited_page)
            return
        if not isinstance(item, DictionaryObject): return
        if id(item) in seen: issues.append("Repeated object/cycle in structure K traversal"); return
        seen.add(id(item))
        declared = page_index(item.get("/Pg"), page_refs); effective = declared if declared is not None else inherited_page
        children = resolve(item.get("/K")); children = children if isinstance(children, ArrayObject) else [children]
        kids = []
        for child in children:
            value = resolve(child)
            if isinstance(value, (int, float)): kids.append({"kind": "mcid", "mcid": int(value), "page_index": effective})
            elif isinstance(value, DictionaryObject) and value.get("/Type") == "/MCR":
                kids.append({"kind": "mcr", "mcid": plain(value.get("/MCID")), "page_index": page_index(value.get("/Pg"), page_refs) if value.get("/Pg") is not None else effective,
                             "stream_ref": reference(value.get("/Stm"))})
            elif value is not None: kids.append({"kind": "object", "object_ref": reference(child), "type": plain(raw_get(value, "/Type"))})
        record = {"object_ref": reference(raw), "type": plain(item.get("/Type")), "role": plain(item.get("/S")),
                  "parent_ref": reference(item.get("/P")), "declared_page_index": declared, "effective_page_index": effective,
                  "alt": plain(item.get("/Alt")), "kids": kids}
        nodes.append(record)
        for child in children:
            if raw_get(child, "/Type") == "/StructElem" or raw_get(child, "/S") is not None: node(child, effective)

    if isinstance(root, DictionaryObject): node(root.get("/K"))
    pairs = walk_pairs(raw_get(root, "/ParentTree"), "/Nums", issues)
    parent_entries = []
    for key, value in pairs:
        values = resolve(value)
        parent_entries.append({"key": int(key), "value_refs": [reference(item) for item in values] if isinstance(values, ArrayObject) else [reference(value)], "array": isinstance(values, ArrayObject)})
    lookup = {row["key"]: row for row in parent_entries}; node_lookup = {row["object_ref"]: row for row in nodes if row["object_ref"] is not None}
    associations = []
    for index, page in enumerate(pages):
        key = plain(page.get("/StructParents")); parent = lookup.get(key)
        for occurrence in mcids[index]:
            mcid = occurrence["mcid"]
            ref = parent["value_refs"][mcid] if parent and parent["array"] and 0 <= mcid < len(parent["value_refs"]) else None
            element = node_lookup.get(ref)
            match = element is not None and any(kid.get("mcid") == mcid and kid.get("page_index") == index for kid in element["kids"])
            associations.append({"page_index": index, "struct_parent_key": key, "mcid": mcid, "tag": occurrence["tag"],
                                 "element_ref": ref, "element_in_structure_tree": element is not None,
                                 "element_page_index": element["effective_page_index"] if element else None,
                                 "element_content_reference_matches": match})
    return {"marked": plain(raw_get(raw_get(catalog, "/MarkInfo"), "/Marked")),
            "root_present": isinstance(root, DictionaryObject), "root_ref": reference(root_raw),
            "language": plain(raw_get(catalog, "/Lang")), "nodes": nodes,
            "parent_tree_entries": parent_entries, "mcid_associations": associations, "issues": issues,
            "certification": "Minimal explicit relationships only; no PDF/UA, accessibility or reading-order certification."}


class WarningLog(logging.Handler):
    def __init__(self): super().__init__(logging.WARNING); self.messages = []
    def emit(self, record): self.messages.append(record.getMessage())


def inspect_pdf(path: Path, render_directory: Path | None = None, dpi: int = 144) -> dict:
    versions = require_versions(); path = path.resolve()
    if not path.is_relative_to(REPO / "tests" / ".work") or not path.is_file():
        raise ValueError("Inspect only explicitly owned synthetic evidence PDFs under tests/.work.")
    if not 72 <= dpi <= 300: raise ValueError("Use a bounded 72..300 DPI development render.")
    if render_directory is not None:
        render_directory = render_directory.resolve()
        if not render_directory.is_relative_to(REPO / "tests" / ".work") or render_directory.exists():
            raise ValueError("Render only into a new explicitly owned tests/.work directory.")
    before = path.read_bytes(); identifiers = []; page_details = []; renders = []
    with pdfium.PdfDocument(path) as document:
        # Disable XFA and provide no JavaScript platform. Corpus URIs are not followed.
        document.init_forms(config=pdfium.raw.FPDF_FORMFILLINFO(version=2, xfa_disabled=True))
        if render_directory is not None: render_directory.mkdir(parents=True, exist_ok=False)
        for index in range(len(document)):
            with closing(document[index]) as page:
                with closing(page.get_textpage()) as text:
                    found = IDENTIFIER.findall(text.get_text_range())
                if len(found) != 1: raise ValueError(f"PDFium page {index + 1} requires exactly one original T19 identifier: {found}")
                identifiers.append(found[0])
                page_details.append({"rotation_degrees": page.get_rotation(), "size_points": list(page.get_size())})
                if render_directory is not None:
                    destination = render_directory / f"page-{index + 1:02d}-{found[0]}.png"
                    with closing(page.render(scale=dpi / 72, may_draw_forms=True, draw_annots=True)) as bitmap:
                        image = bitmap.to_pil(); image.save(destination)
                        renders.append({"page_index": index, "identifier": found[0], "path": str(destination),
                                        "sha256": hashlib.sha256(destination.read_bytes()).hexdigest(), "bytes": destination.stat().st_size,
                                        "pixel_width": image.width, "pixel_height": image.height,
                                        "dpi": dpi, "draw_forms": True, "draw_annots": True})
                        image.close()
    if len(set(identifiers)) != len(identifiers): raise ValueError("Original T19 identifiers must remain distinct.")
    warning_log = WarningLog(); logger = logging.getLogger("pypdf"); logger.addHandler(warning_log)
    try:
        with warnings.catch_warnings(record=True) as warning_records:
            warnings.simplefilter("always")
            reader = PdfReader(io.BytesIO(before), strict=True); pages = list(reader.pages)
            if len(pages) != len(identifiers): raise ValueError("Strict pypdf/native PDFium page counts differ.")
            catalog = reader.trailer["/Root"]; page_refs = {reference(page): index for index, page in enumerate(pages)}
            named, destination_resolver, outlines = inspect_destinations(reader, catalog, page_refs, identifiers)
            forms = inspect_forms(catalog, pages, page_refs); mcids = [content_mcids(page, reader) for page in pages]
            structured_pages = []; action_types = set()
            for index, page in enumerate(pages):
                annotations = []
                raw_annots = resolve(page.get("/Annots", ArrayObject()))
                if not isinstance(raw_annots, ArrayObject): raise ValueError("Page Annots must be an array.")
                for raw in raw_annots:
                    annotation = resolve(raw)
                    if not isinstance(annotation, DictionaryObject): raise ValueError("Annotation must be a dictionary.")
                    action = resolve(annotation.get("/A")); action_type = plain(raw_get(action, "/S"))
                    if action_type: action_types.add(action_type)
                    target = annotation.get("/Dest")
                    if isinstance(action, DictionaryObject) and action_type == "/GoTo": target = action.get("/D")
                    facts = {"object_ref": reference(raw), "subtype": plain(annotation.get("/Subtype")),
                             "rectangle": plain(annotation.get("/Rect")), "name": plain(annotation.get("/NM")),
                             "contents": plain(annotation.get("/Contents")), "title": plain(annotation.get("/T")),
                             "declared_page_index": page_index(annotation.get("/P"), page_refs),
                             "flags": plain(annotation.get("/F")), "quad_points": plain(annotation.get("/QuadPoints")),
                             "color": plain(annotation.get("/C")), "opacity": plain(annotation.get("/CA")),
                             "action_type": action_type, "uri": plain(raw_get(action, "/URI")),
                             "destination": destination_resolver["resolve"](target) if target is not None else None,
                             "appearance": appearance_snapshot(annotation)}
                    if annotation.get("/FS") is not None: facts["file_attachment"] = file_spec_snapshot(annotation.get("/FS"))
                    annotations.append(facts)
                structured_pages.append({"index": index, "identifier": identifiers[index], "pdfium": page_details[index],
                    "structural": {"rotation_degrees": int(page.get("/Rotate", 0)), "media_box": plain(page.get("/MediaBox")),
                                   "crop_box": plain(page.get("/CropBox", page.get("/MediaBox"))), "struct_parents": plain(page.get("/StructParents"))},
                    "annotations": annotations, "content_mcids": mcids[index]})
            embedded_issues = []; names = resolve(catalog.get("/Names"))
            embedded = [{"name": str(name), **file_spec_snapshot(spec)} for name, spec in walk_pairs(raw_get(names, "/EmbeddedFiles"), "/Names", embedded_issues)]
            result = {"schema_version": 1, "versions": versions,
                "file": {"path": str(path), "sha256": hashlib.sha256(before).hexdigest(), "bytes": len(before)},
                "parser": {"strict": True, "warnings": warning_log.messages + [str(item.message) for item in warning_records]},
                "page_count": len(pages), "page_identifiers": identifiers, "pages": structured_pages,
                "forms": forms, "named_destinations": named, "bookmarks": outlines, "embedded_files": embedded,
                "embedded_file_issues": embedded_issues, "tagging": inspect_tags(catalog, pages, page_refs, mcids),
                "signature_observations": {"signature_field_count": sum(row["field_type"] == "/Sig" for row in forms["canonical_fields"]),
                    "xfa_present": raw_get(raw_get(catalog, "/AcroForm"), "/XFA") is not None,
                    "doc_mdp_present": raw_get(raw_get(catalog, "/Perms"), "/DocMDP") is not None,
                    "validation_performed": False},
                "active_content_observations": {"annotation_action_types": sorted(action_types),
                    "catalog_open_action_present": catalog.get("/OpenAction") is not None,
                    "catalog_additional_actions_present": catalog.get("/AA") is not None,
                    "javascript_name_tree_present": raw_get(names, "/JavaScript") is not None,
                    "uris_followed": False, "javascript_platform_provided": False, "xfa_render_disabled": True},
                "renders": renders,
                "limitations": ["Strict parsing is an observation boundary, not a PDF-malware or complete conformance validator.",
                    "Canonical form data, page widgets and normal appearances are inspected separately; rendering alone does not prove interactive field retention.",
                    "Named/local destinations resolve to actual page offsets/identifiers; URI targets are recorded but never followed.",
                    "EmbeddedFiles payload streams and page FileAttachment annotations are distinct observations.",
                    "Minimal tagged structure/ParentTree/MCID associations are not accessibility, reading-order or PDF/UA certification.",
                    "No signature/XFA validity, PDF/A, universal fidelity, archival or feature-preservation guarantee."]}
            # Include warnings raised by the last structure/signature reads too.
            result["parser"]["warnings"] = warning_log.messages + [str(item.message) for item in warning_records]
    finally: logger.removeHandler(warning_log)
    if path.read_bytes() != before: raise ValueError("Source bytes changed while the read-only oracle inspected them.")
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--pdf", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--render-dir", type=Path)
    parser.add_argument("--dpi", type=int, default=144)
    args = parser.parse_args(); output = args.output.resolve()
    if not output.is_relative_to(REPO / "tests" / ".work") or output.exists():
        raise ValueError("Write snapshot only to a new owned JSON file under tests/.work.")
    result = inspect_pdf(args.pdf, args.render_dir, args.dpi)
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(result, indent=2, ensure_ascii=False) + "\n")
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
