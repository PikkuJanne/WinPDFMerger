"""Deterministic development-only feature corpus, original CC0-1.0 content.

This recipe writes two original PDFs only after the caller performs the PDF
artifact-start marker. It reads no external PDF, image, font or network resource.
Generated PDFs stay in a new explicitly owned tests/.work directory. The normal
application never imports this module or depends on Python/ReportLab/pypdf.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
from pathlib import Path
import random
import sys

import pypdf
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (ArrayObject, BooleanObject, ByteStringObject,
                          DecodedStreamObject, DictionaryObject, FloatObject,
                          NameObject, NumberObject, TextStringObject)
import reportlab
from reportlab.pdfgen.canvas import Canvas

VERSIONS = {"python": "3.12.14", "reportlab": "4.4.9", "pypdf": "6.10.0"}
PAGE_SIZE = (612, 792)
COMMENT_HEX_CHARACTERS_PER_PAGE = 65536
FIXED_DATE = "D:20000101000000+00'00'"


def require_versions() -> dict:
    actual = {"python": sys.version.split()[0], "reportlab": reportlab.Version,
              "pypdf": pypdf.__version__}
    if actual != VERSIONS:
        raise RuntimeError(f"Generator requires approved exact development pins: {actual}")
    return actual


def dictionary(**values) -> DictionaryObject:
    return DictionaryObject({NameObject("/" + key): value for key, value in values.items()})


def numbers(values) -> ArrayObject:
    return ArrayObject([FloatObject(value) for value in values])


def comment_padding(key: str, page: int) -> bytes:
    """Nonpainting high-entropy comments retained by PDFtk stream compression.

    Ghostscript can omit comments when rewriting page contents. This artificial
    size weighting exercises email publication; it is not a representative size
    or fidelity benchmark. Each comment starts with %, never an executable op.
    """
    seed = 190000 + (ord(key) - ord("A")) * 100 + page
    hex_data = random.Random(seed).randbytes(COMMENT_HEX_CHARACTERS_PER_PAGE // 2).hex()
    return b"".join(("% T19 inert size weighting " + hex_data[index:index + 80] + "\n").encode("ascii")
                    for index in range(0, len(hex_data), 80))


def base_pages(key: str) -> bytes:
    buffer = io.BytesIO()
    canvas = Canvas(buffer, pagesize=PAGE_SIZE, invariant=1, pageCompression=0,
                    pdfVersion=(1, 7))
    canvas.setTitle("Original T19 feature document " + key)
    canvas.setAuthor("WinPDFMerger original synthetic fixture")
    canvas.setSubject("CC0 feature observations; no signature/XFA/accessibility certification")
    for page in (1, 2):
        canvas.setFillColorRGB(0.08, 0.18, 0.28)
        canvas.setFont("Helvetica-Bold", 19)
        canvas.drawString(52, 748, f"Original feature document {key} / page {page}")
        canvas.setFont("Helvetica", 10)
        canvas.drawString(52, 722, "CC0 synthetic content: inspect logical structures and appearance separately.")
        canvas.setStrokeColorRGB(0.25, 0.4, 0.52)
        canvas.line(52, 708, 560, 708)
        canvas.setFillColorRGB(0, 0, 0)
        canvas.setFont("Helvetica-Bold", 11)
        canvas.drawString(52, 658, f"Shared field name: shared_text / expected canonical value: value-{key}")
        canvas.setFont("Helvetica", 10)
        canvas.drawString(52, 582, "The bordered box above is a Widget appearance, not this static page text.")
        canvas.drawString(52, 550, f"URI link target: https://example.invalid/T19/{key}")
        canvas.drawString(52, 530, f"Named internal link target: document {key}, page 2")
        canvas.drawString(52, 492, f"Highlighted original fixture phrase {key} / page {page}.")
        canvas._code.append("/P <</MCID 0>> BDC")
        canvas.drawString(52, 450, f"Original tagged paragraph {key} / page {page}.")
        canvas._code.append("EMC")
        canvas.setFont("Helvetica", 10)
        canvas.drawString(52, 408, "Sticky note, highlight and attachment icon are annotation observations.")
        canvas.drawString(52, 390, "Links are inert structural targets; the test never follows a URI.")
        canvas.drawString(52, 354, "Named destinations, outlines, field data and tags can differ after rewriting.")
        canvas.drawString(52, 336, "These fixtures contain no XFA, JavaScript or digital signature.")
        canvas.drawString(52, 318, "Minimal tag relationships do not establish accessibility or PDF/UA validity.")
        canvas.setStrokeColorRGB(0.3, 0.3, 0.3)
        for index in range(6):
            canvas.setLineWidth(0.25 + index * 0.2)
            canvas.line(52, 270 - index * 14, 420, 270 - index * 14)
        canvas.setFont("Helvetica", 11)
        canvas.drawString(52, 24, f"T19-{key}-P{page}")
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(560, 24, "Original CC0 corpus; keep originals of feature-rich documents")
        canvas.showPage()
    canvas.save()
    return buffer.getvalue()


def add_annotation(writer: PdfWriter, page_index: int, annotation: DictionaryObject):
    page = writer.pages[page_index]
    annotation[NameObject("/P")] = page.indirect_reference
    ref = writer._add_object(annotation)
    if "/Annots" not in page:
        page[NameObject("/Annots")] = ArrayObject()
    page["/Annots"].append(ref)
    return ref


def add_forms(writer: PdfWriter, key: str) -> None:
    font = writer._add_object(dictionary(Type=NameObject("/Font"), Subtype=NameObject("/Type1"),
                                         BaseFont=NameObject("/Helvetica"), Encoding=NameObject("/WinAnsiEncoding")))
    field = dictionary(FT=NameObject("/Tx"), T=TextStringObject("shared_text"),
                       V=TextStringObject("value-" + key), DV=TextStringObject("value-" + key),
                       Ff=NumberObject(0), DA=TextStringObject("/Helv 12 Tf 0 g"), Kids=ArrayObject())
    field_ref = writer._add_object(field)
    for page in range(2):
        appearance = DecodedStreamObject()
        appearance.set_data(("q 1 1 1 rg 0 0 508 34 re f 0.25 0.4 0.52 RG 1 w "
                             "0.5 0.5 507 33 re S BT /Helv 12 Tf 0 g 10 11 Td "
                             f"(value-{key}) Tj ET Q\n").encode("ascii"))
        appearance.update(dictionary(Type=NameObject("/XObject"), Subtype=NameObject("/Form"),
                                      FormType=NumberObject(1), BBox=numbers([0, 0, 508, 34]),
                                      Resources=dictionary(Font=dictionary(Helv=font))))
        appearance_ref = writer._add_object(appearance)
        widget = dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Widget"),
                            Rect=numbers([52, 610, 560, 644]), Parent=field_ref, F=NumberObject(4),
                            NM=TextStringObject(f"T19-{key}-widget-{page + 1}"),
                            AP=dictionary(N=appearance_ref),
                            MK=dictionary(BC=numbers([0.25, 0.4, 0.52]), BG=numbers([1, 1, 1])))
        field["/Kids"].append(add_annotation(writer, page, widget))
    acroform = dictionary(Fields=ArrayObject([field_ref]), NeedAppearances=BooleanObject(False),
                         DA=TextStringObject("/Helv 12 Tf 0 g"), DR=dictionary(Font=dictionary(Helv=font)))
    writer._root_object[NameObject("/AcroForm")] = writer._add_object(acroform)


def add_links_and_notes(writer: PdfWriter, key: str) -> None:
    # Explicit destination array, with no action or external operation.
    destination = dictionary(D=ArrayObject([writer.pages[1].indirect_reference, NameObject("/Fit")]))
    destination_ref = writer._add_object(destination)
    names = writer._root_object.get("/Names", DictionaryObject())
    names[NameObject("/Dests")] = writer._add_object(dictionary(Names=ArrayObject([
        TextStringObject(f"T19-{key}-second"), destination_ref])))
    writer._root_object[NameObject("/Names")] = names
    for page in range(2):
        add_annotation(writer, page, dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Link"),
            Rect=numbers([52, 546, 440, 564]), Border=numbers([0, 0, 0]),
            A=dictionary(S=NameObject("/URI"), URI=TextStringObject(f"https://example.invalid/T19/{key}"))))
        add_annotation(writer, page, dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Link"),
            Rect=numbers([52, 526, 440, 544]), Border=numbers([0, 0, 0]),
            Dest=TextStringObject(f"T19-{key}-second")))
        add_annotation(writer, page, dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Text"),
            Rect=numbers([542, 668, 566, 692]), T=TextStringObject("Original CC0 fixture"),
            Contents=TextStringObject(f"Original sticky note {key} / page {page + 1}"),
            NM=TextStringObject(f"T19-{key}-note-{page + 1}"), Name=NameObject("/Comment"),
            C=numbers([1, 0.85, 0]), Open=BooleanObject(False), F=NumberObject(4)))
        add_annotation(writer, page, dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/Highlight"),
            Rect=numbers([50, 489, 440, 504]), QuadPoints=numbers([50, 504, 440, 504, 50, 489, 440, 489]),
            C=numbers([1, 1, 0]), CA=FloatObject(0.35),
            Contents=TextStringObject(f"Original highlight {key} / page {page + 1}"),
            NM=TextStringObject(f"T19-{key}-highlight-{page + 1}"), F=NumberObject(4)))
    parent = writer.add_outline_item(f"T19 {key} introduction", 0)
    writer.add_outline_item(f"T19 {key} rotated second page", 1, parent=parent)


def attachment_bytes(key: str) -> bytes:
    return (f"T19 document {key} attachment\nOriginal inert CC0-1.0 synthetic text.\n"
            f"Distinct payload: attachment-{key}. No script, credential or private data.\n").encode("utf-8")


def add_attachment(writer: PdfWriter, key: str) -> None:
    payload = attachment_bytes(key)
    stream = DecodedStreamObject(); stream.set_data(payload)
    stream.update(dictionary(Type=NameObject("/EmbeddedFile"), Subtype=NameObject("/text/plain"),
                             Params=dictionary(Size=NumberObject(len(payload)))))
    stream_ref = writer._add_object(stream)
    filename = f"T19-{key}-attachment.txt"
    spec = dictionary(Type=NameObject("/Filespec"), F=TextStringObject(filename), UF=TextStringObject(filename),
                      Desc=TextStringObject("Original inert CC0 text attachment"),
                      EF=dictionary(F=stream_ref, UF=stream_ref))
    spec_ref = writer._add_object(spec)
    names = writer._root_object["/Names"]
    names[NameObject("/EmbeddedFiles")] = writer._add_object(dictionary(Names=ArrayObject([
        TextStringObject(filename), spec_ref])))
    add_annotation(writer, 1, dictionary(Type=NameObject("/Annot"), Subtype=NameObject("/FileAttachment"),
        Rect=numbers([542, 550, 566, 574]), FS=spec_ref, Name=NameObject("/PushPin"),
        Contents=TextStringObject(f"Original attachment icon {key}"), F=NumberObject(4)))


def add_minimal_tags(writer: PdfWriter, key: str) -> None:
    root = dictionary(Type=NameObject("/StructTreeRoot"), K=ArrayObject(), ParentTreeNextKey=NumberObject(2))
    root_ref = writer._add_object(root)
    document = dictionary(Type=NameObject("/StructElem"), S=NameObject("/Document"), P=root_ref, K=ArrayObject())
    document_ref = writer._add_object(document); root["/K"].append(document_ref)
    nums = ArrayObject()
    for page_index, page in enumerate(writer.pages):
        page[NameObject("/StructParents")] = NumberObject(page_index)
        paragraph = dictionary(Type=NameObject("/StructElem"), S=NameObject("/P"), P=document_ref,
                               Pg=page.indirect_reference, K=NumberObject(0),
                               Alt=TextStringObject(f"Original tagged paragraph {key} / page {page_index + 1}"))
        paragraph_ref = writer._add_object(paragraph); document["/K"].append(paragraph_ref)
        nums.extend([NumberObject(page_index), ArrayObject([paragraph_ref])])
    root[NameObject("/ParentTree")] = writer._add_object(dictionary(Nums=nums))
    writer._root_object[NameObject("/StructTreeRoot")] = root_ref
    writer._root_object[NameObject("/MarkInfo")] = dictionary(Marked=BooleanObject(True))
    writer._root_object[NameObject("/Lang")] = TextStringObject("en-US")


def build_pdf(key: str) -> bytes:
    reader = PdfReader(io.BytesIO(base_pages(key)), strict=True)
    writer = PdfWriter(); writer.clone_document_from_reader(reader)
    writer._header = b"%PDF-1.7"
    writer.add_metadata({"/Title": "Original T19 feature document " + key,
                         "/Author": "WinPDFMerger original synthetic fixture",
                         "/Subject": "CC0 feature corpus; no PDF/UA, signature or XFA certification",
                         "/Producer": "Original recipe using ReportLab 4.4.9 / pypdf 6.10.0",
                         "/CreationDate": FIXED_DATE, "/ModDate": FIXED_DATE})
    identity = hashlib.sha256(("T19 original feature fixture " + key).encode()).digest()[:16]
    writer._ID = ArrayObject([ByteStringObject(identity), ByteStringObject(identity)])
    for page_index, page in enumerate(writer.pages):
        content = page.get_contents().get_data() + comment_padding(key, page_index + 1)
        stream = DecodedStreamObject(); stream.set_data(content)
        page[NameObject("/Contents")] = writer._add_object(stream)
    writer.pages[1][NameObject("/Rotate")] = NumberObject(90 if key == "A" else 270)
    add_forms(writer, key); add_links_and_notes(writer, key); add_attachment(writer, key); add_minimal_tags(writer, key)
    result = io.BytesIO(); writer.write(result)
    return result.getvalue()


def generate(output: Path) -> dict:
    versions = require_versions()
    root = Path(__file__).resolve().parents[3]
    output = output.resolve()
    if not output.is_relative_to(root / "tests" / ".work") or output.exists():
        raise ValueError("Use a new explicitly owned directory under this repository's tests/.work.")
    output.mkdir(parents=True, exist_ok=False)
    fixtures = []
    for index, key in enumerate(("A", "B"), 1):
        name = f"{index}-feature-{key}.pdf"; raw = build_pdf(key)
        with (output / name).open("xb") as stream: stream.write(raw)
        fixtures.append({"file": name, "sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw),
            "page_count": 2, "page_identifiers": [f"T19-{key}-P1", f"T19-{key}-P2"],
            "media_boxes": [[0, 0, 612, 792], [0, 0, 612, 792]],
            "structural_rotations": [0, 90 if key == "A" else 270],
            "canonical_field": {"name": "shared_text", "value": "value-" + key, "field_type": "/Tx", "widget_count": 2},
            "named_destination": {"name": f"T19-{key}-second", "page_index": 1},
            "uri": f"https://example.invalid/T19/{key}",
            "bookmark_titles": [f"T19 {key} introduction", f"T19 {key} rotated second page"],
            "annotation_counts": {"/Widget": 2, "/Link": 4, "/Text": 2, "/Highlight": 2, "/FileAttachment": 1},
            "attachment": {"name": f"T19-{key}-attachment.txt", "bytes": len(attachment_bytes(key)), "sha256": hashlib.sha256(attachment_bytes(key)).hexdigest()},
            "minimal_tags": {"struct_parent_keys": [0, 1], "mcids_per_page": [[0], [0]], "paragraph_elements": 2},
            "artificial_size_weighting": {"kind": "nonpainting seeded high-entropy hex comments in page content", "hex_characters_per_page": COMMENT_HEX_CHARACTERS_PER_PAGE,
                "seeds": [190000 + (ord(key) - ord("A")) * 100 + page for page in (1, 2)],
                "comment_bytes_per_page": [len(comment_padding(key, page)) for page in (1, 2)],
                "purpose": "Exercise publication of a smaller rewritten email derivative; not representative compression or fidelity evidence."}})
    manifest = {"schema_version": 1, "provenance": "Original text, vector marks, form appearances, annotations and embedded attachment bytes authored by this recipe; no external content.",
                "license": "CC0-1.0", "generator": "tests/fixtures/features/generate_features.py", "generator_version": 1,
                "generator_sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(), "versions": versions,
                "fixtures": fixtures, "merge_order": [row["file"] for row in fixtures], "expected_merged_page_count": 4,
                "expected_merged_page_identifiers": [identifier for row in fixtures for identifier in row["page_identifiers"]],
                "excluded_features": {"digital_signatures": "No fabricated signature or validity test; original signatures must be retained separately.", "xfa": "No fabricated XFA fixture or support claim.", "accessibility": "Minimal explicit tag associations only; no PDF/UA/accessibility certification."},
                "limits": ["Structural observations and rendered appearance are separate; neither establishes universal feature retention.", "Repeated same-name fields have distinct canonical values in different original documents.", "URI/internal targets are inert and are never followed by the oracle.", "Generated PDF bytes stay ignored; only recipe, manifest and sanitized observation evidence are tracked."]}
    with (output / "generation.json").open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n")
    return manifest


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(generate(args.output), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
