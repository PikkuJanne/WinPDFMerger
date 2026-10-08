"""Feature oracle trust guards and raw in-memory PDF object graph faults.

These tests do not serialize a PDF, run the application or claim native feature
preservation. They protect observations that page text/rendering cannot prove.
"""
from __future__ import annotations

import importlib.util
import hashlib
from pathlib import Path
import unittest
from unittest import mock

from pypdf import PdfWriter
from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject,
                          NameObject, NumberObject, TextStringObject)

SOURCE = Path(__file__).resolve().parents[1] / "feature_oracle.py"
SPEC = importlib.util.spec_from_file_location("feature_oracle", SOURCE)
oracle = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(oracle)


def dictionary(**values):
    return DictionaryObject({NameObject("/" + key): value for key, value in values.items()})


def graph(page_count=2):
    writer = PdfWriter()
    for _ in range(page_count): writer.add_blank_page(width=612, height=792)
    refs = {oracle.reference(page): index for index, page in enumerate(writer.pages)}
    writer._root_object[NameObject("/AcroForm")] = dictionary(Fields=ArrayObject())
    return writer, refs


def field_with_widget(writer, page_index, value, appearance_value=None):
    """One raw canonical leaf and a separate child widget on its declared page."""
    field = dictionary(FT=NameObject("/Tx"), T=TextStringObject("shared_text"),
                       V=TextStringObject(value), Kids=ArrayObject())
    field_ref = writer._add_object(field)
    stream = DecodedStreamObject()
    stream.set_data(f"BT ({appearance_value or value}) Tj ET\n".encode("ascii"))
    widget = dictionary(Subtype=NameObject("/Widget"), Parent=field_ref,
                        P=writer.pages[page_index].indirect_reference,
                        AP=dictionary(N=writer._add_object(stream)))
    widget_ref = writer._add_object(widget); field["/Kids"].append(widget_ref)
    writer._root_object["/AcroForm"]["/Fields"].append(field_ref)
    page = writer.pages[page_index]
    if "/Annots" not in page: page[NameObject("/Annots")] = ArrayObject()
    page["/Annots"].append(widget_ref)
    return field, field_ref, widget, widget_ref


class FeatureOracleGraphTests(unittest.TestCase):
    def test_actual_current_pinned_dll_is_accepted_and_its_exact_hash_recorded(self):
        library = Path(oracle.pypdfium2_raw.__file__).parent / "pdfium.dll"
        actual_hash = hashlib.sha256(library.read_bytes()).hexdigest()
        result = oracle.require_versions()
        self.assertEqual({key: result[key] for key in oracle.VERSIONS}, oracle.VERSIONS)
        self.assertEqual(result["pdfium_dll_sha256"], actual_hash)
        self.assertIn(actual_hash, {
            "958e5342ed7e2e20fb914adde238bbae0ac8ad4a3267aa49d0b9dd266c7667f2",
            "524ecbe6a7d49103909b1ed39fe512d2d4e612e35dac1336c9274371d20c5d90",
        })

    def test_unapproved_dll_bytes_fail_even_when_reported_versions_match(self):
        with mock.patch.object(Path, "read_bytes", return_value=b"Unapproved synthetic DLL bytes"):
            with self.assertRaisesRegex(RuntimeError, "approved PDFium DLL bytes"):
                oracle.require_versions()

    def test_version_mismatch_fails_before_reading_selected_dll(self):
        with mock.patch.object(oracle.pdfium, "PDFIUM_INFO", "unapproved-version"), \
                mock.patch.object(Path, "read_bytes") as read_bytes:
            with self.assertRaisesRegex(RuntimeError, "approved development pins"):
                oracle.require_versions()
            read_bytes.assert_not_called()

    def test_same_name_canonical_fields_keep_distinct_values_and_widgets(self):
        writer, refs = graph()
        field_with_widget(writer, 0, "value-A"); field_with_widget(writer, 1, "value-B")
        facts = oracle.inspect_forms(writer._root_object, list(writer.pages), refs)
        self.assertEqual([f["value"] for f in facts["canonical_fields"]], ["value-A", "value-B"])
        self.assertEqual(facts["duplicate_canonical_names"], ["shared_text"])
        self.assertEqual([w["effective_value"] for w in facts["widgets"]], ["value-A", "value-B"])
        self.assertEqual(len({w["canonical_field_ref"] for w in facts["widgets"]}), 2)
        for relation in facts["relations"]:
            self.assertTrue(all(value for key, value in relation.items() if key not in ["widget_ref", "canonical_field_ref"]))

    def test_orphan_widget_is_visible_as_orphan_even_with_normal_appearance(self):
        writer, refs = graph()
        _, _, widget, _ = field_with_widget(writer, 0, "value-A")
        del widget["/Parent"]
        facts = oracle.inspect_forms(writer._root_object, list(writer.pages), refs)
        self.assertTrue(facts["widgets"][0]["appearance"]["present"])
        self.assertIsNone(facts["widgets"][0]["canonical_field_ref"])
        self.assertFalse(facts["relations"][0]["canonical_in_field_tree"])
        self.assertFalse(facts["relations"][0]["parent_chain_reaches_canonical"])

    def test_widget_parent_is_not_enough_when_missing_from_canonical_kids(self):
        writer, refs = graph()
        field, _, _, _ = field_with_widget(writer, 0, "value-A")
        field[NameObject("/Kids")] = ArrayObject()
        facts = oracle.inspect_forms(writer._root_object, list(writer.pages), refs)
        self.assertTrue(facts["relations"][0]["parent_chain_reaches_canonical"])
        self.assertFalse(facts["relations"][0]["widget_in_canonical_kids"])

    def test_widget_page_reference_mismatch_is_observed(self):
        writer, refs = graph()
        _, _, widget, _ = field_with_widget(writer, 0, "value-A")
        widget[NameObject("/P")] = writer.pages[1].indirect_reference
        facts = oracle.inspect_forms(writer._root_object, list(writer.pages), refs)
        self.assertEqual(facts["widgets"][0]["page_index"], 0)
        self.assertEqual(facts["widgets"][0]["declared_page_index"], 1)
        self.assertFalse(facts["relations"][0]["page_association_matches"])

    def test_widget_direct_value_and_normal_appearance_are_separate_observations(self):
        writer, refs = graph()
        _, _, widget, _ = field_with_widget(writer, 0, "value-A", "stale-appearance")
        widget[NameObject("/V")] = TextStringObject("conflicting-widget-value")
        facts = oracle.inspect_forms(writer._root_object, list(writer.pages), refs)
        self.assertEqual(facts["canonical_fields"][0]["value"], "value-A")
        self.assertEqual(facts["widgets"][0]["effective_value"], "conflicting-widget-value")
        self.assertEqual(facts["widgets"][0]["appearance"]["normal"]["text_literals"], ["stale-appearance"])
        self.assertFalse(facts["relations"][0]["effective_value_matches_canonical"])

    def test_repeated_or_cyclic_field_tree_does_not_recurse_or_collapse_silently(self):
        writer, refs = graph()
        field, field_ref, _, _ = field_with_widget(writer, 0, "value-A")
        field["/Kids"].append(field_ref)
        facts = oracle.inspect_forms(writer._root_object, list(writer.pages), refs)
        self.assertIn("Repeated object/cycle in canonical Fields tree", facts["issues"])

    def test_raw_name_tree_keeps_duplicate_keys_and_reports_cycle(self):
        writer, _ = graph()
        tree = dictionary(Names=ArrayObject([TextStringObject("same"), NumberObject(1),
                                            TextStringObject("same"), NumberObject(2)]))
        tree_ref = writer._add_object(tree); tree[NameObject("/Kids")] = ArrayObject([tree_ref])
        issues = []; pairs = oracle.walk_pairs(tree_ref, "/Names", issues)
        self.assertEqual([(str(key), int(value)) for key, value in pairs], [("same", 1), ("same", 2)])
        self.assertEqual(issues, ["Cycle in /Names tree"])

    def test_odd_number_tree_leaf_is_reported_without_invented_association(self):
        tree = dictionary(Nums=ArrayObject([NumberObject(0)])); issues = []
        self.assertEqual(oracle.walk_pairs(tree, "/Nums", issues), [])
        self.assertEqual(issues, ["/Nums tree leaf must have key/value pairs"])

    def test_named_destination_resolves_actual_merged_page_offset(self):
        writer, refs = graph(4)
        target = ArrayObject([writer.pages[3].indirect_reference, NameObject("/Fit")])
        writer._root_object[NameObject("/Names")] = dictionary(Dests=dictionary(
            Names=ArrayObject([TextStringObject("T19-B-second"), dictionary(D=target)])))
        class OutlineReader:
            outline = []
        names, resolver, _ = oracle.inspect_destinations(OutlineReader(), writer._root_object, refs,
            ["T19-A-P1", "T19-A-P2", "T19-B-P1", "T19-B-P2"])
        self.assertEqual(names[0]["target"]["page_index"], 3)
        self.assertEqual(names[0]["target"]["page_identifier"], "T19-B-P2")
        self.assertEqual(resolver["resolve"](TextStringObject("T19-B-second"))["page_index"], 3)
        self.assertEqual(oracle.page_index(NumberObject(3), refs), 3)
        self.assertIsNone(oracle.page_index(9, refs))

    def test_named_destination_cycles_and_absent_targets_are_reported(self):
        writer, refs = graph()
        writer._root_object[NameObject("/Names")] = dictionary(Dests=dictionary(Names=ArrayObject([
            TextStringObject("loop"), TextStringObject("loop")])))
        class OutlineReader:
            outline = []
        _, resolver, _ = oracle.inspect_destinations(OutlineReader(), writer._root_object, refs, ["T19-A-P1", "T19-A-P2"])
        self.assertEqual(resolver["resolve"](TextStringObject("loop"))["issue"], "Named destination cycle")
        self.assertEqual(resolver["resolve"](TextStringObject("absent"))["issue"], "Named target absent")

    def test_parent_tree_slot_requires_matching_structure_element_page_and_mcid(self):
        writer, refs = graph()
        paragraph = dictionary(Type=NameObject("/StructElem"), S=NameObject("/P"),
                               Pg=writer.pages[1].indirect_reference, K=NumberObject(0))
        paragraph_ref = writer._add_object(paragraph)
        root = dictionary(K=ArrayObject([paragraph_ref]), ParentTree=dictionary(
            Nums=ArrayObject([NumberObject(0), ArrayObject([paragraph_ref])])))
        writer._root_object[NameObject("/StructTreeRoot")] = writer._add_object(root)
        writer.pages[0][NameObject("/StructParents")] = NumberObject(0)
        facts = oracle.inspect_tags(writer._root_object, list(writer.pages), refs, [[{"tag":"/P","mcid":0}], []])
        association = facts["mcid_associations"][0]
        self.assertTrue(association["element_in_structure_tree"])
        self.assertFalse(association["element_content_reference_matches"])
        self.assertEqual(association["element_page_index"], 1)
        self.assertIn("no PDF/UA", facts["certification"])

    def test_named_marked_content_property_resolves_indirect_resource_dictionary(self):
        writer, _ = graph()
        properties = writer._add_object(dictionary(P0=dictionary(MCID=NumberObject(4))))
        writer.pages[0][NameObject("/Resources")] = dictionary(Properties=properties)
        content = DecodedStreamObject(); content.set_data(b"/P /P0 BDC EMC\n")
        writer.pages[0][NameObject("/Contents")] = writer._add_object(content)
        self.assertEqual(oracle.content_mcids(writer.pages[0], writer), [{"tag":"/P","mcid":4}])


if __name__ == "__main__": unittest.main()
