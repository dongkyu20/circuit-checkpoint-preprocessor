import copy
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
try:
    import preprocess as p
except ImportError:
    p = None


def analysis(text="입력 핀을 접지한다"):
    return {
        "summary": text,
        "checks": [{
            "target_kind": "input_pin", "target_description": "미사용 데이터 입력 핀",
            "part_numbers": [], "applicability": "미사용 입력인 경우",
            "requirement": text, "level": "mandatory",
            "parameters": [], "required_design_facts": ["pin_usage", "connected_net"],
            "evidence": [{"field": "checkpoint", "quote": "GND"}]
        }], "issues": []
    }


@unittest.skipUnless(p, "production module missing")
class PreprocessTests(unittest.TestCase):
    def test_normalization_decodes_entities_preserves_active_low_and_newlines(self):
        self.assertEqual(p.normalize_text(" /OE &amp; DIR\r\n10&#91;ohm&#93; "),
                         "/OE & DIR\n10[ohm]")
        self.assertEqual(p.normalize_text(r"pin\_info 1\. input"), "pin_info 1. input")

    def test_cli_envelope_and_fenced_response_are_parsed(self):
        payload = {"items": [{"request_id": "x", "analysis": analysis()}]}
        output = json.dumps({"response": "```json\n" + json.dumps(payload) + "\n```", "stats": {}})
        self.assertEqual(p.parse_cli_output(output), payload)
        with self.assertRaises(p.CliError):
            p.parse_cli_output(json.dumps({"error": {"message": "quota exhausted"}}))

    def test_at_signs_in_reference_cells_cannot_trigger_cli_file_inclusion(self):
        payload = {"rows": [{"source": {"reference_document": '@첨부파일 "file.pdf"'}}]}
        stdin = p.make_stdin(payload)
        self.assertNotIn("@", stdin)
        self.assertEqual(json.loads(stdin.split("\n\n")[-1]), payload)

    def test_schema_rejects_missing_fields_and_invented_evidence(self):
        good = analysis()
        p.validate_analysis(good, {"checkpoint": "unused input GND"})
        bad = copy.deepcopy(good)
        del bad["checks"][0]["level"]
        with self.assertRaises(p.FormatError):
            p.validate_analysis(bad, {"checkpoint": "GND"})
        with self.assertRaises(p.FormatError):
            p.validate_analysis(good, {"checkpoint": "floating"})
        bad = copy.deepcopy(good)
        bad["checks"][0]["level"] = "PASS"
        with self.assertRaises(p.FormatError):
            p.validate_analysis(bad, {"checkpoint": "GND"})

    def test_cache_changes_when_source_or_prompt_changes(self):
        a = p.request_key({"checkpoint": "GND"}, "v1")
        self.assertNotEqual(a, p.request_key({"checkpoint": "VCC"}, "v1"))
        self.assertNotEqual(a, p.request_key({"checkpoint": "GND"}, "v2"))
        with tempfile.TemporaryDirectory() as tmp:
            with p.Cache(Path(tmp) / "state.sqlite3") as cache:
                cache.put(a, analysis())
            with p.Cache(Path(tmp) / "state.sqlite3") as cache:
                self.assertEqual(cache.get(a), analysis())

    def test_partial_batch_response_keeps_valid_row_and_reports_missing_row(self):
        sources = {"a": {"checkpoint": "GND"}, "b": {"checkpoint": "GND"}}
        valid, invalid = p.validate_items({"items": [{"request_id": "a", "analysis": analysis()}]}, sources)
        self.assertEqual(set(valid), {"a"})
        self.assertIn("b", invalid)
        with self.assertRaises(p.FormatError):
            p.validate_items({"items": [{"request_id": "foreign", "analysis": analysis()}]}, sources)

    def test_duplicate_response_ids_are_rejected(self):
        item = {"request_id": "a", "analysis": analysis()}
        with self.assertRaises(p.FormatError):
            p.validate_items({"items": [item, item]}, {"a": {"checkpoint": "GND"}})

    def test_batching_does_not_truncate_long_rows(self):
        items = [("a", {"checkpoint": "GND"}), ("b", {"checkpoint": "x" * 200})]
        batches = list(p.make_batches(items, 10, 100))
        self.assertEqual(len(batches), 2)
        self.assertEqual(batches[1][0][1]["checkpoint"], "x" * 200)

    def test_csv_cells_cannot_be_excel_formulas(self):
        self.assertEqual(p.csv_safe("=1+1"), "'=1+1")
        self.assertEqual(p.csv_safe(" /OE"), " /OE")

    def test_header_alias_and_duplicate_checkpoint_ids_preserve_rows(self):
        records = p.build_records([
            ["adminno", "seqno", "Checkpoint", "Specification", "설계치", "참조문서"],
            [1336023553, 0, "unused GND", "GND", "Unused input pin:", "file.pdf"],
            [1336023553, 0, "unused GND", "GND", "Unused input pin:", "file.pdf"],
            [1336023553, 1, None, None, None, None]
        ], "체크포인트_REF_GUIDE", 1)
        self.assertEqual(len(records), 3)
        self.assertNotEqual(records[0]["row_uid"], records[1]["row_uid"])
        self.assertEqual(records[0]["source_columns"]["adminno"], 1336023553)
        self.assertIn("duplicate_checkpoint_id", records[0]["source_issues"])
        self.assertFalse(records[2]["eligible"])
        with self.assertRaises(ValueError):
            p.build_records([["Checkpoint", "Specification", "設計値"]], "s", 1)

    def test_duplicate_auxiliary_headers_preserve_both_columns(self):
        records = p.build_records([
            ["Checkpoint", "Specification", "설계치", "category", "category"],
            ["GND", "GND", "pin:", "first", "second"]
        ], "s", 1)
        self.assertEqual(records[0]["source_columns"]["category [col 4]"], "first")
        self.assertEqual(records[0]["source_columns"]["category [col 5]"], "second")
        self.assertEqual(records[0]["renamed_headers"][0]["column"], 4)
        self.assertEqual(records[0]["normalized_source"]["checkpoint"], "GND")

    def test_case_and_space_collisions_in_auxiliary_headers_are_preserved(self):
        records = p.build_records([
            ["Checkpoint", "Specification", "설계치", "Ref Date", "refdate"],
            ["GND", "GND", "pin:", "2012", "2017"]
        ], "s", 1)
        self.assertEqual(records[0]["source_columns"]["Ref Date [col 4]"], "2012")
        self.assertEqual(records[0]["source_columns"]["refdate [col 5]"], "2017")

    def test_duplicate_used_headers_report_names_and_column_positions(self):
        for names in (["pin_info", "pin_info"], ["참조문서", "참조문서"]):
            with self.subTest(names=names):
                with self.assertRaisesRegex(ValueError, "4열.*5열"):
                    p.build_records([["Checkpoint", "Specification", "설계치", *names]], "s", 1)

    def test_generated_header_names_cannot_overwrite_existing_header(self):
        records = p.build_records([
            ["Checkpoint", "Specification", "설계치", "x", "x", "x [col 4]"],
            ["GND", "GND", "pin:", "a", "b", "c"]
        ], "s", 1)
        values = records[0]["source_columns"]
        self.assertEqual(len(values), 6)
        self.assertEqual(values["x [col 4]"], "c")
        self.assertIn("a", values.values())
        self.assertIn("b", values.values())

    def test_identical_joined_adminno_columns_are_used_without_losing_sources(self):
        rows = p.build_records([
            ["Checkpoint", "Specification", "설계치", "adminno", "adminno", "seqno"],
            ["GND", "GND", "pin:", "00123", "00123", "0"]
        ], "s", 1)
        self.assertEqual(rows[0]["checkpoint_id"], "00123:0")
        self.assertEqual(rows[0]["source_columns"]["adminno [col 4]"], "00123")
        self.assertEqual(rows[0]["source_columns"]["adminno [col 5]"], "00123")
        self.assertEqual(rows[0]["adminno_selection"]["mode"], "equal_duplicate")

    def test_differing_joined_adminno_columns_require_explicit_selection(self):
        values = [["Checkpoint", "Specification", "설계치", "adminno", "adminno", "seqno"],
                  ["GND", "GND", "pin:", "10", "20", "0"]]
        with self.assertRaisesRegex(ValueError, "adminno.*4열.*5열"):
            p.build_records(values, "s", 1)
        rows = p.build_records(values, "s", 1, adminno_column=5)
        self.assertEqual(rows[0]["checkpoint_id"], "20:0")
        self.assertEqual(rows[0]["adminno_selection"]["mode"], "explicit")
        self.assertIn("adminno_values_differ", rows[0]["source_issues"])
        with self.assertRaisesRegex(ValueError, "adminno"):
            p.build_records(values, "s", 1, adminno_column=2)

    def test_matching_seqno_columns_remain_distinct_even_when_values_match(self):
        rows = p.build_records([
            ["Checkpoint", "Specification", "설계치", "adminno", "adminno", "seqno", "Seq No"],
            ["GND", "GND", "pin:", "00123", "00123", "0", "0"]
        ], "s", 1)
        self.assertEqual(rows[0]["checkpoint_id"], "row:s:2")
        self.assertEqual(rows[0]["checkpoint_id_basis"], "row")
        self.assertNotIn("seqno_selection", rows[0])
        self.assertEqual([item["value"] for item in rows[0]["seqno_fields"]], ["0", "0"])
        self.assertEqual(rows[0]["source_columns"]["seqno [col 6]"], "0")
        self.assertEqual(rows[0]["source_columns"]["Seq No [col 7]"], "0")

    def test_distinct_seqno_values_are_preserved_and_selection_is_optional(self):
        values = [["Checkpoint", "Specification", "설계치", "adminno", "seqno", "seqno"],
                  ["GND", "GND", "pin:", "123", "0", "7"]]
        automatic = p.build_records(values, "s", 1)[0]
        self.assertEqual(automatic["checkpoint_id"], "row:s:2")
        self.assertEqual(automatic["normalized_source"]["seqno"], "")
        self.assertEqual([item["value"] for item in automatic["seqno_fields"]], ["0", "7"])
        self.assertNotIn("missing_checkpoint_identifiers", automatic["source_issues"])
        rows = p.build_records(values, "s", 1, seqno_column=6)
        self.assertEqual(rows[0]["checkpoint_id"], "123:7")
        self.assertNotIn("seqno_values_differ", rows[0]["source_issues"])
        self.assertEqual(p.build_records(values, "s", 1, seqno_column=5)[0]["checkpoint_id"], "123:0")
        with self.assertRaisesRegex(ValueError, "seqno"):
            p.build_records(values, "s", 1, seqno_column=4)

    def test_separate_seqno_columns_keep_unique_row_ids_and_original_values(self):
        records = p.build_records([
            ["adminno", "seqno", "seqno", "Checkpoint", "Specification", "설계치"],
            ["123", "0", "10", "GND", "GND", "pin:"],
            ["123", "0", "20", "GND", "GND", "pin:"]
        ], "joined", 1)
        self.assertEqual([r["checkpoint_id"] for r in records], ["row:joined:2", "row:joined:3"])
        self.assertEqual([r["seqno_fields"][1]["value"] for r in records], ["10", "20"])

    def test_schema_rejects_unexpected_properties(self):
        bad = analysis()
        bad["pass"] = True
        with self.assertRaises(p.FormatError):
            p.validate_analysis(bad, {"checkpoint": "GND"})


class ModuleExists(unittest.TestCase):
    def test_preprocessor_module_is_available(self):
        self.assertIsNotNone(p, "preprocess.py has not been implemented")


if __name__ == "__main__":
    unittest.main()
