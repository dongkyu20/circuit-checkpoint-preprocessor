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
