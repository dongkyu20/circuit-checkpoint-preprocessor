"""Offline integration: actual XLSX reader + actual subprocess + restart cache."""
import csv
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from xml.sax.saxutils import escape
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]


def make_xlsx(path, rows):
    """Minimal test fixture, not a deliverable workbook."""
    headers = ["adminno", "seqno", "Checkpoint", "Specification", "설계치", "참조문서"]
    def row_xml(index, values):
        return '<row r="%d">%s</row>' % (index, "".join(
            '<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>' %
            (chr(65 + column), index, escape(str(value))) for column, value in enumerate(values)))
    with ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="체크포인트_REF_GUIDE" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + row_xml(1, headers) + "".join(row_xml(i, values) for i, values in enumerate(rows, 2)) + '</sheetData></worksheet>')


FAKE = r'''#!PYTHON
import json, os, pathlib, sys, time
text = sys.stdin.read()
payload = json.loads(text.split("\n\n")[-1])
log = pathlib.Path(os.environ["CHECKPOINT_TEST_CALLS"])
with log.open("a") as f:
    f.write(json.dumps([row["request_id"] for row in payload["rows"]]) + "\n")
mode = os.environ.get("CHECKPOINT_TEST_MODE", "")
if mode == "validate_tools":
    settings = json.loads(pathlib.Path(".gemini/settings.json").read_text())
    core = settings.get("tools", {}).get("core")
    if core == []:
        print("The GenerateContentRequest proto is invalid: tools[0].tool_type: required one_of 'tool_type' must have one initialized field", file=sys.stderr)
        sys.exit(1)
    if not core or set(core) != {"read_file"}:
        print("Only the read-only compatibility tool may be declared", file=sys.stderr)
        sys.exit(1)
if mode == "auth":
    print("authentication required", file=sys.stderr)
    sys.exit(1)
if mode == "timeout":
    time.sleep(10)
if mode == "interrupt":
    import signal
    os.kill(os.getppid(), signal.SIGINT)
    time.sleep(10)
items = []
for row in payload["rows"]:
    source = row["source"]
    analysis = {
      "summary": "미사용 입력 핀 처리",
      "checks": [{"target_kind": "input_pin", "target_description": "미사용 입력",
      "part_numbers": [], "applicability": "미사용인 경우", "requirement": "GND 연결",
      "level": "mandatory", "parameters": [], "required_design_facts": ["connected_net"],
      "evidence": [{"field": "checkpoint", "quote": "GND"}]}], "issues": []}
    if mode == "bad_row" and "BAD" in source["checkpoint"]:
        del analysis["checks"][0]["level"]
    items.append({"request_id": row["request_id"], "analysis": analysis})
if mode == "omit_once" and len(log.read_text().splitlines()) == 1:
    items = items[:1]
print(json.dumps({"response": json.dumps({"items": items}), "stats": {}}))
'''


@unittest.skipIf(os.name == "nt", "test executable fixture uses a POSIX shebang")
class IntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.xlsx = self.folder / "source.xlsx"
        self.output = self.folder / "out"
        self.fake = self.folder / "gemini-test"
        self.fake.write_text(FAKE.replace("#!PYTHON", "#!" + sys.executable), encoding="utf-8")
        self.fake.chmod(0o700)
        self.calls = self.folder / "calls.jsonl"

    def tearDown(self):
        self.temp.cleanup()

    def run_script(self, *extra, mode=""):
        return subprocess.run([sys.executable, str(ROOT / "preprocess.py"), str(self.xlsx),
            "--output", str(self.output), "--gemini", str(self.fake), "--interval", "0", "--backoff", "0", *extra],
            capture_output=True, text=True, encoding="utf-8", timeout=45,
            env=dict(os.environ, CHECKPOINT_TEST_CALLS=str(self.calls), CHECKPOINT_TEST_MODE=mode))

    def results(self):
        return [json.loads(line) for line in (self.output / "checkpoints.jsonl").read_text().splitlines()]

    def test_dedup_resume_and_source_change_reprocess_only_changed_rule(self):
        rows = [["1", "0", "unused GND", "GND", "pin:", "f.pdf"],
                ["2", "0", "unused GND", "GND", "pin:", "f.pdf"]]
        make_xlsx(self.xlsx, rows)
        first = self.run_script()
        self.assertEqual(first.returncode, 0, first.stderr)
        self.assertEqual(len(self.calls.read_text().splitlines()), 1)
        self.assertEqual(len(json.loads(self.calls.read_text().splitlines()[0])), 1)
        results = self.results()
        self.assertEqual(len(results), 2)
        self.assertNotEqual(results[0]["analysis"]["checks"][0]["requirement_id"], results[1]["analysis"]["checks"][0]["requirement_id"])
        resumed = self.run_script()
        self.assertEqual(resumed.returncode, 0, resumed.stderr)
        self.assertEqual(len(self.calls.read_text().splitlines()), 1)
        rows[1][2] = "changed GND"
        make_xlsx(self.xlsx, rows)
        changed = self.run_script()
        self.assertEqual(changed.returncode, 0, changed.stderr)
        self.assertEqual(len(self.calls.read_text().splitlines()), 2)

    def test_missing_response_rows_are_retried_without_repeating_saved_rows(self):
        make_xlsx(self.xlsx, [["1", "0", "first GND", "", "pin:", ""], ["2", "0", "second GND", "", "pin:", ""]])
        result = self.run_script(mode="omit_once")
        self.assertEqual(result.returncode, 0, result.stderr)
        calls = [json.loads(line) for line in self.calls.read_text().splitlines()]
        self.assertEqual([len(c) for c in calls], [2, 1])
        self.assertNotEqual(calls[1][0], calls[0][0])

    def test_bad_single_row_does_not_lose_valid_rows_and_rerun_retries_failure(self):
        make_xlsx(self.xlsx, [["1", "0", "good GND", "", "pin:", ""], ["2", "0", "BAD GND", "", "pin:", ""]])
        result = self.run_script("--attempts", "1", mode="bad_row")
        self.assertEqual(result.returncode, 2)
        self.assertEqual([r["status"] for r in self.results()], ["normalized", "error"])
        result = self.run_script("--attempts", "1")
        self.assertEqual(result.returncode, 0, result.stderr)
        calls = [json.loads(line) for line in self.calls.read_text().splitlines()]
        self.assertEqual([len(c) for c in calls], [2, 1])

    def test_auth_failure_stops_without_retry_storm_and_preserves_pending_rows(self):
        make_xlsx(self.xlsx, [[str(i), "0", f"GND {i}", "", "", ""] for i in range(3)])
        result = self.run_script("--batch-size", "1", mode="auth")
        self.assertEqual(result.returncode, 2)
        self.assertEqual(len(self.calls.read_text().splitlines()), 1)
        self.assertEqual([r["status"] for r in self.results()], ["error", "pending", "pending"])
        self.assertFalse((self.output / "run.lock").exists())

    def test_cli_with_empty_tool_rejection_accepts_read_only_compatibility_settings(self):
        make_xlsx(self.xlsx, [["1", "0", "unused GND", "", "pin:", ""]])
        result = self.run_script("--attempts", "1", mode="validate_tools")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(self.results()[0]["status"], "normalized")
        self.assertEqual(len(self.calls.read_text().splitlines()), 1)

    def test_timeout_is_recorded_and_can_resume(self):
        make_xlsx(self.xlsx, [["1", "0", "GND", "", "", ""]])
        result = self.run_script("--timeout", "0.2", "--attempts", "1", mode="timeout")
        self.assertEqual(result.returncode, 2)
        self.assertEqual(self.results()[0]["status"], "error")
        self.assertIn("timeout", self.results()[0]["error"])
        self.assertEqual(self.run_script().returncode, 0)

    def test_keyboard_interrupt_exports_pending_and_releases_lock(self):
        make_xlsx(self.xlsx, [["1", "0", "GND", "", "", ""]])
        result = self.run_script(mode="interrupt")
        self.assertEqual(result.returncode, 130, result.stderr)
        self.assertEqual(self.results()[0]["status"], "pending")
        self.assertFalse((self.output / "run.lock").exists())
        self.assertEqual(self.run_script().returncode, 0)

    def test_12000_rows_dry_run_preserves_every_row_without_cli(self):
        make_xlsx(self.xlsx, [[str(i), "0", f"unused GND {i}", "", "pin:", ""] for i in range(12000)])
        result = self.run_script("--dry-run")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertFalse(self.calls.exists())
        results = self.results()
        self.assertEqual(len(results), 12000)
        self.assertEqual(results[-1]["excel_row"], 12001)
        self.assertTrue(all(r["status"] == "pending" for r in results))

    def test_txt_uses_cli_and_reuses_equivalent_xlsx_cache(self):
        row = ["1", "0", "unused GND", "GND", "pin:", "f.pdf"]
        make_xlsx(self.xlsx, [row])
        result = self.run_script()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.xlsx = self.folder / "source.txt"
        text = io.StringIO(newline="")
        writer = csv.writer(text)
        writer.writerow(["adminno", "seqno", "Checkpoint", "Specification", "설계치", "참조문서"])
        writer.writerow(row)
        self.xlsx.write_text(text.getvalue(), encoding="utf-8-sig")
        result = self.run_script()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(self.calls.read_text().splitlines()), 1)
        self.assertEqual(self.results()[0]["status"], "normalized")
        self.assertIsNone(self.results()[0]["sheet"])
        # A new CSV/TXT rule really goes through the subprocess, not just the cache.
        self.xlsx.write_text(text.getvalue().replace("unused GND", "changed GND"), encoding="utf-8-sig")
        self.assertEqual(self.run_script().returncode, 0)
        self.assertEqual(len(self.calls.read_text().splitlines()), 2)

    def test_12000_txt_rows_dry_run_preserves_every_row(self):
        self.xlsx = self.folder / "source.txt"
        text = io.StringIO(newline="")
        writer = csv.writer(text, delimiter="\t")
        writer.writerow(["adminno", "seqno", "Checkpoint", "Specification", "설계치", "참조문서"])
        for i in range(12000):
            writer.writerow([str(i), "0", f"입력 GND {i}", "", "pin:", ""])
        self.xlsx.write_bytes(text.getvalue().encode("cp949"))
        result = self.run_script("--dry-run")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertFalse(self.calls.exists())
        rows = self.results()
        self.assertEqual(len(rows), 12000)
        self.assertEqual(rows[-1]["record_row"], 12001)
        self.assertEqual(rows[-1]["input_encoding"], "cp949")


if __name__ == "__main__":
    unittest.main()
