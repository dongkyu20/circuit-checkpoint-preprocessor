import csv
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import preprocess as p

HEADERS = ["adminno", "seqno", "Checkpoint", "Specification", "설계치", "참조문서"]


def table(delimiter=",", checkpoint="미사용 핀은 GND 연결", adminno="00123"):
    output = io.StringIO(newline="")
    writer = csv.writer(output, delimiter=delimiter)
    writer.writerow(HEADERS)
    writer.writerow([adminno, "0", checkpoint, "GND", "Unused input pin:", '@첨부파일 "chip.pdf"'])
    return output.getvalue()


class TextInputTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def read(self, text, file_encoding="utf-8", suffix=".txt", **options):
        path = self.folder / ("rules" + suffix)
        path.write_bytes(text.encode(file_encoding))
        return p.read_input(path, "unused-sheet", 1, **options)

    def test_csv_saved_as_txt_preserves_multiline_quotes_and_identifiers(self):
        checkpoint = '조건 1, 조건 2\n/OE는 "LOW" | 한국어'
        rows = self.read(table(checkpoint=checkpoint), "utf-8-sig")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["source_columns"]["Checkpoint"], checkpoint)
        self.assertEqual(rows[0]["source_columns"]["adminno"], "00123")
        self.assertEqual(rows[0]["checkpoint_id"], "00123:0")
        self.assertIsNone(rows[0]["sheet"])
        self.assertIsNone(rows[0]["excel_row"])
        self.assertEqual(rows[0]["record_row"], 2)

    def test_tab_delimited_cp949_is_detected(self):
        rows = self.read(table("\t"), "cp949")
        self.assertEqual(rows[0]["normalized_source"]["checkpoint"], "미사용 핀은 GND 연결")
        self.assertEqual(rows[0]["input_format"], "csv")

    def test_utf16_excel_text_export_with_bom_is_detected(self):
        rows = self.read(table("\t"), "utf-16")
        self.assertEqual(rows[0]["normalized_source"]["reference_document"], '@첨부파일 "chip.pdf"')

    def test_semicolon_and_pipe_delimiters_are_detected_from_headers(self):
        for delimiter in (";", "|"):
            with self.subTest(delimiter=delimiter):
                self.assertEqual(len(self.read(table(delimiter))), 1)

    def test_manual_encoding_and_delimiter_override(self):
        rows = self.read(table("\t"), "cp949", encoding="cp949", delimiter="tab")
        self.assertEqual(len(rows), 1)
        with self.assertRaisesRegex(ValueError, "인코딩"):
            self.read(table(), "cp949", encoding="utf-8")

    def test_lost_table_structure_is_rejected_with_actionable_message(self):
        with self.assertRaisesRegex(ValueError, "CSV|구분자"):
            self.read("Checkpoint\nSpecification\n설계치\nGND\n")

    def test_unquoted_comma_and_broken_quoted_fields_are_rejected(self):
        header = ",".join(HEADERS)
        with self.assertRaisesRegex(ValueError, "열 수"):
            self.read(header + "\n1,0,first,second,GND,pin:,ref.pdf\n")
        with self.assertRaisesRegex(ValueError, "CSV"):
            self.read(header + '\n1,0,"never closes,GND,pin:,ref.pdf\n')

    def test_explicit_header_row_skips_preamble(self):
        path = self.folder / "rules.txt"
        path.write_text("설명 행\n" + table(), encoding="utf-8")
        rows = p.read_input(path, "unused-sheet", 2)
        self.assertEqual(rows[0]["record_row"], 3)

    def test_mismatched_manual_delimiter_fails_without_silent_fallback(self):
        with self.assertRaisesRegex(ValueError, "헤더|구분자|CSV"):
            self.read(table(), delimiter="tab")

    def test_txt_duplicate_required_header_shows_exact_positions(self):
        with self.assertRaisesRegex(ValueError, "Checkpoint.*1열.*4열"):
            self.read("Checkpoint,Specification,설계치,Checkpoint\nGND,GND,pin:,GND\n")

    def test_txt_duplicate_auxiliary_headers_are_not_lost(self):
        rows = self.read("Checkpoint,Specification,설계치,category,category\nGND,GND,pin:,a,b\n")
        self.assertEqual(rows[0]["source_columns"]["category [col 4]"], "a")
        self.assertEqual(rows[0]["source_columns"]["category [col 5]"], "b")

    def test_joined_txt_adminno_can_be_equal_or_selected_explicitly(self):
        text = "adminno,adminno,seqno,Checkpoint,Specification,설계치\n00123,00123,0,GND,GND,pin:\n"
        rows = self.read(text)
        self.assertEqual(rows[0]["checkpoint_id"], "00123:0")
        rows = self.read(text.replace("00123,00123", "00123,00456"), adminno_column=2)
        self.assertEqual(rows[0]["checkpoint_id"], "00456:0")

    def test_missing_file_error_includes_current_folder(self):
        with self.assertRaisesRegex(ValueError, "현재 폴더"):
            p.read_input(self.folder / "missing.txt", "sheet", 1)

    def test_long_cell_is_not_truncated(self):
        text = "GND " + "x" * 140000
        rows = self.read(table(checkpoint=text))
        self.assertEqual(rows[0]["source_columns"]["Checkpoint"], text)

    def test_txt_cli_dry_run_does_not_need_gemini_or_sheet(self):
        path = self.folder / "input.txt"
        path.write_text(table(), encoding="utf-8-sig")
        output = self.folder / "out"
        result = subprocess.run([sys.executable, str(ROOT / "preprocess.py"), str(path),
            "--output", str(output), "--dry-run", "--gemini", "DOES_NOT_EXIST", "--sheet", "irrelevant"],
            capture_output=True, text=True, encoding="utf-8", timeout=15)
        self.assertEqual(result.returncode, 0, result.stderr)
        row = json.loads((output / "prepared.jsonl").read_text(encoding="utf-8").splitlines()[0])
        self.assertEqual(row["source_columns"]["adminno"], "00123")
        summary = json.loads((output / "summary.json").read_text(encoding="utf-8"))
        self.assertEqual(summary["input_format"], "csv")
        self.assertIsNone(summary["sheet"])

    def test_cli_adminno_column_option_reaches_the_selected_join_key(self):
        path = self.folder / "joined.txt"
        path.write_text("adminno,adminno,seqno,Checkpoint,Specification,설계치\n00123,00456,0,GND,GND,pin:\n", encoding="utf-8")
        output = self.folder / "selected"
        result = subprocess.run([sys.executable, str(ROOT / "preprocess.py"), str(path),
            "--adminno-column", "2", "--output", str(output), "--dry-run"],
            capture_output=True, text=True, encoding="utf-8", timeout=15)
        self.assertEqual(result.returncode, 0, result.stderr)
        record = json.loads((output / "prepared.jsonl").read_text(encoding="utf-8").splitlines()[0])
        self.assertEqual(record["checkpoint_id"], "00456:0")
        self.assertEqual(record["adminno_selection"]["column"], 2)

    def test_cli_can_select_adminno_and_seqno_columns_together(self):
        path = self.folder / "joined-keys.txt"
        path.write_text("adminno,adminno,seqno,seqno,Checkpoint,Specification,설계치\n00123,00456,0,7,GND,GND,pin:\n", encoding="utf-8")
        output = self.folder / "selected-keys"
        result = subprocess.run([sys.executable, str(ROOT / "preprocess.py"), str(path),
            "--adminno-column", "2", "--seqno-column", "4", "--output", str(output), "--dry-run"],
            capture_output=True, text=True, encoding="utf-8", timeout=15)
        self.assertEqual(result.returncode, 0, result.stderr)
        row = json.loads((output / "prepared.jsonl").read_text(encoding="utf-8").splitlines()[0])
        self.assertEqual(row["checkpoint_id"], "00456:7")
        self.assertEqual(row["seqno_selection"], {"column": 4, "mode": "explicit"})


if __name__ == "__main__":
    unittest.main()
