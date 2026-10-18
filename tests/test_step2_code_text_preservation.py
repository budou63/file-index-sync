"""Contracts for preserving formal CSV code text through Steps 1, 2, and allocation."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def proc(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE,
        re.S,
    )
    if not match:
        raise AssertionError(f"missing procedure: {name}")
    return match.group(0)


def write_csv_code(header: str, raw: str) -> tuple[str, str]:
    text_headers = {"ファイル番号", "分類コード２", "分類コード３", "汎用項目１０"}
    return ("@", raw) if header in text_headers else ("unchanged", raw)


class Step2CodeTextPreservationTests(unittest.TestCase):
    def test_required_formal_code_values_are_text_not_numeric(self):
        for header, raw in (("分類コード２", "0321"), ("分類コード３", "0805"), ("汎用項目１０", "090200"), ("分類コード２", "0000")):
            self.assertEqual(write_csv_code(header, raw), ("@", raw))

    def test_only_explicit_code_headers_are_forced_to_text(self):
        self.assertEqual(write_csv_code("年度", "2026"), ("unchanged", "2026"))
        self.assertEqual(write_csv_code("保存期間", "10"), ("unchanged", "10"))

    def test_step1_import_protects_generic_item_10_and_step2_preserves_all_three_values(self):
        importer = proc("CsvHeaderIsIdentifier")
        writer = proc("WriteStep2SourceValue")
        required = proc("Step2HeaderRequiresSourceText")
        self.assertIn('CanonicalHeaderName("汎用項目１０")', importer)
        for header in ("分類コード２", "分類コード３", "汎用項目１０"):
            self.assertIn(f'CanonicalHeaderName("{header}")', required)
        self.assertIn("Step2HeaderRequiresSourceText", writer)
        self.assertIn('targetCell.NumberFormat = "@"', writer)
        self.assertIn("targetCell.Value2 = CStr(sourceValue)", writer)

    def test_second_identical_sync_is_not_changed_after_first_text_write(self):
        stored_format, stored = write_csv_code("分類コード２", "0321")
        self.assertEqual((stored_format, stored), ("@", "0321"))
        self.assertEqual(stored, "0321")
        self.assertNotEqual(stored, "321")
        exporter = proc("手順5_差分インポートCSVを外部出力する")

        self.assertIn('If headerName <> "通し番号" And headerName <> "状態" And _', exporter)
        self.assertIn('headerName <> "最終編集者" And headerName <> "最終編集日時"', exporter)
        self.assertIn("Step2HeaderRequiresSourceText(headerName)", exporter)
        self.assertIn('wsExport.Cells(1, dstCol).Resize(lastRow, 1).NumberFormat = "@"', exporter)
        self.assertIn("wsOut.Cells(usedRng.Row, absCol).Resize(lastRow, 1).Value", exporter)

    def test_actual_value_change_remains_distinguishable(self):
        self.assertNotEqual(write_csv_code("分類コード３", "0805")[1], write_csv_code("分類コード３", "0806")[1])

    def test_classification_recompletion_and_allocation_write_code_cells_as_text_without_invented_padding(self):
        allocation = proc("新ファイル基準表_登録予定分類コード採番")
        refill = proc("新ファイル基準表_分類コード再補完")
        writer = proc("WriteClassificationCodeValue")
        self.assertIn("WriteClassificationCodeValue", allocation)
        self.assertIn("WriteClassificationCodeValue", refill)
        self.assertIn('CanonicalHeaderName("分類コード２")', writer)
        self.assertIn('CanonicalHeaderName("分類コード３")', writer)
        self.assertIn('targetCell.NumberFormat = "@"', writer)
        self.assertIn("targetCell.Value2 = CStr(codeValue)", writer)
        self.assertIn("targetCell.Value = codeValue", writer)
        self.assertNotIn("Format$(", writer)
        self.assertNotIn("Right$(", writer)


if __name__ == "__main__":
    unittest.main()
