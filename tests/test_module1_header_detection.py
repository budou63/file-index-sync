from __future__ import annotations

from pathlib import Path
import re
import unittest


MODULE = Path(__file__).resolve().parents[1] / "Module1"
SOURCE = MODULE.read_text(encoding="utf-8")


def procedure(name: str) -> str:
    pattern = re.compile(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        re.DOTALL,
    )
    match = pattern.search(SOURCE)
    if match is None:
        raise AssertionError(f"{name} が見つかりません")
    return match.group(0)


class HeaderDetectionRegressionTests(unittest.TestCase):
    def test_step2_uses_header_row_bound_for_source_year_search(self) -> None:
        step2 = procedure("手順2_新ファイル基準表を作成する")
        self.assertIn("srcLastCol = GetLastHeaderCol(wsSrc)", step2)
        self.assertIn(
            'srcWarekiCol = FindHeaderColumn(wsSrc, srcLastCol, Array("年度（和暦）", "年度(和暦)"))',
            step2,
        )
        self.assertIn('srcYearCol = FindHeaderColumn(wsSrc, srcLastCol, Array("年度"))', step2)

    def test_header_bound_handles_sparse_header_rows_and_ignores_blank_formula_results(self) -> None:
        helper = procedure("GetLastHeaderCol")
        self.assertIn("ws.Rows(headerRow).Find", helper)
        self.assertIn("LookIn:=xlValues", helper)
        self.assertIn("SearchOrder:=xlByColumns", helper)
        self.assertIn("SearchDirection:=xlPrevious", helper)
        self.assertIn("SearchFormat:=False", helper)

    def test_csv_import_uses_header_row_bound_for_header_mapping(self) -> None:
        csv_import = procedure("CSV取込_シンプル版")
        self.assertIn("srcLastCol = GetLastHeaderCol(wsCSV)", csv_import)
        self.assertIn("dstLastCol = GetLastHeaderCol(ws)", csv_import)
        self.assertIn("dstLastCol = GetLastHeaderCol(ws) + 1", csv_import)

    def test_missing_year_diagnostic_identifies_the_actual_workbook_and_headers(self) -> None:
        step2 = procedure("手順2_新ファイル基準表を作成する")
        self.assertIn("BuildHeaderSearchDiagnostics", step2)
        diagnostics = procedure("BuildHeaderSearchDiagnostics")
        for expected in (
            "ThisWorkbook.Name",
            "ThisWorkbook.FullName",
            "GetLastUsedCol: ",
            "見出し行最終列: ",
            "UsedRange: ",
            "最終見出しセル: ",
            '"A1: "',
            '"B1: "',
            '"C1: "',
            '"D1: "',
            '"E1: "',
            "認識した見出し: ",
        ):
            self.assertIn(expected, diagnostics)


if __name__ == "__main__":
    unittest.main()
