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
    def test_step3_uses_source_data_columns_to_bound_rows_not_find_state(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertIn("srcLastRow = GetLastMokurokuDataRow(wsSrc, class2CodeCol, class2NameCol, class3CodeCol, class3NameCol)", step3)
        self.assertNotIn("srcLastRow = GetLastUsedRow(wsSrc)", step3)
        self.assertLess(step3.index("ValidateGuideCodeSourceHeaders"), step3.index("GetLastMokurokuDataRow"))
        self.assertLess(step3.index("GetLastMokurokuDataRow"), step3.index("Set pendingGuideRows"))

    def test_step3_data_bound_checks_a_through_h_and_guide_columns(self) -> None:
        helper = procedure("GetLastMokurokuDataRow")
        self.assertIn("For col = 1 To 8", helper)
        self.assertIn("class2CodeCol", helper)
        self.assertIn("class2NameCol", helper)
        self.assertIn("class3CodeCol", helper)
        self.assertIn("class3NameCol", helper)
        self.assertIn(".End(xlUp).Row", helper)
        self.assertNotIn(".Find(", helper)

    def test_step3_empty_source_diagnostic_identifies_macro_workbook_and_sheet(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertIn("ThisWorkbook.FullName", step3)
        self.assertIn("wsSrc.Name", step3)
        self.assertIn("GetLastUsedRow(wsSrc)", step3)

    def test_step3_reports_find_disagreement_without_changing_row_bound(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertIn("findLastRow = GetLastUsedRow(wsSrc)", step3)
        self.assertIn("If findLastRow <> srcLastRow Then", step3)
        self.assertIn("行数診断", step3)
        self.assertLess(step3.index("srcLastRow = GetLastMokurokuDataRow"), step3.index("findLastRow = GetLastUsedRow(wsSrc)"))
        self.assertLess(step3.index("findLastRow = GetLastUsedRow(wsSrc)"), step3.index("BuildSystemGuideCodeMasterFromMokuroku"))

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


class SerialAssignmentRegressionTests(unittest.TestCase):
    def test_automatic_reindex_uses_append_only_core(self) -> None:
        automatic = procedure("ReindexSerial_NewFileStandard")
        self.assertIn("AssignMissingSerial_NewFileStandardCore(targetWs)", automatic)
        self.assertNotIn("ReindexSerial_NewFileStandardCore(targetWs)", automatic)

    def test_automatic_core_preserves_existing_values_and_uses_maximum_plus_one(self) -> None:
        automatic_core = procedure("AssignMissingSerial_NewFileStandardCore")
        self.assertNotIn("ClearContents", automatic_core)
        self.assertIn("GetMaximumSerial_NewFileStandard", automatic_core)
        self.assertIn("IsSerialCellBlank", automatic_core)
        self.assertIn("ShouldAssignSerialByRow", automatic_core)
        self.assertIn("nextSerial = nextSerial + 1", automatic_core)
        self.assertIn("Dim nextSerial As Double", automatic_core)
        self.assertIn("nextSerial >= 999999999999999#", automatic_core)
        self.assertIn("ByRef serialValue As Double", procedure("TryGetPositiveSerialValue"))
        self.assertNotIn("2147483647", procedure("TryGetPositiveSerialValue"))
        self.assertLess(
            automatic_core.index("prevCalculation = Application.Calculation"),
            automatic_core.index("If targetWs Is Nothing Then"),
        )

    def test_manual_reindex_keeps_full_renumbering_core(self) -> None:
        manual = procedure("新ファイル基準表_通し番号_手動再採番")
        full_reindex = procedure("ReindexSerial_NewFileStandardCore")
        self.assertIn("ReindexSerial_NewFileStandardCore(targetWs)", manual)
        self.assertIn("ClearContents", full_reindex)

    def test_append_only_serial_model_covers_requested_cases(self) -> None:
        def assign_missing(serials: list[object], eligible: list[bool], visible: list[bool]) -> list[object]:
            del visible  # Hidden rows still participate in the maximum.
            maximum = max((value for value in serials if isinstance(value, int) and value > 0), default=0)
            result = list(serials)
            for index, should_assign in enumerate(eligible):
                if should_assign and result[index] in (None, ""):
                    maximum += 1
                    result[index] = maximum
            return result

        self.assertEqual(assign_missing([872, 873, 874, 875, None], [False, False, False, False, True], [True] * 5), [872, 873, 874, 875, 876])
        self.assertEqual(assign_missing([952, None, None, None], [False, True, True, True], [True] * 4), [952, 953, 954, 955])
        self.assertEqual(assign_missing([1, 2, 5, 10, None], [False, False, False, False, True], [True] * 5), [1, 2, 5, 10, 11])
        self.assertEqual(assign_missing([952, None], [False, True], [False, True]), [952, 953])
        self.assertEqual(assign_missing([874], [True], [True]), [874])
        self.assertEqual(assign_missing([953], [True], [True]), [953])
        self.assertEqual(assign_missing([None], [True], [True]), [1])


class DiffImportHeaderRegressionTests(unittest.TestCase):
    def test_step4_uses_header_row_bounds_for_both_key_sheets(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("srcLastCol = GetLastHeaderCol(wsSrc)", step4)
        self.assertIn("editLastCol = GetLastHeaderCol(wsEdit)", step4)
        self.assertNotIn("srcLastCol = GetLastUsedCol(wsSrc)", step4)
        self.assertNotIn("editLastCol = GetLastUsedCol(wsEdit)", step4)
        confirm_candidates = procedure("ConfirmAndMarkStep4NewRegistrationCandidates")
        self.assertIn("lastCol = GetLastHeaderCol(wsEdit)", confirm_candidates)
        self.assertNotIn("lastCol = GetLastUsedCol(wsEdit)", confirm_candidates)

    def test_step3_uses_header_row_bound_for_code_management_csv(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertIn("srcLastCol = GetLastHeaderCol(wsSrc)", step3)
        self.assertNotIn("srcLastCol = GetLastUsedCol(wsSrc)", step3)

    def test_step4_missing_key_error_identifies_columns_and_header_context(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("BuildDiffKeyMissingHeaders", step4)
        self.assertIn("BuildHeaderSearchDiagnostics", step4)
        helper = procedure("BuildDiffKeyMissingHeaders")
        for expected in ("年度（和暦）または年度", "タイトル", "分類名２", "媒体種別", "保存期間"):
            self.assertIn(expected, helper)


if __name__ == "__main__":
    unittest.main()
