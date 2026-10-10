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


class CsvImportSafetyTests(unittest.TestCase):
    def test_validation_and_review_precede_any_sheet_commit(self) -> None:
        worker = procedure("CSVImportExecute")
        ordered = [
            "AnalyzeCsvImportText",
            "CsvOpenTextAsWorkbook",
            "CsvValidateProtectedSamples",
            "previewText =",
            "ConfirmCsvImport(previewText, False)",
            "SaveCopyAs backupPath",
            "CsvSetProtectedTextFormats ws, formatBackup",
            ".Value2 = writeData",
            "CsvRangeMatches",
        ]
        offsets = [worker.index(token) for token in ordered]
        self.assertEqual(offsets, sorted(offsets))

    def test_no_orphaned_code_or_duplicate_worker_follows_import_procedure(self) -> None:
        match = re.search(r"Private Sub CSVImportExecute.*?End Sub", SOURCE, re.DOTALL)
        self.assertIsNotNone(match)
        assert match is not None
        following = next((line.strip() for line in SOURCE[match.end():].splitlines() if line.strip()), "")
        self.assertEqual(following, "Private Function ReadCsvTextWithDetectedEncoding( _")

    def test_cancel_exits_before_csv_import_call(self) -> None:
        entry = procedure("CSV取込_シンプル版")
        self.assertIn("Application.GetOpenFilename", entry)
        self.assertIn("If VarType(filePath) = vbBoolean Then Exit Sub", entry)
        self.assertLess(entry.index("GetOpenFilename"), entry.index("CSVImportExecute"))

    def test_normal_preview_is_compact_without_removing_safety_checks(self) -> None:
        worker = procedure("CSVImportExecute")
        preview = worker[worker.index("    previewText = "):worker.index("    If Not ConfirmCsvImport(previewText, False)")]
        for token in ("ファイル：", "対象年度：", "の登録件数", "取込前：", "取込後：", "同じ年度のデータを入れ替えます。実行しますか？"):
            self.assertIn(token, preview)
        for token in ("文字コード", "新しいCSVの取込予定件数", "CsvImportIssueSummary", "識別列の文字列保護"):
            self.assertNotIn(token, preview)
        self.assertIn("CSV取込の確認", procedure("ConfirmCsvImport"))
        self.assertIn("CsvImportIssueSummary(stats)", worker[:worker.index("    previewText = ")])
        self.assertIn("CsvValidateProtectedSamples", worker)

    def test_existing_count_uses_this_workbook_and_shared_year_normalization(self) -> None:
        worker = procedure("CSVImportExecute")
        count = worker[worker.index("    existingTargetCount = 0"):worker.index("    requiredHeaders = ")]
        self.assertIn('ThisWorkbook.Worksheets("目録CSV")', worker)
        self.assertIn("CsvGetUnfilteredImportBounds ws, dstLastRow, dstLastCol, dstValueLastCol", worker)
        self.assertNotIn("dstLastRow = GetLastUsedRow(ws)", worker)
        self.assertNotIn("dstLastCol = GetLastHeaderCol(ws)", worker)
        self.assertIn('FindHeaderColumn(ws, dstLastCol, Array("年度（和暦）", "年度(和暦)"))', count)
        self.assertIn('FindHeaderColumn(ws, dstLastCol, Array("年度"))', count)
        self.assertIn("GetNormalizedYearFromRow(ws, oldRow, existingWarekiCol, existingWesternCol)", count)
        self.assertIn("If existingYearValue = targetYearValue Then existingTargetCount = existingTargetCount + 1", count)
        normalizer = procedure("NormalizeYearValue")
        self.assertIn("2018 + n", normalizer)
        self.assertIn("If n >= 1900 And n <= 2999 Then", normalizer)

    def test_import_bounds_scan_all_values_including_filtered_rows(self) -> None:
        worker = procedure("CSVImportExecute")
        bounds = procedure("CsvGetUnfilteredImportBounds")
        self.assertIn("CsvGetUnfilteredImportBounds ws, dstLastRow, dstLastCol, dstValueLastCol", worker)
        self.assertIn("Set usedCells = ws.UsedRange", bounds)
        self.assertIn("CountA(block)", bounds)
        self.assertIn("values = CsvReadRange2D(block)", bounds)
        self.assertIn("CsvValueHasContent(values(r, c))", bounds)
        self.assertIn("If blocksVisited > 10000 Then", bounds)
        self.assertIn("If CDbl(boundRow) * CDbl(boundCol) > 50000000# Then", bounds)
        for forbidden in (".Find(", ".End(", "ShowAllData", "AutoFilterMode = False", "lastDataRow = boundRow"):
            self.assertNotIn(forbidden, bounds)
        self.assertIn("If dstValueLastCol > dstLastCol Then Err.Raise", worker)

    def test_filtered_write_and_restore_address_hidden_rows_without_clearing_filter(self) -> None:
        worker = procedure("CSVImportExecute")
        self.assertIn("filterWasActive = ws.FilterMode", worker)
        self.assertIn("CsvWriteRowsByPosition ws, writeData", worker)
        self.assertIn("CsvClearRowsByPosition ws, outputDataRows + 2", worker)
        writer = procedure("CsvWriteRowsByPosition")
        self.assertIn("For r = 1 To lastRow", writer)
        self.assertIn("ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Value2 = rowValues", writer)
        restore = procedure("CsvRestoreImportSnapshot")
        self.assertIn("CsvClearRowsByPosition ws, 2, clearRows, clearCols", restore)
        self.assertIn("CsvWriteRowsByPosition ws, backupValues, oldLastRow, oldLastCol", restore)
        for forbidden in ("ShowAllData", "AutoFilterMode = False", "AutoFilter.Apply"):
            self.assertNotIn(forbidden, worker + writer + restore)

    def test_zero_count_diagnostic_is_limited_to_existing_rows_and_does_not_block_first_import(self) -> None:
        worker = procedure("CSVImportExecute")
        diagnostic = procedure("CsvShowZeroCountDiagnostic")
        self.assertIn("If existingTargetCount = 0 And examinedRowCount > 0 Then", worker)
        self.assertLess(worker.index("CsvShowZeroCountDiagnostic ws"), worker.index("ConfirmCsvImport(previewText, False)"))
        for field in ("ThisWorkbook.FullName", "initialActiveBookPath", "ActiveWorkbook.FullName", "ws.Name", "lastRow", "lastCol", "warekiCol", "westernCol", "examinedCount", "recognizedCount"):
            self.assertIn(field, diagnostic)
        self.assertNotIn("Err.Raise", diagnostic)
        self.assertNotIn("CsvWriteImportLog", diagnostic)

    def test_count_and_replacement_must_agree_before_any_sheet_write(self) -> None:
        worker = procedure("CSVImportExecute")
        check = worker.index("If deletedCount <> existingTargetCount Then Err.Raise")
        self.assertLess(worker.index("If existingYearValue = targetYearValue Then existingTargetCount = existingTargetCount + 1"), check)
        self.assertLess(worker.index("deletedCount = deletedCount + 1"), check)
        self.assertLess(check, worker.index("SaveCopyAs backupPath"))
        self.assertLess(check, worker.index(".Value2 = writeData"))
        self.assertIn("GetNormalizedYearFromRow(ws, oldRow, existingWarekiCol, existingWesternCol)", worker)

    def test_large_count_reduction_warns_but_does_not_forbid_import(self) -> None:
        worker = procedure("CSVImportExecute")
        self.assertIn("If existingTargetCount >= 10 And importCount * 2 <= existingTargetCount Then", worker)
        self.assertIn("ConfirmCsvImport(reductionWarning, True)", worker)
        self.assertNotIn("Err.Raise", worker[worker.index("If existingTargetCount >= 10"):worker.index("If hasHeader Then", worker.index("If existingTargetCount >= 10"))])

    def test_recoverable_backup_precedes_write_and_failure_is_not_success(self) -> None:
        worker = procedure("CSVImportExecute")
        self.assertIn("Environ(\"LOCALAPPDATA\")", worker)
        self.assertIn("CSVImportBackups", worker)
        self.assertLess(worker.index("SaveCopyAs backupPath"), worker.index(".Value2 = writeData"))
        self.assertIn("CsvRestoreImportSnapshot(ws, backupValues, formatBackup", worker)
        self.assertIn("restore=FAILED;backup=", worker)
        restore = procedure("CsvRestoreImportSnapshot")
        self.assertIn("CsvRangeMatches", restore)
        self.assertIn("CsvFormatSnapshotMatches", restore)

    def test_identifier_columns_are_loaded_and_verified_as_exact_text(self) -> None:
        self.assertIn("FieldInfo:=fieldInfo", procedure("CsvOpenTextAsWorkbook"))
        self.assertIn("Array(col, xlTextFormat)", procedure("CsvOpenTextAsWorkbook"))
        identifier = procedure("CsvHeaderIsIdentifier")
        for token in ("ID", "コード", "番号"):
            self.assertIn(token, identifier)
        worker = procedure("CSVImportExecute")
        self.assertIn("CsvValidateProtectedSamples protectedSamples, textColumns, srcLastRow, srcLastCol", worker)
        restore = procedure("CsvApplyProtectedSamples")
        self.assertIn("sourceValues(rowNo - 1, col) = CStr(protectedSamples(sampleKey))", restore)
        self.assertIn("CsvRangeMatches", worker)

    def test_numeric_date_and_formula_like_values_are_protected_or_rejected(self) -> None:
        self.assertIn("Len(valueText) <= 15", procedure("CsvIsNumericIdentifierRisk"))
        self.assertIn("IsDate(valueText)", procedure("CsvStringNeedsText"))
        record = procedure("CsvProcessImportRecord")
        self.assertIn("If CsvLooksLikeFormulaText(rawValue) Then", record)
        self.assertIn("CsvAddFatal stats", record)
        worker = procedure("CSVImportExecute")
        self.assertIn("CsvConvertDateColumns sourceValues, headers", worker)
        dateConversion = procedure("CsvConvertDateColumns")
        self.assertIn("CsvHeaderIsDate(headerText)", dateConversion)
        self.assertIn("sourceValues(rowNo, col) = CDbl(CDate(valueText))", dateConversion)
        formats = procedure("CsvBuildFormatBackup")
        self.assertIn("CsvRememberNumberFormat", formats)
        self.assertIn("CsvRestoreNumberFormats ws, formatBackup", worker)

    def test_year_and_key_validation_covers_requested_failures(self) -> None:
        record = procedure("CsvProcessImportRecord")
        for text in (
            "年度列がありません",
            "年度が空欄",
            "年度を解釈できない",
            "複数年度が混在",
            "必須見出し",
            "ファイルIDが空欄",
            "ファイルIDが重複",
            "ファイル番号が重複",
            "バーコード番号が重複",
        ):
            self.assertIn(text, record)

    def test_multiline_headers_use_utf8_stage_while_csv_body_is_preserved(self) -> None:
        stage = procedure("CsvPrepareStageCopy")
        self.assertIn("If Not needsFlatten Then Exit Sub", stage)
        self.assertIn("If inQuotes And (ch = vbCr Or ch = vbLf) Then", stage)
        self.assertIn("stageText = Left$(safeBuffer, safeLength) & vbCrLf & Left$(bodyBuffer, bodyLength)", stage)
        self.assertIn('stream.Charset = "utf-8"', stage)
        self.assertIn("stageCodePage = 65001", stage)
        self.assertIn("bodyBuffer = String$(Len(bodyText), vbNullChar)", stage)
        self.assertIn("ElseIf inQuotes And (ch = vbCr Or ch = vbLf) Then", stage)
        self.assertIn("If Len(stagePath) > 0 Then If Len(Dir(stagePath)) > 0 Then Kill stagePath", stage)
        process = procedure("CsvProcessImportRecord")
        self.assertIn('multilineSamples(CStr(col) & "|" & CStr(dataRow)) = rawValue', process)
        worker = procedure("CSVImportExecute")
        self.assertIn("CsvPrepareStageCopy csvText, headers, originCodePage, stagePath, stageCodePage", worker)
        self.assertIn('CsvApplyMultilineSamples sourceValues, stats("MultilineSamples")', worker)
        multiline = procedure("CsvApplyMultilineSamples")
        self.assertIn("sourceValues(rowNo - 1, col) = CStr(multilineSamples(sampleKey))", multiline)
        self.assertIn("If Len(stagePath) > 0 Then", worker)
        self.assertIn("If Len(stagePath) > 0 Then", worker[worker.index("CleanExit:"):])
        self.assertIn("If Len(Dir(stagePath)) > 0 Then Kill stagePath", worker[worker.index("CleanExit:"):])

    def test_import_log_includes_existing_deleted_import_counts(self) -> None:
        log = procedure("CsvWriteImportLog")
        for field in ("既存件数=", "削除件数=", "取込件数="):
            self.assertIn(field, log)
        self.assertIn("WriteProcessLog", log)

    def test_year_replacement_retains_other_rows_and_inserts_in_year_order(self) -> None:
        worker = procedure("CSVImportExecute")
        self.assertIn("If existingYearValue = targetYearValue Then", worker)
        self.assertIn("If Not inserted And existingYearValue > targetYearValue Then", worker)
        self.assertIn("If Not inserted Then", worker)
        self.assertNotIn("ApplyAnnualFolderNameCorrections", worker)
        self.assertNotIn("AnnualFolderNameCorrections", worker)


if __name__ == "__main__":
    unittest.main()
