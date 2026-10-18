"""Source-level contracts for best-effort row last-editor metadata tracking."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8")
SHEET = (ROOT / "vba/新ファイル基準表.txt").read_text(encoding="utf-8")
REVIEW = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")


def procedure(source: str, name: str, kind: str = "Sub") -> str:
    match = re.search(
        rf"(?ms)^(?:Private|Public) {kind} {re.escape(name)}\b.*?^End {kind}",
        source,
    )
    if match is None:
        raise AssertionError(f"missing {kind}: {name}")
    return match.group()


class LastEditTrackingTests(unittest.TestCase):
    def test_sheet_change_adds_tracking_without_removing_existing_handlers(self):
        event = procedure(SHEET, "Worksheet_Change")
        for existing_call in (
            "新ファイル基準表_表示名からID自動補完 Target",
            "新ファイル基準表_廃棄年月日自動計算 Target",
            "新ファイル基準表_登録予定識別番号を自動クリア Target, Me",
            "ShouldReindexSerial(Target, Me)",
            "ReindexSerial_NewFileStandard Me",
        ):
            self.assertIn(existing_call, event)
        tracker = "新ファイル基準表_最終編集情報を記録 Target, Me"
        self.assertIn(tracker, event)
        self.assertLess(event.index("Application.EnableEvents = False"), event.index(tracker))
        self.assertIn("oldEnableEvents = Application.EnableEvents", event)
        self.assertIn("Application.EnableEvents = oldEnableEvents", event)

    def test_tracker_is_header_based_manual_only_rows_and_safe_for_events(self):
        tracker = procedure(MODULE, "新ファイル基準表_最終編集情報を記録")
        for header in ("最終編集者", "最終編集日時"):
            self.assertIn(f'"{header}"', tracker)
        self.assertIn("LastEditFindUniqueColumn", tracker)
        self.assertIn("LastEditAreaHasBusinessChanges(area, editorCol, timestampCol)", tracker)
        self.assertIn("Application.CutCopyMode = xlCut", tracker)
        self.assertIn("editorCol <= 0 Or timestampCol <= 0", tracker)
        self.assertIn("For Each area In Target.Areas", tracker)
        self.assertIn("rowsToStamp.Exists", tracker)
        self.assertIn("If rowNo > 1 Then", tracker)
        self.assertIn('Environ$("Username")', tracker)
        self.assertIn("If Len(userName) = 0 Then", tracker)
        self.assertNotIn('"同期状態"', tracker)
        self.assertNotIn('"登録予定"', tracker)
        self.assertIn("editTimestamp = Now", tracker)
        self.assertIn('NumberFormat = "yyyy/mm/dd hh:mm:ss"', tracker)
        self.assertIn("stampRange.Value2 = editTimestamp", tracker)
        self.assertIn("processedRows", tracker)
        self.assertIn("previousEnableEvents = Application.EnableEvents", tracker)
        self.assertIn("Application.EnableEvents = previousEnableEvents", tracker)
        self.assertIn("On Error GoTo", tracker)
        self.assertNotIn("Application.EnableEvents = True", tracker)

    def test_tracking_ignores_metadata_only_ranges_and_finds_unique_headers(self):
        area_filter = procedure(MODULE, "LastEditAreaHasBusinessChanges", "Function")
        self.assertIn("changedArea.Columns.Count = 1", area_filter)
        self.assertIn("changedArea.Columns.Count = 2", area_filter)
        self.assertIn("editorCol", area_filter)
        self.assertIn("timestampCol", area_filter)
        unique = procedure(MODULE, "LastEditFindUniqueColumn", "Function")
        self.assertIn("GetLastHeaderCol", unique)
        self.assertIn("matchCount > 1", unique)
        self.assertIn("LastEditFindUniqueColumn = 0", unique)

    def test_step2_protects_both_internal_fields_and_only_transfers_mapped_columns(self):
        step2 = procedure(MODULE, "手順2_新ファイル基準表を作成する")
        for header in ("最終編集者", "最終編集日時"):
            self.assertIn(f'"{header}"', step2)
        self.assertIn("IsUserManagedHeader(canonical, srcHeaderMap, baselineUserManagedHeaders)", procedure(MODULE, "BuildStep2SourceTransferMap"))
        update = procedure(MODULE, "UpdateRowFromSourceByHeader")
        append = procedure(MODULE, "AppendRowFromSourceByHeader")
        self.assertIn("srcToDstMap.Keys", update)
        self.assertIn("srcToDstMap.Keys", append)
        self.assertNotIn('"最終編集者"', append)
        self.assertNotIn('"最終編集日時"', append)
        self.assertNotIn(".Rows(", append)
        self.assertIn("appendRow = dstLastRow + 1", step2)
        last_row = procedure(MODULE, "Step2GetLastDataRow", "Function")
        bounds = procedure(MODULE, "CsvGetUnfilteredImportBounds")
        self.assertIn("CsvGetUnfilteredImportBounds ws, Step2GetLastDataRow", last_row)
        self.assertIn("boundCol = usedCells.Column + usedCells.Columns.Count - 1", bounds)
        self.assertIn("If chunkStart + r - 1 > lastDataRow", bounds)
        code_management = procedure(MODULE, "手順3_コード管理CSVを作成する")
        self.assertIn("SHEET_MOKUROKU_CSV", code_management)
        self.assertNotIn('"最終編集者"', code_management)
        self.assertNotIn('"最終編集日時"', code_management)

    def test_nplus1_clears_cloned_editor_metadata_without_stamping(self):
        nplus1 = procedure(MODULE, "nPlus1_次年度フォルダ生成")
        clear = procedure(MODULE, "ClearIdentityColumnsForNewRow")
        copy_at = nplus1.index("Resize(1, lastCol).Value =")
        clear_at = nplus1.index("ClearIdentityColumnsForNewRow wsDst, appendRow")
        self.assertLess(copy_at, clear_at)
        self.assertIn('"最終編集者"', clear)
        self.assertIn('"最終編集日時"', clear)
        self.assertIn("Application.EnableEvents = False", nplus1)

    def test_step5_never_exports_internal_editor_columns(self):
        export = procedure(MODULE, "手順5_差分インポートCSVを外部出力する")
        self.assertIn('headerName <> "最終編集者"', export)
        self.assertIn('headerName <> "最終編集日時"', export)
        self.assertIn('headerName <> "通し番号"', export)
        self.assertIn('headerName <> "状態"', export)

    def test_automatic_color_and_annual_correction_writes_suppress_change_events(self):
        for name in (
            "SyncSecondColorNameFromColor",
            "新ファイル基準表_色名から2色名同期",
            "nプラス1_毎年作成年度補正を実行する",
        ):
            with self.subTest(procedure=name):
                body = procedure(MODULE, name)
                self.assertIn("Application.EnableEvents = False", body)
                self.assertRegex(body, r"(?i)Application\.EnableEvents\s*=\s*(?:oldEnableEvents|previousEnableEvents)")

    def test_sync_review_signatures_ignore_editor_metadata_but_keep_other_fields(self):
        helper = procedure(REVIEW, "ReviewIsInternalMetadataHeader", "Function")
        self.assertIn('"最終編集者"', helper)
        self.assertIn('"最終編集日時"', helper)
        signature = procedure(REVIEW, "ReviewSignature", "Function")
        self.assertIn("ReviewIsInternalMetadataHeader", signature)
        self.assertIn("signatureLastCol", signature)
        self.assertIn("NumberFormat", signature)
        sheet_signature = procedure(REVIEW, "ReviewCopySheetSignature", "Function")
        self.assertIn("ReviewIsInternalMetadataHeader", sheet_signature)
        self.assertIn("ReviewSignature(ws, rowNo)", sheet_signature)


if __name__ == "__main__":
    unittest.main()
