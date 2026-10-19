import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ROW_ID = (ROOT / "vba/modExcelRowId.txt").read_text(encoding="utf-8")
MODULE1 = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8")
SHEET = (ROOT / "vba/新ファイル基準表.txt").read_text(encoding="utf-8")
REVIEW = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")


def procedure(source, name):
    match = re.search(
        r"(?ms)^(?:Public|Private) (?:Sub|Function) "
        + re.escape(name)
        + r"\b.*?^End (?:Sub|Function)",
        source,
    )
    if not match:
        raise AssertionError(f"procedure not found: {name}")
    return match.group(0)


class ExcelManagementRowIdTests(unittest.TestCase):
    def test_module_is_one_cp932_representable_paste_source(self):
        self.assertTrue(ROW_ID.startswith("Option Explicit"))
        self.assertNotIn("Attribute VB_Name", ROW_ID)
        self.assertEqual(ROW_ID.encode("cp932").decode("cp932"), ROW_ID)
        names = re.findall(r"(?m)^(?:Public|Private) (?:Sub|Function) (\w+)", ROW_ID)
        self.assertEqual(len(names), len(set(names)))
        self.assertIn("Private Declare PtrSafe Function CoCreateGuid Lib \"ole32.dll\"", ROW_ID)

    def test_initial_assignment_is_explicit_validated_backed_up_and_idempotent(self):
        body = procedure(ROW_ID, "Excel管理行ID_既存行へ初回付与")
        self.assertIn("ExcelRowIdInspect ws", body)
        self.assertIn("If invalidCount > 0 Or unexpectedCount > 0 Or duplicateCount > 0 Then", body)
        self.assertIn("If MsgBox(", body)
        self.assertLess(body.index("ExcelRowIdCreateBackup()"), body.index(".Value2 = CStr(assignments"))
        self.assertIn("If missingRows.Count = 0 Then", body)
        self.assertIn("既存IDは変更しません", body)
        self.assertIn("ClearContents", body)
        self.assertIn("ExcelRowIdCellIsBlank", body)
        self.assertIn("ExcelRowIdBackups", procedure(ROW_ID, "ExcelRowIdCreateBackup"))
        self.assertIn("ThisWorkbook.SaveCopyAs backupPath", procedure(ROW_ID, "ExcelRowIdCreateBackup"))

    def test_header_resolution_rejects_missing_or_ambiguous_columns_without_inserting(self):
        body = procedure(ROW_ID, "ExcelRowIdRequireColumn")
        self.assertIn("If matchCount = 0 Then Err.Raise", body)
        self.assertIn("If matchCount > 1 Then Err.Raise", body)
        self.assertNotIn("Cells(1, lastCol + 1)", body)
        self.assertIn("Private Const ROW_ID_HEADER As String = \"Excel管理行ID\"", ROW_ID)
        self.assertNotIn("ROW_ID_HEADER_ALIAS", ROW_ID)

    def test_uuid_validation_and_batch_generation_do_not_use_serial_numbers(self):
        self.assertIn("CoCreateGuid(value)", procedure(ROW_ID, "ExcelRowIdCreateGuid"))
        self.assertIn("InStr(1, \"0123456789abcdefABCDEF\", oneChar", procedure(ROW_ID, "ExcelRowIdNormalizeValue"))
        self.assertIn("ExcelRowIdGenerateUnique(usedIds)", procedure(ROW_ID, "ExcelRowIdPrepareBatch"))
        self.assertNotIn("通し番号", ROW_ID)

    def test_new_a_to_t_row_gets_id_when_folder_title_is_entered(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        capture = procedure(ROW_ID, "ExcelRowIdCaptureSelection")
        self.assertIn('ws.Range("A:T")', event)
        self.assertIn("If Len(valueText) = 0 Then", event)
        self.assertIn("If targetIncludesBusinessArea Then", event)
        self.assertIn("rowsToAssign.Add CStr(rowNo), True", event)
        self.assertIn("ROW_ID_STATE_HAS_TITLE", capture)
        self.assertIn("ROW_ID_STATE_ID_VALUE", capture)
        self.assertIn("ExcelRowIdCaptureSelection(Target, Me)", SHEET)
        self.assertIn("Set mRowIdSelectionSnapshot = ExcelRowIdCaptureSelection(Application.Selection, Me)", SHEET)

    def test_a_t_edit_reports_duplicate_ids_without_guessing_a_new_row(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertIn("targetIncludesId Or targetIncludesBusinessArea", event)
        self.assertIn("A～T列の変更行に未解決の重複IDがあります", event)
        self.assertIn("自動修復していません", event)

    def test_ctrl_d_duplicate_resolution_uses_prechange_state_not_clipboard_flag(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        resolver = procedure(ROW_ID, "ExcelRowIdResolveDuplicateGroup")
        self.assertIn("If groupRows.Count > 1 Then", event)
        duplicate_branch = event[event.index("If groupRows.Count > 1 Then"):event.index("ElseIf isCopyOperation")]
        self.assertNotIn("isCopyOperation", duplicate_branch)
        self.assertNotIn("Application.CutCopyMode", resolver)
        self.assertIn("priorHadTitle", resolver)
        self.assertIn("priorIdValue", resolver)
        self.assertIn("rowsToAssign.Add candidateRowKey, True", resolver)
        self.assertIn("groupRows.Count <> 2", resolver)

    def test_existing_row_overwrite_restores_the_prechange_id(self):
        resolver = procedure(ROW_ID, "ExcelRowIdResolveDuplicateGroup")
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertIn("candidateValue = CStr(priorIdValue)", resolver)
        self.assertIn("rowsToRestore.Add candidateRowKey, candidateValue", resolver)
        self.assertIn("If idGroups.Exists(candidateNormalizedId) Then", resolver)
        self.assertIn("rowsToRestore.Add CStr(rowNo), CStr(priorIdValue)", event)
        self.assertIn("ws.Cells(rowNo, idCol).Value2 = CStr(writeValues.Item(CStr(rowKey)))", event)
        self.assertNotIn("Cells(rowNo, titleCol).Value2 =", event)

    def test_a_t_and_whole_row_copy_paths_are_distinguished(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertIn('targetIncludesBusinessArea = Not (Application.Intersect(Target, ws.Range("A:T")) Is Nothing)', event)
        self.assertIn("If targetIncludesBusinessArea Then", event)
        self.assertIn("targetIncludesId = Not (Application.Intersect(Target, ws.Columns(idCol)) Is Nothing)", event)
        self.assertIn('targetIncludesSystemArea = Not (Application.Intersect(Target, ws.Range("U:XFD")) Is Nothing)', event)
        self.assertIn("ExcelRowIdResolveDuplicateGroup", event)

    def test_copy_cells_insert_can_detect_one_row_shift_from_prechange_ids(self):
        resolver = procedure(ROW_ID, "ExcelRowIdResolveDuplicateGroup")
        self.assertIn("ExcelRowIdGroupContainsRow(idGroups, normalizedId, rowNo + 1)", resolver)
        self.assertIn("ExcelRowIdGroupContainsRow(idGroups, priorNormalizedId, rowNo + 1)", resolver)
        self.assertIn("newCount = newCount + 1", resolver)
        self.assertIn("groupRows.Count <> 2", resolver)

    def test_ambiguous_or_inserted_copy_stops_without_auto_repair(self):
        resolver = procedure(ROW_ID, "ExcelRowIdResolveDuplicateGroup")
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertIn("コピー先の変更前状態を確認できません", resolver)
        self.assertIn("groupRows.Count <> 2", resolver)
        self.assertIn("自動修復していません", event)
        self.assertIn("Ctrl+Z", event)
        self.assertNotIn("Application.Undo", event)

    def test_cut_moves_and_sorts_do_not_trigger_id_reissue(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertIn("isCutOperation = (Application.CutCopyMode = xlCut)", event)
        self.assertIn("切り取り・行移動中にID重複", event)
        self.assertIn("ElseIf isCopyOperation And targetIncludesTitle Then", event)
        self.assertIn("If idGroups.Exists(priorNormalizedId) Then", event)
        self.assertIn("並べ替え・移動との区別ができません", event)

    def test_last_edit_tracking_order_and_internal_id_write_are_preserved(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertLess(SHEET.index("新ファイル基準表_最終編集情報を記録 Target, Me"), SHEET.index("ExcelRowIdHandleWorksheetChange Target"))
        self.assertIn("If Application.CutCopyMode = xlCut Then GoTo SafeExit", procedure(MODULE1, "新ファイル基準表_最終編集情報を記録"))
        self.assertIn("Application.EnableEvents = False", event)
        self.assertLess(SHEET.index("ExcelRowIdHandleWorksheetChange Target"), SHEET.index("If ShouldReindexSerial(Target, Me)"))
        self.assertLess(SHEET.index("新ファイル基準表_登録予定識別番号を自動クリア Target, Me"), SHEET.index("ExcelRowIdHandleWorksheetChange Target"))
        self.assertIn("正式ファイルID・ファイル番号・バーコード番号", event)
        self.assertIn("ws.Cells(rowNo, idCol).Value2 = CStr(writeValues.Item(CStr(rowKey)))", event)
        self.assertNotIn("最終編集者", event)
        self.assertNotIn("最終編集日時", event)
        self.assertIn("If ExcelRowIdIsHeaderName(CStr(changedArea.Worksheet.Cells(1, firstCol).Value2)) Then Exit Function", procedure(MODULE1, "LastEditAreaHasBusinessChanges"))

    def test_step2_and_annual_rows_assign_private_ids_not_transfer_them(self):
        step2 = procedure(MODULE1, "手順2_新ファイル基準表を作成する")
        self.assertIn("rowIdCol = ExcelRowIdValidateSheet(wsDst)", step2)
        self.assertIn("newRowId = ExcelRowIdCreateNewValue(wsDst)", step2)
        self.assertIn("wsDst.Cells(appendRow, rowIdCol).ClearContents", step2)
        self.assertIn("Call ExcelRowIdAssignGeneratedRow(wsDst, appendRow, newRowId)", step2)
        mapping = procedure(MODULE1, "BuildStep2SourceTransferMap")
        self.assertLess(mapping.index("ExcelRowIdIsHeaderName(headerText)"), mapping.index("EnsureColumnByHeader"))
        self.assertIn("If ExcelRowIdIsHeaderName(canonicalHeader) Then", procedure(MODULE1, "IsUserManagedHeader"))
        annual = procedure(MODULE1, "nPlus1_次年度フォルダ生成")
        self.assertIn("Set generatedRowIds = ExcelRowIdPrepareBatch(wsDst, targetCount)", annual)
        self.assertIn("Excel管理行ID", procedure(MODULE1, "ClearIdentityColumnsForNewRow"))
        self.assertIn("ExcelRowIdAssignGeneratedRow(wsDst, appendRow", annual)

    def test_legacy_migration_generates_ids_before_inserting_rows(self):
        migration = procedure(MODULE1, "旧基準表_継続文書移行")
        self.assertIn("Set rowIds = ExcelRowIdPrepareBatch(wsNew, totalNewRows)", migration)
        self.assertLess(migration.index("Set rowIds = ExcelRowIdPrepareBatch"), migration.index(".Insert Shift:=xlDown"))
        self.assertIn(".Cells(insertRow + j - 1, newColRowId).ClearContents", migration)
        self.assertIn("Call ExcelRowIdAssignGeneratedRow(wsNew", migration)

    def test_internal_id_is_not_a_sync_key_export_column_or_review_signature(self):
        export = procedure(MODULE1, "手順5_差分インポートCSVを外部出力する")
        self.assertIn("Not ExcelRowIdIsHeaderName(headerName)", export)
        self.assertIn("If ExcelRowIdIsHeaderName(headerName) Then", procedure(MODULE1, "IsStep2UpdateExcludedHeader"))
        self.assertIn("ExcelRowIdIsHeaderName(Trim$(CStr(headerValue)))", procedure(REVIEW, "ReviewIsInternalMetadataHeader"))
        tracking = procedure(MODULE1, "LastEditAreaHasBusinessChanges")
        self.assertIn("ExcelRowIdIsHeaderName(CStr(changedArea.Worksheet.Cells(1, firstCol).Value2))", tracking)

    def test_user_docs_describe_installation_and_privacy(self):
        manual = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8")
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        self.assertIn("Excel管理行ID_点検", manual)
        self.assertIn("既に列がある場合は追加せず", manual)
        self.assertIn("Excel管理行ID_既存行へ初回付与", manual)
        self.assertIn("Excel内部専用", readme)
        self.assertIn("vba/modExcelRowId.txt", readme)
        self.assertIn("Ctrl+D", manual)
        self.assertIn("A～T列だけのコピー", manual)
        self.assertIn("正式ファイルID・ファイル番号・バーコード番号", manual)
        self.assertIn("管理行IDの自動発行だけでは", manual)


if __name__ == "__main__":
    unittest.main()
