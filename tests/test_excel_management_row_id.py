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

    def test_new_rows_receive_ids_but_legacy_blank_rows_need_explicit_first_run(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertIn("priorTitleState.Exists(CStr(rowNo))", event)
        self.assertIn("If Not CBool(priorTitleState(CStr(rowNo))) Then", event)
        self.assertIn("Application.CutCopyMode = xlCopy", event)
        self.assertIn("ElseIf isCopyOperation And targetIncludesTitle Then", event)
        self.assertIn("ElseIf CBool(priorTitleState(CStr(rowNo))) Then", event)
        self.assertIn("コピー先の変更前状態を確認できません", event)
        self.assertIn("groupRows.Count = 2", event)
        self.assertIn("Target.Areas.Count = 1", event)
        self.assertIn("IDは変更していません", event)
        self.assertIn("ExcelRowIdCaptureSelection(Target, Me)", SHEET)
        self.assertIn("ExcelRowIdHandleWorksheetChange Target, Me, mRowIdSelectionSnapshot", SHEET)
        self.assertLess(SHEET.index("ExcelRowIdHandleWorksheetChange Target"), SHEET.index("If ShouldReindexSerial(Target, Me)"))

    def test_existing_ids_are_not_reissued_for_moves_or_sorts(self):
        event = procedure(ROW_ID, "ExcelRowIdHandleWorksheetChange")
        self.assertIn("isCopyOperation = (Application.CutCopyMode = xlCopy)", event)
        self.assertIn("ElseIf targetIncludesId And idGroups.Exists(normalizedId) Then", event)
        self.assertNotIn("Application.CutCopyMode = xlCut", event)
        assignment = procedure(ROW_ID, "ExcelRowIdAssignGeneratedRow")
        preserve = "If ExcelRowIdNormalizeValue(targetCell.Value2, normalizedId) Then"
        self.assertIn(preserve, assignment)
        self.assertLess(assignment.index(preserve), assignment.index("Set usedIds = ExcelRowIdBuildUsedIds"))

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


if __name__ == "__main__":
    unittest.main()
