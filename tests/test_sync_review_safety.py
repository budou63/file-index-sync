import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "modSyncReview.bas").read_text(encoding="utf-8")
FORM = (ROOT / "frmSyncReview.frm").read_text(encoding="utf-8")
STEP2 = (ROOT / "Module1").read_text(encoding="utf-8")


def procedure(code, start, end="End Function"):
    match = re.search(r"(?ms)^" + re.escape(start) + r".*?^" + re.escape(end), code)
    if match is None:
        raise AssertionError(f"missing procedure {start}")
    return match.group()


class SyncReviewSafetyTests(unittest.TestCase):
    def test_independent_modeless_entry_and_no_open_write(self):
        entry = procedure(MODULE, "Public Sub 確認候補を開く()", "End Sub")
        opener = procedure(MODULE, "Private Sub OpenSyncReviewForYear(", "End Sub")
        self.assertIn("OpenSyncReviewForYear yearValue", entry)
        self.assertIn("frmSyncReview.Show vbModeless", opener)
        self.assertIn("Set cases = SyncReviewCandidates(yearValue)", opener)
        self.assertIn("Set copyCases = SyncReviewCopyIdCandidates(yearValue)", opener)
        self.assertIn("cases.Add copyItem", opener)
        self.assertIn("If cases.Count = 0 Then", opener)
        self.assertLess(opener.index("cases.Add copyItem"), opener.index("If cases.Count = 0 Then"))
        self.assertNotIn("ReviewWriteAudit", entry)
        self.assertNotIn("SaveCopyAs", entry)
        self.assertIn("ShowModal = 0", FORM)

    def test_step2_auto_open_inherits_year_and_batch_waits_for_restoration(self):
        step2 = procedure(STEP2, "Public Sub 手順2_新ファイル基準表を作成する()", "End Sub")
        batch = procedure(STEP2, "Public Sub 新ファイル基準表生成()", "End Sub")
        auto = procedure(MODULE, "Public Sub SyncReviewAfterStep2Success(", "End Sub")
        begin = procedure(MODULE, "Public Sub SyncReviewBeginBatch()", "End Sub")
        finish = procedure(MODULE, "Public Sub SyncReviewEndBatch(", "End Sub")

        self.assertEqual(step2.count("SyncReviewAfterStep2Success targetYearValue"), 1)
        self.assertIn("step2Succeeded = True", step2)
        self.assertLess(step2.index("step2Succeeded = True"), step2.index("SafeExit:"))
        self.assertLess(step2.index("RestoreCalculationIfPossible previousCalculationMode"), step2.index("SyncReviewAfterStep2Success targetYearValue"))
        self.assertIn("autoReviewPendingYear = yearValue", auto)
        self.assertNotIn("InputBox", auto)
        self.assertIn("autoReviewPending = False", begin)
        self.assertIn("shouldOpen = (allowOpen And autoReviewPending)", finish)
        self.assertIn("If shouldOpen Then OpenSyncReviewForYear yearValue", finish)

        self.assertLess(batch.index("SyncReviewBeginBatch"), batch.index("手順2_新ファイル基準表を作成する"))
        self.assertLess(batch.index("手順3_コード管理CSVを作成する(True)"), batch.index("Cleanup:"))
        self.assertLess(batch.index("手順3_コード管理CSVを作成する(True)"), batch.index("batchSucceeded = True"))
        self.assertIn("SyncReviewEndBatch batchSucceeded", batch)
        self.assertLess(batch.index("RestoreCalculationIfPossible oldCalculation"), batch.index("SyncReviewEndBatch"))

    def test_csv_import_alone_does_not_auto_open_review(self):
        csv_import = procedure(STEP2, "Private Sub CSVImportExecute(", "End Sub")
        self.assertNotIn("SyncReviewAfterStep2Success", csv_import)
        self.assertNotIn("OpenSyncReviewForYear", csv_import)

    def test_candidates_are_derived_from_live_sheets_not_comments(self):
        build = procedure(MODULE, "Public Function SyncReviewCandidates(")
        for name in ("csvIds", "csvBars", "csvNumbers", "csvNames", "csvFolders", "ReviewField", "ReviewFiscal", "ReviewName"):
            self.assertIn(name, build)
        self.assertNotIn(".Comment", build)
        self.assertIn("If targetCount = 0 Then Err.Raise", build)
        self.assertIn("ReviewRows(csvIds, fid).Count > 0", build)
        self.assertIn("Len(name) = 0", build)
        self.assertIn('cases.Add NewCase("A"', build)

    def test_approval_is_narrow_and_revalidated(self):
        approve = procedure(MODULE, "Public Function SyncReviewApprove(")
        self.assertGreaterEqual(approve.count("ReviewCurrent(item)"), 2)
        self.assertIn('If CStr(item("category")) <> "B"', approve)
        self.assertIn('If CStr(item("rows")) <> CStr(selectedRow)', approve)
        self.assertIn("ReviewRows", MODULE)
        for field in ("ファイルID", "ファイル番号", "バーコード番号", "同期状態"):
            self.assertIn(field, approve)
        self.assertIn("ReviewBackup()", approve)
        self.assertIn("Application.EnableEvents = False", approve)
        self.assertIn("vbDefaultButton2", approve)
        self.assertNotIn("Rows(selectedRow).Delete", approve)

    def test_deletion_is_backup_first_and_descending(self):
        delete = procedure(MODULE, "Public Function SyncReviewDeleteBatch(")
        self.assertIn('If CStr(item("category")) <> "C"', delete)
        self.assertGreaterEqual(delete.count("ReviewCurrent(item, current)"), 2)
        self.assertIn("ReviewSignature(dst, rowNo)", delete)
        self.assertIn("SyncReviewDecision(item)", delete)
        self.assertIn("backup = ReviewBackup()", delete)
        self.assertLess(delete.index("backup = ReviewBackup()"), delete.index("dst.Rows(rowNo).Delete"))
        self.assertIn("rowNo > bestRow", delete)
        self.assertIn('ReviewWriteAudit item, "削除着手"', delete)
        self.assertIn('ReviewWriteAudit item, "削除完了"', delete)
        self.assertIn("dst.FilterMode", delete)
        self.assertIn("vbDefaultButton2", delete)

    def test_decisions_bound_to_full_row_signatures(self):
        self.assertIn('item.Add "sourceSignature"', MODULE)
        self.assertIn('item.Add "destinationSignature"', MODULE)
        signature = procedure(MODULE, "Private Function ReviewSignature(")
        self.assertIn("HasFormula", signature)
        self.assertIn("NumberFormat", signature)
        self.assertIn("TypeName", signature)
        current = procedure(MODULE, "Private Function ReviewCurrent(")
        self.assertIn('SyncReviewCandidates(CLng(item("year")))', current)
        self.assertIn("sourceSignature", current)
        audit = procedure(MODULE, "Private Sub ReviewWriteAudit(", "End Sub")
        self.assertIn("ws.Cells(r, 9).NumberFormat = \"@\"", audit)
        self.assertIn("ws.Cells(r, 10).NumberFormat = \"@\"", audit)
        self.assertIn("destinationSignature", current)

    def test_form_buttons_and_multi_selection(self):
        for control in ("cboCategory", "lstCases", "lstRows", "txtDetails", "chkDelete", "lblCounts", "cmdApprove", "cmdOther", "cmdHold", "cmdNext", "cmdGoTo", "cmdDelete", "cmdRefresh", "cmdFinish"):
            self.assertIn(f'EnsureControl("{control}", "Forms.', FORM)
        self.assertIn("SyncReviewDeleteBatch(selected)", FORM)
        self.assertIn("SyncReviewDecide(item, decision)", FORM)
        self.assertIn("SyncReviewApprove(item, SelectedRow())", FORM)

    def test_step2_pre_add_suspect_hold_not_in_normal_update(self):
        guard = "Set destCandidates = Step2SuspectPlannedRows(wsSrc, srcRow, wsDst, targetYearValue, suspectGuideMap)"
        append = "AppendRowFromSourceByHeader wsSrc, wsDst, srcRow, appendRow"
        self.assertIn(guard, STEP2)
        self.assertIn(append, STEP2)
        self.assertLess(STEP2.index(guard), STEP2.index(append))
        self.assertLess(STEP2.index("Step2PreflightSimilarNameColumns wsSrc"), STEP2.index("EnsureColumnByHeader wsDst"))
        self.assertLess(STEP2.index("BuildStep2SuspectGuideMap wsSrc"), STEP2.index(guard))
        scope = STEP2[STEP2.index(guard):STEP2.index(append)]
        self.assertIn("RecordStep2IdentityHold", scope)
        self.assertIn("GoTo ContinueNextSourceRow", scope)
        helper = procedure(STEP2, "Private Function Step2SuspectPlannedRows(")
        self.assertIn("GetNormalizedYearFromRow", helper)
        self.assertIn('"登録予定"', helper)
        self.assertIn("Len(GetTrimmedCell(wsDst, rowNo, idDst)) <> 0", helper)


    def test_decided_delete_candidates_cannot_be_selected_or_executed(self):
        check = procedure(FORM, "Private Sub mChkDelete_Click()", "End Sub")
        record = procedure(FORM, "Private Sub RecordChoice(", "End Sub")
        batch = procedure(MODULE, "Public Function SyncReviewDeleteBatch(")
        self.assertIn("SyncReviewDecision(item)", check)
        self.assertIn("deletionSelection.Remove", record)
        self.assertIn("mChkDelete.Value = False", record)
        self.assertGreaterEqual(batch.count("SyncReviewDecision(item)"), 3)
        self.assertLess(batch.index("backup = ReviewBackup()"), batch.index("dst.Rows(rowNo).Delete"))
        self.assertIn("approvedKeys.Exists", batch)

    def test_failed_deletion_is_not_terminal_decision(self):
        decision = procedure(MODULE, "Public Function SyncReviewDecision(")
        batch = procedure(MODULE, "Public Function SyncReviewDeleteBatch(")
        self.assertIn("ReviewAuditHashMatches", decision)
        legacy = procedure(MODULE, "Private Function ReviewAuditHashMatches(")
        self.assertIn("IsEmpty(stored)", legacy)
        self.assertIn("CDbl(stored) = 0#", legacy)
        self.assertIn("ReviewWriteAudit item,", batch)
        self.assertIn("削除失敗", batch)
        self.assertLess(batch.index("completed = completed + 1"), batch.index("ReviewWriteAudit item, \"削除完了\""))
        self.assertIn("If rowDeleted Then", batch)
        self.assertIn("削除完了", decision)
        self.assertNotIn("削除失敗", decision)
        self.assertIn("NumberFormat", procedure(MODULE, "Private Sub ReviewWriteAudit(", "End Sub"))

    def test_vba_lines_have_balanced_string_literals(self):
        for code in (MODULE, FORM, STEP2):
            for index, line in enumerate(code.splitlines(), 1):
                self.assertEqual(line.count(chr(34)) % 2, 0, f"unclosed VBA string at line {index}")

    def test_prioritized_aliases_and_indexed_similarity(self):
        col = procedure(MODULE, "Private Function ReviewCol(")
        helper = procedure(STEP2, "Private Function Step2SuspectPlannedRows(")
        self.assertIn("If chosen = 0 Then", col)
        self.assertIn("別名見出しの内容が一致しません", col)
        self.assertIn("If found > 0 Then Err.Raise", col)
        self.assertIn("Step2MapRows(suspectGuideMap, groupKey)", helper)
        self.assertNotIn("For rowNo = 2 To dstLastRow", helper)
        validator = procedure(STEP2, "Private Function Step2ReviewHeaderFamily(")
        self.assertIn("年度(和暦)", validator)
        self.assertIn("作成年度(和暦)", validator)
        preflight = procedure(STEP2, "Private Sub Step2PreflightHeaderAmbiguity(", "End Sub")
        self.assertIn("If key = \"年度（和暦）\" Or key = \"作成年度（和暦）\" Then", preflight)
        self.assertLess(preflight.index("If key ="), preflight.index("For rowNo = 2 To lastRow"))


if __name__ == "__main__":
    unittest.main()
