import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = ROOT / "frmSyncReview_FormCode.txt"
MODULE = ROOT / "modSyncReview.bas"


def proc(text, name):
    start = text.index(f"Private Sub {name}(")
    return text[start:text.index("End Sub", start) + len("End Sub")]


class SyncReviewDFormTests(unittest.TestCase):
    def test_opener_combines_categories_and_fails_closed_if_d_detector_errors(self):
        src = MODULE.read_text(encoding="utf-8")
        opener = proc(src, "OpenSyncReviewForYear")
        self.assertIn("SyncReviewCandidates(yearValue)", opener)
        self.assertIn("SyncReviewCopyIdCandidates(yearValue)", opener)
        self.assertIn("If cases.Count = 0 Then Exit Sub", opener)
        self.assertLess(opener.index("cases.Add copyItem"), opener.index("If cases.Count = 0 Then Exit Sub"))

    def test_four_category_counts_and_category_does_not_call_abc_decisions(self):
        src = FORM.read_text(encoding="utf-8")
        self.assertIn('mCboCategory.AddItem "D：コピー識別番号のクリア候補"', src)
        self.assertIn("counts(1 To 4)", src)
        self.assertIn('Case "D": idx = 4', src)
        self.assertIn('category = "D"', src)

    def test_d_rows_are_role_specific_and_clear_button_calls_only_clear_api(self):
        src = FORM.read_text(encoding="utf-8")
        self.assertIn('item("copyCandidateRows")', src)
        self.assertIn('item("copyKeeperRow")', src)
        self.assertIn("mCmdClear", src)
        click = proc(src, "mCmdClear_Click")
        self.assertIn("SyncReviewClearCopyIdentifiers(item, SelectedRow())", click)
        self.assertNotRegex(click, r"Cells\s*\([^\n]*\)\s*\.\s*(?:ClearContents|Value)")

    def test_d_disables_legacy_decision_delete_and_navigation(self):
        src = FORM.read_text(encoding="utf-8")
        for name in ("mCmdApprove_Click", "mCmdDelete_Click", "mCmdGoTo_Click", "RecordChoice"):
            self.assertIn('CStr(item("category")) = "D"', proc(src, name))
        self.assertIn('RecordChoice "別物"', proc(src, "mCmdOther_Click"))
        self.assertIn('RecordChoice "保留"', proc(src, "mCmdHold_Click"))

    def test_cancel_and_failure_never_claim_success_and_success_refreshes(self):
        src = FORM.read_text(encoding="utf-8")
        click = proc(src, "mCmdClear_Click")
        self.assertIn("If completed Then", click)
        self.assertIn("If Not completed Then mCmdClear.Enabled = IsCopyCandidate(item, SelectedRow())", click)
        self.assertIn("mCmdRefresh_Click", click)
        self.assertIn("Failed:", click)
        self.assertIn("MsgBox", click)
        self.assertIn("If apiInvoked And IsCopyCandidate(item, SelectedRow()) Then mCmdClear.Enabled = True", click)

    def test_control_text_canonical_native_and_blank_form_are_synchronized(self):
        paste = FORM.read_text(encoding="utf-8")
        canonical = (ROOT / "frmSyncReview.frm").read_text(encoding="utf-8")
        native = (ROOT / "frmSyncReview-native.frm").read_bytes().decode("cp932")
        self.assertEqual(canonical[canonical.index("Option Explicit"):], paste)
        self.assertEqual(native[native.index("Option Explicit"):].replace("\r\n", "\n"), paste)
        self.assertIn('OleObjectBlob   =   "frmSyncReview-native.frx":0000', native)
        self.assertIn('EnsureControl("cmdClear", "Forms.CommandButton.1")', paste)

    def test_next_from_keeper_selects_first_candidate_and_from_candidate_advances(self):
        src = FORM.read_text(encoding="utf-8")
        advance = proc(src, "mCmdNext_Click")
        guard = advance.index("If Not IsCopyCandidate(item, selected) Then")
        select_first = advance.index("SelectCaseRow CLng(candidates(1))")
        candidate_loop = advance.index("For j = 1 To candidates.Count")
        self.assertLess(guard, candidate_loop)
        self.assertLess(select_first, candidate_loop)
        self.assertIn("SelectCaseRow CLng(candidates(j + 1))", advance)

        def next_row(selected, keeper, candidates):
            return candidates[0] if selected == keeper else candidates[candidates.index(selected) + 1]

        self.assertEqual(next_row(101, 101, [102, 103]), 102)
        self.assertEqual(next_row(102, 101, [102, 103]), 103)

    def test_refresh_keeps_d_selected_when_only_d_candidates_exist(self):
        src = FORM.read_text(encoding="utf-8")
        refresh = proc(src, "mCmdRefresh_Click")
        self.assertIn("SelectDWhenOnlyCategory", refresh)
        self.assertIn("Private Sub SelectDWhenOnlyCategory()", src)
        self.assertIn("mCboCategory.ListIndex = 3", src)

    def test_critical_rollback_failure_is_propagated_to_form_status(self):
        api_text = (ROOT / "modSyncReview.bas").read_text(encoding="utf-8")
        api_start = api_text.index("Public Function SyncReviewClearCopyIdentifiers(")
        api = api_text[api_start:api_text.index("End Function", api_start) + len("End Function")]
        click = proc(FORM.read_text(encoding="utf-8"), "mCmdClear_Click")
        self.assertIn("criticalFailure", api)
        self.assertIn("If criticalFailure Then", api)
        self.assertIn("Err.Raise vbObjectError + 6155", api)
        self.assertLess(api.index("If criticalFailure Then"), api.index("Err.Raise vbObjectError + 6155"))
        self.assertIn("mLblStatus.Caption = Err.Description", click)
        self.assertIn("apiInvoked", click)
        self.assertIn("Err.Number = vbObjectError + 6155", click)
        self.assertIn("ブックを保存せず終了", click)
        self.assertIn("バックアップを保全", click)

    def test_docs_explain_d_confirmation_backup_and_system_separation(self):
        for name in ("新ファイル基準表.md", "同期候補確認フォーム_導入手順.md"):
            text = (ROOT / name).read_text(encoding="utf-8")
            self.assertIn("管理システム", text)
            self.assertIn("バックアップ", text)
        self.assertIn("コピー識別番号のクリア候補", (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8"))
        self.assertIn("## D：コピー識別番号のクリア", (ROOT / "同期候補確認フォーム_導入手順.md").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
