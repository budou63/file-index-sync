import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = ROOT / "frmSyncReview_FormCode.txt"
MODULE = ROOT / "vba/modSyncReview.txt"


def proc(text, name):
    start = text.index(f"Private Sub {name}(")
    return text[start:text.index("End Sub", start) + len("End Sub")]


class SyncReviewDFormTests(unittest.TestCase):
    def test_opener_combines_categories_and_fails_closed_if_d_detector_errors(self):
        src = MODULE.read_text(encoding="utf-8")
        opener = proc(src, "OpenSyncReviewForYear")
        self.assertIn("SyncReviewCandidates(yearValue)", opener)
        self.assertIn("SyncReviewCopyIdCandidates(yearValue)", opener)
        self.assertIn("If cases.Count = 0 Then", opener)
        self.assertLess(opener.index("cases.Add copyItem"), opener.index("If cases.Count = 0 Then"))

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

    def test_missing_d_header_disables_d_and_warning_survives_selection(self):
        src = FORM.read_text(encoding="utf-8")
        load = src[src.index("Public Sub LoadCases("):src.index("End Sub", src.index("Public Sub LoadCases("))]
        self.assertIn("Optional ByVal unavailableReason As String", load)
        self.assertIn("mDUnavailableReason = unavailableReason", load)
        self.assertIn('EnsureControl("lblDWarning", "Forms.Label.1")', src)
        self.assertIn("mLblDWarning.Caption =", src)
        self.assertIn("mLblDWarning.Visible =", src)
        self.assertIn("If Len(mDUnavailableReason) > 0 Then", proc(src, "mCboCategory_Change"))
        self.assertIn("mCmdClear.Enabled = False", proc(src, "ShowCurrent"))

    def test_reopening_loaded_form_falls_back_from_unavailable_d(self):
        src = FORM.read_text(encoding="utf-8")
        load = src[src.index("Public Sub LoadCases("):src.index("End Sub", src.index("Public Sub LoadCases("))]
        self.assertIn('If Len(mDUnavailableReason) > 0 And mCboCategory.ListIndex = 3 Then', load)
        self.assertIn('rendering = True: mCboCategory.ListIndex = 0: rendering = False', load)
        self.assertLess(load.index('mCboCategory.ListIndex = 0'), load.index('ShowCategory'))

    def test_refresh_stages_candidates_and_only_recovers_missing_d_header(self):
        src = FORM.read_text(encoding="utf-8")
        refresh = proc(src, "mCmdRefresh_Click")
        self.assertIn("Set refreshed = SyncReviewCandidates(targetYear)", refresh)
        self.assertIn("Set copyCases = SyncReviewCopyIdCandidates(targetYear)", refresh)
        self.assertIn("vbObjectError + 6141", refresh)
        self.assertIn("Err.Raise detectorNumber", refresh)
        self.assertIn("refreshed.Add copyItem", refresh)
        self.assertLess(refresh.index("refreshed.Add copyItem"), refresh.index("Set cases = refreshed"))
        self.assertLess(refresh.index("If detectorNumber <> 0 Then"), refresh.index("Set cases = refreshed"))
        self.assertNotIn("cases.Add copyItem", refresh)
        self.assertLess(refresh.index("Set cases = refreshed"), refresh.index("ShowCategory"))
        self.assertIn("mDUnavailableReason = unavailableReason", refresh)
        self.assertIn("UpdateDWarning", refresh)
        self.assertIn('mCboCategory.List(3) = "D：コピー識別番号のクリア候補"', src)
        self.assertIn('mLblDWarning.Caption = "警告：D利用不可: " & mDUnavailableReason', src)
        self.assertIn('MsgBox "D候補を利用できません: " & mDUnavailableReason', src)
        self.assertIn("Set newSelection = CreateObject", refresh)
        self.assertLess(refresh.index("Set newSelection = CreateObject"), refresh.index("Set cases = refreshed"))
        self.assertLess(refresh.index("Err.Raise detectorNumber"), refresh.index("Set cases = refreshed"))
        self.assertIn("Set previousCases = cases", refresh)
        self.assertIn("Set previousSelection = deletionSelection", refresh)
        self.assertIn("previousReason = mDUnavailableReason", refresh)
        self.assertIn("If committed Then", refresh)
        self.assertIn("Set cases = previousCases", refresh)
        self.assertIn("Set deletionSelection = previousSelection", refresh)
        self.assertIn("mDUnavailableReason = previousReason", refresh)

    def test_bottom_controls_are_inside_client_area_with_margin(self):
        src = FORM.read_text(encoding="utf-8")
        build = proc(src, "BuildControls")
        self.assertIn("Me.Height = Me.Height + (440 - Me.InsideHeight)", build)
        self.assertNotIn("Me.InsideHeight =", build)
        inner = 440.0
        bottom = max(float(y) + float(h) for y, h in
                     re.findall(r"PlaceControl mCmd(?:Refresh|Finish), [\d.]+, ([\d.]+), [\d.]+, ([\d.]+)", build))
        self.assertGreater(inner - bottom, 8)
        self.assertGreater(bottom, 403)

    def test_critical_rollback_failure_is_propagated_to_form_status(self):
        api_text = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")
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
