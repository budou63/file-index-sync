"""E multi-selection is display-only; file grouping never changes item decisions."""
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8-sig")
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")


def procedure(text, name, kind="Sub"):
    marker = ("Private " if text is FORM else "Public ") + kind + " " + name + "("
    if marker not in text and text is FORM:
        marker = kind + " " + name + "("
    start = text.index(marker)
    end = text.index("End " + kind, start)
    return text[start:end]


class ESelectionOnlyContracts(unittest.TestCase):
    def test_no_batch_execution_api_or_form_call(self):
        self.assertNotIn("BidirectionalSyncResolveConflictBatch", MODULE)
        self.assertNotIn("BidirectionalSyncResolveConflictBatch", FORM)
        self.assertNotIn("ApplyEChoice", FORM)
        self.assertNotIn("SelectedECases", FORM)

    def test_filter_populated_from_current_year_real_groups(self):
        group = procedure(FORM, "PopulateEGroups")
        for token in ('mCboEGroup.AddItem "すべて"', 'item("groupName")',
                      'CLng(item("year")) = targetYear', 'mCboEGroup.ListIndex = 0'):
            self.assertIn(token, group)
        show = procedure(FORM, "ShowECategory")
        self.assertIn('ERecordHasGroup(record, selectedGroup)', show)
        self.assertIn('mCboEGroup.ListIndex > 0', show)  # a real group named すべて stays selectable
        self.assertIn('CStr(allRecords.Count)', show)
        self.assertIn('CStr(totalItems)', show)
        self.assertIn('mCboEGroup.Visible = True', show)

    def test_multiple_selection_only_on_e_and_counted(self):
        show = procedure(FORM, "ShowCategory")
        e_show = procedure(FORM, "ShowECategory")
        layout = procedure(FORM, "LayoutReviewForE")
        self.assertIn('mLstCases.MultiSelect = fmMultiSelectExtended', e_show)
        self.assertIn('mLstCases.MultiSelect = fmMultiSelectSingle', show)
        self.assertIn('mLstRows.MultiSelect = fmMultiSelectSingle', layout)
        self.assertIn('If categoryIndex <> 4 And mLstCases.ListCount > 0 Then mLstCases.ListIndex = 0', show)
        count = procedure(FORM, "ESelectedCount", "Function")
        self.assertIn('If mCboCategory.ListIndex <> 4 Then Exit Function', count)
        self.assertIn('mLstCases.Selected(i)', count)
        selection = procedure(FORM, "UpdateESelection")
        self.assertIn('CStr(n) & "件選択中"', selection)
        self.assertIn('identitySafe = Not CBool(record("identityIssue"))', selection)
        for name in ('mCmdApprove', 'mCmdOther', 'mCmdHold'):
            self.assertIn(name + '.Enabled = (n = 1 And Not item Is Nothing And identitySafe)', selection)
        self.assertIn('mCmdEClear.Enabled = (n > 0)', selection)
        self.assertIn('複数選択は表示専用です', selection)

    def test_all_e_write_handlers_fail_closed_on_multiple_selection(self):
        approve = procedure(FORM, "mCmdApprove_Click")
        other = procedure(FORM, "RecordChoice")
        for body in (approve, other):
            self.assertIn('If ESelectedCount() <> 1 Then Exit Sub', body)
            self.assertIn('BidirectionalSyncResolveConflict(item,', body)
            self.assertLess(body.index('If ESelectedCount() <> 1 Then Exit Sub'),
                            body.index('BidirectionalSyncResolveConflict(item,'))
        self.assertIn('MsgBox("文書管理システムの値をExcel基準表へ反映します。選択した関連項目グループ全体が対象です。続けますか？"', approve)
        self.assertIn('If MsgBox(', approve)

    def test_e_mouse_keyboard_and_focus_changes_refresh_selected_flags(self):
        for event in ('mLstCases_MouseUp', 'mLstCases_KeyUp', 'mLstCases_Change'):
            body = procedure(FORM, event)
            self.assertIn('If Not rendering And mCboCategory.ListIndex = 4 Then ShowCurrent', body)
            self.assertNotIn('BidirectionalSyncResolveConflict(', body)
        exit_body = procedure(FORM, 'mLstCases_Exit')
        self.assertIn('If Not rendering And mCboCategory.ListIndex = 4 Then UpdateESelection', exit_body)
        self.assertNotIn('BidirectionalSyncResolveConflict(', exit_body)
        self.assertIn('ByVal KeyCode As MSForms.ReturnInteger',
                      procedure(FORM, 'mLstCases_KeyUp'))
        self.assertIn('ByVal Cancel As MSForms.ReturnBoolean',
                      procedure(FORM, 'mLstCases_Exit'))
        self.assertIn('ByVal Button As Integer', procedure(FORM, 'mLstCases_MouseUp'))
        self.assertIn('UpdateESelection', procedure(FORM, 'ShowCurrent'))
        self.assertIn('mLstCases.Selected(i)', procedure(FORM, 'ESelectedCount', 'Function'))
        self.assertIn('If Not rendering Then ShowCurrent', procedure(FORM, 'mLstCases_Click'))

    def test_selection_clear_filter_and_reload_drop_prior_selection(self):
        clear = procedure(FORM, "mCmdEClear_Click")
        self.assertIn('mLstCases.Selected(i) = False', clear)
        self.assertIn('mLstCases.ListIndex = -1', clear)
        self.assertIn('ShowCategory', procedure(FORM, "mCboEGroup_Change"))
        self.assertIn('mLstCases.Clear', procedure(FORM, "ShowCategory"))
        self.assertIn('PopulateEGroups', procedure(FORM, "mCmdRefresh_Click"))
        self.assertIn('PopulateEGroups', procedure(FORM, "LoadCases"))

    def test_single_e_and_a_to_d_existing_entrypoints_retained(self):
        self.assertIn('BidirectionalSyncResolveConflict(item, "システム採用")',
                      procedure(FORM, "mCmdApprove_Click"))
        self.assertIn('BidirectionalSyncResolveConflict(item, decision)',
                      procedure(FORM, "RecordChoice"))
        self.assertIn('SyncReviewApprove(item, SelectedRow())', procedure(FORM, "mCmdApprove_Click"))
        self.assertIn('SyncReviewDecide(item, decision)', procedure(FORM, "RecordChoice"))
        self.assertIn('SelectSingleECase', procedure(FORM, "mCmdNext_Click"))
        self.assertIn('Public Function BidirectionalSyncResolveConflict(', MODULE)

    def test_form_artifacts_and_source_are_exact(self):
        FORM.encode('cp932')
        for name, encoding in [('frmSyncReview.frm', 'utf-8'), ('frmSyncReview-native.frm', 'cp932')]:
            content = (ROOT / name).read_bytes().decode(encoding)
            self.assertEqual(content[content.index('Option Explicit'):].replace(chr(13)+chr(10),chr(10)), FORM)


if __name__ == '__main__':
    unittest.main()
