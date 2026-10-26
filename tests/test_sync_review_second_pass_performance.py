"""Second-pass E display performance and safety contracts (no Excel required)."""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT/'vba/frmSyncReview.txt').read_text(encoding='utf-8')
sys.path.insert(0, str(Path(__file__).parent))
from benchmark_sync_review_second_pass import make_records, render


def proc(name, kind='Sub'):
    import re
    match = re.search(rf'Private {kind} {name}\(', FORM)
    if not match:
        raise AssertionError(name)
    return FORM[match.start():FORM.index('End ' + kind, match.start())]


class ESecondPassPerformanceTests(unittest.TestCase):
    def test_right_mouse_keyboard_each_revalidate_once_per_gesture(self):
        update = proc('UpdateEFieldSelection')
        guard = 'If fieldSelectionKey = eRenderedFieldSelection Then Exit Sub'
        self.assertIn(guard, update)
        self.assertIn('fieldSelectionKey = CStr(mLstRows.ListIndex)', update)
        self.assertNotIn('ESelectionRenderKey()', update)
        self.assertLess(update.index(guard), update.index('Set record = SelectedRecord()'))
        self.assertLess(update.index('UpdateESelection'), update.index('eRenderedFieldSelection = fieldSelectionKey'))
        for handler in ('mLstRows_MouseDown', 'mLstRows_KeyDown'):
            self.assertIn('eRenderedFieldSelection = ""', proc(handler))
        for handler in ('mLstRows_Click', 'mLstRows_Change', 'mLstRows_KeyUp'):
            self.assertIn('UpdateEFieldSelection', proc(handler))
        self.assertIn('eRenderedFieldSelection = ""', proc('mLstCases_Exit'))

    def test_invalidations_and_safety_checks_are_preserved(self):
        for name in ('LoadCases', 'ShowECategory', 'ShowECurrent', 'EDisplayFailed'):
            body = (FORM[FORM.index('Public Sub LoadCases('):FORM.index('End Sub', FORM.index('Public Sub LoadCases('))]
                    if name == 'LoadCases' else proc(name))
            self.assertIn('eRenderedFieldSelection = ""', body)
        select = proc('UpdateESelection')
        self.assertIn('ESelectedDecisionIsSafe("システム採用"', select)
        self.assertIn('ESelectedDecisionIsSafe("Excel採用"', select)
        self.assertIn('BidirectionalSyncClassification1DecisionSafe(item, "保留"', select)
        self.assertIn('BidirectionalSyncRetentionDecisionSafe(item, "保留"', select)
        self.assertIn('If n = 1 Then', select)
        for button in ('mCmdApprove', 'mCmdOther', 'mCmdHold'):
            self.assertIn(button+'.Enabled = (n = 1', select)
        self.assertNotIn('BidirectionalSyncResolveConflict(', proc('ShowECategory'))
        self.assertNotIn('BidirectionalSyncResolveConflict(', proc('UpdateEFieldSelection'))

    def test_visibility_switch_is_display_only_and_layout_still_updates_on_resize(self):
        self.assertIn('ShowECategory True, False', proc('mCboEVisibility_Change'))
        self.assertIn('If applyLayout Then LayoutReviewForE True', proc('ShowECategory'))
        for name in ('mCboEWindowSize_Change', 'mCboEPaneSize_Change'):
            self.assertIn('ShowECategory True', proc(name))
        show = proc('ShowECategory')
        self.assertIn('BeginEDisplayRead', show)
        self.assertIn('EndEDisplayRead', show)
        self.assertIn('needsConfirmation = ERecordNeedsConfirmation(record)', show)
        self.assertEqual(show.count('ERecordNeedsConfirmation(record)'), 1)
        self.assertEqual(show.count('ERecordNeedsInternalInvestigation(record)'), 1)
        self.assertIn('ECountInternalCase item, internalMismatchCount, unresolvedInternalCount', show)

    def test_bulk_render_keeps_order_and_selection_recovery(self):
        show = proc('ShowECategory')
        for token in ('pendingRows.Add Array(guide2Value, folderValue, creationYearValue)',
                      'ReDim listData(0 To shown.Count - 1, 0 To 2)',
                      'mLstCases.List = listData',
                      'If selectedRecordKeys.Exists(recordKey) Then mLstCases.Selected(rowIndex) = True',
                      'mLstCases.TopIndex = priorTopIndex',
                      'StrComp(existingRecordKey, priorRecordKey, vbBinaryCompare) = 0',
                      'ShowCurrent', 'UpdateEFieldSelection'):
            self.assertIn(token, show)
        self.assertLess(show.index('mLstCases.List = listData'), show.index('mLstCases.TopIndex = priorTopIndex'))
        self.assertIn('mLstCases.MultiSelect = fmMultiSelectExtended', show)

    def test_model_same_rows_counts_first_middle_last_and_visibility_round_trip(self):
        records = make_records()
        for full in (False, True, False):
            old = render(records, full, False)
            new = render(records, full, True)
            self.assertEqual((old['keys'], old['rows'], old['count']),
                             (new['keys'], new['rows'], new['count']))
            self.assertEqual(old['count'][0], 385 if full else 379)
            self.assertEqual(old['count'][1:], (380, 561))
            for position in (0, len(old['rows'])//2, len(old['rows'])-1):
                self.assertEqual(old['rows'][position], new['rows'][position])
            selected = {old['keys'][0], old['keys'][len(old['keys'])//2], old['keys'][-1]}
            self.assertEqual([k for k in old['keys'] if k in selected],
                             [k for k in new['keys'] if k in selected])


if __name__ == '__main__':
    unittest.main()
