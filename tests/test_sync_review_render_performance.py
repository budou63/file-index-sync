"""Static contracts + executable display-cost model; no Excel runtime claims."""
from pathlib import Path
import random
import re
import unittest
from benchmark_sync_review_display import Sheet, OldLookup, FrameLookup, SelectionGate, aliases, run_operation

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / 'vba/frmSyncReview.txt').read_text(encoding='utf-8')


def proc(name):
    match = re.search(r'(?ms)^(?:Private|Public) (?:Sub|Function) ' + name + r'\b.*?^End (?:Sub|Function)', FORM)
    assert match, name
    return match.group()


class RenderPerformanceContracts(unittest.TestCase):
    def test_display_cache_has_synchronous_lifetime_and_error_cleanup(self):
        for name in ('ShowECategory', 'ShowECurrent', 'UpdateEFieldSelection'):
            source = proc(name)
            self.assertIn('BeginEDisplayRead', source)
            self.assertIn('EndEDisplayRead', source)
            self.assertIn('On Error GoTo DisplayFailed', source)
            self.assertIn('EDisplayFailed Err.Description', source)
            self.assertNotIn('DoEvents', source)
        end = proc('EndEDisplayRead')
        self.assertIn('Set eDisplayHeaderMaps = Nothing', end)
        self.assertIn('Set eDisplayNormalizedKeys = Nothing', end)
        failed = proc('EDisplayFailed')
        self.assertIn('EndEDisplayRead', failed)
        self.assertIn('rendering = False', failed)
        for button in ('mCmdApprove', 'mCmdOther', 'mCmdHold'):
            self.assertIn(button + '.Enabled = False', failed)

    def test_header_array_read_and_duplicate_fail_closed(self):
        source = proc('EDisplayHeaderMap')
        self.assertIn('values = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value2', source)
        self.assertIn('If lastCol = 1 Then', source)
        self.assertIn('If Not IsError(rawValue) Then', source)
        self.assertIn('headerMap(headerKey) = 0', source)
        lookup = proc('FindCompareColumn')
        self.assertIn('If colNo = 0 Then Exit Function', lookup)
        self.assertIn('If matchedColumns.Count = 1 Then', lookup)
        self.assertNotIn('ws.Cells', lookup)

    def test_cells_are_still_read_live_with_display_format(self):
        cell = proc('CompareCellText')
        self.assertIn('rawValue = ws.Cells(rowNo, colNo).Value2', cell)
        self.assertIn('ws.Cells(rowNo, colNo).Text', cell)
        self.assertNotIn('eDisplayHeaderMaps', cell)
        for name in ('EDisplayHeaderMap', 'FindCompareColumn', 'BeginEDisplayRead', 'EndEDisplayRead'):
            self.assertNotRegex(proc(name), r'(?im)^\s*\w+\.(?:Range|Cells)\(.*\)\.(?:Value2|Formula)\s*=')
            self.assertNotIn('BidirectionalSyncResolveConflict', proc(name))

    def test_event_coalescing_does_not_drop_ctrl_shift_or_keyboard_paths(self):
        current = proc('ShowECurrent')
        self.assertLess(current.index('If selectionKey = eRenderedSelection Then Exit Sub'), current.index('mLstRows.Clear'))
        stamp = proc('ESelectionRenderKey')
        self.assertIn('mCboEVisibility.ListIndex', stamp)
        self.assertIn('mLstCases.ListIndex', stamp)
        self.assertIn('If mLstCases.Selected(i)', stamp)
        for name in ('mLstCases_Change', 'mLstCases_Click', 'mLstCases_MouseUp', 'mLstCases_KeyUp'):
            self.assertIn('ShowCurrent', proc(name))
        for name in ('mLstCases_MouseDown', 'mLstCases_KeyDown', 'LoadCases', 'ShowECategory',
                     'mCmdRefresh_Click', 'SelectSingleECase', 'mCmdEClear_Click'):
            self.assertIn('eRenderedSelection = ""', proc(name))

    def test_adoption_never_trusts_display_cache_or_render_token(self):
        for name in ('ESelectedDecisionIsSafe', 'EConfirmClassification1Decision', 'EConfirmRetentionDecision',
                     'mCmdApprove_Click', 'RecordChoice'):
            self.assertNotIn('eDisplayHeaderMaps', proc(name))
            self.assertNotIn('eRenderedSelection', proc(name))
        self.assertIn('Module1.BidirectionalSyncClassification1DecisionSafe', proc('ESelectedDecisionIsSafe'))
        self.assertIn('Module1.BidirectionalSyncRetentionDecisionSafe', proc('ESelectedDecisionIsSafe'))
        self.assertIn('UpdateESelection', proc('mLstCases_Exit'))
        self.assertIn('UpdateESelection', proc('UpdateEFieldSelection'))

    def test_scroll_and_identity_selection_restoration_retained(self):
        source = proc('ShowECategory')
        for token in ('priorTopIndex = mLstCases.TopIndex', 'mLstCases.TopIndex = priorTopIndex',
                      'ERecordSelectionKey(record)', 'selectedRecordKeys.Exists(recordKey)',
                      'mLstRows.ListIndex = i - 1', 'BuildERecords(cases, targetYear)'):
            self.assertIn(token, source)
        self.assertNotIn('BuildERecords', proc('ShowECurrent'))
        self.assertNotIn('mLstCases.Clear', proc('ShowECurrent'))

    def test_all_restore_code_bodies_match(self):
        for name, encoding in (('frmSyncReview.frm', 'utf-8'), ('frmSyncReview-native.frm', 'cp932')):
            text = (ROOT / name).read_text(encoding=encoding)
            self.assertEqual(text[text.index('Option Explicit'):], FORM)
        FORM.encode('cp932')


class RenderModelEquivalence(unittest.TestCase):
    def test_actual_alias_fixture_is_parsed_including_parentheses(self):
        spec = aliases()
        self.assertEqual(spec[5]['csv'], ['作成年度（和暦）', '作成年度(和暦)'])
        self.assertEqual(len(spec), 9)
        self.assertEqual(spec[9]['csv'], [])

    def test_missing_duplicate_ambiguous_error_and_single_column(self):
        for headers, names in [(['Ａ', 'A'], ['A']), (['A', 'B'], ['A', 'B']),
                               ([None, 'A'], ['Ａ']), (['A'], ['A', 'Ａ', 'A']),
                               (['A'], ['B']), ([''], ['']), ([None], ['A'])]:
            a, b = Sheet('a', headers), Sheet('b', headers)
            self.assertEqual(OldLookup().column(a, names), FrameLookup().column(b, names))

    def test_randomized_lookup_equivalence(self):
        rng = random.Random(147)
        vocabulary = ['A', 'Ａ', ' A ', 'B', '分類名３', '分類名3', 'タイトル', '', None]
        for _ in range(400):
            headers = rng.choices(vocabulary, k=rng.randint(1, 45))
            names = rng.choices(vocabulary[:-1], k=rng.randint(1, 5))
            self.assertEqual(OldLookup().column(Sheet('before', headers), names),
                             FrameLookup().column(Sheet('after', headers), names))

    def test_next_frame_observes_header_change(self):
        sheet = Sheet('CSV', ['A'])
        frame = FrameLookup()
        self.assertEqual(frame.column(sheet, ['A']), 1)
        sheet.headers = ['B']
        self.assertEqual(FrameLookup().column(sheet, ['A']), 0)
        self.assertEqual(FrameLookup().column(sheet, ['B']), 1)

    def test_mouse_keyboard_ctrl_shift_focus_and_mode_transitions(self):
        gate = SelectionGate()
        for mode, focus, selected in [(0, 0, {0}), (0, 4, {0, 4}), (0, 3, {0, 1, 2, 3}),
                                      (0, 3, set()), (1, 3, {3}), (0, 3, {3})]:
            gate.reset()
            self.assertTrue(gate.event(mode, focus, selected))
            self.assertFalse(gate.event(mode, focus, selected))
            self.assertFalse(gate.event(mode, focus, selected))
        self.assertEqual(gate.rebuilds, 6)
        gate.reset()  # clicking the same file again must still refresh
        self.assertTrue(gate.event(0, 3, {3}))
        self.assertTrue(gate.event(0, 3, {2, 3}))  # identical focus, changed selection

    def test_all_requested_operations_have_equal_model_display(self):
        for n in (1, 17, 379):
            for operation in ('open', 'first', 'middle', 'last', 'continuous', 'input_to_all', 'all_to_input'):
                old, old_reads, old_rebuilds = run_operation(operation, False, n, 12)
                new, new_reads, new_rebuilds = run_operation(operation, True, n, 12)
                self.assertEqual(old, new, (n, operation))
                self.assertLessEqual(new_rebuilds, old_rebuilds)
                self.assertLess(new_reads.get('header_range_reads', 0), old_reads['header_cell_value_reads'])


if __name__ == '__main__':
    unittest.main()
