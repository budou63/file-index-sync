"""Static guards for plain-arrow E selection; interactive behavior remains unverified."""
from pathlib import Path
import re
import unittest

FORM = (Path(__file__).resolve().parents[1] / 'vba/frmSyncReview.txt').read_text(encoding='utf-8-sig')


def proc(name):
    match = re.search(r'(?ims)^Private (Sub|Function) ' + re.escape(name) + r'\b.*?^End \1', FORM)
    if not match:
        raise AssertionError(name)
    return match.group(0)


class EKeyboardFocusTests(unittest.TestCase):
    def test_only_unmodified_up_down_arm_keyboard_selection(self):
        down = proc('mLstCases_KeyDown')
        for text in ('ePlainArrowKey = 0', 'mCboCategory.ListIndex = 4', 'Shift = 0',
                     'KeyCode = vbKeyUp Or KeyCode = vbKeyDown', 'ePlainArrowKey = CLng(KeyCode)'):
            self.assertIn(text, down)
        up = proc('mLstCases_KeyUp')
        for text in ('ePlainArrowKey <> 0', 'ePlainArrowKey = CLng(KeyCode)', 'Shift = 0',
                     'ESelectFocusedCase', 'ePlainArrowKey = 0', 'ShowCurrent'):
            self.assertIn(text, up)

    def test_focused_selection_is_guarded_against_reentrancy_and_bad_index(self):
        helper = proc('ESelectFocusedCase')
        for text in ('mLstCases.ListIndex', 'mLstCases.ListCount', 'rendering = True',
                     'mLstCases.Selected(i) = (i = focusIndex)', 'rendering = False',
                     'EDisplayFailed'):
            self.assertIn(text, helper)
        for forbidden in ('BidirectionalSyncResolveConflict', 'SyncReviewDecide', 'Value2 =', '.Save'):
            self.assertNotIn(forbidden, helper)


if __name__ == '__main__':
    unittest.main()
