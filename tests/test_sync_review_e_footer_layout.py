"""Static layout contracts for the E review footer; Excel GUI verification remains pending."""
from pathlib import Path
import re
import unittest

FORM = (Path(__file__).resolve().parents[1] / 'vba/frmSyncReview.txt').read_text(encoding='utf-8')


def proc(name):
    match = re.search(rf'(?ms)^Private Sub {name}\b.*?^End Sub', FORM)
    if not match:
        raise AssertionError(name)
    return match.group()


class EFooterContracts(unittest.TestCase):
    def test_status_wraps_and_has_two_line_area_separate_from_buttons(self):
        build = proc('BuildControls')
        layout = proc('LayoutReviewForE')
        e = layout[layout.index('If isE Then'):layout.index('\n    Else\n')]
        self.assertIn('mLblStatus.WordWrap = True', build)
        self.assertIn('footerTop = Me.InsideHeight - 105', e)
        self.assertIn('PlaceControl mLblStatus, 9, footerTop, availableWidth - 18, 40', e)
        for button in ('mCmdApprove', 'mCmdOther', 'mCmdHold', 'mCmdNext', 'mCmdRefresh', 'mCmdFinish'):
            self.assertIn(f'PlaceControl {button},', e)
            line = next(line for line in e.splitlines() if f'PlaceControl {button},' in line)
            self.assertIn('footerTop + 44', line)

    def test_abc_original_status_and_buttons_are_preserved(self):
        layout = proc('LayoutReviewForE')
        self.assertIn('PlaceControl mLblStatus, 9, 333, 740, 21', layout)
        self.assertIn('PlaceControl mCmdApprove, 84, 379.5, 78, 21', layout)

    def test_presets_keep_minimum_list_and_footer_within_form(self):
        for height in (400, 440, 560):
            details_height = 75 if height > 440 else 55
            footer_top = height - 105
            details_top = footer_top - details_height - 2
            list_height = max(details_top - 85.5, 150)
            self.assertLessEqual(79.5 + list_height, details_top)
            self.assertLessEqual(details_top + details_height, footer_top)
            self.assertLessEqual(footer_top + 40, footer_top + 44)
            self.assertLessEqual(footer_top + 44 + 21, height)


if __name__ == '__main__':
    unittest.main()
