import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class DOpenerFallbackTests(unittest.TestCase):
    def test_only_missing_d_headers_are_recoverable_and_abc_candidates_survive(self):
        src = (ROOT / 'vba/modSyncReview.txt').read_text(encoding='utf-8')
        m = re.search(r'(?ms)^Private Sub OpenSyncReviewForYear\(.*?^End Sub', src)
        self.assertIsNotNone(m)
        opener = m.group()
        self.assertIn('SyncReviewCandidates(yearValue)', opener)
        self.assertIn('SyncReviewCopyIdCandidates(yearValue)', opener)
        self.assertIn('vbObjectError + 6141', opener)
        self.assertIn('Err.Description', opener)
        self.assertRegex(opener, r'(?s)If .*vbObjectError \+ 6141.*Then.*?Resume ')
        self.assertIn('frmSyncReview.LoadCases cases, yearValue,', opener)
        self.assertIn('If cases.Count = 0 Then', opener)
        self.assertIn('If Len(unavailableReason) > 0 Then MsgBox', opener)
        self.assertLess(opener.index('SyncReviewCandidates(yearValue)'), opener.index('SyncReviewCopyIdCandidates(yearValue)'))
        self.assertLess(opener.index('SyncReviewCopyIdCandidates(yearValue)'), opener.index('If cases.Count = 0 Then'))


if __name__ == '__main__':
    unittest.main()
