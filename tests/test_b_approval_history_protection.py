"""B approvals must not erase formulas or bypass retained B decisions (static only)."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / 'vba/modSyncReview.txt').read_text(encoding='utf-8')


def proc(name, kind='Function'):
    match = re.search(rf'(?ms)^(?:Public|Private) {kind} {name}\b.*?^End {kind}', SOURCE)
    if not match:
        raise AssertionError(name)
    return match.group()


class BApprovalProtection(unittest.TestCase):
    def test_prior_hold_and_not_same_are_rechecked_before_mutation(self):
        guard = proc('ReviewBApprovalDecisionAllowed')
        self.assertIn('SyncReviewDecision(item)', guard)
        self.assertIn('SyncReviewDecision(item, selectedRow)', guard)
        self.assertIn('"保留"', guard)
        self.assertIn('"別物"', guard)
        approve = proc('SyncReviewApprove')
        self.assertGreaterEqual(approve.count('ReviewBApprovalDecisionAllowed(item, selectedRow)'), 3)
        last_guard = approve.rindex('ReviewBApprovalDecisionAllowed(item, selectedRow)')
        self.assertLess(last_guard, approve.index('dst.Cells(selectedRow, cols(1)).Value2 = id'))
        self.assertIn('ReviewCurrent(item)', approve)

    def test_formula_cells_are_rejected_before_backup_and_writes(self):
        approve = proc('SyncReviewApprove')
        guard = approve.index('If wasFormula(j) Then Err.Raise')
        self.assertLess(guard, approve.index('backup = ReviewBackup()'))
        self.assertLess(guard, approve.index('dst.Cells(selectedRow, cols(1)).Value2 = id'))
        self.assertIn('ReviewWriteAudit item, "同一として処理"', approve)

    def test_different_folder_and_guide_names_are_not_written(self):
        approve = proc('SyncReviewApprove')
        self.assertIn('ReviewName(src, sRow, True, False)', approve)
        self.assertIn('ガイド・フォルダー・Excel独自列は変更しません', approve)
        self.assertNotIn('ReviewCol(dst, "個別フォルダー")', approve)
        self.assertNotIn('ReviewCol(dst, "第1ガイド")', approve)
        self.assertNotIn('ReviewCol(dst, "第2ガイド")', approve)


if __name__ == '__main__':
    unittest.main()
