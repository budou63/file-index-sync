"""Protect prior review state against auto-clear on non-formal identity matches."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / 'vba/Module1.txt').read_text(encoding='utf-8-sig')


def proc(name):
    match = re.search(r'(?ims)^(?:Private|Public) (Sub|Function) ' + re.escape(name) + r'\b.*?^End \1', SOURCE)
    if not match:
        raise AssertionError(name)
    return match.group(0)


def model(prior, calculated, formal):
    if prior == '要確認' and (calculated != '同期済' or not formal):
        return '要確認'
    return calculated


class AutoClearGuardTests(unittest.TestCase):
    def test_old_review_is_only_cleared_by_formal_identity_in_model(self):
        self.assertEqual(model('要確認', '同期済', True), '同期済')
        self.assertEqual(model('要確認', '同期済', False), '要確認')
        self.assertEqual(model('要確認', '要確認', True), '要確認')
        self.assertEqual(model('要確認', '反映待', True), '要確認')
        self.assertEqual(model('要確認', '反映待', False), '要確認')
        self.assertEqual(model('登録予定', '同期済', False), '同期済')

    def test_context_stores_individual_identity_matches(self):
        start = proc('BidirectionalSyncStartContext')
        self.assertIn('context.Add "StateNonFormalRows"', start)
        track = proc('SyncStateTrackMatchedRow')
        self.assertIn('BidirectionalSyncCurrentMatchKind()', track)
        self.assertIn('"正式ID一意一致"', track)
        self.assertIn('nonFormalRows(rowId) = True', track)

    def test_application_preserves_review_only_for_unsafe_auto_clear(self):
        body = proc('SyncStateApplyMatchedRows')
        self.assertIn('SyncStateCalculate(', body)
        self.assertIn('StateNonFormalRows', body)
        self.assertIn('SyncStateFindStatusColumn(wsDst)', body)
        self.assertIn('GetSyncStatusValue(wsDst, dstRow, statusCol)', body)
        self.assertIn('nonFormalRows.Exists(CStr(key))', body)
        self.assertIn('oldStatus = "要確認"', body)
        self.assertIn('calculatedStatus <> "同期済"', body)
        self.assertIn('calculatedStatus = "要確認"', body)
        self.assertIn('SyncStateWrite wsDst, dstRow, calculatedStatus', body)


if __name__ == '__main__':
    unittest.main()
