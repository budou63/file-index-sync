"""Source-contract regression tests for copy-shaped baseline; not Excel execution."""
import re
import unittest
from pathlib import Path

MODULE = (Path(__file__).resolve().parents[1] / 'vba/Module1.txt').read_text(encoding='utf-8-sig')


def proc(name):
    match = re.search(r'(?ms)^(?:Private|Public) (Sub|Function) ' + name + r'\b.*?^End \1', MODULE)
    if match is None:
        raise AssertionError(f'missing procedure: {name}')
    return match.group(0)


class CopyBaselineContract(unittest.TestCase):
    def test_existing_copy_not_renamed_and_legacy_is_not_migration_source(self):
        code = proc('HorizontalBaselineEnsureSheet')
        self.assertNotIn('HorizontalBaselineDataFromLegacy', code)
        self.assertNotIn('HorizontalBaselineUniqueArchiveName', code)
        self.assertNotIn('Set legacy =', code)

    def test_header_row_one(self):
        line = next(x for x in MODULE.splitlines() if x.startswith('Private Const HORIZONTAL_BASELINE_COPY_HEADER_ROW'))
        self.assertEqual('Private Const HORIZONTAL_BASELINE_COPY_HEADER_ROW As Long = 1', line)

    def test_no_legacy_fallback_when_copy_missing(self):
        code = proc('BidirectionalSyncExistingBaselineSheet')
        self.assertNotIn('BIDIRECTIONAL_SYNC_BASELINE_SHEET', code)

    def test_user_worksheet_event_code_not_copied_to_baseline(self):
        code = proc('HorizontalBaselineEnsureSheet')
        self.assertNotRegex(code, r'(?im)^\s*wsDst\.Copy\b')
        self.assertIn('HorizontalBaselineCopyWithoutSheetEvents', code)

    def test_saved_conflicts_protected_before_any_automatic_updates(self):
        code = proc('BidirectionalSyncApplyManagedGroup')
        self.assertIn('BidirectionalSyncProtectSavedDecision', code)
