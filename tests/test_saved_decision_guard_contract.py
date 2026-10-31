"""Source contracts for history protection; these do not execute Excel VBA."""
import re
import unittest
from pathlib import Path

SOURCE = (Path(__file__).resolve().parents[1] / 'vba/Module1.txt').read_text(encoding='utf-8-sig')

def proc(name):
    found = re.search(r'(?ims)^(?:Private|Public) (Sub|Function) ' + re.escape(name) + r'\b.*?^End \1', SOURCE)
    if not found:
        raise AssertionError('missing procedure: ' + name)
    return found.group(0)

class SavedDecisionSourceContract(unittest.TestCase):
    def test_guard_precedes_match_and_any_csv_write(self):
        code = proc('BidirectionalSyncApplyManagedGroup')
        pos = code.index('BidirectionalSyncProtectSavedDecision')
        self.assertLess(pos, code.index('If excelEqualsCsv Then'))
        self.assertLess(pos, code.index('WriteStep2SourceValue'))

    def test_guard_checks_identity_and_preserves_or_reconfirms(self):
        code = proc('BidirectionalSyncProtectSavedDecision')
        for expected in ['BidirectionalSyncProtectHeldConflict', 'stableKey', 'BidirectionalSyncDestinationRowIdIsUnique', 'BidirectionalSyncBatchSourceKeyIsUnique', 'BidirectionalSyncAppendGroupDecision']:
            self.assertIn(expected, code)

    def test_history_lookup_reads_history_without_live_filter(self):
        code = proc('BidirectionalSyncLoadSavedConflictKeys')
        self.assertIn('BidirectionalSyncConflictSheet', code)
        self.assertIn('BidirectionalSyncUnpackConflictHeaders', code)
        self.assertNotIn('BidirectionalSyncConflictReviewCases', code)

    def test_auto_resolution_is_a_distinct_record_and_skipped_by_form(self):
        self.assertIn('BIDIRECTIONAL_SYNC_CONFLICT_AUTO_RESOLVED', SOURCE)
        self.assertIn('BIDIRECTIONAL_SYNC_CONFLICT_AUTO_RESOLVED', proc('BidirectionalSyncRefreshConflictState'))
        self.assertIn('BIDIRECTIONAL_SYNC_CONFLICT_AUTO_RESOLVED', proc('BidirectionalSyncConflictReviewCases'))

    def test_history_refresh_still_has_error_rollback_and_event_restore(self):
        code = proc('BidirectionalSyncRefreshConflictState')
        for part in ('oldData', 'Rollback', 'Application.EnableEvents = oldEvents'):
            if part == 'Rollback':
                self.assertIn('rollbackError', code)
            else:
                self.assertIn(part, code)

    def test_identity_and_field_validation_precede_code_normalization(self):
        code = proc('BidirectionalSyncApplyManagedGroup')
        self.assertLess(code.index('BidirectionalSyncValidateSavedIdentity'),
                        code.index('BidirectionalSyncPrepareClassificationCodeGroup'))
        validate = proc('BidirectionalSyncValidateSavedIdentity')
        for part in ('正式ID一意一致', 'stableKey', 'BidirectionalSyncDestinationRowIdIsUnique',
                     'BidirectionalSyncBatchSourceKeyIsUnique', 'destinationHeader', 'seen.Exists'):
            self.assertIn(part, validate)

    def test_prior_decision_reused_only_from_latest_state(self):
        code = proc('BidirectionalSyncRefreshConflictState')
        self.assertIn('If Not latestHistoryByCase.Exists(caseKey) Then latestHistoryByCase.Add caseKey, historyKey', code)
        self.assertLess(code.index('CStr(latestHistoryByCase(caseKey)) = historyKey'),
                        code.index('If previousByKey.Exists(historyKey) Then decision ='))
        self.assertIn('preservedRows.Add flatRow', code)

    def test_auto_closure_requires_equal_values_and_protection_checks(self):
        code = proc('BidirectionalSyncProtectSavedDecision')
        self.assertIn('If excelEqualsCsv And classificationCodeReady Then', code)
        self.assertLess(code.index('stableKey'), code.index('resolution.Add "autoResolved", True'))
        self.assertIn('saved("decision")) = BIDIRECTIONAL_SYNC_CONFLICT_AUTO_RESOLVED', code)

if __name__ == '__main__':
    unittest.main()
