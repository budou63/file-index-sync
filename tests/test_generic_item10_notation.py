"""Generic item 10 notation: static contracts and specification models; no Excel run."""
import re
import unittest
from pathlib import Path

SOURCE = (Path(__file__).resolve().parents[1] / 'vba' / 'Module1.txt').read_text(encoding='utf-8-sig')


def proc(name):
    match = re.search(r'(?ms)^Private (?:Function|Sub) ' + re.escape(name)
                      + r'\b.*?^End (?:Function|Sub)\s*$', SOURCE)
    if not match:
        raise AssertionError('Missing procedure: ' + name)
    return match.group(0)


def digits(value):
    if value is None:
        return None
    text = str(value)
    if not text or not text.isascii() or not text.isdigit():
        return None
    return text.lstrip('0') or '0'


def notation_model(csv, excel, *, exact_id=True, row_id_unique=True,
                   decision='none', snapshot_matches=True, baseline_unknown=True, formula=False):
    """Expected outcomes, not an execution of VBA or Excel."""
    if isinstance(csv, str) and isinstance(excel, str) and csv == excel and not formula:
        return excel
    if not exact_id or not row_id_unique or formula or not isinstance(csv, str):
        return excel
    if decision != 'none' and not (decision == 'unresolved' and snapshot_matches and baseline_unknown):
        return excel
    if digits(csv) is None or digits(excel) is None or digits(csv) != digits(excel):
        return excel
    return csv


class Generic10Tests(unittest.TestCase):
    def test_csv_notation_wins_and_second_run_is_stable(self):
        for excel in (90200, '90200'):
            fixed = notation_model('090200', excel)
            self.assertEqual(fixed, '090200')
            self.assertEqual(notation_model('090200', fixed), fixed)
        self.assertEqual(notation_model('0000', 0, decision='unresolved'), '0000')

    def test_blank_different_value_or_explicit_decision_never_autofixes(self):
        for value in ('', None, '90201', 90201, '90X200'):
            self.assertEqual(notation_model('090200', value), value)
        for decision in ('excel', 'system', 'hold'):
            self.assertEqual(notation_model('090200', 90200, decision=decision), 90200)
        for risk in ({'exact_id': False}, {'row_id_unique': False}, {'formula': True},
                     {'decision': 'unresolved', 'snapshot_matches': False},
                     {'decision': 'unresolved', 'baseline_unknown': False}):
            self.assertEqual(notation_model('090200', 90200, **risk), 90200)
        self.assertEqual(notation_model(90200, 90200), 90200)

    def test_only_one_generic_header_uses_existing_guarded_write_path(self):
        caller = proc('BidirectionalSyncApplySingleHeader')
        group = proc('BidirectionalSyncApplyManagedGroup')
        normalizer = proc('BidirectionalSyncPrepareClassificationCodeGroup')
        self.assertIn('BidirectionalSyncApplyManagedGroup', caller)
        self.assertIn('BidirectionalSyncPrepareClassificationCodeGroup', group)
        hold = group.index('If groupName = "汎用項目１０" And Not classificationCodeReady Then')
        self.assertIn('BidirectionalSyncAppendGroupDecision "競合保留"', group[hold:])
        self.assertLess(hold, group.index('For Each key In itemKeys', hold))
        self.assertLess(hold, group.index('ElseIf (Not excelChanged) And csvChanged Then'))
        self.assertIn('Case "汎用項目１０"', normalizer)
        for other in ('汎用項目１', '汎用項目９', '汎用項目１１'):
            self.assertNotIn('Case "' + other + '"', normalizer)
        for token in ('If itemKeys.Count <> 1 Then Exit Function', 'If codeMatchCount <> 1',
                      'BidirectionalSyncCurrentMatchKind()', '"正式ID一意一致"',
                      'BidirectionalSyncDestinationRowIdIsUnique',
                      'BidirectionalSyncDigitsOnlyCodeKey(excelRaw, excelCodeKey)',
                      'BidirectionalSyncHasSavedConflictForGroup',
                      'WriteStep2SourceValue excelCodeCell, codeKey, systemRaw'):
            self.assertIn(token, normalizer)
        self.assertNotIn('WriteStep2SourceValue excelNameCell', normalizer)
        self.assertIn('CanonicalHeaderName(CStr(wsDst.Cells(1, excelCodeCell.Column).Value2))', normalizer)
        self.assertIn('If excelCodeCell.MergeCells Then', normalizer)
        self.assertLess(normalizer.index('If excelCodeCell.MergeCells Then'),
                        normalizer.index('WriteStep2SourceValue excelCodeCell'))
        writer = proc('WriteStep2SourceValue')
        self.assertIn('targetCell.NumberFormat = "@"', writer)
        self.assertIn('targetCell.Value2 = CStr(sourceValue)', writer)

    def test_saved_conflict_requires_one_exact_unresolved_unknown_baseline(self):
        guard = proc('BidirectionalSyncSavedConflictAllowsCodeNotation')
        for token in ('savedHeaders.Count <> 1', 'BIDIRECTIONAL_SYNC_CONFLICT_UNRESOLVED',
                      'savedCase("stableKey")', 'savedHeader("baselineKnown")',
                      'savedHeader("systemValue")', 'savedHeader("excelValue")',
                      'If Not codeSeen Then Exit Function', 'BidirectionalSyncDigitsOnlyCodeKey',
                      'CanonicalHeaderName(destinationHeader)'):
            self.assertIn(token, guard)
        self.assertIn('If groupName = "汎用項目１０" And Not isGeneric10 Then Exit Function', guard)
        normalizer = proc('BidirectionalSyncPrepareClassificationCodeGroup')
        self.assertLess(normalizer.index('BidirectionalSyncHasSavedConflictForGroup'),
                        normalizer.index('BidirectionalSyncRecordGuideCodeSnapshot'))
        self.assertLess(normalizer.index('BidirectionalSyncRecordGuideCodeSnapshot'),
                        normalizer.index('WriteStep2SourceValue excelCodeCell'))
        self.assertIn('BidirectionalSyncRollbackGuideCodeSnapshots', proc('BidirectionalSyncDiscardContext'))

    def test_repeat_run_and_other_columns_stay_unchanged(self):
        group = proc('BidirectionalSyncApplyManagedGroup')
        refresh = proc('BidirectionalSyncRefreshConflictState')
        self.assertIn('If excelEqualsCsv Then', group)
        self.assertIn('BidirectionalSyncStoreBaseline rowId, CStr(key), csvValue', group)
        self.assertIn('For Each item In currentCases', refresh)
        self.assertIn('historyKey = BidirectionalSyncConflictHistoryKey(targetYearValue, rowId, stableKey, groupName, snapshot)', refresh)
        self.assertIn('If previousByKey.Exists(historyKey) Then decision = CStr(previousByKey(historyKey))', refresh)
        row = {'file_id': 'ID-A', 'row_id': 'ROW-A', 'other': 'untouched', 'item10': 90200}
        result = dict(row)
        result['item10'] = notation_model('090200', row['item10'])
        for key in ('file_id', 'row_id', 'other'):
            self.assertEqual(result[key], row[key])
        self.assertEqual(notation_model('090200', result['item10']), result['item10'])


if __name__ == '__main__':
    unittest.main()
