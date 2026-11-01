"""Copy-shaped baseline invariants. Synthetic schema/model plus static VBA checks; no Excel execution."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / 'vba/Module1.txt').read_text(encoding='utf-8-sig')


def proc(name):
    match = re.search(r'(?ims)^(?:Private|Public) (Sub|Function) ' + re.escape(name) + r'\b.*?^End \1', SOURCE)
    if not match:
        raise AssertionError(name)
    return match.group(0)


def model_schema(candidate, destination):
    """Simple synthetic reference model; not CanonicalHeaderName or VBA execution."""
    if len(candidate) < 2 or len(candidate) != len(destination):
        return False
    def normalize(value):
        return value.strip().casefold()
    left = [normalize(value) for value in candidate]
    right = [normalize(value) for value in destination]
    return (all(left) and all(right) and len(set(left)) == len(left) and
            len(set(right)) == len(right) and 'excel管理行id' in left and set(left) == set(right))


class BaselineSchemaTests(unittest.TestCase):
    def test_132_column_schema_and_mismatch_in_synthetic_model(self):
        headers = ['Excel管理行ID'] + [f'列{i}' for i in range(1, 132)]
        self.assertEqual(len(headers), 132)
        self.assertTrue(model_schema(headers, headers))
        self.assertTrue(model_schema(headers[::-1], headers))
        self.assertFalse(model_schema(headers[:-1], headers))
        self.assertFalse(model_schema(headers[:-1] + [headers[1]], headers))
        self.assertFalse(model_schema(headers[:-1] + [''], headers))

    def test_vba_schema_is_dynamic_and_based_on_first_row(self):
        self.assertIn('HORIZONTAL_BASELINE_COPY_HEADER_ROW As Long = 1', SOURCE)
        schema = proc('HorizontalBaselineMatchesDestinationHeaders')
        for token in ('candidateLastCol <> destinationLastCol', 'Array("Excel管理行ID")',
                      'matchCount <> 1', 'mapped.Exists(CStr(mappedCol))',
                      'candidate.Cells(headerRow, 1)', 'destination.Cells(headerRow, 1)'):
            self.assertIn(token, schema)
        self.assertNotRegex(schema, r'\b132\b')

    def test_vba_uses_unique_row_id_and_not_worksheet_row_as_stable_key(self):
        validate = proc('CopyBaselineValidateData')
        for token in ('For r = 2 To lastRow', 'seen.Exists(rowId)', 'Len(rowId) = 0',
                      'ws.Cells(r, idCol).HasFormula'):
            self.assertIn(token, validate)
        loader = proc('CopyBaselineLoadMaps')
        self.assertIn('BidirectionalSyncBaselineKey(rowId, header)', loader)
        self.assertIn('CopyBaselineIsKnownField(flags, rowId, header)', loader)

    def test_vba_pending_update_is_after_csv_write_and_history_refresh(self):
        finish = proc('BidirectionalSyncFinishContext')
        write = finish.index('BidirectionalDiffWriteUtf8Csv outputPath, lines')
        history = finish.index('BidirectionalSyncRefreshConflictState')
        commit = finish.index('BidirectionalSyncCommitPendingBaselines')
        self.assertLess(write, history)
        self.assertLess(history, commit)
        memory = proc('BidirectionalSyncStoreBaseline')
        self.assertIn('PendingBaselines', memory)
        self.assertNotIn('.Value2 =', memory)


if __name__ == '__main__':
    unittest.main()
