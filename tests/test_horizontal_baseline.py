"""Horizontal baseline contracts + independent format model; Excel execution stays pending."""
import copy
from pathlib import Path
import random
import unittest
from horizontal_baseline_model import META, decide, key, legacy_maps, load, pack, pivot, unpack

MODULE = (Path(__file__).resolve().parents[1] / 'vba/Module1.txt').read_text(encoding='utf-8-sig')


def proc(name):
    import re
    match = re.search(r'(?ms)^(?:Private|Public) (Sub|Function) ' + name + r'\b.*?^End \1', MODULE)
    if match is None:
        raise AssertionError(f'VBA procedure not found: {name}')
    return match.group(0)


class HorizontalBaselineFormatTests(unittest.TestCase):
    def setUp(self):
        self.records = [('row-1', '保存期間', '3年', '2026/10/29', '1'), ('row-1', '備考', '', '2026/10/29', '1'), ('row-2', 'タイトル', '00123', '2026/10/29', '1')]

    def test_roundtrip_preserves_all_confirmed_values(self):
        self.assertEqual(load(pivot(self.records)), legacy_maps(self.records))

    def test_confirmed_blank_is_not_an_unknown_field(self):
        result = load(pivot(self.records))
        self.assertIn(key('row-1', '備考'), result)
        self.assertEqual(result[key('row-1', '備考')], '')
        self.assertNotIn(key('row-2', '備考'), result)

    def test_empty_legacy_is_header_only(self):
        self.assertEqual(pivot([]), [META])
        self.assertEqual(load(pivot([])), {})

    def test_sort_and_insert_do_not_change_record_identity(self):
        data = pivot(self.records)
        data[1:] = reversed(data[1:])
        self.assertEqual(load(data), legacy_maps(self.records))
        new_records = self.records + [('row-new', '備考', 'new', '', '1')]
        self.assertEqual(load(pivot(new_records)), legacy_maps(new_records))

    def test_record_id_comparison_and_header_aliases_match_existing_contract(self):
        records = [('Abc', '第1ガイド', '001', '', '1')]
        self.assertEqual(load(pivot(records))[key('abc', '分類名２')], '001')

    def test_duplicate_record_id_in_horizontal_is_rejected(self):
        data = pivot(self.records)
        data.append(data[1].copy())
        with self.assertRaises(ValueError): load(data)

    def test_duplicate_legacy_item_is_rejected(self):
        with self.assertRaises(ValueError): pivot(self.records + [self.records[0]])

    def test_duplicate_canonical_column_is_rejected(self):
        data = pivot(self.records)
        data[0][5] = data[0][4]
        with self.assertRaises(ValueError): load(data)

    def test_current_sheet_copy_is_not_accepted(self):
        with self.assertRaises(ValueError): load([['通し番号', 'タイトル', '同期状態', 'Excel管理行ID']])

    def test_bad_flags_versions_and_timestamps_are_rejected(self):
        for column, value in [(1, 'HB9'), (2, 'x' * 3), (2, '1'), (3, '20:x')]:
            data = pivot(self.records)
            data[1][column] = value
            with self.subTest(column=column, value=value), self.assertRaises(ValueError): load(data)

    def test_unknown_field_cannot_smuggle_a_value(self):
        data = pivot(self.records)
        data[1][-1] = 'unconfirmed'
        with self.assertRaises(ValueError): load(data)

    def test_pack_handles_colons_newlines_and_unicode(self):
        items = ['', '12:東京\nvalue', '\U0001f600']
        self.assertEqual(unpack(pack(items), len(items)), items)

    def test_all_five_user_examples_are_identical_old_vs_new(self):
        old = legacy_maps(self.records)
        new = load(pivot(self.records))
        k = key('row-1', '保存期間')
        examples = [('5年', '3年', 'excel_only_hold'), ('3年', '5年', 'csv_only_apply'), ('5年', '10年', 'conflict_hold'), ('5年', '5年', 'confirmed'), ('5年', '3年', 'excel_only_hold')]
        for excel,csv,expected in examples:
            self.assertEqual(decide(old[k],excel,csv),expected)
            self.assertEqual(decide(new[k],excel,csv),expected)

    def test_sparse_synthetic_1000_records_all_decisions_match(self):
        rng = random.Random(147)
        records = [(f'id-{r}', f'項目{c}', rng.choice(['', '3年', '001', '=literal']), '', '1') for r in range(1000) for c in range(12) if rng.random() > .1]
        old, new = legacy_maps(records), load(pivot(records))
        self.assertEqual(old, new)
        for k in old:
            e,c = rng.choice(['', '3年', '5年']),rng.choice(['', '3年', '10年'])
            self.assertEqual(decide(old[k],e,c),decide(new[k],e,c))


class HorizontalBaselineSourceContracts(unittest.TestCase):
    def test_one_bulk_sheet_read_no_cell_reads_in_map_loop(self):
        self.assertEqual(proc('HorizontalBaselineReadData').count('.Value2'), 1)
        self.assertNotIn('.Cells(', proc('HorizontalBaselineMapsFromData'))
        self.assertIn('baselineRows.Add key, Array(CLng(r), CLng(c))', proc('HorizontalBaselineMapsFromData'))

    def test_migration_reads_only_old_baselines_and_compares_every_value(self):
        body = proc('HorizontalBaselineDataFromLegacy')
        self.assertIn('BidirectionalSyncLoadBaselineMaps legacy, values, locations', body)
        self.assertIn('If values.Count <> locations.Count', body)
        self.assertIn('For Each key In values.Keys', body)
        self.assertNotIn('SHEET_NEW_FILE_STANDARD', body)
        self.assertNotIn('.Delete', body)

    def test_absence_and_confirmed_empty_are_distinct_in_vba(self):
        body = proc('HorizontalBaselineMapsFromData')
        self.assertIn('Case "1"', body)
        self.assertIn('Case "0"', body)
        self.assertIn('baselineValues.Add key, CStr(data(r, c))', body)
        self.assertIn('identities.Exists(rowId)', body)

    def test_vba_rejects_formulas_and_unknown_schema(self):
        self.assertIn('hasFormula = ws.UsedRange.HasFormula',proc('HorizontalBaselineReadData'))
        self.assertIn('HorizontalBaselineMetadataHeaders()',proc('HorizontalBaselineMapsFromData'))
        self.assertIn('CStr(data(r, 2)) <> HORIZONTAL_BASELINE_VERSION',proc('HorizontalBaselineMapsFromData'))


if __name__ == '__main__': unittest.main()
