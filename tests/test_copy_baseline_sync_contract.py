"""VBA source contracts for the copy-shaped baseline; no Excel is executed."""
import re
import unittest
from pathlib import Path

SOURCE = (Path(__file__).resolve().parents[1] / 'vba/Module1.txt').read_text(encoding='utf-8-sig')

def proc(name):
    match = re.search(r'(?ims)^(?:Private|Public) (Sub|Function) ' + re.escape(name) + r'\b.*?^End \1', SOURCE)
    if not match:
        raise AssertionError('missing procedure: ' + name)
    return match.group(0)

class CopyBaselineSyncContract(unittest.TestCase):
    def test_existing_sheet_is_validated_not_renamed_or_replaced(self):
        body = proc('HorizontalBaselineEnsureSheet')
        self.assertIn('HorizontalBaselineMatchesDestinationHeaders', body)
        self.assertNotIn('HorizontalBaselineDataFromLegacy', body)
        self.assertNotIn('HorizontalBaselineUniqueArchiveName', body)
        self.assertIn('xlSheetVeryHidden', body)

    def test_new_sheet_is_data_only_and_snapshots_formulas(self):
        body = proc('HorizontalBaselineCopyWithoutSheetEvents')
        self.assertIn('Worksheets.Add', body)
        self.assertIn('.Value2', body)
        self.assertIn('HorizontalBaselineDataMatches', body)
        self.assertNotRegex(body, r'(?i)wsDst\.Copy\b')
        self.assertIn('Application.EnableEvents = oldEvents', body)

    def test_baseline_load_uses_id_and_aliases_not_row_position(self):
        body = proc('CopyBaselineLoadMaps')
        for token in ['Excel管理行ID', 'BidirectionalSyncBaselineKey', 'GetStep2HeaderCandidates', 'BidirectionalDiffFindCachedColumn', 'NormalizeValueForUpdateCompare']:
            self.assertIn(token, body)

    def test_partial_new_row_knownness_is_durable_and_not_blank_equals_known(self):
        body = proc('CopyBaselineLoadMaps')
        self.assertIn('CopyBaselineLoadKnownFlags', body)
        self.assertIn('CopyBaselineIsKnownField', body)
        commit = proc('CopyBaselineApplyEntries')
        self.assertIn('CopyBaselineWriteKnownFlags', commit)
        self.assertIn('Application.EnableEvents = oldEvents', commit)
        self.assertIn('Rollback', commit)
        self.assertIn('CopyBaselineRestoreUndo', commit)

    def test_pending_values_are_verified_against_live_excel_before_commit(self):
        plan = proc('CopyBaselinePrepareEntries')
        for token in ['SyncStateFindDestinationRowByMap', 'NormalizeValueForUpdateCompare', 'entry(2)', 'Value2']:
            self.assertIn(token, plan)
        apply = proc('CopyBaselineApplyEntries')
        self.assertIn('CopyBaselineCellMatches', apply)
        self.assertLess(apply.index('CopyBaselinePrepareEntries'), apply.index('target.Value2 ='))

    def test_system_acceptance_updates_baseline_transactionally(self):
        body = proc('BidirectionalSyncResolveConflict')
        self.assertIn('CopyBaselineConfirmResolvedSystem', body)
        self.assertLess(body.index('CopyBaselineConfirmResolvedSystem'), body.index('BidirectionalSyncResolveConflict = True'))

    def test_missing_partial_metadata_fails_closed(self):
        loader = proc('CopyBaselineLoadKnownFlags')
        for token in ['COPY_BASELINE_FLAGS_SHEET', 'COPY_BASELINE_FLAGS_MARKER',
                      'COPY_BASELINE_FLAGS_PROPERTY', 'COPY_BASELINE_LOCAL_MARKER',
                      'ws Is Nothing Or marker Is Nothing Or sentinel Is Nothing Or localMarker Is Nothing',
                      'xlSheetVeryHidden', 'Err.Raise']:
            self.assertIn(token, loader)
        cached = proc('CopyBaselineValidateCachedFlag')
        self.assertIn('localMarker.RefersTo', cached)
        self.assertIn('location(7)', cached)
        self.assertIn('location(4)', cached)

    def test_partial_metadata_creation_and_undo_preserve_sentinels(self):
        writer = proc('CopyBaselineWriteKnownFlags')
        undo = proc('CopyBaselineRestoreFlags')
        for token in ['CustomDocumentProperties.Add', 'Names.Add', 'COPY_BASELINE_LOCAL_MARKER',
                      'CopyBaselineRestoreFlags undo', 'Application.EnableEvents = oldEvents']:
            self.assertIn(token, writer)
        for token in ['COPY_BASELINE_LOCAL_MARKER', 'COPY_BASELINE_FLAGS_PROPERTY',
                      'COPY_BASELINE_FLAGS_MARKER', 'Application.EnableEvents = oldEvents']:
            self.assertIn(token, undo)

    def test_comparison_uses_same_value_semantics_as_existing_three_way_engine(self):
        load = proc('CopyBaselineLoadMaps')
        plan = proc('CopyBaselinePrepareEntries')
        persisted = proc('BidirectionalSyncPersistedBaselineValue')
        self.assertIn('Cells(lastRow, lastCol)).Value', load)
        self.assertNotIn('Cells(lastRow, lastCol)).Value2', load)
        self.assertIn('NormalizeValueForUpdateCompare(header, source.Value)', plan)
        self.assertIn('NormalizeValueForUpdateCompare(sourceHeader, ws.Cells(r, c).Value)', persisted)
        self.assertIn('actualValue = source.Value2', plan)

    def test_format_only_paste_and_formula_like_literals_are_safe(self):
        create = proc('HorizontalBaselineCopyWithoutSheetEvents')
        self.assertIn('target.PasteSpecial Paste:=xlPasteFormats', create)
        self.assertNotIn('sourceRange.Copy Destination:=', create)
        self.assertLess(create.index('CopyBaselineRequiresTextFormat'), create.index('target.Value2 = sourceValues'))
        prepare = proc('CopyBaselinePrepareEntries')
        self.assertIn('CopyBaselineRequiresTextFormat(actualValue)', prepare)
        self.assertIn('sourceFormat = "@"', prepare)
        self.assertIn('target.HasFormula', proc('CopyBaselineApplyEntries'))

    def test_batch_adoption_and_failure_restore_same_baseline_transaction(self):
        batch = proc('BidirectionalSyncResolveFileDecisions')
        self.assertIn('CopyBaselineConfirmResolvedBatch plan, baselineUndo', batch)
        self.assertIn('CopyBaselineRestoreUndo baselineUndo', batch)
        self.assertLess(batch.index('CopyBaselineConfirmResolvedBatch'), batch.index('SyncStateUpdateDecision'))
        self.assertIn('CopyBaselineRestoreUndo baselineUndo', proc('BidirectionalSyncResolveConflict'))
        helper = proc('CopyBaselineConfirmResolvedBatch')
        self.assertIn('BIDIRECTIONAL_SYNC_CONFLICT_SYSTEM', helper)
        self.assertIn('CopyBaselineApplyEntries', helper)

    def test_rollback_restores_formula_like_literals_without_evaluating_them(self):
        undo = proc('CopyBaselineRestoreUndo')
        self.assertIn('CopyBaselineRequiresTextFormat(entry("oldValue"))', undo)
        safe = undo.index('CopyBaselineRequiresTextFormat(entry("oldValue"))')
        write = undo.index('target.Value2 = entry("oldValue")')
        restore_format = undo.index('target.NumberFormat = CStr(entry("oldFormat"))', write)
        self.assertLess(safe, write)
        self.assertGreater(restore_format, write)
        self.assertIn('If target.HasFormula Then Err.Raise', undo)

    def test_missing_previously_initialized_baseline_never_gets_recreated(self):
        ensure = proc('HorizontalBaselineEnsureSheet')
        missing = ensure.split('If ws Is Nothing Then', 1)[1].split('End If', 1)[0]
        self.assertLess(missing.index('CopyBaselineHasInitializationMarker'),
                        missing.index('HorizontalBaselineCopyWithoutSheetEvents'))
        self.assertIn('Err.Raise', missing)
        self.assertLess(missing.index('CopyBaselineInstallInitializationMarker'),
                        missing.index('HorizontalBaselineCopyWithoutSheetEvents'))
        self.assertIn('CopyBaselineInstallInitializationMarker', ensure)
        check = proc('CopyBaselineHasInitializationMarker')
        self.assertIn('CustomDocumentProperties', check)
        self.assertIn('COPY_BASELINE_INITIALIZED_PROPERTY', check)

    def test_replacement_baseline_cannot_impersonate_initialized_sheet(self):
        ensure = proc('HorizontalBaselineEnsureSheet')
        self.assertIn('CopyBaselineEnsureInitializationPair ws', ensure)
        pair = proc('CopyBaselineEnsureInitializationPair')
        self.assertIn('CopyBaselineHasInitializationMarker', pair)
        self.assertIn('CopyBaselineHasSheetInitializationMarker', pair)
        self.assertIn('Err.Raise', pair)
        self.assertIn('CopyBaselineInstallSheetInitializationMarker', pair)
        local = proc('CopyBaselineHasSheetInitializationMarker')
        self.assertIn('ws.Names', local)
        self.assertIn('COPY_BASELINE_INITIALIZED_LOCAL', local)

    def test_no_ambiguous_aliases_or_unidentified_data(self):
        load = proc('CopyBaselineLoadMaps')
        self.assertIn('mappedColumns.Exists(CStr(c))', load)
        self.assertIn('CopyBaselineValidateData ws, False', load)
        data = proc('CopyBaselineValidateData')
        self.assertIn('seen.Exists(rowId)', data)
        self.assertIn('ws.Cells(r, idCol).HasFormula', data)
        self.assertIn('IsNull(hasFormula)', data)

if __name__ == '__main__':
    unittest.main()
