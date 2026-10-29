"""Source contracts and a requirements oracle; no VBA/Excel execution."""
import re
import unittest
from pathlib import Path
SRC=(Path(__file__).resolve().parents[1]/'vba/Module1.txt').read_text(encoding='utf-8-sig')
def proc(name):
    m=re.search(r'(?:Public|Private) (?:Sub|Function) '+name+r'\b.*?\nEnd (?:Sub|Function)', SRC,re.S)
    assert m, name
    return m.group()
class StateProjection(unittest.TestCase):
    def test_review_form_accepts_linked_statuses_without_relaxing_identity(self):
        module=(Path(__file__).resolve().parents[1]/'vba/modSyncReview.txt').read_text(encoding='utf-8-sig')
        self.assertIn('Private Function ReviewIsLinkedSyncStatus',module)
        self.assertIn('Case "同期済", "要確認", "反映待"',module)
        self.assertIn('If Not ReviewIsLinkedSyncStatus(ReviewField(dst, matchedRow, "同期状態")) Then Exit Function',module)
        self.assertIn('If ReviewField(dst, matchedRow, "ファイルID") <> fileId Then Exit Function',module)

    def test_related_pair_hold_survives_later_decision(self):
        self.assertIn('SyncStateRelatedPairsComplete(transferMap)',proc('SyncStateCalculate'))
        guard=proc('SyncStateRelatedPairsComplete')
        self.assertIn('BidirectionalSyncCollectMappedKeys',guard)
        self.assertIn('itemKeys.Count <> 0 And itemKeys.Count <> 2',guard)
        for header in ('分類名１','分類コード１','第1ガイド','分類コード２','第2ガイド','分類コード３',
                       '保存期間','保存期間ID','媒体種別','媒体種別ID','原課','原課ID'):
            self.assertIn(header,guard)

    def test_code_notation_hold_survives_decision_outside_sync_context(self):
        s=proc('SyncStateCalculate')
        self.assertIn('SyncStateCodeIsConfirmed',s)
        guard=proc('SyncStateCodeIsConfirmed')
        self.assertIn('vbString',guard)
        self.assertIn('vbBinaryCompare',guard)
        self.assertIn('HasFormula',guard)
        self.assertIn('分類コード２',guard)
        self.assertIn('汎用項目１０',guard)

    def test_finish_projects_after_history_before_baselines(self):
        s=proc('BidirectionalSyncFinishContext')
        self.assertLess(s.index('BidirectionalSyncRefreshConflictState'),s.index('SyncStateApplyMatchedRows'))
        self.assertLess(s.index('SyncStateApplyMatchedRows'),s.index('BidirectionalSyncCommitPendingBaselines'))
    def test_tracking_resolves_stable_ids_after_row_and_column_moves(self):
        s=proc('BidirectionalSyncApplyThreeWayToRow')
        self.assertIn('SyncStateTrackMatchedRow',s)
        s=proc('SyncStateApplyMatchedRows')
        self.assertIn('SyncStateFindDestinationRowByMap',s)
        self.assertNotIn('SyncStateFindDestinationRow(wsDst',s)
        calc=proc('SyncStateCalculate')
        self.assertIn('BidirectionalDiffFindCachedColumn',calc)
        self.assertIn('NormalizeValueForUpdateCompare',calc)
        self.assertIn('BidirectionalSyncConflictReviewCases',calc)
        self.assertIn('IsStep2UpdateExcludedHeader',calc)
    def test_sync_state_projection_reuses_cached_baseline_rows_during_context(self):
        calc = proc('SyncStateCalculate')
        branch = calc[calc.index('If BidirectionalSyncContextIsActive() Then'):calc.index('Set acceptedHeaders')]
        active_context, fallback = branch.split('Else', 1)
        self.assertIn('Set baselineRows = mBidirectionalSyncContext("BaselineRows")', active_context)
        self.assertIn('Set wsBase = mBidirectionalSyncContext("BaselineSheet")', active_context)
        self.assertNotIn('BidirectionalSyncLoadBaselineMaps', active_context)
        self.assertIn('BidirectionalSyncLoadBaselineMaps wsBase, baselineValues, baselineRows', fallback)
        self.assertIn('baselineRows.Exists(baselineKey)', calc)
        self.assertIn('wsBase.Cells(CLng(baselineRows(baselineKey)), 3).Value2', calc)
        apply_rows = proc('SyncStateApplyMatchedRows')
        align_cases = proc('SyncStateAlignCases')
        self.assertIn('For Each key In rows.Keys', apply_rows)
        self.assertIn('SyncStateCalculate', apply_rows)
        self.assertIn('SyncStateCurrentDestinationRowMap(wsDst)', apply_rows)
        self.assertIn('SyncStateCurrentDestinationRowMap(wsDst)', align_cases)
        self.assertIn('SyncStateFindDestinationRowByMap', align_cases)

    def test_destination_row_index_is_cached_and_preserves_duplicate_failure(self):
        current = proc('SyncStateCurrentDestinationRowMap')
        build = proc('SyncStateBuildDestinationRowMap')
        lookup = proc('SyncStateFindDestinationRowByMap')
        self.assertIn('mBidirectionalSyncContext.Exists("DestinationRowMap")', current)
        self.assertIn('mBidirectionalSyncContext.Add "DestinationRowMap"', current)
        self.assertIn('Set rowMap = SyncStateBuildDestinationRowMap(ws)', current)
        self.assertIn('If rowMap.Exists(rowId) Then', build)
        self.assertIn('rowMap(rowId) = CLng(0)', build)
        self.assertIn('If CLng(rowMap(rowId)) <= 0 Then Err.Raise', lookup)
        self.assertIn('If Not rowMap.Exists(rowId) Then Err.Raise', lookup)

    def test_status_participates_in_single_and_batch_rollback(self):
        single=proc('BidirectionalSyncResolveConflict')
        self.assertIn('SyncStateUpdateDecision',single)
        self.assertIn('oldSyncStatus',single)
        self.assertIn('syncStatusCell.Value2 = oldSyncStatus',single)
        self.assertIn('oldSyncStatus',proc('BidirectionalSyncPlanFileDecisions'))
        self.assertIn('syncStatusAddress',proc('BidirectionalSyncRestoreFileDecisions'))
        self.assertIn('SyncStateUpdateDecision',proc('BidirectionalSyncResolveFileDecisions'))
    def test_pending_is_field_scoped_and_review_takes_precedence(self):
        s=proc('SyncStateCalculate')
        self.assertIn('acceptedHeaders',s)
        self.assertIn('baselineValues',s)
        self.assertIn('"反映待"',s)
        self.assertIn('"要確認"',s)
        self.assertIn('"同期済"',s)
    def test_yoshida_sequences_model_only(self):
        def classify(excel,csv,base=None,accepted=False,other_problem=False):
            if other_problem: return '要確認'
            if excel==csv: return '同期済'
            if accepted or (base is not None and csv==base): return '反映待'
            return '要確認'
        e=('工務','3'); c=('土木','2')
        self.assertEqual(classify(e,c),'要確認')
        self.assertEqual(classify(e,c,accepted=True),'反映待')
        self.assertEqual(classify(e,c,accepted=True),'反映待')
        self.assertEqual(classify(e,e),'同期済')
        self.assertEqual(classify(c,c),'同期済')
        self.assertEqual(classify(e,c,accepted=True,other_problem=True),'要確認')
        self.assertEqual(classify(e,c,base=c),'反映待')
        self.assertEqual(classify(('総務','4'),c),'要確認')
