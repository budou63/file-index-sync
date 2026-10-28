# -*- coding: utf-8 -*-
"""同期状態のファイル単位遷移に関する静的契約。"""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / "vba" / "Module1.txt"


def procedure(source: str, name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        source,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"{name} が見つかりません")
    return match.group(0)


def file_status(*, is_new: bool, unresolved: bool, excel_adopted_pending: bool,
                all_synced: bool) -> str:
    """Requirements-level model; VBA runtime behavior is separately unverified."""
    if is_new:
        return "登録予定"
    if unresolved:
        return "要確認"
    if excel_adopted_pending:
        return "反映待"
    if all_synced:
        return "同期済"
    return "要確認"


class TestSyncStatusAutoupdate(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = MODULE.read_text(encoding="utf-8")

    def test_start_context_stores_boolean_as_dictionary_value(self) -> None:
        start = procedure(self.source, "BidirectionalSyncStartContext")
        self.assertIn("context.Add \"BaselineSheetCreated\", CBool(baselineSheetCreated)", start)
        self.assertNotIn("Set context(\"BaselineSheetCreated\") = CBool(baselineSheetCreated)", start)

    def test_step2_binds_transfer_map_after_context_initialization(self) -> None:
        step2 = procedure(self.source, '手順2_新ファイル基準表を作成する')
        build = step2.index('BuildStep2SourceTransferMap wsSrc')
        start = step2.index('BidirectionalSyncStartContext wsDst')
        assignment = 'Set mBidirectionalSyncContext("SyncTransferMap") = srcToDstMap'
        self.assertEqual(step2.count(assignment), 1)
        bind = step2.index(assignment)
        update = step2.index('UpdateSyncStatusColumn wsDst')
        self.assertLess(build, start)
        self.assertLess(start, bind)
        self.assertLess(bind, update)
        self.assertNotIn('If BidirectionalSyncContextIsActive() Then ' + assignment, step2)

    def test_yoshida_case_state_transitions(self) -> None:
        self.assertEqual(file_status(is_new=False, unresolved=True, excel_adopted_pending=False, all_synced=False), "要確認")
        self.assertEqual(file_status(is_new=False, unresolved=False, excel_adopted_pending=True, all_synced=False), "反映待")
        self.assertEqual(file_status(is_new=False, unresolved=False, excel_adopted_pending=True, all_synced=False), "反映待")  # stale CSV
        self.assertEqual(file_status(is_new=False, unresolved=False, excel_adopted_pending=False, all_synced=True), "同期済")
        self.assertEqual(file_status(is_new=False, unresolved=False, excel_adopted_pending=False, all_synced=False), "要確認")
        self.assertEqual(file_status(is_new=True, unresolved=False, excel_adopted_pending=False, all_synced=True), "登録予定")

    def test_new_statuses_are_normalized_and_legacy_states_remain(self) -> None:
        normalize = procedure(self.source, "NormalizeSyncStatusValue")
        for value in ("要確認", "反映待", "判定困難"):
            self.assertIn(f'"{value}"', normalize)
        for legacy in ('Case "現行", "同期済"', 'Case "登録予定", "未登録"', 'Case "削除予定", "削除候補"'):
            self.assertIn(legacy, normalize)
        ignored = procedure(self.source, "IsIgnoredBySyncStatus")
        self.assertIn('= "不明"', ignored)

    def test_excel_accept_sets_pending_but_system_adoption_copies_csv_values(self) -> None:
        resolve = procedure(self.source, "BidirectionalSyncResolveConflict")
        update_status = procedure(self.source, "BidirectionalSyncUpdateStatusAfterDecision")
        self.assertIn("If decision = BIDIRECTIONAL_SYNC_CONFLICT_SYSTEM Then", resolve)
        self.assertIn("WriteStep2SourceValue wsDst.Cells(dstRow, dstCol)", resolve)
        self.assertIn("BidirectionalSyncUpdateStatusAfterDecision", resolve)
        self.assertIn("syncStatusCaptured = True", resolve)
        self.assertIn("oldSyncStatus", procedure(self.source, "BidirectionalSyncResolveConflict"))
        self.assertIn("syncStatusCell.HasFormula Or IsError", resolve)
        self.assertIn("BidirectionalSyncCalculateFileStatus", update_status)
        calc = procedure(self.source, "BidirectionalSyncCalculateFileStatus")
        self.assertIn("BidirectionalSyncRowHasExcelAcceptedConflict", calc)

    def test_step2_aggregates_mapped_values_with_file_level_precedence(self) -> None:
        step2 = procedure(self.source, "手順2_新ファイル基準表を作成する")
        track = procedure(self.source, "BidirectionalSyncTrackMatchedRow")
        calc = procedure(self.source, "BidirectionalSyncCalculateFileStatus")
        remap = procedure(self.source, "BidirectionalSyncRemapTransferMapToCurrentColumns")
        update = procedure(self.source, "UpdateSyncStatusColumn")
        self.assertIn("BidirectionalSyncTrackMatchedRow wsDst, srcRow, targetRow", step2)
        self.assertIn("MatchedSyncRows", track)
        self.assertIn("Excel管理行ID", track)
        self.assertIn("sourceToDestinationMap.Keys", calc)
        self.assertIn("BidirectionalDiffBuildHeaderColumnMap", remap)
        self.assertIn("BidirectionalDiffFindCachedColumn", remap)
        self.assertIn("BidirectionalSyncRemapTransferMapToCurrentColumns", update)
        self.assertIn("NormalizeValueForUpdateCompare", calc)
        self.assertIn("BidirectionalSyncConflictBlocksStep4", calc)
        for value in ("要確認", "反映待", "同期済"):
            self.assertIn(f'"{value}"', calc)
        self.assertIn("BidirectionalSyncCalculateFileStatus", update)

    def test_batch_adoption_updates_and_rolls_back_sync_status(self) -> None:
        plan = procedure(self.source, "BidirectionalSyncPlanFileDecisions")
        restore = procedure(self.source, "BidirectionalSyncRestoreFileDecisions")
        resolve = procedure(self.source, "BidirectionalSyncResolveFileDecisions")
        self.assertIn('entry.Add "oldSyncStatus"', plan)
        self.assertIn('entry.Add "syncStatusAddress"', plan)
        self.assertIn('entry("oldSyncStatus")', restore)
        self.assertIn('entry("syncStatusAddress")', restore)
        self.assertIn("BidirectionalSyncUpdateStatusAfterDecision", resolve)

    def test_stale_csv_preserves_excel_choice_and_deferred_baseline_commit(self) -> None:
        group = procedure(self.source, "BidirectionalSyncApplyManagedGroup")
        self.assertIn("BidirectionalSyncHasSavedConflictForGroup", group)
        self.assertIn("BidirectionalSyncStoreBaseline", group)
        finish = procedure(self.source, "BidirectionalSyncFinishContext")
        self.assertLess(finish.index("BidirectionalSyncRefreshConflictState"), finish.index("BidirectionalSyncCommitPendingBaselines"))
        accepted = procedure(self.source, "BidirectionalSyncRowHasExcelAcceptedConflict")
        context_case = procedure(self.source, "BidirectionalSyncContextCaseIsUnresolved")
        self.assertIn("BidirectionalSyncConflictLiveMatches", accepted)
        self.assertIn("BidirectionalSyncFindConflictStateRow", context_case)
        self.assertIn("Not BidirectionalSyncConflictLiveMatches", context_case)

    def test_unverified_identity_keeps_existing_ambiguity_path(self) -> None:
        step2 = self.source[self.source.index("Private Sub BidirectionalSyncApplyThreeWayToRow"):]
        self.assertIn('"同期基準保存不可"', step2)
        self.assertIn('"Excel管理行IDが空欄のため、CSV値を反映しません。"', step2)
        self.assertIn('"判定困難"', self.source)

    def test_manual_guide_mentions_new_status_transitions(self) -> None:
        guide = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8")
        for value in ("登録予定", "要確認", "反映待", "同期済"):
            self.assertIn(value, guide)
        design = (ROOT / "docs" / "状態管理_登録・競合判断・差異反映_設計方針.md").read_text(encoding="utf-8")
        self.assertIn("旧設計メモ", design)
        self.assertIn("方針が優先される", design)
