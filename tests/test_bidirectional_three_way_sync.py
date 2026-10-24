# -*- coding: utf-8 -*-
"""第2段階: 3方向比較・上書き防止の回帰契約。"""

from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE = ROOT / "vba" / "Module1.txt"


def vba_block(source: str, declaration: str) -> str:
    """指定したVBA Sub/Function本体を静的契約として抜き出す。"""
    start = source.index(declaration)
    end = source.index("End Sub", start) if " Sub " in declaration else source.index("End Function", start)
    return source[start:end]


def decide(baseline: str | None, excel: str, csv: str) -> str:
    """VBAのB/E/C判定仕様を、合成値で独立検証する。"""
    if excel == csv:
        return "confirmed"
    if baseline is None:
        return "initial_hold"
    excel_changed = excel != baseline
    csv_changed = csv != baseline
    if excel_changed and not csv_changed:
        return "excel_only_hold"
    if not excel_changed and csv_changed:
        return "csv_only_apply"
    if excel_changed and csv_changed:
        return "conflict_hold"
    raise AssertionError("Values differ although neither side changed from baseline")


def confirm_initial_group(excel_values: dict[str, str], csv_values: dict[str, str], saved_baselines: dict[str, str], *, saved_conflict: bool = False):
    """Model first-run behavior: equal members may seed baselines; mismatches remain held as a group."""
    baselines = dict(saved_baselines)
    review = []
    for header, excel_value in excel_values.items():
        csv_value = csv_values[header]
        if excel_value == csv_value:
            if not saved_conflict:
                baselines[header] = csv_value
        else:
            review.append(header)
    return baselines, review, 0  # no Excel writes occur during initial mismatch review


class TestBidirectionalThreeWaySync(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = MODULE.read_text(encoding="utf-8")

    def test_five_required_three_way_examples(self) -> None:
        # Excelだけ: 継続→5年はExcelを保持し、システム手動編集待ち。
        self.assertEqual(decide("継続", "5年", "継続"), "excel_only_hold")
        # CSVだけ: 継続→5年はExcelへ反映し、同期基準も5年へ更新。
        self.assertEqual(decide("継続", "継続", "5年"), "csv_only_apply")
        # 両方別値: Excel 5年、CSV 10年は競合保留。
        self.assertEqual(decide("継続", "5年", "10年"), "conflict_hold")
        # 基準未確定の不一致は初期保留。
        self.assertEqual(decide(None, "5年", "継続"), "initial_hold")
        # 両者とも前回値のままなら、値も基準値も維持する。
        self.assertEqual(decide("継続", "継続", "継続"), "confirmed")
        # システム手動編集後に値が一致すれば、その値を基準として確定できる。
        self.assertEqual(decide("継続", "5年", "5年"), "confirmed")

    def test_initial_mismatch_confirms_equal_members_only_and_holds_related_group_atomically(self) -> None:
        body = vba_block(self.source, "Private Sub BidirectionalSyncApplyManagedGroup")
        start = body.index("If Not allBaselineKnown Then")
        stop = body.index("Exit Sub", start) + len("Exit Sub")
        initial_branch = body[start:stop]
        self.assertIn("BidirectionalSyncHasSavedConflictForGroup", initial_branch)
        self.assertIn("If StrComp(excelValue, csvValue, vbTextCompare) = 0 Then", initial_branch)
        self.assertIn("BidirectionalSyncStoreBaseline rowId, CStr(key), csvValue", initial_branch)
        self.assertIn("mismatchedKeys", initial_branch)
        self.assertIn('BidirectionalSyncAppendGroupDecision "初期不一致"', initial_branch)
        self.assertIn("mismatchedKeys, groupName", initial_branch)
        self.assertNotIn("WriteStep2SourceValue", initial_branch)

        excel = {"分類名２": "道路", "分類コード２": "0007"}
        csv = {"分類名２": "道路", "分類コード２": "0008"}
        baselines, review, writes = confirm_initial_group(excel, csv, {})
        self.assertEqual(baselines, {"分類名２": "道路"})
        self.assertEqual(review, ["分類コード２"])
        self.assertEqual(writes, 0)
        self.assertEqual(excel["分類コード２"], "0007")

        # After a user resolves the mismatch so both sides agree, a repeat run confirms the group.
        excel["分類コード２"] = csv["分類コード２"]
        baselines, review, writes = confirm_initial_group(excel, csv, baselines)
        self.assertEqual(baselines, csv)
        self.assertEqual(review, [])
        self.assertEqual(writes, 0)

    def test_initial_mismatch_uses_persisted_e_choices_without_breaking_legacy_state(self) -> None:
        append = vba_block(self.source, "Private Sub BidirectionalSyncAppendGroupDecision")
        self.assertIn('decision = "競合保留" Or decision = "初期不一致"', append)

        capture = vba_block(self.source, "Private Function BidirectionalSyncCaptureConflictCase")
        self.assertIn("baselineKnown = baselineValues.Exists(baselineKey)", capture)
        self.assertIn('fieldItem.Add "baselineKnown", baselineKnown', capture)

        pack = vba_block(self.source, "Private Function BidirectionalSyncPackConflictHeaders")
        unpack = vba_block(self.source, "Private Function BidirectionalSyncUnpackConflictHeaders")
        self.assertIn('BidirectionalSyncPackConflictValue("BSC2")', pack)
        self.assertIn("If hasMissingBaseline Then", pack)
        self.assertIn('If formatMarker = "BSC2" Then', unpack)
        self.assertIn('fieldItem.Add "baselineKnown", True', unpack)
        self.assertIn('fieldItem.Add "baselineKnown", False', unpack)

        live = vba_block(self.source, "Private Function BidirectionalSyncConflictLiveMatches")
        self.assertIn('If baselineKnown Then', live)
        self.assertIn('If baselineValues.Exists(baselineKey) Or Len(CStr(fieldItem("baselineValue"))) > 0 Then', live)
        resolve = vba_block(self.source, "Public Function BidirectionalSyncResolveConflict")
        for option_name in ("BIDIRECTIONAL_SYNC_CONFLICT_EXCEL", "BIDIRECTIONAL_SYNC_CONFLICT_SYSTEM", "BIDIRECTIONAL_SYNC_CONFLICT_HOLD"):
            self.assertIn(option_name, resolve)

        refresh = vba_block(self.source, "Private Sub BidirectionalSyncRefreshConflictState")
        self.assertIn('If CStr(prior(0)) = stableKey And CStr(prior(1)) = snapshot Then decision = CStr(prior(2))', refresh)
        step4_approved = vba_block(self.source, "Private Function BidirectionalSyncConflictApprovedForStep4")
        step4_held = vba_block(self.source, "Private Function BidirectionalSyncConflictBlocksStep4")
        self.assertIn('CStr(ws.Cells(r, 6).Value2) = BIDIRECTIONAL_SYNC_CONFLICT_EXCEL', step4_approved)
        self.assertIn('decision = BIDIRECTIONAL_SYNC_CONFLICT_UNRESOLVED Or decision = BIDIRECTIONAL_SYNC_CONFLICT_HOLD', step4_held)

    def test_partial_baseline_confirmation_does_not_rebase_a_saved_conflict(self) -> None:
        body = vba_block(self.source, "Private Sub BidirectionalSyncApplyManagedGroup")
        start = body.index("If Not allBaselineKnown Then")
        stop = body.index("Exit Sub", start) + len("Exit Sub")
        initial_branch = body[start:stop]
        self.assertLess(initial_branch.index("BidirectionalSyncHasSavedConflictForGroup"), initial_branch.index("BidirectionalSyncStoreBaseline"))
        baselines, review, writes = confirm_initial_group(
            {"分類名２": "道路", "分類コード２": "0007"},
            {"分類名２": "道路", "分類コード２": "0008"},
            {}, saved_conflict=True,
        )
        self.assertEqual(baselines, {})
        self.assertEqual(review, ["分類コード２"])
        self.assertEqual(writes, 0)

    def test_baseline_is_keyed_by_excel_row_id_and_header(self) -> None:
        self.assertIn('BIDIRECTIONAL_SYNC_BASELINE_SHEET As String = "双方向同期基準値"', self.source)
        self.assertIn('"Excel管理行ID"', self.source)
        self.assertIn('"見出し名"', self.source)
        self.assertIn('BidirectionalSyncBaselineKey', self.source)
        self.assertIn('BidirectionalSyncStoreBaseline', self.source)

    def test_excel_csv_and_conflict_paths_are_explicit(self) -> None:
        for required in (
            '"初期不一致"',
            '"システム手動編集待ち"',
            '"競合保留"',
            'WriteStep2SourceValue wsDst.Cells(dstRow, dstCol)',
            'BidirectionalSyncStoreBaseline',
        ):
            self.assertIn(required, self.source)

    def test_related_name_and_id_fields_are_atomic(self) -> None:
        for required in (
            '"保存期間・保存期間ID"',
            '"分類名１・分類コード１"',
            '"第1ガイド・分類コード２"',
            '"第2ガイド・分類コード３"',
            '"媒体種別・媒体種別ID"',
            '"原課・原課ID"',
            '"関連する名称・コード項目の見出しが不足または重複しているため、片方だけを更新しません。"',
        ):
            self.assertIn(required, self.source)

    def test_existing_update_path_uses_context_instead_of_unconditional_write(self) -> None:
        update_start = self.source.index('Private Sub UpdateRowFromSourceByHeader')
        update_end = self.source.index('End Sub', update_start)
        update_body = self.source[update_start:update_end]
        self.assertIn('BidirectionalSyncContextIsActive()', update_body)
        self.assertIn('BidirectionalSyncApplyThreeWayToRow', update_body)
        self.assertIn('BidirectionalSyncStartContext wsDst, targetYearValue, bidirectionalOutputPath', self.source)
        self.assertIn('BidirectionalSyncFinishContext(bidirectionalSummary)', self.source)

    def test_formal_and_name_matches_remain_distinguishable_and_holds_are_listed(self) -> None:
        for required in (
            '"正式ID一意一致"',
            '"名称一意一致（登録予定）"',
            '"同一性未確定"',
            'BidirectionalSyncAppendManualLine',
            'BidirectionalSyncRunForYear',
        ):
            self.assertIn(required, self.source)

    def test_baseline_write_is_deferred_until_report_and_processing_succeed(self) -> None:
        start_body = vba_block(self.source, "Private Sub BidirectionalSyncStartContext")
        store_body = vba_block(self.source, "Private Sub BidirectionalSyncStoreBaseline")
        finish_body = vba_block(self.source, "Private Function BidirectionalSyncFinishContext")
        self.assertIn('"PendingBaselines"', start_body)
        self.assertIn('BidirectionalSyncCommitPendingBaselines', finish_body)
        self.assertLess(
            finish_body.index("BidirectionalDiffWriteUtf8Csv"),
            finish_body.index("BidirectionalSyncCommitPendingBaselines"),
        )
        self.assertNotIn("ws.Cells", store_body)
        self.assertIn('Set pending = mBidirectionalSyncContext("PendingBaselines")', store_body)

    def test_baseline_commit_has_rollback_and_discards_new_sheet_on_failure(self) -> None:
        commit_body = vba_block(self.source, "Private Sub BidirectionalSyncCommitPendingBaselines")
        discard_body = vba_block(self.source, "Private Sub BidirectionalSyncDiscardContext")
        self.assertIn("Rollback:", commit_body)
        self.assertIn("originalFormulas", commit_body)
        self.assertIn("ReDim originalNumberFormats", commit_body)
        self.assertIn("originalNumberFormats(formatRow, formatCol)", commit_body)
        self.assertIn("復元できません", commit_body)
        self.assertIn("BidirectionalSyncDeleteNewBaselineSheet", discard_body)

    def test_new_veryhidden_baseline_is_hidden_before_cleanup_delete(self) -> None:
        delete_body = vba_block(self.source, "Private Sub BidirectionalSyncDeleteNewBaselineSheet")
        self.assertIn("baselineSheet.Parent Is ThisWorkbook", delete_body)
        self.assertIn("ThisWorkbook.ReadOnly Or ThisWorkbook.ProtectStructure", delete_body)
        self.assertIn("If baselineSheet.Visible = xlSheetVeryHidden Then", delete_body)
        self.assertIn("baselineSheet.Visible = xlSheetHidden", delete_body)
        self.assertLess(
            delete_body.index("baselineSheet.Visible = xlSheetHidden"),
            delete_body.index("baselineSheet.Delete"),
        )
        self.assertIn("If visibilityChanged Then baselineSheet.Visible = originalVisibility", delete_body)

    def test_cleanup_failure_retains_the_original_operation_error(self) -> None:
        run_body = vba_block(self.source, "Public Function 双方向編集_3方向比較を実行")
        finish_body = vba_block(self.source, "Private Function BidirectionalSyncFinishContext")
        for body in (run_body, finish_body):
            self.assertIn("originalErrorNumber = Err.Number", body)
            self.assertIn("originalErrorDescription = Err.Description", body)
            self.assertIn("On Error Resume Next", body)
            self.assertIn("cleanupErrorNumber = Err.Number", body)
            self.assertIn("CleanupError:", body)
            self.assertIn("元のエラー", body)
            self.assertIn("後片付けにも失敗しました", body)

    def test_name_match_and_row_id_safety_are_retained(self) -> None:
        run_body = vba_block(self.source, "Private Sub BidirectionalSyncRunForYear")
        self.assertIn("BidirectionalDiffTryFormalIdentityMatch", run_body)
        self.assertIn("BidirectionalDiffTryUniqueNameMatch", run_body)
        self.assertIn("同一性未確定", run_body)
        self.assertIn("BidirectionalDiffFindExcelRowIdColumn", self.source)

    def test_stage1_report_and_existing_candidate_review_are_not_replaced(self) -> None:
        self.assertIn('Public Sub 双方向編集_項目差異一覧CSVを外部出力()', self.source)
        self.assertIn('modSyncReview.SyncReviewAfterStep2Success', self.source)
        self.assertIn('BidirectionalDiffTryUniqueNameMatch', self.source)


if __name__ == "__main__":
    unittest.main()
