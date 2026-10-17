"""Static contracts for Step2 cancellation propagation through batch execution."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def procedure(name: str, kind: str = "Sub") -> str:
    match = re.search(
        rf"(?ms)^(?:Private|Public) {kind} {re.escape(name)}\b.*?^End {kind}",
        SOURCE,
    )
    if match is None:
        raise AssertionError(f"missing {kind}: {name}")
    return match.group()


class Step2BatchCancellationTests(unittest.TestCase):
    def test_public_step2_preserves_compatibility_and_reports_three_outcomes(self):
        step2 = procedure("手順2_新ファイル基準表を作成する")
        self.assertTrue(SOURCE.startswith("Option Explicit"))
        self.assertIn("Public Sub 手順2_新ファイル基準表を作成する()", step2)
        self.assertIn("mStep2Result = STEP2_RESULT_CANCELLED", step2)
        self.assertIn("mStep2Result = STEP2_RESULT_COMPLETED", step2)
        self.assertIn("mStep2Result = STEP2_RESULT_FAILED", step2)
        self.assertLess(step2.index("mStep2Result = STEP2_RESULT_COMPLETED"), step2.index("step2Succeeded = True"))

    def test_step2_preconfirms_before_writes_when_schema_allows_and_defers_comments(self):
        step2 = procedure("手順2_新ファイル基準表を作成する")
        prompt = step2.index("Step2ConfirmMissingSyncStatus")
        self.assertLess(step2.index("If StrPtr(targetYearRaw) = 0 Then GoTo SafeExit"), step2.index("EnsureColumnByHeader"))
        self.assertLess(prompt, step2.index("BuildStep2SourceTransferMap"))
        self.assertLess(prompt, step2.index("For srcRow = 2 To srcLastRow"))
        self.assertIn("If syncStatusPromptReady Then", step2)
        self.assertIn("If Not syncStatusPromptReady Then", step2)
        self.assertGreater(step2.rindex("Step2ConfirmMissingSyncStatus"), step2.index("BuildStep2SourceTransferMap"))
        confirm = procedure("Step2ConfirmMissingSyncStatus", "Function")
        self.assertNotIn("SafeAddComment", confirm)
        comments = procedure("Step2AddMissingSyncStatusComments")
        self.assertIn("SafeAddComment", comments)
        self.assertLess(prompt, step2.index("Step2AddMissingSyncStatusComments"))

    def test_batch_runs_step3_and_success_only_effects_only_after_step2_success(self):
        batch = procedure("新ファイル基準表生成")
        step2_call = batch.index("Call 手順2_新ファイル基準表を作成する")
        gate = batch.index("If mStep2Result <> STEP2_RESULT_COMPLETED Then GoTo Cleanup")
        step3 = batch.index("Call 手順3_コード管理CSVを作成する(True)")
        success = batch.index("新ファイル基準表とコード管理CSVの作成が完了しました。")
        self.assertLess(step2_call, gate)
        self.assertLess(gate, step3)
        self.assertLess(step3, success)
        self.assertLess(success, batch.index("batchSucceeded = True"))
        self.assertIn("SyncReviewBeginBatch", batch)
        self.assertIn("If reviewBatchStarted Then modSyncReview.SyncReviewEndBatch batchSucceeded", batch)

    def test_error_and_cancel_paths_do_not_run_step3_or_report_batch_success(self):
        batch = procedure("新ファイル基準表生成")
        gate = batch.index("If mStep2Result <> STEP2_RESULT_COMPLETED Then GoTo Cleanup")
        self.assertLess(gate, batch.index("Call 手順3_コード管理CSVを作成する(True)"))
        self.assertLess(batch.index("Call 手順3_コード管理CSVを作成する(True)"), batch.index("batchSucceeded = True"))
        step2 = procedure("手順2_新ファイル基準表を作成する")
        error_handler = step2[step2.index("ErrorHandler:"):]
        self.assertLess(error_handler.index("mStep2Result = STEP2_RESULT_FAILED"), error_handler.index('"失敗"'))
        self.assertIn("If Not Step2ConfirmMissingSyncStatus", step2)
        self.assertIn("If StrPtr(targetYearRaw) = 0 Then GoTo SafeExit", step2)
        self.assertLess(step2.index("Step2ConfirmMissingSyncStatus"), step2.index('"成功"'))
        self.assertLess(step2.index("If StrPtr(targetYearRaw) = 0 Then GoTo SafeExit"), step2.index('"成功"'))

    def test_excel_settings_and_review_batch_are_restored_on_every_batch_exit(self):
        batch = procedure("新ファイル基準表生成")
        cleanup = batch[batch.index("Cleanup:"):]
        for restore in ("Application.ScreenUpdating = oldScreenUpdating", "Application.EnableEvents = oldEnableEvents", "RestoreCalculationIfPossible oldCalculation", "SyncReviewEndBatch batchSucceeded"):
            self.assertIn(restore, cleanup)
        self.assertLess(cleanup.index("Application.ScreenUpdating = oldScreenUpdating"), cleanup.index("Application.EnableEvents = oldEnableEvents"))
        self.assertLess(cleanup.index("Application.EnableEvents = oldEnableEvents"), cleanup.index("RestoreCalculationIfPossible oldCalculation"))
        self.assertLess(cleanup.index("RestoreCalculationIfPossible oldCalculation"), cleanup.index("SyncReviewEndBatch batchSucceeded"))
        step2 = procedure("手順2_新ファイル基準表を作成する")
        self.assertIn("step2ApplicationStateCaptured", step2)
        self.assertIn("If step2ApplicationStateCaptured Then", step2)
        self.assertIn("If step2WriteStarted Then RestoreColumnWidthsByHeader", step2)


if __name__ == "__main__":
    unittest.main()
