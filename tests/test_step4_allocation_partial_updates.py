"""Step4 allocation-failure recovery: decision example and VBA source wiring.

This does not execute Excel/VBA; the source assertions bind the expected
failure path to the public Step4 procedure and its targeted checkpoint.
"""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE, re.DOTALL,
    )
    if match is None:
        raise AssertionError(name)
    return match.group(0)


class Step4AllocationFailureTests(unittest.TestCase):
    def test_9400002_no_gap_path_checkpoints_before_refill_and_output(self) -> None:
        # Excel fixture for workplace verification: a registration candidate starts
        # with code 9400002; an empty master lookup clears it, then a missing
        # upper guide (or an exhausted gap) makes allocation fail. Source-level
        # assertions alone cannot prove the actual workbook is restored.
        step4 = procedure("手順4_差分インポートCSVを作成する")
        refill = procedure("新ファイル基準表_分類コード再補完")
        self.assertIn('wsEdit.Cells(r, codeCols(i)).Value = codeValue', refill)
        self.assertIn("CaptureStep4AllocationColumns", step4)
        self.assertIn("RestoreStep4AllocationColumns", step4)
        self.assertIn("RestoreCodeMasterFromBackup", step4)
        self.assertLess(step4.index("CaptureStep4AllocationColumns"),
                        step4.index("ThisWorkbook.Worksheets(SHEET_CODE_MANAGEMENT).Copy"))
        self.assertLess(step4.index("ThisWorkbook.Worksheets(SHEET_CODE_MANAGEMENT).Copy"),
                        step4.index("ApplyStep4ConflictResolution wsEdit, conflictPlan"))
        self.assertLess(step4.index("CaptureStep4AllocationColumns"),
                        step4.index('wsEdit.Cells(CLng(targetRow), syncStatusColEdit).Value = "登録予定"'))
        self.assertLess(step4.index("CaptureStep4AllocationColumns"),
                        step4.index("新ファイル基準表_分類コード再補完 wsEdit"))
        self.assertLess(step4.index("新ファイル基準表_登録予定分類コード採番 wsEdit"),
                        step4.index("wsOut.Cells.Clear"))
        self.assertIn("If allocationInProgress Then", step4)

    def test_refill_errors_reach_step4_failure_handler(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("新ファイル基準表_年度と原課ID再補完 wsEdit, CLng(targetRow), CLng(targetRow), True, True", step4)
        self.assertIn("新ファイル基準表_分類コード再補完 wsEdit, CLng(targetRow), CLng(targetRow), True", step4)
        for name in ("新ファイル基準表_年度と原課ID再補完", "新ファイル基準表_分類コード再補完"):
            helper = procedure(name)
            self.assertIn("Optional ByVal raiseOnError As Boolean = False", helper)
            self.assertIn("If raiseOnError And trappedErrNumber <> 0 Then Err.Raise trappedErrNumber", helper)

    def test_master_restore_preserves_row_dimensions_and_visibility(self) -> None:
        restore = procedure("RestoreCodeMasterFromBackup")
        self.assertIn("wsDst.Rows(rowNo).RowHeight = wsBackup.Rows(rowNo).RowHeight", restore)
        self.assertIn("wsDst.Rows(rowNo).RowHeight = wsUnfiltered.Rows(rowNo).RowHeight", restore)
        self.assertIn("wsDst.Rows(rowNo).Hidden = wsBackup.Rows(rowNo).Hidden", restore)

    def test_checkpoint_covers_working_codes_year_department_and_status(self) -> None:
        checkpoint = procedure("CaptureStep4AllocationColumns")
        restore = procedure("RestoreStep4AllocationColumns")
        for column in ("分類コード１", "分類コード２", "分類コード３", "年度", "原課ID", "同期状態"):
            with self.subTest(column=column):
                self.assertIn('"' + column + '"', checkpoint)
        self.assertIn(".Formula", checkpoint)
        self.assertIn(".Formula", restore)

    def test_prior_step4_safety_and_parent_scoped_allocation_remain(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertLess(step4.index("PreflightStep4ConflictResolution"),
                        step4.index("ApplyStep4ConflictResolution wsEdit, conflictPlan"))
        self.assertLess(step4.index("ValidateStep4GuideCodeStates"),
                        step4.index("新ファイル基準表_登録予定分類コード採番 wsEdit"))
        self.assertIn("If syncStatusChangedCount < 0 Then GoTo SafeExit", step4)
        self.assertIn('guideKey = BuildGuideCodeMasterNameKey("分類名３", NormalizeClassificationGuideKey(guideName), parentCodeValue)',
                      procedure("新ファイル基準表_登録予定分類コード採番"))
        self.assertIn("変更が一部反映された可能性があります。保存せずExcelを終了してください。", step4)


if __name__ == "__main__":
    unittest.main()
