"""Static source-path checks for the D-clear audit gate (not an Excel runtime test)."""
import re
import unittest
from pathlib import Path

MODULE = Path(__file__).resolve().parents[1] / "modSyncReview.bas"


def proc(name, end="End Function"):
    source = MODULE.read_text(encoding="utf-8")
    match = re.search(r"(?ms)^" + re.escape(name) + r".*?^" + end, source)
    if match is None:
        raise AssertionError(f"missing procedure: {name}")
    return match.group()


class AuditPreflightTests(unittest.TestCase):
    def test_nonwriting_preflight_requires_existing_audit_before_backup_and_clear(self):
        api = proc("Public Function SyncReviewClearCopyIdentifiers(")
        preflight = api.index("ReviewPreflightCopyAudit")
        backup = api.index('ReviewBackup("番号クリア前")')
        clear = api.index(".ClearContents")
        self.assertLess(preflight, backup)
        self.assertLess(backup, clear)
        self.assertEqual(api.count("ReviewPreflightCopyAudit"), 2)
        self.assertLess(backup, api.rindex("ReviewPreflightCopyAudit"))
        self.assertLess(api.rindex("ReviewPreflightCopyAudit"), clear)
        helper = proc("Private Sub ReviewPreflightCopyAudit(", "End Sub")
        self.assertIn("ReviewAudit(False)", helper)
        self.assertIn("If ws Is Nothing Then Err.Raise", helper)
        self.assertNotIn("ReviewAudit(True)", helper)
        self.assertIn("ws.ProtectContents", helper)
        self.assertIn("ws.Rows.Count", helper)
        self.assertIn(".MergeCells", helper)
        self.assertNotIn(".Value2 =", helper)
        self.assertNotIn(".ClearContents", helper)

    def test_events_remain_suppressed_during_completion_audit(self):
        api = proc("Public Function SyncReviewClearCopyIdentifiers(")
        self.assertLess(api.index('ReviewWriteCopyClearAudit item, "番号クリア完了"'),
                        api.index("Application.EnableEvents = prevEvents"))

    def test_events_remain_suppressed_during_failure_audit(self):
        api = proc("Public Function SyncReviewClearCopyIdentifiers(")
        failed = api[api.index("Failed:"):]
        self.assertLess(failed.index('ReviewWriteCopyClearAudit item, "番号クリア失敗"'),
                        failed.index("Application.EnableEvents = prevEvents"))
        self.assertIn("If eventsCaptured Then", failed)

    def test_audit_append_finds_formula_cells_without_clobbering_existing_records(self):
        locator = proc("Private Function ReviewCopyAuditNextRow(")
        self.assertIn("LookIn:=xlFormulas", locator)
        self.assertIn("SearchFormat:=False", locator)
        self.assertIn(".Row + 1", locator)
        audit = proc("Private Sub ReviewWriteCopyClearAudit(", "End Sub")
        self.assertIn("r = ReviewCopyAuditNextRow(ws)", audit)
        self.assertNotIn("r = ReviewLast(ws) + 1", audit)
        self.assertLess(audit.index("ws.Cells(r, 12).Value2 = backup"), audit.index("ws.Cells(r, 5).Value2 = decision"))

    def test_copy_audit_does_not_duplicate_keeper_in_row_list(self):
        audit = proc("Private Sub ReviewWriteCopyClearAudit(", "End Sub")
        self.assertIn('ws.Cells(r, 7).Value2 = CStr(item("rows"))', audit)
        self.assertNotIn('CStr(item("copyKeeperRow")) & ";"', audit)

    def test_audit_failure_is_critical_and_three_cell_rollback_retained(self):
        api = proc("Public Function SyncReviewClearCopyIdentifiers(")
        self.assertIn("auditStarted = True", api)
        self.assertRegex(api, r"If auditStarted Then\s+[\s\S]*?criticalFailure = True")
        self.assertIn("If Err.Number <> 0 Then", api)
        self.assertIn("criticalFailure = True", api)
        self.assertIn("ReviewRestoreCopyIdCells", api)
        self.assertIn("ブックを保存せず終了", api)
        self.assertIn("事前バックアップ", api)
        self.assertEqual(api.count(".ClearContents"), 1)
        self.assertIn("For i = 1 To 3: dst.Cells(selectedRow, cols(i)).ClearContents: Next i", api)


if __name__ == "__main__":
    unittest.main()
