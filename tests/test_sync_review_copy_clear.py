import copy
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")


def procedure(code, start, end="End Function"):
    match = re.search(r"(?ms)^" + re.escape(start) + r".*?^" + re.escape(end), code)
    if match is None:
        raise AssertionError(f"missing procedure {start}")
    return match.group()


def simulate_clear(case, row, sheet, system_unregistered=True, final_yes=True,
                   revalidation_ok=True, backup_ok=True, write_failure_at=None,
                   audit_ok=True, restore_ok=True):
    """Failure-injection oracle for the specified one-row, three-field contract."""
    before = copy.deepcopy(sheet)
    if case.get("category") != "D" or row == case.get("keeper") or row not in case.get("candidates", []):
        return False, sheet
    if sheet[row].get("status") != "登録予定" or not revalidation_ok:
        return False, sheet
    if not system_unregistered or not final_yes:
        return False, sheet
    if not revalidation_ok or not backup_ok:
        return False, sheet
    fields = ("id", "number", "barcode")
    try:
        for i, field in enumerate(fields, 1):
            if write_failure_at == i:
                raise RuntimeError("write")
            sheet[row][field] = ""
        if not audit_ok:
            raise RuntimeError("audit")
        return True, sheet
    except RuntimeError:
        if restore_ok:
            sheet.clear()
            sheet.update(before)
        return False, sheet


class CopyIdClearApiTests(unittest.TestCase):
    def test_public_api_is_d_only_candidate_specific_and_non_deleting(self):
        api = procedure(MODULE, "Public Function SyncReviewClearCopyIdentifiers(")
        self.assertIn('CStr(item("category")) <> "D"', api)
        self.assertIn("ReviewCopyCaseCurrent", api)
        self.assertIn("copyKeeperRow", api)
        self.assertIn("copyCandidateRows", api)
        self.assertNotIn(".Rows(", api)
        self.assertNotIn(".Delete", api)
        self.assertNotIn('"同期状態")).ClearContents', api)

    def test_confirmation_revalidation_backup_write_audit_order_is_explicit(self):
        api = procedure(MODULE, "Public Function SyncReviewClearCopyIdentifiers(")
        self.assertGreaterEqual(api.count("ReviewCopyCaseCurrent"), 3)
        self.assertEqual(api.count("MsgBox("), 2)
        self.assertIn("システム", api)
        self.assertIn("未登録", api)
        self.assertIn("vbDefaultButton2", api)
        self.assertLess(api.index("ReviewCopyCaseCurrent"), api.index("MsgBox("))
        first_yes = api.index("MsgBox(")
        second_yes = api.index("MsgBox(", first_yes + 1)
        backup = api.index('ReviewBackup("番号クリア前")')
        first_clear = api.index(".ClearContents")
        audit = api.index('ReviewWriteCopyClearAudit item, "番号クリア完了"')
        self.assertLess(first_yes, second_yes)
        self.assertLess(second_yes, backup)
        self.assertLess(backup, api.index("Application.EnableEvents = False"))
        self.assertLess(api.index("Application.EnableEvents = False"), first_clear)
        self.assertLess(first_clear, audit)
        self.assertLess(audit, api.index("SyncReviewClearCopyIdentifiers = True"))

    def test_three_cells_only_and_error_restoration_keep_status_and_formats(self):
        api = procedure(MODULE, "Public Function SyncReviewClearCopyIdentifiers(")
        self.assertIn('ReviewCol(dst, "ファイルID")', api)
        self.assertIn('ReviewCol(dst, "ファイル番号")', api)
        self.assertIn('ReviewCol(dst, "バーコード番号")', api)
        self.assertEqual(api.count(".ClearContents"), 1)
        self.assertIn("For i = 1 To 3: dst.Cells(selectedRow, cols(i)).ClearContents: Next i", api)
        self.assertIn("ReviewRestoreCopyIdCells", api)
        self.assertIn("Application.EnableEvents = prevEvents", api)
        restore = procedure(MODULE, "Private Function ReviewRestoreCopyIdCells(")
        for term in (".Formula", ".Value2", ".NumberFormat", "HasFormula"):
            self.assertIn(term, restore)

    def test_candidate_role_metadata_is_attached_by_detector_not_inferred_from_rows(self):
        detector = procedure(MODULE, "Public Function SyncReviewCopyIdCandidates(")
        self.assertIn('copyCase.Add "copyKeeperRow", rowNo', detector)
        self.assertIn('copyCase.Add "copyCandidateRows", candidates', detector)

    def test_global_snapshot_rejects_unrelated_sheet_changes_and_audit_commits_last(self):
        snapshot = procedure(MODULE, "Private Function ReviewCopySheetSignature(")
        self.assertIn("ReviewSignature(ws, rowNo)", snapshot)
        self.assertIn("ws.Cells(1, col).Value2", snapshot)
        current = procedure(MODULE, "Private Function ReviewCopyCaseCurrent(")
        self.assertIn("SyncReviewCopyIdCandidates(CLng(expected(\"year\")))", current)
        self.assertIn('CStr(item("copyDatasetSignature")) = CStr(expected("copyDatasetSignature"))', current)
        audit = procedure(MODULE, "Private Sub ReviewWriteCopyClearAudit(", "End Sub")
        self.assertGreater(audit.index("ws.Cells(r, 5).Value2 = decision"), audit.index("ws.Cells(r, 12).Value2 = backup"))

    def test_failure_injection_oracle_never_reports_success_on_cancel_or_failure(self):
        case = {"category": "D", "keeper": 10, "candidates": [11]}
        original = {10: {"id": "I", "number": "01", "barcode": "B", "status": "登録予定", "folder": "keep"},
                    11: {"id": "I", "number": "01", "barcode": "B", "status": "登録予定", "folder": "copy", "format": "@"}}
        success, changed = simulate_clear(case, 11, copy.deepcopy(original))
        self.assertTrue(success)
        self.assertEqual(changed[10], original[10])
        self.assertEqual(changed[11], {**original[11], "id": "", "number": "", "barcode": ""})
        for kwargs in ({"system_unregistered": False}, {"final_yes": False}, {"revalidation_ok": False},
                       {"backup_ok": False}, {"audit_ok": False}, {"write_failure_at": 2}):
            sheet = copy.deepcopy(original)
            success, after = simulate_clear(case, 11, sheet, **kwargs)
            self.assertFalse(success, kwargs)
            self.assertEqual(after, original, kwargs)
        self.assertFalse(simulate_clear(case, 10, copy.deepcopy(original))[0])
        self.assertFalse(simulate_clear(case, 12, copy.deepcopy(original))[0])
        self.assertFalse(simulate_clear({**case, "category": "A"}, 11, copy.deepcopy(original))[0])
        self.assertFalse(simulate_clear(case, 11, copy.deepcopy(original), write_failure_at=2, restore_ok=False)[0])


if __name__ == "__main__":
    unittest.main()
