"""Contract tests for per-Excel-row B-candidate "別物" judgments.

These source-level and small model tests complement (but do not replace) Excel/VBE runtime tests.
"""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8")
MANUAL = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8")


def procedure(source, start, end="End Function"):
    match = re.search(r"(?ms)^" + re.escape(start) + r".*?^" + re.escape(end), source)
    if match is None:
        raise AssertionError(f"missing procedure {start}")
    return match.group()


def next_unconfirmed(rows, decided_rows, after_row):
    """Model the UI's cyclic scan, excluding the currently selected row."""
    if not rows:
        return None
    start = rows.index(after_row) if after_row in rows else -1
    for offset in range(1, len(rows) + 1):
        candidate = rows[(start + offset) % len(rows)]
        if candidate != after_row and candidate not in decided_rows:
            return candidate
    return None


def candidate_audit_matches(audit, year, csv_row, csv_signature, excel_row, excel_signature):
    """Model the complete identity tuple required by the row-audit reader."""
    return (
        audit.get("scope") == "候補行"
        and audit.get("year") == year
        and audit.get("category") == "B"
        and audit.get("csv_row") == csv_row
        and audit.get("excel_row") == excel_row
        and audit.get("csv_signature") == csv_signature
        and audit.get("excel_signature") == excel_signature
    )


class SyncReviewRowDecisionTests(unittest.TestCase):
    def test_one_b_candidate_can_be_judged_without_advancing_to_another_row(self):
        decide = procedure(MODULE, "Public Function SyncReviewDecide(")
        self.assertIn('CStr(item("category")) = "B" And decision = "別物"', decide)
        self.assertIn("ReviewItemContainsRow(item, selectedRow)", decide)
        self.assertIsNone(next_unconfirmed([805], {805}, 805))

    def test_four_candidates_keep_the_other_three_unconfirmed_after_one_judgment(self):
        rows = [805, 806, 807, 808]
        decided = {806}
        self.assertEqual([row for row in rows if row not in decided], [805, 807, 808])
        self.assertEqual(next_unconfirmed(rows, decided, 806), 807)
        update = procedure(FORM, "Private Sub UpdateBRowJudgments(", "End Sub")
        self.assertIn('totalCount & "行／別物確認済" & _', update)
        self.assertIn('confirmedCount & "行／未確認" & (totalCount - confirmedCount)', update)

    def test_record_choice_moves_within_the_same_case_before_next_csv_case(self):
        record = procedure(FORM, "Private Sub RecordChoice(", "End Sub")
        self.assertIn("SyncReviewDecide(item, decision, rowNo)", record)
        self.assertIn("nextRow = BNextUnconfirmedRow(item, rowNo)", record)
        b_branch = record[record.index('If CStr(item("category")) = "B" And decision = "別物" Then'):]
        self.assertLess(b_branch.index("If nextRow > 0 Then"), b_branch.index("mCmdNext_Click"))
        advance = procedure(FORM, "Private Sub mCmdNext_Click()", "End Sub")
        self.assertIn("BNextUnconfirmedRow(item, SelectedRow())", advance)

    def test_audit_appends_pair_identity_and_timestamp_without_replacing_old_rows(self):
        audit = procedure(MODULE, "Private Function ReviewAudit(", "End Function")
        writer = procedure(MODULE, "Private Sub ReviewWriteAudit(", "End Sub")
        self.assertIn('ReviewAuditColumn(ws, "判断単位", True)', audit)
        self.assertIn('ReviewAuditColumn(ws, "選択Excel行署名", True)', audit)
        self.assertIn('Set ws = ReviewAudit(True, selectedRow > 0)', writer)
        self.assertIn("ws.Cells(r, 1).Value2 = Now", writer)
        self.assertIn('ws.Cells(r, 2).Value2 = CLng(item("year"))', writer)
        self.assertIn('ws.Cells(r, 6).Value2 = CLng(item("sourceRow"))', writer)
        self.assertIn('ws.Cells(r, 7).Value2 = CStr(selectedRow)', writer)
        self.assertIn('ws.Cells(r, scopeCol).Value2 = "候補行"', writer)
        self.assertIn("ReviewSignature(dst, selectedRow)", writer)
        self.assertIn('ws.Cells(r, selectedSignatureCol).Value2 = selectedHash', writer)
        self.assertNotIn(".ClearContents", audit + writer)

    def test_rejudge_reuses_only_unchanged_source_and_selected_excel_signatures(self):
        row = {
            "scope": "候補行", "year": 2026, "category": "B", "csv_row": 1399,
            "excel_row": 806, "csv_signature": "csv-a", "excel_signature": "excel-a",
        }
        self.assertTrue(candidate_audit_matches(row, 2026, 1399, "csv-a", 806, "excel-a"))
        self.assertFalse(candidate_audit_matches(row, 2026, 1399, "csv-a", 806, "excel-b"))
        self.assertFalse(candidate_audit_matches(row, 2026, 1399, "csv-b", 806, "excel-a"))
        reader = procedure(MODULE, "Public Function SyncReviewDecision(")
        for identity in ('"候補行"', 'CStr(item("year"))', 'CStr(item("sourceRow"))',
                         'CStr(selectedRow)', 'ReviewAuditHashMatches(ws.Cells(r, 9).Value2, sourceHash)',
                         'selectedSignatureCol'):
            self.assertIn(identity, reader)

    def test_legacy_case_level_distinct_is_kept_separate_not_applied_to_candidate_rows(self):
        reader = procedure(MODULE, "Public Function SyncReviewDecision(")
        legacy = procedure(MODULE, "Public Function SyncReviewLegacyCaseDecision(")
        self.assertIn('If category = "B" And storedDecision = "別物" Then Exit For', reader)
        self.assertIn('ReviewAuditScopeAt(ws, r, scopeCol) = "候補行"', reader)
        self.assertIn('ReviewAuditScopeAt(ws, r, scopeCol) = "案件単位"', legacy)
        self.assertIn('SyncReviewLegacyCaseDecision = "別物"', legacy)
        self.assertIn("旧方式の案件単位判断", FORM)
        self.assertIn("旧仕様で案件全体", MANUAL)

    def test_selected_row_status_is_visible_and_unrelated_safety_apis_remain_scoped(self):
        show = procedure(FORM, "Private Sub ShowCurrent()", "End Sub")
        compare = procedure(FORM, "Private Sub RenderComparison(", "End Sub")
        self.assertIn('mLstRows.AddItem CStr(v) & " 行 [" & rowDecision & "]"', show)
        self.assertIn("選択Excel行の判断", compare)
        approve = procedure(MODULE, "Public Function SyncReviewApprove(")
        delete = procedure(MODULE, "Public Function SyncReviewDeleteBatch(")
        clear = procedure(MODULE, "Public Function SyncReviewClearCopyIdentifiers(")
        self.assertIn('If CStr(item("category")) <> "B"', approve)
        self.assertIn('If CStr(item("category")) <> "C"', delete)
        self.assertIn('If CStr(item("category")) <> "D"', clear)


if __name__ == "__main__":
    unittest.main()
