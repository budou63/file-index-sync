"""Static coverage for the read-only CSV/register comparison display."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8")
MODULE = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")


def procedure(source: str, name: str, kind: str = "Sub") -> str:
    match = re.search(
        rf"(?ms)^Private {kind} {re.escape(name)}\b.*?^End {kind}", source
    )
    if match is None:
        raise AssertionError(f"missing {kind}: {name}")
    return match.group()


class SyncReviewComparisonTableTests(unittest.TestCase):
    def test_comparison_table_has_three_columns_and_explicit_text_statuses(self):
        build = procedure(FORM, "BuildControls")
        render = procedure(FORM, "RenderComparison")
        for caption in ("項目", "目録CSV", "新ファイル基準表"):
            self.assertIn(caption, FORM)
        self.assertIn("mFraCompare", build)
        self.assertIn("mCompareCsvCells", build)
        self.assertIn("mCompareExcelCells", build)
        compare_status = procedure(FORM, "CompareDisplayStatus", "Function")
        for status in ("一致", "相違", "比較対象外"):
            self.assertIn(status, compare_status)
        self.assertIn("RGB(255, 235, 235)", render)
        self.assertIn("statusText", render)

    def test_comparison_reads_selected_rows_without_changing_candidate_logic(self):
        render = procedure(FORM, "RenderComparison")
        self.assertIn("ThisWorkbook.Worksheets(CSV_SHEET)", render)
        self.assertIn("ThisWorkbook.Worksheets(DEST_SHEET)", render)
        self.assertIn("sourceRow", render)
        self.assertIn("excelRow", render)
        for label in ("ファイルID", "ファイル番号", "バーコード番号", "年度", "作成年度", "第1ガイド", "第2ガイド", "個別フォルダー", "同期状態"):
            self.assertIn(label, FORM)
        field_spec = procedure(FORM, "CompareFieldSpec")
        created_year = field_spec[field_spec.index("Case 5"):field_spec.index("Case 6")]
        self.assertIn("作成年度（和暦）", created_year)
        self.assertNotIn('Array("作成年度")', created_year)
        self.assertNotRegex(render, r"(?:\.Value2|\.Value|\.ClearContents)\s*=")
        self.assertIn('item("reason")', render)
        self.assertIn("SyncReviewDecision(item)", render)

    def test_show_current_uses_the_saved_decision_for_delete_checkbox_state(self):
        show = procedure(FORM, "ShowCurrent")
        self.assertRegex(show, r"(?m)^\s*Dim[^\n]*\bdecision\s+As\s+String\b")
        self.assertRegex(show, r'(?m)^\s*If category <> "D" Then decision = SyncReviewDecision\(item\)\s*$')
        self.assertIn("Len(decision) = 0 And deletionSelection.Exists", show)
        self.assertIn('mChkDelete.Visible = (category = "C")', show)

    def test_selecting_an_excel_candidate_refreshes_the_comparison(self):
        click = procedure(FORM, "mLstRows_Click")
        select = procedure(FORM, "ShowCurrentSelection")
        self.assertIn("RenderComparison item, rowNo", select)
        self.assertIn("ShowCurrentSelection", click)
        self.assertIn("For Each v In Split(rows, \";\")", procedure(FORM, "ShowCurrent"))
        self.assertIn('Case "A": idx = 1', procedure(FORM, "ShowCategory"))
        self.assertIn('Case "B": idx = 2', procedure(FORM, "ShowCategory"))
        self.assertIn('Case "C": idx = 3', procedure(FORM, "ShowCategory"))
        self.assertIn('Case "D": idx = 4', procedure(FORM, "ShowCategory"))

    def test_long_values_wrap_inside_a_vertically_scrollable_grid(self):
        build = procedure(FORM, "BuildControls")
        render = procedure(FORM, "RenderComparison")
        self.assertIn("mFraCompare.ScrollBars = fmScrollBarsVertical", build)
        self.assertIn("mFraCompare.ScrollHeight", render)
        self.assertIn(".WordWrap = True", build)
        self.assertIn(".MultiLine = True", build)
        self.assertIn(".ScrollBars = fmScrollBarsVertical", build)

    def test_reason_and_previous_judgment_are_read_from_the_selected_case(self):
        render = procedure(FORM, "RenderComparison")
        self.assertIn('item("reason")', render)
        self.assertIn("SyncReviewDecision(item)", render)
        self.assertIn("SyncReviewDecisionComment(item)", render)
        self.assertNotIn("同一年度・同じガイド内に類似した個別フォルダー", FORM)

    def test_previous_decision_comment_is_read_from_the_matching_audit_entry(self):
        match = re.search(
            r"(?ms)^Public Function SyncReviewDecisionComment\b.*?^End Function",
            MODULE,
        )
        self.assertIsNotNone(match)
        reader = match.group()
        self.assertIn("ReviewAudit(False)", reader)
        self.assertIn("ReviewAuditHashMatches", reader)
        self.assertIn("ws.Cells(r, 11).Value2", reader)
        self.assertNotRegex(reader, r"(?:\.Value2|\.Value|\.ClearContents)\s*=")

    def test_code_values_are_rendered_as_text_without_numeric_coercion(self):
        render = procedure(FORM, "RenderComparison")
        value_reader = procedure(FORM, "CompareCellText", "Function")
        self.assertIn("CStr(rawValue)", value_reader)
        self.assertIn(".Text)", value_reader)
        self.assertNotRegex(value_reader, r"\b(?:CLng|Val|Format\$)\s*\(")
        self.assertIn("CompareCellText", render)

    def test_empty_selection_clears_stale_values_and_comparison_is_read_only(self):
        render = procedure(FORM, "RenderComparison")
        empty_state = render[:render.index("Set src =")]
        self.assertIn('mCompareCsvCells(fieldIndex).Text = "―"', empty_state)
        self.assertIn('mCompareExcelCells(fieldIndex).Text = "―"', empty_state)
        self.assertNotRegex(render, r"(?:\.Value2|\.Value|\.ClearContents)\s*=")

    def test_restore_artifacts_remain_body_synced_and_designer_is_not_rebuilt(self):
        self.assertIn("Public Sub LoadCases(", FORM)
        self.assertIn("Private Sub mCmdRefresh_Click()", FORM)
        self.assertIn("Private Sub mCmdDelete_Click()", FORM)
        self.assertIn("Private Sub mCmdClear_Click()", FORM)


if __name__ == "__main__":
    unittest.main()
