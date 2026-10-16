"""Static and decision-model contracts for nearby Step2 inserts and review actions."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
STEP2 = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8")
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8")
REVIEW = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")


def proc(source: str, name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        source,
        re.S,
    )
    if not match:
        raise AssertionError(f"missing procedure: {name}")
    return match.group(0)


def insert_after(rows, new):
    """Decision model: choose the final row of the strongest matching prior group."""
    same_year = [i for i, row in enumerate(rows) if row["year"] == new["year"] and not row.get("new")]
    if not same_year:
        return len(rows)
    for keys in (("class1", "guide1", "guide2"), ("class1", "guide1"), ("class1",)):
        matches = [i for i in same_year if all(rows[i][key] == new[key] for key in keys)]
        if matches:
            return matches[-1] + 1
    return same_year[-1] + 1


class NearbyInsertAndReviewTests(unittest.TestCase):
    def test_nearby_insert_decision_contract(self):
        rows = [
            {"year": 2025, "class1": "A", "guide1": "x", "guide2": "x1", "id": "old"},
            {"year": 2026, "class1": "A", "guide1": "g", "guide2": "h", "id": "one"},
            {"year": 2026, "class1": "B", "guide1": "z", "guide2": "q", "id": "two"},
            {"year": 2026, "class1": "A", "guide1": "g", "guide2": "h", "id": "three"},
            {"year": 2027, "class1": "A", "guide1": "g", "guide2": "h", "id": "future"},
        ]
        self.assertEqual(insert_after(rows, {"year": 2026, "class1": "A", "guide1": "g", "guide2": "h"}), 4)
        self.assertEqual(insert_after(rows, {"year": 2026, "class1": "A", "guide1": "missing", "guide2": "missing"}), 4)
        self.assertEqual(insert_after(rows, {"year": 2026, "class1": "missing", "guide1": "missing", "guide2": "missing"}), 4)
        self.assertEqual(insert_after(rows, {"year": 2028, "class1": "A", "guide1": "g", "guide2": "h"}), 5)
        self.assertEqual([row["id"] for row in rows], ["old", "one", "two", "three", "future"])

    def test_step2_stages_new_rows_then_moves_only_those_rows_without_global_sort(self):
        step2 = proc(STEP2, "手順2_新ファイル基準表を作成する")
        placement = proc(STEP2, "PlaceStep2NewRowsNearGroups")
        self.assertIn("Set appendedRowKeys = New Collection", step2)
        self.assertIn("appendedRowKeys.Add key", step2)
        self.assertIn("PlaceStep2NewRowsNearGroups wsDst, appendedRowKeys, targetYearValue", step2)
        self.assertNotIn("SortNewFileStandardByFiscalYear wsDst", step2)
        self.assertIn("FindStep2NearbyInsertRow", placement)
        self.assertIn("wsDst.Rows(currentRow).Cut", placement)
        self.assertIn("wsDst.Rows(insertRow).Insert Shift:=xlDown", placement)
        self.assertIn("FindInsertRowByYearOrder", proc(STEP2, "FindStep2NearbyInsertRow"))
        self.assertIn("same year", proc(STEP2, "FindStep2NearbyInsertRow").lower())

    def test_group_priority_and_deterministic_last_matching_group_are_explicit(self):
        helper = proc(STEP2, "FindStep2NearbyInsertRow")
        for header in ("分類名１", "第1ガイド", "第2ガイド"):
            self.assertIn(header, helper)
        self.assertIn("matchLevel", helper)
        self.assertIn("rowNo", helper)
        self.assertIn("candidateRow", helper)
        self.assertIn("= rowNo", helper)

    def test_rejudge_button_reuses_live_detectors_and_does_not_mutate_data(self):
        build = proc(FORM, "BuildControls")
        refresh = proc(FORM, "mCmdRefresh_Click")
        self.assertIn('mCmdRefresh.Caption = "再判定"', build)
        self.assertIn("SyncReviewCandidates(targetYear)", refresh)
        self.assertIn("SyncReviewCopyIdCandidates(targetYear)", refresh)
        self.assertNotRegex(refresh, r"\.(?:Rows|Cells|Worksheets)\b")
        self.assertIn("Set previousCases = cases", refresh)
        self.assertIn("Set cases = previousCases", refresh)

    def test_delete_selection_is_explicit_excel_row_selection_and_batch_stays_backup_first(self):
        show = proc(FORM, "ShowCurrent")
        click = proc(FORM, "mLstRows_Click")
        delete_click = proc(FORM, "mCmdDelete_Click")
        batch = proc(REVIEW, "SyncReviewDeleteBatch")
        self.assertIn("mLstRows.MultiSelect = fmMultiSelectMulti", show)
        self.assertIn("SelectDeleteRow", click)
        self.assertIn("SyncReviewSelectedDeletionItem", delete_click)
        self.assertIn("削除対象のExcel行を選択してください", delete_click)
        self.assertIn("ReviewBackup()", batch)
        self.assertLess(batch.index("ReviewBackup()"), batch.index("dst.Rows(rowNo).Delete"))
        self.assertIn("最終確認：一括削除", batch)


if __name__ == "__main__":
    unittest.main()
