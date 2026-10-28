"""Static contracts and reference model for conflict decision history."""
from __future__ import annotations

import unittest
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")

UNRESOLVED = "競合未解決"
EXCEL = "Excel採用"
SYSTEM = "システム採用"
HOLD = "保留"


def procedure(text: str, name: str, kind: str = "Sub") -> str:
    marker = f"Private {kind} {name}("
    if marker not in text:
        marker = f"{kind} {name}("
    start = text.index(marker)
    end = text.index(f"End {kind}", start)
    return text[start:end]


@dataclass(frozen=True)
class Conflict:
    year: int
    row_id: str
    stable_key: str
    group: str
    excel_value: str
    csv_value: str
    identity_unique: bool = True

    def history_key(self) -> tuple[str, ...] | None:
        if not self.identity_unique or not all((self.row_id, self.stable_key, self.group)):
            return None
        return (str(self.year), self.row_id, self.stable_key, self.group,
                self.excel_value, self.csv_value)


class DecisionHistory:
    """Model the required persist-across-runs contract, not execution of VBA."""

    def __init__(self) -> None:
        self.saved: dict[tuple[str, ...], str] = {}

    def decide(self, conflict: Conflict, decision: str) -> None:
        key = conflict.history_key()
        if key is None:
            raise ValueError("ambiguous identity cannot receive a reusable judgment")
        self.saved[key] = decision

    def refresh(self, current: list[Conflict]) -> list[tuple[Conflict, str]]:
        active: list[tuple[Conflict, str]] = []
        active_keys: set[tuple[str, ...]] = set()
        for conflict in current:
            key = conflict.history_key()
            decision = self.saved.get(key, UNRESOLVED) if key is not None else UNRESOLVED
            if key is not None:
                if key in active_keys:
                    raise ValueError("duplicate current conflict")
                active_keys.add(key)
                self.saved.setdefault(key, UNRESOLVED)
            active.append((conflict, decision))
        return active


class ConflictDecisionHistoryTests(unittest.TestCase):
    def test_excel_decision_is_carried_through_first_second_and_third_run(self) -> None:
        history = DecisionHistory()
        mismatch = Conflict(2026, "RID-1", "FILE-1", "保存期間", "5年", "10年")
        self.assertEqual(history.refresh([mismatch]), [(mismatch, UNRESOLVED)])
        history.decide(mismatch, EXCEL)
        for _ in (2, 3):
            active = history.refresh([mismatch])
            self.assertEqual(active, [(mismatch, EXCEL)])

    def test_hold_remains_hold_without_turning_into_new_unjudged(self) -> None:
        history = DecisionHistory()
        mismatch = Conflict(2026, "RID-1", "FILE-1", "保存期間", "5年", "10年")
        history.decide(mismatch, HOLD)
        self.assertEqual(history.refresh([mismatch]), [(mismatch, HOLD)])

    def test_changed_comparison_value_is_a_new_judgment_and_old_history_remains(self) -> None:
        history = DecisionHistory()
        prior = Conflict(2026, "RID-1", "FILE-1", "保存期間", "5年", "10年")
        changed = Conflict(2026, "RID-1", "FILE-1", "保存期間", "5年", "3年")
        history.decide(prior, EXCEL)
        self.assertEqual(history.refresh([changed]), [(changed, UNRESOLVED)])
        self.assertEqual(history.refresh([prior]), [(prior, EXCEL)])
        self.assertEqual(len(history.saved), 2)

    def test_resolved_candidate_disappears_from_active_list_without_deleting_history(self) -> None:
        history = DecisionHistory()
        mismatch = Conflict(2026, "RID-1", "FILE-1", "保存期間", "5年", "10年")
        history.decide(mismatch, EXCEL)
        self.assertEqual(history.refresh([]), [])
        self.assertEqual(history.refresh([mismatch]), [(mismatch, EXCEL)])

    def test_different_file_with_same_visible_name_does_not_reuse_decision(self) -> None:
        history = DecisionHistory()
        prior = Conflict(2026, "RID-A", "FILE-A", "保存期間", "5年", "10年")
        other = Conflict(2026, "RID-B", "FILE-B", "保存期間", "5年", "10年")
        history.decide(prior, EXCEL)
        self.assertEqual(history.refresh([other]), [(other, UNRESOLVED)])

    def test_ambiguous_identity_never_uses_an_existing_judgment(self) -> None:
        history = DecisionHistory()
        prior = Conflict(2026, "RID-1", "FILE-1", "保存期間", "5年", "10年")
        ambiguous = Conflict(2026, "RID-1", "FILE-1", "保存期間", "5年", "10年", False)
        history.decide(prior, EXCEL)
        self.assertEqual(history.refresh([ambiguous]), [(ambiguous, UNRESOLVED)])

    def test_vba_history_key_includes_year_file_item_and_full_snapshot(self) -> None:
        key = procedure(MODULE, "BidirectionalSyncConflictHistoryKey", "Function")
        for field in ("yearValue", "rowId", "stableKey", "groupName", "snapshot"):
            self.assertIn(field, key)
        refresh = procedure(MODULE, "BidirectionalSyncRefreshConflictState")
        self.assertIn("BidirectionalSyncConflictHistoryKey", refresh)
        self.assertIn("previousByKey.Exists(historyKey)", refresh)
        self.assertIn("preservedRows.Add flatRow", refresh)
        self.assertIn("decision = CStr(previousByKey(historyKey))", refresh)

    def test_current_review_list_filters_archived_history_against_live_values(self) -> None:
        review = procedure(MODULE, "BidirectionalSyncConflictReviewCases", "Function")
        self.assertIn("BidirectionalSyncConflictHistoryKey", review)
        self.assertIn("BidirectionalSyncConflictLiveMatches(item", review)
        self.assertIn("activeCaseKeys.Exists(caseKey)", review)
        self.assertIn("result.Add item", review)
        lookup = procedure(MODULE, "BidirectionalSyncFindConflictStateRow", "Function")
        self.assertIn("BidirectionalSyncConflictHistoryKey", lookup)

    def test_saved_conflict_guard_uses_only_live_cases_not_archived_rows(self) -> None:
        load = procedure(MODULE, "BidirectionalSyncLoadSavedConflictKeys", "Function")
        self.assertIn("BidirectionalSyncConflictReviewCases", load)
        self.assertNotIn("For i = 1 To UBound(values, 1)", load)

    def test_step4_ignores_stale_archives_but_still_blocks_live_unresolved_cases(self) -> None:
        step4 = procedure(MODULE, "BidirectionalSyncConflictBlocksStep4", "Function")
        self.assertIn("BidirectionalSyncConflictReviewCases", step4)
        self.assertIn("BIDIRECTIONAL_SYNC_CONFLICT_UNRESOLVED", step4)
        self.assertIn("BIDIRECTIONAL_SYNC_CONFLICT_HOLD", step4)
        self.assertIn("BidirectionalSyncConflictCurrentValuesDiffer", step4)
        self.assertIn("activeGroups.Exists(groupKey)", step4)

    def test_live_history_reuse_requires_unique_file_and_row_identity(self) -> None:
        live = procedure(MODULE, "BidirectionalSyncConflictLiveMatches", "Function")
        self.assertIn("BidirectionalSyncDestinationRowIdIsUnique", live)
        self.assertIn("BidirectionalSyncBatchSourceKeyIsUnique", live)

    def test_decision_key_uses_current_side_values_not_baseline(self) -> None:
        decision_key = procedure(MODULE, "BidirectionalSyncConflictDecisionSnapshot", "Function")
        self.assertIn('fieldItem("systemValue")', decision_key)
        self.assertIn('fieldItem("excelValue")', decision_key)
        self.assertNotIn('fieldItem("baselineValue")', decision_key)
        history_key = procedure(MODULE, "BidirectionalSyncConflictHistoryKey", "Function")
        self.assertIn("BidirectionalSyncConflictDecisionSnapshot(snapshot)", history_key)


if __name__ == "__main__":
    unittest.main()
