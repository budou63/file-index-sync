"""Safety contract: stale hold decisions must prevent automatic replacement."""

from pathlib import Path
import re
import unittest


SOURCE = (Path(__file__).resolve().parents[1] / "vba" / "Module1.txt").read_text(encoding="utf-8-sig")


def proc(name):
    match = re.search(
        rf"(?ims)^\s*(?:Private|Public) (?:Sub|Function)\s+{re.escape(name)}\b.*?^End (?:Sub|Function)",
        SOURCE,
    )
    if not match:
        raise AssertionError(f"VBA procedure missing: {name}")
    return match.group(0)


class HeldConflictSafetyTests(unittest.TestCase):
    def test_raw_hold_history_checked_before_auto_confirmation_or_overwrite(self):
        body = proc("BidirectionalSyncApplyManagedGroup")
        # A stored hold must be checked regardless of whether the saved
        # baseline / CSV / Excel snapshot still matches live values.
        match = re.search(r"BidirectionalSync(?:Protect|Guard|Check)HeldConflict", body, re.I)
        self.assertIsNotNone(match, "No independent saved-hold guard in group comparison")
        self.assertLess(match.start(), body.index("If excelEqualsCsv Then"))
        self.assertLess(match.start(), body.index("WriteStep2SourceValue"))

    def test_saved_conflict_lookup_must_not_rely_only_on_live_matched_review_cases(self):
        body = proc("BidirectionalSyncLoadSavedConflictKeys")
        # ReviewCases filters out snapshots whenever the live Excel/CSV/base
        # changes; it cannot alone establish that there is no held history.
        self.assertNotIn("Set cases = BidirectionalSyncConflictReviewCases(targetYearValue)", body)


if __name__ == "__main__":
    unittest.main()
