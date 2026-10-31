"""Reference scenarios for copy-shaped snapshots, not an Excel/VBA execution test."""
from copy import deepcopy
import unittest


class Snapshot:
    def __init__(self, rows):
        self.rows = deepcopy(rows)
        self.known = {row_id: set(values) for row_id, values in rows.items()}

    def get(self, row_id, field):
        if row_id not in self.rows or field not in self.known[row_id]:
            return False, None
        return True, self.rows[row_id][field]

    def confirm(self, entries, fail_after=None):
        old = deepcopy((self.rows, self.known))
        try:
            for i, (row_id, field, live, expected) in enumerate(entries, 1):
                if live != expected:
                    raise ValueError("changed after comparison")
                self.rows.setdefault(row_id, {})[field] = live
                self.known.setdefault(row_id, set()).add(field)
                if i == fail_after:
                    raise RuntimeError("injected failure")
        except Exception:
            self.rows, self.known = old
            raise


def decide(snapshot, row_id, field, excel, csv, saved=None):
    known, base = snapshot.get(row_id, field)
    if saved and excel != csv:
        return "review", excel
    if excel == csv:
        snapshot.confirm([(row_id, field, excel, csv)])
        return "auto_resolved" if saved else "confirmed", excel
    if not known:
        return "review", excel
    if excel == base:
        snapshot.confirm([(row_id, field, csv, csv)])
        return "csv_apply", csv
    if csv == base:
        return "excel_pending", excel
    return "review", excel


class SnapshotScenarios(unittest.TestCase):
    def test_partial_new_row_does_not_turn_blank_unknown_into_known(self):
        s = Snapshot({"old": {"period": "5", "title": "old"}})
        s.confirm([("new", "title", "same", "same")])
        self.assertEqual(s.get("new", "period"), (False, None))
        self.assertEqual(decide(s, "new", "period", "", "3"), ("review", ""))
        self.assertEqual(s.get("old", "period"), (True, "5"))

    def test_rows_are_tracked_by_id_even_when_order_changes(self):
        s = Snapshot({"id-1": {"period": "5"}, "id-2": {"period": "9"}})
        s.rows = {key: s.rows[key] for key in reversed(list(s.rows))}
        self.assertEqual(decide(s, "id-1", "period", "5", "3"), ("csv_apply", "3"))
        self.assertEqual(s.get("id-2", "period"), (True, "9"))

    def test_hold_revert_and_excel_adoption_never_auto_overwrite(self):
        for saved, excel in [("hold", "5"), ("hold", "7"), ("excel", "7")]:
            with self.subTest(saved=saved, excel=excel):
                s = Snapshot({"id": {"period": "5"}})
                self.assertEqual(decide(s, "id", "period", excel, "3", saved), ("review", excel))
                self.assertEqual(s.get("id", "period"), (True, "5"))

    def test_resolved_hold_updates_only_confirmed_field(self):
        s = Snapshot({"id": {"period": "5", "title": "do-not-change"}})
        self.assertEqual(decide(s, "id", "period", "3", "3", "hold"), ("auto_resolved", "3"))
        self.assertEqual(s.rows["id"], {"period": "3", "title": "do-not-change"})

    def test_partial_failure_restores_both_values_and_knownness(self):
        s = Snapshot({"old": {"period": "5"}})
        before = deepcopy((s.rows, s.known))
        with self.assertRaises(RuntimeError):
            s.confirm([("new", "title", "A", "A"), ("old", "period", "3", "3")], fail_after=1)
        self.assertEqual((s.rows, s.known), before)

    def test_stale_live_value_prevents_commit(self):
        s = Snapshot({"id": {"period": "5"}})
        with self.assertRaises(ValueError):
            s.confirm([("id", "period", "7", "3")])
        self.assertEqual(s.get("id", "period"), (True, "5"))


if __name__ == "__main__":
    unittest.main()
