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


def detect(csv_rows, excel_rows, target_year):
    """Independent, deliberately strict oracle: each row is a dict of fields."""
    fields = ("id", "number", "barcode")
    scoped = [r for r in csv_rows if r.get("year") == target_year]
    output = []
    for source in scoped:
        values = tuple(source.get(k, "") for k in fields)
        if not all(values) or any(sum(r.get(k, "") == source[k] for r in csv_rows) != 1 for k in fields):
            continue
        matches = [r for r in excel_rows if tuple(r.get(k, "") for k in fields) == values
                   and all(r.get(k) == source.get(k) for k in ("created", "g1", "g2", "folder"))]
        if len(matches) != 1:
            continue
        keeper = matches[0]
        colliders = [r for r in excel_rows if r is not keeper and
                     any(r.get(k) and (r[k].lstrip("0") if k == "number" else r[k]) ==
                         (source[k].lstrip("0") if k == "number" else source[k]) for k in fields)]
        if any(r.get("status") != "登録予定" for r in colliders):
            continue
        if not colliders or any(not all(r.get(k) for k in ("created", "g1", "g2", "folder"))
                                for r in colliders):
            continue
        if any(any(c is not source and any(c.get(k) == r.get(k) for k in fields if r.get(k))
                   for c in csv_rows) for r in colliders):
            continue
        if keeper.get("status") not in {"登録予定", "同期済"}:
            continue
        if any((r.get("created"), r.get("g1"), r.get("g2"), r.get("folder")) ==
               (source.get("created"), source.get("g1"), source.get("g2"), source.get("folder"))
               for r in colliders):
            continue
        output.extend(r["row"] for r in colliders)
    return output


class CopyIdCandidateTests(unittest.TestCase):
    def test_detector_is_isolated_and_never_changes_existing_categories_or_step2(self):
        self.assertIn("Public Function SyncReviewCopyIdCandidates(", MODULE)
        detector = procedure(MODULE, "Public Function SyncReviewCopyIdCandidates(")
        self.assertIn('NewCase("D"', detector)
        self.assertIn('ReviewField(dst, rowNo, "同期状態") <> "登録予定"', detector)
        self.assertIn('Not ReviewIsLinkedSyncStatus(ReviewField(dst, rowNo, "同期状態"))', detector)
        self.assertNotIn("SyncReviewCandidates(", detector)
        self.assertNotIn("ReviewBackup", detector)
        self.assertNotIn(".Rows(", detector)
        self.assertNotIn(".Value2 =", detector)

    def test_three_user_scenarios_are_data_driven_not_row_hardcoded(self):
        scenarios = [
            (1462, 967, [968]),
            (1677, 1362, [1374]),
            (1962, 1923, [1951, 1961]),
        ]
        for src_row, keeper_row, candidate_rows in scenarios:
            src = dict(row=src_row, year=2026, id=f"I{src_row}", number=f"N{src_row}", barcode=f"B{src_row}", created=2020, g1="G1", g2="G2", folder=f"F{src_row}")
            keeper = dict(src, row=keeper_row, status="登録予定")
            copied = [dict(src, row=r, status="登録予定", folder=f"copy{r}") for r in candidate_rows]
            got = detect([src], [keeper, *copied], 2026)
            self.assertEqual(got, candidate_rows)
            self.assertEqual(detect([src], [dict(keeper, status="同期済"), *copied], 2026), candidate_rows)
            self.assertEqual(detect([src], [dict(keeper, status="削除予定"), *copied], 2026), [])
        self.assertNotRegex(MODULE, r"\b(?:1462|967|968|1677|1362|1374|1962|1923|1951|1961)\b")

    def test_ambiguity_and_missing_data_fail_closed(self):
        src = dict(row=1, year=2026, id="I", number="0012", barcode="B", created=2020, g1="G1", g2="G2", folder="F")
        keeper = dict(src, row=2, status="登録予定")
        copy = dict(src, row=3, status="登録予定", folder="Other")
        self.assertEqual(detect([src, dict(src, row=4)], [keeper, copy], 2026), [])
        self.assertEqual(detect([src], [keeper, copy, dict(copy, row=5, folder="Third")], 2026), [3, 5])
        self.assertEqual(detect([src], [keeper, dict(copy, number="12")], 2026), [3])
        self.assertEqual(detect([src], [keeper, dict(keeper, row=7), copy], 2026), [])
        self.assertEqual(detect([src], [keeper, dict(copy, created=2021, folder="F")], 2026), [3])
        self.assertEqual(detect([src], [keeper, dict(copy, g2="OtherProject", folder="F")], 2026), [3])
        common_name = dict(src, folder="起工設計書")
        common_keeper = dict(common_name, row=12, status="登録予定")
        self.assertEqual(detect([common_name], [common_keeper,
                                                dict(common_name, row=13, status="登録予定", created=2021)], 2026), [13])
        self.assertEqual(detect([common_name], [common_keeper,
                                                dict(common_name, row=14, status="登録予定", g2="別工事")], 2026), [14])
        self.assertEqual(detect([common_name], [common_keeper,
                                                dict(common_name, row=15, status="登録予定")], 2026), [])
        self.assertEqual(detect([src], [keeper, dict(copy, folder="F")], 2026), [])
        self.assertEqual(detect([src], [keeper, dict(copy, year=2025, folder="F")], 2026), [])
        self.assertEqual(detect([src], [dict(keeper, folder="Wrong"), copy], 2026), [])
        self.assertEqual(detect([src], [keeper, dict(copy, created=2021)], 2026), [3])
        self.assertEqual(detect([src], [keeper, dict(copy, g1="OtherProject")], 2026), [3])
        self.assertEqual(detect([src], [keeper, dict(copy, folder="F")], 2026), [])
        self.assertEqual(detect([dict(src, id="")], [keeper, copy], 2026), [])
        self.assertEqual(detect([src, dict(src, row=8, id="Other", number="77", barcode="OtherB")],
                                [keeper, copy], 2026), [3])
        self.assertEqual(detect([src, dict(src, row=8, id="Other", number="77", barcode="B")],
                                [keeper, copy], 2026), [])
        self.assertEqual(detect([src], [keeper, dict(copy, number="0012", folder="Other")], 2025), [])

    def test_vba_uses_complete_folder_identity_not_folder_name_alone(self):
        detector = procedure(MODULE, "Public Function SyncReviewCopyIdCandidates(")
        self.assertIn("If ReviewName(dst, dRow, False, True) = sourceName Then valid = False", detector)
        self.assertNotIn('ReviewTextKey(ReviewValue(dst, dRow, ReviewCol(dst, "個別フォルダー"', detector)
        self.assertIn("ReviewName(dst, rowNo, False, False) <> ReviewName(src, sRow, True, False)", detector)
        self.assertIn("ReviewFiscal(dst, dRow) = yearValue", detector)

    def test_matching_and_global_collision_checks_are_explicit(self):
        detector = procedure(MODULE, "Public Function SyncReviewCopyIdCandidates(")
        for term in ("ReviewFiscal", "ReviewName", "作成年度", "第1ガイド", "第2ガイド", "個別フォルダー", "登録予定"):
            self.assertIn(term, detector)
        self.assertIn("ReviewPadding", detector)
        self.assertIn("csv", detector.lower())


if __name__ == "__main__":
    unittest.main()
