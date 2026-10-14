"""D-only year-alias comparison: Python decision model plus VBA source-path checks.

The Python cases do not execute VBA; Excel/VBE runtime verification remains required.
"""
import re
import unittest
from pathlib import Path

MODULE = (Path(__file__).resolve().parents[1] / "vba/modSyncReview.txt").read_text(encoding="utf-8")


def year(raw):
    text = raw.strip().upper().replace("令和", "R").replace("年度", "").replace("年", "")
    if text.startswith("R") and text[1:].isascii() and text[1:].isdigit():
        n = int(text[1:])
        return n + 2018 if 1 <= n <= 100 else 0
    if text.isascii() and text.isdigit():
        n = int(text)
        return n if 1900 <= n <= 2200 else 0
    return 0


def year_alias_equal(a, b):
    left, right = year(a), year(b)
    return left != 0 and right != 0 and left == right


class ReviewYearAliasTests(unittest.TestCase):
    def test_year_value_cases(self):
        for left, right, expected in (
            ("R8", "2026", True), ("R8", "R8", True),
            ("R8", "2025", False), ("R7", "2026", False),
            ("R8", "unknown", False), ("R8", "", False),
            ("R7.5", "2026", False), ("2025.5", "2026", False),
            ("", "2026", False), ("", "", False),
        ):
            with self.subTest(left=left, right=right):
                self.assertEqual(year_alias_equal(left, right), expected)

    def test_created_year_aliases(self):
        self.assertTrue(year_alias_equal("令和8年度", "2026年"))
        self.assertFalse(year_alias_equal("令和7年度", "2026"))
        self.assertFalse(year_alias_equal("", "2026"))

    def test_only_year_aliases_use_year_comparison_in_actual_reviewcol(self):
        start = MODULE.index("Private Function ReviewCol(")
        col = MODULE[start:MODULE.index("End Function", start)]
        self.assertIn("yearAliasGroup", col)
        self.assertIn("ReviewYearAliasGroup(CStr(a))", col)
        self.assertIn("ReviewStrictAliasYear(selectedText)", col)
        self.assertIn("ReviewStrictAliasYear(otherText)", col)
        self.assertIn("yearAliasGroup <> 0", col)
        self.assertIn("If UBound(aliases) = 2 Then yearAliasGroup = ReviewYearAliasGroup(CStr(aliases(0)))", col)
        self.assertIn("StrComp(selectedText, otherText, vbBinaryCompare) <> 0", col)
        self.assertIn('"別名見出しの内容が一致しません:', col)
        self.assertIn("If IsError(ws.Cells(rowNo, chosen).Value2) Or IsError(ws.Cells(rowNo, found).Value2) Then", col)
        self.assertIn("reviewHeaderCache.Add cacheKey, chosen", col)
        group_start = MODULE.index("Private Function ReviewYearAliasGroup(")
        group = MODULE[group_start:MODULE.index("End Function", group_start)]
        for alias in ('"年度（和暦）", "年度(和暦)", "年度"',
                      '"作成年度（和暦）", "作成年度(和暦)", "作成年度"'):
            self.assertIn(alias, group)
        self.assertIn("ReviewYearAliasGroup =", group)
        self.assertNotIn('"ファイルID"', group)

        strict_start = MODULE.index("Private Function ReviewStrictAliasYear(")
        strict = MODULE[strict_start:MODULE.index("End Function", strict_start)]
        self.assertIn('Replace(text, "令和", "R")', strict)
        self.assertIn('Replace(text, "年度", "")', strict)
        self.assertIn('Replace(text, "年", "")', strict)
        self.assertIn('If AscW(c) < 48 Or AscW(c) > 57 Then Exit Function', strict)
        self.assertIn("ReviewStrictAliasYear = ReviewYear(raw)", strict)

    def test_d_preflight_actually_requests_cross_notation_aliases(self):
        start = MODULE.index("Public Function SyncReviewCopyIdCandidates(")
        detector = MODULE[start:MODULE.index("End Function", start)]
        self.assertIn('ReviewCol(src, "年度（和暦）", "年度(和暦)", "年度")', detector)
        self.assertIn('ReviewCol(src, "作成年度（和暦）", "作成年度(和暦)", "作成年度")', detector)
        self.assertIn('ReviewCol(dst, "年度（和暦）", "年度(和暦)", "年度")', detector)
        self.assertIn('ReviewCol(dst, "作成年度（和暦）", "作成年度(和暦)", "作成年度")', detector)


if __name__ == "__main__":
    unittest.main()
