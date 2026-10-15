"""D-only year-alias comparison: Python decision model plus VBA source-path checks.

The Python cases do not execute VBA; Excel/VBE runtime verification remains required.
"""
import unittest
from pathlib import Path

MODULE = (Path(__file__).resolve().parents[1] / "vba/modSyncReview.txt").read_text(encoding="utf-8")


def year(raw):
    text = raw.strip().upper().replace("令和", "R").replace("平成", "H")
    text = text.replace("年度", "").replace("年", "")
    if len(text) < 2 or not text[1:].isascii() or not text[1:].isdigit():
        return 0
    n = int(text[1:])
    offsets = {"R": 2018, "H": 1988}
    if text[0] in offsets:
        value = n + offsets[text[0]]
        return value if 1900 <= value <= 2200 else 0
    if text.isascii() and text.isdigit():
        value = int(text)
        return value if 1900 <= value <= 2200 else 0
    return 0


def year_alias_compatible(a, b):
    if not a.strip() or not b.strip():
        return True
    left, right = year(a), year(b)
    return left != 0 and right != 0 and left == right


def aliases_compatible(*values):
    nonblank = [value for value in values if value.strip()]
    if not nonblank:
        return True
    return all(year_alias_compatible(nonblank[0], value) for value in nonblank)


class ReviewYearAliasTests(unittest.TestCase):
    def test_year_alias_compatibility_cases(self):
        for left, right, expected in (
            ("R8", "2026", True), ("R8", "R8", True),
            ("H12", "2000", True), ("平成12年度", "2000年", True),
            ("H12", "", True), ("", "2026", True), ("", "", True),
            ("R8", "2025", False), ("H12", "2001", False),
            ("R8", "unknown", False), ("H12", "H12.5", False),
            ("R9999999999", "2026", False), ("H9999999999", "2000", False),
            ("R7.5", "2026", False), ("2025.5", "2026", False),
        ):
            with self.subTest(left=left, right=right):
                self.assertEqual(year_alias_compatible(left, right), expected)

    def test_all_nonblank_year_aliases_are_compared_even_when_first_alias_is_blank(self):
        self.assertTrue(aliases_compatible("", "R8", "2026"))
        self.assertTrue(aliases_compatible("H12", "", "2000"))
        self.assertFalse(aliases_compatible("", "R8", "R9"))
        self.assertFalse(aliases_compatible("", "unknown", "2026"))
        start = MODULE.index("Private Function ReviewCol(")
        col = MODULE[start:MODULE.index("End Function", start)]
        self.assertIn("Set yearReferences = CreateObject", col)
        self.assertIn("yearReferences.Exists(CStr(rowNo))", col)
        self.assertIn("yearReferences.Add CStr(rowNo), selectedText", col)

    def test_actual_reviewcol_allows_blank_but_rejects_unknown_or_conflicting_year(self):
        start = MODULE.index("Private Function ReviewCol(")
        col = MODULE[start:MODULE.index("End Function", start)]
        self.assertIn("If Len(otherText) = 0 Then otherText = selectedText", col)
        self.assertIn("If selectedYear = 0 Or otherYear = 0 Or selectedYear <> otherYear Then", col)
        self.assertIn('"別名見出しの内容が一致しません: " & ws.Name', col)
        self.assertIn('" 列" & ws.Cells(1, referenceCol).Address(False, False)', col)
        self.assertIn('"] 値=[" & selectedText & "]"', col)
        self.assertIn('" / 列" & ws.Cells(1, found).Address(False, False)', col)
        self.assertIn('"] 値=[" & otherText & "]"', col)

    def test_year_parser_supports_reiwa_and_heisei_without_decimal_rounding(self):
        strict_start = MODULE.index("Private Function ReviewStrictAliasYear(")
        strict = MODULE[strict_start:MODULE.index("End Function", strict_start)]
        self.assertIn('Replace(text, "令和", "R")', strict)
        self.assertIn('Replace(text, "平成", "H")', strict)
        self.assertIn('If Left$(text, 1) = "R" Or Left$(text, 1) = "H" Then', strict)
        self.assertIn('If AscW(c) < 48 Or AscW(c) > 57 Then Exit Function', strict)
        self.assertIn("ReviewStrictAliasYear = ReviewYear(raw)", strict)
        year_start = MODULE.index("Private Function ReviewYear(")
        parsed = MODULE[year_start:MODULE.index("End Function", year_start)]
        self.assertIn("If Len(text) > 4 Then Exit Function", parsed)
        self.assertIn("If AscW(c) < 48 Or AscW(c) > 57 Then Exit Function", parsed)

    def test_d_preflight_actually_requests_cross_notation_aliases(self):
        start = MODULE.index("Public Function SyncReviewCopyIdCandidates(")
        detector = MODULE[start:MODULE.index("End Function", start)]
        self.assertIn('ReviewCol(src, "年度（和暦）", "年度(和暦)", "年度")', detector)
        self.assertIn('ReviewCol(src, "作成年度（和暦）", "作成年度(和暦)", "作成年度")', detector)
        self.assertIn('ReviewCol(dst, "年度（和暦）", "年度(和暦)", "年度")', detector)
        self.assertIn('ReviewCol(dst, "作成年度（和暦）", "作成年度(和暦)", "作成年度")', detector)


if __name__ == "__main__":
    unittest.main()
