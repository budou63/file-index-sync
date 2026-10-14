"""Stage-3 Step2 missing-row collection: source-linked, not Excel execution."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(rf"(?:Public|Private)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)", SOURCE, re.S)
    if match is None:
        raise AssertionError(f"missing VBA procedure: {name}")
    return match.group(0)


class MissingRowCollectionTests(unittest.TestCase):
    def test_existing_three_paths_and_warning_contract(self):
        step2 = procedure("手順2_新ファイル基準表を作成する")
        warnings = procedure("ShowStep2MissingSourceWarnings")
        first = step2.index("If Len(sourceFileIdValue) = 0 Then")
        next_row = step2.index("GoTo ContinueNextSourceRow", first)
        self.assertIn("If Len(nameSyncKey) = 0 Then", step2[first:next_row])
        self.assertLess(first, step2.index("If Len(nameSyncKey) = 0 Then", first))
        later = step2.index('If nameSyncKey = "" Then', next_row)
        self.assertLess(next_row, later)
        self.assertIn('If missingFileIdCount > 20 Then missingFileIdRows = missingFileIdRows & " ほか"', warnings)
        self.assertIn('If missingNameDataCount > 20 Then missingNameRows = missingNameRows & " ほか"', warnings)
        self.assertIn('"正式ファイルID不足で保留: " & missingFileIdCount & " 件（目録CSV行: " & missingFileIdRows', step2)
        self.assertIn("On Error GoTo ErrorHandler", step2)
        self.assertIn("RestoreCalculationIfPossible previousCalculationMode", step2)

    def test_three_calls_share_the_exact_cap_after_existing_guards(self):
        step2 = procedure("手順2_新ファイル基準表を作成する")
        helper = procedure("RecordStep2MissingSourceRow")
        self.assertIn("ByRef rowCount As Long", helper)
        self.assertIn("ByRef rowList As String", helper)
        self.assertIn("ByVal rowNo As Long", helper)
        self.assertIn("rowCount = rowCount + 1", helper)
        self.assertIn("If rowCount <= 20 Then _", helper)
        self.assertIn('rowList = rowList & IIf(Len(rowList) > 0, "、", "") & CStr(rowNo)', helper)
        first = step2.index("If Len(sourceFileIdValue) = 0 Then")
        next_row = step2.index("GoTo ContinueNextSourceRow", first)
        branch = step2[first:next_row]
        self.assertIn("RecordStep2MissingSourceRow missingFileIdCount, missingFileIdRows, srcRow", branch)
        self.assertIn("If Len(nameSyncKey) = 0 Then\n                    RecordStep2MissingSourceRow missingNameDataCount, missingNameRows, srcRow", branch)
        self.assertLess(branch.index("RecordStep2MissingSourceRow missingNameDataCount"), branch.index("RecordStep2IdentityHold"))
        later = step2.index('If nameSyncKey = "" Then', next_row)
        later_branch = step2[later:step2.index("GoTo ContinueNextSourceRow", later)]
        self.assertIn("RecordStep2MissingSourceRow missingNameDataCount, missingNameRows, srcRow", later_branch)
        self.assertLess(later_branch.index("RecordStep2MissingSourceRow"), later_branch.index("Step2ThreeNameKey"))
        self.assertEqual(step2.count("RecordStep2MissingSourceRow "), 3)
        self.assertNotIn("If missingFileIdCount <= 20 Then _", step2)
        self.assertNotIn("If missingNameDataCount <= 20 Then _", step2)

    def test_decision_example_keeps_first_twenty_in_encounter_order(self):
        # Model only; the source assertions above bind each VBA call site.
        def collect(row_numbers: list[int]) -> tuple[int, str]:
            count, rows = 0, ""
            for row in row_numbers:
                count += 1
                if count <= 20:
                    rows += ("、" if rows else "") + str(row)
            return count, rows
        for limit in (0, 1, 20, 21, 22):
            with self.subTest(limit=limit):
                actual = collect(list(range(2, 2 + limit)))
                displayed = "、".join(str(row) for row in range(2, 2 + min(limit, 20)))
                self.assertEqual(actual, (limit, displayed))


if __name__ == "__main__":
    unittest.main()
