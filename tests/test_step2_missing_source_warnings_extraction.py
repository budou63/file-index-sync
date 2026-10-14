"""Step2 warning extraction: source-path regression, not Excel/VBA execution."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?:Public|Private)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE, re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"procedure missing: {name}")
    return match.group(0)


class MissingSourceWarningsTests(unittest.TestCase):
    def test_current_warning_contract_is_fixed_before_refactor(self):
        step2 = procedure("手順2_新ファイル基準表を作成する")
        try:
            warnings = procedure("ShowStep2MissingSourceWarnings")
        except AssertionError:
            warnings = step2  # Baseline contract passes before extraction.
        expected = (
            'If missingFileIdCount > 20 Then missingFileIdRows = missingFileIdRows & " ほか"',
            'If missingFileIdCount > 0 Then MsgBox "正式ファイルID不足で自動更新・名称同期・新規追加を保留: " & _',
            'CStr(missingFileIdCount) & "件（目録CSV行: " & missingFileIdRows & "）", vbExclamation',
            'If missingNameDataCount > 20 Then missingNameRows = missingNameRows & " ほか"',
            'If missingNameDataCount > 0 Then MsgBox "名称情報・作成年度不足で自動照合と新規追加を保留: " & _',
            'CStr(missingNameDataCount) & "件（目録CSV行: " & missingNameRows & "）" & vbCrLf & _',
            '"目録CSVと登録予定行を確認してください。", vbExclamation',
        )
        positions = [warnings.index(line) for line in expected]
        self.assertEqual(positions, sorted(positions))
        self.assertIn('"正式ファイルID不足で保留: " & missingFileIdCount & " 件（目録CSV行: " & missingFileIdRows', step2)
        self.assertIn("On Error GoTo ErrorHandler", step2)
        self.assertIn("RestoreCalculationIfPossible previousCalculationMode", step2)

    def test_new_helper_preserves_warning_order_and_summary_mutation(self):
        step2 = procedure("手順2_新ファイル基準表を作成する")
        helper = procedure("ShowStep2MissingSourceWarnings")
        self.assertRegex(helper, r"ByVal missingFileIdCount As Long, _\s*ByRef missingFileIdRows As String, _\s*ByVal missingNameDataCount As Long, _\s*ByRef missingNameRows As String")
        self.assertIn("ShowStep2MissingSourceWarnings missingFileIdCount, missingFileIdRows, _\n        missingNameDataCount, missingNameRows", step2)
        self.assertEqual(step2.count("ShowStep2MissingSourceWarnings missingFileIdCount"), 1)
        self.assertLess(step2.index('appendedCount + updatedCount + nameSyncCount, "成功"'),
                        step2.index("ShowStep2MissingSourceWarnings missingFileIdCount"))
        self.assertLess(step2.index("ShowStep2MissingSourceWarnings missingFileIdCount"),
                        step2.index("If Len(missingHeadersText) > 0 Then"))
        self.assertLess(step2.index("If Len(missingHeadersText) > 0 Then"),
                        step2.index("step2Succeeded = True"))
        self.assertNotIn("If missingFileIdCount > 20 Then", step2)
        self.assertNotIn("If missingNameDataCount > 20 Then", step2)

    def test_warning_decision_examples_0_20_21(self):
        # Example only. Source checks above bind thresholds and order to the VBA.
        def warn(count: int, rows: str) -> tuple[bool, str]:
            if count > 20:
                rows = rows + " ほか"
            return count > 0, rows

        for count, expected in ((0, (False, "2、3")), (20, (True, "2、3")), (21, (True, "2、3 ほか"))):
            with self.subTest(count=count):
                self.assertEqual(warn(count, "2、3"), expected)


if __name__ == "__main__":
    unittest.main()
