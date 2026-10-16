"""Contracts for temporary Step2 serial-number font colours."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def proc(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE,
        re.S,
    )
    if not match:
        raise AssertionError(f"missing procedure: {name}")
    return match.group(0)


def colour_for(diff_type: str) -> tuple[int, int, int]:
    return {
        "追加": (0, 128, 0),
        "登録予定現行": (0, 0, 255),
        "名称同期": (0, 0, 255),
        "変更": (255, 0, 0),
    }.get(diff_type, (0, 0, 0))


class Step2SerialFontColourTests(unittest.TestCase):
    def test_priority_colours_for_result_types(self):
        self.assertEqual(colour_for("追加"), (0, 128, 0))
        self.assertEqual(colour_for("登録予定現行"), (0, 0, 255))
        self.assertEqual(colour_for("名称同期"), (0, 0, 255))
        self.assertEqual(colour_for("変更"), (255, 0, 0))
        self.assertEqual(colour_for("判定困難"), (0, 0, 0))
        self.assertEqual(colour_for(""), (0, 0, 0))

    def test_step2_resets_only_target_year_after_processing_not_before_possible_failure(self):
        step2 = proc("手順2_新ファイル基準表を作成する")
        refresh_call = "RefreshSerialDiffMarksByKey wsDst, serialCol, visualDiffMap, serialCommentMap, targetYearValue"
        self.assertEqual(step2.count(refresh_call), 1)
        self.assertLess(step2.index("PlaceStep2NewRowsNearGroups"), step2.index(refresh_call))
        self.assertNotIn("ClearSerialDiffMarksForTargetYear wsDst, serialCol, targetYearValue", step2)
        refresh = proc("RefreshSerialDiffMarksByKey")
        self.assertIn("CaptureSerialDiffMarksForTargetYear", refresh)
        self.assertIn("ClearSerialDiffMarksForTargetYear", refresh)
        self.assertIn("ApplySerialDiffMarksByKey", refresh)
        self.assertIn("RestoreSerialDiffMarksForTargetYear", refresh)
        self.assertIn("Err.Raise", refresh)

    def test_reset_is_target_year_font_only_and_removes_only_legacy_csv_visuals(self):
        clear = proc("ClearSerialDiffMarksForTargetYear")
        self.assertIn("GetNormalizedYearFromRow", clear)
        self.assertIn("targetYearValue", clear)
        self.assertIn(".Font.Color = RGB(0, 0, 0)", clear)
        self.assertIn("SafeDeleteSerialDiffComment", clear)
        self.assertIn("SerialDiffCommentIsLegacyCsvMark(targetCell)", clear)
        self.assertIn("targetCell.Interior.Color = CSV_DIFF_SERIAL_MARK_COLOR", clear)
        self.assertNotIn("SafeDeleteComment targetCell", clear)

    def test_apply_uses_font_colour_not_new_background_fill_and_preserves_serial_comments(self):
        apply = proc("ApplySerialDiffMarksByKey")
        colour = proc("Step2SerialDiffFontColor")
        self.assertIn(".Font.Color = Step2SerialDiffFontColor", apply)
        self.assertIn("SafeReplaceSerialDiffComment", apply)
        self.assertNotIn(".Interior.Pattern = xlSolid", apply)
        self.assertNotIn(".Interior.Color = CSV_DIFF_SERIAL_MARK_COLOR", apply)
        for token in ('Case "追加"', 'Case "登録予定現行", "名称同期"', 'Case "変更"'):
            self.assertIn(token, colour)

    def test_existing_result_map_keeps_new_then_name_sync_then_change_precedence(self):
        step2 = proc("手順2_新ファイル基準表を作成する")
        append = step2.index('visualDiffMap(key) = "追加"')
        name_sync = step2.index('visualDiffMap(key) = "登録予定現行"')
        changed = step2.index('visualDiffMap(key) = "変更"')
        self.assertLess(changed, name_sync)
        self.assertLess(name_sync, append)


if __name__ == "__main__":
    unittest.main()
