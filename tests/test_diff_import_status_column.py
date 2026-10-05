from __future__ import annotations

from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "Module1").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"{name} が見つかりません")
    return match.group(0)


class DiffImportStatusColumnTests(unittest.TestCase):
    def test_all_four_state_outcomes_map_to_the_required_labels(self) -> None:
        status = procedure("GetStep4OutputStatus")
        cases = {
            (True, False): "新規登録",
            (True, True): "新規登録／座作成要",
            (False, False): "既存更新",
            (False, True): "既存更新／座作成要",
        }
        for (is_new, requires_seat), label in cases.items():
            with self.subTest(is_new=is_new, requires_seat=requires_seat):
                self.assertIn(f'GetStep4OutputStatus = "{label}"', status)
        self.assertIn("If isNewRegistration Then", status)
        self.assertIn("If requiresGuideSeat Then", status)

    def test_all_four_status_labels_are_supported(self) -> None:
        status = procedure("GetStep4OutputStatus")
        for label in ("新規登録", "新規登録／座作成要", "既存更新", "既存更新／座作成要"):
            with self.subTest(label=label):
                self.assertIn(f'"{label}"', status)

    def test_guide_seat_check_covers_every_output_candidate(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        start = step4.index("ValidateStep4GuideCodeStates")
        loop_start = step4.index("For Each targetRow In targetRows", start)
        loop_end = step4.index("Next targetRow", loop_start)
        loop = step4[loop_start:loop_end]
        self.assertIn("GetStep4GuideSeatCheckReason", loop)
        self.assertNotIn("If newRegistrationRows.Exists", loop)

    def test_status_is_a_helper_column_before_serial_without_changing_fixed_headers(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertNotIn('"状態"', step4.split("fixedHeaders = Array(", 1)[1].split(")", 1)[0])
        self.assertIn('wsOut.Cells(1, 1).Value = "状態"', step4)
        self.assertIn('wsOut.Cells(1, i - LBound(fixedHeaders) + 2).Value = fixedHeaders(i)', step4)
        self.assertIn('wsOut.Cells(outRow, 1).Value = GetStep4OutputStatus(', step4)

    def test_export_excludes_state_and_serial_by_header_name(self) -> None:
        export = procedure("手順5_差分インポートCSVを外部出力する")
        self.assertIn('If headerName <> "通し番号" And headerName <> "状態" Then', export)

    def test_yellow_fill_is_kept_for_rows_requiring_guide_creation(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("RGB(255, 255, 0)", step4)
        self.assertIn("座作成要確認：", step4)


if __name__ == "__main__":
    unittest.main()
