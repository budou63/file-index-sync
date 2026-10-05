from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import re
import unittest


SOURCE = (Path(__file__).resolve().parents[1] / "Module1").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"{name} が見つかりません")
    return match.group(0)


def resolve(source_rows: list[dict], edit_row: dict, edit_name_counts: dict[tuple, int]) -> str:
    """Model ID-first matching with unique auxiliary-ID and name fallbacks."""
    def value(row: dict, key: str) -> str:
        return row.get(key, "") or ""

    def name_key(row: dict) -> tuple | None:
        parts = tuple(value(row, k) for k in ("year", "class2", "class3", "title"))
        return parts if all(parts) else None

    if value(edit_row, "id"):
        matches = [row for row in source_rows if value(row, "id") == value(edit_row, "id")]
        return "existing" if len(matches) == 1 else "difficult"

    auxiliary_hits: dict[int, dict] = {}
    for field in ("no", "barcode"):
        token = value(edit_row, field)
        if token:
            matches = [row for row in source_rows if value(row, field) == token]
            if len(matches) > 1:
                return "difficult"
            if matches:
                auxiliary_hits[id(matches[0])] = matches[0]
    if len(auxiliary_hits) == 1:
        return "existing"
    if len(auxiliary_hits) > 1:
        return "difficult"

    key = name_key(edit_row)
    if key is None:
        return "difficult"
    matches = [row for row in source_rows if name_key(row) == key]
    if len(matches) == 1 and edit_name_counts.get(key, 0) == 1:
        return "existing"
    if matches:
        return "difficult"
    return "difficult" if edit_row.get("status") == "同期済" else "new"


class Step4IdentityResolutionTests(unittest.TestCase):
    SOURCE_ROW = {"id": "F-1", "no": "10", "barcode": "B-1", "year": "2026", "class2": "総務", "class3": "庶務", "title": "旧題", "period": "継続", "media": "紙"}

    def test_id_match_is_existing_even_if_mutable_fields_change(self) -> None:
        edit = {**self.SOURCE_ROW, "title": "新題", "period": "5年", "media": "電子"}
        self.assertEqual(resolve([self.SOURCE_ROW], edit, defaultdict(int)), "existing")

    def test_blank_id_unique_name_match_is_existing(self) -> None:
        edit = {**self.SOURCE_ROW, "id": "", "no": "", "barcode": ""}
        key = ("2026", "総務", "庶務", "旧題")
        self.assertEqual(resolve([self.SOURCE_ROW], edit, {key: 1}), "existing")

    def test_blank_id_missing_name_is_new_except_synced_status(self) -> None:
        edit = {**self.SOURCE_ROW, "id": "", "no": "", "barcode": "", "title": "新規", "status": ""}
        self.assertEqual(resolve([self.SOURCE_ROW], edit, {("2026", "総務", "庶務", "新規"): 1}), "new")
        self.assertEqual(resolve([self.SOURCE_ROW], {**edit, "status": "同期済"}, {("2026", "総務", "庶務", "新規"): 1}), "difficult")

    def test_id_or_name_ambiguity_is_not_new(self) -> None:
        edit_with_id = {**self.SOURCE_ROW}
        self.assertEqual(resolve([self.SOURCE_ROW, dict(self.SOURCE_ROW)], edit_with_id, defaultdict(int)), "difficult")
        edit_without_id = {**self.SOURCE_ROW, "id": "", "no": "", "barcode": ""}
        key = ("2026", "総務", "庶務", "旧題")
        self.assertEqual(resolve([self.SOURCE_ROW], edit_without_id, {key: 2}), "difficult")

    def test_file_id_match_ignores_missing_file_number(self) -> None:
        source = {**self.SOURCE_ROW}
        edit = {**self.SOURCE_ROW, "no": ""}
        self.assertEqual(resolve([source], edit, defaultdict(int)), "existing")

    def test_file_id_match_takes_priority_over_barcode_mismatch(self) -> None:
        source = {**self.SOURCE_ROW}
        edit = {**self.SOURCE_ROW, "barcode": "B-OTHER"}
        self.assertEqual(resolve([source], edit, defaultdict(int)), "existing")

    def test_missing_file_id_uses_unique_file_number(self) -> None:
        source = {**self.SOURCE_ROW, "id": ""}
        edit = {**source}
        self.assertEqual(resolve([source], edit, defaultdict(int)), "existing")

    def test_missing_file_id_uses_unique_barcode(self) -> None:
        source = {**self.SOURCE_ROW, "id": "", "no": ""}
        edit = {**source}
        self.assertEqual(resolve([source], edit, defaultdict(int)), "existing")

    def test_auxiliary_id_ambiguity_is_difficult(self) -> None:
        source = [{**self.SOURCE_ROW, "id": ""}, {**self.SOURCE_ROW, "id": "", "title": "別"}]
        edit = {**self.SOURCE_ROW, "id": "", "title": "新題"}
        self.assertEqual(resolve(source, edit, defaultdict(int)), "difficult")

    def test_step4_reports_guide_count_after_output_increment(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertGreater(step4.index("Debug.Print BuildStep4SummaryText"), step4.index("guideSeatCheckCount = guideSeatCheckCount + 1"))

    def test_step4_uses_reusable_identity_helpers(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        resolver = procedure("ResolveStep4RowMatch")
        self.assertIn("GetHeaderMapCellValue", step4)
        self.assertIn("BuildFolderNameSyncKeyFromSheetRow", step4)
        self.assertIn("srcFileNoMap", step4)
        self.assertIn("srcBarcodeMap", step4)
        self.assertIn("srcIdMap.Exists(fileIdValue)", resolver)
        self.assertNotIn("BuildHeaderBasedKey(yearVal, titleVal, category2Val, mediaVal, periodVal)", step4)

    def test_step4_compares_mutable_fields_after_identity_resolution(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn('compareDefs = Array("年度", "タイトル", "分類名１", "分類名２", "分類名３", "保存期間", "媒体種別")', step4)
        self.assertIn("newRegistrationRows", step4)
        self.assertIn("ambiguousCount", step4)

    def test_yellow_is_limited_to_confirmed_new_registrations(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("If newRegistrationRows.Exists", step4)
        self.assertIn("GetStep4GuideSeatCheckReason", step4)

def guide_warning(first_guide: str, second_guide: str, first_codes: set[str], second_codes: set[str]) -> str:
    if first_guide not in first_codes:
        return "第1ガイドから座作成が必要"
    if second_guide and second_guide not in second_codes:
        return "第2ガイドから座作成が必要"
    return ""


class Step4GuideSeatRuleTests(unittest.TestCase):
    def test_existing_both_guides_need_no_seat(self) -> None:
        self.assertEqual(guide_warning("A", "B", {"A"}, {"B"}), "")

    def test_missing_first_guide_requires_first_even_when_second_exists(self) -> None:
        self.assertEqual(guide_warning("A", "B", set(), {"B"}), "第1ガイドから座作成が必要")

    def test_missing_both_guides_requires_first(self) -> None:
        self.assertEqual(guide_warning("A", "B", set(), set()), "第1ガイドから座作成が必要")

    def test_existing_first_missing_second_requires_second(self) -> None:
        self.assertEqual(guide_warning("A", "B", {"A"}, set()), "第2ガイドから座作成が必要")

    def test_blank_second_guide_depends_only_on_first(self) -> None:
        self.assertEqual(guide_warning("A", "", {"A"}, set()), "")
        self.assertEqual(guide_warning("A", "", set(), set()), "第1ガイドから座作成が必要")

    def test_existing_name_with_normalized_spelling_is_existing(self) -> None:
        self.assertIn("NormalizeClassificationGuideKey", procedure("GetCodeFromCodeMasterInRows"))
        self.assertIn("NormalizeClassificationGuideKey", procedure("GetStep4GuideSeatCheckReason"))

    def test_code_snapshot_is_taken_before_registration_code_assignment(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertLess(step4.index("BuildCodeMasterNameCodeMap"), step4.index("新ファイル基準表_登録予定分類コード採番"))

    def test_only_new_registration_rows_receive_guide_warning(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("If newRegistrationRows.Exists", step4)
        self.assertIn("GetStep4GuideSeatCheckReason", step4)

    def test_yellow_counter_and_completion_use_same_warning_map(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn('If guideSeatCheckReasons.Exists(CStr(CLng(targetRow))) Then', step4)
        self.assertIn('guideSeatCheckCount = guideSeatCheckCount + 1', step4)
        self.assertGreater(step4.index("Debug.Print BuildStep4SummaryText"), step4.index("guideSeatCheckCount = guideSeatCheckCount + 1"))


if __name__ == "__main__":
    unittest.main()
