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
    """Model the requested safety contract: ID first, then unique name fallback."""
    def id_key(row: dict) -> tuple[str, str, str] | None:
        values = tuple(row.get(k, "") for k in ("id", "no", "barcode"))
        return values if any(values) else None

    def name_key(row: dict) -> tuple | None:
        values = tuple(row.get(k, "") for k in ("year", "class2", "class3", "title"))
        return values if all(values) else None

    row_id = id_key(edit_row)
    if row_id is not None:
        matches = [row for row in source_rows if id_key(row) == row_id]
        return "existing" if len(matches) == 1 else "difficult"

    row_name = name_key(edit_row)
    if row_name is None:
        return "difficult"
    matches = [row for row in source_rows if name_key(row) == row_name]
    if len(matches) == 1 and edit_name_counts[row_name] == 1:
        return "existing"
    if not matches:
        return "difficult" if edit_row.get("status") == "同期済" else "new"
    return "difficult"


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

    def test_step4_uses_reusable_id_and_name_identity_helpers(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("BuildFileKeyFromSheetRow", step4)
        self.assertIn("BuildFolderNameSyncKeyFromSheetRow", step4)
        self.assertIn("ResolveStep4RowMatch", step4)
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


if __name__ == "__main__":
    unittest.main()
