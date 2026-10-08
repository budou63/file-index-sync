from __future__ import annotations

from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "Module1").read_text(encoding="utf-8")


def resolve_level3(rows: list[dict[str, str]], parent_code: str, display_name: str) -> str:
    matches = {row["code"] for row in rows if row["item"] == "分類名３" and row["parent"] == parent_code and row["name"] == display_name and row["state"] == "システム"}
    return next(iter(matches)) if len(matches) == 1 else ""


class Step4ParentGuideResolutionTests(unittest.TestCase):
    SYSTEM_ROWS = [
        {"item": "分類名３", "parent": "230", "name": "用地取得資料", "code": "640", "state": "システム"},
        {"item": "分類名３", "parent": "315", "name": "用地取得資料", "code": "796", "state": "システム"},
        {"item": "分類名３", "parent": "320", "name": "用地取得資料", "code": "799", "state": "システム"},
        {"item": "分類名３", "parent": "321", "name": "用地取得資料", "code": "805", "state": "システム"},
    ]

    def test_same_second_guide_resolves_by_parent_guide_code(self) -> None:
        expected = {"230": "640", "315": "796", "320": "799", "321": "805"}
        for parent, code in expected.items():
            with self.subTest(parent=parent):
                self.assertEqual(resolve_level3(self.SYSTEM_ROWS, parent, "用地取得資料"), code)

    def test_unknown_parent_fails_closed(self) -> None:
        self.assertEqual(resolve_level3(self.SYSTEM_ROWS, "999", "用地取得資料"), "")

    def test_code_master_records_parent_context_and_step4_uses_it(self) -> None:
        self.assertIn("親分類コード", SOURCE)
        self.assertIn('BuildGuideCodeMasterNameKey("分類名３", NormalizeClassificationGuideKey(secondGuideName), firstGuideCode)', SOURCE)
        self.assertIn('GetStep4GuideCodeStateIssue("第2ガイド", "分類名３", secondGuideName, parentCodeValue)', SOURCE)
        self.assertIn('AppendCodeMasterRecord "分類名３", codeValue, guideName, "仮採番", parentCodeValue', SOURCE)

    def test_child_code_recompletion_and_allocation_are_parent_scoped(self) -> None:
        rebuild = re.search(r"Public Sub 新ファイル基準表_分類コード再補完[\s\S]*?End Sub", SOURCE).group(0)
        allocator = re.search(r"Public Sub 新ファイル基準表_登録予定分類コード採番[\s\S]*?End Sub", SOURCE).group(0)
        self.assertIn("GetCodeFromCodeMaster(CStr(pairDefs(i)(0)), inputName, parentCodeValue)", rebuild)
        self.assertIn('BuildCodeMasterNameCodeMap("分類名３", usedCode3, True, True)', allocator)
        self.assertIn("Len(parentCodeValue) = 0", allocator)


if __name__ == "__main__":
    unittest.main()
