from __future__ import annotations

from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "frmSeal_").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"Private\s+Sub\s+{re.escape(name)}\s*\(\s*\).*?\nEnd\s+Sub",
        FORM,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"{name} が見つかりません")
    return match.group(0)


class UserFormCommandButton10Tests(unittest.TestCase):
    def test_command_button_10_calls_existing_generator_once(self) -> None:
        handler = procedure("CommandButton10_Click")
        self.assertEqual(len(re.findall(r"\b新ファイル基準表生成\b", handler)), 1)
        self.assertNotIn("手順2_新ファイル基準表を作成する", handler)
        self.assertNotIn("手順3_コード管理CSVを作成する", handler)

    def test_existing_button_1_to_4_assignments_are_unchanged(self) -> None:
        expected = {
            "CommandButton1_Click": "CSV取込_シンプル版",
            "CommandButton2_Click": "手順2_新ファイル基準表を作成する",
            "CommandButton3_Click": "手順4_差分インポートCSVを作成する",
            "CommandButton4_Click": "手順5_差分インポートCSVを外部出力する",
        }
        for event_name, macro_name in expected.items():
            with self.subTest(event_name=event_name):
                handler = procedure(event_name)
                self.assertEqual(len(re.findall(rf"\b{re.escape(macro_name)}\b", handler)), 1)
                self.assertEqual(len(re.findall(r"(?m)^\s*(?!Private|End|$)[^']+\s*$", handler)), 1)


if __name__ == "__main__":
    unittest.main()
