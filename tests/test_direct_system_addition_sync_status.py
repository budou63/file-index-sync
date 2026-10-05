from __future__ import annotations

from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "Module1").read_text(encoding="utf-8")
MANUAL = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"{name} が見つかりません")
    return match.group(0)


class DirectSystemNewFileSyncStatusTests(unittest.TestCase):
    def test_newly_appended_system_rows_are_marked_synced(self) -> None:
        step2 = procedure("手順2_新ファイル基準表を作成する")
        append_branch = step2.split("AppendRowFromSourceByHeader wsSrc, wsDst, srcRow, appendRow", 1)[1].split("End If", 1)[0]
        self.assertRegex(append_branch, r"wsDst\.Cells\(appendRow,\s*syncCol\)\.Value\s*=\s*\"同期済\"")

    def test_manual_explains_state_for_direct_system_additions(self) -> None:
        self.assertIn("システム側で直接新規作成", MANUAL)
        self.assertIn("同期済", MANUAL)
        self.assertNotIn("同期状態はユーザー管理列のため、空欄になる場合があります", MANUAL)


if __name__ == "__main__":
    unittest.main()
