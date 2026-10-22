"""Static safety contracts for the U-column (ファイルID) direct-edit restriction."""
from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba" / "Module1.txt").read_text(encoding="utf-8")
SHEET = (ROOT / "vba" / "新ファイル基準表.txt").read_text(encoding="utf-8")
REVIEW = (ROOT / "vba" / "modSyncReview.txt").read_text(encoding="utf-8")
MANUAL = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8")


def procedure(source: str, name: str, kind: str = "Sub") -> str:
    match = re.search(
        rf"(?ms)^(?:Private|Public) {kind} {re.escape(name)}\b.*?^End {kind}",
        source,
    )
    if match is None:
        raise AssertionError(f"missing {kind}: {name}")
    return match.group()


class FileIdEditRestrictionTests(unittest.TestCase):
    def test_sheet_captures_prechange_u_column_snapshots(self) -> None:
        activate = procedure(SHEET, "Worksheet_Activate")
        selection = procedure(SHEET, "Worksheet_SelectionChange")
        event = procedure(SHEET, "Worksheet_Change")
        self.assertIn("mFileIdSelectionSnapshot", SHEET)
        self.assertIn("CaptureFileIdSelectionSnapshot Application.ActiveCell, Me", activate)
        self.assertIn("CaptureFileIdSelectionSnapshot Target, Me", selection)
        self.assertIn("CaptureFileIdSelectionSnapshot Application.Selection, Me", event)
        self.assertLess(
            event.index("新ファイル基準表_ファイルID直接編集を復元(Target, Me)"),
            event.index("新ファイル基準表_最終編集情報を記録 Target, Me"),
        )

    def test_direct_file_id_edits_restore_only_u_with_formula_and_format(self) -> None:
        restore = procedure(SHEET, "新ファイル基準表_ファイルID直接編集を復元", "Function")
        capture = procedure(SHEET, "CaptureFileIdCellSnapshot", "Function")
        writeback = procedure(SHEET, "RestoreFileIdCellFromSnapshot")
        self.assertIn('"ファイルID"', restore)
        self.assertIn("Application.Intersect(Target, ws.Columns(fileIdCol))", restore)
        self.assertIn("rowsToRestore", restore)
        self.assertIn("NewFileStandardTargetHasNonFileIdCells", restore)
        self.assertNotIn('"ファイル番号"', restore)
        self.assertNotIn('"バーコード番号"', restore)
        self.assertIn('snapshot.Add "Formula"', capture)
        self.assertIn('snapshot.Add "NumberFormat"', capture)
        self.assertIn("targetCell.NumberFormat", writeback)
        self.assertIn("targetCell.Formula", writeback)
        self.assertIn("復元結果を確認できません", writeback)

    def test_unsafe_partial_restore_undoes_the_whole_user_operation_and_stops(self) -> None:
        undo = procedure(SHEET, "UndoUnsafeFileIdEdit", "Function")
        restore = procedure(SHEET, "新ファイル基準表_ファイルID直接編集を復元", "Function")
        self.assertIn("Application.Undo", undo)
        self.assertIn("編集操作全体を取り消しました", undo)
        self.assertIn("ブックを保存せず終了してください", undo)
        self.assertIn('"登録予定"', restore)
        self.assertIn("UndoUnsafeFileIdEdit(Target, ws)", restore)

    def test_csv_and_sync_macro_updates_invalidate_stale_snapshots_while_events_are_off(self) -> None:
        invalidate = procedure(MODULE, "新ファイル基準表_ファイルID編集スナップショットを無効化")
        epoch = procedure(MODULE, "新ファイル基準表_ファイルID編集スナップショット世代", "Function")
        step2 = procedure(MODULE, "手順2_新ファイル基準表を作成する")
        self.assertGreater(
            MODULE.index("Public Sub 新ファイル基準表_ファイルID編集スナップショットを無効化"),
            MODULE.index("Private gRowMoveClickEndRow As Long"),
        )
        self.assertLess(
            MODULE.index("Public Sub 新ファイル基準表_ファイルID編集スナップショットを無効化"),
            MODULE.index("Public Sub WriteProcessLog"),
        )
        clear_copy = procedure(REVIEW, "SyncReviewClearCopyIdentifiers", "Function")
        approve = procedure(REVIEW, "SyncReviewApprove", "Function")
        self.assertIn("mFileIdEditSnapshotEpoch", invalidate)
        self.assertIn("mFileIdEditSnapshotEpoch", epoch)
        self.assertLess(
            step2.index("新ファイル基準表_ファイルID編集スナップショットを無効化"),
            step2.index("Application.EnableEvents = False"),
        )
        self.assertLess(
            clear_copy.index("新ファイル基準表_ファイルID編集スナップショットを無効化"),
            clear_copy.index("Application.EnableEvents = False"),
        )
        self.assertLess(
            approve.index("新ファイル基準表_ファイルID編集スナップショットを無効化"),
            approve.index("Application.EnableEvents = False"),
        )

    def test_manual_documents_scope_and_preserves_v_and_w_rules(self) -> None:
        self.assertIn("U列「ファイルID」", MANUAL)
        self.assertIn("直接入力・削除・貼り付けはできません", MANUAL)
        self.assertIn("CSV取込、CSV整形、同期候補確認", MANUAL)
        self.assertIn("V列「ファイル番号」とW列「バーコード番号」の直接編集ルールは変更しません", MANUAL)


if __name__ == "__main__":
    unittest.main()