"""Static safety contracts for clearing copied IDs on registration-planned rows."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8")
SHEET = (ROOT / "vba/新ファイル基準表.txt").read_text(encoding="utf-8")


def procedure(source: str, name: str, kind: str = "Sub") -> str:
    match = re.search(
        rf"(?ms)^(?:Private|Public) {kind} {re.escape(name)}\b.*?^End {kind}",
        source,
    )
    if match is None:
        raise AssertionError(f"missing {kind}: {name}")
    return match.group()


class RegistrationIdentifierClearingTests(unittest.TestCase):
    def test_sheet_event_preserves_existing_work_and_restores_prior_event_state(self):
        event = procedure(SHEET, "Worksheet_Change")
        for existing_call in (
            "新ファイル基準表_表示名からID自動補完 Target",
            "新ファイル基準表_廃棄年月日自動計算 Target",
            "ShouldReindexSerial(Target, Me)",
            "ReindexSerial_NewFileStandard Me",
        ):
            self.assertIn(existing_call, event)
        self.assertIn("新ファイル基準表_登録予定識別番号を自動クリア Target, Me", event)
        self.assertIn("oldEnableEvents = Application.EnableEvents", event)
        self.assertIn("Application.EnableEvents = oldEnableEvents", event)

    def test_change_helper_uses_headers_changed_rows_and_registration_status_only(self):
        helper = procedure(MODULE, "新ファイル基準表_登録予定識別番号を自動クリア")
        for header in ("ファイルID", "ファイル番号", "バーコード番号", "同期状態"):
            self.assertIn(f'Array("{header}")', helper)
        self.assertIn("Application.Intersect(Target", helper)
        self.assertIn("For Each changedArea In changedRange.Areas", helper)
        self.assertIn("ws.Cells(ws.Rows.Count, syncCol).End(xlUp).Row", helper)
        self.assertNotIn("Step2GetLastDataRow", helper)
        self.assertIn('GetSyncStatusValue(ws, rowNo, syncCol) = "登録予定"', helper)
        self.assertIn(".ClearContents", helper)
        self.assertIn("oldEnableEvents = Application.EnableEvents", helper)
        self.assertIn("Application.EnableEvents = oldEnableEvents", helper)
        self.assertIn("On Error GoTo", helper)

    def test_step2_confirmation_is_year_scoped_and_no_residual_means_no_prompt(self):
        prepare = procedure(MODULE, "Step2PrepareRegistrationIdentifierCleanup", "Function")
        self.assertIn("GetNormalizedYearFromRow", prepare)
        self.assertIn("<> targetYear", prepare)
        self.assertIn('GetSyncStatusValue(ws, rowNo, syncCol) <> "登録予定"', prepare)
        self.assertIn("residualCellCount = 0", prepare)
        self.assertIn("MsgBox", prepare)
        self.assertIn("登録予定の行に識別番号が残っています。", prepare)
        self.assertIn("vbYesNo", prepare)
        self.assertIn("vbDefaultButton2", prepare)
        self.assertIn("oldValues", prepare)

    def test_confirmation_is_read_only_until_all_prompts_and_backup_precedes_clear(self):
        prepare = procedure(MODULE, "Step2PrepareRegistrationIdentifierCleanup", "Function")
        apply_cleanup = procedure(MODULE, "Step2ApplyRegistrationIdentifierCleanup")
        self.assertNotIn(".ClearContents", prepare)
        self.assertLess(prepare.index("MsgBox"), len(prepare))
        backup = apply_cleanup.index("Step2CreateRegistrationIdentifierBackup")
        clear = apply_cleanup.index(".ClearContents")
        self.assertLess(backup, clear)
        self.assertIn("SaveCopyAs", procedure(MODULE, "Step2CreateRegistrationIdentifierBackup", "Function"))
        self.assertIn("RestoreRegistrationIdentifierValues", apply_cleanup)

    def test_step2_waits_for_existing_confirmation_before_identifier_cleanup(self):
        step2 = procedure(MODULE, "手順2_新ファイル基準表を作成する")
        first_status = step2.index("Step2ConfirmMissingSyncStatus")
        first_prepare = step2.index("Step2PrepareRegistrationIdentifierCleanup")
        first_apply = step2.index("Step2ApplyRegistrationIdentifierCleanup")
        self.assertLess(first_status, first_prepare)
        self.assertLess(first_prepare, first_apply)
        self.assertLess(first_apply, step2.index("BuildHeaderMap wsDst"))
        deferred_status = step2.rindex("Step2ConfirmMissingSyncStatus")
        deferred_apply = step2.rindex("Step2ApplyRegistrationIdentifierCleanup")
        self.assertLess(step2.index("Step2PrepareRegistrationIdentifierCleanup"), step2.index("EnsureColumnByHeader"))
        self.assertLess(deferred_status, deferred_apply)
        self.assertLess(deferred_apply, step2.index("Step2AddMissingSyncStatusComments"))
        self.assertLess(step2.index("Application.EnableEvents = False"), step2.index("For srcRow = 2 To srcLastRow"))
        self.assertLess(step2.index("Application.EnableEvents = False"), step2.index("Step2ApplyRegistrationIdentifierCleanup"))
        self.assertIn("Application.EnableEvents = previousEnableEvents", step2)
        self.assertIn("If Not Step2PrepareRegistrationIdentifierCleanup", step2)


if __name__ == "__main__":
    unittest.main()
