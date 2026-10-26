"""Static safety contracts for linked retention-period/ID review."""
from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8-sig")
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")
MANUAL = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8-sig")


def procedure(source: str, name: str, kind: str) -> str:
    markers = (f"Private {kind} {name}(", f"Public {kind} {name}(", f"{kind} {name}(")
    start = next((source.index(marker) for marker in markers if marker in source), None)
    if start is None:
        raise AssertionError(f"procedure not found: {name}")
    end = source.index(f"End {kind}", start)
    return source[start:end]


class RetentionLinkSafetyTests(unittest.TestCase):
    def test_normal_mode_still_shows_only_period_and_has_a_scoped_linked_group_path(self):
        normal_caption = procedure(FORM, "EHeaderNormalDisplayCaption", "Function")
        self.assertIn('headerKey = CompareHeaderKey("保存期間")', normal_caption)
        self.assertNotIn('CompareHeaderKey("保存期間ID")', normal_caption)
        visible_group = procedure(FORM, "EGroupMatchesNormalDisplay", "Function")
        self.assertIn('Case "保存期間"', visible_group)
        self.assertNotIn('Case "保存期間ID"', visible_group)
        linked_group = procedure(FORM, "EIsRetentionGroup", "Function")
        self.assertIn('"保存期間・保存期間ID"', linked_group)
        safe = procedure(FORM, "ESelectedDecisionIsSafe", "Function")
        self.assertIn("If EIsRetentionGroup(item) Then", safe)
        self.assertIn("Module1.BidirectionalSyncRetentionDecisionSafe(item, decision, reasonText, previewText)", safe)
        self.assertNotIn("BidirectionalSyncClassification1DecisionSafe", safe.split("If EIsRetentionGroup(item) Then", 1)[1].split("End If", 1)[0])

    def test_enabled_buttons_validate_each_decision_and_multi_select_remains_read_only(self):
        selection = procedure(FORM, "UpdateESelection", "Sub")
        self.assertIn("ElseIf EIsRetentionGroup(item) Then", selection)
        for decision in ("システム採用", "Excel採用"):
            self.assertIn(f'ESelectedDecisionIsSafe("{decision}"', selection)
        self.assertIn('Module1.BidirectionalSyncRetentionDecisionSafe(item, "保留", pairReason, pairPreview)', selection)
        self.assertIn("mCmdApprove.Enabled = (n = 1", selection)
        self.assertIn("mCmdOther.Enabled = (n = 1", selection)
        self.assertIn("mCmdHold.Enabled = (n = 1", selection)
        self.assertIn("n > 1", selection)
        self.assertIn("EIsRetentionGroup(item)", procedure(FORM, "SelectNextECase", "Sub"))

    def test_retention_pair_checker_fails_closed_on_missing_duplicate_or_conflicting_mapping(self):
        safety = procedure(MODULE, "BidirectionalSyncRetentionDecisionSafe", "Function")
        mapping = procedure(MODULE, "BidirectionalSyncRetentionPairIsUnique", "Function")
        for token in (
            'CStr(item("groupName")) <> "保存期間・保存期間ID"',
            "BidirectionalSyncConflictLiveMatches(item, reasonText)",
            "BidirectionalSyncFindConflictStateRow(item, stateRow)",
            "既存判断を上書きしません",
            "BidirectionalSyncUnpackConflictHeaders",
            '"保存期間"', '"保存期間ID"',
            "systemMapOk", "excelMapOk", "selectedMapOk",
            "linkedRetentionSignature", "captureSignature", "requireSignature",
        ):
            self.assertIn(token, safety)
        for token in (
            "SHEET_MOKUROKU_CSV", "SHEET_CODE_MANAGEMENT",
            'GetStep2HeaderCandidates("保存期間")', 'Array("保存期間ID")',
            '"項目名"', '"コード"', '"表示名"', '"状態"', '"親分類コード"',
            "sourcePairFound", "exactCount <> 1", "Len(parentCode) > 0",
            "targetId", "targetPeriod", "sourceLastRow", "vbBinaryCompare",
            "別の保存期間にも対応しています", "複数のID候補があります", "未完成対応",
            "exactCount = exactCount + 1", "sourcePairFound = True",
        ):
            self.assertIn(token, mapping)
        self.assertNotIn("NormalizeClassificationGuideKey", mapping)
        retention_impl = safety + mapping
        self.assertNotIn('"1999"', retention_impl)
        self.assertNotIn('"1105"', retention_impl)
        self.assertNotIn('"継続", "1999"', retention_impl)
        self.assertNotIn('"5年", "1105"', retention_impl)

    def test_confirmation_captures_both_pairs_and_cancel_does_not_call_resolver(self):
        confirm = procedure(FORM, "EConfirmRetentionDecision", "Function")
        for token in (
            "対象ファイル（個別フォルダー）", "ERecordIdentityDetails(record)",
            "採用元: ", "保存期間", "保存期間ID", "採用後の判断状態: ",
            "vbYesNo", "True, False", "EClearRetentionPreviewSignature item",
        ):
            self.assertIn(token, confirm)
        self.assertLess(confirm.index("MsgBox("), confirm.index("EConfirmRetentionDecision = True"))
        self.assertIn("<> vbYes Then", confirm)
        self.assertNotIn("BidirectionalSyncResolveConflict(", confirm)
        for name in ("mCmdApprove_Click", "RecordChoice"):
            body = procedure(FORM, name, "Sub")
            self.assertIn("EConfirmRetentionDecision", body)
            self.assertLess(body.index("EConfirmRetentionDecision"), body.index("BidirectionalSyncResolveConflict(item,"))
            self.assertIn("EIsRetentionGroup(item)", body)

    def test_resolver_requires_confirmed_signature_and_reuses_group_rollback(self):
        resolver = procedure(MODULE, "BidirectionalSyncResolveConflict", "Function")
        self.assertIn('CStr(item("groupName")) = "保存期間・保存期間ID"', resolver)
        self.assertIn("BidirectionalSyncRetentionDecisionSafe(item, decision, reason, retentionPreview, False, True)", resolver)
        self.assertIn("item.Remove \"linkedRetentionSignature\"", resolver)
        self.assertIn("snapshots As Collection", resolver)
        self.assertIn("restoreOk", resolver)
        self.assertIn("For Each fieldItem In fieldItems", resolver)
        self.assertIn("WriteStep2SourceValue", resolver)
        self.assertIn("wsConflict.Cells(stateRow, 6).Value2 = decision", resolver)

    def test_manual_documents_only_the_linked_period_workflow_and_runtime_validation(self):
        for text in ("保存期間・保存期間ID", "入力項目のみ", "一意", "採用前"):
            self.assertIn(text, MANUAL)
        self.assertIn("手順4", MANUAL)


if __name__ == "__main__":
    unittest.main()
