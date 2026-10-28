"""Static safety contracts for linked classification-name/code review."""
from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8-sig")
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")
MANUAL = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8-sig")
GUIDE = (ROOT / "同期候補確認フォーム_導入手順.md").read_text(encoding="utf-8-sig")


def procedure(source: str, name: str, kind: str) -> str:
    markers = (f"Private {kind} {name}(", f"Public {kind} {name}(", f"{kind} {name}(")
    start = next((source.index(marker) for marker in markers if marker in source), None)
    if start is None:
        raise AssertionError(f"procedure not found: {name}")
    end = source.index(f"End {kind}", start)
    return source[start:end]


class ClassificationLinkSafetyTests(unittest.TestCase):
    def test_mode_labels_are_consistent_without_changing_internal_indices(self):
        self.assertIn('mCboEVisibility.AddItem "入力項目のみ"', FORM)
        self.assertIn('mCboEVisibility.AddItem "全項目（管理コード含む）"', FORM)
        self.assertNotIn('mCboEVisibility.AddItem "通常表示"', FORM)
        self.assertNotIn('mCboEVisibility.AddItem "すべて表示"', FORM)
        self.assertIn("mCboEVisibility.ListIndex = 0", FORM)
        self.assertIn("EIsAllDisplayMode = (mCboEVisibility.ListIndex = 1)", procedure(FORM, "EIsAllDisplayMode", "Function"))
        for document in (MANUAL, GUIDE):
            self.assertIn("入力項目のみ", document)
            self.assertIn("全項目（管理コード含む）", document)

    def test_only_classification1_can_use_the_hidden_member_path(self):
        safe = procedure(FORM, "ESelectedDecisionIsSafe", "Function")
        details = procedure(FORM, "EConflictFieldDetails", "Function")
        next_case = procedure(FORM, "SelectNextECase", "Sub")
        self.assertIn('If EIsClassification1Group(item) Then', safe)
        self.assertIn("Module1.BidirectionalSyncClassification1DecisionSafe", safe)
        self.assertIn("If Not EIsAllDisplayMode() And ECaseHasOutOfScopeFields(item) Then", safe)
        self.assertIn("非表示の分類コード１も同じ判断単位に含まれます", details)
        self.assertIn("ECaseHasOutOfScopeFields(item) And Not EIsClassification1Group(item)", next_case)

    def test_map_validation_requires_one_exact_data_driven_pair(self):
        safety = procedure(MODULE, "BidirectionalSyncClassification1DecisionSafe", "Function")
        mapping = procedure(MODULE, "BidirectionalSyncClassification1PairIsUnique", "Function")
        for token in (
            "BidirectionalSyncConflictLiveMatches(item, reason)",
            'CStr(item("groupName")) <> "分類名１・分類コード１"',
            'GetStep2HeaderCandidates("分類名１")',
            'Array("分類コード１")',
            "sourceMapOk = BidirectionalSyncClassification1PairIsUnique(sourceCode, sourceName",
            "excelMapOk = BidirectionalSyncClassification1PairIsUnique(excelCode, excelName",
            "selectedMapOk",
        ):
            self.assertIn(token, safety)
        for token in (
            "SHEET_CODE_MANAGEMENT",
            "SHEET_MOKUROKU_CSV", "sourcePairFound", "sourceNameColumn", "sourceCodeColumn",
            '"項目名"', '"コード"', '"表示名"', '"状態"', '"親分類コード"',
            "rowCode, targetCode", "rowName, targetName", "exactCount <> 1",
            "Len(parentCode) > 0", 'StrComp(stateValue, "システム", vbTextCompare)',
        ):
            self.assertIn(token, mapping)
        self.assertNotIn("土木", FORM + MODULE)
        self.assertNotIn("工務", FORM + MODULE)
        self.assertIn("既存判断グループに未対応の関連項目があります", safety)
        self.assertIn("一部だけの判断を防ぐため採用・保留を停止しました", safety)

    def test_confirmation_lists_file_source_pair_related_fields_and_resulting_state(self):
        confirm = procedure(FORM, "EConfirmClassification1Decision", "Function")
        for token in (
            "対象ファイル（個別フォルダー）", "ERecordIdentityDetails(record)",
            "採用元: ", "previewText", "採用後の判断状態: ", "vbYesNo",
            "分類名１・分類コード１の連動判断確認", "EClearClassification1PreviewSignature item",
        ):
            self.assertIn(token, confirm)
        for name, decision in (("mCmdApprove_Click", '"システム採用"'), ("RecordChoice", "decision")):
            body = procedure(FORM, name, "Sub")
            self.assertIn("EConfirmClassification1Decision", body)
            self.assertLess(body.index("EConfirmClassification1Decision"), body.index("BidirectionalSyncResolveConflict(item,"))
            self.assertIn(decision, body)
        self.assertIn("item.Remove \"linkedPairSignature\"", procedure(MODULE, "BidirectionalSyncResolveConflict", "Function"))

    def test_prior_decision_can_be_changed_after_live_revalidation(self):
        safety = procedure(MODULE, "BidirectionalSyncClassification1DecisionSafe", "Function")
        resolver = procedure(MODULE, "BidirectionalSyncResolveConflict", "Function")
        self.assertNotIn("currentDecision", safety)
        self.assertNotIn("既存判断を上書きしません", safety)
        self.assertIn("requireSignature", safety)
        self.assertIn("linkedPairSignature", safety)
        self.assertIn("oldDecision = wsConflict.Cells(stateRow, 6).Value2", resolver)
        self.assertIn("BidirectionalSyncConflictLiveMatches(item, reason)", resolver)
        self.assertIn("BidirectionalSyncFindConflictStateRow(item, stateRow)", resolver)
        self.assertIn("wsConflict.Cells(stateRow, 6).Value2 = decision", resolver)
        self.assertNotIn("previousDecision", resolver)
        self.assertIn("snapshots As Collection", resolver)
        self.assertIn("restoreOk", resolver)

    def test_mapping_failure_blocks_adoption_but_hold_keeps_values_unchanged(self):
        safety = procedure(MODULE, "BidirectionalSyncClassification1DecisionSafe", "Function")
        confirm = procedure(FORM, "EConfirmClassification1Decision", "Function")
        resolver = procedure(MODULE, "BidirectionalSyncResolveConflict", "Function")
        self.assertIn("Not selectedMapOk Then", safety)
        self.assertIn("対応を一意に確認できないため採用できません", safety)
        self.assertIn('Case BIDIRECTIONAL_SYNC_CONFLICT_HOLD', safety)
        self.assertIn("採用元なし（値は変更せず保留）", safety)
        self.assertIn('sourceLabel = "なし（値は変更しません）"', confirm)
        self.assertIn('stateLabel = "保留（関連項目グループ全体）"', confirm)
        self.assertIn("decision = BIDIRECTIONAL_SYNC_CONFLICT_SYSTEM", resolver)

    def test_rejudge_keeps_the_prior_choice_only_for_the_same_key_and_snapshot(self):
        refresh_state = procedure(MODULE, "BidirectionalSyncRefreshConflictState", "Sub")
        self.assertIn("BidirectionalSyncConflictHistoryKey(targetYearValue, rowId, stableKey, groupName, snapshot)", refresh_state)
        self.assertIn("If previousByKey.Exists(historyKey) Then decision = CStr(previousByKey(historyKey))", refresh_state)


if __name__ == "__main__":
    unittest.main()
