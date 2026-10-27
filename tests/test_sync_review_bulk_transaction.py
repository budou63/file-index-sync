"""Source contracts for BUD-147; these are NOT Excel/VBE runtime tests."""
import re
import unittest
from pathlib import Path
from copy import deepcopy

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8-sig")


def body(name, source=MODULE, kind="Function"):
    pattern = rf"(?:Public|Private) {kind} {name}\("
    start = re.search(pattern, source)
    if not start:
        raise AssertionError(f"missing {kind} {name}")
    return source[start.start():source.index("End " + kind, start.end())]


class BulkTransactionContracts(unittest.TestCase):
    def test_plan_is_all_or_nothing_and_rechecks_every_case(self):
        plan = body("BidirectionalSyncPlanFileDecisions")
        for token in ("BidirectionalSyncConflictLiveMatches", "BidirectionalSyncFindConflictStateRow",
                      "BidirectionalSyncBatchLinkedPairSafe", '"rowId"', '"stableKey"',
                      '"year"', '"sourceRow"', '"destinationRow"',
                      "BIDIRECTIONAL_SYNC_CONFLICT_SYSTEM", "BIDIRECTIONAL_SYNC_CONFLICT_EXCEL",
                      "BIDIRECTIONAL_SYNC_CONFLICT_UNRESOLVED", "cells.Add cellState", "seenCells.Exists"):
            self.assertIn(token, plan)
        self.assertNotIn("BidirectionalSyncResolveConflict(item,", plan)
        self.assertIn("BidirectionalSyncDestinationRowIdIsUnique", plan)
        self.assertIn("BidirectionalSyncBatchSourceKeyIsUnique", plan)
        self.assertIn("If groupName = \"第1ガイド・分類コード２\" Or", plan)
        self.assertIn("BidirectionalSyncBatchLinkedPairSafe(item, BIDIRECTIONAL_SYNC_CONFLICT_HOLD", plan)

    def test_transaction_snapshots_before_writes_and_restores_all_stores(self):
        apply = body("BidirectionalSyncResolveFileDecisions")
        commit = body("BidirectionalSyncCommitFilePlan")
        plan = body("BidirectionalSyncPlanFileDecisions")
        for token in ("BidirectionalSyncPlanFileDecisions", "expectedPreview", "Application.EnableEvents = False",
                      "BidirectionalSyncRestoreFileDecisions", "vbCritical"):
            self.assertIn(token, apply)
        for token in ('"oldDecision"', '"oldUpdated"', '"value"', '"format"'):
            self.assertIn(token, plan)
        self.assertIn("WriteStep2SourceValue", commit)
        self.assertIn("wsConflict.Cells", commit)
        self.assertLess(commit.index("WriteStep2SourceValue"), commit.index("wsConflict.Cells"))
        self.assertLess(commit.index("wsConflict.Cells"), commit.index('item("decision") ='))
        restore = body("BidirectionalSyncRestoreFileDecisions")
        self.assertIn("targetCell.Value2 =", restore)
        self.assertIn("targetCell.NumberFormat =", restore)
        self.assertIn("wsConflict.Cells", restore)
        self.assertIn('item("decision") =', restore)

    def test_all_six_linked_pairs_have_data_driven_validation(self):
        validation = body("BidirectionalSyncBatchLinkedPairSafe")
        for group in ("分類名１・分類コード１", "保存期間・保存期間ID", "第1ガイド・分類コード２",
                      "第2ガイド・分類コード３", "媒体種別・媒体種別ID", "原課・原課ID"):
            self.assertIn(group, validation)
        self.assertIn("BidirectionalSyncBatchPairUnique", validation)
        unique = body("BidirectionalSyncBatchPairUnique")
        for token in ("SHEET_MOKUROKU_CSV", "SHEET_CODE_MANAGEMENT", "parentCode", "exactCount",
                      "sourcePairFound", '"システム"', "NormalizeClassificationGuideKey"):
            self.assertIn(token, unique)

    def test_first_guide_master_is_parentless_second_guide_is_parent_scoped(self):
        unique = body("BidirectionalSyncBatchPairUnique")
        master_builder = body("BuildSystemGuideCodeMasterFromMokuroku", kind="Sub")
        self.assertIn('Array("分類名２", class2CodeCol, class2NameCol, 0)', master_builder)
        self.assertIn('Array("分類名３", class3CodeCol, class3NameCol, class2CodeCol)', master_builder)
        self.assertIn('If groupName = "第1ガイド・分類コード２" Then', unique)
        self.assertIn('If Len(masterParent) > 0 Then', unique)
        self.assertIn('If StrComp(masterParent, Trim$(parentCode), vbBinaryCompare) <> 0 Then GoTo NextMaster', unique)

    def test_master_layout_fixture_preserves_guide_context_without_inventing_a_parent(self):
        """Schema fixture, not VBA execution: guide 1 is globally keyed, guide 2 has parent 2."""
        source = [("分類名２", "甲", "201", "101"),
                  ("分類名３", "乙", "305", "201")]
        master = [("分類名２", "201", "甲", "システム", ""),
                  ("分類名３", "305", "乙", "システム", "201")]

        def proven(group, name, code, parent):
            rows = [r for r in source if r[0] == group and r[3] == parent and
                    (r[1] == name or r[2] == code)]
            scope = "" if group == "分類名２" else parent
            masters = [r for r in master if r[0] == group and r[4] == scope and
                       (r[1] == code or r[2] == name)]
            return bool(rows) and all(r[1:3] == (name, code) for r in rows) and \
                len(masters) == 1 and masters[0][1:4] == (code, name, "システム")

        self.assertTrue(proven("分類名２", "甲", "201", "101"))
        self.assertTrue(proven("分類名３", "乙", "305", "201"))
        self.assertFalse(proven("分類名２", "甲", "999", "101"))
        self.assertFalse(proven("分類名３", "乙", "305", "999"))

    def test_existing_accept_buttons_choose_file_or_selected_field_without_numeric_prompts(self):
        approve = body("mCmdApprove_Click", FORM, "Sub")
        other = body("RecordChoice", FORM, "Sub")
        file_apply = body("EApplyFileDecision", FORM, "Sub")
        self.assertIn('EApplyFileDecision "システム採用"', approve)
        self.assertLess(approve.index("If fieldItem Is Nothing Then"), approve.index("Set item = SelectedCase()"))
        self.assertIn('EApplyFileDecision "Excel採用"', other)
        self.assertLess(other.index("If fieldItem Is Nothing Then"), other.index("Set item = SelectedCase()"))
        self.assertIn("BidirectionalSyncPlanFileDecisions", file_apply)
        self.assertIn("BidirectionalSyncResolveFileDecisions", file_apply)
        self.assertIn("vbYesNo", file_apply)
        self.assertIn('ESelectedCount() <> 1', file_apply)
        self.assertIn('record("identityIssue")', file_apply)
        self.assertIn('savedDecision <> "システム採用" And savedDecision <> "Excel採用"', file_apply)
        self.assertIn("choices.Add groupName, decision", file_apply)
        self.assertIn('entry("cells")', file_apply)
        self.assertIn("confirmationText", file_apply)
        self.assertNotIn("InputBox", FORM)
        self.assertNotIn("mCmdEBulk", FORM)
        self.assertNotIn("eBulkChoices", FORM)
        self.assertIn("BidirectionalSyncResolveConflict(item,", FORM)

    def test_file_level_confirmation_is_cancel_safe_and_shows_identity_and_values(self):
        file_apply = body("EApplyFileDecision", FORM, "Sub")
        identity = body("ERecordIdentityDetails", FORM)
        preview = body("BidirectionalSyncPlanFileDecisions")
        for token in ("ERecordIdentityDetails(record)", '"採用対象項目数: "',
                      "採用元: ", "確定済みで変更対象外となる項目:",
                      '" / 確定判断=" & savedDecision', '": 文書管理=" & CStr(fieldItem("systemValue"))',
                      '" / Excel=" & CStr(fieldItem("excelValue"))', "Len(confirmationText)",
                      "If MsgBox(confirmationText", "<> vbYes Then Exit Sub"):
            self.assertIn(token, file_apply)
        self.assertIn('"同期対象年度: "', identity)
        self.assertIn("現在値=", preview)
        self.assertIn("採用予定値=", preview)
        self.assertIn('" ／ 個別フォルダー: "', identity)
        self.assertIn('": 現在値="', preview)
        self.assertIn('" / 採用予定値="', preview)
        self.assertLess(file_apply.index("BidirectionalSyncPlanFileDecisions"),
                        file_apply.index("If MsgBox(confirmationText"))
        self.assertLess(file_apply.index("If MsgBox(confirmationText"),
                        file_apply.index("BidirectionalSyncResolveFileDecisions"))

    def test_file_selection_clears_right_field_before_button_can_route(self):
        render = body("ShowECurrent", FORM, "Sub")
        selection = body("UpdateESelection", FORM, "Sub")
        self.assertIn("mLstRows.ListIndex = -1", render)
        self.assertIn("Not selectedField Is Nothing", selection)
        self.assertIn("selectedField Is Nothing", selection)
        self.assertIn("n = 1 And Not record Is Nothing And identitySafe", selection)

    def test_parent_changes_cannot_leave_unselected_or_held_child_code_stale(self):
        hierarchy = body("BidirectionalSyncBatchHierarchySafe")
        self.assertIn('"分類名１・分類コード１"', hierarchy)
        self.assertIn('"第1ガイド・分類コード２"', hierarchy)
        self.assertIn('"第2ガイド・分類コード３"', hierarchy)
        self.assertIn('"システム採用"', hierarchy)
        self.assertIn("BidirectionalSyncBatchEffectiveParent", hierarchy)
        self.assertIn("BidirectionalSyncBatchHierarchySafe(cases, choices", body("BidirectionalSyncPlanFileDecisions"))

    def test_failure_injection_occurs_after_second_group_and_before_state_commit(self):
        commit = body("BidirectionalSyncCommitFilePlan")
        self.assertIn("groupIndex = groupIndex + 1", commit)
        self.assertIn("failAfterGroup > 0 And groupIndex = failAfterGroup", commit)
        self.assertLess(commit.index("WriteStep2SourceValue"), commit.index("failAfterGroup > 0"))
        self.assertLess(commit.index("failAfterGroup > 0"), commit.index("wsConflict.Cells"))
        self.assertIn("BidirectionalSyncRestoreFileDecisions", body("BidirectionalSyncResolveFileDecisions"))

    def test_form_artifacts_match_canonical_and_are_cp932_compatible(self):
        for source in (MODULE, FORM):
            source.encode("cp932")
        for artifact, encoding in (("frmSyncReview.frm", "utf-8-sig"),
                                   ("frmSyncReview-native.frm", "cp932")):
            exported = (ROOT / artifact).read_bytes().decode(encoding)
            exported = exported.replace("\r\n", "\n")
            self.assertEqual(exported[exported.index("Option Explicit"):], FORM.replace("\r\n", "\n"))


class BulkTransactionModelTests(unittest.TestCase):
    """Synthetic transaction model only; NOT an Excel/VBA execution test."""

    @staticmethod
    def apply_atomic(rows, history, decisions, fail_after=None):
        snapshot = deepcopy((rows, history, decisions))
        try:
            for position, (row_id, field, value, choice) in enumerate(decisions["plan"], 1):
                if decisions.get((row_id, field)) in ("system", "excel"):
                    raise ValueError("finalized")
                if choice == "system":
                    rows[row_id][field] = value
                history.append((row_id, field, choice))
                decisions[(row_id, field)] = choice
                if fail_after == position:
                    raise RuntimeError("injected")
        except (RuntimeError, ValueError):
            rows.clear(); rows.update(snapshot[0])
            history[:] = snapshot[1]
            decisions.clear(); decisions.update(snapshot[2])
            return False
        return True

    def test_second_or_later_failure_restores_cells_history_and_memory(self):
        for fail_after in (2, 3):
            rows = {"2026:a": {"retention": "old", "title": "old", "note": "old"},
                    "2025:a": {"title": "prior"}, "2026:b": {"title": "other"}}
            history = [("2026:a", "earlier", "excel")]
            decisions = {("2026:a", "earlier"): "excel", "plan": [
                ("2026:a", "retention", "new", "system"),
                ("2026:a", "title", "new", "system"),
                ("2026:a", "note", "new", "system"),
            ]}
            before = deepcopy((rows, history, decisions))
            self.assertFalse(self.apply_atomic(rows, history, decisions, fail_after))
            self.assertEqual((rows, history, decisions), before)

    def test_mixed_hold_and_finalized_rejection_do_not_touch_other_years(self):
        rows = {"2026:a": {"title": "old", "note": "old"}, "2025:a": {"title": "prior"}}
        history = []
        decisions = {"plan": [("2026:a", "title", "new", "system"),
                              ("2026:a", "note", "ignored", "hold")]}
        self.assertTrue(self.apply_atomic(rows, history, decisions))
        self.assertEqual(rows["2026:a"], {"title": "new", "note": "old"})
        self.assertEqual(rows["2025:a"]["title"], "prior")
        decisions["plan"] = [("2026:a", "title", "again", "system")]
        before = deepcopy((rows, history, decisions))
        self.assertFalse(self.apply_atomic(rows, history, decisions))
        self.assertEqual((rows, history, decisions), before)


if __name__ == "__main__":
    unittest.main()
