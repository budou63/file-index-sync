"""Safety contracts for the conservative E-review display-only stage."""
from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8-sig")
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")


def procedure(source: str, name: str, kind: str = "Function") -> str:
    markers = (f"Private {kind} {name}(", f"Public {kind} {name}(", f"{kind} {name}(")
    start = next((source.index(marker) for marker in markers if marker in source), None)
    if start is None:
        raise AssertionError(f"procedure not found: {name}")
    end = source.index(f"End {kind}", start)
    return source[start:end]


class SafeInternalScopeTests(unittest.TestCase):
    def test_normal_allowlist_is_staff_input_only_not_the_old_display_set(self):
        body = procedure(FORM, "EHeaderNormalDisplayCaption")
        for field in (
            "毎年度利用区分1:毎年作成 2:単年のみ", "原課", "年度（和暦）", "分類名１",
            "第1ガイド", "第2ガイド", "作成年度（和暦）", "タイトル", "保存期間",
            "媒体種別", "内容・取扱い", "備考",
        ):
            self.assertIn(f'CompareHeaderKey("{field}")', body)
        for non_input in ("色名", "2色名", "保存番号", "廃棄年月日", "同期状態"):
            self.assertNotIn(f'CompareHeaderKey("{non_input}")', body)
        self.assertNotIn("InStr(", body)

    def test_internal_mismatches_stay_unresolved_and_are_identifiable(self):
        show = procedure(FORM, "ShowECategory", "Sub")
        count = procedure(FORM, "ECountInternalCase", "Sub")
        self.assertIn("ERecordNeedsInternalInvestigation(record)", show)
        self.assertIn('"内部のみ・要調査: "', show)
        self.assertIn('"識別要確認: "', FORM)
        self.assertIn("internalMismatchCount = internalMismatchCount + 1", count)
        self.assertIn("unresolvedInternalCount = unresolvedInternalCount + 1", count)
        self.assertIn("safeNoReviewCount = 0", show)
        self.assertIn("自動解決はせず、差異を保持しています。", FORM)
        self.assertIn('If allMode Or Len(displayCaption) > 0 Then', procedure(FORM, "ShowECurrent", "Sub"))

    def test_unrecognized_internal_fields_fail_closed_in_normal_actions(self):
        safe = procedure(FORM, "ESelectedDecisionIsSafe")
        out_of_scope = procedure(FORM, "ECaseHasOutOfScopeFields")
        input_caption = procedure(FORM, "EFieldNormalDisplayCaption")
        self.assertIn('If EIsClassification1Group(item) Then', safe)
        self.assertIn('Module1.BidirectionalSyncClassification1DecisionSafe(item, decision, reasonText, previewText)', safe)
        self.assertIn('If Not EIsAllDisplayMode() And ECaseHasOutOfScopeFields(item) Then', safe)
        self.assertIn('全項目（管理コード含む）に切り替えてください', safe)
        self.assertIn("If Len(EFieldNormalDisplayCaption(caseItem, fieldItem)) = 0 Then", out_of_scope)
        self.assertIn("ECaseHasOutOfScopeFields = True", out_of_scope)
        self.assertIn('CStr(fieldItem("systemValue")), CStr(fieldItem("excelValue"))', input_caption)
        self.assertIn("Then Exit Function", input_caption)

    def test_step2_step4_and_unknown_code_trials_are_not_active(self):
        for trial in (
            "BidirectionalSyncRestoreBlankExcelValuesFromSystem",
            "BidirectionalSyncStep4WouldClearSystemValue",
            "BidirectionalSyncResolveConflictBatch",
            "BidirectionalSyncLastSafeNoReviewCount",
            "BidirectionalSyncConflictWouldBlankSystemValue",
        ):
            with self.subTest(trial=trial):
                self.assertNotIn(trial, MODULE)
                self.assertNotIn(trial, FORM)
        self.assertIn("Public Function BidirectionalSyncResolveConflict(", MODULE)

    def test_form_restore_artifacts_match_authoritative_source(self):
        FORM.encode("cp932")
        for name, encoding in (("frmSyncReview.frm", "utf-8"), ("frmSyncReview-native.frm", "cp932")):
            content = (ROOT / name).read_bytes().decode(encoding)
            self.assertEqual(content[content.index("Option Explicit"):].replace("\r\n", "\n"), FORM)


if __name__ == "__main__":
    unittest.main()
