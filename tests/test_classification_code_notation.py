"""Static contract tests for narrow, safe classification-code notation updates."""
from pathlib import Path
import re
import unittest

MODULE = Path(__file__).resolve().parents[1] / "vba" / "Module1.txt"
SOURCE = MODULE.read_text(encoding="utf-8-sig")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?ms)^Private (?:Sub|Function) {re.escape(name)}\b.*?^End (?:Sub|Function)\s*$",
        SOURCE,
    )
    if not match:
        raise AssertionError(f"missing VBA procedure: {name}")
    return match.group(0)


def digit_key(value):
    text = str(value)
    if not text or not text.isascii() or not text.isdigit():
        return None
    return text.lstrip("0") or "0"


def normalize_model(
    system_code,
    excel_code,
    *,
    exact_id,
    names_match,
    saved_conflict=False,
    saved_decision_unresolved=False,
    saved_snapshot_matches=False,
    saved_baseline_unknown=False,
    formula=False,
):
    """Specification model only; this does not execute VBA or Excel."""
    if isinstance(system_code, str) and isinstance(excel_code, str) and system_code == excel_code and not formula:
        return excel_code
    if not exact_id or not names_match or formula or not isinstance(system_code, str):
        return excel_code
    safe_saved_conflict = (
        saved_decision_unresolved and saved_snapshot_matches and saved_baseline_unknown
    )
    if saved_conflict and not safe_saved_conflict:
        return excel_code
    system_key = digit_key(system_code)
    excel_key = digit_key(excel_code)
    if system_key is not None and system_key == excel_key:
        return system_code
    return excel_code


class ClassificationCodeNotationTests(unittest.TestCase):
    def test_specific_notation_examples_and_idempotence(self):
        self.assertEqual(normalize_model("0000", "0000", exact_id=True, names_match=True), "0000")
        self.assertEqual(normalize_model("0000", 0, exact_id=True, names_match=True), "0000")
        self.assertEqual(normalize_model("090200", "090200", exact_id=True, names_match=True), "090200")
        once = normalize_model(
            "0000", 0, exact_id=True, names_match=True,
            saved_conflict=True, saved_decision_unresolved=True,
            saved_snapshot_matches=True, saved_baseline_unknown=True,
        )
        twice = normalize_model("0000", once, exact_id=True, names_match=True)
        self.assertEqual((once, twice), ("0000", "0000"))

    def test_unproven_identity_name_conflict_or_category_change_holds(self):
        self.assertEqual(normalize_model("0000", 0, exact_id=False, names_match=True), 0)
        self.assertEqual(normalize_model("0000", 0, exact_id=True, names_match=False), 0)
        self.assertEqual(normalize_model("0000", 0, exact_id=True, names_match=True, saved_conflict=True), 0)
        self.assertEqual(normalize_model("0001", 0, exact_id=True, names_match=True), 0)
        self.assertEqual(normalize_model("0000", 0, exact_id=True, names_match=True, formula=True), 0)
        self.assertEqual(normalize_model(0, 0, exact_id=True, names_match=True), 0)

    def test_only_an_exact_unresolved_initial_notation_snapshot_can_clear_a_saved_conflict(self):
        self.assertEqual(
            normalize_model(
                "0000", 0, exact_id=True, names_match=True,
                saved_conflict=True, saved_decision_unresolved=True,
                saved_snapshot_matches=True, saved_baseline_unknown=True,
            ),
            "0000",
        )
        for saved_state in (
            {"saved_decision_unresolved": False, "saved_snapshot_matches": True, "saved_baseline_unknown": True},
            {"saved_decision_unresolved": True, "saved_snapshot_matches": False, "saved_baseline_unknown": True},
            {"saved_decision_unresolved": True, "saved_snapshot_matches": True, "saved_baseline_unknown": False},
        ):
            self.assertEqual(
                normalize_model("0000", 0, exact_id=True, names_match=True, saved_conflict=True, **saved_state),
                0,
            )

        normalizer = procedure("BidirectionalSyncPrepareClassificationCodeGroup")
        saved_check = normalizer.index("BidirectionalSyncHasSavedConflictForGroup")
        exception_check = normalizer.index("BidirectionalSyncSavedConflictAllowsCodeNotation", saved_check)
        write = normalizer.index("BidirectionalSyncRecordGuideCodeSnapshot", exception_check)
        self.assertLess(saved_check, exception_check)
        self.assertLess(exception_check, write)

        saved_guard = procedure("BidirectionalSyncSavedConflictAllowsCodeNotation")
        for required in (
            "BidirectionalSyncConflictReviewCases",
            '"SavedConflictCasesByKey"',
            "For Each savedCase In cases",
            "BIDIRECTIONAL_SYNC_CONFLICT_UNRESOLVED",
            'savedCase("stableKey")',
            'If StrComp(CStr(savedCase("decision")), BIDIRECTIONAL_SYNC_CONFLICT_UNRESOLVED, vbBinaryCompare) <> 0 Then _',
            'savedHeader("baselineKnown")',
            'If CBool(savedHeader("baselineKnown")) Then Exit Function',
            'savedHeader("systemValue")',
            'savedHeader("excelValue")',
            "BidirectionalSyncDigitsOnlyCodeKey",
            "CanonicalHeaderName(destinationHeader)",
            "CanonicalHeaderName(headerName)",
            "If Not codeSeen Then Exit Function",
        ):
            self.assertIn(required, saved_guard)

    def test_real_content_changes_are_not_code_notation_equivalent(self):
        for system_value, excel_value in (
            ("継続", "5年"),
            ("1999", "1105"),
            ("土木", "工務"),
            ("2", "3"),
        ):
            self.assertEqual(
                normalize_model(system_value, excel_value, exact_id=True, names_match=True),
                excel_value,
            )

    def test_vba_normalizer_is_limited_to_classification_two_three_and_generic_ten(self):
        body = procedure("BidirectionalSyncPrepareClassificationCodeGroup")
        self.assertIn("第1ガイド・分類コード２", body)
        self.assertIn("第2ガイド・分類コード３", body)
        self.assertIn('Case "汎用項目１０"', body)
        self.assertIn("BidirectionalSyncCurrentMatchKind()", body)
        self.assertIn("正式ID一意一致", body)
        self.assertIn("BidirectionalSyncDigitsOnlyCodeKey", body)
        self.assertIn("BidirectionalSyncDestinationRowIdIsUnique", body)
        self.assertIn("BidirectionalSyncHasSavedConflictForGroup", body)
        self.assertIn("NormalizeValueForUpdateCompare(nameKey", body)
        self.assertIn("WriteStep2SourceValue excelCodeCell, codeKey, systemRaw", body)
        self.assertNotIn("WriteStep2SourceValue excelNameCell", body)

    def test_unsafe_code_difference_never_uses_the_csv_only_auto_apply_path(self):
        body = procedure("BidirectionalSyncApplyManagedGroup")
        csv_only = body.index("ElseIf (Not excelChanged) And csvChanged Then")
        hold = body.index('"コード表記確認保留"', csv_only)
        apply = body.index("For Each key In itemKeys", hold)
        self.assertLess(hold, apply)
        self.assertIn('"初期不一致"', body)
        self.assertIn('"競合保留"', body)
        self.assertIn('"システム手動編集待ち"', body)

    def test_code_text_import_and_generic_item_10_notation_scope(self):
        importer = procedure("CsvHeaderIsIdentifier")
        opener = procedure("CsvOpenTextAsWorkbook")
        writer = procedure("WriteStep2SourceValue")
        required = procedure("Step2HeaderRequiresSourceText")
        normalizer = procedure("BidirectionalSyncPrepareClassificationCodeGroup")
        for header in ("分類コード２", "分類コード３", "汎用項目１０"):
            self.assertIn(f'CanonicalHeaderName("{header}")', required)
        self.assertIn('CanonicalHeaderName("汎用項目１０")', importer)
        self.assertIn("Array(col, xlTextFormat)", opener)
        self.assertIn("FieldInfo:=fieldInfo", opener)
        self.assertIn('targetCell.NumberFormat = "@"', writer)
        self.assertIn("targetCell.Value2 = CStr(sourceValue)", writer)
        self.assertIn('Case "汎用項目１０"', normalizer)
        self.assertEqual(normalize_model("090200", "090200", exact_id=True, names_match=True), "090200")

    def test_new_code_write_has_a_scoped_value_and_number_format_rollback(self):
        start = procedure("BidirectionalSyncStartContext")
        capture = procedure("BidirectionalSyncRecordGuideCodeSnapshot")
        normalizer = procedure("BidirectionalSyncPrepareClassificationCodeGroup")
        rollback = procedure("BidirectionalSyncRollbackGuideCodeSnapshots")
        discard = procedure("BidirectionalSyncDiscardContext")
        self.assertIn('context.Add "GuideCodeSnapshots"', start)
        self.assertIn("targetCell.Formula", capture)
        self.assertIn("targetCell.NumberFormat", capture)
        self.assertIn("targetCell.Formula = saved(2)", rollback)
        self.assertIn("targetCell.NumberFormat = CStr(saved(3))", rollback)
        self.assertIn("Application.EnableEvents = False", normalizer)
        self.assertIn("Application.EnableEvents = oldEvents", normalizer)
        self.assertIn("Application.EnableEvents = False", rollback)
        self.assertIn("Application.EnableEvents = oldEvents", rollback)
        row_id_guard = procedure("BidirectionalSyncDestinationRowIdIsUnique")
        conflict_guard = procedure("BidirectionalSyncHasSavedConflictForGroup")
        conflict_loader = procedure("BidirectionalSyncLoadSavedConflictKeys")
        self.assertIn("actualRow <> expectedRow", row_id_guard)
        self.assertIn("SavedConflictKeys", conflict_guard)
        review_start = SOURCE.index("Public Function BidirectionalSyncConflictReviewCases(")
        review_cases = SOURCE[review_start:SOURCE.index("End Function", review_start)]
        self.assertIn("BidirectionalSyncConflictReviewCases", conflict_loader)
        self.assertIn("BidirectionalSyncConflictSheet(False, created)", review_cases)
        self.assertIn("BidirectionalSyncConflictDecisionIsValid", review_cases)
        self.assertLess(discard.index("BidirectionalSyncRollbackGuideCodeSnapshots"), discard.index("BidirectionalSyncDeleteNewBaselineSheet"))
        self.assertIn("ブックを保存せず終了", SOURCE)


if __name__ == "__main__":
    unittest.main()
