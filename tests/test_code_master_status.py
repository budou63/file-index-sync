from __future__ import annotations

from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "Module1").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"{name} が見つかりません")
    return match.group(0)


def guide_key(value: str) -> str:
    return value.strip().replace("・", "/").replace("･", "/").replace("／", "/").replace(" ", "").casefold()


def reconcile(provisional: dict, system_rows: list[dict]) -> tuple[list[dict], str]:
    same_name = [r for r in system_rows if r["item"] == provisional["item"] and guide_key(r["name"]) == guide_key(provisional["name"])]
    for row in same_name:
        if row["code"] == provisional["code"]:
            return [], "昇格"
    if not same_name:
        return [{**provisional, "status": "仮採番"}], "維持"
    if len(same_name) == 1:
        system_code = same_name[0]["code"]
        return [{**provisional, "status": f"競合（システム={system_code}）"}], "競合"
    return [{**provisional, "status": "要確認（同名複数）"}], "要確認"


def resolve_code(item: str, display: str, rows: list[dict]) -> str:
    matches = [r for r in rows if r["item"] == item and guide_key(r["name"]) == guide_key(display)]
    for status in ("システム", "仮採番"):
        found = [r["code"] for r in matches if r["status"] == status]
        if found:
            return found[0]
    return ""


class CodeMasterStateBehaviorTests(unittest.TestCase):
    def test_system_rows_take_priority_and_provisional_is_fallback(self) -> None:
        rows = [
            {"item": "分類名３", "code": "811", "name": "資料", "status": "仮採番"},
            {"item": "分類名３", "code": "820", "name": "資料", "status": "システム"},
        ]
        self.assertEqual(resolve_code("分類名３", "資料", rows), "820")
        self.assertEqual(resolve_code("分類名３", "資料", rows[:1]), "811")

    def test_conflict_and_review_codes_are_never_selected(self) -> None:
        rows = [
            {"item": "分類名２", "code": "530", "name": "第一", "status": "競合（システム=531）"},
            {"item": "分類名３", "code": "811", "name": "資料", "status": "要確認（同名複数）"},
        ]
        self.assertEqual(resolve_code("分類名２", "第一", rows), "")
        self.assertEqual(resolve_code("分類名３", "資料", rows), "")

    def test_code_name_and_item_are_all_part_of_match(self) -> None:
        provisional = {"item": "分類名３", "code": "811", "name": "用地取得資料"}
        system = [
            {"item": "分類名３", "code": "640", "name": "用地取得資料"},
            {"item": "分類名３", "code": "796", "name": "用地取得資料"},
        ]
        rows, state = reconcile(provisional, system)
        self.assertEqual(state, "要確認")
        self.assertEqual(rows[0]["status"], "要確認（同名複数）")
        self.assertEqual(reconcile(provisional, [{"item": "分類名２", "code": "811", "name": "用地取得資料"}])[1], "維持")

    def test_exact_item_normalized_name_and_code_promotes_to_system(self) -> None:
        p = {"item": "分類名３", "code": "811", "name": "道路･河川"}
        s = [{"item": "分類名３", "code": "811", "name": "道路／河川"}]
        rows, state = reconcile(p, s)
        self.assertEqual((rows, state), ([], "昇格"))

    def test_single_different_code_is_conflict_and_multiple_names_need_review(self) -> None:
        p = {"item": "分類名３", "code": "811", "name": "資料"}
        rows, state = reconcile(p, [{"item": "分類名３", "code": "820", "name": "資料"}])
        self.assertEqual((state, rows[0]["status"]), ("競合", "競合（システム=820）"))
        many, state = reconcile(p, [
            {"item": "分類名３", "code": "640", "name": "資料"},
            {"item": "分類名３", "code": "796", "name": "資料"},
        ])
        self.assertEqual((state, many[0]["status"]), ("要確認", "要確認（同名複数）"))

    def test_step3_rebuild_preserves_pending_rows_and_writes_four_headers(self) -> None:
        codebook = procedure("手順3_コード管理CSVを作成する")
        self.assertIn('Array("項目名", "コード", "表示名", "状態")', codebook)
        self.assertIn("退避", codebook)
        self.assertIn("CapturePendingGuideCodeRows", codebook)
        self.assertIn("仮採番", procedure("CapturePendingGuideCodeRows"))
        self.assertIn("競合", procedure("CapturePendingGuideCodeRows"))
        self.assertIn("要確認", procedure("CapturePendingGuideCodeRows"))
        self.assertLess(codebook.index("CapturePendingGuideCodeRows"), codebook.index("wsDst.Cells.Clear"))
        self.assertIn("RestorePendingGuideCodeRows", codebook)
        self.assertIn("AutoFilter", codebook)
        self.assertIn('Columns("A:D").AutoFit', codebook)

    def test_legacy_unlabeled_codes_are_captured_for_step3_reconciliation(self) -> None:
        capture = procedure("CapturePendingGuideCodeRows")
        self.assertIn('If Len(stateValue) = 0 Then stateValue = "仮採番"', capture)
        self.assertIn("wsMaster.Cells(r, 2)", capture)
        self.assertIn("wsMaster.Cells(r, 3)", capture)

    def test_new_system_and_provisional_rows_write_explicit_states(self) -> None:
        codebook = procedure("手順3_コード管理CSVを作成する")
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn('Cells(outRow, 4).Value = "システム"', codebook)
        self.assertIn('Set firstGuideCodeMap = BuildCodeMasterNameCodeMap("分類名２", firstGuideUsedCodes, False)', step4)
        self.assertIn('Set secondGuideCodeMap = BuildCodeMasterNameCodeMap("分類名３", secondGuideUsedCodes, False)', step4)
        self.assertIn('AppendCodeMasterRecord "分類名２", codeValue, guideName, "仮採番"', allocator)
        self.assertIn('AppendCodeMasterRecord "分類名３", codeValue, guideName, "仮採番"', allocator)
        self.assertIn('BuildCodeMasterNameCodeMap("分類名２", usedCode2, True)', allocator)
        self.assertIn('BuildCodeMasterNameCodeMap("分類名３", usedCode3, True)', allocator)

    def test_append_helper_writes_status_to_column_d(self) -> None:
        append = procedure("AppendCodeMasterRecord")
        self.assertIn("Cells(nextRow, 4)", append)

    def test_step4_reuses_provisional_and_aborts_conflict_or_review(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn("ValidateStep4GuideCodeStates", step4)
        self.assertLess(step4.index("ValidateStep4GuideCodeStates"), step4.index("新ファイル基準表_登録予定分類コード採番"))
        self.assertIn('AppendCodeMasterRecord "分類名２", codeValue, guideName, "仮採番"', allocator)
        self.assertIn('AppendCodeMasterRecord "分類名３", codeValue, guideName, "仮採番"', allocator)
        self.assertIn('If code2ByName.Exists(guideKey) Then', allocator)

    def test_yellow_requires_system_state_only_for_both_guide_levels(self) -> None:
        reason = procedure("GetStep4GuideSeatCheckReason")
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn('BuildCodeMasterNameCodeMap("分類名２", firstGuideUsedCodes, False)', step4)
        self.assertIn('BuildCodeMasterNameCodeMap("分類名３", secondGuideUsedCodes, False)', step4)
        self.assertIn("firstGuideCodeMap", reason)
        self.assertIn("secondGuideCodeMap", reason)

    def test_get_code_helper_preserves_non_guide_behavior_and_prefers_system(self) -> None:
        lookup = procedure("GetCodeFromCodeMasterInRows")
        self.assertIn("GetCodeFromCodeMasterStatusAware", lookup)
        self.assertIn("GetCodeFromCodeMasterStatusAware", procedure("GetCodeFromCodeMasterStatusAware"))
        self.assertIn('desiredState = "システム"', procedure("GetCodeFromCodeMasterStatusAware"))
        self.assertIn('desiredState = "仮採番"', procedure("GetCodeFromCodeMasterStatusAware"))
        self.assertIn('NormalizeHeaderToken(kubunName)', lookup)
        self.assertIn("GetCodeFromCodeMasterInRows = \"\"", lookup)

    def test_status_cells_only_are_colored(self) -> None:
        color = procedure("ApplyCodeMasterStatusFormatting")
        self.assertIn("Cells(r, 4)", color)
        self.assertNotIn("Cells(r, 1).Interior", color)
        self.assertNotIn("Cells(r, 2).Interior", color)
        self.assertNotIn("Cells(r, 3).Interior", color)
    def test_status_matching_distinguishes_classification_name_2_and_3(self) -> None:
        for item in ("分類名２", "分類名３"):
            rows = [{"item": item, "code": "70", "name": "分類", "status": "仮採番"}]
            self.assertEqual(resolve_code(item, "分類", rows), "70")
            self.assertEqual(resolve_code("分類名１", "分類", rows), "")

    def test_different_code_does_not_promote_provisional(self) -> None:
        p = {"item": "分類名２", "code": "530", "name": "新ガイド"}
        rows, state = reconcile(p, [{"item": "分類名２", "code": "540", "name": "新ガイド"}])
        self.assertEqual(state, "競合")
        self.assertEqual(rows[0]["code"], "530")
        self.assertEqual(rows[0]["status"], "競合（システム=540）")

    def test_no_matching_system_row_keeps_pending_code(self) -> None:
        p = {"item": "分類名３", "code": "811", "name": "新資料"}
        rows, state = reconcile(p, [])
        self.assertEqual(state, "維持")
        self.assertEqual(rows[0]["code"], "811")
        self.assertEqual(rows[0]["status"], "仮採番")

    def test_exact_code_wins_even_if_another_system_code_has_same_name(self) -> None:
        p = {"item": "分類名３", "code": "805", "name": "用地資料"}
        rows, state = reconcile(p, [
            {"item": "分類名３", "code": "640", "name": "用地資料"},
            {"item": "分類名３", "code": "805", "name": "用地資料"},
        ])
        self.assertEqual((rows, state), ([], "昇格"))

    def test_item_mismatch_does_not_promote_a_pending_code(self) -> None:
        p = {"item": "分類名３", "code": "811", "name": "資料"}
        rows, state = reconcile(p, [{"item": "分類名２", "code": "811", "name": "資料"}])
        self.assertEqual((state, rows[0]["status"]), ("維持", "仮採番"))

    def test_rebuild_keeps_each_system_duplicate(self) -> None:
        codebook = procedure("手順3_コード管理CSVを作成する")
        self.assertIn('uniqueKey = pairDefs(i)(0) & "|" & codeValue & "|" & nameValue', codebook)
        self.assertIn('wsDst.Cells(outRow, 4).Value = "システム"', codebook)

    def test_step4_warning_displays_role_name_and_both_codes(self) -> None:
        validator = procedure("GetStep4GuideCodeStateIssue")
        self.assertIn("guideRole", validator)
        self.assertIn("displayName", validator)
        self.assertIn("codeValue", validator)
        self.assertIn("systemCodeList", validator)
        self.assertIn("競合", validator)
        self.assertIn("要確認", validator)

    def test_step4_validates_before_any_code_allocation(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertLess(step4.index("ValidateStep4GuideCodeStates"), step4.index("ConfirmAndMarkStep4NewRegistrationCandidates"))
        self.assertLess(step4.index("ValidateStep4GuideCodeStates"), step4.index("新ファイル基準表_登録予定分類コード採番"))

    def test_step4_status_validator_checks_only_new_registration_rows(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("ValidateStep4GuideCodeStates wsEdit, newRegistrationRowList", step4)
        self.assertNotIn("ValidateStep4GuideCodeStates wsEdit, targetRows", step4)

    def test_step3_reconcile_has_exact_zero_one_and_multiple_candidate_paths(self) -> None:
        restore = procedure("RestorePendingGuideCodeRows")
        self.assertIn('And CStr(wsMaster.Cells(r, 4).Value) = "システム"', restore)
        self.assertIn('And NormalizeClassificationGuideKey(CStr(wsMaster.Cells(r, 3).Value)) = displayKey', restore)
        self.assertIn('If systemCode = codeValue Then exactFound = True', restore)
        self.assertIn('Case 0', restore)
        self.assertIn('stateValue = "仮採番"', restore)
        self.assertIn('Case 1', restore)
        self.assertIn('stateValue = "競合（システム=" & systemCode & "）"', restore)
        self.assertIn('stateValue = "要確認（同名複数）"', restore)

    def test_step3_rejects_error_cells_before_clearing_code_master(self) -> None:
        codebook = procedure("手順3_コード管理CSVを作成する")
        self.assertLess(codebook.index("EnsureCodeMasterRangeHasNoErrors"), codebook.index("wsDst.Cells.Clear"))

    def test_state_paths_guard_excel_error_values(self) -> None:
        validator = procedure("ValidateStep4GuideCodeStates")
        self.assertIn('GetStep4GuideCodeStateIssue(CStr(guideKinds(i)), CStr(guideItems(i)), guideName)', validator)
        self.assertNotIn('guideName, "")', validator)
        self.assertIn("EnsureCodeMasterRangeHasNoErrors", procedure("CapturePendingGuideCodeRows"))
        self.assertIn("EnsureCodeMasterRangeHasNoErrors", procedure("BuildCodeMasterNameCodeMap"))
        self.assertIn("EnsureCodeMasterRangeHasNoErrors", procedure("GetCodeFromCodeMasterStatusAware"))

    def test_non_guide_lookups_keep_legacy_code_name_search(self) -> None:
        lookup = procedure("GetCodeFromCodeMasterInRows")
        self.assertIn('masterValues = wsMaster.Range(wsMaster.Cells(rowStart, 1), wsMaster.Cells(rowEnd, 3)).Value2', lookup)
        self.assertIn("NormalizeClassificationGuideKey", lookup)
        self.assertIn("GetCodeFromCodeMasterInRows = Trim$(CStr(masterValues(r, 2)))", lookup)

    def test_fixed_master_records_remain_in_step3(self) -> None:
        codebook = procedure("手順3_コード管理CSVを作成する")
        for label in ("毎年度利用区分", "年度暦年区分", "完結区分", "公開区分", "歴史資料区分", "廃棄区分", "履歴ファイル区分", "機密ファイル区分"):
            self.assertIn(f'"{label}"', codebook)

    def test_status_color_rule_does_not_repaint_data_columns(self) -> None:
        color = procedure("ApplyCodeMasterStatusFormatting")
        self.assertIn('RGB(255, 242, 204)', color)
        self.assertIn('RGB(244, 204, 204)', color)
        self.assertIn("statusCell.Interior.Pattern = xlNone", color)


if __name__ == "__main__":
    unittest.main()
