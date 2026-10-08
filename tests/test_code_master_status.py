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


def reconcile_final_state(row: dict, system_rows: list[dict]) -> dict:
    same_name = [
        system
        for system in system_rows
        if system["item"] == row["item"] and guide_key(system["name"]) == guide_key(row["name"])
    ]
    if any(system["code"].strip() == row["code"].strip() for system in same_name):
        state = "システム"
    elif not same_name:
        state = "仮採番"
    elif len({system["code"].strip() for system in same_name}) == 1:
        state = f"競合（システム={same_name[0]['code'].strip()}）"
    else:
        state = "要確認（同名複数）"
    return {**row, "status": state}


def reconcile_final_rows(rows: list[dict], system_rows: list[dict]) -> list[dict]:
    reconciled = [reconcile_final_state(row, system_rows) for row in rows]
    unique: dict[tuple[str, str, str], dict] = {}
    for row in reconciled:
        key = (row["item"], row["code"].strip(), guide_key(row["name"]))
        unique.setdefault(key, row)
    return list(unique.values())


def source_first_rebuild(old_rows: list[dict], source_rows: list[dict]) -> list[dict]:
    """Pure-Python contract model; VBA wiring is checked separately below."""
    exact: dict[tuple[str, str, str], dict] = {}
    for row in source_rows:
        item = row["item"].replace("分類名2", "分類名２").replace("分類名3", "分類名３")
        key = item, row["code"].strip(), guide_key(row["name"])
        exact.setdefault(key, {**row, "item": item, "status": "システム"})
    output = list(exact.values())
    for row in old_rows:
        item = row["item"].replace("分類名2", "分類名２").replace("分類名3", "分類名３")
        key = item, row["code"].strip(), guide_key(row["name"])
        if key not in exact:
            output.append(reconcile_final_state({**row, "item": item}, output[:len(exact)]))
    return output


def validate_source_keys(source_rows: list[dict], output: list[dict]) -> None:
    for source in source_rows:
        item = source["item"].replace("分類名2", "分類名２").replace("分類名3", "分類名３")
        matches = [row for row in output if row["item"] == item and row["code"].strip() == source["code"].strip()
                   and guide_key(row["name"]) == guide_key(source["name"])]
        if len(matches) != 1 or matches[0]["status"] != "システム":
            raise ValueError((item, source["code"], source["name"]))


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
        self.assertIn('Array("項目名", "コード", "表示名", "状態", "親分類コード")', codebook)
        self.assertIn("退避", codebook)
        self.assertIn("CapturePendingGuideCodeRows", codebook)
        self.assertIn("仮採番", procedure("CapturePendingGuideCodeRows"))
        self.assertIn("競合", procedure("CapturePendingGuideCodeRows"))
        self.assertIn("要確認", procedure("CapturePendingGuideCodeRows"))
        self.assertLess(codebook.index("CapturePendingGuideCodeRows"), codebook.index("wsDst.Cells.Clear"))
        self.assertIn("RestorePendingGuideCodeRows", codebook)
        self.assertIn("AutoFilter", codebook)
        self.assertIn('Columns("A:E").AutoFit', codebook)

    def test_legacy_unlabeled_codes_are_captured_for_step3_reconciliation(self) -> None:
        capture = procedure("CapturePendingGuideCodeRows")
        self.assertIn('If Len(stateValue) = 0 Then stateValue = "仮採番"', capture)
        self.assertIn("wsMaster.Cells(r, 2)", capture)
        self.assertIn("wsMaster.Cells(r, 3)", capture)

    def test_new_system_and_provisional_rows_write_explicit_states(self) -> None:
        codebook = procedure("手順3_コード管理CSVを作成する")
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn('Cells(outRow, 4).Value = "システム"', procedure("BuildSystemGuideCodeMasterFromMokuroku"))
        self.assertIn('Set firstGuideCodeMap = BuildCodeMasterNameCodeMap("分類名２", firstGuideUsedCodes, False)', step4)
        self.assertIn('Set secondGuideCodeMap = BuildCodeMasterNameCodeMap("分類名３", secondGuideUsedCodes, False, True)', step4)
        self.assertIn('AppendCodeMasterRecord "分類名２", codeValue, guideName, "仮採番"', allocator)
        self.assertIn('AppendCodeMasterRecord "分類名３", codeValue, guideName, "仮採番"', allocator)
        self.assertIn('BuildCodeMasterNameCodeMap("分類名２", usedCode2, True)', allocator)
        self.assertIn('BuildCodeMasterNameCodeMap("分類名３", usedCode3, True, True)', allocator)

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
        self.assertIn('BuildCodeMasterNameCodeMap("分類名３", secondGuideUsedCodes, False, True)', step4)
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


class FinalGuideStateReconciliationTests(unittest.TestCase):
    def test_wholesale_provisional_workplace_rows_become_system_first(self) -> None:
        names = [("0", "全庁共通"), ("10", "計画"), ("20", "人事・議会・監査")]
        source = [{"item": "分類名２", "code": code, "name": name} for code, name in names]
        old = [{**row, "item": "分類名2", "status": "仮採番"} for row in source]
        result = source_first_rebuild(old, source)
        self.assertEqual([(row["code"], row["status"]) for row in result],
                         [(code, "システム") for code, _ in names])
        validate_source_keys(source, result)

    def test_wholesale_provisional_classification3_and_pending_cases(self) -> None:
        source = [{"item": "分類名３", "code": "7", "name": "全般"},
                  {"item": "分類名３", "code": "8", "name": "共通"}]
        old = [{**row, "status": "仮採番"} for row in source]
        old += [{"item": "分類名３", "code": "999", "name": "新規座", "status": "仮採番"},
                {"item": "分類名３", "code": "998", "name": "全般", "status": "仮採番"}]
        result = source_first_rebuild(old, source)
        self.assertEqual([row["status"] for row in result],
                         ["システム", "システム", "仮採番", "競合（システム=7）"])
        validate_source_keys(source, result)

    def test_postcondition_catches_formal_key_provisional_and_duplicate(self) -> None:
        source = [{"item": "分類名２", "code": "10", "name": "計画"}]
        with self.assertRaises(ValueError):
            validate_source_keys(source, [{**source[0], "status": "仮採番"}])
        with self.assertRaises(ValueError):
            validate_source_keys(source, [{**source[0], "status": "システム"}] * 2)

    def test_step3_builds_authoritative_guides_before_merging_pending_and_checks_postcondition(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertLess(step3.index("BuildSystemGuideCodeMasterFromMokuroku"), step3.index("RestorePendingGuideCodeRows"))
        self.assertLess(step3.index("RestorePendingGuideCodeRows"), step3.index("ValidateGuideCodeMasterPostcondition"))
        self.assertLess(step3.index("ValidateGuideCodeMasterPostcondition"), step3.index('MsgBox "コード管理CSV の作成が完了しました。"'))

    def test_step3_builds_offsheet_validates_committed_sheet_and_restores_on_failure(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertIn("Set buildBook = Workbooks.Add(xlWBATWorksheet)", step3)
        self.assertIn("Set wsBuild = buildBook.Worksheets(1)", step3)
        self.assertIn("BuildSystemGuideCodeMasterFromMokuroku wsSrc, wsBuild", step3)
        self.assertIn("RestorePendingGuideCodeRows wsBuild", step3)
        self.assertLess(step3.index("ValidateGuideCodeMasterPostcondition wsBuild"), step3.index("wsDst.Cells.Clear"))
        self.assertIn("ValidateGuideCodeMasterPostcondition wsDst", step3)
        self.assertIn("RestoreCodeMasterFromBackup wsDst, wsBackup", step3)
        restore = procedure("RestoreCodeMasterFromBackup")
        self.assertIn("wsBackup.UsedRange.Copy Destination:=wsDst.Range", restore)
        self.assertIn("wsDst.Columns(col).ColumnWidth", restore)
        self.assertIn("wsBackup.AutoFilter.Filters(filterIndex)", restore)
        self.assertIn("hasCriteria2 = (Err.Number = 0)", restore)
        self.assertIn("If hasCriteria2 Then", restore)
        self.assertIn("keepBackupOpen = True", step3)
        self.assertIn("Not keepBackupOpen Then buildBook.Close", step3)
        self.assertIn("On Error Resume Next", step3)
        self.assertLess(step3.index("wsDst.Copy After:=wsBuild"), step3.index("commitStarted = True"))
        self.assertLess(step3.index("commitStarted = True"), step3.index("wsDst.Cells.Clear"))

    def test_source_first_merge_uses_normalized_identity_and_preserves_distinct_codes(self) -> None:
        build = procedure("BuildSystemGuideCodeMasterFromMokuroku")
        restore = procedure("RestorePendingGuideCodeRows")
        for text in ("NormalizeClassificationGuideKey", 'Cells(outRow, 4).Value = "システム"', 'exactMap.Add'):
            self.assertIn(text, build)
        self.assertIn("If exactMap.Exists(exactKey) Then GoTo ContinuePendingRow", restore)
        self.assertIn('stateValue = "競合（システム="', restore)
        self.assertIn('stateValue = "要確認（同名複数）"', restore)

    def test_postcondition_checks_keys_status_and_duplicates_and_fails_closed(self) -> None:
        check = procedure("ValidateGuideCodeMasterPostcondition")
        for text in ('If stateValue = "システム"', 'If count <> 1', 'Not systemCounts.Exists', 'Err.Raise', '分類名２', '分類名３'):
            self.assertIn(text, check)

    def test_final_reconciliation_promotes_exact_classification2_source_match(self) -> None:
        row = {"item": "分類名２", "code": "490", "name": "河川改修工事", "status": "仮採番"}
        source = [{"item": "分類名２", "code": "490", "name": "河川改修工事"}]
        result = reconcile_final_state(row, source)
        self.assertEqual(result["status"], "システム")

    def test_final_reconciliation_promotes_exact_classification3_source_match(self) -> None:
        row = {"item": "分類名３", "code": "490", "name": "河川改修工事", "status": "仮採番"}
        source = [{"item": "分類名３", "code": "490", "name": "河川改修工事"}]
        self.assertEqual(reconcile_final_state(row, source)["status"], "システム")

    def test_final_reconciliation_keeps_unknown_code_provisional(self) -> None:
        row = {"item": "分類名２", "code": "999", "name": "新規座", "status": "競合（旧）"}
        self.assertEqual(reconcile_final_state(row, [])["status"], "仮採番")

    def test_final_reconciliation_marks_single_different_system_code_as_conflict(self) -> None:
        row = {"item": "分類名２", "code": "999", "name": "河川改修工事", "status": "仮採番"}
        source = [{"item": "分類名２", "code": "490", "name": "河川改修工事"}]
        self.assertEqual(reconcile_final_state(row, source)["status"], "競合（システム=490）")

    def test_final_reconciliation_requires_review_for_multiple_system_codes(self) -> None:
        row = {"item": "分類名２", "code": "999", "name": "河川改修工事", "status": "仮採番"}
        source = [
            {"item": "分類名２", "code": "490", "name": "河川改修工事"},
            {"item": "分類名２", "code": "491", "name": "河川改修工事"},
        ]
        self.assertEqual(reconcile_final_state(row, source)["status"], "要確認（同名複数）")

    def test_exact_system_code_wins_among_multiple_same_name_candidates(self) -> None:
        row = {"item": "分類名２", "code": "490", "name": "河川改修工事", "status": "仮採番"}
        source = [
            {"item": "分類名２", "code": "490", "name": "河川改修工事"},
            {"item": "分類名２", "code": "491", "name": "河川改修工事"},
        ]
        self.assertEqual(reconcile_final_state(row, source)["status"], "システム")

    def test_final_reconciliation_collapses_only_identical_code_name_rows(self) -> None:
        rows = [
            {"item": "分類名２", "code": "490", "name": "河川改修工事", "status": "システム"},
            {"item": "分類名２", "code": "490", "name": "河川改修工事", "status": "仮採番"},
            {"item": "分類名２", "code": "998", "name": "河川改修工事", "status": "仮採番"},
            {"item": "分類名２", "code": "999", "name": "河川改修工事", "status": "仮採番"},
        ]
        source = [{"item": "分類名２", "code": "490", "name": "河川改修工事"}]
        result = reconcile_final_rows(rows, source)
        self.assertEqual([(r["code"], r["status"]) for r in result], [
            ("490", "システム"),
            ("998", "競合（システム=490）"),
            ("999", "競合（システム=490）"),
        ])

    def test_final_reconciliation_does_not_merge_distinct_unregistered_codes(self) -> None:
        rows = [
            {"item": "分類名２", "code": "998", "name": "未登録座", "status": "仮採番"},
            {"item": "分類名２", "code": "999", "name": "未登録座", "status": "仮採番"},
        ]
        result = reconcile_final_rows(rows, [])
        self.assertEqual([(r["code"], r["status"]) for r in result], [
            ("998", "仮採番"),
            ("999", "仮採番"),
        ])

    def test_step3_final_reconciliation_validates_required_headers_before_clearing(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertIn("BuildSystemGuideCodeMasterFromMokuroku", step3)
        self.assertLess(step3.index("ValidateGuideCodeSourceHeaders"), step3.index("wsDst.Cells.Clear"))

    def test_final_reconciliation_handles_all_required_source_pairs_and_state_paths(self) -> None:
        reconciler = procedure("RestorePendingGuideCodeRows") + procedure("BuildSystemGuideCodeMasterFromMokuroku")
        for expected in (
            'exactMap.Exists',
            'candidateCodes.Count',
            'Cells(outRow, 4).Value = "システム"',
            'stateValue = "仮採番"',
            '競合（システム=',
            '要確認（同名複数）',
            'NormalizeClassificationGuideKey',
        ):
            self.assertIn(expected, reconciler)

    def test_required_source_header_validation_rejects_each_missing_pair(self) -> None:
        validate = procedure("ValidateGuideCodeSourceHeaders")
        for expected in (
            'Array("分類コード２")',
            'Array("分類名２")',
            'Array("分類コード３")',
            'Array("分類名３")',
            'Err.Raise vbObjectError + 2604',
            'Err.Raise vbObjectError + 2605',
            '状態を確定できません。',
        ):
            self.assertIn(expected, validate)
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
        self.assertIn('wsMaster.Cells(outRow, 4).Value = "システム"', procedure("BuildSystemGuideCodeMasterFromMokuroku"))

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

    def test_step4_validates_conflicts_for_every_emitted_target_row(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertIn("ValidateStep4GuideCodeStates wsEdit, targetRows", step4)
        self.assertNotIn("ValidateStep4GuideCodeStates wsEdit, newRegistrationRowList", step4)

    def test_step3_reconcile_has_exact_zero_one_and_multiple_candidate_paths(self) -> None:
        restore = procedure("RestorePendingGuideCodeRows")
        self.assertIn('If exactMap.Exists(exactKey) Then GoTo ContinuePendingRow', restore)
        self.assertIn('If systemCodesByName.Exists(nameKey)', restore)
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
        self.assertIn('GetStep4GuideCodeStateIssue("第2ガイド", "分類名３", secondGuideName, parentCodeValue)', validator)
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
