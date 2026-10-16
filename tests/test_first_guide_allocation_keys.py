"""First-guide allocation key contract: decision cases and VBA call-path guards.

The Python decision table is not Excel/VBA execution. Source assertions bind it to
both the map builder and the actual allocator/nearest-neighbor call sites.
"""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def procedure(name: str) -> str:
    found = re.search(
        rf"(?:Public|Private)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE, re.DOTALL,
    )
    if found is None:
        raise AssertionError(f"Procedure not found: {name}")
    return found.group(0)


def master_key(level: str, name: str, parent: str = "") -> str:
    return f"{level}|{name}" + (f"|{parent}" if level == "分類名３" else "")


def allocate_level2(upper: int, lower: int, used: set[int]) -> int:
    if upper and lower:
        if upper >= lower - 1:
            return 0
        return next((candidate for candidate in range(upper + 1, lower) if candidate not in used), 0)
    if upper:
        candidate = upper + 10
        while candidate in used:
            candidate += 1
        return candidate
    return 0


class FirstGuideAllocationKeyTests(unittest.TestCase):
    def test_system_and_provisional_code_reuse_same_master_key(self) -> None:
        for state, code in (("システム", "205"), ("仮採番", "202")):
            with self.subTest(state=state):
                name_map = {master_key("分類名２", "道路計画"): code}
                self.assertEqual(name_map[master_key("分類名２", "道路計画")], code)
        builder = procedure("BuildCodeMasterNameCodeMap")
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn("lookupKey = BuildGuideCodeMasterNameKey(itemName, displayKey", builder)
        self.assertIn('lookupKey = BuildGuideCodeMasterNameKey("分類名２", guideKey)', allocator)
        self.assertIn("If Not code2ByName.Exists(lookupKey) Then", allocator)
        self.assertIn('WriteClassificationCodeValue wsEdit.Cells(r, code2Col), "分類コード２", CStr(code2ByName(lookupKey))', allocator)
        self.assertIn("code2ByName.Add lookupKey, codeValue", allocator)
        self.assertIn("pendingGuide2Rows.Add guideKey, guideRows", allocator)

    def test_nearest_uses_same_key_and_allocates_inside_bounds(self) -> None:
        nearest = procedure("FindNearestClassificationCode")
        self.assertIn('lookupKey = BuildGuideCodeMasterNameKey("分類名２", guideKey)', nearest)
        self.assertIn("If nameCodeMap.Exists(lookupKey) Then", nearest)
        self.assertIn("nameCodeMap(lookupKey)", nearest)
        name_map = {master_key("分類名２", "上側"): "205", master_key("分類名２", "下側"): "215"}
        upper, lower = (int(name_map[master_key("分類名２", n)]) for n in ("上側", "下側"))
        self.assertEqual(allocate_level2(upper, lower, {205, 206, 215}), 207)
        self.assertEqual(allocate_level2(upper, 0, {205, 215}), 216)

    def test_same_named_registration_rows_share_one_pending_code(self) -> None:
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn("If Not pendingGuide2Rows.Exists(guideKey) Then", allocator)
        self.assertIn("pendingGuide2Rows(guideKey).Add r", allocator)
        self.assertIn("For Each guideRow In guideRows", allocator)
        self.assertIn('WriteClassificationCodeValue wsEdit.Cells(CLng(guideRow), code2Col), "分類コード２", codeValue', allocator)
        self.assertLess(allocator.index("pendingGuide2Rows(guideKey).Add r"),
                        allocator.index('AppendCodeMasterRecord "分類名２"'))
        rows = ["新しいガイド", "新しいガイド"]
        pending = {}
        for row_number, name in enumerate(rows, start=2):
            pending.setdefault(name, []).append(row_number)
        self.assertEqual(pending, {"新しいガイド": [2, 3]})

    def test_no_eligible_upper_or_gap_stops_without_new_code(self) -> None:
        self.assertEqual(allocate_level2(0, 205, {205}), 0)
        self.assertEqual(allocate_level2(205, 206, {205, 206}), 0)
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn("If assignedCode <= 0 Then", allocator)
        self.assertIn("If Len(failureText) > 0 Then", allocator)
        self.assertIn("Err.Raise vbObjectError + 2601", allocator)

    def test_conflict_or_unresolved_state_stops_before_first_allocator_write(self) -> None:
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn('GetStep4GuideCodeStateIssue("第1ガイド", "分類名２", guideName)', allocator)
        self.assertIn('Err.Raise vbObjectError + 2602', allocator)
        self.assertLess(allocator.index('GetStep4GuideCodeStateIssue("第1ガイド"'),
                        allocator.index('WriteClassificationCodeValue wsEdit.Cells(r, code2Col), "分類コード２"'))
        self.assertLess(allocator.index('GetStep4GuideCodeStateIssue("第1ガイド"'),
                        allocator.index('AppendCodeMasterRecord "分類名２"'))
        self.assertIn("If ContainsMacroInsertRemark(wsEdit.Cells(r, bikoCol).Value) Then", allocator)
        self.assertIn('NormalizeCellValue("登録予定")', allocator)
        preflight = procedure("GetStep4GuideCodeStateIssue")
        self.assertIn('Left$(stateValue, 2) = "競合"', preflight)
        self.assertIn('Left$(stateValue, 3) = "要確認"', preflight)
        self.assertIn('If systemCodes.Count = 0 And provisionalCodes.Count > 1 Then', preflight)
        self.assertIn("If IsNumeric(codeValue) Then usedCodes(CStr(CLng(codeValue))) = True",
                      procedure("BuildCodeMasterNameCodeMap"))

    def test_all_first_guide_allocations_validate_before_any_code_write(self) -> None:
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        failure_check = allocator.index("If Len(failureText) > 0 Then")
        self.assertLess(failure_check, allocator.index('WriteClassificationCodeValue wsEdit.Cells(r, code2Col), "分類コード２"'))
        self.assertLess(failure_check, allocator.index('AppendCodeMasterRecord "分類名２"'))
        self.assertLess(failure_check, allocator.index('AppendCodeMasterRecord "分類名３"'))
        self.assertIn("plannedGuide2Codes.Add", allocator)
        self.assertIn("code2ByName.Add lookupKey, codeValue", allocator)

    def test_mixed_allocatable_and_unallocatable_guides_do_not_commit(self) -> None:
        # Synthetic decision table. A succeeds between 205 and 215; B has no
        # space between 215 and 216. The plan is discarded before any commit.
        used = {205, 215, 216}
        plan = {}
        for name, upper, lower in (("新規A", 205, 215), ("新規B", 215, 216)):
            code = allocate_level2(upper, lower, used)
            if code == 0:
                plan = None
                break
            plan[name] = code
            used.add(code)
        self.assertIsNone(plan)
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertLess(allocator.index("plannedGuide2Codes.Add"),
                        allocator.index("If Len(failureText) > 0 Then"))
        self.assertLess(allocator.index("If Len(failureText) > 0 Then"),
                        allocator.index('AppendCodeMasterRecord "分類名２"'))

    def test_parent_scoped_child_lookup_and_step4_preflight_unchanged(self) -> None:
        name_map = {master_key("分類名３", "共通", "205"): "31",
                    master_key("分類名３", "共通", "206"): "32"}
        self.assertEqual(name_map[master_key("分類名３", "共通", "205")], "31")
        self.assertEqual(name_map[master_key("分類名３", "共通", "206")], "32")
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn('guideKey = BuildGuideCodeMasterNameKey("分類名３", NormalizeClassificationGuideKey(guideName), parentCodeValue)', allocator)
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertLess(step4.index("PreflightStep4ConflictResolution"),
                        step4.index("新ファイル基準表_登録予定分類コード採番"))
        self.assertLess(step4.index("ValidateStep4GuideCodeStates"),
                        step4.index("新ファイル基準表_登録予定分類コード採番"))
        self.assertIn("保存せずExcelを終了してください", step4)


if __name__ == "__main__":
    unittest.main()
