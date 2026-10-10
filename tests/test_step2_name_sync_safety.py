from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "Module1").read_text(encoding="utf-8")


def body(name: str) -> str:
    match = re.search(rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)", SOURCE, re.S)
    if not match:
        raise AssertionError(name)
    return match.group(0)


class Step2NameSyncSafetyTests(unittest.TestCase):
    def test_name_key_uses_created_year_without_fiscal_fallback(self):
        code = body("BuildStep2NameSyncKey")
        self.assertIn("NormalizeYearValue(ws.Cells(rowNo, createdCol).Value)", code)
        self.assertIn('GetStep2HeaderCandidates("作成年度（和暦）")', code)
        self.assertNotIn("GetNormalizedYearFromRow", code)
        self.assertIn('CStr(yearValue) & "|"', code)
        self.assertIn("srcRowYearValue = GetNormalizedYearFromRow", body("手順2_新ファイル基準表を作成する"))
        self.assertIn("BuildFolderNameSyncKeyFromSheetRow", body("手順4_差分インポートCSVを作成する"))

    def test_name_candidates_are_only_registration_planned_without_formal_file_id(self):
        code = body("BuildDestinationNameSyncMap")
        self.assertIn('GetSyncStatusValue(wsDst, dstRow, syncCol) <> "登録予定"', code)
        self.assertIn("Len(GetTrimmedCell(wsDst, dstRow, fileIdCol)) > 0", code)
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertLess(step2.index("If existingRowMap.Exists(key) Then"), step2.index("Set sourceCandidates ="))
        self.assertIn("sourceCandidates.Count = 1 And destCandidates.Count = 1", step2)
        self.assertIn("Step2NameCandidateIdentityMatches", step2)
        self.assertIn("GoTo ContinueNextSourceRow", step2)

    def test_absent_name_data_and_incomplete_year_do_not_force_append(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn('If nameSyncKey = "" Then', step2)
        self.assertLess(step2.index('If nameSyncKey = "" Then'), step2.index("AppendRowFromSourceByHeader wsSrc"))
        self.assertIn("Step2MapRows(destMissingYearMap,", step2)
        self.assertIn("Step2ConfirmMissingSyncStatus", step2)

    def test_missing_status_requires_fiscal_year_three_names_and_absent_id(self):
        code = body("Step2ConfirmMissingSyncStatus")
        for part in ("GetNormalizedYearFromRow", "GetSyncStatusValue", "fileIdCol", "c2", "c3", "titleCol", "vbDefaultButton2", "SafeAddComment"):
            self.assertIn(part, code)
        self.assertNotIn('= "登録予定"', code)

    def test_ambiguity_comment_has_bounded_examples_counts_and_identity(self):
        code = body("BuildStep2AmbiguousComment")
        detail = body("Step2CandidateDetails")
        for part in ("目録CSV側・Excel側の双方が複数", "目録CSV側が複数", "Excel側が複数", "作成年度（和暦）", "第1ガイド", "第2ガイド", "個別フォルダー", "自動同期せず"):
            self.assertIn(part, code)
        for part in ("rows.Count", "Application.Min(rows.Count, 3)", "ファイルID", "同期状態", "フォルダー"):
            self.assertIn(part, detail)
        self.assertIn("AppendAmbiguousNameSyncRemark", body("手順2_新ファイル基準表を作成する"))

    def test_blank_status_name_collision_is_held_even_after_warning_continues(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("Step2MapRows(destBlankStatusMap,", step2)
        self.assertLess(step2.index("Step2MapRows(destBlankStatusMap,"), step2.index("AppendRowFromSourceByHeader wsSrc"))
        self.assertIn("同期状態未入力・作成年度不足の候補を確認", step2)
        self.assertIn("vbDefaultButton2", body("Step2ConfirmMissingSyncStatus"))

    def test_hidden_destination_header_and_reorder_are_position_safe(self):
        for name in ("BuildExistingRowMap", "ProcessPendingSyncStatusRows", "UpdateSyncStatusColumn", "BuildColumnOrderMap", "ApplySerialDiffMarksByKey"):
            self.assertIn("GetLastHeaderCol", body(name))
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("If wsDst.FilterMode Then", step2)
        self.assertIn("If shouldReorderColumns Then Err.Raise vbObjectError + 2612", step2)
        self.assertLess(step2.index("vbObjectError + 2612"), step2.index("BuildExistingRowMap wsDst"))

    def test_incomplete_candidates_are_checked_before_a_unique_planned_match(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        early = step2.index("Set blockedRows = Step2MapRows(destMissingYearMap")
        self.assertLess(early, step2.index("If sourceCandidates.Count = 1 And destCandidates.Count = 1 And ambiguityNote"))
        self.assertIn("destBlankMissingYearMap", step2)
        self.assertIn('visualDiffMap("THREE|"', step2)
        hold = body("BuildStep2HoldMaps")
        self.assertIn("AddRowToMultiValueMap blankMissingYearMap", hold)

    def test_name_only_diff_comment_never_marks_a_synced_or_deleted_id_row(self):
        marks = body("ApplySerialDiffMarksByKey")
        self.assertIn('statusValue = "登録予定" Or statusValue = ""', marks)
        self.assertIn("Len(GetTrimmedCell(ws, rowNo, fileIdCol)) = 0", marks)
        self.assertIn('visualDiffMap.Exists("THREE|" & threeKey)', marks)

    def test_hidden_rows_are_counted_in_step2_bounds(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("srcLastRow = Step2GetLastDataRow(wsSrc)", step2)
        self.assertIn("dstLastRow = Step2GetLastDataRow(wsDst)", step2)
        self.assertIn("CsvGetUnfilteredImportBounds", body("Step2GetLastDataRow"))


if __name__ == "__main__":
    unittest.main()
