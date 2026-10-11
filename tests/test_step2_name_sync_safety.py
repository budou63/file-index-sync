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
        self.assertLess(step2.index("If existingRowMap.Exists(key) Then"), step2.index("If nameSyncKey <> \"\" And sourceNameSyncMap.Exists(nameSyncKey)"))
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


    def test_exact_three_part_key_remains_first_and_unique_id_fallback_is_strict(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        index = body("BuildStep2IdentityIndex")
        zero_rule = body("Step2FileNumbersDifferOnlyByLeadingZeros")
        self.assertLess(step2.index("If Not existingRowMap.Exists(key) And Len(sourceFileIdValue) > 0 And destIdRows.Count = 1 Then"), step2.index("If existingRowMap.Exists(key) Then"))
        self.assertIn("sourceIdRows.Count > 1 Or destIdRows.Count > 1", step2)
        self.assertIn("sourcePairRows.Count <> 1 Or destPairRows.Count <> 1", step2)
        self.assertIn("GetTrimmedCell(wsDst, targetRow, destBarcodeCol) <> sourceBarcodeValue", step2)
        self.assertIn("Step2FileNumbersDifferOnlyByLeadingZeros", step2)
        self.assertIn("fileId & " + chr(34) + "|" + chr(34) + " & fileNo & " + chr(34) + "|" + chr(34) + " & barcodeNo", body("BuildFileKey"))
        self.assertIn("Step2FileNumberPaddingKey", index)
        self.assertNotIn("Step2FileNumberPaddingKey(fileId)", index)
        self.assertNotIn("Step2FileNumberPaddingKey(barcode)", index)
        self.assertIn("StrComp(sourceValue, destinationValue, vbBinaryCompare) = 0", zero_rule)

    def test_full_file_id_duplicates_and_partial_identity_collisions_are_held(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertLess(step2.index("sourceIdRows.Count > 1 Or destIdRows.Count > 1"), step2.index("If existingRowMap.Exists(key) Then"))
        self.assertIn("正式ファイルIDがCSV側またはExcel側で複数行", step2)
        self.assertIn("destBarcodeMap", step2)
        self.assertIn("destFileNoPaddingMap", step2)
        self.assertIn("RecordStep2IdentityHold", step2)
        self.assertIn("別の対象年度", step2)

    def test_name_sync_is_after_id_identity_and_duplicate_registered_rows_stay_unassigned(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        name_review = body("RecordStep2NameDuplicateReview")
        self.assertLess(step2.index("If existingRowMap.Exists(key) Then"), step2.index("If nameSyncKey <> " + chr(34) + chr(34) + " And sourceNameSyncMap.Exists(nameSyncKey)"))
        self.assertIn("RecordStep2NameDuplicateReview", step2)
        duplicate_guard = step2[step2.index("If sourceIdRows.Count > 1 Or destIdRows.Count > 1"):step2.index("ElseIf sourceExactRows.Count > 1")]
        self.assertIn("Not reviewedNameKeyMap.Exists(nameSyncKey)", duplicate_guard)
        self.assertIn("AppendAmbiguousNameSyncRemark", name_review)
        self.assertIn("登録予定行にはIDを付与せず", name_review)

    def test_file_number_updates_and_new_rows_preserve_source_text(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        writer = body("WriteStep2SourceValue")
        update = body("UpdateRowFromSourceByHeader")
        self.assertIn("targetCell.NumberFormat = " + chr(34) + "@" + chr(34), writer)
        self.assertIn("targetCell.Value2 = CStr(sourceValue)", writer)
        self.assertIn("WriteStep2SourceValue targetCell", update)
        self.assertIn("WriteStep2SourceValue wsDst.Cells", body("AppendRowFromSourceByHeader"))
        self.assertIn("preserveExactIdentity", update)
        self.assertIn("srcToDstMap, colorDiffDetectedCount, True", step2)
        self.assertIn("CanonicalHeaderName(" + chr(34) + "ファイルID" + chr(34) + ")", update)
        self.assertIn("CanonicalHeaderName(" + chr(34) + "バーコード番号" + chr(34) + ")", update)

    def test_new_rows_are_blocked_for_existing_four_part_names_or_identifier_collisions(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertLess(step2.index("destBarcodeMap, sourceBarcodeValue"), step2.index("AppendRowFromSourceByHeader wsSrc"))
        self.assertLess(step2.index("destAllNameMap.Exists(nameSyncKey)"), step2.index("AppendRowFromSourceByHeader wsSrc"))
        self.assertIn("補助ID照合件数", step2)
        new_row_path = step2[step2.index("If Len(sourceBarcodeValue) > 0 Then", step2.index("If existingRowMap.Exists(key) Then")):step2.index("AppendRowFromSourceByHeader wsSrc")]
        self.assertIn("Step2MapRows(destBarcodeMap, sourceBarcodeValue)", new_row_path)
        self.assertIn("Step2MapRows(destFileNoPaddingMap, paddingKey)", new_row_path)
        self.assertIn("Step2MapRows(destFileNoMap, sourceFileNoValue)", new_row_path)
        self.assertIn("If collisionRows.Count > 0 Then", new_row_path)
        self.assertIn("RecordStep2IdentityHold wsSrc", new_row_path)
        self.assertLess(new_row_path.index("If collisionRows.Count > 0 Then"), new_row_path.index("GoTo ContinueNextSourceRow", new_row_path.index("If collisionRows.Count > 0 Then")))

    def test_csv_import_already_protects_file_number_as_text(self):
        code = body("CsvHeaderIsIdentifier")
        self.assertIn("InStr(1, normalized, " + chr(34) + "番号" + chr(34) + ", vbTextCompare)", code)
        importer = body("CsvOpenTextAsWorkbook")
        self.assertIn("fieldInfo(i) = Array(col, xlTextFormat)", importer)
        self.assertIn("FieldInfo:=fieldInfo", importer)




    def test_missing_formal_file_id_holds_before_name_sync_and_append(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        guard = step2.index("If Len(sourceFileIdValue) = 0 Then")
        self.assertLess(guard, step2.index("If nameSyncKey <> " + chr(34) + chr(34) + " And sourceNameSyncMap.Exists"))
        self.assertIn("名称同期・新規追加を保留しました。", step2[guard:])

    def test_name_sync_requires_unique_complete_name_and_safe_cross_year_distinction(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("blockedRows.Count <> destCandidates.Count", step2)
        self.assertIn("Not Step2CrossYearNameCandidatesAreDistinct", step2)
        self.assertIn("Step2RowsForTargetYear", step2)
        self.assertIn("Step2HasOtherRows(collisionRows, targetRow)", step2)
        self.assertIn("Private Function Step2HasOtherRows", SOURCE)
        cross_year = body("Step2CrossYearNameCandidatesAreDistinct")
        for part in ("dstYearValue <= 0", "Len(sourceId) = 0", "Len(sourceFileNo) = 0", "Len(sourceBarcode) = 0", "StrComp(sourceId, destId", "StrComp(sourceBarcode, destBarcode", "sourceFileNoKey", "destFileNoKey"):
            self.assertIn(part, cross_year)

    def test_alphanumeric_file_number_conflicts_use_exact_text_map(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("Step2MapRows(destFileNoMap, sourceFileNoValue)", step2)
        self.assertIn("Step2MapRows(sourceFileNoMap, sourceFileNoValue)", step2)
        self.assertIn("Set collisionRows = Step2MapRows(destFileNoMap, sourceFileNoValue)", step2)

    def test_supplemental_file_number_mapping_is_not_used_for_other_identifiers(self):
        padding = body("Step2FileNumberPaddingKey")
        self.assertIn("If ch < " + chr(34) + "0" + chr(34) + " Or ch > " + chr(34) + "9" + chr(34), padding)
        self.assertIn("Step2FileNumbersDifferOnlyByLeadingZeros", SOURCE)
        self.assertNotIn("NormalizeIdentifierByRemovingZeros", SOURCE)


    def test_supplemental_match_requires_unique_secondary_identifiers_on_both_sides(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("sourceBarcodeRows.Count <> 1 Or destBarcodeRows.Count <> 1", step2)
        self.assertIn("CLng(destBarcodeRows.Item(1)) <> targetRow", step2)
        self.assertIn("sourceFileNoRows.Count <> 1 Or destFileNoRows.Count <> 1", step2)
        self.assertIn("CLng(destFileNoRows.Item(1)) <> targetRow", step2)
        self.assertIn("Step2MapRows(sourceFileNoPaddingMap, paddingKey)", step2)
        self.assertIn("Step2MapRows(destFileNoPaddingMap, paddingKey)", step2)

    def test_supplemental_match_keeps_pending_deletion_present(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        start = step2.index("ElseIf beforeSyncStatus = \"削除予定\" Then")
        branch = step2[start:step2.index("End If", start)]
        self.assertIn("matchedKey = BuildFileKeyFromSheetRow", branch)
        self.assertIn("sourceKeyMap.Add matchedKey, True", branch)
        self.assertIn("GoTo ContinueNextSourceRow", branch)

    def test_identity_hold_lists_all_candidates_but_marks_only_target_year(self):
        hold = body("RecordStep2IdentityHold")
        self.assertIn("Optional ByVal targetYearValue As Long = 0", hold)
        self.assertIn("BuildStep2AmbiguousComment(wsSrc, sourceRows, srcLastCol, wsDst, destinationRows", hold)
        self.assertIn("GetNormalizedYearFromRow(wsDst, CLng(rowValue), warekiYearCol, westernYearCol) <> targetYearValue", hold)
        self.assertLess(hold.index("GetNormalizedYearFromRow(wsDst, CLng(rowValue)"), hold.index("AppendRemarkOnce wsDst"))
        self.assertIn("visualDiffMap(key) = " + chr(34) + "判定困難" + chr(34), hold)

    def test_fiscal_year_aliases_are_used_for_identity_and_candidate_maps(self):
        aliases = "Array(\"年度（和暦）\", \"年度(和暦)\")"
        for name in ("BuildStep2IdentityIndex", "BuildDestinationAllNameMap", "RecordStep2IdentityHold", "ApplySerialDiffMarksByKey"):
            self.assertIn(aliases, body(name))

    def test_name_duplicate_comment_includes_the_id_row_and_planned_row(self):
        review = body("RecordStep2NameDuplicateReview")
        self.assertIn("ByVal idMatchedRows As Collection", review)
        self.assertIn("candidateRows.Add rowValue", review)
        self.assertIn("For Each rowValue In idMatchedRows", review)
        self.assertIn("For Each rowValue In plannedRows", review)
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("Set destCandidates = allNameCandidates", step2)


    def test_exact_key_candidate_checks_all_identifier_uniqueness_before_update(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        id_guard = step2.index("If sourceIdRows.Count > 1 Or destIdRows.Count > 1 Then")
        full_key_guard = step2.index("ElseIf sourceExactRows.Count > 1 Or destExactRows.Count > 1 Then")
        secondary_guard = step2.index("If destExactRows.Count = 1 And (Len(sourceFileIdValue) = 0")
        supplemental = step2.index("If Not existingRowMap.Exists(key) And Len(sourceFileIdValue) > 0")
        update = step2.index("If existingRowMap.Exists(key) Then")
        self.assertLess(id_guard, full_key_guard)
        self.assertLess(full_key_guard, secondary_guard)
        self.assertLess(secondary_guard, supplemental)
        self.assertLess(supplemental, update)
        for part in ("sourceBarcodeRows.Count <> 1", "destBarcodeRows.Count <> 1", "sourceFileNoRows.Count <> 1", "destFileNoRows.Count <> 1", "完全一致候補の正式ID・ファイル番号・バーコード番号"):
            self.assertIn(part, step2)
        self.assertIn("sourceFileNoPaddingMap, paddingKey", step2)
        self.assertIn("destFileNoPaddingMap, paddingKey", step2)

    def test_held_source_id_prevents_pending_delete_from_becoming_deletable(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        pending = body("ProcessPendingSyncStatusRows")
        exists = body("RowExistsInSourceByIdentityOrName")
        self.assertIn("sourceFileIdMap, sourceBarcodeMap", step2)
        self.assertIn("ByVal sourceFileIdMap As Object", pending)
        self.assertIn("GetTrimmedCell(wsDst, rowNo, fileIdCol)", pending)
        self.assertIn("If sourceFileIdMap.Exists(fileId) Then", exists)
        self.assertLess(exists.index("If sourceKeyMap.Exists(key) Then"), exists.index("If sourceFileIdMap.Exists(fileId) Then"))
        self.assertIn("identityHold = True", exists)
        self.assertIn("sourceBarcodeMap.Exists(barcode)", exists)
        self.assertIn("sourceFileNoPaddingMap.Exists(paddingKey)", exists)
        self.assertIn("sourceFileNoMap.Exists(fileNo)", exists)
        self.assertIn("sourceFileNoMap, sourceFileNoPaddingMap, visualDiffMap", step2)
        self.assertIn("If identityHold Then", pending)
        self.assertIn("visualDiffMap(markKey) = " + chr(34) + "判定困難" + chr(34), pending)

    def test_missing_id_with_incomplete_name_is_counted_before_early_hold(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        guard = step2.index("If Len(sourceFileIdValue) = 0 Then")
        leave = step2.index("GoTo ContinueNextSourceRow", guard)
        branch = step2[guard:leave]
        self.assertIn("If Len(nameSyncKey) = 0 Then", branch)
        self.assertIn("missingNameDataCount = missingNameDataCount + 1", branch)
        self.assertIn("CStr(srcRow)", branch)
        self.assertIn("missingNameRows", branch)

    def test_missing_created_year_header_is_reported_without_column_zero_access(self):
        code = body("BuildStep2AmbiguousComment")
        self.assertIn("If createdCol > 0 Then", code)
        self.assertIn("createdYearText = " + chr(34) + "（列なし）" + chr(34), code)
        self.assertIn("normalizedCreatedYear = " + chr(34) + "判定不可" + chr(34), code)
        self.assertLess(code.index("If createdCol > 0 Then"), code.index("wsSrc.Cells(r, createdCol).Value"))
        self.assertIn("createdYearText &", code)
        self.assertIn("normalizedCreatedYear &", code)

    def test_cross_year_name_candidate_is_not_unconditionally_a_conflict(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        helper = body("Step2CrossYearNameCandidatesAreDistinct")
        self.assertIn("blockedRows.Count > 0 Or _", step2)
        self.assertIn("Not Step2CrossYearNameCandidatesAreDistinct", step2)
        self.assertIn("Step2RowsForTargetYear", step2)
        self.assertIn("If dstYearValue <= 0 Then", helper)
        self.assertIn("If Len(sourceId) = 0", helper)
        self.assertIn("If StrComp(sourceId, destId", helper)
        self.assertIn("StrComp(sourceBarcode, destBarcode", helper)
        self.assertIn("StrComp(sourceFileNoKey, destFileNoKey", helper)
        self.assertIn("Step2RowsForTargetYear(wsDst, destCandidates", step2)

    def test_source_only_and_cross_year_holds_show_candidate_details(self):
        hold = body("RecordStep2IdentityHold")
        detail = body("Step2CandidateDetails")
        step2 = body("手順2_新ファイル基準表を作成する")
        self.assertIn("If eligibleDestinationCount = 0 Then", hold)
        self.assertIn("unreportedSummary", hold)
        self.assertIn("BuildStep2AmbiguousComment", hold)
        for part in (" / 年度=", " / 状態=", " / ID=", " / フォルダー=", chr(34) + "  行" + chr(34) + " & CStr(rowNo)"):
            self.assertIn(part, detail)
        self.assertIn("unreportedHoldCount", step2)
        self.assertIn("holdCandidatesSummary", step2)
        self.assertIn("GetNormalizedYearFromRow(ws, rowNo, warekiYearCol, westernYearCol)", detail)
        self.assertIn("If yearText = \"0\" Then yearText = \"判定不可\"", detail)


    def test_pending_delete_requires_absence_of_exact_and_partial_identity_evidence(self):
        pending = body("ProcessPendingSyncStatusRows")
        exists = body("RowExistsInSourceByIdentityOrName")
        self.assertIn("sourceFileIdMap.Exists(fileId)", exists)
        self.assertIn("sourceBarcodeMap.Exists(barcode)", exists)
        self.assertIn("sourceFileNoPaddingMap.Exists(paddingKey)", exists)
        self.assertIn("sourceFileNoMap.Exists(fileNo)", exists)
        self.assertIn("ByRef identityHold As Boolean", exists)
        self.assertIn("If identityHold Then", pending)
        self.assertIn("削除判定保留（CSV側に識別候補あり", pending)
        self.assertLess(pending.index("If identityHold Then"), pending.index("Set markRange = wsDst.Range"))
        self.assertIn("StrComp(CStr(visualDiffMap(markKey)), " + chr(34) + "判定困難" + chr(34), pending)


    def test_missing_id_with_or_without_name_and_identifiers_is_reported_before_any_update(self):
        step2 = body("手順2_新ファイル基準表を作成する")
        guard = step2.index("If Len(sourceFileIdValue) = 0 Then")
        no_key = step2.index("If key = " + chr(34) + chr(34) + " Then", guard)
        source_map = step2.index("If Not sourceKeyMap.Exists(key) Then")
        update = step2.index("If existingRowMap.Exists(key) Then")
        branch = step2[guard:no_key]
        self.assertLess(guard, no_key)
        self.assertLess(no_key, source_map)
        self.assertLess(source_map, update)
        for required in (
            "missingFileIdCount = missingFileIdCount + 1",
            "missingFileIdRows = missingFileIdRows",
            "If Len(nameSyncKey) = 0 Then",
            "missingNameDataCount = missingNameDataCount + 1",
            "RecordStep2IdentityHold wsSrc, sourceCandidates",
            "If Len(key) = 0 Then skippedNoKeyCount = skippedNoKeyCount + 1",
            "GoTo ContinueNextSourceRow",
        ):
            self.assertIn(required, branch)
        self.assertLess(branch.index("missingFileIdCount = missingFileIdCount + 1"), branch.index("GoTo ContinueNextSourceRow"))
        self.assertIn("目録CSV行: " , step2)
        self.assertIn("正式ファイルID不足で保留: ", step2)

if __name__ == "__main__":
    unittest.main()
