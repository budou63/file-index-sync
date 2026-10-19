import unittest
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/modSyncReview.txt").read_text(encoding="utf-8")
MANUAL = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8")

def proc(source, start, end="End Function"):
    i = source.index(start)
    return source[i:source.index(end, i) + len(end)]

def number_key(value):
    value = str(value or "")
    return (value.lstrip("0") or "0") if value.isdigit() and value else value

def guid_key(value):
    try:
        return str(uuid.UUID(str(value).strip("{}"))).upper()
    except (ValueError, TypeError, AttributeError):
        return None

def stable_key(source, rows):
    file_id = source.get("id", "")
    if not file_id or sum(row.get("id") == file_id for row in rows) != 1:
        return None
    year = source["year"]
    number = number_key(source.get("number"))
    barcode = source.get("barcode", "")
    if not number or not barcode:
        return None
    if number and sum(row["year"] == year and number_key(row.get("number")) == number for row in rows) != 1:
        return None
    if barcode and sum(row["year"] == year and row.get("barcode") == barcode for row in rows) != 1:
        return None
    return "SR1|{}|{}:{}|{}:{}|{}:{}".format(year, len(file_id), file_id, len(number), number, len(barcode), barcode)

def history_decision(history, source, csv_rows, excel_row, excel_rows, legacy=None):
    key = stable_key(source, csv_rows)
    row_id = guid_key(excel_row.get("id"))
    if not key or not row_id or sum(guid_key(row.get("id")) == row_id for row in excel_rows) != 1:
        return None
    for entry in reversed(history):
        if entry.get("format") == "IDv1" and entry.get("csv_key") == key and entry.get("row_id") == row_id:
            return entry.get("decision")
        if not entry.get("format") and not entry.get("csv_key") and not entry.get("row_id") and entry.get("snapshot") == legacy:
            return entry.get("decision")
    return None

def exact_match(source, csv_rows, excel_rows):
    fid, number, barcode, year = source.get("id", ""), number_key(source.get("number")), source.get("barcode", ""), source["year"]
    if not fid or not number or not barcode or sum(r.get("id") == fid for r in csv_rows) != 1:
        return None
    if sum(r["year"] == year and number_key(r.get("number")) == number for r in csv_rows) != 1:
        return None
    if sum(r["year"] == year and r.get("barcode") == barcode for r in csv_rows) != 1:
        return None
    matches = [r for r in excel_rows if r.get("id") == fid]
    if len(matches) != 1:
        return None
    row = matches[0]
    if row.get("year") != year or number_key(row.get("number")) != number or row.get("barcode") != barcode or row.get("status") != "同期済":
        return None
    for other in excel_rows:
        if other is not row and other.get("id") != fid and (other.get("barcode") == barcode or number_key(other.get("number")) == number):
            return None
    return row

def filter_candidates(source, csv_rows, excel_rows, candidates):
    if exact_match(source, csv_rows, excel_rows) is None:
        return list(candidates)
    title = "".join(str(source.get("folder", "")).split()).casefold()
    return [r for r in candidates if r.get("status") == "登録予定" and not r.get("id") and
            "".join(str(r.get("folder", "")).split()).casefold() == title]

def assigned_title_conflicts(source, candidates):
    title = "".join(str(source.get("folder", "")).split()).casefold()
    return [row for row in candidates if row.get("status") == "登録予定" and row.get("id") and
            row.get("id") != source.get("id") and
            "".join(str(row.get("folder", "")).split()).casefold() == title]


class StableIdentityTests(unittest.TestCase):
    def setUp(self):
        self.source = {"year": 2026, "id": "SYS-1", "number": "00123", "barcode": "BC-9", "folder": "旧名", "guide1": "第1", "guide2": "第2"}
        self.csv = [dict(self.source)]
        self.row = {"id": "11111111-1111-4111-8111-111111111111", "folder": "候補", "guide1": "第1", "guide2": "第2", "status": "登録予定"}
        self.excel = [self.row]
        self.history = [{"format": "IDv1", "csv_key": stable_key(self.source, self.csv),
                         "row_id": self.row["id"].upper(), "decision": "別物"}]

    def test_csv_and_excel_row_reordering_and_nonidentity_edits_keep_history(self):
        source = dict(self.source, folder="改名", guide1="新ガイド1", guide2="新ガイド2", csv_row=80, retention="永年")
        row = dict(self.row, excel_row=900, folder="再編集", guide1="編集後1", guide2="編集後2", retention="10年")
        other_source = {"year": 2025, "id": "SYS-2", "number": "888", "barcode": "BC-8"}
        other_row = dict(self.row, id="22222222-2222-4222-8222-222222222222", excel_row=2)
        self.assertEqual(history_decision(self.history, source, [other_source, self.source], row, [other_row, row]), "別物")
        reader = proc(MODULE, "Public Function SyncReviewDecision(")
        self.assertIn("storedManagementId = managementId And storedStableKey = stableKey", reader)
        self.assertIn('keyVersion = "IDv1"', reader)
        id_branch = reader.split('If keyVersion = "IDv1" Then', 1)[1].split("ElseIf", 1)[0]
        self.assertNotIn("sourceHash", id_branch)
        self.assertNotIn("selectedHash", id_branch)
        comment = proc(MODULE, "Public Function SyncReviewDecisionComment(")
        comment_branch = comment.split('If keyVersion = "IDv1" Then', 1)[1].split("ElseIf", 1)[0]
        self.assertNotIn("sourceHash", comment_branch)
        self.assertNotIn("selectedHash", comment_branch)

    def test_copy_or_changed_system_identity_does_not_inherit_history(self):
        copied = dict(self.row, id="22222222-2222-4222-8222-222222222222")
        self.assertIsNone(history_decision(self.history, self.source, self.csv, copied, [self.row, copied]))
        for field, value in (("id", "SYS-2"), ("number", "124"), ("barcode", "BC-10"), ("year", 2027)):
            changed = dict(self.source, **{field: value})
            self.assertIsNone(history_decision(self.history, changed, [changed], self.row, self.excel))

    def test_invalid_duplicate_or_missing_ids_and_csv_identity_fail_closed(self):
        for raw_id in ("", "invalid"):
            row = dict(self.row, id=raw_id)
            self.assertIsNone(history_decision(self.history, self.source, self.csv, row, [row]))
        self.assertIsNone(history_decision(self.history, self.source, self.csv, self.row, [self.row, dict(self.row)]))
        missing = dict(self.source, id="")
        self.assertIsNone(history_decision(self.history, missing, [missing], self.row, self.excel))
        self.assertIsNone(history_decision(self.history, self.source, self.csv + [dict(self.source)], self.row, self.excel))
        for field in ("number", "barcode"):
            missing = dict(self.source, **{field: ""})
            self.assertIsNone(history_decision(self.history, missing, [missing], self.row, self.excel))
        self.assertIn("CLng(reviewManagementIdCounts(normalizedId)) <> 1", proc(MODULE, "Private Function ReviewManagementRowId("))
        stable = proc(MODULE, "Private Function ReviewSourceStableKey(")
        self.assertIn("ReviewRows(csvIds, fileId).Count <> 1", stable)
        self.assertIn("Len(fileId) = 0 Or Len(fileNumber) = 0 Or Len(barcode) = 0", stable)

    def test_old_audit_rows_only_use_unchanged_strict_signature_path(self):
        old = [{"snapshot": "same", "decision": "別物"}]
        self.assertEqual(history_decision(old, self.source, self.csv, self.row, self.excel, "same"), "別物")
        self.assertIsNone(history_decision(old, self.source, self.csv, self.row, self.excel, "moved"))
        reader = proc(MODULE, "Public Function SyncReviewDecision(")
        self.assertIn("ReviewAuditHashMatches(ws.Cells(r, 9).Value2, sourceHash)", reader)
        self.assertIn("ReviewAuditHashMatches(ws.Cells(r, selectedSignatureCol).Value2, selectedHash)", reader)
        self.assertIn("Len(keyVersion) = 0 And Len(storedManagementId) = 0", reader)
        self.assertIn("ReviewAuditTextAt(ws, r, 11)", proc(MODULE, "Public Function SyncReviewDecisionComment("))
        self.assertIn("If IsError(stored) Then Exit Function", proc(MODULE, "Private Function ReviewAuditHashMatches("))
        self.assertIn("If IsNull(stored) Then Exit Function", proc(MODULE, "Private Function ReviewAuditHashMatches("))
        for header in ("判断キー形式", "選択Excel管理行ID", "CSV安定識別キー"):
            self.assertIn('ReviewAuditColumn(ws, "' + header + '", True)', proc(MODULE, "Private Function ReviewAudit("))
        writer = proc(MODULE, "Private Sub ReviewWriteAudit(", "End Sub")
        self.assertIn('ws.Cells(r, keyVersionCol).Value2 = "IDv1"', writer)
        self.assertIn("ws.Cells(r, managementIdCol).Value2 = managementId", writer)
        self.assertIn("ws.Cells(r, stableKeyCol).Value2 = stableKey", writer)

    def test_exact_formal_match_suppresses_weak_but_not_same_title_registration(self):
        source = dict(self.source, folder="正式名")
        official = {"year": 2026, "id": "SYS-1", "number": "123", "barcode": "BC-9", "status": "同期済"}
        weak = {"status": "登録予定", "id": "", "folder": "似た名"}
        same_title = dict(weak, folder="正式名", guide="別ガイド")
        self.assertEqual(filter_candidates(source, self.csv, [official], [weak]), [])
        self.assertEqual(filter_candidates(source, self.csv, [official], [weak, same_title]), [same_title])
        candidates = proc(MODULE, "Public Function SyncReviewCandidates(")
        self.assertIn("ReviewFindUniqueFormalMatch", candidates)
        self.assertIn("ReviewExactFolderPlannedRows", candidates)
        self.assertIn('cases.Add NewCase("A"', candidates)
        helper = proc(MODULE, "Private Function ReviewExactFolderPlannedRows(")
        self.assertIn("ReviewRows(folderIndex, sourceFolder)", helper)
        self.assertIn("ReviewFiscal(dst, CLng(rowValue)) = yearValue", helper)
        build = proc(MODULE, "Public Function SyncReviewCandidates(")
        self.assertIn('Set folders = CreateObject("Scripting.Dictionary")', build)
        self.assertIn("ReviewMapAdd folders, ReviewTextKey", build)

    def test_exact_formal_match_accepts_real_sync_status_and_rejects_old_label(self):
        source = dict(self.source, folder="正式名")
        official = {"year": 2026, "id": "SYS-1", "number": "123", "barcode": "BC-9", "status": "同期済"}
        weak = {"status": "登録予定", "id": "", "folder": "似た名"}
        self.assertIsNotNone(exact_match(source, self.csv, [official]))
        self.assertEqual(filter_candidates(source, self.csv, [official], [weak]), [])
        old_label = dict(official, status="同期済み")
        self.assertIsNone(exact_match(source, self.csv, [old_label]))
        self.assertEqual(filter_candidates(source, self.csv, [old_label], [weak]), [weak])
        formal = proc(MODULE, "Private Function ReviewFindUniqueFormalMatch(")
        candidates = proc(MODULE, "Public Function SyncReviewCandidates(")
        self.assertIn('ReviewField(dst, matchedRow, "同期状態") <> "同期済"', formal)
        self.assertIn('ReviewField(dst, matchedFormalRow, "同期状態") <> "同期済"', candidates)
        self.assertNotIn('"同期済み"', formal + candidates)

    def test_duplicate_or_conflicting_identifiers_and_wrong_year_do_not_suppress(self):
        source = dict(self.source, folder="正式名")
        official = {"year": 2026, "id": "SYS-1", "number": "123", "barcode": "BC-9", "status": "同期済"}
        weak = {"status": "登録予定", "id": "", "folder": "似た名"}
        cases = [(self.csv + [dict(self.source)], [official]),
                 (self.csv, [dict(official, barcode="wrong")]),
                 (self.csv, [dict(official, year=2025)]),
                 (self.csv, [official, dict(official, id="SYS-2", barcode="BC-9")]),
                 ([dict(self.source, barcode="")], [official])]
        for csv_rows, excel_rows in cases:
            self.assertEqual(filter_candidates(source, csv_rows, excel_rows, [weak]), [weak])

    def test_leading_zero_file_number_match_is_supported(self):
        official = {"year": 2026, "id": "SYS-1", "number": "123", "barcode": "BC-9", "status": "同期済"}
        self.assertIsNotNone(exact_match(self.source, self.csv, [official]))
        self.assertIn("ReviewPadding(ReviewField(dst, matchedRow,", proc(MODULE, "Private Function ReviewFindUniqueFormalMatch("))

    def test_candidate_revalidation_snapshots_row_ids_and_source_key(self):
        item = proc(MODULE, "Private Function NewCase(")
        snapshot = proc(MODULE, "Private Function ReviewManagementIdSnapshot(")
        current = proc(MODULE, "Private Function ReviewCurrent(")
        index = proc(MODULE, "Private Sub ReviewBuildManagementIdIndex(", "End Sub")
        self.assertIn('item.Add "destinationIdSnapshot"', item)
        self.assertIn('item.Add "sourceStableKey"', item)
        self.assertIn("ReviewManagementRowId(ws, CLng(rowValue))", snapshot)
        self.assertIn('candidate("destinationIdSnapshot")', current)
        self.assertIn('candidate("sourceStableKey")', current)
        self.assertIn('CreateObject("Scripting.Dictionary")', index)
        self.assertIn("For rowNo = 2 To lastRow", index)

    def test_manual_documents_stable_ids_and_formal_match_rules(self):
        self.assertIn("Excel管理行ID", MANUAL)
        self.assertIn("正式識別", MANUAL)

    def test_exact_title_lookup_is_indexed_not_limited_to_weak_candidate_rows(self):
        helper = proc(MODULE, "Private Function ReviewExactFolderPlannedRows(")
        self.assertIn("folderIndex As Object", helper)
        self.assertIn("ReviewRows(folderIndex, sourceFolder)", helper)
        self.assertNotIn("candidates As Collection", helper)
        build = proc(MODULE, "Public Function SyncReviewCandidates(")
        self.assertLess(build.index("ReviewMapAdd folders"), build.index("For sRow = 2 To srcLast", build.index("ReviewMapAdd folders")))

    def test_same_title_planned_row_with_another_id_remains_an_a_conflict(self):
        source = dict(self.source, folder="正式名")
        planned = {"status": "登録予定", "id": "SYS-2", "folder": "正式名"}
        self.assertEqual(assigned_title_conflicts(source, [planned]), [planned])
        helper = proc(MODULE, "Private Function ReviewExactFolderAssignedRows(")
        self.assertIn("rowFileId <> sourceFileId", helper)
        self.assertIn("ReviewRows(folderIndex, sourceFolder)", helper)
        self.assertIn("ReviewFiscal(dst, CLng(rowValue)) = yearValue", helper)
        build = proc(MODULE, "Public Function SyncReviewCandidates(")
        self.assertIn("ReviewExactFolderAssignedRows", build)
        self.assertIn('cases.Add NewCase("A"', build)

if __name__ == "__main__":
    unittest.main()
