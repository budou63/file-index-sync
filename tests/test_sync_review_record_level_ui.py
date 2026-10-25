"""Static contracts and a reference model for record-level E conflict display."""
from __future__ import annotations

import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8-sig")
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")


def procedure(text: str, name: str, kind: str = "Sub") -> str:
    marker = f"Private {kind} {name}("
    if marker not in text:
        marker = f"{kind} {name}("
    start = text.index(marker)
    end = text.index(f"End {kind}", start)
    return text[start:end]


def record_key(case: dict) -> tuple | None:
    row_id = str(case.get("rowId", ""))
    stable_key = str(case.get("stableKey", ""))
    year = int(case.get("year", 0))
    source_row = int(case.get("sourceRow", 0))
    destination_row = int(case.get("destinationRow", 0))
    if not row_id.strip() or not stable_key or source_row < 2 or destination_row < 2:
        return None
    return year, row_id, stable_key, source_row, destination_row


def build_records(cases: list[dict], target_year: int) -> list[dict]:
    records: list[dict] = []
    by_key: dict[tuple, dict] = {}
    for index, case in enumerate(cases):
        if case.get("category") != "E" or int(case.get("year", 0)) != target_year:
            continue
        key = record_key(case)
        if key is None:
            record = {"cases": [], "identityIssue": True, "sourceIndex": index}
            records.append(record)
        elif key in by_key:
            record = by_key[key]
        else:
            record = {"cases": [], "identityIssue": False, "key": key}
            by_key[key] = record
            records.append(record)
        record["cases"].append(case)

    row_counts: dict[tuple, int] = {}
    stable_counts: dict[tuple, int] = {}
    for record in records:
        if record.get("key"):
            year, row_id, stable_key, _, _ = record["key"]
            row_key = (year, row_id)
            stable_key_id = (year, stable_key)
            row_counts[row_key] = row_counts.get(row_key, 0) + 1
            stable_counts[stable_key_id] = stable_counts.get(stable_key_id, 0) + 1
    for record in records:
        if record.get("key"):
            year, row_id, stable_key, _, _ = record["key"]
            if row_counts[(year, row_id)] > 1 or stable_counts[(year, stable_key)] > 1:
                record["identityIssue"] = True
    return records


def conflict(*, row_id="RID-1", stable_key="FILE-1", year=2026, source_row=41,
             destination_row=18, name="same visible folder", group="group-1", headers=None):
    return {
        "category": "E",
        "year": year,
        "rowId": row_id,
        "stableKey": stable_key,
        "sourceRow": source_row,
        "destinationRow": destination_row,
        "display": (name, 2026),
        "groupName": group,
        "headers": headers if headers is not None else [{"sourceHeader": group}],
    }


class RecordLevelConflictUiTests(unittest.TestCase):
    def test_ten_conflict_items_are_one_file_and_all_items_remain(self):
        cases = [
            conflict(group=f"group-{index}", headers=[{"sourceHeader": f"field-{index}"}])
            for index in range(10)
        ]
        records = build_records(cases, 2026)
        self.assertEqual(len(records), 1)
        self.assertEqual(len(records[0]["cases"]), 10)
        self.assertEqual(
            [case["headers"][0]["sourceHeader"] for case in records[0]["cases"]],
            [f"field-{index}" for index in range(10)],
        )

    def test_580_conflict_items_produce_58_file_rows_without_loss(self):
        cases = [
            conflict(
                row_id=f"RID-{file_index}",
                stable_key=f"FILE-{file_index}",
                source_row=100 + file_index,
                destination_row=300 + file_index,
                group=f"group-{field_index}",
                headers=[{"sourceHeader": f"field-{field_index}"}],
            )
            for file_index in range(58)
            for field_index in range(10)
        ]
        records = build_records(cases, 2026)
        self.assertEqual(len(records), 58)
        self.assertEqual(sum(len(record["cases"]) for record in records), 580)
        self.assertTrue(all(len(record["cases"]) == 10 for record in records))

    def test_same_name_and_year_but_different_management_ids_stay_separate(self):
        cases = [
            conflict(row_id="RID-A", stable_key="FILE-A", source_row=41, destination_row=18),
            conflict(row_id="RID-B", stable_key="FILE-B", source_row=42, destination_row=19),
        ]
        records = build_records(cases, 2026)
        self.assertEqual(len(records), 2)
        self.assertEqual({r["cases"][0]["display"] for r in records}, {("same visible folder", 2026)})

    def test_conflicting_stable_identity_never_merges_and_is_flagged(self):
        cases = [
            conflict(row_id="RID-A", stable_key="FILE-A"),
            conflict(row_id="RID-A", stable_key="FILE-B", source_row=42, destination_row=19),
        ]
        records = build_records(cases, 2026)
        self.assertEqual(len(records), 2)
        self.assertTrue(all(record["identityIssue"] for record in records))

    def test_same_stable_key_with_different_management_ids_never_merges(self):
        cases = [
            conflict(row_id="RID-A", stable_key="FILE-SAME"),
            conflict(row_id="RID-B", stable_key="FILE-SAME", source_row=42, destination_row=19),
        ]
        records = build_records(cases, 2026)
        self.assertEqual(len(records), 2)
        self.assertTrue(all(record["identityIssue"] for record in records))

    def test_missing_identity_fields_are_kept_as_individual_records(self):
        cases = [
            conflict(row_id="", stable_key="FILE-SAME"),
            conflict(row_id="", stable_key="FILE-SAME"),
            conflict(row_id="RID-C", stable_key=""),
        ]
        records = build_records(cases, 2026)
        self.assertEqual(len(records), 3)
        self.assertTrue(all(record["identityIssue"] for record in records))

    def test_other_target_year_is_excluded(self):
        cases = [
            conflict(row_id="RID-26", stable_key="FILE-26", year=2026),
            conflict(row_id="RID-25", stable_key="FILE-25", year=2025),
        ]
        self.assertEqual(len(build_records(cases, 2026)), 1)
        self.assertEqual(len(build_records(cases, 2025)), 1)

    def test_group_filter_keeps_all_cases_for_a_matching_file(self):
        cases = [conflict(group=f"group-{index}") for index in range(10)]
        record = build_records(cases, 2026)[0]
        filtered = [
            record for record in [record]
            if any(case["groupName"] == "group-3" for case in record["cases"])
        ]
        self.assertEqual(len(filtered), 1)
        self.assertEqual(len(filtered[0]["cases"]), 10)

    def test_vba_groups_only_by_explicit_identity_not_display_names(self):
        builder = procedure(FORM, "BuildERecords", "Function")
        key = procedure(FORM, "ERecordBucketKey", "Function")
        show = procedure(FORM, "ShowCategory")
        for token in ('rowId', 'stableKey', 'year', 'sourceRow', 'destinationRow'):
            self.assertIn(f'item("{token}")', builder + key)
        self.assertIn('If CStr(item("category")) = "E" And CLng(item("year")) = yearValue Then', builder)
        self.assertNotIn('groupName', key)
        self.assertNotIn('ERecordCaption', builder + key)
        self.assertIn('If categoryIndex = 4 Then', show)
        self.assertIn('ShowECategory', show)
        self.assertIn('LayoutReviewForE False', show)
        display = procedure(FORM, "ERecordDisplayValues")
        self.assertLess(display.index('EIdentityPair(src, dst, sourceRow, destinationRow, 7)'),
                        display.index('EIdentityPair(src, dst, sourceRow, destinationRow, 8)'))
        self.assertLess(display.index('EIdentityPair(src, dst, sourceRow, destinationRow, 8)'),
                        display.index('EIdentityPair(src, dst, sourceRow, destinationRow, 5)'))
        layout = procedure(FORM, "LayoutReviewForE")
        self.assertIn('mLstCases.ColumnCount = 3', layout)
        self.assertIn('mLstCases.ColumnCount = 1', layout)
        e_show = procedure(FORM, "ShowECategory")
        self.assertIn('mLstCases.List(rowIndex, 1) = folderValue', e_show)
        self.assertIn('mLstCases.List(rowIndex, 2) = creationYearValue', e_show)

    def test_selected_file_renders_every_item_and_both_sides_and_baseline(self):
        current = procedure(FORM, "ShowECurrent")
        for token in ('For Each item In recordCases', 'Set fieldItems = item("headers")',
                      'For Each fieldItem In fieldItems', 'baselineValue', 'systemValue',
                      'excelValue', 'eFieldCases.Add item', 'eFieldItems.Add fieldItem'):
            self.assertIn(token, current)
        self.assertIn('mLstRows.List(rowIndex, 4) = decision', current)
        self.assertIn('mLstRows.ColumnCount = 5', current)
        self.assertIn('ERecordIdentityDetails(record)', current)

    def test_selection_actions_remain_item_level_and_multi_select_is_read_only(self):
        selection = procedure(FORM, "UpdateESelection")
        self.assertIn('identitySafe = Not CBool(record("identityIssue"))', selection)
        self.assertIn('mCmdApprove.Enabled = (n = 1 And Not item Is Nothing And identitySafe And groupSafe)', selection)
        self.assertIn('mCmdOther.Enabled = (n = 1 And Not item Is Nothing And identitySafe And groupSafe)', selection)
        self.assertIn('mCmdHold.Enabled = (n = 1 And Not item Is Nothing And identitySafe And groupSafe)', selection)
        pending = procedure(FORM, "ERecordNeedsConfirmation", "Function")
        self.assertIn('If CBool(record("identityIssue")) Then', pending)
        self.assertIn('識別情報の一貫性を確認できないため', selection)
        manual = (ROOT / "新ファイル基準表.md").read_text(encoding="utf-8")
        for label in ("第2ガイド／個別フォルダー名／作成年度", "不一致項目数", "確認待ちファイル数",
                      "ファイル一覧", "複数ファイルの選択は表示専用", "関連項目全体"):
            self.assertIn(label, manual)
        self.assertIn("A～D表示では従来のカテゴリ別件数・未確認件数を維持し", manual)
        self.assertIn("識別情報が欠落・重複・不整合", manual)
        self.assertIn("第2ガイド・個別フォルダー名・作成年度", (ROOT / "同期候補確認フォーム_導入手順.md").read_text(encoding="utf-8"))
        self.assertIn('If ESelectedCount() <> 1 Then Exit Sub', procedure(FORM, "mCmdApprove_Click"))
        self.assertIn('If ESelectedCount() <> 1 Then Exit Sub', procedure(FORM, "RecordChoice"))
        self.assertIn('BidirectionalSyncResolveConflict(item, decision)', procedure(FORM, "RecordChoice"))
        self.assertNotIn('BidirectionalSyncResolveConflictBatch', FORM + MODULE)

    def test_selected_item_explains_related_group_and_previous_value(self):
        details = procedure(FORM, "EConflictFieldDetails", "Function")
        self.assertIn('EConflictReasonDisplay(item, allMode)', details)
        self.assertIn('item("reason")', procedure(FORM, "EConflictReasonDisplay", "Function"))
        self.assertIn('item("groupName")', details)
        self.assertIn('baselineValue', details)
        self.assertIn('関連項目', details)
        self.assertIn('全体', details)
        self.assertIn('ERecordIdentityDetails(record)', procedure(FORM, "UpdateEFieldSelection"))

    def test_staff_labels_explain_decisions_reasons_and_code_values(self):
        current = procedure(FORM, "ShowECurrent")
        decision = procedure(FORM, "EConflictDecisionDisplay", "Function")
        reason = procedure(FORM, "EConflictReasonDisplay", "Function")
        details = procedure(FORM, "EConflictFieldDetails", "Function")
        self.assertIn('EConflictDecisionDisplay(item)', current)
        for label in ('未判断', '文書管理を採用', 'Excel基準表を採用', '保留中'):
            self.assertIn(label, decision)
        for explanation in ('前回の同期確認値がありません', '両方の値が変更されています',
                            'コード差を安全に判断できない', '保存済みの競合を再表示'):
            self.assertIn(explanation, reason)
        self.assertIn('検出時の説明: ', reason)
        self.assertIn('コード欄:', details)

    def test_display_path_is_read_only_and_actions_keep_existing_persistence(self):
        read_only_helpers = (
            "ShowECategory", "ShowECurrent", "UpdateEFieldSelection",
            "ERecordIdentityDetails", "EConflictFieldDetails",
            "EConflictDecisionDisplay", "EConflictReasonDisplay",
        )
        forbidden = (
            "BidirectionalSyncResolveConflict(", "SyncReviewDecide(",
            "SyncReviewApprove(", "SyncReviewDeleteBatch(",
        )
        for name in read_only_helpers:
            kind = "Function" if name.startswith("ERecord") or name.startswith("EConflict") else "Sub"
            body = procedure(FORM, name, kind)
            for token in forbidden:
                self.assertNotIn(token, body, name)

        reload_ui = procedure(FORM, "mCmdRefresh_Click")
        load_cases = procedure(MODULE, "BidirectionalSyncConflictReviewCases", "Function")
        resolve = procedure(MODULE, "BidirectionalSyncResolveConflict", "Function")
        self.assertIn('Module1.BidirectionalSyncConflictReviewCases(targetYear)', reload_ui)
        self.assertIn('CStr(values(1, 6)) <> BIDIRECTIONAL_SYNC_CONFLICT_SYSTEM', load_cases)
        self.assertIn('item.Add "decision", CStr(values(1, 6))', load_cases)
        self.assertIn('wsConflict.Cells(stateRow, 6).Value2 = decision', resolve)
        self.assertIn('item("decision") = decision', resolve)

    def test_a_to_d_layout_and_single_column_lists_are_restored(self):
        layout = procedure(FORM, "LayoutReviewForE")
        restored = layout[layout.index("Else") :]
        for token in (
            "Me.Width = 760", "mLstCases.ColumnCount = 1", "mLstRows.ColumnCount = 1",
            "PlaceControl mLstCases, 9, 64.5, 187.5, 228",
            "PlaceControl mLstRows, 205, 64.5, 110, 228",
            "mCmdGoTo.Visible = True", "mCmdDelete.Visible = True", "mCmdClear.Visible = True",
            'mCmdApprove.Caption = "同一として処理"', 'mCmdOther.Caption = "別物"',
        ):
            self.assertIn(token, restored)


if __name__ == "__main__":
    unittest.main()
