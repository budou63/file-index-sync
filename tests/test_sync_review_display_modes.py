"""Static contracts for E-conflict normal/all display filtering and counts."""
from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "vba/frmSyncReview.txt").read_text(encoding="utf-8-sig")
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8-sig")

EXPECTED_HEADERS = {
    "毎年度利用区分1:毎年作成 2:単年のみ": "毎年度利用区分（1：毎年作成、2：単年のみ）",
    "原課": "原課",
    "年度（和暦）": "年度（和暦）",
    "年度(和暦)": "年度（和暦）",
    "分類名１": "分類名１",
    "分類名1": "分類名１",
    "第1ガイド": "第1ガイド",
    "第１ガイド": "第1ガイド",
    "分類名２": "第1ガイド",
    "分類名2": "第1ガイド",
    "第2ガイド": "第2ガイド",
    "第２ガイド": "第2ガイド",
    "分類名３": "第2ガイド",
    "分類名3": "第2ガイド",
    "作成年度（和暦）": "作成年度（和暦）",
    "作成年度(和暦)": "作成年度（和暦）",
    "個別フォルダー": "個別フォルダー（タイトル）",
    "個別フォルダ": "個別フォルダー（タイトル）",
    "タイトル": "個別フォルダー（タイトル）",
    "保存期間": "保存期間",
    "媒体種別": "媒体種別",
    "内容・取扱い": "内容・取扱い",
    "備考": "備考",
}

PAIRED_GROUPS = {
    "原課": ("原課・原課ID", ("原課",)),
    "分類名１": ("分類名１・分類コード１", ("分類名１", "分類名1")),
    "第1ガイド": ("第1ガイド・分類コード２", ("第1ガイド", "第１ガイド", "分類名２", "分類名2")),
    "第2ガイド": ("第2ガイド・分類コード３", ("第2ガイド", "第２ガイド", "分類名３", "分類名3")),
    "保存期間": ("保存期間・保存期間ID", ("保存期間",)),
    "媒体種別": ("媒体種別・媒体種別ID", ("媒体種別",)),
}


def procedure(name: str, kind: str = "Sub") -> str:
    marker = f"Private {kind} {name}("
    if marker not in FORM:
        marker = f"{kind} {name}("
    start = FORM.index(marker)
    end = FORM.index(f"End {kind}", start)
    return FORM[start:end]


def module_procedure(name: str) -> str:
    marker = f"Private Function {name}("
    if marker not in MODULE:
        marker = f"Function {name}("
    start = MODULE.index(marker)
    end = MODULE.index("End Function", start)
    return MODULE[start:end]


def parse_header_display_map() -> dict[str, str]:
    body = procedure("EHeaderNormalDisplayCaption", "Function")
    parsed: dict[str, str] = {}
    branch = re.compile(
        r"(?ims)^\s*(?:If|ElseIf)\s+(.+?)\s+Then\s*\n\s*"
        r"EHeaderNormalDisplayCaption\s*=\s*\"([^\"]+)\""
    )
    for condition, caption in branch.findall(body):
        for header in re.findall(r'CompareHeaderKey\("([^"]+)"\)', condition):
            parsed[header] = caption
    return parsed


class NormalDisplayTests(unittest.TestCase):
    def test_exact_header_allowlist_and_aliases(self):
        actual = parse_header_display_map()
        self.assertEqual(actual, EXPECTED_HEADERS)
        body = procedure("EHeaderNormalDisplayCaption", "Function")
        self.assertNotIn("InStr(", body)  # classification must not use partial matches
        for hidden in ("保管期間ID", "履歴番号", "分類コード１", "image", "画像"):
            self.assertNotIn(f'CompareHeaderKey("{hidden}")', body)

    def test_group_scope_is_explicit_and_header_source_and_destination_must_agree(self):
        classifier = procedure("EFieldNormalDisplayCaption", "Function")
        group_map = procedure("EGroupMatchesNormalDisplay", "Function")
        self.assertIn('sourceCaption = EHeaderNormalDisplayCaption(CStr(fieldItem("sourceHeader")))', classifier)
        self.assertIn('destinationCaption = EHeaderNormalDisplayCaption(CStr(fieldItem("destinationHeader")))', classifier)
        self.assertIn("StrComp(sourceCaption, destinationCaption, vbBinaryCompare)", classifier)
        self.assertIn("EGroupMatchesNormalDisplay(CStr(caseItem(\"groupName\"))", classifier)
        for caption, (group, headers) in PAIRED_GROUPS.items():
            with self.subTest(caption=caption):
                self.assertIn(f'Case "{caption}"', group_map)
                self.assertIn(f'CompareHeaderKey("{group}")', group_map)
                self.assertIn(f'"{group}"', MODULE)
                for header in headers:
                    self.assertIn(f'CompareHeaderKey("{header}")', group_map)
        self.assertNotIn("InStr(", group_map)

    def test_requested_fields_and_aliases_follow_existing_header_resolution(self):
        displayed = module_procedure("GetStep2DisplayHeadersOrder")
        candidates = module_procedure("GetStep2HeaderCandidates")
        canonical = module_procedure("CanonicalHeaderName")
        display_headers = (
            "毎年度利用区分1:毎年作成 2:単年のみ", "原課", "年度（和暦）", "分類名１",
            "第1ガイド", "第2ガイド", "作成年度（和暦）", "個別フォルダー",
            "保存期間", "媒体種別", "内容・取扱い", "備考",
        )
        for header in display_headers:
            with self.subTest(header=header):
                self.assertIn(f'"{header}"', displayed)
        self.assertIn('Array("分類名１", "分類名1")', candidates)
        self.assertIn('Array("第1ガイド", "第１ガイド", "分類名２", "分類名2")', candidates)
        self.assertIn('Array("第2ガイド", "第２ガイド", "分類名３", "分類名3")', candidates)
        self.assertIn('Array("個別フォルダー", "個別フォルダ", "タイトル")', candidates)
        self.assertIn('Array("作成年度（和暦）", "作成年度(和暦)")', candidates)
        self.assertIn('Case "分類名2", "分類名２", "第1ガイド", "第１ガイド"', canonical)
        self.assertIn('Case "分類名3", "分類名３", "第2ガイド", "第２ガイド"', canonical)

    def test_display_mode_defaults_to_normal_and_toggle_only_rerenders(self):
        self.assertIn('mCboEVisibility.AddItem "通常表示"', FORM)
        self.assertIn('mCboEVisibility.AddItem "すべて表示"', FORM)
        self.assertIn("mCboEVisibility.ListIndex = 0", FORM)
        self.assertIn("EIsAllDisplayMode = (mCboEVisibility.ListIndex = 1)", procedure("EIsAllDisplayMode", "Function"))
        event = procedure("mCboEVisibility_Change")
        self.assertIn("ShowECategory True", event)
        for forbidden in ("BidirectionalSyncResolveConflict", "SyncReviewDecide(", "WriteStep2", "SaveConflict"):
            self.assertNotIn(forbidden, event)

    def test_api_free_size_and_pane_presets_resize_lists_headers_and_details(self):
        self.assertIn('mCboEWindowSize.AddItem "コンパクト"', FORM)
        self.assertIn('mCboEWindowSize.AddItem "標準"', FORM)
        self.assertIn('mCboEWindowSize.AddItem "拡大"', FORM)
        self.assertIn('mCboEPaneSize.AddItem "一覧を広く"', FORM)
        self.assertIn('mCboEPaneSize.AddItem "比較を広く"', FORM)
        layout = procedure("LayoutReviewForE")
        for token in (
            "Me.Width = 960", "targetInsideHeight = 400", "Me.Width = 1100",
            "Me.Width = 1280", "targetInsideHeight = 560",
            "Me.Height = Me.Height + (targetInsideHeight - Me.InsideHeight)",
            "PlaceControl mLblCounts, 9, 6, availableWidth - 18, 34",
            "PlaceControl mLstCases, 9, 79.5, leftWidth, listHeight",
            "PlaceControl mLstRows, rightLeft, 79.5, rightWidth, listHeight",
            "PlaceControl mEColumnHeaders(i), headerLeft, 65, columnWidths(i), 14",
            "PlaceControl mTxtDetails, rightLeft, detailsTop, rightWidth, detailsHeight",
            "PlaceControl mLblStatus, 9, footerTop, availableWidth - 18, 21",
        ):
            self.assertIn(token, layout)
        self.assertIn("mLstRows.ColumnWidths = CStr(CLng(columnWidths(0)))", layout)
        self.assertNotIn('mLstRows.ColumnWidths = "130 pt;145 pt;145 pt;145 pt;180 pt"', procedure("ShowECurrent"))
        self.assertIn('mLblRows.Caption = "項目群"', layout)
        self.assertIn('mLblCase.Caption = "確認が必要な入力項目"', layout)
        self.assertIn('ControlTipText = "項目グループで対象ファイルを絞り込みます。判断状態は変更しません。"', FORM)
        top_row = ((205, 42), (251, 160), (419, 42), (464, 112), (582, 35), (620, 88), (716, 28), (748, 112))
        for (left, width), (next_left, _) in zip(top_row, top_row[1:]):
            self.assertLessEqual(left + width, next_left)
        self.assertLessEqual(top_row[-1][0] + top_row[-1][1], 900)
        self.assertIn("footerTop = Me.InsideHeight - 84", layout)
        self.assertIn("PlaceControl mCmdApprove, 9, footerTop + 23.5, 110, 21", layout)
        for event_name in ("mCboEWindowSize_Change", "mCboEPaneSize_Change"):
            self.assertIn("ShowECategory True", procedure(event_name))
        for forbidden in ("SetWindowLong", "FindWindow", "CreateWindow", "PtrSafe"):
            self.assertNotIn(forbidden, FORM)
        self.assertNotRegex(FORM, r"(?im)^\s*(?:Private\s+)?Declare\s+")

    def test_counts_are_dynamic_existing_judgment_units_and_partition_the_total(self):
        show = procedure("ShowECategory")
        self.assertIn("totalItems = totalItems + recordCases.Count", show)
        self.assertIn("ECaseNeedsHumanDecision(item)", show)
        self.assertIn("ECaseHasNormalDisplayFields(item)", show)
        self.assertIn("inputFileKeys.Count", show)
        self.assertIn("CStr(inputItemCount)", show)
        self.assertIn("CStr(internalMismatchCount)", show)
        self.assertIn("CStr(safeNoReviewCount)", show)
        self.assertIn("CStr(unresolvedInternalCount)", show)
        self.assertIn("CStr(internalOnlyFileCount)", show)
        self.assertIn("CStr(identityIssueFileCount)", show)
        self.assertIn("safeNoReviewCount = 0", show)
        self.assertNotIn("BidirectionalSyncLastSafeNoReviewCount", FORM)
        self.assertNotRegex(show, r"\b(?:385|580)\b")
        self.assertIn('"確認が必要なファイル: "', show)
        self.assertIn("確認が必要な入力項目:", show)
        self.assertIn("内部項目の不一致:", show)
        self.assertIn("安全に確認不要:", show)
        self.assertIn("未確認・要調査:", show)
        internal_count = procedure("ECountInternalCase")
        self.assertIn("internalMismatchCount = internalMismatchCount + 1", internal_count)
        self.assertIn("unresolvedInternalCount = unresolvedInternalCount + 1", internal_count)

    def test_normal_view_lists_only_open_input_files_and_all_view_keeps_internal_records(self):
        show_records = procedure("ShowECategory")
        self.assertIn("(EIsAllDisplayMode() Or ERecordNeedsConfirmation(record)) And ERecordHasGroup(record, selectedGroup)", show_records)
        needs_input = procedure("ERecordNeedsConfirmation", "Function")
        self.assertIn("ECaseNeedsHumanDecision(item)", needs_input)
        self.assertIn("ECaseHasNormalDisplayFields(item)", needs_input)
        self.assertIn("ERecordNeedsInternalInvestigation(record)", show_records)
        self.assertIn("内部のみ・要調査: ", show_records)
        self.assertIn("識別要確認: ", FORM)
        self.assertIn("自動解決はせず、差異を保持しています。", FORM)
        show_current = procedure("ShowECurrent")
        self.assertIn("For Each item In recordCases", show_current)
        self.assertIn("If allMode Or Len(displayCaption) > 0 Then", show_current)
        self.assertIn("eFieldCases.Add item", show_current)
        self.assertIn("eFieldItems.Add fieldItem", show_current)
        self.assertIn("通常表示の対象項目に不一致はありません。その他の項目に", show_current)
        self.assertIn("ファイル一覧には残しています", show_current)

    def test_selection_is_restored_by_file_and_visible_item_without_changing_judgments(self):
        show = procedure("ShowECategory")
        for token in (
            "If preserveFileSelection Then",
            "selectedRecordKeys(ERecordBucketKey(record)) = True",
            "selectedRecordKeys.Exists(recordKey)",
            "mLstCases.Selected(rowIndex) = True",
            'priorCaseKey = CStr(priorCase("key"))',
            "CompareHeaderKey(priorHeaderName)",
            "mLstRows.ListIndex = i - 1",
        ):
            self.assertIn(token, show)
        self.assertIn('mLstRows.ListIndex = -1', procedure("ShowECurrent"))
        self.assertNotIn("BidirectionalSyncResolveConflict", show)
        self.assertNotIn("WriteStep2", show)

    def test_hidden_members_make_normal_mode_actions_fail_closed(self):
        selection = procedure("UpdateESelection")
        safe = procedure("ESelectedDecisionIsSafe", "Function")
        for handler in (procedure("mCmdApprove_Click"), procedure("RecordChoice")):
            self.assertIn("ESelectedDecisionIsSafe()", handler)
            self.assertLess(handler.index("ESelectedDecisionIsSafe()"), handler.index("BidirectionalSyncResolveConflict(item,"))
        self.assertIn("ECaseHasOutOfScopeFields(item)", safe)
        self.assertIn("If Not EIsAllDisplayMode() And ECaseHasOutOfScopeFields(item) Then Exit Function", safe)
        out_of_scope = procedure("ECaseHasOutOfScopeFields", "Function")
        self.assertIn("If Len(EFieldNormalDisplayCaption(caseItem, fieldItem)) = 0 Then", out_of_scope)
        self.assertIn("ECaseHasOutOfScopeFields = True", out_of_scope)
        self.assertIn("identitySafe And groupSafe", selection)
        self.assertIn("全項目を確認するため「すべて表示」に切り替えてください", selection)

    def test_next_unresolved_skips_hidden_only_units_in_normal_mode(self):
        next_case = procedure("SelectNextECase")
        self.assertIn("Not EIsAllDisplayMode() And (Not ECaseHasNormalDisplayFields(item) Or ECaseHasOutOfScopeFields(item))", next_case)
        self.assertIn("通常表示で操作できる未解決項目はありません", next_case)
        self.assertIn("すべて表示", next_case)

    def test_existing_resolution_engine_and_a_to_d_paths_are_not_replaced(self):
        self.assertIn("Public Function BidirectionalSyncResolveConflict(", MODULE)
        self.assertIn('SyncReviewApprove(item, SelectedRow())', procedure("mCmdApprove_Click"))
        self.assertIn('SyncReviewDecide(item, decision)', procedure("RecordChoice"))
        self.assertIn("If mCboCategory.ListIndex = 4 Then", procedure("mCmdNext_Click"))
        self.assertIn('If CStr(item("category")) = "B" And decision = "別物" Then', procedure("RecordChoice"))
        layout = procedure("LayoutReviewForE")
        a_to_d = layout.split("\n    Else\n", 1)[1]
        self.assertIn("mLblEVisibility.Visible = False", a_to_d)
        self.assertIn("mCboEVisibility.Visible = False", a_to_d)
        self.assertIn("mCboEWindowSize.Visible = False", a_to_d)
        self.assertIn("mCboEPaneSize.Visible = False", a_to_d)
        self.assertIn("mLblCompareRows.Visible = True", a_to_d)
        self.assertIn("Me.Width = 760", a_to_d)
        self.assertIn("Me.Height = Me.Height + (440 - Me.InsideHeight)", a_to_d)


if __name__ == "__main__":
    unittest.main()
