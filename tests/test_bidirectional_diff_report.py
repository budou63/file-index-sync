# -*- coding: utf-8 -*-
"""第1段階の読取専用・項目差異一覧出力のソース契約テスト。"""

from __future__ import annotations

import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE1 = ROOT / "vba" / "Module1.txt"
MANUAL = ROOT / "新ファイル基準表.md"


class BidirectionalDiffSourceContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.source = MODULE1.read_text(encoding="utf-8")
        cls.manual = MANUAL.read_text(encoding="utf-8")

    def procedure(self, name: str, kind: str) -> str:
        pattern = rf"(?ms)^(?:Public |Private )?{kind} {re.escape(name)}\b.*?^End {kind}$"
        match = re.search(pattern, self.source)
        if match is None:
            self.fail(f"{kind} {name} が見つかりません")
        return match.group(0)

    def test_policy_defines_all_a_to_t_headers(self) -> None:
        block = self.procedure("BidirectionalDiffHeaderPolicies", "Function")
        entries = re.findall(r'Array\("([^"]+)", "([^"]+)", (True|False)\)', block)
        self.assertEqual(
            [header for header, _policy, _compare in entries],
            [
                "通し番号",
                "キャビネット番号",
                "ｎ＋1",
                "毎年度利用区分1:毎年作成 2:単年のみ",
                "原課",
                "年度（和暦）",
                "分類名１",
                "第1ガイド",
                "色名",
                "第2ガイド",
                "2色名",
                "作成年度（和暦）",
                "個別フォルダー",
                "保存期間",
                "媒体種別",
                "保存番号",
                "廃棄年月日",
                "同期状態",
                "内容・取扱い",
                "備考",
            ],
        )
        policies = {header: (policy, comparable) for header, policy, comparable in entries}
        self.assertEqual(policies["保存期間"], ("システム連携", "True"))
        self.assertEqual(policies["備考"], ("システム連携", "True"))
        self.assertEqual(policies["キャビネット番号"], ("職員管理", "False"))
        self.assertEqual(policies["同期状態"], ("マクロ管理", "False"))

    def test_extract_uses_existing_identity_and_name_safety_helpers(self) -> None:
        block = self.procedure("双方向編集_項目差異一覧を抽出", "Function")
        for helper in (
            "BuildStep2IdentityIndex",
            "BidirectionalDiffBuildSourceNameSyncMapCached",
            "BidirectionalDiffBuildDestinationNameSyncMapCached",
            "BidirectionalDiffBuildDestinationAllNameMapCached",
            "BidirectionalDiffTryFormalIdentityMatch",
            "BidirectionalDiffTryUniqueNameMatch",
            "srcLastRow = GetLastUsedRow(wsSrc)",
            "dstLastRow = GetLastUsedRow(wsDst)",
        ):
            self.assertIn(helper, block)
        self.assertNotIn("Step2GetLastDataRow", block)

        formal = self.procedure("BidirectionalDiffTryFormalIdentityMatch", "Function")
        for token in (
            "sourceFullKeyMap",
            "destFullKeyMap",
            "sourceFileIdMap",
            "destFileIdMap",
            "sourceBarcodeMap",
            "destBarcodeMap",
            "sourceFileNoMap",
            "destFileNoMap",
            ".Count <> 1",
            "GetNormalizedYearFromRow",
        ):
            self.assertIn(token, formal)

        name_match = self.procedure("BidirectionalDiffTryUniqueNameMatch", "Function")
        self.assertIn("ByVal sourceHeaderMap As Object, ByVal destHeaderMap As Object", name_match)
        for token in (
            "BidirectionalDiffBuildNameSyncKeyCached",
            "sourceCandidates.Count <> 1 Or destCandidates.Count <> 1",
            "BidirectionalDiffNameCandidateIdentityMatchesCached",
            "BidirectionalDiffRowsForTargetYearCached",
            "BidirectionalDiffCrossYearNameCandidatesAreDistinctCached",
            "Step2HasOtherRows",
        ):
            self.assertIn(token, name_match)

    def test_ambiguous_candidates_are_reported_as_hold_not_normal_match(self) -> None:
        block = self.procedure("双方向編集_項目差異一覧を抽出", "Function")
        hold = self.procedure("BidirectionalDiffAppendIdentityHoldLine", "Sub")
        self.assertIn('"同一性未確定", "比較保留"', hold)
        self.assertIn("BidirectionalDiffAppendIdentityHoldLine", block)
        self.assertNotIn("frmSyncReview.Show", block)
        self.assertNotIn("UpdateRowFromSourceByHeader", block)

    def test_difference_comparison_is_header_based_and_normalized(self) -> None:
        block = self.procedure("BidirectionalDiffAppendDifferences", "Function")
        self.assertIn("BidirectionalDiffHeaderPolicies", block)
        self.assertIn("GetStep2HeaderCandidates(headerName)", block)
        self.assertIn("NormalizeValueForUpdateCompare", block)
        self.assertIn('"項目差異"', block)
        self.assertIn('"比較不能"', block)

    def test_report_is_external_and_does_not_write_to_the_workbook(self) -> None:
        entry = self.procedure("双方向編集_項目差異一覧CSVを外部出力", "Sub")
        extract = self.procedure("双方向編集_項目差異一覧を抽出", "Function")
        writer = self.procedure("BidirectionalDiffWriteUtf8Csv", "Sub")
        combined = "\n".join((entry, extract, writer))

        self.assertIn("Application.GetSaveAsFilename", entry)
        self.assertIn("EnsureCsvExtension", entry)
        self.assertIn("CreateObject(\"ADODB.Stream\")", writer)
        self.assertIn("stream.Charset = \"utf-8\"", writer)
        self.assertIn("stream.SaveToFile outputPath, 2", writer)
        self.assertIn("既存ファイルを保護するため上書きしません", writer)
        csv_line = self.procedure("BidirectionalDiffCsvLine", "Function")
        self.assertIn('""""', csv_line)
        self.assertNotIn(r'\"', csv_line)

        for forbidden in (
            "UpdateRowFromSourceByHeader",
            "AppendRecord",
            "WriteProcessLog",
            "frmSyncReview",
            "ExcelRowIdEnsureSheet",
            "SetSyncStatus",
        ):
            self.assertNotIn(forbidden, combined)

        # 読取りの .Value2 と区別して、セル値・数式・行を代入/削除しないことを確認する。
        self.assertNotRegex(
            combined,
            r"(?m)\.Cells\([^\r\n]+\)\.(?:Value|Value2|Formula|Formula2)\s*=",
        )
        self.assertNotRegex(combined, r"(?m)\.Rows\([^\r\n]+\)\.Delete")
        self.assertNotRegex(combined, r"(?m)\.Cells\([^\r\n]+\)\.Clear")

    def test_manual_describes_first_stage_scope(self) -> None:
        for text in (
            "双方向編集・第1段階：項目差異一覧CSV",
            "双方向編集_項目差異一覧CSVを外部出力",
            "正式ID一意一致",
            "名称一意一致（登録予定）",
            "同一性未確定",
            "読み取り専用",
            "自動反映・自動上書きは行いません",
        ):
            self.assertIn(text, self.manual)


class BidirectionalDiffBehaviorContractTests(unittest.TestCase):
    """VBAの入出力契約を小さな合成データで固定する。"""

    COMPARABLE_HEADERS = ("保存期間", "備考")

    @staticmethod
    def _name_key(row: dict[str, str]) -> tuple[str, str, str, str]:
        return (
            row["作成年度（和暦）"],
            row["第1ガイド"],
            row["第2ガイド"],
            row["個別フォルダー"],
        )

    @classmethod
    def _classify(cls, source: dict[str, str], excel_rows: list[dict[str, str]]) -> tuple[str, dict[str, str] | None]:
        formal = [
            row
            for row in excel_rows
            if all(source.get(key) and row.get(key) == source.get(key) for key in ("ファイルID", "ファイル番号", "バーコード番号"))
        ]
        if len(formal) == 1:
            return "正式ID一意一致", formal[0]

        candidates = [
            row
            for row in excel_rows
            if row.get("同期状態") == "登録予定"
            and not row.get("ファイルID")
            and cls._name_key(row) == cls._name_key(source)
        ]
        if source.get("ファイルID") and len(candidates) == 1:
            return "名称一意一致（登録予定）", candidates[0]
        return "同一性未確定", None

    def test_formal_id_difference_contract_keeps_both_values(self) -> None:
        source = {
            "ファイルID": "ID-1",
            "ファイル番号": "10",
            "バーコード番号": "BC-10",
            "保存期間": "継続",
            "備考": "",
        }
        excel = {
            "ファイルID": "ID-1",
            "ファイル番号": "10",
            "バーコード番号": "BC-10",
            "保存期間": "5年",
            "備考": "Excelで追記",
        }
        kind, matched = self._classify(source, [excel])
        self.assertEqual(kind, "正式ID一意一致")
        self.assertIsNotNone(matched)
        assert matched is not None
        differences = [header for header in self.COMPARABLE_HEADERS if source[header] != matched[header]]
        self.assertEqual(differences, ["保存期間", "備考"])

    def test_unique_name_registration_contract_is_comparable(self) -> None:
        source = {
            "ファイルID": "ID-2",
            "ファイル番号": "20",
            "バーコード番号": "BC-20",
            "作成年度（和暦）": "R8",
            "第1ガイド": "総務",
            "第2ガイド": "庶務",
            "個別フォルダー": "会議資料",
        }
        registration = {
            "ファイルID": "",
            "同期状態": "登録予定",
            "作成年度（和暦）": "R8",
            "第1ガイド": "総務",
            "第2ガイド": "庶務",
            "個別フォルダー": "会議資料",
        }
        self.assertEqual(self._classify(source, [registration])[0], "名称一意一致（登録予定）")

    def test_duplicate_name_registration_contract_is_held(self) -> None:
        source = {
            "ファイルID": "ID-3",
            "作成年度（和暦）": "R8",
            "第1ガイド": "総務",
            "第2ガイド": "庶務",
            "個別フォルダー": "会議資料",
        }
        duplicate = {
            "ファイルID": "",
            "同期状態": "登録予定",
            "作成年度（和暦）": "R8",
            "第1ガイド": "総務",
            "第2ガイド": "庶務",
            "個別フォルダー": "会議資料",
        }
        kind, matched = self._classify(source, [duplicate, duplicate.copy()])
        self.assertEqual(kind, "同一性未確定")
        self.assertIsNone(matched)


if __name__ == "__main__":
    unittest.main()
