"""Step2 stage-1 extraction: source-bound contract, not an Excel runtime test."""
from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / "vba/Module1.txt").read_text(encoding="utf-8")


def proc(name: str) -> str:
    match = re.search(rf"(?:Public|Private)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)", SOURCE, re.S)
    if not match:
        raise AssertionError(f"missing procedure: {name}")
    return match.group(0)


class SourceTransferExtractionTests(unittest.TestCase):
    def test_existing_transfer_contract_before_and_after_extraction(self):
        step2 = proc("手順2_新ファイル基準表を作成する")
        try:
            body = proc("BuildStep2SourceTransferMap")
        except AssertionError:
            body = step2  # GREEN before extraction: freeze the original decisions.
        required_in_order = (
            "BuildSourceHeaderMap wsSrc, srcLastCol, srcHeaderMap",
            "For srcCol = 1 To srcLastCol",
            "headerText = Trim$(CStr(wsSrc.Cells(1, srcCol).Value))",
            'If headerText <> "" Then',
            "canonical = CanonicalHeaderName(headerText)",
            "dstCol = EnsureColumnByHeader(wsDst, existingDstHeaders, headerText)",
            "If Not IsUserManagedHeader(canonical, srcHeaderMap, baselineUserManagedHeaders) Then",
            "srcToDstMap(canonical) = dstCol",
            "Next srcCol",
        )
        indices = [body.index(item) for item in required_in_order]
        self.assertEqual(indices, sorted(indices))
        self.assertIn('baselineUserManagedHeaders = Array("通し番号", "キャビネット番号", "ｎ＋1", "保存番号", "廃棄年月日", "同期状態")', step2)
        self.assertIn("BuildHeaderMap wsDst, dstLastCol, existingDstHeaders", step2)
        self.assertLess(step2.index("BuildHeaderMap wsDst"), step2.index("' 4) 既存行マップ"))
        self.assertIn("IsUserManagedHeader = Not srcHeaderMap.Exists(canonicalHeader)", proc("IsUserManagedHeader"))

    def test_decision_example_preserves_user_columns_and_skips_empty_headers(self):
        # Python decision example only; the source-bound tests verify VBA wiring.
        headers = ["年度", "タイトル", "", "通し番号", "同期状態", "独自列"]
        managed = {"通し番号", "同期状態"}
        transfer = {}
        for col, header in enumerate(headers, 1):
            if header and header not in managed:
                transfer[header] = col
        self.assertEqual(transfer, {"年度": 1, "タイトル": 2, "独自列": 6})

    def test_stage_one_wires_private_helper_between_user_headers_and_identity_index(self):
        step2 = proc("手順2_新ファイル基準表を作成する")
        helper = proc("BuildStep2SourceTransferMap")
        self.assertLess(step2.index("Step2PreflightSimilarNameColumns"), step2.index("BuildStep2SourceTransferMap wsSrc"))
        self.assertLess(step2.index("GetUniqueSyncStatusColumnOrRaise"), step2.index("BuildStep2SourceTransferMap wsSrc"))
        self.assertEqual(step2.count("BuildStep2SourceTransferMap wsSrc"), 1)
        self.assertIn("BuildStep2SourceTransferMap wsSrc, wsDst, srcLastCol, srcHeaderMap, existingDstHeaders, _\n        baselineUserManagedHeaders, srcToDstMap", step2)
        self.assertLess(step2.index("Step2ConfirmMissingSyncStatus"), step2.index("BuildStep2SourceTransferMap wsSrc"))
        self.assertLess(step2.index("BuildStep2SourceTransferMap wsSrc"), step2.index("BuildExistingRowMap wsDst"))
        self.assertNotIn("For srcCol = 1 To srcLastCol", step2)
        self.assertIn("ByVal baselineUserManagedHeaders As Variant", helper)
        self.assertIn("ByVal srcToDstMap As Object", helper)
        self.assertIn("On Error GoTo ErrorHandler", step2)
        self.assertIn("Application.ScreenUpdating = previousScreenUpdating", step2)
        self.assertIn("Application.EnableEvents = previousEnableEvents", step2)
        self.assertIn("RestoreCalculationIfPossible previousCalculationMode", step2)


if __name__ == "__main__":
    unittest.main()
