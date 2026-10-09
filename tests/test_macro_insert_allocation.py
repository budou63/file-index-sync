"""Macro-insert rows must not allocate new guide codes (static VBA wiring + decision cases)."""
from pathlib import Path
import re
import unittest

MODULE = Path(__file__).resolve().parents[1] / "Module1"
SOURCE = MODULE.read_text(encoding="utf-8")


def procedure(name: str) -> str:
    match = re.search(
        rf"(?:Private|Public)\s+(?:Sub|Function)\s+{re.escape(name)}\b.*?\nEnd\s+(?:Sub|Function)",
        SOURCE,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"Procedure not found: {name}")
    return match.group(0)


def proposed_allocations(rows: list[dict], system_names: set[tuple[str, str]]) -> set[tuple[str, str]]:
    """Decision table, not an Excel/VBA execution substitute."""
    return {
        (level, row[level])
        for row in rows
        if row["sync"] == "登録予定" and "マクロ差込" not in row["remarks"]
        for level in ("分類名２", "分類名３")
        if row[level] and (level, row[level]) not in system_names
    }


class MacroInsertAllocationTests(unittest.TestCase):
    def test_decision_table_keeps_only_normal_rows_and_existing_system_codes(self) -> None:
        system = {("分類名２", "正式座"), ("分類名３", "正式な第2ガイド")}
        rows = [
            {"sync": "登録予定", "remarks": "", "分類名２": "通常座", "分類名３": "通常下位座"},
            {"sync": "登録予定", "remarks": "マクロ差込", "分類名２": "差込だけ", "分類名３": "差込下位"},
            {"sync": "登録予定", "remarks": "マクロ差込／補助", "分類名２": "共有座", "分類名３": "共有下位"},
            {"sync": "登録予定", "remarks": "", "分類名２": "共有座", "分類名３": "共有下位"},
            {"sync": "登録予定", "remarks": "", "分類名２": "正式座", "分類名３": "正式な第2ガイド"},
        ]
        self.assertEqual(proposed_allocations(rows, system), {
            ("分類名２", "通常座"), ("分類名３", "通常下位座"),
            ("分類名２", "共有座"), ("分類名３", "共有下位"),
        })

    def test_step4_already_excludes_macro_rows_before_allocator(self) -> None:
        step4 = procedure("手順4_差分インポートCSVを作成する")
        self.assertLess(step4.index("If ContainsMacroInsertRemark(bikoVal) Then"),
                        step4.index("newRegistrationRowList.Add r"))
        self.assertIn("新ファイル基準表_登録予定分類コード採番 wsEdit, 2, editLastRow, newRegistrationRowList", step4)

    def test_allocator_requires_remarks_header_before_writing_codes(self) -> None:
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn('bikoCol = GetUniqueRemarkColumnOrRaise(wsEdit, lastCol)', allocator)
        self.assertIn('If bikoCol <= 0 Then Err.Raise', allocator)
        self.assertLess(allocator.index('If bikoCol <= 0 Then Err.Raise'),
                        allocator.index('AppendCodeMasterRecord "分類名２"'))

    def test_duplicate_remark_headers_fail_closed_before_step4_or_allocator_changes(self) -> None:
        helper = procedure("GetUniqueRemarkColumnOrRaise")
        self.assertIn("For colNo = 1 To lastCol", helper)
        self.assertIn('CanonicalHeaderName("備考")', helper)
        self.assertIn("If foundCount > 1 Then", helper)
        self.assertIn("Err.Raise", helper)
        step4 = procedure("手順4_差分インポートCSVを作成する")
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        self.assertIn('bikoColEdit = GetUniqueRemarkColumnOrRaise(wsEdit, editLastCol)', step4)
        self.assertLess(step4.index('GetUniqueRemarkColumnOrRaise'), step4.index('macroInsertRows.Add r'))
        self.assertLess(step4.index('macroInsertRows.Add r'), step4.index('wsEdit.Cells(CLng(targetRow), syncStatusColEdit).Value = "文書管理対象外"'))
        self.assertIn('bikoCol = GetUniqueRemarkColumnOrRaise(wsEdit, lastCol)', allocator)
        self.assertLess(allocator.index('GetUniqueRemarkColumnOrRaise'), allocator.index('AppendCodeMasterRecord "分類名２"'))

    def test_both_guide_levels_skip_macro_rows_before_reuse_or_allocation(self) -> None:
        allocator = procedure("新ファイル基準表_登録予定分類コード採番")
        guard = 'If ContainsMacroInsertRemark(wsEdit.Cells(r, bikoCol).Value) Then GoTo '
        self.assertIn(guard + 'ContinueCollectGuide2', allocator)
        self.assertIn(guard + 'ContinueNextRow', allocator)
        self.assertLess(allocator.index(guard + 'ContinueCollectGuide2'),
                        allocator.index('AppendCodeMasterRecord "分類名２"'))
        self.assertLess(allocator.index(guard + 'ContinueNextRow'),
                        allocator.index('AppendCodeMasterRecord "分類名３"'))
        self.assertIn('If NormalizeCellValue(wsEdit.Cells(r, syncCol).Value) <> NormalizeCellValue("登録予定")', allocator)

    def test_source_master_and_pending_reconciliation_remain_independent(self) -> None:
        step3 = procedure("手順3_コード管理CSVを作成する")
        self.assertIn("BuildSystemGuideCodeMasterFromMokuroku wsSrc, wsBuild", step3)
        self.assertIn("RestorePendingGuideCodeRows wsBuild, pendingGuideRows", step3)
        self.assertIn('If exactMap.Exists(exactKey) Then GoTo ContinuePendingRow',
                      procedure("RestorePendingGuideCodeRows"))
