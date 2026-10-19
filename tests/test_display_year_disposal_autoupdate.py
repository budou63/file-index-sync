from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "vba/Module1.txt").read_text(encoding="utf-8")
SHEET = (ROOT / "vba/新ファイル基準表.txt").read_text(encoding="utf-8")


def procedure(source: str, name: str, kind: str = "Sub") -> str:
    match = re.search(
        rf"(?ms)^(?:Private|Public) {kind} {re.escape(name)}\b.*?^End {kind}",
        source,
    )
    if match is None:
        raise AssertionError(f"missing {kind}: {name}")
    return match.group()


class DisplayYearDisposalAutoUpdateTests(unittest.TestCase):
    def test_creation_year_is_independent_from_fiscal_year(self):
        auto = procedure(MODULE, "新ファイル基準表_表示名からID自動補完")
        self.assertIn("createdWarekiCol = FindUniqueHeaderColumn", auto)
        self.assertIn("createdYearCol = FindUniqueHeaderColumn", auto)
        self.assertIn("If changedCell.Column = createdWarekiCol Then", auto)
        self.assertIn("ConvertReiwaToWesternYear(CStr(changedCell.Value2))", auto)
        self.assertIn("wsEdit.Cells(changedCell.Row, createdYearCol).Value2 = CLng(westernYear)", auto)
        self.assertIn("warekiCol = FindUniqueHeaderColumn", auto)
        self.assertIn("wsEdit.Cells(changedCell.Row, seirekiCol).Value2 = CLng(westernYear)", auto)
        branch = auto.split("If createdWarekiCol > 0 And createdYearCol > 0 Then", 1)[1].split("For i = LBound(pairDefs)", 1)[0]
        self.assertNotIn("seirekiCol", branch)

    def test_multicell_edits_are_limited_to_changed_rows(self):
        helper = procedure(MODULE, "新ファイル基準表_廃棄年月日自動計算")
        for contract in (
            "Set workRange = Intersect(Target, monitorRange)",
            "For Each changedCell In workRange.Cells",
            "changedCell.Row >= 2",
            "recalculateRows.Add CStr(changedCell.Row)",
            "dateOnlyRows.Add CStr(changedCell.Row)",
            "For Each key In dateOnlyRows.Keys",
        ):
            self.assertIn(contract, helper)
        self.assertNotIn("For rowNo = 2 To", helper)
        event = procedure(SHEET, "Worksheet_Change")
        self.assertIn("新ファイル基準表_表示名からID自動補完 Target", event)
        self.assertNotIn("新ファイル基準表_廃棄年月日自動計算 Target", event)
        auto = procedure(MODULE, "新ファイル基準表_表示名からID自動補完")
        self.assertEqual(auto.count("Call 新ファイル基準表_廃棄年月日自動計算(Target)"), 1)

    def test_existing_disposal_calculation_is_the_single_source_for_years(self):
        calc = procedure(MODULE, "CalculateHaikiDateText", "Function")
        recalc = procedure(MODULE, "RecalculateHaikiDateByRowMap", "Function")
        self.assertEqual(recalc.count("CalculateHaikiDateText(warekiText, saveTermText)"), 1)
        self.assertIn("wsEdit.Cells(rowNo, haikiCol).Value2 = calcResult", recalc)
        self.assertIn("WriteHaikiYearFieldsFromDateValue(wsEdit, rowNo, calcResult", recalc)
        self.assertIn("haikiWesternYear = startWesternYear + 1", calc)
        self.assertIn('CalculateHaikiDateText = haikiWarekiLabel & ".3.31"', calc)
        self.assertIn("haikiWesternYear = startWesternYear + saveYears + 1", calc)
        self.assertIn('CalculateHaikiDateText = haikiWarekiLabel & ".4.1"', calc)
        self.assertIn("If isPermanent Then Exit Function", calc)

    def test_expected_disposal_cases_and_period_change_sequence(self):
        def to_label(year):
            if year >= 2019:
                return f"R{year - 2018}"
            if year >= 1989:
                return f"H{year - 1988}"
            return ""

        def year_from_label(label):
            match = re.fullmatch(r"([RH])(\d+)", label)
            if match is None:
                return None
            return int(match.group(2)) + (2018 if match.group(1) == "R" else 1988)

        def expected(start, term):
            start_year = year_from_label(start)
            if term in ("継続", "永年"):
                return "", None, None
            if term == "随時":
                year = start_year + 1
                return f"{to_label(year)}.3.31", year, to_label(year)
            match = re.fullmatch(r"(1|3|5|10)年", term)
            if match is None:
                return "", None, None
            year = start_year + int(match.group(1)) + 1
            return f"{to_label(year)}.4.1", year, to_label(year)

        cases = {
            ("R8", "1年"): ("R10.4.1", 2028, "R10"),
            ("R8", "3年"): ("R12.4.1", 2030, "R12"),
            ("R8", "5年"): ("R14.4.1", 2032, "R14"),
            ("R8", "10年"): ("R19.4.1", 2037, "R19"),
            ("R8", "随時"): ("R9.3.31", 2027, "R9"),
            ("R8", "継続"): ("", None, None),
            ("R8", "永年"): ("", None, None),
            ("H27", "5年"): ("R3.4.1", 2021, "R3"),
        }
        for inputs, output in cases.items():
            with self.subTest(inputs=inputs):
                self.assertEqual(expected(*inputs), output)

        transition = [expected("R8", term) for term in ("1年", "継続", "5年")]
        self.assertEqual([result[1:] for result in transition], [(2028, "R10"), (None, None), (2032, "R14")])
        calculate = procedure(MODULE, "CalculateHaikiDateText", "Function")
        self.assertIn("GetSaveTermYears(saveTermText, isPermanent, isZuiji)", calculate)

    def test_manual_date_year_parsing_examples_and_invalid_values(self):
        from datetime import date

        def to_label(year):
            if year >= 2019:
                return f"R{year - 2018}"
            if year >= 1989:
                return f"H{year - 1988}"
            return None

        def parse(value):
            if not value:
                return None, None
            match = re.fullmatch(r"(R|H|令和|平成)?(\d+)[./-](\d{1,2})[./-](\d{1,2})", value, re.IGNORECASE)
            if match is None:
                return None, None
            prefix, era_year, month, day = match.groups()
            if not prefix and len(era_year) != 4:
                return None, None
            if prefix and prefix.upper() in ("R", "令和"):
                year = 2018 + int(era_year)
                label = f"R{int(era_year)}"
            elif prefix and prefix.upper() in ("H", "平成"):
                year = 1988 + int(era_year)
                label = f"H{int(era_year)}"
            else:
                year = int(era_year)
                label = to_label(year)
            try:
                date(year, int(month), int(day))
            except ValueError:
                return None, None
            return year, label

        self.assertEqual(parse("R10.4.1"), (2028, "R10"))
        self.assertEqual(parse("H27.4.1"), (2015, "H27"))
        self.assertEqual(parse("令和10.4.1"), (2028, "R10"))
        self.assertEqual(parse("2028-4-1"), (2028, "R10"))
        for invalid in ("R10.4.31", "R10.13.1", "10.4.1", "R10.4.1x"):
            with self.subTest(invalid=invalid):
                self.assertEqual(parse(invalid), (None, None))
        self.assertEqual(parse(""), (None, None))

    def test_direct_date_updates_only_year_fields_and_invalid_date_is_preserved(self):
        parser = procedure(MODULE, "TryParseHaikiDateValue", "Function")
        writer = procedure(MODULE, "WriteHaikiYearFieldsFromDateValue", "Function")
        self.assertIn("ParseWarekiYear(yearPart, eraName, eraYear, westernYear)", parser)
        self.assertIn("If UBound(parts) <> 2 Then Exit Function", parser)
        self.assertIn("If dayNo > Day(DateSerial(westernYear, monthNo + 1, 0)) Then Exit Function", parser)
        self.assertIn("Len(ToHalfWidthAlphaNumeric(yearPart)) <> 4", parser)
        self.assertIn("If Not TryParseHaikiDateValue(dateValue, westernYear, warekiLabel) Then Exit Function", writer)
        self.assertIn("wsEdit.Cells(rowNo, scheduledYearCol).Value2 = westernYear", writer)
        self.assertIn("wsEdit.Cells(rowNo, scheduledWarekiCol).Value2 = warekiLabel", writer)
        self.assertNotIn("年度", writer)
        self.assertNotIn("保存期間", writer)
        self.assertIn("If IsEmpty(dateValue) Then", writer)
        self.assertIn("If IsError(dateValue) Then Exit Function", writer)

    def test_extension_fields_preserve_scheduled_years_and_warn(self):
        guard = procedure(MODULE, "HasHaikiExtension", "Function")
        recalc = procedure(MODULE, "RecalculateHaikiDateByRowMap", "Function")
        direct = procedure(MODULE, "新ファイル基準表_廃棄年月日自動計算")
        warning = procedure(MODULE, "ShowHaikiExtensionWarning")
        self.assertIn("extensionYearCol", guard)
        self.assertIn("extensionReasonCol", guard)
        self.assertIn("If updateScheduledYears Then", recalc)
        self.assertIn("If HasHaikiExtension(wsEdit, rowNo, extensionYearCol, extensionReasonCol) Then", recalc)
        self.assertIn("AddHaikiExtensionWarningRow skippedExtensionRows, rowNo", recalc)
        self.assertIn("If HasHaikiExtension(wsEdit, rowNo, extensionYearCol, extensionReasonCol) Then", direct)
        self.assertIn("ShowHaikiExtensionWarning skippedExtensionRows", direct)
        self.assertIn("廃棄予定年を自動更新していません", warning)
        self.assertNotIn(".Value2 =", guard)
        self.assertNotIn(".ClearContents", guard)

    def test_full_row_legacy_recalculation_remains_date_only(self):
        all_rows = procedure(MODULE, "RecalculateHaikiDateAllRows_NewFileStandard", "Function")
        self.assertIn("RecalculateHaikiDateByRowMap(wsEdit, rowMap, Nothing, False)", all_rows)
        legacy = procedure(MODULE, "UpdateHaikiDate")
        self.assertIn("新ファイル基準表_廃棄年月日自動計算(Target)", legacy)

    def test_missing_headers_and_event_state_are_safe(self):
        unique = procedure(MODULE, "FindUniqueHeaderColumn", "Function")
        helper = procedure(MODULE, "新ファイル基準表_廃棄年月日自動計算")
        auto = procedure(MODULE, "新ファイル基準表_表示名からID自動補完")
        self.assertIn("FindUniqueHeaderColumn = 0", unique)
        self.assertIn("If matchedCol > 0 Then Exit Function", unique)
        self.assertIn("If haikiDateCol <= 0 Then GoTo SafeExit", helper)
        self.assertIn("scheduledYearCol <= 0 Or scheduledWarekiCol <= 0 Then Exit Function", procedure(MODULE, "WriteHaikiYearFieldsFromDateValue", "Function"))
        self.assertIn("Application.EnableEvents = previousEnableEvents", helper)
        self.assertIn("Application.EnableEvents = previousEnableEvents", auto)


if __name__ == "__main__":
    unittest.main()
