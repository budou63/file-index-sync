import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class HomeLauncherTests(unittest.TestCase):
    def test_launcher_targets_this_workbook_without_replacing_an_unknown_sheet(self):
        source = (ROOT / "modFrmSealNewLayout_Code.txt").read_text(encoding="utf-8")
        launcher = source.split("Public Sub InstallFrmSealNewLauncher()", 1)[1]
        q = chr(34)
        self.assertIn("ThisWorkbook.Worksheets(" + q + "操作パネル" + q + ")", launcher)
        self.assertIn("Set btn = ws.Buttons(" + q + "btnOpenFrmSealNew" + q + ")", launcher)
        self.assertIn("If btn Is Nothing Then", launcher)
        self.assertIn("Err.Raise vbObjectError + 9271", launcher)
        self.assertIn("Replace(ThisWorkbook.Name, quote, quote & quote)", launcher)
        self.assertIn("!modFrmSealNewLayout.ShowFrmSealNew", launcher)
        self.assertNotIn("ActiveWorkbook", launcher)

    def test_calculation_optimization_does_not_abort_unsupported_excel(self):
        source = (ROOT / "Module1").read_text(encoding="utf-8")
        self.assertEqual(source.count("    TrySetManualCalculation" + chr(10)), 10)
        self.assertEqual(source.count("    RestoreCalculationIfPossible "), 11)
        self.assertIn("If Application.Calculation <> xlCalculationManual Then Application.Calculation = xlCalculationManual", source)
        self.assertIn("If Application.Calculation <> savedMode Then Application.Calculation = savedMode", source)
        self.assertNotIn("    Application.Calculation = xlCalculationManual" + chr(10), source)


if __name__ == "__main__":
    unittest.main()
