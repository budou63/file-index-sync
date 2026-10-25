"""Static checks for blank-form paste deployment and native export consistency."""
import re
import subprocess
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CANONICAL = ROOT / "frmSyncReview.frm"
PASTE = ROOT / "vba/frmSyncReview.txt"
NATIVE = ROOT / "frmSyncReview-native.frm"
FRX = ROOT / "frmSyncReview-native.frx"
MODULE1 = ROOT / "vba/Module1.txt"

CONTROLS = {
    "lblCounts": "Label", "cboCategory": "ComboBox", "lblCase": "Label",
    "cboEGroup": "ComboBox", "lblEVisibility": "Label", "cboEVisibility": "ComboBox",
    "lblESelection": "Label", "cmdEClear": "CommandButton",
    "lstCases": "ListBox", "lblRows": "Label", "lstRows": "ListBox",
    "lblCompareRows": "Label", "fraCompare": "Frame", "txtDetails": "TextBox",
    "chkDelete": "CheckBox", "lblStatus": "Label",
    "lblDWarning": "Label",
    "cmdGoTo": "CommandButton", "cmdApprove": "CommandButton",
    "cmdOther": "CommandButton", "cmdHold": "CommandButton",
    "cmdNext": "CommandButton", "cmdDelete": "CommandButton",
    "cmdClear": "CommandButton", "cmdRefresh": "CommandButton", "cmdFinish": "CommandButton",
}
EVENTS = {"cboCategory", "cboEGroup", "cboEVisibility", "lstCases", "lstRows", "chkDelete"} | {n for n in CONTROLS if n.startswith("cmd")}


class BlankFormTests(unittest.TestCase):
    def test_independent_paste_only_code_and_typed_controls(self):
        source = PASTE.read_text(encoding="utf-8")
        self.assertTrue(source.startswith("Option Explicit\n"))
        self.assertNotIn("Attribute VB_", source)
        self.assertNotIn("VERSION 5.00", source)
        self.assertNotIn("Begin VB.", source)
        for name, kind in CONTROLS.items():
            modifier = "WithEvents " if name in EVENTS else ""
            field = "m" + name[0].upper() + name[1:]
            self.assertRegex(source, rf"(?m)^Private {modifier}{field} As MSForms\.{kind}$")
            self.assertIn(f'Set {field} = EnsureControl("{name}", "Forms.{kind}.1")', source)
            self.assertNotRegex(source, rf"(?m)^Private (?:WithEvents )?{name} As ")
            if name in EVENTS:
                suffix = "Change" if name in ("cboCategory", "cboEGroup", "cboEVisibility") else "Click"
                self.assertIn(f"Private Sub {field}_{suffix}()", source)
        self.assertIn("Private mEColumnHeaders(0 To 4) As MSForms.Label", source)
        self.assertIn('EnsureControl("lblEColumnHeader" & CStr(i), "Forms.Label.1")', source)
        self.assertEqual(source.count('EnsureControl("'), len(CONTROLS) + 1)
        self.assertIn("Me.Controls.Add(progId, controlName, True)", source)
        self.assertIn('If existing.Name = controlName Then', source)
        self.assertIn('If TypeName(existing) <> expectedType Then Err.Raise', source)
        self.assertIn('rendering = True\n    BuildControls', source)
        self.assertIn('mCboCategory.ListIndex = 0', source)
        self.assertIn('mCboEVisibility.ListIndex = 0\n    rendering = False', source)
        self.assertLess(source.index('Set mCboCategory = EnsureControl'), source.index('mCboCategory.AddItem "A'))
        show_category = source[source.index('Private Sub ShowCategory()'):source.index('End Sub', source.index('Private Sub ShowCategory()'))]
        self.assertLess(show_category.index('If categoryIndex <> 4 And mLstCases.ListCount > 0 Then mLstCases.ListIndex = 0'), show_category.index('ShowCurrent'))
        self.assertIn("Me.Width = 760", source)
        self.assertIn("Me.Height = Me.Height + (440 - Me.InsideHeight)", source)
        self.assertNotIn("Me.InsideHeight =", source)
        self.assertIn("mTxtDetails.MultiLine = True", source)
        self.assertIn("mTxtDetails.ScrollBars = fmScrollBarsVertical", source)
        self.assertIn("mTxtDetails.Locked = True", source)
        self.assertIn('mChkDelete.Caption = "選択したExcel行を削除対象にする"', source)
        self.assertIn("SyncReviewDeleteBatch(selected)", source)
        self.assertIn("SyncReviewApprove(item, SelectedRow())", source)

    def test_form_defines_sheet_constants_in_every_restore_source(self):
        sources = {
            "paste": PASTE.read_text(encoding="utf-8"),
            "reference": CANONICAL.read_text(encoding="utf-8").split("Option Explicit", 1)[1],
            "native": NATIVE.read_bytes().decode("cp932").split("Option Explicit", 1)[1].replace("\r\n", "\n"),
        }
        for source_name, source in sources.items():
            with self.subTest(source=source_name):
                for identifier, value in (("CSV_SHEET", "目録CSV"), ("DEST_SHEET", "新ファイル基準表")):
                    declaration = f'Private Const {identifier} As String = "{value}"'
                    self.assertEqual(source.count(declaration), 1, identifier)
                    self.assertIn(f"Worksheets({identifier})", source)

    def test_no_vba_component_references_module1_private_constants(self):
        module = MODULE1.read_text(encoding="utf-8-sig")
        private_constants = re.findall(
            r"(?im)^\s*Private\s+Const\s+([A-Za-z_]\w*)\s+As\s+", module
        )
        violations = []
        for component in (ROOT / "vba").glob("*.txt"):
            if component.name == "Module1.txt":
                continue
            source = component.read_text(encoding="utf-8-sig")
            for identifier in private_constants:
                if re.search(rf"\b{re.escape(identifier)}\b", source, re.IGNORECASE):
                    violations.append(f"{component.name}: {identifier}")
        self.assertEqual(violations, [])

    def test_opener_loads_before_show_and_form_initializes_before_data_render(self):
        module = (ROOT / 'vba/modSyncReview.txt').read_text(encoding='utf-8')
        opener = module[module.index('Private Sub OpenSyncReviewForYear('):]
        opener = opener[:opener.index('End Sub')]
        self.assertLess(opener.index('frmSyncReview.LoadCases cases, yearValue'),
                        opener.index('frmSyncReview.Show vbModeless'))
        source = PASTE.read_text(encoding='utf-8')
        initialization = source[source.index('Private Sub UserForm_Initialize()'):]
        initialization = initialization[:initialization.index('End Sub')]
        self.assertIn('BuildControls', initialization)
        self.assertLess(source.index('Private Sub UserForm_Initialize()'), source.index('Public Sub LoadCases('))
        self.assertIn('ShowCategory', source[source.index('Public Sub LoadCases('):source.index('End Sub', source.index('Public Sub LoadCases('))])

    def test_both_exports_share_exact_paste_code_and_native_metadata(self):
        paste = PASTE.read_text(encoding="utf-8")
        canonical = CANONICAL.read_text(encoding="utf-8")
        raw = NATIVE.read_bytes()
        native = raw.decode("cp932")
        self.assertIn(b"\r\n", raw)
        self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"))
        self.assertEqual(raw, native.encode("cp932"))
        self.assertEqual(canonical[canonical.index("Option Explicit"):], paste)
        self.assertEqual(native[native.index("Option Explicit"):].replace("\r\n", "\n"), paste)
        self.assertIn('OleObjectBlob   =   "frmSyncReview-native.frx":0000', native)
        baseline = subprocess.check_output(['git', 'show', 'HEAD:frmSyncReview-native.frm'], cwd=ROOT)
        designer = lambda data: data[:data.index(b'Option Explicit')].replace(b'\r\n', b'\n')
        self.assertEqual(designer(raw), designer(baseline))
        self.assertIn('Attribute VB_PredeclaredId = True', native)
        self.assertIn('ShowModal       =   0', native)
        self.assertGreater(FRX.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
