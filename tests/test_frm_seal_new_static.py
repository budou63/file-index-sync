from __future__ import annotations
from pathlib import Path
import re
import unittest
ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / "frmSealNew_FormCode.txt").read_text(encoding="utf-8")
LAYOUT = (ROOT / "modFrmSealNewLayout_Code.txt").read_text(encoding="utf-8")
EVENT_CLASS = (ROOT / "clsFrmSealNewEvent_Code.txt").read_text(encoding="utf-8")
def project_source() -> str:
    return "\n".join((ROOT / name).read_text(encoding="utf-8-sig", errors="replace") for name in ("Module1", "frmSeal_", "固定印刷設定.bas"))
class FrmSealNewStaticTests(unittest.TestCase):
    def test_existing_click_handlers_match_dispatch_exactly(self):
        existing = set(re.findall(r"(?im)^Private Sub (\w+)_Click\(", (ROOT / "frmSeal_").read_text(encoding="utf-8")))
        dispatch = re.findall(r"(?im)^\s*Case [\"](\w+)[\"]: (\w+)_Click", FORM)
        self.assertEqual(len(dispatch), len(existing))
        self.assertEqual({name for name, _ in dispatch}, existing)
        self.assertTrue(all(name == target for name, target in dispatch))
    def test_click_handler_bodies_preserve_original_logic(self):
        baseline = (ROOT / "frmSeal_").read_text(encoding="utf-8")
        names = re.findall(r"(?im)^Private Sub (\w+)_Click\(", baseline)
        def body(source, name):
            return re.search(r"(?ims)^Private Sub " + re.escape(name) + r"_Click\(\)(.*?)^End Sub", source).group(1).strip()
        for name in names:
            current = body(FORM, name)
            current = re.sub(r'Me\.Controls\("(tgl(?:MoveMode|DuplicateCheckMode))"\)', r"Me.\1", current)
            original = body(baseline, name)
            original = re.sub(r'Me\.Controls\("(tgl(?:MoveMode|DuplicateCheckMode))"\)', r"Me.\1", original)
            with self.subTest(handler=name):
                self.assertEqual(current, original)
    def test_every_dispatched_action_is_generated_and_has_event_sink(self):
        keys = {name for name, _ in re.findall(r"(?im)^\s*Case [\"](\w+)[\"]: (\w+)_Click", FORM)}
        self.assertTrue(all((chr(34) + name + chr(34)) in LAYOUT for name in keys))
        self.assertIn("sinks.Add sink", LAYOUT)
        self.assertIn("Private WithEvents mButton As MSForms.CommandButton", EVENT_CLASS)
        self.assertIn("Private WithEvents mToggle As MSForms.ToggleButton", EVENT_CLASS)
    def test_all_23_generated_action_keys_have_dispatch_handlers(self):
        dispatched = re.findall(r"(?im)^\s*Case [\"](\w+)[\"]: (\w+)_Click", FORM)
        self.assertEqual(len(dispatched), 23)
        self.assertEqual(len({key for key, _ in dispatched}), 23)
        generated = re.findall(r"(?im)^\s*names = Array\((.+)\)", LAYOUT)
        direct = re.findall(r"(?im)^\s*AddAction target, sinks, [\"](\w+)[\"]", LAYOUT)
        generated_names = [name for row in generated for name in re.findall(chr(34) + r"(\w+)" + chr(34), row)]
        self.assertEqual(len(generated_names) + len(direct), 23)
        self.assertEqual(set(generated_names + direct), {key for key, _ in dispatched})
    def test_old_form_control_refs_are_late_bound_for_empty_copy(self):
        old_form = (ROOT / "frmSeal_").read_text(encoding="utf-8")
        for name in ("tglDuplicateCheckMode", "tglMoveMode"):
            self.assertNotIn("Me." + name, old_form)
            self.assertIn("Me.Controls(" + chr(34) + name + chr(34) + ")", old_form)
    def test_missing_sheet_event_module_dependency_is_present(self):
        sheet = (ROOT / "新ファイル基準表_SheetModule.bas").read_text(encoding="utf-8-sig")
        self.assertIn("Private Sub Worksheet_Change", sheet)
        self.assertIn("Private Sub Worksheet_SelectionChange", sheet)
        self.assertIn("行移動_移動先クリックで確定", sheet)
    def test_codepage_sensitive_symbols_use_chrW(self):
        module = (ROOT / "Module1").read_text(encoding="utf-8")
        for literal in ("ChrW(&H2212)", "ChrW(&H2713)", "ChrW(&HB7)", "ChrW(&H2011)", "ChrW(&H2013)", "ChrW(&H2014)"):
            self.assertIn(literal, module)
        for symbol in ("−", "✓", "·", "‑", "–", "—"):
            self.assertNotIn(chr(34) + symbol + chr(34), module)
    def test_shared_seal_helpers_match_legacy_and_are_used_without_form_instantiation(self):
        legacy = (ROOT / "frmSeal_").read_text(encoding="utf-8-sig")
        shared = (ROOT / "modSealShared_Code.txt").read_text(encoding="utf-8")
        start = "Public Sub BuildIndexes_NewSpec("
        self.assertEqual(shared[shared.index(start):], legacy[legacy.index(start):])
        module = (ROOT / "Module1").read_text(encoding="utf-8")
        self.assertEqual(re.findall(r"\bfrmSeal\.(\w+)", module), ["Show"])
        for name in ("BuildIndexes_NewSpec", "WriteOneSeal_NewSpec", "ExtractWarekiNumber", "SaveTermToKei"):
            with self.subTest(helper=name):
                self.assertIn("modSealShared." + name, module)
                self.assertIn("modSealShared." + name, FORM)
        for name in ("BuildIndexes_NewSpec", "WriteOneSeal_NewSpec", "ExtractWarekiNumber", "SaveTermToKei", "FindHeaderByCandidates"):
            with self.subTest(public_api=name):
                self.assertRegex(FORM, r"(?im)^Public (?:Sub|Function) " + name + r"\b")
        self.assertIn("GroupBase slot, baseCol, baseRow, offSlash, offG1", shared)
        self.assertIn("PutMergeTopLeft wsSeal, baseRow, baseCol", shared)

    def test_button_caption_arrays_are_initialized_before_use(self):
        self.assertIn("Dim labels As Variant, names As Variant", LAYOUT)
        self.assertEqual(LAYOUT.count("labels = Array("), 3)
        self.assertEqual(LAYOUT.count("CStr(labels(i))"), 3)
        for name in ("CSVを取り込む", "重複確認", "移動開始"):
            self.assertIn(name, LAYOUT)
    def test_initialization_is_idempotent_and_display_initializes_first(self):
        self.assertIn("Public Sub EnsureFrmSealNewInitialized()", FORM)
        self.assertIn("If Not mSinks Is Nothing Then Exit Sub", FORM)
        self.assertIn("EnsureFrmSealNewInitialized", FORM.split("Private Sub UserForm_Initialize()", 1)[1])
        self.assertRegex(LAYOUT, r"Public Sub ShowFrmSealNew\(\)\s+frmSealNew\.EnsureFrmSealNewInitialized\s+frmSealNew\.Show")
    def test_production_sources_do_not_contain_transient_probes(self):
        for source in (FORM, LAYOUT, EVENT_CLASS):
            self.assertNotRegex(source, r"(?i)Hermes[A-Za-z]*Probe|HermesAudit|__probe|HERMES_EVENT_CONNECTED")
        self.assertIn("Set sink = New clsFrmSealNewEvent", LAYOUT)
        self.assertIn("sink.Connect c, frm", LAYOUT)
    def test_single_slots_and_bulk_fields_keep_existing_names(self):
        for name in ("tbID", "TextBox1", "TextBox2"):
            self.assertIn(name, LAYOUT)
        for name in ("GetSerialBySlot", "GetBulkStartFromForm", "GetBulkEndFromForm"):
            self.assertIn(name, FORM)
    def test_existing_macro_entry_points_are_defined_in_repository(self):
        source = project_source()
        names = ["CSV取込_シンプル版", "手順2_新ファイル基準表を作成する", "手順4_差分インポートCSVを作成する", "手順5_差分インポートCSVを外部出力する", "新ファイル基準表生成", "キャビネットガイド作成_実行", "旧基準表_継続文書移行", "新ファイル基準表_通し番号_手動再採番", "固定印刷設定を登録する", "固定設定で印刷する", "色ルールチェック", "重複フォルダ候補をマーキングして並び替え", "重複確認モードを開始する", "重複確認モードを完了する", "行移動モード開始_または確定", "行移動_キャンセル", "行移動_上へ", "行移動_下へ", "行移動_指定行へ", "行移動_移動先クリック開始", "行移動_元に戻す", "GenerateMultiSealSheets_NewSpec", "ClearAllSeals_NewSpec"]
        for name in names:
            with self.subTest(macro=name):
                self.assertRegex(source, r"(?im)^\s*(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function)\s+" + re.escape(name) + r"\b")
    def test_paste_files_have_no_probe_or_vbe_export_attributes(self):
        for source in (FORM, LAYOUT, EVENT_CLASS):
            self.assertNotIn("ProbeFrmSealNew", source)
            self.assertNotIn("Attribute VB_", source)
        self.assertIn("Public Sub ShowFrmSealNew()", LAYOUT)
if __name__ == "__main__":
    unittest.main()
