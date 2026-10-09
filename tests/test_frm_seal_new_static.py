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
            with self.subTest(handler=name):
                self.assertEqual(current, body(baseline, name))
    def test_every_dispatched_action_is_generated_and_has_event_sink(self):
        keys = {name for name, _ in re.findall(r"(?im)^\s*Case [\"](\w+)[\"]: (\w+)_Click", FORM)}
        self.assertTrue(all((chr(34) + name + chr(34)) in LAYOUT for name in keys))
        self.assertIn("sinks.Add sink", LAYOUT)
        self.assertIn("Private WithEvents mButton As MSForms.CommandButton", EVENT_CLASS)
        self.assertIn("Private WithEvents mToggle As MSForms.ToggleButton", EVENT_CLASS)
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
