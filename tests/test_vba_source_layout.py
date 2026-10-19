"""One paste-ready source per managed VBA component; restore artifacts stay separate."""
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
STANDARD = ("Module1", "modSealShared", "modFrmSealNewLayout", "modSyncReview",
            "modExcelRowId", "固定印刷設定", "modDummyExport")
OTHER_CODE = ("clsFrmSealNewEvent", "frmSeal", "frmSealNew", "frmSyncReview", "新ファイル基準表")
FORM_ARTIFACTS = ("frmSyncReview.frm", "frmSyncReview-native.frm", "frmSyncReview-native.frx")


class SourceLayoutTests(unittest.TestCase):
    def test_each_managed_component_has_one_paste_ready_source(self):
        for name in STANDARD + OTHER_CODE:
            with self.subTest(name=name):
                path = ROOT / "vba" / (name + ".txt")
                self.assertTrue(path.is_file(), str(path))
                text = path.read_text(encoding="utf-8-sig")
                self.assertTrue(text.startswith("Option Explicit\n") or
                                text.startswith("Option Explicit\r\n"))
                self.assertNotIn("Attribute VB_Name", text)
                self.assertEqual(text.encode("cp932").decode("cp932"), text)

    def test_form_artifacts_remain_at_their_existing_paths(self):
        for name in FORM_ARTIFACTS:
            with self.subTest(name=name):
                self.assertTrue((ROOT / name).is_file(), name)

    def test_no_legacy_paste_source_at_root(self):
        for name in ("Module1", "modSealShared_Code.txt", "modFrmSealNewLayout_Code.txt",
                     "modSyncReview.bas", "固定印刷設定.bas", "modDummyExport.bas",
                     "clsFrmSealNewEvent_Code.txt", "frmSealNew_FormCode.txt", "frmSeal_",
                     "frmSyncReview_FormCode.txt", "新ファイル基準表_SheetModule.bas"):
            with self.subTest(name=name):
                self.assertFalse((ROOT / name).exists(), name)

    def test_readme_maps_destinations_and_does_not_point_to_removed_files(self):
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        for name in STANDARD + OTHER_CODE:
            self.assertIn("vba/" + name + ".txt", readme)
        for name in ("frmSeal", "frmSealNew", "frmSyncReview"):
            self.assertIn(name, readme)
        self.assertNotIn("frmSeal_新仕様差し替え関数.bas", readme)
        self.assertNotIn("frmSeal_btnApply_Click_差し替え.bas", readme)


if __name__ == "__main__":
    unittest.main()
