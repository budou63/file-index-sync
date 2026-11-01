"""Read-only A/B/C editor information must follow the selected source and destination rows."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
FORM = (ROOT / 'vba/frmSyncReview.txt').read_text(encoding='utf-8')


def procedure(name, kind='Sub'):
    match = re.search(rf'(?ms)^Private {kind} {name}\b.*?^End {kind}', FORM)
    if not match:
        raise AssertionError(f'missing {kind} {name}')
    return match.group()


class EditorMetadataTests(unittest.TestCase):
    def test_abc_metadata_is_appended_after_decision_text(self):
        body = procedure('RenderComparison')
        self.assertIn('If CStr(item("category")) = "A" Or CStr(item("category")) = "B" Or _', body)
        self.assertIn('CStr(item("category")) = "C" Then', body)
        self.assertIn('CompareEditorInfo(src, sourceRow, dst, excelRow)', body)
        self.assertGreater(body.index('CompareEditorInfo('), body.index('【前回判断】'))
        self.assertNotRegex(body, r'(?:\.Value2|\.Value|\.ClearContents)\s*=')

    def test_distinct_correct_labels_and_columns(self):
        info = procedure('CompareEditorInfo', 'Function')
        self.assertIn('Array("最終編集者")', info)
        self.assertIn('Array("更新者メンバーID")', info)
        self.assertIn('Excel側の最終編集者', info)
        self.assertIn('CSV側の更新者メンバーID', info)
        self.assertNotIn('作成者', info)
        self.assertNotRegex(info, r'(?:\.Value2|\.Value|\.ClearContents)\s*=')

    def test_missing_row_or_header_is_explicit_and_selection_refreshes(self):
        read = procedure('CompareEditorValue', 'Function')
        self.assertIn('rowNo < 2', read)
        self.assertIn('If Not headerFound Then', read)
        self.assertIn('未取得', read)
        self.assertIn('（空欄）', read)
        selection = procedure('ShowCurrentSelection')
        self.assertIn('RenderComparison item, rowNo', selection)
        build = procedure('BuildControls')
        self.assertIn('mTxtDetails.Locked = True', FORM)
        self.assertIn('mTxtDetails.ScrollBars = fmScrollBarsVertical', FORM)


if __name__ == '__main__':
    unittest.main()
