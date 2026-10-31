"""Static regression contracts; no Excel runtime executed."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / 'vba/Module1.txt').read_text(encoding='utf-8-sig')
SHEET = (ROOT / 'vba/新ファイル基準表.txt').read_text(encoding='utf-8-sig')
ROWID = (ROOT / 'vba/modExcelRowId.txt').read_text(encoding='utf-8-sig')


def proc(source, name):
    m = re.search(r'(?ms)^(?:Public|Private) (Sub|Function) ' + re.escape(name) + r'\b.*?^End \1', source)
    if not m:
        raise AssertionError('missing VBA procedure: ' + name)
    return m.group(0)


class EventGuards(unittest.TestCase):
    def test_sheet_event_entry_rejects_any_nonstandard_sheet(self):
        for name in ('Worksheet_Activate', 'Worksheet_Change', 'Worksheet_SelectionChange'):
            body = proc(SHEET, name)
            self.assertIn('IsNewFileStandardWorksheet(Me)', body, name)

    def test_old_copy_entry_predicates_reject_active_baseline(self):
        for name in ('IsRowMoveClickModeActive', 'IsRowMoveModeActive'):
            body = proc(MODULE, name)
            self.assertIn('IsNewFileStandardActiveSheet()', body, name)

    def test_row_move_rechecks_origin_before_cut(self):
        body = proc(MODULE, '行移動_移動先クリックで確定')
        self.assertIn('IsNewFileStandardActiveSheet()', body)
        body = proc(MODULE, 'ExecuteRowMoveToTargetRow')
        self.assertLess(body.index('IsNewFileStandardWorksheet(ws)'), body.index('.Cut'))
        self.assertLess(body.index('IsNewFileStandardActiveSheet()'), body.index('.Cut'))

    def test_baseline_change_events_cannot_mutate_metadata_or_ids(self):
        for name in ('新ファイル基準表_登録予定識別番号を自動クリア', '新ファイル基準表_最終編集情報を記録'):
            self.assertIn('IsNewFileStandardWorksheet(ws)', proc(MODULE, name), name)
        self.assertIn('IsNewFileStandardWorksheet(ws)', proc(ROWID, 'ExcelRowIdHandleWorksheetChange'))

    def test_worksheet_identity_requires_correct_workbook_and_named_sheet(self):
        body = proc(MODULE, 'IsNewFileStandardWorksheet')
        self.assertIn('ThisWorkbook', body)
        self.assertIn('SHEET_NEW_FILE_STANDARD', body)
        self.assertIn('ws Is', body)
