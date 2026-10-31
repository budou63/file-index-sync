"""Independent scenario oracle; this does not execute or compile VBA."""
import unittest

def decide(base, excel, csv, saved=None, *, identity=True, classification_ready=True):
    if saved is not None and not identity:
        return ('stop_identity', excel, base)
    if saved not in (None, '自動解決'):
        if excel == csv and classification_ready:
            return ('auto_resolved', excel, csv)
        return ('review', excel, base)
    if excel == csv:
        return ('confirmed', excel, csv)
    if excel == base:
        return ('csv_apply', csv, csv)
    if csv == base:
        return ('excel_pending', excel, base)
    return ('review', excel, base)

class SavedDecisionScenarios(unittest.TestCase):
    def test_a_hold_then_excel_reverts(self):
        self.assertEqual(decide('5', '5', '3', '保留'), ('review', '5', '5'))

    def test_b_hold_then_values_agree(self):
        self.assertEqual(decide('5', '3', '3', '保留'), ('auto_resolved', '3', '3'))

    def test_c_hold_then_excel_changes(self):
        self.assertEqual(decide('5', '7', '3', '保留'), ('review', '7', '5'))

    def test_d_excel_accepted_then_excel_changes(self):
        self.assertEqual(decide('5', '7', '3', 'Excel採用'), ('review', '7', '5'))

    def test_e_no_history_keeps_csv_only_automation(self):
        self.assertEqual(decide('5', '5', '3'), ('csv_apply', '3', '3'))

    def test_unchanged_hold_or_excel_choice_keeps_values(self):
        for decision in ('保留', 'Excel採用', '未解決'):
            self.assertEqual(decide('5', '10', '3', decision), ('review', '10', '5'))

    def test_identity_mismatch_never_auto_resolves_equal_values(self):
        self.assertEqual(decide('5', '3', '3', '保留', identity=False), ('stop_identity', '3', '5'))

    def test_code_protection_prevents_automatic_resolution(self):
        self.assertEqual(decide('5', '3', '3', '保留', classification_ready=False), ('review', '3', '5'))

    def test_closed_history_does_not_block_future_csv_changes(self):
        self.assertEqual(decide('3', '3', '8', '自動解決'), ('csv_apply', '8', '8'))

if __name__ == '__main__':
    unittest.main()
