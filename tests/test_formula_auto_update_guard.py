"""Safety contract: automatic CSV reflection must not erase Excel formulas."""
import re
import unittest
from pathlib import Path

SOURCE = (Path(__file__).resolve().parents[1] / 'vba/Module1.txt').read_text(encoding='utf-8-sig')

def proc(name):
    m = re.search(r'(?ims)^(?:Private|Public) (Sub|Function) ' + re.escape(name) + r'\b.*?^End \1', SOURCE)
    if not m:
        raise AssertionError('missing procedure: ' + name)
    return m.group(0)

class AutoUpdateFormulaSafety(unittest.TestCase):
    def test_whole_group_preflight_is_before_first_auto_write(self):
        body = proc('BidirectionalSyncApplyManagedGroup')
        branch = body[body.index('ElseIf (Not excelChanged) And csvChanged Then'):]
        self.assertLess(branch.index('BidirectionalSyncAutoUpdateTargetsAreSafe'), branch.index('WriteStep2SourceValue'))
        self.assertIn('BidirectionalSyncAppendGroupDecision "競合保留"', branch)

    def test_preflight_checks_every_target_for_formula_merge_and_error(self):
        body = proc('BidirectionalSyncAutoUpdateTargetsAreSafe')
        for token in ('For Each key In itemKeys', '.HasFormula', '.MergeCells', 'IsError(', 'Next key'):
            self.assertIn(token, body)
        self.assertNotIn('WriteStep2SourceValue', body)

    def test_identity_updates_preflight_all_write_targets_before_any_identity_write(self):
        body = proc('BidirectionalSyncApplyThreeWayToRow')
        self.assertLess(body.index('BidirectionalSyncAutoUpdateTargetsAreSafe'), body.index('WriteStep2SourceValue'))
        self.assertIn('identityWriteKeys.Add', body)
        self.assertIn('identityWriteKeys.Count > 0', body)
        self.assertIn('Err.Raise', body[body.index('BidirectionalSyncAutoUpdateTargetsAreSafe'):body.index('WriteStep2SourceValue')])
        self.assertIn('preserveExactIdentity', body)

if __name__ == '__main__':
    unittest.main()
