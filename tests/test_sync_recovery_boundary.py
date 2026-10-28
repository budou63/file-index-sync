"""Static contracts, not Excel execution or proof of SaveCopyAs fidelity."""
import re
import unittest
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
SRC = (ROOT / 'vba/Module1.txt').read_text(encoding='utf-8-sig')
def proc(name):
    m = re.search(r'(?:Public|Private) (?:Sub|Function) ' + re.escape(name) + r'\b.*?\nEnd (?:Sub|Function)', SRC, re.S)
    assert m, name
    return m.group()
class RecoveryBoundary(unittest.TestCase):
    def test_internal_commit_failure_model(self):
        # Requirements model only; no VBA, file I/O or Excel execution.
        from copy import deepcopy
        initial={'input': '未保存入力', 'history': ['Excel採用'], 'baseline': ['土木・2']}
        for phase in ('history', 'baseline', 'summary'):
            with self.subTest(phase=phase):
                recovery=deepcopy(initial)
                live=deepcopy(initial)
                snapshots={k:deepcopy(initial[k]) for k in ('history','baseline')}
                with self.assertRaises(RuntimeError):
                    live['input']='途中まで整形'
                    live['history']=['判断履歴更新後']
                    if phase=='history': raise RuntimeError('injected')
                    live['baseline']=['工務・3']
                    raise RuntimeError('injected')
                live.update(snapshots)
                self.assertEqual(live['history'],initial['history'])
                self.assertEqual(live['baseline'],initial['baseline'])
                self.assertEqual(recovery,initial)
                self.assertNotEqual(live['input'],initial['input'])

    def test_copy_precedes_first_step2_mutation(self):
        s = proc('手順2_新ファイル基準表を作成する')
        self.assertLess(s.index('SyncRecoveryCreateCopy'), s.index('Step2ApplyRegistrationIdentifierCleanup wsDst'))
        self.assertIn('SyncRecoveryFailureNotice', s)
    def test_copy_does_not_save_or_replace_original(self):
        s = proc('SyncRecoveryCreateCopy')
        self.assertIn('ThisWorkbook.SaveCopyAs', s)
        self.assertIn('FileLen', s)
        self.assertIn('FileExists', s)
        self.assertNotIn('ThisWorkbook.Save\n', s)
        self.assertIn('Err.Raise', s)
    def test_context_restores_history_and_baseline_before_discard(self):
        start = proc('BidirectionalSyncStartContext')
        discard = proc('BidirectionalSyncDiscardContext')
        self.assertLess(start.index('SyncRecoveryCaptureInternalSheets'), start.index('BidirectionalSyncEnsureBaselineSheet'))
        self.assertIn('SyncRecoveryRestoreInternalSheets', discard)
        restore = proc('SyncRecoveryRestoreInternalSheets')
        self.assertIn('vbBinaryCompare', restore)
        self.assertIn('Err.Raise', restore)
    def test_failed_session_blocks_another_sync(self):
        s = proc('SyncRecoveryCreateCopy')
        self.assertIn('mSyncRecoveryRequired', s)
        self.assertIn('mSyncRecoveryRequired = True', proc('SyncRecoveryFailureNotice'))
    def test_failure_model_preserves_preoperation_inputs(self):
        # Recovery-copy policy oracle only: never claims a partial live book is restored.
        from copy import deepcopy
        before = {'input': 'unsaved', 'history': ['Excel採用'], 'baseline': 'old'}
        for phase in ('input', 'history', 'baseline'):
            live = deepcopy(before); recovery = deepcopy(before)
            live[phase] = 'partial'
            self.assertEqual(recovery, before)
            self.assertNotEqual(live, before)
