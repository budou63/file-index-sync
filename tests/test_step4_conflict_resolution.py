from pathlib import Path
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / 'Module1').read_text(encoding='utf-8')


def proc(name):
    match = re.search(rf'(?:Private|Public) (?:Sub|Function) {name}\b.*?\nEnd (?:Sub|Function)', SOURCE, re.S)
    assert match, name
    return match.group(0)


def plan(rows, source, targets):
    """Decision model: scoped identity, verified authority, all-or-nothing preflight."""
    changes = {}
    for item, name, parent in targets:
        identity = (item, name.strip().casefold(), parent if item == '3' else '')
        matching = [r for r in rows if r[:3] == identity]
        if not any(r[4].startswith(('競合', '要確認')) for r in matching):
            continue
        authoritative = {r[3] for r in matching if r[4] == 'システム'}
        if len(authoritative) != 1 or authoritative != {s[3] for s in source if s[:3] == identity}:
            raise ValueError('ambiguous or unverified')
        new = next(iter(authoritative))
        for row in matching:
            if row[4].startswith(('競合', '要確認')):
                if any(other[:3] != identity and other[0] == item and other[2] == parent and other[3] == new for other in rows):
                    raise ValueError('identity collision')
                changes[(identity, row[3])] = new
    return changes


def child_code_supported(code, child_name, new_parent, source, pending, moved_pending):
    """A nonblank child code must be source-backed or a scoped pending move."""
    if not code:
        return True
    source_codes = {row[3] for row in source if row[:3] == ('3', child_name, new_parent)}
    if source_codes:
        return len(source_codes) == 1 and code in source_codes
    return (child_name, new_parent, code) in pending or (child_name, code) in moved_pending


def reject_unplanned_ambiguity(rows, targets, changes):
    for identity in targets:
        if any(planned_identity == identity for planned_identity, _old in changes):
            continue
        matching = [row for row in rows if row[:3] == identity]
        system_codes = {row[3] for row in matching if row[4] == 'システム'}
        pending_codes = {row[3] for row in matching if row[4] == '仮採番'}
        if any(row[4].startswith(('競合', '要確認')) for row in matching) or len(system_codes) > 1 or (system_codes and pending_codes - system_codes):
            raise ValueError('unresolved unplanned scope')


class DecisionTests(unittest.TestCase):
    def test_destination_pending_blocks_child_conflict_parent_migration(self):
        destination = [('3', 'plan', '20', '30', 'システム'),
                       ('3', 'plan', '20', '40', '仮採番')]
        self.assertTrue(any(row[4] != 'システム' for row in destination if row[:3] == ('3', 'plan', '20')))
        source_path = proc('PreflightStep4ConflictResolution')
        self.assertIn('If Trim$(CStr(wsMaster.Cells(r, 4).Value)) <> "システム" Then Err.Raise', source_path)
        self.assertIn('親変更先の第2ガイドに未解決の仮採番・競合があります', source_path)
        self.assertLess(source_path.index('親変更先の第2ガイドに未解決の仮採番・競合があります'), source_path.index('plan(CStr(masterRow)) = Array('))

    def test_extra_provisional_in_planned_scope_is_not_silently_dropped(self):
        rows = [('2', 'road', '', '90', '競合（システム=20）'),
                ('2', 'road', '', '20', 'システム'), ('2', 'road', '', '30', '仮採番')]
        self.assertTrue(any(row[4] == '仮採番' for row in rows if row[:3] == ('2', 'road', '')))
        preflight = proc('PreflightStep4ConflictResolution')
        self.assertIn('If conflictRows.Count > 0 And hasOtherState Then Err.Raise', preflight)
        self.assertIn('If conflictRows.Count > 1 Then Err.Raise', preflight)

    def test_unplanned_system_and_different_provisional_must_stop_before_yes(self):
        rows = [('2', 'road', '', '90', '競合'), ('2', 'road', '', '20', 'システム'),
                ('3', 'child', '20', '88', '仮採番'), ('3', 'child', '20', '30', 'システム')]
        changes = plan(rows, [rows[1], rows[3]], [('2', 'road', '')])
        with self.assertRaises(ValueError):
            reject_unplanned_ambiguity(rows, [('2', 'road', ''), ('3', 'child', '20')], changes)
        check = proc('PreflightStep4ConflictResolution')
        self.assertIn('unresolvedIssue = GetStep4GuideCodeStateIssue(itemName, itemName, displayName, oldParent)', check)
        self.assertIn('If conflictRows.Count > 0 And hasOtherState Then Err.Raise', check)
        self.assertLess(check.index('If conflictRows.Count > 0 And hasOtherState Then Err.Raise'), check.index('plan(CStr(masterRow)) = Array('))
        self.assertLess(check.index('unresolvedIssue = GetStep4GuideCodeStateIssue'), check.index('GoTo NextScope', check.index('unresolvedIssue = GetStep4GuideCodeStateIssue')))

    def test_stale_system_child_from_old_parent_is_not_valid_under_new_parent(self):
        self.assertFalse(child_code_supported('88', 'plan', '20', [], set(), set()))
        self.assertTrue(child_code_supported('88', 'plan', '20', [('3', 'plan', '20', '88', 'システム')], set(), set()))
        self.assertTrue(child_code_supported('88', 'plan', '20', [], {('plan', '20', '88')}, set()))

    def test_two_scoped_conflicts_and_unrelated_pending(self):
        rows = [('2', 'road', '', '90', '競合'), ('2', 'road', '', '20', 'システム'),
                ('3', 'plan', '90', '88', '競合'), ('3', 'plan', '90', '30', 'システム'),
                ('3', 'plan', 'other', '89', '仮採番')]
        source = [rows[1], rows[3]]
        self.assertEqual(plan(rows, source, [('2', 'road', ''), ('3', 'plan', '90')]),
                         {(('2', 'road', ''), '90'): '20', (('3', 'plan', '90'), '88'): '30'})
    def test_ambiguous_and_missing_source_abort(self):
        rows = [('2', 'road', '', '90', '競合'), ('2', 'road', '', '20', 'システム')]
        for source in ([], [('2', 'road', '', '20', 'システム'), ('2', 'road', '', '21', 'システム')]):
            with self.assertRaises(ValueError):
                plan(rows, source, [('2', 'road', '')])
    def test_collision_rejected(self):
        rows = [('3', 'plan', '10', '88', '競合'), ('3', 'plan', '10', '30', 'システム'), ('3', 'other', '10', '30', '仮採番')]
        with self.assertRaises(ValueError):
            plan(rows, [rows[1]], [('3', 'plan', '10')])

    def test_same_child_name_under_other_parent_is_untouched(self):
        rows = [('3', 'plan', '10', '88', '競合'), ('3', 'plan', '10', '30', 'システム'),
                ('3', 'plan', '11', '89', '仮採番')]
        self.assertEqual(plan(rows, [rows[1]], [('3', 'plan', '10')]),
                         {(('3', 'plan', '10'), '88'): '30'})


class SourcePathTests(unittest.TestCase):
    def test_preflight_confirmation_and_revalidation_order(self):
        step4 = proc('手順4_差分インポートCSVを作成する')
        for token in ('PreflightStep4ConflictResolution', 'vbYesNo', 'ApplyStep4ConflictResolution',
                      'ValidateStep4GuideCodeStates wsEdit, targetRows'):
            self.assertIn(token, step4)
        self.assertLess(step4.index('PreflightStep4ConflictResolution'), step4.index('ApplyStep4ConflictResolution'))
        self.assertLess(step4.index('ConfirmAndMarkStep4NewRegistrationCandidates'), step4.index('ApplyStep4ConflictResolution'))
        self.assertLess(step4.index('ApplyStep4ConflictResolution'), step4.index('Set firstGuideCodeMap = BuildCodeMasterNameCodeMap("分類名２", firstGuideUsedCodes, False)', step4.index('ValidateStep4GuideCodeStates wsEdit, targetRows')))
        self.assertLess(step4.index('Set secondGuideCodeMap = BuildCodeMasterNameCodeMap("分類名３", secondGuideUsedCodes, False, True)', step4.index('ValidateStep4GuideCodeStates wsEdit, targetRows')), step4.index('GetStep4GuideSeatCheckReason'))
        self.assertLess(step4.index('ApplyStep4ConflictResolution'), step4.index('ValidateStep4GuideCodeStates wsEdit, targetRows'))
        self.assertLess(step4.index('ValidateStep4GuideCodeStates wsEdit, targetRows'), step4.index('wsOut.Cells.Clear'))
        self.assertLess(step4.index('PreflightStep4ConflictResolution'), step4.index('ConfirmAndMarkStep4NewRegistrationCandidates'))
        self.assertLess(step4.index('wsOut.Cells.Clear'), step4.index('WriteProcessLog "差分インポート抽出"'))
    def test_preconfirmation_has_no_mutations_and_code_recompletion_is_not_relied_on(self):
        step4 = proc('手順4_差分インポートCSVを作成する')
        before = step4[:step4.index('ApplyStep4ConflictResolution')]
        self.assertNotIn('wsEdit.Cells(r, syncStatusColEdit).Value = "文書管理対象外"', before)
        apply = proc('ApplyStep4ConflictResolution')
        self.assertIn('wsEdit.Cells(r, code2Col).Value', apply)
        self.assertIn('wsEdit.Cells(r, code3Col).Value', apply)
        self.assertIn('wsMaster.Rows', apply)
        self.assertIn('WriteStep4ConflictSuccessLog', step4)
        self.assertIn('WriteProcessLog', proc('WriteStep4ConflictSuccessLog'))
        self.assertIn('pendingMarks.Add CLng(rowNo)', proc('ConfirmAndMarkStep4NewRegistrationCandidates'))
        self.assertNotIn('wsEdit.Cells(CLng(rowNo), syncCol).Value = "登録予定"', proc('ConfirmAndMarkStep4NewRegistrationCandidates'))
        self.assertLess(step4.index('If syncStatusChangedCount < 0 Then GoTo SafeExit'), step4.index('ApplyStep4ConflictResolution'))
        self.assertLess(step4.index('ValidateStep4GuideCodeStates wsEdit, targetRows'), step4.index('WriteStep4ConflictSuccessLog'))
        self.assertNotIn('.Save', step4 + apply)
    def test_source_and_scope_guards(self):
        check = proc('PreflightStep4ConflictResolution')
        self.assertIn('SHEET_MOKUROKU_CSV', check)
        self.assertIn('NormalizeClassificationGuideKey', check)
        self.assertIn('親分類コード', check)
        self.assertIn('Err.Raise', check)
        self.assertIn('sourceCodes', check)
        self.assertIn('For Each rowValue In targetRows', check)
        self.assertIn('For r = 2 To lastEdit', check)
        self.assertIn('parentChanges', check)
        self.assertIn('plan(CStr(r)) = Array("親参照"', check)
        self.assertIn('Step4OriginalParentCode', check)
        self.assertIn('Step4HasScopedProvisional', check)
        self.assertIn('If parentChanges.Exists(identity) And sourceCodes.Exists(sourceIdentity) Then', check)
        self.assertIn('If Not coveredByPlan Then Err.Raise', check)
        self.assertIn('If Left$(stateValue, 3) = "要確認" Then Err.Raise', check)
        self.assertIn('"競合（システム=" & newCode & "）"', check)
        self.assertIn('旧親コードが別の第1ガイドでも使用されています', check)
        self.assertIn('旧親コードを別の第1ガイド行でも使用しています', check)
        self.assertIn('旧親コードの第2ガイドを安全に移せません', check)
        self.assertIn('ElseIf parentChanges.Exists(identity) Then', check)
        self.assertIn('親変更後の第2ガイドコードを確認できません', check)
        self.assertIn('対象" & CStr(affectedCount) & "行', check)
        self.assertIn('If previewCount <= 8 Then', check)
        step4 = proc('手順4_差分インポートCSVを作成する')
        self.assertIn('If CStr(conflictPlan(targetRow)(0)) <> "親参照" Then conflictCount = conflictCount + 1', step4)
        apply = proc('ApplyStep4ConflictResolution')
        self.assertIn('Step4OriginalParentCode', apply)
        self.assertIn('If CStr(entry(0)) <> "親参照" Then wsMaster.Rows(r).Delete', apply)

if __name__ == '__main__':
    unittest.main()
