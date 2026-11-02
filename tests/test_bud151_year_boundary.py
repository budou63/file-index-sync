"""Year-boundary source contracts and static decision examples; no Excel/VBA runtime."""
from pathlib import Path
import re, unittest
ROOT=Path(__file__).resolve().parents[1]
REVIEW=(ROOT/'vba/modSyncReview.txt').read_text(encoding='utf-8-sig')
MOD=(ROOT/'vba/Module1.txt').read_text(encoding='utf-8-sig')
def proc(text,name):
 m=re.search(r'(?:Public|Private)\s+(?:Sub|Function)\s+'+re.escape(name)+r'\b.*?\nEnd\s+(?:Sub|Function)',text,re.I|re.S)
 if not m: raise AssertionError(name)
 return m.group()
class YearBoundary(unittest.TestCase):
 def test_reorder_guard_precedes_first_step2_sheet_write(self):
  s=proc(MOD,'手順2_新ファイル基準表を作成する')
  self.assertLess(s.index('Step2PreflightFormulaReorder wsDst'),s.index('Step2ApplyRegistrationIdentifierCleanup'))
  self.assertLess(s.index('Step2PreflightFormulaReorder wsDst'),s.index('EnsureColumnByHeader'))
  guard=proc(MOD,'Step2PreflightFormulaReorder')
  self.assertIn('If possibleReorder Then',guard)
  self.assertIn('Err.Raise',guard)
  self.assertIn('BuildColumnOrderMap',guard)
  self.assertNotIn('ReorderSheetColumnsByMap wsDst',s)
  self.assertIn('予期しない列順変更',s)
  self.assertIn('Step2HeaderAffectsColumnOrder',guard)
  self.assertIn('For Each header In managedHeaders',guard)
  self.assertIn('For srcCol = 1 To GetLastHeaderCol(wsSrc)',guard)
  self.assertNotIn('SpecialCells',guard)
  helper=proc(MOD,'Step2HeaderAffectsColumnOrder')
  self.assertIn('For Each headerGroup In groups',helper)
  self.assertIn('Next headerGroup',helper)
  self.assertNotIn('Next group',helper)
  self.assertIn('GetStep2HeaderCandidates',helper)
  self.assertIn('CanonicalHeaderName',helper)
 def test_review_candidate_indexes_scoped_to_target_year(self):
  s=proc(REVIEW,'SyncReviewCandidates')
  self.assertIn('If ReviewFiscal(dst, dRow) = yearValue Then',s)
  self.assertIn('If sourceFiscalYear = yearValue Then',s)
 def test_review_b_uniqueness_checks_scope_year(self):
  s=proc(REVIEW,'SyncReviewApprove')
  self.assertIn('If i <> sRow And ReviewFiscal(src, i) = CLng(item("year")) Then',s)
  self.assertIn('If i <> selectedRow And ReviewFiscal(dst, i) = CLng(item("year")) Then',s)
 def test_step2_destination_maps_scope_target(self):
  s=proc(MOD,'手順2_新ファイル基準表を作成する')
  self.assertIn('BuildStep2IdentityIndex wsDst, dstLastRow, GetLastHeaderCol(wsDst), targetYearValue',s)
  self.assertIn('BuildDestinationAllNameMap wsDst, dstLastRow, GetLastHeaderCol(wsDst), targetYearValue',s)
 def test_new_row_lookup_never_selects_other_year(self):
  caller=proc(MOD,'PlaceStep2NewRowsNearGroups')
  self.assertIn('FindStep2RowByFileKey(wsDst, headers, CStr(rowKey), targetYearValue)',caller)
  s=proc(MOD,'FindStep2RowByFileKey')
  self.assertIn('GetNormalizedYearFromRow(ws, rowNo, warekiCol, yearCol) = targetYearValue',s)
 def test_d_snapshot_ignores_other_year_values(self):
  s=proc(REVIEW,'ReviewCopySheetSignature')
  self.assertIn('ByVal yearValue As Long',s)
  self.assertIn('If ReviewFiscal(ws, rowNo) = yearValue Then',s)
  self.assertIn('ReviewCopyDatasetSignature(src, dst, yearValue)',proc(REVIEW,'SyncReviewCopyIdCandidates'))
 def test_e_stable_key_uniqueness_scoped_to_source_year(self):
  s=proc(MOD,'BidirectionalSyncBatchSourceKeyIsUnique')
  self.assertIn('targetYear = GetNormalizedYearFromRow(wsSrc, sourceRow, warekiCol, yearCol)',s)
  self.assertIn('If GetNormalizedYearFromRow(wsSrc, rowNo, warekiCol, yearCol) <> targetYear Then GoTo NextSourceKeyRow',s)
 def test_unknown_year_data_preflight_before_step2_and_review(self):
  s=proc(MOD,'手順2_新ファイル基準表を作成する')
  self.assertLess(s.index('ValidateFiscalDataRows wsDst'),s.index('SyncRecoveryCreateCopy'))
  self.assertIn('ValidateFiscalDataRows wsSrc',s)
  r=proc(REVIEW,'SyncReviewCandidates')
  self.assertIn('ValidateFiscalDataRows src',r)
  self.assertIn('ValidateFiscalDataRows dst',r)
  guard=proc(REVIEW,'ValidateFiscalDataRows')
  self.assertIn('CountA',guard)
  self.assertIn('年度（和暦）',guard)
  self.assertIn('Err.Raise',guard)
  self.assertIn('warekiText',guard)
  self.assertIn('westernText',guard)
  self.assertIn('LookIn:=xlFormulas',guard)
  self.assertIn('SearchFormat:=False',guard)
  self.assertIn('lastUsedCol',guard)
  self.assertIn('Len(warekiText) = 0',guard)
  for reason in ('年度（和暦）が空欄です','年度（和暦）が不正です','年度（和暦）と年度が一致しません'):
   self.assertIn(reason,guard)

 def test_preflight_existing_ids_use_only_target_year(self):
  s=proc(MOD,'Step2PreflightSimilarNameColumns')
  self.assertIn('GetNormalizedYearFromRow(wsDst, rowNo, dstWarekiCol, dstYearCol) = targetYear',s)
 def test_diff_export_maps_are_scoped(self):
  s=proc(MOD,'双方向編集_項目差異一覧を抽出')
  self.assertIn('BuildStep2IdentityIndex wsDst, dstLastRow, dstLastCol, targetYearValue',s)
  self.assertIn('BidirectionalDiffBuildDestinationAllNameMapCached wsDst, dstLastRow, destHeaderMap, targetYearValue',s)
 def test_existing_name_key_keeps_creation_year(self):
  s=proc(REVIEW,'ReviewName')
  for token in ('作成年度（和暦）','第1ガイド','第2ガイド','個別フォルダー'):
   self.assertIn(token,s)
 def test_scope_examples(self):
  def candidates(year,rows): return [row for row in rows if row['fiscal']==year]
  rows=[{'fiscal':2026,'created':2025},{'fiscal':2025,'created':2026},{'fiscal':2027,'created':2026},{'fiscal':None,'created':2026}]
  self.assertEqual(candidates(2026,rows),[rows[0]])
