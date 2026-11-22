/* 1회용: 빈강의실_DB_최종 2 반영 (rooms M열, survey 통합본, 과거_사용인원 탭, 설명)
 * 다른 탭과 운영 기록(rentals, visits, reports, reviews 등)은 건드리지 않음.
 * 실행 후 이 파일은 지워도 됨. */
const MIGRATE2_SRC_ID = '1hevlvENEo6u7l9TF0jzM7AHJCfbsfY12i0cMM_SxuIc';

function applyDbFinal2() {
  const src = SpreadsheetApp.openById(MIGRATE2_SRC_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // 새 탭 만들기 (survey 바로 뒤)
  if (!ss.getSheetByName('과거_사용인원')) {
    const sv = ss.getSheetByName('survey');
    ss.insertSheet('과거_사용인원', sv.getIndex());
    log.push('새 탭: 과거_사용인원');
  }

  // 수식이 참조하는 survey부터 교체
  ['survey', '과거_사용인원', 'rooms', '설명'].forEach(function (n) {
    copySheet2_(src.getSheetByName(n), ss.getSheetByName(n));
    log.push('교체: ' + n);
  });

  SpreadsheetApp.flush();
  Logger.log(log.join('\n'));
}

function copySheet2_(s, t) {
  const rows = Math.max(s.getLastRow(), 1), cols = Math.max(s.getLastColumn(), 1);
  t.clear();
  t.clearConditionalFormatRules();
  t.getRange(1, 1, t.getMaxRows(), t.getMaxColumns()).clearDataValidations();
  if (t.getMaxRows() < rows) t.insertRowsAfter(t.getMaxRows(), rows - t.getMaxRows());
  if (t.getMaxColumns() < cols) t.insertColumnsAfter(t.getMaxColumns(), cols - t.getMaxColumns());

  const sr = s.getRange(1, 1, rows, cols), tr = t.getRange(1, 1, rows, cols);
  const values = sr.getValues(), formulas = sr.getFormulas();
  const merged = values.map(function (r, i) { return r.map(function (v, j) { return formulas[i][j] || v; }); });
  tr.setNumberFormats(sr.getNumberFormats());
  tr.setValues(merged);
  tr.setFontWeights(sr.getFontWeights());
  tr.setFontColors(sr.getFontColors());
  tr.setBackgrounds(sr.getBackgrounds());
  tr.setHorizontalAlignments(sr.getHorizontalAlignments());
  tr.setWraps(sr.getWraps());
  try { tr.setDataValidations(sr.getDataValidations()); } catch (e) { Logger.log('검증 규칙 복사 실패(' + t.getName() + '): ' + e); }
  for (let c = 1; c <= cols; c++) t.setColumnWidth(c, s.getColumnWidth(c));
  t.setFrozenRows(s.getFrozenRows());
  t.setFrozenColumns(s.getFrozenColumns());
  try {
    t.setConditionalFormatRules(s.getConditionalFormatRules().map(function (rule) {
      return rule.copy().setRanges(rule.getRanges().map(function (rg) { return t.getRange(rg.getA1Notation()); })).build();
    }));
  } catch (e) { Logger.log('조건부 서식 복사 실패(' + t.getName() + '): ' + e); }
  if (s.getTabColor()) t.setTabColor(s.getTabColor());
}
