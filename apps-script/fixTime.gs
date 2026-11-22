function fixTimeText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  [['timetable', 3, 2], ['survey', 4, 1]].forEach(function (t) {
    const sh = ss.getSheetByName(t[0]);
    const n = sh.getLastRow() - 1;
    if (n < 1) return;
    const rng = sh.getRange(2, t[1], n, t[2]);
    const shown = rng.getDisplayValues().map(function (r) {
      return r.map(function (v) { return String(v).trim().replace(/^(\d):/, '0$1:'); });   // 9:00 → 09:00
    });
    rng.setNumberFormat('@').setValues(shown);
    log.push(t[0] + ': ' + n + '줄 글자로 변환 (예: ' + shown[0].join(', ') + ')');
  });
  SpreadsheetApp.flush();
  Logger.log(log.join('\n'));
}
