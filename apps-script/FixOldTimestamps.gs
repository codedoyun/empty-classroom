/* 1회용: 시간대를 서울로 바꾸기 전에 쌓인 기록 시각(미국 태평양 시간으로 적힌 것)을 한국 시간으로 고침.
 * 바꾸기 전 기록은 '2026-11-05 10:00' 이전 시각으로 찍혀 있고, 바꾼 뒤 기록은 그 이후라서 둘을 구분할 수 있음.
 * 한 번만 실행하세요 (두 번 실행하면 또 밀림). 실행 후 이 파일은 지워도 됨. */
function fixOldTimestamps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const FMT = 'yyyy-MM-dd HH:mm:ss';
  const CUTOFF = '2026-11-05 10:00:00';
  const targets = [['visits', [1]], ['reports', [2]], ['checkins', [2, 3]], ['reviews', [2]], ['rentals', [7]]];
  const log = [];
  targets.forEach(function (t) {
    const sh = ss.getSheetByName(t[0]);
    if (!sh || sh.getLastRow() < 2) return;
    t[1].forEach(function (col) {
      const rng = sh.getRange(2, col, sh.getLastRow() - 1, 1);
      const vals = rng.getValues();
      let n = 0;
      vals.forEach(function (r) {
        if (!(r[0] instanceof Date)) return;
        const wall = Utilities.formatDate(r[0], 'Asia/Seoul', FMT);   // 시트에 적혀 있는 그대로의 시각
        if (wall >= CUTOFF) return;                                    // 시간대 바꾼 뒤 기록은 그대로
        r[0] = Utilities.parseDate(wall, 'America/Los_Angeles', FMT);  // 그 시각을 미국 시간으로 보고 실제 시점 계산
        n++;
      });
      rng.setValues(vals);
      log.push(t[0] + ' ' + col + '열: ' + n + '개 고침');
    });
  });
  Logger.log(log.join('\n'));
}
