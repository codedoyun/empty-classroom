/***************************************************************
 * 빈 강의실 찾기 - 서버 코드 (Code.gs) 최종본
 * 스프레드시트 탭: rooms, timetable, events, checkins, reports, visits,
 *                  reviews, survey (+ rentals는 없으면 자동 생성)
 * 화면: GitHub Pages(index.html) → doPost API / 이 주소로 직접 열면 → Index.html
 ***************************************************************/

/* ===== 설정 (필요하면 여기만 바꾸세요) ===== */
const MAP_FILE_ID = '여기에_지도_이미지_파일_ID';  // 구글 드라이브에 올린 campus_map.jpg의 파일 ID (GitHub에는 비워 둠)
const OPEN_START = 8 * 60;        // 학교 개방 시작 08:00
const OPEN_END = 22 * 60;         // 학교 개방 종료 22:00
const CHECKIN_MINUTES = 90;       // 체크인 자동 만료(분)
const REPORT_MINUTES = 30;        // 신고가 화면에 표시되는 시간(분)
const SLOT_MINUTES = 50;          // 조회 시간 칸 길이 (정각 ~ 50분)
const DAYS = '일월화수목금토';
const RENTAL_FIRST_HOUR = 9;      // 대여 가능 첫 시간 칸 9:00~9:50
const RENTAL_LAST_HOUR = 21;      // 대여 가능 마지막 시간 칸 21:00~21:50
const RENTAL_MAX_HOURS = 3;       // 한 번에 최대 3시간(3칸)

const SS = SpreadsheetApp.getActiveSpreadsheet();

/* ===== 웹 페이지 열기 ===== */
function doGet(e) {
  let room = (e && e.parameter && e.parameter.room) || '';
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(room)) room = '';   // 이상한 값 차단
  try { SS.getSheetByName('visits').appendRow([new Date(), room]); } catch (err) {}

  const t = HtmlService.createTemplateFromFile('Index');
  t.room = room;
  t.mapUrl = mapDataUrl_();
  return t.evaluate()
    .setTitle('빈 강의실 찾기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* 드라이브의 지도 이미지를 페이지에 바로 넣을 수 있는 형태로 변환 */
function mapDataUrl_() {
  try {
    const blob = DriveApp.getFileById(MAP_FILE_ID).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (err) {
    return '';
  }
}

/* ===== 시트 읽기 도우미 ===== */
// 1행 헤더 이름으로 값을 꺼낼 수 있게 객체 배열로 변환 (열 순서가 바뀌어도 동작)
function table_(name) {
  const sheet = SS.getSheetByName(name);
  if (!sheet) throw new Error('"' + name + '" 탭을 찾을 수 없어요. 탭 이름을 확인하세요.');
  const values = sheet.getDataRange().getValues();
  const head = values.shift().map(function (h) { return String(h).trim(); });
  return values.map(function (r, i) {
    const o = { _row: i + 2 };
    head.forEach(function (h, j) {
      o[h] = r[j];
      o[h.replace(/\s/g, '')] = r[j];   // '평소 혼잡도'와 '평소혼잡도' 둘 다로 찾을 수 있게
    });
    return o;
  });
}

function toMin_(v) {
  // 시트가 '14:00'을 시간 형식으로 바꿔 저장했어도 시트에 보이는 값 그대로 읽음
  if (v instanceof Date) v = Utilities.formatDate(v, SS.getSpreadsheetTimeZone(), 'HH:mm');
  const p = String(v).trim().split(':');
  return Number(p[0]) * 60 + Number(p[1] || 0);
}

function fmt_(min) {
  return ('0' + Math.floor(min / 60)).slice(-2) + ':' + ('0' + (min % 60)).slice(-2);
}

function dateStr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v).trim();
}

const id_ = function (v) { return String(v).trim(); };
const txt_ = function (v) { return v === undefined || v === null ? '' : String(v).trim(); };

/* ===== 현장조사 기록(survey 탭) → 교시별 과거 사용 인원 =====
 * survey 헤더: 강의실 | 요일 | 조사시각 | 인원 | 상태
 * 조사시각의 '시'로 교시를 정함: 12:03, 12:40, 12:55 → 모두 12:00~12:50 교시
 * 상태가 '잠김'이면 인원 대신 "잠겨 있었어요"로 표시
 * 인원도 없고 잠김도 아닌 기록(수업중 등)과 시각이 없는 기록은 건너뜀
 */
function pastUsage_(day) {
  const map = {};
  const sheet = SS.getSheetByName('survey');
  if (!sheet || sheet.getLastRow() < 2) return map;

  const values = sheet.getDataRange().getDisplayValues();   // 화면에 보이는 글자 그대로
  const head = values.shift().map(function (h) { return String(h).replace(/\s/g, ''); });
  const cRoom = head.indexOf('강의실'), cDay = head.indexOf('요일'), cTime = head.indexOf('조사시각');
  const cN = head.indexOf('인원'), cSt = head.indexOf('상태');
  if (cRoom < 0 || cTime < 0 || cN < 0) return map;

  values.forEach(function (r) {
    const id = id_(r[cRoom]);
    const m = String(r[cTime]).match(/(\d{1,2})\s*[:시]/);
    if (!id || !m) return;
    const locked = cSt >= 0 && String(r[cSt]).indexOf('잠') >= 0;
    const raw = String(r[cN]).replace(/[^\d]/g, '');
    if (!locked && raw === '') return;

    const key = id + '|' + Number(m[1]);
    if (!map[key]) map[key] = { same: [], all: [] };
    const rec = { locked: locked, n: locked ? 0 : Number(raw) };
    map[key].all.push(rec);
    const d = cDay >= 0 ? String(r[cDay]).trim() : '';
    if (d && d.charAt(0) === String(day).charAt(0)) map[key].same.push(rec);
  });
  return map;
}

function pastText_(past, id, hour) {
  const rec = past[id + '|' + hour];
  if (!rec) return '과거 기록 없음';
  const arr = rec.same.length ? rec.same : rec.all;   // 같은 요일 기록이 있으면 그것만 사용
  const open = arr.filter(function (x) { return !x.locked; });
  if (!open.length) return '과거 이 시간엔 잠겨 있었어요';
  const avg = Math.round(open.reduce(function (s, x) { return s + x.n; }, 0) / open.length);
  return '과거 이 시간엔 ' + avg + '명 사용중';
}

/* 그 요일에 이 강의실의 마지막 수업이 끝나는 시각 (수업이 없으면 -1) */
function lastClassEnd_(id, classes) {
  const ends = classes
    .filter(function (c) { return id_(c['강의실']) === id; })
    .map(function (c) { return toMin_(c['종료']); });
  return ends.length ? Math.max.apply(null, ends) : -1;
}

/* ===== 강의실 상태 조회 =====
 * code: 건물 코드 (N, M, K, BA, V, D, B, C, SH, NP, CH)
 * day : '월' ~ '일'
 * time: 'HH:MM' (기본: 그 시각부터 50분 동안을 조회. 예: 13:00 → 13:00~13:50)
 */
// exact가 true면 시간 칸(50분)이 아니라 그 시각 '한 순간'의 상태를 조회 (현재 시간 버튼용)
function getRooms(code, day, time, exact) {
  return rooms_(code, day, toMin_(time), exact ? 1 : SLOT_MINUTES);
}

// len: 조회할 시간 길이(분). 시간 칸 조회는 50, QR 화면의 '지금'은 1
function rooms_(code, day, min, len) {
  if (!/^[A-Z]{1,3}$/.test(code)) throw new Error('건물 코드가 올바르지 않아요.');
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayDay = DAYS[now.getDay()];
  const nowMin = now.getHours() * 60 + now.getMinutes();
  const end = min + len;

  // 오늘 + 지금 시간대를 고른 경우에만 체크인·신고(실시간 정보)를 반영
  const isLive = day === todayDay && Math.floor(min / 60) === Math.floor(nowMin / 60);
  const closed = min < OPEN_START || min >= OPEN_END;

  const re = new RegExp('^' + code + '\\d');
  const rooms = table_('rooms').filter(function (r) { return re.test(id_(r['강의실'])); });
  const ids = {};
  rooms.forEach(function (r) { ids[id_(r['강의실'])] = true; });

  const classes = table_('timetable').filter(function (c) {
    return ids[id_(c['강의실'])] && String(c['요일']).trim() === day;
  });
  // 고른 요일의 가장 가까운 날짜(오늘 포함)의 일정·대여를 반영
  const target = dateForDay_(day, now, tz);
  const events = table_('events').filter(function (e) {
    return ids[id_(e['강의실'])] && dateStr_(e['날짜'], tz) === target;
  });
  // 학생 대여(선착순)는 '사용중'으로 보여주고, 대여 인원을 현재 인원에 더함
  const rentals = activeRentals_(target, tz).filter(function (r) { return ids[r['강의실']]; });
  const checkins = isLive ? table_('checkins').filter(function (c) {
    return ids[id_(c['강의실'])] && new Date(c['만료시각']) > now;
  }) : [];
  const reports = isLive ? table_('reports').filter(function (r) {
    return ids[id_(r['강의실'])] && now - new Date(r['신고시각']) < REPORT_MINUTES * 60000;
  }) : [];
  const past = pastUsage_(day);
  const hour = Math.floor(min / 60);

  const list = rooms.map(function (r) {
    const id = id_(r['강의실']);
    const rentNow = rentals.filter(function (x) {
      return x['강의실'] === id && toMin_(x['시작']) < end && toMin_(x['종료']) > min;
    });
    // '이미 사용 중' 신고는 사용 불가가 아니라 인원 정보로 씀 (rooms E열 수식처럼 체크인 합계와 신고 인원 중 큰 값)
    const myReports = reports.filter(function (x) { return id_(x['강의실']) === id; });
    const busy = myReports.filter(function (x) { return txt_(x['신고유형']) === '이미 사용 중'; });
    const blockReports = myReports.filter(function (x) { return txt_(x['신고유형']) !== '이미 사용 중'; });
    const checkinPeople = checkins
      .filter(function (c) { return id_(c['강의실']) === id; })
      .reduce(function (s, c) { return s + (Number(c['인원']) || 1); }, 0);
    const reportPeople = busy.length ? Number(busy[busy.length - 1]['사용인원(명)']) || 0 : 0;
    const people = Math.max(checkinPeople, reportPeople)
      + rentNow.reduce(function (s, x) { return s + (Number(x['인원']) || 1); }, 0);
    const limited = txt_(r['이용제한']).toUpperCase() === 'Y';
    const st = status_(id, min, end, closed, events, classes, blockReports, people, rentals, rentNow, limited);
    return {
      id: id,
      capacity: r['수용인원'] === '' ? '' : Number(r['수용인원']),
      people: people,
      status: st.status,
      color: st.color,
      // 비었음·사용중 카드는 "문 닫을 때까지" 대신 현장조사 기준 과거 사용 인원을 보여줌
      note: st.past ? pastText_(past, id, hour) : (st.note || ''),
      // 팝업 세부정보용 (시트에 비어 있으면 '')
      building: txt_(r['건물']),
      type: txt_(r['강의실유형']),
      outlets: txt_(r['콘센트개수']),
      board: txt_(r['칠판']),
      crowd: txt_(r['평소혼잡도'])
    };
  });

  const order = { green: 0, orange: 1, red: 2, gray: 3 };
  list.sort(function (a, b) { return order[a.color] - order[b.color] || a.id.localeCompare(b.id); });

  return { rooms: list, isLive: isLive, day: day, time: len === 1 ? fmt_(min) + ' (현재 시각)' : fmt_(min) + '~' + fmt_(end), closed: closed };
}

/* 상태 판단 순서:
 * 개방시간 → 이용 제한 → 시험/대여 일정 → 수업 → 신고 → 학생 대여
 * → 마지막 수업 종료 후(사용 불가) → 체크인(사용중) → 비었음
 * [min, end) 구간과 조금이라도 겹치는 수업·일정이 있으면 그 상태로 표시 */
function status_(id, min, end, closed, events, classes, reports, people, rentals, rentNow, limited) {
  if (closed) return { status: '개방 시간 아님', color: 'gray' };
  if (limited) return { status: '사용 불가', color: 'red', note: '이용 제한 강의실' };

  const ev = events.filter(function (e) {
    return id_(e['강의실']) === id && toMin_(e['시작']) < end && toMin_(e['종료']) > min;
  })[0];
  if (ev) {
    const name = String(ev['상태명']);
    if (name.indexOf('대여') >= 0) return { status: '대여중', color: 'red', note: fmt_(toMin_(ev['종료'])) + '까지' };
    return { status: '시험중', color: 'red', note: fmt_(toMin_(ev['종료'])) + '까지' };
  }

  const cls = classes.filter(function (c) {
    return id_(c['강의실']) === id && toMin_(c['시작']) < end && toMin_(c['종료']) > min;
  })[0];
  if (cls) return { status: '수업중', color: 'gray', note: fmt_(toMin_(cls['종료'])) + '까지' };

  const rep = reports.filter(function (r) { return id_(r['강의실']) === id; }).pop();
  if (rep) {
    const why = txt_(rep['상세내용']) || txt_(rep['신고유형']) || txt_(rep['사유']) || '신고 접수';
    return { status: '사용 불가', color: 'red', note: why };
  }

  if (rentNow && rentNow.length) {
    const last = Math.max.apply(null, rentNow.map(function (x) { return toMin_(x['종료']); }));
    return { status: '사용중', color: 'orange', note: '대여 ' + fmt_(last) + '까지' };
  }

  // 그날 마지막 수업이 끝난 뒤에는 강의실이 잠기므로 사용 불가 (그날 수업이 없는 강의실은 해당 없음)
  const lastEnd = lastClassEnd_(id, classes);
  if (lastEnd >= 0 && min >= lastEnd) {
    return { status: '사용 불가', color: 'red', note: '오늘 수업 종료 (' + fmt_(lastEnd) + ')' };
  }

  if (people > 0) return { status: '사용중', color: 'orange', past: true };
  return { status: '비었음', color: 'green', past: true };
}

/* 요일 → 오늘부터 가장 가까운 그 요일의 날짜 (예: 오늘 목요일, '월' → 다음 주 월요일) */
function dateForDay_(day, now, tz) {
  const off = (DAYS.indexOf(day) - now.getDay() + 7) % 7;
  return Utilities.formatDate(new Date(now.getTime() + off * 86400000), tz, 'yyyy-MM-dd');
}

/* ===== 대여 (선착순) ===== */
// rentals 탭이 없으면 자동으로 만듦
function rentalSheet_() {
  let sh = SS.getSheetByName('rentals');
  if (!sh) {
    sh = SS.insertSheet('rentals');
    sh.getRange('B:D').setNumberFormat('@');   // 날짜·시간은 글자 그대로 저장
    sh.appendRow(['강의실', '날짜', '시작', '종료', '인원', '목적', '신청시각', '취소', '확인코드']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 특정 날짜의 취소되지 않은 대여를 events와 같은 모양으로 반환
function activeRentals_(date, tz) {
  rentalSheet_();
  return table_('rentals').filter(function (r) {
    return dateStr_(r['날짜'], tz) === date && String(r['취소']).trim() === '';
  }).map(function (r) {
    return { '강의실': id_(r['강의실']), '시작': r['시작'], '종료': r['종료'], '인원': r['인원'], '상태명': '대여 중' };
  });
}

/* 대여 신청 (오늘만): 그 시간에 수업·일정·다른 대여가 없으면 바로 등록 (선착순) */
function rent(id, startHour, hours, people) {
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(id)) throw new Error('강의실 번호가 올바르지 않아요.');
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const date = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const day = DAYS[now.getDay()];
  startHour = Number(startHour); hours = Number(hours);
  if (!(startHour >= RENTAL_FIRST_HOUR && hours >= 1 && hours <= RENTAL_MAX_HOURS &&
        startHour + hours - 1 <= RENTAL_LAST_HOUR)) throw new Error('대여 시간이 올바르지 않아요.');
  people = Math.min(200, Math.max(1, Number(people) || 1));

  const start = startHour * 60;
  const end = (startHour + hours) * 60 - 10;          // 예: 14시부터 2시간 → 14:00~15:50
  if (end <= now.getHours() * 60 + now.getMinutes()) throw new Error('이미 지난 시간이에요.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const roomRow = table_('rooms').filter(function (r) { return id_(r['강의실']) === id; })[0];
    if (roomRow && txt_(roomRow['이용제한']).toUpperCase() === 'Y') throw new Error('이용 제한 강의실이라 대여할 수 없어요.');

    // 겹치는 수업·일정·대여 확인
    const overlap = function (x) { return toMin_(x['시작']) < end && toMin_(x['종료']) > start; };
    const dayClasses = table_('timetable').filter(function (c) {
      return id_(c['강의실']) === id && String(c['요일']).trim() === day;
    });
    const cls = dayClasses.filter(overlap)[0];
    if (cls) throw new Error('그 시간에는 수업이 있어요 (' + fmt_(toMin_(cls['시작'])) + '~' + fmt_(toMin_(cls['종료'])) + ').');
    const lastEnd = lastClassEnd_(id, dayClasses);
    if (lastEnd >= 0 && start >= lastEnd) throw new Error('오늘 수업이 ' + fmt_(lastEnd) + '에 끝나 잠기는 강의실이라 대여할 수 없어요.');
    const ev = table_('events').filter(function (e) {
      return id_(e['강의실']) === id && dateStr_(e['날짜'], tz) === date && overlap(e);
    }).concat(activeRentals_(date, tz).filter(function (r) { return r['강의실'] === id && overlap(r); }))[0];
    if (ev) throw new Error('이미 대여나 시험 일정이 있어요 (' + fmt_(toMin_(ev['시작'])) + '~' + fmt_(toMin_(ev['종료'])) + ').');

    const token = Utilities.getUuid().slice(0, 8);
    const sh = rentalSheet_();
    const row = sh.getLastRow() + 1;
    const range = sh.getRange(row, 1, 1, 9);
    // 날짜·시간이 시트에서 자동 변환되지 않도록 글자 형식으로 먼저 지정한 뒤 저장
    range.setNumberFormats([['@', '@', '@', '@', '0', '@', 'yyyy-MM-dd HH:mm:ss', '@', '@']]);
    range.setValues([[id, date, fmt_(start), fmt_(end), people, '', now, '', token]]);   // 목적 칸은 비워 둠
    SpreadsheetApp.flush();   // 바로 다음 조회에 반영되도록 즉시 저장
    return { row: row, token: token, id: id, date: date, start: fmt_(start), end: fmt_(end) };
  } finally {
    lock.releaseLock();
  }
}

/* 대여 취소: 신청한 기기에 저장된 확인코드가 맞아야 취소됨 */
function cancelRental(row, token) {
  const sh = rentalSheet_();
  row = Number(row);
  if (!row || row < 2 || row > sh.getLastRow()) return false;
  const vals = sh.getRange(row, 1, 1, 9).getValues()[0];
  if (String(vals[8]) !== String(token) || String(vals[7]).trim() !== '') return false;
  sh.getRange(row, 8).setValue('Y');
  return true;
}

/* 나의 사용 상태 확인: 기기에 저장된 대여들이 아직 유효한지(관리자가 취소했는지) 알려줌 */
function rentalStatus(items) {
  const values = rentalSheet_().getDataRange().getValues();
  return (items || []).map(function (it) {
    const v = values[Number(it.row) - 1];
    if (!v || Number(it.row) < 2 || String(v[8]) !== String(it.token)) return 'missing';
    return String(v[7]).trim() === '' ? 'active' : 'cancelled';
  });
}

/* QR로 들어온 강의실 1개의 현재 상태 */
function getRoom(id) {
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(id)) throw new Error('강의실 번호가 올바르지 않아요.');
  const now = new Date();
  const code = id.match(/^[A-Z]+/)[0];
  const res = rooms_(code, DAYS[now.getDay()], now.getHours() * 60 + now.getMinutes(), 1);
  return res.rooms.filter(function (r) { return r.id === id; })[0] || null;
}

/* 신고 화면에서 쓸 건물별 강의실 목록 */
function listRooms(code) {
  if (!/^[A-Z]{1,3}$/.test(code)) return [];
  const re = new RegExp('^' + code + '\\d');
  return table_('rooms').map(function (r) { return id_(r['강의실']); })
    .filter(function (id) { return re.test(id); }).sort();
}

/* ===== 체크인 / 퇴실 / 신고 ===== */
function checkIn(id, people) {
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(id)) throw new Error('강의실 번호가 올바르지 않아요.');
  people = Math.min(10, Math.max(1, Number(people) || 1));
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SS.getSheetByName('checkins');
    const now = new Date();
    const expire = new Date(now.getTime() + CHECKIN_MINUTES * 60000);
    sheet.appendRow([id, now, expire, people]);
    return { row: sheet.getLastRow(), expire: fmt_(expire.getHours() * 60 + expire.getMinutes()) };
  } finally {
    lock.releaseLock();
  }
}

function checkOut(id, row) {
  const sheet = SS.getSheetByName('checkins');
  row = Number(row);
  if (!row || row < 2 || row > sheet.getLastRow()) return false;
  if (id_(sheet.getRange(row, 1).getValue()) !== id) return false;   // 다른 강의실 기록은 건드리지 않음
  sheet.getRange(row, 3).setValue(new Date());                      // 만료시각을 지금으로
  return true;
}

/* 화면의 신고 버튼 → reports 탭 (강의실 | 신고시각 | 상세 내용 | 신고 유형 | 사용 인원(명)) */
function report(id, reason) {
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(id)) throw new Error('강의실 번호가 올바르지 않아요.');
  reason = String(reason || '').trim().slice(0, 40) || '기타';
  let type = '기타';
  if (reason.indexOf('잠겨') >= 0) type = '잠김';
  else if (reason.indexOf('망가') >= 0 || reason.indexOf('시설') >= 0) type = '파손';
  else if (reason.indexOf('시험') >= 0 || reason.indexOf('수업') >= 0 || reason.indexOf('행사') >= 0) type = '시간표 불일치';
  SS.getSheetByName('reports').appendRow([id, new Date(), reason, type, '']);
  return true;
}

/* ===============================================================
 * 추가 기능 (접속 기록 · 유형별 신고 · 후기/이벤트 · 추첨 · 개인정보 파기)
 * GitHub 화면(index.html)은 submitReport·submitReview를 아래 API로 호출하고,
 * Apps Script 화면(Index.html)은 위의 report()를 그대로 써요.
 * =============================================================== */
const REPORT_TYPES = ['위험', '파손', '잠김', '시간표 불일치', '이미 사용 중', '기타'];

/* 1) 접속 기록 (중복 허용) — doGet에서 이미 기록하므로, 강의실 상세를 열 때만 따로 호출 */
function logVisit(room) {
  SS.getSheetByName('visits').appendRow([new Date(), String(room || '').slice(0, 20)]);
}

/* 2) 유형별 신고 (사용 인원 신고 포함)
 * form = { room, type, detail, count } / type이 '이미 사용 중'이면 count 필수 */
function submitReport(form) {
  form = form || {};
  const room = String(form.room || '').trim();
  const type = String(form.type || '').trim();
  const detail = String(form.detail || '').trim().slice(0, 300);
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(room)) throw new Error('강의실을 선택해 주세요.');
  if (type === '기타' && !detail) throw new Error('기타 사유를 적어 주세요.');
  if (REPORT_TYPES.indexOf(type) === -1) throw new Error('신고 유형을 선택해 주세요.');

  let count = '';
  if (type === '이미 사용 중') {
    count = parseInt(form.count, 10);
    if (!(count >= 1 && count <= 300)) throw new Error('사용 인원을 1명 이상으로 입력해 주세요.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SS.getSheetByName('reports').appendRow([room, new Date(), detail, type, count]);
  } finally {
    lock.releaseLock();
  }
  return '신고가 접수됐어요. 감사합니다!';
}

/* 3) 후기 + 이벤트 응모
 * form = { room, rating, text, entry('Y'|'N'), name, phone, dept, consent('Y'|'N') }
 * 이름·연락처·학과는 이벤트 응모 + 개인정보 동의일 때만 저장 */
function submitReview(form) {
  form = form || {};
  const room = String(form.room || '').trim();
  const rating = parseInt(form.rating, 10);
  const text = String(form.text || '').trim().slice(0, 500);
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(room)) throw new Error('강의실 정보가 없어요.');
  if (!(rating >= 1 && rating <= 5)) throw new Error('별점을 선택해 주세요.');
  if (!text) throw new Error('후기 내용을 입력해 주세요.');

  let entry = 'N', name = '', phone = '', dept = '', consent = 'N';
  if (form.entry === 'Y') {
    if (form.consent !== 'Y') throw new Error('이벤트 응모를 하려면 개인정보 수집에 동의해 주세요.');
    name = String(form.name || '').trim().slice(0, 20);
    phone = String(form.phone || '').replace(/\s/g, '');
    dept = String(form.dept || '').trim().slice(0, 30);
    if (!name || !dept) throw new Error('이름과 학과를 입력해 주세요.');
    if (!/^01[016789]-?\d{3,4}-?\d{4}$/.test(phone)) throw new Error('연락처 형식을 확인해 주세요. 예) 010-1234-5678');
    entry = 'Y'; consent = 'Y';
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SS.getSheetByName('reviews').appendRow([room, new Date(), rating, text, entry, name, phone, dept, consent]);
  } finally {
    lock.releaseLock();
  }
  return entry === 'Y' ? '후기 제출과 이벤트 응모가 완료됐어요!' : '후기가 제출됐어요. 감사합니다!';
}

/* 4) 추첨 (관리자가 편집기에서 runDraw 선택 → 실행) — 결과는 '추첨결과' 시트 */
function runDraw() {
  const WINNERS = 3; // 당첨 인원
  const out = SS.getSheetByName('추첨결과') || SS.insertSheet('추첨결과');
  if (out.getLastRow() > 0) throw new Error('추첨결과가 이미 있어요. 재추첨하려면 시트 내용을 지우세요.');

  const rows = SS.getSheetByName('reviews').getDataRange().getValues().slice(1);
  const seen = {}, pool = [];
  rows.forEach(function (r) {
    const phone = String(r[6]).replace(/-/g, '');
    if (r[4] === 'Y' && r[8] === 'Y' && phone && !seen[phone]) { seen[phone] = true; pool.push(r); }
  });
  for (let i = pool.length - 1; i > 0; i--) {          // Fisher–Yates 셔플
    const j = Math.floor(Math.random() * (i + 1));
    const t = pool[i]; pool[i] = pool[j]; pool[j] = t;
  }
  out.appendRow(['추첨일시', new Date(), '응모자 수(중복 제외)', pool.length]);
  out.appendRow(['순번', '이름', '연락처', '학과', '강의실', '후기 작성시각']);
  pool.slice(0, WINNERS).forEach(function (r, k) { out.appendRow([k + 1, r[5], r[6], r[7], r[0], r[1]]); });
}

/* 5) 이벤트 종료 후 개인정보 파기 (관리자 실행) — reviews F~H열과 추첨결과 시트를 지움 */
function purgePersonalInfo() {
  const rv = SS.getSheetByName('reviews');
  const n = rv.getLastRow() - 1;
  if (n > 0) rv.getRange(2, 6, n, 3).clearContent();
  const out = SS.getSheetByName('추첨결과');
  if (out) out.clear();
}

/* ===== 외부 화면(GitHub Pages)용 API =====
 * 화면이 fetch로 { fn: '함수이름', args: [...] }를 보내면 그 함수를 실행해 JSON으로 돌려줌.
 * 여기 적힌 함수만 바깥에서 부를 수 있음 (추첨·개인정보 파기 같은 관리자 함수는 차단). */
const API = {
  getRooms: getRooms,
  getRoom: getRoom,
  listRooms: listRooms,
  checkIn: checkIn,
  checkOut: checkOut,
  report: report,              // 예전 화면 호환용
  submitReport: submitReport,
  submitReview: submitReview,
  rent: rent,
  cancelRental: cancelRental,
  rentalStatus: rentalStatus,
  visit: visit_
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (!Object.prototype.hasOwnProperty.call(API, req.fn)) throw new Error('알 수 없는 요청이에요.');
    out = { ok: true, data: API[req.fn].apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* 방문 기록 (외부 화면이 처음 열릴 때 한 번 호출) */
function visit_(room) {
  room = String(room || '');
  if (!/^[A-Z]{1,3}\d{2,4}$/.test(room)) room = '';
  SS.getSheetByName('visits').appendRow([new Date(), room]);
  return true;
}

/* ===== 편집기에서 테스트용 ===== */
function test() {
  Logger.log(JSON.stringify(getRooms('N', '금', '17:00'), null, 2));
}
