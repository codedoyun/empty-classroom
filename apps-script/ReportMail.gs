/***************************************************************
 * 공강 정류장 - 신고 자동 메일 (ReportMail.gs)
 *
 * 학생이 사이트에서 남긴 신고(reports 탭)를 학교 담당 부서로 자동 전송해요.
 *  - '위험' 신고        → 10분 안에 바로 전송
 *  - '시설 파손' 신고   → 매일 오전 9시에 하루치를 모아서 한 통
 *
 * 보내는 사람은 이 스크립트를 만든 "공강 정류장 운영 계정"이에요.
 * 신고한 학생의 메일·이름은 쓰지도, 저장하지도 않아요.
 *
 * 설치 위치: 운영 계정의 script.google.com → '새 프로젝트'(독립 실행형)
 *           사이트용 Apps Script(Code.gs)는 건드리지 않아요.
 * 설치 방법: 운영 매뉴얼 PDF의 '신고 자동 메일 설정' 참고
 *   1) 아래 SHEET_ID에 스프레드시트 ID를 넣는다
 *   2) 위 메뉴에서 setup 선택 → 실행 → 권한 허용
 *   3) 운영 계정 메일함에 '[테스트]' 메일이 왔는지 확인
 ***************************************************************/

/* ===== 설정 (필요하면 여기만 바꾸세요) ===== */
const SHEET_ID = '여기에_스프레드시트_ID';      // 스프레드시트 주소의 /d/ 와 /edit 사이 글자
const MAIL_TO = 'cukcap@catholic.ac.kr';        // 신고를 받는 학교 메일
const SENDER_NAME = '공강 정류장 (빈 강의실 찾기)'; // 받는 사람에게 보이는 보낸 사람 이름
const SITE_URL = 'https://codedoyun.github.io/empty-classroom/';
const TZ = 'Asia/Seoul';
const DAILY_HOUR = 9;            // 하루 모음 메일 보내는 시각 (오전 9시)
const URGENT_EVERY_MIN = 10;     // 위험 신고 확인 주기 (분): 1, 5, 10, 15, 30 중 하나
const MAX_MAILS_PER_DAY = 15;    // 하루 최대 메일 수 (장난 신고가 몰려도 학교에 메일이 쏟아지지 않게)
const MAX_ROWS_PER_MAIL = 40;    // 메일 한 통에 넣는 최대 신고 수
const SENT_HEADER = '메일발송';   // reports 탭 맨 오른쪽에 자동으로 생기는 열 (보낸 시각 기록)

/* ===== 어떤 신고를 보낼지 =====
 * 사이트 버튼 문구('위험해요', '시설이 망가졌어요')와 신고 유형('위험', '파손') 둘 다로 판단 */
function kindOf_(type, detail) {
  const t = String(type || '');
  const d = String(detail || '');
  if (t === '위험' || d.indexOf('위험') >= 0) return 'urgent';
  if (t === '파손' || d.indexOf('망가') >= 0 || d.indexOf('시설') >= 0 || d.indexOf('파손') >= 0) return 'daily';
  return '';   // 잠김·시험·수업 중·이미 사용 중 등은 사이트 화면에만 반영하고 메일은 보내지 않음
}

/* ===== 트리거가 부르는 함수 ===== */
// 10분마다: 아직 안 보낸 위험 신고가 있으면 바로 한 통
function sendUrgentReports() {
  sendKind_('urgent');
}
// 매일 오전 9시: 아직 안 보낸 시설 파손 신고를 모아서 한 통 (혹시 못 보낸 위험 신고도 함께)
function sendDailyReports() {
  sendKind_('daily');
}

/* ===== 처음 한 번 실행: 트리거 설치 + 기존 신고 정리 + 테스트 메일 ===== */
function setup() {
  if (SHEET_ID.indexOf('여기에') === 0) throw new Error('먼저 맨 위 SHEET_ID에 스프레드시트 ID를 넣어 주세요.');
  // 1) 이 프로젝트의 기존 트리거를 지우고 새로 설치 (여러 번 실행해도 중복되지 않게)
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendUrgentReports').timeBased().everyMinutes(URGENT_EVERY_MIN).create();
  ScriptApp.newTrigger('sendDailyReports').timeBased().atHour(DAILY_HOUR).everyDays(1).inTimezone(TZ).create();

  // 2) 설치 전에 쌓인 신고(테스트 신고 포함)는 학교로 보내지 않도록 '설치 전'으로 표시
  const r = readReports_();
  const marks = [];
  r.rows.forEach(function (row, i) { if (!row.sent) marks.push(i); });
  markSent_(r, marks, '설치 전 (보내지 않음)');

  // 3) 학교가 아니라 운영 계정 본인에게 테스트 메일
  const me = Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: me,
    name: SENDER_NAME,
    subject: '[테스트] 공강 정류장 신고 자동 메일 설치 완료',
    body: '설치가 끝났어요.\n\n' +
      '- 위험 신고: ' + URGENT_EVERY_MIN + '분마다 확인해서 바로 ' + MAIL_TO + ' 로 전송\n' +
      '- 시설 파손 신고: 매일 ' + DAILY_HOUR + '시에 모아서 전송\n' +
      '- 설치 전 신고 ' + marks.length + '건은 보내지 않음으로 표시했어요.\n\n' +
      '실제 메일 모양을 보려면 previewMail 을 실행하세요 (본인에게만 와요).'
  });
  Logger.log('설치 완료. 트리거 2개 설치, 설치 전 신고 ' + marks.length + '건 표시, 테스트 메일 → ' + me);
}

/* 학교에 보낼 메일 모양을 운영 계정 본인에게만 미리 보내 봄 (기록은 바꾸지 않음) */
function previewMail() {
  const r = readReports_();
  const sample = r.rows.slice(-5);
  if (!sample.length) throw new Error('reports 탭에 신고가 없어서 미리 볼 수 없어요.');
  const mail = buildMail_('urgent', sample);
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), name: SENDER_NAME,
    subject: '[미리보기] ' + mail.subject, body: mail.body, htmlBody: mail.html });
}

/* 자동 메일 끄기 (학교 요청이 있거나 운영을 멈출 때) */
function uninstall() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  Logger.log('자동 메일을 껐어요. 다시 켜려면 setup 을 실행하세요.');
}

/* ===== 내부 동작 ===== */
function sendKind_(kind) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;      // 다른 실행이 돌고 있으면 이번엔 건너뜀
  try {
    const r = readReports_();
    const picks = [];
    r.rows.forEach(function (row, i) {
      if (row.sent || !row.room) return;
      const k = kindOf_(row.type, row.detail);
      // 하루 모음 메일은 혹시 놓친 위험 신고도 함께 보냄
      if (k === kind || (kind === 'daily' && k === 'urgent')) picks.push(i);
    });
    if (!picks.length) return;
    if (!canSendToday_()) { Logger.log('오늘 메일 한도(' + MAX_MAILS_PER_DAY + '통)를 넘어서 다음에 보내요.'); return; }

    const batch = picks.slice(0, MAX_ROWS_PER_MAIL);
    const mail = buildMail_(kind, batch.map(function (i) { return r.rows[i]; }));
    MailApp.sendEmail({ to: MAIL_TO, name: SENDER_NAME, subject: mail.subject, body: mail.body, htmlBody: mail.html });
    countSent_();
    markSent_(r, batch, now_() + (kind === 'urgent' ? ' (즉시)' : ' (하루 모음)'));
  } finally {
    lock.releaseLock();
  }
}

function readReports_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('reports');
  if (!sheet) throw new Error('"reports" 탭을 찾을 수 없어요.');
  let lastCol = sheet.getLastColumn();
  const head = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).replace(/\s/g, ''); });
  let cSent = head.indexOf(SENT_HEADER);
  if (cSent < 0) {                        // 처음이면 맨 오른쪽에 '메일발송' 열을 만듦
    lastCol += 1;
    sheet.getRange(1, lastCol).setValue(SENT_HEADER);
    head.push(SENT_HEADER);
    cSent = lastCol - 1;
  }
  const col = function (name, fallback) { const i = head.indexOf(name); return i >= 0 ? i : fallback; };
  const cRoom = col('강의실', 0), cTime = col('신고시각', 1), cDetail = col('상세내용', 2), cType = col('신고유형', 3);
  const n = sheet.getLastRow() - 1;
  const values = n > 0 ? sheet.getRange(2, 1, n, lastCol).getValues() : [];
  return {
    sheet: sheet,
    sentCol: cSent + 1,
    rows: values.map(function (v) {
      return {
        room: String(v[cRoom] || '').trim(),
        time: v[cTime],
        detail: String(v[cDetail] || '').trim(),
        type: String(v[cType] || '').trim(),
        sent: String(v[cSent] || '').trim() !== ''
      };
    })
  };
}

function markSent_(r, indexes, text) {
  indexes.forEach(function (i) { r.sheet.getRange(i + 2, r.sentCol).setValue(text); });
  SpreadsheetApp.flush();
}

function buildMail_(kind, rows) {
  const urgent = rows.some(function (x) { return kindOf_(x.type, x.detail) === 'urgent'; });
  const today = Utilities.formatDate(new Date(), TZ, 'M월 d일');
  const subject = urgent && kind === 'urgent'
    ? '[공강 정류장] 강의실 위험 신고 ' + rows.length + '건 (' + roomsOf_(rows) + ')'
    : '[공강 정류장] ' + today + ' 강의실 시설 신고 ' + rows.length + '건';
  const label = function (x) { return kindOf_(x.type, x.detail) === 'urgent' ? '위험' : '시설 파손'; };
  const when = function (x) { return x.time instanceof Date ? Utilities.formatDate(x.time, TZ, 'M/d HH:mm') : String(x.time); };

  const intro = '안녕하세요. 교내 빈 강의실 찾기 서비스 「공강 정류장」입니다.\n' +
    '학생들이 사이트에서 남긴 강의실 ' + (kind === 'urgent' ? '위험' : '시설') + ' 신고를 자동으로 전달드립니다.';
  const foot = '이 메일은 자동으로 발송됐습니다. 신고자 개인정보는 수집하지 않습니다.\n' +
    '서비스: ' + SITE_URL + '\n' +
    '문의·수신 중단 요청: 이 메일에 회신해 주세요 (공강 정류장 운영팀).';

  const body = intro + '\n\n' + rows.map(function (x, k) {
    return (k + 1) + '. [' + label(x) + '] ' + x.room + ' · ' + when(x) + ' · ' + (x.detail || x.type || '-');
  }).join('\n') + '\n\n' + foot;

  const esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const td = 'style="padding:8px 10px;border-bottom:1px solid #e5e5e5;text-align:left"';
  const html = '<div style="font-family:sans-serif;font-size:14px;color:#222">' +
    '<p>' + esc(intro).replace(/\n/g, '<br>') + '</p>' +
    '<table style="border-collapse:collapse;min-width:420px">' +
    '<tr style="background:#2E3440;color:#fff"><th ' + td + '>구분</th><th ' + td + '>강의실</th><th ' + td + '>신고 시각</th><th ' + td + '>내용</th></tr>' +
    rows.map(function (x) {
      const red = label(x) === '위험' ? 'color:#C0392B;font-weight:bold' : '';
      return '<tr><td ' + td.replace('"', '"' + red + ';') + '>' + label(x) + '</td><td ' + td + '>' + esc(x.room) +
        '</td><td ' + td + '>' + esc(when(x)) + '</td><td ' + td + '>' + esc(x.detail || x.type || '-') + '</td></tr>';
    }).join('') +
    '</table><p style="color:#777;font-size:12px;margin-top:16px">' + esc(foot).replace(/\n/g, '<br>') + '</p></div>';
  return { subject: subject, body: body, html: html };
}

function roomsOf_(rows) {
  const ids = rows.map(function (x) { return x.room; }).filter(function (v, i, a) { return a.indexOf(v) === i; });
  return ids.length > 3 ? ids.slice(0, 3).join(', ') + ' 외' : ids.join(', ');
}

function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

// 하루 메일 수 세기 (스크립트 속성에 날짜별로 저장)
function canSendToday_() {
  const p = PropertiesService.getScriptProperties();
  return Number(p.getProperty('sent_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd')) || 0) < MAX_MAILS_PER_DAY;
}
function countSent_() {
  const p = PropertiesService.getScriptProperties();
  const key = 'sent_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  p.setProperty(key, String(Number(p.getProperty(key) || 0) + 1));
}
