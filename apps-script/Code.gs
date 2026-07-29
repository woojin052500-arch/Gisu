/**
 * 기수제 프로그램 신청 폼 — 신청서 수신용 Apps Script
 *
 * 설치 (5분)
 * 1. 구글 스프레드시트를 새로 만든다
 * 2. 확장 프로그램 > Apps Script 클릭
 * 3. 기본 코드를 지우고 이 파일 전체를 붙여넣는다
 * 4. 배포 > 새 배포 > 유형 '웹 앱'
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자   ← 반드시 이걸로. 아니면 신청서가 안 들어온다
 * 5. 배포 후 나오는 /exec 로 끝나는 URL을 복사해서
 *    apply.html 운영자 화면의 '웹훅 URL' 칸에 붙여넣는다
 *
 * 주의: 코드를 수정하면 반드시 '새 배포'를 다시 해야 반영된다.
 */

var SHEET_NAME = '신청자';       // apply.html — 프로그램 신청자
var WAITLIST_NAME = '사전등록';  // landing.html — 출시 알림 신청자

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);

    // 랜딩 페이지 사전 등록은 별도 시트로 보낸다
    if (payload.type === 'waitlist') return saveWaitlist_(payload);

    var sheet = getSheet_();
    var answers = payload.answers || {};

    // 첫 실행이거나 질문 항목이 바뀌면 헤더를 다시 맞춘다
    var fixed = ['접수시각', '프로그램', '기수', '참가비', '입금확인', '메모'];
    var keys = Object.keys(answers);
    var header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var needHeader = header.filter(String).length === 0;

    if (!needHeader) {
      for (var i = 0; i < keys.length; i++) {
        if (header.indexOf(keys[i]) === -1) { needHeader = true; break; }
      }
    }
    if (needHeader) {
      var newHeader = fixed.slice(0, 4).concat(keys).concat(['입금확인', '메모']);
      sheet.getRange(1, 1, 1, newHeader.length).setValues([newHeader]);
      sheet.getRange(1, 1, 1, newHeader.length)
        .setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      header = newHeader;
    }

    var row = header.map(function (h) {
      if (h === '접수시각') return formatKST_(payload.submittedAt);
      if (h === '프로그램') return payload.program || '';
      if (h === '기수') return payload.cohort ? payload.cohort + '기' : '';
      if (h === '참가비') return Number(payload.fee || 0);
      if (h === '입금확인') return '';   // 수동 체크용 빈 칸
      if (h === '메모') return '';
      return answers[h] != null ? answers[h] : '';
    });

    sheet.appendRow(row);

    // 신규 신청 알림 메일 (필요 없으면 아래 두 줄을 지운다)
    var name = answers['이름'] || '이름없음';
    notify_(payload, name);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * 신청 현황 조회.
 * 신청 페이지와 운영자 화면이 이 함수를 호출해 현재 신청자 수를 가져간다.
 * 브라우저에서 응답을 읽어야 하므로 callback 파라미터가 오면 JSONP 로 응답한다.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'count') {
    if (p.type === 'waitlist') return respond_(countWaitlist_(), p.callback);
    return respond_(countFor_(p.program || '', p.cohort || ''), p.callback);
  }
  return respond_({ ok: true, msg: 'alive' }, p.callback);
}

/* ── 랜딩 페이지 사전 등록 ────────────────────────── */

function saveWaitlist_(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(WAITLIST_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(WAITLIST_NAME);
    var header = ['등록시각', '이름', '전화번호', '운영 중인 프로그램', '유입 위치', '연락함', '메모'];
    sheet.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#2f4fe8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 130);
  }

  var phone = String(payload.phone || '').replace(/\D/g, '');

  // 같은 번호가 이미 있으면 중복 저장하지 않는다
  var last = sheet.getLastRow();
  if (last > 1) {
    var existing = sheet.getRange(2, 3, last - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).replace(/\D/g, '') === phone) {
        return json_({ ok: true, duplicated: true });
      }
    }
  }

  sheet.appendRow([
    formatKST_(payload.submittedAt),
    payload.name || '',
    "'" + (payload.phone || ''),   // 앞의 0이 잘리지 않도록 텍스트로 저장
    payload.kind || '',
    payload.from || '',
    '',                            // 연락함 — 수동 체크
    ''                             // 메모
  ]);

  try {
    var to = Session.getEffectiveUser().getEmail();
    if (to) {
      MailApp.sendEmail({
        to: to,
        subject: '[다음기수] 사전 등록 ' + (payload.name || '') + ' ' + (payload.phone || ''),
        body: '이름 : ' + (payload.name || '(미입력)')
          + '\n전화 : ' + (payload.phone || '')
          + '\n프로그램 : ' + (payload.kind || '(미선택)')
          + '\n유입 : ' + (payload.from || '')
      });
    }
  } catch (e) { /* 메일 실패해도 저장은 유지 */ }

  return json_({ ok: true });
}

function countWaitlist_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(WAITLIST_NAME);
  if (!sheet) return { ok: true, count: 0 };
  return { ok: true, count: Math.max(0, sheet.getLastRow() - 1) };
}

function countFor_(program, cohort) {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return { ok: true, count: 0, paid: 0 };

  var width = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, width).getValues()[0];
  var iProgram = header.indexOf('프로그램');
  var iCohort = header.indexOf('기수');
  var iPaid = header.indexOf('입금확인');
  var rows = sheet.getRange(2, 1, last - 1, width).getValues();

  var cohortLabel = cohort ? cohort + '기' : '';
  var count = 0, paid = 0;

  rows.forEach(function (r) {
    if (!String(r.join('')).trim()) return;                                  // 빈 줄 무시
    if (program && iProgram >= 0 && String(r[iProgram]) !== program) return;  // 다른 프로그램 제외
    if (cohortLabel && iCohort >= 0 && String(r[iCohort]) !== cohortLabel) return; // 지난 기수 제외
    count++;
    if (iPaid >= 0 && String(r[iPaid]).trim() !== '') paid++;
  });

  return { ok: true, count: count, paid: paid };
}

function respond_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function formatKST_(iso) {
  var d = iso ? new Date(iso) : new Date();
  return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

function notify_(payload, name) {
  try {
    var to = Session.getEffectiveUser().getEmail();
    if (!to) return;
    var lines = Object.keys(payload.answers || {}).map(function (k) {
      return k + ' : ' + payload.answers[k];
    });
    MailApp.sendEmail({
      to: to,
      subject: '[신청] ' + (payload.program || '') + ' ' + (payload.cohort ? payload.cohort + '기 ' : '') + '- ' + name,
      body: lines.join('\n') + '\n\n입금 예정 금액 : ' + Number(payload.fee || 0).toLocaleString() + '원\n'
        + '입금자명은 ' + name + ' 으로 안내되었습니다.'
    });
  } catch (e) { /* 메일 할당량 초과 등은 무시하고 저장은 계속 */ }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
