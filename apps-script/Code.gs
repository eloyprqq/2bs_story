/**
 * 2BS 신청자·질문 저장 (Firebase 없이)
 *
 * 1. Google 스프레드시트 새로 만들기
 * 2. 확장 프로그램 > Apps Script 에 이 파일 붙여넣기
 * 3. 배포 > 새 배포 > 웹 앱
 *    - 실행 계정: 나
 *    - 액세스: 모든 사용자
 * 4. 나온 URL을 index.html 의 APPS_SCRIPT_URL 에 붙여넣기
 */

var QUESTION_HEADERS = ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName', 'text', 'answer', 'answerDate', 'answerRead'];
var DELETE_PASSWORD = '2026';

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function questionsSheet_() {
  var sh = sheet_('questions', QUESTION_HEADERS);
  var lastCol = Math.max(sh.getLastColumn(), QUESTION_HEADERS.length);
  var existing = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  QUESTION_HEADERS.forEach(function(h, i) {
    if (String(existing[i] || '') !== h) {
      sh.getRange(1, i + 1).setValue(h);
    }
  });
  return sh;
}

function rowsToObjects_(sh) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var keys = values[0];
  var items = [];
  for (var i = values.length - 1; i >= 1; i--) {
    var row = {};
    keys.forEach(function(k, idx) { row[k] = values[i][idx]; });
    items.push(row);
  }
  return items;
}

function findUser_(username) {
  var sh = sheet_('users', ['username', 'passHash', 'created']);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(username)) {
      return { username: String(values[i][0]), passHash: String(values[i][1]) };
    }
  }
  return null;
}

function findQuestionRow_(sh, timestamp) {
  var values = sh.getDataRange().getValues();
  var ts = String(timestamp);
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === ts) return i + 1;
  }
  return -1;
}

function setQuestionField_(sh, row, field, value) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf(field);
  if (col >= 0) sh.getRange(row, col + 1).setValue(value);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'listApplicants';

  if (action === 'signup') {
    var username = String(p.username || '').trim();
    var passHash = String(p.passHash || '');
    if (!username || !passHash) return json_({ ok: false, error: 'missing' });
    if (findUser_(username)) return json_({ ok: false, error: 'exists' });
    var userSh = sheet_('users', ['username', 'passHash', 'created']);
    userSh.appendRow([username, passHash, new Date().toISOString()]);
    return json_({ ok: true, username: username });
  }

  if (action === 'login') {
    var loginName = String(p.username || '').trim();
    var loginHash = String(p.passHash || '');
    var user = findUser_(loginName);
    if (!user || user.passHash !== loginHash) return json_({ ok: false, error: 'bad' });
    return json_({ ok: true, username: user.username });
  }

  if (action === 'listQuestions') {
    var qSh = questionsSheet_();
    return json_({ items: rowsToObjects_(qSh) });
  }

  var appSh = sheet_('applicants', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName']);
  return json_({ items: rowsToObjects_(appSh) });
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);

  if (data.action === 'signup') {
    var suName = String(data.username || '').trim();
    var suHash = String(data.passHash || '');
    if (!suName || !suHash) return json_({ ok: false, error: 'missing' });
    if (findUser_(suName)) return json_({ ok: false, error: 'exists' });
    var suSh = sheet_('users', ['username', 'passHash', 'created']);
    suSh.appendRow([suName, suHash, new Date().toISOString()]);
    return json_({ ok: true, username: suName });
  }

  if (data.action === 'login') {
    var liName = String(data.username || '').trim();
    var liHash = String(data.passHash || '');
    var liUser = findUser_(liName);
    if (!liUser || liUser.passHash !== liHash) return json_({ ok: false, error: 'bad' });
    return json_({ ok: true, username: liUser.username });
  }

  if (data.action === 'question') {
    var qAdd = questionsSheet_();
    qAdd.appendRow([
      data.timestamp || Date.now(),
      data.date || '',
      data.name || '',
      data.uid || '',
      data.courseId || '',
      data.courseName || '',
      data.text || '',
      '',
      '',
      ''
    ]);
    return json_({ ok: true });
  }

  if (data.action === 'answerQuestion') {
    var qAns = questionsSheet_();
    var ansRow = findQuestionRow_(qAns, data.timestamp);
    if (ansRow < 0) return json_({ ok: false, error: 'notfound' });
    setQuestionField_(qAns, ansRow, 'answer', data.answer || '');
    setQuestionField_(qAns, ansRow, 'answerDate', data.answerDate || '');
    setQuestionField_(qAns, ansRow, 'answerRead', '');
    return json_({ ok: true });
  }

  if (data.action === 'deleteQuestion') {
    if (String(data.password || '') !== DELETE_PASSWORD) return json_({ ok: false, error: 'badpass' });
    var qDel = questionsSheet_();
    var delRow = findQuestionRow_(qDel, data.timestamp);
    if (delRow < 0) return json_({ ok: false, error: 'notfound' });
    qDel.deleteRow(delRow);
    return json_({ ok: true });
  }

  if (data.action === 'markAnswerRead') {
    var qRead = questionsSheet_();
    var readRow = findQuestionRow_(qRead, data.timestamp);
    if (readRow < 0) return json_({ ok: false, error: 'notfound' });
    setQuestionField_(qRead, readRow, 'answerRead', '1');
    return json_({ ok: true });
  }

  var enSh = sheet_('applicants', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName']);
  enSh.appendRow([data.timestamp || Date.now(), data.date || '', data.name || '', data.uid || '', data.courseId || '', data.courseName || '']);
  return json_({ ok: true });
}
