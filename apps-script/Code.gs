/**
 * 2BS 신청자·질문 저장
 * 배포: 실행=나, 액세스=모든 사용자
 */

var QUESTION_HEADERS = ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName', 'text', 'answer', 'answerDate', 'answerRead'];
var DELETE_PASSWORD = '2026';

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
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
    if (String(existing[i] || '') !== h) sh.getRange(1, i + 1).setValue(h);
  });
  return sh;
}

function normalizeTs_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return String(v.getTime());
  if (typeof v === 'number' && !isNaN(v)) return String(Math.round(v));
  var s = String(v == null ? '' : v).trim();
  if (/^\d+(\.\d+)?$/.test(s)) return String(Math.round(parseFloat(s)));
  return s;
}

function rowsToObjects_(sh) {
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var keys = values[0];
  var items = [];
  for (var i = values.length - 1; i >= 1; i--) {
    var row = { _row: i + 1 };
    keys.forEach(function(k, idx) { row[k] = values[i][idx]; });
    row.timestamp = normalizeTs_(row.timestamp);
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
  var target = normalizeTs_(timestamp);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (normalizeTs_(values[i][0]) === target) return i + 1;
  }
  return -1;
}

function setQuestionField_(sh, row, field, value) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf(field);
  if (col >= 0) sh.getRange(row, col + 1).setValue(value);
}

function deleteQuestion_(data) {
  if (String(data.password || '') !== DELETE_PASSWORD) return json_({ ok: false, error: 'badpass' });
  var sh = questionsSheet_();
  var target = normalizeTs_(data.timestamp);
  if (!target) return json_({ ok: false, error: 'notfound' });
  var rowNum = parseInt(data.row, 10);
  if (rowNum > 1 && rowNum <= sh.getLastRow()) {
    var cellTs = normalizeTs_(sh.getRange(rowNum, 1).getValue());
    if (cellTs === target) {
      sh.deleteRow(rowNum);
      return json_({ ok: true });
    }
  }
  var delRow = findQuestionRow_(sh, target);
  if (delRow < 0) return json_({ ok: false, error: 'notfound' });
  sh.deleteRow(delRow);
  return json_({ ok: true });
}

function findApplicantRow_(sh, timestamp) {
  var target = normalizeTs_(timestamp);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (normalizeTs_(values[i][0]) === target) return i + 1;
  }
  return -1;
}

function deleteApplicant_(data) {
  if (String(data.password || '') !== DELETE_PASSWORD) return json_({ ok: false, error: 'badpass' });
  var sh = sheet_('applicants', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName']);
  var target = normalizeTs_(data.timestamp);
  if (!target) return json_({ ok: false, error: 'notfound' });
  var rowNum = parseInt(data.row, 10);
  if (rowNum > 1 && rowNum <= sh.getLastRow()) {
    var cellTs = normalizeTs_(sh.getRange(rowNum, 1).getValue());
    if (cellTs === target) {
      sh.deleteRow(rowNum);
      return json_({ ok: true });
    }
  }
  var delRow = findApplicantRow_(sh, target);
  if (delRow < 0) return json_({ ok: false, error: 'notfound' });
  sh.deleteRow(delRow);
  return json_({ ok: true });
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'listApplicants';

  if (action === 'signup') {
    var username = String(p.username || '').trim();
    var passHash = String(p.passHash || '');
    if (!username || !passHash) return json_({ ok: false, error: 'missing' });
    if (findUser_(username)) return json_({ ok: false, error: 'exists' });
    sheet_('users', ['username', 'passHash', 'created']).appendRow([username, passHash, new Date().toISOString()]);
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
    return json_({ items: rowsToObjects_(questionsSheet_()) });
  }

  if (action === 'listApplicants') {
    return json_({ items: rowsToObjects_(sheet_('applicants', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName'])) });
  }

  if (action === 'deleteQuestion') {
    return deleteQuestion_(p);
  }

  if (action === 'deleteApplicant') {
    return deleteApplicant_(p);
  }

  if (action === 'answerQuestion') {
    var ansSh = questionsSheet_();
    var ansRowNum = p.row ? parseInt(p.row, 10) : findQuestionRow_(ansSh, p.timestamp);
    if (ansRowNum < 2) return json_({ ok: false, error: 'notfound' });
    setQuestionField_(ansSh, ansRowNum, 'answer', p.answer || '');
    setQuestionField_(ansSh, ansRowNum, 'answerDate', p.answerDate || '');
    setQuestionField_(ansSh, ansRowNum, 'answerRead', '');
    return json_({ ok: true });
  }

  if (action === 'markAnswerRead') {
    var readSh = questionsSheet_();
    var readRowNum = p.row ? parseInt(p.row, 10) : findQuestionRow_(readSh, p.timestamp);
    if (readRowNum < 2) return json_({ ok: false, error: 'notfound' });
    setQuestionField_(readSh, readRowNum, 'answerRead', '1');
    return json_({ ok: true });
  }

  return json_({ ok: false, error: 'unknown' });
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);

  if (data.action === 'signup') {
    var suName = String(data.username || '').trim();
    var suHash = String(data.passHash || '');
    if (!suName || !suHash) return json_({ ok: false, error: 'missing' });
    if (findUser_(suName)) return json_({ ok: false, error: 'exists' });
    sheet_('users', ['username', 'passHash', 'created']).appendRow([suName, suHash, new Date().toISOString()]);
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
    questionsSheet_().appendRow([
      normalizeTs_(data.timestamp || Date.now()),
      data.date || '', data.name || '', data.uid || '',
      data.courseId || '', data.courseName || '', data.text || '',
      '', '', ''
    ]);
    return json_({ ok: true });
  }

  if (data.action === 'answerQuestion') {
    var sh = questionsSheet_();
    var ansRow = data.row ? parseInt(data.row, 10) : findQuestionRow_(sh, data.timestamp);
    if (ansRow < 2) return json_({ ok: false, error: 'notfound' });
    setQuestionField_(sh, ansRow, 'answer', data.answer || '');
    setQuestionField_(sh, ansRow, 'answerDate', data.answerDate || '');
    setQuestionField_(sh, ansRow, 'answerRead', '');
    return json_({ ok: true });
  }

  if (data.action === 'deleteQuestion') {
    return deleteQuestion_(data);
  }

  if (data.action === 'deleteApplicant') {
    return deleteApplicant_(data);
  }

  if (data.action === 'markAnswerRead') {
    var qRead = questionsSheet_();
    var readRow = data.row ? parseInt(data.row, 10) : findQuestionRow_(qRead, data.timestamp);
    if (readRow < 2) return json_({ ok: false, error: 'notfound' });
    setQuestionField_(qRead, readRow, 'answerRead', '1');
    return json_({ ok: true });
  }

  if (data.action === 'enroll') {
    sheet_('applicants', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName']).appendRow([
      normalizeTs_(data.timestamp || Date.now()),
      data.date || '', data.name || '', data.uid || '', data.courseId || '', data.courseName || ''
    ]);
    return json_({ ok: true });
  }

  return json_({ ok: false, error: 'unknown' });
}
