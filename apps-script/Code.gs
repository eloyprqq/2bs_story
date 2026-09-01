/**
 * 동아플 — 강의·질문·후기 저장
 * 배포: 실행=나, 액세스=모든 사용자
 */

var DELETE_PASSWORD = '2026';
var TEACHER_SIGNUP_CODE = '2026';
var USER_HEADERS = ['username', 'passHash', 'created', 'role', 'displayName', 'grade', 'points'];
var COURSE_HEADERS = ['id', 'title', 'desc', 'maxWeeks', 'content', 'materials', 'homework', 'teacherUid', 'teacherName', 'created'];
var ENROLL_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'currentWeek'];
var QUESTION_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'text', 'answer', 'answerDate', 'answerRead'];
var REVIEW_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'rating', 'text'];
var EVAL_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'text', 'teacherReply'];

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
    if (row.timestamp != null) row.timestamp = normalizeTs_(row.timestamp);
    items.push(row);
  }
  return items;
}

function findUser_(username) {
  var sh = sheet_('users', USER_HEADERS);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(username)) {
      return {
        username: String(values[i][0]),
        passHash: String(values[i][1]),
        role: String(values[i][3] || 'student'),
        displayName: String(values[i][4] || values[i][0]),
        grade: String(values[i][5] || '초급'),
        points: Number(values[i][6] || 0)
      };
    }
  }
  return null;
}

function signupUser_(username, passHash, role, teacherCode, displayName) {
  if (!username || !passHash) return json_({ ok: false, error: 'missing' });
  if (findUser_(username)) return json_({ ok: false, error: 'exists' });
  role = String(role || 'student');
  if (role === 'teacher' && String(teacherCode || '') !== TEACHER_SIGNUP_CODE) {
    return json_({ ok: false, error: 'badteacher' });
  }
  if (role !== 'teacher') role = 'student';
  var name = String(displayName || username);
  sheet_('users', USER_HEADERS).appendRow([username, passHash, new Date().toISOString(), role, name, role === 'teacher' ? '선생님' : '초급', 0]);
  return json_({ ok: true, username: username, role: role, displayName: name, grade: role === 'teacher' ? '선생님' : '초급', points: 0 });
}

function loginUser_(username, passHash, role) {
  var user = findUser_(username);
  if (!user || user.passHash !== passHash) return json_({ ok: false, error: 'bad' });
  if (role && String(role) !== user.role) return json_({ ok: false, error: 'badrole' });
  return json_({ ok: true, username: user.username, role: user.role, displayName: user.displayName, grade: user.grade, points: user.points });
}

function findQuestionRow_(sh, timestamp) {
  var target = normalizeTs_(timestamp);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (normalizeTs_(values[i][0]) === target) return i + 1;
  }
  return -1;
}

function setField_(sh, row, field, value) {
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf(field);
  if (col >= 0) sh.getRange(row, col + 1).setValue(value);
}

function deleteQuestion_(data) {
  if (String(data.password || '') !== DELETE_PASSWORD) return json_({ ok: false, error: 'badpass' });
  var sh = sheet_('questions', QUESTION_HEADERS);
  var target = normalizeTs_(data.timestamp);
  if (!target) return json_({ ok: false, error: 'notfound' });
  var delRow = findQuestionRow_(sh, target);
  if (delRow < 0) return json_({ ok: false, error: 'notfound' });
  sh.deleteRow(delRow);
  return json_({ ok: true });
}

function clearAllData_(data) {
  if (String(data.password || '') !== DELETE_PASSWORD) return json_({ ok: false, error: 'badpass' });
  ['users', 'courses', 'enrollments', 'questions', 'reviews', 'evaluations'].forEach(function(name) {
    var sh = SpreadsheetApp.getActive().getSheetByName(name);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  return json_({ ok: true });
}

function handleGet_(action, p) {
  if (action === 'signup') {
    return signupUser_(String(p.username || '').trim(), String(p.passHash || ''), p.role, p.teacherCode, p.displayName);
  }
  if (action === 'login') {
    return loginUser_(String(p.username || '').trim(), String(p.passHash || ''), p.role);
  }
  if (action === 'listCourses') {
    return json_({ items: rowsToObjects_(sheet_('courses', COURSE_HEADERS)) });
  }
  if (action === 'listEnrollments') {
    return json_({ items: rowsToObjects_(sheet_('enrollments', ENROLL_HEADERS)) });
  }
  if (action === 'listQuestions') {
    return json_({ items: rowsToObjects_(sheet_('questions', QUESTION_HEADERS)) });
  }
  if (action === 'listReviews') {
    return json_({ items: rowsToObjects_(sheet_('reviews', REVIEW_HEADERS)) });
  }
  if (action === 'listEvaluations') {
    return json_({ items: rowsToObjects_(sheet_('evaluations', EVAL_HEADERS)) });
  }
  if (action === 'deleteQuestion') return deleteQuestion_(p);
  if (action === 'clearAllData') return clearAllData_(p);
  return json_({ ok: false, error: 'unknown' });
}

function handlePost_(data) {
  if (data.action === 'signup') {
    return signupUser_(String(data.username || '').trim(), String(data.passHash || ''), data.role, data.teacherCode, data.displayName);
  }
  if (data.action === 'login') {
    return loginUser_(String(data.username || '').trim(), String(data.passHash || ''), data.role);
  }
  if (data.action === 'createCourse') {
    var id = String(data.id || ('c' + Date.now()));
    sheet_('courses', COURSE_HEADERS).appendRow([
      id, data.title || '', data.desc || '', Number(data.maxWeeks || 1),
      data.content || '', data.materials || '', data.homework || '',
      data.teacherUid || '', data.teacherName || '', new Date().toISOString()
    ]);
    return json_({ ok: true, id: id });
  }
  if (data.action === 'enroll') {
    sheet_('enrollments', ENROLL_HEADERS).appendRow([
      normalizeTs_(data.timestamp || Date.now()), data.date || '',
      data.uid || '', data.name || '', data.courseId || '', data.courseName || '', Number(data.currentWeek || 1)
    ]);
    return json_({ ok: true });
  }
  if (data.action === 'question') {
    sheet_('questions', QUESTION_HEADERS).appendRow([
      normalizeTs_(data.timestamp || Date.now()), data.date || '', data.uid || '', data.name || '',
      data.courseId || '', data.courseName || '', data.text || '', '', '', ''
    ]);
    return json_({ ok: true });
  }
  if (data.action === 'answerQuestion') {
    var qSh = sheet_('questions', QUESTION_HEADERS);
    var ansRow = data.row ? parseInt(data.row, 10) : findQuestionRow_(qSh, data.timestamp);
    if (ansRow < 2) return json_({ ok: false, error: 'notfound' });
    setField_(qSh, ansRow, 'answer', data.answer || '');
    setField_(qSh, ansRow, 'answerDate', data.answerDate || '');
    setField_(qSh, ansRow, 'answerRead', '');
    return json_({ ok: true });
  }
  if (data.action === 'markAnswerRead') {
    var readSh = sheet_('questions', QUESTION_HEADERS);
    var readRow = data.row ? parseInt(data.row, 10) : findQuestionRow_(readSh, data.timestamp);
    if (readRow < 2) return json_({ ok: false, error: 'notfound' });
    setField_(readSh, readRow, 'answerRead', '1');
    return json_({ ok: true });
  }
  if (data.action === 'review') {
    sheet_('reviews', REVIEW_HEADERS).appendRow([
      normalizeTs_(data.timestamp || Date.now()), data.date || '', data.uid || '', data.name || '',
      data.courseId || '', data.courseName || '', Number(data.rating || 5), data.text || ''
    ]);
    return json_({ ok: true });
  }
  if (data.action === 'evaluation') {
    sheet_('evaluations', EVAL_HEADERS).appendRow([
      normalizeTs_(data.timestamp || Date.now()), data.date || '', data.uid || '', data.name || '',
      data.courseId || '', data.courseName || '', data.text || '', ''
    ]);
    return json_({ ok: true });
  }
  if (data.action === 'deleteQuestion') return deleteQuestion_(data);
  if (data.action === 'clearAllData') return clearAllData_(data);
  if (data.action === 'updateProfile') {
    var user = findUser_(String(data.username || '').trim());
    if (!user) return json_({ ok: false, error: 'notfound' });
    var uSh = sheet_('users', USER_HEADERS);
    var vals = uSh.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === user.username) {
        if (data.displayName) uSh.getRange(i + 1, 5).setValue(String(data.displayName));
        return json_({ ok: true });
      }
    }
    return json_({ ok: false, error: 'notfound' });
  }
  return json_({ ok: false, error: 'unknown' });
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  return handleGet_(p.action || 'listCourses', p);
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    return handlePost_(JSON.parse(raw));
  } catch (err) {
    return json_({ ok: false, error: 'badjson' });
  }
}
