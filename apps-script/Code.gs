/**
 * 동아플 — 강의·질문·후기 저장
 * 배포: 실행=나, 액세스=모든 사용자
 *
 * 브라우저 POST는 구글 리다이렉트 때문에 저장이 안 됩니다.
 * 모든 저장은 GET 쿼리로도 처리합니다.
 * 코드 수정 후: 배포 → 배포 관리 → 연필 → 새 버전 → 배포
 */

var DELETE_PASSWORD = '2026';
var TEACHER_SIGNUP_CODE = '2026';
var USER_HEADERS = ['username', 'passHash', 'created', 'role', 'displayName', 'grade', 'points'];
var COURSE_HEADERS = ['id', 'title', 'desc', 'maxWeeks', 'content', 'materials', 'homework', 'teacherUid', 'teacherName', 'created'];
var ENROLL_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'currentWeek', 'hwDoneWeek'];
var QUESTION_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'text', 'answer', 'answerDate', 'answerRead'];
var REVIEW_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'rating', 'text'];
var EVAL_HEADERS = ['timestamp', 'date', 'uid', 'name', 'courseId', 'courseName', 'text', 'teacherReply', 'authorRole'];

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
    return sh;
  }
  ensureHeaders_(sh, headers);
  return sh;
}

function ensureHeaders_(sh, headers) {
  var colCount = sh.getLastColumn();
  if (colCount < 1) {
    sh.appendRow(headers);
    return;
  }
  var existing = sh.getRange(1, 1, 1, colCount).getValues()[0].map(function(h) { return String(h); });
  headers.forEach(function(h) {
    if (existing.indexOf(h) === -1) {
      colCount += 1;
      sh.getRange(1, colCount).setValue(h);
      existing.push(h);
    }
  });
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
    var any = false;
    keys.forEach(function(k, idx) {
      row[k] = values[i][idx];
      if (k && String(values[i][idx] == null ? '' : values[i][idx]).trim()) any = true;
    });
    if (!any) continue;
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
  var col = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]) === field) { col = i; break; }
  }
  if (col >= 0) sh.getRange(row, col + 1).setValue(value);
}

function findEnrollment_(uid, courseId) {
  var sh = sheet_('enrollments', ENROLL_HEADERS);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return null;
  var headers = values[0].map(function(h) { return String(h); });
  var uidCol = headers.indexOf('uid');
  var courseCol = headers.indexOf('courseId');
  if (uidCol < 0 || courseCol < 0) return null;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][uidCol]) === String(uid) && String(values[i][courseCol]) === String(courseId)) {
      return { sh: sh, row: i + 1 };
    }
  }
  return null;
}

function courseMaxWeeks_(courseId) {
  var items = rowsToObjects_(sheet_('courses', COURSE_HEADERS));
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].id) === String(courseId)) {
      var n = Math.round(Number(items[i].maxWeeks || 8));
      return n > 0 ? n : 8;
    }
  }
  return 52;
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

function createCourse_(data) {
  var title = String(data.title || '').trim();
  if (!title) return json_({ ok: false, error: 'missing' });
  var id = String(data.id || ('c' + Date.now()));
  var teacherUid = String(data.teacherUid || '');
  var existing = rowsToObjects_(sheet_('courses', COURSE_HEADERS));
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i].id) === id) return json_({ ok: true, id: id, duplicate: true });
    if (teacherUid && String(existing[i].teacherUid) === teacherUid && String(existing[i].title || '').trim() === title) {
      return json_({ ok: false, error: 'duplicate' });
    }
  }
  var weeks = Math.round(Number(data.maxWeeks || 8));
  if (!weeks || weeks < 1) weeks = 8;
  if (weeks > 52) weeks = 52;
  sheet_('courses', COURSE_HEADERS).appendRow([
    id, title, data.desc || '', weeks,
    data.content || '', data.materials || '', data.homework || '',
    teacherUid, data.teacherName || '', new Date().toISOString()
  ]);
  return json_({ ok: true, id: id });
}

function enroll_(data) {
  if (!data.uid || !data.courseId) return json_({ ok: false, error: 'missing' });
  if (findEnrollment_(data.uid, data.courseId)) return json_({ ok: true, duplicate: true });
  var week = Math.round(Number(data.currentWeek || 1));
  if (!week || week < 1) week = 1;
  sheet_('enrollments', ENROLL_HEADERS).appendRow([
    normalizeTs_(data.timestamp || Date.now()), data.date || '',
    String(data.uid), data.name || '', String(data.courseId), data.courseName || '',
    week, Math.round(Number(data.hwDoneWeek || 0))
  ]);
  return json_({ ok: true });
}

function setWeek_(data) {
  var found = findEnrollment_(data.uid, data.courseId);
  if (!found) return json_({ ok: false, error: 'notfound' });
  var week = Math.round(Number(data.currentWeek || 1));
  if (!week || week < 1) week = 1;
  var max = courseMaxWeeks_(data.courseId);
  if (week > max) week = max;
  setField_(found.sh, found.row, 'currentWeek', week);
  return json_({ ok: true, currentWeek: week });
}

function markHomework_(data) {
  var found = findEnrollment_(data.uid, data.courseId);
  if (!found) return json_({ ok: false, error: 'notfound' });
  var done = Math.round(Number(data.hwDoneWeek || 0));
  if (done < 0) done = 0;
  setField_(found.sh, found.row, 'hwDoneWeek', done);
  return json_({ ok: true, hwDoneWeek: done });
}

function addQuestion_(data) {
  var text = String(data.text || '').trim();
  if (!text) return json_({ ok: false, error: 'missing' });
  var ts = normalizeTs_(data.timestamp || Date.now());
  var sh = sheet_('questions', QUESTION_HEADERS);
  if (findQuestionRow_(sh, ts) > 0) return json_({ ok: true, duplicate: true });
  sh.appendRow([
    ts, data.date || '', data.uid || '', data.name || '',
    data.courseId || '', data.courseName || '', text, '', '', ''
  ]);
  return json_({ ok: true });
}

function answerQuestion_(data) {
  var qSh = sheet_('questions', QUESTION_HEADERS);
  var ansRow = data.row ? parseInt(data.row, 10) : findQuestionRow_(qSh, data.timestamp);
  if (ansRow < 2) return json_({ ok: false, error: 'notfound' });
  setField_(qSh, ansRow, 'answer', data.answer || '');
  setField_(qSh, ansRow, 'answerDate', data.answerDate || '');
  setField_(qSh, ansRow, 'answerRead', '');
  return json_({ ok: true });
}

function markAnswerRead_(data) {
  var readSh = sheet_('questions', QUESTION_HEADERS);
  var readRow = data.row ? parseInt(data.row, 10) : findQuestionRow_(readSh, data.timestamp);
  if (readRow < 2) return json_({ ok: false, error: 'notfound' });
  setField_(readSh, readRow, 'answerRead', '1');
  return json_({ ok: true });
}

function addReview_(data) {
  var ts = normalizeTs_(data.timestamp || Date.now());
  var sh = sheet_('reviews', REVIEW_HEADERS);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (normalizeTs_(values[i][0]) === ts) return json_({ ok: true, duplicate: true });
  }
  sh.appendRow([
    ts, data.date || '', data.uid || '', data.name || '',
    data.courseId || '', data.courseName || '', Number(data.rating || 5), data.text || ''
  ]);
  return json_({ ok: true });
}

function addEvaluation_(data) {
  var text = String(data.text || '').trim();
  if (!text) return json_({ ok: false, error: 'missing' });
  var ts = normalizeTs_(data.timestamp || Date.now());
  var sh = sheet_('evaluations', EVAL_HEADERS);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (normalizeTs_(values[i][0]) === ts) return json_({ ok: true, duplicate: true });
  }
  sh.appendRow([
    ts, data.date || '', data.uid || '', data.name || '',
    data.courseId || '', data.courseName || '', text, '',
    data.authorRole || 'student'
  ]);
  return json_({ ok: true });
}

function sendResult_(data) {
  var text = String(data.text || '').trim();
  if (!data.uid || !data.courseId || !text) return json_({ ok: false, error: 'missing' });
  var ts = normalizeTs_(data.timestamp || Date.now());
  var sh = sheet_('evaluations', EVAL_HEADERS);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (normalizeTs_(values[i][0]) === ts) return json_({ ok: true, duplicate: true });
  }
  sh.appendRow([
    ts, data.date || '', String(data.uid), String(data.teacherName || data.name || '선생님'),
    String(data.courseId), data.courseName || '', text, '', 'teacher'
  ]);
  return json_({ ok: true });
}

function updateProfile_(data) {
  var username = String(data.username || '').trim();
  var sh = sheet_('users', USER_HEADERS);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === username) {
      var row = i + 1;
      if (data.displayName) setField_(sh, row, 'displayName', String(data.displayName));
      if (data.grade) setField_(sh, row, 'grade', String(data.grade));
      if (data.points != null && String(data.points) !== '') setField_(sh, row, 'points', Number(data.points) || 0);
      return json_({ ok: true });
    }
  }
  return json_({ ok: false, error: 'notfound' });
}

function handlePost_(data) {
  data = data || {};
  if (data.action === 'signup') {
    return signupUser_(String(data.username || '').trim(), String(data.passHash || ''), data.role, data.teacherCode, data.displayName);
  }
  if (data.action === 'login') {
    return loginUser_(String(data.username || '').trim(), String(data.passHash || ''), data.role);
  }
  if (data.action === 'createCourse') return createCourse_(data);
  if (data.action === 'enroll') return enroll_(data);
  if (data.action === 'setWeek') return setWeek_(data);
  if (data.action === 'markHomework') return markHomework_(data);
  if (data.action === 'question') return addQuestion_(data);
  if (data.action === 'answerQuestion') return answerQuestion_(data);
  if (data.action === 'markAnswerRead') return markAnswerRead_(data);
  if (data.action === 'review') return addReview_(data);
  if (data.action === 'evaluation') return addEvaluation_(data);
  if (data.action === 'sendResult') return sendResult_(data);
  if (data.action === 'deleteQuestion') return deleteQuestion_(data);
  if (data.action === 'clearAllData') return clearAllData_(data);
  if (data.action === 'updateProfile') return updateProfile_(data);
  return json_({ ok: false, error: 'unknown' });
}

function handleGet_(action, p) {
  if (action === 'listCourses') return json_({ items: rowsToObjects_(sheet_('courses', COURSE_HEADERS)) });
  if (action === 'listEnrollments') return json_({ items: rowsToObjects_(sheet_('enrollments', ENROLL_HEADERS)) });
  if (action === 'listQuestions') return json_({ items: rowsToObjects_(sheet_('questions', QUESTION_HEADERS)) });
  if (action === 'listReviews') return json_({ items: rowsToObjects_(sheet_('reviews', REVIEW_HEADERS)) });
  if (action === 'listEvaluations') return json_({ items: rowsToObjects_(sheet_('evaluations', EVAL_HEADERS)) });
  var data = {};
  for (var k in p) if (Object.prototype.hasOwnProperty.call(p, k)) data[k] = p[k];
  data.action = action;
  return handlePost_(data);
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
