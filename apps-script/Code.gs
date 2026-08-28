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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function rowsToObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const keys = values[0];
  const items = [];
  for (let i = values.length - 1; i >= 1; i--) {
    const row = {};
    keys.forEach((k, idx) => { row[k] = values[i][idx]; });
    items.push(row);
  }
  return items;
}

function findUser_(username) {
  const sh = sheet_('users', ['username', 'passHash', 'created']);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(username)) {
      return { username: String(values[i][0]), passHash: String(values[i][1]) };
    }
  }
  return null;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'listApplicants';

  if (action === 'signup') {
    const username = String(p.username || '').trim();
    const passHash = String(p.passHash || '');
    if (!username || !passHash) return json_({ ok: false, error: 'missing' });
    if (findUser_(username)) return json_({ ok: false, error: 'exists' });
    const sh = sheet_('users', ['username', 'passHash', 'created']);
    sh.appendRow([username, passHash, new Date().toISOString()]);
    return json_({ ok: true, username: username });
  }

  if (action === 'login') {
    const username = String(p.username || '').trim();
    const passHash = String(p.passHash || '');
    const user = findUser_(username);
    if (!user || user.passHash !== passHash) return json_({ ok: false, error: 'bad' });
    return json_({ ok: true, username: user.username });
  }

  if (action === 'listQuestions') {
    const sh = sheet_('questions', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName', 'text']);
    return json_({ items: rowsToObjects_(sh) });
  }
  const sh = sheet_('applicants', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName']);
  return json_({ items: rowsToObjects_(sh) });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'signup') {
    const username = String(data.username || '').trim();
    const passHash = String(data.passHash || '');
    if (!username || !passHash) return json_({ ok: false, error: 'missing' });
    if (findUser_(username)) return json_({ ok: false, error: 'exists' });
    const sh = sheet_('users', ['username', 'passHash', 'created']);
    sh.appendRow([username, passHash, new Date().toISOString()]);
    return json_({ ok: true, username: username });
  }

  if (data.action === 'login') {
    const username = String(data.username || '').trim();
    const passHash = String(data.passHash || '');
    const user = findUser_(username);
    if (!user || user.passHash !== passHash) return json_({ ok: false, error: 'bad' });
    return json_({ ok: true, username: user.username });
  }

  if (data.action === 'question') {
    const sh = sheet_('questions', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName', 'text']);
    sh.appendRow([data.timestamp || Date.now(), data.date || '', data.name || '', data.uid || '', data.courseId || '', data.courseName || '', data.text || '']);
    return json_({ ok: true });
  }
  const sh = sheet_('applicants', ['timestamp', 'date', 'name', 'uid', 'courseId', 'courseName']);
  sh.appendRow([data.timestamp || Date.now(), data.date || '', data.name || '', data.uid || '', data.courseId || '', data.courseName || '']);
  return json_({ ok: true });
}
