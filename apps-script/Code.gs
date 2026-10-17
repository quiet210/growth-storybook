/**
 * 나의 성장 그림책 도우미 — 데이터 서버 (Google Apps Script)
 *
 * 화면(GitHub Pages)이 이 스크립트에 데이터를 보내고 받습니다.
 * 데이터는 이 스크립트가 연결된 구글 시트의 '수업방', '학생기록' 시트에 저장됩니다.
 *
 * 설치 순서는 저장소의 README.md 를 보세요.
 */

// ===== 처음 한 번 바꾸는 부분 =====
const CONFIG = {
  // 선생님이 '새 수업방 만들기'를 할 때 입력해야 하는 학교 코드입니다.
  // 함께 쓸 선생님들에게만 알려 주세요.
  SCHOOL_CODE: 'change-me',
  WORK_TITLE: '멍키 스패너',
  STORYBOOK_URL: 'https://gemini.google.com/gem/storybook',
};
// ==================================

const ROOM_SHEET = '수업방';
const DATA_SHEET = '학생기록';

const ROOM_FIELDS = [
  ['room', '방 코드'],
  ['teacher', '선생님'],
  ['title', '수업 이름'],
  ['grade', '학년'],
  ['classes', '반 목록'],
  ['padletUrl', '패들렛 주소'],
  ['pwHash', '비밀번호(암호화됨)'],
  ['createdAt', '만든 시각'],
];

const FIELDS = [
  ['id', 'ID'],
  ['room', '방 코드'],
  ['updatedAt', '저장 시각'],
  ['cls', '반'],
  ['num', '번호'],
  ['name', '이름'],
  ['step', '진행 단계'],
  ['hardship', '1. 내가 겪었던 어려움'],
  ['pastAction', '2. 그때 나는 이렇게 행동했다'],
  ['pastView', '3. 지금 생각하는 그때의 나'],
  ['nowAction', '4. 같은 상황이라면'],
  ['reason', '5. 그렇게 행동하려는 이유'],
  ['growth', '6. 나의 성장, 한 문장'],
  ['style', '그림 스타일'],
  ['hero', '주인공 표현'],
  ['prompt', '생성 프롬프트'],
  ['check1', '점검: 어려움이 드러난다'],
  ['check2', '점검: 변화가 드러난다'],
  ['check3', '점검: 전하고 싶은 생각'],
  ['keyScene', '가장 중요한 장면'],
  ['sceneWhy', '장면이 중요한 까닭'],
  ['insight', '새롭게 깨달은 점'],
  ['friendBook', '내가 읽은 친구 그림책'],
  ['friendMoment', '친구의 어려움/변화의 순간'],
  ['growthAlso', '성장에는 ~도 있다'],
  ['compare', '비슷한 점/다른 점'],
  ['realized', '삶을 돌아보며 깨달은 것'],
  ['literature', '문학은 나에게 ~을 가능하게 해 준다'],
  ['anonymous', '익명 공유'],
];
// 학생이 직접 바꿀 수 있는 칸 (id, room, 반, 번호, 이름, 시각은 서버가 정함)
const STUDENT_EDITABLE = FIELDS.map(f => f[0]).filter(k =>
  ['id', 'room', 'updatedAt', 'cls', 'num', 'name', 'step'].indexOf(k) === -1);

// ---------------------------------------------------------------- 진입점

function doGet() {
  return json_({ ok: true, data: { app: 'growth-storybook', version: 2 } });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const handler = ACTIONS[body.action];
    if (!handler) throw new UserError('알 수 없는 요청입니다.');
    return json_({ ok: true, data: handler(body) });
  } catch (err) {
    const msg = err instanceof UserError ? err.message : '서버 오류가 났어요. 잠시 후 다시 시도해 주세요.';
    if (!(err instanceof UserError)) console.error(err && err.stack || err);
    return json_({ ok: false, error: msg });
  }
}

/** 처음 한 번 실행: 시트를 만들고 권한을 승인합니다. */
function setup() {
  getSheet_(ROOM_SHEET, ROOM_FIELDS);
  getSheet_(DATA_SHEET, FIELDS);
  salt_();
}

const ACTIONS = {
  ping: () => ({ app: 'growth-storybook', version: 2 }),
  roomInfo: b => publicRoom_(getRoom_(b.room)),
  createRoom: createRoom_,
  teacherLogin: b => privateRoom_(auth_(b.room, b.password)),
  updateRoom: updateRoom_,
  loadStudent: loadStudent_,
  saveStudent: saveStudent_,
  teacherData: teacherData_,
  deleteStudent: deleteStudent_,
};

// ---------------------------------------------------------------- 공통

function UserError(message) { this.message = message; }
UserError.prototype = Object.create(Error.prototype);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name, fields) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, fields.length).setValues([fields.map(f => f[1])])
      .setFontWeight('bold').setBackground('#fde68a');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), fields.length).setNumberFormat('@');
  }
  return sh;
}

function readAll_(sh, fields) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, fields.length).getValues().map((row, i) => {
    const o = { _row: i + 2 };
    fields.forEach(([k], j) => { o[k] = row[j]; });
    return o;
  });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new UserError('지금 저장하는 친구들이 많아요. 잠시 후 다시 눌러 주세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function str_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max || 200);
}

function normNum_(v) {
  const n = parseInt(String(v == null ? '' : v).trim(), 10);
  return isNaN(n) || n < 1 || n > 99 ? '' : String(n);
}

function normName_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').slice(0, 20);
}

function normRoom_(v) {
  return String(v == null ? '' : v).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function cell_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'boolean' || typeof v === 'number') return v;
  let s = String(v).slice(0, 3000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 시트에서 수식으로 실행되는 것 방지
  return s;
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function salt_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('SALT');
  if (!s) { s = Utilities.getUuid(); props.setProperty('SALT', s); }
  return s;
}

function hash_(pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + '|' + pw, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

// ---------------------------------------------------------------- 수업방

function parseClasses_(v) {
  const list = String(v == null ? '' : v).split(/[,\s/]+/).map(normNum_).filter(Boolean);
  return list.filter((c, i) => list.indexOf(c) === i).sort((a, b) => a - b);
}

function getRoom_(code) {
  code = normRoom_(code);
  if (!code) throw new UserError('수업방 코드를 입력해 주세요.');
  const room = readAll_(getSheet_(ROOM_SHEET, ROOM_FIELDS), ROOM_FIELDS).filter(r => normRoom_(r.room) === code)[0];
  if (!room) throw new UserError('수업방 코드를 찾을 수 없어요. 선생님께 다시 확인해 주세요.');
  room.room = code;
  room.classesList = parseClasses_(room.classes);
  return room;
}

function publicRoom_(r) {
  return {
    room: r.room, teacher: r.teacher, title: r.title, grade: String(r.grade),
    classes: r.classesList, padletUrl: r.padletUrl,
    workTitle: CONFIG.WORK_TITLE, storybookUrl: CONFIG.STORYBOOK_URL,
  };
}

function privateRoom_(r) {
  return Object.assign(publicRoom_(r), { createdAt: r.createdAt });
}

function auth_(code, password) {
  const room = getRoom_(code);
  if (!password || hash_(String(password)) !== room.pwHash) throw new UserError('비밀번호가 맞지 않아요.');
  return room;
}

function validRoomInput_(b) {
  const out = {
    teacher: str_(b.teacher, 30),
    title: str_(b.title, 60) || CONFIG.WORK_TITLE + ' 성장 그림책',
    grade: normNum_(b.grade) || '1',
    classes: parseClasses_(b.classes).join(','),
    padletUrl: str_(b.padletUrl, 300),
  };
  if (!out.teacher) throw new UserError('선생님 이름을 입력해 주세요.');
  if (!out.classes) throw new UserError('수업하는 반을 하나 이상 입력해 주세요. 예) 5, 8, 15');
  if (out.padletUrl && !/^https:\/\//.test(out.padletUrl)) throw new UserError('패들렛 주소는 https:// 로 시작해야 해요.');
  return out;
}

function createRoom_(b) {
  if (String(b.schoolCode || '').trim() !== String(CONFIG.SCHOOL_CODE)) throw new UserError('학교 코드가 맞지 않아요.');
  const pw = String(b.password || '');
  if (pw.length < 4) throw new UserError('비밀번호는 4자 이상으로 정해 주세요.');
  const input = validRoomInput_(b);
  return withLock_(() => {
    const sh = getSheet_(ROOM_SHEET, ROOM_FIELDS);
    const used = readAll_(sh, ROOM_FIELDS).map(r => normRoom_(r.room));
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    let code;
    do {
      code = '';
      for (let i = 0; i < 6; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
    } while (used.indexOf(code) !== -1);
    const rec = Object.assign({ room: code, pwHash: hash_(pw), createdAt: now_() }, input);
    sh.appendRow(ROOM_FIELDS.map(([k]) => cell_(rec[k])));
    return privateRoom_(getRoom_(code));
  });
}

function updateRoom_(b) {
  const room = auth_(b.room, b.password);
  const input = validRoomInput_(Object.assign({}, room, b));
  return withLock_(() => {
    const sh = getSheet_(ROOM_SHEET, ROOM_FIELDS);
    const rec = Object.assign({}, room, input);
    if (b.newPassword) {
      if (String(b.newPassword).length < 4) throw new UserError('새 비밀번호는 4자 이상이어야 해요.');
      rec.pwHash = hash_(String(b.newPassword));
    }
    sh.getRange(room._row, 1, 1, ROOM_FIELDS.length).setValues([ROOM_FIELDS.map(([k]) => cell_(rec[k]))]);
    return privateRoom_(getRoom_(room.room));
  });
}

// ---------------------------------------------------------------- 학생

function studentKey_(b) {
  const room = getRoom_(b.room);
  const cls = normNum_(b.cls), num = normNum_(b.num), name = normName_(b.name);
  if (!cls || !num) throw new UserError('반과 번호를 확인해 주세요.');
  if (!name) throw new UserError('이름을 입력해 주세요.');
  if (room.classesList.indexOf(cls) === -1) throw new UserError('이 수업방에 없는 반이에요.');
  return { room: room.room, cls, num, name };
}

function sameStudent_(r, k) {
  return normRoom_(r.room) === k.room && normNum_(r.cls) === k.cls &&
         normNum_(r.num) === k.num && normName_(r.name) === k.name;
}

function cleanRecord_(r) {
  const o = {};
  FIELDS.forEach(([k]) => { o[k] = r[k] instanceof Date ? Utilities.formatDate(r[k], 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss') : r[k]; });
  o.cls = normNum_(o.cls); o.num = normNum_(o.num); o.step = Number(o.step || 0);
  ['check1', 'check2', 'check3', 'anonymous'].forEach(k => { o[k] = o[k] === true || o[k] === 'TRUE'; });
  return o;
}

/** 반·번호·이름이 모두 같은 기록만 불러옵니다. */
function loadStudent_(b) {
  const k = studentKey_(b);
  const rows = readAll_(getSheet_(DATA_SHEET, FIELDS), FIELDS)
    .filter(r => normRoom_(r.room) === k.room && normNum_(r.cls) === k.cls && normNum_(r.num) === k.num);
  const mine = rows.filter(r => sameStudent_(r, k))[0];
  return {
    record: mine ? cleanRecord_(mine) : null,
    // 같은 반·번호로 다른 이름의 기록이 있으면 알려 줌 (이름은 알려 주지 않음)
    otherNameExists: !mine && rows.length > 0,
  };
}

/** 반·번호·이름이 같으면 이어서 저장, 아니면 새 학생으로 추가합니다. */
function saveStudent_(b) {
  const k = studentKey_(b);
  const data = b.data || {};
  return withLock_(() => {
    const sh = getSheet_(DATA_SHEET, FIELDS);
    const existing = readAll_(sh, FIELDS).filter(r => sameStudent_(r, k))[0];
    const rec = existing ? cleanRecord_(existing) : { id: Utilities.getUuid() };
    STUDENT_EDITABLE.forEach(key => { if (data[key] !== undefined) rec[key] = data[key]; });
    Object.assign(rec, k);
    rec.updatedAt = now_();
    rec.step = Math.max(Number(rec.step || 0), Math.min(4, Number(b.step || 0)));
    const values = FIELDS.map(([key]) => cell_(rec[key]));
    if (existing) sh.getRange(existing._row, 1, 1, FIELDS.length).setValues([values]);
    else sh.appendRow(values);
    return { id: rec.id, step: rec.step, savedAt: rec.updatedAt.slice(11), isNew: !existing };
  });
}

// ---------------------------------------------------------------- 교사

function teacherData_(b) {
  const room = auth_(b.room, b.password);
  const rows = readAll_(getSheet_(DATA_SHEET, FIELDS), FIELDS)
    .filter(r => normRoom_(r.room) === room.room)
    .map(cleanRecord_)
    .sort((x, y) => (x.cls - y.cls) || (x.num - y.num) || String(x.name).localeCompare(String(y.name)));
  return { room: privateRoom_(room), rows };
}

function deleteStudent_(b) {
  const room = auth_(b.room, b.password);
  return withLock_(() => {
    const sh = getSheet_(DATA_SHEET, FIELDS);
    const target = readAll_(sh, FIELDS).filter(r => r.id === b.id && normRoom_(r.room) === room.room)[0];
    if (!target) throw new UserError('이미 지워졌거나 찾을 수 없는 기록이에요.');
    sh.deleteRow(target._row);
    return { deleted: b.id };
  });
}
