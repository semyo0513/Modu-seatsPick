// ============================================================
// 창순기획 자리뽑기 - Code.gs  (계정정보 변경 기능 추가)
// ============================================================

const SHEET_NAME_STUDENTS = '학생목록';
const SHEET_NAME_SEATING  = '자리배치';
const SHEET_NAME_HISTORY  = '배치기록';
const SHEET_NAME_RULES    = '규칙설정';
const SHEET_NAME_CONFIG   = '설정';
const SHEET_NAME_USERS    = '회원목록';

function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    try {
      var action = e.parameter.action;
      var payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
      var data = dispatch(action, payload);
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return HtmlService.createHtmlOutput(
    "<h1>모두의 자리뽑기 백엔드 API</h1>" +
    "<p>이 URL은 구글 앱스 스크립트 웹앱 API 엔드포인트입니다.</p>" +
    "<p>깃허브에 배포된 프론트엔드 웹페이지의 'API 설정'에 이 웹앱 URL을 등록하여 사용해주세요.</p>"
  ).setTitle('모두의 자리뽑기 API');
}

function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var payload = requestData.payload || {};
    
    var result = dispatch(action, payload);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ──────────────────────────────────────────────
// 시트 초기화
// ──────────────────────────────────────────────
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  function ensureSheet(name, headers) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      if (headers && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        sh.getRange(1, 1, 1, headers.length)
          .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
        sh.setFrozenRows(1);
      }
    }
    return sh;
  }
  ensureSheet(SHEET_NAME_USERS,    ['이메일','비밀번호해시','이름','가입일','마지막로그인']);
  ensureSheet(SHEET_NAME_STUDENTS, ['계정ID','번호','이름','메모','제외여부']);
  ensureSheet(SHEET_NAME_SEATING,  ['계정ID','배치ID','배치날짜','배치JSON']);
  ensureSheet(SHEET_NAME_HISTORY,  ['계정ID','배치ID','날짜','학생번호','학생이름','행','열']);
  ensureSheet(SHEET_NAME_RULES,    ['계정ID','규칙ID','규칙종류','학생A','학생B','메모']);
  ensureSheet(SHEET_NAME_CONFIG,   ['계정ID','키','값']);
  return { ok: true };
}

// ──────────────────────────────────────────────
// 해시
// ──────────────────────────────────────────────
function hashPassword(password) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8
  );
  return bytes.map(function(b){ return ('0'+(b&0xff).toString(16)).slice(-2); }).join('');
}

// ──────────────────────────────────────────────
// 회원가입
// ──────────────────────────────────────────────
function signUp(payload) {
  initSheets();
  var email    = (payload.email    || '').trim().toLowerCase();
  var password = (payload.password || '').trim();
  var name     = (payload.name     || '').trim();
  if (!email || !password || !name) return { ok:false, error:'모든 항목을 입력해주세요.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok:false, error:'올바른 이메일 형식이 아닙니다.' };
  if (password.length < 6) return { ok:false, error:'비밀번호는 6자 이상이어야 합니다.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME_USERS);
  var last = sh.getLastRow();
  if (last >= 2) {
    var emails = sh.getRange(2,1,last-1,1).getValues().map(function(r){ return r[0]; });
    if (emails.indexOf(email) >= 0) return { ok:false, error:'이미 가입된 이메일입니다.' };
  }
  var now = Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
  sh.appendRow([email, hashPassword(password), name, now, now]);
  return { ok:true, user:{ email:email, name:name } };
}

// ──────────────────────────────────────────────
// 로그인
// ──────────────────────────────────────────────
function signIn(payload) {
  initSheets();
  var email    = (payload.email    || '').trim().toLowerCase();
  var password = (payload.password || '').trim();
  if (!email || !password) return { ok:false, error:'이메일과 비밀번호를 입력해주세요.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME_USERS);
  var last = sh.getLastRow();
  if (last < 2) return { ok:false, error:'등록된 계정이 없습니다.' };
  var data   = sh.getRange(2,1,last-1,5).getValues();
  var hashed = hashPassword(password);
  for (var i=0; i<data.length; i++) {
    if (data[i][0]===email && data[i][1]===hashed) {
      var now = Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
      sh.getRange(i+2,5).setValue(now);
      return { ok:true, user:{ email:email, name:data[i][2] } };
    }
  }
  return { ok:false, error:'이메일 또는 비밀번호가 올바르지 않습니다.' };
}

// ──────────────────────────────────────────────
// ★ 계정 정보 변경 (이름 / 비밀번호)
// ──────────────────────────────────────────────
function updateAccount(payload) {
  initSheets();
  var email       = (payload.email       || '').trim().toLowerCase();
  var currentPw   = (payload.currentPassword || '').trim();
  var newName     = (payload.newName     || '').trim();
  var newPassword = (payload.newPassword || '').trim();

  if (!email || !currentPw) return { ok:false, error:'현재 비밀번호를 입력해주세요.' };
  if (!newName && !newPassword) return { ok:false, error:'변경할 이름 또는 비밀번호를 입력해주세요.' };
  if (newPassword && newPassword.length < 6) return { ok:false, error:'새 비밀번호는 6자 이상이어야 합니다.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME_USERS);
  var last = sh.getLastRow();
  if (last < 2) return { ok:false, error:'등록된 계정이 없습니다.' };

  var data   = sh.getRange(2,1,last-1,5).getValues();
  var hashed = hashPassword(currentPw);

  for (var i=0; i<data.length; i++) {
    if (data[i][0]===email && data[i][1]===hashed) {
      var row = i+2;
      if (newName)     sh.getRange(row,3).setValue(newName);
      if (newPassword) sh.getRange(row,2).setValue(hashPassword(newPassword));
      var updatedName = newName || data[i][2];
      return { ok:true, user:{ email:email, name:updatedName } };
    }
  }
  return { ok:false, error:'현재 비밀번호가 올바르지 않습니다.' };
}


// 계정 정보 변경 (이름 / 비밀번호)
function changeProfile(payload) {
  var email = (payload.accountId || '').trim().toLowerCase();
  var newName = (payload.newName || '').trim();
  var currentPw = (payload.currentPassword || '').trim();
  var newPw = (payload.newPassword || '').trim();
 
  if (!email) return { ok: false, error: '로그인 정보가 없습니다.' };
 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME_USERS);
  var last = sh.getLastRow();
  if (last < 2) return { ok: false, error: '계정을 찾을 수 없습니다.' };
 
  var data = sh.getRange(2, 1, last - 1, 5).getValues();
  var rowIdx = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === email) { rowIdx = i; break; }
  }
  if (rowIdx < 0) return { ok: false, error: '계정을 찾을 수 없습니다.' };
 
  // 현재 비밀번호 확인 (비밀번호 변경 시 필수)
  if (newPw) {
    if (!currentPw) return { ok: false, error: '현재 비밀번호를 입력해주세요.' };
    var currentHash = hashPassword(currentPw);
    if (data[rowIdx][1] !== currentHash) return { ok: false, error: '현재 비밀번호가 올바르지 않습니다.' };
    if (newPw.length < 6) return { ok: false, error: '새 비밀번호는 6자 이상이어야 합니다.' };
    sh.getRange(rowIdx + 2, 2).setValue(hashPassword(newPw));
  }
 
  if (newName) {
    sh.getRange(rowIdx + 2, 3).setValue(newName);
  }
 
  var updatedName = newName || data[rowIdx][2];
  return { ok: true, user: { email: email, name: updatedName } };
}





// ──────────────────────────────────────────────
// 학생
// ──────────────────────────────────────────────
function getStudents(accountId) {
  initSheets();
  if (!accountId) return [];
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_STUDENTS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2,1,last-1,5).getValues();
  return data
    .filter(function(r){ return r[0]===accountId && r[1]!==''; })
    .map(function(r){
      return { number:r[1], name:r[2], memo:r[3], excluded:r[4]===true||r[4]==='TRUE'||r[4]==='제외' };
    });
}

function saveStudents(accountId, students) {
  initSheets();
  if (!accountId) return { ok:false, error:'로그인이 필요합니다.' };
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_STUDENTS);
  var last = sh.getLastRow();
  if (last >= 2) {
    var data = sh.getRange(2,1,last-1,1).getValues();
    for (var i=data.length-1; i>=0; i--) { if(data[i][0]===accountId) sh.deleteRow(i+2); }
  }
  if (!students || students.length===0) return { ok:true };
  var rows = students.map(function(s){
    return [accountId, s.number||'', s.name||'', s.memo||'', s.excluded?'제외':''];
  });
  sh.getRange(sh.getLastRow()+1,1,rows.length,5).setValues(rows);
  return { ok:true };
}

// ──────────────────────────────────────────────
// 규칙
// ──────────────────────────────────────────────
function getRules(accountId) {
  initSheets();
  if (!accountId) return [];
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_RULES);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2,1,last-1,6).getValues();
  return data
    .filter(function(r){ return r[0]===accountId && r[1]!==''; })
    .map(function(r){ return { id:r[1], type:r[2], studentA:r[3], studentB:r[4], memo:r[5] }; });
}

function saveRules(accountId, rules) {
  initSheets();
  if (!accountId) return { ok:false, error:'로그인이 필요합니다.' };
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_RULES);
  var last = sh.getLastRow();
  if (last >= 2) {
    var data = sh.getRange(2,1,last-1,1).getValues();
    for (var i=data.length-1; i>=0; i--) { if(data[i][0]===accountId) sh.deleteRow(i+2); }
  }
  if (!rules || rules.length===0) return { ok:true };
  var rows = rules.map(function(r){
    return [accountId, r.id||'', r.type||'', r.studentA||'', r.studentB||'', r.memo||''];
  });
  sh.getRange(sh.getLastRow()+1,1,rows.length,6).setValues(rows);
  return { ok:true };
}

// ──────────────────────────────────────────────
// 배치 기록
// ──────────────────────────────────────────────
function getSeatingHistory(accountId) {
  initSheets();
  if (!accountId) return [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SEATING);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var data = sh.getRange(2, 1, last - 1, 4).getValues();
  var history = [];

  // 이메일 공백 제거 및 소문자 통일
  var acc = String(accountId).trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var rowAcc = String(data[i][0]).trim().toLowerCase();
    
    if (rowAcc === acc) {
      var parsed = null;
      try {
        parsed = JSON.parse(data[i][3]);
      } catch (e) {
        // 파싱 실패 시 건너뛰지 않고 null로 유지
      }

      // ★ 핵심: 날짜 객체를 완벽한 문자열로 변환 (통신 오류 방지)
      var dateVal = data[i][2];
      var dateStr = "";
      if (Object.prototype.toString.call(dateVal) === '[object Date]') {
        dateStr = Utilities.formatDate(dateVal, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      } else {
        dateStr = String(dateVal);
      }

      history.push({
        id: String(data[i][1]),
        date: dateStr,
        layout: parsed
      });
    }
  }
  
  // 가공된 데이터를 최신순으로 뒤집어서 반환
  return history.reverse();
}



function saveSeating(accountId, layoutJson) {
  initSheets();
  if (!accountId) return { ok:false, error:'로그인이 필요합니다.' };
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName(SHEET_NAME_SEATING);
  var hSh = ss.getSheetByName(SHEET_NAME_HISTORY);
  var now = new Date();
  var id  = Utilities.formatDate(now,'Asia/Seoul','yyyyMMdd_HHmmss');
  var dateStr = Utilities.formatDate(now,'Asia/Seoul','yyyy-MM-dd HH:mm:ss');
  sh.appendRow([accountId, id, dateStr, JSON.stringify(layoutJson)]);
  if (layoutJson && layoutJson.seats) {
    layoutJson.seats.forEach(function(seat){
      if (seat.student) hSh.appendRow([accountId,id,dateStr,seat.student.number,seat.student.name,seat.row,seat.col]);
    });
  }
  var last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2,1,last-1,2).getValues();
    var accountRows = [];
    rows.forEach(function(r,i){ if(r[0]===accountId) accountRows.push(i+2); });
    if (accountRows.length > 10) {
      var toDelete = accountRows.slice(0, accountRows.length-10);
      for (var i=toDelete.length-1; i>=0; i--) sh.deleteRow(toDelete[i]);
    }
  }
  return { ok:true, id:id };
}

function getLastSeatingMap(accountId) {
  var history = getSeatingHistory(accountId);
  if (!history.length) return {};
  var last = history[history.length-1];
  var map  = {};
  if (last.layout && last.layout.seats) {
    last.layout.seats.forEach(function(seat){
      if (seat.student) map[seat.student.number+''] = { row:seat.row, col:seat.col };
    });
  }
  return map;
}

// ──────────────────────────────────────────────
// 설정
// ──────────────────────────────────────────────
function getConfig(accountId) {
  initSheets();
  if (!accountId) return {};
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_CONFIG);
  var last = sh.getLastRow();
  if (last < 2) return {};
  var data = sh.getRange(2,1,last-1,3).getValues();
  var cfg  = {};
  data.forEach(function(r){ if(r[0]===accountId && r[1]) cfg[r[1]]=r[2]; });
  return cfg;
}

function saveConfig(accountId, cfg) {
  initSheets();
  if (!accountId) return { ok:false, error:'로그인이 필요합니다.' };
  var sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_CONFIG);
  var last = sh.getLastRow();
  if (last >= 2) {
    var data = sh.getRange(2,1,last-1,1).getValues();
    for (var i=data.length-1; i>=0; i--) { if(data[i][0]===accountId) sh.deleteRow(i+2); }
  }
  var rows = Object.keys(cfg).map(function(k){ return [accountId, k, cfg[k]]; });
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,3).setValues(rows);
  return { ok:true };
}

// ──────────────────────────────────────────────
// 랜덤 배치
// ──────────────────────────────────────────────
function generateRandomSeating(accountId, seatLayout, students, rules, avoidLastSeating) {
  var activeSeats      = seatLayout.filter(function(s){ return s.active; });
  var eligibleStudents = students.filter(function(s){ return !s.excluded; });
  var lastMap = (avoidLastSeating && accountId) ? getLastSeatingMap(accountId) : {};

  // 활성 좌석의 최소 행 번호 계산 (칠판 쪽 = 행 번호가 가장 작은 쪽)
  var rowNums = activeSeats.map(function(s){ return s.row; });
  var minRow = rowNums.length ? Math.min.apply(null, rowNums) : 1;

  var MAX_ATTEMPTS = 500;
  for (var attempt=0; attempt<MAX_ATTEMPTS; attempt++) {
    var shuffled = shuffleArray(eligibleStudents.slice());
    var result   = assignSeats(activeSeats, shuffled);
    if (checkRules(result, rules, minRow) && checkAvoidLast(result, lastMap)) return { ok:true, seats:result };
  }
  var shuffled = shuffleArray(eligibleStudents.slice());
  return { ok:true, seats:assignSeats(activeSeats, shuffled), warning:'일부 규칙을 충족하지 못했습니다.' };
}

function assignSeats(activeSeats, students) {
  return activeSeats.map(function(seat, i){
    return { row:seat.row, col:seat.col, student: i<students.length ? students[i] : null };
  });
}

function checkRules(seats, rules, minRow) {
  if (!rules || rules.length===0) return true;
  // minRow가 없으면 1로 폴백 (칠판 = 가장 작은 행 번호)
  var frontMin = (typeof minRow === 'number') ? minRow : 1;
  var frontMax = frontMin + 1; // 앞 2행을 앞자리로 간주
  var posMap = {};
  seats.forEach(function(s){ if(s.student) posMap[s.student.number+'']={row:s.row,col:s.col}; });
  return rules.every(function(rule){
    var posA = posMap[rule.studentA+''];
    if (rule.type==='front') {
      if (!posA) return true;
      // 칠판 쪽(행 번호가 작은 쪽) 1~2행에 배치되어야 함
      return posA.row <= frontMax;
    }
    var posB = posMap[rule.studentB+''];
    if (!posA || !posB) return true;
    var adjacent = Math.abs(posA.row-posB.row)+Math.abs(posA.col-posB.col) <= 1;
    if (rule.type==='separate') return !adjacent;
    if (rule.type==='together') return adjacent;
    return true;
  });
}

function checkAvoidLast(seats, lastMap) {
  if (!lastMap || Object.keys(lastMap).length===0) return true;
  return seats.every(function(s){
    if (!s.student) return true;
    var last = lastMap[s.student.number+''];
    if (!last) return true;
    return !(last.row===s.row && last.col===s.col);
  });
}

function shuffleArray(arr) {
  for (var i=arr.length-1; i>0; i--) {
    var j = Math.floor(Math.random()*(i+1));
    var tmp=arr[i]; arr[i]=arr[j]; arr[j]=tmp;
  }
  return arr;
}

// ──────────────────────────────────────────────
// dispatch
// ──────────────────────────────────────────────
function dispatch(action, payload) {
  payload = payload || {};
  var accountId = payload.accountId || null;
  switch(action) {
    case 'initSheets':        return initSheets();
    case 'signUp':            return signUp(payload);
    case 'signIn':            return signIn(payload);
    case 'updateAccount':     return updateAccount(payload);          // ★ 신규
    case 'getStudents':       return getStudents(accountId);
    case 'saveStudents':      return saveStudents(accountId, payload.students);
    case 'getRules':          return getRules(accountId);
    case 'saveRules':         return saveRules(accountId, payload.rules);
    case 'getSeatingHistory': return getSeatingHistory(accountId);
    case 'saveSeating':       return saveSeating(accountId, payload.layout);
    case 'getLastSeatingMap': return getLastSeatingMap(accountId);
    case 'getConfig':         return getConfig(accountId);
    case 'saveConfig':        return saveConfig(accountId, payload.config);
    case 'changeProfile':     return changeProfile(payload);
    case 'generateRandomSeating':
      return generateRandomSeating(accountId, payload.seatLayout, payload.students, payload.rules, payload.avoidLastSeating);
    default:
      return { error:'Unknown action: '+action };
  }
}