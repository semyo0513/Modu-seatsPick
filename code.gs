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
// 랜덤 배치 (규칙 우선 제약 충족 최적화 솔버)
// ──────────────────────────────────────────────
function generateRandomSeating(accountId, seatLayout, students, rules, avoidLastSeating) {
  var activeSeats      = seatLayout.filter(function(s){ return s.active; });
  var eligibleStudents = students.filter(function(s){ return !s.excluded; });
  var lastMap = (avoidLastSeating && accountId) ? getLastSeatingMap(accountId) : {};

  return solveSeatingAlgorithm(activeSeats, eligibleStudents, rules, lastMap);
}

function solveSeatingAlgorithm(activeSeats, eligibleStudents, rules, lastMap) {
  if (!activeSeats || !activeSeats.length || !eligibleStudents || !eligibleStudents.length) {
    return { ok: true, seats: [] };
  }

  // 좌석 그리드 크기 및 칠판(앞자리) 행 계산
  var rowNums = activeSeats.map(function(s){ return s.row; });
  var colNums = activeSeats.map(function(s){ return s.col; });
  var minRow = Math.min.apply(null, rowNums);
  var maxCol = Math.max.apply(null, colNums);
  var frontMax = minRow + 1; // 앞 2행 (칠판 쪽 = 가장 작은 행 번호)

  // 1. 유니버설 학생 식별자 해결기
  function resolveStudent(ref) {
    if (ref === undefined || ref === null) return null;
    var raw = String(ref).trim();
    if (!raw) return null;

    var numFromRaw = parseInt(raw.replace(/[^0-9]/g, ''), 10);

    for (var i = 0; i < eligibleStudents.length; i++) {
      var s = eligibleStudents[i];
      var sNumStr = String(s.number).trim();
      if (sNumStr === raw) return s;
      if (!isNaN(numFromRaw) && parseInt(s.number, 10) === numFromRaw) {
        if (/^\d+(\.0+)?$/.test(raw) || /^\d+번/.test(raw)) {
          return s;
        }
      }
    }

    for (var i = 0; i < eligibleStudents.length; i++) {
      var s = eligibleStudents[i];
      var sName = s.name ? String(s.name).trim() : '';
      if (sName && (sName === raw || raw.indexOf(sName) >= 0)) return s;
    }

    if (!isNaN(numFromRaw)) {
      for (var i = 0; i < eligibleStudents.length; i++) {
        if (parseInt(eligibleStudents[i].number, 10) === numFromRaw) {
          return eligibleStudents[i];
        }
      }
    }

    return null;
  }

  // 2. 규칙 정제 및 표준 학생 식별자 매핑
  var cleanRules = [];
  (rules || []).forEach(function(r){
    if (!r || !r.type) return;
    var stA = resolveStudent(r.studentA);
    if (!stA) return;
    var idA = String(stA.number).trim();

    if (r.type === 'front') {
      cleanRules.push({ type: 'front', studentA: idA, studentB: '' });
    } else {
      var stB = resolveStudent(r.studentB);
      if (!stB) return;
      var idB = String(stB.number).trim();
      if (idA !== idB) {
        cleanRules.push({ type: r.type, studentA: idA, studentB: idB });
      }
    }
  });

  // 3. 인접 필수 (짝꿍): 같은 행, 짝수열인 경우 1-2열, 3-4열, 5-6열 등 짝지 배치 검증
  function isValidTogether(posA, posB) {
    if (!posA || !posB) return false;
    if (posA.row !== posB.row) return false;
    var c1 = Math.min(posA.col, posB.col);
    var c2 = Math.max(posA.col, posB.col);
    if (c2 - c1 !== 1) return false;

    if (maxCol % 2 === 0) {
      // 짝수열인 경우: 1-2열, 3-4열, 5-6열
      return c1 % 2 === 1;
    } else {
      return true;
    }
  }

  // 4. 인접 금지: 양옆(좌우) 및 앞뒤 좌석 배치 절대 금지 검증
  function isAdjacent(posA, posB) {
    if (!posA || !posB) return false;
    // 양옆: 같은 행, 열 차이 1
    if (posA.row === posB.row && Math.abs(posA.col - posB.col) === 1) return true;
    // 앞뒤: 같은 열, 행 차이 1
    if (posA.col === posB.col && Math.abs(posA.row - posB.row) === 1) return true;
    return false;
  }

  // 5. 페널티 계산 함수
  function calcPenalty(assignment) {
    var posMap = {};
    for (var i = 0; i < assignment.length; i++) {
      var s = assignment[i];
      if (s.student) {
        posMap[String(s.student.number).trim()] = s;
      }
    }

    var ruleViolations = 0;
    var rulePenalty = 0;
    var conflictedStudentIds = {};

    for (var r = 0; r < cleanRules.length; r++) {
      var rule = cleanRules[r];
      var posA = posMap[rule.studentA];
      if (rule.type === 'front') {
        if (posA) {
          if (posA.row > frontMax) {
            ruleViolations++;
            rulePenalty += 20000 + (posA.row - frontMax) * 5000;
            conflictedStudentIds[rule.studentA] = true;
          } else if (posA.row > minRow) {
            rulePenalty += 1;
          }
        }
      } else {
        var posB = posMap[rule.studentB];
        if (posA && posB) {
          if (rule.type === 'together') {
            if (!isValidTogether(posA, posB)) {
              ruleViolations++;
              var dist = Math.abs(posA.row - posB.row) + Math.abs(posA.col - posB.col);
              rulePenalty += 20000 + dist * 1000;
              conflictedStudentIds[rule.studentA] = true;
              conflictedStudentIds[rule.studentB] = true;
            }
          } else if (rule.type === 'separate') {
            if (isAdjacent(posA, posB)) {
              ruleViolations++;
              rulePenalty += 20000;
              conflictedStudentIds[rule.studentA] = true;
              conflictedStudentIds[rule.studentB] = true;
            }
          }
        }
      }
    }

    var avoidViolations = 0;
    if (lastMap && Object.keys(lastMap).length > 0) {
      for (var i = 0; i < assignment.length; i++) {
        var st = assignment[i].student;
        if (st) {
          var last = lastMap[String(st.number).trim()];
          if (last && last.row === assignment[i].row && last.col === assignment[i].col) {
            avoidViolations++;
          }
        }
      }
    }

    return {
      ruleViolations: ruleViolations,
      rulePenalty: rulePenalty,
      avoidViolations: avoidViolations,
      conflictedStudentIds: conflictedStudentIds,
      score: rulePenalty * 10000 + avoidViolations
    };
  }

  function shuffle(arr) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = a[i]; a[i] = a[j]; a[j] = t;
    }
    return a;
  }

  // 6. 스마트 초기 배치
  function buildInitialPlacement() {
    var seats = activeSeats.map(function(s){
      return { row: s.row, col: s.col, student: null };
    });

    var studentMap = {};
    for (var i = 0; i < eligibleStudents.length; i++) {
      studentMap[String(eligibleStudents[i].number).trim()] = eligibleStudents[i];
    }

    var placedStudentNums = {};
    var occupiedSeatIndices = {};

    function getEmptySeatIndices(filterFn) {
      var res = [];
      for (var i = 0; i < seats.length; i++) {
        if (!occupiedSeatIndices[i]) {
          if (!filterFn || filterFn(seats[i])) {
            res.push(i);
          }
        }
      }
      return res;
    }

    function canPlaceWithoutSeparateViolation(sNum, seat) {
      for (var r = 0; r < cleanRules.length; r++) {
        var rule = cleanRules[r];
        if (rule.type === 'separate') {
          var otherId = null;
          if (rule.studentA === sNum) otherId = rule.studentB;
          else if (rule.studentB === sNum) otherId = rule.studentA;

          if (otherId && placedStudentNums[otherId]) {
            for (var i = 0; i < seats.length; i++) {
              if (seats[i].student && String(seats[i].student.number).trim() === otherId) {
                if (isAdjacent(seat, seats[i])) {
                  return false;
                }
                break;
              }
            }
          }
        }
      }
      return true;
    }

    // A. '인접 필수(together)' 짝지 배치 (1-2열, 3-4열, 5-6열 분단 구조 엄격 준수)
    var togetherRules = cleanRules.filter(function(r){ return r.type === 'together'; });
    var frontRuleStudents = {};
    cleanRules.forEach(function(r){
      if (r.type === 'front') frontRuleStudents[r.studentA] = true;
    });

    var shuffledTogether = shuffle(togetherRules);
    shuffledTogether.forEach(function(rule){
      var sA = rule.studentA;
      var sB = rule.studentB;
      if (!studentMap[sA] || !studentMap[sB]) return;
      if (placedStudentNums[sA] || placedStudentNums[sB]) return;

      var hasFront = frontRuleStudents[sA] || frontRuleStudents[sB];
      var emptyIndices = shuffle(getEmptySeatIndices());

      var bestPair = null;
      for (var i = 0; i < emptyIndices.length; i++) {
        var idx1 = emptyIndices[i];
        var seat1 = seats[idx1];
        if (hasFront && seat1.row > frontMax) continue;

        for (var j = 0; j < emptyIndices.length; j++) {
          if (i === j) continue;
          var idx2 = emptyIndices[j];
          var seat2 = seats[idx2];
          if (hasFront && seat2.row > frontMax) continue;

          if (isValidTogether(seat1, seat2)) {
            if (canPlaceWithoutSeparateViolation(sA, seat1) && canPlaceWithoutSeparateViolation(sB, seat2)) {
              bestPair = [idx1, idx2];
              break;
            }
          }
        }
        if (bestPair) break;
      }

      if (bestPair) {
        seats[bestPair[0]].student = studentMap[sA];
        seats[bestPair[1]].student = studentMap[sB];
        placedStudentNums[sA] = true;
        placedStudentNums[sB] = true;
        occupiedSeatIndices[bestPair[0]] = true;
        occupiedSeatIndices[bestPair[1]] = true;
      }
    });

    // B. 나머지 '앞자리(front)' 학생 배치
    for (var fNum in frontRuleStudents) {
      if (placedStudentNums[fNum]) continue;
      if (!studentMap[fNum]) continue;

      var frontEmptyIndices = shuffle(getEmptySeatIndices(function(s){ return s.row <= frontMax; }));
      var chosenIdx = -1;
      for (var f = 0; f < frontEmptyIndices.length; f++) {
        var idx = frontEmptyIndices[f];
        if (canPlaceWithoutSeparateViolation(fNum, seats[idx])) {
          chosenIdx = idx;
          break;
        }
      }
      if (chosenIdx < 0 && frontEmptyIndices.length > 0) {
        chosenIdx = frontEmptyIndices[0];
      }

      if (chosenIdx >= 0) {
        seats[chosenIdx].student = studentMap[fNum];
        placedStudentNums[fNum] = true;
        occupiedSeatIndices[chosenIdx] = true;
      }
    }

    // C. 나머지 분리(separate) 규칙 대상 학생 배치
    var separateStudents = {};
    cleanRules.forEach(function(r){
      if (r.type === 'separate') {
        if (!placedStudentNums[r.studentA]) separateStudents[r.studentA] = true;
        if (!placedStudentNums[r.studentB]) separateStudents[r.studentB] = true;
      }
    });

    for (var sepNum in separateStudents) {
      if (placedStudentNums[sepNum]) continue;
      if (!studentMap[sepNum]) continue;

      var emptyIndices = shuffle(getEmptySeatIndices());
      var chosenIdx = -1;
      for (var e = 0; e < emptyIndices.length; e++) {
        var idx = emptyIndices[e];
        if (canPlaceWithoutSeparateViolation(sepNum, seats[idx])) {
          chosenIdx = idx;
          break;
        }
      }
      if (chosenIdx < 0 && emptyIndices.length > 0) {
        chosenIdx = emptyIndices[0];
      }

      if (chosenIdx >= 0) {
        seats[chosenIdx].student = studentMap[sepNum];
        placedStudentNums[sepNum] = true;
        occupiedSeatIndices[chosenIdx] = true;
      }
    }

    // D. 나머지 일반 학생들 무작위 배치
    var remainingStudents = eligibleStudents.filter(function(s){
      return !placedStudentNums[String(s.number).trim()];
    });
    remainingStudents = shuffle(remainingStudents);

    var remainingEmptyIndices = shuffle(getEmptySeatIndices());
    for (var rIdx = 0; rIdx < remainingStudents.length && rIdx < remainingEmptyIndices.length; rIdx++) {
      var seatIdx = remainingEmptyIndices[rIdx];
      seats[seatIdx].student = remainingStudents[rIdx];
      occupiedSeatIndices[seatIdx] = true;
    }

    return seats;
  }

  // 7. 로컬 서치
  var bestSeats = null;
  var bestPenalty = { score: Infinity, ruleViolations: Infinity, avoidViolations: Infinity };

  var NUM_RESTARTS = 12;
  var MAX_STEPS = 1500;

  for (var restart = 0; restart < NUM_RESTARTS; restart++) {
    var currentSeats = buildInitialPlacement();
    var currentPenalty = calcPenalty(currentSeats);

    if (currentPenalty.score < bestPenalty.score) {
      bestPenalty = currentPenalty;
      bestSeats = JSON.parse(JSON.stringify(currentSeats));
      if (bestPenalty.ruleViolations === 0 && (!lastMap || bestPenalty.avoidViolations === 0)) {
        break;
      }
    }

    for (var step = 0; step < MAX_STEPS; step++) {
      if (currentPenalty.ruleViolations === 0 && (!lastMap || currentPenalty.avoidViolations === 0)) {
        break;
      }

      var idxA = -1;
      var conflictedIds = Object.keys(currentPenalty.conflictedStudentIds || {});
      if (conflictedIds.length > 0 && Math.random() < 0.85) {
        var targetId = conflictedIds[Math.floor(Math.random() * conflictedIds.length)];
        for (var i = 0; i < currentSeats.length; i++) {
          if (currentSeats[i].student && String(currentSeats[i].student.number).trim() === targetId) {
            idxA = i;
            break;
          }
        }
      }
      if (idxA < 0) {
        idxA = Math.floor(Math.random() * currentSeats.length);
      }

      var idxB = Math.floor(Math.random() * currentSeats.length);
      while (idxB === idxA && currentSeats.length > 1) {
        idxB = Math.floor(Math.random() * currentSeats.length);
      }

      var tmp = currentSeats[idxA].student;
      currentSeats[idxA].student = currentSeats[idxB].student;
      currentSeats[idxB].student = tmp;

      var newPenalty = calcPenalty(currentSeats);

      var accept = false;
      if (newPenalty.score < currentPenalty.score) {
        accept = true;
      } else if (newPenalty.score === currentPenalty.score && Math.random() < 0.20) {
        accept = true;
      }

      if (accept) {
        currentPenalty = newPenalty;
        if (currentPenalty.score < bestPenalty.score) {
          bestPenalty = currentPenalty;
          bestSeats = JSON.parse(JSON.stringify(currentSeats));
          if (bestPenalty.ruleViolations === 0 && (!lastMap || bestPenalty.avoidViolations === 0)) {
            break;
          }
        }
      } else {
        // 롤백
        currentSeats[idxB].student = currentSeats[idxA].student;
        currentSeats[idxA].student = tmp;
      }
    }

    if (bestPenalty.ruleViolations === 0 && (!lastMap || bestPenalty.avoidViolations === 0)) {
      break;
    }
  }

  var warningMsg = null;
  if (bestPenalty.ruleViolations > 0) {
    warningMsg = '일부 규칙이 서로 상충되어 가능한 최대 범위 내에서 규칙을 우선 적용했습니다.';
  }

  return {
    ok: true,
    seats: bestSeats || activeSeats.map(function(s){ return { row: s.row, col: s.col, student: null }; }),
    warning: warningMsg
  };
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