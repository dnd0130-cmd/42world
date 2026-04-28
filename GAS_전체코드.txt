function doGet(e) { return handleRequest(e); }

function doPost(e) {
  if (e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body.action === 'submitMission') {
        return ContentService.createTextOutput(JSON.stringify(submitMission(body)))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (body.action === 'updateMissionSubmission') {
        return ContentService.createTextOutput(JSON.stringify(updateMissionSubmission(body)))
          .setMimeType(ContentService.MimeType.JSON);
      }
    } catch(err) {}
  }
  return handleRequest(e);
}

function handleRequest(e) {
  var params = e.parameter;
  var action = params.action;
  var result;
  try {
    switch (action) {
      case 'writeGuestbook':               result = writeGuestbook(params.fromCode, params.toCode, params.content, params.authorType); break;
      case 'getGuestbook':                 result = getGuestbook(params.code); break;
      case 'login':                        result = login(params.code); break;
      case 'getStudent':                   result = getStudent(params.code); break;
      case 'getClassmates':                result = getClassmates(params.code); break;
      case 'getShop':                      result = getShop(); break;
      case 'buyItem':                      result = buyItem(params.code, params.itemName); break;
      case 'sendPraise':                   result = sendPraise(params.fromCode, params.toCode, params.message); break;
      case 'submitDiary':                  result = submitDiary(params.code, params.emotion, params.content); break;
      case 'getAllStudents':               result = getAllStudents(); break;
      case 'giveAcorn':                    result = giveAcorn(params.code, params.amount, params.reason); break;
      case 'getLog':                       result = getLog(); break;
      case 'getDiaries':                   result = getDiaries(); break;
      case 'addShopItem':                  result = addShopItem(params.name, params.price, params.qty, params.desc); break;
      case 'deleteShopItem':              result = deleteShopItem(params.name); break;
      case 'getVouchers':                  result = getVouchers(params.code); break;
      case 'useVoucher':                   result = useVoucher(params.id, params.code); break;
      case 'getPendingVouchers':          result = getPendingVouchers(); break;
      case 'approveVoucher':              result = approveVoucher(params.id); break;
      case 'rejectVoucher':               result = rejectVoucher(params.id); break;
      case 'getPurchaseHistory':          result = getPurchaseHistory(params.code); break;
      case 'cancelAcorn':                  result = cancelAcorn(params.rowNum); break;
      case 'archiveLogs':                  result = archiveLogs(); break;
      case 'getMissions':                  result = getMissions(params); break;
      case 'createMission':               result = createMission(params); break;
      case 'toggleMission':               result = toggleMission(params); break;
      case 'deleteMission':               result = deleteMission(params); break;
      case 'getMissionSubmissions':       result = getMissionSubmissions(params); break;
      case 'getPendingMissionSubmissions': result = getPendingMissionSubmissions(); break;
      case 'approveMissionSubmission':    result = approveMissionSubmission(params); break;
      case 'rejectMissionSubmission':     result = rejectMissionSubmission(params); break;
      case 'withdrawMissionSubmission':   result = withdrawMissionSubmission(params); break;
      default:                             result = { success: false, error: '알 수 없는 요청입니다.' }; break;
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ===== 월별 보관 =====
function archiveMonthlyLogs() { archiveLogs(); }

function archiveLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var year = lastMonth.getFullYear();
  var month = lastMonth.getMonth() + 1;
  var label = year + '_' + (month < 10 ? '0' + month : month);
  var targets = [
    { sheetName: '도토리 로그', dateCol: 0 },
    { sheetName: '칭찬 기록',   dateCol: 0 },
    { sheetName: '방명록',      dateCol: 0 },
    { sheetName: '감정일기',    dateCol: 0 }
  ];
  var archived = [];
  targets.forEach(function(t) {
    var src = ss.getSheetByName(t.sheetName);
    if (!src) return;
    var data = src.getDataRange().getValues();
    if (data.length <= 1) return;
    var header = data[0];
    var toMove = [];
    var toKeep = [header];
    for (var i = 1; i < data.length; i++) {
      var d = new Date(data[i][t.dateCol]);
      if (d.getFullYear() === year && d.getMonth() + 1 === month) {
        toMove.push(data[i]);
      } else {
        toKeep.push(data[i]);
      }
    }
    if (toMove.length === 0) return;
    var archiveName = t.sheetName + '_' + label;
    var archiveSheet = ss.getSheetByName(archiveName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveName);
      archiveSheet.appendRow(header);
    }
    toMove.forEach(function(row) { archiveSheet.appendRow(row); });
    src.clearContents();
    toKeep.forEach(function(row) { src.appendRow(row); });
    archived.push(t.sheetName + ' ' + toMove.length + '행');
  });
  if (archived.length === 0) return { success: true, message: '보관할 지난달 데이터가 없어요.' };
  return { success: true, message: label + ' 보관 완료: ' + archived.join(', ') };
}

// ===== 레벨 =====
function getLevelInfo(exp) {
  var levels = [
    { level:1,  exp:0,    name:'새싹' },
    { level:2,  exp:100,  name:'도토리' },
    { level:3,  exp:300,  name:'동잎' },
    { level:4,  exp:600,  name:'은별' },
    { level:5,  exp:1000, name:'금별' },
    { level:6,  exp:1500, name:'루비' },
    { level:7,  exp:2200, name:'사파이어' },
    { level:8,  exp:3000, name:'에메랄드' },
    { level:9,  exp:4000, name:'다이아' },
    { level:10, exp:5500, name:'크리스탈' },
    { level:11, exp:7000, name:'플래티넘' },
    { level:12, exp:9000, name:'레전드' }
  ];
  var result = levels[0];
  for (var i = 0; i < levels.length; i++) {
    if (exp >= levels[i].exp) result = levels[i];
  }
  return result;
}

// ===== 학생 =====
function login(code) {
  if (!code) return { success: false, error: '코드를 입력하세요.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('학생 명단');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(code)) {
      return { success: true, student: {
        학번: data[i][0], 이름: data[i][1], 고유코드: data[i][2],
        도토리: data[i][3] || 0, 경험치: data[i][4] || 0,
        레벨: data[i][5] || 1, 뱃지: data[i][6] || ''
      }};
    }
  }
  return { success: false, error: '일촌 번호를 찾을 수 없습니다.' };
}

function getStudent(code) { return login(code); }

function getClassmates(code) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('학생 명단');
  var data = sheet.getDataRange().getValues();
  var classmates = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) !== String(code)) {
      classmates.push({ 학번: data[i][0], 이름: data[i][1], 고유코드: data[i][2] });
    }
  }
  return { success: true, classmates: classmates };
}

function getAllStudents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('학생 명단');
  var data = sheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      students.push({
        학번: data[i][0], 이름: data[i][1], 고유코드: data[i][2],
        도토리: data[i][3] || 0, 경험치: data[i][4] || 0, 레벨: data[i][5] || 1, 뱃지: data[i][6] || ''
      });
    }
  }
  return { success: true, students: students };
}

// ===== 도토리 =====
function giveAcorn(code, amount, reason) {
  if (!code || !amount) return { success: false, error: '잘못된 요청입니다.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('학생 명단');
  var logSheet = ss.getSheetByName('도토리 로그');
  var data = studentSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(code)) {
      var currentDotori = Number(data[i][3]) || 0;
      var currentExp = Number(data[i][4]) || 0;
      var oldLevel = Number(data[i][5]) || 1;
      var delta = Number(amount);
      var newDotori = Math.max(0, currentDotori + delta);
      var expGain = delta > 0 ? delta * 10 : 0;
      var newExp = currentExp + expGain;
      var levelInfo = getLevelInfo(newExp);
      studentSheet.getRange(i + 1, 4).setValue(newDotori);
      studentSheet.getRange(i + 1, 5).setValue(newExp);
      studentSheet.getRange(i + 1, 6).setValue(levelInfo.level);
      studentSheet.getRange(i + 1, 7).setValue(levelInfo.name);
      logSheet.appendRow([new Date(), data[i][1], delta, reason || '교사 지급', '교사']);
      return { success: true, 도토리: newDotori, 경험치: newExp, 레벨: levelInfo.level, 뱃지: levelInfo.name, 레벨업: levelInfo.level > oldLevel, 이전레벨: oldLevel };
    }
  }
  return { success: false, error: '학생을 찾을 수 없습니다.' };
}

function getLog() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('도토리 로그');
  var data = sheet.getDataRange().getValues();
  var log = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0]) {
      log.push({ rowNum: i + 1, 날짜: data[i][0], 이름: data[i][1], 변화량: data[i][2], 사유: data[i][3] });
      if (log.length >= 20) break;
    }
  }
  return { success: true, log: log };
}

function cancelAcorn(rowNum) {
  if (!rowNum) return { success: false, error: '잘못된 요청입니다.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('도토리 로그');
  var studentSheet = ss.getSheetByName('학생 명단');
  var row = parseInt(rowNum);
  var logRow = logSheet.getRange(row, 1, 1, 5).getValues()[0];
  var studentName = logRow[1];
  var amount = Number(logRow[2]);
  var reason = logRow[3];
  if (String(reason).indexOf('취소:') === 0) return { success: false, error: '이미 취소된 항목입니다.' };
  var studentData = studentSheet.getDataRange().getValues();
  for (var i = 1; i < studentData.length; i++) {
    if (studentData[i][1] === studentName) {
      var currentDotori = Number(studentData[i][3]) || 0;
      studentSheet.getRange(i + 1, 4).setValue(Math.max(0, currentDotori - amount));
      logSheet.appendRow([new Date(), studentName, -amount, '취소: ' + reason, '교사']);
      return { success: true, message: studentName + '의 [' + reason + '] 취소 완료!' };
    }
  }
  return { success: false, error: '학생을 찾을 수 없습니다.' };
}

// ===== 칭찬 =====
function sendPraise(fromCode, toCode, message) {
  if (!fromCode || !toCode || !message) return { success: false, error: '잘못된 요청입니다.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('학생 명단');
  var praiseSheet = ss.getSheetByName('칭찬 기록');
  var logSheet = ss.getSheetByName('도토리 로그');
  var data = studentSheet.getDataRange().getValues();
  var fromRow = -1, fromName = '', toRow = -1, toName = '', toDotori = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(fromCode)) { fromRow = i + 1; fromName = data[i][1]; }
    if (String(data[i][2]) === String(toCode))   { toRow = i + 1; toName = data[i][1]; toDotori = Number(data[i][3]) || 0; }
  }
  if (fromRow === -1) return { success: false, error: '보내는 학생을 찾을 수 없습니다.' };
  if (toRow === -1)   return { success: false, error: '받는 친구를 찾을 수 없습니다.' };
  var today = new Date(); today.setHours(0,0,0,0);
  var praiseData = praiseSheet.getDataRange().getValues();
  var todayCount = 0;
  for (var j = 1; j < praiseData.length; j++) {
    var pd = new Date(praiseData[j][0]); pd.setHours(0,0,0,0);
    if (praiseData[j][1] === fromName && pd.getTime() === today.getTime()) todayCount++;
  }
  if (todayCount >= 1) return { success: false, error: '오늘 칭찬을 모두 사용했어요! 내일 다시 해주세요.' };
  studentSheet.getRange(toRow, 4).setValue(toDotori + 1);
  var toExp = Number(data[toRow - 1][4]) || 0;
  var newToExp = toExp + 10;
  var toLevelInfo = getLevelInfo(newToExp);
  studentSheet.getRange(toRow, 5).setValue(newToExp);
  studentSheet.getRange(toRow, 6).setValue(toLevelInfo.level);
  studentSheet.getRange(toRow, 7).setValue(toLevelInfo.name);
  praiseSheet.appendRow([new Date(), fromName, toName, message, 1]);
  logSheet.appendRow([new Date(), toName, 1, fromName + '의 칭찬', fromName]);
  return { success: true, message: toName + '에게 칭찬과 도토리 1개를 보냈어요!', 받는친구: toName, 남은칭찬: 0 };
}

// ===== 감정일기 =====
function submitDiary(code, emotion, content) {
  if (!code || !emotion || !content) return { success: false, error: '모든 항목을 입력해주세요.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('학생 명단');
  var diarySheet = ss.getSheetByName('감정일기');
  var data = studentSheet.getDataRange().getValues();
  var studentName = '', studentCode = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(code)) {
      studentName = data[i][1]; studentCode = data[i][2]; break;
    }
  }
  if (!studentName) return { success: false, error: '학생을 찾을 수 없습니다.' };
  diarySheet.appendRow([new Date(), studentCode, studentName, emotion, content]);
  return { success: true, message: '감정일기가 저장되었습니다!', 감정: emotion };
}

function getDiaries() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('감정일기');
  var data = sheet.getDataRange().getValues();
  var diaries = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0]) {
      diaries.push({ 날짜: data[i][0], 학번: data[i][1], 이름: data[i][2], 감정: data[i][3], 내용: data[i][4] });
    }
  }
  return { success: true, diaries: diaries };
}

// ===== 상점 =====
function getShop() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('상점 리스트');
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      items.push({ 아이템이름: data[i][0], 필요도토리: data[i][1], 상세설명: data[i][2], 남은수량: data[i][3] });
    }
  }
  return { success: true, items: items };
}

function buyItem(code, itemName) {
  if (!code || !itemName) return { success: false, error: '잘못된 요청입니다.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('학생 명단');
  var shopSheet = ss.getSheetByName('상점 리스트');
  var logSheet = ss.getSheetByName('도토리 로그');
  var studentData = studentSheet.getDataRange().getValues();
  var studentRow = -1, studentName = '', currentDotori = 0;
  for (var i = 1; i < studentData.length; i++) {
    if (String(studentData[i][2]) === String(code)) {
      studentRow = i + 1; studentName = studentData[i][1]; currentDotori = Number(studentData[i][3]) || 0; break;
    }
  }
  if (studentRow === -1) return { success: false, error: '학생을 찾을 수 없습니다.' };
  var shopData = shopSheet.getDataRange().getValues();
  var itemRow = -1, itemPrice = 0, itemQty = 0;
  for (var j = 1; j < shopData.length; j++) {
    if (String(shopData[j][0]) === String(itemName)) {
      itemRow = j + 1; itemPrice = Number(shopData[j][1]) || 0; itemQty = shopData[j][3]; break;
    }
  }
  if (itemRow === -1) return { success: false, error: '아이템을 찾을 수 없습니다.' };
  if (itemQty !== '' && Number(itemQty) <= 0) return { success: false, error: '품절된 아이템입니다.' };
  if (currentDotori < itemPrice) return { success: false, error: '도토리가 부족합니다. (보유: ' + currentDotori + ' / 필요: ' + itemPrice + ')' };
  studentSheet.getRange(studentRow, 4).setValue(currentDotori - itemPrice);
  if (itemQty !== '') shopSheet.getRange(itemRow, 4).setValue(Number(itemQty) - 1);
  logSheet.appendRow([new Date(), studentName, -itemPrice, '상점 구매: ' + itemName, '시스템']);
  var voucherSheet = ss.getSheetByName('이용권');
  if (voucherSheet) {
    voucherSheet.appendRow([new Date().getTime(), new Date(), code, studentName, itemName, '보관중', '']);
  }
  return { success: true, message: itemName + ' 구매 완료!', 남은도토리: currentDotori - itemPrice };
}

function addShopItem(name, price, qty, desc) {
  if (!name || !price) return { success: false, error: '이름과 가격을 입력하세요.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('상점 리스트');
  sheet.appendRow([name, Number(price), desc || '', qty !== '' && qty ? Number(qty) : '']);
  return { success: true, message: name + ' 등록 완료!' };
}

function deleteShopItem(name) {
  if (!name) return { success: false, error: '잘못된 요청입니다.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('상점 리스트');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(name)) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, error: '아이템을 찾을 수 없습니다.' };
}

// ===== 이용권 =====
function getVouchers(code) {
  if (!code) return { success: false, error: '잘못된 요청입니다.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이용권');
  if (!sheet) return { success: true, vouchers: [] };
  var data = sheet.getDataRange().getValues();
  var vouchers = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]) === String(code) && data[i][5] !== '사용완료' && data[i][5] !== '거절') {
      vouchers.push({ id: data[i][0], 구매일: data[i][1], 품목: data[i][4], 상태: data[i][5], 신청일: data[i][6] });
    }
  }
  return { success: true, vouchers: vouchers };
}

function useVoucher(id, code) {
  if (!id || !code) return { success: false, error: '잘못된 요청입니다.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이용권');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id) && String(data[i][2]) === String(code)) {
      if (data[i][5] === '사용신청') return { success: false, error: '이미 사용 신청 중입니다.' };
      if (data[i][5] === '사용완료') return { success: false, error: '이미 사용된 이용권입니다.' };
      sheet.getRange(i + 1, 6).setValue('사용신청');
      sheet.getRange(i + 1, 7).setValue(new Date());
      return { success: true, message: '사용 신청 완료! 선생님 승인을 기다려주세요.' };
    }
  }
  return { success: false, error: '이용권을 찾을 수 없습니다.' };
}

function getPendingVouchers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이용권');
  if (!sheet) return { success: true, vouchers: [] };
  var data = sheet.getDataRange().getValues();
  var vouchers = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] === '사용신청') {
      vouchers.push({ id: data[i][0], 구매일: data[i][1], 학번: data[i][2], 이름: data[i][3], 품목: data[i][4], 신청일: data[i][6] });
    }
  }
  return { success: true, vouchers: vouchers };
}

function approveVoucher(id) {
  if (!id) return { success: false, error: '잘못된 요청입니다.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이용권');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue('사용완료');
      return { success: true, message: data[i][3] + '의 ' + data[i][4] + ' 이용권을 승인했어요!' };
    }
  }
  return { success: false, error: '이용권을 찾을 수 없습니다.' };
}

function rejectVoucher(id) {
  if (!id) return { success: false, error: '잘못된 요청입니다.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('이용권');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue('거절');
      return { success: true, message: data[i][3] + '의 ' + data[i][4] + ' 이용권 신청을 거절했어요.' };
    }
  }
  return { success: false, error: '이용권을 찾을 수 없습니다.' };
}

function getPurchaseHistory(code) {
  if (!code) return { success: false, error: '잘못된 요청입니다.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('학생 명단');
  var logSheet = ss.getSheetByName('도토리 로그');
  var studentData = studentSheet.getDataRange().getValues();
  var studentName = '';
  for (var i = 1; i < studentData.length; i++) {
    if (String(studentData[i][2]) === String(code)) { studentName = studentData[i][1]; break; }
  }
  if (!studentName) return { success: false, error: '학생을 찾을 수 없습니다.' };
  var logData = logSheet.getDataRange().getValues();
  var history = [];
  for (var i = logData.length - 1; i >= 1; i--) {
    if (logData[i][1] === studentName && String(logData[i][3]).indexOf('상점 구매') === 0) {
      history.push({ 날짜: logData[i][0], 품목: String(logData[i][3]).replace('상점 구매: ', ''), 도토리: logData[i][2] });
    }
  }
  return { success: true, history: history, 이름: studentName };
}

// ===== 방명록 =====
function writeGuestbook(fromCode, toCode, content, authorType) {
  if (!toCode || !content) return { success: false, error: '잘못된 요청입니다.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('학생 명단');
  var gbSheet = ss.getSheetByName('방명록');
  var data = studentSheet.getDataRange().getValues();
  var fromName = authorType === '교사' ? '선생님' : '';
  var toName = '';
  for (var i = 1; i < data.length; i++) {
    if (authorType !== '교사' && String(data[i][2]) === String(fromCode)) fromName = data[i][1];
    if (String(data[i][2]) === String(toCode)) toName = data[i][1];
  }
  if (!fromName) return { success: false, error: '작성자를 찾을 수 없습니다.' };
  if (!toName)   return { success: false, error: '받는 학생을 찾을 수 없습니다.' };
  gbSheet.appendRow([new Date(), fromName, toCode, content, authorType || '학생']);
  return { success: true, message: toName + '의 방명록에 글을 남겼어요!' };
}

function getGuestbook(code) {
  if (!code) return { success: false, error: '잘못된 요청입니다.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('방명록');
  var data = sheet.getDataRange().getValues();
  var entries = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][2]) === String(code)) {
      entries.push({ 날짜: data[i][0], 작성자: data[i][1], 내용: data[i][3], 작성자유형: data[i][4] });
      if (entries.length >= 20) break;
    }
  }
  return { success: true, entries: entries };
}

// ===== 미션 =====
function getMissions(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('미션');
  if (!sheet) {
    sheet = ss.insertSheet('미션');
    sheet.appendRow(['ID','제목','내용','보상도토리','마감일','상태','생성일']);
  }
  var data = sheet.getDataRange().getValues();
  var showAll = params && params.all === '1';
  var missions = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (!showAll && data[i][5] !== '활성') continue;
    missions.push({
      id: data[i][0], 제목: data[i][1], 내용: data[i][2],
      보상도토리: Number(data[i][3]) || 0,
      마감일: data[i][4] ? Utilities.formatDate(new Date(data[i][4]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      상태: data[i][5]
    });
  }
  return { success: true, missions: missions };
}

function createMission(params) {
  var title = params.title || '';
  var content = params.content || '';
  var reward = parseInt(params.reward) || 0;
  var deadline = params.deadline || '';
  if (!title || !content || reward < 1) return { success: false, error: '필수 항목이 누락됐어요.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('미션');
  if (!sheet) {
    sheet = ss.insertSheet('미션');
    sheet.appendRow(['ID','제목','내용','보상도토리','마감일','상태','생성일']);
  }
  var id = new Date().getTime().toString();
  sheet.appendRow([id, title, content, reward, deadline, '활성', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')]);
  return { success: true, message: '미션이 만들어졌어요!' };
}

function toggleMission(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션');
  if (!sheet) return { success: false, error: '미션 시트가 없어요.' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id)) {
      var next = data[i][5] === '활성' ? '비활성' : '활성';
      sheet.getRange(i + 1, 6).setValue(next);
      return { success: true };
    }
  }
  return { success: false, error: '미션을 찾을 수 없어요.' };
}

function deleteMission(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션');
  if (!sheet) return { success: false, error: '미션 시트가 없어요.' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: '미션이 삭제됐어요.' };
    }
  }
  return { success: false, error: '미션을 찾을 수 없어요.' };
}

function submitMission(body) {
  var missionId = body.missionId;
  var code = body.code;
  var memo = body.memo || '';
  var fileData = body.fileData || '';
  var fileName = body.fileName || '';
  var fileType = body.fileType || '';
  if (!missionId || !code) return { success: false, error: '필수 항목이 누락됐어요.' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var missionSheet = ss.getSheetByName('미션');
  if (!missionSheet) return { success: false, error: '미션 시트가 없어요.' };
  var mData = missionSheet.getDataRange().getValues();
  var mission = null;
  for (var i = 1; i < mData.length; i++) {
    if (String(mData[i][0]) === String(missionId)) { mission = mData[i]; break; }
  }
  if (!mission) return { success: false, error: '미션을 찾을 수 없어요.' };
  if (mission[5] !== '활성') return { success: false, error: '비활성화된 미션이에요.' };
  var studentSheet = ss.getSheetByName('학생 명단');
  var sData = studentSheet.getDataRange().getValues();
  var studentName = code;
  for (var j = 1; j < sData.length; j++) {
    if (String(sData[j][2]) === String(code)) { studentName = sData[j][1]; break; }
  }
  var fileUrl = '';
  if (fileData && fileName) {
    try {
      var folders = DriveApp.getFoldersByName('미션제출파일');
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('미션제출파일');
      var base64 = fileData.replace(/^data:[^;]+;base64,/, '');
      var blob = Utilities.newBlob(Utilities.base64Decode(base64), fileType || 'application/octet-stream', fileName);
      var saved = folder.createFile(blob);
      saved.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = saved.getUrl();
    } catch(fe) { fileUrl = ''; }
  }
  var subSheet = ss.getSheetByName('미션제출');
  if (!subSheet) {
    subSheet = ss.insertSheet('미션제출');
    subSheet.appendRow(['ID','미션ID','미션제목','학생코드','학생이름','제출일','파일링크','메모','상태','처리일','보상도토리']);
  }
  subSheet.appendRow([
    new Date().getTime().toString(), missionId, mission[1], code, studentName,
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    fileUrl, memo, '대기중', '', Number(mission[3]) || 0
  ]);
  return { success: true, message: '미션 제출 완료! 선생님 확인을 기다려주세요.' };
}

function getMissionSubmissions(params) {
  var code = params.code;
  if (!code) return { success: false, error: '학생 코드가 필요해요.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션제출');
  if (!sheet) return { success: true, submissions: [] };
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][3]) === String(code)) {
      list.push({
        id: data[i][0], 미션제목: data[i][2], 제출일: data[i][5],
        파일링크: data[i][6], 메모: data[i][7], 상태: data[i][8],
        보상도토리: Number(data[i][10]) || 0
      });
    }
  }
  return { success: true, submissions: list };
}

function getPendingMissionSubmissions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션제출');
  if (!sheet) return { success: true, submissions: [] };
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === '대기중') {
      list.push({
        id: data[i][0], 미션ID: data[i][1], 미션제목: data[i][2],
        학생코드: data[i][3], 학생이름: data[i][4],
        파일링크: data[i][6], 메모: data[i][7], 보상도토리: Number(data[i][10]) || 0
      });
    }
  }
  return { success: true, submissions: list };
}

function approveMissionSubmission(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션제출');
  if (!sheet) return { success: false, error: '미션제출 시트가 없어요.' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id)) {
      if (data[i][8] !== '대기중') return { success: false, error: '이미 처리된 제출이에요.' };
      var reward = Number(data[i][10]) || 0;
      var studentCode = data[i][3];
      var studentName = data[i][4];
      var missionTitle = data[i][2];
      sheet.getRange(i + 1, 9).setValue('승인');
      sheet.getRange(i + 1, 10).setValue(Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'));
      giveAcorn(studentCode, reward, '미션: ' + missionTitle);
      return { success: true, message: studentName + '에게 도토리 ' + reward + '개를 지급했어요! 🌰' };
    }
  }
  return { success: false, error: '제출 내역을 찾을 수 없어요.' };
}

function rejectMissionSubmission(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션제출');
  if (!sheet) return { success: false, error: '미션제출 시트가 없어요.' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id)) {
      if (data[i][8] !== '대기중') return { success: false, error: '이미 처리된 제출이에요.' };
      sheet.getRange(i + 1, 9).setValue('거절');
      sheet.getRange(i + 1, 10).setValue(Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'));
      return { success: true, message: '거절 처리 완료.' };
    }
  }
  return { success: false, error: '제출 내역을 찾을 수 없어요.' };
}

function withdrawMissionSubmission(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션제출');
  if (!sheet) return { success: false, error: '제출 내역이 없어요.' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id) && String(data[i][3]) === String(params.code)) {
      if (data[i][8] !== '대기중') return { success: false, error: '이미 처리된 제출은 회수할 수 없어요.' };
      sheet.deleteRow(i + 1);
      return { success: true, message: '제출이 회수됐어요.' };
    }
  }
  return { success: false, error: '제출 내역을 찾을 수 없어요.' };
}

function updateMissionSubmission(body) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('미션제출');
  if (!sheet) return { success: false, error: '제출 내역이 없어요.' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.id) && String(data[i][3]) === String(body.code)) {
      if (data[i][8] !== '대기중') return { success: false, error: '이미 처리된 제출은 수정할 수 없어요.' };
      sheet.getRange(i + 1, 8).setValue(body.memo || '');
      if (body.fileData && body.fileName) {
        try {
          var folders = DriveApp.getFoldersByName('미션제출파일');
          var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('미션제출파일');
          var base64 = body.fileData.replace(/^data:[^;]+;base64,/, '');
          var blob = Utilities.newBlob(Utilities.base64Decode(base64), body.fileType || 'application/octet-stream', body.fileName);
          var saved = folder.createFile(blob);
          saved.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          sheet.getRange(i + 1, 7).setValue(saved.getUrl());
        } catch(fe) {}
      }
      return { success: true, message: '수정이 완료됐어요!' };
    }
  }
  return { success: false, error: '제출 내역을 찾을 수 없어요.' };
}
