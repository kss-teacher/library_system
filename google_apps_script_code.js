/**
 * ==============================================================================
 * Project: 꿈빛 (Ggumbit) - 도서관 봉사인력 자원관리 시스템 (Google Apps Script)
 * Version: v1.19 (출석부 6열 영구 보존, KST 당일 출석 제한, 긴급 구인 SOS 선착순 엔진 통합)
 * Author / Copyright: (c) 2026 Sung soo Kim. All Rights Reserved.
 * ==============================================================================
 */

const SHEET_NAMES = {
  CONFIG: '시스템 설정',
  MEMBERS: '봉사자 정보',
  ROSTER: '봉사자 순번',
  ATTENDANCE: '출석부',
  OVERRIDES: '대타 및 변경'
};

/**
 * [0] 초기 권한 승인 및 시스템 점검
 */
function testSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupStandardSheets_(ss, 2);
  SpreadsheetApp.flush();
  Logger.log('꿈빛 백엔드 v1.19 초기화 및 시스템 점검 완료: ' + ss.getName());
} // [end of testSystem]

/**
 * [1] 웹앱 GET 요청 처리 (데이터 조회)
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getSystemData';

    if (action === 'getSystemData') {
      const config = getSystemConfig_(ss);
      const members = getMembersList_(ss);
      const roster = getRosterList_(ss);
      const overrides = getOverridesMap_(ss);
      const attendance = getAttendanceMap_(ss);
      const serverTodayKst = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

      const result = {
        status: 'success',
        version: 'v1.19',
        serverTodayKst: serverTodayKst,
        spreadsheetUrl: ss.getUrl(),
        spreadsheetId: ss.getId(),
        libraryName: config.libraryName || '꿈빛도서관',
        adminName: config.adminName || '사서교사',
        adminId: config.adminId || 'admin',
        adminPassword: config.adminPassword || 'admin1234',
        cycleType: config.cycleType || '2주',
        slotCapacity: config.slotCapacity || 2,
        operatingDays: config.operatingDays || ['월', '화', '수', '목', '금'],
        autoHolidaysClose: config.autoHolidaysClose !== undefined ? config.autoHolidaysClose : true,
        cycleBaseDate: config.cycleBaseDate || '2026-03-02',
        cycleBaseWeek: config.cycleBaseWeek || 1,
        timeSlots: config.timeSlots || null,
        members: members,
        roster: roster,
        overrides: overrides,
        attendance: attendance
      };

      return createJsonResponse_(result);
    }

    return createJsonResponse_({ status: 'error', message: 'Unknown GET action: ' + action });
  } catch (err) {
    return createJsonResponse_({ status: 'error', message: err.toString() });
  }
} // [end of doGet]

/**
 * [2] 웹앱 POST 요청 처리 (LockService 기반 트랜잭션 동시성 제어)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (lockErr) {
    return createJsonResponse_({ status: 'error', message: '서버가 혼잡합니다. 잠시 후 다시 시도해주세요.' });
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const payload = JSON.parse(rawBody);
    const action = payload.action || '';

    if (action === 'setupStandardSheets') {
      const cap = parseInt(payload.slotCapacity, 10) || 2;
      setupStandardSheets_(ss, cap);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '표준 5대 탭 구성 완료' });
    }

    if (action === 'saveSystemConfig') {
      saveSystemConfig_(ss, payload);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '시스템 기본 설정 저장 완료' });
    }

    if (action === 'saveAdminAccount') {
      saveAdminAccount_(ss, payload);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '관리자 계정 정보 저장 완료' });
    }

    if (action === 'saveMembers') {
      saveMembersList_(ss, payload.members || []);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '봉사위원 명부 저장 완료' });
    }

    if (action === 'saveRoster') {
      const cap = parseInt(payload.slotCapacity, 10) || 2;
      saveRosterList_(ss, payload.roster || [], cap);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '봉사자 순번 저장 완료' });
    }

    if (action === 'saveExcelBatch') {
      const cap = parseInt(payload.slotCapacity, 10) || 2;
      saveMembersList_(ss, payload.members || []);
      saveRosterList_(ss, payload.roster || [], cap);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '엑셀 일괄 데이터 반영 완료' });
    }

    if (action === 'saveAttendanceCheck') {
      const saveRes = saveAttendanceCheck_(ss, payload);
      SpreadsheetApp.flush();
      return createJsonResponse_(saveRes);
    }

    if (action === 'setOverride') {
      setOverrideRecord_(ss, payload.slotKey, payload.data || {});
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '대타 및 변경 이력 반영 완료' });
    }

    if (action === 'requestSOS') {
      const res = requestSOS_(ss, payload);
      SpreadsheetApp.flush();
      return createJsonResponse_(res);
    }

    if (action === 'applySOS') {
      const res = applySOS_(ss, payload);
      SpreadsheetApp.flush();
      return createJsonResponse_(res);
    }

    if (action === 'cancelSOS') {
      setOverrideRecord_(ss, payload.slotKey, { status: 'RESET' });
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '구인 요청이 취소되었습니다.' });
    }

    if (action === 'createMonthlySheet') {
      createMonthlySheets_(ss, payload);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '월별 배정표 및 달력 탭 생성 완료' });
    }

    if (action === 'resetNewSemester') {
      const cap = parseInt(payload.slotCapacity, 10) || 2;
      resetNewSemesterData_(ss, cap);
      SpreadsheetApp.flush();
      return createJsonResponse_({ status: 'success', message: '신학기 데이터 초기화 완료' });
    }

    return createJsonResponse_({ status: 'error', message: 'Unknown POST action: ' + action });
  } catch (err) {
    return createJsonResponse_({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
} // [end of doPost]

/**
 * [2-1] 긴급 구인(SOS) 코어 엔진
 */
function requestSOS_(ss, payload) {
  const slotKey = String(payload.slotKey || '').trim();
  const origWorker = String(payload.origWorker || '공석').trim();
  if (!slotKey) return { status: 'error', message: '슬롯키가 없습니다.' };

  setOverrideRecord_(ss, slotKey, {
    status: 'SOS_OPEN',
    requester: origWorker,
    substitute: '',
    reason: '긴급 구인'
  });
  return { status: 'success', message: '긴급 구인 알림이 발송되었습니다.' };
} // [end of requestSOS_]

function applySOS_(ss, payload) {
  const slotKey = String(payload.slotKey || '').trim();
  const applicantName = String(payload.applicantName || '').trim();
  if (!slotKey || !applicantName) return { status: 'error', message: '잘못된 요청입니다.' };

  let sheet = ss.getSheetByName(SHEET_NAMES.OVERRIDES);
  if (!sheet) return { status: 'error', message: '시스템 오류' };

  const lastRow = sheet.getLastRow();
  let foundRow = -1;
  let currentStatus = '';
  let requester = '';

  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === slotKey) {
        foundRow = i + 2;
        currentStatus = String(keys[i][1]).trim();
        requester = String(keys[i][2]).trim();
        break;
      }
    }
  }

  if (foundRow > 0 && currentStatus === 'SOS_OPEN') {
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(foundRow, 2, 1, 5).setNumberFormat('@').setValues([
      ['ACCEPTED', requester, applicantName, '긴급 구인 지원', nowStr]
    ]);
    return { status: 'success', message: '지원 감사합니다. 일정표를 확인해주세요.' };
  } else {
    return { status: 'error', message: '마감되었습니다. 지원 감사합니다.' };
  }
} // [end of applySOS_]

/**
 * [3] 표준 5대 탭 초기화 및 서식 구성
 */
function setupStandardSheets_(ss, slotCapacity) {
  const cap = Math.max(1, parseInt(slotCapacity, 10) || 2);
  let cfgSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!cfgSheet) cfgSheet = ss.insertSheet(SHEET_NAMES.CONFIG);
  if (cfgSheet.getLastRow() === 0) {
    const defaultRows = [
      ['설정 항목 (Key)', '설정 값 (Value)', '항목 설명'],
      ['libraryName', '꿈빛도서관', '도서관 공식 명칭 (B2 고정)'],
      ['adminName', '사서교사', '사서교사 / 관리자 성함 (B3 고정)'],
      ['adminId', 'admin', '관리자 로그인 아이디 (B4 고정)'],
      ['adminPassword', 'admin1234', '관리자 로그인 비밀번호 (B5 고정)'],
      ['cycleType', '2주', '순번 운영 주기'],
      ['slotCapacity', String(cap), '슬롯당 정원'],
      ['operatingDays', '월,화,수,목,금', '정기 운영 요일'],
      ['autoHolidaysClose', 'true', '공휴일 자동 휴관'],
      ['cycleBaseDate', '2026-03-02', '기준일'],
      ['cycleBaseWeek', '1', '기준 주차']
    ];
    cfgSheet.getRange(1, 1, defaultRows.length, 3).setNumberFormat('@').setValues(defaultRows);
    styleHeaderRow_(cfgSheet, 3, '#fef3c7');
    cfgSheet.setColumnWidth(1, 180);
    cfgSheet.setColumnWidth(2, 240);
    cfgSheet.setColumnWidth(3, 280);
  }

  let memSheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  if (!memSheet) memSheet = ss.insertSheet(SHEET_NAMES.MEMBERS);
  if (memSheet.getLastRow() === 0) {
    memSheet.getRange(1, 1, 1, 3).setValues([['순번', '이름', '연락처']]);
    styleHeaderRow_(memSheet, 3, '#e0f2fe');
  }

  let rosSheet = ss.getSheetByName(SHEET_NAMES.ROSTER);
  if (!rosSheet) rosSheet = ss.insertSheet(SHEET_NAMES.ROSTER);
  if (rosSheet.getLastRow() === 0) {
    const headers = ['순번', '주차', '요일', '시간대'];
    for (let r = 1; r <= cap; r++) {
      headers.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
    }
    rosSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeaderRow_(rosSheet, headers.length, '#e0e7ff');
  }

  let attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (!attSheet) attSheet = ss.insertSheet(SHEET_NAMES.ATTENDANCE);
  ensureAttendanceSheetSchema_(attSheet);

  let ovrSheet = ss.getSheetByName(SHEET_NAMES.OVERRIDES);
  if (!ovrSheet) ovrSheet = ss.insertSheet(SHEET_NAMES.OVERRIDES);
  if (ovrSheet.getLastRow() === 0) {
    ovrSheet.getRange(1, 1, 1, 6).setValues([['슬롯키(slotKey)', '상태', '원래배정자', '대타참여자', '사유', '기록일시']]);
    styleHeaderRow_(ovrSheet, 6, '#ffe4e6');
  }
} // [end of setupStandardSheets_]

/**
 * [4] 시스템 설정 읽기 (B2~B5 직접 셀 및 키 기반 병행 조회)
 */
function getSystemConfig_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  const cfg = {
    libraryName: '꿈빛도서관',
    adminName: '사서교사',
    adminId: 'admin',
    adminPassword: 'admin1234',
    cycleType: '2주',
    slotCapacity: 2,
    operatingDays: ['월', '화', '수', '목', '금'],
    autoHolidaysClose: true,
    cycleBaseDate: '2026-03-02',
    cycleBaseWeek: 1
  };
  if (!sheet || sheet.getLastRow() < 2) return cfg;
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 2).getDisplayValues();
  if (values.length >= 2 && values[1][1]) cfg.libraryName = String(values[1][1]).trim();
  if (values.length >= 3 && values[2][1]) cfg.adminName = String(values[2][1]).trim();
  if (values.length >= 4 && values[3][1]) cfg.adminId = String(values[3][1]).trim();
  if (values.length >= 5 && values[4][1]) cfg.adminPassword = String(values[4][1]).trim();

  for (let i = 1; i < values.length; i++) {
    const rawKey = String(values[i][0] || '').trim();
    const val = String(values[i][1] || '').trim();
    if (!rawKey || !val) continue;
    if (rawKey === 'libraryName' || rawKey === '도서관명') cfg.libraryName = val;
    else if (rawKey === 'adminName' || rawKey === '사서교사') cfg.adminName = val;
    else if (rawKey === 'adminId' || rawKey === '아이디') cfg.adminId = val;
    else if (rawKey === 'adminPassword' || rawKey === '비밀번호') cfg.adminPassword = val;
    else if (rawKey === 'cycleType') cfg.cycleType = val;
    else if (rawKey === 'slotCapacity') cfg.slotCapacity = parseInt(val, 10) || 2;
    else if (rawKey === 'operatingDays') cfg.operatingDays = val.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    else if (rawKey === 'autoHolidaysClose') cfg.autoHolidaysClose = (val.toLowerCase() === 'true');
    else if (rawKey === 'cycleBaseDate') cfg.cycleBaseDate = normalizeDateStr_(val);
    else if (rawKey === 'cycleBaseWeek') cfg.cycleBaseWeek = parseInt(val, 10) || 1;
    else if (rawKey === 'timeSlots') {
      try { cfg.timeSlots = JSON.parse(val); } catch (e) {}
    }
  }
  return cfg;
} // [end of getSystemConfig_]

/**
 * [5] 시스템 기본 설정 저장
 */
function saveSystemConfig_(ss, payload) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) {
    setupStandardSheets_(ss, payload.slotCapacity || 2);
    sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  }
  if (payload.libraryName) upsertConfigKey_(sheet, 'libraryName', String(payload.libraryName).trim(), 2, '도서관 공식 명칭');
  if (payload.cycleType) upsertConfigKey_(sheet, 'cycleType', String(payload.cycleType).trim(), 6, '순번 운영 주기');
  if (payload.slotCapacity !== undefined) upsertConfigKey_(sheet, 'slotCapacity', String(payload.slotCapacity), 7, '슬롯당 정원');
  if (Array.isArray(payload.operatingDays)) upsertConfigKey_(sheet, 'operatingDays', payload.operatingDays.join(','), 8, '정기 운영 요일');
  if (payload.autoHolidaysClose !== undefined) upsertConfigKey_(sheet, 'autoHolidaysClose', String(payload.autoHolidaysClose), 9, '공휴일 자동 휴관');
  if (payload.cycleBaseDate) upsertConfigKey_(sheet, 'cycleBaseDate', normalizeDateStr_(payload.cycleBaseDate), 10, '기준 시작일자');
  if (payload.cycleBaseWeek !== undefined) upsertConfigKey_(sheet, 'cycleBaseWeek', String(payload.cycleBaseWeek), 11, '기준 순번 주차');
  if (Array.isArray(payload.timeSlots)) upsertConfigKey_(sheet, 'timeSlots', JSON.stringify(payload.timeSlots), 12, '운영 시간대(JSON)');
} // [end of saveSystemConfig_]

/**
 * [6] 관리자 계정 정보 저장
 */
function saveAdminAccount_(ss, payload) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) {
    setupStandardSheets_(ss, 2);
    sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  }
  if (payload.adminName) upsertConfigKey_(sheet, 'adminName', String(payload.adminName).trim(), 3, '사서교사 / 관리자 성함');
  if (payload.adminId) upsertConfigKey_(sheet, 'adminId', String(payload.adminId).trim(), 4, '관리자 아이디');
  if (payload.adminPassword) upsertConfigKey_(sheet, 'adminPassword', String(payload.adminPassword).trim(), 5, '관리자 비밀번호');
} // [end of saveAdminAccount_]

function upsertConfigKey_(sheet, keyName, valueStr, preferredRow, descStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === keyName) {
        sheet.getRange(i + 2, 2).setNumberFormat('@').setValue(String(valueStr));
        return;
      }
    }
  }
  const targetRow = (preferredRow && sheet.getRange(preferredRow, 1).getDisplayValue() === '')
    ? preferredRow
    : Math.max(lastRow + 1, preferredRow || 2);
  sheet.getRange(targetRow, 1, 1, 3).setNumberFormat('@').setValues([[keyName, String(valueStr), descStr || '']]);
} // [end of upsertConfigKey_]

/**
 * [7] 봉사자 정보(명부) 읽기 및 쓰기
 */
function getMembersList_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
  const list = [];
  for (let i = 0; i < rows.length; i++) {
    const name = String(rows[i][1] || '').trim();
    if (name) {
      list.push({
        id: parseInt(rows[i][0], 10) || (list.length + 1),
        name: name,
        phone: String(rows[i][2] || '').trim()
      });
    }
  }
  return list;
} // [end of getMembersList_]

function saveMembersList_(ss, members) {
  let sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.MEMBERS);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['순번', '이름', '연락처']]);
  styleHeaderRow_(sheet, 3, '#e0f2fe');
  if (!members || members.length === 0) return;
  const rows = members.map(function(m, idx) {
    return [String(idx + 1), String(m.name || '').trim(), String(m.phone || '').trim()];
  });
  sheet.getRange(2, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
} // [end of saveMembersList_]

/**
 * [8] 봉사자 순번 읽기 및 쓰기
 */
function getRosterList_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ROSTER);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(6, sheet.getLastColumn());
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getDisplayValues();
  const list = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const weekMatch = String(r[1] || '').match(/\d+/);
    const day = String(r[2] || '').trim();
    const time = normalizeTimeStr_(r[3]);
    if (!day || !time) continue;
    const item = {
      id: parseInt(r[0], 10) || (list.length + 1),
      week: weekMatch ? parseInt(weekMatch[0], 10) : 1,
      day: day,
      time: time
    };
    for (let c = 4; c < lastCol; c++) {
      const roleIdx = c - 3;
      item['name' + roleIdx] = String(r[c] || '').trim();
      if (roleIdx === 1) item.name1 = item['name' + roleIdx];
      if (roleIdx === 2) item.name2 = item['name' + roleIdx];
    }
    list.push(item);
  }
  return list;
} // [end of getRosterList_]

function saveRosterList_(ss, roster, slotCapacity) {
  let sheet = ss.getSheetByName(SHEET_NAMES.ROSTER);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.ROSTER);
  const cap = Math.max(1, parseInt(slotCapacity, 10) || 2);
  sheet.clearContents();
  const headers = ['순번', '주차', '요일', '시간대'];
  for (let r = 1; r <= cap; r++) {
    headers.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeaderRow_(sheet, headers.length, '#e0e7ff');
  if (!roster || roster.length === 0) return;
  const rows = roster.map(function(item, idx) {
    const row = [
      String(idx + 1),
      String(item.week || 1) + '주차',
      String(item.day || '').trim(),
      normalizeTimeStr_(item.time)
    ];
    for (let r = 1; r <= cap; r++) {
      row.push(String(item['name' + r] || (r === 1 ? item.name1 : (r === 2 ? item.name2 : '')) || '').trim());
    }
    return row;
  });
  sheet.getRange(2, 1, rows.length, headers.length).setNumberFormat('@').setValues(rows);
} // [end of saveRosterList_]

/**
 * [9] 출석부 6열 스키마 보장, 밀린 행 자동 복구 및 중복 병합 엔진
 */
function ensureAttendanceSheetSchema_(sheet) {
  const expectedHeaders = ['슬롯키(slotKey)', '활동일자', '시간대', '슬롯순번', '봉사위원 성함', '출석상태'];
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, 6).setValues([expectedHeaders]);
    styleHeaderRow_(sheet, 6, '#dcfce7');
    formatAttendanceColumns_(sheet);
    return;
  }
  const colAHeader = String(sheet.getRange(1, 1).getDisplayValue() || '').trim();
  const hasValidHeader = (colAHeader === '슬롯키(slotKey)');
  if (lastRow === 1) {
    if (!hasValidHeader) {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, 6).setValues([expectedHeaders]);
      styleHeaderRow_(sheet, 6, '#dcfce7');
      formatAttendanceColumns_(sheet);
    }
    return;
  }
  const maxCols = Math.max(6, sheet.getLastColumn());
  const rawValues = sheet.getRange(2, 1, lastRow - 1, maxCols).getDisplayValues();
  const normalizedMap = {};
  const orderedKeys = [];
  let needsRewrite = !hasValidHeader;

  for (let i = 0; i < rawValues.length; i++) {
    const c0 = String(rawValues[i][0] || '').trim(),
          c1 = String(rawValues[i][1] || '').trim(),
          c2 = String(rawValues[i][2] || '').trim(),
          c3 = String(rawValues[i][3] || '').trim(),
          c4 = String(rawValues[i][4] || '').trim(),
          c5 = String(rawValues[i][5] || '').trim();
    if (!c0 && !c1 && !c2 && !c3 && !c4 && !c5) {
      needsRewrite = true;
      continue;
    }

    let slotKey = '', dateStr = '', timeStr = '', roleIdx = '1', workerName = '', status = 'O';
    if (c0 && c0.indexOf('_') === -1 && /^\d{4}[-./]\d{1,2}[-./]\d{1,2}/.test(c0)) {
      needsRewrite = true;
      dateStr = normalizeDateStr_(c0);
      timeStr = normalizeTimeStr_(c1);
      workerName = c2;
      status = (c3.toUpperCase() === 'X') ? 'X' : 'O';
      roleIdx = '1';
      slotKey = dateStr + '_' + timeStr + '_' + roleIdx;
    } else {
      const parts = c0.split('_');
      dateStr = normalizeDateStr_(c1 || parts[0] || '');
      timeStr = normalizeTimeStr_(c2 || parts[1] || '');
      roleIdx = String(c3 || parts[2] || '1').trim() || '1';
      workerName = c4;
      status = (c5.toUpperCase() === 'X') ? 'X' : 'O';
      slotKey = (dateStr && timeStr) ? (dateStr + '_' + timeStr + '_' + roleIdx) : c0;
      if (c0 !== slotKey || c1 !== dateStr || c2 !== timeStr || c3 !== roleIdx || c5 !== status) needsRewrite = true;
    }
    if (!dateStr || !timeStr) {
      needsRewrite = true;
      continue;
    }
    const uniquePersonKey = workerName ? (dateStr + '_' + timeStr + '_' + workerName) : slotKey;
    const cleanRow = [slotKey, dateStr, timeStr, roleIdx, workerName, status];

    if (normalizedMap[uniquePersonKey]) {
      needsRewrite = true;
      if (normalizedMap[uniquePersonKey][5] !== 'O' && status === 'O') normalizedMap[uniquePersonKey] = cleanRow;
      else if (normalizedMap[uniquePersonKey][5] === status) normalizedMap[uniquePersonKey] = cleanRow;
    } else {
      normalizedMap[uniquePersonKey] = cleanRow;
      orderedKeys.push(uniquePersonKey);
    }
  }
  if (needsRewrite) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 6).setValues([expectedHeaders]);
    styleHeaderRow_(sheet, 6, '#dcfce7');
    formatAttendanceColumns_(sheet);
    const finalRows = orderedKeys.map(function(k) { return normalizedMap[k]; });
    if (finalRows.length > 0) {
      sheet.getRange(2, 1, finalRows.length, 6).setNumberFormat('@').setValues(finalRows);
    }
  }
} // [end of ensureAttendanceSheetSchema_]

function formatAttendanceColumns_(sheet) {
  sheet.setColumnWidth(1, 195);
  sheet.setColumnWidth(2, 115);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 85);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 95);
} // [end of formatAttendanceColumns_]

function saveAttendanceCheck_(ss, payload) {
  let sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAMES.ATTENDANCE);
  ensureAttendanceSheetSchema_(sheet);

  const dateStr = normalizeDateStr_(payload.date || '');
  const timeStr = normalizeTimeStr_(payload.time || '');
  const workerName = String(payload.name || '').trim();
  const status = (String(payload.status || 'O').trim().toUpperCase() === 'X') ? 'X' : 'O';
  const isAdminReq = (payload.isAdmin === true);

  let slotKey = String(payload.slotKey || '').trim();
  let roleIdx = String(payload.roleIdx || '').trim();
  if (slotKey) {
    const parts = slotKey.split('_');
    if (!roleIdx && parts.length >= 3) roleIdx = parts[2];
    slotKey = normalizeDateStr_(parts[0] || dateStr) + '_' + normalizeTimeStr_(parts[1] || timeStr) + '_' + (roleIdx || '1');
  } else {
    slotKey = dateStr + '_' + timeStr + '_' + (roleIdx || '1');
  }

  const finalDate = dateStr || slotKey.split('_')[0] || '';
  const finalTime = timeStr || slotKey.split('_')[1] || '';
  const finalRole = roleIdx || slotKey.split('_')[2] || '1';

  if (payload.isAdmin === false) {
    const todayKst = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    if (finalDate && finalDate !== todayKst) {
      return { status: 'error', message: '오늘 봉사만 출석이 됩니다.' };
    }
  }

  const rowPayload = [slotKey, finalDate, finalTime, finalRole, workerName, status];
  const lastRow = sheet.getLastRow();
  const matchedRows = [];
  let alreadyConfirmedO = false;

  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();
    for (let i = 0; i < existing.length; i++) {
      const exKey = String(existing[i][0] || '').trim(),
            exDate = normalizeDateStr_(existing[i][1]),
            exTime = normalizeTimeStr_(existing[i][2]),
            exName = String(existing[i][4] || '').trim(),
            exStatus = String(existing[i][5] || '').trim().toUpperCase();

      const samePersonSameTime = (workerName && exDate === finalDate && exTime === finalTime && exName === workerName);
      const sameSlot = (exKey === slotKey && (!workerName || !exName || exName === workerName));
      if (samePersonSameTime || sameSlot) {
        matchedRows.push(i + 2);
        if (exStatus === 'O') alreadyConfirmedO = true;
      }
    }
  }

  if (!isAdminReq && payload.isAdmin === false && alreadyConfirmedO && status === 'X') {
    return { status: 'locked', message: '이미 출석 확인이 완료되었습니다.' };
  }

  if (matchedRows.length > 0) {
    sheet.getRange(matchedRows[0], 1, 1, 6).setNumberFormat('@').setValues([rowPayload]);
    for (let k = matchedRows.length - 1; k >= 1; k--) sheet.deleteRow(matchedRows[k]);
  } else {
    sheet.getRange(lastRow + 1, 1, 1, 6).setNumberFormat('@').setValues([rowPayload]);
  }
  return {
    status: 'success',
    message: status === 'O' ? '출석 상태 저장 완료' : '출석 상태 변경 완료',
    record: {
      slotKey: slotKey,
      date: finalDate,
      time: finalTime,
      roleIdx: finalRole,
      name: workerName,
      status: status
    }
  };
} // [end of saveAttendanceCheck_]

function getAttendanceMap_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  ensureAttendanceSheetSchema_(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    const slotKey = rows[i][0] || (normalizeDateStr_(rows[i][1]) + '_' + normalizeTimeStr_(rows[i][2]) + '_' + String(rows[i][3] || '1').trim());
    const record = {
      slotKey: slotKey,
      date: normalizeDateStr_(rows[i][1]),
      time: normalizeTimeStr_(rows[i][2]),
      roleIdx: String(rows[i][3] || '1').trim(),
      name: String(rows[i][4] || '').trim(),
      status: (String(rows[i][5] || 'O').trim().toUpperCase() === 'X') ? 'X' : 'O'
    };
    if (slotKey && (!map[slotKey] || record.status === 'O')) map[slotKey] = record;
    if (record.date && record.time && record.name && (!map[record.date + '_' + record.time + '_' + record.name] || record.status === 'O')) {
      map[record.date + '_' + record.time + '_' + record.name] = record;
    }
  }
  return map;
} // [end of getAttendanceMap_]

/**
 * [10] 대타 및 변경 이력 읽기/쓰기
 */
function getOverridesMap_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.OVERRIDES);
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    const slotKey = String(rows[i][0] || '').trim(),
          status = String(rows[i][1] || '').trim();
    if (!slotKey || !status || status === 'CANCELLED' || status === 'RESET') continue;
    map[slotKey] = {
      status: status,
      requester: String(rows[i][2] || '').trim(),
      substitute: String(rows[i][3] || '').trim(),
      reason: String(rows[i][4] || '').trim()
    };
  }
  return map;
} // [end of getOverridesMap_]

function setOverrideRecord_(ss, slotKey, data) {
  if (!slotKey) return;
  let sheet = ss.getSheetByName(SHEET_NAMES.OVERRIDES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.OVERRIDES);
    sheet.getRange(1, 1, 1, 6).setValues([['슬롯키(slotKey)', '상태', '원래배정자', '대타참여자', '사유', '기록일시']]);
    styleHeaderRow_(sheet, 6, '#ffe4e6');
  }
  const status = String(data.status || '').trim(),
        lastRow = sheet.getLastRow();
  let foundRow = -1;
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === slotKey) {
        foundRow = i + 2;
        break;
      }
    }
  }
  if (status === 'CANCELLED' || status === 'RESET') {
    if (foundRow > 0) sheet.deleteRow(foundRow);
    return;
  }
  const rowValues = [
    slotKey,
    status,
    String(data.requester || '').trim(),
    String(data.substitute || '').trim(),
    String(data.reason || '').trim(),
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
  ];
  sheet.getRange(foundRow > 0 ? foundRow : lastRow + 1, 1, 1, 6).setNumberFormat('@').setValues([rowValues]);
} // [end of setOverrideRecord_]

/**
 * [11] 월별 봉사배정표 및 7열 달력 탭 자동 생성
 */
function createMonthlySheets_(ss, payload) {
  const year = parseInt(payload.year, 10) || new Date().getFullYear();
  const month = parseInt(payload.month, 10) || (new Date().getMonth() + 1);
  const cap = Math.max(1, parseInt(payload.slotCapacity, 10) || 2);
  const ymStr = year + '-' + String(month).padStart(2, '0');
  const scheduleRows = Array.isArray(payload.scheduleRows) ? payload.scheduleRows : [];
  let gridRows = Array.isArray(payload.gridRows) ? payload.gridRows : (scheduleRows.length > 0 ? buildGridRowsFromSchedule_(year, month, scheduleRows, cap) : []);

  let listSheet = ss.getSheetByName(ymStr + ' 봉사배정표');
  if (listSheet) listSheet.clear();
  else listSheet = ss.insertSheet(ymStr + ' 봉사배정표');
  const listHeaders = ['일자', '요일', '주차', '시간대'];
  for (let r = 1; r <= cap; r++) {
    listHeaders.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
  }
  listHeaders.push('비고');
  listSheet.getRange(1, 1, 1, listHeaders.length).setValues([listHeaders]);
  styleHeaderRow_(listSheet, listHeaders.length, '#fef3c7');
  if (scheduleRows.length > 0) {
    const normList = scheduleRows.map(function(r) {
      const row = r.slice(0, listHeaders.length);
      while (row.length < listHeaders.length) row.push('');
      return row.map(function(v) { return String(v || ''); });
    });
    listSheet.getRange(2, 1, normList.length, listHeaders.length).setNumberFormat('@').setValues(normList);
  }

  let calSheet = ss.getSheetByName(ymStr + ' 달력');
  if (calSheet) calSheet.clear();
  else calSheet = ss.insertSheet(ymStr + ' 달력');
  calSheet.getRange(1, 1, 1, 7).setValues([['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일']]);
  styleHeaderRow_(calSheet, 7, '#e0e7ff');
  if (gridRows.length > 0) {
    const normGrid = gridRows.map(function(w) {
      const row = w.slice(0, 7);
      while (row.length < 7) row.push('');
      return row.map(function(v) { return String(v || ''); });
    });
    const gridRange = calSheet.getRange(2, 1, normGrid.length, 7);
    gridRange.setNumberFormat('@').setValues(normGrid);
    gridRange.setWrap(true).setVerticalAlignment('top');
    for (let rIdx = 2; rIdx <= normGrid.length + 1; rIdx++) calSheet.setRowHeight(rIdx, 95);
    for (let cIdx = 1; cIdx <= 7; cIdx++) calSheet.setColumnWidth(cIdx, 145);
  }
} // [end of createMonthlySheets_]

function buildGridRowsFromSchedule_(year, month, scheduleRows, cap) {
  const byDate = {};
  for (let i = 0; i < scheduleRows.length; i++) {
    const dStr = normalizeDateStr_(scheduleRows[i][0]);
    if (dStr) {
      if (!byDate[dStr]) byDate[dStr] = [];
      byDate[dStr].push(scheduleRows[i]);
    }
  }
  const gridRows = [], lastDate = new Date(year, month, 0).getDate();
  let weekRow = ['', '', '', '', '', '', ''];
  for (let d = 1; d <= lastDate; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    const dStr = year + '-' + String(month).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    const items = byDate[dStr] || [];
    if (items.length === 0) weekRow[dow] = d + '일';
    else if (items[0][2] === '-') weekRow[dow] = d + '일\n(' + (items[0][4 + cap] || '휴관') + ')';
    else {
      let cellText = d + '일 [' + items[0][2] + ']\n';
      for (let j = 0; j < items.length; j++) {
        const workers = [];
        for (let c = 4; c < 4 + cap; c++) {
          const w = String(items[j][c] || '').trim();
          if (w && w !== '공석' && w !== '-') workers.push(w);
        }
        if (workers.length > 0) cellText += String(items[j][3] || '').slice(0, 5) + ' ' + workers.join('/') + '\n';
      }
      weekRow[dow] = cellText.trim();
    }
    if (dow === 6 || d === lastDate) {
      gridRows.push(weekRow);
      weekRow = ['', '', '', '', '', '', ''];
    }
  }
  return gridRows;
} // [end of buildGridRowsFromSchedule_]

/**
 * [12] 신학기 데이터 초기화
 */
function resetNewSemesterData_(ss, slotCapacity) {
  setupStandardSheets_(ss, slotCapacity);
  [SHEET_NAMES.MEMBERS, SHEET_NAMES.ROSTER, SHEET_NAMES.ATTENDANCE, SHEET_NAMES.OVERRIDES].forEach(function(t) {
    const sh = ss.getSheetByName(t);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
} // [end of resetNewSemesterData_]

/**
 * [13] 공통 유틸리티
 */
function normalizeDateStr_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const m = String(val).trim().match(/(\d{4})[-./\s]+(\d{1,2})[-./\s]+(\d{1,2})/);
  return m ? m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0') : String(val).trim();
} // [end of normalizeDateStr_]

function normalizeTimeStr_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'HH:mm');
  }
  const m = String(val).trim().match(/(\d{1,2}):(\d{2})/);
  return m ? String(m[1]).padStart(2, '0') + ':' + m[2] : String(val).trim();
} // [end of normalizeTimeStr_]

function styleHeaderRow_(sheet, colCount, bgColor) {
  sheet.getRange(1, 1, 1, colCount)
    .setBackground(bgColor || '#f1f5f9')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
} // [end of styleHeaderRow_]

function createJsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
} // [end of createJsonResponse_]
