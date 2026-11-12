/**
 * ==============================================================================
 * Project: 꿈빛 (Ggumbit) - 도서관 봉사인력 자원관리 시스템 (Google Apps Script)
 * Version: v1.13 (45.1단계: 출석부 6열 자동 복구, 동일인 1회 출석 제한, KST 당일 및 잠금 검증)
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
 * [0] 초기 권한 승인, 시스템 점검 및 밀린 출석부 행 즉시 복구 함수
 * Apps Script 편집기 상단에서 이 함수를 선택하고 [실행]을 누르면 권한 승인과 시트 정비가 완료됩니다.
 */
function testSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupStandardSheets_(ss, 2);
  SpreadsheetApp.flush();
  Logger.log('꿈빛 백엔드 v1.13 초기화 및 출석부 6열 정렬 점검 완료: ' + ss.getName());
} // [end of testSystem]

/**
 * [1] 웹앱 GET 요청 처리 (데이터 조회 및 한국 표준시 기준 오늘 날짜 제공)
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
        version: 'v1.13',
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
 * [2] 웹앱 POST 요청 처리 (데이터 저장, 출석체크, 대타, 월별 시트 생성)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (lockErr) {
    return createJsonResponse_({ status: 'error', message: 'Server busy, please retry.' });
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
 * [3] 표준 5대 탭 초기화 및 서식 구성
 */
function setupStandardSheets_(ss, slotCapacity) {
  const cap = Math.max(1, parseInt(slotCapacity, 10) || 2);

  // 1. 시스템 설정 탭 (B2~B5 고정 셀 규격)
  let cfgSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet(SHEET_NAMES.CONFIG);
  }
  if (cfgSheet.getLastRow() === 0) {
    const defaultRows = [
      ['설정 항목 (Key)', '설정 값 (Value)', '항목 설명'],
      ['libraryName', '꿈빛도서관', '도서관 공식 명칭 (B2 고정)'],
      ['adminName', '사서교사', '사서교사 / 관리자 성함 (B3 고정)'],
      ['adminId', 'admin', '관리자 로그인 아이디 (B4 고정)'],
      ['adminPassword', 'admin1234', '관리자 로그인 비밀번호 (B5 고정)'],
      ['cycleType', '2주', '순번 운영 주기 (1주 / 2주 / 1달)'],
      ['slotCapacity', String(cap), '슬롯당 봉사위원 정원 (명)'],
      ['operatingDays', '월,화,수,목,금', '정기 운영 요일'],
      ['autoHolidaysClose', 'true', '법정 공휴일 자동 휴관 여부'],
      ['cycleBaseDate', '2026-03-02', '순번 주차 기준 시작일자'],
      ['cycleBaseWeek', '1', '기준 시작일의 순번 주차']
    ];
    cfgSheet.getRange(1, 1, defaultRows.length, 3).setNumberFormat('@').setValues(defaultRows);
    styleHeaderRow_(cfgSheet, 3, '#fef3c7');
    cfgSheet.setColumnWidth(1, 180);
    cfgSheet.setColumnWidth(2, 240);
    cfgSheet.setColumnWidth(3, 280);
  }

  // 2. 봉사자 정보 탭
  let memSheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  if (!memSheet) {
    memSheet = ss.insertSheet(SHEET_NAMES.MEMBERS);
  }
  if (memSheet.getLastRow() === 0) {
    memSheet.getRange(1, 1, 1, 3).setValues([['순번', '이름', '연락처']]);
    styleHeaderRow_(memSheet, 3, '#e0f2fe');
  }

  // 3. 봉사자 순번 탭
  let rosSheet = ss.getSheetByName(SHEET_NAMES.ROSTER);
  if (!rosSheet) {
    rosSheet = ss.insertSheet(SHEET_NAMES.ROSTER);
  }
  if (rosSheet.getLastRow() === 0) {
    const headers = ['순번', '주차', '요일', '시간대'];
    for (let r = 1; r <= cap; r++) {
      headers.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
    }
    rosSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeaderRow_(rosSheet, headers.length, '#e0e7ff');
  }

  // 4. 출석부 탭 (6열 slotKey 영구 보존 및 밀린 행 자동 복구 규격)
  let attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (!attSheet) {
    attSheet = ss.insertSheet(SHEET_NAMES.ATTENDANCE);
  }
  ensureAttendanceSheetSchema_(attSheet);

  // 5. 대타 및 변경 탭
  let ovrSheet = ss.getSheetByName(SHEET_NAMES.OVERRIDES);
  if (!ovrSheet) {
    ovrSheet = ss.insertSheet(SHEET_NAMES.OVERRIDES);
  }
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

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();

  if (values.length >= 2 && values[1][1]) cfg.libraryName = String(values[1][1]).trim();
  if (values.length >= 3 && values[2][1]) cfg.adminName = String(values[2][1]).trim();
  if (values.length >= 4 && values[3][1]) cfg.adminId = String(values[3][1]).trim();
  if (values.length >= 5 && values[4][1]) cfg.adminPassword = String(values[4][1]).trim();

  for (let i = 1; i < values.length; i++) {
    const rawKey = String(values[i][0] || '').trim();
    const val = String(values[i][1] || '').trim();
    if (!rawKey || !val) continue;

    if (rawKey === 'libraryName' || rawKey === '도서관명' || rawKey === '도서관 이름') {
      cfg.libraryName = val;
    } else if (rawKey === 'adminName' || rawKey === '관리자성함' || rawKey === '사서교사') {
      cfg.adminName = val;
    } else if (rawKey === 'adminId' || rawKey === '관리자아이디' || rawKey === '아이디') {
      cfg.adminId = val;
    } else if (rawKey === 'adminPassword' || rawKey === '비밀번호' || rawKey === '패스워드') {
      cfg.adminPassword = val;
    } else if (rawKey === 'cycleType') {
      cfg.cycleType = val;
    } else if (rawKey === 'slotCapacity') {
      cfg.slotCapacity = parseInt(val, 10) || 2;
    } else if (rawKey === 'operatingDays') {
      cfg.operatingDays = val.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    } else if (rawKey === 'autoHolidaysClose') {
      cfg.autoHolidaysClose = (val.toLowerCase() === 'true');
    } else if (rawKey === 'cycleBaseDate') {
      cfg.cycleBaseDate = normalizeDateStr_(val);
    } else if (rawKey === 'cycleBaseWeek') {
      cfg.cycleBaseWeek = parseInt(val, 10) || 1;
    } else if (rawKey === 'timeSlots') {
      try { cfg.timeSlots = JSON.parse(val); } catch (e) {}
    }
  }

  return cfg;
} // [end of getSystemConfig_]

/**
 * [5] 시스템 기본 설정 저장 (도서관명 B2 고정 및 운영 파라미터 갱신)
 */
function saveSystemConfig_(ss, payload) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) {
    setupStandardSheets_(ss, payload.slotCapacity || 2);
    sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  }

  if (payload.libraryName) {
    upsertConfigKey_(sheet, 'libraryName', String(payload.libraryName).trim(), 2, '도서관 공식 명칭');
  }
  if (payload.cycleType) {
    upsertConfigKey_(sheet, 'cycleType', String(payload.cycleType).trim(), 6, '순번 운영 주기');
  }
  if (payload.slotCapacity !== undefined) {
    upsertConfigKey_(sheet, 'slotCapacity', String(payload.slotCapacity), 7, '슬롯당 봉사위원 정원');
  }
  if (Array.isArray(payload.operatingDays)) {
    upsertConfigKey_(sheet, 'operatingDays', payload.operatingDays.join(','), 8, '정기 운영 요일');
  }
  if (payload.autoHolidaysClose !== undefined) {
    upsertConfigKey_(sheet, 'autoHolidaysClose', String(payload.autoHolidaysClose), 9, '법정 공휴일 자동 휴관 여부');
  }
  if (payload.cycleBaseDate) {
    upsertConfigKey_(sheet, 'cycleBaseDate', normalizeDateStr_(payload.cycleBaseDate), 10, '순번 주차 기준 시작일자');
  }
  if (payload.cycleBaseWeek !== undefined) {
    upsertConfigKey_(sheet, 'cycleBaseWeek', String(payload.cycleBaseWeek), 11, '기준 시작일의 순번 주차');
  }
  if (Array.isArray(payload.timeSlots)) {
    upsertConfigKey_(sheet, 'timeSlots', JSON.stringify(payload.timeSlots), 12, '운영 시간대 설정(JSON)');
  }
} // [end of saveSystemConfig_]

/**
 * [6] 관리자 계정 정보 저장 (B3: 성함, B4: 아이디, B5: 비밀번호 직결)
 */
function saveAdminAccount_(ss, payload) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) {
    setupStandardSheets_(ss, 2);
    sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  }

  if (payload.adminName) {
    upsertConfigKey_(sheet, 'adminName', String(payload.adminName).trim(), 3, '사서교사 / 관리자 성함');
  }
  if (payload.adminId) {
    upsertConfigKey_(sheet, 'adminId', String(payload.adminId).trim(), 4, '관리자 로그인 아이디');
  }
  if (payload.adminPassword) {
    upsertConfigKey_(sheet, 'adminPassword', String(payload.adminPassword).trim(), 5, '관리자 로그인 비밀번호');
  }
} // [end of saveAdminAccount_]

function upsertConfigKey_(sheet, keyName, valueStr, preferredRow, descStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === keyName) {
        const targetRow = i + 2;
        sheet.getRange(targetRow, 2).setNumberFormat('@').setValue(String(valueStr));
        return;
      }
    }
  }

  const targetRow = (preferredRow && sheet.getRange(preferredRow, 1).getDisplayValue() === '')
    ? preferredRow
    : Math.max(lastRow + 1, preferredRow || 2);

  sheet.getRange(targetRow, 1, 1, 3)
    .setNumberFormat('@')
    .setValues([[keyName, String(valueStr), descStr || '']]);
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
    const phone = String(rows[i][2] || '').trim();
    if (name) {
      list.push({
        id: parseInt(rows[i][0], 10) || (list.length + 1),
        name: name,
        phone: phone
      });
    }
  }
  return list;
} // [end of getMembersList_]

function saveMembersList_(ss, members) {
  let sheet = ss.getSheetByName(SHEET_NAMES.MEMBERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.MEMBERS);
  }
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
    const week = weekMatch ? parseInt(weekMatch[0], 10) : 1;
    const day = String(r[2] || '').trim();
    const time = normalizeTimeStr_(r[3]);

    if (!day || !time) continue;

    const item = {
      id: parseInt(r[0], 10) || (list.length + 1),
      week: week,
      day: day,
      time: time
    };

    for (let c = 4; c < lastCol; c++) {
      const roleIdx = c - 3;
      const worker = String(r[c] || '').trim();
      item['name' + roleIdx] = worker;
      if (roleIdx === 1) item.name1 = worker;
      if (roleIdx === 2) item.name2 = worker;
    }

    list.push(item);
  }
  return list;
} // [end of getRosterList_]

function saveRosterList_(ss, roster, slotCapacity) {
  let sheet = ss.getSheetByName(SHEET_NAMES.ROSTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ROSTER);
  }
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
      const val = item['name' + r] || (r === 1 ? item.name1 : (r === 2 ? item.name2 : '')) || '';
      row.push(String(val).trim());
    }
    return row;
  });

  sheet.getRange(2, 1, rows.length, headers.length).setNumberFormat('@').setValues(rows);
} // [end of saveRosterList_]

/**
 * [9] 출석부 6열 스키마 보장, 밀린 행 자동 복구 및 중복 병합 엔진
 * 표준 열 순서: [1: 슬롯키(slotKey), 2: 활동일자, 3: 시간대, 4: 슬롯순번, 5: 봉사위원 성함, 6: 출석상태]
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

  // 1행 헤더가 다르더라도 기존 2행 이하 데이터를 추출해 6열 표준 구조로 정규화
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
    const row = rawValues[i];
    const c0 = String(row[0] || '').trim();
    const c1 = String(row[1] || '').trim();
    const c2 = String(row[2] || '').trim();
    const c3 = String(row[3] || '').trim();
    const c4 = String(row[4] || '').trim();
    const c5 = String(row[5] || '').trim();

    if (!c0 && !c1 && !c2 && !c3 && !c4 && !c5) {
      needsRewrite = true;
      continue;
    }

    let slotKey = '';
    let dateStr = '';
    let timeStr = '';
    let roleIdx = '1';
    let workerName = '';
    let status = 'O';

    // [밀린 4열 행 감지] c0가 날짜(예: 2026-11-12)이고 '_'가 없으며 c4·c5가 비어 있는 경우
    if (c0 && c0.indexOf('_') === -1 && /^\d{4}[-./]\d{1,2}[-./]\d{1,2}/.test(c0)) {
      needsRewrite = true;
      dateStr = normalizeDateStr_(c0);
      timeStr = normalizeTimeStr_(c1);
      workerName = c2;
      status = (c3.toUpperCase() === 'X') ? 'X' : 'O';
      roleIdx = '1';
      slotKey = dateStr + '_' + timeStr + '_' + roleIdx;
    } else {
      // 정상 6열 구조 파싱
      const parts = c0.split('_');
      dateStr = normalizeDateStr_(c1 || parts[0] || '');
      timeStr = normalizeTimeStr_(c2 || parts[1] || '');
      roleIdx = String(c3 || parts[2] || '1').trim() || '1';
      workerName = c4;
      status = (c5.toUpperCase() === 'X') ? 'X' : 'O';
      slotKey = (dateStr && timeStr) ? (dateStr + '_' + timeStr + '_' + roleIdx) : c0;

      if (c0 !== slotKey || c1 !== dateStr || c2 !== timeStr || c3 !== roleIdx || c5 !== status) {
        needsRewrite = true;
      }
    }

    if (!dateStr || !timeStr) {
      needsRewrite = true;
      continue;
    }

    // 동일인·동일시간대 중복 행 병합 키 (날짜_시간_성함 또는 슬롯키)
    const uniquePersonKey = workerName ? (dateStr + '_' + timeStr + '_' + workerName) : slotKey;
    const cleanRow = [slotKey, dateStr, timeStr, roleIdx, workerName, status];

    if (normalizedMap[uniquePersonKey]) {
      needsRewrite = true;
      // 이미 'O'가 기록되어 있으면 'O' 우선 보존
      if (normalizedMap[uniquePersonKey][5] !== 'O' && status === 'O') {
        normalizedMap[uniquePersonKey] = cleanRow;
      } else if (normalizedMap[uniquePersonKey][5] === status) {
        normalizedMap[uniquePersonKey] = cleanRow;
      }
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
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ATTENDANCE);
  }
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
    if (!roleIdx && parts.length >= 3) {
      roleIdx = parts[2];
    }
    const normDate = normalizeDateStr_(parts[0] || dateStr);
    const normTime = normalizeTimeStr_(parts[1] || timeStr);
    const normRole = roleIdx || '1';
    slotKey = normDate + '_' + normTime + '_' + normRole;
  } else {
    roleIdx = roleIdx || '1';
    slotKey = dateStr + '_' + timeStr + '_' + roleIdx;
  }

  const finalDate = dateStr || slotKey.split('_')[0] || '';
  const finalTime = timeStr || slotKey.split('_')[1] || '';
  const finalRole = roleIdx || slotKey.split('_')[2] || '1';

  // 일반 사용자(isAdmin === false)인 경우 한국 표준시(KST) 기준 '오늘'만 출석 허용
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
      const exKey = String(existing[i][0] || '').trim();
      const exDate = normalizeDateStr_(existing[i][1]);
      const exTime = normalizeTimeStr_(existing[i][2]);
      const exName = String(existing[i][4] || '').trim();
      const exStatus = String(existing[i][5] || '').trim().toUpperCase();

      const samePersonSameTime = (workerName && exDate === finalDate && exTime === finalTime && exName === workerName);
      const sameSlot = (exKey === slotKey && (!workerName || !exName || exName === workerName));

      if (samePersonSameTime || sameSlot) {
        matchedRows.push(i + 2);
        if (exStatus === 'O') {
          alreadyConfirmedO = true;
        }
      }
    }
  }

  // 일반 사용자는 이미 'O'로 확정된 출석을 'X'로 취소 불가
  if (!isAdminReq && payload.isAdmin === false && alreadyConfirmedO && status === 'X') {
    return { status: 'locked', message: '이미 출석 확인이 완료되었습니다.' };
  }

  if (matchedRows.length > 0) {
    const primaryRow = matchedRows[0];
    sheet.getRange(primaryRow, 1, 1, 6).setNumberFormat('@').setValues([rowPayload]);

    for (let k = matchedRows.length - 1; k >= 1; k--) {
      sheet.deleteRow(matchedRows[k]);
    }
  } else {
    const nextRow = lastRow + 1;
    sheet.getRange(nextRow, 1, 1, 6).setNumberFormat('@').setValues([rowPayload]);
  }

  return {
    status: 'success',
    message: status === 'O' ? '출석 상태 영구 저장 완료' : '출석 상태 변경 완료',
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
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    const rawKey = String(rows[i][0] || '').trim();
    const dateStr = normalizeDateStr_(rows[i][1]);
    const timeStr = normalizeTimeStr_(rows[i][2]);
    const roleIdx = String(rows[i][3] || '1').trim();
    const name = String(rows[i][4] || '').trim();
    const status = (String(rows[i][5] || 'O').trim().toUpperCase() === 'X') ? 'X' : 'O';

    const slotKey = rawKey || (dateStr + '_' + timeStr + '_' + roleIdx);
    const record = {
      slotKey: slotKey,
      date: dateStr,
      time: timeStr,
      roleIdx: roleIdx,
      name: name,
      status: status
    };

    if (slotKey) {
      if (!map[slotKey] || status === 'O') {
        map[slotKey] = record;
      }
    }

    if (dateStr && timeStr && name) {
      const personKey = dateStr + '_' + timeStr + '_' + name;
      if (!map[personKey] || status === 'O') {
        map[personKey] = record;
      }
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
    const slotKey = String(rows[i][0] || '').trim();
    const status = String(rows[i][1] || '').trim();
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

  const status = String(data.status || '').trim();
  const lastRow = sheet.getLastRow();
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
    if (foundRow > 0) {
      sheet.deleteRow(foundRow);
    }
    return;
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const rowValues = [
    slotKey,
    status,
    String(data.requester || '').trim(),
    String(data.substitute || '').trim(),
    String(data.reason || '').trim(),
    nowStr
  ];

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, 6).setNumberFormat('@').setValues([rowValues]);
  } else {
    sheet.getRange(lastRow + 1, 1, 1, 6).setNumberFormat('@').setValues([rowValues]);
  }
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
  let gridRows = Array.isArray(payload.gridRows) ? payload.gridRows : [];

  if (gridRows.length === 0 && scheduleRows.length > 0) {
    gridRows = buildGridRowsFromSchedule_(year, month, scheduleRows, cap);
  }

  const listTabName = ymStr + ' 봉사배정표';
  let listSheet = ss.getSheetByName(listTabName);
  if (listSheet) {
    listSheet.clear();
  } else {
    listSheet = ss.insertSheet(listTabName);
  }

  const listHeaders = ['일자', '요일', '주차', '시간대'];
  for (let r = 1; r <= cap; r++) {
    listHeaders.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
  }
  listHeaders.push('비고');

  listSheet.getRange(1, 1, 1, listHeaders.length).setValues([listHeaders]);
  styleHeaderRow_(listSheet, listHeaders.length, '#fef3c7');

  if (scheduleRows.length > 0) {
    const normalizedList = scheduleRows.map(function(r) {
      const row = r.slice(0, listHeaders.length);
      while (row.length < listHeaders.length) row.push('');
      return row.map(function(v) { return String(v === null || v === undefined ? '' : v); });
    });
    listSheet.getRange(2, 1, normalizedList.length, listHeaders.length)
      .setNumberFormat('@')
      .setValues(normalizedList);
  }

  const calTabName = ymStr + ' 달력';
  let calSheet = ss.getSheetByName(calTabName);
  if (calSheet) {
    calSheet.clear();
  } else {
    calSheet = ss.insertSheet(calTabName);
  }

  const calHeaders = ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'];
  calSheet.getRange(1, 1, 1, 7).setValues([calHeaders]);
  styleHeaderRow_(calSheet, 7, '#e0e7ff');

  if (gridRows.length > 0) {
    const normalizedGrid = gridRows.map(function(weekArr) {
      const row = weekArr.slice(0, 7);
      while (row.length < 7) row.push('');
      return row.map(function(v) { return String(v === null || v === undefined ? '' : v); });
    });
    const gridRange = calSheet.getRange(2, 1, normalizedGrid.length, 7);
    gridRange.setNumberFormat('@').setValues(normalizedGrid);
    gridRange.setWrap(true).setVerticalAlignment('top');
    for (let rIdx = 2; rIdx <= normalizedGrid.length + 1; rIdx++) {
      calSheet.setRowHeight(rIdx, 95);
    }
    for (let cIdx = 1; cIdx <= 7; cIdx++) {
      calSheet.setColumnWidth(cIdx, 145);
    }
  }
} // [end of createMonthlySheets_]

function buildGridRowsFromSchedule_(year, month, scheduleRows, cap) {
  const byDate = {};
  for (let i = 0; i < scheduleRows.length; i++) {
    const r = scheduleRows[i];
    const dStr = normalizeDateStr_(r[0]);
    if (!dStr) continue;
    if (!byDate[dStr]) byDate[dStr] = [];
    byDate[dStr].push(r);
  }

  const lastDate = new Date(year, month, 0).getDate();
  const gridRows = [];
  let weekRow = ['', '', '', '', '', '', ''];

  for (let d = 1; d <= lastDate; d++) {
    const cur = new Date(year, month - 1, d);
    const dow = cur.getDay();
    const dStr = year + '-' + String(month).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    const items = byDate[dStr] || [];

    if (items.length === 0) {
      weekRow[dow] = d + '일';
    } else if (items[0][2] === '-') {
      weekRow[dow] = d + '일\n(' + (items[0][4 + cap] || '휴관') + ')';
    } else {
      let cellText = d + '일 [' + items[0][2] + ']\n';
      for (let j = 0; j < items.length; j++) {
        const it = items[j];
        const tShort = String(it[3] || '').slice(0, 5);
        const workers = [];
        for (let c = 4; c < 4 + cap; c++) {
          const w = String(it[c] || '').trim();
          if (w && w !== '공석' && w !== '-') workers.push(w);
        }
        if (workers.length > 0) {
          cellText += tShort + ' ' + workers.join('/') + '\n';
        }
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
 * [12] 신학기 데이터 초기화 (헤더 유지, 데이터 행만 안전 초기화)
 */
function resetNewSemesterData_(ss, slotCapacity) {
  setupStandardSheets_(ss, slotCapacity);
  const targets = [SHEET_NAMES.MEMBERS, SHEET_NAMES.ROSTER, SHEET_NAMES.ATTENDANCE, SHEET_NAMES.OVERRIDES];
  for (let i = 0; i < targets.length; i++) {
    const sh = ss.getSheetByName(targets[i]);
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
  }
} // [end of resetNewSemesterData_]

/**
 * [13] 공통 유틸리티 (날짜/시간 정규화, 헤더 스타일, JSON 응답)
 */
function normalizeDateStr_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  const m = s.match(/(\d{4})[-./\s]+(\d{1,2})[-./\s]+(\d{1,2})/);
  if (m) {
    return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');
  }
  return s;
} // [end of normalizeDateStr_]

function normalizeTimeStr_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'HH:mm');
  }
  const s = String(val).trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    return String(m[1]).padStart(2, '0') + ':' + m[2];
  }
  return s;
} // [end of normalizeTimeStr_]

function styleHeaderRow_(sheet, colCount, bgColor) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground(bgColor || '#f1f5f9')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');
} // [end of styleHeaderRow_]

function createJsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
} // [end of createJsonResponse_]
