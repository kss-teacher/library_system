/**
 * ==============================================================================
 * [프로그램 정보]
 * Project: 꿈빛 (Ggumbit) - 도서관 봉사인력 자원관리 시스템
 * Version: Google Apps Script 클라우드 백엔드 엔진 v1.03
 * Author / Copyright: (c) 2026 Sung soo Kim. All Rights Reserved.
 * License: Open Education License (Based on MIT License)
 * ==============================================================================
 */

// -------------------------------------------------------------
// [1] 스프레드시트 기동 시 전용 관리 메뉴 등록
// -------------------------------------------------------------
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 꿈빛 도서관 관리')
    .addItem('🚀 표준 5대 탭 자동 구성하기', 'setupStandardSheetsMenu')
    .addItem('📅 이번 달 배정표 & 달력 탭 생성', 'createCurrentMonthSheetMenu')
    .addSeparator()
    .addItem('🌸 신학기 봉사단 시트 초기화', 'resetNewSemesterMenu')
    .addItem('🔧 시스템 권한 및 연결 테스트', 'testSystem')
    .addToUi();
}

function setupStandardSheetsMenu() {
  setupStandardSheets(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('🎉 표준 5대 탭 구성이 완료되었습니다.');
}

function createCurrentMonthSheetMenu() {
  const now = new Date();
  createMonthlySheet(now.getFullYear(), now.getMonth() + 1, 2);
  SpreadsheetApp.getUi().alert('🎉 이번 달 배정표 및 달력 탭 생성이 완료되었습니다.');
}

function resetNewSemesterMenu() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('신학기 초기화', '정말 초기화하시겠습니까? "초기화"를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK && res.getResponseText().trim() === '초기화') {
    resetNewSemesterSheets(SpreadsheetApp.getActiveSpreadsheet());
    ui.alert('초기화가 완료되었습니다.');
  }
}

function testSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('연결 성공: ' + ss.getName() + ' (' + ss.getUrl() + ')');
  return 'SUCCESS: ' + ss.getUrl();
}

// -------------------------------------------------------------
// [2] 웹앱 통신 GET 요청 처리기 (doGet)
// -------------------------------------------------------------
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getSystemConfig(ss.getSheetByName('시스템 설정'));

    // 1. 봉사자 정보
    const members = [];
    const memberSheet = ss.getSheetByName('봉사자 정보');
    if (memberSheet) {
      const mData = memberSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < mData.length; i++) {
        if (!mData[i][1]) continue;
        members.push({
          id: i,
          name: String(mData[i][1]).trim(),
          phone: String(mData[i][2] || '').trim()
        });
      }
    }

    // 2. 봉사자 순번 (N인 가변 열 파서 연동)
    let roster = [];
    let detectedCap = config.slotCapacity || 2;
    const rosterSheet = ss.getSheetByName('봉사자 순번');
    if (rosterSheet) {
      const res = findRosterColumns(rosterSheet);
      roster = res.roster || [];
      if (res.slotCapacity && res.slotCapacity > detectedCap) {
        detectedCap = res.slotCapacity;
      }
    }

    // 3. 출석 및 변경 기록 (대타 / 직권교체)
    const overrides = {};
    const logSheet = ss.getSheetByName('출석 및 변경 기록');
    if (logSheet) {
      const lData = logSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < lData.length; i++) {
        const k = String(lData[i][1] || '').trim();
        const stat = String(lData[i][2] || '').trim();
        const req = String(lData[i][3] || '').trim();
        const sub = String(lData[i][4] || '').trim();
        const reas = String(lData[i][5] || '').trim();
        if (k) {
          overrides[k] = { status: stat, requester: req, substitute: sub, reason: reas };
        }
      }
    }

    // 4. 출석부
    const attendance = {};
    const attSheet = ss.getSheetByName('출석부');
    if (attSheet) {
      const aData = attSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < aData.length; i++) {
        const d = String(aData[i][0] || '').trim();
        const t = normalizeTimeStr(aData[i][1]);
        const n = String(aData[i][2] || '').trim();
        const s = String(aData[i][3] || '').trim();
        if (d && n) {
          attendance[`${d}_${t}_${n}`] = { date: d, time: t, name: n, status: s };
        }
      }
    }

    const payload = {
      status: 'success',
      spreadsheetUrl: ss.getUrl(),
      libraryName: config.libraryName,
      adminName: config.adminName,
      adminId: config.adminId,
      adminPassword: config.adminPassword,
      adminEmail: config.adminEmail,
      cycleType: config.cycleType,
      slotCapacity: detectedCap,
      operatingDays: config.operatingDays,
      autoHolidaysClose: config.autoHolidaysClose,
      timeSlots: config.timeSlots,
      noticeMemo: config.defaultNoticeMemo,
      cycleBaseDate: config.cycleBaseDate,
      cycleBaseWeek: config.cycleBaseWeek,
      members: members,
      roster: roster,
      overrides: overrides,
      attendance: attendance
    };

    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------
// [3] 웹앱 통신 POST 요청 처리기 (doPost)
// -------------------------------------------------------------
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const postData = JSON.parse(e.postData.contents || '{}');
    const action = postData.action;

    // 1. 5대 표준 탭 자동 구성
    if (action === 'setupStandardSheets') {
      setupStandardSheets(ss);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 시스템 기본 파라미터 저장
    if (action === 'saveSystemConfig' || action === 'saveSettings') {
      const sheet = ss.getSheetByName('시스템 설정') || ss.insertSheet('시스템 설정');
      const data = sheet.getDataRange().getDisplayValues();
      const updateKey = (key, val) => {
        let found = false;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]).trim() === key) {
            sheet.getRange(i + 1, 2).setValue(val);
            found = true;
            break;
          }
        }
        if (!found) sheet.appendRow([key, val]);
      };

      if (postData.libraryName) updateKey('libraryName', postData.libraryName);
      if (postData.cycleType) updateKey('cycleType', postData.cycleType);
      if (postData.cycleBaseDate) updateKey('cycleBaseDate', postData.cycleBaseDate);
      if (postData.cycleBaseWeek !== undefined) updateKey('cycleBaseWeek', String(postData.cycleBaseWeek));
      if (postData.slotCapacity !== undefined) updateKey('slotCapacity', String(postData.slotCapacity));
      if (postData.operatingDays) updateKey('operatingDays', JSON.stringify(postData.operatingDays));
      if (postData.autoHolidaysClose !== undefined) updateKey('autoHolidaysClose', String(postData.autoHolidaysClose));
      if (postData.timeSlots) updateKey('timeSlots', JSON.stringify(postData.timeSlots));

      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. 봉사위원 명부 저장
    if (action === 'saveMembers') {
      const sheet = ss.getSheetByName('봉사자 정보') || ss.insertSheet('봉사자 정보');
      const members = postData.members || [];
      sheet.clearContents();
      sheet.appendRow(['순번', '이름', '연락처']);
      const rows = members.map((m, idx) => [idx + 1, m.name || '', m.phone || '']);
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 3).setValues(rows);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. 봉사자 순번표 저장 (N인 가변 열 완비)
    if (action === 'saveRoster') {
      const sheet = ss.getSheetByName('봉사자 순번') || ss.insertSheet('봉사자 순번');
      const roster = postData.roster || [];
      const cap = Math.max(1, parseInt(postData.slotCapacity, 10) || 2);

      sheet.clearContents();
      const header = ['순번', '주차', '요일', '시간대'];
      for (let r = 1; r <= cap; r++) {
        header.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
      }
      sheet.appendRow(header);

      const rows = roster.map((r, idx) => {
        const row = [idx + 1, r.week + '주차', r.day, normalizeTimeStr(r.time)];
        for (let i = 1; i <= cap; i++) {
          row.push(r['name' + i] || (i === 1 ? r.name1 : (i === 2 ? r.name2 : '')) || '');
        }
        return row;
      });

      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 5. 대타 및 직권교체 기록 저장
    if (action === 'setOverride') {
      const sheet = ss.getSheetByName('출석 및 변경 기록') || ss.insertSheet('출석 및 변경 기록');
      const k = postData.slotKey;
      const d = postData.data || {};
      const timeStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['기록일시', '슬롯키', '상태', '요청자', '대타/교체자', '사유']);
      }

      sheet.appendRow([timeStr, k, d.status || '', d.requester || '', d.substitute || '', d.reason || '']);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 6. 출석 체크(O/X) 저장
    if (action === 'saveAttendanceCheck') {
      const sheet = ss.getSheetByName('출석부') || ss.insertSheet('출석부');
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['날짜', '시간', '이름', '출석확인']);
      }
      sheet.appendRow([postData.date, normalizeTimeStr(postData.time), postData.name, postData.status]);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 7. 월별 배정표 및 달력 듀얼 탭 원격 생성 (세부 데이터 완전 적재)
    if (action === 'createMonthlySheet') {
      const year = parseInt(postData.year, 10);
      const month = parseInt(postData.month, 10);
      const cap = Math.max(1, parseInt(postData.slotCapacity, 10) || 2);
      const scheduleRows = postData.scheduleRows || [];
      const gridRows = postData.gridRows || [];
      createMonthlySheet(year, month, cap, scheduleRows, gridRows);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 8. 신학기 시트 탭 안전 초기화
    if (action === 'resetNewSemester') {
      resetNewSemesterSheets(ss);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'unknown_action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------
// [4] 데이터 파싱 및 헬퍼 엔진
// -------------------------------------------------------------
function normalizeTimeStr(t) {
  if (!t) return '';
  const s = String(t).trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    return String(m[1]).padStart(2, '0') + ':' + m[2];
  }
  return s;
}

function findRosterColumns(sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { roster: [], slotCapacity: 2 };

  const header = data[0] || [];
  let maxCap = 2;
  const workerCols = [];

  for (let c = 4; c < header.length; c++) {
    const h = String(header[c] || '').trim();
    const m = h.match(/봉사위원\s*(\d+)/) || h.match(/봉사자\s*(\d+)/);
    const rIdx = m ? parseInt(m[1], 10) : (c - 3);
    workerCols.push({ colIdx: c, roleIdx: rIdx });
    if (rIdx > maxCap) maxCap = rIdx;
  }

  if (workerCols.length === 0) {
    workerCols.push({ colIdx: 4, roleIdx: 1 });
    workerCols.push({ colIdx: 5, roleIdx: 2 });
  }

  const roster = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1] && !row[2]) continue;

    let weekVal = 1;
    const mWeek = String(row[1] || '').match(/\d+/);
    if (mWeek) weekVal = parseInt(mWeek[0], 10);

    const item = {
      id: i,
      week: weekVal,
      day: String(row[2] || '').trim(),
      time: normalizeTimeStr(row[3])
    };

    workerCols.forEach(wc => {
      const val = String(row[wc.colIdx] || '').trim();
      item['name' + wc.roleIdx] = (val === '공석') ? '' : val;
      if (wc.roleIdx === 1) item.name1 = item['name' + wc.roleIdx];
      if (wc.roleIdx === 2) item.name2 = item['name' + wc.roleIdx];
    });

    roster.push(item);
  }
  return { roster, slotCapacity: maxCap };
}

function getSystemConfig(sheet) {
  const config = {
    libraryName: '꿈빛도서관',
    adminName: '사서교사',
    adminId: 'admin',
    adminPassword: 'admin1234',
    adminEmail: '',
    cycleType: '2주',
    slotCapacity: 2,
    operatingDays: ['월', '화', '수', '목', '금'],
    autoHolidaysClose: true,
    timeSlots: [
      { id: 't_1', name: '오전 1차', startTime: '08:30', endTime: '10:30' },
      { id: 't_2', name: '오전 2차', startTime: '10:30', endTime: '12:30' }
    ],
    defaultNoticeMemo: '꿈빛도서관 봉사단에 오신 것을 환영합니다.',
    cycleBaseDate: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'),
    cycleBaseWeek: 1
  };
  if (!sheet) return config;

  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    const v = String(data[i][1] || '').trim();
    if (k === 'libraryName' && v) config.libraryName = v;
    if (k === 'adminName' && v) config.adminName = v;
    if (k === 'adminId' && v) config.adminId = v;
    if (k === 'adminPassword' && v) config.adminPassword = v;
    if (k === 'adminEmail' && v) config.adminEmail = v;
    if (k === 'cycleType' && v) config.cycleType = v;
    if (k === 'slotCapacity' && v) config.slotCapacity = parseInt(v, 10) || 2;
    if (k === 'operatingDays' && v) {
      try { config.operatingDays = JSON.parse(v); } catch (e) {}
    }
    if (k === 'autoHolidaysClose' && v) config.autoHolidaysClose = (v === 'true' || v === true);
    if (k === 'cycleBaseDate' && v) config.cycleBaseDate = v;
    if (k === 'cycleBaseWeek' && v) config.cycleBaseWeek = parseInt(v, 10) || 1;
    if (k === 'timeSlots' && v) {
      try { config.timeSlots = JSON.parse(v); } catch (e) {}
    }
  }
  return config;
}

// -------------------------------------------------------------
// [5] 표준 5대 탭 자동 구성 및 신학기 초기화
// -------------------------------------------------------------
function setupStandardSheets(ss) {
  // 1. 시스템 설정
  let s1 = ss.getSheetByName('시스템 설정') || ss.insertSheet('시스템 설정');
  if (s1.getLastRow() === 0) {
    s1.appendRow(['항목키', '설정값']);
    s1.appendRow(['libraryName', '꿈빛도서관']);
    s1.appendRow(['adminName', '사서교사']);
    s1.appendRow(['adminId', 'admin']);
    s1.appendRow(['adminPassword', 'admin1234']);
    s1.appendRow(['cycleType', '2주']);
    s1.appendRow(['slotCapacity', '2']);
    s1.appendRow(['operatingDays', '["월","화","수","목","금"]']);
    s1.appendRow(['autoHolidaysClose', 'true']);
    s1.appendRow(['cycleBaseDate', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd')]);
    s1.appendRow(['cycleBaseWeek', '1']);
  }

  // 2. 봉사자 정보
  let s2 = ss.getSheetByName('봉사자 정보') || ss.insertSheet('봉사자 정보');
  if (s2.getLastRow() === 0) {
    s2.appendRow(['순번', '이름', '연락처']);
  }

  // 3. 봉사자 순번
  let s3 = ss.getSheetByName('봉사자 순번') || ss.insertSheet('봉사자 순번');
  if (s3.getLastRow() === 0) {
    s3.appendRow(['순번', '주차', '요일', '시간대', '봉사위원 1 (정)', '봉사위원 2 (부)']);
  }

  // 4. 출석부
  let s4 = ss.getSheetByName('출석부') || ss.insertSheet('출석부');
  if (s4.getLastRow() === 0) {
    s4.appendRow(['날짜', '시간', '이름', '출석확인']);
  }

  // 5. 출석 및 변경 기록
  let s5 = ss.getSheetByName('출석 및 변경 기록') || ss.insertSheet('출석 및 변경 기록');
  if (s5.getLastRow() === 0) {
    s5.appendRow(['기록일시', '슬롯키', '상태', '요청자', '대타/교체자', '사유']);
  }
}

function resetNewSemesterSheets(ss) {
  const clearTab = (name, headers) => {
    const s = ss.getSheetByName(name);
    if (s) {
      s.clearContents();
      if (headers) s.appendRow(headers);
    }
  };
  clearTab('봉사자 정보', ['순번', '이름', '연락처']);
  clearTab('봉사자 순번', ['순번', '주차', '요일', '시간대', '봉사위원 1 (정)', '봉사위원 2 (부)']);
  clearTab('출석부', ['날짜', '시간', '이름', '출석확인']);
  clearTab('출석 및 변경 기록', ['기록일시', '슬롯키', '상태', '요청자', '대타/교체자', '사유']);
}

// -------------------------------------------------------------
// [6] 월별 배정표 및 7열 격자 달력 탭 동시 생성 엔진
// -------------------------------------------------------------
function createMonthlySheet(year, month, slotCapacity, scheduleRows, gridRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cap = slotCapacity || 2;
  const sheetName = year + '-' + String(month).padStart(2, '0') + ' 봉사배정표';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  const header = ['활동 날짜', '요일', '순번 주차', '시간대'];
  for (let r = 1; r <= cap; r++) {
    header.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
  }
  header.push('비고 (대타 및 교체 내역)');

  sheet.appendRow(header);
  sheet.getRange(1, 1, 1, header.length)
    .setBackground('#059669')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 세부 배정 데이터 적재
  if (scheduleRows && scheduleRows.length > 0) {
    const dataRange = sheet.getRange(2, 1, scheduleRows.length, header.length);
    dataRange.setValues(scheduleRows);
    dataRange.setVerticalAlignment('middle');
    sheet.getRange(2, 1, scheduleRows.length, 3).setHorizontalAlignment('center');
  }

  // 열 너비 자동 조정
  for (let c = 1; c <= header.length; c++) {
    sheet.autoResizeColumn(c);
  }

  // 달력형 그리드 탭 동시 생성
  createMonthlyCalendarGridSheet(ss, year, month, cap, gridRows);
}

function createMonthlyCalendarGridSheet(ss, year, month, slotCapacity, gridRows) {
  const gridSheetName = year + '-' + String(month).padStart(2, '0') + ' 달력';
  let gSheet = ss.getSheetByName(gridSheetName);
  if (gSheet) ss.deleteSheet(gSheet);
  gSheet = ss.insertSheet(gridSheetName);

  const headers = ['일', '월', '화', '수', '목', '금', '토'];
  gSheet.appendRow(headers);
  gSheet.getRange(1, 1, 1, 7)
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 일요일 빨강, 토요일 파랑 헤더 강조
  gSheet.getRange(1, 1).setBackground('#e11d48');
  gSheet.getRange(1, 7).setBackground('#2563eb');

  if (gridRows && gridRows.length > 0) {
    const range = gSheet.getRange(2, 1, gridRows.length, 7);
    range.setValues(gridRows);
    range.setVerticalAlignment('top');
    range.setWrap(true);
    for (let r = 2; r <= gridRows.length + 1; r++) {
      gSheet.setRowHeight(r, 85);
    }
  }

  for (let c = 1; c <= 7; c++) {
    gSheet.setColumnWidth(c, 135);
  }
}
