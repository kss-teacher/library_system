/**
 * ==============================================================================
 * [프로그램 정보]
 * Project: 꿈빛 (Ggumbit) - 도서관 봉사인력 자원관리 시스템
 * Version: Google Apps Script 클라우드 백엔드 엔진 v1.10 (B2:도서관, B3:성함, B4:아이디, B5:비밀번호 다이렉트 고정셀 & flush() 완비)
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
} // [end of onOpen]

function setupStandardSheetsMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSystemConfig(ss.getSheetByName('시스템 설정'));
  setupStandardSheets(ss, config.slotCapacity || 2);
  SpreadsheetApp.getUi().alert('🎉 표준 5대 탭 구성이 완료되었습니다.');
} // [end of setupStandardSheetsMenu]

function createCurrentMonthSheetMenu() {
  const now = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSystemConfig(ss.getSheetByName('시스템 설정'));
  createMonthlySheet(now.getFullYear(), now.getMonth() + 1, config.slotCapacity || 2);
  SpreadsheetApp.getUi().alert('🎉 이번 달 배정표 및 달력 탭 생성이 완료되었습니다.');
} // [end of createCurrentMonthSheetMenu]

function resetNewSemesterMenu() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('신학기 초기화', '정말 초기화하시겠습니까? "초기화"를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK && res.getResponseText().trim() === '초기화') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getSystemConfig(ss.getSheetByName('시스템 설정'));
    resetNewSemesterSheets(ss, config.slotCapacity || 2);
    ui.alert('초기화가 완료되었습니다.');
  }
} // [end of resetNewSemesterMenu]

function testSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('연결 성공: ' + ss.getName() + ' (' + ss.getUrl() + ')');
  return 'SUCCESS: ' + ss.getUrl();
} // [end of testSystem]

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
      spreadsheetId: ss.getId(), // [v1.07] 원터치 엑셀 백업 다운로드 직결용 공식 고유 ID
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
} // [end of doGet]

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
      const cap = Math.max(1, parseInt(postData.slotCapacity, 10) || 2);
      setupStandardSheets(ss, cap);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 시스템 설정 및 관리자 계정 저장 (B2~B5 고정셀 정렬 & flush 장착)
    if (action === 'saveSystemConfig' || action === 'saveSettings' || action === 'saveAdminAccount') {
      const sheet = ss.getSheetByName('시스템 설정') || ss.insertSheet('시스템 설정');
      
      // 기존 데이터 로드
      let data = sheet.getDataRange().getDisplayValues();
      if (data.length === 0 || !data[0][0]) {
        sheet.appendRow(['항목키', '설정값']);
        data = [['항목키', '설정값']];
      }

      // 키-값 맵 구성
      const configMap = {};
      for (let i = 1; i < data.length; i++) {
        const k = String(data[i][0] || '').trim();
        const v = String(data[i][1] || '').trim();
        if (k) configMap[k] = v;
      }

      // 4대 핵심 계정 및 도서관 정보 업데이트
      if (postData.libraryName !== undefined && postData.libraryName !== null && String(postData.libraryName).trim() !== '') {
        configMap['libraryName'] = String(postData.libraryName).trim();
      }
      if (postData.adminName !== undefined && postData.adminName !== null && String(postData.adminName).trim() !== '') {
        configMap['adminName'] = String(postData.adminName).trim();
      }
      if (postData.adminId !== undefined && postData.adminId !== null && String(postData.adminId).trim() !== '') {
        configMap['adminId'] = String(postData.adminId).trim();
      }
      if (postData.adminPassword !== undefined && postData.adminPassword !== null && String(postData.adminPassword).trim() !== '') {
        configMap['adminPassword'] = String(postData.adminPassword).trim();
      }

      if (postData.cycleType) configMap['cycleType'] = postData.cycleType;
      if (postData.cycleBaseDate) configMap['cycleBaseDate'] = postData.cycleBaseDate;
      if (postData.cycleBaseWeek !== undefined) configMap['cycleBaseWeek'] = String(postData.cycleBaseWeek);
      if (postData.slotCapacity !== undefined) configMap['slotCapacity'] = String(postData.slotCapacity);
      if (postData.operatingDays) configMap['operatingDays'] = JSON.stringify(postData.operatingDays);
      if (postData.autoHolidaysClose !== undefined) configMap['autoHolidaysClose'] = String(postData.autoHolidaysClose);
      if (postData.timeSlots) configMap['timeSlots'] = JSON.stringify(postData.timeSlots);

      // 기본값 Fallback 보장
      if (!configMap['libraryName']) configMap['libraryName'] = '꿈빛도서관';
      if (!configMap['adminName']) configMap['adminName'] = '사서교사';
      if (!configMap['adminId']) configMap['adminId'] = 'admin';
      if (!configMap['adminPassword']) configMap['adminPassword'] = 'admin1234';

      // 순서 고정 배열: Row 2(B2)는 libraryName, Row 3(B3)는 adminName, Row 4(B4)는 adminId, Row 5(B5)는 adminPassword
      const orderedKeys = [
        'libraryName',      // Row 2 -> B2
        'adminName',        // Row 3 -> B3
        'adminId',          // Row 4 -> B4
        'adminPassword',    // Row 5 -> B5
        'cycleType',        // Row 6 -> B6
        'slotCapacity',     // Row 7 -> B7
        'operatingDays',    // Row 8 -> B8
        'autoHolidaysClose',// Row 9 -> B9
        'cycleBaseDate',    // Row 10 -> B10
        'cycleBaseWeek',    // Row 11 -> B11
        'timeSlots'         // Row 12 -> B12
      ];

      const rowsToWrite = [['항목키', '설정값']];
      const writtenKeySet = {};

      orderedKeys.forEach(k => {
        if (configMap[k] !== undefined) {
          rowsToWrite.push([k, configMap[k]]);
          writtenKeySet[k] = true;
        }
      });

      // 기타 사용자 정의 키가 있다면 아래에 안전하게 보존
      Object.keys(configMap).forEach(k => {
        if (!writtenKeySet[k]) {
          rowsToWrite.push([k, configMap[k]]);
        }
      });

      sheet.clearContents();
      sheet.getRange(1, 1, rowsToWrite.length, 2).setValues(rowsToWrite);

      // 서식 지정 및 고정 컬럼 폭
      sheet.getRange(1, 1, 1, 2).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
      sheet.getRange(2, 1, rowsToWrite.length - 1, 1).setFontWeight('bold').setFontColor('#475569');
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 340);

      // 🚨 핵심: 변경 사항 즉시 시트에 박아 넣는 강제 플러시
      SpreadsheetApp.flush();

      return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'config_saved_flushed' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2-1. 엑셀 일괄 동기화 (명부 + 순번 2개 탭 원자적 덮어쓰기 & LockService 완비)
    if (action === 'saveExcelBatch') {
      const lock = LockService.getScriptLock();
      try {
        lock.waitLock(15000);

        // 1) 봉사자 정보 탭 덮어쓰기
        const memberSheet = ss.getSheetByName('봉사자 정보') || ss.insertSheet('봉사자 정보');
        const members = postData.members || [];
        memberSheet.clearContents();
        memberSheet.appendRow(['순번', '이름', '연락처']);
        const mRows = members.map((m, idx) => [idx + 1, m.name || '', m.phone || '']);
        if (mRows.length > 0) {
          memberSheet.getRange(2, 1, mRows.length, 3).setValues(mRows);
        }

        // 2) 봉사자 순번 탭 덮어쓰기 (N인 가변 열 동적 구성)
        const rosterSheet = ss.getSheetByName('봉사자 순번') || ss.insertSheet('봉사자 순번');
        const roster = postData.roster || [];
        const cap = Math.max(1, parseInt(postData.slotCapacity, 10) || 2);

        rosterSheet.clearContents();
        const header = ['순번', '주차', '요일', '시간대'];
        for (let r = 1; r <= cap; r++) {
          header.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
        }
        rosterSheet.appendRow(header);

        const rRows = roster.map((r, idx) => {
          const row = [idx + 1, r.week + '주차', r.day, normalizeTimeStr(r.time)];
          for (let i = 1; i <= cap; i++) {
            row.push(r['name' + i] || (i === 1 ? r.name1 : (i === 2 ? r.name2 : '')) || '');
          }
          return row;
        });

        if (rRows.length > 0) {
          rosterSheet.getRange(2, 1, rRows.length, header.length).setValues(rRows);
        }

        SpreadsheetApp.flush();

        return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'excel_batch_saved' }))
          .setMimeType(ContentService.MimeType.JSON);

      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
      } finally {
        lock.releaseLock();
      }
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
      SpreadsheetApp.flush();
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
      SpreadsheetApp.flush();
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
      SpreadsheetApp.flush();
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
      SpreadsheetApp.flush();
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

    // 8. 신학기 시트 탭 안전 초기화 (N인 가변 정원 헤더 완비)
    if (action === 'resetNewSemester') {
      const cap = Math.max(1, parseInt(postData.slotCapacity, 10) || 2);
      resetNewSemesterSheets(ss, cap);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'unknown_action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
} // [end of doPost]

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
} // [end of normalizeTimeStr]

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
} // [end of findRosterColumns]

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

  // 💡 다이렉트 고정 셀 직결 감지 (사서 선생님이 시트 B2~B5를 직접 수정하여 계정을 복구하는 경우)
  if (data.length >= 2 && data[1] && data[1][1] && (String(data[1][0]).trim() === 'libraryName' || !data[1][0])) {
    config.libraryName = String(data[1][1]).trim();
  }
  if (data.length >= 3 && data[2] && data[2][1] && (String(data[2][0]).trim() === 'adminName' || !data[2][0])) {
    config.adminName = String(data[2][1]).trim();
  }
  if (data.length >= 4 && data[3] && data[3][1] && (String(data[3][0]).trim() === 'adminId' || !data[3][0])) {
    config.adminId = String(data[3][1]).trim();
  }
  if (data.length >= 5 && data[4] && data[4][1] && (String(data[4][0]).trim() === 'adminPassword' || !data[4][0])) {
    config.adminPassword = String(data[4][1]).trim();
  }

  return config;
} // [end of getSystemConfig]

// -------------------------------------------------------------
// [5] 표준 5대 탭 자동 구성 및 신학기 초기화 (N인 가변 정원 완비)
// -------------------------------------------------------------
function setupStandardSheets(ss, slotCapacity) {
  const cap = Math.max(1, parseInt(slotCapacity, 10) || 2);

  // 1. 시스템 설정 (B2:도서관, B3:성함, B4:아이디, B5:비밀번호 고정 배치)
  let s1 = ss.getSheetByName('시스템 설정') || ss.insertSheet('시스템 설정');
  if (s1.getLastRow() === 0) {
    const defaultRows = [
      ['항목키', '설정값'],
      ['libraryName', '꿈빛도서관'],
      ['adminName', '사서교사'],
      ['adminId', 'admin'],
      ['adminPassword', 'admin1234'],
      ['cycleType', '2주'],
      ['slotCapacity', String(cap)],
      ['operatingDays', '["월","화","수","목","금"]'],
      ['autoHolidaysClose', 'true'],
      ['cycleBaseDate', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd')],
      ['cycleBaseWeek', '1']
    ];
    s1.getRange(1, 1, defaultRows.length, 2).setValues(defaultRows);
    s1.getRange(1, 1, 1, 2).setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
    s1.getRange(2, 1, defaultRows.length - 1, 1).setFontWeight('bold').setFontColor('#475569');
    s1.setColumnWidth(1, 160);
    s1.setColumnWidth(2, 340);
  }

  // 2. 봉사자 정보
  let s2 = ss.getSheetByName('봉사자 정보') || ss.insertSheet('봉사자 정보');
  if (s2.getLastRow() === 0) {
    s2.appendRow(['순번', '이름', '연락처']);
  }

  // 3. 봉사자 순번 (N인 동적 헤더)
  let s3 = ss.getSheetByName('봉사자 순번') || ss.insertSheet('봉사자 순번');
  if (s3.getLastRow() === 0) {
    const rosterHeader = ['순번', '주차', '요일', '시간대'];
    for (let r = 1; r <= cap; r++) {
      rosterHeader.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
    }
    s3.appendRow(rosterHeader);
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

  SpreadsheetApp.flush();
} // [end of setupStandardSheets]

function resetNewSemesterSheets(ss, slotCapacity) {
  const cap = Math.max(1, parseInt(slotCapacity, 10) || 2);
  const clearTab = (name, headers) => {
    const s = ss.getSheetByName(name);
    if (s) {
      s.clearContents();
      if (headers) s.appendRow(headers);
    }
  };

  clearTab('봉사자 정보', ['순번', '이름', '연락처']);

  const rosterHeader = ['순번', '주차', '요일', '시간대'];
  for (let r = 1; r <= cap; r++) {
    rosterHeader.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
  }
  clearTab('봉사자 순번', rosterHeader);

  clearTab('출석부', ['날짜', '시간', '이름', '출석확인']);
  clearTab('출석 및 변경 기록', ['기록일시', '슬롯키', '상태', '요청자', '대타/교체자', '사유']);

  SpreadsheetApp.flush();
} // [end of resetNewSemesterSheets]

// -------------------------------------------------------------
// [6] 월별 배정표 및 7열 격자 달력 탭 동시 생성 엔진
// -------------------------------------------------------------
function createMonthlySheet(year, month, slotCapacity, scheduleRows, gridRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getSystemConfig(ss.getSheetByName('시스템 설정'));
  const cap = Math.max(1, parseInt(slotCapacity, 10) || config.slotCapacity || 2);
  const sheetName = year + '-' + String(month).padStart(2, '0') + ' 봉사배정표';
  
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    try {
      ss.deleteSheet(sheet);
      sheet = ss.insertSheet(sheetName);
    } catch (e) {
      sheet.clear();
    }
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // 헤더 정의
  const header = ['활동 날짜', '요일', '순번 주차', '시간대'];
  for (let r = 1; r <= cap; r++) {
    header.push(r === 1 ? '봉사위원 1 (정)' : (r === 2 ? '봉사위원 2 (부)' : '봉사위원 ' + r));
  }
  header.push('비고 (대타 및 교체 내역)');

  sheet.appendRow(header);

  // 헤더 스타일링
  sheet.getRange(1, 1, 1, header.length)
    .setBackground('#059669')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  // 세부 배정 데이터 조립
  let finalScheduleRows = scheduleRows;
  let finalGridRows = gridRows;

  if (!finalScheduleRows || finalScheduleRows.length === 0) {
    const generated = generateMonthlyDataFromSheet(ss, year, month, cap, config);
    finalScheduleRows = generated.scheduleRows;
    finalGridRows = generated.gridRows;
  } else if (!finalGridRows || finalGridRows.length === 0) {
    finalGridRows = buildGridRowsFromSchedule(year, month, cap, finalScheduleRows);
  }

  // 데이터 적재 및 셀 서식 적용
  if (finalScheduleRows && finalScheduleRows.length > 0) {
    const dataRange = sheet.getRange(2, 1, finalScheduleRows.length, header.length);
    dataRange.setValues(finalScheduleRows);
    dataRange.setVerticalAlignment('middle');

    sheet.getRange(2, 1, finalScheduleRows.length, 3).setHorizontalAlignment('center');
    sheet.getRange(2, 5, finalScheduleRows.length, cap).setHorizontalAlignment('center');
    sheet.getRange(2, 4 + cap + 1, finalScheduleRows.length, 1).setHorizontalAlignment('left');

    for (let r = 2; r <= finalScheduleRows.length + 1; r++) {
      sheet.setRowHeight(r, 28);
      if (r % 2 === 1) {
        sheet.getRange(r, 1, 1, header.length).setBackground('#f8fafc');
      }
    }
  }

  const colWidths = [115, 75, 85, 140];
  for (let r = 1; r <= cap; r++) colWidths.push(130);
  colWidths.push(220);

  for (let c = 1; c <= header.length; c++) {
    sheet.setColumnWidth(c, colWidths[c - 1] || 120);
  }

  SpreadsheetApp.flush();

  createMonthlyCalendarGridSheet(ss, year, month, cap, finalGridRows);
} // [end of createMonthlySheet]

function createMonthlyCalendarGridSheet(ss, year, month, slotCapacity, gridRows) {
  const gridSheetName = year + '-' + String(month).padStart(2, '0') + ' 달력';
  let gSheet = ss.getSheetByName(gridSheetName);
  if (gSheet) {
    try {
      ss.deleteSheet(gSheet);
      gSheet = ss.insertSheet(gridSheetName);
    } catch (e) {
      gSheet.clear();
    }
  } else {
    gSheet = ss.insertSheet(gridSheetName);
  }

  const headers = ['일', '월', '화', '수', '목', '금', '토'];
  gSheet.appendRow(headers);
  gSheet.setRowHeight(1, 36);

  gSheet.getRange(1, 1, 1, 7)
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  gSheet.getRange(1, 1).setBackground('#e11d48');
  gSheet.getRange(1, 7).setBackground('#2563eb');

  if (gridRows && gridRows.length > 0) {
    const range = gSheet.getRange(2, 1, gridRows.length, 7);
    range.setValues(gridRows);
    range.setVerticalAlignment('top');
    range.setWrap(true);
    for (let r = 2; r <= gridRows.length + 1; r++) {
      gSheet.setRowHeight(r, 95);
    }
  }

  for (let c = 1; c <= 7; c++) {
    gSheet.setColumnWidth(c, 145);
  }

  SpreadsheetApp.flush();
} // [end of createMonthlyCalendarGridSheet]

function buildGridRowsFromSchedule(year, month, cap, scheduleRows) {
  const lastDay = new Date(year, month, 0).getDate();
  const dayMap = {};

  scheduleRows.forEach(row => {
    const dateStr = String(row[0] || '').trim();
    if (!dateStr || !dateStr.includes('-')) return;
    const parts = dateStr.split('-');
    const d = parseInt(parts[2], 10);
    if (isNaN(d)) return;

    if (!dayMap[d]) {
      const isClosed = String(row[2] || '').trim() === '-' || String(row[row.length - 1] || '').includes('휴관');
      if (isClosed) {
        const note = String(row[row.length - 1] || '휴관');
        dayMap[d] = d + '일\n(' + note + ')';
        return;
      } else {
        const weekStr = String(row[2] || '');
        dayMap[d] = d + '일 [' + weekStr + ']\n';
      }
    }

    if (String(row[2] || '').trim() !== '-') {
      const timeStr = String(row[3] || '').slice(0, 5);
      const workers = [];
      for (let r = 4; r < 4 + cap; r++) {
        const w = String(row[r] || '').trim();
        if (w && w !== '공석' && w !== '-') workers.push(w);
      }
      if (workers.length > 0) {
        dayMap[d] += timeStr + ' ' + workers.join('/') + '\n';
      }
    }
  });

  const gridRows = [];
  let weekGrid = ['', '', '', '', '', '', ''];

  for (let d = 1; d <= lastDay; d++) {
    const curDate = new Date(year, month - 1, d);
    const dow = curDate.getDay();
    weekGrid[dow] = (dayMap[d] || (d + '일')).trim();

    if (dow === 6 || d === lastDay) {
      gridRows.push(weekGrid);
      weekGrid = ['', '', '', '', '', '', ''];
    }
  }
  return gridRows;
} // [end of buildGridRowsFromSchedule]

// -------------------------------------------------------------
// [7] 스프레드시트 원본 데이터 기반 월간 일정 자동 생성기
// -------------------------------------------------------------
function generateMonthlyDataFromSheet(ss, year, month, cap, config) {
  const rosterRes = findRosterColumns(ss.getSheetByName('봉사자 순번') || ss.insertSheet('봉사자 순번'));
  const roster = rosterRes.roster || [];
  
  const overrides = {};
  const logSheet = ss.getSheetByName('출석 및 변경 기록');
  if (logSheet) {
    const lData = logSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < lData.length; i++) {
      const k = String(lData[i][1] || '').trim();
      if (k) overrides[k] = { status: lData[i][2], substitute: lData[i][4], requester: lData[i][3] };
    }
  }

  const lastDay = new Date(year, month, 0).getDate();
  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  const opDays = config.operatingDays || ['월', '화', '수', '목', '금'];
  const timeSlots = config.timeSlots || [{ name: '오전 1차', startTime: '08:30' }, { name: '오전 2차', startTime: '10:30' }];

  const scheduleRows = [];
  const gridRows = [];
  let weekGrid = ['', '', '', '', '', '', ''];

  for (let d = 1; d <= lastDay; d++) {
    const curDate = new Date(year, month - 1, d);
    const dow = curDate.getDay();
    const dayShort = dayNames[dow];
    const dateStr = Utilities.formatDate(curDate, 'Asia/Seoul', 'yyyy-MM-dd');
    const isOp = opDays.includes(dayShort);

    if (!isOp) {
      const row = [dateStr, dayShort + '요일', '-', '-'];
      for (let r = 1; r <= cap; r++) row.push('-');
      row.push(dayShort + '요일 정기휴관');
      scheduleRows.push(row);
      weekGrid[dow] = d + '일\n(' + dayShort + '요휴관)';
    } else {
      const bDate = new Date(config.cycleBaseDate || '2026-03-02');
      const diffWeeks = Math.floor((curDate - bDate) / (7 * 24 * 60 * 60 * 1000));
      const limit = (config.cycleType === '1주') ? 1 : ((config.cycleType === '1달') ? 4 : 2);
      let weekCycle = ((config.cycleBaseWeek - 1 + diffWeeks) % limit);
      if (weekCycle < 0) weekCycle += limit;
      weekCycle += 1;

      let cellText = d + '일 [' + weekCycle + '주]\n';

      timeSlots.forEach(slot => {
        const sNorm = normalizeTimeStr(slot.startTime);
        const rMatch = roster.find(r => r.week === weekCycle && r.day.startsWith(dayShort) && normalizeTimeStr(r.time) === sNorm);
        const workers = [];
        const notes = [];

        for (let r = 1; r <= cap; r++) {
          const k = dateStr + '_' + sNorm + '_' + r;
          const rec = overrides[k];
          const orig = rMatch ? (rMatch['name' + r] || '') : '';
          const eff = (rec && rec.status === 'ACCEPTED' && rec.substitute) ? rec.substitute : (rec?.requester || orig || '공석');
          workers.push(eff || '공석');
          if (rec && rec.status === 'ACCEPTED') notes.push('[대타] ' + rec.substitute);
        }

        scheduleRows.push([
          dateStr,
          dayShort + '요일',
          weekCycle + '주차',
          slot.startTime + ' (' + (slot.name || '') + ')',
          ...workers,
          notes.join(' / ')
        ]);

        cellText += slot.startTime.slice(0, 5) + ' ' + workers.join('/') + '\n';
      });

      weekGrid[dow] = cellText.trim();
    }

    if (dow === 6 || d === lastDay) {
      gridRows.push(weekGrid);
      weekGrid = ['', '', '', '', '', '', ''];
    }
  }

  return { scheduleRows, gridRows };
} // [end of generateMonthlyDataFromSheet]
