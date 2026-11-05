/**
 * ==============================================================================
 * [꿈빛 - 도서관 봉사인력 관리 시스템]
 * Google Apps Script 클라우드 통합 백엔드 엔진 v33.1
 * ==============================================================================
 * 
 * [주요 기능]
 * 1. 스프레드시트 상단 맞춤형 메뉴 자동 생성 (표준 5대 시트 원클릭 세팅 & 신학기 초기화)
 * 2. 가변 타임슬롯(1일 1타임, 2타임, 3타임 등) 및 운영 주기(1주/2주/1달) 지원
 * 3. 월별 일자별 봉사 배정표 시트 동적 자동 생성 (웹앱 원격 트리거 API 결합)
 * 4. 웹앱 REST API (doGet / doPost): 명부, 순번표, 공지사항, 출석부, 대타로그, 원격 배정표 생성
 * 5. 관리자 계정(ID/PW) 안전 저장 및 비상 초기화 지원
 */

// ==========================================
// 1. 구글 스프레드시트 상단 전용 메뉴
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 꿈빛 도서관 봉사 관리')
    .addItem('🌟 [최초 1회] 표준 5대 시트 탭 자동 구성', 'setupNewSchoolTemplate')
    .addItem('🌸 [신학기] 봉사단 시트 안전 초기화', 'menuResetNewSemester')
    .addSeparator()
    .addItem('📅 이번 달 배정표(목록+달력) 시트 생성', 'menuCreateThisMonthSheet')
    .addItem('⏭️ 다음 달 배정표(목록+달력) 시트 즉시 생성', 'menuCreateNextMonthSheet')
    .addSeparator()
    .addItem('⏰ [자동화] 매달 말일 배정표 자동 생성 예약', 'setupAutoMonthlyTrigger')
    .addItem('🔑 관리자 계정 초기화 (admin / admin1234)', 'menuResetAdminAccount')
    .addItem('🔍 시스템 연동 및 데이터 상태 진단', 'testSystem')
    .addToUi();
}

/**
 * 기본 표준 타임슬롯 정의
 */
function getDefaultTimeSlots() {
  return [
    { id: 't1', name: '오전 1차', startTime: '08:30', endTime: '10:30' },
    { id: 't2', name: '오전 2차', startTime: '10:30', endTime: '12:30' }
  ];
}

/**
 * 시간 형식 정규화 유틸리티 (08:30 <-> 8:30 통일)
 */
function normalizeTimeStr(t) {
  if (!t) return '';
  const s = String(t).trim();
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) {
    return String(m[1]).padStart(2, '0') + ':' + m[2];
  }
  return s;
}

// ==========================================
// 2. 표준 5대 시트 자동 구성 마법사
// ==========================================

function setupNewSchoolTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '🌟 표준 5대 시트 자동 구성',
    '이 스프레드시트에 꿈빛 시스템 운영에 필요한 필수 시트\n' +
    '(1. 시스템 설정  2. 봉사자 순번  3. 봉사자 정보  4. 출석부  5. 출석 및 변경 기록  6. 개인정보_동의이력  7. 관리자용 매뉴얼)를\n' +
    '표준 규격으로 자동 생성하시겠습니까?\n\n' +
    '※ 기존에 작성된 데이터가 있는 경우 덮어씌워질 수 있으니 주의하세요.',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  createOrResetStandardSheets(ss, '2주', '도서관', '관리자', '', getDefaultTimeSlots());

  ui.alert(
    '🎉 시트 자동 구성 완료!',
    '도서관 봉사단 운영에 필요한 표준 시트가 성공적으로 구성되었습니다.\n\n' +
    '1. [시스템 설정]에서 도서관 명칭과 시간대를 확인하세요.\n' +
    '2. [봉사자 순번] 및 [봉사자 정보]에 봉사위원님 명단을 등록하세요.\n' +
    '3. [배포] ➔ [웹 앱]으로 배포한 URL을 꿈빛 앱에 연결하시면 됩니다.',
    ui.ButtonSet.OK
  );
}

function createOrResetStandardSheets(ss, cycleType, libraryName, adminName, adminEmail, timeSlots) {
  cycleType = cycleType || '2주';
  libraryName = libraryName || '도서관';
  adminName = adminName || '관리자';
  adminEmail = adminEmail || '';
  timeSlots = (Array.isArray(timeSlots) && timeSlots.length > 0) ? timeSlots : getDefaultTimeSlots();

  // 1. [시스템 설정] 시트
  let cfgSheet = ss.getSheetByName('시스템 설정');
  if (!cfgSheet) cfgSheet = ss.insertSheet('시스템 설정');
  cfgSheet.clear();
  cfgSheet.setColumnWidth(1, 180);
  cfgSheet.setColumnWidth(2, 450);
  cfgSheet.getRange('A1:B1').setValues([['설정 항목', '설정 값']])
    .setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  
  const defaultConfigs = [
    ['libraryName', libraryName],
    ['adminName', adminName],
    ['adminId', 'admin'],
    ['adminPassword', 'admin1234'],
    ['adminEmail', adminEmail],
    ['cycleType', cycleType],
    ['timeSlots', JSON.stringify(timeSlots)],
    ['defaultNoticeMemo', '오늘도 아이들의 꿈을 위해 따뜻한 사랑으로 함께해 주시는 모든 도서관 봉사위원님들께 진심으로 감사드립니다.'],
    ['cycleBaseDate', Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd')],
    ['cycleBaseWeek', '1']
  ];
  cfgSheet.getRange(2, 1, defaultConfigs.length, 2).setValues(defaultConfigs);

  // 2. [봉사자 순번] 시트
  let rSheet = ss.getSheetByName('봉사자 순번');
  if (!rSheet) rSheet = ss.insertSheet('봉사자 순번');
  rSheet.clear();
  rSheet.setColumnWidth(1, 60);
  rSheet.setColumnWidth(2, 80);
  rSheet.setColumnWidth(3, 80);
  rSheet.setColumnWidth(4, 130);
  rSheet.setColumnWidth(5, 140);
  rSheet.setColumnWidth(6, 140);

  const rHeaders = ['순번', '주차', '요일', '시간대', '봉사위원 1 (정)', '봉사위원 2 (부)'];
  rSheet.getRange('A1:F1').setValues([rHeaders])
    .setBackground('#15803d').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  let weekLimit = (cycleType === '1주') ? 1 : ((cycleType === '1달') ? 4 : 2);
  const days = ['월요일', '화요일', '수요일', '목요일', '금요일'];
  const rosterRows = [];
  let seqId = 1;

  for (let w = 1; w <= weekLimit; w++) {
    for (let d = 0; d < days.length; d++) {
      for (let t = 0; t < timeSlots.length; t++) {
        const slot = timeSlots[t];
        const timeLabel = slot.name ? `${slot.startTime} (${slot.name})` : slot.startTime;
        rosterRows.push([seqId++, `${w}주차`, days[d], timeLabel, '', '']);
      }
    }
  }
  rSheet.getRange(2, 1, rosterRows.length, 6).setValues(rosterRows).setHorizontalAlignment('center');

  // 3. [봉사자 정보] 시트
  let memSheet = ss.getSheetByName('봉사자 정보');
  if (!memSheet) memSheet = ss.insertSheet('봉사자 정보');
  memSheet.clear();
  memSheet.setColumnWidth(1, 60);
  memSheet.setColumnWidth(2, 120);
  memSheet.setColumnWidth(3, 160);
  memSheet.getRange('A1:C1').setValues([['순번', '이름', '연락처']])
    .setBackground('#0369a1').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 4. [출석부] 시트
  let attSheet = ss.getSheetByName('출석부');
  if (!attSheet) attSheet = ss.insertSheet('출석부');
  attSheet.clear();
  attSheet.setColumnWidth(1, 130);
  attSheet.setColumnWidth(2, 220);
  attSheet.setColumnWidth(3, 130);
  attSheet.setColumnWidth(4, 100);
  attSheet.getRange('A1:D1').setValues([['날짜', '시간', '이름', '출석확인']])
    .setBackground('#047857').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 5. [출석 및 변경 기록] 시트
  let logSheet = ss.getSheetByName('출석 및 변경 기록');
  if (!logSheet) logSheet = ss.insertSheet('출석 및 변경 기록');
  logSheet.clear();
  logSheet.setColumnWidth(1, 150);
  logSheet.setColumnWidth(2, 100);
  logSheet.setColumnWidth(3, 90);
  logSheet.setColumnWidth(4, 110);
  logSheet.setColumnWidth(5, 120);
  logSheet.setColumnWidth(6, 150);
  logSheet.setColumnWidth(7, 160);
  logSheet.getRange('A1:G1').setValues([['슬롯키', '일자', '시간', '역할', '대상자/대타 성명', '사유', '기록일시']])
    .setBackground('#475569').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 6. [개인정보_동의이력] 시트 (PIPA 감사로그)
  let consentSheet = ss.getSheetByName('개인정보_동의이력');
  if (!consentSheet) consentSheet = ss.insertSheet('개인정보_동의이력');
  consentSheet.clear();
  consentSheet.setColumnWidth(1, 160);
  consentSheet.setColumnWidth(2, 110);
  consentSheet.setColumnWidth(3, 160);
  consentSheet.setColumnWidth(4, 120);
  consentSheet.setColumnWidth(5, 180);
  consentSheet.getRange('A1:E1').setValues([['기록일시', '구분', '성명 (식별값)', '동의 규격', '비고']])
    .setBackground('#7c3aed').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 7. [관리자용 매뉴얼] 시트
  let manualSheet = ss.getSheetByName('관리자용 매뉴얼');
  if (!manualSheet) manualSheet = ss.insertSheet('관리자용 매뉴얼');
  manualSheet.clear();
  manualSheet.setColumnWidth(1, 140);
  manualSheet.setColumnWidth(2, 600);
  manualSheet.getRange('A1:B1').merge()
    .setValue('📖 [꿈빛 시스템] 관리자 필수 운영 매뉴얼')
    .setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  manualSheet.setRowHeight(1, 38);

  const manualContent = [
    ['1. 시스템 개요', '꿈빛은 전국의 학교 도서관에서 서버 비용 0원으로 운영할 수 있는 자원봉사자 관리 솔루션입니다.'],
    ['2. 명부 작성 요령', '[봉사자 정보] 탭에 학부모님 성함과 연락처(010-0000-0000)를 입력하시면 웹앱과 실시간 연동됩니다.'],
    ['3. 순번표 작성 요령', '[봉사자 순번] 탭의 봉사위원 1 및 봉사위원 2 열에 배정할 위원 성함을 입력하세요.'],
    ['4. 배정표 시트 생성', '웹앱의 [월별 캘린더] 탭에서 [구글 시트에 배정표 탭 생성]을 누르면 해당 월의 배정표 탭이 자동 생성됩니다.'],
    ['5. 비밀번호 분실 시', '상단 메뉴 [📚 꿈빛 도서관 봉사 관리] ➔ [관리자 계정 초기화]를 누르면 admin / admin1234로 즉시 리셋됩니다.'],
    ['6. 배포 팁', 'Apps Script [배포] ➔ [배포 관리] ➔ [편집] ➔ [새 버전]으로 배포한 웹 앱 URL을 사용합니다.']
  ];
  manualSheet.getRange(2, 1, manualContent.length, 2).setValues(manualContent).setVerticalAlignment('middle');
  manualSheet.getRange('A2:A7').setBackground('#f1f5f9').setFontWeight('bold').setHorizontalAlignment('center');
  for (let r = 2; r <= 7; r++) { manualSheet.setRowHeight(r, 32); }

  const defaultSheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (defaultSheet1 && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet1); } catch (e) {}
  }

  ss.setActiveSheet(rSheet);
  return { status: 'success', cycleType, timeSlots };
}

// ==========================================
// 3. 관리자 계정 비상 리셋 (admin / admin1234)
// ==========================================

function menuResetAdminAccount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfgSheet = ss.getSheetByName('시스템 설정');
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet('시스템 설정');
    cfgSheet.appendRow(['설정 항목', '설정 값']);
  }

  const cData = cfgSheet.getDataRange().getDisplayValues();
  let idRow = -1;
  let pwRow = -1;

  for (let i = 1; i < cData.length; i++) {
    const k = String(cData[i][0]).trim();
    if (k === 'adminId') idRow = i + 1;
    if (k === 'adminPassword') pwRow = i + 1;
  }

  if (idRow > 0) cfgSheet.getRange(idRow, 2).setValue('admin');
  else cfgSheet.appendRow(['adminId', 'admin']);

  if (pwRow > 0) cfgSheet.getRange(pwRow, 2).setValue('admin1234');
  else cfgSheet.appendRow(['adminPassword', 'admin1234']);

  SpreadsheetApp.getUi().alert(
    '✅ 계정 초기화 완료',
    '관리자 계정이 기본값으로 재설정되었습니다.\n\n' +
    '아이디: admin\n비밀번호: admin1234\n\n' +
    '웹앱에 접속하여 최초 로그인 후 새 비밀번호를 변경하세요.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================
// 3-1. 신학기 봉사단 시트 탭 안전 초기화 마법사
// ==========================================

function menuResetNewSemester() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    '🌸 신학기 봉사단 시트 초기화',
    '새 학년도 봉사단 출범을 위해 이전 활동 데이터\n' +
    '([봉사자 순번 배정], [봉사자 정보], [출석부], [출석 및 변경 기록])를 깨끗이 초기화하시겠습니까?\n\n' +
    '※ 주의: 스프레드시트 상단 메뉴 [파일] ➔ [사본 만들기]로 이전 학기 백업본을 먼저 만들어 두시는 것을 강력히 권장합니다.',
    ui.ButtonSet.YES_NO
  );

  if (resp === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    resetNewSemesterSheets(ss);
    ui.alert(
      '✅ 신학기 초기화 완료',
      '봉사단 명부, 순번 배정, 출석부, 변경 기록이 새 학년을 위해 깨끗하게 초기화되었습니다.',
      ui.ButtonSet.OK
    );
  }
}

function resetNewSemesterSheets(ss) {
  // 1) 봉사자 정보 탭 초기화 (헤더 1행 보존, 2행 이하 공백화)
  let memSheet = ss.getSheetByName('봉사자 정보');
  if (memSheet && memSheet.getLastRow() > 1) {
    memSheet.getRange(2, 1, memSheet.getLastRow() - 1, memSheet.getLastColumn()).clearContent();
  }

  // 2) 출석부 탭 초기화 (헤더 1행 보존, 2행 이하 공백화)
  let attSheet = ss.getSheetByName('출석부');
  if (attSheet && attSheet.getLastRow() > 1) {
    attSheet.getRange(2, 1, attSheet.getLastRow() - 1, attSheet.getLastColumn()).clearContent();
  }

  // 3) 출석 및 변경 기록 탭 초기화 (헤더 1행 보존, 2행 이하 공백화)
  let logSheet = ss.getSheetByName('출석 및 변경 기록');
  if (logSheet && logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).clearContent();
  }

  // 4) 봉사자 순번 탭의 봉사위원1(E열), 봉사위원2(F열) 공백화 (헤더 및 순번/주차/요일/시간 구조 100% 보존)
  let rSheet = ss.getSheetByName('봉사자 순번');
  if (rSheet && rSheet.getLastRow() > 1) {
    rSheet.getRange(2, 5, rSheet.getLastRow() - 1, 2).clearContent();
  }
}

// ==========================================
// 4. 주차 연산 및 월별 배정표 시트 생성기
// ==========================================

function calculateContinuousWeekCycle(date, baseDateStr, baseWeekNum, cycleType) {
  cycleType = cycleType || '2주';
  if (cycleType === '1주') return 1;

  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  let bDate = new Date(2026, 2, 2);

  if (baseDateStr) {
    const cleanStr = String(baseDateStr).trim();
    const match = cleanStr.match(/(\d{4})[^\d]+(\d{1,2})[^\d]+(\d{1,2})/);
    if (match) {
      bDate = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10));
    }
  }

  const startWeek = (parseInt(baseWeekNum, 10) >= 1) ? parseInt(baseWeekNum, 10) : 1;

  const getMonday = (dt) => {
    const m = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
    const day = m.getDay();
    const diff = (day === 0 ? -6 : 1 - day);
    m.setDate(m.getDate() + diff);
    m.setHours(0, 0, 0, 0);
    return m;
  };

  const targetMonday = getMonday(d);
  const baseMonday = getMonday(bDate);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;
  const diffWeeks = Math.round((targetMonday.getTime() - baseMonday.getTime()) / msPerWeek);

  if (cycleType === '1달') {
    return ((startWeek - 1 + diffWeeks) % 4 + 4) % 4 + 1;
  }
  return ((startWeek - 1 + diffWeeks) % 2 + 2) % 2 + 1;
}

function menuCreateThisMonthSheet() {
  const now = new Date();
  const sName = createMonthlySheet(now.getFullYear(), now.getMonth() + 1);
  const cName = createMonthlyCalendarGridSheet(now.getFullYear(), now.getMonth() + 1);
  SpreadsheetApp.getUi().alert('✅ 시트 생성 완료', `[${sName}] 및 [${cName}] 탭이 성공적으로 생성되었습니다.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuCreateNextMonthSheet() {
  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const sName = createMonthlySheet(nextMonth.getFullYear(), nextMonth.getMonth() + 1);
  const cName = createMonthlyCalendarGridSheet(nextMonth.getFullYear(), nextMonth.getMonth() + 1);
  SpreadsheetApp.getUi().alert('✅ 시트 생성 완료', `[${sName}] 및 [${cName}] 탭이 성공적으로 생성되었습니다.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupAutoMonthlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkAndRunLastDayTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('checkAndRunLastDayTrigger').timeBased().everyDays(1).atHour(9).create();
  SpreadsheetApp.getUi().alert('✅ 자동 예약 완료', '매달 말일 오전 9시에 다음 달 봉사 배정표 시트가 자동 생성됩니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function checkAndRunLastDayTrigger() {
  const today = new Date();
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  if (tomorrow.getDate() === 1) {
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    createMonthlySheet(nextMonth.getFullYear(), nextMonth.getMonth() + 1);
    createMonthlyCalendarGridSheet(nextMonth.getFullYear(), nextMonth.getMonth() + 1);
  }
}

function createMonthlySheet(year, month, clientRows, libName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthStr = year + '-' + String(month).padStart(2, '0');
  const sheetName = monthStr + ' 봉사배정표';

  // 1. 기존 동일 명칭 탭이 있을 경우 선 삭제 후 재생성 (기존 병합 셀 충돌 및 예외 원천 방어)
  let targetSheet = ss.getSheetByName(sheetName);
  if (targetSheet) {
    try {
      ss.deleteSheet(targetSheet);
    } catch (e) {
      targetSheet.clear();
      try { targetSheet.getRange('A1:G1').breakApart(); } catch (e2) {}
    }
  }
  targetSheet = ss.insertSheet(sheetName);

  // 2. 가독성 높은 열 너비 세팅
  targetSheet.setColumnWidth(1, 110);
  targetSheet.setColumnWidth(2, 80);
  targetSheet.setColumnWidth(3, 90);
  targetSheet.setColumnWidth(4, 150);
  targetSheet.setColumnWidth(5, 140);
  targetSheet.setColumnWidth(6, 140);
  targetSheet.setColumnWidth(7, 220);

  const displayLibName = libName || (readSystemConfig(ss).libraryName || '도서관');

  // 3. 상단 대형 타이틀
  targetSheet.getRange('A1:G1').merge()
    .setValue(`[${displayLibName}] ${year}년 ${month}월 봉사 배정표`)
    .setBackground('#15803d').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  targetSheet.setRowHeight(1, 40);

  // 4. 헤더 행
  const headers = ['활동 날짜', '요일', '순번 주차', '시간대', '봉사위원 1 (정)', '봉사위원 2 (부)', '비고'];
  targetSheet.getRange(2, 1, 1, 7).setValues([headers])
    .setBackground('#e2e8f0').setFontColor('#334155').setFontWeight('bold').setHorizontalAlignment('center');
  targetSheet.setRowHeight(2, 28);

  let rows = [];

  // 5. 웹앱에서 정밀 계산된 실데이터가 넘어온 경우 1순위 즉시 적재
  if (clientRows && Array.isArray(clientRows) && clientRows.length > 0) {
    rows = clientRows;
  } else {
    // 스프레드시트 메뉴 등에서 단독 실행된 경우 시트 내부 데이터 기반 계산
    const rosterSheet = ss.getSheetByName('봉사자 순번') || ss.getSheets()[0];
    const { roster } = findRosterColumns(rosterSheet);
    const holidays = readHolidaysList(ss);
    const config = readSystemConfig(ss);
    const timeSlots = config.timeSlots || getDefaultTimeSlots();

    const logSheet = ss.getSheetByName('출석 및 변경 기록');
    const overrides = {};
    if (logSheet) {
      const logData = logSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < logData.length; i++) {
        const k = String(logData[i][0] || '').trim();
        const ovr = String(logData[i][4] || '').trim();
        if (k && ovr) overrides[k] = ovr;
      }
    }

    const lastDate = new Date(year, month, 0).getDate();
    const dayNames = ['일', '월', '화', '수', '목', '금', '토'];

    for (let d = 1; d <= lastDate; d++) {
      const curDate = new Date(year, month - 1, d);
      const dateStr = Utilities.formatDate(curDate, 'Asia/Seoul', 'yyyy-MM-dd');
      const dayOfWeek = curDate.getDay();
      const dayName = dayNames[dayOfWeek];
      const status = checkDayStatus(curDate, dateStr, holidays);

      if (status.isClosed) {
        rows.push([dateStr, dayName, '-', '-', '-', '-', status.title]);
      } else {
        const weekCycle = calculateContinuousWeekCycle(curDate, config.cycleBaseDate, config.cycleBaseWeek, config.cycleType);
        timeSlots.forEach(slot => {
          const sNorm = normalizeTimeStr(slot.startTime);
          const rMatch = roster.find(r => r.week === weekCycle && r.day.startsWith(dayName) && normalizeTimeStr(r.time) === sNorm);
          const ovr1 = overrides[`${dateStr}_${sNorm}_1`];
          const ovr2 = overrides[`${dateStr}_${sNorm}_2`];
          const n1 = ovr1 || (rMatch ? rMatch.name1 : '') || '공석';
          const n2 = ovr2 || (rMatch ? rMatch.name2 : '') || '(1인 근무)';
          rows.push([dateStr, dayName, `${weekCycle}주차`, slot.startTime, n1, n2, '']);
        });
      }
    }
  }

  // 6. 데이터 기입 및 정돈된 테두리 설정
  if (rows.length > 0) {
    targetSheet.getRange(3, 1, rows.length, 7).setValues(rows).setHorizontalAlignment('center');
    targetSheet.getRange(3, 1, rows.length, 7).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  }

  ss.setActiveSheet(targetSheet);
  return sheetName;
}

/**
 * 7열 격자형 월간 벽걸이 달력 시트 자동 생성 (A4 가로 인쇄 최적화)
 */
function createMonthlyCalendarGridSheet(year, month, clientRows, libName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthStr = year + '-' + String(month).padStart(2, '0');
  const sheetName = monthStr + ' 달력';

  // 1. 기존 동일 명칭 탭이 있을 경우 선 삭제 후 재생성 (병합 셀 충돌 방지)
  let targetSheet = ss.getSheetByName(sheetName);
  if (targetSheet) {
    try {
      ss.deleteSheet(targetSheet);
    } catch (e) {
      targetSheet.clear();
      try { targetSheet.getRange('A1:G1').breakApart(); } catch (e2) {}
    }
  }
  targetSheet = ss.insertSheet(sheetName);

  // 2. 가로 A4 인쇄에 최적인 7개 열(일~토) 너비 균등 설정
  for (let c = 1; c <= 7; c++) {
    targetSheet.setColumnWidth(c, 138);
  }

  const displayLibName = libName || (readSystemConfig(ss).libraryName || '도서관');

  // 3. 상단 대형 타이틀
  targetSheet.getRange('A1:G1').merge()
    .setValue(`[${displayLibName}] ${year}년 ${month}월 도서관 봉사활동 달력`)
    .setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  targetSheet.setRowHeight(1, 40);

  // 4. 요일 헤더 (일~토)
  const daysHeader = ['일 (Sun)', '월 (Mon)', '화 (Tue)', '수 (Wed)', '목 (Thu)', '금 (Fri)', '토 (Sat)'];
  targetSheet.getRange('A2:G2').setValues([daysHeader])
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  targetSheet.setRowHeight(2, 28);
  
  targetSheet.getRange('A2').setBackground('#ffe4e6').setFontColor('#be123c'); // 일요일 (레드)
  targetSheet.getRange('B2:F2').setBackground('#f1f5f9').setFontColor('#334155'); // 평일 (슬레이트)
  targetSheet.getRange('G2').setBackground('#dbeafe').setFontColor('#1d4ed8'); // 토요일 (블루)

  // 5. 당월 데이터 맵 구성
  const scheduleMap = {};
  if (clientRows && Array.isArray(clientRows) && clientRows.length > 0) {
    clientRows.forEach(r => {
      const dStr = String(r[0] || '').trim();
      if (!dStr) return;
      if (!scheduleMap[dStr]) {
        scheduleMap[dStr] = {
          isClosed: (r[2] === '-'),
          title: r[6] || '',
          weekCycle: r[2] || '',
          slots: []
        };
      }
      if (r[2] !== '-') {
        const timePart = String(r[3] || '').split(' ')[0];
        const n1 = String(r[4] || '공석').trim();
        const n2 = String(r[5] || '공석').trim();
        scheduleMap[dStr].slots.push({
          time: timePart,
          worker1: n1,
          worker2: n2
        });
      }
    });
  } else {
    const rosterSheet = ss.getSheetByName('봉사자 순번') || ss.getSheets()[0];
    const { roster } = findRosterColumns(rosterSheet);
    const holidays = readHolidaysList(ss);
    const config = readSystemConfig(ss);
    const timeSlots = config.timeSlots || getDefaultTimeSlots();

    const logSheet = ss.getSheetByName('출석 및 변경 기록');
    const overrides = {};
    if (logSheet) {
      const logData = logSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < logData.length; i++) {
        const k = String(logData[i][0] || '').trim();
        const ovr = String(logData[i][4] || '').trim();
        if (k && ovr) overrides[k] = ovr;
      }
    }

    const lastDate = new Date(year, month, 0).getDate();
    const dayNames = ['일', '월', '화', '수', '목', '금', '토'];

    for (let d = 1; d <= lastDate; d++) {
      const curDate = new Date(year, month - 1, d);
      const dateStr = Utilities.formatDate(curDate, 'Asia/Seoul', 'yyyy-MM-dd');
      const dayOfWeek = curDate.getDay();
      const dayName = dayNames[dayOfWeek];
      const status = checkDayStatus(curDate, dateStr, holidays);

      scheduleMap[dateStr] = {
        isClosed: status.isClosed,
        title: status.title,
        weekCycle: status.isClosed ? '' : `${calculateContinuousWeekCycle(curDate, config.cycleBaseDate, config.cycleBaseWeek, config.cycleType)}주차`,
        slots: []
      };

      if (!status.isClosed) {
        const weekCycle = calculateContinuousWeekCycle(curDate, config.cycleBaseDate, config.cycleBaseWeek, config.cycleType);
        timeSlots.forEach(slot => {
          const sNorm = normalizeTimeStr(slot.startTime);
          const rMatch = roster.find(r => r.week === weekCycle && r.day.startsWith(dayName) && normalizeTimeStr(r.time) === sNorm);
          const ovr1 = overrides[`${dateStr}_${sNorm}_1`];
          const ovr2 = overrides[`${dateStr}_${sNorm}_2`];
          const n1 = ovr1 || (rMatch ? rMatch.name1 : '') || '공석';
          const n2 = ovr2 || (rMatch ? rMatch.name2 : '') || '(1인)';
          scheduleMap[dateStr].slots.push({
            time: sNorm,
            worker1: n1,
            worker2: n2
          });
        });
      }
    }
  }

  // 6. 달력 그리드 행/열 배치
  const firstDayOfWeek = new Date(year, month - 1, 1).getDay();
  const lastDateNum = new Date(year, month, 0).getDate();

  let curRow = 3;
  let dayCounter = 1;

  while (dayCounter <= lastDateNum) {
    const dateNumRow = [];
    const contentRow = [];

    for (let col = 0; col < 7; col++) {
      if ((curRow === 3 && col < firstDayOfWeek) || dayCounter > lastDateNum) {
        dateNumRow.push('');
        contentRow.push('');
      } else {
        const curDate = new Date(year, month - 1, dayCounter);
        const dateStr = Utilities.formatDate(curDate, 'Asia/Seoul', 'yyyy-MM-dd');
        const info = scheduleMap[dateStr];

        dateNumRow.push(dayCounter);

        if (!info) {
          contentRow.push('');
        } else if (info.isClosed) {
          contentRow.push(`\n[${info.title || '휴관'}]`);
        } else {
          let lines = [];
          if (info.weekCycle) lines.push(`[${info.weekCycle}]`);
          info.slots.forEach(s => {
            lines.push(`${s.time} ${s.worker1}, ${s.worker2}`);
          });
          contentRow.push(lines.join('\n'));
        }

        dayCounter++;
      }
    }

    // 날짜 번호 행 (높이 24)
    const dateRange = targetSheet.getRange(curRow, 1, 1, 7);
    dateRange.setValues([dateNumRow])
      .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setBackground('#f8fafc');
    targetSheet.setRowHeight(curRow, 24);

    targetSheet.getRange(curRow, 1).setFontColor('#be123c'); // 일요일 레드
    targetSheet.getRange(curRow, 7).setFontColor('#1d4ed8'); // 토요일 블루

    // 배정 내용 행 (높이 88, 자동 줄바꿈)
    const contentRange = targetSheet.getRange(curRow + 1, 1, 1, 7);
    contentRange.setValues([contentRow])
      .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('top')
      .setWrap(true);
    targetSheet.setRowHeight(curRow + 1, 88);

    // 격자 테두리
    targetSheet.getRange(curRow, 1, 2, 7)
      .setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

    curRow += 2;
  }

  // 7. 하단 A4 인쇄 가이드 배너
  targetSheet.getRange(curRow, 1, 1, 7).merge()
    .setValue('💡 [인쇄 팁] 상단 메뉴 [파일] ➔ [인쇄]에서 [가로 방향] 및 [페이지에 맞춤]으로 설정하시면 A4 1장에 꼭 맞게 인쇄됩니다.')
    .setFontColor('#64748b').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#f8fafc');
  targetSheet.setRowHeight(curRow, 28);

  return sheetName;
}

// ==========================================
// 5. 시트 열 스마트 감지 및 데이터 파서
// ==========================================

function findRosterColumns(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!sheet) sheet = ss.getSheetByName('봉사자 순번') || ss.getSheets()[0];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { roster: [] };

  const roster = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1] && !row[2]) continue;

    let weekVal = 1;
    const mWeek = String(row[1] || '').match(/\d+/);
    if (mWeek) weekVal = parseInt(mWeek[0], 10);

    roster.push({
      id: i,
      week: weekVal,
      day: String(row[2] || '').trim(),
      time: normalizeTimeStr(row[3]),
      name1: String(row[4] || '').trim(),
      name2: String(row[5] || '').trim()
    });
  }
  return { roster };
}

function readMembersList(ss) {
  const memSheet = ss.getSheetByName('봉사자 정보');
  if (!memSheet) return [];
  const data = memSheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];

  const members = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][1] || '').trim();
    if (name) {
      members.push({
        id: i,
        name: name,
        phone: String(data[i][2] || '').trim()
      });
    }
  }
  return members;
}

function readSystemConfig(ss) {
  let sheet = ss.getSheetByName('시스템 설정');
  const config = {
    libraryName: '도서관',
    adminName: '관리자',
    adminId: 'admin',
    adminPassword: 'admin1234',
    adminEmail: '',
    cycleType: '2주',
    timeSlots: getDefaultTimeSlots(),
    defaultNoticeMemo: '오늘도 아이들의 꿈을 위해 따뜻한 사랑으로 함께해 주시는 모든 도서관 봉사위원님들께 진심으로 감사드립니다.',
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
    if (k === 'timeSlots' && v) {
      try {
        const parsed = JSON.parse(v);
        if (Array.isArray(parsed) && parsed.length > 0) config.timeSlots = parsed;
      } catch (e) {}
    }
    if (k === 'defaultNoticeMemo' && v) config.defaultNoticeMemo = v;
    if (k === 'cycleBaseDate' && v) config.cycleBaseDate = v;
    if (k === 'cycleBaseWeek' && v) config.cycleBaseWeek = parseInt(v, 10) || 1;
  }
  return config;
}

function checkDayStatus(curDate, dateStr, holidays) {
  const day = curDate.getDay();
  if (day === 0 || day === 6) {
    return { isClosed: true, title: day === 0 ? '일요일 휴관' : '토요일 휴관' };
  }
  const statutoryHolidays = {
    '2026-01-01': '신정', '2026-02-16': '설연휴', '2026-02-17': '설날', '2026-02-18': '설연휴',
    '2026-03-01': '3·1절', '2026-03-02': '대체공휴일', '2026-05-05': '어린이날', '2026-05-24': '부처님오신날',
    '2026-05-25': '대체공휴일', '2026-06-06': '현충일', '2026-08-15': '광복절', '2026-08-17': '대체공휴일',
    '2026-09-24': '추석연휴', '2026-09-25': '추석', '2026-09-26': '추석연휴', '2026-10-03': '개천절',
    '2026-10-05': '대체공휴일', '2026-10-09': '한글날', '2026-12-25': '성탄절'
  };
  if (statutoryHolidays[dateStr]) {
    return { isClosed: true, title: statutoryHolidays[dateStr] };
  }
  return { isClosed: false, title: '' };
}

function readHolidaysList(ss) {
  return [];
}

// ==========================================
// 6. 웹앱 REST API 통신 (doGet / doPost)
// ==========================================

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: '활성화된 스프레드시트를 찾을 수 없습니다. 스프레드시트에 포함된 Apps Script인지 확인해 주세요.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let roster = [];
    try {
      const rosterSheet = ss.getSheetByName('봉사자 순번');
      if (rosterSheet) {
        roster = findRosterColumns(rosterSheet).roster || [];
      }
    } catch (e1) {}

    let members = [];
    try {
      members = readMembersList(ss) || [];
    } catch (e2) {}

    let config = {
      libraryName: '도서관',
      adminName: '관리자',
      adminId: 'admin',
      adminPassword: 'admin1234',
      adminEmail: '',
      cycleType: '2주',
      timeSlots: getDefaultTimeSlots(),
      defaultNoticeMemo: '오늘도 아이들의 꿈을 위해 따뜻한 사랑으로 함께해 주시는 모든 도서관 봉사위원님들께 진심으로 감사드립니다.',
      cycleBaseDate: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'),
      cycleBaseWeek: 1
    };
    try {
      const c = readSystemConfig(ss);
      if (c) config = Object.assign(config, c);
    } catch (e3) {}

    const overrides = {};
    try {
      const logSheet = ss.getSheetByName('출석 및 변경 기록');
      if (logSheet) {
        const logData = logSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < logData.length; i++) {
          const k = String(logData[i][0] || '').trim();
          const ovr = String(logData[i][4] || '').trim();
          const rsn = String(logData[i][5] || '').trim();
          if (k && ovr) overrides[k] = { name: ovr, reason: rsn };
        }
      }
    } catch (e4) {}

    // 출석부 동기화 데이터 조회 (O / X 상태)
    const attendance = {};
    try {
      const attSheet = ss.getSheetByName('출석부');
      if (attSheet) {
        const attData = attSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < attData.length; i++) {
          const d = String(attData[i][0] || '').trim();
          const t = String(attData[i][1] || '').trim();
          const n = String(attData[i][2] || '').trim();
          const st = String(attData[i][3] || '').trim();
          const k = attData[i][4] ? String(attData[i][4]).trim() : '';
          if (k) {
            attendance[k] = { status: st, name: n, time: '' };
          }
        }
      }
    } catch (e5) {}

    let sheetUrl = '';
    try {
      sheetUrl = ss.getUrl();
    } catch (e6) {}

    const payload = {
      status: 'success',
      spreadsheetUrl: sheetUrl,
      libraryName: config.libraryName,
      adminName: config.adminName,
      adminId: config.adminId,
      adminPassword: config.adminPassword,
      adminEmail: config.adminEmail,
      cycleType: config.cycleType,
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

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: '데이터가 없습니다.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    // 0. 표준 5대 시트 탭 원격 자동 구성 (웹앱 STEP 05 연동)
    if (action === 'setupStandardSheets' || action === 'setupSheets') {
      const res = createOrResetStandardSheets(
        ss,
        postData.cycleType || '2주',
        postData.libraryName || '도서관',
        postData.adminName || '관리자',
        postData.adminEmail || '',
        postData.timeSlots || getDefaultTimeSlots()
      );
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: '표준 5대 탭이 성공적으로 자동 구성되었습니다.',
        details: res
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 1. 개인정보 동의 감사로그 저장 (PIPA)
    if (action === 'saveConsentLog') {
      let sheet = ss.getSheetByName('개인정보_동의이력');
      if (!sheet) {
        sheet = ss.insertSheet('개인정보_동의이력');
        sheet.appendRow(['기록일시', '구분', '성명 (식별값)', '동의 규격', '비고']);
      }
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([
        nowStr,
        postData.userType || '정규봉사위원',
        postData.userName || '',
        postData.version || 'PIPA-v1.0',
        postData.note || '모바일 본인인증 완료'
      ]);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 공지사항 인라인 실시간 저장
    if (action === 'saveNoticeMemo' || action === 'saveNotice') {
      let cfgSheet = ss.getSheetByName('시스템 설정') || ss.insertSheet('시스템 설정');
      const cData = cfgSheet.getDataRange().getDisplayValues();
      let found = false;
      const text = postData.noticeMemo || postData.defaultNoticeMemo || '';

      for (let i = 1; i < cData.length; i++) {
        if (String(cData[i][0]).trim() === 'defaultNoticeMemo') {
          cfgSheet.getRange(i + 1, 2).setValue(text);
          found = true;
          break;
        }
      }
      if (!found) cfgSheet.appendRow(['defaultNoticeMemo', text]);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', noticeMemo: text })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. 도서관 운영 설정 및 관리자 계정 저장
    if (action === 'saveSystemConfig' || action === 'saveSettings') {
      let cfgSheet = ss.getSheetByName('시스템 설정') || ss.insertSheet('시스템 설정');
      const cData = cfgSheet.getDataRange().getDisplayValues();
      const map = {};
      for (let i = 1; i < cData.length; i++) {
        map[String(cData[i][0]).trim()] = i + 1;
      }

      const updateKey = (k, v) => {
        if (v === undefined || v === null) return;
        if (map[k]) cfgSheet.getRange(map[k], 2).setValue(v);
        else {
          cfgSheet.appendRow([k, v]);
          map[k] = cfgSheet.getLastRow();
        }
      };

      if (postData.libraryName) updateKey('libraryName', postData.libraryName);
      if (postData.adminName) updateKey('adminName', postData.adminName);
      if (postData.adminId) updateKey('adminId', String(postData.adminId).trim().toLowerCase());
      if (postData.adminPassword) updateKey('adminPassword', postData.adminPassword);
      if (postData.adminEmail !== undefined) updateKey('adminEmail', postData.adminEmail);
      if (postData.cycleType) updateKey('cycleType', postData.cycleType);
      if (postData.cycleBaseDate) updateKey('cycleBaseDate', postData.cycleBaseDate);
      if (postData.cycleBaseWeek !== undefined) updateKey('cycleBaseWeek', String(postData.cycleBaseWeek));

      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }

    // 4. 대타 지정 및 해제 로그 (관리자 직권 교체 사유 연동 강화)
    if (action === 'setOverride') {
      let logSheet = ss.getSheetByName('출석 및 변경 기록') || ss.insertSheet('출석 및 변경 기록');
      const data = logSheet.getDataRange().getDisplayValues();
      let foundRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === postData.slotKey) {
          foundRow = i + 1;
          break;
        }
      }
      if (postData.name) {
        const parts = postData.slotKey.split('_');
        const dt = parts[0] || '';
        const tm = parts[1] || '';
        const role = parts[2] === '1' ? '봉사위원 1' : '봉사위원 2';
        const nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        const rsn = postData.reason || '대타/직권 교체';
        if (foundRow > 0) {
          logSheet.getRange(foundRow, 5).setValue(postData.name);
          logSheet.getRange(foundRow, 6).setValue(rsn);
          logSheet.getRange(foundRow, 7).setValue(nowStr);
        } else {
          logSheet.appendRow([postData.slotKey, dt, tm, role, postData.name, rsn, nowStr]);
        }
      } else {
        if (foundRow > 0) logSheet.deleteRow(foundRow);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }

    // 5. 순번표 일괄 저장
    if (action === 'saveRoster' && Array.isArray(postData.roster)) {
      let rSheet = ss.getSheetByName('봉사자 순번') || ss.insertSheet('봉사자 순번');
      rSheet.clear();
      const rows = [['순번', '주차', '요일', '시간대', '봉사위원 1 (정)', '봉사위원 2 (부)']];
      postData.roster.forEach(r => {
        const dStr = String(r.day || '').replace(/요일/, '') + '요일';
        rows.push([r.id, `${r.week}주차`, dStr, r.time, r.name1 || '', r.name2 || '']);
      });
      rSheet.getRange(1, 1, rows.length, 6).setValues(rows).setHorizontalAlignment('center');
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
    }

    // 6. 봉사자 명부 실시간 일괄 저장 (CRUD 동기화)
    if (action === 'saveMembers' && Array.isArray(postData.members)) {
      let memSheet = ss.getSheetByName('봉사자 정보');
      if (!memSheet) {
        memSheet = ss.insertSheet('봉사자 정보');
        memSheet.setColumnWidth(1, 60);
        memSheet.setColumnWidth(2, 120);
        memSheet.setColumnWidth(3, 160);
      }
      memSheet.clear();
      memSheet.getRange('A1:C1').setValues([['순번', '이름', '연락처']])
        .setBackground('#0369a1').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

      const rows = [];
      postData.members.forEach((m, idx) => {
        if (m && m.name) {
          rows.push([idx + 1, String(m.name).trim(), String(m.phone || '').trim()]);
        }
      });
      if (rows.length > 0) {
        memSheet.getRange(2, 1, rows.length, 3).setValues(rows).setHorizontalAlignment('center');
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', count: rows.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 7. 출석부 실시간 저장 (날짜 | 시간 | 이름 | 출석확인 4대 표준 규격)
    if (action === 'saveAttendanceCheck' || action === 'saveAttendance') {
      let attSheet = ss.getSheetByName('출석부');
      if (!attSheet) {
        attSheet = ss.insertSheet('출석부');
        attSheet.setColumnWidth(1, 130);
        attSheet.setColumnWidth(2, 220);
        attSheet.setColumnWidth(3, 130);
        attSheet.setColumnWidth(4, 100);
        attSheet.getRange('A1:D1').setValues([['날짜', '시간', '이름', '출석확인']])
          .setBackground('#047857').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      }

      const dateVal = String(postData.date || '').trim();
      const timeVal = String(postData.time || '').trim();
      const nameVal = String(postData.name || '').trim();
      const statusVal = String(postData.status || 'O').trim();
      const slotKeyVal = String(postData.slotKey || '').trim();

      const data = attSheet.getDataRange().getDisplayValues();
      let foundRow = -1;

      for (let i = 1; i < data.length; i++) {
        const rowSlotKey = data[i][4] ? String(data[i][4]).trim() : '';
        const rowDate = String(data[i][0] || '').trim();
        const rowTime = String(data[i][1] || '').trim();
        const rowName = String(data[i][2] || '').trim();

        if (slotKeyVal && rowSlotKey === slotKeyVal) {
          foundRow = i + 1;
          break;
        } else if (rowDate === dateVal && rowTime.includes(timeVal.split(' ')[0]) && (rowName === nameVal || !nameVal)) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow > 0) {
        attSheet.getRange(foundRow, 3).setValue(nameVal);
        attSheet.getRange(foundRow, 4).setValue(statusVal).setHorizontalAlignment('center').setFontWeight('bold');
        if (slotKeyVal) attSheet.getRange(foundRow, 5).setValue(slotKeyVal);
      } else {
        attSheet.appendRow([dateVal, timeVal, nameVal, statusVal, slotKeyVal]);
        const lastRow = attSheet.getLastRow();
        attSheet.getRange(lastRow, 1, 1, 4).setHorizontalAlignment('center');
        attSheet.getRange(lastRow, 4).setFontWeight('bold');
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        date: dateVal,
        name: nameVal,
        attendance: statusVal
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 8. 신학기 봉사단 안전 초기화 원격 트리거
    if (action === 'resetNewSemester') {
      resetNewSemesterSheets(ss);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: '신학기 봉사단 시트가 깨끗하게 초기화되었습니다.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 9. 월별 봉사 배정표 시트 탭 원격 자동 생성 (웹앱 월별 캘린더 연동)
    if (action === 'createMonthlySheet') {
      const targetYear = parseInt(postData.year, 10) || new Date().getFullYear();
      const targetMonth = parseInt(postData.month, 10) || (new Date().getMonth() + 1);
      const libName = postData.libraryName || '도서관';
      const clientRows = Array.isArray(postData.scheduleRows) ? postData.scheduleRows : null;
      
      const createdSheetName = createMonthlySheet(targetYear, targetMonth, clientRows, libName);
      const createdCalName = createMonthlyCalendarGridSheet(targetYear, targetMonth, clientRows, libName);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        sheetName: createdSheetName,
        calSheetName: createdCalName,
        message: `[${createdSheetName}] 및 [${createdCalName}] 탭이 구글 시트에 성공적으로 동시 생성되었습니다.`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'unknown_action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 7. 시스템 연동 및 데이터 상태 진단
// ==========================================

function testSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readSystemConfig(ss);
  const rosterSheet = ss.getSheetByName('봉사자 순번');
  const roster = rosterSheet ? findRosterColumns(rosterSheet).roster : [];
  const members = readMembersList(ss);

  Logger.log(`📂 현재 스프레드시트: ${ss.getName()} (${ss.getUrl()})`);
  Logger.log(`⚙️ 도서관명: "${config.libraryName}", 관리자ID: "${config.adminId}"`);
  Logger.log(`👥 등록된 봉사위원 수: ${members.length}명`);
  Logger.log(`📋 배정된 순번 슬롯 수: ${roster.length}개`);
  Logger.log(`🎉 [진단 성공] 스프레드시트 접근 권한이 정상 승인되었습니다. 이제 웹 앱을 [새 버전]으로 배포하세요!`);
}
