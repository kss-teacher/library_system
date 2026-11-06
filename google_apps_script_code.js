<<<<<<< SEARCH
 * [꿈빛 - 도서관 봉사인력 관리 시스템]
 * Google Apps Script 클라우드 통합 백엔드 엔진 v33.1
=======
 * [꿈빛 - 도서관 봉사인력 관리 시스템]
 * Google Apps Script 클라우드 통합 백엔드 엔진 v1.01
>>>>>>> REPLACE
<<<<<<< SEARCH
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
=======
function findRosterColumns(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!sheet) sheet = ss.getSheetByName('봉사자 순번') || ss.getSheets()[0];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { roster: [], slotCapacity: 2 };

  const header = data[0] || [];
  let maxCap = 2;
  const workerCols = [];

  // 헤더 5열부터(E열 이상) 존재하는 모든 봉사위원 컬럼 동적 탐지
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

    // 1부터 N번째 봉사위원까지 누락 없이 모두 수집
    workerCols.forEach(wc => {
      const val = String(row[wc.colIdx] || '').trim();
      item[`name${wc.roleIdx}`] = (val === '공석') ? '' : val;
      if (wc.roleIdx === 1) item.name1 = item[`name${wc.roleIdx}`];
      if (wc.roleIdx === 2) item.name2 = item[`name${wc.roleIdx}`];
    });

    roster.push(item);
  }
  return { roster, slotCapacity: maxCap };
}
>>>>>>> REPLACE
<<<<<<< SEARCH
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
=======
    cycleBaseDate: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'),
    cycleBaseWeek: 1,
    slotCapacity: 2,
    operatingDays: ['월', '화', '수', '목', '금'],
    autoHolidaysClose: true
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
>>>>>>> REPLACE
<<<<<<< SEARCH
    let roster = [];
    try {
      const rosterSheet = ss.getSheetByName('봉사자 순번');
      if (rosterSheet) {
        roster = findRosterColumns(rosterSheet).roster || [];
      }
    } catch (e1) {}
=======
    let roster = [];
    let detectedCap = 2;
    try {
      const rosterSheet = ss.getSheetByName('봉사자 순번');
      if (rosterSheet) {
        const res = findRosterColumns(rosterSheet);
        roster = res.roster || [];
        if (res.slotCapacity) detectedCap = res.slotCapacity;
      }
    } catch (e1) {}
>>>>>>> REPLACE
<<<<<<< SEARCH
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
=======
      adminPassword: config.adminPassword,
      adminEmail: config.adminEmail,
      cycleType: config.cycleType,
      slotCapacity: config.slotCapacity || detectedCap,
      operatingDays: config.operatingDays || ['월', '화', '수', '목', '금'],
      autoHolidaysClose: (config.autoHolidaysClose !== false),
      timeSlots: config.timeSlots,
      noticeMemo: config.defaultNoticeMemo,
      cycleBaseDate: config.cycleBaseDate,
      cycleBaseWeek: config.cycleBaseWeek,
      members: members,
      roster: roster, // N인 위원이 온전히 담긴 무결성 배열 전송
      overrides: overrides,
      attendance: attendance
    };
>>>>>>> REPLACE
<<<<<<< SEARCH
      if (postData.cycleBaseDate) updateKey('cycleBaseDate', postData.cycleBaseDate);
      if (postData.cycleBaseWeek !== undefined) updateKey('cycleBaseWeek', String(postData.cycleBaseWeek));

      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
=======
      if (postData.cycleBaseDate) updateKey('cycleBaseDate', postData.cycleBaseDate);
      if (postData.cycleBaseWeek !== undefined) updateKey('cycleBaseWeek', String(postData.cycleBaseWeek));
      if (postData.slotCapacity !== undefined) updateKey('slotCapacity', String(postData.slotCapacity));
      if (postData.operatingDays) updateKey('operatingDays', JSON.stringify(postData.operatingDays));
      if (postData.autoHolidaysClose !== undefined) updateKey('autoHolidaysClose', String(postData.autoHolidaysClose));

      return ContentService.createTextOutput(JSON.stringify({ status: 'success' })).setMimeType(ContentService.MimeType.JSON);
>>>>>>> REPLACE
