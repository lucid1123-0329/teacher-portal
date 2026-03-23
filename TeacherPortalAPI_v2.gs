/**
 * ============================================================
 * 명불허전학원 통합 관리 포털 API v2.0 (TeacherPortalAPI.gs)
 * ============================================================
 * 실제 데이터베이스 구조 기반 (2026-03-20 분석)
 * 
 * 시트 매핑:
 *   [CONFIG] 설정           — 평가 드롭다운 옵션
 *   [DB] 학생               — 학생 마스터 (144명)
 *   [DB] 반                 — 반 마스터 (29개)
 *   [DB] 수강               — 학생-반 매핑 (181건)
 *   [ALL] 전체 로그 취합     — 일일 평가 로그 (7,727건)
 *   [DB] 성향분석           — AI 분석 + 레벨테스트
 *   [MASTER] 강사계정        — 신규 생성 (로그인/역할)
 *   Master List (Helper)    — 조인 뷰 (읽기 전용)
 * 
 * 별도 시트:
 *   신규 상담 시트 (1C3hlNGzE9qoBYfcHfSCMRQEjyz7aFyofj8TKgrz-eCc)
 */

// ─── 설정 ───────────────────────────────────────────────
const TP = {
  // 시트명 (기존 데이터베이스)
  SHEET_CONFIG: '[CONFIG] 설정',
  SHEET_STUDENTS: '[DB] 학생',
  SHEET_CLASSES: '[DB] 반',
  SHEET_ENROLLMENT: '[DB] 수강',
  SHEET_DAILY_LOG: '[ALL] 전체 로그 취합',
  SHEET_PORTAL_LOG: '[ALL] 전체 로그 취합',   // ★ 포털도 [ALL]에 직접 저장
  PORTAL_START_ROW: 10001,                     // ★ 기존 데이터 10000행 이후부터 쓰기
  SHEET_PERSONALITY: '[DB] 성향분석',
  SHEET_TEACHERS: '[MASTER] 강사계정',
  SHEET_MASTER_LIST: 'Master List (Helper)',
  // 메인 시트의 상담 관련 탭
  SHEET_CONSULT_LOG: '[DB] 상담로그',
  SHEET_DASHBOARD: '[DASHBOARD] 상담 주기 관리',
  // 별도 스프레드시트
  CONSULT_SHEET_ID: '1C3hlNGzE9qoBYfcHfSCMRQEjyz7aFyofj8TKgrz-eCc',   // Sheet 2: 신규 상담
  CONSULT_SYS_ID: '1m1I9XtviE5aesy9DYqlm_G_Fxp2LJQov84S-gkyMgNI',     // Sheet 3: 통합 상담 시스템
  // 토큰
  TOKEN_SECRET: 'mbhj-portal-2026-v2',
  TOKEN_EXPIRY_HOURS: 24,
};

// ─── doGet / doPost (Sheet 1 전용 — 기존 Code.gs/JSONParser.gs와 충돌 없음) ───
// ⚠️ Sheet 1(데이터베이스)에는 기존 doGet이 없으므로 여기서 새로 정의합니다.
//    기존 Code.gs의 onOpen() 메뉴, JSONParser.gs의 맞춤함수는 그대로 작동합니다.

function doGet(e) {
  var action = (e.parameter.action || '');
  if (action.indexOf('tp_') === 0) {
    return tpHandleGet_(e);
  }
  // tp_ 접두어가 아닌 요청은 안내 메시지 반환
  return ContentService.createTextOutput(JSON.stringify({
    service: '명불허전학원 강사 포털 API',
    status: 'running',
    usage: 'action=tp_login&name=이름&pw=비밀번호'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var pd = JSON.parse(e.postData.contents);
    if (pd.action && pd.action.indexOf('tp_') === 0) {
      return tpHandlePost_(pd);
    }
  } catch(err) {}
  return ContentService.createTextOutput(JSON.stringify({error: 'Unknown request'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function tpHandleGet_(e) {
  const p = e.parameter;
  let r;
  try {
    switch (p.action) {
      case 'tp_login':        r = tpLogin_(p.name, p.pw); break;
      case 'tp_init':         r = tpInit_(p.teacher, p.role, p.token); break;
      case 'tp_config':       r = tpGetConfig_(p.token); break;
      case 'tp_classes':      r = tpGetClasses_(p.teacher, p.role, p.token); break;
      case 'tp_students':     r = tpGetStudents_(p.className, p.token); break;
      case 'tp_todayStatus':  r = tpGetTodayStatus_(p.teacher, p.role, p.token); break;
      case 'tp_studentProfile': r = tpGetStudentProfile_(p.student, p.token); break;
      case 'tp_studentReports': r = tpGetStudentReports_(p.student, p.token); break;
      case 'tp_myStudents':    r = tpGetMyStudents_(p.teacher, p.role, p.token); break;
      // 관리자 전용
      case 'tp_admin_overview':   r = tpAdminOverview_(p.token); break;
      case 'tp_admin_evalLog':    r = tpAdminEvalLog_(p.token, p.date, p.teacher, p.className); break;
      case 'tp_admin_allStudents':r = tpAdminAllStudents_(p.token); break;
      case 'tp_admin_allClasses': r = tpAdminAllClasses_(p.token); break;
      case 'tp_admin_teachers':   r = tpAdminTeachers_(p.token); break;
      case 'tp_admin_consultList':r = tpAdminConsultList_(p.token); break;
      case 'tp_admin_syncCounsel':r = tpAdminSyncCounsel_(p.token); break;
      case 'tp_admin_counselStudents':r = tpAdminCounselStudents_(p.token); break;
      case 'tp_counselStudentDetail': r = tpCounselStudentDetail_(p.student, p.token); break;
      case 'tp_counselDashboard':    r = tpCounselDashboard_(p.token); break;
      case 'tp_evalTrend':       r = tpEvalTrend_(p.student, p.days||'30', p.token); break;
      case 'tp_evalStats':       r = tpEvalStats_(p.days||'14', p.token); break;
      case 'tp_getMyEvals':      r = tpGetMyEvals_(p.teacherLogId, p.date, p.className, p.token); break;
      default: r = { ok: false, error: 'Unknown action: ' + p.action };
    }
  } catch (err) {
    r = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);
}

function tpHandlePost_(d) {
  let r;
  try {
    switch (d.action) {
      case 'tp_saveEval':           r = tpSaveEval_(d); break;
      case 'tp_admin_saveStudent':  r = tpAdminSaveStudent_(d); break;
      case 'tp_admin_saveTeacher':  r = tpAdminSaveTeacher_(d); break;
      case 'tp_changePassword':     r = tpChangePassword_(d); break;
      case 'tp_admin_saveClass':    r = tpAdminSaveClass_(d); break;
      case 'tp_admin_editEval':     r = tpAdminEditEval_(d); break;
      case 'tp_admin_enrollStudent': r = tpAdminEnrollStudent_(d.student, d.classes, d.token); break;
      case 'tp_admin_saveEnrollment': r = tpAdminSaveEnrollment_(d); break;
      case 'tp_counselSubmit':       r = tpCounselSubmit_(d.formData, d.token); break;
      default: r = { ok: false, error: 'Unknown action: ' + d.action };
    }
  } catch (err) {
    r = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════
// AUTH
// ═══════════════════════════════════════════════════════
function tpLogin_(name, pw) {
  if (!name || !pw) return { ok: false, error: '이름과 비밀번호를 입력해주세요.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_TEACHERS);
  if (!sh) return { ok: false, error: '강사계정 시트가 없습니다. tpSetup()을 먼저 실행하세요.' };
  const data = sh.getDataRange().getValues();
  // 헤더: [강사명, 과목, 비밀번호, 역할, 로그식별자, 상태]
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === name && String(data[i][2]) === String(pw) && data[i][5] !== '비활성') {
      return {
        ok: true,
        teacher: { name: data[i][0], subject: data[i][1], role: data[i][3], logId: data[i][4] },
        token: tpGenToken_(name),
        needPasswordChange: String(pw) === '1234',
      };
    }
  }
  return { ok: false, error: '이름 또는 비밀번호가 올바르지 않습니다.' };
}

// ═══════════════════════════════════════════════════════
// 배치 초기 로딩 — config + todayStatus 한번에
// ═══════════════════════════════════════════════════════
function tpInit_(teacherName, role, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  var config = tpGetConfig_(token);
  var today = tpGetTodayStatus_(teacherName, role, token);
  return { ok: true, config: config.ok ? config : null, todayStatus: today.ok ? today : null };
}

// ═══════════════════════════════════════════════════════
// CONFIG — 평가 입력 옵션 로딩
// ═══════════════════════════════════════════════════════
function tpGetConfig_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_CONFIG);
  if (!sh) return { ok: false, error: 'CONFIG 시트 없음' };
  const data = sh.getDataRange().getValues();
  
  const strip = (s) => String(s).replace(/^\[/, '').replace(/\]$/, ''); // 대괄호 제거
  const scales = [], weaknesses = [], mgmtAreas = [], urgentActions = [], growthPts = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) scales.push(data[i][0]); // 점수 척도는 대괄호 없음
    if (data[i][1]) weaknesses.push(strip(data[i][1]));
    if (data[i][2]) mgmtAreas.push(strip(data[i][2]));
    if (data[i][3]) urgentActions.push(strip(data[i][3]));
    if (data[i][4]) growthPts.push(strip(data[i][4]));
  }
  return { ok: true, scales, weaknesses, mgmtAreas, urgentActions, growthPts };
}

// ═══════════════════════════════════════════════════════
// 담당반 목록
// ═══════════════════════════════════════════════════════
function tpGetClasses_(teacherName, role, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // [DB] 반 읽기
  const classSheet = ss.getSheetByName(TP.SHEET_CLASSES);
  const classData = classSheet.getDataRange().getValues();
  // [DB] 수강 읽기 (학생 수 카운트용)
  const enrollSheet = ss.getSheetByName(TP.SHEET_ENROLLMENT);
  const enrollData = enrollSheet.getDataRange().getValues();
  
  // 학생 수 카운트
  const studentCount = {};
  for (let i = 1; i < enrollData.length; i++) {
    const cn = enrollData[i][1]; // 반 이름
    studentCount[cn] = (studentCount[cn] || 0) + 1;
  }
  
  const classes = [];
  for (let i = 1; i < classData.length; i++) {
    const [className, subject, level, teachers] = classData[i];
    if (!className) continue;
    
    // 관리자는 전체, 강사는 담당 반만
    const isMyClass = role === '관리자' || (teachers && teachers.includes(teacherName));
    if (!isMyClass) continue;
    
    classes.push({
      name: className,
      subject: subject,
      level: level,
      teachers: teachers,
      studentCount: studentCount[className] || 0,
    });
  }
  return { ok: true, classes };
}

// ═══════════════════════════════════════════════════════
// 반별 학생 목록
// ═══════════════════════════════════════════════════════
function tpGetStudents_(className, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // [DB] 수강에서 해당 반 학생 조회
  const enrollSheet = ss.getSheetByName(TP.SHEET_ENROLLMENT);
  const enrollData = enrollSheet.getDataRange().getValues();
  const studentNames = [];
  for (let i = 1; i < enrollData.length; i++) {
    if (enrollData[i][1] === className) {
      studentNames.push({ name: enrollData[i][0], grade: enrollData[i][2] });
    }
  }
  
  // [DB] 학생에서 학교 정보 보강
  const studentSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  const studentMap = {};
  for (let i = 1; i < studentData.length; i++) {
    studentMap[studentData[i][0]] = {
      school: studentData[i][2],
      status: studentData[i][5] || '재원',
    };
  }
  
  const students = studentNames
    .filter(s => {
      const info = studentMap[s.name];
      return !info || info.status !== '퇴원';
    })
    .map(s => ({
      name: s.name,
      grade: s.grade,
      school: (studentMap[s.name] || {}).school || '',
    }))
    .sort((a, b) => a.name.localeCompare(b.name, 'ko'));
  
  // [DB] 반에서 과목 정보
  const classSheet = ss.getSheetByName(TP.SHEET_CLASSES);
  const classData = classSheet.getDataRange().getValues();
  let subject = '';
  for (let i = 1; i < classData.length; i++) {
    if (classData[i][0] === className) { subject = classData[i][1]; break; }
  }
  
  return { ok: true, className, subject, students };
}

// ═══════════════════════════════════════════════════════
// 오늘 평가 현황
// ═══════════════════════════════════════════════════════
function tpGetTodayStatus_(teacherName, role, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy. M. d');
  // 대체 형식도 체크 (스프레드시트 날짜 포맷 다양)
  const todayAlt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy. MM. dd');
  const todayISO = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  
  // 내 담당반 계산
  const classSheet = ss.getSheetByName(TP.SHEET_CLASSES);
  const classData = classSheet.getDataRange().getValues();
  const myClasses = {};
  for (let i = 1; i < classData.length; i++) {
    const [cn, subj, lvl, teachers] = classData[i];
    if (!cn) continue;
    if (role === '관리자' || (teachers && teachers.includes(teacherName))) {
      myClasses[cn] = { subject: subj, total: 0, completed: 0 };
    }
  }
  
  // [DB] 수강에서 학생 수
  const enrollSheet = ss.getSheetByName(TP.SHEET_ENROLLMENT);
  const enrollData = enrollSheet.getDataRange().getValues();
  // 퇴원 학생 제외를 위한 학생 시트 로딩
  const studentSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  const quitSet = new Set();
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][5] === '퇴원') quitSet.add(studentData[i][0]);
  }
  
  let totalStudents = 0;
  for (let i = 1; i < enrollData.length; i++) {
    const cn = enrollData[i][1];
    const sn = enrollData[i][0];
    if (myClasses[cn] && !quitSet.has(sn)) {
      myClasses[cn].total++;
      totalStudents++;
    }
  }
  
  // [ALL] + [포털] 두 탭 모두에서 오늘 입력 확인 (★ #3 핵심)
  const completedSet = new Set();
  let completedStudents = 0;
  
  // 함수: 한 시트의 최근 행만 스캔 (전체 10000+ 읽기 방지)
  function scanLogSheet_(sheetName) {
    const logSheet = ss.getSheetByName(sheetName);
    if (!logSheet || logSheet.getLastRow() < 2) return;
    const lastRow = logSheet.getLastRow();
    // 오늘 데이터는 최근에 추가되었으므로 마지막 500행만 스캔
    const startRow = Math.max(2, lastRow - 499);
    const numRows = lastRow - startRow + 1;
    const logData = logSheet.getRange(startRow, 1, numRows, 13).getValues();
    for (let i = logData.length - 1; i >= 0; i--) {
      const rowDate = tpNormalizeDate_(logData[i][1]);
      if (rowDate !== todayISO) continue;
      const cn = logData[i][2];
      const sn = logData[i][3];
      if (myClasses[cn]) {
        const key = cn + '|' + sn;
        if (!completedSet.has(key)) {
          completedSet.add(key);
          myClasses[cn].completed++;
          completedStudents++;
        }
      }
    }
  }
  
  scanLogSheet_(TP.SHEET_DAILY_LOG);   // [ALL] 전체 로그 취합 (포털 데이터 포함)
  
  return {
    ok: true,
    date: todayISO,
    totalStudents,
    completedStudents,
    classes: Object.entries(myClasses).map(([name, info]) => ({
      name, ...info,
    })),
  };
}

// ═══════════════════════════════════════════════════════
// 일일 평가 저장
// ═══════════════════════════════════════════════════════
function tpSaveEval_(postData) {
  const { evaluations, teacherLogId, evalDate, token } = postData;
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  if (!evaluations || !evaluations.length) return { ok: false, error: '저장할 데이터 없음' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { ok: false, error: '서버가 바쁩니다. 잠시 후 다시 시도해주세요.' };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TP.SHEET_PORTAL_LOG);
    if (!sheet) return { ok: false, error: '[포털] 일일로그 탭이 없습니다.' };
    
    // evalDate가 있으면 해당 날짜 사용, 없으면 오늘
    var targetDate, targetISO;
    if (evalDate) {
      // evalDate: "2026-03-20" 형식
      var parts = evalDate.split('-');
      targetDate = parts[0] + '. ' + parseInt(parts[1]) + '. ' + parseInt(parts[2]);
      targetISO = evalDate;
    } else {
      var now = new Date();
      targetDate = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy. M. d');
      targetISO = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd');
    }
    
    // 해당 날짜+같은 교사+같은 학생 행 찾기 (수정용 — 10001행 이후만)
    const data = sheet.getDataRange().getValues();
    const existingRows = {};
    const startIdx = Math.max(1, TP.PORTAL_START_ROW - 1); // 0-based index
    for (let i = data.length - 1; i >= startIdx; i--) {
      const rowDate = tpNormalizeDate_(data[i][1]);
      if (rowDate !== targetISO) continue;
      if (data[i][0] !== teacherLogId) continue;
      const key = data[i][2] + '|' + data[i][3];
      if (!existingRows[key]) existingRows[key] = i + 1;
    }
    
    const toAppend = [];
    const toUpdate = [];
    
    for (const ev of evaluations) {
      const isAbsent = ev.isAbsent === true;
      const row = [
        teacherLogId,
        targetDate,
        ev.className,
        ev.studentName,
        isAbsent ? '' : (ev.homework || ''),
        isAbsent ? '' : (ev.understanding || ''),
        isAbsent ? '' : (ev.focus || ''),
        isAbsent ? '' : (ev.persistence || ''),
        isAbsent ? '' : (ev.growthPoints || ''),
        isAbsent ? '' : (ev.weaknesses || ''),
        isAbsent ? '결석' : (ev.mgmtAreas || ''),
        isAbsent ? '' : (ev.urgentAction || ''),
        isAbsent ? '결석' : (ev.comment || ''),
      ];
      
      const key = ev.className + '|' + ev.studentName;
      if (existingRows[key]) {
        toUpdate.push({ row: existingRows[key], data: row });
      } else {
        toAppend.push(row);
      }
    }
    
    for (const u of toUpdate) {
      sheet.getRange(u.row, 1, 1, u.data.length).setValues([u.data]);
    }
    if (toAppend.length > 0) {
      var appendStart = Math.max(sheet.getLastRow() + 1, TP.PORTAL_START_ROW);
      sheet.getRange(appendStart, 1, toAppend.length, toAppend[0].length)
        .setValues(toAppend);
    }
    
    return {
      ok: true,
      message: evaluations.length + '명 평가 저장 완료' + (toUpdate.length ? ' (수정 ' + toUpdate.length + '건)' : ''),
      updated: toUpdate.length,
      appended: toAppend.length,
    };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════
// 교사 본인의 기존 평가 조회 (날짜 + 반 기준)
// ═══════════════════════════════════════════════════════
function tpGetMyEvals_(teacherLogId, dateStr, className, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  if (!teacherLogId || !className) return { ok: false, error: '필수 파라미터 누락' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetDate = dateStr || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var evals = {};
  
  function scan(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== teacherLogId) continue;
      var d = tpNormalizeDate_(data[i][1]);
      if (d !== targetDate) continue;
      if (String(data[i][2]).trim() !== className) continue;
      
      var studentName = String(data[i][3]).trim();
      evals[studentName] = {
        homework: String(data[i][4] || ''),
        understanding: String(data[i][5] || ''),
        focus: String(data[i][6] || ''),
        persistence: String(data[i][7] || ''),
        growthPoints: String(data[i][8] || ''),
        weaknesses: String(data[i][9] || ''),
        mgmtAreas: String(data[i][10] || ''),
        urgentAction: String(data[i][11] || ''),
        comment: String(data[i][12] || ''),
        isAbsent: String(data[i][10]).trim() === '결석' && !data[i][4],
      };
    }
  }
  
  scan(TP.SHEET_DAILY_LOG);  // [ALL] 전체 로그 취합 (포털 데이터 포함)
  
  return { ok: true, date: targetDate, className: className, evals: evals };
}

// ═══════════════════════════════════════════════════════
// 담당 학생 조회 — 강사: 담당 반 학생만 / 관리자: 전체
// ═══════════════════════════════════════════════════════
function tpGetMyStudents_(teacherName, role, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 관리자면 전체 재원 학생
  if (role === '관리자') {
    var stSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
    if (!stSheet) return { ok: true, students: [] };
    var stData = stSheet.getDataRange().getValues();
    var all = [];
    for (var i = 1; i < stData.length; i++) {
      if (!stData[i][0]) continue;
      var st = stData[i][5] || '재원';
      if (st === '퇴원') continue;
      all.push({ name: String(stData[i][0]).trim(), grade: String(stData[i][1]||''), school: String(stData[i][2]||'') });
    }
    return { ok: true, students: all };
  }
  
  // 강사: 담당 반 찾기
  var classSheet = ss.getSheetByName(TP.SHEET_CLASSES);
  if (!classSheet) return { ok: true, students: [] };
  var clData = classSheet.getDataRange().getValues();
  var myClasses = [];
  for (var i = 1; i < clData.length; i++) {
    var teachers = String(clData[i][3] || '');
    if (teachers.indexOf(teacherName) >= 0) {
      myClasses.push(String(clData[i][0]).trim());
    }
  }
  
  if (!myClasses.length) return { ok: true, students: [] };
  
  // 담당 반의 학생 찾기
  var enrollSheet = ss.getSheetByName(TP.SHEET_ENROLLMENT);
  if (!enrollSheet) return { ok: true, students: [] };
  var enData = enrollSheet.getDataRange().getValues();
  var nameSet = new Set();
  for (var i = 1; i < enData.length; i++) {
    if (myClasses.indexOf(String(enData[i][1]).trim()) >= 0) {
      nameSet.add(String(enData[i][0]).trim());
    }
  }
  
  // 학생 정보 조회
  var stSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
  var students = [];
  if (stSheet) {
    var stData = stSheet.getDataRange().getValues();
    for (var i = 1; i < stData.length; i++) {
      var nm = String(stData[i][0]).trim();
      if (!nameSet.has(nm)) continue;
      var st = stData[i][5] || '재원';
      if (st === '퇴원') continue;
      students.push({ name: nm, grade: String(stData[i][1]||''), school: String(stData[i][2]||'') });
    }
  }
  
  return { ok: true, students: students, classes: myClasses };
}

// ═══════════════════════════════════════════════════════
// 학생 프로필 (성향 데이터)
// ═══════════════════════════════════════════════════════
function tpGetStudentProfile_(studentName, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const studentSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
  const studentData = studentSheet.getDataRange().getValues();
  let studentInfo = null;
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][0] === studentName) {
      studentInfo = {
        name: studentData[i][0], grade: studentData[i][1],
        school: studentData[i][2], parentPhone: studentData[i][3],
        studentPhone: studentData[i][4], status: studentData[i][5] || '재원',
      };
      break;
    }
  }
  if (!studentInfo) return { ok: false, error: '학생을 찾을 수 없습니다.' };
  
  // [DB] 수강
  const enrollSheet = ss.getSheetByName(TP.SHEET_ENROLLMENT);
  const enrollData = enrollSheet.getDataRange().getValues();
  const classes = [];
  for (let i = 1; i < enrollData.length; i++) {
    if (enrollData[i][0] === studentName) classes.push(enrollData[i][1]);
  }
  
  // [DB] 성향분석 — 최신 1건 (★ 전체 JSON 반환)
  const pSheet = ss.getSheetByName(TP.SHEET_PERSONALITY);
  if (pSheet) {
    const pData = pSheet.getDataRange().getValues();
    // 마지막 행이 최신이므로 역순 탐색
    for (let i = pData.length - 1; i >= 1; i--) {
      if (pData[i][0] === studentName) {
        studentInfo.personalityDate = pData[i][1] || '';
        studentInfo.reliability = pData[i][2] || '';
        // D열: 전체 JSON을 문자열 그대로 반환
        studentInfo.personalityJson = pData[i][3] ? String(pData[i][3]) : '';
        studentInfo.levelTestEnglish = pData[i][4] ? String(pData[i][4]) : '';
        studentInfo.levelTestMath = pData[i][5] ? String(pData[i][5]) : '';
        break;
      }
    }
  }
  
  studentInfo.classes = classes;
  return { ok: true, student: studentInfo };
}

// ═══════════════════════════════════════════════════════
// 학생 2주 리포트 목록
// ═══════════════════════════════════════════════════════
function tpGetStudentReports_(studentName, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  var ss3;
  try {
    ss3 = SpreadsheetApp.openById(TP.CONSULT_SYS_ID);
  } catch(e) {
    return { ok: false, error: 'Sheet 3 접근 불가: ' + e.message };
  }
  var sh = ss3.getSheetByName(TP.SHEET_CONSULT_LOG);
  if (!sh) return { ok: true, reports: [] };
  
  var data = sh.getDataRange().getValues();
  var reports = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(studentName).trim()) {
      // 브리핑에서 핵심만 추출 (전체 JSON은 너무 큼)
      var briefSummary = '';
      try {
        var bj = JSON.parse(data[i][4]);
        briefSummary = bj.aiCounselingAgenda ? bj.aiCounselingAgenda.mainTheme || '' : '';
      } catch(e) { briefSummary = String(data[i][4] || '').substring(0, 200); }
      
      // AI 분석에서 핵심만 추출
      var aiSummary = '';
      var metrics = null;
      var counselResult = null;
      var actionPlan = null;
      try {
        var aj = JSON.parse(data[i][6]);
        if (aj.performanceAnalysis) {
          aiSummary = aj.performanceAnalysis.aiGeneralAssessment || '';
          metrics = aj.performanceAnalysis.quantitativeMetrics || null;
        }
        if (aj.counselingResult) counselResult = aj.counselingResult;
        if (aj.actionPlanForNextTwoWeeks) actionPlan = aj.actionPlanForNextTwoWeeks;
      } catch(e) { aiSummary = String(data[i][6] || '').substring(0, 200); }
      
      reports.push({
        id: data[i][0],
        date: tpNormalizeDate_(data[i][2]),
        counselor: data[i][3],
        briefSummary: briefSummary,
        checklist: String(data[i][5] || ''),
        aiSummary: aiSummary,
        metrics: metrics,
        counselResult: counselResult,
        actionPlan: actionPlan,
        goals: String(data[i][7] || ''),
        reportUrl: data[i][9] || '',
        shortUrl: data[i][10] || '',
      });
    }
  }
  return { ok: true, reports: reports.reverse() };
}

// ═══════════════════════════════════════════════════════
// 관리자: 전체 현황
// ═══════════════════════════════════════════════════════
function tpAdminOverview_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todayISO = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  
  // 강사 목록
  const tSheet = ss.getSheetByName(TP.SHEET_TEACHERS);
  const tData = tSheet.getDataRange().getValues();
  const teachers = [];
  for (let i = 1; i < tData.length; i++) {
    if (tData[i][5] !== '비활성') {
      teachers.push({ name: tData[i][0], subject: tData[i][1], logId: tData[i][4], role: tData[i][3] });
    }
  }
  
  // 오늘 로그 집계
  const logSheet = ss.getSheetByName(TP.SHEET_DAILY_LOG);
  const logData = logSheet.getDataRange().getValues();
  const todayByTeacher = {};
  for (let i = logData.length - 1; i >= 1; i--) {
    const rowDate = tpNormalizeDate_(logData[i][1]);
    if (rowDate !== todayISO) continue;
    const t = logData[i][0]; // 담당교사
    todayByTeacher[t] = (todayByTeacher[t] || 0) + 1;
  }
  
  // 학생/반 통계
  const studentSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
  const sData = studentSheet.getDataRange().getValues();
  let activeStudents = 0;
  for (let i = 1; i < sData.length; i++) {
    if (sData[i][5] !== '퇴원') activeStudents++;
  }
  
  return {
    ok: true,
    date: todayISO,
    activeStudents,
    totalLogs: logData.length - 1,
    todayByTeacher,
    teachers,
  };
}

// ═══════════════════════════════════════════════════════
// 관리자: 일일 평가 로그 조회
// ═══════════════════════════════════════════════════════
function tpAdminEvalLog_(token, filterDate, filterTeacher, filterClass) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetDate = filterDate || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const logs = [];
  
  // ★ 두 탭 모두 스캔
  function scanSheet_(sheetName, source) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1 && logs.length < 300; i--) {
      const rowDate = tpNormalizeDate_(data[i][1]);
      if (filterDate && rowDate !== targetDate) continue;
      if (filterTeacher && data[i][0] !== filterTeacher) continue;
      if (filterClass && data[i][2] !== filterClass) continue;
      logs.push({
        source, rowIndex: i + 1, sheetName,
        teacher: data[i][0], date: rowDate, className: data[i][2],
        student: data[i][3], homework: data[i][4], understanding: data[i][5],
        focus: data[i][6], persistence: data[i][7],
        growthPoints: data[i][8], weaknesses: data[i][9],
        mgmtAreas: data[i][10], urgentAction: data[i][11], comment: data[i][12],
      });
    }
  }
  scanSheet_(TP.SHEET_DAILY_LOG, 'all');
  
  // 날짜 내림차순 정렬
  logs.sort((a, b) => b.date.localeCompare(a.date));
  return { ok: true, logs };
}

// ═══════════════════════════════════════════════════════
// 관리자: 전체 학생 목록
// ═══════════════════════════════════════════════════════
function tpAdminAllStudents_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_STUDENTS);
  const data = sh.getDataRange().getValues();
  
  // 수강 정보 로딩
  const enrollSh = ss.getSheetByName(TP.SHEET_ENROLLMENT);
  const enrollData = enrollSh.getDataRange().getValues();
  const classMap = {};
  for (let i = 1; i < enrollData.length; i++) {
    const sn = enrollData[i][0];
    if (!classMap[sn]) classMap[sn] = [];
    classMap[sn].push(enrollData[i][1]);
  }
  
  const students = [];
  for (let i = 1; i < data.length; i++) {
    students.push({
      name: data[i][0], grade: data[i][1], school: data[i][2],
      parentPhone: data[i][3], studentPhone: data[i][4],
      status: data[i][5] || '재원',
      classes: classMap[data[i][0]] || [],
    });
  }
  return { ok: true, students };
}

// ═══════════════════════════════════════════════════════
// 관리자: 반 목록
// ═══════════════════════════════════════════════════════
function tpAdminAllClasses_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_CLASSES);
  const data = sh.getDataRange().getValues();
  
  const enrollSh = ss.getSheetByName(TP.SHEET_ENROLLMENT);
  const enrollData = enrollSh.getDataRange().getValues();
  const countMap = {};
  for (let i = 1; i < enrollData.length; i++) {
    const cn = enrollData[i][1];
    countMap[cn] = (countMap[cn] || 0) + 1;
  }
  
  const classes = [];
  for (let i = 1; i < data.length; i++) {
    classes.push({
      name: data[i][0], subject: data[i][1], level: data[i][2],
      teachers: data[i][3], studentCount: countMap[data[i][0]] || 0,
    });
  }
  return { ok: true, classes };
}

// ═══════════════════════════════════════════════════════
// 관리자: 강사 목록
// ═══════════════════════════════════════════════════════
function tpAdminTeachers_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_TEACHERS);
  const data = sh.getDataRange().getValues();
  const teachers = [];
  for (let i = 1; i < data.length; i++) {
    teachers.push({
      name: data[i][0], subject: data[i][1], role: data[i][3],
      logId: data[i][4], status: data[i][5],
    });
  }
  return { ok: true, teachers };
}

// ═══════════════════════════════════════════════════════
// 관리자: 신규 상담 목록 (별도 시트)
// ═══════════════════════════════════════════════════════
function tpAdminConsultList_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    const extSS = SpreadsheetApp.openById(TP.CONSULT_SHEET_ID);
    const sheets = extSS.getSheets();
    // 첫 번째 시트에서 상담 목록 읽기 (구조는 counsel 시스템에 따라 다름)
    const sh = sheets[0];
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const consults = [];
    for (let i = 1; i < Math.min(data.length, 51); i++) { // 최근 50건
      const row = {};
      headers.forEach((h, j) => { row[h] = data[i][j]; });
      consults.push(row);
    }
    return { ok: true, consults, headers };
  } catch (err) {
    return { ok: false, error: '신규 상담 시트 접근 불가: ' + err.message };
  }
}

// ═══════════════════════════════════════════════════════
// 관리자: 학생 저장 (추가/수정)
// ═══════════════════════════════════════════════════════
function tpAdminSaveStudent_(d) {
  if (!tpValidToken_(d.token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_STUDENTS);
  const data = sh.getDataRange().getValues();
  const s = d.student;
  
  // 기존 학생 찾기
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === s.name) {
      // 업데이트
      sh.getRange(i + 1, 1, 1, 6).setValues([[
        s.name, s.grade, s.school, s.parentPhone, s.studentPhone, s.status
      ]]);
      return { ok: true, message: s.name + ' 학생 정보 수정 완료' };
    }
  }
  
  // 신규 추가
  sh.appendRow([s.name, s.grade, s.school, s.parentPhone, s.studentPhone, s.status || '재원']);
  return { ok: true, message: s.name + ' 학생 추가 완료' };
}

// ═══════════════════════════════════════════════════════
// 관리자: 강사 저장
// ═══════════════════════════════════════════════════════
function tpAdminSaveTeacher_(d) {
  if (!tpValidToken_(d.token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_TEACHERS);
  const data = sh.getDataRange().getValues();
  const t = d.teacher;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === t.name) {
      sh.getRange(i + 1, 1, 1, 6).setValues([[
        t.name, t.subject, t.password || data[i][2], t.role, t.logId, t.status
      ]]);
      return { ok: true, message: t.name + ' 강사 정보 수정 완료' };
    }
  }
  sh.appendRow([t.name, t.subject, t.password, t.role, t.logId, t.status || '활성']);
  return { ok: true, message: t.name + ' 강사 추가 완료' };
}

// ═══════════════════════════════════════════════════════
// 학생별 평가 추이 (최근 N일)
// ═══════════════════════════════════════════════════════
function tpEvalTrend_(studentName, days, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  if (!studentName) return { ok: false, error: '학생명 필수' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daysNum = parseInt(days) || 30;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysNum);
  var cutoffISO = Utilities.formatDate(cutoff, 'Asia/Seoul', 'yyyy-MM-dd');
  
  // 점수 텍스트에서 숫자 추출 ("5점 (완벽)" → 5)
  function extractScore(val) {
    if (!val) return null;
    var m = String(val).match(/(\d)점/);
    return m ? parseInt(m[1]) : null;
  }
  
  var entries = [];
  
  // 두 탭 모두 스캔
  function scan(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() !== studentName) continue;
      var d = tpNormalizeDate_(data[i][1]);
      if (d < cutoffISO) continue;
      var hw = extractScore(data[i][4]);
      var ud = extractScore(data[i][5]);
      var fc = extractScore(data[i][6]);
      var ps = extractScore(data[i][7]);
      if (hw === null && ud === null) continue; // 결석 등 빈 데이터 제외
      entries.push({
        date: d,
        teacher: String(data[i][0]),
        className: String(data[i][2]),
        homework: hw, understanding: ud, focus: fc, persistence: ps,
        growthPoints: String(data[i][8] || ''),
        weaknesses: String(data[i][9] || ''),
      });
    }
  }
  scan(TP.SHEET_DAILY_LOG);
  
  // 날짜 오름차순 정렬
  entries.sort(function(a, b) { return a.date.localeCompare(b.date); });
  
  // 날짜별 평균 계산 (같은 날 여러 수업이 있을 수 있으므로)
  var byDate = {};
  entries.forEach(function(e) {
    if (!byDate[e.date]) byDate[e.date] = { hw: [], ud: [], fc: [], ps: [], classes: [] };
    if (e.homework !== null) byDate[e.date].hw.push(e.homework);
    if (e.understanding !== null) byDate[e.date].ud.push(e.understanding);
    if (e.focus !== null) byDate[e.date].fc.push(e.focus);
    if (e.persistence !== null) byDate[e.date].ps.push(e.persistence);
    if (byDate[e.date].classes.indexOf(e.className) < 0) byDate[e.date].classes.push(e.className);
  });
  
  function avg(arr) { return arr.length ? Math.round(arr.reduce(function(a,b){return a+b;},0) / arr.length * 10) / 10 : null; }
  
  var trend = Object.keys(byDate).sort().map(function(d) {
    var v = byDate[d];
    return { date: d, homework: avg(v.hw), understanding: avg(v.ud), focus: avg(v.fc), persistence: avg(v.ps), classes: v.classes };
  });
  
  // 전체 평균
  var allHw = [], allUd = [], allFc = [], allPs = [];
  entries.forEach(function(e) {
    if (e.homework !== null) allHw.push(e.homework);
    if (e.understanding !== null) allUd.push(e.understanding);
    if (e.focus !== null) allFc.push(e.focus);
    if (e.persistence !== null) allPs.push(e.persistence);
  });
  
  // 자주 나오는 약점/성장포인트 집계
  var wpCount = {}, gpCount = {};
  entries.forEach(function(e) {
    if (e.weaknesses) e.weaknesses.split(',').forEach(function(w) { w = w.trim(); if (w) wpCount[w] = (wpCount[w] || 0) + 1; });
    if (e.growthPoints) e.growthPoints.split(',').forEach(function(g) { g = g.trim(); if (g) gpCount[g] = (gpCount[g] || 0) + 1; });
  });
  
  return {
    ok: true,
    student: studentName,
    period: daysNum + '일',
    totalEntries: entries.length,
    averages: { homework: avg(allHw), understanding: avg(allUd), focus: avg(allFc), persistence: avg(allPs) },
    trend: trend,
    topWeaknesses: Object.keys(wpCount).sort(function(a,b){return wpCount[b]-wpCount[a];}).slice(0, 5).map(function(k){return {name:k,count:wpCount[k]};}),
    topGrowthPoints: Object.keys(gpCount).sort(function(a,b){return gpCount[b]-gpCount[a];}).slice(0, 5).map(function(k){return {name:k,count:gpCount[k]};}),
  };
}

// ═══════════════════════════════════════════════════════
// 강사별 평가 통계 (최근 N일)
// ═══════════════════════════════════════════════════════
function tpEvalStats_(days, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daysNum = parseInt(days) || 14;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysNum);
  var cutoffISO = Utilities.formatDate(cutoff, 'Asia/Seoul', 'yyyy-MM-dd');
  
  function extractScore(val) {
    if (!val) return null;
    var m = String(val).match(/(\d)점/);
    return m ? parseInt(m[1]) : null;
  }
  
  // 강사별 집계
  var byTeacher = {};
  var byDate = {}; // 날짜별 전체 입력수
  
  function scan(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var d = tpNormalizeDate_(data[i][1]);
      if (d < cutoffISO) continue;
      var teacher = String(data[i][0]).trim();
      if (!teacher) continue;
      
      if (!byTeacher[teacher]) byTeacher[teacher] = { count: 0, hw: [], ud: [], fc: [], ps: [], dates: {} };
      byTeacher[teacher].count++;
      var hw = extractScore(data[i][4]);
      var ud = extractScore(data[i][5]);
      var fc = extractScore(data[i][6]);
      var ps = extractScore(data[i][7]);
      if (hw !== null) byTeacher[teacher].hw.push(hw);
      if (ud !== null) byTeacher[teacher].ud.push(ud);
      if (fc !== null) byTeacher[teacher].fc.push(fc);
      if (ps !== null) byTeacher[teacher].ps.push(ps);
      byTeacher[teacher].dates[d] = (byTeacher[teacher].dates[d] || 0) + 1;
      
      byDate[d] = (byDate[d] || 0) + 1;
    }
  }
  scan(TP.SHEET_DAILY_LOG);
  
  function avg(arr) { return arr.length ? Math.round(arr.reduce(function(a,b){return a+b;},0) / arr.length * 10) / 10 : null; }
  
  var teachers = Object.keys(byTeacher).map(function(t) {
    var v = byTeacher[t];
    return {
      teacher: t,
      totalEntries: v.count,
      activeDays: Object.keys(v.dates).length,
      avgHomework: avg(v.hw),
      avgUnderstanding: avg(v.ud),
      avgFocus: avg(v.fc),
      avgPersistence: avg(v.ps),
    };
  }).sort(function(a, b) { return b.totalEntries - a.totalEntries; });
  
  // 일별 전체 입력 추이
  var dailyTrend = Object.keys(byDate).sort().map(function(d) { return { date: d, count: byDate[d] }; });
  
  return {
    ok: true,
    period: daysNum + '일',
    teachers: teachers,
    dailyTrend: dailyTrend,
    totalEntries: teachers.reduce(function(s, t) { return s + t.totalEntries; }, 0),
  };
}

// ═══════════════════════════════════════════════════════
// 관리자: 수강 관리 (반 배정 추가/삭제)
// ═══════════════════════════════════════════════════════
function tpAdminSaveEnrollment_(d) {
  if (!tpValidToken_(d.token)) return { ok: false, error: '인증 만료' };
  var action = d.enrollAction; // 'add' or 'remove'
  var studentName = d.student;
  var className = d.className;
  
  if (!studentName || !className) return { ok: false, error: '학생명과 반이름 필수' };
  
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e) { return { ok: false, error: '서버 바쁨' }; }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(TP.SHEET_ENROLLMENT);
    if (!sh) return { ok: false, error: '[DB] 수강 탭 없음' };
    
    if (action === 'add') {
      // 중복 확인
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === studentName && String(data[i][1]).trim() === className) {
          return { ok: true, message: studentName + '은 이미 ' + className + '에 등록되어 있습니다.' };
        }
      }
      // 학생의 학년 조회
      var stSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
      var grade = '';
      if (stSheet) {
        var stData = stSheet.getDataRange().getValues();
        for (var i = 1; i < stData.length; i++) {
          if (String(stData[i][0]).trim() === studentName) { grade = stData[i][1]; break; }
        }
      }
      sh.appendRow([studentName, className, grade]);
      return { ok: true, message: studentName + ' → ' + className + ' 배정 완료' };
      
    } else if (action === 'remove') {
      var data = sh.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).trim() === studentName && String(data[i][1]).trim() === className) {
          sh.deleteRow(i + 1);
          return { ok: true, message: studentName + ' → ' + className + ' 배정 해제' };
        }
      }
      return { ok: false, error: '해당 수강 정보를 찾을 수 없습니다.' };
    }
    
    return { ok: false, error: '알 수 없는 action: ' + action };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════
// 관리자: 반 추가/수정 (#4)
// ═══════════════════════════════════════════════════════
function tpAdminSaveClass_(d) {
  if (!tpValidToken_(d.token)) return { ok: false, error: '인증 만료' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TP.SHEET_CLASSES);
  const data = sh.getDataRange().getValues();
  const c = d.classInfo;
  if (!c || !c.name) return { ok: false, error: '반 이름 필수' };
  
  // 기존 반 수정
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === c.originalName || data[i][0] === c.name) {
      sh.getRange(i + 1, 1, 1, 4).setValues([[c.name, c.subject, c.level || '', c.teachers || '']]);
      return { ok: true, message: c.name + ' 반 수정 완료' };
    }
  }
  // 신규 반 추가
  sh.appendRow([c.name, c.subject, c.level || '', c.teachers || '']);
  return { ok: true, message: c.name + ' 반 추가 완료' };
}

// ═══════════════════════════════════════════════════════
// 관리자: 일일 평가 수정 (#4)
// ═══════════════════════════════════════════════════════
function tpAdminEditEval_(d) {
  if (!tpValidToken_(d.token)) return { ok: false, error: '인증 만료' };
  const { sheetName, rowIndex, evalData } = d;
  if (!sheetName || !rowIndex || !evalData) return { ok: false, error: '필수 데이터 없음' };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { ok: false, error: '서버 바쁨. 재시도 해주세요.' };
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: '시트 없음: ' + sheetName };
    
    const row = [
      evalData.teacher, evalData.date, evalData.className, evalData.student,
      evalData.homework, evalData.understanding, evalData.focus, evalData.persistence,
      evalData.growthPoints || '', evalData.weaknesses || '', evalData.mgmtAreas || '',
      evalData.urgentAction || '', evalData.comment || '',
    ];
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    return { ok: true, message: '평가 수정 완료' };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════
// 비밀번호 변경
// ═══════════════════════════════════════════════════════
function tpChangePassword_(d) {
  if (!tpValidToken_(d.token)) return { ok: false, error: '인증 만료' };
  var name = d.name;
  var currentPw = d.currentPassword;
  var newPw = d.newPassword;
  
  if (!name || !currentPw || !newPw) return { ok: false, error: '필수 항목 누락' };
  if (newPw.length < 4) return { ok: false, error: '비밀번호는 4자 이상이어야 합니다.' };
  if (newPw === '1234') return { ok: false, error: '초기 비밀번호와 다른 비밀번호를 설정하세요.' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TP.SHEET_TEACHERS);
  var data = sh.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) {
      if (String(data[i][2]) !== String(currentPw)) {
        return { ok: false, error: '현재 비밀번호가 일치하지 않습니다.' };
      }
      sh.getRange(i + 1, 3).setValue(newPw);
      return { ok: true, message: '비밀번호가 변경되었습니다.' };
    }
  }
  return { ok: false, error: '계정을 찾을 수 없습니다.' };
}

// ═══════════════════════════════════════════════════════
// 유틸리티
function tpGenToken_(name) {
  const ts = new Date().getTime();
  const payload = name + ':' + ts + ':' + TP.TOKEN_SECRET;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload, Utilities.Charset.UTF_8);
  const hStr = hash.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').substring(0, 16);
  return Utilities.base64Encode(name + ':' + ts + ':' + hStr);
}

function tpValidToken_(token) {
  if (!token) return false;
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split(':');
    if (parts.length < 3) return false;
    const ts = parseInt(parts[1]);
    return (new Date().getTime() - ts) / 3600000 < TP.TOKEN_EXPIRY_HOURS;
  } catch (e) { return false; }
}

/** 다양한 날짜 형식을 YYYY-MM-DD로 정규화 */
function tpNormalizeDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  // "2026. 3. 20" → "2026-03-20"
  const m = s.match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (m) {
    return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
  }
  // "2026-03-20" 그대로
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

// ═══════════════════════════════════════════════════════
// 초기 설정 (1회 실행) — [MASTER] 강사계정 시트 생성
// ═══════════════════════════════════════════════════════
function tpSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ── 1. [MASTER] 강사계정 시트 ──
  let sh = ss.getSheetByName(TP.SHEET_TEACHERS);
  if (!sh) {
    sh = ss.insertSheet(TP.SHEET_TEACHERS);
    sh.getRange('A1:F1').setValues([['강사명', '과목', '비밀번호', '역할', '로그식별자', '상태']]);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
    const teachers = [
      ['김윤재', '국어', '1234', '관리자', '국어_김윤재', '활성'],
      ['김남선', '수학', '1234', '강사', '수학_김남선', '활성'],
      ['남제식', '수학', '1234', '강사', '수학_남제식', '활성'],
      ['강민수', '수학', '1234', '강사', '수학_강민수', '활성'],
      ['노영훈', '수학', '1234', '강사', '수학_노영훈', '활성'],
      ['김학균', '영어', '1234', '강사', '영어_김학균', '활성'],
      ['김진송', '영어', '1234', '강사', '영어_김진송', '활성'],
      ['박건영', '영어', '1234', '강사', '영어_박건영', '활성'],
      ['김현수', '영어', '1234', '강사', '영어_김현수', '활성'],
      ['권영준', '영어', '1234', '강사', '영어_권영준', '활성'],
    ];
    sh.getRange(2, 1, teachers.length, 6).setValues(teachers);
    sh.setColumnWidths(1, 6, 120);
    sh.setFrozenRows(1);
    Logger.log('✅ [MASTER] 강사계정 시트 생성 완료 (10명, 초기 비밀번호: 1234)');
  } else {
    Logger.log('[MASTER] 강사계정 이미 존재');
  }

  // ── 2. [포털] 일일로그 시트 (★ 핵심: 포털 전용 쓰기 탭) ──
  let psh = ss.getSheetByName(TP.SHEET_PORTAL_LOG);
  if (!psh) {
    psh = ss.insertSheet(TP.SHEET_PORTAL_LOG);
    psh.getRange('A1:M1').setValues([['담당교사', '날짜', '반 이름', '학생명', '과제 완료도', '수업 이해도', '수업 집중도', '학습 끈기', '오늘의 성장 포인트', '주요 학업 약점', '주요 관리 영역', '교사 긴급 조치', '종합코멘트']]);
    psh.getRange('A1:M1').setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
    psh.setFrozenRows(1);
    Logger.log('✅ [포털] 일일로그 시트 생성 완료');
    Logger.log('⚠️ [ALL] 전체 로그 취합의 QUERY에 이 탭도 포함시켜야 합니다.');
    Logger.log('   또는 Sheet 3에서 이 탭을 별도 IMPORTRANGE로 읽으세요.');
  } else {
    Logger.log('[포털] 일일로그 이미 존재');
  }
  
  Logger.log('');
  Logger.log('⚠️ 배포 전 반드시 비밀번호를 변경하세요!');
}
