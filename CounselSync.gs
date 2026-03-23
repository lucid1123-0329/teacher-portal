// ═══════════════════════════════════════════════════════════════
// CounselSync.gs — counsel 시트 → 메인 DB 성향분석 자동 동기화
// 
// 설치 위치: 메인 데이터베이스(14R-N4QZ...)의 Apps Script
// 역할: counsel 시트(1C3h...)의 [DB] 학생 마스터에서
//       성향검사/레벨테스트 결과를 메인 DB의 [DB] 성향분석에 동기화
//
// 실행 방법:
//   1. 수동: 포털 관리자 메뉴에서 "동기화" 버튼
//   2. 자동: setupSyncTrigger() 1회 실행 → 매일 오전 6시 자동 실행
// ═══════════════════════════════════════════════════════════════

const SYNC = {
  // counsel 시트 (신규 상담 시스템)
  COUNSEL_SHEET_ID: '1C3hlNGzE9qoBYfcHfSCMRQEjyz7aFyofj8TKgrz-eCc',
  COUNSEL_TAB: '[DB] 학생 마스터',
  
  // counsel [DB] 학생 마스터 컬럼 인덱스 (0-based)
  C_NAME: 2,           // C열: 학생 이름
  C_PHONE: 3,          // D열: 학생 연락처
  C_PARENT_PHONE: 4,   // E열: 학부모 연락처
  C_SCHOOL: 6,         // G열: 학교
  C_GRADE: 7,          // H열: 학년
  C_SURVEY_DATE: 20,   // U열: 성향검사 실시일
  C_RELIABILITY: 21,   // V열: 신뢰도 (상/중/하)
  C_SURVEY_JSON: 22,   // W열: Gemini 성향분석 전체 JSON
  C_LT_ENG_JSON: 25,   // Z열: 영어 레벨테스트 JSON
  C_LT_MATH_JSON: 28,  // AC열: 수학 레벨테스트 JSON
  C_STATUS: 32,        // AG열: 파이프라인 상태
  
  // 메인 DB
  MAIN_TAB: '[DB] 성향분석',
  MAIN_STUDENTS: '[DB] 학생',
  // [DB] 학생 컬럼: A=학생명, B=학년, C=학교, D=학부모연락처, E=학생연락처, F=재원여부
  // [DB] 성향분석 컬럼: A=학생명, B=검사일, C=신뢰도, D=분석결과JSON, E=레벨(영어), F=레벨(수학)
};


/**
 * ── 핵심 동기화 함수 ──
 * counsel [DB] 학생 마스터 → 메인 DB [DB] 성향분석
 * 
 * 동작:
 *  1. counsel 시트에서 성향검사 완료된 학생 목록 읽기
 *  2. 메인 DB의 기존 성향분석 데이터와 비교
 *  3. 새 학생이면 추가, 기존 학생인데 날짜가 다르면 업데이트
 * 
 * @returns {Object} { ok, added, updated, skipped, errors }
 */
function syncCounselToMainDB() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { ok: false, error: '서버 바쁨. 잠시 후 재시도해주세요.' };
  }
  
  try {
    // ── 1. counsel 시트 읽기 ──
    const counselSS = SpreadsheetApp.openById(SYNC.COUNSEL_SHEET_ID);
    const counselSheet = counselSS.getSheetByName(SYNC.COUNSEL_TAB);
    if (!counselSheet) {
      return { ok: false, error: 'counsel 시트에서 [DB] 학생 마스터 탭을 찾을 수 없습니다.' };
    }
    
    const counselData = counselSheet.getDataRange().getValues();
    if (counselData.length < 2) {
      return { ok: true, added: 0, updated: 0, skipped: 0, message: 'counsel 데이터 없음' };
    }
    
    // 성향검사가 완료된 학생만 추출 (survey_json이 있는 행)
    const counselStudents = [];
    for (let i = 1; i < counselData.length; i++) {
      const row = counselData[i];
      const name = row[SYNC.C_NAME];
      const surveyJson = row[SYNC.C_SURVEY_JSON];
      
      if (!name || !surveyJson) continue; // 성향검사 미완료
      
      counselStudents.push({
        name: String(name).trim(),
        surveyDate: formatSyncDate_(row[SYNC.C_SURVEY_DATE]),
        reliability: String(row[SYNC.C_RELIABILITY] || ''),
        surveyJson: String(surveyJson),
        ltEngJson: row[SYNC.C_LT_ENG_JSON] ? String(row[SYNC.C_LT_ENG_JSON]) : '',
        ltMathJson: row[SYNC.C_LT_MATH_JSON] ? String(row[SYNC.C_LT_MATH_JSON]) : '',
      });
    }
    
    if (!counselStudents.length) {
      return { ok: true, added: 0, updated: 0, skipped: 0, message: '동기화할 성향검사 데이터 없음' };
    }
    
    // ── 2. 메인 DB 성향분석 읽기 ──
    const mainSS = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = mainSS.getSheetByName(SYNC.MAIN_TAB);
    if (!mainSheet) {
      return { ok: false, error: '메인 DB에서 [DB] 성향분석 탭을 찾을 수 없습니다.' };
    }
    
    const mainData = mainSheet.getDataRange().getValues();
    
    // 기존 데이터를 이름+검사일로 인덱싱
    // key: "이름|검사일" → row index (1-based)
    const existingMap = {};
    // 이름만으로도 최신 데이터 추적
    const latestByName = {};
    
    for (let i = 1; i < mainData.length; i++) {
      const name = String(mainData[i][0]).trim();
      const date = formatSyncDate_(mainData[i][1]);
      if (!name) continue;
      
      existingMap[name + '|' + date] = i + 1; // 1-based row
      
      // 이름별 최신 날짜 추적
      if (!latestByName[name] || date > latestByName[name].date) {
        latestByName[name] = { date, row: i + 1 };
      }
    }
    
    // ── 3. 동기화 실행 ──
    let added = 0, updated = 0, skipped = 0;
    const errors = [];
    
    for (const cs of counselStudents) {
      try {
        const key = cs.name + '|' + cs.surveyDate;
        
        const newRow = [
          cs.name,           // A: 학생명
          cs.surveyDate,     // B: 검사일
          cs.reliability,    // C: 신뢰도
          cs.surveyJson,     // D: 분석결과 JSON
          cs.ltEngJson,      // E: 레벨테스트(영어)
          cs.ltMathJson,     // F: 레벨테스트(수학)
        ];
        
        if (existingMap[key]) {
          // 같은 이름 + 같은 날짜 → 이미 동기화됨
          skipped++;
        } else if (latestByName[cs.name] && latestByName[cs.name].date === cs.surveyDate) {
          // 같은 이름, 같은 날짜인데 키가 다른 경우(형식 차이) → 스킵
          skipped++;
        } else if (latestByName[cs.name] && cs.surveyDate > latestByName[cs.name].date) {
          // 같은 이름이지만 counsel 쪽이 더 최신 → 기존 행 업데이트
          const targetRow = latestByName[cs.name].row;
          mainSheet.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
          updated++;
        } else {
          // 새 학생 또는 새 검사 → 행 추가
          mainSheet.appendRow(newRow);
          added++;
        }
      } catch (e) {
        errors.push(cs.name + ': ' + e.message);
      }
    }
    
    const result = {
      ok: true,
      total: counselStudents.length,
      added,
      updated,
      skipped,
      message: `동기화 완료 — 추가 ${added}명, 갱신 ${updated}명, 기존 ${skipped}명`,
    };
    
    if (errors.length) {
      result.errors = errors;
      result.message += ` (오류 ${errors.length}건)`;
    }
    
    // 로그 기록
    Logger.log('🔄 counsel → 메인DB 동기화: ' + JSON.stringify(result));
    
    return result;
    
  } catch (e) {
    Logger.log('❌ 동기화 오류: ' + e.message);
    return { ok: false, error: '동기화 실패: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}


/**
 * ── 날짜 형식 통일 ──
 * Date 객체든 문자열이든 'YYYY-MM-DD' 형식으로 변환
 */
function formatSyncDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = (val.getMonth() + 1).toString().padStart(2, '0');
    const d = val.getDate().toString().padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  const s = String(val).trim();
  // "2026. 3. 20" → "2026-03-20"
  const match = s.match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (match) {
    return match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0');
  }
  // "2026-03-20" 그대로
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}


// ═══════════════════════════════════════════════════════════════
// 트리거 설정 — 매일 오전 6시 자동 동기화
// ═══════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════
// counsel 학생 목록 조회 (입학 확정 대상)
// ═══════════════════════════════════════════════════════════════

/**
 * counsel [DB] 학생 마스터에서 학생 목록 조회
 * 포털에서 입학 확정 대상을 보여줄 때 사용
 */
function tpAdminCounselStudents_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  
  try {
    var counselSS = SpreadsheetApp.openById(SYNC.COUNSEL_SHEET_ID);
    var counselSheet = counselSS.getSheetByName(SYNC.COUNSEL_TAB);
    if (!counselSheet) return { ok: false, error: 'counsel [DB] 학생 마스터 탭 없음' };
    
    var data = counselSheet.getDataRange().getValues();
    
    // 메인 DB [DB] 학생에 이미 있는 이름 Set
    var mainSS = SpreadsheetApp.getActiveSpreadsheet();
    var studentSheet = mainSS.getSheetByName(SYNC.MAIN_STUDENTS);
    var existingNames = new Set();
    if (studentSheet) {
      var sData = studentSheet.getDataRange().getValues();
      for (var i = 1; i < sData.length; i++) {
        if (sData[i][0]) existingNames.add(String(sData[i][0]).trim());
      }
    }
    
    var students = [];
    for (var i = 1; i < data.length; i++) {
      var name = data[i][SYNC.C_NAME];
      if (!name) continue;
      name = String(name).trim();
      
      var hasSurvey = !!data[i][SYNC.C_SURVEY_JSON];
      var hasEngLT = !!data[i][SYNC.C_LT_ENG_JSON];
      var hasMathLT = !!data[i][SYNC.C_LT_MATH_JSON];
      var status = String(data[i][SYNC.C_STATUS] || '');
      var alreadyEnrolled = existingNames.has(name);
      
      students.push({
        name: name,
        grade: String(data[i][SYNC.C_GRADE] || ''),
        school: String(data[i][SYNC.C_SCHOOL] || ''),
        parentPhone: String(data[i][SYNC.C_PARENT_PHONE] || ''),
        studentPhone: String(data[i][SYNC.C_PHONE] || ''),
        status: status,
        hasSurvey: hasSurvey,
        hasEngLT: hasEngLT,
        hasMathLT: hasMathLT,
        alreadyEnrolled: alreadyEnrolled,
      });
    }
    
    return { ok: true, students: students };
  } catch (e) {
    return { ok: false, error: 'counsel 접근 실패: ' + e.message };
  }
}


// ═══════════════════════════════════════════════════════════════
// counsel 학생 상세 — 성향검사/레벨테스트 원본 JSON 조회
// ═══════════════════════════════════════════════════════════════
function tpCounselStudentDetail_(studentName, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  if (!studentName) return { ok: false, error: '학생명 필요' };
  
  try {
    var counselSS = SpreadsheetApp.openById(SYNC.COUNSEL_SHEET_ID);
    var counselSheet = counselSS.getSheetByName(SYNC.COUNSEL_TAB);
    if (!counselSheet) return { ok: false, error: 'counsel 탭 없음' };
    
    var data = counselSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][SYNC.C_NAME]).trim() !== studentName.trim()) continue;
      
      return {
        ok: true,
        student: {
          name: studentName.trim(),
          grade: String(data[i][SYNC.C_GRADE] || ''),
          school: String(data[i][SYNC.C_SCHOOL] || ''),
          surveyDate: String(data[i][SYNC.C_SURVEY_DATE] || ''),
          reliability: String(data[i][SYNC.C_RELIABILITY] || ''),
          surveyJson: data[i][SYNC.C_SURVEY_JSON] ? String(data[i][SYNC.C_SURVEY_JSON]) : null,
          engLtJson: data[i][SYNC.C_LT_ENG_JSON] ? String(data[i][SYNC.C_LT_ENG_JSON]) : null,
          mathLtJson: data[i][SYNC.C_LT_MATH_JSON] ? String(data[i][SYNC.C_LT_MATH_JSON]) : null,
        }
      };
    }
    return { ok: false, error: studentName + ' 학생을 찾을 수 없습니다.' };
  } catch (e) {
    return { ok: false, error: 'counsel 접근 실패: ' + e.message };
  }
}


// ═══════════════════════════════════════════════════════════════
// 입학 확정 — counsel → 메인 DB 3곳에 한번에 기록
// ═══════════════════════════════════════════════════════════════

/**
 * counsel [DB] 학생 마스터에서 특정 학생을 찾아
 * 메인 DB의 [DB] 학생 + [DB] 성향분석에 동시 추가
 * 
 * @param {string} studentName - 입학 확정할 학생명
 * @param {string} token - 인증 토큰
 * @returns {Object} { ok, message, details }
 */
function tpAdminEnrollStudent_(studentName, classNames, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  if (!studentName) return { ok: false, error: '학생명이 필요합니다.' };
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { ok: false, error: '서버 바쁨. 잠시 후 재시도해주세요.' };
  }
  
  try {
    // ── 1. counsel에서 학생 데이터 읽기 ──
    var counselSS = SpreadsheetApp.openById(SYNC.COUNSEL_SHEET_ID);
    var counselSheet = counselSS.getSheetByName(SYNC.COUNSEL_TAB);
    if (!counselSheet) return { ok: false, error: 'counsel [DB] 학생 마스터 탭 없음' };
    
    var counselData = counselSheet.getDataRange().getValues();
    var found = null;
    
    for (var i = 1; i < counselData.length; i++) {
      if (String(counselData[i][SYNC.C_NAME]).trim() === studentName.trim()) {
        found = counselData[i];
        break;
      }
    }
    
    if (!found) return { ok: false, error: 'counsel에서 "' + studentName + '" 학생을 찾을 수 없습니다.' };
    
    var mainSS = SpreadsheetApp.getActiveSpreadsheet();
    var details = { student: false, personality: false };
    
    // ── 2. [DB] 학생 탭에 추가 (중복 체크) ──
    var studentSheet = mainSS.getSheetByName(SYNC.MAIN_STUDENTS);
    if (studentSheet) {
      var sData = studentSheet.getDataRange().getValues();
      var exists = false;
      for (var i = 1; i < sData.length; i++) {
        if (String(sData[i][0]).trim() === studentName.trim()) {
          exists = true;
          break;
        }
      }
      
      if (!exists) {
        studentSheet.appendRow([
          studentName.trim(),                                    // A: 학생명
          String(found[SYNC.C_GRADE] || ''),                    // B: 학년
          String(found[SYNC.C_SCHOOL] || ''),                   // C: 학교
          String(found[SYNC.C_PARENT_PHONE] || ''),             // D: 학부모 연락처
          String(found[SYNC.C_PHONE] || ''),                    // E: 학생 연락처
          '재원',                                                // F: 재원 여부
        ]);
        details.student = true;
      } else {
        details.student = 'already_exists';
      }
    }
    
    // ── 3. [DB] 성향분석 탭에 추가 (성향검사 있을 때만) ──
    var surveyJson = found[SYNC.C_SURVEY_JSON];
    if (surveyJson) {
      var pSheet = mainSS.getSheetByName(SYNC.MAIN_TAB);
      if (pSheet) {
        // 중복 체크: 같은 이름 + 같은 날짜
        var pData = pSheet.getDataRange().getValues();
        var surveyDate = formatSyncDate_(found[SYNC.C_SURVEY_DATE]);
        var pExists = false;
        
        for (var i = 1; i < pData.length; i++) {
          if (String(pData[i][0]).trim() === studentName.trim() 
              && formatSyncDate_(pData[i][1]) === surveyDate) {
            pExists = true;
            break;
          }
        }
        
        if (!pExists) {
          pSheet.appendRow([
            studentName.trim(),                                  // A: 학생명
            surveyDate,                                          // B: 검사일
            String(found[SYNC.C_RELIABILITY] || ''),             // C: 신뢰도
            String(surveyJson),                                  // D: 분석결과 JSON
            found[SYNC.C_LT_ENG_JSON] ? String(found[SYNC.C_LT_ENG_JSON]) : '',   // E: 레벨(영어)
            found[SYNC.C_LT_MATH_JSON] ? String(found[SYNC.C_LT_MATH_JSON]) : '', // F: 레벨(수학)
          ]);
          details.personality = true;
        } else {
          details.personality = 'already_exists';
        }
      }
    }
    
    // ── 4. [DB] 수강 탭에 반 배정 (classNames 배열) ──
    details.enrollment = [];
    if (classNames && classNames.length > 0) {
      var enrollSheet = mainSS.getSheetByName('[DB] 수강');
      if (enrollSheet) {
        var eData = enrollSheet.getDataRange().getValues();
        var existingEnroll = new Set();
        for (var i = 1; i < eData.length; i++) {
          existingEnroll.add(String(eData[i][0]).trim() + '|' + String(eData[i][1]).trim());
        }
        
        var grade = String(found[SYNC.C_GRADE] || '');
        for (var c = 0; c < classNames.length; c++) {
          var cn = classNames[c];
          if (!cn) continue;
          var key = studentName.trim() + '|' + cn;
          if (existingEnroll.has(key)) {
            details.enrollment.push(cn + ' (이미 배정)');
          } else {
            enrollSheet.appendRow([studentName.trim(), cn, grade]);
            details.enrollment.push(cn + ' (배정 완료)');
          }
        }
      }
    }
    
    // ── 결과 메시지 ──
    var msgs = [];
    if (details.student === true) msgs.push('[DB] 학생 추가');
    else if (details.student === 'already_exists') msgs.push('[DB] 학생 이미 존재');
    if (details.personality === true) msgs.push('성향분석 추가');
    else if (details.personality === 'already_exists') msgs.push('성향분석 이미 존재');
    else if (!surveyJson) msgs.push('성향검사 미완료');
    if (details.enrollment.length > 0) msgs.push('반 배정: ' + details.enrollment.join(', '));
    
    return {
      ok: true,
      message: studentName + ' 입학 확정: ' + msgs.join(', '),
      details: details,
    };
    
  } catch (e) {
    return { ok: false, error: '입학 확정 실패: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}


/**
 * 1회만 실행하면 됩니다.
 * 매일 오전 6시~7시 사이에 syncCounselToMainDB()가 자동 실행됩니다.
 */
function setupSyncTrigger() {
  // 기존 동기화 트리거 제거 (중복 방지)
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'syncCounselToMainDB') {
      ScriptApp.deleteTrigger(t);
      Logger.log('기존 동기화 트리거 삭제');
    }
  }
  
  // 새 트리거 생성 — 매일 오전 6시
  ScriptApp.newTrigger('syncCounselToMainDB')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  
  Logger.log('✅ 매일 오전 6시 자동 동기화 트리거 설정 완료');
  Logger.log('수동 실행: syncCounselToMainDB() 함수를 직접 실행하거나, 포털에서 동기화 버튼 클릭');
}


/**
 * 트리거 해제 (필요 시)
 */
function removeSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'syncCounselToMainDB') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  }
  Logger.log('동기화 트리거 ' + removed + '개 삭제');
}


// ═══════════════════════════════════════════════════════════════
// 포털 API 연동 — GET 핸들러에서 호출
// ═══════════════════════════════════════════════════════════════
// 
// TeacherPortalAPI_v2.gs의 tpHandleGet_ switch문에 아래를 추가:
//
//   case 'tp_admin_syncCounsel':
//     r = tpAdminSyncCounsel_(e.parameter.token);
//     break;
//

/**
 * 포털에서 호출되는 래퍼 함수
 * 관리자만 실행 가능
 */
function tpAdminSyncCounsel_(token) {
  // 토큰 검증 (TeacherPortalAPI_v2.gs의 함수 재사용)
  if (!tpValidToken_(token)) {
    return { ok: false, error: '인증 만료. 다시 로그인해주세요.' };
  }
  
  // 동기화 실행
  return syncCounselToMainDB();
}


// ═══════════════════════════════════════════════════════════════
// counsel 대시보드 — 파이프라인 현황 조회
// ═══════════════════════════════════════════════════════════════
function tpCounselDashboard_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  
  try {
    var counselSS = SpreadsheetApp.openById(SYNC.COUNSEL_SHEET_ID);
    var counselSheet = counselSS.getSheetByName(SYNC.COUNSEL_TAB);
    if (!counselSheet) return { ok: false, error: 'counsel 탭 없음' };
    
    var data = counselSheet.getDataRange().getValues();
    var pipeline = { '접수': 0, '성향완료': 0, '레벨완료': 0, '분석완료': 0 };
    var recent = [];
    
    for (var i = 1; i < data.length; i++) {
      var name = data[i][SYNC.C_NAME];
      if (!name) continue;
      name = String(name).trim();
      var status = String(data[i][SYNC.C_STATUS] || '접수');
      
      if (pipeline.hasOwnProperty(status)) pipeline[status]++;
      else pipeline['접수']++;
      
      recent.push({
        student_id: String(data[i][0] || ''),
        name: name,
        grade: String(data[i][SYNC.C_GRADE] || ''),
        school: String(data[i][SYNC.C_SCHOOL] || ''),
        phone: String(data[i][SYNC.C_PHONE] || ''),
        parentPhone: String(data[i][SYNC.C_PARENT_PHONE] || ''),
        status: status,
        hasSurvey: !!data[i][SYNC.C_SURVEY_JSON],
        hasEngLT: !!data[i][SYNC.C_LT_ENG_JSON],
        hasMathLT: !!data[i][SYNC.C_LT_MATH_JSON],
        hasAnalysis: !!data[i][29],  // AD열 (분석날짜)
        created: data[i][1] ? Utilities.formatDate(new Date(data[i][1]), 'Asia/Seoul', 'yyyy-MM-dd') : ''
      });
    }
    
    // 최근순 정렬
    recent.sort(function(a, b) { return b.student_id.localeCompare(a.student_id); });
    
    return {
      ok: true,
      total: recent.length,
      pipeline: pipeline,
      recent: recent.slice(0, 30)
    };
  } catch (e) {
    return { ok: false, error: 'counsel 접근 실패: ' + e.message };
  }
}


// ═══════════════════════════════════════════════════════════════
// counsel 상담접수 — counsel [DB] 학생 마스터에 직접 기록
// (counsel의 Consultation.gs submitConsultation과 동일한 데이터 구조)
// ═══════════════════════════════════════════════════════════════
function tpCounselSubmit_(formData, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  if (!formData || !formData.name) return { ok: false, error: '학생 이름이 필요합니다.' };
  
  try {
    var counselSS = SpreadsheetApp.openById(SYNC.COUNSEL_SHEET_ID);
    var counselSheet = counselSS.getSheetByName(SYNC.COUNSEL_TAB);
    if (!counselSheet) return { ok: false, error: 'counsel 탭 없음' };
    
    var studentId = Date.now().toString();
    var now = new Date();
    var TOTAL_COLS = 41;
    
    // counsel Code.gs의 COL 매핑과 동일하게 구성
    var newRow = new Array(TOTAL_COLS).fill('');
    
    // A: student_id, B: created_at
    newRow[0] = studentId;
    newRow[1] = now;
    
    // C~N: 기본 프로필
    newRow[2]  = formData.name || '';          // C: 이름
    newRow[3]  = formData.phone || '';         // D: 학생 전화
    newRow[4]  = formData.parent_phone || '';  // E: 부모 연락처
    newRow[5]  = formData.address || '';       // F: 주소
    newRow[6]  = formData.school || '';        // G: 학교
    newRow[7]  = formData.grade || '';         // H: 학년
    newRow[8]  = formData.grades || '';        // I: 성적
    newRow[9]  = formData.consent ? true : false; // J: 동의
    newRow[10] = formData.route || '';         // K: 유입경로
    newRow[11] = formData.reason || '';        // L: 상담사유
    newRow[12] = formData.subject || '';       // M: 과목
    newRow[13] = formData.wants || '';         // N: 희망사항
    
    // O: 목표대학
    newRow[14] = formData.target_univ || '';
    
    // P~S: 모의고사 (고등학생)
    var grade = String(formData.grade || '');
    if (grade.indexOf('고') === 0) {
      newRow[15] = formData.mock_korean || '';
      newRow[16] = formData.mock_english || '';
      newRow[17] = formData.mock_math || '';
      newRow[18] = formData.mock_science || '';
    }
    
    // AG: 상태, AH: 접수일시
    newRow[32] = '접수';
    newRow[33] = now;
    
    counselSheet.appendRow(newRow);
    
    return {
      ok: true,
      studentId: studentId,
      message: (formData.name) + ' 학생의 상담이 접수되었습니다.'
    };
  } catch (e) {
    return { ok: false, error: '접수 실패: ' + e.message };
  }
}
