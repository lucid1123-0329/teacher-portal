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
  
  // counsel [DB] 학생 마스터 컬럼 인덱스 (0-based) — Code.gs COL과 동일
  C_ID: 0,              // A열: student_id
  C_CREATED: 1,         // B열: created_at
  C_NAME: 2,            // C열: 학생 이름
  C_PHONE: 3,           // D열: 학생 연락처
  C_PARENT_PHONE: 4,    // E열: 학부모 연락처
  C_ADDRESS: 5,         // F열: 주소
  C_SCHOOL: 6,          // G열: 학교
  C_GRADE: 7,           // H열: 학년
  C_GRADES: 8,          // I열: 성적
  C_CONSENT: 9,         // J열: 동의
  C_ROUTE: 10,          // K열: 유입경로
  C_REASON: 11,         // L열: 상담사유
  C_SUBJECT: 12,        // M열: 과목
  C_WANTS: 13,          // N열: 희망사항
  C_TARGET_UNIV: 14,    // O열: 목표대학
  C_MOCK_KOR: 15,       // P열: 모의 국어
  C_MOCK_ENG: 16,       // Q열: 모의 영어
  C_MOCK_MATH: 17,      // R열: 모의 수학
  C_MOCK_SCI: 18,       // S열: 모의 과학
  C_SURVEY_TYPE: 19,    // T열: 검사 유형
  C_SURVEY_DATE: 20,    // U열: 성향검사 실시일
  C_RELIABILITY: 21,    // V열: 신뢰도
  C_SURVEY_JSON: 22,    // W열: 성향분석 JSON
  C_LT_ENG_TYPE: 23,    // X열: 영어 LT 유형
  C_LT_ENG_DATE: 24,    // Y열: 영어 LT 날짜
  C_LT_ENG_JSON: 25,    // Z열: 영어 LT JSON
  C_LT_MATH_TYPE: 26,   // AA열: 수학 LT 유형
  C_LT_MATH_DATE: 27,   // AB열: 수학 LT 날짜
  C_LT_MATH_JSON: 28,   // AC열: 수학 LT JSON
  C_ANALYSIS_DATE: 29,  // AD열: 분석일
  C_ANALYSIS_JSON: 30,  // AE열: 분석결과 JSON
  C_REPORT_TOKEN: 31,   // AF열: 리포트 토큰
  C_STATUS: 32,         // AG열: 파이프라인 상태
  C_STEP_CONSULT: 33,   // AH열
  C_STEP_SURVEY: 34,    // AI열
  C_STEP_LT_ENG: 35,    // AJ열
  C_STEP_LT_MATH: 36,   // AK열
  C_STEP_ANALYSIS: 37,  // AL열
  C_MEMO: 38,           // AM열
  C_TOTAL_COLS: 41,
  
  // counsel 시트 추가 탭
  TAB_RAW_SURVEY: '[RAW] 성향응답',
  TAB_REF_SURVEY_Q: '[REF] 설문문항',
  TAB_REF_ANSWER: '[REF] 정답지',
  TAB_REF_ROADMAP: '[REF] 로드맵',
  TAB_REF_LEVEL: '[REF] 레벨기준',
  TAB_REF_STUDY: '[REF] 학습방향',
  TAB_LOG_GRADE: '[LOG] 채점이력',
  TAB_OUT_ANALYSIS: '[OUT] 분석결과',
  TAB_CONFIG: '[CONFIG]',
  
  // 메인 DB
  MAIN_TAB: '[DB] 성향분석',
  MAIN_STUDENTS: '[DB] 학생',
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


// ═══════════════════════════════════════════════════════════════
// COUNSEL 시트 헬퍼 함수 (counsel Apps Script 무변경 원칙)
// ═══════════════════════════════════════════════════════════════

function counselSS_() {
  return SpreadsheetApp.openById(SYNC.COUNSEL_SHEET_ID);
}

function counselSheet_(tabName) {
  var ss = counselSS_();
  var sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error('counsel 탭을 찾을 수 없습니다: ' + tabName);
  return sh;
}

function counselStudentById_(studentId) {
  var sheet = counselSheet_(SYNC.COUNSEL_TAB);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(studentId).trim()) {
      return { row_index: i + 1, data: data[i] };
    }
  }
  return null;
}

function counselUpdateCells_(rowIndex, updates) {
  var sheet = counselSheet_(SYNC.COUNSEL_TAB);
  for (var col in updates) {
    sheet.getRange(rowIndex, parseInt(col)).setValue(updates[col]);
  }
}

function counselConfig_(key) {
  var sheet = counselSheet_(SYNC.TAB_CONFIG);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) return data[i][1];
  }
  throw new Error('CONFIG 키를 찾을 수 없습니다: ' + key);
}

function counselToday_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
}


// ═══════════════════════════════════════════════════════════════
// 성향검사 API (Survey.gs 동일 로직)
// ═══════════════════════════════════════════════════════════════

function tpCounselSurveyStudents_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var sheet = counselSheet_(SYNC.COUNSEL_TAB);
    var data = sheet.getDataRange().getValues();
    var students = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][SYNC.C_ID]) continue;
      students.push({
        student_id: String(data[i][SYNC.C_ID]),
        name: String(data[i][SYNC.C_NAME] || ''),
        grade: String(data[i][SYNC.C_GRADE] || ''),
        school: String(data[i][SYNC.C_SCHOOL] || ''),
        status: String(data[i][SYNC.C_STATUS] || ''),
        has_survey: !!data[i][SYNC.C_SURVEY_JSON],
        survey_type: String(data[i][SYNC.C_SURVEY_TYPE] || ''),
        survey_date: String(data[i][SYNC.C_SURVEY_DATE] || ''),
        survey_reliability: String(data[i][SYNC.C_RELIABILITY] || '')
      });
    }
    return { ok: true, students: students };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function tpCounselSurveyQuestions_(surveyType, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var sheet = counselSheet_(SYNC.TAB_REF_SURVEY_Q);
    var data = sheet.getDataRange().getValues();
    var questions = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === surveyType) {
        questions.push({
          index: parseInt(data[i][1]) || (questions.length + 1),
          text: String(data[i][2] || '').trim()
        });
      }
    }
    var expectedQ = surveyType === 'elementary' ? 150 : 170;
    return { ok: true, type: surveyType, questions: questions, totalQ: questions.length, expectedQ: expectedQ };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function tpCounselSurveySubmit_(data, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var student = counselStudentById_(data.studentId);
    if (!student) throw new Error('학생을 찾을 수 없습니다.');
    var studentName = student.data[SYNC.C_NAME] || '';
    var now = new Date();
    var todayStr = counselToday_();

    // [RAW] 성향응답 저장
    var rawSheet = counselSheet_(SYNC.TAB_RAW_SURVEY);
    var rawRow = [data.studentId, studentName, data.surveyType, now];
    (data.answers || []).forEach(function(a) { rawRow.push(a); });
    rawSheet.appendRow(rawRow);

    // Gemini 분석
    var result = counselRunSurveyGemini_(data.studentId, data.surveyType);

    // [DB] 학생 마스터 업데이트
    var updates = {};
    updates[SYNC.C_SURVEY_TYPE + 1] = data.surveyType;
    updates[SYNC.C_SURVEY_DATE + 1] = todayStr;
    updates[SYNC.C_RELIABILITY + 1] = result.reliability || '';
    updates[SYNC.C_SURVEY_JSON + 1] = result.json || '';
    updates[SYNC.C_STEP_SURVEY + 1] = now;
    var currentStatus = String(student.data[SYNC.C_STATUS] || '');
    if (currentStatus === '접수') updates[SYNC.C_STATUS + 1] = '성향완료';
    counselUpdateCells_(student.row_index, updates);

    return { ok: true, message: '성향검사 분석 완료! (신뢰도: ' + result.reliability + ')', reliability: result.reliability };
  } catch (e) {
    return { ok: false, error: '분석 실패: ' + e.message };
  }
}

function tpCounselSurveyRerun_(studentId, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var student = counselStudentById_(studentId);
    if (!student) throw new Error('학생을 찾을 수 없습니다.');
    var rawSheet = counselSheet_(SYNC.TAB_RAW_SURVEY);
    var rawData = rawSheet.getDataRange().getValues();
    var surveyType = null;
    for (var i = rawData.length - 1; i >= 1; i--) {
      if (String(rawData[i][0]).trim() === String(studentId).trim()) { surveyType = String(rawData[i][2]).trim(); break; }
    }
    if (!surveyType) throw new Error('성향검사 응답을 찾을 수 없습니다.');
    var result = counselRunSurveyGemini_(studentId, surveyType);
    var now = new Date(); var todayStr = counselToday_();
    var updates = {};
    updates[SYNC.C_SURVEY_TYPE + 1] = surveyType;
    updates[SYNC.C_SURVEY_DATE + 1] = todayStr;
    updates[SYNC.C_RELIABILITY + 1] = result.reliability || '';
    updates[SYNC.C_SURVEY_JSON + 1] = result.json || '';
    updates[SYNC.C_STEP_SURVEY + 1] = now;
    if (String(student.data[SYNC.C_STATUS] || '') === '접수') updates[SYNC.C_STATUS + 1] = '성향완료';
    counselUpdateCells_(student.row_index, updates);
    return { ok: true, message: '재분석 완료! (신뢰도: ' + result.reliability + ')' };
  } catch (e) {
    return { ok: false, error: '재분석 실패: ' + e.message };
  }
}

function tpCounselSurveyResult_(studentId, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var student = counselStudentById_(studentId);
    if (!student) return { ok: false, error: '학생을 찾을 수 없습니다.' };
    var jsonStr = student.data[SYNC.C_SURVEY_JSON];
    if (!jsonStr) return { ok: false, error: '성향검사 결과가 없습니다.' };
    return { ok: true, dataJson: String(jsonStr), type: String(student.data[SYNC.C_SURVEY_TYPE] || ''), date: String(student.data[SYNC.C_SURVEY_DATE] || ''), reliability: String(student.data[SYNC.C_RELIABILITY] || '') };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Survey.gs의 Gemini 분석 로직 복제 (counsel 시트 전용)
var COUNSEL_FRAMEWORK_ELEM = '\n[신뢰도 척도 규칙 (초등)]:\n1. L-Scale (착한 어린이):\n    - 문항: 15, 35, 56, 114\n    - 판정: ④⑤ 응답 많으면 방어적 태도 (신뢰도 저하)\n2. S-Scale (모범생 과장):\n    - 문항: 24, 49, 66, 75, 89, 98, 106, 143, 150\n    - 판정: ④⑤ 응답 많으면 학업 약점 숨김 (신뢰도 저하)\n3. VRIN (비일관성): 모순 2개 이상 → 신뢰도 하\n';
var COUNSEL_FRAMEWORK_SEC = '\n[신뢰도 척도 규칙 (중/고등)]:\n1. L-Scale: 문항 9,15,35,41,51,66,90,125,144,156,161,166 (12개)\n2. S-Scale: 문항 46,72,84,95,101,108,116,138 (8개)\n3. VRIN: 모순 3개 이상 → 신뢰도 하\n';

function counselRunSurveyGemini_(studentId, surveyType) {
  var rawSheet = counselSheet_(SYNC.TAB_RAW_SURVEY);
  var rawData = rawSheet.getDataRange().getValues();
  var studentRow = null, studentName = '';
  for (var i = rawData.length - 1; i >= 1; i--) {
    if (String(rawData[i][0]).trim() === String(studentId).trim()) { studentRow = rawData[i]; studentName = String(rawData[i][1]); break; }
  }
  if (!studentRow) throw new Error('응답 데이터 없음');

  var refSheet = counselSheet_(SYNC.TAB_REF_SURVEY_Q);
  var refData = refSheet.getDataRange().getValues();
  var qTexts = [];
  for (var i = 1; i < refData.length; i++) { if (String(refData[i][0]).trim() === surveyType) qTexts.push(String(refData[i][2] || '').trim()); }
  var answers = studentRow.slice(4);
  var framework = surveyType === 'elementary' ? COUNSEL_FRAMEWORK_ELEM : COUNSEL_FRAMEWORK_SEC;
  var rawStr = '';
  for (var i = 0; i < qTexts.length; i++) { if (!qTexts[i]) break; rawStr += '[Q] ' + qTexts[i] + '\n[A] ' + (answers[i] || '-') + '\n\n'; }

  var apiKey = counselConfig_('GEMINI_API_KEY');
  var modelPrimary = counselConfig_('GEMINI_MODEL_PRIMARY');
  var modelFallback = ''; try { modelFallback = counselConfig_('GEMINI_MODEL_FALLBACK'); } catch(e) {}
  var maxRetries = 5; try { maxRetries = parseInt(counselConfig_('GEMINI_MAX_RETRIES')); } catch(e) {}
  var temperature = 0.2; try { temperature = parseFloat(counselConfig_('GEMINI_TEMPERATURE')); } catch(e) {}
  var todayStr = counselToday_();

  var prompt = '당신은 교육 심리 데이터 분석 전문가입니다.\n"' + studentName + '" 학생의 응답 데이터를 분석하여 JSON 형식으로 출력하십시오.\n\n[Step 1: 신뢰도 판정]\n[Step 2: 점수 산출 (100점 만점)]\n[Step 3: 서술형 코멘트 (문항번호 언급 금지, 학부모 대상 상담체)]\n\n[분석 가이드라인]:\n' + framework + '\n\n[학생 응답 데이터]:\n' + rawStr;

  var schema = counselBuildSurveySchema_(todayStr);
  var models = [modelPrimary]; if (modelFallback) models.push(modelFallback);
  var lastError = null;

  for (var m = 0; m < models.length; m++) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[m] + ':generateContent?key=' + apiKey;
    var payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: temperature, topP: 0.8, topK: 40, responseMimeType: 'application/json', responseSchema: schema } };
    var options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

    for (var attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        var response = UrlFetchApp.fetch(url, options);
        var code = response.getResponseCode();
        var body = response.getContentText();
        if (code === 200) {
          var json = JSON.parse(body);
          if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) {
            var resultText = json.candidates[0].content.parts[0].text;
            var resultObj = JSON.parse(resultText);
            return { json: resultText, reliability: resultObj.reliability ? resultObj.reliability.score : '알 수 없음' };
          }
        } else if (code === 429 || code === 503) {
          if (attempt < maxRetries) { Utilities.sleep(2000 * Math.pow(2, attempt - 1)); continue; }
        } else { lastError = new Error('API 오류 (' + models[m] + ', HTTP ' + code + ')'); break; }
      } catch (e) { lastError = e; if (attempt < maxRetries) { Utilities.sleep(2000); continue; } }
    }
  }
  throw lastError || new Error('Gemini API 호출 실패');
}

function counselBuildSurveySchema_(dateStr) {
  return {type:'OBJECT',properties:{analysis_date:{type:'STRING'},reliability:{type:'OBJECT',properties:{score:{type:'STRING',enum:['상','중','하']},comment:{type:'STRING'},detected_issues:{type:'ARRAY',items:{type:'STRING'}}},required:['score','comment','detected_issues']},profileScores:{type:'OBJECT',properties:{motivation:{type:'INTEGER'},planning:{type:'INTEGER'},metacognition:{type:'INTEGER'},self_regulation:{type:'INTEGER'},help_seeking:{type:'INTEGER'}},required:['motivation','planning','metacognition','self_regulation','help_seeking']},profileAnalysis:{type:'OBJECT',properties:{summary_title:{type:'STRING'},summary_text:{type:'STRING'},motivation_text:{type:'STRING'},planning_text:{type:'STRING'},metacognition_text:{type:'STRING'},self_regulation_text:{type:'STRING'},help_seeking_text:{type:'STRING'}},required:['summary_title','summary_text']},sectionAnalysis:{type:'OBJECT',properties:{section1_motivation:{type:'OBJECT',properties:{sectionName:{type:'STRING'},sectionSummary:{type:'STRING'},detailedScores:{type:'OBJECT',properties:{intrinsic_motivation:{type:'INTEGER'},identified_regulation:{type:'INTEGER'},self_efficacy:{type:'INTEGER'},achievement_pressure:{type:'INTEGER'},academic_lethargy:{type:'INTEGER'}}}}},section2_cognition:{type:'OBJECT',properties:{sectionName:{type:'STRING'},sectionSummary:{type:'STRING'},detailedScores:{type:'OBJECT',properties:{metacognition:{type:'INTEGER'},planning_ability:{type:'INTEGER'},memory_strategy:{type:'INTEGER'},comprehension_strategy:{type:'INTEGER'},error_analysis:{type:'INTEGER'}}}}},section3_behavior:{type:'OBJECT',properties:{sectionName:{type:'STRING'},sectionSummary:{type:'STRING'},detailedScores:{type:'OBJECT',properties:{execution_power:{type:'INTEGER'},persistence:{type:'INTEGER'},env_control:{type:'INTEGER'},attention_control:{type:'INTEGER'},digital_detox:{type:'INTEGER'}}}}},section4_emotion:{type:'OBJECT',properties:{sectionName:{type:'STRING'},sectionSummary:{type:'STRING'},detailedScores:{type:'OBJECT',properties:{academic_self_esteem:{type:'INTEGER'},test_anxiety:{type:'INTEGER'},resilience:{type:'INTEGER'},emotion_regulation:{type:'INTEGER'},optimism:{type:'INTEGER'}}}}},section5_environment:{type:'OBJECT',properties:{sectionName:{type:'STRING'},sectionSummary:{type:'STRING'},detailedScores:{type:'OBJECT',properties:{parent_support:{type:'INTEGER'},autonomy_respect:{type:'INTEGER'},help_seeking:{type:'INTEGER'},teacher_rapport:{type:'INTEGER'},peer_relation:{type:'INTEGER'}}}}}},required:['section1_motivation','section2_cognition','section3_behavior','section4_emotion','section5_environment']},consultStrategy:{type:'OBJECT',properties:{priority:{type:'STRING',enum:['상','중','하']},recommend_tag_A:{type:'STRING'},recommend_tag_B:{type:'STRING'},recommend_text:{type:'STRING'}},required:['priority','recommend_tag_A','recommend_tag_B','recommend_text']}},required:['analysis_date','reliability','profileScores','profileAnalysis','sectionAnalysis','consultStrategy']};
}


// ═══════════════════════════════════════════════════════════════
// 레벨테스트 API (LevelTest.gs 동일 로직)
// ═══════════════════════════════════════════════════════════════

function tpCounselLtStudents_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var sheet = counselSheet_(SYNC.COUNSEL_TAB);
    var data = sheet.getDataRange().getValues();
    var students = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][SYNC.C_ID]) continue;
      students.push({
        student_id: String(data[i][SYNC.C_ID]),
        name: String(data[i][SYNC.C_NAME] || ''),
        grade: String(data[i][SYNC.C_GRADE] || ''),
        school: String(data[i][SYNC.C_SCHOOL] || ''),
        status: String(data[i][SYNC.C_STATUS] || ''),
        has_eng: !!data[i][SYNC.C_LT_ENG_TYPE],
        has_math: !!data[i][SYNC.C_LT_MATH_TYPE]
      });
    }
    return { ok: true, students: students };
  } catch (e) { return { ok: false, error: e.message }; }
}

function tpCounselAnswerKey_(subject, level, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var sheet = counselSheet_(SYNC.TAB_REF_ANSWER);
    var data = sheet.getDataRange().getValues();
    var testType = subject === 'math' ? 'math_' + level : 'eng_' + level;

    if (subject === 'math') {
      var result = { subject: 'math', testType: testType, totalQ: 0, answers: [], parts: [], areas: [], scopes: [] };
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() !== testType) continue;
        var rowType = String(data[i][2]).trim();
        var totalQ = parseInt(data[i][1]) || 0;
        if (rowType === 'answer') { result.totalQ = totalQ; result.answers = data[i].slice(3, 3 + totalQ).map(function(v){return String(v).trim();}); }
        else if (rowType === 'part') { result.parts = data[i].slice(3, 3 + totalQ).map(function(v){return String(v).trim();}); }
        else if (rowType === 'area') { result.areas = data[i].slice(3, 3 + totalQ).map(function(v){return String(v).trim();}); }
        else if (rowType === 'scope') { result.scopes = data[i].slice(3, 3 + totalQ).map(function(v){return String(v).trim();}); }
      }
      return { ok: true, data: result };
    } else {
      var items = [];
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() !== testType) continue;
        var qNum = parseInt(data[i][1]);
        if (isNaN(qNum)) continue;
        items.push({ q_number: qNum, behavior: String(data[i][2]).trim(), content: String(data[i][3]).trim(), answer: String(data[i][4]).trim(), description: String(data[i][5] || '').trim() });
      }
      return { ok: true, data: { subject: 'english', testType: testType, totalQ: items.length, items: items } };
    }
  } catch (e) { return { ok: false, error: e.message }; }
}

function tpCounselLtSave_(gradeData, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var student = counselStudentById_(gradeData.studentId);
    if (!student) throw new Error('학생을 찾을 수 없습니다.');
    var now = new Date(); var todayStr = counselToday_();
    var studentName = student.data[SYNC.C_NAME] || '';
    var studentSchool = student.data[SYNC.C_SCHOOL] || '';
    var studentGrade = student.data[SYNC.C_GRADE] || '';

    // [LOG] 채점이력
    var logSheet = counselSheet_(SYNC.TAB_LOG_GRADE);
    logSheet.appendRow([Date.now().toString(), gradeData.studentId, studentName, studentSchool, studentGrade, gradeData.subject, gradeData.testType, todayStr, gradeData.totalScore, gradeData.totalQuestions, gradeData.percentage, gradeData.resultJson, gradeData.scoreArray]);

    // [DB] 학생 마스터 업데이트
    var updates = {};
    if (gradeData.subject === 'english') {
      updates[SYNC.C_LT_ENG_TYPE + 1] = gradeData.testType;
      updates[SYNC.C_LT_ENG_DATE + 1] = todayStr;
      updates[SYNC.C_LT_ENG_JSON + 1] = gradeData.resultJson;
      updates[SYNC.C_STEP_LT_ENG + 1] = now;
    } else {
      updates[SYNC.C_LT_MATH_TYPE + 1] = gradeData.testType;
      updates[SYNC.C_LT_MATH_DATE + 1] = todayStr;
      updates[SYNC.C_LT_MATH_JSON + 1] = gradeData.resultJson;
      updates[SYNC.C_STEP_LT_MATH + 1] = now;
    }
    var hasEng = gradeData.subject === 'english' ? true : !!student.data[SYNC.C_LT_ENG_TYPE];
    var hasMath = gradeData.subject === 'math' ? true : !!student.data[SYNC.C_LT_MATH_TYPE];
    if (hasEng && hasMath) updates[SYNC.C_STATUS + 1] = '레벨완료';
    counselUpdateCells_(student.row_index, updates);

    var subjectKo = gradeData.subject === 'math' ? '수학' : '영어';
    return { ok: true, message: subjectKo + ' ' + gradeData.testType + ' 채점 완료! (' + gradeData.totalScore + '/' + gradeData.totalQuestions + ', ' + gradeData.percentage + '%)' };
  } catch (e) { return { ok: false, error: '저장 실패: ' + e.message }; }
}


// ═══════════════════════════════════════════════════════════════
// 분석/리포트 API (Analysis.gs 동일 로직)
// ═══════════════════════════════════════════════════════════════

function tpCounselAnalysisStudents_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var sheet = counselSheet_(SYNC.COUNSEL_TAB);
    var data = sheet.getDataRange().getValues();
    var students = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][SYNC.C_ID]) continue;
      students.push({
        student_id: String(data[i][SYNC.C_ID]),
        name: String(data[i][SYNC.C_NAME] || ''),
        grade: String(data[i][SYNC.C_GRADE] || ''),
        school: String(data[i][SYNC.C_SCHOOL] || ''),
        status: String(data[i][SYNC.C_STATUS] || ''),
        has_survey: !!data[i][SYNC.C_SURVEY_JSON],
        has_eng: !!data[i][SYNC.C_LT_ENG_JSON],
        has_math: !!data[i][SYNC.C_LT_MATH_JSON],
        has_analysis: !!data[i][SYNC.C_ANALYSIS_JSON],
        analysis_date: String(data[i][SYNC.C_ANALYSIS_DATE] || ''),
        report_token: String(data[i][SYNC.C_REPORT_TOKEN] || '')
      });
    }
    return { ok: true, students: students };
  } catch (e) { return { ok: false, error: e.message }; }
}

function tpCounselAnalysisResult_(studentId, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var student = counselStudentById_(studentId);
    if (!student) return { ok: false, error: '학생을 찾을 수 없습니다.' };
    var jsonStr = student.data[SYNC.C_ANALYSIS_JSON];
    if (!jsonStr) return { ok: false, error: '분석 결과가 없습니다.' };
    return { ok: true, dataJson: String(jsonStr), date: String(student.data[SYNC.C_ANALYSIS_DATE] || ''), token: String(student.data[SYNC.C_REPORT_TOKEN] || '') };
  } catch (e) { return { ok: false, error: e.message }; }
}

function tpCounselVerifyAdmin_(password, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var adminPw = counselConfig_('ADMIN_PASSWORD');
    if (String(password).trim() === String(adminPw).trim()) return { ok: true };
    return { ok: false, error: '비밀번호가 올바르지 않습니다.' };
  } catch (e) { return { ok: false, error: e.message }; }
}

function tpCounselRunAnalysis_(studentId, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var masterSheet = counselSheet_(SYNC.COUNSEL_TAB);
    var student = counselStudentById_(studentId);
    if (!student) throw new Error('학생을 찾을 수 없습니다.');
    var row = student.data;
    var studentName = row[SYNC.C_NAME] || '';
    var grade = String(row[SYNC.C_GRADE] || '');
    var isHS = grade.trim().indexOf('고') === 0;

    var profile = { '학생 이름': studentName, '학교명': row[SYNC.C_SCHOOL]||'', '학년': grade, '학교 성적': row[SYNC.C_GRADES]||'', '유입 경로': row[SYNC.C_ROUTE]||'', '상담 사유': row[SYNC.C_REASON]||'', '희망 과목': row[SYNC.C_SUBJECT]||'', '희망 사항': row[SYNC.C_WANTS]||'', '희망 목표 대학': row[SYNC.C_TARGET_UNIV]||'' };
    var mockScores = { korean: row[SYNC.C_MOCK_KOR]||'정보 없음', english: row[SYNC.C_MOCK_ENG]||'정보 없음', math: row[SYNC.C_MOCK_MATH]||'정보 없음', science: row[SYNC.C_MOCK_SCI]||'정보 없음' };
    var surveyJson = row[SYNC.C_SURVEY_JSON] || '{}';
    var engJson = row[SYNC.C_LT_ENG_JSON] || '{}';
    var mathJson = row[SYNC.C_LT_MATH_JSON] || '{}';
    var engType = row[SYNC.C_LT_ENG_TYPE] || '';
    var mathType = row[SYNC.C_LT_MATH_TYPE] || '';

    // 참조 데이터
    var refRoadmap = counselGetRefData_(SYNC.TAB_REF_ROADMAP, 'Tier_Name', row[SYNC.C_TARGET_UNIV]||'');
    var refStudy = counselGetRefData_(SYNC.TAB_REF_STUDY, 'Grade_Level', grade);
    var refLevel = counselGetRefLevel_();

    // 프롬프트 (Analysis.gs와 동일)
    var propData = surveyJson; var engData = isHS ? '생략' : engJson; var mathData = isHS ? '생략' : mathJson;
    var gradeInst = isHS ? '[고등학생 상담 모드] mockTest에 숫자만, 수능/내신 관점' : '[초중학생 상담 모드] mockTest 비워두기, 기초 학력 형성 초점';

    var prompt = '당신은 대한민국 상위 1% 입시 컨설턴트입니다.\n' + gradeInst + '\n\n[학생 데이터]\n1. 프로필: ' + JSON.stringify(profile) + '\n2. 모의고사: ' + JSON.stringify(mockScores) + '\n3. 성향검사: ' + propData + '\n4. 영어LT: ' + engData + '\n5. 수학LT: ' + mathData + '\n\n[참조]\n로드맵: ' + refRoadmap + '\n레벨기준: ' + refLevel + '\n학습방향: ' + refStudy + '\n\n위 데이터를 분석하여 상담 리포트 JSON을 작성하세요. responseMimeType: application/json';

    var apiKey = counselConfig_('GEMINI_API_KEY');
    var models = [counselConfig_('GEMINI_MODEL_PRIMARY')];
    try { var fb = counselConfig_('GEMINI_MODEL_FALLBACK'); if(fb) models.push(fb); } catch(e){}
    var temperature = 0.2; try { temperature = parseFloat(counselConfig_('GEMINI_TEMPERATURE')); } catch(e) {}

    var analysisResult = counselCallGemini_(prompt, apiKey, models, 5, temperature);
    var now = new Date(); var todayStr = counselToday_();
    analysisResult.id = Date.now().toString();
    analysisResult.name = studentName;
    analysisResult.date = todayStr;
    analysisResult.consultType = isHS ? 'HIGH_SCHOOL' : 'MIDDLE_SCHOOL';
    if (!analysisResult.basicProfile) analysisResult.basicProfile = {};
    analysisResult.basicProfile.targetUniv = row[SYNC.C_TARGET_UNIV] || '';

    var resultJson = JSON.stringify(analysisResult);
    var reportToken = Utilities.getUuid();

    // [OUT] 분석결과 저장
    var outSheet = counselSheet_(SYNC.TAB_OUT_ANALYSIS);
    outSheet.appendRow([studentId, studentName, isHS ? 'HIGH_SCHOOL' : 'MIDDLE_SCHOOL', todayStr, resultJson]);

    // [DB] 학생 마스터 업데이트
    var updates = {};
    updates[SYNC.C_ANALYSIS_DATE + 1] = todayStr;
    updates[SYNC.C_ANALYSIS_JSON + 1] = resultJson;
    updates[SYNC.C_REPORT_TOKEN + 1] = reportToken;
    updates[SYNC.C_STATUS + 1] = '분석완료';
    updates[SYNC.C_STEP_ANALYSIS + 1] = now;
    counselUpdateCells_(student.row_index, updates);

    return { ok: true, message: studentName + ' 학생 종합분석 완료!', token: reportToken, consultType: isHS ? 'HIGH_SCHOOL' : 'MIDDLE_SCHOOL' };
  } catch (e) {
    return { ok: false, error: '분석 실패: ' + e.message };
  }
}

function counselGetRefData_(tabName, filterCol, filterVal) {
  try {
    var sheet = counselSheet_(tabName);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return '데이터 없음';
    var headers = data[0];
    var colIdx = headers.indexOf(filterCol);
    if (colIdx === -1) colIdx = headers.findIndex(function(h){return String(h).indexOf(filterCol.substring(0,4))>-1;});
    if (colIdx === -1) return '필터 컬럼 없음';
    var filtered = [headers.join(',')];
    for (var i = 1; i < data.length; i++) {
      if (filterVal && String(data[i][colIdx]||'').toLowerCase().indexOf(String(filterVal).toLowerCase()) > -1) filtered.push(data[i].join(','));
    }
    return filtered.length > 1 ? filtered.join('\n') : data.map(function(r){return r.join(',');}).join('\n');
  } catch (e) { return '참조 데이터 없음'; }
}

function counselGetRefLevel_() {
  try { var sheet = counselSheet_(SYNC.TAB_REF_LEVEL); return sheet.getDataRange().getValues().map(function(r){return r.join(',');}).join('\n'); } catch(e){ return '레벨기준 없음'; }
}

function counselCallGemini_(prompt, apiKey, models, maxRetries, temperature) {
  var lastError = null;
  for (var m = 0; m < models.length; m++) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[m] + ':generateContent?key=' + apiKey;
    var payload = { contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: temperature, responseMimeType: 'application/json' } };
    var options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
    for (var attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        var response = UrlFetchApp.fetch(url, options);
        var code = response.getResponseCode();
        if (code === 200) {
          var json = JSON.parse(response.getContentText());
          if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) {
            var rawText = json.candidates[0].content.parts[0].text;
            return JSON.parse(rawText.replace(/```json/g,'').replace(/```/g,'').trim());
          }
        } else if (code === 429 || code === 503) { if (attempt < maxRetries) { Utilities.sleep(2000*Math.pow(2,attempt-1)); continue; } }
        else { lastError = new Error('HTTP ' + code); break; }
      } catch(e) { lastError = e; if (attempt < maxRetries) { Utilities.sleep(2000); continue; } }
    }
  }
  throw lastError || new Error('Gemini API 실패');
}
// ═══════════════════════════════════════════════════════════════
// [추가] 재원생 통합 검색 — CounselSync.gs 맨 끝에 붙여넣기
// ═══════════════════════════════════════════════════════════════

/**
 * 재원생 목록 조회 (teacher 메인 DB의 [DB] 학생 + [DB] 수강)
 * counsel 시트에 이미 있는 학생은 제외 (중복 방지)
 */
function tpCounselEnrolledStudents_(token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // [DB] 학생 읽기
    var stSheet = ss.getSheetByName(TP.SHEET_STUDENTS);
    var stData = stSheet.getDataRange().getValues();
    
    // [DB] 수강 읽기 — 학생별 수강 반/과목
    var enrollSheet = ss.getSheetByName(TP.SHEET_ENROLLMENT);
    var enrollData = enrollSheet.getDataRange().getValues();
    var classMap = {};
    for (var i = 1; i < enrollData.length; i++) {
      var sn = enrollData[i][0];
      if (!classMap[sn]) classMap[sn] = [];
      classMap[sn].push(String(enrollData[i][1]));
    }
    
    // [DB] 반 읽기 — 반명 → 과목 매핑
    var classSheet = ss.getSheetByName(TP.SHEET_CLASSES);
    var classData = classSheet.getDataRange().getValues();
    var subjMap = {};
    for (var i = 1; i < classData.length; i++) {
      subjMap[String(classData[i][0])] = String(classData[i][1] || '');
    }
    
    // counsel [DB] 학생 마스터에서 이름 목록 (중복 제거용)
    var counselSheet = counselSheet_(SYNC.COUNSEL_TAB);
    var counselData = counselSheet.getDataRange().getValues();
    var counselNames = {};
    for (var i = 1; i < counselData.length; i++) {
      var n = String(counselData[i][SYNC.C_NAME] || '').trim();
      if (n) counselNames[n] = true;
    }
    
    // 재원생 목록 구성 (퇴원 제외, counsel 미등록자만)
    var enrolled = [];
    for (var i = 1; i < stData.length; i++) {
      var name = String(stData[i][0] || '').trim();
      var status = String(stData[i][5] || '재원');
      if (!name || status === '퇴원') continue;
      if (counselNames[name]) continue; // 이미 counsel에 있으면 스킵
      
      var classes = classMap[name] || [];
      var subjects = [];
      classes.forEach(function(cn) {
        var subj = subjMap[cn];
        if (subj && subjects.indexOf(subj) < 0) subjects.push(subj);
      });
      
      enrolled.push({
        name: name,
        grade: String(stData[i][1] || ''),
        school: String(stData[i][2] || ''),
        parentPhone: String(stData[i][3] || ''),
        studentPhone: String(stData[i][4] || ''),
        subjects: subjects,
        classes: classes,
        source: 'enrolled'
      });
    }
    
    enrolled.sort(function(a, b) { return a.name.localeCompare(b.name, 'ko'); });
    return { ok: true, students: enrolled };
  } catch (e) {
    return { ok: false, error: '재원생 조회 실패: ' + e.message };
  }
}


/**
 * 재원생을 counsel [DB] 학생 마스터에 등록
 * 최초 검사/분석 시 한 번만 호출
 * 이후에는 일반 counsel 학생과 동일하게 처리됨
 */
function tpCounselRegisterEnrolled_(data, token) {
  if (!tpValidToken_(token)) return { ok: false, error: '인증 만료' };
  if (!data || !data.name) return { ok: false, error: '학생 이름 필수' };
  
  try {
    // 이미 등록되어 있는지 확인
    var sheet = counselSheet_(SYNC.COUNSEL_TAB);
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][SYNC.C_NAME] || '').trim() === String(data.name).trim()) {
        // 이미 있으면 기존 student_id 반환
        return { ok: true, studentId: String(existing[i][SYNC.C_ID]), message: '이미 등록된 학생입니다.', existing: true };
      }
    }
    
    // 새 행 생성
    var studentId = Date.now().toString();
    var now = new Date();
    var newRow = new Array(SYNC.C_TOTAL_COLS).fill('');
    
    newRow[SYNC.C_ID] = studentId;
    newRow[SYNC.C_CREATED] = now;
    newRow[SYNC.C_NAME] = data.name || '';
    newRow[SYNC.C_PHONE] = data.studentPhone || '';
    newRow[SYNC.C_PARENT_PHONE] = data.parentPhone || '';
    newRow[SYNC.C_SCHOOL] = data.school || '';
    newRow[SYNC.C_GRADE] = data.grade || '';
    newRow[SYNC.C_SUBJECT] = (data.subjects || []).join('+') || '';
    newRow[SYNC.C_ROUTE] = '재원생 과목추가';
    newRow[SYNC.C_CONSENT] = true;
    newRow[SYNC.C_STATUS] = '재원생등록';
    newRow[SYNC.C_STEP_CONSULT] = now;
    
    sheet.appendRow(newRow);
    
    return {
      ok: true,
      studentId: studentId,
      message: data.name + ' 학생이 상담 시스템에 등록되었습니다.',
      existing: false
    };
  } catch (e) {
    return { ok: false, error: '등록 실패: ' + e.message };
  }
}
