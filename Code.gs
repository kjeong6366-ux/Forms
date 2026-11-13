/** 충청대학교 AI·DX 교수 설문 / Google Apps Script 서버 */
const APP = { sheet: '응답원본', settings: '설정' };
const HEADERS = ['접수시간','응답ID','소속분야','교원구분','교육경력','담당교과목유형', ...Array.from({length:10},(_,i)=>`Q${i+1}`)];
const MULTI = [4,5,6,7,9];
const OPTIONS = {
  1:['활용하지 않음','가끔 활용','일부 수업 활용','정기적 활용','교육 전반에 활용'],
  2:['매우 낮음','낮음','보통','높음','매우 높음'],
  3:['매우 낮음','낮음','보통','높음','매우 높음'],
  4:['AI 기초교양','전공별 X+AI 융합교과','AI 직무모듈·마이크로디그리','산학협력 프로젝트','AI 비교과교육'],
  5:['AI 활용 수업설계','개인별 맞춤형 학습','AI 기반 프로젝트학습','AI 활용 피드백·평가','AI 협력형 문제해결학습'],
  6:['교수자의 전문역량 부족','교육콘텐츠 부족','실습환경·SW 부족','교육과정 개편의 어려움','학습자 역량 차이','기타'],
  7:['단계별 AI 연수','산업체 현장연수','교과목 개발비·운영지원','AI 도구·실습환경','전문가 공동수업 지원'],
  8:['전혀 없음','낮음','보통','높음','매우 높음'],
  9:['학생 AI 역량 향상','전공직무 활용능력','교육과정 개발·운영 성과','산학 프로젝트 성과','취업·산업체 만족도']
};
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('충청대학교 AI·DX 교수 설문조사')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function db_() {
  const prop = PropertiesService.getScriptProperties();
  let id = prop.getProperty('SURVEY_SHEET_ID');
  let ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('충청대학교_AI_DX_교수설문_응답');
  if (!id) prop.setProperty('SURVEY_SHEET_ID', ss.getId());
  let ws = ss.getSheetByName(APP.sheet) || ss.insertSheet(APP.sheet);
  if (ws.getLastRow() === 0) {
    ws.appendRow(HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1,1,1,HEADERS.length).setBackground('#183A71').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return {ss,ws};
}
function initializeSurvey() {
  const db=db_();
  return { spreadsheetUrl:db.ss.getUrl(), surveyUrl:ScriptApp.getService().getUrl() || '' };
}
function submitSurvey(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('잘못된 요청입니다.');
  const profile=payload.profile || {};
  const validProfile={
    field:['자연계열','공학계열','인문사회계열','예체능계열'],
    staff:['전임','비전임'],
    career:['5년 미만','5~10년 미만','10~20년 미만','20년 이상'],
    course:['전공','교양','전공+교양']
  };
  Object.keys(validProfile).forEach(k=>{if(!validProfile[k].includes(profile[k])) throw new Error('기본정보를 모두 작성해 주세요.');});
  const answers=payload.answers || {};
  for(let n=1;n<=10;n++) {
    const a=answers[n];
    if(n===10){if(typeof a!=='string'||!a.trim()||a.length>2000)throw new Error('10번 서술형 답변을 입력해 주세요(2000자 이하).');continue;}
    if(MULTI.includes(n)){
      if(!Array.isArray(a)||a.length!==2||new Set(a).size!==2||a.some(x=>!OPTIONS[n].includes(x))) throw new Error(`${n}번은 서로 다른 항목 2개를 선택해 주세요.`);
    } else if(!OPTIONS[n].includes(a)) throw new Error(`${n}번 응답을 확인해 주세요.`);
  }
  const lock=LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const {ws}=db_();
    const id=Utilities.getUuid();
    const value=n=>Array.isArray(answers[n])?answers[n].join(' | '):String(answers[n]).trim();
    ws.appendRow([new Date(),id,profile.field,profile.staff,profile.career,profile.course,...Array.from({length:10},(_,i)=>value(i+1))]);
    return {ok:true,id};
  } finally {lock.releaseLock();}
}
function verifyAdmin_(pin) {
  const saved=PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if(!saved) throw new Error('관리자 PIN 미설정: Apps Script 프로젝트 설정 > 스크립트 속성에 ADMIN_PIN을 등록해 주세요.');
  if(String(pin)!==saved) throw new Error('관리자 인증에 실패했습니다.');
}
function getReport(pin) {
  verifyAdmin_(pin);
  const {ws,ss}=db_();
  const data=ws.getDataRange().getDisplayValues();
  return { headers:HEADERS, rows:data.slice(1), count:Math.max(0,data.length-1), spreadsheetUrl:ss.getUrl() };
}
