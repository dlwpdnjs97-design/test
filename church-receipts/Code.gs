/******************************************************
 * 초등부 영수증 자동 정리 (Google Apps Script + Gemini 무료 API)
 * 드라이브 폴더에 올린 영수증 사진을 Gemini가 읽어서
 * '출납장' 시트에 한 줄씩 채워 넣어요.
 ******************************************************/

/***** ① 설정: 여기만 바꾸면 돼요 *****/
const CONFIG = {
  ROOT_FOLDER_ID: '여기에_폴더_ID를_붙여넣으세요',
  SHEET_NAME: '출납장',
  MODEL: 'gemini-3.5-flash-lite',
  MAX_FILES_PER_RUN: 8,           // 무료 사용량 제한 때문에 한 번에 조금씩 처리
  WAIT_BETWEEN_MS: 6000,          // 영수증 사이에 6초 쉬기 (분당 요청 제한 대비)
};

/***** 아래는 고치지 않아도 돼요 *****/
const DONE_FOLDER_NAME = '처리완료';
const UNASSIGNED = '미지정';
const SUPPORTED = [
  'image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif',   // 아이폰 HEIC도 OK
  'application/pdf',
];
const MAX_FILE_MB = 14;           // 요청 한도(20MB)를 넘지 않도록

// 출납장 열 순서 (A열부터). 맨 뒤 3칸은 확인용으로 추가한 칸이에요.
const HEADERS = ['지출결의서 연월', '증빙번호', '사용일자', '비목', '적요(사용내역)',
                 '지출금액(원)', '수입금액(원)', '거래 후 잔액', '사용구분', '사용자', '거래처/결재처',
                 '영수증 사진', '확인', 'AI 메모'];
const C = { MONTH: 1, PROOF: 2, DATE: 3, OUT: 6, IN: 7, BALANCE: 8, USER: 10,
            PHOTO: 12, CHECK: 13, MEMO: 14 };
const MAIN_COLS = 11;             // 원래 출납장 칸 수 (A~K)
let stopRun = false;              // 사용량 한도 등에 걸리면 이번 실행은 여기서 멈춤

/** 시트를 열면 위쪽 메뉴에 '🧾 영수증'이 생겨요 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧾 영수증')
    .addItem('지금 영수증 처리하기', 'processReceipts')
    .addItem('영수증 출력물 만들기', 'makeReceiptDoc')
    .addToUi();
}

/** 처음 한 번만 실행: 시트 만들기 + 10분마다 자동 실행 등록 */
function setup() {
  getSheet_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processReceipts')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processReceipts').timeBased().everyMinutes(10).create();
  Logger.log('설정 완료! 이제 10분마다 새 영수증을 확인해요.');
}

/** 폴더에서 새 영수증을 찾아 처리 */
function processReceipts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {                 // 다른 실행이 끝날 때까지 30초 기다려 봄
    Logger.log('다른 실행이 아직 진행 중이라 이번에는 건너뛰어요.');
    return;
  }
  try {
    const sheet = getSheet_();
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    let count = 0;

    // 선생님별 폴더 (폴더 이름 = 사용자)
    const people = root.getFolders();
    while (people.hasNext() && count < CONFIG.MAX_FILES_PER_RUN && !stopRun) {
      const folder = people.next();
      if (folder.getName() === DONE_FOLDER_NAME || folder.getName() === '출력물') continue;   // 처리완료·출력물 폴더는 건너뜀
      count += processFolder_(sheet, folder, folder.getName(), CONFIG.MAX_FILES_PER_RUN - count);
    }
    // 최상위 폴더에 바로 올린 사진
    if (count < CONFIG.MAX_FILES_PER_RUN && !stopRun) {
      count += processFolder_(sheet, root, UNASSIGNED, CONFIG.MAX_FILES_PER_RUN - count);
    }
    Logger.log('이번에 처리한 영수증: ' + count + '장');
  } finally {
    lock.releaseLock();
  }
}

function processFolder_(sheet, folder, person, limit) {
  const files = folder.getFiles();
  let done = null;
  let n = 0;
  while (files.hasNext() && n < limit) {
    const file = files.next();
    if (!handleFile_(sheet, file, person)) {           // 일시적 오류: 다음 실행 때 다시 시도
      stopRun = true;
      break;
    }
    done = done || getOrCreateSub_(folder, DONE_FOLDER_NAME);
    file.moveTo(done);
    n++;
  }
  return n;
}

/** 파일 하나 처리. 시트에 기록했으면 true, 다음에 다시 시도할 거면 false */
function handleFile_(sheet, file, person) {
  const mime = file.getMimeType();
  const rec = { month: '', proof: '', date: '', amount: '', person: person,
                link: file.getUrl(), memo: '', warn: true };

  if (SUPPORTED.indexOf(mime) === -1) {
    rec.memo = '⚠️ 읽을 수 없는 파일 형식(' + mime + '). 사진(JPG, PNG)이나 PDF로 올려주세요.';
    writeRow_(sheet, rec);
    return true;
  }
  const blob = file.getBlob();
  if (blob.getBytes().length > MAX_FILE_MB * 1024 * 1024) {
    rec.memo = '⚠️ 파일이 너무 커요(' + MAX_FILE_MB + 'MB 초과). 사진 크기를 줄여서 다시 올려주세요.';
    writeRow_(sheet, rec);
    return true;
  }

  let r;
  try {
    r = readReceipt_(blob, mime);
  } catch (e) {
    if (e.retryable) { Logger.log('다음에 다시 시도: ' + e.message); return false; }
    rec.memo = '⚠️ 자동 읽기 실패: ' + e.message;
    writeRow_(sheet, rec);
    return true;
  } finally {
    Utilities.sleep(CONFIG.WAIT_BETWEEN_MS);
  }

  const d = parseDate_(r.date);
  const hhmm = parseTime_(r.time);
  rec.date = d || '';
  rec.month = d ? new Date(d.getFullYear(), d.getMonth(), 1) : '';
  rec.proof = (d ? pad2_(d.getMonth() + 1) + pad2_(d.getDate()) : '????') + ' ' + (hhmm || '????');
  rec.amount = Number(r.total) || '';

  const warnings = [];
  if (r.confidence !== '높음') warnings.push('읽기 신뢰도 ' + (r.confidence || '?') + ': 꼭 확인해 주세요');
  if (!d) warnings.push('날짜를 못 읽었어요');
  if (!hhmm) warnings.push('결제 시각을 못 읽었어요(증빙번호 확인)');
  if (!rec.amount) warnings.push('금액을 못 읽었어요');
  if (r.note) warnings.push(r.note);
  if (isDuplicate_(sheet, rec)) warnings.push('⚠️ 중복 의심: 같은 증빙번호·금액이 이미 있어요');

  rec.warn = warnings.length > 0;
  const ref = '참고: ' + (r.store || '구매처 모름') + (r.items ? ' / ' + r.items : '');
  rec.memo = warnings.concat([ref]).join(' / ');

  writeRow_(sheet, rec);
  return true;
}

/** Gemini에게 영수증을 보내서 정해진 형식(JSON)으로 받아오기 */
function readReceipt_(blob, mime) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('API 키가 없어요. 프로젝트 설정 → 스크립트 속성에 GEMINI_API_KEY를 넣어주세요.');

  const schema = {
    type: 'OBJECT',
    properties: {
      date: { type: 'STRING', description: '결제 날짜. YYYY-MM-DD 형식. 안 보이면 빈 문자열' },
      time: { type: 'STRING', description: '결제 시각. 24시간제 HH:MM 형식(예: 오후 11시 10분 → 23:10). 안 보이면 빈 문자열' },
      store: { type: 'STRING', description: '가게 이름. 안 보이면 빈 문자열' },
      total: { type: 'INTEGER', description: '실제로 결제한 최종 금액(원). 할인 반영. 모르면 0' },
      items: { type: 'STRING', description: '산 물건을 짧게 요약. 예: 초코파이 2상자, 우유 24개' },
      confidence: { type: 'STRING', format: 'enum', enum: ['높음', '보통', '낮음'] },
      note: { type: 'STRING', description: '사람이 확인해야 할 점. 없으면 빈 문자열' },
    },
    required: ['date', 'time', 'store', 'total', 'items', 'confidence', 'note'],
    propertyOrdering: ['date', 'time', 'store', 'total', 'items', 'confidence', 'note'],
  };

  const prompt =
    '교회 초등부에서 쓴 비용의 영수증입니다. 영수증에서 결제 날짜, 결제 시각, 가게 이름, ' +
    '최종 결제 금액, 산 물건을 읽어 주세요.\n' +
    '- 결제 시각은 영수증에 찍힌 거래 시각을 24시간제로 쓰세요. 카드 전표라면 승인 시각을 쓰세요.\n' +
    '- 금액은 부가세를 포함한 최종 결제 금액(받을 금액, 합계, 결제 금액)을 쓰세요. 거스름돈이나 받은 돈이 아니에요.\n' +
    '- 글자가 흐리거나 잘려서 확실하지 않으면 confidence를 낮추고 note에 어떤 부분이 불확실한지 쓰세요.\n' +
    '- 영수증이 아닌 사진이면 confidence를 낮음으로 하고 note에 그렇게 쓰세요.\n' +
    '- 지어내지 마세요. 안 보이는 값은 빈 문자열이나 0으로 두세요.';

  const body = {
    contents: [{
      role: 'user',
      parts: [
        { inline_data: { mime_type: mime, data: Utilities.base64Encode(blob.getBytes()) } },
        { text: prompt },
      ],
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: schema,
      temperature: 0,
    },
  };

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
              encodeURIComponent(CONFIG.MODEL) + ':generateContent';
  let res;
  try {
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
  } catch (e) {
    throw retryable_('인터넷 연결 문제: ' + e.message);
  }

  const code = res.getResponseCode();
  let json = null;
  try { json = JSON.parse(res.getContentText()); } catch (e) { /* JSON이 아닌 응답 */ }
  const apiMsg = json && json.error ? json.error.message : res.getContentText().slice(0, 200);

  if (code === 429) throw retryable_('무료 사용량 한도에 걸렸어요. 잠시 후 자동으로 다시 시도해요.');
  if (code >= 500) throw retryable_('AI 서버가 바빠요(' + code + ')');
  if (code === 404) throw retryable_('모델(' + CONFIG.MODEL + ')을 쓸 수 없어요. 코드 맨 위 MODEL을 바꿔 주세요. 구글 메시지: ' + apiMsg);
  if (code !== 200) throw new Error('API 오류 ' + code + ': ' + apiMsg);

  if (json.promptFeedback && json.promptFeedback.blockReason) {
    throw new Error('AI가 이 파일 읽기를 거절했어요(' + json.promptFeedback.blockReason + '). 직접 입력해 주세요.');
  }
  const cand = json.candidates && json.candidates[0];
  if (!cand || !cand.content || !cand.content.parts) {
    throw new Error('AI 응답이 비어 있어요' + (cand && cand.finishReason ? '(' + cand.finishReason + ')' : ''));
  }
  if (cand.finishReason === 'MAX_TOKENS') throw new Error('AI 응답이 중간에 잘렸어요.');

  const text = cand.content.parts.filter(p => p.text && !p.thought).map(p => p.text).join('');
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error('AI 응답을 읽지 못했어요: ' + text.slice(0, 100));
  }
}

/***** 시트에 쓰기 *****/
function writeRow_(sheet, rec) {
  const r = sheet.getLastRow() + 1;
  const values = new Array(HEADERS.length).fill('');
  values[C.MONTH - 1] = rec.month;
  values[C.PROOF - 1] = rec.proof;
  values[C.DATE - 1] = rec.date;
  values[C.OUT - 1] = rec.amount;
  values[C.USER - 1] = rec.person;
  values[C.PHOTO - 1] = rec.link;
  values[C.CHECK - 1] = false;
  values[C.MEMO - 1] = rec.memo;

  sheet.getRange(r, C.PROOF).setNumberFormat('@');            // 증빙번호는 글자 그대로
  sheet.getRange(r, 1, 1, HEADERS.length).setValues([values]);
  sheet.getRange(r, C.MONTH).setNumberFormat('yyyy.mm');
  sheet.getRange(r, C.DATE).setNumberFormat('mm/dd');
  sheet.getRange(r, C.OUT, 1, 3).setNumberFormat('#,##0');
  // 거래 후 잔액 = 윗줄 잔액 + 수입 - 지출
  sheet.getRange(r, C.BALANCE).setFormulaR1C1('=N(R[-1]C)+N(RC[-1])-N(RC[-2])');
  sheet.getRange(r, C.CHECK).insertCheckboxes();
  if (rec.warn) {
    sheet.getRange(r, 1, 1, HEADERS.length).setBackground('#fff4d6');   // 확인 필요한 줄은 노란색
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true);
    sheet.getRange(1, 1, 1, MAIN_COLS).setBackground('#d9d9d9');
    sheet.getRange(1, MAIN_COLS + 1, 1, HEADERS.length - MAIN_COLS).setBackground('#f2c230');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/***** 도우미 함수들 *****/
function retryable_(msg) {
  const e = new Error(msg);
  e.retryable = true;
  return e;
}

function getOrCreateSub_(folder, name) {
  const it = folder.getFoldersByName(name);
  return it.hasNext() ? it.next() : folder.createFolder(name);
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

function parseDate_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function parseTime_(s) {
  const m = /^(\d{1,2}):(\d{2})/.exec(s || '');
  if (!m || Number(m[1]) > 23 || Number(m[2]) > 59) return '';
  return pad2_(Number(m[1])) + m[2];
}

function isDuplicate_(sheet, rec) {
  const last = sheet.getLastRow();
  if (last < 2 || !rec.amount || rec.proof.indexOf('????') === 0) return false;
  const values = sheet.getRange(2, C.PROOF, last - 1, C.OUT - C.PROOF + 1).getValues();
  return values.some(v => String(v[0]) === rec.proof && Number(v[C.OUT - C.PROOF]) === rec.amount);
}

/***** 문제 찾기용 함수들 *****/
/** 폴더 안에 무엇이 보이는지 실행 로그에 보여줘요 */
function checkSetup() {
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  Logger.log('바깥 폴더 이름: ' + root.getName());
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    Logger.log('📂 ' + f.getName());
    const files = f.getFiles();
    let n = 0;
    while (files.hasNext()) {
      const file = files.next();
      n++;
      Logger.log('   - ' + file.getName() + ' (' + file.getMimeType() + ')');
    }
    if (n === 0) Logger.log('   (파일 없음)');
  }
  const loose = root.getFiles();
  while (loose.hasNext()) {
    const file = loose.next();
    Logger.log('바깥 폴더에 바로 있는 파일: ' + file.getName() + ' (' + file.getMimeType() + ')');
  }
  Logger.log('API 키 등록됨: ' + Boolean(PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')));
}

/** 이 API 키로 쓸 수 있는 모델 이름을 보여줘요 */
function listModels() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true,
  });
  const json = JSON.parse(res.getContentText());
  if (!json.models) { Logger.log('오류: ' + res.getContentText()); return; }
  json.models
    .filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') !== -1)
    .filter(m => m.name.indexOf('flash') !== -1)
    .forEach(m => Logger.log(m.name.replace('models/', '') + '   ← ' + (m.displayName || '')));
}
