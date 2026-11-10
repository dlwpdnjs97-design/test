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
  FIRST_COL: 2,                   // 출납장 표가 시작하는 열 (A열=1, B열=2)

  // 적요·비목 자동 분류: 예산표 탭에서 항목 목록을 읽어와요
  BUDGET_SHEET: '2026 최종(고고고 변경내역)',   // 예산표 탭 이름 (정확히 똑같이)
  BUDGET_GROUP_COL: 2,            // 예산표에서 비목(기본운영비, 행사 …)이 있는 열 (B열=2)
  BUDGET_ITEM_COL: 3,             // 예산표에서 적요(오후간식, 생일 …)가 있는 열 (C열=3)

  // 적요를 고를 때 AI가 참고할 우리 부서 규칙. 말로 자유롭게 고치거나 더 적어도 돼요.
  CATEGORY_HINTS: [
    '빵집·제과점에서 산 빵은 오전간식(빵)',
    '과자·음료·우유 같은 간식은 오후간식',
    '생일 케이크·생일 선물은 생일',
    '교사 모임 식사나 교사용 물품은 교사 활동비',
    '어린이날 무렵(4월 말~5월 초) 선물은 어린이날, 12월 선물은 성탄절',
    '행사 이름 옆 날짜 앞뒤 3주 안의 구매는 그 행사 물품인지 먼저 생각하기',
  ],
};

/***** 아래는 고치지 않아도 돼요 *****/
const DONE_FOLDER_NAME = '처리완료';
const SKIP_FOLDERS = [DONE_FOLDER_NAME, '출력물'];   // 영수증 폴더로 보지 않는 폴더
const UNASSIGNED = '미지정';
const SUPPORTED = [
  'image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif',   // 아이폰 HEIC도 OK
  'application/pdf',
];
const MAX_FILE_MB = 14;           // 요청 한도(20MB)를 넘지 않도록

// 출납장 열 순서 (FIRST_COL 열부터). 맨 뒤 3칸은 확인용으로 추가한 칸이에요.
const HEADERS = ['지출결의서 연월', '증빙번호', '사용일자', '비목', '적요(사용내역)',
                 '지출금액(원)', '수입금액(원)', '거래 후 잔액', '사용구분', '사용자', '거래처/결재처',
                 '영수증 사진', '확인', 'AI 메모'];
const C = { MONTH: 1, PROOF: 2, DATE: 3, BIMOK: 4, JEOKYO: 5, OUT: 6, IN: 7, BALANCE: 8, USER: 10,
            VENDOR: 11, PHOTO: 12, CHECK: 13, MEMO: 14 };
const UNKNOWN_ITEM = '모름';
let budgetCache = null;           // 예산표 항목 목록 (한 번 실행에 한 번만 읽음)
const MAIN_COLS = 11;             // 원래 출납장 칸 수 (A~K)
let stopRun = false;              // 사용량 한도 등에 걸리면 이번 실행은 여기서 멈춤

/** 시트를 열면 위쪽 메뉴에 '🧾 영수증'이 생겨요 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧾 영수증')
    .addItem('지금 영수증 처리하기', 'processReceipts')
    .addItem('영수증 출력물 만들기', 'makeReceiptDoc')
    .addSeparator()
    .addItem('선생님 폴더에 항목 폴더 만들기', 'makeCategoryFolders')
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
      if (SKIP_FOLDERS.indexOf(folder.getName()) !== -1) continue;   // 처리완료·출력물 폴더는 건너뜀
      count += processTeacher_(sheet, folder, CONFIG.MAX_FILES_PER_RUN - count);
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

/** 선생님 폴더 하나 처리: 폴더에 바로 넣은 사진 + 항목 폴더(심방, 생일 …)에 넣은 사진 */
function processTeacher_(sheet, teacherFolder, limit) {
  const person = teacherFolder.getName();
  let n = processFolder_(sheet, teacherFolder, person, limit, null);   // 항목 폴더 밖 → AI가 적요 추측
  const subs = teacherFolder.getFolders();
  while (subs.hasNext() && n < limit && !stopRun) {
    const sub = subs.next();
    if (SKIP_FOLDERS.indexOf(sub.getName()) !== -1) continue;
    const item = matchBudgetItem_(sub.getName());
    n += processFolder_(sheet, sub, person, limit - n, item || { unmatched: sub.getName() });
  }
  return n;
}

/** fixed: 항목 폴더에서 온 사진이면 {group, item}, 이름이 안 맞는 폴더면 {unmatched}, 아니면 null */
function processFolder_(sheet, folder, person, limit, fixed) {
  const files = folder.getFiles();
  let done = null;
  let n = 0;
  while (files.hasNext() && n < limit) {
    const file = files.next();
    if (!handleFile_(sheet, file, person, fixed)) {    // 일시적 오류: 다음 실행 때 다시 시도
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
function handleFile_(sheet, file, person, fixed) {
  const mime = file.getMimeType();
  const rec = { month: '', proof: '', date: '', amount: '', person: person,
                bimok: '', jeokyo: '', vendor: '', link: file.getUrl(), memo: '', warn: true };

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

  const budget = getBudgetItems_();
  const folderItem = fixed && fixed.item ? fixed : null;
  let r;
  try {
    // 항목 폴더에서 온 사진은 적요가 이미 정해져 있으니 AI에게 묻지 않아요
    r = readReceipt_(blob, mime, folderItem ? [] : budget);
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
  rec.vendor = (r.store || '') + (r.items ? ' / ' + r.items : '');

  // 적요 → 비목: 항목 폴더가 있으면 그걸로, 없으면 AI가 고른 것으로
  const hit = folderItem || budget.filter(b => b.item === r.category)[0];
  if (hit) {
    rec.jeokyo = hit.item;
    rec.bimok = hit.group;
  }

  const warnings = [];
  if (r.confidence !== '높음') warnings.push('읽기 신뢰도 ' + (r.confidence || '?') + ': 꼭 확인해 주세요');
  if (!d) warnings.push('날짜를 못 읽었어요');
  if (!hhmm) warnings.push('결제 시각을 못 읽었어요(증빙번호 확인)');
  if (!rec.amount) warnings.push('금액을 못 읽었어요');
  if (fixed && fixed.unmatched) warnings.push('폴더 이름(' + fixed.unmatched + ')이 예산표 항목과 안 맞아서 AI가 적요를 추측했어요');
  if (!folderItem) {
    if (!budget.length) warnings.push('예산표 탭(' + CONFIG.BUDGET_SHEET + ')을 못 찾아 적요를 비워 뒀어요');
    else if (!hit) warnings.push('적요를 정하지 못했어요');
    else if (r.category_sure !== true) warnings.push('적요 확인 필요(AI 추측: ' + (r.category_reason || '') + ')');
  }
  if (r.note) warnings.push(r.note);
  if (isDuplicate_(sheet, rec)) warnings.push('⚠️ 중복 의심: 같은 증빙번호·금액이 이미 있어요');

  rec.warn = warnings.length > 0;
  rec.memo = warnings.join(' / ');

  writeRow_(sheet, rec);
  return true;
}

/** Gemini에게 영수증을 보내서 정해진 형식(JSON)으로 받아오기 */
function readReceipt_(blob, mime, budget) {
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
  if (budget.length) {
    schema.properties.category = { type: 'STRING', format: 'enum',
      enum: budget.map(b => b.item).concat([UNKNOWN_ITEM]),
      description: '이 지출이 해당하는 예산 적요. 판단이 어려우면 모름' };
    schema.properties.category_sure = { type: 'BOOLEAN', description: '적요를 확신하면 true' };
    schema.properties.category_reason = { type: 'STRING', description: '그 적요를 고른 이유를 한 줄로' };
    ['category', 'category_sure', 'category_reason'].forEach(k => {
      schema.required.push(k);
      schema.propertyOrdering.push(k);
    });
  }

  const prompt =
    '교회 초등부에서 쓴 비용의 영수증입니다. 영수증에서 결제 날짜, 결제 시각, 가게 이름, ' +
    '최종 결제 금액, 산 물건을 읽어 주세요.\n' +
    '- 결제 시각은 영수증에 찍힌 거래 시각을 24시간제로 쓰세요. 카드 전표라면 승인 시각을 쓰세요.\n' +
    '- 금액은 부가세를 포함한 최종 결제 금액(받을 금액, 합계, 결제 금액)을 쓰세요. 거스름돈이나 받은 돈이 아니에요.\n' +
    '- 글자가 흐리거나 잘려서 확실하지 않으면 confidence를 낮추고 note에 어떤 부분이 불확실한지 쓰세요.\n' +
    '- 영수증이 아닌 사진이면 confidence를 낮음으로 하고 note에 그렇게 쓰세요.\n' +
    '- 지어내지 마세요. 안 보이는 값은 빈 문자열이나 0으로 두세요.' +
    (budget.length ? categoryPrompt_(budget) : '');

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
  values[C.BIMOK - 1] = rec.bimok;
  values[C.JEOKYO - 1] = rec.jeokyo;
  values[C.VENDOR - 1] = rec.vendor;
  values[C.PHOTO - 1] = rec.link;
  values[C.CHECK - 1] = false;
  values[C.MEMO - 1] = rec.memo;

  sheet.getRange(r, col_(C.PROOF)).setNumberFormat('@');      // 증빙번호는 글자 그대로
  sheet.getRange(r, col_(1), 1, HEADERS.length).setValues([values]);
  sheet.getRange(r, col_(C.MONTH)).setNumberFormat('yyyy.mm');
  sheet.getRange(r, col_(C.DATE)).setNumberFormat('mm/dd');
  sheet.getRange(r, col_(C.OUT), 1, 3).setNumberFormat('#,##0');
  // 거래 후 잔액 = 윗줄 잔액 + 수입 - 지출
  sheet.getRange(r, col_(C.BALANCE)).setFormulaR1C1('=N(R[-1]C)+N(RC[-1])-N(RC[-2])');
  sheet.getRange(r, col_(C.CHECK)).insertCheckboxes();
  sheet.getRange(r, col_(1), 1, MAIN_COLS).setBorder(true, true, true, true, true, false);   // 표 테두리
  if (rec.warn) {
    sheet.getRange(r, col_(1), 1, HEADERS.length).setBackground('#fff4d6');   // 확인 필요한 줄은 노란색
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, col_(1), 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true);
    sheet.getRange(1, col_(1), 1, MAIN_COLS).setBackground('#d9d9d9');
    sheet.setFrozenRows(1);
  }
  // 이미 있던 출납장이면 오른쪽 확인용 3칸 제목만 채워요
  const extra = sheet.getRange(1, col_(MAIN_COLS + 1), 1, HEADERS.length - MAIN_COLS);
  if (extra.getValues()[0].join('') === '') {
    extra.setValues([HEADERS.slice(MAIN_COLS)])
      .setFontWeight('bold').setHorizontalAlignment('center').setBackground('#f2c230')
      .setBorder(true, true, true, true, true, true);
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

/** 예산표 탭에서 [비목, 적요] 목록 읽기. '월별구분' 줄 아래부터 '월계' 줄 전까지 */
function getBudgetItems_() {
  if (budgetCache) return budgetCache;
  budgetCache = [];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.BUDGET_SHEET);
  if (!sh) {
    Logger.log('예산표 탭을 찾을 수 없어요: ' + CONFIG.BUDGET_SHEET);
    return budgetCache;
  }
  const g = CONFIG.BUDGET_GROUP_COL, it = CONFIG.BUDGET_ITEM_COL;
  const left = Math.min(g, it);
  const rows = sh.getRange(1, left, sh.getLastRow(), Math.abs(g - it) + 1).getValues();
  let started = false;
  let group = '';
  for (const row of rows) {
    const gv = String(row[g - left]).trim();
    const iv = String(row[it - left]).trim();
    if (!started) { if (gv === '월별구분') started = true; continue; }
    if (gv === '월계') break;
    if (gv) group = gv;                         // 합쳐진 칸은 첫 줄에만 값이 있어서 이어서 씀
    if (iv) budgetCache.push({ group: group, item: iv });
  }
  if (!budgetCache.length) Logger.log('예산표에서 항목을 못 읽었어요. 월별구분/월계 줄과 열 번호를 확인해 주세요.');
  return budgetCache;
}

/** 폴더 이름을 예산표 적요와 맞춰 보기. 띄어쓰기와 괄호 속 날짜는 무시 (예: '여름성경학교' = '여름성경학교(7/24금~25토)') */
function matchBudgetItem_(name) {
  const norm = t => String(t).replace(/\(.*?\)/g, '').replace(/\s+/g, '');
  const key = norm(name);
  if (!key) return null;
  return getBudgetItems_().filter(b => b.item === name || norm(b.item) === key)[0] || null;
}

/** 폴더 이름으로 쓸 짧은 이름 (괄호 속 날짜 빼기) */
function shortItemName_(item) {
  return item.replace(/\(.*?\)/g, '').trim() || item;
}

/** 메뉴에서 실행: 모든 선생님 폴더 안에 예산 항목 폴더를 만들어요 (이미 있는 건 건너뜀) */
function makeCategoryFolders() {
  const budget = getBudgetItems_();
  const ui = SpreadsheetApp.getUi();
  if (!budget.length) {
    ui.alert('예산표 탭(' + CONFIG.BUDGET_SHEET + ')에서 항목을 못 읽었어요. 코드의 BUDGET_SHEET 이름을 확인해 주세요.');
    return;
  }
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const people = root.getFolders();
  const made = [];
  while (people.hasNext()) {
    const teacher = people.next();
    if (SKIP_FOLDERS.indexOf(teacher.getName()) !== -1) continue;
    const have = [];
    const subs = teacher.getFolders();
    while (subs.hasNext()) {
      const hit = matchBudgetItem_(subs.next().getName());
      if (hit) have.push(hit.item);
    }
    let n = 0;
    budget.forEach(b => {
      if (have.indexOf(b.item) === -1) { teacher.createFolder(shortItemName_(b.item)); n++; }
    });
    made.push(teacher.getName() + ': ' + n + '개');
  }
  ui.alert('항목 폴더를 만들었어요.\n' + (made.join('\n') || '선생님 폴더가 없어요.'));
}

/** 적요 고르기 안내문 */
function categoryPrompt_(budget) {
  return '\n\n이 지출이 아래 예산 항목 중 어디에 해당하는지 category에 골라 주세요. ' +
    '결제 날짜와 산 물건을 함께 보고 판단하세요.\n' +
    budget.map(b => '- ' + b.item + ' (' + b.group + ')').join('\n') +
    '\n판단 규칙:\n' + CONFIG.CATEGORY_HINTS.map(h => '- ' + h).join('\n') +
    '\n- 여러 항목이 가능하거나 근거가 약하면 가장 가까운 것을 고르되 category_sure를 false로 하세요.' +
    '\n- 전혀 모르겠으면 ' + UNKNOWN_ITEM + '를 고르세요.';
}

/** 표 안의 몇 번째 칸(1부터)을 시트의 실제 열 번호로 바꾸기 */
function col_(c) {
  return c + CONFIG.FIRST_COL - 1;
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
  const values = sheet.getRange(2, col_(C.PROOF), last - 1, C.OUT - C.PROOF + 1).getValues();
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
