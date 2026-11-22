/******************************************************
 * 월별 지출결의서 만들기
 * '지출결의, 증빙내역' 탭을 복사해서 'N월' 탭을 만들고
 *  - 왼쪽 지출결의서: 날짜, 청구 금액(그 달 예산), 일금, 지급목적,
 *    예산표 탭을 옮겨 그린 월별 예산표(그림 대신 진짜 칸)
 *  - 오른쪽 증빙내역: 출납장에서 고른 행들, 지출합계, 월 예산, 잔액
 * 을 채워요. 칸 위치는 제목 글자('증빙 번호', '지출합계' …)를 찾아서 정해요.
 ******************************************************/

const REPORT = {
  TEMPLATE_SHEET: '지출결의, 증빙내역',   // 복사할 서식 탭
  // 왼쪽 예산표를 그릴 열 (서식 탭 기준)
  TABLE_GROUP_COLS: 'C',                  // 비목(기본운영비 …)
  TABLE_ITEM_COLS: 'D:F',                 // 적요(오후간식 …)
  TABLE_MONTH_COLS: 'G:H',                // 그 달 금액
  TABLE_FONT_SIZE: 9,
  ALWAYS_GROUPS: ['기본운영비'],           // 그 달 예산이 없어도 항상 표에 넣을 비목
};

/** 메뉴에서 실행: 몇 월인지, 출납장 몇 행을 넣을지 물어보고 탭을 만들어요 */
function makeMonthlyReport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();

  // 출납장에서 행을 선택해 둔 상태면 그 범위를 기본값으로
  const ledger = getSheet_();
  const sel = ss.getActiveSheet().getName() === CONFIG.SHEET_NAME ? ss.getActiveRange() : null;
  const selText = sel && sel.getLastRow() >= 2 ? Math.max(2, sel.getRow()) + '-' + sel.getLastRow() : '';

  const a1 = ui.prompt('월별 지출결의서 만들기 (1/2)',
    '몇 월 지출결의서인가요? 예: 11 (비워 두면 이번 달)', ui.ButtonSet.OK_CANCEL);
  if (a1.getSelectedButton() !== ui.Button.OK) return;
  const t1 = a1.getResponseText().replace(/월/g, '').trim();
  const month = t1 ? Number(t1) : now.getMonth() + 1;
  if (!(month >= 1 && month <= 12)) { ui.alert('1부터 12 사이 숫자로 써 주세요.'); return; }
  const year = now.getFullYear();

  const last = ledger.getLastRow();
  const a2 = ui.prompt('월별 지출결의서 만들기 (2/2)',
    '증빙내역에 넣을 출납장 행을 써 주세요. 예: 2-15\n' +
    (selText ? '비워 두면 지금 선택한 ' + selText + '행을 넣어요.'
             : '비워 두면 "지출결의서 연월"이 ' + year + '.' + pad2_(month) + '인 행을 넣어요.') +
    '\n(출납장은 2행부터 ' + last + '행까지 있어요)', ui.ButtonSet.OK_CANCEL);
  if (a2.getSelectedButton() !== ui.Button.OK) return;
  const t2 = a2.getResponseText().trim() || selText;
  let range = null;
  if (t2) {
    range = parseRowRange_(t2, last);
    if (!range) { ui.alert('2-15 처럼 "시작행-끝행"으로 써 주세요. 출납장은 2행부터 ' + last + '행까지 있어요.'); return; }
  }

  const name = month + '월';
  const old = ss.getSheetByName(name);
  if (old) {
    const yes = ui.alert(name + ' 탭이 이미 있어요. 지우고 새로 만들까요?\n(직접 고친 내용도 사라져요)', ui.ButtonSet.YES_NO);
    if (yes !== ui.Button.YES) return;
    ss.deleteSheet(old);
  }
  const result = buildMonthlyReport_(year, month, range);
  ss.setActiveSheet(result.sheet);
  ui.alert(name + ' 지출결의서를 만들었어요.\n' +
           '· 증빙 ' + result.count + '건, 지출 ' + won_(result.spent) + '\n' +
           '· ' + name + ' 예산(청구 금액) ' + won_(result.budget) +
           (result.moved ? '\n· 출납장 ' + result.moved + '줄의 "지출결의서 연월"을 ' + year + '.' + pad2_(month) + '로 바꿨어요' : '') +
           (result.notes.length ? '\n\n확인해 주세요:\n· ' + result.notes.join('\n· ') : ''));
}

/** 실제로 탭을 만드는 부분. range가 없으면 '지출결의서 연월'이 그 달인 행을 써요 */
function buildMonthlyReport_(year, month, range) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tpl = ss.getSheetByName(REPORT.TEMPLATE_SHEET);
  if (!tpl) throw new Error('서식 탭(' + REPORT.TEMPLATE_SHEET + ')을 찾을 수 없어요.');
  const sh = tpl.copyTo(ss).setName(month + '월');
  const prev = ss.getSheetByName((month - 1) + '월');      // 앞 달 탭 바로 뒤에 놓기
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(prev ? prev.getIndex() + 1 : ss.getNumSheets());

  const notes = [];
  const today = new Date();
  const all = getBudgetItems_();
  if (!all.length) notes.push('예산표 탭(' + setting_('BUDGET_SHEET') + ')을 못 읽어 예산을 0으로 뒀어요');
  const items = all.filter(b => (b.months[month] || 0) > 0 || REPORT.ALWAYS_GROUPS.indexOf(b.group) !== -1);
  const budget = all.reduce((a, b) => a + (b.months[month] || 0), 0);

  // ── 왼쪽: 날짜, 지급목적 ──
  findAll_(sh, '\\d{4}년\\s*\\d{1,2}월\\s*\\d{1,2}일').forEach(c => {
    const spaced = /년\s/.test(String(c.getDisplayValue()));
    c.setNumberFormat('@').setValue(Utilities.formatDate(today, Session.getScriptTimeZone(),
      spaced ? 'yyyy년 MM월 dd일' : 'yyyy년MM월dd일'));
  });
  const purpose = findOne_(sh, '지급목적');
  if (purpose) purpose.setValue(String(purpose.getValue()).replace(/\d{1,2}\s*월/, month + '월'));

  // ── 왼쪽: 월별 예산표 (그림 대신 진짜 칸으로) ──
  if (purpose) drawBudgetTable_(sh, purpose.getRow() + 1, month, items, notes);
  else notes.push('"지급목적" 칸을 못 찾아 왼쪽 예산표를 그리지 못했어요');

  // ── 왼쪽: 청구 금액(₩), 일금, 'N월 예산' ──
  const wonCell = findOne_(sh, '^\\s*₩\\s*$');
  if (wonCell) rightOf_(sh, wonCell).setValue(budget); else notes.push('₩ 칸을 못 찾았어요');
  const ilgum = findOne_(sh, '^\\s*일\\s*금\\s*$');
  if (ilgum) rightOf_(sh, ilgum).setValue(koreanMoney_(budget)); else notes.push('일금 칸을 못 찾았어요');
  const budgetLabel = findOne_(sh, '^\\s*\\d{1,2}\\s*월\\s*예산\\s*$');
  if (budgetLabel) {
    budgetLabel.setValue(month + '월 예산');
    sh.getRange(budgetLabel.getRow() + 1, budgetLabel.getColumn()).setValue(budget);
  }

  // ── 오른쪽: 증빙내역 ──
  const picked = pickLedgerRows_(year, month, range);
  const spent = fillProofTable_(sh, picked.rows, budget, today, notes);

  // 고른 행의 '지출결의서 연월'을 이 결의서 달로 맞추기 (늦게 처리한 지난달 영수증 등)
  let moved = 0;
  if (range) {
    const ledger = getSheet_();
    const monthDate = new Date(year, month - 1, 1);
    picked.sheetRows.forEach(r => {
      const cell = ledger.getRange(r, col_(C.MONTH));
      const v = cell.getValue();
      if (!(v instanceof Date) || v.getFullYear() !== year || v.getMonth() + 1 !== month) {
        cell.setValue(monthDate).setNumberFormat('yyyy.mm');
        moved++;
      }
    });
  }
  return { sheet: sh, count: picked.rows.length, spent: spent, budget: budget, notes: notes, moved: moved };
}

/** 출납장에서 증빙내역에 넣을 행 고르기 (지출금액이 있는 행만) */
function pickLedgerRows_(year, month, range) {
  const ledger = getSheet_();
  const last = ledger.getLastRow();
  if (last < 2) return { rows: [], sheetRows: [] };
  const start = range ? range.start : 2;
  const end = range ? range.end : last;
  const vals = ledger.getRange(start, col_(1), end - start + 1, C.OUT).getValues();
  const rows = [], sheetRows = [];
  vals.forEach((r, i) => {
    if (!(Number(r[C.OUT - 1]) > 0)) return;
    if (!range) {
      const m = r[C.MONTH - 1];
      if (!(m instanceof Date) || m.getFullYear() !== year || m.getMonth() + 1 !== month) return;
    }
    rows.push({ proof: String(r[C.PROOF - 1] || ''), date: r[C.DATE - 1],
                jeokyo: String(r[C.JEOKYO - 1] || ''), amount: Number(r[C.OUT - 1]) });
    sheetRows.push(start + i);
  });
  return { rows: rows, sheetRows: sheetRows };     // 출납장 행 순서 그대로
}

/** 오른쪽 증빙내역 표 채우기. 지출 합계를 돌려줘요 */
function fillProofTable_(sh, rows, budget, today, notes) {
  const head = findOne_(sh, '^\\s*증빙\\s*번호\\s*$');
  if (!head) { notes.push('증빙내역의 "증빙 번호" 칸을 못 찾았어요'); return 0; }
  const hr = head.getRow();
  const headVals = sh.getRange(hr, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const colOf = re => { const i = headVals.findIndex(v => re.test(v)); return i === -1 ? null : i + 1; };
  const cProof = head.getColumn(), cDate = colOf(/사용\s*일자/), cJeokyo = colOf(/적요/), cAmount = colOf(/사용\s*금액/);
  if (!cDate || !cJeokyo || !cAmount) { notes.push('증빙내역 제목(사용일자/적요/사용금액)을 못 찾았어요'); return 0; }

  let totalLabel = findOne_(sh, '^\\s*지출\\s*합계\\s*$');
  if (!totalLabel) { notes.push('"지출합계" 칸을 못 찾았어요'); return 0; }
  const first = hr + 1;
  let capacity = totalLabel.getRow() - first;
  if (rows.length > capacity) {                       // 칸이 모자라면 줄 추가
    const add = rows.length - capacity;
    sh.insertRowsBefore(totalLabel.getRow(), add);
    sh.getRange(first, cProof, 1, cAmount - cProof + 1)
      .copyTo(sh.getRange(first + capacity, cProof, add, cAmount - cProof + 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    capacity = rows.length;
    totalLabel = findOne_(sh, '^\\s*지출\\s*합계\\s*$');
    notes.push('증빙이 많아 ' + add + '줄을 추가했어요 (왼쪽 지출결의서 모양도 확인해 주세요)');
  }
  [cProof, cDate, cJeokyo, cAmount].forEach(c => sh.getRange(first, c, capacity, 1).clearContent());
  rows.forEach((r, i) => {
    sh.getRange(first + i, cProof).setNumberFormat('@').setValue(r.proof);
    sh.getRange(first + i, cDate).setValue(r.date).setNumberFormat('mm/dd');
    sh.getRange(first + i, cJeokyo).setValue(r.jeokyo);
    sh.getRange(first + i, cAmount).setValue(r.amount).setNumberFormat('#,##0');
  });
  const totalCell = rightOf_(sh, totalLabel);
  totalCell.setFormula('=SUM(' + sh.getRange(first, cAmount, capacity, 1).getA1Notation() + ')').setNumberFormat('#,##0');

  const budgetLabel = findOne_(sh, '월\\s*예산\\s*->');
  let budgetCell = null;
  if (budgetLabel) { budgetCell = rightOf_(sh, budgetLabel); budgetCell.setValue(budget).setNumberFormat('#,##0'); }
  else notes.push('"월 예산->" 칸을 못 찾았어요');
  const restLabel = findOne_(sh, '^\\s*잔\\s*액\\s*$');
  if (restLabel && budgetCell) {
    rightOf_(sh, restLabel).setFormula('=' + budgetCell.getA1Notation() + '-' + totalCell.getA1Notation()).setNumberFormat('#,##0');
  }
  const dateLabel = findOne_(sh, '^\\s*지출결의서\\s*$');
  if (dateLabel && dateLabel.getColumn() >= cProof - 1) {
    rightOf_(sh, dateLabel).setValue(today).setNumberFormat('yyyy. m. d');
  }
  return rows.reduce((a, r) => a + r.amount, 0);
}

/** 왼쪽 월별 예산표를 예산표 탭에서 옮겨 그리기 (startRow부터: 제목 줄, 항목들, 월계) */
function drawBudgetTable_(sh, startRow, month, items, notes) {
  const bsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(setting_('BUDGET_SHEET'));
  if (!bsheet || !budgetMeta.headerRow) { notes.push('예산표 탭을 못 읽어 왼쪽 예산표를 그리지 못했어요'); return; }
  const g = colSpan_(REPORT.TABLE_GROUP_COLS), it = colSpan_(REPORT.TABLE_ITEM_COLS), mo = colSpan_(REPORT.TABLE_MONTH_COLS);
  const left = g.start, width = mo.start + mo.width - g.start;

  // 그릴 수 있는 줄: '지급목적' 다음 줄부터 '담 당'(지급승인) 줄 전까지
  const stop = findOne_(sh, '^\\s*담\\s*당\\s*$') || findOne_(sh, '지급\\s*승인');
  let have = stop ? stop.getRow() - startRow : items.length + 2;
  const need = items.length + 2;                           // 제목 줄 + 항목 + 월계
  if (need > have) { sh.insertRowsBefore(startRow + have, need - have); have = need; }

  // 그 자리에 붙어 있던 캡처 그림 지우기
  sh.getImages().forEach(img => {
    const a = img.getAnchorCell();
    if (a.getRow() >= startRow - 1 && a.getRow() < startRow + have && a.getColumn() <= left + width) img.remove();
  });

  const area = sh.getRange(startRow, left, have, width);
  area.breakApart().clearContent().clearFormat();

  const mcol = budgetMeta.monthCols[month];
  const B = CONFIG.BUDGET_GROUP_COL, I = CONFIG.BUDGET_ITEM_COL;
  // 한 줄 그리기: 예산표의 같은 줄 서식을 복사하고 값 넣기
  const drawRow = (r, srcRow, groupVal, itemVal, monthVal) => {
    bsheet.getRange(srcRow, B).copyTo(sh.getRange(r, g.start, 1, g.width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    bsheet.getRange(srcRow, I).copyTo(sh.getRange(r, it.start, 1, it.width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    if (mcol) bsheet.getRange(srcRow, mcol).copyTo(sh.getRange(r, mo.start, 1, mo.width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    if (it.width > 1) sh.getRange(r, it.start, 1, it.width).merge();
    if (mo.width > 1) sh.getRange(r, mo.start, 1, mo.width).merge();
    if (g.width > 1) sh.getRange(r, g.start, 1, g.width).merge();
    if (groupVal !== null) sh.getRange(r, g.start).setValue(groupVal);
    if (itemVal !== null) sh.getRange(r, it.start).setValue(itemVal);
    if (monthVal !== null) sh.getRange(r, mo.start).setValue(monthVal);
  };

  // 제목 줄
  drawRow(startRow, budgetMeta.headerRow, '월별구분', '', month + '월');
  // 항목 줄
  items.forEach((b, i) => drawRow(startRow + 1 + i, b.row, '', b.item, b.months[month] || ''));
  // 비목: 같은 비목끼리 세로로 합치기
  let i = 0;
  while (i < items.length) {
    let j = i;
    while (j + 1 < items.length && items[j + 1].group === items[i].group) j++;
    const grp = sh.getRange(startRow + 1 + i, g.start, j - i + 1, g.width);
    grp.breakApart().merge();
    grp.setValue(items[i].group).setVerticalAlignment('middle');
    i = j + 1;
  }
  // 월계 줄
  const totalRow = startRow + 1 + items.length;
  drawRow(totalRow, budgetMeta.totalRow || budgetMeta.headerRow, '월계', '', null);
  sh.getRange(totalRow, mo.start).setFormula('=SUM(' +
    sh.getRange(startRow + 1, mo.start, Math.max(items.length, 1), 1).getA1Notation() + ')');

  const table = sh.getRange(startRow, left, items.length + 2, width);
  table.setFontSize(REPORT.TABLE_FONT_SIZE).setHorizontalAlignment('center')
       .setBorder(true, true, true, true, true, true);
  sh.getRange(startRow + 1, mo.start, items.length + 1, mo.width).setHorizontalAlignment('right').setNumberFormat('#,##0');
}

/***** 도우미 *****/
function findAll_(sh, regex) {
  return sh.createTextFinder(regex).useRegularExpression(true).findAll();
}

function findOne_(sh, regex) {
  return sh.createTextFinder(regex).useRegularExpression(true).findNext();
}

/** 'D:F' → { start: 4, width: 3 } */
function colSpan_(spec) {
  const toNum = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const parts = String(spec).toUpperCase().split(':');
  const a = toNum(parts[0].trim()), b = toNum((parts[1] || parts[0]).trim());
  return { start: Math.min(a, b), width: Math.abs(b - a) + 1 };
}

/** 제목 칸 바로 오른쪽(합쳐진 칸이면 그 끝 다음) 칸 */
function rightOf_(sh, cell) {
  if (cell.isPartOfMerge()) {
    const m = cell.getMergedRanges()[0];
    return sh.getRange(m.getRow(), m.getLastColumn() + 1);
  }
  return sh.getRange(cell.getRow(), cell.getColumn() + 1);
}

function won_(n) {
  return Number(n || 0).toLocaleString('ko-KR') + '원';
}

/** 1770000 → '백칠십칠만원정' */
function koreanMoney_(n) {
  n = Math.floor(Number(n) || 0);
  if (n === 0) return '영원정';
  const digits = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
  const small = ['', '십', '백', '천'];
  const big = ['', '만', '억', '조'];
  let out = '';
  let g = 0;
  while (n > 0) {
    const part = n % 10000;
    if (part > 0) {
      let s = '';
      String(part).split('').reverse().forEach((d, k) => {
        const v = Number(d);
        if (v) s = (v === 1 && k > 0 ? '' : digits[v]) + small[k] + s;
      });
      out = s + big[g] + out;
    }
    n = Math.floor(n / 10000);
    g++;
  }
  return out + '원정';
}
