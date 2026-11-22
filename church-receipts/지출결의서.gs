/******************************************************
 * 월별 지출결의서 만들기
 * '지출결의, 증빙내역' 탭을 복사해서 'N월' 탭을 만들고
 *  - 왼쪽 지출결의서: 날짜, 청구 금액(그 달 예산), 일금, 지급목적, 월별 예산표
 *  - 오른쪽 증빙내역: 출납장에서 그 달 영수증 목록, 지출합계, 월 예산, 잔액
 * 을 채워요. 칸 위치는 제목 글자('증빙 번호', '지출합계' …)를 찾아서 정해요.
 ******************************************************/

const REPORT = {
  TEMPLATE_SHEET: '지출결의, 증빙내역',   // 복사할 서식 탭
};

/** 메뉴에서 실행: 몇 월 지출결의서를 만들지 물어보고 탭을 만들어요 */
function makeMonthlyReport() {
  const ui = SpreadsheetApp.getUi();
  const now = new Date();
  const answer = ui.prompt('월별 지출결의서 만들기',
    '몇 월 지출결의서를 만들까요? 예: 11 (비워 두면 이번 달)', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const text = answer.getResponseText().replace(/월/g, '').trim();
  const month = text ? Number(text) : now.getMonth() + 1;
  if (!(month >= 1 && month <= 12)) { ui.alert('1부터 12 사이 숫자로 써 주세요.'); return; }
  const year = now.getFullYear();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = month + '월';
  const old = ss.getSheetByName(name);
  if (old) {
    const yes = ui.alert(name + ' 탭이 이미 있어요. 지우고 새로 만들까요?\n(직접 고친 내용도 사라져요)', ui.ButtonSet.YES_NO);
    if (yes !== ui.Button.YES) return;
    ss.deleteSheet(old);
  }
  const result = buildMonthlyReport_(year, month);
  ss.setActiveSheet(result.sheet);
  ui.alert(name + ' 지출결의서를 만들었어요.\n' +
           '· 영수증 ' + result.count + '건, 지출 ' + won_(result.spent) + '\n' +
           '· ' + name + ' 예산(청구 금액) ' + won_(result.budget) +
           (result.notes.length ? '\n\n확인해 주세요:\n· ' + result.notes.join('\n· ') : ''));
}

/** 실제로 탭을 만드는 부분 */
function buildMonthlyReport_(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tpl = ss.getSheetByName(REPORT.TEMPLATE_SHEET);
  if (!tpl) throw new Error('서식 탭(' + REPORT.TEMPLATE_SHEET + ')을 찾을 수 없어요.');
  const sh = tpl.copyTo(ss).setName(month + '월');
  // 앞 달 탭 바로 뒤에 놓기
  const prev = ss.getSheetByName((month - 1) + '월');
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(prev ? prev.getIndex() + 1 : ss.getNumSheets());

  const notes = [];
  const today = new Date();
  const budgetItems = getBudgetItems_().filter(b => (b.months[month] || 0) > 0);
  const budget = budgetItems.reduce((a, b) => a + b.months[month], 0);
  if (!getBudgetItems_().length) notes.push('예산표 탭(' + setting_('BUDGET_SHEET') + ')을 못 읽어 예산을 0으로 뒀어요');

  // ── 왼쪽: 날짜 ──
  findAll_(sh, '\\d{4}년\\s*\\d{1,2}월\\s*\\d{1,2}일').forEach(c => {
    const spaced = /년\s/.test(String(c.getDisplayValue()));
    c.setNumberFormat('@').setValue(Utilities.formatDate(today, Session.getScriptTimeZone(),
      spaced ? 'yyyy년 MM월 dd일' : 'yyyy년MM월dd일'));
  });
  // ── 왼쪽: 지급목적의 '○월' ──
  findAll_(sh, '지급목적').forEach(c => {
    c.setValue(String(c.getValue()).replace(/\d{1,2}\s*월/, month + '월'));
  });

  // ── 왼쪽: 월별 예산표 ──
  const table = fillBudgetTable_(sh, month, budgetItems, notes);

  // ── 왼쪽: 청구 금액(₩)과 일금 ──
  const wonCell = findOne_(sh, '^\\s*₩\\s*$');
  if (wonCell) {
    const v = rightOf_(sh, wonCell);
    v.setValue(budget);
  } else notes.push('₩ 칸을 못 찾았어요');
  const ilgum = findOne_(sh, '^\\s*일\\s*금\\s*$');
  if (ilgum) rightOf_(sh, ilgum).setValue(koreanMoney_(budget));
  else notes.push('일금 칸을 못 찾았어요');

  // ── 왼쪽 아래: '○월 예산' 칸 ──
  const budgetLabel = findOne_(sh, '^\\s*\\d{1,2}\\s*월\\s*예산\\s*$');
  if (budgetLabel) {
    budgetLabel.setValue(month + '월 예산');
    sh.getRange(budgetLabel.getRow() + 1, budgetLabel.getColumn()).setValue(budget);
  }

  // ── 오른쪽: 증빙내역 ──
  const rows = monthRows_(year, month);
  const spent = fillProofTable_(sh, rows, budget, today, notes);

  return { sheet: sh, count: rows.length, spent: spent, budget: budget, notes: notes, table: table };
}

/** 출납장에서 그 달 지출 줄들 (증빙번호 순) */
function monthRows_(year, month) {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, col_(1), last - 1, C.OUT).getValues()
    .filter(r => r[C.MONTH - 1] instanceof Date &&
                 r[C.MONTH - 1].getFullYear() === year &&
                 r[C.MONTH - 1].getMonth() + 1 === month &&
                 Number(r[C.OUT - 1]) > 0)
    .map(r => ({ proof: String(r[C.PROOF - 1] || ''), date: r[C.DATE - 1],
                 jeokyo: String(r[C.JEOKYO - 1] || ''), amount: Number(r[C.OUT - 1]) }))
    .sort((a, b) => sortProof_(a.proof) < sortProof_(b.proof) ? -1 : sortProof_(a.proof) > sortProof_(b.proof) ? 1 : 0);
}

function sortProof_(p) {
  return p.replace(/\?/g, '~');
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
    const fmtSrc = sh.getRange(first, cProof, 1, cAmount - cProof + 1);
    fmtSrc.copyTo(sh.getRange(first + capacity, cProof, add, cAmount - cProof + 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    capacity = rows.length;
    totalLabel = findOne_(sh, '^\\s*지출\\s*합계\\s*$');
    notes.push('영수증이 많아 증빙내역에 ' + add + '줄을 추가했어요 (왼쪽 지출결의서 모양도 확인해 주세요)');
  }
  // 예전 내용(#REF! 수식 포함) 지우고 새로 쓰기
  [cProof, cDate, cJeokyo, cAmount].forEach(c => sh.getRange(first, c, capacity, 1).clearContent());
  rows.forEach((r, i) => {
    sh.getRange(first + i, cProof).setNumberFormat('@').setValue(r.proof);
    sh.getRange(first + i, cDate).setValue(r.date).setNumberFormat('mm/dd');
    sh.getRange(first + i, cJeokyo).setValue(r.jeokyo);
    sh.getRange(first + i, cAmount).setValue(r.amount).setNumberFormat('#,##0');
  });
  const amountA1 = sh.getRange(first, cAmount, capacity, 1).getA1Notation();
  const totalCell = rightOf_(sh, totalLabel);
  totalCell.setFormula('=SUM(' + amountA1 + ')').setNumberFormat('#,##0');

  const budgetLabel = findOne_(sh, '월\\s*예산\\s*->');
  let budgetCell = null;
  if (budgetLabel) { budgetCell = rightOf_(sh, budgetLabel); budgetCell.setValue(budget).setNumberFormat('#,##0'); }
  else notes.push('"월 예산->" 칸을 못 찾았어요');
  const restLabel = findOne_(sh, '^\\s*잔\\s*액\\s*$');
  if (restLabel && budgetCell) {
    rightOf_(sh, restLabel).setFormula('=' + budgetCell.getA1Notation() + '-' + totalCell.getA1Notation()).setNumberFormat('#,##0');
  }
  // 증빙내역 위쪽 '지출결의서 | 날짜' 칸
  const dateLabel = sh.createTextFinder('^\\s*지출결의서\\s*$').useRegularExpression(true).matchEntireCell(true).findNext();
  if (dateLabel && dateLabel.getColumn() >= cProof - 1) {
    rightOf_(sh, dateLabel).setValue(today).setNumberFormat('yyyy. m. d');
  }
  return rows.reduce((a, r) => a + r.amount, 0);
}

/** 왼쪽 월별 예산표('월별구분' ~ '월계')를 그 달 예산 항목으로 다시 그리기 */
function fillBudgetTable_(sh, month, items, notes) {
  const top = findOne_(sh, '^\\s*월별\\s*구분\\s*$');
  const bottom = findOne_(sh, '^\\s*월\\s*계\\s*$');
  if (!top || !bottom || bottom.getRow() <= top.getRow()) { notes.push('왼쪽 월별 예산표(월별구분~월계)를 못 찾았어요'); return false; }
  const r0 = top.getRow();
  const cGroup = top.getColumn();
  const wGroup = mergeWidth_(top);
  const cItem = cGroup + wGroup;
  const headVals = sh.getRange(r0, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  let cMonth = headVals.findIndex((v, i) => i + 1 > cItem && /^\s*\d{1,2}\s*월\s*$/.test(v)) + 1;
  if (!cMonth) { notes.push('왼쪽 예산표에서 "○월" 제목 칸을 못 찾았어요'); return false; }
  const wItem = cMonth - cItem;
  const wMonth = mergeWidth_(sh.getRange(r0, cMonth));
  const width = cMonth + wMonth - cGroup;
  sh.getRange(r0, cMonth).setValue(month + '월');

  // 줄 수 맞추기 (모자라면 월계 위에 추가)
  let have = bottom.getRow() - r0 - 1;
  const need = Math.max(items.length, 1);
  if (need > have) {
    sh.insertRowsBefore(bottom.getRow(), need - have);
    have = need;
  }
  const body = sh.getRange(r0 + 1, cGroup, have, width);
  const sample = sh.getRange(r0 + 1, cGroup, 1, width);         // 첫 줄 모양을 기본으로
  body.breakApart();
  sample.copyTo(body, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  body.clearContent();

  // 예산표 탭의 비목 칸 색을 그대로 가져오기
  const bsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(setting_('BUDGET_SHEET'));
  const groupFirstRow = {};
  getBudgetItems_().forEach(b => { if (!(b.group in groupFirstRow)) groupFirstRow[b.group] = b.row; });

  items.forEach((b, i) => {
    const r = r0 + 1 + i;
    sh.getRange(r, cItem, 1, wItem).merge().setValue(b.item).setHorizontalAlignment('center');
    sh.getRange(r, cMonth, 1, wMonth).merge().setValue(b.months[month]).setNumberFormat('#,##0');
  });
  // 비목 칸: 같은 비목끼리 세로로 합치기
  let i = 0;
  while (i < items.length) {
    let j = i;
    while (j + 1 < items.length && items[j + 1].group === items[i].group) j++;
    const g = sh.getRange(r0 + 1 + i, cGroup, j - i + 1, wGroup).merge();
    g.setValue(items[i].group).setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (bsheet && groupFirstRow[items[i].group]) {
      g.setBackground(bsheet.getRange(groupFirstRow[items[i].group], CONFIG.BUDGET_GROUP_COL).getBackground());
    }
    i = j + 1;
  }
  if (items.length < have) notes.push('왼쪽 예산표에 빈 줄이 ' + (have - items.length) + '개 남아 있어요. 필요 없으면 지워 주세요');

  // 월계
  const sumCell = sh.getRange(bottom.getRow(), cMonth);
  sumCell.setFormula('=SUM(' + sh.getRange(r0 + 1, cMonth, have, 1).getA1Notation() + ')').setNumberFormat('#,##0');
  return true;
}

/***** 도우미 *****/
function findAll_(sh, regex) {
  return sh.createTextFinder(regex).useRegularExpression(true).findAll();
}

function findOne_(sh, regex) {
  return sh.createTextFinder(regex).useRegularExpression(true).findNext();
}

/** 합쳐진 칸이면 그 너비(칸 수), 아니면 1 */
function mergeWidth_(cell) {
  return cell.isPartOfMerge() ? cell.getMergedRanges()[0].getNumColumns() : 1;
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
      const ds = String(part).split('').reverse();
      ds.forEach((d, k) => {
        const v = Number(d);
        if (!v) return;
        s = (v === 1 && k > 0 ? '' : digits[v]) + small[k] + s;
      });
      out = s + big[g] + out;
    }
    n = Math.floor(n / 10000);
    g++;
  }
  return out + '원정';
}
