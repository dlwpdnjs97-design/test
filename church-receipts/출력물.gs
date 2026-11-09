/******************************************************
 * 영수증 출력물 만들기 (행 범위 선택)
 * 출납장에서 고른 행들의 영수증 사진을 A4 한 장에 4장(2×2) 또는 2장씩
 * 행 순서대로 배치한 구글 문서와 PDF를 만들어요.
 ******************************************************/

const PRINT = {
  OUTPUT_FOLDER_NAME: '출력물',   // '영수증 올리는 곳' 안에 이 이름의 폴더를 만들어 저장
  PER_PAGE: 4,                    // 한 장에 넣을 영수증 수: 4(2×2) 또는 2(옆으로 2장, 크게)
  MARGIN: 20,                     // 종이 여백 (pt, 20pt ≈ 0.7cm)
};
// A4 크기와 여백으로 사진 칸 크기를 자동 계산 (pt)
const PAGE_W = 595.28, PAGE_H = 841.89;
const ROWS_PER_PAGE = PRINT.PER_PAGE === 2 ? 1 : 2;
const IMG_MAX_W = Math.floor((PAGE_W - PRINT.MARGIN * 2) / 2 - 12);
const IMG_MAX_H = Math.floor((PAGE_H - PRINT.MARGIN * 2) / ROWS_PER_PAGE - 48);

/** 메뉴에서 실행: 몇 행부터 몇 행까지 출력할지 물어보고 문서를 만들어요 */
function makeReceiptDoc() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) {
    ui.alert('출납장에 아직 영수증이 없어요.');
    return;
  }

  // 출납장에서 행을 선택해 둔 상태면 그 범위를 기본값으로 써요
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sel = ss.getActiveSheet().getName() === CONFIG.SHEET_NAME ? ss.getActiveRange() : null;
  const selText = sel && sel.getLastRow() >= 2
    ? Math.max(2, sel.getRow()) + '-' + sel.getLastRow() : '';

  const answer = ui.prompt('영수증 출력물 만들기',
    '출납장 몇 행부터 몇 행까지 출력할까요? 예: 2-15\n' +
    (selText ? '비워 두면 지금 선택한 ' + selText + '행을 출력해요.' : '(2행부터 ' + last + '행까지 있어요)'),
    ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;

  const range = parseRowRange_(answer.getResponseText() || selText, last);
  if (!range) {
    ui.alert('2-15 처럼 "시작행-끝행"으로 써 주세요. 출납장에는 2행부터 ' + last + '행까지 있어요.');
    return;
  }
  const result = buildReceiptDoc_(range.start, range.end);
  if (!result) {
    ui.alert(range.start + '~' + range.end + '행에 영수증 사진이 있는 줄이 없어요.');
    return;
  }

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px;line-height:1.8">' +
    '<p>' + range.start + '~' + range.end + '행의 영수증 ' + result.count + '장을 ' + result.pages + '쪽에 담았어요.' +
    (result.skipped ? '<br>사진 링크가 없는 ' + result.skipped + '줄은 건너뛰었어요.' : '') +
    (result.missing ? '<br>⚠️ 사진을 넣지 못한 영수증 ' + result.missing + '장은 칸에 표시해 뒀어요.' : '') + '</p>' +
    '<p>📄 <a href="' + result.docUrl + '" target="_blank">구글 문서로 열기</a> (내용 고치기)<br>' +
    '🖨️ <a href="' + result.pdfUrl + '" target="_blank">PDF로 열기</a> (바로 인쇄)</p></div>')
    .setWidth(400).setHeight(200);
  ui.showModalDialog(html, '출력물이 만들어졌어요');
}

/** 실제로 문서를 만드는 부분: 출납장 start~end행을 행 순서대로 배치 */
function buildReceiptDoc_(start, end) {
  const sheet = getSheet_();
  const all = sheet.getRange(start, 1, end - start + 1, C.PHOTO).getValues();
  const rows = all
    .filter(r => r[C.PHOTO - 1])
    .map(r => ({ proof: String(r[C.PROOF - 1] || ''), url: String(r[C.PHOTO - 1]) }));
  if (rows.length === 0) return null;

  const title = '영수증 증빙 ' + rows[0].proof + ' ~ ' + rows[rows.length - 1].proof +
                ' (' + start + '-' + end + '행)';
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.setPageWidth(PAGE_W).setPageHeight(PAGE_H)          // A4
      .setMarginTop(PRINT.MARGIN).setMarginBottom(PRINT.MARGIN)
      .setMarginLeft(PRINT.MARGIN).setMarginRight(PRINT.MARGIN);

  let missing = 0;
  let pages = 0;
  const perPage = ROWS_PER_PAGE * 2;
  for (let i = 0; i < rows.length; i += perPage) {
    if (pages > 0) body.appendPageBreak();
    pages++;
    const group = rows.slice(i, i + perPage);
    const table = body.appendTable(ROWS_PER_PAGE === 1 ? [['', '']] : [['', ''], ['', '']]);
    table.setBorderWidth(0);

    for (let k = 0; k < perPage; k++) {
      const cell = table.getCell(Math.floor(k / 2), k % 2);
      cell.setPaddingTop(2).setPaddingBottom(4).setPaddingLeft(4).setPaddingRight(4)
          .setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
      const label = cell.getChild(0).asParagraph();
      if (k >= group.length) continue;                     // 마지막 쪽 빈칸

      const item = group[k];
      label.setText(item.proof);
      label.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      label.editAsText().setFontSize(11).setBold(true);

      const p = cell.appendParagraph('');
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      const blob = imageBlob_(item.url);
      if (!blob) {
        p.appendText('⚠️ 사진을 넣지 못했어요. 직접 넣어 주세요.\n' + item.url).setFontSize(8).setBold(false);
        missing++;
        continue;
      }
      try {
        const img = p.appendInlineImage(blob);
        // 문서의 그림 크기는 픽셀 단위라서 pt × 4/3 으로 바꿔서 계산
        const scale = Math.min(IMG_MAX_W * 4 / 3 / img.getWidth(), IMG_MAX_H * 4 / 3 / img.getHeight());
        img.setWidth(Math.round(img.getWidth() * scale)).setHeight(Math.round(img.getHeight() * scale));
      } catch (e) {
        p.appendText('⚠️ 사진을 넣지 못했어요(' + e.message + ').\n' + item.url).setFontSize(8).setBold(false);
        missing++;
      }
    }
  }
  doc.saveAndClose();

  // 문서와 PDF를 '출력물' 폴더에 저장
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const out = getOrCreateSub_(root, PRINT.OUTPUT_FOLDER_NAME);
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(out);
  const pdf = out.createFile(docFile.getAs('application/pdf').setName(title + '.pdf'));

  return { count: rows.length, skipped: all.length - rows.length, pages: pages, missing: missing,
           docUrl: doc.getUrl(), pdfUrl: pdf.getUrl() };
}

/** 드라이브 사진을 문서에 넣을 수 있는 그림으로 가져오기 (아이폰 HEIC, PDF 영수증 포함) */
function imageBlob_(url) {
  const m = /\/d\/([-\w]{10,})/.exec(url) || /[?&]id=([-\w]{10,})/.exec(url);
  if (!m) return null;
  const id = m[1];

  // 1순위: 드라이브 미리보기 그림 (사진 방향이 바로잡히고 용량도 작아요)
  try {
    const res = UrlFetchApp.fetch('https://drive.google.com/thumbnail?id=' + id + '&sz=w1600', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200) {
      const b = res.getBlob();
      if (/^image\/(jpeg|png|gif)/.test(b.getContentType())) return b;
    }
  } catch (e) { /* 아래 방법으로 */ }

  // 2순위: 원본 파일 (JPG, PNG, GIF만 문서에 넣을 수 있어요)
  try {
    const file = DriveApp.getFileById(id);
    if (['image/jpeg', 'image/png', 'image/gif'].indexOf(file.getMimeType()) !== -1) return file.getBlob();
  } catch (e) { /* 파일을 못 찾음 */ }
  return null;
}

/** "2-15", "2~15", "2 15", "7" 같은 입력을 행 범위로 바꾸기 */
function parseRowRange_(text, last) {
  const m = /^\s*(\d+)\s*(?:[-~,\s]\s*(\d+))?\s*$/.exec(String(text || ''));
  if (!m) return null;
  let a = Number(m[1]);
  let b = m[2] ? Number(m[2]) : a;
  if (a > b) { const t = a; a = b; b = t; }
  if (a < 2 || b > last) return null;
  return { start: a, end: b };
}
