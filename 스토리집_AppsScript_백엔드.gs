/**
 * 스토리집 도서 검색·대여 시스템 — Apps Script 백엔드
 *
 * 배포 전 준비:
 * 1. 이 코드를 "스토리집 도서 목록" 스프레드시트에 바인딩된 Apps Script 프로젝트에 붙여넣는다.
 * 2. 프로젝트 설정 > 스크립트 속성(Script Properties)에 ADMIN_PASSWORD 키를 추가하고 관리자 비밀번호를 값으로 넣는다.
 * 3. 시트1에 ISBN / 표지URL / 소개 열이 있어야 한다(헤더 이름으로 찾으므로 순서는 상관없음).
 * 4. "추천도서" 시트를 만든다 (구조는 아래 SHEET_RECOMMEND 관련 함수 참고).
 * 5. 배포 > 웹 앱으로 배포: 실행 대상 "나", 액세스 권한 "모든 사용자".
 */

const SHEET_BOOKS = '시트1';
const SHEET_LOANS = '대여기록';
const SHEET_RECOMMEND = '추천도서';

const RENTAL_DAYS = 14;
const MAX_RENTAL = 3;
const MAX_REC = 1;

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'books') return getBooks();
    if (action === 'status') return getStatus();
    if (action === 'myloans') return getMyLoans(e.parameter.phone || '');
    if (action === 'recommend') return getRecommend();
    if (action === 'admin') return getAdminRecords(e.parameter.pw || '');
    return jsonOut({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: '요청 형식이 올바르지 않습니다.' });
  }
  try {
    if (payload.action === 'apply') return applyBooks(payload);
    if (payload.action === 'return') return requestReturn(payload);
    if (payload.action === 'admin_confirm_pay') return adminConfirmPay(payload);
    if (payload.action === 'admin_complete_return') return adminCompleteReturn(payload);
    return jsonOut({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── 공통: 헤더 이름 기준으로 시트를 객체 배열로 읽기 ──
function readSheetAsObjects(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const rowValues = values[i];
    const obj = { row: i + 1 };
    headers.forEach((h, idx) => { obj[h] = rowValues[idx]; });
    rows.push(obj);
  }
  return { sheet, headers, rows };
}

function toIso(v) {
  if (v instanceof Date) return v.toISOString();
  return v ? String(v) : '';
}

function formatMonth_(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0');
  }
  return v ? String(v) : '';
}

// ── 도서 목록 ──
function getBooks() {
  const { rows } = readSheetAsObjects(SHEET_BOOKS);
  const books = rows
    .filter(r => String(r['제목'] || '').trim())
    .map(r => ({
      id: String(r['구분']),
      title: String(r['제목'] || ''),
      author: String(r['저자'] || ''),
      cat: String(r['장르'] || ''),
      isbn: String(r['ISBN'] || ''),
      coverUrl: String(r['표지URL'] || ''),
      description: String(r['소개'] || ''),
    }));
  return jsonOut({ ok: true, books });
}

// ── 대출 상태 (대여중/반납신청인 책의 반납예정일) ──
function getActiveLoanRows() {
  const { rows } = readSheetAsObjects(SHEET_LOANS);
  return rows.filter(r => r['이용형태'] === '대여' && (r['상태'] === '대여중' || r['상태'] === '반납신청'));
}

function getStatus() {
  const active = getActiveLoanRows();
  const status = {};
  active.forEach(r => {
    status[String(r['도서관리번호'])] = { dueDate: toIso(r['반납예정일']) };
  });
  return jsonOut({ ok: true, status });
}

// ── 내 대여 목록 ──
function getMyLoans(phone) {
  if (!phone) return jsonOut({ ok: true, loans: [] });
  const { rows } = readSheetAsObjects(SHEET_LOANS);
  const loans = rows
    .filter(r => r['이용형태'] === '대여' && String(r['전화번호']) === String(phone) && r['상태'] !== '반납완료')
    .map(r => ({
      bookId: String(r['도서관리번호']),
      title: r['도서명'],
      author: r['저자'],
      loanDate: toIso(r['대출일']),
      dueDate: toIso(r['반납예정일']),
      status: r['상태'],
      returnRequestDate: toIso(r['반납신청일시']),
      returnTimeSlot: r['반납신청시간대'] || '',
    }));
  return jsonOut({ ok: true, loans });
}

// ── 이달의 추천도서 (추천도서 시트: A열=라벨, B열=값 / 6행부터 관리번호 목록) ──
function getRecommend() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECOMMEND);
  if (!sheet) return jsonOut({ ok: true, month: '', theme: '', posterUrl: '', books: [] });

  const values = sheet.getDataRange().getValues();
  const meta = {};
  for (let i = 0; i < values.length; i++) {
    const label = String(values[i][0] || '').trim();
    if (label === '연월') meta.month = formatMonth_(values[i][1]);
    if (label === '주제') meta.theme = String(values[i][1] || '');
    if (label === '포스터URL') meta.posterUrl = String(values[i][1] || '');
    if (label === '관리번호') { meta.idsStartRow = i + 1; }
  }

  const ids = [];
  if (meta.idsStartRow !== undefined) {
    for (let i = meta.idsStartRow; i < values.length; i++) {
      const v = String(values[i][0] || '').trim();
      if (v) ids.push(v);
    }
  }

  const { rows: bookRows } = readSheetAsObjects(SHEET_BOOKS);
  const byId = {};
  bookRows.forEach(r => { byId[String(r['구분'])] = r; });

  const books = ids
    .map(id => byId[id])
    .filter(r => r)
    .map(r => ({
      id: String(r['구분']),
      title: r['제목'],
      author: r['저자'],
      isbn: String(r['ISBN'] || ''),
      coverUrl: String(r['표지URL'] || ''),
      description: String(r['소개'] || ''),
    }));

  return jsonOut({
    ok: true,
    month: meta.month || '',
    theme: meta.theme || '',
    posterUrl: meta.posterUrl || '',
    books,
  });
}

// ── 관리자 인증 ──
function getAdminPassword() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
}

function adminAuth(pw) {
  const real = getAdminPassword();
  return !!real && pw === real;
}

function getAdminRecords(pw) {
  if (!adminAuth(pw)) return jsonOut({ ok: false, error: '비밀번호가 올바르지 않아요.' });
  const { rows } = readSheetAsObjects(SHEET_LOANS);
  const records = rows.map(r => ({
    row: r.row,
    name: r['이름'],
    phone: r['전화번호'],
    bookId: String(r['도서관리번호']),
    title: r['도서명'],
    author: r['저자'],
    type: r['이용형태'],
    loanDate: toIso(r['대출일']),
    dueDate: toIso(r['반납예정일']),
    status: r['상태'],
    returnRequestDate: toIso(r['반납신청일시']),
    returnTimeSlot: r['반납신청시간대'] || '',
    paymentConfirmed: r['입금확인'] || 'N',
  }));
  return jsonOut({ ok: true, records });
}

// ── 대여/구매 신청 ──
function applyBooks(payload) {
  const name = String(payload.name || '').trim();
  const phone = String(payload.phone || '').trim();
  const type = payload.type;
  const items = payload.items || [];

  if (!name || !phone) return jsonOut({ ok: false, error: '이름과 연락처를 입력해주세요.' });
  if (!items.length) return jsonOut({ ok: false, error: '선택한 도서가 없어요.' });
  if (type !== '대여' && type !== '구매') return jsonOut({ ok: false, error: '이용 형태가 올바르지 않아요.' });

  if (type === '대여') {
    const activeForPhone = getActiveLoanRows().filter(r => String(r['전화번호']) === phone).length;
    if (activeForPhone + items.length > MAX_RENTAL) {
      return jsonOut({ ok: false, error: '대여는 1인당 최대 ' + MAX_RENTAL + '권까지 가능해요.' });
    }
    const activeBookIds = new Set(getActiveLoanRows().map(r => String(r['도서관리번호'])));
    const alreadyOut = items.find(it => activeBookIds.has(String(it.bookId)));
    if (alreadyOut) {
      return jsonOut({ ok: false, error: '이미 대여 중인 도서가 포함되어 있어요. 새로고침 후 다시 시도해주세요.' });
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const now = new Date();
  const dueDate = new Date(now.getTime() + RENTAL_DAYS * 24 * 60 * 60 * 1000);

  items.forEach(it => {
    sheet.appendRow([
      Utilities.getUuid(),
      name,
      phone,
      it.bookId,
      it.title,
      it.author,
      type,
      type === '대여' ? now : '',
      type === '대여' ? dueDate : '',
      type === '대여' ? '대여중' : '구매신청',
      '',
      '',
      'N',
    ]);
  });

  return jsonOut({ ok: true });
}

// ── 반납 신청 ──
function requestReturn(payload) {
  const phone = String(payload.phone || '');
  const bookId = String(payload.bookId || '');
  const returnDate = payload.returnDate;
  const returnTimeSlot = payload.returnTimeSlot;

  const { sheet, rows } = readSheetAsObjects(SHEET_LOANS);
  const target = rows.find(r => String(r['전화번호']) === phone && String(r['도서관리번호']) === bookId && r['상태'] === '대여중');
  if (!target) return jsonOut({ ok: false, error: '대여 중인 도서를 찾을 수 없어요.' });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const statusCol = headers.indexOf('상태') + 1;
  const reqDateCol = headers.indexOf('반납신청일시') + 1;
  const reqSlotCol = headers.indexOf('반납신청시간대') + 1;

  sheet.getRange(target.row, statusCol).setValue('반납신청');
  sheet.getRange(target.row, reqDateCol).setValue(returnDate);
  sheet.getRange(target.row, reqSlotCol).setValue(returnTimeSlot);

  return jsonOut({ ok: true });
}

// ── 관리자: 입금 확인 ──
function adminConfirmPay(payload) {
  if (!adminAuth(payload.pw)) return jsonOut({ ok: false, error: '권한이 없어요.' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = headers.indexOf('입금확인') + 1;
  sheet.getRange(payload.row, col).setValue('Y');
  return jsonOut({ ok: true });
}

// ── 관리자: 반납 완료 처리 ──
function adminCompleteReturn(payload) {
  if (!adminAuth(payload.pw)) return jsonOut({ ok: false, error: '권한이 없어요.' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const col = headers.indexOf('상태') + 1;
  sheet.getRange(payload.row, col).setValue('반납완료');
  return jsonOut({ ok: true });
}

// ═══════════════════════════════════════════════════════════════
// 신간 자동화: 시트1에 ISBN을 입력하면 표지·소개를 알라딘 API로 자동으로 채운다.
// 스크립트 속성에 ALADIN_TTBKEY를 등록해야 동작한다.
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().addMenu('스토리집', [
    { name: '① 신간 자동화 켜기 (최초 1회만 실행)', functionName: 'installEditTrigger' },
    { name: '표지·소개 다시 채우기 (ISBN 입력된 행 전체)', functionName: 'fillAllMissingCovers' },
  ]);
}

// onEdit(e)라는 이름으로 직접 정의하면 "단순 트리거"가 되어 보안상 외부 API 호출(UrlFetchApp)이
// 막힌다. 그래서 이 함수는 다른 이름으로 두고, installEditTrigger()로 "설치 트리거"에 등록해서 쓴다.
function installEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onBookSheetEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onBookSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('신간 자동화가 켜졌습니다. 이제부터 ISBN을 입력하면 표지·소개가 자동으로 채워져요.');
}

function onBookSheetEdit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_BOOKS) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const isbnCol = headers.indexOf('ISBN') + 1;
    if (isbnCol === 0) return;

    const editedStartCol = e.range.getColumn();
    const editedEndCol = editedStartCol + e.range.getNumColumns() - 1;
    if (isbnCol < editedStartCol || isbnCol > editedEndCol) return;

    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    for (let r = startRow; r < startRow + numRows; r++) {
      if (r === 1) continue;
      fillCoverForRow(sheet, headers, r);
    }
  } catch (err) {
    // 트리거 오류가 시트 편집 자체를 막지 않도록 조용히 무시한다.
  }
}

// 반환값: 알라딘 API를 실제로 호출했으면 true (시간 제한 계산에 사용)
function fillCoverForRow(sheet, headers, row) {
  const isbnCol = headers.indexOf('ISBN') + 1;
  const coverCol = headers.indexOf('표지URL') + 1;
  const descCol = headers.indexOf('소개') + 1;
  if (!isbnCol || !coverCol || !descCol) return false;

  const isbn = String(sheet.getRange(row, isbnCol).getValue() || '').trim();
  if (!isbn) return false;
  const cover = String(sheet.getRange(row, coverCol).getValue() || '').trim();
  const desc = String(sheet.getRange(row, descCol).getValue() || '').trim();
  if (cover && desc) return false;

  const info = lookupAladin(isbn);
  if (!info) return true;
  if (!cover && info.cover) sheet.getRange(row, coverCol).setValue(info.cover);
  if (!desc && info.description) sheet.getRange(row, descCol).setValue(info.description);
  return true;
}

function lookupAladin(isbn) {
  const ttbkey = PropertiesService.getScriptProperties().getProperty('ALADIN_TTBKEY');
  if (!ttbkey) return null;
  const url = 'https://www.aladin.co.kr/ttb/api/ItemLookUp.aspx?ttbkey=' + ttbkey +
    '&ItemId=' + isbn + '&ItemIdType=ISBN13&output=js&Version=20131101&Cover=Big';
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());
    const item = data.item && data.item[0];
    if (!item) return null;
    return { cover: item.cover, description: item.description };
  } catch (err) {
    return null;
  }
}

// 붙여넣기 등으로 onEdit이 놓친 행이 있을 때 메뉴에서 수동 실행.
// 앱스크립트는 실행당 6분 제한이 있어서, 5분을 넘기면 안전하게 멈추고 어디까지 했는지 알려준다.
// 남은 행이 있으면 메뉴를 다시 누르면 이어서 처리된다(이미 채워진 행은 건너뛰므로 중복 호출 안 됨).
function fillAllMissingCovers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BOOKS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const lastRow = sheet.getLastRow();
  const startedAt = Date.now();
  const TIME_LIMIT_MS = 5 * 60 * 1000;

  let processed = 0;
  let stoppedEarly = false;
  for (let r = 2; r <= lastRow; r++) {
    if (Date.now() - startedAt > TIME_LIMIT_MS) { stoppedEarly = true; break; }
    const calledApi = fillCoverForRow(sheet, headers, r);
    if (calledApi) {
      processed++;
      Utilities.sleep(250);
    }
  }

  if (stoppedEarly) {
    SpreadsheetApp.getUi().alert('시간이 오래 걸려서 일부만 처리하고 멈췄습니다(' + processed + '건 처리). 메뉴를 다시 눌러 이어서 진행해주세요.');
  } else {
    SpreadsheetApp.getUi().alert('표지·소개 채우기를 완료했습니다(' + processed + '건 처리).');
  }
}
