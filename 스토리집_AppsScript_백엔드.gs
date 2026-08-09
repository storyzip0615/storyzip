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
const VALID_RETURN_SLOTS = ['오전 (10~13시)', '오후 (13~17시)'];

// 관리자 비밀번호 무차별 대입 방어용 (실패 누적 시 일정 시간 잠금)
const ADMIN_FAIL_LIMIT = 5;
const ADMIN_LOCK_MINUTES = 5;

// 관리자 세션 토큰 TTL(초) — 로그인 이후에는 원문 비밀번호 대신 이 토큰으로 인증한다.
const ADMIN_TOKEN_TTL_SEC = 1800;

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'books') return getBooks();
    if (action === 'status') return getStatus();
    if (action === 'myloans') return getMyLoans(e.parameter.phone || '');
    if (action === 'recommend') return getRecommend();
    if (action === 'bookDetail') return getBookDetail(e.parameter.id || '');
    // 관리자 비밀번호는 GET 쿼리에 남기지 않는다 — doPost의 'admin_records'로만 조회한다.
    return jsonOut({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    // 내부 오류 원문(시트 구조·경로 등)을 클라이언트에 그대로 노출하지 않는다.
    // 실제 원인은 Apps Script 실행 기록(Executions 로그)에서 확인한다.
    console.error('doGet error: ' + err);
    return jsonOut({ ok: false, error: '요청 처리 중 문제가 발생했어요. 잠시 후 다시 시도해주세요.' });
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
    if (payload.action === 'admin_records') return getAdminRecords(payload);
    if (payload.action === 'admin_confirm_pay') return adminConfirmPay(payload);
    if (payload.action === 'admin_complete_return') return adminCompleteReturn(payload);
    if (payload.action === 'admin_complete_purchase') return adminCompletePurchase(payload);
    return jsonOut({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    console.error('doPost error: ' + err);
    return jsonOut({ ok: false, error: '요청 처리 중 문제가 발생했어요. 잠시 후 다시 시도해주세요.' });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ── 청크 캐싱 (CacheService 키당 값 100KB 제한 우회) ──
// getBooks() 응답은 583권 규모라 소개글을 120자로 잘라도 100KB를 넘을 수 있어 통짜로는
// 캐싱할 수 없다. 여러 조각으로 나눠 저장/재조합한다.
// ★100KB 제한은 JS 문자열 length(문자 수)가 아니라 실제 UTF-8 바이트 기준이다 — 한글은
// 문자당 최대 3바이트라 문자 수로만 자르면 바이트 기준 한도를 넘을 수 있다. 그래서
// 문자 수 기준으로 넉넉히 자른 뒤, Utilities.newBlob(...).getBytes().length로 실제
// 바이트를 재보고 넘으면 잘라낸 조각을 줄여나간다(안전망 — 순수 한글 텍스트라도 절대
// 100KB를 넘지 않게 보장).
const CACHE_CHUNK_CHARS = 25000; // 넉넉한 여유를 둔 문자 수 단위(최종 안전판은 바이트 재검증)
const CACHE_MAX_BYTES = 100 * 1024; // CacheService 값 1개당 한도

function utf8ByteLength_(str) {
  return Utilities.newBlob(str).getBytes().length;
}

function cachePutChunked_(key, value, ttlSec) {
  const chunks = [];
  let i = 0;
  while (i < value.length) {
    let end = Math.min(i + CACHE_CHUNK_CHARS, value.length);
    let piece = value.slice(i, end);
    while (utf8ByteLength_(piece) > CACHE_MAX_BYTES) {
      end = i + Math.floor((end - i) * 0.9);
      piece = value.slice(i, end);
    }
    chunks.push(piece);
    i = end;
  }
  const entries = {};
  entries[key + ':meta'] = JSON.stringify({ n: chunks.length });
  chunks.forEach((c, idx) => { entries[key + ':' + idx] = c; });
  CacheService.getScriptCache().putAll(entries, ttlSec);
}

function cacheGetChunked_(key) {
  const cache = CacheService.getScriptCache();
  const metaRaw = cache.get(key + ':meta');
  if (!metaRaw) return null;
  let meta;
  try { meta = JSON.parse(metaRaw); } catch (err) { return null; }
  const keys = [];
  for (let idx = 0; idx < meta.n; idx++) keys.push(key + ':' + idx);
  const got = keys.length ? cache.getAll(keys) : {};
  let out = '';
  for (let idx = 0; idx < meta.n; idx++) {
    const piece = got[key + ':' + idx];
    if (piece === undefined || piece === null) return null; // 조각 하나라도 만료됐으면 실패로 취급(부분 조합 방지)
    out += piece;
  }
  return out;
}

// ── 시간버킷 캐시 키 (2026-08-08, 오너 결정 · R3) ──
// 이 방식은 완전한 정합성이 아니라 짧은 자연 지연을 받아들이는 단순화다 —
// reviewer-codex R1/R2에서 세대추적+명시적 무효화 방식에 동시성 결함(HIGH 2건, MEDIUM
// 1건: 세대 read-modify-write 사이 TOCTOU, 자동채움 중간상태 캐싱, 락 없는 세대증가의
// lost update)이 반복 발견되어, 이 프로젝트(소규모·저트래픽) 규모에 맞게 오너가
// 의도적으로 단순한 방식을 선택했다. 만약 실사용 중 캐시 지연이 실제로 문제가 되면,
// 완전한 수정은 ①캐시 키 자체에 세대를 붙이고(books_cache:<generation>) ②자동채움처럼
// 파생 쓰기가 있는 편집은 시작과 완료 양쪽에서 세대를 발행하며 ③세대 증가는
// ScriptLock 또는 UUID로 원자화해야 한다(2026-08-08 reviewer-codex R2 리뷰 참고).
//
// 명시적 무효화(캐시 삭제·세대 카운터) 자체를 없애고, 캐시 키에 시간버킷을 섞어 넣는다
// — 같은 버킷 안의 요청은 캐시를 공유하고, 버킷이 바뀌면(60초마다) 새 키를 쓰므로 옛
// 버킷 값은 절대 조회되지 않는다(가시 stale은 버킷 잔여시간 미만 — 60초 미만). TTL
// 70초는 옛 키가 물리적으로 살아있는 시간일 뿐 가시성을 늘리지 않는다 — 폐기 여유일
// 뿐이다. onBookSheetEdit은 캐시를 전혀 건드리지 않는다(각 시트 쓰기 완료 후 최대
// 60초 이내 반영 — 자동채움의 외부 API 처리시간 자체는 이 상한에 포함되지 않는다).
// (reviewer-codex R3 ACCEPT — 미사용 cacheRemoveChunked_ 삭제 및 이 문구 정교화 반영)
function cacheBucketKey_(base) {
  const BUCKET_SEC = 60;
  return base + ':' + Math.floor(Date.now() / (BUCKET_SEC * 1000));
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

// http(s) URL 형식이 아니면 빈 문자열로 바꾼다 — 표지URL 열은 대부분 알라딘
// API가 자동으로 채우지만 수동 입력도 가능해서, 프론트가 이스케이프해서 그리는
// 것과 별개로 서버 쪽에서도 애초에 위험한 값이 나가지 않게 막는다(근본 방어 —
// 프론트 이스케이프가 "여러 겹 방어" 중 한 겹이라면 이건 그 앞단 방어).
function sanitizeImageUrl_(url) {
  const s = String(url || '').trim();
  // 스킴(http/https)만 보는 게 아니라 "'<> 같은 속성 이탈 문자도 여기서 걸러낸다
  // — 프론트 escapeHtml()이 나중에 회귀하더라도 이 함수 하나만으로 안전하게.
  return /^https?:\/\/[^\s"'<>]+$/.test(s) ? s : '';
}

// 목록 화면(.result-desc/.rec-desc)은 CSS -webkit-line-clamp:2로 2줄만 보여주는데,
// 지금은 583권 전체의 소개 전문(약 194KB, 전체 응답의 58%)을 매번 통째로 내려보내고
// 있었다. 화면에 어차피 안 보이는 텍스트라 목록 응답에서는 짧게 잘라 보낸다 — 전체
// 텍스트는 getBookDetail()에서 별도로 지연 로딩한다(평균 소개글 길이 137자 참고, 2줄
// 클램프는 120자면 넉넉히 채우고도 남는다).
// ★String.prototype.slice(0,120)은 UTF-16 code unit 기준이라, 자르는 경계가 이모지 등
// surrogate pair 중간이면 고아 서로게이트가 생겨 깨진 문자가 만들어질 수 있다
// (reviewer-codex R1 LOW). Array.from(s)는 유니코드 코드포인트 단위로 나누므로 그
// 배열을 자르면 항상 온전한 문자 경계에서 끊긴다 — 길이 판정도 코드포인트 개수 기준으로
// 맞춘다(s.length 그대로 쓰면 판정 기준과 자르는 기준이 어긋난다).
function truncateDesc_(text, maxLen) {
  const s = String(text || '');
  const codepoints = Array.from(s);
  if (codepoints.length <= maxLen) return s;
  return codepoints.slice(0, maxLen).join('') + '…';
}

// ── 도서 목록 ──
function getBooks() {
  const cacheKey = cacheBucketKey_('books_cache');
  const cached = cacheGetChunked_(cacheKey);
  if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

  const { rows } = readSheetAsObjects(SHEET_BOOKS);
  const books = rows
    .filter(r => String(r['제목'] || '').trim())
    .map(r => ({
      id: String(r['구분']),
      title: String(r['제목'] || ''),
      author: String(r['저자'] || ''),
      cat: String(r['장르'] || ''),
      isbn: String(r['ISBN'] || ''),
      coverUrl: sanitizeImageUrl_(r['표지URL']),
      description: truncateDesc_(r['소개'], 120),
    }));
  const json = JSON.stringify({ ok: true, books });
  cachePutChunked_(cacheKey, json, 70); // 버킷(60초)보다 살짝 여유있게 — 버킷 경계 직전 요청도 안전하게 만료
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── 도서 상세(전체 소개글) — 목록에서 잘린 소개글을 상세보기 모달에서만 온전히 보여준다 ──
function getBookDetail(id) {
  const idStr = String(id || '').trim();
  if (!idStr) return jsonOut({ ok: false, error: '잘못된 요청입니다.' });
  const { rows } = readSheetAsObjects(SHEET_BOOKS);
  const r = rows.find(row => String(row['구분']) === idStr);
  if (!r) return jsonOut({ ok: false, error: '해당 도서를 찾을 수 없습니다.' });
  return jsonOut({
    ok: true,
    book: {
      id: String(r['구분']),
      title: String(r['제목'] || ''),
      author: String(r['저자'] || ''),
      cat: String(r['장르'] || ''),
      coverUrl: sanitizeImageUrl_(r['표지URL']),
      description: String(r['소개'] || ''),
    },
  });
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
  const normalizedPhone = normalizePhone_(phone);
  const { rows } = readSheetAsObjects(SHEET_LOANS);
  // 대여는 반납 전까지("반납완료" 이전) 계속 보여주고, 구매는 관리자가 "구매완료"
  // 처리한 뒤에만 "구매하신 도서"에 노출한다(프론트 index.html의 rentalLoans/purchases
  // 분리 로직이 이 전제를 그대로 따른다 — 구매신청 단계는 아직 마이페이지에 안 뜬다).
  const loans = rows
    .filter(r => normalizePhone_(r['전화번호']) === normalizedPhone && (
      (r['이용형태'] === '대여' && r['상태'] !== '반납완료') ||
      (r['이용형태'] === '구매' && r['상태'] === '구매완료')
    ))
    .map(r => ({
      row: r.row,
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

// ── 이달의 추천도서 ──
// 추천도서 시트 구조:
//   A1: 연월        B1: 값
//   A2: 주제        B2: 값
//   A3: 포스터URL   (라벨만, 값은 아래 행에 여러 장 나열 가능)
//   A4~: 포스터 이미지 URL (한 줄에 1개씩, 여러 장 가능)
//   ...: 관리번호    (라벨만)
//   그 아래: 추천도서 관리번호 목록 (한 줄에 1개씩)
function getRecommend() {
  const cacheKey = cacheBucketKey_('recommend_cache');
  const cached = cacheGetChunked_(cacheKey);
  if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECOMMEND);
  if (!sheet) return jsonOut({ ok: true, month: '', theme: '', posterUrls: [], books: [] });

  const values = sheet.getDataRange().getValues();
  const meta = {};
  const labelRows = {};
  for (let i = 0; i < values.length; i++) {
    const label = String(values[i][0] || '').trim();
    if (label === '연월') meta.month = formatMonth_(values[i][1]);
    if (label === '주제') meta.theme = String(values[i][1] || '');
    if (label === '포스터URL') labelRows.posterStart = i + 1;
    if (label === '관리번호') labelRows.idsStart = i + 1;
  }

  // 라벨을 못 찾으면(오타·행 삭제 등) 조용히 빈 값으로 넘어가지 않고 원인을
  // 실행 기록(Executions 로그)에 남긴다 — "추천도서가 안 뜬다" 문의가 오면 여기부터 확인.
  if (labelRows.posterStart === undefined) {
    console.error('추천도서 시트: "포스터URL" 라벨을 A열에서 못 찾았습니다. 철자·공백을 확인하세요.');
  }
  if (labelRows.idsStart === undefined) {
    console.error('추천도서 시트: "관리번호" 라벨을 A열에서 못 찾았습니다. 철자·공백을 확인하세요.');
  }

  // 포스터URL/관리번호 두 목록 구간은 시트에 어느 라벨이 먼저 오든 서로의 다음
  // 라벨 시작 행 앞까지로 계산한다 — 라벨 순서가 바뀌어도 깨지지 않게.
  const listStarts = [labelRows.posterStart, labelRows.idsStart].filter(v => v !== undefined).sort((a, b) => a - b);
  function rangeEnd_(start) {
    const next = listStarts.find(v => v > start);
    return next !== undefined ? next - 1 : values.length;
  }

  const posterUrls = [];
  if (labelRows.posterStart !== undefined) {
    const end = rangeEnd_(labelRows.posterStart);
    for (let i = labelRows.posterStart; i < end; i++) {
      const v = sanitizeImageUrl_(values[i][0]);
      if (v) posterUrls.push(v);
    }
  }

  const ids = getRecommendBookIds_();

  const { rows: bookRows } = readSheetAsObjects(SHEET_BOOKS);
  const byId = {};
  bookRows.forEach(r => { byId[String(r['구분'])] = r; });

  const books = ids
    .map(id => {
      const b = byId[id];
      if (!b) console.error('추천도서 시트의 관리번호 "' + id + '"에 해당하는 책을 시트1에서 찾을 수 없습니다 — 오타이거나 시트1에서 삭제된 책일 수 있습니다.');
      return b;
    })
    .filter(r => r)
    .map(r => ({
      id: String(r['구분']),
      title: r['제목'],
      author: r['저자'],
      isbn: String(r['ISBN'] || ''),
      coverUrl: sanitizeImageUrl_(r['표지URL']),
      description: String(r['소개'] || ''),
    }));

  const json = JSON.stringify({
    ok: true,
    month: meta.month || '',
    theme: meta.theme || '',
    posterUrls,
    books,
  });
  cachePutChunked_(cacheKey, json, 70); // 버킷(60초)보다 살짝 여유있게 — 버킷 경계 직전 요청도 안전하게 만료
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── 관리자 인증 ──
function getAdminPassword() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
}

// 실패가 ADMIN_FAIL_LIMIT회 쌓이면 ADMIN_LOCK_MINUTES분간 잠근다(무차별 대입 방어).
// ★호출부가 반드시 락(withAdminLock_)을 쥔 상태에서 불러야 한다 — 이 함수 자체는
// 락을 잡지 않는다(read-modify-write를 호출부의 임계구역 안에 두어 원자성을 보장하고,
// 관리자 액션 본문까지 같은 락으로 묶어 인증-확인과 실제 변경 사이 경합도 막는다).
// 반환값: { ok: true } 또는 { ok: false, error: '...' }
function adminAuthCheck(pw) {
  const props = PropertiesService.getScriptProperties();
  const lockUntil = Number(props.getProperty('ADMIN_LOCK_UNTIL') || 0);
  if (Date.now() < lockUntil) {
    const minutesLeft = Math.ceil((lockUntil - Date.now()) / 60000);
    return { ok: false, error: '비밀번호를 너무 많이 틀렸어요. ' + minutesLeft + '분 후 다시 시도해주세요.' };
  }

  const real = getAdminPassword();
  const matched = !!real && pw === real;
  if (matched) {
    props.deleteProperty('ADMIN_FAIL_COUNT');
    props.deleteProperty('ADMIN_LOCK_UNTIL');
    return { ok: true };
  }

  const fails = Number(props.getProperty('ADMIN_FAIL_COUNT') || 0) + 1;
  if (fails >= ADMIN_FAIL_LIMIT) {
    props.setProperty('ADMIN_LOCK_UNTIL', String(Date.now() + ADMIN_LOCK_MINUTES * 60 * 1000));
    props.deleteProperty('ADMIN_FAIL_COUNT');
    return { ok: false, error: '비밀번호를 너무 많이 틀렸어요. ' + ADMIN_LOCK_MINUTES + '분 후 다시 시도해주세요.' };
  }
  props.setProperty('ADMIN_FAIL_COUNT', String(fails));
  return { ok: false, error: '비밀번호가 올바르지 않아요.' };
}

// 토큰 검증 전용 — CacheService 존재 여부만 본다. ★ADMIN_FAIL_COUNT/ADMIN_LOCK_UNTIL
// (PropertiesService, adminAuthCheck 전용)은 이 함수 안 어디에서도 읽거나 쓰지 않는다 —
// 토큰 만료·미제출은 무차별 대입 시도가 아니므로 로그인 실패 카운터·5분 잠금에 영향을
// 주면 안 된다(이 프로젝트에서 실수로 5분 잠금을 두 번 유발한 전례가 있어 특히 분리한다).
function adminAuthCheckToken_(token) {
  if (!token) return { ok: false, error: '세션이 만료됐어요. 다시 로그인해주세요.' };
  const cached = CacheService.getScriptCache().get(token);
  if (!cached) return { ok: false, error: '세션이 만료됐어요. 다시 로그인해주세요.' };
  return { ok: true };
}

// 관리자 액션 공통 진입점 — 인증 확인과 실제 시트 변경을 하나의 락 임계구역으로 묶는다.
// 이걸 안 하면 ①동시 요청이 실패 카운터를 덮어써 무차별 대입 제한이 무뎌지고 ②반납
// 신청과 관리자 처리가 서로 끼어들어 상태가 역행할 수 있다(예: 반납완료가 반납신청으로
// 되돌아감).
// auth = { pw, token } — 최초 로그인은 pw로, 이후 액션은 token으로 인증한다. token이
// 있으면 adminAuthCheckToken_(카운터 미접촉)로, 없으면 adminAuthCheck(pw)(카운터 접촉)로
// 검증한다. fn은 (viaToken) 하나를 받는다 — viaToken=false면 방금 pw로 새로 로그인한
// 것이므로 호출부가 새 토큰을 발급해도 된다는 신호다.
// ★분기는 반드시 '필드 존재 여부'로 한다 — '값이 truthy한가'로 하지 않는다(reviewer-codex
// R1 REVISE). token 필드가 payload에 있으면(null·빈문자열이라도) 무조건 token 경로로
// 보내 무효 토큰은 그 안에서 거절한다(카운터 미접촉 유지).
// ★3-way 분기(reviewer-codex R2 REVISE): token/pw 필드가 '둘 다 없는' 경우를 pw 경로로
// 합쳐버리면 adminAuthCheck(undefined)가 호출돼 오답 취급으로 카운터가 오염된다. 이
// 백엔드 URL은 이제 GitHub Pages로 공개돼 있어(발견 난이도 하락), 인증 필드를 아예
// 빼고 admin_confirm_pay 등을 5회 찌르는 것만으로 진짜 관리자를 5분 잠글 수 있는
// 구멍이었다. 그래서 ①token 필드 있음→adminAuthCheckToken_(카운터 미접촉)
// ②token 없고 pw 필드 있음→adminAuthCheck(pw)(카운터 접촉, 기존 로그인 동작)
// ③둘 다 없음→adminAuthCheck 자체를 호출하지 않고 카운터 미접촉인 채로 즉시 거절 —
// 세 갈래로 명시적으로 나눈다.
function withAdminLock_(auth, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return jsonOut({ ok: false, error: '지금 요청이 몰리고 있어요. 잠시 후 다시 시도해주세요.' });
  }
  try {
    const hasToken = !!(auth && Object.prototype.hasOwnProperty.call(auth, 'token'));
    const hasPw = !!(auth && Object.prototype.hasOwnProperty.call(auth, 'pw'));
    let result;
    if (hasToken) {
      result = adminAuthCheckToken_(auth.token);
    } else if (hasPw) {
      result = adminAuthCheck(auth.pw);
    } else {
      result = { ok: false, error: '인증 정보가 없습니다.' };
    }
    if (!result.ok) return jsonOut({ ok: false, error: result.error });
    return fn(hasToken);
  } finally {
    lock.releaseLock();
  }
}

function getAdminRecords(payload) {
  return withAdminLock_(payload, function (viaToken) {
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
    const result = { ok: true, records };
    if (!viaToken) {
      // pw로 방금 새로 로그인한 경우에만 새 토큰을 발급한다(토큰 경로 재조회는 재발급하지 않음).
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put(token, '1', ADMIN_TOKEN_TTL_SEC);
      result.token = token;
    }
    return jsonOut(result);
  });
}

// 010/011/016~019로 시작하는 국내 휴대폰 번호 형식만 통과시킨다(하이픈 유무 무관).
// 프론트에도 동일한 검증이 있지만, POST body는 클라이언트가 임의로 보낼 수 있으므로
// 서버에서도 반드시 확인한다.
function isValidPhone_(phone) {
  return /^01[016789]-?\d{3,4}-?\d{4}$/.test(String(phone || '').trim());
}

// 숫자만 남긴다 — "010-1234-5678"과 "01012345678"을 같은 사람으로 인식하게 해서,
// 하이픈 표기만 바꿔가며 대여 3권 한도를 우회하는 것을 막는다. 시트에 과거부터
// 섞여 있는 표기(하이픈 있음/없음)도 비교 시점에 이걸 거치면 동일하게 취급된다.
function normalizePhone_(phone) {
  return String(phone || '').replace(/\D/g, '');
}

// 이달의 추천도서 관리번호 목록만 필요할 때 쓰는 경량 버전 — getRecommend()와
// 같은 파싱 로직을 공유한다(로직 두 곳에 따로 두면 드리프트 위험).
function getRecommendBookIds_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECOMMEND);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const labelRows = {};
  for (let i = 0; i < values.length; i++) {
    const label = String(values[i][0] || '').trim();
    if (label === '포스터URL') labelRows.posterStart = i + 1;
    if (label === '관리번호') labelRows.idsStart = i + 1;
  }
  if (labelRows.idsStart === undefined) return [];
  const listStarts = [labelRows.posterStart, labelRows.idsStart].filter(v => v !== undefined).sort((a, b) => a - b);
  const next = listStarts.find(v => v > labelRows.idsStart);
  const end = next !== undefined ? next - 1 : values.length;
  const ids = [];
  for (let i = labelRows.idsStart; i < end; i++) {
    const v = String(values[i][0] || '').trim();
    if (v) ids.push(v);
  }
  return ids;
}

// ── 대여/구매 신청 ──
function applyBooks(payload) {
  const name = String(payload.name || '').trim();
  const rawPhone = String(payload.phone || '').trim();
  const type = payload.type;
  const items = payload.items || [];

  if (!name || !rawPhone) return jsonOut({ ok: false, error: '이름과 연락처를 입력해주세요.' });
  if (!isValidPhone_(rawPhone)) return jsonOut({ ok: false, error: '전화번호 형식을 확인해주세요. 예: 010-1234-5678' });
  if (!items.length) return jsonOut({ ok: false, error: '선택한 도서가 없어요.' });
  if (type !== '대여' && type !== '구매') return jsonOut({ ok: false, error: '이용 형태가 올바르지 않아요.' });

  // "010-1234-5678"과 "01012345678"을 같은 사람으로 취급한다 — 하이픈 표기만
  // 바꿔가며 대여 한도(전화번호별 집계)를 우회하지 못하게. 시트에도 이 정규화된
  // 형태로 저장해서 다음 조회 때도 계속 같은 사람으로 인식되게 한다.
  const phone = normalizePhone_(rawPhone);

  // 클라이언트가 보낸 도서명·저자를 그대로 믿지 않는다 — 실제 시트1과 대조해
  // 서버가 아는 값으로 덮어쓴다. bookId가 실존하지 않으면 신청 자체를 막는다.
  // (위조된 bookId나 대여기록에 임의 텍스트가 남는 것을 원천 차단 — 관리자 화면
  // 등에서 그 값을 보여줄 때 XSS 방어를 여러 겹으로 하는 것보다 근본적인 방어.)
  const { rows: bookRows } = readSheetAsObjects(SHEET_BOOKS);
  const bookById = {};
  bookRows.forEach(r => { bookById[String(r['구분'])] = r; });

  const resolvedItems = [];
  const seenBookIds = new Set();
  for (const it of items) {
    const bookId = String(it.bookId);
    if (seenBookIds.has(bookId)) {
      return jsonOut({ ok: false, error: '같은 책이 목록에 두 번 포함되어 있어요. 새로고침 후 다시 시도해주세요.' });
    }
    seenBookIds.add(bookId);
    const book = bookById[bookId];
    if (!book) return jsonOut({ ok: false, error: '선택한 도서 중 존재하지 않는 항목이 있어요. 새로고침 후 다시 시도해주세요.' });
    resolvedItems.push({
      bookId: bookId,
      title: String(book['제목'] || ''),
      author: String(book['저자'] || ''),
    });
  }

  // "이미 대여중인지/권수초과인지 확인 → 기록 추가" 사이에 다른 신청이 끼어들면
  // 같은 책이 동시에 두 사람에게 대여 처리될 수 있어 잠근다(레이스 컨디션 방지).
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return jsonOut({ ok: false, error: '지금 신청이 몰리고 있어요. 잠시 후 다시 시도해주세요.' });
  }
  try {
    if (type === '대여') {
      const activeLoans = getActiveLoanRows();
      const activeForPhone = activeLoans.filter(r => normalizePhone_(r['전화번호']) === phone).length;
      if (activeForPhone + resolvedItems.length > MAX_RENTAL) {
        return jsonOut({ ok: false, error: '대여는 1인당 최대 ' + MAX_RENTAL + '권까지 가능해요.' });
      }
      const activeBookIds = new Set(activeLoans.map(r => String(r['도서관리번호'])));
      const alreadyOut = resolvedItems.find(it => activeBookIds.has(it.bookId));
      if (alreadyOut) {
        return jsonOut({ ok: false, error: '이미 대여 중인 도서가 포함되어 있어요. 새로고침 후 다시 시도해주세요.' });
      }

      // 이달의 추천도서는 1인당 MAX_REC권까지만 — "이미 대여 중인 추천도서 수"와
      // "이번에 새로 신청하는 추천도서 수"를 합산해서 확인해야 한다(따로 두 번
      // 나눠 신청하는 우회를 막으려면). 활성 대출 스냅샷과 같은 락 안에서 계산해야
      // 정확하다(락 밖에서 계산하면 그 사이 다른 신청이 끼어들 수 있다).
      const recIds = new Set(getRecommendBookIds_());
      const activeRecForPhone = activeLoans.filter(r => normalizePhone_(r['전화번호']) === phone && recIds.has(String(r['도서관리번호']))).length;
      const newRecCount = resolvedItems.filter(it => recIds.has(it.bookId)).length;
      if (activeRecForPhone + newRecCount > MAX_REC) {
        return jsonOut({ ok: false, error: '이달의 추천도서는 1인당 최대 ' + MAX_REC + '권까지 대여할 수 있어요.' });
      }
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
    const now = new Date();
    const dueDate = new Date(now.getTime() + RENTAL_DAYS * 24 * 60 * 60 * 1000);

    resolvedItems.forEach(it => {
      sheet.appendRow([
        Utilities.getUuid(),
        name,
        // 앞의 '(작은따옴표)는 시트에 저장은 안 되고 "이 값은 숫자로 변환하지 말고
        // 텍스트 그대로 저장하라"는 표시다 — 안 붙이면 스프레드시트가 "010-..."을
        // 숫자로 인식해 맨 앞 0을 지워버린다. 그러면 이후 모든 조회가
        // normalizePhone_()로 재정규화한 값과 맞지 않아 "내 대여목록"·중복대여
        // 검사가 전부 깨진다(단순 표시 문제가 아니라 매칭 자체가 틀어짐).
        "'" + phone,
        it.bookId,
        it.title,
        it.author,
        type,
        now,
        type === '대여' ? dueDate : '',
        type === '대여' ? '대여중' : '구매신청',
        '',
        '',
        'N',
      ]);
    });

    return jsonOut({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// ── 반납 신청 ──
function requestReturn(payload) {
  const phone = normalizePhone_(payload.phone || '');
  const bookId = String(payload.bookId || '');
  const rowNum = payload.row ? Number(payload.row) : null;
  const returnDate = String(payload.returnDate || '');
  const returnTimeSlot = String(payload.returnTimeSlot || '');

  // 서버도 날짜·시간대 값을 확인한다 — 프론트가 과거 날짜·임의 문자열을 못 보내게
  // 막고 있지만, POST body는 클라이언트가 임의로 구성할 수 있다.
  // 엄격한 YYYY-MM-DD 형식만 통과시킨다 — "2026-1-1"처럼 0-패딩 안 된 값은
  // 문자열 비교("2026-1-1" > "2026-08-06")에서 미래 날짜로 잘못 판정될 수 있었다.
  // 형식이 맞아도 "2027-02-30"처럼 실존하지 않는 날짜는 JS Date가 3월 2일로
  // 조용히 보정해버려서 형식·과거날짜 검사를 다 통과할 수 있다 — round-trip으로
  // (문자열→Date→다시 문자열) 원래 값과 같은지 대조해 이런 값도 거른다.
  if (!/^\d{4}-\d{2}-\d{2}$/.test(returnDate)) {
    return jsonOut({ ok: false, error: '반납 예정 일시를 올바르게 선택해주세요.' });
  }
  const parsedDate = new Date(returnDate + 'T00:00:00');
  const roundTrip = parsedDate.getFullYear() + '-' +
    String(parsedDate.getMonth() + 1).padStart(2, '0') + '-' +
    String(parsedDate.getDate()).padStart(2, '0');
  if (isNaN(parsedDate.getTime()) || roundTrip !== returnDate) {
    return jsonOut({ ok: false, error: '반납 예정 일시를 올바르게 선택해주세요.' });
  }
  // 서울(한국) 기준 오늘 날짜와 비교한다 — UTC 기준으로 비교하면 자정부터
  // 오전 9시 사이엔 한국은 이미 다음 날인데 하루 전으로 오판될 수 있다.
  const todaySeoul = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  if (returnDate < todaySeoul) {
    return jsonOut({ ok: false, error: '반납 예정일은 오늘 이후 날짜로 선택해주세요.' });
  }
  if (VALID_RETURN_SLOTS.indexOf(returnTimeSlot) === -1) {
    return jsonOut({ ok: false, error: '방문 시간대를 올바르게 선택해주세요.' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return jsonOut({ ok: false, error: '지금 요청이 몰리고 있어요. 잠시 후 다시 시도해주세요.' });
  }
  try {
    const { sheet, rows } = readSheetAsObjects(SHEET_LOANS);
    // row(시트 고유 행 번호)가 있으면 그걸로 정확히 특정하되, bookId도 반드시
    // 함께 대조한다 — bookId를 비워 보내는 방식으로 이 대조를 생략시킬 수
    // 없도록(동일 전화번호의 다른 대여건을 잘못 반납 처리하는 경로 차단).
    // row가 없는 옛 클라이언트 호출만 전화번호+bookId 방식으로 폴백한다.
    const target = rowNum
      ? (bookId
          ? rows.find(r => r.row === rowNum && normalizePhone_(r['전화번호']) === phone && r['상태'] === '대여중' && String(r['도서관리번호']) === bookId)
          : null)
      : rows.find(r => normalizePhone_(r['전화번호']) === phone && String(r['도서관리번호']) === bookId && r['상태'] === '대여중');
    if (!target) return jsonOut({ ok: false, error: '대여 중인 도서를 찾을 수 없어요.' });

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const statusCol = headers.indexOf('상태') + 1;
    const reqDateCol = headers.indexOf('반납신청일시') + 1;
    const reqSlotCol = headers.indexOf('반납신청시간대') + 1;

    sheet.getRange(target.row, statusCol).setValue('반납신청');
    sheet.getRange(target.row, reqDateCol).setValue(returnDate);
    sheet.getRange(target.row, reqSlotCol).setValue(returnTimeSlot);

    return jsonOut({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// ── 관리자: 입금 확인 ──
function adminConfirmPay(payload) {
  return withAdminLock_(payload, function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const col = headers.indexOf('입금확인') + 1;
    sheet.getRange(payload.row, col).setValue('Y');
    return jsonOut({ ok: true });
  });
}

// ── 관리자: 반납 완료 처리 ──
function adminCompleteReturn(payload) {
  return withAdminLock_(payload, function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const col = headers.indexOf('상태') + 1;
    sheet.getRange(payload.row, col).setValue('반납완료');
    return jsonOut({ ok: true });
  });
}

// ── 관리자: 구매 완료 처리 ──
function adminCompletePurchase(payload) {
  return withAdminLock_(payload, function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOANS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const col = headers.indexOf('상태') + 1;
    sheet.getRange(payload.row, col).setValue('구매완료');
    return jsonOut({ ok: true });
  });
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
    // 트리거 오류가 시트 편집 자체를 막지는 않되(그래서 여기서 다시 던지지 않음),
    // 원인은 실행 기록에 남긴다 — "ISBN 입력했는데 표지가 안 채워진다" 문의 시 확인용.
    console.error('onBookSheetEdit 오류: ' + err);
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
  // 알라딘 API 응답을 사람 검수 없이 그대로 저장하므로, 저장 시점에 형식을
  // 확인한다 — http(s) URL이 아니면 애초에 셀에 쓰지 않는다.
  const safeCover = sanitizeImageUrl_(info.cover);
  if (!cover && safeCover) sheet.getRange(row, coverCol).setValue(safeCover);
  if (!desc && info.description) sheet.getRange(row, descCol).setValue(info.description);
  return true;
}

// 실패해도 항상 null만 반환하지만(호출부는 "표지 없음"으로 취급, 다음 실행 때
// 자동 재시도됨), 원인별로 실행 기록(Executions 로그)에 남겨서 "왜 이 책만
// 표지가 안 채워지나"를 admin이 나중에 로그로 구분할 수 있게 한다 —
// 알라딘 서버 오류/요청제한(일시적, 재시도하면 될 일)인지, 정말 알라딘에
// 없는 책(독립출판 등, ISBN을 다시 확인해야 할 일)인지 구분.
function lookupAladin(isbn) {
  const ttbkey = PropertiesService.getScriptProperties().getProperty('ALADIN_TTBKEY');
  if (!ttbkey) {
    console.error('ALADIN_TTBKEY가 스크립트 속성에 없습니다 — 표지·소개 자동화가 전부 동작하지 않습니다. 재시도로 해결되지 않으니 스크립트 속성을 확인하세요.');
    return null;
  }
  const url = 'https://www.aladin.co.kr/ttb/api/ItemLookUp.aspx?ttbkey=' + ttbkey +
    '&ItemId=' + isbn + '&ItemIdType=ISBN13&output=js&Version=20131101&Cover=Big';
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 401 || code === 403) {
      console.error('알라딘 API 인증 오류 (ISBN ' + isbn + '): HTTP ' + code + ' — TTBKey가 잘못됐거나 만료됐을 가능성이 높습니다. 재시도로는 해결되지 않으니 키를 재발급·재확인하세요.');
      return null;
    }
    if (code !== 200) {
      console.error('알라딘 API 응답 오류 (ISBN ' + isbn + '): HTTP ' + code + ' — 일시적 오류일 수 있음, 다음 실행 때 자동 재시도됨');
      return null;
    }
    const data = JSON.parse(res.getContentText());
    if (data.errorCode) {
      console.error('알라딘 API 에러 (ISBN ' + isbn + '): ' + data.errorCode + ' ' + (data.errorMessage || ''));
      return null;
    }
    const item = data.item && data.item[0];
    if (!item) {
      console.error('알라딘에 해당 ISBN이 없습니다: ' + isbn + ' (독립출판·자체제작물 등일 수 있음 — ISBN 재확인 필요)');
      return null;
    }
    return { cover: item.cover, description: item.description };
  } catch (err) {
    console.error('알라딘 API 호출 실패 (ISBN ' + isbn + '): ' + err);
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
