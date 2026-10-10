/**
 * 도서신청 자동화 웹앱 (교보문고 검색 → 신청 → 학교도서관 소장 확인 → 주문서)
 * 구두방장선생 · Google Apps Script
 *
 * [설치]  자세한 내용은 배포 매뉴얼 참고
 *  1) 사본 만들기 → 시트 메뉴 [📚 도서신청 > 처음 세팅]
 *  2) [학교도서관 연결하기] 에 read365 학교 검색 결과 주소 붙여넣기
 *  3) 웹앱설정 시트에서 관리자비밀번호 변경 → [설정 점검]
 *  4) Apps Script > 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자)
 */

/* ===================== 설정 ===================== */
var VERSION = '1.1.0';
var VERSION_DATE = '2026-10-07';
var AUTHOR = '구두방장선생';
var AUTHOR_URL = 'https://blog.naver.com/qkfkekdl';
// 새 버전 알림: 최신 버전 정보를 담은 JSON 주소 (비우면 확인 안 함)
// 형식: {"latest":"1.2.0","date":"2026-11-01","url":"블로그 글 주소","notes":"바뀐 점 한 줄"}
var UPDATE_URL = 'https://raw.githubusercontent.com/GDB5124/book-request-webapp/main/version.json';
var DEFAULT_PW = '1234';
var SHEET_ORDERS = '신청내역';
var SHEET_CONF = '웹앱설정';
var SHEET_SUMMARY = '주문서';

var ORDER_HEADERS = ['신청일시', '신청자', '수량', '도서명', '출판사', '지은이', '할인가', '정가',
  '금액', 'ISBN', '상품코드', '출간일', '교보링크', '상태', '신청ID', '학교소장'];
var COL = {}; ORDER_HEADERS.forEach(function (h, i) { COL[h] = i; });

var DEFAULT_CONF = [
  ['제목', '도서신청', '웹앱 상단에 표시되는 제목 (예: 2026 ○○고 도서신청)'],
  ['신청받기', 'Y', 'N 으로 바꾸면 신청이 막힙니다'],
  ['마감일시', '', '예: 2026-10-10 18:00 (비워두면 마감 없음)'],
  ['금액기준', '정가', '정가 또는 할인가 (주문서 금액 계산 기준)'],
  ['총예산', 0, '0 이면 예산 표시 안 함'],
  ['검색대상', 'KOR', 'KOR=국내도서, ENG=외국도서, POD=POD (쉼표로 여러 개)'],
  ['관리자비밀번호', DEFAULT_PW, '웹앱 [현황] 탭 비밀번호 — 1234 그대로면 현황 탭이 열리지 않습니다'],
  ['안내문구', '교보문고에서 책을 검색해서 신청하세요. 금액은 교보문고 기준입니다.', '신청 화면 상단 안내'],
  ['도서관_학교명', '', '메뉴 [학교도서관 연결하기]로 자동 입력 (read365 주소의 schoolName)'],
  ['도서관_지역코드', '', '메뉴 [학교도서관 연결하기]로 자동 입력 (read365 주소의 provCode)'],
  ['도서관_학교코드', '', '메뉴 [학교도서관 연결하기]로 자동 입력 (read365 주소의 neisCode) — 비우면 소장 확인 끔']
];

var UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36';

/* ===================== 웹앱 진입 ===================== */
function doGet() {
  var t = HtmlService.createTemplateFromFile('index');
  return t.evaluate()
    .setTitle(getConf_()['제목'] || '도서신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📚 도서신청')
    .addItem('① 처음 세팅', 'setup')
    .addItem('② 학교도서관 연결하기', 'connectLibrary')
    .addItem('③ 설정 점검', 'checkSetup')
    .addSeparator()
    .addItem('주문서 시트 만들기', 'buildSummarySheet')
    .addItem('웹앱 주소 보기', 'showWebAppUrl')
    .addItem('버전 정보', 'showVersion')
    .addToUi();
}

/* ===================== 초기 세팅 ===================== */
function setup() {
  var ss = SpreadsheetApp.getActive();

  var o = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS);
  if (o.getLastRow() === 0) {
    o.appendRow(ORDER_HEADERS);
    o.getRange(1, 1, 1, ORDER_HEADERS.length).setFontWeight('bold').setBackground('#e8f0fe');
    o.setFrozenRows(1);
    o.getRange('G:I').setNumberFormat('#,##0');
    o.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  }

  var c = ss.getSheetByName(SHEET_CONF) || ss.insertSheet(SHEET_CONF);
  if (c.getLastRow() === 0) {
    c.appendRow(['항목', '값', '설명']);
    c.getRange(2, 1, DEFAULT_CONF.length, 3).setValues(DEFAULT_CONF);
    c.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
    c.setColumnWidth(1, 140); c.setColumnWidth(2, 260); c.setColumnWidth(3, 360);
  }
  SpreadsheetApp.getActive().toast('시트 준비 완료! 다음은 [② 학교도서관 연결하기] 입니다.', '도서신청', 8);
  if (!libConf_().neis) {
    try { connectLibrary(); } catch (e) { /* 편집기에서 실행한 경우 등 UI 없음 */ }
  }
}

/** 웹앱설정 시트의 값 하나를 바꿈 (없으면 행 추가) */
function setConf_(key, value) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONF);
  if (!sh) { setup(); sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONF); }
  var n = sh.getLastRow();
  var keys = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues() : [];
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  }
  var desc = (DEFAULT_CONF.filter(function (r) { return r[0] === key; })[0] || [])[2] || '';
  sh.appendRow([key, value, desc]);
}

/** read365 학교 검색 결과 주소 → 지역코드·학교코드·학교명 */
function parseRead365Url_(url) {
  var q = {};
  String(url || '').replace(/^[^?]*\?/, '').split('&').forEach(function (kv) {
    var i = kv.indexOf('=');
    if (i < 0) return;
    try { q[kv.slice(0, i)] = decodeURIComponent(kv.slice(i + 1).replace(/\+/g, ' ')); } catch (e) { }
  });
  return { prov: (q.provCode || '').trim(), neis: (q.neisCode || '').trim(), school: (q.schoolName || '').trim() };
}

/** 메뉴: 학교도서관 연결하기 */
function connectLibrary() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('② 학교도서관 연결하기',
    'read365(독서교육종합지원시스템)에서 우리 학교를 고른 뒤 아무 책이나 검색하고,\n' +
    '그 검색 결과 화면의 주소를 통째로 복사해서 붙여넣으세요.\n\n' +
    '예) https://read365.edunet.net/PureScreen/SchoolSearchResult?searchKeyword=...&provCode=...&neisCode=...&schoolName=...\n\n' +
    '학교도서관 확인을 쓰지 않으려면 [취소]를 누르세요.',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var lc = parseRead365Url_(r.getResponseText());
  if (!lc.prov || !lc.neis) {
    ui.alert('주소에서 provCode, neisCode 를 찾지 못했습니다.\n학교를 고른 뒤 "검색 결과" 화면의 주소인지 확인해 주세요.');
    return;
  }
  setConf_('도서관_지역코드', lc.prov);
  setConf_('도서관_학교코드', lc.neis);
  if (lc.school) setConf_('도서관_학교명', lc.school);
  var t = checkLibraryConnection_();
  ui.alert((t.ok ? '✅ ' : '⚠️ ') + t.msg);
}

/** 학교코드가 맞는지: 흔한 낱말로 검색해서 우리 학교 책이 나오는지 */
function checkLibraryConnection_() {
  var lc = libConf_();
  if (!lc.neis || !lc.prov) return { ok: false, msg: '학교도서관이 연결되지 않았습니다 (소장 확인 꺼짐).' };
  var words = ['책', '이야기', '수학'];
  for (var i = 0; i < words.length; i++) {
    var list = libParseList_(libFetch_(libSearchRequest_(words[i], lc, 5)));
    if (list.length) {
      var name = list[0].schoolName || list[0].libName || lc.school;
      return { ok: true, msg: '학교도서관 연결됨: ' + name + ' (' + lc.prov + ' / ' + lc.neis + ')' };
    }
  }
  return { ok: false, msg: '학교도서관 검색 결과가 없습니다. 코드(' + lc.prov + ' / ' + lc.neis + ')가 맞는지 확인하세요.' };
}

function getConf_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONF);
  var conf = {};
  DEFAULT_CONF.forEach(function (r) { conf[r[0]] = r[1]; });
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0] !== '') conf[String(r[0]).trim()] = r[1];
    });
  }
  return conf;
}

/* ===================== 교보문고 가져오기 ===================== */
function fetchHtml_(url) {
  var res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': UA, 'Accept-Language': 'ko-KR,ko;q=0.9' }
  });
  var code = res.getResponseCode();
  if (code !== 200) throw new Error('교보문고 응답 오류 (HTTP ' + code + ')');
  return res.getContentText('UTF-8');
}

function cacheGet_(key) {
  try { var v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; }
}
function cachePut_(key, obj, sec) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(obj), sec || 21600); } catch (e) { }
}

/** 검색어(제목·저자·ISBN) 또는 교보 링크 → 결과 목록 */
function searchBooks(query) {
  query = String(query || '').trim();
  if (!query) return { items: [] };

  // 교보 링크나 상품코드를 붙여넣은 경우 → 바로 상세
  var pidInLink = query.match(/S\d{12}/);
  if (pidInLink) {
    var b = getBook(pidInLink[0]);
    return { items: [b], direct: true };
  }

  var conf = getConf_();
  var types = String(conf['검색대상'] || 'KOR').split(/[,\s]+/).filter(String);
  var key = 'q:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, query + '|' + types.join(','), Utilities.Charset.UTF_8));
  var cached = cacheGet_(key);
  if (cached) return { items: cached };

  var url = 'https://search.kyobobook.co.kr/search?keyword=' + encodeURIComponent(query) + '&gbCode=TOT&target=total';
  var items = parseSearchHtml_(fetchHtml_(url), types).slice(0, 20);
  cachePut_(key, items, 3600);
  return { items: items };
}

/** 상품코드 → 상세 정보 (가격·저자·출판사 확정) */
function getBook(pid) {
  pid = String(pid || '').trim();
  if (!/^S\d{12}$/.test(pid)) throw new Error('상품코드 형식이 아닙니다: ' + pid);
  var cached = cacheGet_('d:' + pid);
  if (cached) return cached;
  var book = parseDetailHtml_(fetchHtml_('https://product.kyobobook.co.kr/detail/' + pid), pid);
  if (!book.ok) throw new Error('책 정보를 읽지 못했습니다. 교보문고 페이지 구조가 바뀌었을 수 있어요. (' + pid + ')');
  cachePut_('d:' + pid, book, 21600);
  return book;
}

/** 메뉴: 설정 점검 — 배포 전 한 번, 이상할 때마다 실행 */
function checkSetup() {
  var lines = [], bad = 0;
  function ok(cond, good, fix) { lines.push((cond ? '✅ ' : '❌ ') + (cond ? good : fix)); if (!cond) bad++; }
  var conf = getConf_();

  ok(Session.getScriptTimeZone() === 'Asia/Seoul', '시간대: Asia/Seoul',
    '시간대가 ' + Session.getScriptTimeZone() + ' 입니다 → Apps Script > 프로젝트 설정에서 (GMT+09:00) 서울로 바꾸세요.');
  ok(String(conf['관리자비밀번호']) !== DEFAULT_PW && String(conf['관리자비밀번호']).trim() !== '',
    '관리자 비밀번호 변경됨', '관리자 비밀번호가 초기값(1234)입니다 → 웹앱설정 시트에서 바꾸세요.');

  var due = conf['마감일시'];
  var dueD = parseDue_(due);
  var dueOk = dueD === null || !isNaN(dueD);
  ok(dueOk, '마감일시: ' + (due === '' ? '없음' : (due instanceof Date ? Utilities.formatDate(due, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : due)),
    '마감일시 형식이 이상합니다 → 2026-10-10 18:00 처럼 쓰세요.');
  lines.push((String(conf['신청받기']).toUpperCase() === 'Y' ? '✅ 신청받기: Y (신청 가능)' : 'ℹ️ 신청받기: N (지금은 신청이 막혀 있음)'));

  var book = null;
  try {
    var r = searchBooks('구글 시트');
    ok(r.items.length > 0, '교보문고 검색: ' + r.items.length + '건', '교보문고 검색 결과가 없습니다 → 교보 화면 구조가 바뀌었을 수 있습니다.');
    if (r.items.length) {
      book = getBook(r.items[0].pid);
      ok(book.ok && book.salePrice > 0, '교보문고 책 정보: ' + book.title + ' (' + book.salePrice + '원)', '교보문고 책 정보를 읽지 못했습니다.');
    }
  } catch (e) { ok(false, '', '교보문고 연결 실패: ' + e.message); }

  var lib = checkLibraryConnection_();
  lines.push((lib.ok ? '✅ ' : (libConf_().neis ? '❌ ' : 'ℹ️ ')) + lib.msg);
  if (!lib.ok && libConf_().neis) bad++;

  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { }
  lines.push(url ? '✅ 웹앱 배포됨' : 'ℹ️ 아직 웹앱으로 배포하지 않았습니다 (5단계).');

  var up = checkUpdate_();
  if (up.newer) lines.push('🆕 새 버전 v' + up.latest + ' 이 있습니다 (지금 v' + VERSION + ') → ' + (up.url || AUTHOR_URL));
  else if (up.latest) lines.push('✅ 최신 버전 사용 중 (v' + VERSION + ')');

  var msg = '도서신청 웹앱 v' + VERSION + ' 설정 점검\n\n' + lines.join('\n') +
    '\n\n' + (bad ? '❌ 표시를 먼저 고쳐 주세요.' : '모두 정상입니다.');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { }
  return msg;
}

/** 예전 이름 호환 */
function testKyobo() { return checkSetup(); }

/** 메뉴: 웹앱 주소 보기 */
function showWebAppUrl() {
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { }
  SpreadsheetApp.getUi().alert(url ? '교직원에게 이 주소를 공유하세요:\n\n' + url :
    '아직 웹앱으로 배포하지 않았습니다.\n확장 프로그램 > Apps Script > 배포 > 새 배포 > 웹 앱');
}

/** 메뉴: 버전 정보 */
function showVersion() {
  var up = checkUpdate_();
  var tail = up.newer ? '\n\n🆕 새 버전 v' + up.latest + (up.date ? ' (' + up.date + ')' : '') + ' 이 있습니다.' +
    (up.notes ? '\n' + up.notes : '') + '\n업데이트 방법: ' + (up.url || AUTHOR_URL)
    : (up.latest ? '\n\n✅ 최신 버전입니다.' : '');
  SpreadsheetApp.getUi().alert('도서신청 웹앱\n버전 ' + VERSION + ' (' + VERSION_DATE + ')\n만든 이: ' + AUTHOR +
    '\n블로그: ' + AUTHOR_URL + tail);
}

/** '1.10.0' > '1.9.2' 처럼 숫자로 비교 */
function compareVersion_(a, b) {
  var x = String(a).split('.'), y = String(b).split('.');
  for (var i = 0; i < 3; i++) {
    var d = (parseInt(x[i], 10) || 0) - (parseInt(y[i], 10) || 0);
    if (d) return d > 0 ? 1 : -1;
  }
  return 0;
}

/** 새 버전 확인 (6시간 캐시). 실패하거나 UPDATE_URL 이 비어 있으면 조용히 넘어감 */
function checkUpdate_() {
  if (!UPDATE_URL) return {};
  var info = cacheGet_('update');
  if (!info) {
    try {
      var res = UrlFetchApp.fetch(UPDATE_URL, { muteHttpExceptions: true, followRedirects: true });
      info = res.getResponseCode() === 200 ? JSON.parse(res.getContentText('UTF-8')) : {};
    } catch (e) { info = {}; }
    cachePut_('update', info, 21600);
  }
  var latest = String(info.latest || '');
  var url = /^https?:\/\//.test(String(info.url || '')) ? String(info.url) : '';
  return { latest: latest, date: String(info.date || ''), url: url, notes: String(info.notes || '').slice(0, 200),
    newer: !!latest && compareVersion_(latest, VERSION) > 0 };
}

/* ===================== 교보문고 파서 (순수 함수) ===================== */

function decodeEntities_(s) {
  if (!s) return '';
  return String(s)
    .replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(parseInt(d, 10)); })
    .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&');
}

function stripTags_(s) {
  return decodeEntities_(String(s || '').replace(/<!--[\s\S]*?-->/g, '').replace(/<[^>]+>/g, ' '))
    .replace(/\s+/g, ' ').trim();
}

function toNum_(s) {
  var n = parseInt(String(s || '').replace(/[^\d]/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

/** 검색 결과 HTML → [{pid, isbn, title, type, author, publisher, pubDate, salePrice, cover}] */
function parseSearchHtml_(html, allowTypes) {
  var parts = html.split('class="prod_item"');
  var out = [], seen = {};
  for (var i = 1; i < parts.length; i++) {
    var s = parts[i];
    var pid = (s.match(/data-pid="(S\d{12})"/) || [])[1];
    if (!pid || seen[pid]) continue;
    var type = (s.match(/data-code="([A-Z]+)"/) || [])[1] || '';
    if (allowTypes && allowTypes.indexOf(type) < 0) continue;
    seen[pid] = true;
    var isbn = (s.match(/data-bid="(\d{10,13})"/) || [])[1] || '';
    var title = decodeEntities_((s.match(/data-name="([^"]*)"/) || [])[1] || '');
    // 저자: <a class="author rep">이름</a>, <a class="author">이름</a> ... (팝업 영역 전까지)
    var aArea = s.split('dialog_author')[0];
    var names = [], am, are = /class="author(?: rep)?"[^>]*>([^<]+)</g;
    while ((am = are.exec(aArea)) !== null) {
      var nm = decodeEntities_(am[1]).trim();
      if (nm && names.indexOf(nm) < 0) names.push(nm);
    }
    var author = names.join(', ');
    var pub = (s.match(/class="prod_publish"[\s\S]*?class="text"[^>]*>([^<]*)<\/a>/) || [])[1] || '';
    var date = (s.match(/class="prod_publish"[\s\S]*?class="date"[^>]*>([^<]*)</) || [])[1] || '';
    var price = (s.match(/class="price">\s*<span class="val">([\d,]+)/) || [])[1] || '';
    out.push({
      pid: pid, isbn: isbn, title: title, type: type,
      author: author, publisher: decodeEntities_(pub).trim(), pubDate: date.trim(),
      salePrice: toNum_(price),
      cover: isbn ? 'https://contents.kyobobook.co.kr/sih/fit-in/200x0/pdt/' + isbn + '.jpg' : '',
      url: 'https://product.kyobobook.co.kr/detail/' + pid
    });
  }
  return out;
}

/** 상세 페이지 HTML → {pid, isbn, title, author, publisher, pubDate, salePrice, originalPrice, cover, url} */
function parseDetailHtml_(html, pid) {
  // Next.js RSC 데이터는 JS 문자열 안에 \" 로 이스케이프되어 있음 → 따옴표 복원
  var t = html.replace(/\\"/g, '"');
  var anchor = t.indexOf('"infoSaleCmdtId":"' + pid + '"');
  if (anchor < 0) anchor = t.indexOf('"salePrice":');
  var before = anchor >= 0 ? t.slice(Math.max(0, anchor - 4000), anchor) : '';
  var after = anchor >= 0 ? t.slice(anchor, anchor + 8000) : '';

  function str(re, src) {
    var m = (src || after).match(re);
    if (!m) return '';
    try { return JSON.parse('"' + m[1] + '"'); } catch (e) { return m[1]; }
  }
  function lastMatch(re, src) {
    var m, last = null, g = new RegExp(re.source, 'g');
    while ((m = g.exec(src)) !== null) last = m[1];
    return last || '';
  }

  var r = { pid: pid, url: 'https://product.kyobobook.co.kr/detail/' + pid };
  r.title = str(/"title":"((?:[^"\\]|\\.)*)"/);
  r.salePrice = toNum_((after.match(/"salePrice":(\d+)/) || [])[1]);
  r.originalPrice = toNum_((after.match(/"originalPrice":(\d+)/) || [])[1]);
  r.isbn = lastMatch(/"cmdtCode":"(\d{10,13})"/, before);

  // 저자: "authors":[ ... ] 구간의 chrcName 들 + 역할
  var aSeg = (after.match(/"authors":(\[[\s\S]*?\]\}\]),"authorSectionTitle"/) || [])[1];
  var authors = [];
  if (aSeg) {
    try {
      JSON.parse(aSeg).forEach(function (grp) {
        var role = grp.role && grp.role.roleCodeName ? grp.role.roleCodeName : '';
        var names = (grp.info || []).map(function (x) { return x.chrcName; }).filter(Boolean);
        if (names.length) authors.push(names.join(', ') + (role ? ' ' + role : ''));
      });
    } catch (e) {
      var g = /"chrcName":"([^"]*)"/g, m;
      while ((m = g.exec(aSeg)) !== null) authors.push(m[1]);
    }
  }
  r.author = authors.join(' / ');
  r.publisher = str(/"publisher":\{"label":\{"text":"((?:[^"\\]|\\.)*)"/);
  var pd = (after.match(/"pubDate":"(\d{8})"/) || [])[1] || '';
  r.pubDate = pd ? pd.slice(0, 4) + '-' + pd.slice(4, 6) + '-' + pd.slice(6, 8) : '';

  // ---- 대체 경로 (구조가 또 바뀌었을 때) ----
  if (!r.title) {
    r.title = stripTags_((html.match(/<span class="prod_title">([\s\S]*?)<\/span>/) || [])[1] ||
      ((html.match(/<meta property="og:title" content="([^"]*)"/) || [])[1] || '').replace(/\s*-\s*교보문고\s*$/, ''));
  }
  var ogImg = (html.match(/<meta property="og:image" content="([^"]*)"/) || [])[1] || '';
  if (!r.isbn) r.isbn = (ogImg.match(/pdt\/(\d{13})\./) || [])[1] || '';
  r.cover = ogImg || (r.isbn ? 'https://contents.kyobobook.co.kr/sih/fit-in/400x0/pdt/' + r.isbn + '.jpg' : '');
  if (!r.originalPrice) r.originalPrice = r.salePrice;
  r.ok = !!(r.title && r.salePrice);
  return r;
}

/* ===================== 학교도서관 소장 확인 (read365 독서교육종합지원시스템) ===================== */
var LIB_BASE = 'https://read365.edunet.net';

function libConf_() {
  var c = getConf_();
  return {
    school: String(c['도서관_학교명'] || '').trim(),
    prov: String(c['도서관_지역코드'] || '').trim(),
    neis: String(c['도서관_학교코드'] || '').trim()
  };
}

function libSearchRequest_(keyword, lc, display) {
  return {
    url: LIB_BASE + '/alpasq/api/search',
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { 'User-Agent': UA, 'Accept': 'application/json' },
    payload: JSON.stringify({
      searchKeyword: keyword, page: 1, display: display || 20,
      neisCode: [lc.neis], provCode: lc.prov, schoolName: lc.school, coverYn: 'N', facet: 'Y'
    })
  };
}

function libFetch_(req) {
  var opt = {};
  Object.keys(req).forEach(function (k) { if (k !== 'url') opt[k] = req[k]; });
  return UrlFetchApp.fetch(req.url, opt);
}

function libParseList_(res) {
  if (res.getResponseCode() !== 200) return [];
  try {
    var j = JSON.parse(res.getContentText('UTF-8'));
    return (j && j.data && j.data.bookList) || [];
  } catch (e) { return []; }
}

function libSearchUrl_(keyword, lc) {
  return LIB_BASE + '/PureScreen/SchoolSearchResult?searchKeyword=' + encodeURIComponent(keyword) +
    '&provCode=' + lc.prov + '&neisCode=' + lc.neis + '&schoolName=' + encodeURIComponent(lc.school);
}

/** 제목 비교용: 부제(: 뒤) 제거, 공백·기호 제거 */
function normTitle_(t) {
  return String(t || '').split(/[:：]/)[0].replace(/\([^)]*\)|\[[^\]]*\]/g, '').replace(/[\s\-\(\)\[\]<>《》〈〉'"“”‘’!?.,·~]/g, '').toLowerCase();
}

/**
 * 한 권의 소장 여부 + 권별 대출 상태
 * 1) ISBN 으로 찾기 → 2) 없으면 같은 제목(판이 다른 책) 찾기
 * 반환: { match: 'isbn'|'title'|'none', holdings: [...], url, school }
 */
function checkLibrary(isbn, title) {
  var lc = libConf_();
  if (!lc.neis || !lc.prov) return { match: 'off', holdings: [] };
  isbn = String(isbn || '').replace(/[^\dX]/gi, '');
  var key = 'lib:' + isbn + ':' + normTitle_(title).slice(0, 40);
  var cached = cacheGet_(key);
  if (cached) return cached;

  var match = 'none', list = [];
  if (isbn) {
    list = libParseList_(libFetch_(libSearchRequest_(isbn, lc, 30)))
      .filter(function (b) { return String(b.isbn || '').replace(/[^\dX]/gi, '') === isbn; });
    if (list.length) match = 'isbn';
  }
  if (!list.length && title) {
    var nt = normTitle_(title);
    var q = String(title).split(/[:：]/)[0].trim();
    list = libParseList_(libFetch_(libSearchRequest_(q, lc, 30)))
      .filter(function (b) { return nt && normTitle_(b.title) === nt; });
    if (list.length) match = 'title';
  }
  list = list.slice(0, 10);

  // 권별 대출 상태 (한 번에 병렬 요청)
  var states = list.length ? UrlFetchApp.fetchAll(list.map(function (b) {
    return {
      url: LIB_BASE + '/alpasq/api/search/book/state?bookKey=' + encodeURIComponent(b.bookKey) +
        '&provCode=' + lc.prov + '&neisCode=' + lc.neis,
      muteHttpExceptions: true, headers: { 'User-Agent': UA, 'Accept': 'application/json' }
    };
  })) : [];

  var holdings = list.map(function (b, i) {
    var st = {};
    try { st = JSON.parse(states[i].getContentText('UTF-8')).data || {}; } catch (e) { }
    return {
      title: b.title, author: b.author, publisher: b.publisher, pubYear: b.pubYear, isbn: b.isbn,
      callNo: b.callNo || '', regNo: b.regNo || '',
      location: st.locationName || b.locationName || '',
      status: st.status || '확인불가',
      returnDate: st.returnPlanDate || '',
      cover: st.coverUrl || ''
    };
  });

  var result = { match: match, holdings: holdings, school: lc.school,
    url: libSearchUrl_(match === 'title' ? String(title).split(/[:：]/)[0].trim() : (isbn || title), lc) };
  cachePut_(key, result, 600); // 대출 상태는 자주 바뀌므로 10분만 캐시
  return result;
}

/** 검색 결과 목록용: ISBN 들을 한 번에 조회해서 소장 권수만 돌려줌 { isbn: 권수 } */
function checkLibraryBatch(isbns) {
  var lc = libConf_();
  if (!lc.neis || !lc.prov || !isbns || !isbns.length) return {};
  isbns = isbns.map(function (x) { return String(x || '').replace(/[^\dX]/gi, ''); }).filter(String).slice(0, 20);
  var out = {}, need = [];
  isbns.forEach(function (i) {
    var c = cacheGet_('libn:' + i);
    if (c !== null) out[i] = c; else need.push(i);
  });
  if (need.length) {
    var res = UrlFetchApp.fetchAll(need.map(function (i) { return libSearchRequest_(i, lc, 30); }));
    need.forEach(function (i, k) {
      var n = libParseList_(res[k]).filter(function (b) { return String(b.isbn || '').replace(/[^\dX]/gi, '') === i; }).length;
      out[i] = n;
      cachePut_('libn:' + i, n, 1800);
    });
  }
  return out;
}

/** 시트 기록용 한 줄 요약 */
function libSummaryText_(r) {
  if (!r || r.match === 'off') return '';
  if (!r.holdings.length) return '미소장';
  var ok = r.holdings.filter(function (h) { return h.status === '대출가능'; }).length;
  return (r.match === 'title' ? '다른판 ' : '소장 ') + r.holdings.length + '권 (대출가능 ' + ok + ')';
}

/* ===================== 시트 데이터 ===================== */
/** 시트에 쓸 글자: = + - @ 로 시작하면 수식으로 실행되지 않도록 앞에 ' 를 붙임 */
function safeCell_(v) {
  if (typeof v !== 'string') return v;
  v = v.replace(/[\u0000-\u001f]/g, ' ').slice(0, 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function ordersSheet_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS);
  if (!sh) { setup(); sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ORDERS); }
  var last = ORDER_HEADERS.length;
  if (sh.getRange(1, last).getValue() !== ORDER_HEADERS[last - 1]) {   // 예전 버전 시트에 새 열 제목 추가
    sh.getRange(1, last).setValue(ORDER_HEADERS[last - 1]).setFontWeight('bold').setBackground('#e8f0fe');
  }
  return sh;
}

function readOrders_() {
  var sh = ordersSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, ORDER_HEADERS.length).getValues().map(function (r, i) {
    return {
      row: i + 2,
      time: r[COL['신청일시']] instanceof Date ? Utilities.formatDate(r[COL['신청일시']], 'Asia/Seoul', 'MM-dd HH:mm') : String(r[COL['신청일시']]),
      name: String(r[COL['신청자']]).trim(),
      qty: Number(r[COL['수량']]) || 0,
      title: String(r[COL['도서명']]),
      publisher: String(r[COL['출판사']]),
      author: String(r[COL['지은이']]),
      salePrice: Number(r[COL['할인가']]) || 0,
      originalPrice: Number(r[COL['정가']]) || 0,
      isbn: String(r[COL['ISBN']]).replace(/^'/, ''),
      pid: String(r[COL['상품코드']]),
      status: String(r[COL['상태']] || '신청'),
      id: String(r[COL['신청ID']]),
      lib: String(r[COL['학교소장']] || '')
    };
  });
}

/** 이름 정리: 공백 제거 ('김 교사' → '김교사') */
function normName_(n) {
  return String(n || '').replace(/\s+/g, '').slice(0, 20);
}

/** 마감일시 값 → Date (없으면 null, 형식 오류면 Invalid Date) */
function parseDue_(due) {
  if (due === '' || due === null || due === undefined) return null;
  if (due instanceof Date) return due;
  var s = String(due).trim().replace(/[./]/g, '-').replace(' ', 'T');
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) s += 'T23:59:59';
  s = s.replace(/^(\d{4})-(\d{1,2})-(\d{1,2})/, function (m, y, mo, d) { return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2); });
  return new Date(s + '+09:00');
}

function isOpen_(conf) {
  if (String(conf['신청받기']).toUpperCase() !== 'Y') return { open: false, msg: '지금은 신청 기간이 아닙니다.' };
  var d = parseDue_(conf['마감일시']);
  if (d && !isNaN(d) && new Date() > d) return { open: false, msg: '신청이 마감되었습니다.' };
  return { open: true };
}

/** 첫 화면에 필요한 정보 */
function getInitData() {
  var conf = getConf_();
  var st = isOpen_(conf);
  var due = conf['마감일시'];
  return {
    version: VERSION,
    author: AUTHOR,
    authorUrl: AUTHOR_URL,
    title: conf['제목'],
    notice: conf['안내문구'],
    open: st.open, closedMsg: st.msg || '',
    due: due instanceof Date ? Utilities.formatDate(due, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(due || '')
  };
}

/** 내 신청 현황 */
function getMyOrders(name) {
  name = normName_(name);
  var mine = readOrders_().filter(function (o) { return normName_(o.name) === name && o.status !== '취소'; });
  var total = mine.reduce(function (s, o) { return s + o.qty; }, 0);
  return { name: name, total: total, orders: mine.reverse() };
}

/** 신청 저장 — 가격은 서버에서 다시 교보에서 읽어 확정 (화면 값 조작 방지) */
function submitOrder(name, pid, qty) {
  var v = validateOrder_(name, qty);
  var book = getBook(pid);
  return writeOrder_(v.name, v.qty, book, v.conf);
}

/** 교보에서 정보를 못 읽을 때: 직접 입력한 책으로 신청 */
function submitManualOrder(name, info, qty) {
  var v = validateOrder_(name, qty);
  info = info || {};
  var title = String(info.title || '').trim();
  var price = Math.floor(Number(String(info.price || '').replace(/[^\d]/g, '')));
  if (!title) throw new Error('도서명을 입력하세요.');
  if (!(price > 0)) throw new Error('정가를 숫자로 입력하세요.');
  var isbn = String(info.isbn || '').replace(/[^\dX]/gi, '');
  var link = String(info.url || '').trim();
  var book = {
    pid: '직접입력:' + (isbn || normTitle_(title).slice(0, 30)),
    title: title.slice(0, 200),
    author: String(info.author || '').trim().slice(0, 100),
    publisher: String(info.publisher || '').trim().slice(0, 100),
    salePrice: price, originalPrice: price,
    isbn: isbn, pubDate: '',
    url: /^https?:\/\//.test(link) ? link.slice(0, 300) : ''
  };
  return writeOrder_(v.name, v.qty, book, v.conf);
}

function validateOrder_(name, qty) {
  name = normName_(name);
  qty = Math.floor(Number(qty));
  if (!name) throw new Error('이름을 입력하세요.');
  if (!(qty >= 1)) throw new Error('신청 권수를 1 이상으로 입력하세요.');
  if (qty > 100) throw new Error('신청 권수가 너무 많습니다. 담당자에게 문의하세요.');
  var conf = getConf_();
  var st = isOpen_(conf);
  if (!st.open) throw new Error(st.msg);
  return { name: name, qty: qty, conf: conf };
}

function writeOrder_(name, qty, book, conf) {
  var libText = '';
  try { libText = libSummaryText_(checkLibrary(book.isbn, book.title)); } catch (e) { libText = '확인실패'; }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var mine = readOrders_().filter(function (o) { return normName_(o.name) === name && o.status !== '취소'; });
    if (mine.some(function (o) { return o.pid === book.pid; })) {
      throw new Error('이미 신청한 책입니다: ' + book.title + ' (권수를 바꾸려면 내 신청내역에서 취소 후 다시 신청하세요)');
    }
    var basis = String(conf['금액기준']) === '할인가' ? book.salePrice : book.originalPrice;
    var id = Utilities.getUuid().slice(0, 8);
    ordersSheet_().appendRow([
      new Date(), safeCell_(name), qty, safeCell_(book.title), safeCell_(book.publisher), safeCell_(book.author),
      book.salePrice, book.originalPrice, basis * qty,
      book.isbn ? "'" + book.isbn : '', safeCell_(book.pid), safeCell_(book.pubDate), safeCell_(book.url), '신청', id, libText
    ]);
  } finally {
    lock.releaseLock();
  }
  return getMyOrders(name);
}

/** 신청 취소 (행은 남기고 상태만 '취소') */
function cancelOrder(name, id) {
  var conf = getConf_();
  if (!isOpen_(conf).open) throw new Error('신청 기간이 끝나 취소할 수 없습니다. 담당자에게 문의하세요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var o = readOrders_().filter(function (x) { return x.id === String(id) && normName_(x.name) === normName_(name); })[0];
    if (!o) throw new Error('신청 내역을 찾을 수 없습니다.');
    ordersSheet_().getRange(o.row, COL['상태'] + 1).setValue('취소');
  } finally {
    lock.releaseLock();
  }
  return getMyOrders(name);
}

/* ===================== 관리자 ===================== */
function summarize_() {
  var conf = getConf_();
  var basisSale = String(conf['금액기준']) === '할인가';
  var orders = readOrders_().filter(function (o) { return o.status !== '취소'; });
  var byBook = {}, byPerson = {};
  orders.forEach(function (o) {
    var k = o.pid || o.title;
    if (!byBook[k]) byBook[k] = { title: o.title, publisher: o.publisher, author: o.author, isbn: o.isbn,
      originalPrice: o.originalPrice, salePrice: o.salePrice, qty: 0, people: [], lib: o.lib };
    byBook[k].qty += o.qty;
    byBook[k].people.push(o.name + (o.qty > 1 ? '(' + o.qty + ')' : ''));
    byPerson[normName_(o.name)] = (byPerson[normName_(o.name)] || 0) + o.qty;
  });
  var books = Object.keys(byBook).map(function (k) {
    var b = byBook[k];
    b.unit = basisSale ? b.salePrice : b.originalPrice;
    b.amount = b.unit * b.qty;
    return b;
  }).sort(function (a, b) { return b.qty - a.qty || a.title.localeCompare(b.title); });
  return {
    basis: basisSale ? '할인가' : '정가',
    budget: Number(conf['총예산']) || 0,
    books: books,
    totalQty: books.reduce(function (s, b) { return s + b.qty; }, 0),
    totalAmount: books.reduce(function (s, b) { return s + b.amount; }, 0),
    people: Object.keys(byPerson).length
  };
}

function checkAdminPw_(pw) {
  var real = String(getConf_()['관리자비밀번호'] || '').trim();
  if (!real || real === DEFAULT_PW) throw new Error('관리자 비밀번호가 초기값입니다. 시트의 [웹앱설정]에서 비밀번호를 바꾼 뒤 사용하세요.');
  if (String(pw) !== real) throw new Error('비밀번호가 맞지 않습니다.');
}

function getAdminSummary(pw) {
  checkAdminPw_(pw);
  var s = summarize_();
  s.update = checkUpdate_();
  s.version = VERSION;
  return s;
}

function buildSummaryFromWeb(pw) {
  checkAdminPw_(pw);
  buildSummarySheet_();
  return SpreadsheetApp.getActive().getUrl() + '#gid=' + SpreadsheetApp.getActive().getSheetByName(SHEET_SUMMARY).getSheetId();
}

/**
 * 메뉴 전용: [주문서 시트 만들기]
 * 이름 끝에 _ 가 없어 google.script.run 으로도 부를 수 있으므로, 시트 화면에서만 동작하게 막는다.
 * 1) 웹앱에서 부르면 SpreadsheetApp.getUi() 가 예외를 던져 여기서 멈춘다.
 * 2) 그래도 UI 가 잡히는 경우를 대비해, 시트 앞의 사람이 [예]를 눌러야만 진행한다.
 */
function buildSummarySheet() {
  var ui = SpreadsheetApp.getUi();
  var ans = ui.alert('주문서 시트 만들기',
    '[' + SHEET_SUMMARY + '] 시트의 기존 내용을 지우고 새로 만듭니다. 계속할까요?', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  buildSummarySheet_();
}

/** 주문서 시트를 새로 만듭니다 (도서별 합계). 기존 주문서 내용은 지워짐 — 관리자 확인을 거친 진입점에서만 부를 것 */
function buildSummarySheet_() {
  var s = summarize_();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_SUMMARY) || ss.insertSheet(SHEET_SUMMARY);
  sh.clear();
  var head = ['순번', '도서명', '출판사', '지은이', 'ISBN', '정가', '할인가', '수량', '금액(' + s.basis + ')', '신청자', '학교소장'];
  var rows = s.books.map(function (b, i) {
    return [i + 1, safeCell_(b.title), safeCell_(b.publisher), safeCell_(b.author), b.isbn ? "'" + b.isbn : '', b.originalPrice, b.salePrice, b.qty, b.amount, safeCell_(b.people.join(', ')), b.lib || ''];
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8f0fe');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    var tr = rows.length + 2;
    sh.getRange(tr, 1, 1, head.length).setValues([['합계', '', '', '', '', '', '', s.totalQty, s.totalAmount, s.people + '명 신청', '']])
      .setFontWeight('bold').setBackground('#fef7e0');
    sh.getRange(2, 6, rows.length + 1, 4).setNumberFormat('#,##0');
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, head.length);
  sh.setColumnWidth(2, 320);
  try { ss.toast('주문서 시트를 만들었습니다 (' + rows.length + '종 ' + s.totalQty + '권)', '도서신청', 5); } catch (e) { }
}
