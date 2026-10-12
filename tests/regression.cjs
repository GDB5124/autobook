const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const root = path.resolve(__dirname, '..');
const code = fs.readFileSync(path.join(root, 'Code.gs'), 'utf8');
let passed = 0;
function context(extra = {}) {
  return vm.createContext({
    Utilities: {
      DigestAlgorithm: {SHA_256: 'sha256'}, Charset: {UTF_8: 'utf8'},
      computeDigest: (algorithm, value) => crypto.createHash(algorithm).update(value).digest(),
      base64EncodeWebSafe: value => Buffer.from(value).toString('base64url'),
      getUuid: () => '12345678-1234-1234-1234-123456789012'
    }, ...extra
  });
}
function load(extra) { const c = context(extra); vm.runInContext(code, c); return c; }
function test(name, fn) { fn(); passed++; console.log('PASS ' + name); }
function properties(initial = {}) {
  const data = {...initial};
  return {getProperty: key => data[key], setProperty: (key, value) => {data[key] = value;}};
}
function sheetWithConf(rows) {
  return {getLastRow: () => rows.length + 1, getRange: () => ({getValues: () => rows})};
}
test('web app reads saved sheet ID with no active spreadsheet', () => {
  const props = properties({SPREADSHEET_ID: 'saved-sheet'});
  let opened;
  const c = load({PropertiesService: {getScriptProperties: () => props}, SpreadsheetApp: {
    getActive: () => null,
    openById: id => {opened = id; return {getSheetByName: () => sheetWithConf([['제목', '우리학교']])};}
  }});
  assert.equal(c.getInitData().title, '우리학교'); assert.equal(opened, 'saved-sheet');
});
test('menu context remembers the current copy instead of an old sheet ID', () => {
  const props = properties({SPREADSHEET_ID: 'old-sheet'});
  const active = {getId: () => 'new-copy'};
  const c = load({PropertiesService: {getScriptProperties: () => props}, SpreadsheetApp: {getActive: () => active}});
  assert.equal(c.getSpreadsheet_(), active); assert.equal(props.getProperty('SPREADSHEET_ID'), 'new-copy');
});
test('missing installation gives a setup instruction', () => {
  const c = load({PropertiesService: {getScriptProperties: () => properties()}, SpreadsheetApp: {getActive: () => null}});
  assert.throws(() => c.getInitData(), /처음 세팅/);
});
test('summary generator is private and web entry rejects invalid passwords before writes', () => {
  const c = load(); c.getConf_ = () => ({'관리자비밀번호': 'secret'});
  let builds = 0; c.buildSummarySheet_ = () => {builds++;};
  c.getSpreadsheet_ = () => ({getUrl: () => 'https://sheet', getSheetByName: () => ({getSheetId: () => 7})});
  assert.equal(typeof c.buildSummarySheet, 'undefined');
  assert.throws(() => c.buildSummaryFromWeb('wrong'), /비밀번호/); assert.equal(builds, 0);
  assert.equal(c.buildSummaryFromWeb('secret'), 'https://sheet#gid=7'); assert.equal(builds, 1);
  assert.match(code, /addItem\('주문서 시트 만들기', 'buildSummarySheet_'\)/);
});
test('manual titles retain subtitles, authors and publishers in their identity', () => {
  const c = load(); c.getConf_ = () => ({'신청받기': 'Y'});
  const saved = []; c.writeOrder_ = (name, qty, book) => {saved.push(book); return book;};
  c.submitManualOrder('교사', {title: '수학: 상', price: '10000'}, 1);
  c.submitManualOrder('교사', {title: '수학: 하', price: '10000'}, 1);
  c.submitManualOrder('교사', {title: '수학: 상', publisher: '다른출판사', price: '10000'}, 1);
  assert.equal(new Set(saved.map(b => b.pid)).size, 3);
  assert.equal(c.bookKey_({pid: '직접입력:수학', title: '수학: 상'}), c.bookKey_(saved[0]));
  assert.notEqual(c.bookKey_({pid: '직접입력:수학', title: '수학: 하'}), c.bookKey_(saved[0]));
});
test('same ISBN has stable manual identity despite title variations', () => {
  const c = load();
  assert.equal(c.bookKey_({pid: '직접입력:old', isbn: '978-1234567890', title: '상'}),
    c.bookKey_({pid: '직접입력:new', isbn: '9781234567890', title: '다른 제목'}));
});
function order(extra = {}) {
  return {pid: 'S123456789012', title: '책', name: '교사', qty: 1, originalPrice: 10000,
    salePrice: 9000, status: '신청', ...extra};
}
test('summary sums each order price and excludes cancellations for both price bases', () => {
  const c = load(); c.readOrders_ = () => [order(), order({name: '다른교사', qty: 2, originalPrice: 20000, salePrice: 15000}), order({status: '취소'})];
  for (const [basis, expected] of [['정가', 50000], ['할인가', 39000]]) {
    c.getConf_ = () => ({'금액기준': basis}); const s = c.summarize_();
    assert.equal(s.totalQty, 3); assert.equal(s.people, 2); assert.equal(s.totalAmount, expected);
    assert.equal(s.books.length, 1); assert.equal(s.books[0].mixedPrice, true);
  }
});
test('legacy colliding manual IDs produce distinct summary rows', () => {
  const c = load(); c.getConf_ = () => ({'금액기준': '정가'});
  c.readOrders_ = () => [order({pid: '직접입력:수학', title: '수학: 상'}), order({pid: '직접입력:수학', title: '수학: 하'})];
  assert.equal(c.summarize_().books.length, 2); assert.equal(c.summarize_().totalAmount, 20000);
});
function response(status, body) {return {getResponseCode: () => status, getContentText: () => body};}
test('library accepts real empty results but rejects transport and schema failures', () => {
  const c = load();
  assert.equal(c.libParseList_(response(200, '{"data":{"bookList":[]}}')).length, 0);
  for (const r of [response(503, '{}'), response(200, 'not json'), response(200, '{"data":{}}'), response(200, '{"data":{"bookList":{}}}')]) {
    assert.throws(() => c.libParseList_(r), /학교도서관/);
  }
});
test('library failure is not cached as absence and saved orders record 확인실패', () => {
  const c = load(); let cached = 0; let row;
  c.libConf_ = () => ({neis: 'school', prov: 'region'}); c.cacheGet_ = () => null;
  c.cachePut_ = () => {cached++;}; c.libFetch_ = () => response(503, '{}');
  assert.throws(() => c.checkLibrary('9781234567890', '책'), /HTTP 503/); assert.equal(cached, 0);
  c.LockService = {getScriptLock: () => ({waitLock() {}, releaseLock() {}})};
  c.readOrders_ = () => []; c.ordersSheet_ = () => ({appendRow: value => {row = value;}}); c.getMyOrders = () => ({});
  c.writeOrder_('교사', 1, order(), {'금액기준': '정가'});
  assert.equal(row[15], '확인실패'); assert.equal(c.checkLibraryConnection_().ok, false);
});
test('legacy same-book requests remain blocked but different subtitles can be saved', () => {
  const c = load(); let writes = 0;
  c.checkLibrary = () => ({match: 'off'}); c.LockService = {getScriptLock: () => ({waitLock() {}, releaseLock() {}})};
  c.readOrders_ = () => [order({name: '교사', pid: '직접입력:수학', title: '수학: 상'})];
  c.ordersSheet_ = () => ({appendRow: () => {writes++;}}); c.getMyOrders = () => ({});
  assert.throws(() => c.writeOrder_('교사', 1, order({pid: '직접입력:new', title: '수학: 상'}), {}), /이미 신청/);
  c.writeOrder_('교사', 1, order({pid: '직접입력:new', title: '수학: 하'}), {}); assert.equal(writes, 1);
});
test('exported summary preserves mixed price labels and the correct total', () => {
  const c = load(); const values = [];
  c.getConf_ = () => ({'금액기준': '정가'});
  c.readOrders_ = () => [order(), order({originalPrice: 20000})];
  const range = {setValues(rows) {values.push(rows); return this;}, setFontWeight() {return this;},
    setBackground() {return this;}, setNumberFormat() {return this;}, setWrap() {return this;}, setVerticalAlignment() {return this;}};
  const sheet = {clear() {}, getRange: () => range, setFrozenRows() {}, autoResizeColumns() {}, setColumnWidth() {},
    getColumnWidth: () => 100, autoResizeRows() {}};
  c.SpreadsheetApp = {flush() {}};
  c.getSpreadsheet_ = () => ({getSheetByName: () => sheet, toast() {}});
  c.buildSummarySheet_();
  assert.equal(values[1][0][5], '신청별 상이'); assert.equal(values[1][0][8], 30000);
  assert.equal(values[2][0][8], 30000);
});
test('Kyobo requests include a search referrer and reject HTTP 200 empty responses', () => {
  let options; let body = '<html>book</html>';
  const c = load({UrlFetchApp: {fetch: (url, opt) => {options = opt; return response(200, body);}}});
  assert.equal(c.fetchHtml_('https://product.kyobobook.co.kr/detail/S000216232523'), body);
  assert.equal(options.headers.Referer, 'https://search.kyobobook.co.kr/');
  assert.match(options.headers.Accept, /text\/html/);
  body = ' '; assert.throws(() => c.fetchHtml_('https://product.kyobobook.co.kr/detail/S000216232523'), /빈 응답/);
});
test('manifest, release metadata and both scripts parse', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
  assert.ok(manifest.oauthScopes.includes('https://www.googleapis.com/auth/spreadsheets'));
  const release = JSON.parse(fs.readFileSync(path.join(root, 'version.json'), 'utf8'));
  assert.equal(load().VERSION, release.latest);
  new vm.Script(fs.readFileSync(path.join(root, 'index.html'), 'utf8').match(/<script>([\s\S]*?)<\/script>/)[1]);
});
console.log(`${passed} regression checks passed`);
