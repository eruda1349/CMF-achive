/**
 * CMF 채번 시스템 — 웹앱 버전
 * 아카이브 + CMF 등록 + 샘플실 + 작업지시서 소재 추출
 *
 * v18 변경
 *  · 화면 밝기(Light/Dark)를 사용자 속성에 저장 — 앱스스크립트 화면은 열 때마다
 *    googleusercontent 하위 도메인이 바뀌어 localStorage 가 남지 않는다.
 *  · 런처(Design Applicators)가 ?home= 으로 넘겨 준 주소를 받아
 *    화면 왼쪽 위에 '← Applicators' 링크를 띄운다.
 *  · GUI 를 디자인 시뮬레이터·회의록과 같은 규칙으로 통일 (Archive.html)
 *
 * v17 변경
 *  · 채번 규칙: 코드 앞자리를 연도 2자리만 사용 (26-FB-SG-EM-MA-01)
 *  · 작업지시서(PDF/PPTX/이미지) → 브라우저에서 텍스트 추출 → CMF 후보 리스트 → 일괄 등록
 *  · submitEntry 내부 로직 분리(submitEntryCore_)로 일괄 등록 시 잠금 1회만 사용
 *
 * v19 변경 (hida 스타일 리디자인)
 *  · Archive.html 을 hida(hidalab.com)의 갤러리형 CMF 라이브러리 UX를 참고해 리터치
 *  · 기본 화면 밝기를 Light 로 변경 (밝고 미니멀한 갤러리 톤에 맞춤) — 사용자가 저장한
 *    선호값이 있으면 그 값이 항상 우선한다
 */

const APP_VERSION = 'v19';

const CONFIG = {
  SHEET_MASTER: '②등록대장',
  SHEET_COLOR: '③컬러사전',
  SHEET_CLASS: '④분류사전',
  SHEET_APPLY: '⑤적용이력',
  SHEET_COUNTER: '⑥채번카운터',
  ADMIN_EMAIL: 'eevee6@coway.com',
  PHOTO_FOLDER_NAME: 'CMF 아카이브 사진'
};

const GLOSSES = [
  { code: 'MA', name: '무광', legacy: 'M' },
  { code: 'SG', name: '반광', legacy: 'S' },
  { code: 'GL', name: '유광', legacy: 'G' }
];

const BOX_ORDER = [
  'PU',        // A
  '패브릭',    // B
  '레깅스',    // C
  '가죽',      // D
  '스웨이드',  // E
  '사출',      // F
  '금속',      // G
  '스프레이',  // H
  '실리콘',    // I
  '우드'       // J
];

const BOX_ALIAS = {
  '페브릭': '패브릭',
  'PU레더': 'PU',
  'PU 레더': 'PU',
  'PU(레더)': 'PU',
  'PU (레더)': 'PU',
  '레더': 'PU',
  '피유': 'PU',
  '사출(플라스틱)': '사출',
  '사출 (플라스틱)': '사출',
  '플라스틱': '사출',
  '메탈': '금속',
  '우드(목재)': '우드',
  '목재': '우드'
};

const COL_BOX = 26;
const COL_PRODUCT = 20;
const COL_MANUFACTURER = 27;
const COL_PRICE = 28;
const COL_UPDATED_AT = 29;
const COL_UPDATED_BY = 30;
const COL_LEGACY_CODE = 31;
const COL_DEPT = 32;

/* 현재 형식: 연도 2자리 + 4구간 + 순번 */
const CODE_RE_YEAR = /^(\d{2})-([A-Z0-9]{2})-([A-Z0-9]{2})-([A-Z0-9]{2})-([A-Z0-9]{2})-(\d+)$/;
/* 이전 형식: 연월일 6자리 */
const CODE_RE_FULLDATE = /^(\d{6})-([A-Z0-9]{2})-([A-Z0-9]{2})-([A-Z0-9]{2})-([A-Z0-9]{2})-(\d+)$/;
const NEW_CODE_RE = CODE_RE_YEAR;

/* ────────── 웹앱 진입점 ────────── */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Archive');

  const raw = (e && e.parameter && e.parameter.code) ? String(e.parameter.code) : '';
  t.initialCode = /^[A-Za-z0-9-]+$/.test(raw) ? raw : '';

  // 런처가 ?home= 으로 자기 주소를 넘겨 준다. 주소창에 아무 주소나 적어 넣어도
  // 엉뚱한 곳으로 튀지 않도록, 구글 스크립트 주소만 받아들인다.
  const home = (e && e.parameter && e.parameter.home) ? String(e.parameter.home) : '';
  t.homeUrl = /^https:\/\/script\.google\.com\/[A-Za-z0-9\/_\-.?=&]*$/.test(home) ? home : '';

  t.theme = getTheme();

  return t.evaluate()
    .setTitle('BEREX CMF LAB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ────────── 화면 밝기 ──────────
 * 계정별로 갈리므로 사람마다 자기 설정만 본다.
 * (런처·회의록은 다른 프로젝트라 설정이 따로 저장된다)
 * v19: 저장된 값이 없는 첫 방문자는 Light 로 시작 (hida 스타일 갤러리 톤에 맞춤) */
function getTheme() {
  const v = PropertiesService.getUserProperties().getProperty('THEME');
  return v === 'dark' ? 'dark' : 'light';
}
function setTheme(theme) {
  PropertiesService.getUserProperties()
    .setProperty('THEME', theme === 'dark' ? 'dark' : 'light');
  return theme;
}

let _classCache = null;
function classRows_(kind) {
  if (!_classCache) {
    _classCache = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_CLASS).getDataRange().getValues();
  }
  return _classCache.filter(r => r[0] === kind);
}
function clearClassCache_() { _classCache = null; }

function boxName(i) {
  let s = '';
  do { s = String.fromCharCode(65 + i % 26) + s; i = Math.floor(i / 26) - 1; } while (i >= 0);
  return s;
}

function canonMatName_(name) {
  const n = String(name || '').trim();
  return BOX_ALIAS[n] || n;
}

function extraMatNames_() {
  const names = classRows_('소재').map(r => canonMatName_(r[2]));
  return [...new Set(names)]
    .filter(n => n && BOX_ORDER.indexOf(n) < 0)
    .sort((a, b) => a.localeCompare(b, 'ko'));
}

function assignBox(matCode) {
  const row = classRows_('소재').find(r => r[1] === matCode);
  const name = canonMatName_(row ? row[2] : '');
  const i = BOX_ORDER.indexOf(name);
  if (i >= 0) return boxName(i);
  const extras = extraMatNames_();
  const p = extras.indexOf(name);
  return boxName(BOX_ORDER.length + (p >= 0 ? p : extras.length));
}

function boxLayout_() {
  const list = BOX_ORDER.map((name, i) => ({ box: boxName(i), name: name }));
  extraMatNames_().forEach((name, i) => {
    list.push({ box: boxName(BOX_ORDER.length + i), name: name });
  });
  return list;
}

function getMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mats = classRows_('소재');
  const fins = classRows_('마감');
  let appUrl = '';
  try { appUrl = ScriptApp.getService().getUrl() || ''; } catch (err) {}

  const gloss = [];
  const seen = new Set();
  GLOSSES.forEach(g => {
    gloss.push({ code: g.code, name: g.name });
    seen.add(g.code); seen.add(g.legacy); seen.add(normName(g.name));
  });
  classRows_('광택').forEach(r => {
    const c = String(r[1] || '').toUpperCase(), n = normName(r[2] || '');
    if (!c || !n || seen.has(c) || seen.has(n)) return;
    gloss.push({ code: r[1], name: r[2] });
    seen.add(c); seen.add(n);
  });

  // 소재는 정식 이름으로 병합 (PU(레더)/PU 처럼 중복된 행을 하나로)
  const byName = {};
  mats.forEach(r => {
    const name = canonMatName_(r[2]);
    if (!name || !r[1]) return;
    const exact = String(r[2]).trim() === name;
    if (!byName[name] || (exact && !byName[name].exact)) byName[name] = { code: r[1], exact: exact };
  });
  const materials = Object.keys(byName)
    .map(n => ({ code: byName[n].code, name: n, box: assignBox(byName[n].code) }))
    .sort((a, b) => a.box.length - b.box.length || a.box.localeCompare(b.box));

  return {
    version: APP_VERSION,
    materials: materials,
    finishes: fins.map(r => ({ code: r[1], name: r[2] })),
    glosses: gloss,
    boxes: boxLayout_(),
    sheetUrl: ss.getUrl(),
    appUrl: appUrl,
    user: Session.getActiveUser().getEmail() || ''
  };
}

function getProductList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const set = new Set();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER).getDataRange().getValues();
  for (let i = 1; i < master.length; i++) {
    const p = master[i][COL_PRODUCT - 1];
    if (p && p !== '미정') set.add(p);
  }
  const apply = getOrCreateSheet(ss, CONFIG.SHEET_APPLY, ['CMF코드','제품','','','','등록경로','일시','비고']);
  const rows = apply.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (rows[i][1]) set.add(rows[i][1]);
  return [...set].sort();
}

function getArchiveData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER).getDataRange().getValues();

  const applySheet = getOrCreateSheet(ss, CONFIG.SHEET_APPLY, ['CMF코드','제품','','','','등록경로','일시','비고']);
  const applyRows = applySheet.getDataRange().getValues();
  const applyMap = {};
  for (let i = 1; i < applyRows.length; i++) {
    const code = applyRows[i][0];
    if (!code) continue;
    if (!applyMap[code]) applyMap[code] = [];
    applyMap[code].push({
      product: applyRows[i][1] || '',
      date: applyRows[i][6] ? Utilities.formatDate(new Date(applyRows[i][6]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      note: applyRows[i][7] || ''
    });
  }

  const rows = [];
  for (let i = 1; i < master.length; i++) {
    const r = master[i];
    if (!r[0]) continue;
    const products = applyMap[r[0]] || [];
    const prodDisplay = products.length
      ? products.map(p => p.product).filter((v, idx, arr) => arr.indexOf(v) === idx).join(', ')
      : (r[COL_PRODUCT - 1] || '미정');
    rows.push({
      code: r[0], status: r[1], matCode: r[2], mat: canonMatName_(r[3]), abbr: r[4], color: r[5],
      finish: r[7], gloss: r[9], ser: String(r[10]), hex: r[11], pantone: r[12],
      sup: r[14], supcode: r[15], prod: prodDisplay, products: products,
      photo: r[21], note: r[22],
      by: r[23], at: r[24] ? Utilities.formatDate(new Date(r[24]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      box: r[25] || '',
      price: r[COL_PRICE - 1] || '',
      updatedAt: r[COL_UPDATED_AT - 1] ? Utilities.formatDate(new Date(r[COL_UPDATED_AT - 1]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      updatedBy: r[COL_UPDATED_BY - 1] || '',
      legacy: r[COL_LEGACY_CODE - 1] || '',
      dept: r[COL_DEPT - 1] || ''
    });
  }
  return rows;
}

function checkNewItems(d) {
  const mat = resolveMaterial(d.material, d.materialEtc, d.materialEtcEng, true);
  const finish = resolveFinish(d.finishCode, d.finishEtc, d.finishEtcEng, true);
  const gloss = resolveGloss(d.gloss, d.glossEtc, d.glossEtcEng, true);
  const color = resolveColor(d.colorName, d.colorEng, d.hex, true);
  return {
    matIsNew: mat.isNew, matName: mat.name, matCode: mat.code,
    finIsNew: finish.isNew, finName: finish.name, finCode: finish.code,
    gloIsNew: gloss.isNew, gloName: gloss.name, gloCode: gloss.code,
    colorIsNew: color.isNew, colorName: color.stdName
  };
}

/* 등록 본체 — 잠금은 호출부에서 처리 (일괄 등록 시 잠금 1회로 묶기 위함) */
function submitEntryCore_(ss, d) {
  const mat = resolveMaterial(d.material, d.materialEtc, d.materialEtcEng, false);
  const glo = resolveGloss(d.gloss, d.glossEtc, d.glossEtcEng, false);
  const color = resolveColor(d.colorName, d.colorEng, d.hex, false);
  const finish = resolveFinish(d.finishCode, d.finishEtc, d.finishEtcEng, false);

  const now = new Date();

  // 코드 앞자리는 연도 2자리만 사용한다.
  // 등록 날짜(년·월·일)는 그대로 Y열에 저장되고, 코드에는 연도만 반영된다.
  const dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(d.date || '').trim());
  const yy = dm ? dm[1].slice(2) : Utilities.formatDate(now, 'Asia/Seoul', 'yy');
  const regDate = dm ? new Date(dm[1] + '-' + dm[2] + '-' + dm[3] + 'T12:00:00+09:00') : now;

  const mc = two_(mat.code), cc = two_(color.abbr), fc = two_(finish.code), gc = two_(glo.code);
  const serial = nextSerial(yy, mc, cc, fc, gc);
  const cmfCode = [yy, mc, cc, fc, gc, serial].join('-');

  const status = (d.status === '검토중') ? '검토중' : '확정';
  const by = d.by || Session.getActiveUser().getEmail() || '';
  const dept = d.dept || '';
  const product = d.product || '미정';
  const box = assignBox(mat.code);

  const row = [
    cmfCode, status, mat.code, mat.name, color.abbr, color.stdName,
    finish.code, finish.name, glo.code, glo.name, serial,
    color.hex || d.hex, d.pantone || '', '', d.supplier || '', d.supplierCode || '',
    '', '', '', product, '', d.photo || '', d.note || '', by, regDate, box,
    '', d.price || '', now, by, '', dept
  ];
  ss.getSheetByName(CONFIG.SHEET_MASTER).appendRow(row);

  if (product && product !== '미정') {
    appendApplyRow(ss, cmfCode, product, d.source || '초기 등록', d.note || '');
  }

  if ((color.isNew || finish.isNew || mat.isNew || glo.isNew) && CONFIG.ADMIN_EMAIL) {
    try {
      MailApp.sendEmail(CONFIG.ADMIN_EMAIL, '[CMF] 신규 사전 항목: ' + cmfCode,
        '신규 컬러/소재/마감/광택이 등록되었습니다.\n코드: ' + cmfCode + '\n등록자: ' + by);
    } catch (err) {}
  }

  return {
    code: cmfCode, status: status, serial: serial, box: box, year: yy,
    matName: mat.name, colorName: color.stdName, finishName: finish.name, glossName: glo.name,
    colorIsNew: color.isNew, finishIsNew: finish.isNew, matIsNew: mat.isNew, glossIsNew: glo.isNew,
    abbr: color.abbr, finishCode: finish.code, matCode: mat.code, glossCode: glo.code
  };
}

function submitEntry(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return submitEntryCore_(SpreadsheetApp.getActiveSpreadsheet(), d);
  } finally {
    lock.releaseLock();
  }
}

/* 작업지시서 추출분 일괄 등록 */
function submitBatch(items, common) {
  const lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const c = common || {};
    const done = [], failed = [];
    (items || []).forEach((it, i) => {
      try {
        const d = {
          date: c.date, by: c.by, dept: c.dept,
          status: c.status || '검토중',
          photo: '',
          source: '작업지시서 추출',
          material: it.material, materialEtc: '', materialEtcEng: '',
          finishCode: it.finishCode, finishEtc: '', finishEtcEng: '',
          gloss: it.gloss, glossEtc: '', glossEtcEng: '',
          colorName: it.colorName, colorEng: it.colorEng,
          hex: it.hex, pantone: it.pantone,
          supplier: it.supplier || '', supplierCode: it.supplierCode || '',
          price: it.price || '',
          product: it.product || '미정',
          note: it.note || ''
        };
        done.push(submitEntryCore_(ss, d));
      } catch (e) {
        failed.push('#' + (i + 1) + ' ' + (it.colorName || it.hex || '') + ' — ' + (e.message || e));
      }
    });
    return { ok: done.length, codes: done.map(x => x.code), failed: failed };
  } finally {
    lock.releaseLock();
  }
}

function comboKey_(yy, matCode, colorAbbr, finishCode, glossCode) {
  return [yy, matCode, colorAbbr, finishCode, glossCode].join('|');
}

function nextSerial(yy, matCode, colorAbbr, finishCode, glossCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_COUNTER, ['조합키', '최대발급번호']);
  const data = sheet.getDataRange().getValues();
  const key = comboKey_(yy, matCode, colorAbbr, finishCode, glossCode);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      const next = (parseInt(data[i][1], 10) || 0) + 1;
      sheet.getRange(i + 1, 2).setValue(next);
      return String(next).padStart(2, '0');
    }
  }
  sheet.appendRow([key, 1]);
  return '01';
}

function initCounterFromExisting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER).getDataRange().getValues();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_COUNTER, ['조합키', '최대발급번호']);
  const counterData = sheet.getDataRange().getValues();
  const rowIndexByKey = {};
  for (let i = 1; i < counterData.length; i++) rowIndexByKey[counterData[i][0]] = i + 1;

  const maxByKey = {};
  for (let i = 1; i < master.length; i++) {
    const m = CODE_RE_YEAR.exec(String(master[i][0] || '').toUpperCase());
    if (!m) continue;
    const key = comboKey_(m[1], m[2], m[3], m[4], m[5]);
    const ser = parseInt(m[6], 10) || 0;
    maxByKey[key] = Math.max(maxByKey[key] || 0, ser);
  }

  Object.keys(maxByKey).forEach(key => {
    const val = maxByKey[key];
    if (rowIndexByKey[key]) {
      const rowNum = rowIndexByKey[key];
      const cur = parseInt(sheet.getRange(rowNum, 2).getValue(), 10) || 0;
      if (val > cur) sheet.getRange(rowNum, 2).setValue(val);
    } else {
      sheet.appendRow([key, val]);
    }
  });
  return say_('채번카운터 갱신 — 조합 ' + Object.keys(maxByKey).length + '건');
}

function getPhotoFolder_() {
  const it = DriveApp.getRootFolder().getFoldersByName(CONFIG.PHOTO_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.getRootFolder().createFolder(CONFIG.PHOTO_FOLDER_NAME);
}

function uploadPhoto(base64Data, mimeType, fileName) {
  const folder = getPhotoFolder_();
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName || ('cmf_' + Date.now() + '.jpg'));
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {}
  return {
    url: 'https://lh3.googleusercontent.com/d/' + file.getId() + '=s800',
    fileId: file.getId()
  };
}

function getOrCreateSheet(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (header && header.length) sheet.appendRow(header);
  }
  return sheet;
}

function findRow(sheet, code) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (data[i][0] === code) return i + 1;
  return -1;
}

function updateEntry(code, f) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MASTER);
  const row = findRow(sheet, code);
  if (row < 0) throw new Error('코드를 찾을 수 없습니다: ' + code);

  const hex = String(f.hex || '').trim();
  if (hex && !/^#[0-9a-fA-F]{6}$/.test(hex)) throw new Error('HEX 형식이 올바르지 않습니다 (#RRGGBB): ' + hex);

  sheet.getRange(row, 2).setValue(f.status || '확정');
  sheet.getRange(row, 12).setValue(hex);
  sheet.getRange(row, 13).setValue(f.pantone || '');
  sheet.getRange(row, 15).setValue(f.sup || '');
  sheet.getRange(row, 16).setValue(f.supcode || '');
  sheet.getRange(row, 23).setValue(f.note || '');
  sheet.getRange(row, COL_BOX).setValue(f.box || '');
  if (f.photo) sheet.getRange(row, 22).setValue(f.photo);
  sheet.getRange(row, COL_PRICE).setValue(f.price || '');
  sheet.getRange(row, COL_DEPT).setValue(f.dept || '');

  if (f.date) {
    const dm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(f.date).trim());
    if (dm) {
      sheet.getRange(row, 25).setValue(new Date(dm[1] + '-' + dm[2] + '-' + dm[3] + 'T12:00:00+09:00'));
    }
  }

  sheet.getRange(row, COL_UPDATED_AT).setValue(new Date());
  sheet.getRange(row, COL_UPDATED_BY).setValue(f.by || Session.getActiveUser().getEmail() || '');
  return true;
}

function deleteEntry(code) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const row = findRow(master, code);
  if (row < 0) throw new Error('코드를 찾을 수 없습니다: ' + code);
  master.deleteRow(row);
  const apply = ss.getSheetByName(CONFIG.SHEET_APPLY);
  if (apply) {
    const data = apply.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === code) apply.deleteRow(i + 1);
    }
  }
  return true;
}

function toggleStatusSrv(code) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MASTER);
  const row = findRow(sheet, code);
  if (row < 0) throw new Error('코드를 찾을 수 없습니다: ' + code);
  const next = sheet.getRange(row, 2).getValue() === '확정' ? '검토중' : '확정';
  sheet.getRange(row, 2).setValue(next);
  return next;
}

function appendApplyRow(ss, code, product, source, note) {
  const apply = getOrCreateSheet(ss, CONFIG.SHEET_APPLY, ['CMF코드','제품','','','','등록경로','일시','비고']);
  apply.appendRow([code, product, '', '', '', source || '웹앱 추가', new Date(), note || '']);
}

function addApplication(code, product, note) {
  if (!product || !product.trim()) throw new Error('제품명을 입력해주세요.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
  const row = findRow(master, code);
  if (row < 0) throw new Error('코드를 찾을 수 없습니다: ' + code);
  appendApplyRow(ss, code, product.trim(), '웹앱 추가', note || '');
  const cur = master.getRange(row, COL_PRODUCT).getValue();
  if (!cur || cur === '미정') master.getRange(row, COL_PRODUCT).setValue(product.trim());
  return true;
}

function normName(s) {
  return s.toString().toLowerCase().replace(/\s+/g, '');
}

function two_(s) {
  const t = String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return (t || 'X').padEnd(2, 'X').slice(0, 2);
}

function glossCode2_(v) {
  const s = String(v || '').toUpperCase().trim();
  const g = GLOSSES.find(x => x.code === s || x.legacy === s);
  return g ? g.code : two_(s);
}

function make2_(eng, existing) {
  const clean = String(eng || '').toUpperCase().replace(/[^A-Z\s]/g, '').trim();
  const w = clean ? clean.split(/\s+/).filter(Boolean) : [];
  const a = w[0] || 'X';
  const cands = [];
  if (w.length >= 2) cands.push(a[0] + w[1][0]);
  cands.push(a.padEnd(2, 'X').slice(0, 2));
  if (a.length >= 3) cands.push(a[0] + a[2]);
  if (a.length >= 2) cands.push(a[0] + a[a.length - 1]);
  if (w.length >= 2 && w[1].length >= 2) cands.push(a[0] + w[1][1]);

  for (const c of cands) if (c.length === 2 && existing.indexOf(c) < 0) return c;
  for (let i = 2; i <= 9; i++) { const c = a[0] + i; if (existing.indexOf(c) < 0) return c; }
  const AZ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  for (const y of AZ) { const c = a[0] + y; if (existing.indexOf(c) < 0) return c; }
  for (const x of AZ) for (const y of AZ) { const c = x + y; if (existing.indexOf(c) < 0) return c; }
  return 'ZZ';
}

function resolveColor(rawName, engName, hex, dryRun) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_COLOR);
  const data = sheet.getDataRange().getValues();
  const key = normName(rawName);

  for (let i = 1; i < data.length; i++) {
    const names = [data[i][0], data[i][1]].concat((data[i][3] || '').split(','));
    if (names.some(n => n && normName(n) === key)) {
      return { stdName: data[i][0], abbr: data[i][2], hex: data[i][4] || hex, isNew: false };
    }
  }

  let eng = engName;
  if (!eng && /^[a-zA-Z\s]+$/.test(rawName)) eng = rawName;
  const existing = data.slice(1).map(r => String(r[2] || '').toUpperCase());
  const abbr = make2_(eng || 'X', existing);

  if (!dryRun) {
    sheet.appendRow([rawName, eng ? eng.toUpperCase().trim() : '(영문명 확인 필요)', abbr, '', hex || '', '자동 등록 — 검토 필요']);
  }
  return { stdName: rawName, abbr: abbr, hex: hex, isNew: true };
}

function resolveMaterial(matCode, etcName, etcEng, dryRun) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CLASS);
  const rows = classRows_('소재');

  if (matCode && matCode !== '__etc') {
    const r = rows.find(x => x[1] === matCode);
    if (r) return { code: r[1], name: canonMatName_(r[2]), isNew: false };
    throw new Error('알 수 없는 소재 코드: ' + matCode);
  }
  if (!etcName) throw new Error('기타 소재명을 입력해 주세요.');
  const exist = rows.find(r => normName(r[2]) === normName(etcName));
  if (exist) return { code: exist[1], name: canonMatName_(exist[2]), isNew: false };

  const existing = rows.map(r => String(r[1] || '').toUpperCase());
  const code = make2_(etcEng || 'X', existing);
  if (!dryRun) {
    sheet.appendRow(['소재', code, etcName, etcEng ? etcEng.toUpperCase().trim() : '(영문명 확인 필요)', '-', '기타 응답 자동 등록 — 검토 필요']);
    clearClassCache_();
  }
  return { code: code, name: etcName, isNew: true };
}

function resolveFinish(finishCode, etcName, etcEng, dryRun) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CLASS);
  const rows = classRows_('마감');

  if (finishCode && finishCode !== '__etc') {
    const r = rows.find(x => x[1] === finishCode);
    if (r) return { code: r[1], name: r[2], isNew: false };
    throw new Error('알 수 없는 마감 코드: ' + finishCode);
  }
  if (!etcName) throw new Error('기타 마감명을 입력해 주세요.');
  const exist = rows.find(r => normName(r[2]) === normName(etcName));
  if (exist) return { code: exist[1], name: exist[2], isNew: false };

  const existing = rows.map(r => String(r[1] || '').toUpperCase());
  const code = make2_(etcEng || 'X', existing);
  if (!dryRun) {
    sheet.appendRow(['마감', code, etcName, etcEng ? etcEng.toUpperCase().trim() : '(영문명 확인 필요)', '', '기타 응답 자동 등록 — 검토 필요']);
    clearClassCache_();
  }
  return { code: code, name: etcName, isNew: true };
}

function resolveGloss(glossCode, etcName, etcEng, dryRun) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CLASS);
  const rows = classRows_('광택');

  if (glossCode && glossCode !== '__etc') {
    const builtin = GLOSSES.find(g => g.code === glossCode);
    if (builtin) return { code: builtin.code, name: builtin.name, isNew: false };
    const r = rows.find(x => x[1] === glossCode);
    if (r) return { code: r[1], name: r[2], isNew: false };
    throw new Error('알 수 없는 광택 코드: ' + glossCode);
  }
  if (!etcName) throw new Error('기타 광택명을 입력해 주세요.');

  const sameBuiltin = GLOSSES.find(g => normName(g.name) === normName(etcName));
  if (sameBuiltin) return { code: sameBuiltin.code, name: sameBuiltin.name, isNew: false };
  const exist = rows.find(r => normName(r[2]) === normName(etcName));
  if (exist) return { code: exist[1], name: exist[2], isNew: false };

  const existing = GLOSSES.map(g => g.code).concat(rows.map(r => String(r[1] || '').toUpperCase()));
  const code = make2_(etcEng || 'X', existing);
  if (!dryRun) {
    sheet.appendRow(['광택', code, etcName, etcEng ? etcEng.toUpperCase().trim() : '(영문명 확인 필요)', '', '기타 응답 자동 등록 — 검토 필요']);
    clearClassCache_();
  }
  return { code: code, name: etcName, isNew: true };
}

/* ═════════════════════════════════════════════
   디자인 작업지시서 소재 추출
   텍스트 추출은 브라우저에서 처리하고(pdf.js / JSZip / OCR),
   서버는 사전에 등록된 용어와 HEX·Pantone 패턴으로 CMF 후보만 뽑는다.
   ※ 문서를 '이해'하는 것이 아니라 용어와 패턴을 찾는 방식이므로
     결과는 반드시 사람이 확인한 뒤 등록해야 한다.
   ═════════════════════════════════════════════ */

/* 브라우저에서 뽑아낸 텍스트를 받아 CMF 후보로 정리한다.
   파일 자체는 서버로 올리지 않으므로 Drive API·업로드 권한이 필요 없다. */
function parseWorkOrderText(text, fileName) {
  const t = String(text || '');
  return {
    fileName: String(fileName || ''),
    chars: t.length,
    items: parseCmfCandidates_(t),
    preview: t.slice(0, 800)
  };
}

function parseCmfCandidates_(text) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mats = classRows_('소재').map(r => ({ code: r[1], name: canonMatName_(r[2]), eng: String(r[3] || '') }));
  const fins = classRows_('마감').map(r => ({ code: r[1], name: r[2], eng: String(r[3] || '') }));
  const glos = GLOSSES.map(g => ({ code: g.code, name: g.name, eng: '' }))
    .concat(classRows_('광택').map(r => ({ code: r[1], name: r[2], eng: String(r[3] || '') })));
  const cRows = ss.getSheetByName(CONFIG.SHEET_COLOR).getDataRange().getValues();
  const colors = cRows.slice(1).filter(r => r[0])
    .map(r => ({ code: '', name: String(r[0]), eng: String(r[1] || ''), hex: String(r[4] || '') }));

  const lines = String(text || '').split(/\r?\n/)
    .map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
  const win = i => lines.slice(Math.max(0, i - 3), i + 4).join('  ');

  // 사전 용어 중 가장 긴 것을 채택 ('가죽'보다 '인조가죽'을 우선)
  function findTerm(hay, list) {
    const h = String(hay).toLowerCase();
    let best = null;
    list.forEach(t => {
      [t.name, t.eng].filter(Boolean).forEach(c => {
        const cc = String(c).toLowerCase().trim();
        if (cc.length < 2) return;
        if (h.indexOf(cc) >= 0 && (!best || cc.length > best.len)) {
          best = { code: t.code, name: t.name, len: cc.length };
        }
      });
    });
    return best;
  }

  const HEX = /#([0-9a-fA-F]{6})\b/g;
  const PANTONE = /PANTONE\s*([0-9]{2,4}\s*[CUcu]?)/i;
  const SUPCODE = /\b([A-Z]{1,4}-?\d{3,6})\b/;
  const PRODUCT = /(?:제품|모델|품명|MODEL)\s*[:：]\s*([^\s|][^|]{1,25}?)(?:\s{2,}|$)/i;

  const out = [], seen = {};
  function add(i, hex, colorName, lineOnly) {
    const line = lines[i], w = win(i);
    const pick = list => lineOnly ? findTerm(line, list) : (findTerm(line, list) || findTerm(w, list));
    const grab = re => (re.exec(line) || (lineOnly ? null : re.exec(w)) || [])[1] || '';
    const grabLine = re => (re.exec(line) || [])[1] || '';

    const mat = pick(mats), fin = pick(fins), glo = pick(glos);
    let cn = colorName;
    if (!cn) { const c = pick(colors); cn = c ? c.name : ''; }
    const pan = grabLine(PANTONE), sup = grabLine(SUPCODE), pr = grab(PRODUCT);
    let hx = hex ? ('#' + String(hex).replace('#', '').toUpperCase()) : '';

    let hexFrom = hx ? 'doc' : '';
    if (!hx && cn) {
      const c = colors.find(x => x.name === cn);
      if (c && /^#?[0-9a-fA-F]{6}$/.test(String(c.hex).trim())) {
        hx = '#' + String(c.hex).trim().replace('#', '').toUpperCase();
        hexFrom = 'dict';
      }
    }

    const key = (hx || cn) ? (hx + '|' + cn + '|' + (mat ? mat.code : '')) : ('line' + i);
    if (seen[key]) return;
    seen[key] = 1;

    out.push({
      li: i,
      low: !hx && !cn,
      hex: hx,
      hexFrom: hexFrom,
      colorName: cn,
      colorEng: '',
      matCode: mat ? mat.code : '', matName: mat ? mat.name : '',
      finCode: fin ? fin.code : '', finName: fin ? fin.name : '',
      gloCode: glo ? glo.code : '', gloName: glo ? glo.name : '',
      pantone: pan ? ('PANTONE ' + pan.replace(/\s+/g, '').toUpperCase()) : '',
      supcode: sup,
      product: pr.trim(),
      source: line
    });
  }

  lines.forEach((ln, i) => {
    let m; HEX.lastIndex = 0;
    while ((m = HEX.exec(ln)) !== null) add(i, m[1], '');
  });
  if (!out.length) {
    lines.forEach((ln, i) => {
      const c = findTerm(ln, colors);
      if (c) add(i, '', c.name);
    });
  }

  const covered = {};
  out.forEach(o => { covered[o.li] = 1; });

  const numbered = [];
  lines.forEach((ln, i) => { if (/^\s*\d{1,2}\s*[.)\]]?\s+\S/.test(ln)) numbered.push(i); });
  if (numbered.some(i => covered[i])) {
    numbered.forEach(i => { if (!covered[i]) { add(i, '', '', true); covered[i] = 1; } });
  }

  lines.forEach((ln, i) => {
    if (covered[i]) return;
    if (findTerm(ln, mats)) { add(i, '', '', true); covered[i] = 1; }
  });

  return out.sort((a, b) => a.li - b.li).slice(0, 40);
}

/* ═════════════════════════════════════════════
   일회성 유지보수 함수 — Apps Script 편집기에서 함수 선택 후 "실행"
   결과는 편집기 아래 '실행 로그' 에 찍힌다.
   ═════════════════════════════════════════════ */

function say_(msg) {
  Logger.log(msg);
  return msg;
}

function planCodeMigration_(data) {
  const used = {};
  for (let i = 1; i < data.length; i++) {
    const m = CODE_RE_YEAR.exec(String(data[i][0] || '').toUpperCase());
    if (!m) continue;
    const key = comboKey_(m[1], m[2], m[3], m[4], m[5]);
    used[key] = Math.max(used[key] || 0, parseInt(m[6], 10) || 0);
  }

  const plan = [], skipped = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const old = String(r[0] || '').trim();
    if (!old) continue;
    if (CODE_RE_YEAR.test(old.toUpperCase())) continue;

    let yy, mc, cc, fc, gc;
    const fd = CODE_RE_FULLDATE.exec(old.toUpperCase());
    if (fd) {
      yy = fd[1].slice(0, 2); mc = fd[2]; cc = fd[3]; fc = fd[4]; gc = fd[5];
    } else {
      const at = r[24];
      const d = at ? new Date(at) : null;
      if (!d || isNaN(d.getTime())) { skipped.push(old + ' (등록일 없음 — 수동 확인 필요)'); continue; }
      yy = Utilities.formatDate(d, 'Asia/Seoul', 'yy');
      mc = two_(r[2]); cc = two_(r[4]); fc = two_(r[6]); gc = glossCode2_(r[8]);
    }

    const key = comboKey_(yy, mc, cc, fc, gc);
    const n = (used[key] || 0) + 1;
    used[key] = n;
    const ser = String(n).padStart(2, '0');

    plan.push({ rowNum: i + 1, old: old, next: [yy, mc, cc, fc, gc, ser].join('-'), gloss: gc, serial: ser });
  }
  return { plan: plan, skipped: skipped };
}

function previewCodeMigration() {
  const master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MASTER);
  const { plan, skipped } = planCodeMigration_(master.getDataRange().getValues());
  const out = [
    '변환 대상 ' + plan.length + '건',
    plan.map(p => p.old + '  →  ' + p.next).join('\n') || '(없음)',
    skipped.length ? '\n건너뜀 ' + skipped.length + '건\n' + skipped.join('\n') : ''
  ].join('\n');
  Logger.log(out);
  return out;
}

function migrateCodesToDated() {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(CONFIG.SHEET_MASTER);
    const data = master.getDataRange().getValues();
    const { plan, skipped } = planCodeMigration_(data);
    if (!plan.length) return say_('변환할 코드가 없습니다. (건너뜀 ' + skipped.length + '건)');

    const n = master.getLastRow() - 1;
    if (!String(master.getRange(1, COL_LEGACY_CODE).getValue() || '').trim()) {
      master.getRange(1, COL_LEGACY_CODE).setValue('이전코드');
    }

    const colA = master.getRange(2, 1, n, 1).getValues();
    const colGloss = master.getRange(2, 9, n, 1).getValues();
    const colSerial = master.getRange(2, 11, n, 1).getValues();
    const colLegacy = master.getRange(2, COL_LEGACY_CODE, n, 1).getValues();

    const remap = {};
    plan.forEach(p => {
      const i = p.rowNum - 2;
      colA[i][0] = p.next;
      colGloss[i][0] = p.gloss;
      colSerial[i][0] = p.serial;
      if (!String(colLegacy[i][0] || '').trim()) colLegacy[i][0] = p.old;
      remap[p.old] = p.next;
    });

    master.getRange(2, 1, n, 1).setValues(colA);
    master.getRange(2, 9, n, 1).setValues(colGloss);
    master.getRange(2, 11, n, 1).setValues(colSerial);
    master.getRange(2, COL_LEGACY_CODE, n, 1).setValues(colLegacy);

    let applyFixed = 0;
    const apply = ss.getSheetByName(CONFIG.SHEET_APPLY);
    if (apply && apply.getLastRow() > 1) {
      const m = apply.getLastRow() - 1;
      const codes = apply.getRange(2, 1, m, 1).getValues();
      codes.forEach(row => {
        const c = String(row[0] || '').trim();
        if (remap[c]) { row[0] = remap[c]; applyFixed++; }
      });
      apply.getRange(2, 1, m, 1).setValues(codes);
    }

    const counter = ss.getSheetByName(CONFIG.SHEET_COUNTER);
    if (counter) counter.clear().appendRow(['조합키', '최대발급번호']);
    initCounterFromExisting();

    return say_([
      '코드 ' + plan.length + '건 변환 완료',
      '적용이력 ' + applyFixed + '행 갱신',
      skipped.length ? '건너뜀 ' + skipped.length + '건: ' + skipped.join(' / ') : '',
      '옛 코드는 AE열(이전코드)에 보관되어 아카이브 검색에서도 찾을 수 있습니다.'
    ].filter(Boolean).join('\n'));
  } finally {
    lock.releaseLock();
  }
}

function cleanDuplicateGloss() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CLASS);
  const data = sheet.getDataRange().getValues();
  const builtin = new Set();
  GLOSSES.forEach(g => { builtin.add(g.code); builtin.add(g.legacy); builtin.add(normName(g.name)); });

  const removed = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] !== '광택') continue;
    const c = String(data[i][1] || '').toUpperCase(), n = normName(data[i][2] || '');
    if (builtin.has(c) || builtin.has(n)) {
      sheet.deleteRow(i + 1);
      removed.push(data[i][2]);
    }
  }
  clearClassCache_();
  return say_(removed.length ? '삭제: ' + removed.reverse().join(', ') : '중복 광택 없음');
}

function checkSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.SHEET_MASTER).getDataRange().getValues();
  let ok = 0, needFix = 0, blank = 0;
  const samples = [];
  for (let i = 1; i < master.length; i++) {
    const c = String(master[i][0] || '').trim();
    if (!c) { blank++; continue; }
    if (CODE_RE_YEAR.test(c.toUpperCase())) ok++;
    else { needFix++; if (samples.length < 5) samples.push(c); }
  }
  let url = '(배포 안 됨)';
  try { url = ScriptApp.getService().getUrl() || url; } catch (err) {}

  return say_([
    '── CMF 시스템 진단 ──',
    'Code.gs 버전: ' + APP_VERSION + '   (코드 형식: 26-FB-SG-EM-MA-01)',
    '화면 밝기: ' + getTheme(),
    '',
    '등록대장 ' + (master.length - 1) + '행',
    '  연도 형식(정상) : ' + ok + '건',
    '  변환 필요       : ' + needFix + '건' + (samples.length ? '   예) ' + samples.join(', ') : ''),
    '  코드 없음       : ' + blank + '건',
    '',
    '작업지시서 분석: 브라우저에서 처리 (Drive API 불필요)',
    '웹앱 URL: ' + url,
    '',
    '※ 이 주소를 런처(Design Applicators) 프로젝트의 Code.gs 안',
    '   CONFIG.CMF_URL 에 붙여넣으면 CMF LAB 카드가 이 화면을 엽니다.',
    '※ 코드를 붙여넣은 뒤 배포 → 배포 관리 → 연필 → 버전 "새 버전" → 배포 를 해야 화면에 반영됩니다.'
  ].join('\n'));
}

function setupNewColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MASTER);
  const set = (col, name) => {
    if (!String(sheet.getRange(1, col).getValue() || '').trim()) {
      sheet.getRange(1, col).setValue(name);
      return name;
    }
    return null;
  };
  const added = [set(COL_LEGACY_CODE, '이전코드'), set(COL_DEPT, '부서')].filter(Boolean);
  return say_(added.length ? '머리글 추가: ' + added.join(', ') : '이미 설정되어 있습니다.');
}

function renameMaterialTerms() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CLASS);
  const data = sheet.getDataRange().getValues();
  const changes = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== '소재') continue;
    const cur = String(data[i][2] || '').trim();
    const next = canonMatName_(cur);
    if (next && next !== cur) {
      sheet.getRange(i + 1, 3).setValue(next);
      changes.push(cur + ' → ' + next);
    }
  }
  clearClassCache_();
  return say_(changes.length ? changes.join('\n') : '통일할 소재명 없음');
}

function seedMaterialDictionary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_CLASS);
  const rows = classRows_('소재');
  const have = new Set(rows.map(r => canonMatName_(r[2])));
  const codes = rows.map(r => String(r[1] || '').toUpperCase());
  const seed = [
    ['PU',       'PU',        'PU'],
    ['패브릭',   'FABRIC',    'FB'],
    ['레깅스',   'LEGGINGS',  'LG'],
    ['가죽',     'LEATHER',   'LT'],
    ['스웨이드', 'SUEDE',     'SD'],
    ['사출',     'INJECTION', 'IJ'],
    ['금속',     'METAL',     'MT'],
    ['스프레이', 'SPRAY',     'SP'],
    ['실리콘',   'SILICONE',  'SI'],
    ['우드',     'WOOD',      'WD']
  ];
  const added = [];
  seed.forEach(([ko, en, pref]) => {
    if (have.has(ko)) return;
    const code = (codes.indexOf(pref) < 0) ? pref : make2_(en, codes);
    codes.push(code);
    sheet.appendRow(['소재', code, ko, en, '-', 'BOX_ORDER 시드']);
    added.push(ko + '=' + code);
  });
  clearClassCache_();
  return say_(added.length ? added.join(', ') : '추가할 소재 없음 (10종 모두 존재)');
}

function normalizeDictCodes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_CLASS);
  const data = sheet.getDataRange().getValues();
  const changes = [];

  ['소재', '마감'].forEach(kind => {
    const used = data.filter(r => r[0] === kind).map(r => String(r[1] || '').toUpperCase());
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== kind) continue;
      const cur = String(data[i][1] || '').toUpperCase();
      if (cur.length === 2) continue;
      const next = make2_(data[i][3] || cur, used);
      used.push(next);
      sheet.getRange(i + 1, 2).setValue(next);
      changes.push(kind + ' ' + data[i][2] + ': ' + cur + ' → ' + next);
    }
  });

  const cs = ss.getSheetByName(CONFIG.SHEET_COLOR);
  const cd = cs.getDataRange().getValues();
  const cUsed = cd.slice(1).map(r => String(r[2] || '').toUpperCase());
  for (let i = 1; i < cd.length; i++) {
    const cur = String(cd[i][2] || '').toUpperCase();
    if (!cur || cur.length === 2) continue;
    const next = make2_(cd[i][1] || cur, cUsed);
    cUsed.push(next);
    cs.getRange(i + 1, 3).setValue(next);
    changes.push('컬러 ' + cd[i][0] + ': ' + cur + ' → ' + next);
  }

  clearClassCache_();
  return say_(changes.length ? changes.join('\n') : '2자가 아닌 코드 없음');
}

function backfillBoxes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MASTER);
  const data = sheet.getDataRange().getValues();
  let filled = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    sheet.getRange(i + 1, COL_BOX).setValue(assignBox(data[i][2]));
    filled++;
  }
  return say_('보관 장소 다시 계산 — ' + filled + '행');
}

function fixPhotoUrls() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MASTER);
  const data = sheet.getDataRange().getValues();
  const COL_PHOTO = 22;
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const url = data[i][21];
    if (url && String(url).indexOf('drive.google.com/uc') !== -1) {
      const m = String(url).match(/id=([-\w]{20,})/);
      if (m) {
        sheet.getRange(i + 1, COL_PHOTO).setValue('https://lh3.googleusercontent.com/d/' + m[1] + '=s800');
        count++;
      }
    }
  }
  return say_('사진 주소 고침 — ' + count + '행');
}
