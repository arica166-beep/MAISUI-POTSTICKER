/**
 * 麥穗鍋貼所｜線上訂單 → Google Sheet（Google Apps Script 後端）
 *
 * 功能
 *  - doPost：接收網站結帳資料，驗證後寫入「訂單」「訂單明細」，回傳訂單編號
 *  - 金額以「菜單」工作表為準重新計算（防止前端竄改價格）
 *  - 寄出顧客確認信、門市新訂單通知
 *  - 在「訂單」把狀態改成「可取餐／已出貨／已取消」時，自動通知顧客
 *  - 「設定」工作表可一鍵暫停接單、調整運費與免運門檻
 *
 * 第一次使用：在 Apps Script 編輯器選擇 setup 函式並執行，再部署成網頁應用程式。
 * 詳細步驟見 gas/README.md。
 */

const SHEET = {
  ORDERS: '訂單',
  ITEMS: '訂單明細',
  MENU: '菜單',
  STORES: '門市',
  SETTINGS: '設定',
  DAILY: '每日統計',
};

const ORDER_HEADERS = [
  '訂單編號', '建立時間', '狀態', '取餐方式', '門市', '取餐時間／到貨日', '用餐人數',
  '姓名', '手機', 'Email', '縣市', '地址', '付款方式', '付款狀態',
  '品項摘要', '小計', '運費', '總額', '前台金額', '金額檢查', '備註',
  '狀態更新時間', '已通知狀態', '請求ID',
];
const COL = Object.fromEntries(ORDER_HEADERS.map((h, i) => [h, i + 1]));

const STATUS = ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'];
const NOTIFY_STATUS = ['可取餐', '已出貨', '已取消'];
const METHOD_LABEL = { dine: '內用', takeout: '外帶', delivery: '冷凍宅配' };
const PAY_OPTIONS = ['店內付款', 'LINE Pay', '信用卡'];

// 第一次 setup 時寫入「菜單」工作表的資料（與網站上的品項一致）
const SEED_MENU = [
  ['p1', '招牌黃金鍋貼', '招牌鍋貼', 80, '10 顆'],
  ['p2', '韭菜鮮肉鍋貼', '招牌鍋貼', 85, '10 顆'],
  ['p3', '麻辣起司鍋貼', '招牌鍋貼', 105, '10 顆'],
  ['p4', '鮮蝦玉米鍋貼', '招牌鍋貼', 120, '10 顆'],
  ['d1', '高麗菜鮮肉水餃', '精選水餃', 75, '10 顆'],
  ['d2', '玉米鮮肉水餃', '精選水餃', 75, '10 顆'],
  ['d3', '紅油抄手', '精選水餃', 90, '8 顆'],
  ['d4', '香菇素食水餃', '精選水餃', 80, '10 顆'],
  ['s1', '酸辣湯', '湯品飲料', 45, '碗'],
  ['s2', '玉米濃湯', '湯品飲料', 45, '碗'],
  ['s3', '古早味豆漿', '湯品飲料', 35, '杯'],
  ['s4', '冬瓜檸檬', '湯品飲料', 40, '杯'],
  ['c1', '涼拌小黃瓜', '經典小菜', 40, '份'],
  ['c2', '皮蛋豆腐', '經典小菜', 50, '份'],
  ['c3', '麻辣鴨血', '經典小菜', 60, '份'],
  ['c4', '滷味拼盤', '經典小菜', 70, '份'],
  ['g1', '招財金鍋貼禮盒', '冷凍禮盒', 680, '50 入'],
  ['g2', '闔家團圓綜合禮盒', '冷凍禮盒', 980, '60 入'],
  ['g3', '小家庭常備組', '冷凍禮盒', 720, '60 入'],
  ['g4', '麻辣控挑戰組', '冷凍禮盒', 890, '54 入'],
];
const SEED_STORES = [
  ['flagship', '大安旗艦店', ''],
  ['xinyi', '信義店', ''],
  ['zhongshan', '中山店', ''],
  ['banqiao', '板橋店', ''],
];

/* =====================================================================
 * 初始化：建立工作表、表頭、格式、下拉選單與觸發器（可重複執行）
 * ===================================================================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Taipei');

  // 訂單
  const orders = ensureSheet_(ss, SHEET.ORDERS, ORDER_HEADERS);
  orders.getRange(2, COL['手機'], orders.getMaxRows() - 1, 1).setNumberFormat('@');
  orders.getRange(2, COL['建立時間'], orders.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange(2, COL['狀態更新時間'], orders.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  ['小計', '運費', '總額', '前台金額'].forEach(h =>
    orders.getRange(2, COL[h], orders.getMaxRows() - 1, 1).setNumberFormat('#,##0'));
  orders.getRange(2, COL['狀態'], orders.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build());
  orders.getRange(2, COL['付款狀態'], orders.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未付款', '已付款', '已退款'], true).build());
  [COL['品項摘要'], COL['地址'], COL['備註']].forEach(c => orders.setColumnWidth(c, 260));
  orders.hideColumns(COL['已通知狀態']);
  orders.hideColumns(COL['請求ID']);
  addStatusColors_(orders);

  // 訂單明細
  const items = ensureSheet_(ss, SHEET.ITEMS, ['訂單編號', '建立時間', '品項代碼', '品名', '單價', '數量', '金額']);
  items.getRange(2, 2, items.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');

  // 菜單（已有資料就不覆蓋）
  const menu = ensureSheet_(ss, SHEET.MENU, ['品項代碼', '品名', '分類', '單價', '單位', '上架']);
  if (menu.getLastRow() < 2) {
    menu.getRange(2, 1, SEED_MENU.length, 6).setValues(SEED_MENU.map(r => r.concat([true])));
  }
  menu.getRange(2, 6, menu.getMaxRows() - 1, 1).insertCheckboxes();

  // 門市（各門市可填自己的通知信箱）
  const stores = ensureSheet_(ss, SHEET.STORES, ['門市代碼', '門市名稱', '新訂單通知 Email']);
  if (stores.getLastRow() < 2) stores.getRange(2, 1, SEED_STORES.length, 3).setValues(SEED_STORES);
  stores.setColumnWidth(3, 260);

  // 設定
  const settings = ensureSheet_(ss, SHEET.SETTINGS, ['項目', '值', '說明']);
  const defaults = [
    ['接單', '開', '改成「關」即可暫停網站接單'],
    ['店家名稱', '麥穗鍋貼所', '顯示在通知信中'],
    ['總部通知 Email', Session.getEffectiveUser().getEmail() || '', '所有新訂單都會通知這個信箱（可留空）'],
    ['寄送顧客確認信', '是', '顧客有填 Email 時寄出訂單確認信'],
    ['狀態變更通知顧客', '是', '狀態改成可取餐／已出貨／已取消時通知顧客'],
    ['宅配運費', 160, '冷凍宅配的運費'],
    ['免運門檻', 1000, '冷凍宅配滿額免運'],
  ];
  const existing = settings.getLastRow() > 1 ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().flat() : [];
  defaults.filter(d => !existing.includes(d[0])).forEach(d => settings.appendRow(d));
  settings.setColumnWidth(1, 160); settings.setColumnWidth(2, 260); settings.setColumnWidth(3, 320);

  // 每日統計（公式自動彙整，不含已取消）
  const daily = ss.getSheetByName(SHEET.DAILY) || ss.insertSheet(SHEET.DAILY);
  daily.clear();
  daily.getRange('A1').setFormula(
    `=IFERROR(QUERY({ARRAYFORMULA(IF('${SHEET.ORDERS}'!B2:B="",,TO_DATE(INT('${SHEET.ORDERS}'!B2:B)))),'${SHEET.ORDERS}'!C2:C,'${SHEET.ORDERS}'!R2:R},` +
    `"select Col1, count(Col1), sum(Col3) where Col1 is not null and Col2 <> '已取消' group by Col1 order by Col1 desc label Col1 '日期', count(Col1) '訂單數', sum(Col3) '營業額'",0),"尚無訂單")`);
  daily.getRange('A:A').setNumberFormat('yyyy/mm/dd');
  daily.getRange('C:C').setNumberFormat('#,##0');
  daily.setFrozenRows(1);

  const sheet1 = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);
  ss.setActiveSheet(orders);

  // 狀態變更通知：可安裝的 onEdit 觸發器（避免重複建立）
  const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onOrderEdit');
  if (!has) ScriptApp.newTrigger('onOrderEdit').forSpreadsheet(ss).onEdit().create();

  return '設定完成';
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#43291a').setFontColor('#fbf6ee');
  sh.setFrozenRows(1);
  return sh;
}

function addStatusColors_(sh) {
  const range = sh.getRange(2, COL['狀態'], sh.getMaxRows() - 1, 1);
  const colors = { '新訂單': '#fde2c4', '製作中': '#fff3c4', '可取餐': '#d6f0d0', '已出貨': '#d4e6f7', '已完成': '#e6e6e6', '已取消': '#f6d0d0' };
  const rules = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(x => x.getColumn() === COL['狀態']));
  Object.keys(colors).forEach(s => rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([range]).build()));
  sh.setConditionalFormatRules(rules);
}

/* =====================================================================
 * 網頁應用程式入口
 * ===================================================================== */
function doGet() {
  return json_({ ok: true, service: 'maisui-orders', accepting: getSettings_()['接單'] !== '關' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_(createOrder_(body));
  } catch (err) {
    console.error(err && err.stack || err);
    return json_({ ok: false, error: '系統忙碌中，請稍後再試，或直接來電門市訂購。' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================================
 * 建立訂單
 * ===================================================================== */
function createOrder_(d) {
  // 機器人防護：隱藏欄位被填寫就拒絕
  if (d.website) return { ok: false, error: '訂單送出失敗，請重新整理後再試。' };

  const settings = getSettings_();
  if (settings['接單'] === '關') return { ok: false, error: '目前暫停線上接單，請直接來電門市，造成不便敬請見諒。' };

  const err = validate_(d);
  if (err) return { ok: false, error: err };

  const cache = CacheService.getScriptCache();
  const reqKey = 'req_' + String(d.requestId).slice(0, 64);
  const dup = d.requestId && cache.get(reqKey);
  if (dup) return JSON.parse(dup);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 以「菜單」工作表的價格重新計算
  const menu = readMenu_(ss);
  const lines = [];
  for (const it of d.items) {
    const m = menu[it.id];
    if (!m || !m.active) return { ok: false, error: `「${it.name || it.id}」目前無法訂購，請從購物車移除後再送出。` };
    lines.push({ id: it.id, name: m.name, price: m.price, qty: it.qty, amount: m.price * it.qty });
  }
  const subtotal = lines.reduce((s, l) => s + l.amount, 0);
  const shipFee = Number(settings['宅配運費']) || 0, freeShip = Number(settings['免運門檻']) || 0;
  const shipping = d.method === 'delivery' && subtotal < freeShip ? shipFee : 0;
  const total = subtotal + shipping;
  const clientTotal = Number(d.clientTotal) || 0;

  const stores = readStores_(ss);
  const store = stores[d.store];

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let orderNo;
  try {
    const again = d.requestId && cache.get(reqKey);
    if (again) return JSON.parse(again);

    const now = new Date();
    orderNo = nextOrderNo_(now);
    const when = d.method === 'delivery' ? (d.date || '盡快出貨') : d.time;
    const row = [];
    row[COL['訂單編號'] - 1] = orderNo;
    row[COL['建立時間'] - 1] = now;
    row[COL['狀態'] - 1] = '新訂單';
    row[COL['取餐方式'] - 1] = METHOD_LABEL[d.method];
    row[COL['門市'] - 1] = d.method === 'delivery' ? '' : (store ? store.name : d.store);
    row[COL['取餐時間／到貨日'] - 1] = when;
    row[COL['用餐人數'] - 1] = d.method === 'dine' ? d.people : '';
    row[COL['姓名'] - 1] = d.name;
    row[COL['手機'] - 1] = d.phone;
    row[COL['Email'] - 1] = d.email || '';
    row[COL['縣市'] - 1] = d.method === 'delivery' ? d.city : '';
    row[COL['地址'] - 1] = d.method === 'delivery' ? d.address : '';
    row[COL['付款方式'] - 1] = d.pay;
    row[COL['付款狀態'] - 1] = '未付款';
    row[COL['品項摘要'] - 1] = lines.map(l => `${l.name}×${l.qty}`).join('、');
    row[COL['小計'] - 1] = subtotal;
    row[COL['運費'] - 1] = shipping;
    row[COL['總額'] - 1] = total;
    row[COL['前台金額'] - 1] = clientTotal;
    row[COL['金額檢查'] - 1] = clientTotal === total ? '一致' : '⚠ 不一致（以總額為準）';
    row[COL['備註'] - 1] = d.note || '';
    row[COL['狀態更新時間'] - 1] = now;
    row[COL['已通知狀態'] - 1] = '';
    row[COL['請求ID'] - 1] = d.requestId || '';

    const orders = ss.getSheetByName(SHEET.ORDERS);
    const r = nextRow_(orders, 1);
    orders.getRange(r, COL['手機']).setNumberFormat('@');   // 保留手機號碼開頭的 0
    orders.getRange(r, 1, 1, ORDER_HEADERS.length).setValues([row.map(safe_)]);
    const items = ss.getSheetByName(SHEET.ITEMS);
    items.getRange(nextRow_(items, lines.length), 1, lines.length, 7)
      .setValues(lines.map(l => [orderNo, now, l.id, l.name, l.price, l.qty, l.amount].map(safe_)));
    SpreadsheetApp.flush();

    const result = { ok: true, orderNo, subtotal, shipping, total };
    if (d.requestId) cache.put(reqKey, JSON.stringify(result), 21600);
    lock.releaseLock();

    // 寄信失敗不影響訂單成立
    const order = { orderNo, d, lines, subtotal, shipping, total, when, storeName: store ? store.name : '' };
    try { notifyStore_(order, settings, store); } catch (e) { console.warn('門市通知失敗', e); }
    if (d.email && settings['寄送顧客確認信'] !== '否') {
      try { sendCustomerConfirm_(order, settings); result.emailSent = true; } catch (e) { console.warn('確認信失敗', e); }
    }
    return result;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function validate_(d) {
  if (!METHOD_LABEL[d.method]) return '請選擇取餐方式。';
  if (!d.name || String(d.name).trim().length > 30) return '請填寫姓名（30 字以內）。';
  if (!/^09\d{8}$/.test(String(d.phone || ''))) return '請輸入正確的手機號碼（09 開頭共 10 碼）。';
  if (d.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(d.email))) return 'Email 格式不正確。';
  if (!PAY_OPTIONS.includes(d.pay)) return '請選擇付款方式。';
  if (d.method === 'delivery') {
    if (!d.address || String(d.address).trim().length < 4) return '請填寫宅配地址。';
  } else if (!d.store || !d.time) {
    return '請選擇門市與時間。';
  }
  if (!Array.isArray(d.items) || !d.items.length || d.items.length > 50) return '購物車是空的。';
  for (const it of d.items) {
    if (!Number.isInteger(it.qty) || it.qty < 1 || it.qty > 99) return '商品數量不正確。';
  }
  ['name', 'address', 'note', 'email', 'city', 'time', 'date', 'people'].forEach(k => {
    if (d[k] != null) d[k] = String(d[k]).trim().slice(0, k === 'note' || k === 'address' ? 200 : 60);
  });
  return '';
}

// 取得下一個空白列；列數不夠時自動擴充
function nextRow_(sh, need) {
  const r = sh.getLastRow() + 1;
  if (r + need - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), Math.max(need, 200));
  return r;
}

// 避免使用者輸入被試算表當成公式執行（=、+、-、@ 開頭）
function safe_(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function nextOrderNo_(now) {
  const day = Utilities.formatDate(now, 'Asia/Taipei', 'yyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq_' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return `MS${day}-${String(n).padStart(3, '0')}`;
}

function getSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return {};
  return Object.fromEntries(sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().map(r => [String(r[0]).trim(), typeof r[1] === 'string' ? r[1].trim() : r[1]]));
}

function readMenu_(ss) {
  const sh = ss.getSheetByName(SHEET.MENU);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(r => {
    if (r[0]) out[String(r[0]).trim()] = { name: r[1], price: Number(r[3]) || 0, active: r[5] === true || r[5] === 'TRUE' };
  });
  return out;
}

function readStores_(ss) {
  const sh = ss.getSheetByName(SHEET.STORES);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    if (r[0]) out[String(r[0]).trim()] = { name: r[1], email: String(r[2] || '').trim() };
  });
  return out;
}

/* =====================================================================
 * 通知信
 * ===================================================================== */
function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
const money_ = n => 'NT$' + Number(n).toLocaleString('en-US');

function orderTable_(o) {
  const rows = o.lines.map(l => `<tr><td style="padding:6px 0">${esc_(l.name)} × ${l.qty}</td><td style="text-align:right">${money_(l.amount)}</td></tr>`).join('');
  return `<table style="width:100%;border-collapse:collapse;font-size:14px">${rows}
    <tr><td style="padding-top:10px;border-top:1px solid #e9dcc6">小計</td><td style="padding-top:10px;border-top:1px solid #e9dcc6;text-align:right">${money_(o.subtotal)}</td></tr>
    <tr><td>運費</td><td style="text-align:right">${o.d.method === 'delivery' ? (o.shipping ? money_(o.shipping) : '免運') : '—'}</td></tr>
    <tr><td style="font-weight:bold;padding-top:6px">應付總額</td><td style="font-weight:bold;text-align:right;color:#c95f18">${money_(o.total)}</td></tr></table>`;
}
function pickupText_(o) {
  return o.d.method === 'delivery'
    ? `${esc_(o.d.city)}${esc_(o.d.address)}（希望到貨日：${esc_(o.when)}）`
    : `${esc_(o.storeName)}・${esc_(o.when)}${o.d.method === 'dine' ? '・' + esc_(o.d.people) : ''}`;
}
function mailWrap_(shop, inner) {
  return `<div style="max-width:560px;margin:auto;font-family:'Noto Sans TC',sans-serif;color:#2c1a10;background:#fbf6ee;padding:24px;border-radius:16px">
    <h2 style="margin:0 0 4px;font-family:'Noto Serif TC',serif">${esc_(shop)}</h2>${inner}
    <p style="margin-top:24px;font-size:12px;color:#8a5f3c">此信件由系統自動寄出，如有問題請直接回覆或來電門市。</p></div>`;
}

function sendCustomerConfirm_(o, settings) {
  const shop = settings['店家名稱'] || '麥穗鍋貼所';
  const html = mailWrap_(shop, `
    <p>${esc_(o.d.name)} 您好，感謝您的訂購！我們已收到您的訂單：</p>
    <p style="font-size:20px;font-weight:bold;color:#c95f18;margin:8px 0">${o.orderNo}</p>
    <p>取餐方式：${METHOD_LABEL[o.d.method]}<br>取餐資訊：${pickupText_(o)}<br>付款方式：${esc_(o.d.pay)}</p>
    ${orderTable_(o)}`);
  MailApp.sendEmail({ to: o.d.email, subject: `【${shop}】訂單確認 ${o.orderNo}`, htmlBody: html, name: shop });
}

function notifyStore_(o, settings, store) {
  const to = [settings['總部通知 Email'], store && store.email].filter(Boolean).filter((v, i, a) => a.indexOf(v) === i).join(',');
  if (!to) return;
  const shop = settings['店家名稱'] || '麥穗鍋貼所';
  const html = mailWrap_(shop, `
    <p style="font-size:18px;font-weight:bold">新訂單 ${o.orderNo}</p>
    <p>${METHOD_LABEL[o.d.method]}｜${pickupText_(o)}<br>
    顧客：${esc_(o.d.name)}（${esc_(o.d.phone)}）${o.d.email ? '｜' + esc_(o.d.email) : ''}<br>
    付款：${esc_(o.d.pay)}${o.d.note ? '<br>備註：' + esc_(o.d.note) : ''}</p>
    ${orderTable_(o)}
    <p><a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="color:#c95f18">開啟訂單試算表</a></p>`);
  MailApp.sendEmail({ to, subject: `【新訂單】${o.orderNo}｜${METHOD_LABEL[o.d.method]}｜${money_(o.total)}`, htmlBody: html, name: shop + ' 線上訂單' });
}

/* =====================================================================
 * 狀態變更：記錄時間，並在可取餐／已出貨／已取消時通知顧客
 * （由 setup 建立的可安裝 onEdit 觸發器呼叫）
 * ===================================================================== */
function onOrderEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET.ORDERS || e.range.getColumn() > COL['狀態'] || e.range.getLastColumn() < COL['狀態'] || e.range.getRow() < 2) return;
  const settings = getSettings_();
  const shop = settings['店家名稱'] || '麥穗鍋貼所';
  for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
    const row = sh.getRange(r, 1, 1, ORDER_HEADERS.length).getValues()[0];
    const status = row[COL['狀態'] - 1];
    if (!row[0] || !status) continue;
    sh.getRange(r, COL['狀態更新時間']).setValue(new Date());
    const email = row[COL['Email'] - 1];
    if (!NOTIFY_STATUS.includes(status) || !email || settings['狀態變更通知顧客'] === '否') continue;
    const notified = String(row[COL['已通知狀態'] - 1] || '').split(',');
    if (notified.includes(status)) continue;   // 同一個狀態只通知一次
    const msg = {
      '可取餐': `您的餐點已經準備好囉！請至 <b>${esc_(row[COL['門市'] - 1])}</b> 取餐，現場告知訂單編號即可。`,
      '已出貨': '您的冷凍宅配已經出貨，將由黑貓宅急便冷凍配送，請留意收件。',
      '已取消': '您的訂單已取消。如有疑問，請直接回覆此信或來電門市。',
    }[status];
    try {
      MailApp.sendEmail({
        to: email, name: shop,
        subject: `【${shop}】訂單 ${row[0]} ${status}`,
        htmlBody: mailWrap_(shop, `<p>${esc_(row[COL['姓名'] - 1])} 您好，</p><p>訂單 <b>${esc_(row[0])}</b>：${msg}</p>`),
      });
      sh.getRange(r, COL['已通知狀態']).setValue(notified.concat(status).filter(Boolean).join(','));
    } catch (err) { console.warn('狀態通知失敗', err); }
  }
}
