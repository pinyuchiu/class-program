/**
 * 金煎坊 Golden Pan — 線上訂單 → Google Sheet
 * ------------------------------------------------------------
 * 部署方式見 gas/README.md
 *   1. 在 Google Sheet 開啟「擴充功能 → Apps Script」，貼上本檔
 *   2. 執行一次 setup()（建立工作表與表頭、授權）
 *   3. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我／存取權：所有人）
 *   4. 把網址貼到 index.html 的 GAS_URL
 */

// ===== 設定 =====
const CONFIG = {
  ORDER_SHEET: '訂單',
  ITEM_SHEET: '訂單明細',
  NOTIFY_EMAIL: '',          // 有新訂單時寄信通知（留空＝不寄）
  TIMEZONE: 'Asia/Taipei',
  STATUSES: ['新訂單', '製作中', '可取餐／已出貨', '已完成', '已取消'],
  MAX_ITEMS: 50,             // 單筆訂單最多品項數（防濫用）
  MAX_QTY: 99,               // 單一品項最多數量
};

const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '手機', '門市', '預計時間', '宅配地址', '備註', '品項摘要', '小計', '運費', '應付金額'];
const ITEM_HEADERS  = ['訂單編號', '建立時間', '品項ID', '品名', '單價', '數量', '小計'];
const MODE_LABEL = { dine: '內用', takeout: '外帶', delivery: '冷凍宅配' };

// ===== 初始化：建立工作表、表頭、格式與狀態下拉選單 =====
function setup() {
  const ss = SpreadsheetApp.getActive();
  const orders = ensureSheet_(ss, CONFIG.ORDER_SHEET, ORDER_HEADERS);
  ensureSheet_(ss, CONFIG.ITEM_SHEET, ITEM_HEADERS);

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).build();
  orders.getRange(2, 3, orders.getMaxRows() - 1, 1).setDataValidation(rule);
  orders.getRange('L:N').setNumberFormat('"NT$"#,##0');
  ss.getSheetByName(CONFIG.ITEM_SHEET).getRange('E:G').setNumberFormat('"NT$"#,##0');
  orders.getRange('F:F').setNumberFormat('@');   // 手機保留開頭的 0
  return '完成：已建立「' + CONFIG.ORDER_SHEET + '」與「' + CONFIG.ITEM_SHEET + '」';
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4A3226').setFontColor('#FAF6EF');
  sh.setFrozenRows(1);
  return sh;
}

// ===== 健康檢查：用瀏覽器開網址可確認部署成功 =====
function doGet() {
  return json_({ ok: true, service: 'golden-pan-orders' });
}

// ===== 接收訂單 =====
function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.website) return json_({ ok: true, id: 'IGNORED' });   // 蜜罐欄位：機器人才會填

    const order = validate_(data);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);                        // 避免同時下單時編號重複
    let id, now;
    try {
      now = new Date();
      id = nextOrderId_(now);
      writeOrder_(id, now, order);
    } finally {
      lock.releaseLock();
    }
    notify_(id, order);
    return json_({ ok: true, id: id, total: order.total });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function validate_(d) {
  const s = v => String(v == null ? '' : v).trim().slice(0, 300);
  const mode = s(d.mode);
  if (!MODE_LABEL[mode]) throw new Error('取餐方式錯誤');
  const name = s(d.name);
  if (!name) throw new Error('請填寫姓名');
  const phone = s(d.phone).replace(/[\s-]/g, '');
  if (!/^09\d{8}$/.test(phone)) throw new Error('手機號碼格式錯誤');
  const address = s(d.address);
  if (mode === 'delivery' && address.length < 5) throw new Error('請填寫宅配地址');

  const items = Array.isArray(d.items) ? d.items.slice(0, CONFIG.MAX_ITEMS) : [];
  if (!items.length) throw new Error('購物車是空的');
  const clean = items.map(it => {
    const qty = Math.floor(Number(it.qty));
    const price = Math.round(Number(it.price));
    if (!(qty >= 1 && qty <= CONFIG.MAX_QTY) || !(price >= 0)) throw new Error('品項數量或價格錯誤');
    return { id: s(it.id), name: s(it.name), price: price, qty: qty, sub: price * qty };
  });
  const subtotal = clean.reduce((n, it) => n + it.sub, 0);
  const shipping = Math.max(0, Math.round(Number(d.shipping) || 0));

  return {
    mode: mode, name: name, phone: phone, address: mode === 'delivery' ? address : '',
    store: mode === 'delivery' ? '' : s(d.store), time: mode === 'delivery' ? '' : s(d.time),
    note: s(d.note), items: clean, subtotal: subtotal, shipping: shipping, total: subtotal + shipping,
  };
}

// 訂單編號：GP + 日期 + 當日流水號，例如 GP20261127-0001
function nextOrderId_(now) {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyyMMdd');
  const key = 'seq_' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return 'GP' + day + '-' + ('000' + n).slice(-4);
}

function writeOrder_(id, now, o) {
  const ss = SpreadsheetApp.getActive();
  const orders = ss.getSheetByName(CONFIG.ORDER_SHEET) || ensureSheet_(ss, CONFIG.ORDER_SHEET, ORDER_HEADERS);
  const itemsSh = ss.getSheetByName(CONFIG.ITEM_SHEET) || ensureSheet_(ss, CONFIG.ITEM_SHEET, ITEM_HEADERS);
  const summary = o.items.map(it => it.name + ' ×' + it.qty).join('、');
  // 以 ' 開頭避免手機與文字被試算表當成數字或公式
  const txt = v => (v === '' ? '' : "'" + v);
  orders.appendRow([id, now, CONFIG.STATUSES[0], MODE_LABEL[o.mode], txt(o.name), txt(o.phone), o.store, o.time,
                    txt(o.address), txt(o.note), txt(summary), o.subtotal, o.shipping, o.total]);
  const rows = o.items.map(it => [id, now, it.id, txt(it.name), it.price, it.qty, it.sub]);
  itemsSh.getRange(itemsSh.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);
}

function notify_(id, o) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  try {
    const lines = o.items.map(it => '・' + it.name + ' ×' + it.qty + '　NT$' + it.sub);
    MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, '【金煎坊新訂單】' + id + '（' + MODE_LABEL[o.mode] + '）NT$' + o.total,
      ['訂單編號：' + id, '取餐方式：' + MODE_LABEL[o.mode], '姓名：' + o.name, '手機：' + o.phone,
       o.mode === 'delivery' ? '地址：' + o.address : '門市：' + o.store + '　時間：' + o.time,
       '備註：' + (o.note || '—'), '', lines.join('\n'), '', '應付金額：NT$' + o.total].join('\n'));
  } catch (err) { console.warn('通知信寄送失敗', err); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
