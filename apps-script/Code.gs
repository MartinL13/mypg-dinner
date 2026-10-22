/**
 * MYPG 四组晚餐服事 · Google Apps Script 后端
 *
 * 部署方式：Extensions → Apps Script → 粘贴本文件 → Deploy → New deployment
 *   类型 Web app · Execute as: Me · Who has access: Anyone
 * 拿到 .../exec 网址后填进 index.html 的 API_URL。
 *
 * 读写的是现有 Sheet2：
 *   D5 「小组成员」，E5 起每列一个服事日期；D6 起名字；格子里填 A / C / NA
 *   汇总区：D29 日期表头行，「拼菜总数」「取餐同工」「订餐餐厅」「预订菜品」「自制佳肴」按 D 列标签定位
 * 另外自动建两个标签页：
 *   「菜单」   餐厅菜单，换餐厅时改这里
 *   「网页数据」每人每月的细项（点了哪几道、带了什么、付款勾选）
 */

const SHEET_MAIN = 'Sheet2';
const SHEET_MENU = '菜单';
const SHEET_DATA = '网页数据';
const DEADLINE_HOUR = 19;          // 服事日前一天（周四）晚 7 点截止
const PRICE_PER_PERSON = 25;
const PER_DISH = 50;               // 一道菜 ≈ $50 ≈ 两人
const MAX_VOTES = 3;
const RESTAURANT = '香港园';

// ---------- HTTP ----------
function doGet(e) {
  return json(safe(() => state()));
}
function doPost(e) {
  return json(safe(() => {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'save') return save(body);
    throw new Error('unknown action');
  }));
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function safe(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return { ok: true, data: fn() }; }
  catch (err) { return { ok: false, error: String(err && err.message || err) }; }
  finally { lock.releaseLock(); }
}

// ---------- 读 ----------
function state() {
  ensureTabs();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_MAIN);
  const lay = layout(sh);
  const svc = currentService(sh, lay);
  const roster = lay.names.map(r => r.name);

  const people = {};
  lay.names.forEach(r => {
    const raw = String(sh.getRange(r.row, svc.col).getValue() || '').trim().toUpperCase();
    people[r.name] = { mode: modeFromCell(raw), votes: [], bring: [], paid: false, method: '', any: false };
  });
  // 细项覆盖
  const data = ss.getSheetByName(SHEET_DATA);
  const rows = data.getLastRow() > 1 ? data.getRange(2, 1, data.getLastRow() - 1, 9).getValues() : [];
  rows.forEach(r => {
    if (fmtDate(r[0]) !== svc.date) return;
    const name = String(r[1]).trim(); if (!name) return;
    if (!people[name]) { people[name] = { mode: null, votes: [], bring: [], paid: false, method: '', any: false }; roster.push(name); }
    const p = people[name];
    p.mode = r[2] ? String(r[2]) : p.mode;
    p.votes = split(r[3]); p.bring = split(r[4]); p.paid = isTrue(r[5]);
    p.method = String(r[7] || ''); p.any = isTrue(r[8]);
  });

  return {
    service: { date: svc.date, deadline: svc.deadline.toISOString(), locked: new Date() > svc.deadline, restaurant: RESTAURANT },
    roster, people, menu: readMenu(ss), config: { max: MAX_VOTES, price: PRICE_PER_PERSON, perDish: PER_DISH }
  };
}

function modeFromCell(raw) {
  if (!raw || raw === '-') return null;
  if (raw === 'NA' || raw === 'N/A') return 'NA';
  if (raw.indexOf('C') >= 0) return 'C';
  if (raw.indexOf('A') >= 0) return 'A';
  return null;
}
function split(v) { return String(v || '').split(';').map(s => s.trim()).filter(Boolean); }

function readMenu(ss) {
  const sh = ss.getSheetByName(SHEET_MENU);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues()
    .filter(r => r[0] !== '' && String(r[8]).toUpperCase() !== 'N')
    .map((r, i) => ({ id: String(r[0]), cat: String(r[1]), zh: String(r[2]), en: String(r[3]),
      price: r[4] === '' ? null : Number(r[4]), spicy: yes(r[5]), rec: yes(r[6]), note: String(r[7] || ''), order: i }));
}
function yes(v) { return String(v).trim().toUpperCase() === 'Y'; }
function isTrue(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }

// ---------- 写 ----------
function save(body) {
  ensureTabs();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_MAIN);
  let lay = layout(sh);
  const svc = currentService(sh, lay);
  if (new Date() > svc.deadline) throw new Error('已截止，改动请直接联系负责人');

  const name = String(body.name || '').trim().slice(0, 12);
  if (!name) throw new Error('缺少名字');
  const mode = ['A', 'C', 'NA'].indexOf(body.mode) >= 0 ? body.mode : null;
  const votes = (body.votes || []).map(String).slice(0, MAX_VOTES);
  const bring = (body.bring || []).map(s => String(s).trim().slice(0, 20)).filter(Boolean).slice(0, 8);
  const paid = !!body.paid;
  const method = ['Venmo', 'Zelle'].indexOf(body.method) >= 0 ? body.method : '';
  const any = !!body.any;

  // 1. 名单矩阵：新朋友插一行
  let who = lay.names.find(r => r.name === name);
  if (!who) {
    const last = lay.names[lay.names.length - 1].row;
    sh.insertRowAfter(last);
    sh.getRange(last + 1, lay.nameCol).setValue(name);
    lay = layout(sh);
    who = lay.names.find(r => r.name === name);
  }
  sh.getRange(who.row, svc.col).setValue(mode || '');

  // 2. 细项表
  const data = ss.getSheetByName(SHEET_DATA);
  const rows = data.getLastRow() > 1 ? data.getRange(2, 1, data.getLastRow() - 1, 2).getValues() : [];
  let rowIdx = -1;
  rows.forEach((r, i) => { if (fmtDate(r[0]) === svc.date && String(r[1]).trim() === name) rowIdx = i + 2; });
  const rec = [svc.date, name, mode || '', mode === 'C' ? votes.join(';') : '', mode === 'A' ? bring.join(';') : '', mode === 'C' && paid, new Date(),
    mode === 'C' && paid ? method : '', mode === 'C' && any];
  if (rowIdx < 0) data.appendRow(rec); else data.getRange(rowIdx, 1, 1, 9).setValues([rec]);

  // 3. 汇总区
  writeSummary(ss, sh, lay, svc);
  return state();
}

function writeSummary(ss, sh, lay, svc) {
  const st = stateLite(ss, sh, lay, svc);
  const C = Object.keys(st).filter(n => st[n].mode === 'C');
  const k = Math.floor(C.length / 2);
  // 票数排行
  const cnt = {};
  C.forEach(n => st[n].votes.forEach(id => { cnt[id] = (cnt[id] || 0) + 1; }));
  const menu = readMenu(ss);
  const byId = {}; menu.forEach(d => { byId[d.id] = d; });
  const ranked = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a] || ((byId[a] || {}).order || 999) - ((byId[b] || {}).order || 999));
  const picks = ranked.slice(0, k).map(id => byId[id] ? byId[id].zh : id);
  const bring = [];
  Object.keys(st).forEach(n => { if (st[n].mode === 'A' && st[n].bring.length) bring.push(n + '：' + st[n].bring.join('、')); });

  if (lay.sum.count) sh.getRange(lay.sum.count, svc.col).setValue(C.length || '');
  if (lay.sum.rest && !sh.getRange(lay.sum.rest, svc.col).getValue()) sh.getRange(lay.sum.rest, svc.col).setValue(RESTAURANT);
  fillBlock(sh, lay.sum.dishes, svc.col, picks);
  fillBlock(sh, lay.sum.bring, svc.col, bring);
}
function fillBlock(sh, block, col, lines) {
  if (!block) return;
  const n = block.end - block.start + 1;
  const vals = []; for (let i = 0; i < n; i++) vals.push([lines[i] || '']);
  sh.getRange(block.start, col, n, 1).setValues(vals);
}
function stateLite(ss, sh, lay, svc) {
  const st = {};
  lay.names.forEach(r => {
    st[r.name] = { mode: modeFromCell(String(sh.getRange(r.row, svc.col).getValue() || '').trim().toUpperCase()), votes: [], bring: [] };
  });
  const data = ss.getSheetByName(SHEET_DATA);
  const rows = data.getLastRow() > 1 ? data.getRange(2, 1, data.getLastRow() - 1, 7).getValues() : [];
  rows.forEach(r => {
    if (fmtDate(r[0]) !== svc.date) return;
    const name = String(r[1]).trim(); if (!name) return;
    st[name] = st[name] || { mode: null, votes: [], bring: [] };
    if (r[2]) st[name].mode = String(r[2]);
    st[name].votes = split(r[3]); st[name].bring = split(r[4]);
  });
  return st;
}

// ---------- 表结构 ----------
function layout(sh) {
  const vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  let headerRow = -1, nameCol = -1;
  for (let r = 0; r < vals.length && headerRow < 0; r++)
    for (let c = 0; c < vals[r].length; c++)
      if (String(vals[r][c]).trim() === '小组成员') { headerRow = r + 1; nameCol = c + 1; break; }
  if (headerRow < 0) throw new Error('找不到「小组成员」表头');
  const dates = [];
  for (let c = nameCol; c < vals[headerRow - 1].length; c++) {
    const d = toDate(vals[headerRow - 1][c]);
    if (d) dates.push({ col: c + 1, date: d });
  }
  const names = [];
  for (let r = headerRow; r < vals.length; r++) {
    const v = String(vals[r][nameCol - 1]).trim();
    if (!v) break;
    names.push({ row: r + 1, name: v });
  }
  const sum = {};
  const label = (r) => String(vals[r][nameCol - 1]).trim();
  for (let r = headerRow + names.length; r < vals.length; r++) {
    const v = label(r);
    if (v === '拼菜总数') sum.count = r + 1;
    else if (v === '取餐同工') sum.pickup = r + 1;
    else if (v === '订餐餐厅') sum.rest = r + 1;
    else if (v === '预订菜品' && !sum.dishes) sum.dishes = block(sh, r + 1, nameCol);
    else if (v === '自制佳肴' && !sum.bring) sum.bring = block(sh, r + 1, nameCol);
  }
  return { headerRow, nameCol, dates, names, sum };
}
function block(sh, row, col) {
  const m = sh.getRange(row, col).getMergedRanges();
  if (m.length) return { start: m[0].getRow(), end: m[0].getLastRow() };
  return { start: row, end: row };
}
function currentService(sh, lay) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  const cutoff = new Date(now.getTime() - 36 * 3600 * 1000); // 周五当天到周六中午前仍算本次
  const next = lay.dates.filter(d => d.date >= cutoff).sort((a, b) => a.date - b.date)[0];
  if (!next) throw new Error('表头里没有未来的服事日期');
  const deadline = new Date(next.date.getTime());
  deadline.setDate(deadline.getDate() - 1);
  deadline.setHours(DEADLINE_HOUR, 0, 0, 0);
  return { col: next.col, date: fmtDate(next.date), deadline, tz };
}
function toDate(v) {
  if (v instanceof Date) return v;
  const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]));
  return null;
}
function fmtDate(v) {
  const d = toDate(v); if (!d) return String(v);
  return Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

// ---------- 建表 ----------
function ensureTabs() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SHEET_DATA)) {
    const sh = ss.insertSheet(SHEET_DATA);
    sh.appendRow(['日期', '名字', '参与', '点菜编号', '自带', '已付款', '更新时间', '付款方式', '不挑']);
    sh.setFrozenRows(1);
  }
  const d = ss.getSheetByName(SHEET_DATA);
  if (!d.getRange(1, 8).getValue()) d.getRange(1, 8, 1, 2).setValues([['付款方式', '不挑']]);
  if (!ss.getSheetByName(SHEET_MENU)) {
    const sh = ss.insertSheet(SHEET_MENU);
    const rows = [['编号', '分类', '中文名', '英文名', '价格', '辣', '团餐推荐', '备注', '启用']];
    MENU_SEED.trim().split('\n').forEach(l => {
      const f = l.split('|');
      rows.push([f[0], f[1], f[2], f[3], f[4] === '' ? '' : Number(f[4]), f[5] || '', f[6] || '', f[7] || '', 'Y']);
    });
    sh.getRange(1, 1, rows.length, 9).setValues(rows);
    sh.setFrozenRows(1);
  }
}

// 首次运行时自动写入「菜单」标签页（香港园，$18 以下）。之后以标签页为准。
const MENU_SEED = `
S10|招牌菜|左宗鸡|General Tso's Chicken|16|Y|Y|
S11|招牌菜|芝麻鸡|Sesame Chicken|16||Y|
S13|招牌菜|宫保鸡|Kung Po Chicken|16|Y|Y|
S12|招牌菜|陈皮鸡|Chicken w. Orange Flavor|16|Y||
S14|招牌菜|鱼香鸡|Chicken w. Garlic Sauce|16|Y||
S15|招牌菜|腰果鸡|Chicken w. Cashew Nuts|16|||
S16|招牌菜|黑椒鸡|Chicken w. Black Pepper Sauce|16|||
S30|招牌菜|椒盐鱿鱼|Salt & Pepper Squid|17|Y|Y|
S17|招牌菜|番茄牛|Beef w. Tomatoes|17|||
S24|招牌菜|京都肉排|Pork Chop Beijing Style|17|||
S25|招牌菜|木须肉|Moo Shu Pork|16|||配4张薄饼
S26|招牌菜|木须鸡|Moo Shu Chicken|16|||配4张薄饼
S27|招牌菜|木须虾|Moo Shu Shrimp|16|||配4张薄饼
S28|招牌菜|豉椒鱿鱼|Squid w. Black Bean Sauce|17|||
45|鸡|蘑菇鸡片|Moo Goo Gai Pan|15|||
46|鸡|芥兰鸡|Chicken w. Broccoli|15|||
47|鸡|雪豆鸡|Chicken w. Snow Peas|15|||
48|鸡|咖喱鸡|Curry Chicken|15|Y||
49|鸡|豆豉鸡|Chicken w. Black Bean Sauce|15|||
50|鸡|豆腐鸡|Chicken w. Tofu|15|||
51|鸡|杂菜鸡|Chicken w. Mixed Vegetables|15|||
52|鸡|四季豆鸡|Chicken w. String Beans|15|||
53|鸡|甜酸鸡|Sweet & Sour Chicken|15|||
62|牛肉|青椒牛|Pepper Steak w. Onion|16|||
63|牛肉|杂菜牛|Beef w. Mixed Vegetables|16|||
64|牛肉|芥兰牛|Beef w. Broccoli|16||Y|
65|牛肉|蘑菇牛|Beef w. Mushroom|16|||
66|牛肉|豆腐牛|Beef w. Bean Curd|16|||
67|牛肉|咖喱牛|Curry Beef w. Onions|16|Y||
68|牛肉|雪豆牛|Beef w. Snow Peas|16|||
69|牛肉|四季豆牛|Beef w. String Beans|16|||
54|猪肉|叉烧杂菜|Roast Pork w. Mixed Vegetables|15|||
55|猪肉|叉烧芥兰|Roast Pork w. Broccoli|15|||
56|猪肉|叉烧芽菜|Roast Pork w. Bean Sprouts|15|||
57|猪肉|叉烧豆腐|Roast Pork w. Bean Curd|15|||
58|猪肉|叉烧蘑菇|Roast Pork w. Mushroom|15|||
59|猪肉|叉烧四季豆|Roast Pork w. String Beans|15|||
60|猪肉|回锅肉|Double Cooked Pork|16|||
61|猪肉|麻婆豆腐|Ma Po Tofu with Pork|16|Y|Y|
70|海鲜|虾龙糊|Shrimp w. Lobster Sauce|17||Y|
71|海鲜|杂菜虾|Shrimp w. Mixed Vegetables|17|||
72|海鲜|芥兰虾|Shrimp w. Broccoli|17|||
73|海鲜|豆豉虾|Shrimp w. Black Bean Sauce|17|||
74|海鲜|雪豆虾|Shrimp w. Snow Peas|17|||
75|海鲜|茄子虾|Shrimp w. Eggplant|17|||
76|海鲜|豆腐虾|Shrimp w. Bean Curd|17|||
78|海鲜|咖喱虾|Curry Shrimp|17|Y||
87|蔬菜|鱼香茄子|Eggplant w. Garlic Sauce|15|Y|Y|
83|蔬菜|干扁四季豆|Sauteed String Bean w. Pork|15||Y|
80|蔬菜|杂菜|Mixed Vegetables|13|||
81|蔬菜|茄子豆腐|Egg Plant Tofu in Garlic Sauce|15|Y||
82|蔬菜|鱼香芥兰|Broccoli w. Garlic Sauce|13|Y||
84|蔬菜|木须菜|Moo Shu Vegetable|15|||
86|蔬菜|左宗豆腐|General Tso's Fried Tofu|15|Y||
H1|健康清蒸|清蒸杂菜豆腐|Steamed Mixed Vegetables w. Bean Curd|14|||酱汁另放
H2|健康清蒸|清蒸鸡杂菜|Steamed Chicken w. Mixed Vegetables|15|||酱汁另放
H3|健康清蒸|清蒸芥兰鸡|Steamed Chicken w. Broccoli|15|||酱汁另放
H4|健康清蒸|清蒸大虾杂菜|Steamed Jumbo Shrimp w. Mixed Vegetables|17|||酱汁另放
27|粉面|星洲米粉|Singapore Rice Noodle|16|Y|Y|
28|粉面|干炒牛河|Beef Chow Fun|16||Y|
26|粉面|本楼炒米粉|House Special Rice Noodle|16|||
29|粉面|豉汁牛河|Beef Chow Fun w. Black Bean Sauce|17|||
30|粉面|广东炒面|Cantonese Fried Noodles|17|||
31|粉面|泰式鸡炒粉|Pad Thai (Chicken)|16|||
32|粉面|泰式菜炒粉|Pad Thai (Vegetable)|16|||
38|捞面|本楼捞面|House Special Lo Mein|15||Y|
33|捞面|素菜捞面|Vegetable Lo Mein|12|||
34|捞面|鸡捞面|Chicken Lo Mein|13|||
35|捞面|叉烧捞面|Roast Pork Lo Mein|13|||
36|捞面|虾捞面|Shrimp Lo Mein|15|||
37|捞面|牛肉捞面|Beef Lo Mein|15|||
24|炒饭|本楼炒饭|House Special Fried Rice|15||Y|
25|炒饭|扬州炒饭|Young Chow Fried Rice|15|||
18|炒饭|净炒饭|Plain Fried Rice w. Egg|6.5|||
19|炒饭|叉烧炒饭|Roast Pork Fried Rice|12|||
20|炒饭|鸡炒饭|Chicken Fried Rice|12|||
21|炒饭|素菜炒饭|Vegetable Fried Rice|12|||
22|炒饭|虾炒饭|Shrimp Fried Rice|14|||
23|炒饭|牛肉炒饭|Beef Fried Rice|14|||
N1|汤面|牛腩汤面|Beef Stew Noodle Soup|13|||
N2|汤面|云吞汤面|Wonton Noodle Soup|13|||
N3|汤面|鲜虾汤面|Shrimp Noodle Soup|13|||
N4|汤面|海鲜汤面|Seafood Noodle Soup|15|||
39|芙蓉蛋/炒面|肉芙蓉蛋|Pork Egg Foo Young|14|||
40|芙蓉蛋/炒面|鸡芙蓉蛋|Chicken Egg Foo Young|14|||
41|芙蓉蛋/炒面|牛肉芙蓉蛋|Beef Egg Foo Young|14|||
42|芙蓉蛋/炒面|虾芙蓉蛋|Shrimp Egg Foo Young|14|||
43|芙蓉蛋/炒面|肉或鸡炒面|Pork or Chicken Chow Mein|14|||
44|芙蓉蛋/炒面|虾炒面|Shrimp Chow Mein|14|||
16|小吃|排骨|Bar-B-Q Spare Ribs (5)|12||Y|5条
17|小吃|无骨排|Boneless Ribs|12|||
15|小吃|炸鸡翼|Fried Chicken Wings (8)|10|||8个
14|小吃|锅贴或水饺|Steamed or Fried Dumplings (6)|8|||6个
13|小吃|炸云吞|Fried Wonton (10)|6.5|||10个
12|小吃|虾吐司|Shrimp Toast|6|||
10|小吃|春卷|Egg Roll|2.75|||单个
11|小吃|菜卷|Spring Roll|2.25|||单个
4|汤|酸辣汤|Hot & Sour Soup|6|Y||
1|汤|云吞汤|Wonton Soup|6|||
2|汤|蛋花汤|Egg Drop Soup|5|||
3|汤|云吞蛋花汤|Wonton Egg Drop Soup|6|||
5|汤|瓜菜汤|Vegetable Soup|5.5|||
6|汤|八宝汤|Eight Ingredients Soup|12|||
7|汤|榨菜肉丝汤|Shredded Pork & Pickled Cabbage Soup|12|Y||
8|汤|鸡面汤|Chicken Noodle Soup|10|||
9|汤|鸡蓉粟米汤|Chicken & Egg Corn Soup|10|||
`;
