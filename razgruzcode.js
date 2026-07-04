/* OZON FBO SUPPLY AUTOMATION — CENTRAL CODE v2.9 — 04.07.2026 */
/* Умная обработка 429 (пауза и повтор вместо остановки), пауза между заявками
   настраивается свободно. Ручные правки в «Разгрузе» и «Своде». Код обезличен. */

/**
 * Разгруз Ozon v2 — обновление прямо из Google Таблицы (Apps Script).
 * Кнопка/меню «Обновить» → скрипт идёт в Ozon API, тянет «Управление остатками»,
 * пересобирает листы Разгруз / Свод по товару / Исключения / Короба / Методика.
 * Введённое «Доступно», отметки «Исключений» и заполненные «Короба» сохраняются.
 *
 * v2:
 *  - МУЛЬТИБРЕНД: константа BRAND + встроенные карты коробов по брендам +
 *    лист «Короба» (артикул → шт/кор), который правится прямо в таблице.
 *  - МИНИ-ТАБЛИЦА: меню «Сформировать таблицу поставок» — отдельная Google-таблица,
 *    лист на товар (Направление / Кол-во шт / Кол-во коробов + итог),
 *    как ручной файл «Разгруз <бренд> <дата>».
 *  - Модуль автосоздания поставок — в отдельном файле postavki_appscript.gs
 *    (вставляется в этот же проект Apps Script вторым файлом).
 *
 * Ключи Ozon бренда — в Project Settings → Script Properties:
 *   OZON_CLIENT_ID, OZON_API_KEY
 * Одна Google-таблица = один бренд (ключи у каждой таблицы свои).
 */

// BRAND задаётся в загрузчике таблицы или на листе «Ключи» (B1)
var TARGET_FOLDER_ID = "1tFeJIfCQRgB29cF5HJSRiY-WZWWnIyj2";
var HORIZON = 60;   // горизонт потребности, дней
var PRI = 30;       // порог приоритета (для подсветки «Остаток дней»)

/** Карты коробов, зашитые в код (бренд → { артикул: шт в коробке }).
 *  В публичной версии пусто: фасовки задаются на листе «Короба» каждой таблицы
 *  (приоритет: лист «Короба» > эта карта > авто-правило по фасовке из артикула). */
var BOX_OVERRIDE_BY_BRAND = {};

var RAZ = 'Разгруз', SVOD = 'Свод по товару', EXC = 'Исключения', MET = 'Методика', KOR = 'Короба';

/** Меню при открытии таблицы. */
function onOpen() {
  var m = SpreadsheetApp.getUi().createMenu('Разгруз')
    .addItem('🔄 Обновить из Ozon', 'refreshRazgruz')
    .addItem('📦 Сформировать таблицу поставок', 'buildSupplyTable');
  try {
    if (typeof addPostavkiMenu_ === 'function') addPostavkiMenu_(m); // из postavki_appscript.gs
  } catch (e) {}
  m.addToUi();
}

/** Главная функция — её вешаем на кнопку и в меню. */
function refreshRazgruz() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Тяну остатки из Ozon…', 'Разгруз ' + brand_(), 60);
  var creds = ozonCreds_();
  if (!creds) return;
  try { ss.setSpreadsheetLocale('en_US'); } catch (e) {}
  var rows = fetchFromApi_(creds.cid, creds.key);
  var keep = readExisting_(ss);
  buildSheets_(ss, rows, keep.avail, keep.closed, keep.boxes, keep.over, keep.razMan);
  ss.toast('Готово: обновлено ' + rows.filter(function (r) { return r.need > 0; }).length + ' строк.', 'Разгруз', 6);
}

var KEYS = 'Ключи';

/** Лист «Ключи» (скрытый): B1 бренд, B2 Client-Id, B3 Api-Key.
 *  Позволяет настраивать копию таблицы без захода в Apps Script. */
function ensureKeysSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(KEYS);
  if (!s) {
    s = ss.insertSheet(KEYS);
    s.getRange(1, 1, 4, 2).setValues([
      ['Бренд', ''],
      ['OZON_CLIENT_ID', ''],
      ['OZON_API_KEY', ''],
      ['После заполнения лист можно скрыть (ПКМ по ярлыку → Скрыть)', '']
    ]);
    s.getRange(1, 2, 3, 1).setBackground('#FFF2CC');
    s.setColumnWidth(1, 200); s.setColumnWidth(2, 360);
  }
  return s;
}

/** Бренд: лист «Ключи» (B1) > переменная BRAND загрузчика/скрипта. */
function brand_() {
  try {
    var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYS);
    if (s) { var b = String(s.getRange('B1').getValue()).trim(); if (b) return b; }
  } catch (e) {}
  return (typeof BRAND !== 'undefined' && BRAND) ? BRAND : '(бренд не задан)';
}

/** Ключи: Script Properties > лист «Ключи» (B2/B3). */
function ozonCreds_() {
  var props = PropertiesService.getScriptProperties();
  var cid = props.getProperty('OZON_CLIENT_ID');
  var key = props.getProperty('OZON_API_KEY');
  if (!cid || !key) {
    var s = ensureKeysSheet_();
    cid = String(s.getRange('B2').getValue()).trim();
    key = String(s.getRange('B3').getValue()).trim();
  }
  if (!cid || !key) {
    try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYS).showSheet(); } catch (e) {}
    SpreadsheetApp.getUi().alert('Заполните лист «Ключи»: B1 — бренд, B2 — Client-Id, B3 — Api-Key.\n' +
      '(Либо задайте OZON_CLIENT_ID / OZON_API_KEY в Script Properties.)');
    return null;
  }
  return { cid: cid, key: key };
}

/* ---------------- короб: лист «Короба» > карта бренда > фасовка из артикула ---------------- */
function isSet_(a) { return String(a).toLowerCase().indexOf('set') >= 0; }

function boxSize_(art, boxMap) {
  if (boxMap && boxMap[art]) return boxMap[art];
  var ov = BOX_OVERRIDE_BY_BRAND[String(brand_()).toUpperCase()] || {};
  if (ov[art]) return ov[art];
  var n = String(art).toLowerCase();
  var m = n.match(/(\d+)\s*caps/);
  if (m) { var c = +m[1]; if (c === 60 || c === 90 || c === 120) return 60; if (c === 180 || c === 240) return 30; if (c === 500) return 12; return null; }
  m = n.match(/(\d+)\s*g\b/);
  if (m) { var g = +m[1]; if (g === 200 || g === 300) return 36; if (g === 500) return 18; }
  return null;
}

/** Читает лист «Короба» → { артикул: шт_в_коробе }. */
function readBoxSheet_(ss) {
  var map = {};
  var s = ss.getSheetByName(KOR);
  if (s && s.getLastRow() > 1) {
    var d = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < d.length; i++) {
      var art = d[i][0], v = d[i][1];
      if (art && typeof v === 'number' && v > 0) map[art] = v;
    }
  }
  return map;
}

/* ---------------- запросы к Ozon ---------------- */
function ozon_(url, body, cid, key) {
  for (var t = 0; t < 4; t++) {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json',
      headers: { 'Client-Id': cid, 'Api-Key': key },
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 200) return JSON.parse(resp.getContentText());
    if (code === 429 || code >= 500) { Utilities.sleep(1500 * (t + 1)); continue; }
    throw new Error('Ozon ' + url + ' -> ' + code + ': ' + resp.getContentText().slice(0, 200));
  }
  throw new Error('Ozon ' + url + ' — не удалось после ретраев');
}

function fetchFromApi_(cid, key) {
  // 1) список товаров
  var pids = [], lastId = '';
  while (true) {
    var r = ozon_('https://api-seller.ozon.ru/v3/product/list',
      { filter: { visibility: 'ALL' }, limit: 1000, last_id: lastId }, cid, key);
    var items = (r.result && r.result.items) || [];
    for (var i = 0; i < items.length; i++) pids.push(items[i].product_id);
    lastId = (r.result && r.result.last_id) || '';
    if (!lastId || !items.length) break;
  }
  // 2) sku-номера (батч 100)
  var skus = [];
  for (var a = 0; a < pids.length; a += 100) {
    var inf = ozon_('https://api-seller.ozon.ru/v3/product/info/list',
      { product_id: pids.slice(a, a + 100) }, cid, key);
    var its = inf.items || [];
    for (var j = 0; j < its.length; j++) if (its[j].sku) skus.push(its[j].sku);
  }
  // 3) остатки по складам (батч 50) -> агрегация по (offer_id, кластер)
  var agg = {};
  for (var b = 0; b < skus.length; b += 50) {
    var rs = ozon_('https://api-seller.ozon.ru/v1/analytics/stocks',
      { skus: skus.slice(b, b + 50).map(function (s) { return Number(s); }) }, cid, key);
    var arr = rs.items || [];
    for (var k = 0; k < arr.length; k++) {
      var x = arr[k], art = x.offer_id;
      if (art == null || isSet_(art)) continue;
      var keyk = art + '||' + x.cluster_name;
      if (!agg[keyk]) agg[keyk] = { art: art, name: x.name, sku: x.sku, cluster: x.cluster_name, ssd: 0, ost: 0 };
      agg[keyk].ssd = x.ads_cluster || 0;
      agg[keyk].ost += (x.available_stock_count || 0) + (x.valid_stock_count || 0) +
        (x.requested_stock_count || 0) + (x.transit_stock_count || 0) + (x.return_from_customer_stock_count || 0);
    }
  }
  var rows = [];
  for (var kk in agg) {
    var v = agg[kk];
    rows.push({ art: v.art, name: v.name, sku: v.sku, cluster: v.cluster, ssd: v.ssd, ost: v.ost,
                need: Math.round(v.ssd * HORIZON - v.ost) });
  }
  return rows;
}

/* ---------------- сохранить введённое «Доступно», «Исключения», «Короба» ---------------- */
function readExisting_(ss) {
  var avail = {}, closed = {};
  var razMan = {};
  var rz0 = ss.getSheetByName(RAZ);
  if (rz0 && rz0.getLastRow() > 1 && rz0.getLastColumn() >= 22) {
    var dm = rz0.getRange(2, 1, rz0.getLastRow() - 1, 22).getValues();
    for (var q = 0; q < dm.length; q++) {
      var vv = dm[q][21];
      if (dm[q][0] && dm[q][2] && typeof vv === 'number' && vv > 0)
        razMan[String(dm[q][0]) + '||' + String(dm[q][2])] = vv;
    }
  }
  var over = {};
  var sv = ss.getSheetByName(SVOD);
  if (sv && sv.getLastRow() > 1) {
    var width = Math.min(Math.max(sv.getLastColumn(), 7), 10);
    var d = sv.getRange(2, 1, sv.getLastRow() - 1, width).getValues();
    for (var i = 0; i < d.length; i++) {
      var art = d[i][0], val = d[i][6];
      if (art && typeof val === 'number' && val > 0) avail[art] = val;
      var ov = d[i].length > 9 ? d[i][9] : '';
      if (art && typeof ov === 'number' && ov > 0) over[art] = ov;
    }
  }
  var ex = ss.getSheetByName(EXC);
  if (ex && ex.getLastRow() > 1) {
    var e = ex.getRange(2, 1, ex.getLastRow() - 1, 2).getValues();
    for (var j = 0; j < e.length; j++)
      if (e[j][0] && String(e[j][1]).trim().toUpperCase() === 'ДА') closed[e[j][0]] = true;
  }
  return { avail: avail, closed: closed, boxes: readBoxSheet_(ss), over: over, razMan: razMan };
}

/* ---------------- сборка листов ---------------- */
function sheet_(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  s.clear();
  s.clearConditionalFormatRules();
  return s;
}

function buildSheets_(ss, rows, avail, closed, boxMap, over, razMan) {
  over = over || {}; razMan = razMan || {};
  // строки с потребностью, сортировка по (артикул, -потребность, sku)
  var rs = rows.filter(function (r) { return r.need > 0; });
  rs.sort(function (p, q) {
    if (String(p.art) !== String(q.art)) return String(p.art) < String(q.art) ? -1 : 1;
    if (q.need !== p.need) return q.need - p.need;
    return String(p.sku) < String(q.sku) ? -1 : 1;
  });
  var last = rs.length + 1;

  // ----- Разгруз -----
  var head = ['Артикул','SKU','Кластер','ССП/28дн','Остаток (всего)','Остаток дней','Потребность 60дн',
              'К отгрузке (банки)','К отгрузке (короба)','Остаток дней после поставки',
              '','_tier','_box','_avail','_needOpen','_p','_manual','_ideal','_okey','_cum','_excl',
              'К отгрузке (ручное)'];
  var rz = [head];
  for (var i = 0; i < rs.length; i++) {
    var r = i + 2, x = rs[i];
    rz.push([
      x.art, x.sku, x.cluster, Math.round(x.ssd * 10) / 10, Math.round(x.ost),
      '=IF(D'+r+'=0,"",ROUND(E'+r+'/D'+r+',0))', x.need,
      '=IF(V'+r+'>0,IF(M'+r+'=0,V'+r+',M'+r+'*ROUNDUP(V'+r+'/M'+r+',0)),' +
        'IF(OR(N'+r+'=0,N'+r+'="",M'+r+'=0),0,M'+r+'*(INT(T'+r+'/M'+r+')-INT((T'+r+'-R'+r+')/M'+r+'))))',
      '=IF(M'+r+'=0,"",H'+r+'/M'+r+')',
      '=IF(D'+r+'=0,"",ROUND((E'+r+'+H'+r+')/D'+r+',0))',
      '',
      '=IF(F'+r+'="",1,IF(F'+r+'<'+PRI+',0,1))',
      "=SUMIF('"+SVOD+"'!$A:$A,A"+r+",'"+SVOD+"'!$B:$B)",
      "=MAX(0,IF(Q"+r+">0,Q"+r+",SUMIF('"+SVOD+"'!$A:$A,A"+r+",'"+SVOD+"'!$G:$G))" +
        "-SUMIFS($V$2:$V$"+last+",$A$2:$A$"+last+",A"+r+"))",
      '=SUMIFS($G$2:$G$'+last+',$A$2:$A$'+last+',A'+r+',$U$2:$U$'+last+',0)',
      '', "=SUMIF('"+SVOD+"'!$A:$A,A"+r+",'"+SVOD+"'!$J:$J)",
      '=IF(U'+r+'=1,0,IF(O'+r+'=0,0,IF(Q'+r+'>0,N'+r+'*G'+r+'/O'+r+',G'+r+'*MIN(1,N'+r+'/O'+r+'))))',
      i + 1,
      '=SUMIFS($R$2:$R$'+last+',$A$2:$A$'+last+',A'+r+',$S$2:$S$'+last+',"<="&S'+r+')',
      '=IF(OR(IFERROR(VLOOKUP(C'+r+","+EXC+'!$A:$B,2,FALSE),"")="ДА",V'+r+'>0),1,0)',
      (razMan[String(x.art) + '||' + String(x.cluster)] || '')
    ]);
  }
  var sRz = sheet_(ss, RAZ);
  sRz.getRange(1, 1, rz.length, 22).setValues(rz);

  // ----- Свод -----
  var aggm = {}, order = [];
  for (var t2 = 0; t2 < rows.length; t2++) {
    var y = rows[t2];
    if (!aggm[y.art]) { aggm[y.art] = { need: 0, ost: 0, ssd: 0 }; order.push(y.art); }
    aggm[y.art].need += Math.max(0, y.need); aggm[y.art].ost += y.ost; aggm[y.art].ssd += y.ssd;
  }
  order.sort(function (p, q) { return aggm[q].need - aggm[p].need; });
  var noBox = [];
  var sv = [['Артикул','Короб, шт/кор','Остаток всего','ССП всего','Остаток дней','Потребность 60дн',
             'Доступно к поставке','Покрытие %','Распределено','Отгрузить всего (ручное)']];
  for (var o = 0; o < order.length; o++) {
    var art = order[o], a = aggm[art], r2 = o + 2, bx = boxSize_(art, boxMap), av = avail[art];
    if (!bx) noBox.push(art);
    var ovv = over[art];
    sv.push([art, (bx ? bx : ''), Math.round(a.ost), Math.round(a.ssd * 10) / 10,
      '=IF(D'+r2+'=0,"",ROUND(C'+r2+'/D'+r2+',0))', Math.round(a.need),
      (typeof av === 'number' ? av : ''),
      '=IF(F'+r2+'=0,"",IF(G'+r2+'="","",MIN(1,G'+r2+'/F'+r2+')))',
      "=SUMIF('"+RAZ+"'!$A:$A,A"+r2+",'"+RAZ+"'!$H:$H)",
      (typeof ovv === 'number' && ovv > 0 ? ovv : '')]);
  }
  var sSv = sheet_(ss, SVOD);
  sSv.getRange(1, 1, sv.length, 10).setValues(sv);

  // ----- Исключения -----
  var clset = {}, clusters = [];
  for (var c2 = 0; c2 < rows.length; c2++) if (rows[c2].cluster && !clset[rows[c2].cluster]) { clset[rows[c2].cluster] = 1; clusters.push(rows[c2].cluster); }
  clusters.sort();
  var ex = [['Кластер', 'Закрыт (ДА = не грузить)']];
  for (var e2 = 0; e2 < clusters.length; e2++) ex.push([clusters[e2], closed[clusters[e2]] ? 'ДА' : '']);
  var sEx = sheet_(ss, EXC);
  sEx.getRange(1, 1, ex.length, 2).setValues(ex);

  // ----- Короба (артикул → шт/кор; ручное заполнение имеет высший приоритет) -----
  var kor = [['Артикул', 'Шт в коробе (ручной ввод)', 'Определено авто-правилом']];
  var allArts = order.slice();
  for (var k2 = 0; k2 < allArts.length; k2++) {
    var art2 = allArts[k2];
    var manual = boxMap[art2] || '';
    var auto = boxSize_(art2, null) || '';
    kor.push([art2, manual, auto]);
  }
  var sKor = sheet_(ss, KOR);
  sKor.getRange(1, 1, kor.length, 3).setValues(kor);

  // ----- Методика -----
  var met = [['Разгруз Ozon — методика (бренд: ' + brand_() + ')'], [''],
    ['Потребность = ССП(28дн) × ' + HORIZON + ' − Остаток (Доступно+Готовим+Заявки+В пути+Возвраты)'],
    ['Остаток дней = Остаток ÷ ССП. Подсветка: красный мало дней → зелёный много.'],
    ['Короб: лист «Короба» (ручной) > карта бренда в скрипте > фасовка из артикула'],
    ['  (60/90/120 caps→60 | 180/240→30 | 200/300г→36 | 500г→18).'],
    ['Распределение: пропорционально потребности по ВСЕМ открытым кластерам, кратно коробам;'],
    ['«Отгрузить всего (ручное)» (Свод, кол. J): принудительный объём по товару —'],
    ['распределяется пропорционально потребности даже СВЕРХ неё (загрузка впрок);'],
    ['имеет приоритет над «Доступно к поставке».'],
    ['«К отгрузке (ручное)» (Разгруз, после кол. J): жёсткий объём по конкретному'],
    ['кластеру (кратно коробу вверх); такой кластер выпадает из пропорций, его объём'],
    ['вычитается из общего, остальные делят остаток. Кластер без строки в разгрузе —'],
    ['только через лист «Корректировки».'],
    ['кумулятивное округление — «Распределено» ≤ «Доступно». Дефицитные получают больше (больше потребность).'],
    ['Исключения: «ДА» у кластера — он выпадает из распределения.'],
    ['«Доступно к поставке» вписывается в листе «Свод» (колонка G). Наборы SET исключены.'],
    ['Мини-таблица: «Разгруз → Сформировать таблицу поставок» — отдельный файл, лист на товар.'],
    ['Источник: Ozon API analytics/stocks. Обновление — кнопкой «Разгруз → Обновить из Ozon».']];
  var sMet = sheet_(ss, MET);
  sMet.getRange(1, 1, met.length, 1).setValues(met);

  formatSheets_(sRz, sSv, sEx, sKor, last, sv.length, clusters.length, kor.length);

  if (noBox.length) {
    ss.toast('Без короба: ' + noBox.length + ' артикулов — заполните лист «Короба».', 'Разгруз', 8);
  }
}

/* ---------------- оформление ---------------- */
function gradRule_(rng) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#E67C73', SpreadsheetApp.InterpolationType.NUMBER, '14')
    .setGradientMidpointWithValue('#FFD966', SpreadsheetApp.InterpolationType.NUMBER, '21')
    .setGradientMaxpointWithValue('#93C47D', SpreadsheetApp.InterpolationType.NUMBER, '30')
    .setRanges([rng]).build();
}

function formatSheets_(sRz, sSv, sEx, sKor, last, svRows, nCl, korRows) {
  // --- Разгруз ---
  sRz.setFrozenRows(1);
  sRz.getRange(1, 1, 1, 10).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  var wRz = [300, 90, 180, 75, 95, 85, 95, 95, 95, 110];
  for (var i = 0; i < wRz.length; i++) sRz.setColumnWidth(i + 1, wRz[i]);
  sRz.hideColumns(11, 11); // K..U (V остаётся видимой — визуально сразу после J)
  sRz.setColumnWidth(22, 120);
  sRz.getRange(1, 22).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  if (last > 1) {
    sRz.getRange(2, 7, last - 1, 1).setBackground('#FFF2CC');         // Потребность — жёлтая
    sRz.getRange(2, 4, last - 1, 1).setNumberFormat('0.0');           // ССП
    sRz.getRange(2, 22, last - 1, 1).setBackground('#FFF2CC');        // К отгрузке (ручное)
  }
  var rulesRz = [
    gradRule_(sRz.getRange('F2:F')),   // Остаток дней
    gradRule_(sRz.getRange('J2:J')),   // Остаток дней после поставки
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#B6D7A8').setRanges([sRz.getRange('H2:H')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#B6D7A8').setRanges([sRz.getRange('I2:I')]).build()
  ];
  sRz.setConditionalFormatRules(rulesRz);

  // --- Свод ---
  sSv.setFrozenRows(1);
  sSv.getRange(1, 1, 1, 10).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  var wSv = [300, 95, 95, 90, 90, 110, 120, 80, 95, 130];
  for (var j = 0; j < wSv.length; j++) sSv.setColumnWidth(j + 1, wSv[j]);
  if (svRows > 1) {
    sSv.getRange(2, 7, svRows - 1, 1).setBackground('#FFF2CC');       // Доступно — жёлтая
    sSv.getRange(2, 8, svRows - 1, 1).setNumberFormat('0%');          // Покрытие
    sSv.getRange(2, 10, svRows - 1, 1).setBackground('#FFF2CC');      // Отгрузить всего (ручное)
  }
  sSv.setConditionalFormatRules([ gradRule_(sSv.getRange('E2:E')) ]); // Остаток дней
  sSv.getRange(2, 1, Math.max(svRows - 1, 1), 1).setHorizontalAlignment('left');

  // --- Исключения ---
  sEx.setFrozenRows(1);
  sEx.getRange(1, 1, 1, 2).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sEx.setColumnWidth(1, 230); sEx.setColumnWidth(2, 190);
  if (nCl > 0) {
    var rng = sEx.getRange(2, 2, nCl, 1);
    rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ДА'], true).setAllowInvalid(true).build());
    rng.setBackground('#FFF2CC');
  }

  // --- Короба ---
  sKor.setFrozenRows(1);
  sKor.getRange(1, 1, 1, 3).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sKor.setColumnWidth(1, 300); sKor.setColumnWidth(2, 180); sKor.setColumnWidth(3, 180);
  if (korRows > 1) {
    sKor.getRange(2, 2, korRows - 1, 1).setBackground('#FFF2CC'); // ручной ввод — жёлтый
    // подсветить строки, где нет ни ручного, ни авто-короба
    var noneRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A2<>"",$B2="",$C2="")')
      .setBackground('#F4CCCC')
      .setRanges([sKor.getRange(2, 1, korRows - 1, 3)]).build();
    sKor.setConditionalFormatRules([noneRule]);
  }
}

/* ================================================================
 *  МИНИ-ТАБЛИЦА ПОСТАВОК («Разгруз <бренд> <дата>»)
 *  Лист на товар: Направление / Кол-во шт / Кол-во коробов (+ итог).
 *  Плюс лист «Свод по кластерам» — состав будущих заявок.
 * ================================================================ */

/** Снимок распределения из листа «Разгруз»: [{art, sku, cluster, units, boxes, box}] где units>0. */
function collectAllocation_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(RAZ);
  if (!s || s.getLastRow() < 2) throw new Error('Лист «Разгруз» пуст — сначала «Обновить из Ozon».');
  var d = s.getRange(2, 1, s.getLastRow() - 1, 13).getValues(); // A..M (M = _box)
  var out = [];
  for (var i = 0; i < d.length; i++) {
    var units = Number(d[i][7]) || 0; // H: К отгрузке (банки)
    if (units <= 0) continue;
    var box = Number(d[i][12]) || 0;  // M: _box
    out.push({
      art: d[i][0], sku: d[i][1], cluster: d[i][2],
      units: units, box: box, boxes: box > 0 ? Math.round(units / box) : ''
    });
  }
  if (!out.length) throw new Error('Нет распределённых количеств: заполните «Доступно к поставке» на листе «Свод».');
  return out;
}

/** Меню: «Сформировать таблицу поставок». */
function buildSupplyTable() {
  var ui = SpreadsheetApp.getUi();
  var alloc = collectAllocation_();

  // группировка по товару
  var byArt = {}, artOrder = [];
  alloc.forEach(function (r) {
    if (!byArt[r.art]) { byArt[r.art] = []; artOrder.push(r.art); }
    byArt[r.art].push(r);
  });

  var dd = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Moscow', 'dd.MM');
  var name = 'Разгруз ' + brand_() + ' ' + dd;
  var out = SpreadsheetApp.create(name);

  // лист на товар
  artOrder.forEach(function (art, idx) {
    var rows = byArt[art].slice().sort(function (a, b) { return b.units - a.units; });
    var sh = idx === 0 ? out.getSheets()[0] : out.insertSheet();
    sh.setName(sheetNameSafe_(art, idx));
    var data = [['Направление', 'Кол-во шт', 'Кол-во коробов']];
    rows.forEach(function (r) { data.push([r.cluster, r.units, r.boxes]); });
    sh.getRange(1, 1, data.length, 3).setValues(data);
    var lastRow = data.length;
    // пустая строка + итог (формулами, как в ручном файле)
    sh.getRange(lastRow + 2, 2).setFormula('=SUM(B2:B' + lastRow + ')');
    sh.getRange(lastRow + 2, 3).setFormula('=SUM(C2:C' + lastRow + ')');
    sh.getRange(lastRow + 2, 2, 1, 2).setFontWeight('bold');
    sh.getRange(1, 1, 1, 3).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
      .setHorizontalAlignment('center');
    sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 130);
    sh.setFrozenRows(1);
  });

  // свод по кластерам (состав будущих заявок)
  var byCl = {}, clOrder = [];
  alloc.forEach(function (r) {
    if (!byCl[r.cluster]) { byCl[r.cluster] = []; clOrder.push(r.cluster); }
    byCl[r.cluster].push(r);
  });
  clOrder.sort();
  var sw = out.insertSheet('Свод по кластерам');
  var rows2 = [['Кластер', 'Артикул', 'SKU', 'Кол-во шт', 'Кол-во коробов']];
  clOrder.forEach(function (cl) {
    byCl[cl].sort(function (a, b) { return b.units - a.units; })
      .forEach(function (r) { rows2.push([cl, r.art, r.sku, r.units, r.boxes]); });
  });
  sw.getRange(1, 1, rows2.length, 5).setValues(rows2);
  sw.getRange(1, 1, 1, 5).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sw.setColumnWidth(1, 240); sw.setColumnWidth(2, 300); sw.setColumnWidth(3, 100);
  sw.setFrozenRows(1);

  // положить файл рядом с текущей таблицей (в ту же папку Drive)
  labelsFolder_().addFile(out);
   function labelsFolder_() {
  return DriveApp.getFolderById(TARGET_FOLDER_ID);
}

  ui.alert('Таблица поставок готова', name + '\n\n' + out.getUrl(), ui.ButtonSet.OK);
}

/** Имя листа: без запрещённых символов, ≤ 90 знаков, уникальность по индексу. */
function sheetNameSafe_(art, idx) {
  var n = String(art).replace(/[\[\]\*\/\\\?:]/g, ' ').replace(/\s+/g, ' ').trim();
  if (n.length > 90) n = n.slice(0, 90);
  return n || ('Товар ' + (idx + 1));
}


/* ==================== МОДУЛЬ ПОСТАВОК ==================== */

/**
 * Поставки Ozon FBO — автосоздание заявок из разгруза (Apps Script, файл 2/2).
 * Вставляется В ТОТ ЖЕ проект Apps Script, что и razgruz_appscript.gs
 * (Файл → + → Скрипт → назвать «postavki» → вставить этот код).
 *
 * Что делает:
 *   1. «Сформировать план поставок» — лист «План поставок»: строка = кластер,
 *      состав (SKU × шт) берётся из текущего распределения листа «Разгруз».
 *   2. «Создать заявки в Ozon» — для каждого отмеченного кластера:
 *      черновик → расчёт → склад → таймслот → заявка на поставку. Результат —
 *      в лист «Поставки» (номер заявки, склад, слот, статус).
 *   3. «Передать грузоместа + этикетки» — по строке листа «Поставки»: короба
 *      (каждый короб = фасовка одного SKU) + PDF-этикетки в папку Drive.
 *   4. «Обновить статусы», «Склады отгрузки», «Диагностика API».
 *
 * Эндпоинты — актуальные на июнь 2026 (старые /v1/draft/* отключены 16.03.2026):
 *   /v1/draft/direct/create | /v1/draft/crossdock/create   (черновик)
 *   /v2/draft/create/info                                   (расчёт черновика)
 *   /v2/draft/timeslot/info                                 (таймслоты)
 *   /v2/draft/supply/create , /v2/draft/supply/create/status(заявка)
 *   /v3//v2/supply-order/get , /v2/supply-order/list        (статусы)
 *   /v2//v1/cluster/list , /v1/warehouse/fbo/seller/list    (справочники)
 *   /v1/cargoes/create(/info) , /v1/cargoes-label/*         (грузоместа, этикетки)
 * Все запросы идут через ozonTry_() со списком путей: если Ozon переедет на новую
 * версию метода, добавьте путь в CFG_EP — больше ничего менять не нужно.
 *
 * ЛИМИТЫ OZON: для актуальных draft-методов опубликованных лимитов нет (общий —
 * 50 rps на аккаунт), но 429 на практике случаются. Скрипт держит паузу между
 * заявками (настройка), при 429 ждёт и повторяет, а при устойчивом лимите
 * останавливается с понятным статусом. Продолжение — повторным запуском или
 * авто-триггером (лимит Apps Script ~6 минут на запуск).
 */

var NAST = 'Настройки поставок', PLAN = 'План поставок', REESTR = 'Поставки', KORR = 'Корректировки';

/* Пути методов: первый — основной, дальше — запасные (на случай смены версии). */
var CFG_EP = {
  draftDirect:    ['/v1/draft/direct/create'],
  draftCrossdock: ['/v1/draft/crossdock/create'],
  draftInfo:      ['/v2/draft/create/info'],
  timeslotInfo:   ['/v2/draft/timeslot/info'],
  supplyCreate:   ['/v2/draft/supply/create'],
  supplyStatus:   ['/v2/draft/supply/create/status'],
  orderGet:       ['/v3/supply-order/get', '/v2/supply-order/get'],
  orderList:      ['/v2/supply-order/list'],
  clusterList:    ['/v2/cluster/list', '/v1/cluster/list'],
  fboWarehouses:  ['/v1/warehouse/fbo/seller/list', '/v1/warehouse/fbo/list'],
  dropoffPoints:  ['/v1/warehouse/fbo/list'], // точки отгрузки Ozon (хабы/СЦ); seller/list — это СВОИ склады
  cargoesCreate:  ['/v1/cargoes/create'],
  cargoesInfo:    ['/v1/cargoes/create/info', '/v1/cargoes/create/status'],
  cargoesGet:     ['/v1/cargoes/get'],
  labelCreate:    ['/v1/cargoes-label/create'],
  labelGet:       ['/v1/cargoes-label/get']
};
var API = 'https://api-seller.ozon.ru';

/** Пункты меню (вызывается из onOpen в razgruz_appscript.gs). */
function addPostavkiMenu_(m) {
  m.addSeparator()
   .addItem('🧾 1. Сформировать план поставок', 'buildSupplyPlan')
   .addItem('🚚 2. Создать заявки в Ozon', 'createSupplies')
   .addItem('📦 3. Грузоместа + этикетки (строка «Поставки»)', 'sendCargoesForActiveRow')
   .addItem('🏷 Скачать этикетки (строка «Поставки»)', 'downloadLabelsForActiveRow')
   .addItem('🧩 Объединить PDF этикеток за сегодня', 'mergeTodayLabelFolders')
   .addSeparator()
   .addItem('🔁 Обновить статусы заявок', 'updateSupplyStatuses')
   .addItem('🎯 Точка отгрузки из последних заявок', 'dropoffFromRecentOrders')
   .addItem('🗺 Кластеры (справочник)', 'showClusters')
   .addItem('🏭 Склады отгрузки (справочник)', 'showDropoffWarehouses')
   .addItem('🩺 Диагностика API поставок', 'supplyDiagnostics');
}

/* ---------------- утилиты ---------------- */
function toastSafe_(msg, title) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, title || 'Поставки', 8); } catch (e) {}
}
function alertSafe_(title, msg) {
  try { SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { toastSafe_(title + ': ' + msg); }
}
function nowIso_(d) { return Utilities.formatDate(d, 'GMT', "yyyy-MM-dd'T'HH:mm:ss'Z'"); }
function normName_(s) {
  return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim();
}

/** POST с перебором путей-кандидатов (404/«unknown method» → следующий путь). */
function ozonTry_(paths, body, cid, key) {
  var lastErr = null;
  for (var p = 0; p < paths.length; p++) {
    for (var t = 0; t < 4; t++) {
      var resp = UrlFetchApp.fetch(API + paths[p], {
        method: 'post', contentType: 'application/json',
        headers: { 'Client-Id': cid, 'Api-Key': key },
        payload: JSON.stringify(body || {}), muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      var txt = resp.getContentText();
      if (code === 200) return JSON.parse(txt || '{}');
      if (code === 404) { lastErr = paths[p] + ' -> 404'; break; } // пробуем следующий путь
      if (code === 429 || code >= 500) { lastErr = paths[p] + ' -> ' + code; Utilities.sleep(2000 * (t + 1)); continue; }
      throw new Error('Ozon ' + paths[p] + ' -> ' + code + ': ' + txt.slice(0, 1000));
    }
    if (lastErr && lastErr.indexOf('429') >= 0) throw new Error('Ozon 429 (лимит запросов): ' + paths[p]);
  }
  throw new Error('Ozon: метод недоступен (' + (lastErr || paths.join(',')) + '). Проверьте права API-ключа на «Поставки FBO».');
}

/* ---------------- настройки ---------------- */
function ensureSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(NAST);
  if (!s) {
    s = ss.insertSheet(NAST);
    var rows = [
      ['Настройка', 'Значение', 'Подсказка'],
      ['Тип поставки', 'CREATE_TYPE_CROSSDOCK', 'CREATE_TYPE_CROSSDOCK (сдаёте всё в одну точку, Ozon развозит) или CREATE_TYPE_DIRECT (везёте на склад кластера сами)'],
      ['ID точки отгрузки (кросс-докинг)', '', 'warehouse_id вашей точки сдачи — возьмите из «Склады отгрузки (справочник)»'],
      ['Таймслот: искать от (дней вперёд)', 1, 'обычно 1 = с завтрашнего дня'],
      ['Таймслот: искать до (дней вперёд)', 14, 'максимум 28'],
      ['Предпочтительное время с (час)', 10, 'первый слот, начинающийся не раньше этого часа; если нет — самый ранний'],
      ['Пауза между кластерами (сек)', 20, 'пауза между заявками; при 429 скрипт сам подождёт и повторит'],
      ['Авто-продолжение (ДА/НЕТ)', 'ДА', 'если кластеров много и 6 минут не хватило — скрипт сам продолжит через минуту']
    ];
    s.getRange(1, 1, rows.length, 3).setValues(rows);
    s.getRange(1, 1, 1, 3).setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold');
    s.getRange(2, 2).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['CREATE_TYPE_CROSSDOCK', 'CREATE_TYPE_DIRECT'], true).build());
    s.getRange(2, 2, 7, 1).setBackground('#FFF2CC');
    s.setColumnWidth(1, 260); s.setColumnWidth(2, 220); s.setColumnWidth(3, 520);
  }
  return s;
}

function getSettings_() {
  var s = ensureSettings_();
  var v = s.getRange(2, 2, 7, 1).getValues();
  return {
    supplyType: String(v[0][0] || 'CREATE_TYPE_CROSSDOCK').trim(),
    dropoffId:  Number(v[1][0]) || 0,
    fromDays:   Math.max(0, Number(v[2][0]) || 1),
    toDays:     Math.min(28, Number(v[3][0]) || 14),
    prefHour:   Math.min(23, Math.max(0, Number(v[4][0]) || 10)),
    pauseSec:   Math.max(5, Number(v[5][0]) || 20),
    autoCont:   String(v[6][0] || 'ДА').trim().toUpperCase() === 'ДА'
  };
}

/* ---------------- ручные корректировки (лист «Корректировки») ---------------- */
function ensureAdjust_(ss) {
  var s = ss.getSheetByName(KORR);
  if (!s) {
    s = ss.insertSheet(KORR);
    s.getRange(1, 1, 1, 4).setValues([['Артикул', 'Кластер (пусто = на все)', 'Отгрузить, шт', 'Статус']])
      .setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    s.setColumnWidth(1, 300); s.setColumnWidth(2, 230); s.setColumnWidth(3, 120); s.setColumnWidth(4, 280);
    s.getRange(2, 1, 50, 3).setBackground('#FFF2CC');
    s.setFrozenRows(1);
  }
  return s;
}

function readAdjustments_(ss) {
  var s = ss.getSheetByName(KORR);
  if (!s || s.getLastRow() < 2) return [];
  var d = s.getRange(2, 1, s.getLastRow() - 1, 3).getValues();
  var out = [];
  for (var i = 0; i < d.length; i++) {
    var art = String(d[i][0] || '').trim(), cl = String(d[i][1] || '').trim(), q = Number(d[i][2]);
    if (art && q > 0) out.push({ art: art, cluster: cl, qty: q, row: i + 2 });
  }
  return out;
}

/** «Отгрузить всего (ручное)» из Свода (колонка J) → корректировки без кластера. */
function readSvodOverrides_(ss) {
  var s = ss.getSheetByName(SVOD);
  if (!s || s.getLastRow() < 2 || s.getLastColumn() < 10) return [];
  var d = s.getRange(2, 1, s.getLastRow() - 1, 10).getValues();
  var out = [];
  for (var i = 0; i < d.length; i++) {
    var art = String(d[i][0] || '').trim(), q = Number(d[i][9]);
    if (art && q > 0) out.push({ art: art, cluster: '', qty: q });
  }
  return out;
}

/** Метаданные всех артикулов из листа «Разгруз» (включая строки без распределения). */
function artMeta_(ss) {
  var s = ss.getSheetByName(RAZ), out = {};
  if (!s || s.getLastRow() < 2) return out;
  var d = s.getRange(2, 1, s.getLastRow() - 1, 13).getValues();
  for (var i = 0; i < d.length; i++) {
    var art = String(d[i][0] || ''); if (!art) continue;
    if (!out[art]) out[art] = { sku: d[i][1], box: Number(d[i][12]) || 0 };
  }
  return out;
}

/** Применение корректировок: с кластером — замена/добавление количества (кратно коробу,
 *  вверх); без кластера — пропорциональное масштабирование текущего распределения. */
function applyAdjustments_(alloc, adj, ss) {
  var s = ss.getSheetByName(KORR);
  var meta = artMeta_(ss);
  adj.forEach(function (a) {
    var m = meta[a.art] || {};
    var box = m.box || 0;
    var rows = alloc.filter(function (r) {
      return String(r.art) === a.art && (!a.cluster || normName_(r.cluster) === normName_(a.cluster));
    });
    var status = '';
    if (a.cluster) {
      var units = box > 0 ? Math.ceil(a.qty / box) * box : a.qty;
      if (rows.length) {
        rows[0].units = units;
        rows[0].boxes = box > 0 ? Math.round(units / box) : '';
        status = 'заменено: ' + units + ' шт' + (units !== a.qty ? ' (кратно коробу ' + box + ')' : '');
      } else if (m.sku) {
        alloc.push({ art: a.art, sku: m.sku, cluster: a.cluster, units: units, box: box,
                     boxes: box > 0 ? Math.round(units / box) : '' });
        status = 'добавлено: ' + units + ' шт';
      } else {
        status = 'артикул не найден в «Разгрузе»';
      }
    } else {
      var cur = 0; rows.forEach(function (r) { cur += r.units; });
      if (!rows.length || cur <= 0) {
        status = 'нет распределения — укажите кластер в колонке B';
      } else {
        var factor = a.qty / cur, total = 0;
        rows.forEach(function (r) {
          var u = r.units * factor;
          if (box > 0) u = Math.ceil(u / box) * box;
          r.units = Math.round(u);
          r.boxes = box > 0 ? Math.round(r.units / box) : '';
          total += r.units;
        });
        status = 'масштабировано: ' + cur + ' → ' + total + ' шт';
      }
    }
    if (s && a.row) s.getRange(a.row, 4).setValue(status);
  });
  return alloc;
}

/* ---------------- палеты и папка этикеток ---------------- */
/** Настройки палет (в «Настройках поставок», колонки E/F). */
function palletSettings_() {
  var s = ensureSettings_();
  if (!String(s.getRange('E2').getValue())) {
    s.getRange('E2').setValue('Палеты: если банок в поставке больше');
    s.getRange('F2').setValue(1000).setBackground('#FFF2CC');
    s.getRange('E3').setValue('Банок на одну палету (макс)');
    s.getRange('F3').setValue(1000).setBackground('#FFF2CC');
    s.setColumnWidth(5, 280); s.setColumnWidth(6, 90);
  }
  return { threshold: Number(s.getRange('F2').getValue()) || 1000,
           perPallet: Math.max(1, Number(s.getRange('F3').getValue()) || 1000) };
}

/** Папка «Этикетки поставок <бренд>» рядом с таблицей разгруза (создаётся один раз).
 *  Расшарьте её складу — все PDF будут падать туда. */
function labelsFolder_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'Этикетки поставок ' + brand_();
  var parent = null;
  try {
    var cur = DriveApp.getFileById(ss.getId());
    var parents = cur.getParents();
    if (parents.hasNext()) parent = parents.next();
  } catch (e) {}
  if (!parent) parent = DriveApp.getRootFolder();
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** Имя PDF: «Кластер №заявки (слот дата)» — этикетки ВСЕХ грузомест заявки в одном файле. */
function labelName_(d) {
  var slot = String(d[7] || '').slice(0, 10);
  return (d[2] ? d[2] + ' ' : '') + '№' + (d[3] || d[5] || '') + (slot ? ' (слот ' + slot + ')' : '');
}

/* ---------------- справочник кластеров (имя → id), кэш 12ч ---------------- */
function clusterMap_(cid, key) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('clusterMapV2');
  if (hit) { var cached = JSON.parse(hit); if (Object.keys(cached).length) return cached; }
  var bodies = [{ cluster_type: 'CLUSTER_TYPE_OZON' }, {}, { cluster_ids: [] }];
  var map = {};
  for (var b = 0; b < bodies.length; b++) {
    var r;
    try { r = ozonTry_(CFG_EP.clusterList, bodies[b], cid, key); } catch (e) { continue; }
    var arr = r.clusters || (r.result && r.result.clusters) || extractList_(r);
    for (var i = 0; i < arr.length; i++) {
      var c = arr[i] || {};
      // форматы: {id,name} | {cluster_id,cluster_name} |
      //          {macrolocal_cluster_id, data:{macrolocal_cluster:{name}}} (v2, июнь 2026)
      var id = (c.id != null) ? c.id : (c.cluster_id != null) ? c.cluster_id : c.macrolocal_cluster_id;
      var nm = (c.name != null) ? c.name : (c.cluster_name != null) ? c.cluster_name
             : (c.data && c.data.macrolocal_cluster && c.data.macrolocal_cluster.name);
      if (nm != null && id != null) map[normName_(nm)] = String(id);
    }
    if (Object.keys(map).length) break;
  }
  if (Object.keys(map).length) { try { cache.put('clusterMapV2', JSON.stringify(map), 21600); } catch (e) {} }
  return map;
}

/** Меню: показать все кластеры Ozon (ID + имя) в «Настройки поставок». */
function showClusters() {
  var creds = ozonCreds_(); if (!creds) return;
  var s = ensureSettings_();
  var map = clusterMap_(creds.cid, creds.key);
  var names = Object.keys(map).sort();
  var rows = [['— Кластеры Ozon (ID → имя) —', '', '']];
  names.forEach(function (n) { rows.push([map[n], n, '']); });
  if (!names.length) {
    var raw = '';
    try { raw = JSON.stringify(ozonTry_(CFG_EP.clusterList, {}, creds.cid, creds.key)).slice(0, 900); }
    catch (e) { raw = String(e.message || e).slice(0, 300); }
    rows.push(['(парсер не нашёл кластеров — сырой ответ ниже, пришлите его Клоду)', '', '']);
    rows.push(['', '', raw]);
  }
  s.getRange(10, 1, rows.length, 3).setValues(rows);
  toastSafe_('Кластеров: ' + names.length + ' — список в «' + NAST + '» с 10-й строки.');
}
function clusterIdByName_(name, map) {
  var n = normName_(name);
  if (map[n]) return map[n];
  for (var k in map) if (k.indexOf(n) >= 0 || n.indexOf(k) >= 0) return map[k];
  return '';
}

/* ================================================================
 *  ШАГ 1. ПЛАН ПОСТАВОК
 * ================================================================ */
var PLAN_HEAD = ['Артикул', 'Кластер', 'Cluster ID', 'Банок', 'Коробов', 'Создавать?',
                 'Статус', 'Заявка №', 'Order ID', 'Склад', 'Таймслот', 'Ошибка', '_items'];

function buildSupplyPlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettings_();
  var creds = ozonCreds_(); if (!creds) return;
  var alloc = collectAllocation_(); // из razgruz_appscript.gs

  // «Отгрузить всего (ручное)» из Свода (кол. J) — масштабирует объём по товару
  var svodAdj = readSvodOverrides_(ss);
  if (svodAdj.length) alloc = applyAdjustments_(alloc, svodAdj, ss);
  // лист «Корректировки» — точечные правки; при конфликте выигрывает он
  ensureAdjust_(ss);
  var adj = readAdjustments_(ss);
  if (adj.length) alloc = applyAdjustments_(alloc, adj, ss);

  // ОДНА ЗАЯВКА = ОДИН АРТИКУЛ + ОДИН КЛАСТЕР (правило селлера, 03.07.2026)
  var byKey = {}, order = [];
  alloc.forEach(function (r) {
    var k = String(r.art) + '||' + r.cluster;
    if (!byKey[k]) { byKey[k] = { art: String(r.art), cluster: r.cluster, items: [], units: 0, boxes: 0 }; order.push(k); }
    var g = byKey[k];
    g.items.push({ sku: Number(r.sku), qty: r.units, art: String(r.art), box: r.box });
    g.units += r.units; g.boxes += (Number(r.boxes) || 0);
  });
  order.sort();

  var cmap = {};
  try { cmap = clusterMap_(creds.cid, creds.key); }
  catch (e) { toastSafe_('Справочник кластеров недоступен: ' + e.message); }

  var s = sheet_(ss, PLAN);
  var rows = [PLAN_HEAD];
  order.forEach(function (k) {
    var g = byKey[k];
    rows.push([g.art, g.cluster, clusterIdByName_(g.cluster, cmap), g.units, g.boxes,
               'ДА', 'К созданию', '', '', '', '', '', JSON.stringify(g.items)]);
  });
  s.getRange(1, 1, rows.length, PLAN_HEAD.length).setValues(rows);
  s.getRange(1, 1, 1, PLAN_HEAD.length).setBackground('#1F4E78').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  s.setFrozenRows(1);
  var w = [280, 220, 110, 90, 90, 100, 200, 130, 110, 200, 190, 320, 40];
  for (var i = 0; i < w.length; i++) s.setColumnWidth(i + 1, w[i]);
  s.hideColumns(PLAN_HEAD.length); // _items
  if (rows.length > 1) {
    s.getRange(2, 6, rows.length - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['ДА', 'НЕТ'], true).build())
      .setBackground('#FFF2CC');
    s.getRange(2, 3, rows.length - 1, 1).setBackground('#FFF2CC'); // Cluster ID можно поправить руками
  }
  toastSafe_('План: ' + (rows.length - 1) + ' заявок (артикул × кластер). Проверьте и запустите «Создать заявки».');
}

/* ================================================================
 *  ШАГ 2. СОЗДАНИЕ ЗАЯВОК
 * ================================================================ */
function createSupplies() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(PLAN);
  if (!s || s.getLastRow() < 2) { alertSafe_('Поставки', 'Сначала «Сформировать план поставок».'); return; }
  var n = countPending_(s);
  if (!n) { alertSafe_('Поставки', 'Нет строк со статусом «К созданию» и отметкой «ДА».'); return; }
  var st = getSettings_();
  if (st.supplyType === 'CREATE_TYPE_CROSSDOCK' && !st.dropoffId) {
    alertSafe_('Поставки', 'Для кросс-докинга укажите «ID точки отгрузки» в «' + NAST + '» (справочник — в меню).');
    return;
  }
  var resp = ui.alert('Создать заявки на поставку?',
    'Будет создано заявок: ' + n + ' (тип: ' + st.supplyType.replace('CREATE_TYPE_', '') + ').\n' +
    'Пауза между кластерами ' + st.pauseSec + ' сек — примерно ' +
    Math.ceil(n * st.pauseSec / 60) + ' мин.\n\nЭто реальные заявки в кабинете Ozon. Продолжить?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  processSupplyQueue_();
}

/** Продолжение после авто-триггера (без UI). */
function continueCreateSupplies() {
  killTriggers_('continueCreateSupplies');
  processSupplyQueue_();
}

function countPending_(s) {
  var d = s.getRange(2, 1, s.getLastRow() - 1, PLAN_HEAD.length).getValues();
  var n = 0;
  for (var i = 0; i < d.length; i++)
    if (String(d[i][5]).toUpperCase() === 'ДА' && isPendingStatus_(d[i][6])) n++;
  return n;
}
function isPendingStatus_(st) {
  st = String(st || '');
  return st === '' || st === 'К созданию' || st === 'В очереди' ||
         st.indexOf('Ошибка') === 0 || st.indexOf('Лимит') === 0;
}

function processSupplyQueue_() {
  var started = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(PLAN);
  var creds = ozonCreds_(); if (!creds) return;
  var st = getSettings_();
  var lastRow = s.getLastRow();
  var first = true;
  var retries429 = {};

  for (var row = 2; row <= lastRow; row++) {
    var d = s.getRange(row, 1, 1, PLAN_HEAD.length).getValues()[0];
    if (String(d[5]).toUpperCase() !== 'ДА' || !isPendingStatus_(d[6])) continue;

    // предохранитель: 6-минутный лимит Apps Script
    if (Date.now() - started > 4.5 * 60 * 1000) {
      markQueued_(s, row, lastRow);
      if (st.autoCont) {
        ScriptApp.newTrigger('continueCreateSupplies').timeBased().after(60 * 1000).create();
        toastSafe_('Лимит времени: продолжу автоматически через минуту.');
      } else {
        toastSafe_('Лимит времени: запустите «Создать заявки в Ozon» ещё раз.');
      }
      return;
    }
    if (!first) Utilities.sleep(st.pauseSec * 1000); // лимит Ozon 2/мин на черновики
    first = false;

    try {
      createOneSupply_(s, row, d, creds, st);
    } catch (e) {
      var msg = String(e.message || e);
      if (msg.indexOf('429') >= 0 || msg.indexOf('лимит') >= 0) {
        retries429[row] = (retries429[row] || 0) + 1;
        if (retries429[row] <= 1) {
          // временный лимит запросов: подождать минуту и повторить ЭТУ же строку
          setPlan_(s, row, { status: 'Пауза: лимит запросов Ozon…', err: '' });
          Utilities.sleep(60 * 1000);
          row--;
          continue;
        }
        setPlan_(s, row, { status: 'Лимит Ozon — повторить позже', err: msg.slice(0, 500) });
        markQueued_(s, row + 1, lastRow);
        toastSafe_('Ozon стабильно отвечает 429. Запустите «Создать заявки» чуть позже (обычно хватает 10–30 минут).');
        return;
      }
      setPlan_(s, row, { status: 'Ошибка', err: msg.slice(0, 6000) });
    }
  }
  toastSafe_('Готово. Смотрите листы «' + PLAN + '» и «' + REESTR + '».');
}

function markQueued_(s, fromRow, lastRow) {
  for (var r = fromRow; r <= lastRow; r++) {
    var v = s.getRange(r, 1, 1, PLAN_HEAD.length).getValues()[0];
    if (String(v[5]).toUpperCase() === 'ДА' && isPendingStatus_(v[6]))
      s.getRange(r, 7).setValue('В очереди');
  }
  SpreadsheetApp.flush();
}

function setPlan_(s, row, o) {
  if (o.status !== undefined) s.getRange(row, 7).setValue(o.status);
  if (o.num !== undefined) s.getRange(row, 8).setValue(o.num);
  if (o.orderId !== undefined) s.getRange(row, 9).setValue(o.orderId);
  if (o.wh !== undefined) s.getRange(row, 10).setValue(o.wh);
  if (o.slot !== undefined) s.getRange(row, 11).setValue(o.slot);
  if (o.err !== undefined) s.getRange(row, 12).setValue(o.err);
  SpreadsheetApp.flush();
}

/** Полный цикл по одному кластеру. */
function createOneSupply_(s, row, d, creds, st) {
  var artName = String(d[0] || '');
  var clusterName = d[1];
  var clusterId = String(d[2] || '').trim();
  var items = JSON.parse(d[12] || '[]');
  if (!items.length) throw new Error('Пустой состав (_items) — пересоберите план.');
  if (!clusterId) throw new Error('Не определён Cluster ID — впишите вручную в колонку C (справочник: «🗺 Кластеры»).');

  // --- 1. черновик (перебор форм payload; удачная запоминается) ---
  setPlan_(s, row, { status: 'Черновик…', err: '' });
  var props = PropertiesService.getScriptProperties();
  var prefName = props.getProperty('DRAFT_VARIANT_' + st.supplyType) || '';
  var attempts = draftAttempts_(st, clusterId, items);
  if (prefName) attempts.sort(function (a, b) { return (b.name === prefName ? 1 : 0) - (a.name === prefName ? 1 : 0); });
  var dr = null, draftErrs = [];
  for (var ai = 0; ai < attempts.length && !dr; ai++) {
    for (var tryN = 0; tryN < 2 && !dr; tryN++) {
      try {
        dr = ozonTry_(attempts[ai].eps, attempts[ai].body, creds.cid, creds.key);
        props.setProperty('DRAFT_VARIANT_' + st.supplyType, attempts[ai].name);
      } catch (e) {
        var em = String(e.message || e);
        if ((em.indexOf('429') >= 0 || em.indexOf('лимит') >= 0) && tryN === 0) {
          setPlan_(s, row, { status: 'Черновик… (пауза: лимит Ozon)' });
          Utilities.sleep(st.pauseSec * 1000);
          continue; // повтор того же варианта после паузы
        }
        draftErrs.push(attempts[ai].name + ' → ' + em.slice(0, 900));
        if (em.indexOf('400') < 0 && em.indexOf('429') < 0 && em.indexOf('лимит') < 0)
          throw new Error(draftErrs.join('\n').slice(0, 6000)); // не-валидационная ошибка
        break;
      }
    }
    if (!dr && ai < attempts.length - 1) Utilities.sleep(3000);
  }
  if (!dr) throw new Error('Черновик не создан:\n' + draftErrs.join('\n').slice(0, 6000));
  var opId = dr.operation_id || (dr.result && dr.result.operation_id);
  var newDraftId = dr.draft_id || (dr.result && dr.result.draft_id); // новый API (2026) отдаёт draft_id сразу
  if (!opId && !newDraftId) throw new Error('draft/create: нет draft_id/operation_id: ' + JSON.stringify(dr).slice(0, 200));

  // --- 2. расчёт черновика (склады кластера) ---
  setPlan_(s, row, { status: 'Расчёт черновика…' });
  var info = pollDraftInfo_(opId, newDraftId, creds);
  var draftId = info.draft_id || newDraftId;
  var wh = pickWarehouse_(info, clusterId);
  if (!wh) throw new Error('Нет доступного склада в кластере (draft ' + draftId + '). ' +
    (shortDraftErrors_(info) || '') + ' RAW: ' + JSON.stringify(info.clusters || info).slice(0, 800));

  // --- 3. таймслот ---
  setPlan_(s, row, { status: 'Выбор таймслота…', wh: wh.name || wh.id || 'выберет Ozon' });
  var slot = pickTimeslot_(draftId, wh, clusterId, creds, st);

  // --- 4. заявка ---
  setPlan_(s, row, { status: 'Создаём заявку…', slot: slot.from + ' → ' + slot.to });
  var sc = createSupplyFromDraft_(draftId, wh, clusterId, slot, creds, st);
  var scRoot = sc.result || sc;
  if (scRoot.error_reasons && scRoot.error_reasons.length)
    throw new Error('supply/create: ' + JSON.stringify(scRoot.error_reasons).slice(0, 900));
  var opId2 = scRoot.operation_id || '';
  var orderIds = scRoot.order_ids || [];
  if (!orderIds.length) {
    // метод асинхронный: заявка принята, номер забираем опросом create/status
    setPlan_(s, row, { status: 'Ждём номер заявки…' });
    orderIds = pollSupplyStatus_(opId2 ? { operation_id: opId2 } : { draft_id: Number(draftId) }, creds);
  }
  if (!orderIds.length) throw new Error('supply/create: нет order_id: ' + JSON.stringify(sc).slice(0, 500));
  var orderId = orderIds[0];

  // --- 5. карточка заявки ---
  var num = '', state = '', supplyId = '';
  if (orderId) {
    try {
      var og = ozonTry_(CFG_EP.orderGet, { order_ids: [String(orderId)] }, creds.cid, creds.key);
      var orders = og.orders || (og.result && og.result.orders) || [];
      if (orders.length) {
        num = orders[0].supply_order_number || '';
        state = orders[0].state || '';
        var sups = orders[0].supplies || [];
        if (sups.length) supplyId = sups[0].supply_id || sups[0].id || '';
      }
    } catch (e) { /* карточка подтянется в «Обновить статусы» */ }
  }

  setPlan_(s, row, { status: 'Создана ✓', num: num || orderId, orderId: orderId });
  appendReestr_([new Date(), brand_(), clusterName, num || '', orderId, supplyId,
                 wh.name || wh.id, slot.from + ' → ' + slot.to, state || 'создана', '',
                 sumBoxes_(items), '', JSON.stringify(items)]);
}

/** Варианты payload для создания черновика (Ozon не публикует точную схему новых
 *  методов; требование ClusterInfo подтверждено ошибкой валидации 03.07.2026). */
function draftAttempts_(st, clusterId, items) {
  var it = items.map(function (x) { return { sku: Number(x.sku), quantity: Number(x.qty) }; });
  var cidS = String(clusterId);
  var out = [];
  // ПОДТВЕРЖДЕНО ошибками валидации Ozon 03.07.2026:
  //   cluster_info = { macrolocal_cluster_id: число > 0, items: [{sku, quantity}] (1..5000) }
  var ci = { macrolocal_cluster_id: Number(clusterId), items: it };
  if (st.supplyType === 'CREATE_TYPE_DIRECT') {
    out.push({ name: 'direct.macro', eps: CFG_EP.draftDirect,
               body: { cluster_info: ci } });
    out.push({ name: 'direct.macro+topIds', eps: CFG_EP.draftDirect,
               body: { cluster_info: ci, cluster_ids: [cidS] } });
  } else {
    // Подтверждено валидацией 03.07.2026: delivery_info.type ≠ 0 (шлём 1 — прошло),
    // drop_off_warehouse.warehouse_type ≠ 0 — значения перебираем числами
    // (protobuf допускает; при неверном значении валидатор вернёт допустимый список).
    // deletion_sku_mode ≠ 0 (подтверждено 03.07.2026) — режим обработки SKU,
    // которые нельзя поставить: 1/2 перебираем, warehouse_type=1 уже прошёл валидацию.
    function dwv(wt) { return { warehouse_id: Number(st.dropoffId), warehouse_type: wt }; }
    out.push({ name: 'cross.macro.wt1.del1', eps: CFG_EP.draftCrossdock,
               body: { cluster_info: ci, delivery_info: { type: 1, drop_off_warehouse: dwv(1) }, deletion_sku_mode: 1 } });
    out.push({ name: 'cross.macro.wt1.del2', eps: CFG_EP.draftCrossdock,
               body: { cluster_info: ci, delivery_info: { type: 1, drop_off_warehouse: dwv(1) }, deletion_sku_mode: 2 } });
    out.push({ name: 'cross.macro.wt2.del1', eps: CFG_EP.draftCrossdock,
               body: { cluster_info: ci, delivery_info: { type: 1, drop_off_warehouse: dwv(2) }, deletion_sku_mode: 1 } });
    out.push({ name: 'cross.macro.wt3.del1', eps: CFG_EP.draftCrossdock,
               body: { cluster_info: ci, delivery_info: { type: 1, drop_off_warehouse: dwv(3) }, deletion_sku_mode: 1 } });
  }
  return out;
}

function sumBoxes_(items) {
  var n = 0;
  items.forEach(function (it) { if (it.box > 0) n += Math.round(it.qty / it.box); });
  return n;
}
function shortDraftErrors_(info) {
  try {
    var e = info.errors || [];
    return e.map(function (x) { return x.error_message || JSON.stringify(x); }).join('; ').slice(0, 200);
  } catch (x) { return ''; }
}

function pollDraftInfo_(opId, draftId, creds) {
  // v2/draft/create/info принимает draft_id (новый флоу) или operation_id (старый)
  var body = draftId ? { draft_id: Number(draftId) } : { operation_id: opId };
  for (var t = 0; t < 30; t++) {
    var r = ozonTry_(CFG_EP.draftInfo, body, creds.cid, creds.key);
    var status = String(r.status || '');
    if (status.indexOf('ERROR') >= 0 || status.indexOf('FAILED') >= 0)
      throw new Error('Черновик не рассчитан: ' + (shortDraftErrors_(r) || status));
    var cls = r.clusters || [];
    if (cls.length) { if (!r.draft_id && draftId) r.draft_id = Number(draftId); return r; }
    if (r.draft_id && !draftId) { draftId = r.draft_id; body = { draft_id: Number(draftId) }; }
    Utilities.sleep(8000); // реже опрашиваем create/info — меньше шансов словить 429
  }
  throw new Error('Черновик: расчёт не завершился за 2 минуты (draft ' + (draftId || opId) + ')');
}

/** Лучший доступный склад кластера: минимальный total_rank (1 = лучший), иначе max score.
 *  Парсер всеяден к формам ответа v2/draft/create/info. */
function pickWarehouse_(info, clusterId) {
  var clusters = info.clusters || [];
  var best = null;
  for (var i = 0; i < clusters.length; i++) {
    var c = clusters[i];
    var ccid = (c.cluster_id != null) ? c.cluster_id
             : (c.macrolocal_cluster_id != null) ? c.macrolocal_cluster_id : c.id;
    if (clusterId && ccid != null && String(ccid) !== String(clusterId) && clusters.length > 1) continue;
    var ws = c.warehouses || (c.data && c.data.warehouses) || [];
    for (var j = 0; j < ws.length; j++) {
      var w = ws[j];
      if (!whAvailable_(w)) continue;
      // при кросс-доке storage_warehouse = null: склад хранения распределяет Ozon —
      // кандидат без id допустим (id пустой, есть bundle_id)
      var wid = w.warehouse_id || w.id ||
                (w.storage_warehouse && (w.storage_warehouse.warehouse_id || w.storage_warehouse.id)) ||
                (w.supply_warehouse && (w.supply_warehouse.warehouse_id || w.supply_warehouse.id)) || '';
      var nm = w.name || (w.supply_warehouse && w.supply_warehouse.name) ||
               (w.storage_warehouse && w.storage_warehouse.name) || (wid ? '' : 'выберет Ozon');
      var cand = { id: wid, name: nm, bundle: w.bundle_id || '',
                   rank: (w.total_rank == null ? 9999 : Number(w.total_rank)), score: Number(w.total_score) || 0 };
      if (!best || cand.rank < best.rank || (cand.rank === best.rank && cand.score > best.score)) best = cand;
    }
  }
  return best;
}

/** Доступность склада: status/availability_status — объект {is_available,state}, строка или отсутствует. */
function whAvailable_(w) {
  var st = (w.status != null) ? w.status : w.availability_status;
  if (st == null) return true;
  if (typeof st === 'string') return st.indexOf('UNAVAILABLE') < 0 && st.indexOf('NOT_AVAILABLE') < 0;
  if (typeof st === 'object') {
    if (st.is_available === false) return false;
    if (typeof st.state === 'string' &&
        (st.state.indexOf('UNAVAILABLE') >= 0 || st.state.indexOf('NOT_AVAILABLE') >= 0)) return false;
  }
  return true;
}

function ext_(a, b) { var o = {}, k; for (k in a) o[k] = a[k]; for (k in b) o[k] = b[k]; return o; }

/** Формы selected_cluster_warehouses (обязателен в v2-методах; состав не документирован,
 *  собираем из того, что дал черновик: macrolocal_cluster_id + bundle_id). */
function scwVariants_(wh, clusterId) {
  var mc = Number(clusterId), out = [];
  if (wh && wh.bundle) out.push({ name: 'mc+bundle', v: [{ macrolocal_cluster_id: mc, bundle_id: wh.bundle }] });
  out.push({ name: 'mc', v: [{ macrolocal_cluster_id: mc }] });
  if (wh && wh.bundle) out.push({ name: 'bundle', v: [{ bundle_id: wh.bundle }] });
  if (wh && wh.id) out.push({ name: 'mc+wid', v: [{ macrolocal_cluster_id: mc, storage_warehouse_id: Number(wh.id) }] });
  out.push({ name: 'empty', v: [{}] });
  return out;
}
function scwPreferred_(list) {
  var pref = PropertiesService.getScriptProperties().getProperty('SCW_SHAPE') || '';
  if (pref) list.sort(function (a, b) { return (b.name === pref ? 1 : 0) - (a.name === pref ? 1 : 0); });
  return list;
}

function pickTimeslot_(draftId, wh, clusterId, creds, st) {
  var from = new Date(); from.setDate(from.getDate() + st.fromDays);
  var to = new Date();   to.setDate(to.getDate() + st.toDays);
  // формат дат: строго YYYY-MM-DD (подтверждено валидацией 03.07.2026)
  var base = { draft_id: Number(draftId),
               date_from: Utilities.formatDate(from, 'GMT', 'yyyy-MM-dd'),
               date_to: Utilities.formatDate(to, 'GMT', 'yyyy-MM-dd') };
  var scws = scwPreferred_(scwVariants_(wh, clusterId));
  // supply_type обязателен (≠ 0, подтверждено 03.07.2026): сначала имя enum, затем числа
  var isCross = st.supplyType !== 'CREATE_TYPE_DIRECT';
  var typeVals = isCross ? ['CROSSDOCK', 2, 1] : ['DIRECT', 1, 2];
  var props = PropertiesService.getScriptProperties();
  var errs = [];
  for (var ti = 0; ti < typeVals.length; ti++) {
    for (var b = 0; b < scws.length; b++) {
      try {
        var r = ozonTry_(CFG_EP.timeslotInfo,
          ext_(base, { supply_type: typeVals[ti], selected_cluster_warehouses: scws[b].v }),
          creds.cid, creds.key);
        props.setProperty('SCW_SHAPE', scws[b].name);
        props.setProperty('TSLOT_TYPE', String(typeVals[ti]));
        var slot = slotScan_(r, st.prefHour);
        if (slot) return slot;
        throw new Error('СХЕМА ОК (type=' + typeVals[ti] + ', ' + scws[b].name +
          '), но слотов нет в окне ' + st.fromDays + '–' + st.toDays +
          ' дн. RAW-ответ: ' + JSON.stringify(r).slice(0, 2500));
      } catch (e) {
        var em = String(e.message || e);
        if (em.indexOf('СХЕМА ОК') === 0) throw e; // схема верная, слотов нет — перебор не поможет
        errs.push('[type=' + typeVals[ti] + ', ' + scws[b].name + '] ' + em.slice(0, 900));
        if (em.indexOf('SupplyType') >= 0) break; // тип не подходит — пробуем следующий тип
        // ошибка не про scw — перебор scw не поможет, идём к следующему типу
        if (em.indexOf('SelectedClusterWarehouses') < 0) break;
      }
    }
  }
  throw new Error('Таймслот не получен:\n' + errs.join('\n').slice(0, 6000));
}

function slotScan_(r, prefHour) {
  // ответ бывает обёрнут в result; drop_off_warehouse_timeslots — объект ИЛИ массив
  var root = r.result || r;
  var whs = root.drop_off_warehouse_timeslots || root.timeslots || [];
  if (whs && whs.days) whs = [whs]; // объект с days → оборачиваем в массив
  if (!whs || !whs.length) whs = [];
  var fallback = null;
  for (var i = 0; i < whs.length; i++) {
    var days = whs[i].days || [];
    for (var dIdx = 0; dIdx < days.length; dIdx++) {
      var slots = days[dIdx].timeslots || [];
      for (var sIdx = 0; sIdx < slots.length; sIdx++) {
        var sl = slots[sIdx];
        if (!sl.from_in_timezone) continue;
        if (!fallback) fallback = sl;
        var hour = Number(String(sl.from_in_timezone).replace(/^.*T(\d{2}).*$/, '$1'));
        if (hour >= prefHour) return { from: sl.from_in_timezone, to: sl.to_in_timezone };
      }
    }
  }
  return fallback ? { from: fallback.from_in_timezone, to: fallback.to_in_timezone } : null;
}

/** Создание заявки из черновика: перебор форм v2/draft/supply/create, удачная запоминается.
 *  selected_cluster_warehouses — та же форма, что сработала в timeslot/info (SCW_SHAPE). */
function createSupplyFromDraft_(draftId, wh, clusterId, slot, creds, st) {
  var isCross = st.supplyType !== 'CREATE_TYPE_DIRECT';
  var ts = { from_in_timezone: slot.from, to_in_timezone: slot.to };
  var scws = scwPreferred_(scwVariants_(wh, clusterId));
  var types = [isCross ? 'CROSSDOCK' : 'DIRECT', st.supplyType, null];
  // тип, сработавший в timeslot/info — первым
  var storedType = PropertiesService.getScriptProperties().getProperty('TSLOT_TYPE');
  if (storedType) types.unshift(isNaN(Number(storedType)) ? storedType : Number(storedType));
  var bodies = [];
  for (var ti = 0; ti < types.length; ti++) {
    var b0 = { draft_id: Number(draftId), timeslot: ts, selected_cluster_warehouses: scws[0].v };
    if (types[ti]) b0.supply_type = types[ti];
    bodies.push({ name: 'type:' + (types[ti] || 'нет') + '|' + scws[0].name, body: b0 });
  }
  if (scws.length > 1) {
    var b1 = { draft_id: Number(draftId), timeslot: ts, selected_cluster_warehouses: scws[1].v,
               supply_type: types[0] };
    bodies.push({ name: 'type:' + types[0] + '|' + scws[1].name, body: b1 });
  }
  var props = PropertiesService.getScriptProperties();
  var pref = props.getProperty('SUPPLY_VARIANT') || '';
  if (pref) bodies.sort(function (a, b2) { return (b2.name === pref ? 1 : 0) - (a.name === pref ? 1 : 0); });
  var errs = [];
  for (var i = 0; i < bodies.length; i++) {
    try {
      var r = ozonTry_(CFG_EP.supplyCreate, bodies[i].body, creds.cid, creds.key);
      props.setProperty('SUPPLY_VARIANT', bodies[i].name);
      return r;
    } catch (e) {
      var em = String(e.message || e);
      errs.push(bodies[i].name + ' → ' + em.slice(0, 900));
      if (em.indexOf('400') < 0) break; // не-валидационная ошибка — перебор не поможет
    }
  }
  throw new Error('Заявка не создана:\n' + errs.join('\n').slice(0, 6000));
}

function pollSupplyStatus_(body, creds) {
  for (var t = 0; t < 30; t++) {
    var r = ozonTry_(CFG_EP.supplyStatus, body, creds.cid, creds.key);
    var root = r.result || r;
    var ids = root.order_ids || [];
    if (!ids.length && root.order_id) ids = [root.order_id];
    var status = String(root.status || '');
    if (ids.length) return ids;
    var er = root.error_reasons || root.errors;
    if ((er && er.length) || status.indexOf('ERROR') >= 0 || status.indexOf('FAILED') >= 0)
      throw new Error('Заявка не создана: ' + JSON.stringify(er || status).slice(0, 900));
    Utilities.sleep(4000);
  }
  throw new Error('Заявка: статус не получен за 2 минуты (' + JSON.stringify(body) + ')');
}

/* ---------------- реестр «Поставки» ---------------- */
var REESTR_HEAD = ['Дата', 'Бренд', 'Кластер', 'Заявка №', 'Order ID', 'Supply ID',
                   'Склад', 'Таймслот', 'Статус', 'Этикетки (PDF)', 'Коробов', 'Срок годности', '_items'];

function ensureReestr_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(REESTR);
  if (!s) {
    s = ss.insertSheet(REESTR);
    s.getRange(1, 1, 1, REESTR_HEAD.length).setValues([REESTR_HEAD])
      .setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold')
      .setHorizontalAlignment('center').setWrap(true);
    s.setFrozenRows(1);
    var w = [130, 90, 220, 130, 110, 110, 200, 190, 150, 220, 80, 120, 40];
    for (var i = 0; i < w.length; i++) s.setColumnWidth(i + 1, w[i]);
    s.getRange(2, 12, 500, 1).setBackground('#FFF2CC').setNumberFormat('dd.mm.yyyy'); // Срок годности
    s.hideColumns(REESTR_HEAD.length);
  } else if (String(s.getRange(1, 12).getValue()) === '_items') {
    // миграция старого реестра: вставить «Срок годности» перед _items
    s.insertColumnBefore(12);
    s.getRange(1, 12).setValue('Срок годности').setBackground('#1F4E78').setFontColor('#FFFFFF')
      .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    s.setColumnWidth(12, 120);
    s.getRange(2, 12, 500, 1).setBackground('#FFF2CC').setNumberFormat('dd.mm.yyyy');
  }
  return s;
}
function appendReestr_(row) {
  var s = ensureReestr_();
  s.getRange(s.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();
}

/* ================================================================
 *  ШАГ 3. ГРУЗОМЕСТА + ЭТИКЕТКИ
 *  Один пункт меню, два режима: выделенная строка «Поставок»
 *  или ВСЕ новые заявки подряд (выбор при запуске).
 * ================================================================ */
function sendCargoesForActiveRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ensureReestr_();
  var ui = SpreadsheetApp.getUi();
  var creds = ozonCreds_(); if (!creds) return;
  var onReestr = ss.getActiveSheet().getName() === REESTR;
  var row = (onReestr && s.getActiveRange()) ? s.getActiveRange().getRow() : 0;
  var canSingle = onReestr && row >= 2;
  var resp = ui.alert('Грузоместа + этикетки',
    canSingle
      ? 'ДА — только выделенная строка (' + row + ')\nНЕТ — все новые заявки подряд\nОТМЕНА — выход'
      : 'Обработать ВСЕ новые заявки подряд?\n(Чтобы обработать одну — выделите её строку на листе «' + REESTR + '».)',
    canSingle ? ui.ButtonSet.YES_NO_CANCEL : ui.ButtonSet.YES_NO);
  if (resp === ui.Button.CANCEL) return;
  if (canSingle && resp === ui.Button.YES) {
    var d = s.getRange(row, 1, 1, REESTR_HEAD.length).getValues()[0];
    try {
      alertSafe_('Готово', cargoesForRow_(s, row, d, creds));
    } catch (e) {
      s.getRange(row, 10).setValue('ОШИБКА: ' + String(e.message || e).slice(0, 900));
      alertSafe_('Ошибка', String(e.message || e).slice(0, 1200));
    }
    return;
  }
  if (!canSingle && resp !== ui.Button.YES) return;
  processAllNewCargoes_(s, creds);
}

/** Все заявки без переданных грузомест — подряд, с лимитом времени Apps Script. */
function processAllNewCargoes_(s, creds) {
  var n = s.getLastRow() - 1;
  if (n < 1) { alertSafe_('Грузоместа', 'Реестр «' + REESTR + '» пуст.'); return; }
  var data = s.getRange(2, 1, n, REESTR_HEAD.length).getValues();
  var pend = [];
  for (var i = 0; i < data.length; i++) {
    var st = String(data[i][8] || '');
    var itemsJson = String(data[i][12] || data[i][11] || '');
    if (data[i][5] && itemsJson.indexOf('[') === 0 && st.indexOf('грузоместа переданы') < 0) pend.push(i + 2);
  }
  if (!pend.length) { alertSafe_('Грузоместа', 'Новых заявок нет — у всех уже переданы грузоместа.'); return; }
  var started = Date.now(), done = 0, failed = 0;
  for (var pIdx = 0; pIdx < pend.length; pIdx++) {
    if (Date.now() - started > 4.5 * 60 * 1000) {
      alertSafe_('Грузоместа', 'Обработано ' + done + ' из ' + pend.length +
        ' (лимит времени Apps Script). Запустите пункт ещё раз — продолжу с оставшихся.');
      return;
    }
    var row = pend[pIdx];
    var d = s.getRange(row, 1, 1, REESTR_HEAD.length).getValues()[0];
    toastSafe_('Заявка ' + (d[3] || d[4]) + ' (' + (pIdx + 1) + '/' + pend.length + ')…');
    try { cargoesForRow_(s, row, d, creds); done++; }
    catch (e) {
      failed++;
      s.getRange(row, 10).setValue('ОШИБКА: ' + String(e.message || e).slice(0, 900));
      SpreadsheetApp.flush();
    }
  }
  alertSafe_('Грузоместа + этикетки', 'Готово: ' + done + ' успешно' +
    (failed ? ', ошибок: ' + failed + ' (текст — в колонке «Этикетки»)' : '') + '.');
}

/** Полный цикл по одной строке реестра: короба/палеты → передача → этикетки по артикулам. */
function cargoesForRow_(s, row, d, creds) {
  var supplyId = d[5];
  if (!supplyId) throw new Error('В строке нет Supply ID — сначала «Обновить статусы заявок».');
  var items = JSON.parse(d[12] || d[11] || '[]'); // d[12] после миграции, d[11] — старый реестр
  if (!items.length) throw new Error('В строке нет состава (_items).');

  // палеты, если банок больше порога (Настройки, F2); иначе короба по фасовке
  var pal = palletSettings_();
  var totalUnits = 0;
  items.forEach(function (it) { totalUnits += Number(it.qty) || 0; });
  var usePallet = totalUnits > pal.threshold;
  var cargoType = usePallet ? 'PALLET' : 'BOX';

  // СХЕМА ИЗ ОФИЦИАЛЬНОЙ ДОКУМЕНТАЦИИ (03.07.2026):
  // { cargoes: [{ key, value: { items: [{offer_id, quant, quantity}], type } }],
  //   delete_current_version, supply_id }
  // срок годности — колонка «Срок годности» ЭТОЙ строки реестра (заявка = один артикул)
  var expDate = d[11];
  var hasExp = (expDate instanceof Date && !isNaN(expDate.getTime()));
  function withExp_(o) {
    if (hasExp) o.expires_at = Utilities.formatDate(expDate, 'GMT', "yyyy-MM-dd'T'23:59:59'Z'");
    return o;
  }
  var typeVariants = [
    { name: 'value.type(offer_id)', make: function (it, qty) {
        return { key: Utilities.getUuid(),
                 value: { type: cargoType, items: [withExp_({ offer_id: String(it.art), quant: 1, quantity: qty })] } }; } },
    { name: 'value.type(sku)', make: function (it, qty) {
        return { key: Utilities.getUuid(),
                 value: { type: cargoType, items: [withExp_({ sku: Number(it.sku), quantity: qty })] } }; } }
  ];
  var propsC = PropertiesService.getScriptProperties();
  var prefC = propsC.getProperty('CARGO_TYPE_SHAPE') || '';
  if (prefC) typeVariants.sort(function (a, b) { return (b.name === prefC ? 1 : 0) - (a.name === prefC ? 1 : 0); });

  function buildCargoes_(make) {
    var out = [];
    items.forEach(function (it) {
      var box = Number(it.box) || 0, qty = Number(it.qty) || 0;
      var chunk = usePallet ? pal.perPallet : box;
      if (chunk <= 0 || qty <= 0) return;
      var nFull = Math.floor(qty / chunk), rest = qty - nFull * chunk;
      for (var b = 0; b < nFull; b++) out.push(make(it, chunk));
      if (rest > 0) out.push(make(it, rest));
    });
    return out;
  }
  var cargoes = buildCargoes_(typeVariants[0].make);
  if (!cargoes.length) throw new Error('Не получилось собрать грузоместа (нет фасовок в «Коробах»).');

  var cr = null, cErrs = [];
  for (var tv = 0; tv < typeVariants.length && !cr; tv++) {
    try {
      cr = ozonTry_(CFG_EP.cargoesCreate,
        { supply_id: Number(supplyId), delete_current_version: true,
          cargoes: buildCargoes_(typeVariants[tv].make) },
        creds.cid, creds.key);
      propsC.setProperty('CARGO_TYPE_SHAPE', typeVariants[tv].name);
    } catch (e) {
      var emC = String(e.message || e);
      cErrs.push(typeVariants[tv].name + ' → ' + emC.slice(0, 600));
      if (emC.indexOf('400') < 0) throw new Error(cErrs.join('\n').slice(0, 5000));
    }
  }
  if (!cr) throw new Error('Грузоместа не переданы:\n' + cErrs.join('\n').slice(0, 5000));

  var opId = cr.operation_id || (cr.result && cr.result.operation_id);
  if (opId) {
    try {
      for (var t = 0; t < 20; t++) {
        var st2 = ozonTry_(CFG_EP.cargoesInfo, { operation_id: opId }, creds.cid, creds.key);
        var root2 = st2.result || st2;
        var stat = String(root2.status || '');
        if (stat.indexOf('SUCCESS') >= 0 || (root2.cargoes && root2.cargoes.length)) break;
        if (stat.indexOf('ERROR') >= 0 || stat.indexOf('FAILED') >= 0)
          throw new Error('Грузоместа не приняты: ' + JSON.stringify(root2.errors || stat).slice(0, 900));
        Utilities.sleep(3000);
      }
    } catch (eInfo) {
      var emI = String(eInfo.message || eInfo);
      if (emI.indexOf('404') >= 0 || emI.indexOf('недоступен') >= 0) Utilities.sleep(5000);
      else throw eInfo;
    }
  }

  // этикетки: по-артикульные PDF в папки «ДД.ММ.ГГГГ — <артикул>»
  var lab = saveLabelsByArt_(supplyId, creds, d, items);
  s.getRange(row, 10).setValue(lab.links.length ? lab.links.join('\n') : (lab.err || ''));
  s.getRange(row, 9).setValue('грузоместа переданы' + (hasExp ? '' : ' (⚠ без срока годности)'));
  SpreadsheetApp.flush();
  return 'Заявка ' + (d[3] || d[4]) + ': ' + (usePallet ? 'ПАЛЕТЫ' : 'короба') + ', грузомест: ' + cargoes.length +
         (lab.links.length ? ', этикеток PDF: ' + lab.links.length : (lab.err ? ' (' + lab.err + ')' : '')) +
         (hasExp ? ' | срок годности ✓'
                  : ' | ⚠ срок годности не указан — колонка «Срок годности» в «Поставках»');
}

/** Скачивание PDF этикеток по телу labelCreate ({supply_id: N} или {cargo_ids: [...]}) → Blob. */
function labelPdfBlob_(body, creds) {
  var lc = ozonTry_(CFG_EP.labelCreate, body, creds.cid, creds.key);
  var lcRoot = lc.result || lc;
  var lOp = lcRoot.operation_id || '';
  var guid = lcRoot.file_guid || '';
  var directUrl = lcRoot.file_url || lcRoot.url || '';
  var lastRaw = JSON.stringify(lc).slice(0, 700);
  for (var t = 0; t < 20 && !guid && !directUrl; t++) {
    Utilities.sleep(3000);
    var lg = ozonTry_(CFG_EP.labelGet, lOp ? { operation_id: lOp } : body, creds.cid, creds.key);
    var r = lg.result || lg;
    lastRaw = JSON.stringify(lg).slice(0, 700);
    guid = r.file_guid || '';
    directUrl = r.file_url || r.url || '';
    var lst = String(r.status || '');
    if (lst.indexOf('ERROR') >= 0 || lst.indexOf('FAILED') >= 0) throw new Error('Этикетки: ' + lastRaw);
  }
  var urls = [];
  if (directUrl) urls.push(directUrl);
  if (guid) urls.push(API + '/v1/cargoes-label/file/' + guid);
  if (!urls.length) return { blob: null, guid: '', err: 'не получен file_guid. Ответ: ' + lastRaw };
  var lastCode = 0;
  for (var u = 0; u < urls.length; u++) {
    for (var dR = 0; dR < 10; dR++) {
      var pdf = UrlFetchApp.fetch(urls[u], {
        headers: { 'Client-Id': creds.cid, 'Api-Key': creds.key }, muteHttpExceptions: true });
      lastCode = pdf.getResponseCode();
      if (lastCode === 200) return { blob: pdf.getBlob(), guid: guid, err: '' };
      Utilities.sleep(4000);
    }
  }
  return { blob: null, guid: guid, err: 'PDF не скачался (HTTP ' + lastCode + '). Ответ: ' + lastRaw };
}

/** Папка дня по артикулу: «Этикетки поставок <бренд>» / «ДД.ММ.ГГГГ — <артикул>». */
function artDayFolder_(art) {
  var root = labelsFolder_();
  var day = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Moscow', 'dd.MM.yyyy');
  var name = day + ' — ' + String(art).replace(/[\\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

/** Этикетки заявки, разложенные ПО АРТИКУЛАМ: на каждый артикул — свой PDF в папке
 *  «ДД.ММ.ГГГГ — <артикул>». Если нарезка недоступна (cargoes/get или labelCreate по
 *  cargo_ids не сработали) — общий PDF заявки кладётся в папку каждого её артикула. */
function saveLabelsByArt_(supplyId, creds, d, items) {
  var links = [], errs = [];
  var fileName = 'Этикетки ' + labelName_(d) + '.pdf';
  var artBySku = {};
  (items || []).forEach(function (it) { artBySku[String(it.sku)] = String(it.art); });

  // 1) грузоместа заявки → группировка cargo_id по артикулу
  var groups = null;
  try {
    Utilities.sleep(2000);
    var cg = ozonTry_(CFG_EP.cargoesGet, { supply_id: Number(supplyId), limit: 200 }, creds.cid, creds.key);
    var list = extractList_(cg.result || cg);
    var g = {};
    for (var i = 0; i < list.length; i++) {
      var c = list[i] || {};
      var cid = c.cargo_id || c.id || deepFind_(c, /cargo.*id/i);
      if (!cid) continue;
      var ci = c.items || (c.value && c.value.items) || [];
      var art = '';
      if (ci.length) art = artBySku[String(ci[0].sku || '')] || String(ci[0].offer_id || '') || '';
      if (!art) art = 'Прочее';
      if (!g[art]) g[art] = [];
      g[art].push(cid);
    }
    if (Object.keys(g).length) groups = g;
    else errs.push('cargoes/get: пустой список грузомест');
  } catch (e) { errs.push('cargoes/get: ' + String(e.message || e).slice(0, 200)); }

  if (groups) {
    for (var art2 in groups) {
      try {
        var res = labelPdfBlob_({ cargo_ids: groups[art2] }, creds);
        if (!res.blob) throw new Error(res.err);
        links.push(artDayFolder_(art2).createFile(res.blob.setName(fileName)).getUrl());
      } catch (e2) { errs.push(art2 + ': ' + String(e2.message || e2).slice(0, 200)); }
    }
    if (links.length) return { links: links, err: errs.join(' | ') };
  }

  // 2) фолбэк: общий PDF заявки в папку каждого артикула
  var res2 = labelPdfBlob_({ supply_id: Number(supplyId) }, creds);
  if (!res2.blob) return { links: [], err: errs.concat([res2.err]).join(' | ') };
  res2.blob.setName(fileName);
  var arts = {};
  (items || []).forEach(function (it) { arts[String(it.art)] = 1; });
  var artList = Object.keys(arts);
  if (!artList.length) artList = ['Прочее'];
  artList.forEach(function (a) { links.push(artDayFolder_(a).createFile(res2.blob).getUrl()); });
  return { links: links,
           err: (errs.length ? errs.join(' | ') + ' — ' : '') + 'сохранён общий PDF заявки (без нарезки по артикулам)' };
}

/** Меню: докачать этикетки по выделенной строке «Поставок» (без повторной передачи коробов). */
function downloadLabelsForActiveRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(REESTR);
  var row = s && s.getActiveRange() ? s.getActiveRange().getRow() : 0;
  if (!s || ss.getActiveSheet().getName() !== REESTR || row < 2) {
    alertSafe_('Этикетки', 'Выделите строку заявки на листе «' + REESTR + '».'); return;
  }
  var creds = ozonCreds_(); if (!creds) return;
  var d = s.getRange(row, 1, 1, REESTR_HEAD.length).getValues()[0];
  if (!d[5]) { alertSafe_('Этикетки', 'В строке нет Supply ID.'); return; }
  toastSafe_('Формирую этикетки…');
  var items2 = [];
  try { items2 = JSON.parse(d[12] || d[11] || '[]'); } catch (e) {}
  var lab = saveLabelsByArt_(d[5], creds, d, items2);
  s.getRange(row, 10).setValue(lab.links.length ? lab.links.join('\n') : (lab.err || ''));
  alertSafe_('Этикетки', lab.links.length
    ? lab.links.length + ' PDF, папки по артикулам' + (lab.err ? '\n(' + lab.err + ')' : '')
    : (lab.err || 'нет данных'));
}

/* ---------------- объединение PDF этикеток (pdf-lib с CDN) ---------------- */
var PDFLIB_URL = 'https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js';
function loadPdfLib_() {
  if (typeof PDFLib !== 'undefined') return;
  eval(UrlFetchApp.fetch(PDFLIB_URL).getContentText());
}

/** Склейка PDF-блобов в один файл. */
async function mergePdfBlobs_(blobs, name) {
  loadPdfLib_();
  var merged = await PDFLib.PDFDocument.create();
  for (var i = 0; i < blobs.length; i++) {
    var src = await PDFLib.PDFDocument.load(new Uint8Array(blobs[i].getBytes()));
    var pages = await merged.copyPages(src, src.getPageIndices());
    pages.forEach(function (p) { merged.addPage(p); });
  }
  var bytes = await merged.save();
  return Utilities.newBlob(Array.from(bytes), 'application/pdf', name);
}

/** Меню: в каждой сегодняшней папке артикула склеить все PDF в один файл
 *  «ВСЕ ЭТИКЕТКИ — <папка>.pdf». Старая склейка удаляется, файлы-источники остаются. */
async function mergeTodayLabelFolders() {
  var day = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Moscow', 'dd.MM.yyyy');
  var root = labelsFolder_();
  var folders = root.getFolders();
  var done = [];
  toastSafe_('Склеиваю PDF в папках за ' + day + '…');
  while (folders.hasNext()) {
    var f = folders.next();
    if (f.getName().indexOf(day) !== 0) continue;
    var blobs = [];
    var files = f.getFilesByType(MimeType.PDF);
    while (files.hasNext()) {
      var file = files.next();
      if (file.getName().indexOf('ВСЕ ЭТИКЕТКИ') >= 0) { file.setTrashed(true); continue; }
      blobs.push(file.getBlob());
    }
    if (!blobs.length) continue;
    try {
      var mergedBlob = await mergePdfBlobs_(blobs, 'ВСЕ ЭТИКЕТКИ — ' + f.getName() + '.pdf');
      f.createFile(mergedBlob);
      done.push('✓ ' + f.getName() + ' (' + blobs.length + ' PDF → 1)');
    } catch (e) {
      done.push('✗ ' + f.getName() + ' — ' + String(e.message || e).slice(0, 150));
    }
  }
  alertSafe_('Объединение этикеток',
    done.length ? done.join('\n') : 'Папок с PDF за ' + day + ' не нашлось.');
}

/* ================================================================
 *  СТАТУСЫ, СПРАВОЧНИКИ, ДИАГНОСТИКА
 * ================================================================ */
function updateSupplyStatuses() {
  var s = ensureReestr_();
  if (s.getLastRow() < 2) { toastSafe_('Реестр пуст.'); return; }
  var creds = ozonCreds_(); if (!creds) return;
  var n = s.getLastRow() - 1;
  var data = s.getRange(2, 1, n, REESTR_HEAD.length).getValues();
  var ids = [];
  data.forEach(function (r) { if (r[4]) ids.push(String(r[4])); });
  if (!ids.length) { toastSafe_('Нет Order ID в реестре.'); return; }
  for (var a = 0; a < ids.length; a += 50) {
    var og = ozonTry_(CFG_EP.orderGet, { order_ids: ids.slice(a, a + 50) }, creds.cid, creds.key);
    var orders = og.orders || (og.result && og.result.orders) || [];
    orders.forEach(function (o) {
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][4]) === String(o.supply_order_id || o.id)) {
          if (o.supply_order_number) s.getRange(i + 2, 4).setValue(o.supply_order_number);
          if (o.state) s.getRange(i + 2, 9).setValue(o.state);
          var sups = o.supplies || [];
          if (sups.length && (sups[0].supply_id || sups[0].id)) s.getRange(i + 2, 6).setValue(sups[0].supply_id || sups[0].id);
        }
      }
    });
  }
  toastSafe_('Статусы обновлены.');
}

function showDropoffWarehouses() {
  var creds = ozonCreds_(); if (!creds) return;
  var s = ensureSettings_();
  // поиск по названию точки отгрузки
  var q = '';
  try {
    var ui = SpreadsheetApp.getUi();
    var pr = ui.prompt('Поиск точки отгрузки',
      'Введите название или его часть (например: ХАБ или город).\nПусто — весь список.', ui.ButtonSet.OK_CANCEL);
    if (pr.getSelectedButton() !== ui.Button.OK) return;
    q = String(pr.getResponseText() || '').trim();
  } catch (e) {}
  // точки отгрузки Ozon — /v1/warehouse/fbo/list (seller/list вернул бы СВОИ склады)
  var r;
  try { r = ozonTry_(CFG_EP.dropoffPoints, { filter_by_supply_type: ['CREATE_TYPE_CROSSDOCK'], search: q }, creds.cid, creds.key); }
  catch (e) {
    try { r = ozonTry_(CFG_EP.dropoffPoints, { search: q }, creds.cid, creds.key); }
    catch (e2) { r = ozonTry_(CFG_EP.fboWarehouses, {}, creds.cid, creds.key); }
  }
  var list = extractList_(r);
  // доп. фильтр на нашей стороне
  if (q) {
    var qn = q.toLowerCase();
    var filtered = list.filter(function (w) { return JSON.stringify(w).toLowerCase().indexOf(qn) >= 0; });
    if (filtered.length) list = filtered;
  }
  var rows = [['— Справочник складов/точек отгрузки' + (q ? ' (поиск: ' + q + ')' : '') + ' —', '', '']];
  for (var i = 0; i < list.length && i < 300; i++) {
    var w = list[i];
    var id = w.warehouse_id || w.id || deepFind_(w, /warehouse.*id|dropoff.*id|point.*id/i) || '';
    var name = w.name || deepFind_(w, /(^|\.)name$/i) || '';
    rows.push([id, name, JSON.stringify(w).slice(0, 900)]);
  }
  if (rows.length === 1) rows.push(['(пусто — попробуйте другой запрос или проверьте права ключа)', '', '']);
  s.getRange(10, 1, Math.max(1, rows.length), 3).setValues(rows);
  toastSafe_('Справочник записан в «' + NAST + '» (с 10-й строки): ' + (rows.length - 1) + ' шт. ID — колонка A.');
}

/** Достаёт массив из ответа неизвестной формы. */
function extractList_(r) {
  if (!r) return [];
  if (Object.prototype.toString.call(r) === '[object Array]') return r;
  var keys = ['search', 'result', 'clusters', 'warehouses', 'items', 'list'];
  for (var i = 0; i < keys.length; i++) {
    var v = r[keys[i]];
    if (v) {
      if (Object.prototype.toString.call(v) === '[object Array]') return v;
      if (typeof v === 'object') { var inner = extractList_(v); if (inner.length) return inner; }
    }
  }
  return [];
}

/** Рекурсивный поиск первого скалярного значения, чей путь-ключ подходит под regex. */
function deepFind_(obj, re, path) {
  path = path || '';
  for (var k in obj) {
    var v = obj[k], p = path ? path + '.' + k : k;
    if (v !== null && typeof v === 'object') {
      var got = deepFind_(v, re, p);
      if (got !== '') return got;
    } else if (re.test(p) && v !== '' && v != null) {
      return v;
    }
  }
  return '';
}

/** РЕКОМЕНДУЕМЫЙ способ найти ID точки отгрузки: берём его из последних заявок,
 *  созданных вручную в кабинете (dropoff_warehouse_id). Самую частую точку
 *  скрипт сам подставит в «Настройки поставок», если поле пустое. */
function dropoffFromRecentOrders() {
  var creds = ozonCreds_(); if (!creds) return;
  var s = ensureSettings_();
  var bodies = [
    { filter: { states: [] }, paging: { from_supply_order_id: 0, limit: 100 } },
    { filter: { states: ['ORDER_STATE_DATA_FILLING', 'ORDER_STATE_READY_TO_SUPPLY', 'ORDER_STATE_IN_TRANSIT', 'ORDER_STATE_ACCEPTED_AT_SUPPLY_WAREHOUSE', 'ORDER_STATE_COMPLETED'] }, paging: { from_supply_order_id: 0, limit: 100 } },
    {}
  ];
  var ids = [], errs = [];
  for (var b = 0; b < bodies.length && !ids.length; b++) {
    try {
      var r = ozonTry_(CFG_EP.orderList, bodies[b], creds.cid, creds.key);
      ids = r.supply_order_id || (r.result && r.result.supply_order_id) || r.order_ids || [];
      if (!ids.length) errs.push('вариант ' + (b + 1) + ': ответ без заявок');
    } catch (e) { errs.push('вариант ' + (b + 1) + ': ' + String(e.message || e).slice(0, 160)); }
  }
  if (!ids.length) {
    alertSafe_('Точка отгрузки', 'Заявки не получены. Причины по вариантам запроса:\n\n' + errs.join('\n') +
      '\n\nЕсли везде 403 — у API-ключа нет прав на поставки FBO.\nАльтернатива: «Склады отгрузки (справочник)» → поиск по названию точки.');
    return;
  }
  var count = {}, whName = {};
  for (var a = 0; a < ids.length && a < 100; a += 50) {
    var og = ozonTry_(CFG_EP.orderGet, { order_ids: ids.slice(a, a + 50).map(String) }, creds.cid, creds.key);
    var orders = og.orders || (og.result && og.result.orders) || [];
    var whs = og.warehouses || (og.result && og.result.warehouses) || [];
    whs.forEach(function (w) {
      var wid = String(w.warehouse_id || w.id || '');
      if (wid) whName[wid] = (w.name || '') + (typeof w.address === 'string' && w.address ? ' | ' + w.address : '');
    });
    orders.forEach(function (o) {
      var d = o.dropoff_warehouse_id || deepFind_(o, /dropoff.*warehouse.*id/i);
      if (d) { d = String(d); count[d] = (count[d] || 0) + 1; }
    });
  }
  var arr = Object.keys(count).sort(function (x, y) { return count[y] - count[x]; });
  if (!arr.length) {
    alertSafe_('Точка отгрузки', 'В заявках не нашлось dropoff_warehouse_id — запустите «Склады отгрузки (справочник)» и пришлите Клоду строку из колонки C.');
    return;
  }
  var rows = [['— Точки отгрузки из ваших последних заявок —', '', '']];
  arr.forEach(function (id) { rows.push([id, whName[id] || '', 'использована в заявках: ' + count[id] + ' раз']); });
  s.getRange(10, 1, rows.length, 3).setValues(rows);
  if (!s.getRange(3, 2).getValue()) s.getRange(3, 2).setValue(Number(arr[0]) || arr[0]);
  alertSafe_('Точка отгрузки найдена',
    'Самая частая: ' + arr[0] + (whName[arr[0]] ? '\n' + whName[arr[0]] : '') +
    '\n\nЗаписал список в «' + NAST + '» (с 10-й строки) и подставил ID в настройку, если она была пустой.');
}

function supplyDiagnostics() {
  var creds = ozonCreds_(); if (!creds) return;
  var out = [];
  function check(name, fn) {
    try { var x = fn(); out.push('✓ ' + name + (x ? ' — ' + x : '')); }
    catch (e) { out.push('✗ ' + name + ' — ' + String(e.message || e).slice(0, 140)); }
  }
  check('Ключи / товары (product/list)', function () {
    var r = ozon_(API + '/v3/product/list', { filter: { visibility: 'ALL' }, limit: 1, last_id: '' }, creds.cid, creds.key);
    return 'ok';
  });
  check('Кластеры (cluster/list)', function () {
    var m = clusterMap_(creds.cid, creds.key);
    return Object.keys(m).length + ' кластеров';
  });
  check('Склады отгрузки (warehouse/fbo)', function () {
    var r = ozonTry_(CFG_EP.fboWarehouses, {}, creds.cid, creds.key);
    return 'ok';
  });
  check('Методы черновиков (draft/create/info)', function () {
    try { ozonTry_(CFG_EP.draftInfo, { operation_id: 'diagnostic-probe' }, creds.cid, creds.key); return 'доступен'; }
    catch (e) {
      var msg = String(e.message || e);
      if (msg.indexOf('403') >= 0) throw new Error('403 — у API-ключа нет прав на поставки FBO');
      if (msg.indexOf('404') >= 0) throw new Error('404 — метод не найден (сменилась версия?)');
      return 'доступен (ответил: ' + msg.slice(0, 60) + ')'; // 400 на фейковый id = метод жив
    }
  });
  alertSafe_('Диагностика API поставок', out.join('\n'));
}

function killTriggers_(fnName) {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++)
    if (ts[i].getHandlerFunction() === fnName) ScriptApp.deleteTrigger(ts[i]);
}
