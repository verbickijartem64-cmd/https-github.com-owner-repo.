const pptxgen = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');
const K = require('./content3.js');

const THEME = {
  name: 'CS616 Chernozem',
  headFontFace: 'Cambria',
  bodyFontFace: 'Calibri',
  colors: { dk1:'2A1E16', lt1:'FFFFFF', dk2:'5B4636', lt2:'F1F3F2',
    accent1:'1C7C8C', accent2:'D9A441', accent3:'B8552F', accent4:'8C6A4F', accent5:'6E7F80', accent6:'4E8A5B',
    hlink:'1C7C8C', folHlink:'8C6A4F' }
};
const H = THEME.colors; // hex — only for chart options that take hex
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = 'Когда датчик влажности «льстит» почве: CS616 в незасолённых почвах';
pres.author = 'Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е.';
pres.company = 'Кубанский ГАУ';
const C = pres.SchemeColor;
const DK = C.text1, MUTED = C.text2, WH = C.background1, LT = C.background2;
const TEAL = C.accent1, OCHRE = C.accent2, TERRA = C.accent3, BROWN = C.accent4, SLATE = C.accent5, GREEN = C.accent6;

const titlePh = (color) => ({ placeholder:{ options:{ name:'title', type:'title', x:0.6, y:0.36, w:12.1, h:1.05, fontSize:30, bold:true, color, align:'left', valign:'top', margin:0 }, text:'' } });
pres.defineSlideMaster({ title:'LIGHT', background:{ color: WH }, margin:[0.5, 0.6, 0.6, 0.6], objects:[titlePh(DK)],
  slideNumber:{ x:12.03, y:7.02, w:0.7, h:0.3, fontSize:11, color:MUTED, align:'right' } });
pres.defineSlideMaster({ title:'DARK', background:{ color: DK }, margin:[0.5, 0.6, 0.6, 0.6], objects:[titlePh(WH)],
  slideNumber:{ x:12.03, y:7.02, w:0.7, h:0.3, fontSize:11, color:OCHRE, align:'right' } });
pres.defineSlideMaster({ title:'COVER', background:{ color: DK }, margin:[0.5, 0.6, 0.6, 0.6],
  objects:[{ placeholder:{ options:{ name:'title', type:'title', x:0.7, y:1.15, w:7.9, h:2.0, fontSize:40, bold:true, color:WH, align:'left', valign:'top', margin:0 }, text:'' } }] });

// ---------- helpers ----------
const fmt = n => String(n).replace('.', ',');
const wc = s => s.split(/\s+/).filter(Boolean).length;
function T(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({ x, y, w, h, fontSize:16, color:DK, margin:0, valign:'top', isTextBox:true }, o));
}
function card(s, x, y, w, h, fill = LT, name = 'card', o = {}) {
  s.addShape(pres.ShapeType.roundRect, Object.assign({ x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, rectRadius:0.08, objectName:name }, o));
}
function circ(s, x, y, d, fill, label, fs = 18, color = WH) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w:d, h:d, fill:{ color: fill }, line:{ type:'none' }, objectName:'badge' });
  if (label) s.addText(label, { x, y, w:d, h:d, align:'center', valign:'middle', fontSize:fs, bold:true, color, margin:0, isTextBox:true });
}
// straight line between two points; flips instead of negative sizes
function L(s, x1, y1, x2, y2, color, o = {}) {
  s.addShape(pres.ShapeType.line, { x:Math.min(x1, x2), y:Math.min(y1, y2), w:Math.abs(x2 - x1), h:Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
    line:Object.assign({ color, width:2 }, o), objectName:o.name || 'line' });
}
function arrowDown(s, x, y, w = 0.4, h = 0.3, fill = SLATE) {
  s.addShape(pres.ShapeType.downArrow, { x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, objectName:'arrow' });
}
function arrowRight(s, x, y, w = 0.3, h = 0.4, fill = SLATE) {
  s.addShape(pres.ShapeType.rightArrow, { x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, objectName:'arrow' });
}
function tracker(s, act, dark) {
  if (!act) return;
  const runs = [];
  K.ACTS.forEach((a, i) => {
    const on = a === act;
    runs.push({ text:a, options:{ bold:on, color:on ? (dark ? OCHRE : TEAL) : (dark ? LT : SLATE) } });
    if (i < K.ACTS.length - 1) runs.push({ text:'   ·   ', options:{ color:dark ? LT : SLATE } });
  });
  T(s, runs, 0.6, 7.02, 10.5, 0.3, { fontSize:11, valign:'middle' });
}
const KIND = { core:'основной', deep:'углубление: при нехватке времени — короткая версия', optional:'можно пропустить' };
let curSection = null;
function addSlide(n, section) {
  const d = K.slides[n - 1];
  if (section) { pres.addSection({ title: section }); curSection = section; }
  const master = n === 1 ? 'COVER' : (d.master || 'LIGHT');
  const s = pres.addSlide({ masterName: master, sectionTitle: curSection });
  s.addText(d.title, { placeholder:'title' });
  const mins = Math.max(0.5, Math.round(wc(d.text.join(' ')) / 145 * 2) / 2);
  s.addNotes(`[≈ ${fmt(mins)} мин · ${KIND[d.kind]}]\n\nГлавная мысль: ${d.takeaway}\n\n` + d.text.join('\n\n') + `\n\nЕсли времени мало: ${d.short}`);
  tracker(s, d.act, master !== 'LIGHT');
  return s;
}
const chartBase = { catAxisLabelColor:H.dk2, valAxisLabelColor:H.dk2, catAxisLabelFontSize:12, valAxisLabelFontSize:12,
  catAxisLabelFontFace:'+mn-lt', valAxisLabelFontFace:'+mn-lt', legendFontFace:'+mn-lt', legendFontSize:12, legendColor:H.dk2,
  valGridLine:{ color:'D9DEDC', size:0.5 }, catGridLine:{ style:'none' }, dataLabelFontFace:'+mn-lt', titleFontFace:'+mn-lt',
  catAxisTitleFontFace:'+mn-lt', valAxisTitleFontFace:'+mn-lt', catAxisTitleFontSize:12, valAxisTitleFontSize:12,
  catAxisTitleColor:H.dk2, valAxisTitleColor:H.dk2, titleFontSize:14, titleColor:H.dk1 };
const cb = o => Object.assign({}, chartBase, o);
const TV = { // Table V of the article: Δθ at six θ levels, ρb = 1.35
  lv:['0,05', '0,10', '0,20', '0,30', '0,38', '0,40'],
  A:[0.000, 0.001, 0.003, 0.005, 0.010, 0.011], B:[0.001, 0.002, 0.005, 0.009, 0.017, 0.019], C:[0.001, 0.002, 0.006, 0.012, 0.022, 0.027] };

// ================= 1. cover =================
{
  const s = addSlide(1, 'Зачем');
  T(s, 'Систематическая погрешность заводской калибровки датчиков влажности CS616 в незасолённых почвах разного гранулометрического состава', 0.7, 3.5, 7.5, 1.3, { fontSize:18, color:OCHRE, italic:true });
  T(s, 'Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е.', 0.7, 5.15, 7.5, 0.4, { fontSize:15, color:WH, bold:true });
  T(s, 'Кубанский государственный аграрный университет имени И. Т. Трубилина · Краснодар, 2026', 0.7, 5.6, 7.5, 0.7, { fontSize:14, color:WH });
  const sx = 9.0;
  [['A', TEAL, 'чернозём обыкновенный'], ['B', BROWN, 'чернозём типичный'], ['C', TERRA, 'лугово-чернозёмная']].forEach(([l, col, nm], k) => {
    card(s, sx, 2.25 + k * 1.35, 3.6, 1.2, col, 'soil-' + l);
    T(s, l, sx + 0.25, 2.25 + k * 1.35, 0.6, 1.2, { fontSize:36, bold:true, color:WH, valign:'middle' });
    T(s, nm, sx + 0.95, 2.25 + k * 1.35, 1.4, 1.2, { fontSize:13, color:WH, valign:'middle' });
  });
  s.addShape(pres.ShapeType.roundRect, { x:sx + 2.4, y:1.2, w:0.9, h:0.65, fill:{ color: WH }, line:{ type:'none' }, rectRadius:0.06, objectName:'probe-head' });
  T(s, 'CS616', sx + 2.4, 1.2, 0.9, 0.65, { fontSize:12, bold:true, align:'center', valign:'middle' });
  [2.55, 3.15].forEach(dx => s.addShape(pres.ShapeType.rect, { x:sx + dx - 0.04, y:1.85, w:0.08, h:4.3, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
}

// ================= 2. main message =================
{
  const s = addSlide(2);
  const st = [['≤ 0,014', 'наибольшая типичная ошибка (RMSE) среди шести вариантов — ниже критерия 0,02', TEAL, 44],
              ['+0,006…+0,010', 'среднее завышение влажности в тяжёлых почвах — всегда в одну сторону', TERRA, 32],
              ['× 2,6', 'во столько раз быстрее растёт ошибка с увлажнением в самой тяжёлой почве, чем в лёгкой', BROWN, 44]];
  st.forEach((r, k) => {
    const x = 0.6 + k * 4.1;
    card(s, x, 1.6, 3.9, 2.6, LT, 'stat-' + k);
    T(s, r[0], x + 0.25, 1.75, 3.4, 0.95, { fontSize:r[3], bold:true, color:r[2], valign:'middle' });
    T(s, r[1], x + 0.25, 2.8, 3.4, 1.3, { fontSize:16 });
  });
  card(s, 0.6, 4.4, 12.1, 1.1, DK, 'conclusion');
  T(s, [{ text:'Вывод: ', options:{ bold:true, color:OCHRE } }, { text:'в тяжёлых почвах поправка к датчику должна зависеть от влажности — постоянное число не подойдёт', options:{ color:WH } }], 0.9, 4.4, 11.5, 1.1, { fontSize:20, valign:'middle' });
  T(s, 'Маршрут доклада', 0.6, 5.68, 4, 0.3, { fontSize:12, bold:true, color:MUTED });
  K.ACTS.forEach((a, k) => s.addText(`${k + 1}. ${a}`, { shape:pres.ShapeType.chevron, x:0.6 + k * 2.41, y:6.0, w:2.46, h:0.8, fill:{ color: k % 2 ? LT : 'E3E8E6' }, line:{ type:'none' }, fontSize:15, bold:true, color:DK, align:'center', valign:'middle', margin:0, objectName:'route' }));
}

// ================= 3. cost of 1% =================
{
  const s = addSlide(3);
  [['0,01', '1 л воды на 100 л почвы'], ['10 мм', 'воды в метровом слое'], ['100 м³', 'воды на гектар']].forEach((r, k) => {
    const y = 1.6 + k * 1.15;
    card(s, 0.6, y, 5.6, 1.0, LT, 'conv-' + k);
    T(s, r[0], 0.85, y, 2.1, 1.0, { fontSize:34, bold:true, color:TEAL, valign:'middle' });
    T(s, r[1], 3.05, y, 3.0, 1.0, { fontSize:18, bold:true, valign:'middle' });
  });
  card(s, 0.6, 5.1, 5.6, 1.75, DK, 'passport');
  T(s, [{ text:'Паспорт: ±0,025', options:{ bold:true, color:OCHRE, fontSize:22, breakLine:true } }, { text:'это до 25 мм воды в метровом слое; точность установлена в основном на супесях и более лёгких почвах', options:{ color:WH, fontSize:16 } }], 0.85, 5.18, 5.1, 1.6, { valign:'middle' });
  // schematic
  T(s, 'Почему завышение опасно для полива (схема)', 6.6, 1.55, 6.1, 0.4, { fontSize:16, bold:true, color:MUTED });
  L(s, 6.9, 5.45, 6.9, 2.0, SLATE, { width:1.5, endArrowType:'triangle' });
  L(s, 6.9, 5.45, 12.65, 5.45, SLATE, { width:1.5, endArrowType:'triangle' });
  T(s, 'влажность', 5.95, 3.2, 1.6, 0.3, { fontSize:12, color:MUTED, rotate:270, align:'center' });
  T(s, 'время: почва подсыхает →', 9.6, 5.5, 3.1, 0.3, { fontSize:12, color:MUTED, align:'right' });
  L(s, 6.9, 4.3, 12.5, 4.3, OCHRE, { width:2.25, dashType:'dash', name:'threshold' });
  T(s, 'порог полива', 7.0, 3.95, 1.8, 0.3, { fontSize:13, bold:true, color:BROWN });
  L(s, 7.0, 2.6, 12.3, 5.3, TEAL, { width:3, name:'true-theta' });
  L(s, 7.0, 2.25, 12.3, 5.15, TERRA, { width:3, dashType:'dash', name:'sensor-theta' });
  circ(s, 10.237, 4.2, 0.2, TEAL);
  circ(s, 10.646, 4.2, 0.2, TERRA);
  T(s, 'пора поливать', 8.4, 4.42, 1.75, 0.3, { fontSize:13, bold:true, color:TEAL, align:'right' });
  T(s, 'датчик разрешает полив', 10.85, 3.8, 1.85, 0.45, { fontSize:13, bold:true, color:TERRA });
  L(s, 9.3, 2.15, 9.8, 2.15, TEAL, { width:3 });
  T(s, 'фактическая влажность (эталон)', 9.9, 2.0, 2.8, 0.3, { fontSize:12, color:DK, valign:'middle' });
  L(s, 9.3, 2.5, 9.8, 2.5, TERRA, { width:3, dashType:'dash' });
  T(s, 'показания датчика', 9.9, 2.35, 2.8, 0.3, { fontSize:12, color:DK, valign:'middle' });
  card(s, 6.6, 5.9, 6.1, 0.95, LT, 'real-numbers');
  T(s, [{ text:'Почва C, наш опыт: ', options:{ bold:true } }, { text:'эталон 0,40 → датчик 0,427; эталон 0,30 → датчик 0,312' }], 6.8, 5.9, 5.8, 0.95, { fontSize:16, valign:'middle' });
}

// ================= 4. sensor measures time =================
{
  const s = addSlide(4);
  card(s, 0.6, 1.6, 5.8, 1.95, BROWN, 'soil-block');
  s.addShape(pres.ShapeType.roundRect, { x:0.95, y:1.6, w:1.3, h:0.5, fill:{ color: DK }, line:{ type:'none' }, rectRadius:0.05, objectName:'probe-head' });
  T(s, 'датчик', 0.95, 1.6, 1.3, 0.5, { fontSize:13, color:WH, align:'center', valign:'middle' });
  [1.2, 2.0].forEach(x => s.addShape(pres.ShapeType.rect, { x, y:2.1, w:0.07, h:1.3, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
  L(s, 2.6, 2.2, 6.0, 2.2, WH, { width:2.5, dashType:'dash', endArrowType:'triangle' });
  L(s, 6.0, 3.0, 2.6, 3.0, OCHRE, { width:2.5, dashType:'dash', endArrowType:'triangle' });
  T(s, 'импульс бежит по штангам и возвращается', 2.6, 2.35, 3.6, 0.5, { fontSize:14, color:WH });
  T(s, 'почва', 5.3, 3.1, 1.0, 0.35, { fontSize:13, color:WH, align:'right' });
  s.addChart(pres.charts.BAR, [{ name:'√ε', labels:['Воздух', 'Минералы', 'Вода'], values:[1, 2.17, 8.94] }], cb({
    x:0.6, y:3.75, w:5.8, h:3.1, barDir:'bar', chartColors:[H.accent1], showValue:true, dataLabelFormatCode:'0.0', dataLabelColor:H.dk1, dataLabelFontSize:14, dataLabelPosition:'outEnd',
    showTitle:true, title:'Во сколько раз волна медленнее, чем в воздухе (√ε)', showLegend:false, valAxisHidden:true, valGridLine:{ style:'none' }, barGapWidthPct:40, catAxisOrientation:'maxMin', catAxisLabelFontSize:14 }));
  ['Больше воды в почве', 'Волна идёт медленнее', 'Период τ длиннее'].forEach((c, k) => {
    const y = 1.6 + k * 1.1;
    card(s, 6.8, y, 5.9, 0.8, k === 2 ? TEAL : LT, 'chain-' + k);
    T(s, c, 7.1, y, 5.3, 0.8, { fontSize:20, bold:true, color:k === 2 ? WH : DK, valign:'middle' });
    if (k < 2) arrowDown(s, 9.55, y + 0.79, 0.4, 0.3);
  });
  card(s, 6.8, 4.95, 5.9, 1.9, DK, 'formula');
  T(s, 'Заводской «переводчик»: время → влажность', 7.1, 5.07, 5.3, 0.35, { fontSize:15, color:OCHRE, bold:true });
  T(s, 'θ = −0,0663 − 0,0063·τ + 0,0007·τ²', 7.1, 5.5, 5.3, 0.55, { fontSize:22, bold:true, color:WH });
  T(s, 'ε: воздух 1 · минералы 4,7 · вода 80; τ — в микросекундах', 7.1, 6.15, 5.3, 0.5, { fontSize:14, color:WH });
}

// ================= 5. two assumptions =================
{
  const s = addSlide(5);
  [['1', 'Почва для волны — изолятор', 'Пробег удлиняет только вода; по дороге сигнал не теряет энергию.'], ['2', 'Переводчик универсален', 'Уравнение, настроенное на одни почвы, верно и для других.']].forEach((c, k) => {
    const x = 0.6 + k * 6.2;
    card(s, x, 1.6, 5.9, 2.3, LT, 'assumption-' + k);
    circ(s, x + 0.3, 1.88, 0.75, k ? TERRA : TEAL, c[0], 24);
    T(s, c[1], x + 1.3, 1.83, 4.4, 0.85, { fontSize:22, bold:true, valign:'middle' });
    T(s, c[2], x + 0.3, 2.85, 5.3, 0.9, { fontSize:18 });
  });
  T(s, 'Где производитель гарантирует стандартную шкалу', 0.6, 4.15, 8, 0.4, { fontSize:16, bold:true, color:MUTED });
  [['< 0,5 дСм/м', 'объёмная электропроводность почвы'], ['< 1,55 г/см³', 'плотность сложения'], ['< 30 %', 'глины (частицы мельче 0,002 мм)']].forEach((c, k) => {
    const x = 0.6 + k * 4.1;
    card(s, x, 4.6, 3.9, 1.3, WH, 'limit-' + k, { line:{ color: SLATE, width:1.25 } });
    T(s, c[0], x + 0.25, 4.68, 3.4, 0.65, { fontSize:28, bold:true, color:TEAL, valign:'middle' });
    T(s, c[1], x + 0.25, 5.35, 3.4, 0.45, { fontSize:15, color:MUTED });
  });
  T(s, [{ text:'Для наших тяжёлых почв это не гарантировано: ', options:{ bold:true, color:TERRA } }, { text:'физической глины до 65 %, а глину < 0,002 мм и объёмную электропроводность мы не измеряли' }], 0.6, 6.12, 12.1, 0.75, { fontSize:18 });
}

// ================= 6. mechanism =================
{
  const s = addSlide(6);
  card(s, 0.6, 1.6, 6.0, 3.75, LT, 'dl-card');
  const cx = 2.2, cy = 3.45;
  circ(s, cx - 0.75, cy - 0.75, 1.5, BROWN, '− − −\n− − −', 16);
  for (let k = 0; k < 10; k++) {
    const a = k * Math.PI / 5;
    circ(s, cx + 1.05 * Math.cos(a) - 0.16, cy + 1.05 * Math.sin(a) - 0.16, 0.32, TEAL, '+', 14);
  }
  T(s, 'Частица глины несёт заряд (−); вокруг неё — облако подвижных ионов (+), двойной электрический слой. Он проводит ток.', 3.75, 1.8, 2.7, 1.9, { fontSize:15 });
  T(s, 'Поверхность смектита 600–800 м²/г — на порядки больше, чем у песка (сахар и сахарная пудра)', 3.75, 3.8, 2.7, 1.4, { fontSize:15, bold:true, color:TERRA });
  const t = Array.from({ length:25 }, (_, i) => i);
  s.addChart(pres.charts.LINE, [
    { name:'сильный сигнал (почва-изолятор)', labels:t.map(() => ''), values:t.map(v => +(1 - Math.exp(-v / 2.5)).toFixed(3)) },
    { name:'ослабленный сигнал (проводящая почва)', labels:t.map(() => ''), values:t.map(v => +(0.6 * (1 - Math.exp(-v / 4))).toFixed(3)) },
    { name:'порог срабатывания', labels:t.map(() => ''), values:t.map(() => 0.45) } ],
    cb({ x:6.8, y:1.6, w:5.9, h:2.95, chartColors:[H.accent1, H.accent3, H.accent2], lineSize:3, lineDataSymbol:'none', showLegend:true, legendPos:'b',
      valAxisHidden:true, catAxisHidden:true, valGridLine:{ style:'none' }, valAxisMinVal:0, valAxisMaxVal:1.05, showTitle:true, title:'Схема: прибор срабатывает, когда сигнал пересекает порог' }));
  card(s, 6.8, 4.65, 5.9, 0.7, TEAL, 'mech-result');
  T(s, 'Слабый сигнал пересекает порог позже → τ длиннее → «лишняя вода»', 7.0, 4.65, 5.5, 0.7, { fontSize:16, bold:true, color:WH, valign:'middle' });
  card(s, 0.6, 5.55, 12.1, 1.3, WH, 'counter', { line:{ color: OCHRE, width:1.5 } });
  T(s, [{ text:'Обратная сила: ', options:{ bold:true, color:TERRA } }, { text:'вода, связанная с поверхностью, тормозит волну слабее свободной и тянет показания вниз. Кто сильнее — решает опыт. Наше предположение: завышение, растущее с влажностью.' }], 0.9, 5.55, 11.5, 1.3, { fontSize:17, valign:'middle' });
}

// ================= 7. theory map =================
{
  const s = addSlide(7);
  const node = (x, y, head, sub, fill = LT, dark = false) => {
    card(s, x, y, 2.1, 1.4, fill, 'node');
    T(s, [{ text:head, options:{ bold:true, fontSize:15, color:dark ? WH : DK, breakLine:true } }, { text:sub, options:{ fontSize:14, color:dark ? WH : MUTED } }], x + 0.12, y + 0.05, 1.86, 1.3, { valign:'middle' });
  };
  node(0.6, 1.6, 'Физика почв', 'доли фаз, плотность, набухание');
  node(0.6, 3.35, 'Химия глин', 'двойной слой, соли; прокси: MH, ECe');
  node(3.1, 1.6, 'Теория смеси (CRIM)', '√ε смеси = Σ доля × √ε');
  node(3.1, 3.35, 'Электро­динамика', 'скорость ∝ 1/√ε; потери сигнала');
  node(5.6, 2.475, 'Датчик + уравнение', 'τ → θ: эмпирический перевод', DK, true);
  node(8.1, 1.6, 'Метрология', 'эталон U = 0,005; MB, RMSE');
  node(8.1, 3.35, 'Статистика', 'дисперсионный анализ, блоки, регрессия');
  node(10.6, 2.475, 'Мелиорация', 'решение: порог полива', TEAL, true);
  const A = { endArrowType:'triangle', width:2.5 };
  L(s, 2.7, 2.3, 3.1, 2.3, OCHRE, A);
  L(s, 2.7, 4.05, 3.1, 4.05, TEAL, A);
  L(s, 4.15, 3.35, 4.15, 3.0, OCHRE, A);
  L(s, 5.2, 2.3, 5.6, 2.9, OCHRE, A);
  L(s, 5.2, 4.05, 5.6, 3.45, TEAL, A);
  L(s, 7.7, 2.9, 8.1, 2.3, SLATE, A);
  L(s, 9.15, 3.0, 9.15, 3.35, SLATE, A);
  L(s, 10.2, 4.05, 10.6, 3.45, SLATE, A);
  const leg = [[TEAL, 'Механизм', 'причина → следствие. Подводит, если есть обратная сила (связанная вода)', null],
               [OCHRE, 'Модель', 'даёт число, его можно проверить. Подводит, если нарушены допущения (плотность)', null],
               [TERRA, 'Косвенный показатель', 'судим о скрытом по видимому (MH, ECe). Подводит, если прокси ходят парой', 'dash'],
               [SLATE, 'Статистический вывод', 'от колонок к общему. Подводит, если наблюдения не независимы', null]];
  leg.forEach((g, k) => {
    const x = 0.6 + k * 3.07;
    card(s, x, 5.0, 2.92, 1.85, LT, 'legend-' + k);
    L(s, x + 0.2, 5.3, x + 0.8, 5.3, g[0], { width:3, dashType:g[3] || 'solid' });
    T(s, g[1], x + 0.9, 5.14, 1.95, 0.35, { fontSize:15, bold:true, valign:'middle' });
    T(s, g[2], x + 0.2, 5.58, 2.6, 1.22, { fontSize:14 });
  });
}

// ================= 8. terms trap =================
{
  const s = addSlide(8);
  T(s, 'Размер частиц, мм (в масштабе)', 0.6, 1.55, 6, 0.35, { fontSize:14, bold:true, color:MUTED });
  const scale = 900; // inches per mm
  [['глина по Качинскому < 0,001', 0.001, SLATE], ['«clay» по USDA и в паспорте < 0,002', 0.002, TEAL], ['физическая глина по Качинскому < 0,01', 0.01, TERRA]].forEach((b, k) => {
    const y = 2.0 + k * 0.95;
    s.addShape(pres.ShapeType.rect, { x:0.6, y, w:b[1] * scale, h:0.6, fill:{ color: b[2] }, line:{ type:'none' }, objectName:'size-bar' });
    if (b[1] > 0.005) T(s, b[0], 0.85, y, 8, 0.6, { fontSize:17, bold:true, valign:'middle', color:WH });
    else T(s, b[0], 0.6 + b[1] * scale + 0.2, y, 7, 0.6, { fontSize:17, bold:true, valign:'middle' });
  });
  card(s, 10.0, 1.6, 2.7, 1.85, DK, 'gap');
  T(s, 'Для чернозёмов края количественных оценок не было', 10.2, 1.6, 2.3, 1.85, { fontSize:16, bold:true, color:WH, valign:'middle' });
  card(s, 0.6, 4.9, 5.9, 1.95, LT, 'trap');
  T(s, [{ text:'Ловушка перевода', options:{ bold:true, color:TERRA, breakLine:true } }, { text:'30 % «clay» из паспорта и 30 % физической глины — разные почвы. «Лёгкая глина» по Качинскому ≠ «clay» по USDA.' }], 0.85, 5.0, 5.4, 1.75, { fontSize:17 });
  card(s, 6.8, 4.9, 5.9, 1.95, LT, 'consequence');
  T(s, [{ text:'Что из этого следует', options:{ bold:true, color:TEAL, breakLine:true } }, { text:'В группе A физической глины ≤ 20 %, значит, и USDA-глины ≤ 20 %: условие паспорта выполнено. Для B и C проверить нельзя.' }], 7.05, 5.0, 5.4, 1.75, { fontSize:17 });
}

// ================= 9. hypotheses =================
{
  const s = addSlide(9);
  [['H1', 'Ошибка растёт с содержанием физической глины', 'Опровергнет: ошибка одинакова в лёгких и тяжёлых почвах', TEAL],
   ['H2', 'Ошибка растёт с влажностью, всё быстрее к насыщению', 'Опровергнет: ошибка постоянна или растёт по прямой', TERRA],
   ['H3', 'Плотность влияет, и по-разному в разных почвах', 'Опровергнет: плотность не меняет ошибку или меняет её одинаково везде', BROWN]].forEach((h, k) => {
    const y = 1.6 + k * 1.6;
    card(s, 0.6, y, 12.1, 1.4, LT, 'h-' + k);
    circ(s, 0.9, y + 0.33, 0.75, h[3], h[0], 20);
    T(s, h[1], 2.0, y + 0.12, 10.4, 0.7, { fontSize:22, bold:true, valign:'middle' });
    T(s, h[2], 2.0, y + 0.82, 10.4, 0.45, { fontSize:17, color:MUTED, italic:true });
  });
  T(s, 'Каждое предсказание запрещает часть результатов. Итоги — на последнем слайде.', 0.6, 6.45, 12.1, 0.45, { fontSize:17, bold:true, color:TEAL });
}

// ================= 10. experiment =================
{
  const s = addSlide(10, 'Как проверяли');
  T(s, 'Схема опыта: 3 почвы × 2 плотности', 0.6, 1.5, 6, 0.35, { fontSize:14, bold:true, color:MUTED });
  [['A', '10–20 % физ. глины', TEAL], ['B', '40–50 %', BROWN], ['C', '50–65 %', TERRA]].forEach((c, k) => {
    T(s, [{ text:c[0], options:{ bold:true, fontSize:20, breakLine:true } }, { text:c[1], options:{ fontSize:13 } }], 1.9 + k * 1.7, 1.95, 1.6, 0.72, { align:'center' });
    [1.35, 1.52].forEach((d, r) => {
      card(s, 1.9 + k * 1.7, 2.75 + r * 0.85, 1.6, 0.75, c[2], 'cell');
      T(s, fmt(d) + ' г/см³', 1.9 + k * 1.7, 2.75 + r * 0.85, 1.6, 0.75, { fontSize:15, bold:true, color:WH, align:'center', valign:'middle' });
    });
  });
  T(s, 'плотность', 0.6, 2.75, 1.2, 0.75, { fontSize:14, color:MUTED, valign:'middle' });
  T(s, 'все почвы незасолённые: ECe 0,28–0,61 дСм/м', 1.9, 4.5, 5.1, 0.4, { fontSize:14, color:MUTED });
  card(s, 7.3, 1.5, 5.4, 1.45, DK, 'count');
  T(s, '6 × 3 × 10 = 180', 7.55, 1.55, 4.9, 0.8, { fontSize:36, bold:true, color:WH, valign:'middle' });
  T(s, 'вариантов × колонок × ступеней влажности', 7.55, 2.35, 4.9, 0.5, { fontSize:15, color:OCHRE });
  [['Датчик с заводским уравнением', '= студент', TEAL], ['Термостатно-весовой метод', '= ключ с ответами (неопределённость 0,005)', TERRA]].forEach((r, k) => {
    const y = 3.15 + k * 0.95;
    card(s, 7.3, y, 5.4, 0.8, LT, 'exam-' + k);
    T(s, [{ text:r[0], options:{ bold:true, fontSize:16, breakLine:true } }, { text:r[1], options:{ fontSize:15, color:r[2], bold:true } }], 7.5, y, 5.0, 0.8, { valign:'middle' });
  });
  ['Насыщение снизу, 48 ч', 'Ступенчатое иссушение при 20 °C', 'Ожидание равновесия 8–36 ч', 'Снимок: период τ + масса колонки'].forEach((t, k) => {
    const x = 0.6 + k * 3.1;
    card(s, x, 5.3, 2.85, 1.3, k === 3 ? TEAL : LT, 'proc-' + k);
    T(s, t, x + 0.2, 5.3, 2.45, 1.3, { fontSize:17, bold:true, color:k === 3 ? WH : DK, valign:'middle' });
    if (k < 3) arrowRight(s, x + 2.88, 5.78, 0.24, 0.35);
  });
  T(s, 'Колонка 152 мм × 300 мм, три слоя уплотнения. Каждое решение убирает одну постороннюю причину.', 0.6, 6.68, 12.1, 0.3, { fontSize:13, color:MUTED });
}

// ================= 11. target =================
{
  const s = addSlide(11);
  const cx = 3.4, cy = 4.15;
  [2.3, 1.55, 0.8].forEach((r, k) => s.addShape(pres.ShapeType.ellipse, { x:cx - r, y:cy - r, w:2 * r, h:2 * r, fill:{ color: k % 2 ? WH : LT }, line:{ color: SLATE, width:1 }, objectName:'ring' }));
  circ(s, cx - 0.07, cy - 0.07, 0.14, DK);
  const gx = cx + 0.75, gy = cy - 0.55;
  [[0, 0], [0.3, -0.2], [-0.25, 0.3], [0.2, 0.35], [-0.3, -0.15], [0.45, 0.05], [-0.05, -0.4]].forEach(([dx, dy]) => circ(s, gx + dx - 0.09, gy + dy - 0.09, 0.18, TERRA));
  L(s, cx, cy, gx, gy, DK, { width:2.5, endArrowType:'triangle' });
  T(s, 'MB — сдвиг центра кучки от яблочка', 0.9, 6.55, 5.0, 0.35, { fontSize:15, bold:true, color:TERRA, align:'center' });
  card(s, 6.9, 1.6, 5.8, 1.5, DK, 'identity');
  T(s, 'RMSE² = MB² + разброс²', 7.15, 1.6, 5.3, 1.0, { fontSize:30, bold:true, color:WH, valign:'middle' });
  T(s, 'типичный промах складывается из сдвига и разброса', 7.15, 2.55, 5.3, 0.45, { fontSize:15, color:OCHRE });
  card(s, 6.9, 3.3, 5.8, 1.3, LT, 'mb0');
  T(s, [{ text:'MB ≈ 0: ', options:{ bold:true, color:TEAL } }, { text:'ошибки случайны, то выше, то ниже — при усреднении гасятся' }], 7.15, 3.3, 5.3, 1.3, { fontSize:17, valign:'middle' });
  card(s, 6.9, 4.8, 5.8, 1.3, LT, 'mbpos');
  T(s, [{ text:'MB > 0: ', options:{ bold:true, color:TERRA } }, { text:'систематическое завышение — усреднение не поможет; опасно для порога полива' }], 7.15, 4.8, 5.3, 1.3, { fontSize:17, valign:'middle' });
  T(s, 'Критерий приемлемости: RMSE ≤ 0,02', 6.9, 6.3, 5.8, 0.45, { fontSize:20, bold:true, color:TEAL });
}

// ================= 12. our own assumptions =================
{
  const s = addSlide(12);
  const hd = { bold:true, color:WH, fill:{ color: DK }, fontSize:14, valign:'middle' };
  const c = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:14, valign:'middle', color:DK, fill:{ color: LT } }, o) });
  const st = (t, col, txt = WH) => ({ text:t, options:{ fontSize:14, bold:true, align:'center', valign:'middle', color:txt, fill:{ color: col } } });
  const rows = [
    [{ text:'Допущение', options:hd }, { text:'Что мы сделали или знаем', options:hd }, { text:'Статус', options:Object.assign({}, hd, { align:'center' }) }],
    [c('«Эталон — это истина»', { bold:true }), c('Два способа (вся колонка и микрокерны): расхождение ≤ 0,008, в среднем 0,003; U = 0,005'), st('частично', OCHRE, DK)],
    [c('«Десять уровней — независимые наблюдения»', { bold:true }), c('Нет: это последовательные состояния одних и тех же колонок'), st('p описательные', TERRA)],
    [c('«Одна почва представляет группу»', { bold:true }), c('В каждой группе одна почва'), st('экстраполяция', TERRA)],
    [c('«Колонка не мешает датчику»', { bold:true }), c('До стенки 59–73 мм (рекомендуется ≥ 100), штанги доходят до дна'), st('не проверено', TERRA)],
    [c('«Объём почвы постоянен»', { bold:true }), c('Нарушено при набухании (почва C, 1,52 г/см³) — эталон по реальному объёму'), st('учтено', GREEN)],
    [c('«Округление опубликованных чисел не меняет выводов»', { bold:true }), c('20 000 симуляций: доля эффекта почвы 0,07–0,11 держится; граница A–B неустойчива'), st('проверено', GREEN)]
  ];
  s.addTable(rows, { x:0.6, y:1.6, w:12.1, colW:[3.6, 6.0, 2.5], rowH:[0.45, 0.72, 0.72, 0.72, 0.72, 0.72, 0.72], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.08, objectName:'own-assumptions' });
  T(s, 'Поэтому дальше: «описательно» и «в изученных почвах»', 0.6, 6.5, 12.1, 0.4, { fontSize:18, bold:true, color:TEAL });
}

// ================= 13. result 1 =================
{
  const s = addSlide(13, 'Что нашли');
  const labels = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2'];
  s.addChart([
    { type:pres.charts.BAR, data:[{ name:'MB (средний сдвиг)', labels, values:[0.004, 0.002, 0.007, 0.006, 0.010, 0.008] }, { name:'RMSE (типичный промах)', labels, values:[0.007, 0.006, 0.011, 0.009, 0.014, 0.012] }],
      options:{ barDir:'col', chartColors:[H.accent4, H.accent1], barGapWidthPct:60, showValue:true, dataLabelFormatCode:'0.000', dataLabelFontSize:11, dataLabelColor:H.dk1, dataLabelPosition:'outEnd' } },
    { type:pres.charts.LINE, data:[{ name:'Критерий 0,02', labels, values:labels.map(() => 0.02) }], options:{ chartColors:[H.accent3], lineSize:2, lineDash:['dash'], lineDataSymbol:'none' } }
  ], cb({ x:0.6, y:1.55, w:7.6, h:5.3, valAxisMinVal:0, valAxisMaxVal:0.025, valAxisLabelFormatCode:'0.000', showLegend:true, legendPos:'b',
    showCatAxisTitle:true, catAxisTitle:'вариант: почва (A, B, C) и плотность (1 — 1,35; 2 — 1,52 г/см³)', valAxisTitle:'м³/м³', showValAxisTitle:true }));
  [['0,006–0,014', 'RMSE во всех шести вариантах — ниже критерия 0,02', TEAL], ['1 : 1,5 : 2', 'рост RMSE от A к B и C (0,0065; 0,0100; 0,0130)', BROWN], ['MB > 0', 'везде завышение; в A неотличимо от нуля, в B и C +0,006…+0,010', TERRA]].forEach((c, i) => {
    const y = 1.55 + i * 1.8;
    card(s, 8.6, y, 4.1, 1.65, LT, 'kpi-' + i);
    T(s, c[0], 8.85, y + 0.1, 3.7, 0.6, { fontSize:28, bold:true, color:c[2], valign:'middle' });
    T(s, c[1], 8.85, y + 0.75, 3.7, 0.85, { fontSize:15 });
  });
}

// ================= 14. result 2 =================
{
  const s = addSlide(14);
  const X = [0.05, 0.10, 0.20, 0.30, 0.38, 0.40, 0.44, 0.48];
  s.addChart(pres.charts.SCATTER, [
    { name:'θ', values:X },
    { name:'A (10–20 % физ. глины)', values:TV.A.concat([null, null]) },
    { name:'B (40–50 %)', values:TV.B.concat([null, null]) },
    { name:'C (50–65 %)', values:TV.C.concat([null, null]) },
    { name:'C: состояния вне анализа', values:[null, null, null, null, null, 0.027, 0.031, 0.038] },
    { name:'Паспортный предел 0,025', values:X.map(() => 0.025) }
  ], cb({ x:0.6, y:1.55, w:8.0, h:5.3, lineSize:2.25, lineDataSymbolSize:8, chartColors:[H.accent1, H.accent4, H.accent3, 'E3A98F', H.accent5],
    valAxisMinVal:0, valAxisMaxVal:0.04, valAxisMajorUnit:0.005, catAxisMinVal:0, catAxisMaxVal:0.5, catAxisMajorUnit:0.1, valAxisLabelFormatCode:'General',
    showLegend:true, legendPos:'b', showCatAxisTitle:true, catAxisTitle:'влажность по эталону θ, м³/м³', showValAxisTitle:true, valAxisTitle:'ошибка Δθ, м³/м³' }));
  T(s, 'Прирост ошибки при θ от 0,20 до 0,40', 8.9, 1.5, 3.8, 0.4, { fontSize:15, bold:true, color:MUTED });
  [['A', '+0,008', TEAL], ['B', '+0,014', BROWN], ['C', '+0,021', TERRA]].forEach((r, k) => {
    const y = 1.95 + k * 0.85;
    card(s, 8.9, y, 3.8, 0.7, LT, 'inc-' + k);
    circ(s, 9.05, y + 0.1, 0.5, r[2], r[0], 16);
    T(s, r[1] + ' м³/м³', 9.75, y, 2.8, 0.7, { fontSize:22, bold:true, valign:'middle' });
  });
  card(s, 8.9, 4.6, 3.8, 2.25, DK, 'ratio');
  T(s, '× 2,6', 9.15, 4.65, 3.3, 0.85, { fontSize:44, bold:true, color:OCHRE, valign:'middle' });
  T(s, 'во столько раз ошибка в C прирастает сильнее, чем в A. Бледные точки (0,44 и 0,48, вне анализа): +0,031 и +0,038 — плато нет', 9.15, 5.5, 3.3, 1.3, { fontSize:14, color:WH });
}

// ================= 15. "not significant" =================
{
  const s = addSlide(15, 'Как доказали');
  const lv = TV.lv;
  const series = lv.map((l, i) => ({ name:'c' + (i + 1), labels:['A', 'B', 'C'], values:[TV.A[i], TV.B[i], TV.C[i]] }));
  series.push({ name:'среднее по группе', labels:['A', 'B', 'C'], values:[0.0050, 0.0088, 0.0117] });
  s.addChart(pres.charts.LINE, series, cb({ x:0.6, y:1.55, w:6.0, h:4.85, chartColors:['9AA6A4', '9AA6A4', '9AA6A4', '9AA6A4', '9AA6A4', '9AA6A4', H.dk1], lineSize:3, lineDataSymbol:'circle', lineDataSymbolSize:11,
    showLegend:false, valAxisMinVal:0, valAxisMaxVal:0.03, valAxisMajorUnit:0.005, valAxisLabelFormatCode:'General', showTitle:true, title:'Все уровни влажности в одной куче: облака перекрываются', catAxisLabelFontSize:16,
    showValAxisTitle:true, valAxisTitle:'ошибка Δθ, м³/м³' }));
  T(s, [{ text:'●  ', options:{ color:SLATE } }, { text:'отдельные уровни влажности    ' }, { text:'●—  ', options:{ color:DK, bold:true } }, { text:'среднее по группе' }], 0.8, 6.42, 5.8, 0.3, { fontSize:13 });
  T(s, 'Иллюстрация: табл. V, 6 уровней, плотность 1,35 г/см³', 0.8, 6.68, 5.8, 0.28, { fontSize:12, color:MUTED, italic:true });
  s.addChart(pres.charts.BAR, [
    { name:'Почва 9,1 %', labels:[''], values:[9.1] }, { name:'Плотность 1,0 %', labels:[''], values:[1.0] },
    { name:'Взаимодействие 0,1 %', labels:[''], values:[0.1] }, { name:'Внутри вариантов 89,8 %', labels:[''], values:[89.8] } ],
    cb({ x:6.9, y:1.55, w:5.8, h:1.95, barDir:'bar', barGrouping:'stacked', chartColors:[H.accent3, H.accent4, H.accent2, 'C9D1CF'], showLegend:true, legendPos:'b',
      valAxisHidden:true, catAxisHidden:true, valGridLine:{ style:'none' }, valAxisMaxVal:100, showTitle:true, title:'Доли изменчивости ошибки (все 60 значений)', barGapWidthPct:20 }));
  card(s, 6.9, 3.65, 5.8, 1.6, DK, 'verdict');
  T(s, 'Почва: F(2; 54) = 2,74; p = 0,074', 7.15, 3.72, 5.3, 0.55, { fontSize:20, bold:true, color:WH, valign:'middle' });
  T(s, [{ text:'«Не значимо» ', options:{ bold:true, color:OCHRE } }, { text:'= «не слышно на этом шуме», а не «эффекта нет»', options:{ color:WH } }], 7.15, 4.3, 5.3, 0.85, { fontSize:17 });
  card(s, 6.9, 5.4, 5.8, 1.45, LT, 'analogy');
  T(s, [{ text:'Разговор на концерте: ', options:{ bold:true, color:TEAL } }, { text:'разговор есть, но его заглушает музыка — рост ошибки с влажностью' }], 7.15, 5.4, 5.3, 1.45, { fontSize:17, valign:'middle' });
}

// ================= 16. same-moisture comparison =================
{
  const s = addSlide(16);
  const seq = ['9FD0D6', '72BAC3', '4CA3AF', '2C8C9A', '1C7484', '0E4F5A'];
  s.addChart(pres.charts.LINE, TV.lv.map((l, i) => ({ name:'θ = ' + l, labels:['A', 'B', 'C'], values:[TV.A[i], TV.B[i], TV.C[i]] })),
    cb({ x:0.6, y:1.55, w:6.2, h:5.3, chartColors:seq, lineSize:3, lineDataSymbol:'circle', lineDataSymbolSize:9, showLegend:true, legendPos:'b',
      valAxisMinVal:0, valAxisMaxVal:0.03, valAxisMajorUnit:0.005, valAxisLabelFormatCode:'General', showTitle:true, title:'Те же точки, соединённые по уровню влажности', catAxisLabelFontSize:16,
      showValAxisTitle:true, valAxisTitle:'ошибка Δθ, м³/м³' }));
  card(s, 7.1, 1.55, 5.6, 1.55, LT, 'intuition');
  T(s, [{ text:'Интуиция: ', options:{ bold:true, color:TEAL } }, { text:'если бы почва не влияла, A оказалась бы наименьшей на всех 6 уровнях с шансом (1/3)⁶ ≈ 1 из 729' }], 7.35, 1.55, 5.1, 1.55, { fontSize:17, valign:'middle' });
  [['Влажность — блок', 'F(2; 10) = 6,92\np = 0,013'], ['Фридман: только порядок', 'χ²(2) = 11,27\np = 0,004']].forEach((r, k) => {
    const x = 7.1 + k * 2.85;
    card(s, x, 3.25, 2.75, 1.65, DK, 'test-' + k);
    T(s, [{ text:r[0], options:{ fontSize:14, color:OCHRE, bold:true, breakLine:true } }, { text:r[1], options:{ fontSize:20, color:WH, bold:true } }], x + 0.2, 3.25, 2.4, 1.65, { valign:'middle' });
  });
  card(s, 7.1, 5.05, 5.6, 1.8, WH, 'caveat', { line:{ color: OCHRE, width:1.5 } });
  T(s, [{ text:'Граница вывода: ', options:{ bold:true, color:TERRA } }, { text:'уровни — состояния одних колонок, почва в группе одна. Порядок устойчив в изученных почвах; перенос на другие — гипотеза' }], 7.35, 5.05, 5.1, 1.8, { fontSize:16, valign:'middle' });
}

// ================= 17. clay or salts =================
{
  const s = addSlide(17);
  const mb = [0.004, 0.002, 0.007, 0.006, 0.010, 0.008];
  const MH = [2.1, 2.3, 6.5, 6.8, 11.4, 11.1], EC = [0.35, 0.28, 0.52, 0.45, 0.61, 0.49];
  const chart = (x0, xv, a, b, xt, ttl, minx, maxx, major, lx) => s.addChart(pres.charts.SCATTER, [
    { name:'x', values:xv.concat(lx) }, { name:'шесть вариантов', values:mb.concat([null, null]) }, { name:'аппроксимация', values:xv.map(() => null).concat(lx.map(v => a + b * v)) }
  ], cb({ x:x0, y:1.55, w:5.95, h:3.85, lineSize:2, lineDataSymbolSize:11, chartColors:[H.accent3, H.accent1], showLegend:false, valAxisMinVal:0, valAxisMaxVal:0.012, valAxisMajorUnit:0.002, valAxisLabelFormatCode:'General',
    catAxisMinVal:minx, catAxisMaxVal:maxx, catAxisMajorUnit:major, showCatAxisTitle:true, catAxisTitle:xt, showTitle:true, title:ttl, showValAxisTitle:true, valAxisTitle:'MB, м³/м³' }));
  chart(0.6, MH, 0.0017, 0.00066, 'MH, %', 'Подозреваемый 1 — поверхность глин (MH): R² = 0,88', 0, 12, 2, [0, 12]);
  chart(6.75, EC, -0.0044, 0.0235, 'ECe, дСм/м', 'Подозреваемый 2 — соли (ECe): R² = 0,96', 0, 0.7, 0.1, [0.25, 0.65]);
  card(s, 0.6, 5.6, 4.0, 1.25, LT, 'tied');
  T(s, [{ text:'r = 0,88', options:{ bold:true, fontSize:24, color:TERRA, breakLine:true } }, { text:'между MH и ECe: подозреваемые всегда вместе', options:{ fontSize:15 } }], 0.85, 5.6, 3.6, 1.25, { valign:'middle' });
  card(s, 4.8, 5.6, 4.0, 1.25, LT, 'fragile');
  T(s, [{ text:'p = 0,017', options:{ bold:true, fontSize:24, color:TEAL, breakLine:true } }, { text:'в модели с обоими значима ECe, но вывод хрупкий; почв три, а не шесть', options:{ fontSize:15 } }], 5.05, 5.6, 3.6, 1.25, { valign:'middle' });
  card(s, 9.0, 5.6, 3.7, 1.25, DK, 'need');
  T(s, 'Нужны почвы, где глина и соли не ходят парой', 9.2, 5.6, 3.3, 1.25, { fontSize:17, bold:true, color:WH, valign:'middle' });
}

// ================= 18. CRIM prediction =================
{
  const s = addSlide(18);
  s.addChart(pres.charts.BAR, [
    { name:'Твёрдая фаза', labels:['1,35 г/см³', '1,52 г/см³'], values:[0.513, 0.578] },
    { name:'Вода (θ = 0,25)', labels:['1,35 г/см³', '1,52 г/см³'], values:[0.25, 0.25] },
    { name:'Воздух', labels:['1,35 г/см³', '1,52 г/см³'], values:[0.237, 0.172] } ],
    cb({ x:0.6, y:1.55, w:5.6, h:5.3, barDir:'col', barGrouping:'stacked', chartColors:['B9C3C1', '8FC5CD', 'EEF1F0'], showValue:true, dataLabelPosition:'ctr', dataLabelFormatCode:'0.00', dataLabelColor:H.dk1, dataLabelFontSize:14,
      valAxisMinVal:0, valAxisMaxVal:1, valAxisMajorUnit:0.2, valAxisLabelFormatCode:'0.0', showLegend:true, legendPos:'b', showTitle:true, title:'Доли объёма при θ = 0,25 (ρs = 2,63 г/см³)', barGapWidthPct:50, catAxisLabelFontSize:14 }));
  const steps = ['Модель: √ε смеси = φтв·√4,7 + θ·√80 + φвозд·√1',
                 'Уплотнение 1,35 → 1,52: твёрдого больше на 0,17 / 2,63 ≈ 0,065 — вместо воздуха',
                 'Пробег растёт: 0,065 × (√4,7 − 1) = 0,065 × 1,17 ≈ 0,076',
                 'Датчик читает это как воду вместо воздуха: 0,076 / (√80 − 1) = 0,076 / 7,94 ≈ 0,0095'];
  steps.forEach((t, k) => {
    const y = 1.55 + k * 0.97;
    card(s, 6.6, y, 6.1, 0.85, LT, 'step-' + k);
    circ(s, 6.75, y + 0.17, 0.5, TEAL, String(k + 1), 16);
    T(s, t, 7.4, y, 5.2, 0.85, { fontSize:15, valign:'middle', bold:k === 3 });
  });
  card(s, 6.6, 5.5, 6.1, 0.85, DK, 'prediction');
  T(s, [{ text:'Прогноз: ≈ +0,0095 ', options:{ bold:true, color:OCHRE } }, { text:'(при ρs 2,60–2,67: +0,0094…+0,0096)', options:{ color:WH } }], 6.85, 5.5, 5.7, 0.85, { fontSize:17, valign:'middle' });
  T(s, 'Скрыто: та же θ, та же проводимость, нет потерь в пути', 6.6, 6.45, 6.1, 0.42, { fontSize:15, bold:true, color:TERRA });
}

// ================= 19. prediction failed =================
{
  const s = addSlide(19);
  s.addChart(pres.charts.BAR, [
    { name:'Прогноз CRIM', labels:['A', 'B', 'C'], values:[9.5, 9.5, 9.5] },
    { name:'Наблюдалось', labels:['A', 'B', 'C'], values:[-2, -1, -2] } ],
    cb({ x:0.6, y:1.55, w:5.9, h:5.3, barDir:'col', chartColors:[H.accent5, H.accent3], barGapWidthPct:60, showValue:true, dataLabelFontSize:14, dataLabelColor:H.dk1, dataLabelPosition:'outEnd', dataLabelFormatCode:'+0.0;−0.0',
      valAxisMinVal:-4, valAxisMaxVal:12, valAxisMajorUnit:2, showLegend:true, legendPos:'b', showCatAxisTitle:true, catAxisTitle:'почва', catAxisLabelPos:'low',
      showValAxisTitle:true, valAxisTitle:'сдвиг при 1,52 против 1,35 г/см³, ×10⁻³ м³/м³' }));
  card(s, 6.9, 1.55, 5.8, 1.4, DK, 'logic');
  T(s, [{ text:'модель ∧ та же θ ∧ та же проводимость ⇒ +0,0095', options:{ bold:true, color:WH, breakLine:true } }, { text:'наблюдали −0,0017 (ДИ −0,006…+0,003) ⇒ ложно хотя бы одно звено', options:{ color:OCHRE } }], 7.15, 1.55, 5.35, 1.4, { fontSize:16, valign:'middle' });
  [['×', TERRA, '«Та же проводимость» нарушена', 'ECe у плотных ниже на 0,07–0,12 дСм/м; с ECe в регрессии эффект плотности +0,0004 (p = 0,51)'],
   ['×', TERRA, '«Та же влажность» нарушена', 'средние посчитаны по разным диапазонам влажности, а ошибка растёт с θ'],
   ['?', OCHRE, 'Сама модель', 'в таком виде не содержит проводимости — описывает изолятор']].forEach((r, k) => {
    const y = 3.1 + k * 1.15;
    card(s, 6.9, y, 5.8, 1.02, LT, 'link-' + k);
    circ(s, 7.05, y + 0.26, 0.5, r[1], r[0], 20, r[1] === OCHRE ? DK : WH);
    T(s, [{ text:r[2], options:{ bold:true, fontSize:15, breakLine:true } }, { text:r[3], options:{ fontSize:14 } }], 7.7, y, 4.9, 1.02, { valign:'middle' });
  });
  T(s, 'Малый эффект плотности нельзя ни подтвердить, ни исключить', 6.9, 6.55, 5.8, 0.4, { fontSize:13, color:MUTED, italic:true });
}

// ================= 20. swelling =================
{
  const s = addSlide(20);
  const draw = (x, title) => {
    T(s, title, x, 1.5, 3.8, 0.4, { fontSize:17, bold:true, align:'center' });
    s.addShape(pres.ShapeType.rect, { x:x + 0.7, y:2.35, w:2.4, h:3.2, fill:{ color: TERRA }, line:{ type:'none' }, objectName:'soil' });
    s.addShape(pres.ShapeType.rect, { x:x + 0.7, y:2.35, w:2.4, h:3.2, fill:{ type:'none' }, line:{ color: SLATE, width:2 }, objectName:'column' });
    [1.4, 2.4].forEach(dx => s.addShape(pres.ShapeType.rect, { x:x + dx, y:2.35, w:0.07, h:3.2, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
  };
  draw(0.6, 'Сухая: 300 мм');
  draw(4.7, 'Влажная: ≈ 325 мм');
  s.addShape(pres.ShapeType.rect, { x:5.4, y:2.08, w:2.4, h:0.27, fill:{ color: TERRA }, line:{ type:'none' }, objectName:'swollen-top' });
  T(s, 'почва = длина штанг', 0.6, 5.65, 3.8, 0.4, { fontSize:15, align:'center', color:MUTED });
  T(s, 'почва выше штанг на ≈ 25 мм', 4.7, 5.65, 3.8, 0.4, { fontSize:15, align:'center', color:TERRA, bold:true });
  [['≈ 25 мм', 'подъём поверхности'], ['8,3 %', 'объёмная деформация'], ['0,460 > 0,415', 'θmax больше исходной пористости']].forEach((r, k) => {
    const y = 1.5 + k * 1.3;
    card(s, 8.9, y, 3.8, 1.15, LT, 'swell-' + k);
    T(s, r[0], 9.1, y + 0.05, 3.4, 0.6, { fontSize:26, bold:true, color:TERRA, valign:'middle' });
    T(s, r[1], 9.1, y + 0.65, 3.4, 0.42, { fontSize:15 });
  });
  card(s, 0.6, 6.12, 12.1, 0.8, WH, 'swell-note', { line:{ color: OCHRE, width:1.5 } });
  T(s, 'Делить воду на прежний объём — значит завысить эталон и занизить видимую ошибку датчика. Эталон считали по реальному объёму: дополнительной ошибки не нашли.', 0.85, 6.12, 11.6, 0.8, { fontSize:15, valign:'middle' });
}

// ================= 21. traffic light =================
{
  const s = addSlide(21, 'Что делать');
  [[GREEN, 'Лёгкие, 10–20 % физ. глины', 'заводская калибровка без поправки'],
   [OCHRE, 'Тяжёлые, 40–65 %', 'в среднем +0,006…+0,010 (6–10 мм в метровом слое), до +0,027 во влажной: калибровка под почву или 1–2 весовые проверки за сезон'],
   [TERRA, 'Постоянная поправка', 'не годится: ошибка растёт с влажностью (график справа)'],
   [SLATE, 'Серая зона', '20–40 % физ. глины не испытаны; в поле нужна температурная поправка']].forEach((r, k) => {
    const y = 1.55 + k * 1.35;
    card(s, 0.6, y, 6.9, 1.22, LT, 'tl-' + k);
    s.addShape(pres.ShapeType.ellipse, { x:0.8, y:y + 0.26, w:0.7, h:0.7, fill:{ color: r[0] }, line:{ type:'none' }, objectName:'light' });
    T(s, [{ text:r[1], options:{ bold:true, fontSize:17, breakLine:true } }, { text:r[2], options:{ fontSize:14, color:MUTED } }], 1.75, y, 5.6, 1.22, { valign:'middle' });
  });
  s.addChart(pres.charts.LINE, [
    { name:'Ошибка в почве C', labels:TV.lv, values:TV.C },
    { name:'Остаток после вычитания +0,010', labels:TV.lv, values:TV.C.map(v => +(v - 0.010).toFixed(3)) } ],
    cb({ x:7.8, y:1.5, w:4.9, h:3.5, chartColors:[H.accent3, H.accent1], lineSize:2.5, lineDataSymbolSize:7, valAxisMinVal:-0.01, valAxisMaxVal:0.03, valAxisMajorUnit:0.01, valAxisLabelFormatCode:'General',
      showLegend:true, legendPos:'b', showTitle:true, title:'Постоянная поправка +0,010 в почве C', showCatAxisTitle:true, catAxisTitle:'θ, м³/м³', catAxisLabelPos:'low' }));
  card(s, 7.8, 5.15, 4.9, 1.7, DK, 'two-point');
  T(s, [{ text:'Проверять в двух точках: ', options:{ bold:true, color:OCHRE } }, { text:'после полива (влажно) и перед поливом (сухо). Одна точка даёт сдвиг, две — ещё и наклон. Следствие из данных; в поле не проверено.', options:{ color:WH } }], 8.0, 5.15, 4.5, 1.7, { fontSize:15, valign:'middle' });
}

// ================= 22. transferable ideas =================
{
  const s = addSlide(22);
  [['Ошибка — функция, а не число', 'проверять и калибровать по всему рабочему диапазону влажности'],
   ['Сравнивай при равных условиях', 'блокирование по влажности — для любого калибровочного опыта'],
   ['Прокси ходят парами', 'разводить связанные факторы планом опыта, а не расчётом'],
   ['Несбывшийся прогноз — подарок', 'количественная модель указывает, какое допущение проверить'],
   ['Ошибка — в единицах решения', '«10 мм воды в метровом слое» понятнее, чем «0,010 м³/м³»'],
   ['Тот же вопрос — другим датчикам', 'все, кто судит о воде по проницаемости, исходят из «почва — изолятор»']].forEach((c, k) => {
    const x = 0.6 + (k % 3) * 4.1, y = 1.6 + Math.floor(k / 3) * 2.68;
    card(s, x, y, 3.9, 2.5, LT, 'idea-' + k);
    circ(s, x + 0.25, y + 0.25, 0.6, [TEAL, BROWN, TERRA, SLATE, GREEN, DK][k], String(k + 1), 20);
    T(s, c[0], x + 0.25, y + 0.98, 3.4, 0.7, { fontSize:19, bold:true });
    T(s, c[1], x + 0.25, y + 1.65, 3.4, 0.8, { fontSize:15, color:MUTED });
  });
}

// ================= 23. if-then predictions =================
{
  const s = addSlide(23);
  const hd = { bold:true, color:WH, fill:{ color: DK }, fontSize:14, valign:'middle' };
  const c = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:14, valign:'middle', color:DK, fill:{ color: LT } }, o) });
  const rows = [
    [{ text:'Если верно…', options:hd }, { text:'…то увидим', options:hd }, { text:'Решающий опыт', options:hd }],
    [c('Завышение создаёт проводимость порового раствора', { bold:true }), c('ошибка растёт вместе с ECe при засолении одной и той же почвы'), c('ступенчатое засоление (ECe 1–4 дСм/м — сделано, отдельная статья)')],
    [c('Работает поверхность глин', { bold:true }), c('при равной ECe ошибка больше в почве с большей долей смектита'), c('минералогия глин; пары почв с равной ECe')],
    [c('Эффект плотности «съела» разница в ECe', { bold:true }), c('при выровненной ECe плотный вариант ближе к прогнозу +0,009…+0,010'), c('уплотнение одной почвы при одинаковой ECe')],
    [c('Эффект почвы реален, а не артефакт колонок', { bold:true }), c('он сохранится в смешанной модели (колонка — случайный эффект)'), c('переанализ 180 колоночных значений — данные уже есть', { bold:true, color:TEAL })],
    [c('Идея: завышение задаёт проводимость', { bold:true, color:TERRA }), c('поправка по двум входам (τ и объёмная ЭП) уберёт большую часть ошибки'), c('параллельно мерить объёмную ЭП (в этом опыте не измеряли)')]
  ];
  s.addTable(rows, { x:0.6, y:1.55, w:12.1, colW:[3.7, 4.4, 4.0], rowH:[0.45, 0.9, 0.8, 0.9, 0.9, 0.9], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.1, objectName:'if-then' });
  T(s, 'Попробуйте сами выбрать решающий опыт: какой из них различит гипотезы с наименьшими затратами?', 0.6, 6.55, 12.1, 0.3, { fontSize:13, color:MUTED, italic:true });
}

// ================= 24. evidence ledger =================
{
  const s = addSlide(24);
  const hd = { bold:true, color:WH, fill:{ color: DK }, fontSize:14, valign:'middle' };
  const c = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:14, valign:'middle', color:DK, fill:{ color: LT } }, o) });
  const dots = { 'высокая':['●●●', TEAL], 'средняя':['●●○', BROWN], 'низкая':['●○○', TERRA] };
  const rows = [[{ text:'Утверждение', options:hd }, { text:'Твёрдость', options:Object.assign({}, hd, { align:'center' }) }, { text:'Что мешает сказать сильнее', options:hd }]];
  K.ledger.forEach(r => rows.push([c(r[0], { bold:true }), c(dots[r[3]][0], { fontSize:20, align:'center', color:dots[r[3]][1] }), c(r[4])]));
  s.addTable(rows, { x:0.6, y:1.55, w:12.1, colW:[4.6, 1.6, 5.9], rowH:[0.42].concat(K.ledger.map(() => 0.69)), border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.08, objectName:'ledger' });
}

// ================= 25. three take-aways =================
{
  const s = addSlide(25);
  [['1', 'Годится, но не бесплатно', 'RMSE 0,006–0,014 < 0,02 везде. Но при 40–65 % физ. глины датчик завышает, и сильнее во влажной почве'],
   ['2', 'Доказывали устойчивость, а не только p', 'порядок почв сохранялся на каждом уровне влажности, два разных теста согласны; слабые места показаны'],
   ['3', 'Гипотеза подтверждена частично', 'связь с почвой и рост с влажностью — да; плотность — нет; механизм (глина или соли) не разделён']].forEach((t, k) => {
    const x = 0.6 + k * 4.1;
    card(s, x, 1.55, 3.9, 2.95, '3D2E22', 'take-' + k);
    circ(s, x + 0.3, 1.78, 0.7, OCHRE, t[0], 22, DK);
    T(s, t[1], x + 0.3, 2.6, 3.3, 0.8, { fontSize:20, bold:true, color:WH });
    T(s, t[2], x + 0.3, 3.38, 3.3, 1.05, { fontSize:15, color:WH });
  });
  [['H1 · глина', 'частично', OCHRE, DK], ['H2 · влажность', 'подтверждена', GREEN, WH], ['H3 · плотность', 'не подтверждена', TERRA, WH]].forEach((h, k) => {
    const x = 0.6 + k * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y:4.7, w:3.9, h:0.75, fill:{ color: h[2] }, line:{ type:'none' }, rectRadius:0.08, objectName:'verdict-' + k });
    T(s, [{ text:h[0] + ': ', options:{ bold:true } }, { text:h[1] }], x + 0.2, 4.7, 3.5, 0.75, { fontSize:17, color:h[3], valign:'middle' });
  });
  T(s, 'Работа — карта: где датчик надёжен, где нужна проверка и какой опыт закроет пробел', 0.6, 5.7, 12.1, 0.5, { fontSize:20, bold:true, color:OCHRE });
  T(s, 'Спасибо за внимание!', 0.6, 6.25, 12.1, 0.6, { fontSize:26, bold:true, color:WH });
}

// ================= backup slides =================
function addBackup(i) {
  const b = K.backup[i];
  if (i === 0) { pres.addSection({ title:'Запасные слайды' }); curSection = 'Запасные слайды'; }
  const s = pres.addSlide({ masterName:'LIGHT', sectionTitle:curSection });
  s.addText(b.title, { placeholder:'title' });
  s.addNotes('[Запасной слайд — показывать только при вопросах]\n\n' + b.note);
  T(s, 'Запасные слайды', 0.6, 7.02, 6, 0.3, { fontSize:11, color:SLATE, valign:'middle' });
  return s;
}
const tblHd = t => ({ text:t, options:{ bold:true, color:WH, fill:{ color: DK }, fontSize:14, align:'center', valign:'middle' } });
const tc = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:14, align:'center', valign:'middle', color:DK, fill:{ color: LT } }, o) });
{
  const s = addBackup(0);
  const rows = [['Вариант', 'ECe, дСм/м', 'MB', 'RMSE', 'Δθ min / max', '95 % ДИ для MB', 'p (MB = 0)', 'J_nl'].map(tblHd)];
  [['A1-n', '0,35', '+0,004', '0,007', '−0,002 / +0,013', '−0,0003…+0,0083', '0,066', '−0,263'],
   ['A2-n', '0,28', '+0,002', '0,006', '−0,004 / +0,011', '−0,0023…+0,0063', '0,32', '−0,268'],
   ['B1-n', '0,52', '+0,007', '0,011', '−0,003 / +0,019', '+0,0006…+0,0134', '0,035', '−0,118'],
   ['B2-n', '0,45', '+0,006', '0,009', '−0,004 / +0,016', '+0,0009…+0,0111', '0,025', '0,000'],
   ['C1-n', '0,61', '+0,010', '0,014', '−0,001 / +0,027', '+0,0026…+0,0174', '0,014', '−0,306'],
   ['C2-n', '0,49', '+0,008', '0,012', '−0,002 / +0,022', '+0,0013…+0,0147', '0,025', '−0,225']].forEach(r => rows.push(r.map((v, k) => tc(v, k === 0 ? { bold:true } : {}))));
  s.addTable(rows, { x:0.6, y:1.6, w:12.1, colW:[1.3, 1.3, 1.1, 1.1, 2.0, 2.5, 1.4, 1.4], rowH:[0.5, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.05, objectName:'table-III' });
  T(s, 'м³/м³; MB и RMSE — по средним трёх колонок, n = 10 уровней влажности. ДИ и p описательные (уровни — состояния одних колонок). J_nl < 0 — правосторонняя асимметрия: крупные положительные ошибки во влажной части. Код варианта: группа, плотность (1 — 1,35; 2 — 1,52 г/см³), n — незасолённая.', 0.6, 5.6, 12.1, 1.1, { fontSize:15, color:MUTED });
}
{
  const s = addBackup(1);
  const rows = [['Источник изменчивости', 'SS, (м³/м³)²', 'df', 'MS, (м³/м³)²', 'F', 'p', 'η²'].map(tblHd)];
  [['Почва (T)', '3,63·10⁻⁴', '2', '1,82·10⁻⁴', '2,74', '0,074', '0,091'],
   ['Плотность (ρb)', '4,17·10⁻⁵', '1', '4,17·10⁻⁵', '0,63', '0,43', '0,010'],
   ['Взаимодействие T × ρb', '3,3·10⁻⁶', '2', '1,7·10⁻⁶', '0,03', '0,98', '0,001'],
   ['Внутри вариантов', '3,58·10⁻³', '54', '6,63·10⁻⁵', '—', '—', '0,898'],
   ['Всего', '3,99·10⁻³', '59', '—', '—', '—', '1,000']].forEach(r => rows.push(r.map((v, k) => tc(v, k === 0 ? { bold:true, align:'left' } : {}))));
  s.addTable(rows, { x:0.6, y:1.6, w:12.1, colW:[3.6, 1.9, 0.9, 1.9, 1.2, 1.3, 1.3], rowH:[0.5, 0.55, 0.55, 0.55, 0.55, 0.55], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.08, objectName:'table-IV' });
  T(s, 'N = 60 значений (6 вариантов × 10 уровней, средние трёх колонок); SS восстановлены по MB и RMSE вариантов. Внутривариантный член включает рост ошибки с влажностью. 20 000 симуляций округления: p для почвы 0,04–0,12, η² 0,07–0,11. Эффект плотности −0,0017, 95 % ДИ −0,006…+0,003.', 0.6, 5.0, 12.1, 1.3, { fontSize:15, color:MUTED });
}
{
  const s = addBackup(2);
  const rows = [['Вариант', 'Группа (физ. глина, %)', 'ρs, г/см³', 'MH, %', 'ρb, г/см³', 'σ1:5, дСм/м', 'ECe, дСм/м'].map(tblHd)];
  [['A1-n', 'A (10–20)', '2,67 ± 0,01', '2,1 ± 0,2', '1,35 ± 0,02', '0,044 ± 0,003', '0,35 ± 0,02'],
   ['A2-n', 'A (10–20)', '2,67 ± 0,01', '2,3 ± 0,2', '1,52 ± 0,02', '0,036 ± 0,002', '0,28 ± 0,02'],
   ['B1-n', 'B (40–50)', '2,63 ± 0,01', '6,5 ± 0,3', '1,35 ± 0,03', '0,066 ± 0,004', '0,52 ± 0,03'],
   ['B2-n', 'B (40–50)', '2,63 ± 0,01', '6,8 ± 0,4', '1,52 ± 0,03', '0,058 ± 0,004', '0,45 ± 0,03'],
   ['C1-n', 'C (50–65)', '2,60 ± 0,02', '11,4 ± 0,6', '1,35 ± 0,03', '0,078 ± 0,005', '0,61 ± 0,04'],
   ['C2-n', 'C (50–65)', '2,60 ± 0,02', '11,1 ± 0,6', '1,52 ± 0,03', '0,062 ± 0,004', '0,49 ± 0,03']].forEach(r => rows.push(r.map((v, k) => tc(v, k === 0 ? { bold:true } : {}))));
  s.addTable(rows, { x:0.6, y:1.6, w:12.1, colW:[1.3, 2.4, 1.6, 1.5, 1.7, 1.8, 1.8], rowH:[0.5, 0.55, 0.55, 0.55, 0.55, 0.55, 0.55], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.05, objectName:'table-II' });
  T(s, 'n = 3, среднее ± SD. Все ECe ниже 1 дСм/м и в 6,6–14 раз ниже порога засоления 4 дСм/м. В плотных вариантах ECe ниже на 0,07–0,12 дСм/м. Почва C2-n набухала при θ > 0,40.', 0.6, 5.6, 12.1, 0.9, { fontSize:15, color:MUTED });
}
{
  const s = addBackup(3);
  const half = Math.ceil(K.limits.length / 2);
  [K.limits.slice(0, half), K.limits.slice(half)].forEach((col, j) => {
    card(s, 0.6 + j * 6.2, 1.6, 5.9, 4.6, LT, 'limits-' + j);
    T(s, col.map((t, i) => ({ text:`${j * half + i + 1}. ${t}`, options:{ breakLine:i < col.length - 1, paraSpaceAfter:10 } })), 0.85 + j * 6.2, 1.8, 5.4, 4.2, { fontSize:16 });
  });
}

// per-series styling that pptxgenjs cannot express
async function patchSeries(file) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const rules = [
    [/^шесть вариантов$/, { noLine:true }],
    [/^c\d$/, { noLine:true }],
    [/^аппроксимация$/, { noMarker:true }],
    [/^Паспортный предел 0,025$/, { dash:'dash', noMarker:true }],
    [/^C: состояния вне анализа$/, { dash:'sysDash' }],
    [/^порог срабатывания$/, { dash:'dash' }]
  ];
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    let xml = await zip.file(name).async('string');
    xml = xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ser => {
      const m = ser.match(/<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/); if (!m) return ser;
      const hit = rules.find(([re]) => re.test(m[1])); if (!hit) return ser;
      const r = hit[1];
      if (r.noLine) ser = ser.replace(/(<c:spPr>[\s\S]*?)<a:ln[\s\S]*?<\/a:ln>/, '$1<a:ln w="28575"><a:noFill/></a:ln>');
      if (r.dash) ser = ser.replace(/(<c:spPr>[\s\S]*?<a:ln[^>]*>[\s\S]*?)<a:prstDash val="solid"\/>/, `$1<a:prstDash val="${r.dash}"/>`);
      if (r.noMarker) ser = ser.replace(/<c:marker>[\s\S]*?<\/c:marker>/, '<c:marker><c:symbol val="none"/></c:marker>');
      return ser;
    });
    zip.file(name, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' }));
}

(async () => {
  const out = path.resolve(process.argv[2] || 'deck3.pptx');
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  await patchSeries(out);
  console.log('written', out);
})();
