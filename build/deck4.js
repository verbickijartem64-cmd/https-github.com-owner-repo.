// Версия 4: доклад на 10 минут (12 слайдов) + 9 запасных слайдов для вопросов.
const pptxgen = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');
const K3 = require('./content3.js');           // проверенные данные версии 3 (таблицы, ограничения)
const K = require('./content4.json');          // текст 10-минутной версии

const THEME = {
  name: 'CS616 Chernozem', headFontFace: 'Cambria', bodyFontFace: 'Calibri',
  colors: { dk1:'2A1E16', lt1:'FFFFFF', dk2:'5B4636', lt2:'F1F3F2',
    accent1:'1C7C8C', accent2:'D9A441', accent3:'B8552F', accent4:'8C6A4F', accent5:'6E7F80', accent6:'4E8A5B', hlink:'1C7C8C', folHlink:'8C6A4F' }
};
const H = THEME.colors;
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = 'Когда датчик влажности «льстит» почве: CS616 в незасолённых почвах (10 минут)';
pres.author = 'Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е.';
pres.company = 'Кубанский ГАУ';
const C = pres.SchemeColor;
const DK = C.text1, MUTED = C.text2, WH = C.background1, LT = C.background2;
const TEAL = C.accent1, OCHRE = C.accent2, TERRA = C.accent3, BROWN = C.accent4, SLATE = C.accent5, GREEN = C.accent6;

const titlePh = color => ({ placeholder:{ options:{ name:'title', type:'title', x:0.6, y:0.36, w:12.1, h:1.05, fontSize:30, bold:true, color, align:'left', valign:'top', margin:0 }, text:'' } });
pres.defineSlideMaster({ title:'LIGHT', background:{ color: WH }, margin:[0.5, 0.6, 0.6, 0.6], objects:[titlePh(DK)],
  slideNumber:{ x:12.03, y:7.02, w:0.7, h:0.3, fontSize:11, color:MUTED, align:'right' } });
pres.defineSlideMaster({ title:'DARK', background:{ color: DK }, margin:[0.5, 0.6, 0.6, 0.6], objects:[titlePh(WH)],
  slideNumber:{ x:12.03, y:7.02, w:0.7, h:0.3, fontSize:11, color:OCHRE, align:'right' } });
pres.defineSlideMaster({ title:'COVER', background:{ color: DK }, margin:[0.5, 0.6, 0.6, 0.6],
  objects:[{ placeholder:{ options:{ name:'title', type:'title', x:0.7, y:1.15, w:7.9, h:2.0, fontSize:40, bold:true, color:WH, align:'left', valign:'top', margin:0 }, text:'' } }] });

// ---------- helpers ----------
const fmt = n => String(n).replace('.', ',');
const wc = s => s.split(/\s+/).filter(Boolean).length;
function T(s, text, x, y, w, h, o = {}) { s.addText(text, Object.assign({ x, y, w, h, fontSize:16, color:DK, margin:0, valign:'top', isTextBox:true }, o)); }
function card(s, x, y, w, h, fill = LT, name = 'card', o = {}) { s.addShape(pres.ShapeType.roundRect, Object.assign({ x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, rectRadius:0.08, objectName:name }, o)); }
function circ(s, x, y, d, fill, label, fs = 18, color = WH) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w:d, h:d, fill:{ color: fill }, line:{ type:'none' }, objectName:'badge' });
  if (label) s.addText(label, { x, y, w:d, h:d, align:'center', valign:'middle', fontSize:fs, bold:true, color, margin:0, isTextBox:true });
}
function L(s, x1, y1, x2, y2, color, o = {}) {
  s.addShape(pres.ShapeType.line, { x:Math.min(x1, x2), y:Math.min(y1, y2), w:Math.abs(x2 - x1), h:Math.abs(y2 - y1), flipH:x2 < x1, flipV:y2 < y1,
    line:Object.assign({ color, width:2 }, o), objectName:o.name || 'line' });
}
function arrowDown(s, x, y, w = 0.4, h = 0.3, fill = SLATE) { s.addShape(pres.ShapeType.downArrow, { x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, objectName:'arrow' }); }
function arrowRight(s, x, y, w = 0.3, h = 0.4, fill = SLATE) { s.addShape(pres.ShapeType.rightArrow, { x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, objectName:'arrow' }); }
const ACTS = ['Зачем', 'Как проверяли', 'Что нашли', 'Как доказали', 'Что делать'];
const ACT_OF = n => n === 1 ? null : n <= 4 ? 'Зачем' : n === 5 ? 'Как проверяли' : n === 6 ? 'Что нашли' : n <= 9 ? 'Как доказали' : 'Что делать';
function tracker(s, act, dark) {
  if (!act) return;
  const runs = [];
  ACTS.forEach((a, i) => {
    const on = a === act;
    runs.push({ text:a, options:{ bold:on, color:on ? (dark ? OCHRE : TEAL) : (dark ? LT : SLATE) } });
    if (i < ACTS.length - 1) runs.push({ text:'   ·   ', options:{ color:dark ? LT : SLATE } });
  });
  T(s, runs, 0.6, 7.02, 10.5, 0.3, { fontSize:11, valign:'middle' });
}
const WPM = 140;
const secs = K.slides.map(d => Math.round(wc(d.text.join(' ')) / WPM * 60));
const mmss = t => `${Math.floor(t / 60)}:${String(t % 60).padStart(2, '0')}`;
let curSection = null;
function addSlide(n, section) {
  const d = K.slides[n - 1];
  if (section) { pres.addSection({ title: section }); curSection = section; }
  const master = n === 1 ? 'COVER' : n === 12 ? 'DARK' : 'LIGHT';
  const s = pres.addSlide({ masterName: master, sectionTitle: curSection });
  s.addText(d.title, { placeholder:'title' });
  const cum = secs.slice(0, n).reduce((a, b) => a + b, 0) + 4 * n; // + ~4 с на смену слайда
  s.addNotes(`[≈ ${secs[n - 1]} с · к концу слайда ≈ ${mmss(cum)} из 10:00]\n\nГлавная мысль: ${d.takeaway}\n\n` + d.text.join('\n\n'));
  tracker(s, ACT_OF(n), master !== 'LIGHT');
  return s;
}
const chartBase = { catAxisLabelColor:H.dk2, valAxisLabelColor:H.dk2, catAxisLabelFontSize:12, valAxisLabelFontSize:12,
  catAxisLabelFontFace:'+mn-lt', valAxisLabelFontFace:'+mn-lt', legendFontFace:'+mn-lt', legendFontSize:12, legendColor:H.dk2,
  valGridLine:{ color:'D9DEDC', size:0.5 }, catGridLine:{ style:'none' }, dataLabelFontFace:'+mn-lt', titleFontFace:'+mn-lt',
  catAxisTitleFontFace:'+mn-lt', valAxisTitleFontFace:'+mn-lt', catAxisTitleFontSize:12, valAxisTitleFontSize:12,
  catAxisTitleColor:H.dk2, valAxisTitleColor:H.dk2, titleFontSize:14, titleColor:H.dk1 };
const cb = o => Object.assign({}, chartBase, o);
const TV = { lv:['0,05', '0,10', '0,20', '0,30', '0,38', '0,40'],
  A:[0.000, 0.001, 0.003, 0.005, 0.010, 0.011], B:[0.001, 0.002, 0.005, 0.009, 0.017, 0.019], C:[0.001, 0.002, 0.006, 0.012, 0.022, 0.027] };

// ================= 1. cover =================
{
  const s = addSlide(1, 'Доклад (10 минут)');
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
  [['≤ 0,014', 'наибольшая типичная ошибка (RMSE) среди шести вариантов — ниже критерия 0,02', TEAL, 44],
   ['+0,006…+0,010', 'среднее завышение влажности в тяжёлых почвах — всегда в одну сторону', TERRA, 32],
   ['× 2,6', 'во столько раз сильнее растёт ошибка с увлажнением в самой тяжёлой почве, чем в лёгкой', BROWN, 44]].forEach((r, k) => {
    const x = 0.6 + k * 4.1;
    card(s, x, 1.6, 3.9, 2.55, LT, 'stat-' + k);
    T(s, r[0], x + 0.25, 1.72, 3.4, 0.95, { fontSize:r[3], bold:true, color:r[2], valign:'middle' });
    T(s, r[1], x + 0.25, 2.75, 3.4, 1.3, { fontSize:16 });
  });
  card(s, 0.6, 4.4, 12.1, 0.95, WH, 'units', { line:{ color: TEAL, width:1.5 } });
  T(s, [{ text:'0,01 м³/м³', options:{ bold:true, color:TEAL } }, { text:'  =  1 л воды на 100 л почвы  =  ' }, { text:'10 мм воды в метровом слое', options:{ bold:true, color:TEAL } }, { text:'  =  100 м³ на гектар' }], 0.9, 4.4, 11.5, 0.95, { fontSize:20, valign:'middle' });
  card(s, 0.6, 5.6, 12.1, 1.25, DK, 'conclusion');
  T(s, [{ text:'Вывод: ', options:{ bold:true, color:OCHRE } }, { text:'в тяжёлых почвах поправка должна зависеть от влажности. Завышение показаний = полив запаздывает.', options:{ color:WH } }], 0.9, 5.6, 11.5, 1.25, { fontSize:20, valign:'middle' });
}

// ================= 3. sensor measures time + hidden assumptions =================
{
  const s = addSlide(3);
  card(s, 0.6, 1.55, 5.8, 1.75, BROWN, 'soil-block');
  s.addShape(pres.ShapeType.roundRect, { x:0.95, y:1.55, w:1.3, h:0.5, fill:{ color: DK }, line:{ type:'none' }, rectRadius:0.05, objectName:'probe-head' });
  T(s, 'датчик', 0.95, 1.55, 1.3, 0.5, { fontSize:13, color:WH, align:'center', valign:'middle' });
  [1.2, 2.0].forEach(x => s.addShape(pres.ShapeType.rect, { x, y:2.05, w:0.07, h:1.1, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
  L(s, 2.6, 2.1, 6.0, 2.1, WH, { width:2.5, dashType:'dash', endArrowType:'triangle' });
  L(s, 6.0, 2.85, 2.6, 2.85, OCHRE, { width:2.5, dashType:'dash', endArrowType:'triangle' });
  T(s, 'импульс бежит по штангам и возвращается: прибор меряет время', 2.6, 2.22, 3.7, 0.55, { fontSize:13, color:WH });
  s.addChart(pres.charts.BAR, [{ name:'√ε', labels:['Воздух', 'Минералы', 'Вода'], values:[1, 2.17, 8.94] }], cb({
    x:0.6, y:3.45, w:5.8, h:3.4, barDir:'bar', chartColors:[H.accent1], showValue:true, dataLabelFormatCode:'0.0', dataLabelColor:H.dk1, dataLabelFontSize:14, dataLabelPosition:'outEnd',
    showTitle:true, title:'Во сколько раз волна медленнее, чем в воздухе (√ε)', showLegend:false, valAxisHidden:true, valGridLine:{ style:'none' }, barGapWidthPct:40, catAxisOrientation:'maxMin', catAxisLabelFontSize:14 }));
  T(s, 'Скрытые допущения заводской шкалы', 6.8, 1.5, 5.9, 0.4, { fontSize:16, bold:true, color:MUTED });
  [['1', 'Почва для волны — изолятор', 'пробег удлиняет только вода; сигнал по пути не теряет энергию', TEAL],
   ['2', 'Переводчик универсален', 'заводское уравнение τ → θ верно для любой почвы', TERRA]].forEach((c, k) => {
    const y = 1.95 + k * 1.6;
    card(s, 6.8, y, 5.9, 1.45, LT, 'assumption-' + k);
    circ(s, 7.0, y + 0.35, 0.7, c[3], c[0], 22);
    T(s, [{ text:c[1], options:{ bold:true, fontSize:19, breakLine:true } }, { text:c[2], options:{ fontSize:15, color:MUTED } }], 7.9, y, 4.65, 1.45, { valign:'middle' });
  });
  card(s, 6.8, 5.2, 5.9, 1.65, WH, 'passport', { line:{ color: SLATE, width:1.25 } });
  T(s, [{ text:'Где производитель это гарантирует', options:{ bold:true, color:TEAL, breakLine:true } },
        { text:'объёмная электропроводность < 0,5 дСм/м · плотность < 1,55 г/см³ · глина (< 0,002 мм) < 30 %', options:{ breakLine:true } },
        { text:'В наших тяжёлых почвах физической глины до 65 %', options:{ bold:true, color:TERRA } }], 7.05, 5.25, 5.45, 1.55, { fontSize:15, valign:'middle' });
}

// ================= 4. mechanism + predictions =================
{
  const s = addSlide(4);
  card(s, 0.6, 1.55, 5.5, 3.15, LT, 'dl-card');
  const cx = 1.85, cy = 2.95;
  circ(s, cx - 0.6, cy - 0.6, 1.2, BROWN, '− − −\n− − −', 13);
  for (let k = 0; k < 10; k++) { const a = k * Math.PI / 5; circ(s, cx + 0.88 * Math.cos(a) - 0.14, cy + 0.88 * Math.sin(a) - 0.14, 0.28, TEAL, '+', 12); }
  T(s, 'Частица глины (−) и облако ионов (+) — двойной электрический слой: он проводит ток. Смектит: 600–800 м²/г поверхности.', 3.15, 1.75, 2.8, 2.0, { fontSize:15 });
  T(s, 'Обратная сила: связанная вода тянет показания вниз', 0.85, 4.05, 5.1, 0.55, { fontSize:14, bold:true, color:TERRA, valign:'middle' });
  const t = Array.from({ length:25 }, (_, i) => i);
  s.addChart(pres.charts.LINE, [
    { name:'сильный сигнал (почва-изолятор)', labels:t.map(() => ''), values:t.map(v => +(1 - Math.exp(-v / 2.5)).toFixed(3)) },
    { name:'ослабленный сигнал (проводящая почва)', labels:t.map(() => ''), values:t.map(v => +(0.6 * (1 - Math.exp(-v / 4))).toFixed(3)) },
    { name:'порог срабатывания', labels:t.map(() => ''), values:t.map(() => 0.45) } ],
    cb({ x:6.3, y:1.5, w:6.4, h:2.65, chartColors:[H.accent1, H.accent3, H.accent2], lineSize:3, lineDataSymbol:'none', showLegend:true, legendPos:'r',
      valAxisHidden:true, catAxisHidden:true, valGridLine:{ style:'none' }, valAxisMinVal:0, valAxisMaxVal:1.05, showTitle:true, title:'Схема: прибор срабатывает, когда сигнал пересекает порог' }));
  card(s, 6.3, 4.2, 6.4, 0.5, TEAL, 'mech-result');
  T(s, 'Слабый сигнал позже пересекает порог → τ длиннее → «лишняя вода»', 6.45, 4.2, 6.15, 0.5, { fontSize:14, bold:true, color:WH, valign:'middle' });
  [['H1', 'Ошибка растёт с физической глиной', 'опровергнет: одинакова в лёгких и тяжёлых почвах', TEAL],
   ['H2', 'Растёт с влажностью, нелинейно', 'опровергнет: постоянна или растёт по прямой', TERRA],
   ['H3', 'Плотность влияет, по-разному в почвах', 'опровергнет: эффекта нет или он одинаков везде', BROWN]].forEach((h, k) => {
    const x = 0.6 + k * 4.1;
    card(s, x, 4.9, 3.9, 1.95, LT, 'h-' + k);
    circ(s, x + 0.2, 5.05, 0.6, h[3], h[0], 15);
    T(s, h[1], x + 0.95, 5.0, 2.8, 0.8, { fontSize:16, bold:true, valign:'middle' });
    T(s, h[2], x + 0.2, 5.85, 3.5, 0.9, { fontSize:14, color:MUTED, italic:true });
  });
}

// ================= 5. experiment =================
{
  const s = addSlide(5, 'Опыт и результаты');
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
  card(s, 7.3, 1.5, 5.4, 1.3, DK, 'count');
  T(s, '6 × 3 × 10 = 180', 7.55, 1.52, 4.9, 0.75, { fontSize:34, bold:true, color:WH, valign:'middle' });
  T(s, 'вариантов × колонок × ступеней влажности', 7.55, 2.27, 4.9, 0.45, { fontSize:15, color:OCHRE });
  card(s, 7.3, 2.95, 5.4, 0.95, LT, 'key');
  T(s, [{ text:'Ключ ответов: ', options:{ bold:true, color:TERRA } }, { text:'термостатно-весовой метод (сушка при 105 °C), неопределённость 0,005' }], 7.5, 2.95, 5.0, 0.95, { fontSize:15, valign:'middle' });
  card(s, 7.3, 4.05, 5.4, 1.1, LT, 'metrics');
  T(s, [{ text:'MB — сдвиг, RMSE — типичный промах', options:{ bold:true, breakLine:true } }, { text:'RMSE² = MB² + разброс²; критерий RMSE ≤ 0,02', options:{ color:TEAL, bold:true } }], 7.5, 4.05, 5.0, 1.1, { fontSize:15, valign:'middle' });
  ['Насыщение снизу, 48 ч', 'Ступенчатое иссушение при 20 °C', 'Равновесие 8–36 ч', 'Снимок: период τ + масса колонки'].forEach((t, k) => {
    const x = 0.6 + k * 3.1;
    card(s, x, 5.35, 2.85, 1.25, k === 3 ? TEAL : LT, 'proc-' + k);
    T(s, t, x + 0.2, 5.35, 2.45, 1.25, { fontSize:17, bold:true, color:k === 3 ? WH : DK, valign:'middle' });
    if (k < 3) arrowRight(s, x + 2.88, 5.8, 0.24, 0.35);
  });
  T(s, 'Колонка 152 мм × 300 мм; температура постоянна; на каждой ступени ждём, пока период и масса перестанут меняться', 0.6, 6.68, 12.1, 0.3, { fontSize:13, color:MUTED });
}

// ================= 6. results =================
{
  const s = addSlide(6);
  const X = [0.05, 0.10, 0.20, 0.30, 0.38, 0.40, 0.44, 0.48];
  s.addChart(pres.charts.SCATTER, [
    { name:'θ', values:X },
    { name:'A (10–20 % физ. глины)', values:TV.A.concat([null, null]) },
    { name:'B (40–50 %)', values:TV.B.concat([null, null]) },
    { name:'C (50–65 %)', values:TV.C.concat([null, null]) },
    { name:'C: состояния вне анализа', values:[null, null, null, null, null, 0.027, 0.031, 0.038] },
    { name:'Паспортный предел 0,025', values:X.map(() => 0.025) }
  ], cb({ x:0.6, y:1.5, w:7.7, h:5.35, lineSize:2.25, lineDataSymbolSize:8, chartColors:[H.accent1, H.accent4, H.accent3, 'E3A98F', H.accent5],
    valAxisMinVal:0, valAxisMaxVal:0.04, valAxisMajorUnit:0.005, catAxisMinVal:0, catAxisMaxVal:0.5, catAxisMajorUnit:0.1, valAxisLabelFormatCode:'General',
    showLegend:true, legendPos:'b', showCatAxisTitle:true, catAxisTitle:'влажность по эталону θ, м³/м³', showValAxisTitle:true, valAxisTitle:'ошибка Δθ, м³/м³' }));
  const kp = [['0,006–0,014', 'RMSE во всех вариантах — ниже критерия 0,02', TEAL],
              ['MB > 0', 'везде завышение: в A неотличимо от нуля, в B и C +0,006…+0,010', TERRA],
              ['+0,008 · +0,014 · +0,021', 'прирост при θ 0,20 → 0,40 (A · B · C): в C в 2,6 раза больше, чем в A', BROWN],
              ['+0,027', 'максимум при θ 0,40 в C; бледные точки вне анализа: +0,031 и +0,038', DK]];
  kp.forEach((c, i) => {
    const y = 1.5 + i * 1.36;
    card(s, 8.55, y, 4.15, 1.24, LT, 'kpi-' + i);
    T(s, c[0], 8.75, y + 0.06, 3.8, 0.5, { fontSize:i === 2 ? 21 : 24, bold:true, color:c[2], valign:'middle' });
    T(s, c[1], 8.75, y + 0.58, 3.8, 0.62, { fontSize:13 });
  });
}

// ================= 7. proof: same-moisture comparison =================
{
  const s = addSlide(7, 'Доказательства');
  s.addChart(pres.charts.BAR, [
    { name:'Почва 9,1 %', labels:[''], values:[9.1] }, { name:'Плотность 1,0 %', labels:[''], values:[1.0] },
    { name:'Взаимодействие 0,1 %', labels:[''], values:[0.1] }, { name:'Внутри вариантов 89,8 %', labels:[''], values:[89.8] } ],
    cb({ x:0.6, y:1.45, w:5.6, h:1.85, barDir:'bar', barGrouping:'stacked', chartColors:[H.accent3, H.accent4, H.accent2, 'C9D1CF'], showLegend:true, legendPos:'b',
      valAxisHidden:true, catAxisHidden:true, valGridLine:{ style:'none' }, valAxisMaxVal:100, showTitle:true, title:'Доли изменчивости ошибки (все 60 значений)', barGapWidthPct:20 }));
  card(s, 0.6, 3.45, 5.6, 1.3, DK, 'verdict');
  T(s, [{ text:'Почва: p = 0,074 → «не значимо»', options:{ bold:true, color:WH, breakLine:true } }, { text:'но «не значимо» ≠ «эффекта нет»: его заглушает рост ошибки с влажностью', options:{ color:OCHRE } }], 0.85, 3.45, 5.1, 1.3, { fontSize:16, valign:'middle' });
  card(s, 0.6, 4.9, 5.6, 0.85, LT, 'analogy');
  T(s, [{ text:'Разговор на концерте: ', options:{ bold:true, color:TEAL } }, { text:'музыка — рост ошибки с влажностью' }], 0.85, 4.9, 5.1, 0.85, { fontSize:16, valign:'middle' });
  card(s, 0.6, 5.9, 5.6, 0.95, WH, 'caveat', { line:{ color: OCHRE, width:1.5 } });
  T(s, [{ text:'Граница вывода: ', options:{ bold:true, color:TERRA } }, { text:'уровни — состояния одних колонок, почва в группе одна → описательно' }], 0.85, 5.9, 5.1, 0.95, { fontSize:14, valign:'middle' });
  const seq = ['9FD0D6', '72BAC3', '4CA3AF', '2C8C9A', '1C7484', '0E4F5A'];
  s.addChart(pres.charts.LINE, TV.lv.map((l, i) => ({ name:'θ = ' + l, labels:['A', 'B', 'C'], values:[TV.A[i], TV.B[i], TV.C[i]] })),
    cb({ x:6.5, y:1.45, w:6.2, h:3.75, chartColors:seq, lineSize:3, lineDataSymbol:'circle', lineDataSymbolSize:8, showLegend:true, legendPos:'r',
      valAxisMinVal:0, valAxisMaxVal:0.03, valAxisMajorUnit:0.01, valAxisLabelFormatCode:'General', showTitle:true, title:'На равной влажности все линии идут вверх', catAxisLabelFontSize:15 }));
  [['Влажность — блок', 'p = 0,013'], ['Фридман: только порядок', 'p = 0,004'], ['Интуиция: A меньше всех 6 раз', '(1/3)⁶ ≈ 1 из 729']].forEach((r, k) => {
    const x = 6.5 + k * 2.1;
    card(s, x, 5.35, 2.0, 1.5, k === 2 ? LT : DK, 'test-' + k);
    T(s, [{ text:r[0], options:{ fontSize:13, color:k === 2 ? TEAL : OCHRE, bold:true, breakLine:true } }, { text:r[1], options:{ fontSize:k === 2 ? 15 : 20, color:k === 2 ? DK : WH, bold:true } }], x + 0.15, 5.35, 1.75, 1.5, { valign:'middle' });
  });
}

// ================= 8. why: clay or salts; density =================
{
  const s = addSlide(8);
  T(s, 'Глина или соли?', 0.6, 1.45, 5.5, 0.4, { fontSize:17, bold:true, color:MUTED });
  [['R² = 0,88', 'сдвиг MB растёт с гигроскопичностью MH — прокси поверхности глин', TEAL],
   ['R² = 0,96', 'сдвиг MB растёт с электропроводностью ECe — прокси солей', BROWN],
   ['r = 0,88', 'MH и ECe связаны: два подозреваемых всегда вместе — не разделить', TERRA]].forEach((c, k) => {
    const y = 1.9 + k * 1.2;
    card(s, 0.6, y, 5.5, 1.08, LT, 'suspect-' + k);
    T(s, c[0], 0.8, y, 1.7, 1.08, { fontSize:22, bold:true, color:c[2], valign:'middle' });
    T(s, c[1], 2.55, y, 3.45, 1.08, { fontSize:14, valign:'middle' });
  });
  card(s, 0.6, 5.55, 5.5, 1.3, DK, 'need');
  T(s, 'Нужны почвы, где глина и соли не ходят парой', 0.85, 5.55, 5.0, 1.3, { fontSize:18, bold:true, color:WH, valign:'middle' });
  T(s, 'Плотность: прогноз модели смеси CRIM', 6.5, 1.45, 6.2, 0.4, { fontSize:17, bold:true, color:MUTED });
  s.addChart(pres.charts.BAR, [
    { name:'Прогноз CRIM', labels:['A', 'B', 'C'], values:[9.5, 9.5, 9.5] },
    { name:'Наблюдалось', labels:['A', 'B', 'C'], values:[-2, -1, -2] } ],
    cb({ x:6.5, y:1.85, w:6.2, h:2.85, barDir:'col', chartColors:[H.accent5, H.accent3], barGapWidthPct:60, showValue:true, dataLabelFontSize:13, dataLabelColor:H.dk1, dataLabelPosition:'outEnd', dataLabelFormatCode:'+0.0;−0.0',
      valAxisMinVal:-4, valAxisMaxVal:12, valAxisMajorUnit:4, showLegend:true, legendPos:'r', catAxisLabelPos:'low', showValAxisTitle:true, valAxisTitle:'×10⁻³ м³/м³' }));
  card(s, 6.5, 4.85, 6.2, 2.0, DK, 'logic');
  T(s, [{ text:'модель ∧ та же θ ∧ та же проводимость ⇒ +0,0095', options:{ bold:true, color:WH, breakLine:true } },
        { text:'наблюдали −0,0017 (ДИ −0,006…+0,003) ⇒ ложно хотя бы одно звено', options:{ color:OCHRE, breakLine:true } },
        { text:'ECe у плотных ниже на 0,07–0,12 дСм/м; диапазоны влажности разные', options:{ color:WH } }], 6.75, 4.85, 5.75, 2.0, { fontSize:15, valign:'middle' });
}

// ================= 9. evidence ledger =================
{
  const s = addSlide(9);
  const hd = { bold:true, color:WH, fill:{ color: DK }, fontSize:14, valign:'middle' };
  const c = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:14, valign:'middle', color:DK, fill:{ color: LT } }, o) });
  const dots = { 'высокая':['●●●', TEAL], 'средняя':['●●○', BROWN], 'низкая':['●○○', TERRA] };
  const rows = [[{ text:'Утверждение', options:hd }, { text:'Тип связи', options:hd }, { text:'Твёрдость', options:Object.assign({}, hd, { align:'center' }) }, { text:'Что мешает сказать сильнее', options:hd }]];
  K3.ledger.forEach(r => rows.push([c(r[0], { bold:true }), c(r[2], { color:MUTED }), c(dots[r[3]][0], { fontSize:20, align:'center', color:dots[r[3]][1] }), c(r[4])]));
  s.addTable(rows, { x:0.6, y:1.5, w:12.1, colW:[3.9, 2.0, 1.4, 4.8], rowH:[0.42].concat(K3.ledger.map(() => 0.64)), border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.07, objectName:'ledger' });
  T(s, 'Допущения нашего собственного опыта (эталон ±0,005, уровни не независимы, одна почва на группу, стенки колонки) — на запасном слайде', 0.6, 6.55, 12.1, 0.4, { fontSize:13, color:MUTED, italic:true });
}

// ================= 10. traffic light =================
{
  const s = addSlide(10, 'Что делать');
  [[GREEN, 'Лёгкие, 10–20 % физ. глины', 'заводская калибровка без поправки'],
   [OCHRE, 'Тяжёлые, 40–65 %', 'завышение 6–10 мм воды в метровом слое, до 27 мм во влажной: калибровка под почву или 1–2 весовые проверки за сезон'],
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
  T(s, [{ text:'Проверять в двух точках: ', options:{ bold:true, color:OCHRE } }, { text:'после полива и перед поливом. Одна точка даёт сдвиг, две — ещё и наклон. Следствие из данных; в поле не проверено.', options:{ color:WH } }], 8.0, 5.15, 4.5, 1.7, { fontSize:15, valign:'middle' });
}

// ================= 11. if-then =================
{
  const s = addSlide(11);
  const hd = { bold:true, color:WH, fill:{ color: DK }, fontSize:14, valign:'middle' };
  const c = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:15, valign:'middle', color:DK, fill:{ color: LT } }, o) });
  const rows = [
    [{ text:'Если верно…', options:hd }, { text:'…то увидим', options:hd }, { text:'Решающий опыт', options:hd }],
    [c('Завышение создаёт проводимость порового раствора', { bold:true }), c('ошибка растёт вместе с ECe при засолении одной и той же почвы'), c('ступенчатое засоление, ECe 1–4 дСм/м (сделано, отдельная статья)')],
    [c('Работает поверхность глин', { bold:true }), c('при равной ECe ошибка больше в почве с большей долей смектита'), c('минералогия глин; пары почв с равной ECe')],
    [c('Эффект плотности «съела» разница в ECe', { bold:true }), c('при выровненной ECe плотный вариант ближе к +0,009…+0,010'), c('уплотнение одной почвы при одинаковой ECe')],
    [c('Эффект почвы реален, а не артефакт колонок', { bold:true }), c('сохранится в смешанной модели (колонка — случайный эффект)'), c('переанализ 180 колоночных значений — данные уже есть', { bold:true, color:TEAL })],
    [c('Идея: завышение задаёт проводимость', { bold:true, color:TERRA }), c('поправка по двум входам (τ и объёмная ЭП) уберёт большую часть ошибки'), c('мерить объёмную ЭП рядом с датчиком (в опыте не мерили)')]
  ];
  s.addTable(rows, { x:0.6, y:1.5, w:12.1, colW:[3.7, 4.4, 4.0], rowH:[0.45, 0.95, 0.85, 0.95, 0.95, 0.95], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.1, objectName:'if-then' });
}

// ================= 12. three take-aways =================
{
  const s = addSlide(12);
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
let bi = 0;
function addBackup(title, note) {
  if (bi++ === 0) { pres.addSection({ title:'Запасные слайды' }); curSection = 'Запасные слайды'; }
  const s = pres.addSlide({ masterName:'LIGHT', sectionTitle:curSection });
  s.addText(title, { placeholder:'title' });
  s.addNotes('[Запасной слайд — показывать только при вопросах; во время доклада не используется]\n\n' + note);
  T(s, 'Запасные слайды — для вопросов', 0.6, 7.02, 6, 0.3, { fontSize:11, color:SLATE, valign:'middle' });
  return s;
}
// B1 theory map
{
  const s = addBackup('Запасной: шесть теорий и типы связей', 'Показывать, если спрашивают, на чём стоит работа и как связаны теории. Связи четырёх типов: механизм, количественная модель, косвенный показатель, статистический вывод; у каждого своё слабое место.');
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
  L(s, 2.7, 2.3, 3.1, 2.3, OCHRE, A); L(s, 2.7, 4.05, 3.1, 4.05, TEAL, A); L(s, 4.15, 3.35, 4.15, 3.0, OCHRE, A);
  L(s, 5.2, 2.3, 5.6, 2.9, OCHRE, A); L(s, 5.2, 4.05, 5.6, 3.45, TEAL, A); L(s, 7.7, 2.9, 8.1, 2.3, SLATE, A);
  L(s, 9.15, 3.0, 9.15, 3.35, SLATE, A); L(s, 10.2, 4.05, 10.6, 3.45, SLATE, A);
  [[TEAL, 'Механизм', 'причина → следствие. Подводит, если есть обратная сила (связанная вода)', null],
   [OCHRE, 'Модель', 'даёт число, его можно проверить. Подводит, если нарушены допущения (плотность)', null],
   [TERRA, 'Косвенный показатель', 'судим о скрытом по видимому (MH, ECe). Подводит, если прокси ходят парой', 'dash'],
   [SLATE, 'Статистический вывод', 'от колонок к общему. Подводит, если наблюдения не независимы', null]].forEach((g, k) => {
    const x = 0.6 + k * 3.07;
    card(s, x, 5.0, 2.92, 1.85, LT, 'legend-' + k);
    L(s, x + 0.2, 5.3, x + 0.8, 5.3, g[0], { width:3, dashType:g[3] || 'solid' });
    T(s, g[1], x + 0.9, 5.14, 1.95, 0.35, { fontSize:15, bold:true, valign:'middle' });
    T(s, g[2], x + 0.2, 5.58, 2.6, 1.22, { fontSize:14 });
  });
}
// B2 own assumptions
{
  const s = addBackup('Запасной: допущения нашего опыта', 'Показывать, если спрашивают о надёжности эталона, статистики или колонок. Эта таблица объясняет, почему p-значения названы описательными, а выводы относятся к изученным почвам.');
  const hd = { bold:true, color:WH, fill:{ color: DK }, fontSize:14, valign:'middle' };
  const c = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:14, valign:'middle', color:DK, fill:{ color: LT } }, o) });
  const st = (t, col, txt = WH) => ({ text:t, options:{ fontSize:14, bold:true, align:'center', valign:'middle', color:txt, fill:{ color: col } } });
  s.addTable([
    [{ text:'Допущение', options:hd }, { text:'Что мы сделали или знаем', options:hd }, { text:'Статус', options:Object.assign({}, hd, { align:'center' }) }],
    [c('«Эталон — это истина»', { bold:true }), c('Два способа (вся колонка и микрокерны): расхождение ≤ 0,008, в среднем 0,003; U = 0,005'), st('частично', OCHRE, DK)],
    [c('«Десять уровней — независимые наблюдения»', { bold:true }), c('Нет: это последовательные состояния одних и тех же колонок'), st('p описательные', TERRA)],
    [c('«Одна почва представляет группу»', { bold:true }), c('В каждой группе одна почва'), st('экстраполяция', TERRA)],
    [c('«Колонка не мешает датчику»', { bold:true }), c('До стенки 59–73 мм (рекомендуется ≥ 100), штанги доходят до дна'), st('не проверено', TERRA)],
    [c('«Объём почвы постоянен»', { bold:true }), c('Нарушено при набухании (почва C, 1,52 г/см³) — эталон по реальному объёму'), st('учтено', GREEN)],
    [c('«Округление опубликованных чисел не меняет выводов»', { bold:true }), c('20 000 симуляций: доля эффекта почвы 0,07–0,11 держится; граница A–B неустойчива'), st('проверено', GREEN)]
  ], { x:0.6, y:1.6, w:12.1, colW:[3.6, 6.0, 2.5], rowH:[0.45, 0.72, 0.72, 0.72, 0.72, 0.72, 0.72], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.08, objectName:'own-assumptions' });
}
// B3 CRIM calculation
{
  const s = addBackup('Запасной: расчёт прогноза модели смеси', 'Показывать, если просят объяснить, откуда +0,0095. Модель: время пробега складывается из времён через фазы; уплотнение при той же влажности заменяет воздух минералами; заводская шкала читает удлинение пробега как воду вместо воздуха.');
  s.addChart(pres.charts.BAR, [
    { name:'Твёрдая фаза', labels:['1,35 г/см³', '1,52 г/см³'], values:[0.513, 0.578] },
    { name:'Вода (θ = 0,25)', labels:['1,35 г/см³', '1,52 г/см³'], values:[0.25, 0.25] },
    { name:'Воздух', labels:['1,35 г/см³', '1,52 г/см³'], values:[0.237, 0.172] } ],
    cb({ x:0.6, y:1.55, w:5.6, h:5.3, barDir:'col', barGrouping:'stacked', chartColors:['B9C3C1', '8FC5CD', 'EEF1F0'], showValue:true, dataLabelPosition:'ctr', dataLabelFormatCode:'0.00', dataLabelColor:H.dk1, dataLabelFontSize:14,
      valAxisMinVal:0, valAxisMaxVal:1, valAxisMajorUnit:0.2, valAxisLabelFormatCode:'0.0', showLegend:true, legendPos:'b', showTitle:true, title:'Доли объёма при θ = 0,25 (ρs = 2,63 г/см³)', barGapWidthPct:50, catAxisLabelFontSize:14 }));
  ['Модель: √ε смеси = φтв·√4,7 + θ·√80 + φвозд·√1',
   'Уплотнение 1,35 → 1,52: твёрдого больше на 0,17 / 2,63 ≈ 0,065 — вместо воздуха',
   'Пробег растёт: 0,065 × (√4,7 − 1) = 0,065 × 1,17 ≈ 0,076',
   'Датчик читает это как воду вместо воздуха: 0,076 / (√80 − 1) = 0,076 / 7,94 ≈ 0,0095'].forEach((t, k) => {
    const y = 1.55 + k * 0.97;
    card(s, 6.6, y, 6.1, 0.85, LT, 'step-' + k);
    circ(s, 6.75, y + 0.17, 0.5, TEAL, String(k + 1), 16);
    T(s, t, 7.4, y, 5.2, 0.85, { fontSize:15, valign:'middle', bold:k === 3 });
  });
  card(s, 6.6, 5.5, 6.1, 0.85, DK, 'prediction');
  T(s, [{ text:'Прогноз: ≈ +0,0095 ', options:{ bold:true, color:OCHRE } }, { text:'(при ρs 2,60–2,67: +0,0094…+0,0096)', options:{ color:WH } }], 6.85, 5.5, 5.7, 0.85, { fontSize:17, valign:'middle' });
  T(s, 'Скрыто: та же θ, та же проводимость, нет потерь в пути', 6.6, 6.45, 6.1, 0.42, { fontSize:15, bold:true, color:TERRA });
}
// B4 swelling
{
  const s = addBackup('Запасной: набухание и пересчёт эталона', 'Показывать, если спрашивают о набухании. Почва C при 1,52 г/см³ набухла: подъём ≈ 25 мм, объёмная деформация 8,3 %, θmax 0,460 > исходной пористости 0,415. Эталон считали по реальному объёму; дополнительной ошибки не нашли, но сравнение смешано с плотностью и ECe, а почва поднималась выше штанг.');
  const draw = (x, title) => {
    T(s, title, x, 1.5, 3.8, 0.4, { fontSize:17, bold:true, align:'center' });
    s.addShape(pres.ShapeType.rect, { x:x + 0.7, y:2.35, w:2.4, h:3.2, fill:{ color: TERRA }, line:{ type:'none' }, objectName:'soil' });
    s.addShape(pres.ShapeType.rect, { x:x + 0.7, y:2.35, w:2.4, h:3.2, fill:{ type:'none' }, line:{ color: SLATE, width:2 }, objectName:'column' });
    [1.4, 2.4].forEach(dx => s.addShape(pres.ShapeType.rect, { x:x + dx, y:2.35, w:0.07, h:3.2, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
  };
  draw(0.6, 'Сухая: 300 мм'); draw(4.7, 'Влажная: ≈ 325 мм');
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
// B5 terms trap
{
  const s = addBackup('Запасной: ловушка терминов «глина»', 'Показывать, если спрашивают о сравнении с паспортом и зарубежными работами. Физическая глина по Качинскому — частицы < 0,01 мм; глина по USDA и в паспорте — < 0,002 мм. В группе A условие паспорта (< 30 % глины) выполнено заведомо; для B и C проверить нельзя.');
  T(s, 'Размер частиц, мм (в масштабе)', 0.6, 1.55, 6, 0.35, { fontSize:14, bold:true, color:MUTED });
  [['глина по Качинскому < 0,001', 0.001, SLATE], ['«clay» по USDA и в паспорте < 0,002', 0.002, TEAL], ['физическая глина по Качинскому < 0,01', 0.01, TERRA]].forEach((b, k) => {
    const y = 2.0 + k * 0.95;
    s.addShape(pres.ShapeType.rect, { x:0.6, y, w:b[1] * 900, h:0.6, fill:{ color: b[2] }, line:{ type:'none' }, objectName:'size-bar' });
    if (b[1] > 0.005) T(s, b[0], 0.85, y, 8, 0.6, { fontSize:17, bold:true, valign:'middle', color:WH });
    else T(s, b[0], 0.6 + b[1] * 900 + 0.2, y, 7, 0.6, { fontSize:17, bold:true, valign:'middle' });
  });
  card(s, 0.6, 4.9, 5.9, 1.95, LT, 'trap');
  T(s, [{ text:'Ловушка перевода', options:{ bold:true, color:TERRA, breakLine:true } }, { text:'30 % «clay» из паспорта и 30 % физической глины — разные почвы. «Лёгкая глина» по Качинскому ≠ «clay» по USDA.' }], 0.85, 5.0, 5.4, 1.75, { fontSize:17 });
  card(s, 6.8, 4.9, 5.9, 1.95, LT, 'consequence');
  T(s, [{ text:'Что из этого следует', options:{ bold:true, color:TEAL, breakLine:true } }, { text:'В группе A физической глины ≤ 20 %, значит, и USDA-глины ≤ 20 %: условие паспорта выполнено. Для B и C проверить нельзя.' }], 7.05, 5.0, 5.4, 1.75, { fontSize:17 });
}
// B6–B9 tables and limits
const tblHd = t => ({ text:t, options:{ bold:true, color:WH, fill:{ color: DK }, fontSize:14, align:'center', valign:'middle' } });
const tc = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:14, align:'center', valign:'middle', color:DK, fill:{ color: LT } }, o) });
{
  const s = addBackup('Запасной: статистика ошибки по вариантам', K3.backup[0].note);
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
  const s = addBackup('Запасной: дисперсионный анализ ошибки', K3.backup[1].note);
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
  const s = addBackup('Запасной: свойства почв в вариантах', K3.backup[2].note);
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
  const s = addBackup('Запасной: ограничения исследования', K3.backup[3].note);
  const half = Math.ceil(K3.limits.length / 2);
  [K3.limits.slice(0, half), K3.limits.slice(half)].forEach((col, j) => {
    card(s, 0.6 + j * 6.2, 1.6, 5.9, 4.6, LT, 'limits-' + j);
    T(s, col.map((t, i) => ({ text:`${j * half + i + 1}. ${t}`, options:{ breakLine:i < col.length - 1, paraSpaceAfter:10 } })), 0.85 + j * 6.2, 1.8, 5.4, 4.2, { fontSize:16 });
  });
}

async function patchSeries(file) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const rules = [
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
  const out = path.resolve(process.argv[2] || 'deck4.pptx');
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  await patchSeries(out);
  const total = secs.reduce((a, b) => a + b, 0);
  console.log('written', out, '; speech', mmss(total), '; with transitions', mmss(total + 4 * K.slides.length));
})();
