// Версия 6: профессиональный дизайн. 12 слайдов на 10 минут + 9 запасных.
// Графики нарисованы фигурами (а не диаграммами), чтобы подписи стояли точно у линий и одинаково выглядели в PowerPoint и LibreOffice.
'use strict';
const pptxgen = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const sharp = require('sharp');
const React = require('react');
const { renderToStaticMarkup } = require('react-dom/server');
const TB = require('react-icons/tb');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');
const ILL = require('./illus6.js');
const K3 = require('./content3.js');           // проверенные данные версии 3 (таблицы, ограничения)
const K = require('./content6.json');          // текст версии 6 (= версия 5 + подсказки под новый дизайн)

const THEME = {
  name: 'CS616 Soil and Water', headFontFace: 'Cambria', bodyFontFace: 'Calibri',
  colors: { dk1:'231C17', lt1:'FFFFFF', dk2:'5E5650', lt2:'F3F2F0',
    accent1:'008C9E', accent2:'D9A441', accent3:'C0562B', accent4:'9A6440', accent5:'7A8587', accent6:'4E8A5B', hlink:'008C9E', folHlink:'9A6440' }
};
const HX = { ink:'231C17', ink2:'5E5650', teal:'008C9E', ochre:'D9A441', terra:'C0562B', brown:'9A6440', slate:'7A8587', green:'4E8A5B', white:'FFFFFF' };
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = 'Когда датчик влажности «льстит» почве: CS616 в незасолённых почвах (доклад 10 минут)';
pres.author = 'Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е.';
pres.company = 'Кубанский ГАУ';
const C = pres.SchemeColor;
const INK = C.text1, INK2 = C.text2, WH = C.background1, PANEL = C.background2;
const TEAL = C.accent1, OCHRE = C.accent2, TERRA = C.accent3, BROWN = C.accent4, SLATE = C.accent5, GREEN = C.accent6;
const HAIR = 'E3DFDA', GRID = 'ECE8E4', AXIS = 'B5ADA5', DARKLINE = '453A32', DARKTEXT = 'D8D1C9';
const SOIL = ['C49A6C', '9A6440', '6B3220'], SOIL_T = ['DFC7AE', 'C7AA96', 'AE8E84'];
const WATER = ['7FBFC6', '5EAAB4', '3F93A0', '26798A', '165F70', '0A4452'];
const MX = 0.75, SW = 13.333, CW = SW - 2 * MX, RX = SW - MX;

// ---------- мастер-слайды ----------
const titlePh = (color, y = 0.74, h = 0.8, fs = 30, w = CW) => ({ placeholder:{ options:{ name:'title', type:'title', x:MX, y, w, h, fontSize:fs, bold:true, color,
  align:'left', valign:'top', margin:0, fontFace:'Cambria' }, text:'' } });
const footRule = color => ({ line:{ x:MX, y:6.96, w:CW, h:0, line:{ color, width:0.75 } } });
const slideNo = color => ({ x:RX - 0.6, y:7.02, w:0.6, h:0.3, fontSize:10, color, align:'right' });
pres.defineSlideMaster({ title:'CONTENT', background:{ color:'FFFFFF' }, objects:[footRule(HAIR), titlePh(INK)], slideNumber:slideNo(HX.ink2) });
pres.defineSlideMaster({ title:'DARK', background:{ color:HX.ink }, objects:[footRule(DARKLINE), titlePh(WH)], slideNumber:slideNo(HX.ochre) });
pres.defineSlideMaster({ title:'COVER', background:{ color:'FFFFFF' }, objects:[titlePh(INK, 1.4, 2.1, 40, 7.2)] });
pres.defineSlideMaster({ title:'BACKUP', background:{ color:'FFFFFF' }, objects:[footRule(HAIR), titlePh(INK)], slideNumber:slideNo(HX.ink2) });

// ---------- примитивы ----------
function T(s, text, x, y, w, h, o = {}) { s.addText(text, Object.assign({ x, y, w, h, fontSize:16, color:INK, margin:0, valign:'top', isTextBox:true }, o)); }
function R(s, x, y, w, h, fill, o = {}) { s.addShape(pres.ShapeType.rect, Object.assign({ x, y, w, h, fill:{ color:fill } }, o)); }
function RR(s, x, y, w, h, fill, r = 0.08, o = {}) { s.addShape(pres.ShapeType.roundRect, Object.assign({ x, y, w, h, fill:{ color:fill }, rectRadius:r }, o)); }
function L(s, x1, y1, x2, y2, color, o = {}) {
  s.addShape(pres.ShapeType.line, { x:Math.min(x1, x2), y:Math.min(y1, y2), w:Math.abs(x2 - x1), h:Math.abs(y2 - y1), flipH:x2 < x1, flipV:y2 < y1,
    line:{ color, width:o.width || 1.5, dashType:o.dashType, endArrowType:o.endArrowType, beginArrowType:o.beginArrowType }, objectName:o.name || 'line' });
}
const hline = (s, x1, y, x2, color = HAIR, width = 0.75) => L(s, x1, y, x2, y, color, { width, name:'rule' });
const vline = (s, x, y1, y2, color = HAIR, width = 0.75) => L(s, x, y1, x, y2, color, { width, name:'rule' });
// ломаная (custGeom) в координатах слайда
function poly(s, pts, color, o = {}) {
  const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys);
  const w = Math.max(Math.max(...xs) - x0, 0.001), h = Math.max(Math.max(...ys) - y0, 0.001);
  s.addShape(pres.ShapeType.custGeom, { x:x0, y:y0, w, h, objectName:o.name || 'curve',
    points:pts.map((p, i) => Object.assign({ x:p[0] - x0, y:p[1] - y0 }, i === 0 ? { moveTo:true } : {})),
    line:{ color, width:o.width || 2, dashType:o.dashType, endArrowType:o.endArrowType } });
}
function dot(s, cx, cy, d, fill, o = {}) {
  s.addShape(pres.ShapeType.ellipse, { x:cx - d / 2, y:cy - d / 2, w:d, h:d, fill:{ color:fill }, line:{ color:o.ring || 'FFFFFF', width:o.ringW === undefined ? 1.25 : o.ringW }, objectName:o.name || 'marker' });
}
function badge(s, cx, cy, d, fill, label, fs, color = WH) {
  s.addShape(pres.ShapeType.ellipse, { x:cx - d / 2, y:cy - d / 2, w:d, h:d, fill:{ color:fill }, objectName:'badge' });
  T(s, label, cx - d / 2, cy - d / 2, d, d, { fontSize:fs, bold:true, color, align:'center', valign:'middle' });
}
function kicker(s, text, dark) { T(s, text, MX, 0.42, CW, 0.28, { fontSize:12, bold:true, color:dark ? OCHRE : TEAL, charSpacing:1.5 }); }
function label(s, text, x, y, w, o = {}) { T(s, text, x, y, w, 0.28, Object.assign({ fontSize:11, bold:true, color:INK2, charSpacing:1 }, o)); }
// подстрочные индексы: 't|1' → t₁ (через run со subscript)
const sub = (base, idx, o = {}) => [{ text:base, options:o }, { text:idx, options:Object.assign({}, o, { subscript:true }) }];

// ---------- иконки (Tabler) ----------
const ICONS = {};
async function prepIcons(list) {
  for (const [name, color] of list) {
    const key = name + color; if (ICONS[key]) continue;
    if (!TB[name]) throw new Error('нет иконки ' + name);
    let svg = renderToStaticMarkup(React.createElement(TB[name], { size:256, color:'#' + color }));
    svg = svg.replace('stroke-width="2"', 'stroke-width="1.8"');
    ICONS[key] = 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
  }
}
function icon(s, name, color, x, y, size) {
  const data = ICONS[name + color]; if (!data) throw new Error('иконка не подготовлена: ' + name + ' ' + color);
  s.addImage({ data, x, y, w:size, h:size, altText:name.replace(/^Tb/, '') });
}
const ICON_LIST = [
  ['TbCertificate', HX.ink2], ['TbClock', HX.teal], ['TbBook', HX.teal], ['TbDroplet', HX.teal], ['TbAlertTriangle', HX.terra],
  ['TbCube', HX.brown], ['TbScale', HX.ink2], ['TbScale', HX.teal], ['TbTarget', HX.teal], ['TbGauge', HX.teal], ['TbMusic', HX.terra],
  ['TbVolumeOff', HX.teal], ['TbAlertCircle', HX.ochre], ['TbDeviceCctv', HX.ink2], ['TbLinkOff', HX.terra], ['TbRuler2', HX.ink2],
  ['TbEye', HX.ink2], ['TbChartDots', HX.ink2], ['TbMathFunction', HX.ink2], ['TbFingerprint', HX.ink2], ['TbArrowRight', HX.teal],
  ['TbBulb', HX.teal], ['TbMessageQuestion', HX.terra], ['TbDatabase', HX.white]
];

// ---------- хронометраж и заметки ----------
const WPM = 130;
// числа вслух длиннее: «0,014» ≈ 3–4 слова, диапазон «0,006–0,010» ≈ 5
const spoken = t => t.split(/\s+/).filter(Boolean).reduce((n, w) => n + (/\d/.test(w) ? (/[–…]/.test(w) ? 5 : 3) : 1), 0);
const secs = K.slides.map(d => Math.round(spoken(d.text.join(' ')) / WPM * 60));
const mmss = t => `${Math.floor(t / 60)}:${String(t % 60).padStart(2, '0')}`;
const ACTS = ['Зачем', 'Как проверяли', 'Что нашли', 'Как доказали', 'Что делать'];
const ACT_OF = n => n === 1 || n === 12 ? null : n <= 4 ? 'Зачем' : n === 5 ? 'Как проверяли' : n === 6 ? 'Что нашли' : n <= 9 ? 'Как доказали' : 'Что делать';
const KICK = { 2:'ЦЕНА ВОПРОСА', 3:'КАК РАБОТАЕТ ДАТЧИК', 4:'ГИПОТЕЗА И ПРЕДСКАЗАНИЯ', 5:'ЭКСПЕРИМЕНТ', 6:'РЕЗУЛЬТАТ',
  7:'ДОКАЗАТЕЛЬСТВО 1 · ЭФФЕКТ ПОЧВЫ', 8:'ДОКАЗАТЕЛЬСТВО 2 · МЕХАНИЗМ И ПЛОТНОСТЬ', 9:'ЧЕСТНАЯ ОЦЕНКА', 10:'ПРАКТИКА', 11:'ПЕРСПЕКТИВЫ', 12:'ИТОГ' };
function tracker(s, act) {
  if (!act) return;
  const runs = [];
  ACTS.forEach((a, i) => {
    const on = a === act;
    runs.push({ text:a, options:{ bold:on, color:on ? TEAL : 'A39B94' } });
    if (i < ACTS.length - 1) runs.push({ text:'    ', options:{} });
  });
  T(s, runs, MX, 7.02, 9, 0.3, { fontSize:10, valign:'middle' });
}
let curSection = null;
function addSlide(n, section, titleRuns) {
  const d = K.slides[n - 1];
  if (section) { pres.addSection({ title:section }); curSection = section; }
  const master = n === 1 ? 'COVER' : n === 12 ? 'DARK' : 'CONTENT';
  const s = pres.addSlide({ masterName:master, sectionTitle:curSection });
  s.addText(titleRuns || d.title, { placeholder:'title' });
  if (KICK[n]) kicker(s, KICK[n], master === 'DARK');
  const cum = secs.slice(0, n).reduce((a, b) => a + b, 0) + 4 * n; // + ~4 с на смену слайда
  s.addNotes(`[≈ ${secs[n - 1]} с · к концу слайда ≈ ${mmss(cum)} из 10:00]\n\nПоказать: ${d.cue}\n\nГлавная мысль: ${d.takeaway}\n\n` + d.text.join('\n\n'));
  tracker(s, ACT_OF(n));
  return s;
}

// ---------- данные (табл. V статьи, плотность 1,35 г/см³) ----------
const TV = { th:[0.05, 0.10, 0.20, 0.30, 0.38, 0.40], lv:['0,05', '0,10', '0,20', '0,30', '0,38', '0,40'],
  A:[0.000, 0.001, 0.003, 0.005, 0.010, 0.011], B:[0.001, 0.002, 0.005, 0.009, 0.017, 0.019], C:[0.001, 0.002, 0.006, 0.012, 0.022, 0.027] };
const lin = (d0, d1, r0, r1) => v => r0 + (v - d0) / (d1 - d0) * (r1 - r0);

function build(ill) {
// ================= 1. Титул =================
{
  const s = addSlide(1, 'Доклад (10 минут)', [{ text:'Когда датчик', options:{ breakLine:true } }, { text:'влажности', options:{ breakLine:true } }, { text:'«льстит» почве' }]);
  s.addImage({ data:ill.cover, x:8.1, y:0, w:5.233, h:7.5, altText:'Почвенный керн с датчиком CS616: импульс бежит по штангам и возвращается' });
  T(s, 'ДАТЧИКИ ВЛАЖНОСТИ CS616 · ТРИ ПОЧВЫ КУБАНИ', MX, 1.0, 7.0, 0.3, { fontSize:12, bold:true, color:TEAL, charSpacing:1.5 });
  T(s, 'Систематическая погрешность заводской калибровки датчиков влажности CS616 в незасолённых почвах разного гранулометрического состава',
    MX, 3.62, 6.6, 1.0, { fontSize:17, color:INK2 });
  hline(s, MX, 5.0, MX + 0.7, TEAL, 2.5);
  T(s, 'Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е.', MX, 5.2, 7.0, 0.35, { fontSize:15, bold:true });
  T(s, 'Кубанский государственный аграрный университет имени И. Т. Трубилина', MX, 5.6, 7.0, 0.3, { fontSize:13, color:INK2 });
  T(s, 'Краснодар, 2026', MX, 5.9, 7.0, 0.3, { fontSize:13, color:INK2 });
}

// ================= 2. Цена вопроса =================
{
  const s = addSlide(2);
  T(s, 'Ошибка 0,01 м³/м³ — это', MX, 1.95, 5, 0.3, { fontSize:15, color:INK2 });
  T(s, '10 мм', MX, 2.2, 5, 0.95, { fontSize:66, bold:true, valign:'middle' });
  T(s, [{ text:'воды в метровом слое', options:{ bold:true, color:INK, breakLine:true } }, { text:'1 л на 100 л почвы · 100 м³ на гектар' }], MX, 3.15, 4.7, 0.62, { fontSize:15, color:INK2 });
  hline(s, MX, 3.98, MX + 4.7);
  T(s, 'Наш опыт: самая тяжёлая почва, влажный край', MX, 4.15, 5, 0.3, { fontSize:15, color:INK2 });
  T(s, '27 мм', MX, 4.4, 5, 0.95, { fontSize:66, bold:true, color:TERRA, valign:'middle' });
  T(s, [{ text:'воды, которой нет: ' }, { text:'эталон 0,40 → датчик 0,427', options:{ bold:true, color:INK } }], MX, 5.38, 4.7, 0.35, { fontSize:15, color:INK2 });
  icon(s, 'TbCertificate', HX.ink2, MX, 6.14, 0.36);
  T(s, [{ text:'Паспорт ±0,025 проверен в основном на супесях; ' }, { text:'у нас физической глины до 65 %', options:{ bold:true, color:INK } }],
    MX + 0.5, 6.02, 4.4, 0.6, { fontSize:13.5, color:INK2, valign:'middle' });

  label(s, 'СХЕМА · ПОЧЕМУ ЗАВЫШЕНИЕ ОПАСНО ДЛЯ ПОЛИВА', 6.25, 1.95, 6.33);
  const X0 = 6.55, X1 = 12.5, YT = 2.45, YB = 5.68, THR = 4.55;
  L(s, X0, YB, X0, YT, AXIS, { width:1, endArrowType:'triangle' });
  L(s, X0, YB, X1, YB, AXIS, { width:1, endArrowType:'triangle' });
  T(s, 'влажность', X0 - 1.02, (YT + YB) / 2 - 0.14, 1.6, 0.28, { fontSize:12, color:INK2, rotate:270, align:'center' });
  T(s, 'время', X1 - 0.8, YB + 0.07, 0.8, 0.26, { fontSize:12, color:INK2, align:'right' });
  L(s, X0, THR, X1 - 0.25, THR, OCHRE, { width:2, dashType:'dash', name:'threshold' });
  T(s, 'порог полива', X0 + 0.12, THR - 0.33, 1.8, 0.28, { fontSize:12.5, bold:true });
  const tru = [[6.65, 3.0], [12.25, 5.62]], sen = [[6.65, 2.55], [12.25, 5.42]];
  const cross = ([[x1, y1], [x2, y2]]) => x1 + (THR - y1) * (x2 - x1) / (y2 - y1);
  const c1 = cross(tru), c2 = cross(sen);
  L(s, ...tru[0], ...tru[1], TEAL, { width:2.75, name:'true-theta' });
  L(s, ...sen[0], ...sen[1], TERRA, { width:2.75, dashType:'dash', name:'sensor-theta' });
  L(s, 9.6, 2.62, 10.05, 2.62, TEAL, { width:2.75 });
  T(s, 'фактическая влажность (эталон)', 10.15, 2.49, 2.43, 0.26, { fontSize:12 });
  L(s, 9.6, 2.95, 10.05, 2.95, TERRA, { width:2.75, dashType:'dash' });
  T(s, 'показания датчика', 10.15, 2.82, 2.43, 0.26, { fontSize:12 });
  L(s, c1, THR, c1, YB, TEAL, { width:1.25, dashType:'sysDot' });
  L(s, c2, THR, c2, YB, TERRA, { width:1.25, dashType:'sysDot' });
  R(s, c1, YB - 0.045, c2 - c1, 0.09, TERRA, { objectName:'delay' });
  dot(s, c1, THR, 0.18, TEAL); dot(s, c2, THR, 0.18, TERRA);
  T(s, sub('t', '1', { bold:true }), c1 - 0.3, YB + 0.07, 0.6, 0.3, { fontSize:14, align:'center' });
  T(s, sub('t', '2', { bold:true }), c2 - 0.3, YB + 0.07, 0.6, 0.3, { fontSize:14, align:'center' });
  T(s, [...sub('t', '1', { bold:true, color:INK }), { text:' — пора поливать;   ' }, ...sub('t', '2', { bold:true, color:INK }), { text:' — датчик «разрешает» полив' }],
    6.25, 6.0, 6.33, 0.3, { fontSize:13, color:INK2 });
  T(s, [{ text:'Отрезок ' }, ...sub('t', '2'), { text:' − ' }, ...sub('t', '1'), { text:' — запаздывание: ', options:{} },
        { text:'полив начинается, когда воды уже меньше, чем задумано', options:{ bold:true } }], 6.25, 6.32, 6.33, 0.55, { fontSize:13.5 });
}

// ================= 3. Датчик измеряет время =================
{
  const s = addSlide(3);
  const IX = MX, IY = 1.85, IW = 5.6, IH = IW * ILL.SENSOR.H / ILL.SENSOR.W;
  s.addImage({ data:ill.sensor, x:IX, y:IY, w:IW, h:IH, altText:'Датчик в почве и увеличенный фрагмент: минеральные зёрна, вода, воздух' });
  const P = ([x, y]) => [IX + x / ILL.SENSOR.W * IW, IY + y / ILL.SENSOR.H * IH];
  const LX = 6.5;
  [['минералы', ILL.SENSOR_ANCHORS.mineral, -0.05], ['вода', ILL.SENSOR_ANCHORS.water, 0], ['воздух', ILL.SENSOR_ANCHORS.air, 0.18]].forEach(([t, a, dy]) => {
    const [ax, ay] = P(a), ly = ay + dy;
    L(s, ax, ay, LX - 0.06, ly, HX.ink, { width:0.9 });
    T(s, t, LX, ly - 0.14, 1.0, 0.28, { fontSize:13, bold:true, valign:'middle' });
  });
  T(s, 'импульс бежит по штангам туда и обратно', IX + 0.15, IY + IH * 0.968 + 0.06, 4.0, 0.26, { fontSize:12, color:INK2 });
  label(s, 'ВО СКОЛЬКО РАЗ ВОЛНА ИДЁТ МЕДЛЕННЕЕ, ЧЕМ В ВОЗДУХЕ (√ε)', MX, 5.4, 6.4);
  const unit = 0.52;
  [['воздух', 1, 'EFE9DF', '1'], ['минералы', 2.17, 'D2B184', '≈ 2,2'], ['вода', 8.94, '46B7C6', '≈ 8,9']].forEach(([t, v, col, lab], k) => {
    const y = 5.78 + k * 0.38;
    T(s, t, MX, y - 0.02, 1.15, 0.3, { fontSize:13, valign:'middle' });
    R(s, 1.95, y + 0.03, v * unit, 0.24, col, k === 0 ? { line:{ color:'CFC6BB', width:0.75 } } : {});
    T(s, lab, 1.95 + v * unit + 0.1, y - 0.02, 0.8, 0.3, { fontSize:13, bold:true, valign:'middle' });
  });

  const RXc = 7.65, RW = RX - RXc;
  label(s, 'ПРИБОР — «ПЕРЕВОДЧИК»', RXc, 1.95, RW);
  [['TbClock', 'время пробега τ'], ['TbBook', 'заводское уравнение — «словарь»'], ['TbDroplet', 'влажность θ']].forEach(([ic, t], k) => {
    const cx = RXc + 0.4 + k * (RW - 0.8) / 2;
    s.addShape(pres.ShapeType.ellipse, { x:cx - 0.36, y:2.35, w:0.72, h:0.72, fill:{ color:'E2F1F3' }, objectName:'icon-disc' });
    icon(s, ic, HX.teal, cx - 0.21, 2.5, 0.42);
    T(s, t, cx - 0.95, 3.13, 1.9, 0.5, { fontSize:12.5, align:'center', color:INK2 });
    if (k < 2) L(s, cx + 0.48, 2.71, cx + (RW - 0.8) / 2 - 0.48, 2.71, AXIS, { width:1.25, endArrowType:'triangle' });
  });
  T(s, 'Два скрытых допущения переводчика', RXc, 3.8, RW, 0.35, { fontSize:18, bold:true });
  [['1', 'Почва для волны — изолятор', 'время пробега меняет только вода; сигнал по пути не теряет энергию'],
   ['2', 'Словарь универсален', 'уравнение, проверенное в основном на супесях, верно для любой почвы']].forEach((a, k) => {
    const y = 4.3 + k * 0.85;
    T(s, a[0], RXc, y - 0.08, 0.45, 0.6, { fontSize:34, bold:true, color:TEAL, fontFace:'Cambria' });
    T(s, [{ text:a[1], options:{ bold:true, fontSize:16, breakLine:true } }, { text:a[2], options:{ fontSize:13.5, color:INK2 } }], RXc + 0.55, y, RW - 0.55, 0.78);
  });
  hline(s, RXc, 6.05, RX);
  icon(s, 'TbAlertTriangle', HX.terra, RXc, 6.17, 0.34);
  T(s, [{ text:'Шкала гарантирована при электропроводности < 0,5 дСм/м, плотности < 1,55 г/см³, глине (< 0,002 мм) < 30 %. ' },
        { text:'В наших тяжёлых почвах это не гарантировано', options:{ bold:true, color:TERRA } }], RXc + 0.48, 6.12, RW - 0.48, 0.78, { fontSize:12.5, color:INK2 });
}

// ================= 4. Подозреваемый — глина =================
{
  const s = addSlide(4);
  const IW = 4.75, IH = IW * 640 / 900;
  s.addImage({ data:ill.clay, x:MX - 0.1, y:1.8, w:IW, h:IH, altText:'Частица глины с отрицательным зарядом, облако катионов и соли в поровом растворе' });
  // легенда к иллюстрации
  const ly = 5.3;
  badge(s, MX + 0.12, ly + 0.14, 0.24, TEAL, '+', 12);
  T(s, 'облако катионов', MX + 0.32, ly, 1.45, 0.28, { fontSize:12, valign:'middle' });
  RR(s, MX + 1.78, ly + 0.03, 0.34, 0.22, BROWN, 0.04);
  T(s, '−', MX + 1.78, ly + 0.03, 0.34, 0.22, { fontSize:12, bold:true, color:WH, align:'center', valign:'middle' });
  T(s, 'частица глины', MX + 2.18, ly, 1.25, 0.28, { fontSize:12, valign:'middle' });
  badge(s, MX + 3.5, ly + 0.14, 0.22, TEAL, '+', 11); badge(s, MX + 3.72, ly + 0.14, 0.22, TERRA, '−', 11);
  T(s, 'соли', MX + 3.9, ly, 0.75, 0.28, { fontSize:12, valign:'middle' });
  icon(s, 'TbCube', HX.brown, MX, 5.82, 0.34);
  T(s, [{ text:'Поверхность смектита 600–800 м²/г: ', options:{ bold:true, color:INK } }, { text:'как сахарная пудра против куска сахара' }],
    MX + 0.45, 5.77, 4.45, 0.5, { fontSize:13, color:INK2 });
  icon(s, 'TbScale', HX.ink2, MX, 6.38, 0.34);
  T(s, [{ text:'Обратная сила: ', options:{ bold:true, color:INK } }, { text:'связанная вода тянет показания вниз — поэтому решает опыт' }],
    MX + 0.45, 6.33, 4.45, 0.5, { fontSize:13, color:INK2 });

  const XL = 6.2;
  label(s, 'СХЕМА · ПРИБОР СРАБАТЫВАЕТ ПО ПОРОГУ, КАК БУДИЛЬНИК ПО ГРОМКОСТИ', XL, 1.95, RX - XL);
  const X0 = 6.45, X1 = 10.95, YB = 4.12, YT = 2.5;
  const sx = lin(0, 7, X0, X1), sy = lin(0, 1, YB, YT);
  L(s, X0, YB, X0, YT - 0.15, AXIS, { width:1, endArrowType:'triangle' });
  L(s, X0, YB, X1 + 0.1, YB, AXIS, { width:1, endArrowType:'triangle' });
  T(s, 'сила сигнала', X0 - 1.02, (YT + YB) / 2 - 0.14, 1.6, 0.28, { fontSize:11.5, color:INK2, rotate:270, align:'center' });
  const thr = 0.35;
  L(s, X0, sy(thr), X1, sy(thr), OCHRE, { width:2, dashType:'dash', name:'threshold' });
  const ts = Array.from({ length:57 }, (_, i) => i * 7 / 56);
  poly(s, ts.map(t => [sx(t), sy(1 - Math.exp(-t / 1.0))]), TEAL, { width:2.75, name:'strong-signal' });
  poly(s, ts.map(t => [sx(t), sy(0.6 * (1 - Math.exp(-t / 1.9)))]), TERRA, { width:2.75, name:'weak-signal' });
  const t1 = -Math.log(1 - thr), t2 = -1.9 * Math.log(1 - thr / 0.6);
  [[t1, TEAL], [t2, TERRA]].forEach(([t, col]) => { L(s, sx(t), sy(thr), sx(t), YB, col, { width:1.25, dashType:'sysDot' }); });
  R(s, sx(t1), YB - 0.045, sx(t2) - sx(t1), 0.09, TERRA, { objectName:'delay' });
  dot(s, sx(t1), sy(thr), 0.17, TEAL); dot(s, sx(t2), sy(thr), 0.17, TERRA);
  T(s, sub('τ', '1', { bold:true }), sx(t1) - 0.3, YB + 0.05, 0.6, 0.3, { fontSize:14, align:'center' });
  T(s, sub('τ', '2', { bold:true }), sx(t2) - 0.3, YB + 0.05, 0.6, 0.3, { fontSize:14, align:'center' });
  const EX = X1 + 0.15;
  T(s, [{ text:'сильный сигнал', options:{ bold:true, breakLine:true } }, { text:'почва-изолятор', options:{ color:INK2, fontSize:11.5 } }], EX, sy(1) - 0.22, RX - EX, 0.46, { fontSize:12.5 });
  T(s, [{ text:'ослабленный', options:{ bold:true, breakLine:true } }, { text:'проводящая почва', options:{ color:INK2, fontSize:11.5 } }], EX, sy(0.585) - 0.22, RX - EX, 0.46, { fontSize:12.5 });
  T(s, 'порог', EX, sy(thr) - 0.02, RX - EX, 0.28, { fontSize:12.5, bold:true });
  T(s, [...sub('τ', '2'), { text:' > ' }, ...sub('τ', '1'), { text:': прибор решает, что дорога длиннее, и «видит» ' }, { text:'лишнюю воду', options:{ bold:true, color:TERRA } }],
    sx(t2) + 0.45, YB + 0.06, RX - sx(t2) - 0.45, 0.5, { fontSize:13 });
  T(s, 'Три проверяемых предсказания', XL, 4.74, RX - XL, 0.32, { fontSize:16, bold:true });
  [['1', 'Тяжелее почва → больше ошибка', 'опровергнет: ошибка одинакова в лёгкой и тяжёлой'],
   ['2', 'Влажнее → ошибка растёт всё быстрее', 'опровергнет: ошибка постоянна или растёт линейно'],
   ['3', 'Плотнее → минералы вместо воздуха', 'опровергнет: плотность не сдвигает показания']].forEach((p, k) => {
    const y = 5.12 + k * 0.58;
    hline(s, XL, y - 0.04, RX);
    T(s, p[0], XL, y, 0.35, 0.48, { fontSize:20, bold:true, color:TEAL, fontFace:'Cambria', valign:'middle' });
    T(s, p[1], XL + 0.42, y, 3.55, 0.48, { fontSize:13.5, bold:true, valign:'middle' });
    T(s, p[2], XL + 4.02, y, RX - XL - 4.02, 0.5, { fontSize:11.5, color:INK2, italic:true, valign:'middle' });
  });
}

// ================= 5. Эксперимент =================
{
  const s = addSlide(5, 'Опыт и результаты');
  label(s, 'СХЕМА ОПЫТА · 18 КОЛОНОК', MX, 1.95, 6);
  const CX0 = 2.05, CWc = 1.55;
  [['A · лёгкая', '10–20 %'], ['B · тяжёлая', '40–50 %'], ['C · тяжёлая', '50–65 %']].forEach((h, k) => {
    T(s, [{ text:h[0], options:{ bold:true, fontSize:15, breakLine:true } }, { text:h[1] + ' физ. глины', options:{ fontSize:12, color:INK2 } }],
      CX0 + k * CWc, 2.3, CWc, 0.55, { align:'center' });
  });
  ['1,35 г/см³', '1,52 г/см³'].forEach((t, r) => T(s, t, MX, 3.05 + r * 0.98, 1.25, 0.6, { fontSize:13.5, color:INK2, valign:'middle' }));
  for (let k = 0; k < 3; k++) for (let r = 0; r < 2; r++) for (let c = 0; c < 3; c++) {
    const x = CX0 + k * CWc + 0.28 + c * 0.38, y = 2.98 + r * 0.98;
    R(s, x, y + 0.07, 0.26, 0.6, SOIL[k], { objectName:'column' });
    s.addShape(pres.ShapeType.ellipse, { x, y:y + 0.6, w:0.26, h:0.13, fill:{ color:SOIL[k] }, objectName:'column-bottom' });
    s.addShape(pres.ShapeType.ellipse, { x, y, w:0.26, h:0.13, fill:{ color:SOIL_T[k] }, objectName:'column-top' });
  }
  T(s, '3 почвы × 2 плотности × 3 колонки · колонка 152 × 300 мм', MX, 5.0, 5.9, 0.28, { fontSize:12.5, color:INK2 });
  T(s, 'физическая глина — частицы мельче 0,01 мм; все почвы незасолённые', MX, 5.28, 5.9, 0.28, { fontSize:12.5, color:INK2 });

  const XR = 7.25;
  T(s, '180', XR, 1.85, 1.9, 1.0, { fontSize:66, bold:true, valign:'middle' });
  T(s, [{ text:'измерений', options:{ bold:true, fontSize:18, breakLine:true } },
        { text:'6 вариантов × 3 колонки × 10 ступеней влажности → 60 средних', options:{ fontSize:13.5, color:INK2 } }], XR + 2.0, 1.98, RX - XR - 2.0, 0.85, { valign:'middle' });
  hline(s, XR, 3.05, RX);
  [['TbScale', 'Ключ ответов — весовой метод', 'сушка при 105 °C; точность ±0,005'],
   ['TbTarget', 'Сдвиг прицела (MB) и типичный промах (RMSE)', 'критерий годности: RMSE ≤ 0,02']].forEach((r, k) => {
    const y = 3.25 + k * 0.9;
    icon(s, r[0], HX.teal, XR, y + 0.04, 0.42);
    T(s, [{ text:r[1], options:{ bold:true, fontSize:16, breakLine:true } }, { text:r[2], options:{ fontSize:13.5, color:INK2 } }], XR + 0.6, y, RX - XR - 0.6, 0.75);
  });
  T(s, 'Каждое решение убирает постороннюю причину: пестроту поля, температуру, перетекание воды', XR, 4.98, RX - XR, 0.45, { fontSize:12.5, color:INK2, italic:true });
  // ход опыта
  L(s, 3.75, 5.82, RX, 5.82, AXIS, { width:1 });
  L(s, 3.75, 5.82, 3.75, 5.9, AXIS, { width:1 }); L(s, RX, 5.82, RX, 5.9, AXIS, { width:1 });
  T(s, 'повторяется на каждой из 10 ступеней влажности', 3.75, 5.5, RX - 3.75, 0.28, { fontSize:12, color:INK2, align:'center' });
  [['Насыщение снизу', '48 ч'], ['Иссушение ступенями', 'при 20 °C'], ['Равновесие', '8–36 ч на ступени'], ['Снимок', 'время пробега + масса колонки']].forEach((st, k) => {
    const x = MX + k * 3.0, last = k === 3;
    R(s, x, 6.0, 2.83, 0.05, last ? TEAL : 'BFDFE3', { objectName:'step-bar' });
    T(s, String(k + 1), x, 6.12, 0.35, 0.5, { fontSize:22, bold:true, color:TEAL, fontFace:'Cambria' });
    T(s, [{ text:st[0], options:{ bold:true, fontSize:14, breakLine:true } }, { text:st[1], options:{ fontSize:12.5, color:INK2 } }], x + 0.4, 6.12, 2.43, 0.62);
  });
}

// ================= 6. Результат =================
{
  const s = addSlide(6);
  const X0 = 1.45, X1 = 8.05, YT = 2.35, YB = 5.82;
  const sx = lin(0, 0.5, X0, X1), sy = lin(0, 0.04, YB, YT);
  T(s, 'ошибка датчика Δθ, м³/м³', MX, 1.93, 4, 0.28, { fontSize:12.5, color:INK2 });
  [0.01, 0.02, 0.03, 0.04].forEach(v => hline(s, X0, sy(v), X1, GRID, 0.75));
  hline(s, X0, sy(0), X1, AXIS, 1);
  [0, 0.01, 0.02, 0.03, 0.04].forEach(v => T(s, v === 0 ? '0' : String(v).replace('.', ','), MX, sy(v) - 0.13, 0.6, 0.26, { fontSize:12, color:INK2, align:'right' }));
  [0, 0.1, 0.2, 0.3, 0.4, 0.5].forEach(v => T(s, v === 0 ? '0' : String(v).replace('.', ','), sx(v) - 0.3, YB + 0.06, 0.6, 0.26, { fontSize:12, color:INK2, align:'center' }));
  T(s, 'влажность по эталону θ, м³/м³ · плотность 1,35 г/см³', X0, YB + 0.36, X1 - X0, 0.28, { fontSize:12.5, color:INK2, align:'center' });
  L(s, X0, sy(0.025), X1, sy(0.025), SLATE, { width:1.25, dashType:'dash', name:'passport-limit' });
  T(s, 'паспортный предел 0,025', X0 + 0.08, sy(0.025) - 0.3, 2.6, 0.26, { fontSize:12, color:INK2 });
  // продолжение C (вне анализа)
  const cEnd = [sx(0.40), sy(0.027)];
  poly(s, [cEnd, [sx(0.44), sy(0.031)], [sx(0.48), sy(0.038)]], SOIL[2], { width:1.5, dashType:'sysDot', name:'C-outside' });
  ['A', 'B', 'C'].forEach((g, k) => {
    poly(s, TV.th.map((t, i) => [sx(t), sy(TV[g][i])]), SOIL[k], { width:2.5, name:'series-' + g });
    TV.th.forEach((t, i) => dot(s, sx(t), sy(TV[g][i]), 0.12, SOIL[k]));
  });
  [[0.44, 0.031], [0.48, 0.038]].forEach(([t, v]) => dot(s, sx(t), sy(v), 0.13, 'FFFFFF', { ring:SOIL[2], ringW:1.75 }));
  T(s, 'состояния вне анализа', sx(0.48) - 2.05, sy(0.038) - 0.13, 1.88, 0.26, { fontSize:11.5, color:INK2, align:'right' });
  [['A', '10–20 %', 0.011], ['B', '40–50 %', 0.019]].forEach(([g, p, v]) =>
    T(s, [{ text:g + ' ', options:{ bold:true } }, { text:p, options:{ color:INK2 } }], sx(0.40) + 0.13, sy(v) - 0.13, 1.2, 0.26, { fontSize:12.5 }));
  T(s, [{ text:'C ', options:{ bold:true } }, { text:'50–65 %', options:{ color:INK2 } }], sx(0.245), sy(0.0215) - 0.13, sx(0.345) - sx(0.245), 0.26, { fontSize:12.5, align:'right' });
  // легенда
  const LGY = 6.58;
  [['A · 10–20 % физ. глины', SOIL[0]], ['B · 40–50 %', SOIL[1]], ['C · 50–65 %', SOIL[2]]].forEach(([t, col], k) => {
    const x = X0 + [0, 2.25, 3.65][k];
    L(s, x, LGY + 0.13, x + 0.35, LGY + 0.13, col, { width:2.5 }); dot(s, x + 0.175, LGY + 0.13, 0.1, col);
    T(s, t, x + 0.43, LGY, 1.9, 0.26, { fontSize:11.5, color:INK2 });
  });
  dot(s, X0 + 5.1, LGY + 0.13, 0.11, 'FFFFFF', { ring:SOIL[2], ringW:1.5 });
  T(s, 'вне анализа', X0 + 5.22, LGY, 1.3, 0.26, { fontSize:11.5, color:INK2 });

  const XR = 8.75, RW = RX - XR;
  [['≤ 0,014', TEAL, 'типичный промах (RMSE) везде ниже критерия 0,02', 1.92],
   ['+6…10 мм', TERRA, 'средний сдвиг вверх в тяжёлых почвах на метровый слой (MB +0,006…+0,010, описательно); в лёгкой — на уровне погрешности', 3.27],
   ['× 2,6', INK, 'во столько раз быстрее растёт ошибка в самой тяжёлой почве, чем в лёгкой (θ 0,20 → 0,40)', 4.82]].forEach(([big, col, cap, y], k) => {
    if (k) hline(s, XR, y - 0.12, RX);
    T(s, big, XR, y, RW, 0.62, { fontSize:34, bold:true, color:col, valign:'middle' });
    T(s, cap, XR, y + 0.62, RW, 0.7, { fontSize:13, color:INK2 });
  });
  RR(s, XR, 6.02, RW, 0.84, PANEL, 0.06, { objectName:'analogy' });
  icon(s, 'TbGauge', HX.teal, XR + 0.15, 6.22, 0.44);
  T(s, [{ text:'Как спидометр: ', options:{ bold:true } }, { text:'чем быстрее, тем сильнее завышает → постоянная поправка не сработает' }],
    XR + 0.72, 6.02, RW - 0.85, 0.84, { fontSize:12.5, valign:'middle' });
}

// ================= 7. «Не значимо» ≠ «невиновен» =================
{
  const s = addSlide(7, 'Доказательства');
  const LW = 4.85;
  icon(s, 'TbMusic', HX.terra, MX, 1.95, 0.38);
  T(s, 'Если смешать всё', MX + 0.5, 1.95, LW - 0.5, 0.38, { fontSize:18, bold:true, valign:'middle' });
  T(s, 'доли разброса ошибки · 60 средних по колонкам', MX, 2.5, LW, 0.26, { fontSize:12, color:INK2 });
  const parts = [[9.1, TERRA, 'почва'], [1.0, BROWN, 'плотность'], [0.1, OCHRE, 'взаимодействие'], [89.8, 'DDD8D2', 'внутри вариантов']];
  let bx = MX; const BY = 2.85, BH = 0.5, GAP = 0.025, tot = LW - GAP * 3;
  parts.forEach(([v, col], k) => { const w = Math.max(v / 100 * tot, 0.012); R(s, bx, BY, w, BH, col, { objectName:'share-' + k }); bx += w + GAP; });
  T(s, [{ text:'внутри вариантов 89,8 %', options:{ bold:true, color:INK, breakLine:true } }, { text:'в основном рост ошибки с влажностью', options:{ fontSize:11 } }],
    MX + 0.75, BY + 0.02, LW - 0.85, BH - 0.04, { fontSize:12, color:INK2, valign:'middle' });
  [['почва 9,1 %', TERRA], ['плотность 1,0 %', BROWN], ['взаимодействие 0,1 %', OCHRE]].forEach(([t, col], k) => {
    const x = MX + [0, 1.35, 2.95][k];
    R(s, x, 3.53, 0.16, 0.16, col);
    T(s, t, x + 0.22, 3.47, 1.7, 0.28, { fontSize:12, color:k === 0 ? INK : INK2, bold:k === 0 });
  });
  T(s, 'p = 0,074 → «не значимо»', MX, 3.95, LW, 0.45, { fontSize:22, bold:true });
  T(s, 'Но «нет доказательств» ≠ «невиновен»', MX, 4.43, LW, 0.35, { fontSize:17, bold:true, color:TERRA });
  T(s, 'Эффект почвы заглушён ростом ошибки с влажностью — как разговор на концерте заглушает музыка', MX, 4.85, LW, 0.55, { fontSize:13.5, color:INK2 });
  icon(s, 'TbAlertCircle', HX.ochre, MX, 5.95, 0.34);
  T(s, [{ text:'Граница вывода: ', options:{ bold:true, color:INK } }, { text:'уровни влажности — состояния одних и тех же колонок, почва в группе одна → вывод описательный' }],
    MX + 0.47, 5.9, LW - 0.47, 0.9, { fontSize:13, color:INK2 });

  vline(s, 6.05, 1.95, 6.85);
  const XL = 6.4;
  icon(s, 'TbVolumeOff', HX.teal, XL, 1.95, 0.38);
  T(s, 'Выключим музыку: та же влажность', XL + 0.5, 1.95, RX - XL - 0.5, 0.38, { fontSize:18, bold:true, valign:'middle' });
  T(s, 'ошибка Δθ, м³/м³', XL, 2.47, 2.0, 0.26, { fontSize:12, color:INK2 });
  T(s, 'ни одна линия не идёт вниз: A < B ≤ C', 8.4, 2.47, RX - 8.4, 0.26, { fontSize:12.5, bold:true, align:'right' });
  const GX = [7.6, 9.35, 11.1], sy = lin(0, 0.03, 5.55, 2.85);
  [0.01, 0.02, 0.03].forEach(v => hline(s, 7.3, sy(v), 11.4, GRID, 0.75));
  hline(s, 7.3, sy(0), 11.4, AXIS, 1);
  [0, 0.01, 0.02, 0.03].forEach(v => T(s, v === 0 ? '0' : String(v).replace('.', ','), XL, sy(v) - 0.13, 0.75, 0.26, { fontSize:12, color:INK2, align:'right' }));
  TV.th.forEach((t, i) => {
    const pts = ['A', 'B', 'C'].map((g, k) => [GX[k], sy(TV[g][i])]);
    poly(s, pts, WATER[i], { width:2, name:'level-' + TV.lv[i] });
    pts.forEach(p => dot(s, p[0], p[1], 0.11, WATER[i]));
  });
  [['θ = 0,40', 0.027], ['0,38', 0.022], ['0,30', 0.012], ['0,20', 0.006], ['0,05 и 0,10', 0.0015]].forEach(([t, v]) =>
    T(s, t, 11.25, sy(v) - 0.13, RX - 11.25, 0.26, { fontSize:12, color:INK2 }));
  ['A', 'B', 'C'].forEach((g, k) => {
    R(s, GX[k] - 0.2, 5.72, 0.14, 0.14, SOIL[k]);
    T(s, g, GX[k] - 0.02, 5.62, 0.4, 0.32, { fontSize:15, bold:true });
  });
  [['по уровням (блоки)', 'p = 0,013'], ['только порядок (Фридман)', 'p = 0,004'], ['если уровни независимы', '(1/3)⁶ ≈ 1 из 729']].forEach((r, k) => {
    const x = XL + k * 2.07;
    if (k) vline(s, x - 0.1, 6.12, 6.82);
    T(s, [{ text:r[0], options:{ fontSize:11.5, color:INK2, breakLine:true } }, { text:r[1], options:{ fontSize:k === 2 ? 15 : 18, bold:true } }], x, 6.1, 1.95, 0.75);
  });
}

// ================= 8. Второй подозреваемый и несбывшийся прогноз =================
{
  const s = addSlide(8);
  const LW = 5.3;
  icon(s, 'TbDeviceCctv', HX.ink2, MX, 1.95, 0.38);
  T(s, 'Глина и соли всегда в кадре вместе', MX + 0.5, 1.95, LW - 0.5, 0.38, { fontSize:17, bold:true, valign:'middle' });
  T(s, 'ECe, дСм/м — след солей', MX, 2.47, 2.4, 0.26, { fontSize:12, color:INK2 });
  dot(s, 3.6, 2.6, 0.13, SOIL[1]); T(s, '1,35 г/см³', 3.72, 2.47, 1.0, 0.26, { fontSize:11.5, color:INK2 });
  dot(s, 4.85, 2.6, 0.13, 'FFFFFF', { ring:SOIL[1], ringW:1.75 }); T(s, '1,52 г/см³', 4.97, 2.47, 1.1, 0.26, { fontSize:11.5, color:INK2 });
  const X0 = 1.4, X1 = 5.95, sx = lin(0, 12, X0, X1), sy = lin(0.2, 0.7, 5.2, 2.95);
  [0.3, 0.4, 0.5, 0.6, 0.7].forEach(v => hline(s, X0, sy(v), X1, GRID, 0.75));
  hline(s, X0, sy(0.2), X1, AXIS, 1);
  [0.2, 0.3, 0.4, 0.5, 0.6, 0.7].forEach(v => T(s, String(v).replace('.', ','), MX, sy(v) - 0.13, 0.55, 0.26, { fontSize:12, color:INK2, align:'right' }));
  [0, 4, 8, 12].forEach(v => T(s, String(v), sx(v) - 0.3, 5.26, 0.6, 0.26, { fontSize:12, color:INK2, align:'center' }));
  T(s, 'гигроскопичность MH, % — след глины', X0, 5.52, X1 - X0, 0.26, { fontSize:12, color:INK2, align:'center' });
  const PTS = [[2.1, 0.35, 0, 1], [2.3, 0.28, 0, 0], [6.5, 0.52, 1, 1], [6.8, 0.45, 1, 0], [11.4, 0.61, 2, 1], [11.1, 0.49, 2, 0]];
  PTS.forEach(([m, e, k, filled]) => dot(s, sx(m), sy(e), 0.17, filled ? SOIL[k] : 'FFFFFF', filled ? {} : { ring:SOIL[k], ringW:2 }));
  [['A', 2.6, 0.33], ['B', 7.05, 0.5], ['C', 10.2, 0.57]].forEach(([g, m, e]) => T(s, g, sx(m), sy(e) - 0.14, 0.35, 0.28, { fontSize:14, bold:true }));
  T(s, 'r = 0,88', X0 + 0.1, sy(0.7) + 0.05, 1.3, 0.3, { fontSize:14, bold:true });
  T(s, [{ text:'сдвиг ~ MH: ', options:{ color:INK2 } }, { text:'R² = 0,88', options:{ bold:true } }, { text:'      сдвиг ~ ECe: ', options:{ color:INK2 } }, { text:'R² = 0,96', options:{ bold:true } }],
    MX, 5.92, LW, 0.3, { fontSize:14 });
  T(s, 'Всего 6 точек (3 почвы × 2 плотности): кто виноват — данные не скажут', MX, 6.3, LW, 0.55, { fontSize:13, color:INK2 });

  vline(s, 6.35, 1.95, 6.85);
  const XL = 6.7, RW = RX - XL;
  icon(s, 'TbLinkOff', HX.terra, XL, 1.95, 0.38);
  T(s, 'Прогноз для плотности не сбылся', XL + 0.5, 1.95, RW - 0.5, 0.38, { fontSize:17, bold:true, valign:'middle' });
  T(s, 'сдвиг при уплотнении 1,35 → 1,52, мм в метровом слое', XL, 2.47, RW, 0.26, { fontSize:12, color:INK2 });
  R(s, XL, 2.84, 0.16, 0.16, 'BCC2C3'); T(s, 'прогноз теории смеси', XL + 0.22, 2.78, 2.2, 0.28, { fontSize:12, color:INK2 });
  R(s, XL + 2.45, 2.84, 0.16, 0.16, TERRA); T(s, 'наблюдалось', XL + 2.67, 2.78, 1.6, 0.28, { fontSize:12, color:INK2 });
  const by = lin(-4, 12, 4.82, 3.2);
  [-4, 4, 8, 12].forEach(v => hline(s, XL + 0.55, by(v), RX, GRID, 0.75));
  hline(s, XL + 0.55, by(0), RX, AXIS, 1);
  [-4, 0, 4, 8, 12].forEach(v => T(s, v > 0 ? '+' + v : v < 0 ? '−' + (-v) : '0', XL, by(v) - 0.13, 0.45, 0.26, { fontSize:12, color:INK2, align:'right' }));
  [['A', -2], ['B', -1], ['C', -2]].forEach(([g, obs], k) => {
    const cx = XL + 1.35 + k * 1.75;
    R(s, cx - 0.5, by(9.5), 0.46, by(0) - by(9.5), 'BCC2C3', { objectName:'predicted' });
    R(s, cx + 0.04, by(0), 0.46, by(obs) - by(0), TERRA, { objectName:'observed' });
    T(s, '+9,5', cx - 0.62, by(9.5) - 0.3, 0.7, 0.26, { fontSize:12.5, bold:true, align:'center' });
    T(s, '−' + (-obs), cx - 0.08, by(obs) + 0.02, 0.7, 0.26, { fontSize:12.5, bold:true, align:'center' });
    T(s, g, cx - 0.3, 4.92, 0.6, 0.28, { fontSize:14, bold:true, align:'center' });
  });
  const links = [['модель смеси', '?', OCHRE, INK], ['та же влажность', '×', TERRA, WH], ['та же ECe', '×', TERRA, WH]];
  links.forEach((l, k) => {
    const x = XL + k * 2.0, w = 1.86;
    s.addShape(pres.ShapeType.roundRect, { x, y:5.33, w, h:0.44, fill:{ color:'FFFFFF' }, line:{ color:l[2], width:1.5 }, rectRadius:0.22, objectName:'link-' + k });
    T(s, l[0], x + 0.15, 5.33, w - 0.5, 0.44, { fontSize:12.5, bold:true, valign:'middle' });
    badge(s, x + w - 0.24, 5.55, 0.3, l[2], l[1], 13, l[3]);
    if (k < 2) L(s, x + w, 5.55, x + 2.0, 5.55, AXIS, { width:2 });
  });
  T(s, [{ text:'Все звенья целы ⇒ прогноз сбывается. Не сбылся ⇒ порвалось хотя бы одно. ', options:{ bold:true, breakLine:true } },
        { text:'Кандидаты: у плотных вариантов ECe ниже на 0,07–0,12 дСм/м, диапазоны влажности разные — но полностью разрыв они не объясняют', options:{ color:INK2 } }],
    XL, 5.92, RW, 0.95, { fontSize:13 });
}

// ================= 9. Насколько твёрды улики =================
{
  const s = addSlide(9);
  const cols = [MX, 4.75, 7.95, 9.45], ws = [3.85, 3.05, 1.4, RX - 9.45];
  ['Утверждение', 'Тип связи', 'Твёрдость', 'Слабое место'].forEach((h, k) => label(s, h.toUpperCase(), cols[k], 1.98, ws[k]));
  hline(s, MX, 2.33, RX, INK, 1);
  const lev = { 3:[TEAL, 'высокая'], 2:[OCHRE, 'средняя'], 1:[TERRA, 'низкая'] };
  const ROWS = [
    ['Типичный промах (RMSE) < 0,02 везде', 'TbRuler2', 'прямое измерение', 'видим прямо', 3, 'лаборатория; средние трёх колонок, без разброса между датчиками'],
    ['Ошибка растёт с влажностью', 'TbEye', 'прямое наблюдение', 'видим прямо', 3, 'до θ 0,40 (и два состояния вне анализа); только иссушение'],
    ['На равной влажности A < B ≤ C', 'TbChartDots', 'статистический вывод', 'обобщаем', 2, 'уровни — состояния одних колонок; одна почва на группу'],
    ['Прогноз смеси для плотности не сбылся', 'TbMathFunction', 'количественная модель', 'предсказываем', 2, 'ECe и диапазоны влажности вариантов различались'],
    ['Виновата именно поверхность глин', 'TbFingerprint', 'косвенный показатель', 'судим по следам', 1, 'не отделена от солей (r = 0,88); всего три почвы']];
  ROWS.forEach((r, i) => {
    const y = 2.4 + i * 0.78, h = 0.7;
    T(s, r[0], cols[0], y, ws[0] - 0.15, h, { fontSize:15, bold:true, valign:'middle' });
    icon(s, r[1], HX.ink2, cols[1], y + 0.17, 0.36);
    T(s, [{ text:r[2], options:{ breakLine:true } }, { text:r[3], options:{ italic:true, color:INK2, fontSize:12.5 } }], cols[1] + 0.5, y, ws[1] - 0.5, h, { fontSize:14, valign:'middle' });
    const [col, word] = lev[r[4]];
    for (let q = 0; q < 3; q++) RR(s, cols[2] + q * 0.38, y + 0.2, 0.32, 0.12, q < r[4] ? col : 'E6E2DD', 0.03, { objectName:'strength' });
    T(s, word, cols[2], y + 0.38, 1.3, 0.25, { fontSize:11.5, color:INK2 });
    T(s, r[5], cols[3], y, ws[3], h, { fontSize:13, color:INK2, valign:'middle' });
    hline(s, MX, y + 0.74, RX);
  });
  icon(s, 'TbArrowRight', HX.teal, MX, 6.42, 0.36);
  T(s, [{ text:'Слабые места — это адреса следующих опытов ', options:{ bold:true } }, { text:'(слайд 11)', options:{ color:INK2 } }], MX + 0.5, 6.4, CW - 0.5, 0.4, { fontSize:16, valign:'middle' });
}

// ================= 10. Что делать сейчас =================
{
  const s = addSlide(10, 'Что делать');
  const LW = 6.1;
  [[GREEN, 'ЗЕЛЁНЫЙ', 'Лёгкие незасолённые (10–20 % физ. глины)', 'как наша почва A: заводская калибровка годится без поправки'],
   [OCHRE, 'ЖЁЛТЫЙ', 'Тяжёлые почвы, 40–65 % физ. глины', 'завышение 6–10 мм воды в метровом слое, до 27 мм во влажной: калибровка под почву или 1–2 весовые проверки за сезон'],
   [TERRA, 'КРАСНЫЙ', 'Постоянная поправка', 'не годится: ошибается и в сухой, и во влажной почве (график справа)'],
   [SLATE, 'СЕРЫЙ', 'Серая зона', '20–40 % физ. глины не испытаны; в поле нужна температурная поправка']].forEach((r, k) => {
    const y = 1.98 + k * 1.2;
    s.addShape(pres.ShapeType.ellipse, { x:MX, y:y + 0.05, w:0.3, h:0.3, fill:{ color:r[0] }, objectName:'light' });
    T(s, r[1], MX + 0.45, y + 0.06, 2, 0.26, { fontSize:11, bold:true, color:INK2, charSpacing:1 });
    T(s, [{ text:r[2], options:{ bold:true, fontSize:16, breakLine:true } }, { text:r[3], options:{ fontSize:13, color:INK2 } }], MX + 0.45, y + 0.31, LW - 0.45, 0.8);
    if (k < 3) hline(s, MX, y + 1.12, MX + LW);
  });
  const XL = 7.35, RW = RX - XL;
  T(s, 'Почему постоянная поправка не годится', XL, 1.95, RW, 0.35, { fontSize:16, bold:true });
  T(s, 'почва C, 1,35 г/см³ · Δθ, м³/м³', XL, 2.32, RW, 0.26, { fontSize:12, color:INK2 });
  const X0 = 7.95, X1 = 11.15, sx = lin(0, 0.4, X0, X1), sy = lin(-0.01, 0.03, 5.15, 2.8);
  [-0.01, 0.01, 0.02, 0.03].forEach(v => hline(s, X0, sy(v), X1, GRID, 0.75));
  hline(s, X0, sy(0), X1, AXIS, 1);
  [-0.01, 0, 0.01, 0.02, 0.03].forEach(v => T(s, v === 0 ? '0' : (v < 0 ? '−' : '') + String(Math.abs(v)).replace('.', ','), XL, sy(v) - 0.13, 0.5, 0.26, { fontSize:12, color:INK2, align:'right' }));
  [0, 0.1, 0.2, 0.3, 0.4].forEach(v => T(s, v === 0 ? '0' : String(v).replace('.', ','), sx(v) - 0.3, 5.22, 0.6, 0.26, { fontSize:12, color:INK2, align:'center' }));
  T(s, 'влажность θ, м³/м³', X0, 5.48, X1 - X0, 0.26, { fontSize:12, color:INK2, align:'center' });
  const res = TV.C.map(v => +(v - 0.010).toFixed(3));
  poly(s, TV.th.map((t, i) => [sx(t), sy(TV.C[i])]), TERRA, { width:2.5, name:'error-C' });
  poly(s, TV.th.map((t, i) => [sx(t), sy(res[i])]), TEAL, { width:2.5, name:'residual-C' });
  TV.th.forEach((t, i) => { dot(s, sx(t), sy(TV.C[i]), 0.11, TERRA); dot(s, sx(t), sy(res[i]), 0.11, TEAL); });
  T(s, [{ text:'ошибка', options:{ bold:true, breakLine:true } }, { text:'0…+0,027', options:{ color:INK2 } }], X1 + 0.15, sy(0.027) - 0.24, RX - X1 - 0.15, 0.48, { fontSize:12 });
  T(s, [{ text:'после поправки', options:{ bold:true, breakLine:true } }, { text:'−0,009…+0,017', options:{ color:INK2 } }], X1 + 0.15, sy(0.017) - 0.12, RX - X1 - 0.15, 0.48, { fontSize:12 });
  RR(s, XL, 5.95, RW, 0.9, PANEL, 0.06, { objectName:'proposal' });
  icon(s, 'TbBulb', HX.teal, XL + 0.15, 6.17, 0.44);
  T(s, [{ text:'Наше предложение (в поле не проверено): ', options:{ bold:true } }, { text:'пробы после полива и перед поливом — одна точка даёт сдвиг, две — ещё и наклон' }],
    XL + 0.72, 5.95, RW - 0.85, 0.9, { fontSize:12.5, valign:'middle' });
}

// ================= 11. Как закрыть дело =================
{
  const s = addSlide(11);
  const cols = [MX, 4.6, 8.85], ws = [3.7, 4.1, RX - 8.85];
  ['Если верно…', '…то увидим', 'Решающий опыт'].forEach((h, k) => label(s, h.toUpperCase(), cols[k], 1.98, ws[k]));
  hline(s, MX, 2.33, RX, INK, 1);
  const ROWS = [
    ['Виноваты соли (проводимость раствора)', 'ошибка растёт вместе с ECe', 'засолённые варианты, ECe 1–4 дСм/м (отдельная статья)'],
    ['Виновата поверхность глин', 'при равной ECe ошибка больше там, где больше смектита', 'минералогия глин; пары почв с равной ECe'],
    ['Плотность «спрятана» разницей ECe', 'при выровненной ECe сдвиг ближе к прогнозу +0,009…+0,010', 'уплотнение одной почвы при одинаковой ECe'],
    ['Эффект почвы реален, а не артефакт колонок', 'сохранится в смешанной модели, учитывающей разброс колонок', 'переанализ 180 измерений'],
    ['Идея: завышение задаёт проводимость', 'поправка по двум входам (время пробега + объёмная ЭП) уберёт большую часть ошибки', 'мерить объёмную ЭП рядом с датчиком (в опыте не мерили)']];
  ROWS.forEach((r, i) => {
    const y = 2.4 + i * 0.72, h = 0.66, idea = i === 4;
    T(s, r[0], cols[0], y, ws[0] - 0.15, h, { fontSize:15, bold:true, color:idea ? TERRA : INK, valign:'middle' });
    T(s, r[1], cols[1], y, ws[1] - 0.2, h, { fontSize:13.5, color:INK2, valign:'middle' });
    if (i === 3) {
      T(s, r[2], cols[2], y + 0.03, ws[2], 0.3, { fontSize:13.5, bold:true });
      RR(s, cols[2], y + 0.35, 1.72, 0.27, TEAL, 0.135, { objectName:'data-ready' });
      icon(s, 'TbDatabase', HX.white, cols[2] + 0.1, y + 0.385, 0.2);
      T(s, 'данные уже есть', cols[2] + 0.36, y + 0.35, 1.32, 0.27, { fontSize:11.5, bold:true, color:WH, valign:'middle' });
    } else T(s, r[2], cols[2], y, ws[2], h, { fontSize:13.5, valign:'middle' });
    hline(s, MX, y + 0.69, RX);
  });
  icon(s, 'TbMessageQuestion', HX.terra, MX, 6.15, 0.46);
  T(s, [{ text:'Ваш ход: какой опыт различит гипотезы дешевле всего?', options:{ bold:true, fontSize:17, breakLine:true } },
        { text:'Так же стоит проверять любой датчик, который судит о воде по проницаемости', options:{ fontSize:13.5, color:INK2 } }], MX + 0.62, 6.08, CW - 0.62, 0.8);
}

// ================= 12. Три мысли =================
{
  const s = addSlide(12);
  const CWk = 3.75, GAPk = (CW - 3 * CWk) / 2;
  [['1', 'Годится, но не бесплатно', 'По критерию 0,02 — да. Но при 40–65 % физ. глины средний сдвиг вверх, и сильнее во влажной почве'],
   ['2', 'Доказывать устойчивость, а не только p', 'Порядок почв сохранялся на каждом уровне влажности, два разных теста согласны; слабые места названы'],
   ['3', 'Гипотеза подтверждена частично', 'Связь с почвой и влажностью есть; механизм (глина или соли) не разделён; плотность — нет']].forEach((t, k) => {
    const x = MX + k * (CWk + GAPk);
    if (k) vline(s, x - GAPk / 2, 2.0, 4.6, DARKLINE, 0.75);
    T(s, t[0], x, 1.85, 0.8, 0.8, { fontSize:44, bold:true, color:OCHRE, fontFace:'Cambria' });
    T(s, t[1], x, 2.72, CWk, 0.7, { fontSize:19, bold:true, color:WH });
    T(s, t[2], x, 3.45, CWk, 1.15, { fontSize:14.5, color:DARKTEXT });
  });
  [['H1 · почва', 'частично', OCHRE, INK], ['H2 · влажность', 'подтверждена', GREEN, WH], ['H3 · плотность', 'не подтверждена', TERRA, WH]].forEach((h, k) => {
    const x = MX + k * (CWk + GAPk);
    RR(s, x, 4.9, CWk, 0.5, h[2], 0.25, { objectName:'verdict-' + k });
    T(s, [{ text:h[0] + ': ', options:{ bold:true } }, { text:h[1] }], x + 0.25, 4.9, CWk - 0.5, 0.5, { fontSize:15.5, color:h[3], valign:'middle' });
  });
  T(s, 'Дело не закрыто — и теперь понятно, какой опыт его закроет', MX, 5.72, CW, 0.42, { fontSize:20, color:OCHRE, italic:true, fontFace:'Cambria' });
  T(s, 'Спасибо за внимание!', MX, 6.22, CW, 0.55, { fontSize:28, bold:true, color:WH });
}

// ================= запасные слайды =================
let bi = 0;
function addBackup(title, note) {
  if (bi++ === 0) { pres.addSection({ title:'Запасные слайды' }); curSection = 'Запасные слайды'; }
  const s = pres.addSlide({ masterName:'BACKUP', sectionTitle:curSection });
  s.addText(title, { placeholder:'title' });
  kicker(s, 'ЗАПАСНОЙ СЛАЙД · ДЛЯ ВОПРОСОВ');
  s.addNotes('[Запасной слайд — показывать только при вопросах; во время доклада не используется]\n\n' + note);
  T(s, 'Запасные слайды — для вопросов', MX, 7.02, 6, 0.3, { fontSize:10, color:'A39B94', valign:'middle' });
  return s;
}
// таблица в минималистичном стиле: заголовок без заливки, тонкие линии между строками
function minTable(s, head, rows, colW, o = {}) {
  const bNone = { type:'none' }, bHair = { type:'solid', pt:0.75, color:HAIR }, bInk = { type:'solid', pt:1, color:HX.ink };
  const hd = head.map((t, k) => ({ text:t, options:{ bold:true, fontSize:12, color:INK2, valign:'bottom', align:k === 0 && o.leftFirst ? 'left' : (o.align || 'center'), border:[bNone, bNone, bInk, bNone] } }));
  const body = rows.map(r => r.map((t, k) => ({ text:t, options:Object.assign({ fontSize:o.fontSize || 14, color:INK, valign:'middle', align:k === 0 && o.leftFirst ? 'left' : (o.align || 'center'),
    border:[bNone, bNone, bHair, bNone] }, k === 0 ? { bold:true } : {}) })));
  s.addTable([hd].concat(body), { x:MX, y:o.y || 1.95, w:colW.reduce((a, b) => a + b, 0), colW, rowH:o.rowH, margin:[3, 6, 3, 6], objectName:o.name || 'table' });
}
// B1 карта теорий
{
  const s = addBackup('Шесть теорий и типы связей', 'Показывать, если спрашивают, на чём стоит работа и как связаны теории. Связи четырёх типов: механизм, количественная модель, косвенный показатель, статистический вывод; у каждого своё слабое место.');
  const NW = 2.0, NH = 1.25, GX = (CW - 5 * NW) / 4, cx = k => MX + k * (NW + GX), TOP = 1.95, BOT = 3.55, MID = (TOP + BOT) / 2;
  const node = (k, y, head, sub2, mode) => {
    const fill = mode === 'dark' ? HX.ink : mode === 'teal' ? HX.teal : 'FFFFFF';
    s.addShape(pres.ShapeType.roundRect, { x:cx(k), y, w:NW, h:NH, fill:{ color:fill }, line:mode ? undefined : { color:'CFC8C0', width:1 }, rectRadius:0.08, objectName:'node' });
    T(s, [{ text:head, options:{ bold:true, fontSize:14.5, color:mode ? WH : INK, breakLine:true } }, { text:sub2, options:{ fontSize:12, color:mode ? 'E8E3DD' : INK2 } }],
      cx(k) + 0.13, y + 0.05, NW - 0.26, NH - 0.1, { valign:'middle' });
  };
  node(0, TOP, 'Физика почв', 'доли фаз, плотность, набухание');
  node(0, BOT, 'Химия глин', 'двойной слой, соли; прокси: MH, ECe');
  node(1, TOP, 'Теория смеси (CRIM)', '√ε смеси = Σ доля × √ε');
  node(1, BOT, 'Электродинамика', 'скорость ∝ 1/√ε; потери сигнала');
  node(2, MID, 'Датчик + уравнение', 'τ → θ: эмпирический перевод', 'dark');
  node(3, TOP, 'Метрология', 'эталон U = 0,005; MB, RMSE');
  node(3, BOT, 'Статистика', 'дисперсионный анализ, блоки, регрессия');
  node(4, MID, 'Мелиорация', 'решение: порог полива', 'teal');
  const A = { endArrowType:'triangle', width:2 }, yT = TOP + NH / 2, yB = BOT + NH / 2, yM = MID + NH / 2;
  L(s, cx(0) + NW, yT, cx(1), yT, OCHRE, A);                       // физика → смесь (модель)
  L(s, cx(0) + NW, yB, cx(1), yB, TEAL, A);                        // химия → электродинамика (механизм)
  L(s, cx(1) + NW / 2, BOT, cx(1) + NW / 2, TOP + NH, OCHRE, A);   // электродинамика → смесь
  L(s, cx(1) + NW, yT, cx(2), yM - 0.25, OCHRE, A);                // смесь → датчик
  L(s, cx(1) + NW, yB, cx(2), yM + 0.25, TEAL, A);                 // электродинамика → датчик
  L(s, cx(2) + NW, yM - 0.25, cx(3), yT, SLATE, A);                // датчик → метрология
  L(s, cx(3) + NW / 2, TOP + NH, cx(3) + NW / 2, BOT, SLATE, A);   // метрология → статистика
  L(s, cx(3) + NW, yB, cx(4), yM + 0.25, SLATE, A);                // статистика → мелиорация
  // косвенный показатель: о химии глин судим по MH и ECe в регрессиях
  poly(s, [[cx(0) + NW / 2, BOT + NH], [cx(0) + NW / 2, BOT + NH + 0.2], [cx(3) + NW / 2, BOT + NH + 0.2], [cx(3) + NW / 2, BOT + NH + 0.02]], TERRA,
    { width:2, dashType:'dash', endArrowType:'triangle', name:'proxy-link' });
  T(s, 'прокси MH, ECe → регрессии', cx(1) + 0.2, BOT + NH + 0.24, 3.2, 0.26, { fontSize:11.5, color:INK2 });
  [[TEAL, 'Механизм', 'причина → следствие. Подводит, если есть обратная сила (связанная вода)', undefined],
   [OCHRE, 'Модель', 'даёт число, его можно проверить. Подводит, если нарушены допущения (плотность)', undefined],
   [TERRA, 'Косвенный показатель', 'судим о скрытом по видимому (MH, ECe). Подводит, если прокси ходят парой', 'dash'],
   [SLATE, 'Статистический вывод', 'от колонок к общему. Подводит, если наблюдения не независимы', undefined]].forEach((g, k) => {
    const x = MX + k * 3.0;
    hline(s, x, 5.5, x + 2.8, HAIR);
    L(s, x, 5.78, x + 0.55, 5.78, g[0], { width:3, dashType:g[3] });
    T(s, g[1], x + 0.68, 5.63, 2.1, 0.3, { fontSize:14, bold:true, valign:'middle' });
    T(s, g[2], x, 6.0, 2.8, 0.9, { fontSize:12, color:INK2 });
  });
}
// B2 допущения нашего опыта
{
  const s = addBackup('Допущения нашего опыта', 'Показывать, если спрашивают о надёжности эталона, статистики или колонок. Эта таблица объясняет, почему p-значения названы описательными, а выводы относятся к изученным почвам.');
  const ROWS = [
    ['«Эталон — это истина»', 'Два способа (вся колонка и микрокерны): расхождение ≤ 0,008, в среднем 0,003; U = 0,005', 'частично', OCHRE, INK],
    ['«Десять уровней — независимые наблюдения»', 'Нет: это последовательные состояния одних и тех же колонок', 'p описательные', TERRA, WH],
    ['«Одна почва представляет группу»', 'В каждой группе одна почва', 'экстраполяция', TERRA, WH],
    ['«Колонка не мешает датчику»', 'До стенки 59–73 мм (рекомендуется ≥ 100), штанги доходят до дна', 'не проверено', TERRA, WH],
    ['«Объём почвы постоянен»', 'Нарушено при набухании (почва C, 1,52 г/см³) — эталон по реальному объёму', 'учтено частично', OCHRE, INK],
    ['«Округление опубликованных чисел не меняет выводов»', '20 000 симуляций: доля эффекта почвы 0,07–0,11 держится; граница A–B неустойчива', 'проверено', GREEN, WH]];
  ['ДОПУЩЕНИЕ', 'ЧТО МЫ СДЕЛАЛИ ИЛИ ЗНАЕМ', 'СТАТУС'].forEach((h, k) => label(s, h, [MX, 4.75, 10.45][k], 1.98, [3.9, 5.5, 2.1][k]));
  hline(s, MX, 2.33, RX, INK, 1);
  ROWS.forEach((r, i) => {
    const y = 2.4 + i * 0.72;
    T(s, r[0], MX, y, 3.85, 0.66, { fontSize:14, bold:true, valign:'middle' });
    T(s, r[1], 4.75, y, 5.5, 0.66, { fontSize:13.5, color:INK2, valign:'middle' });
    RR(s, 10.45, y + 0.17, 2.13, 0.34, r[3], 0.17, { objectName:'status' });
    T(s, r[2], 10.45, y + 0.17, 2.13, 0.34, { fontSize:12, bold:true, color:r[4], align:'center', valign:'middle' });
    hline(s, MX, y + 0.69, RX);
  });
}
// B3 расчёт прогноза теории смеси
{
  const s = addBackup('Расчёт прогноза модели смеси', 'Показывать, если просят объяснить, откуда +0,0095. Модель: время пробега складывается из времён через фазы; уплотнение при той же влажности заменяет воздух минералами; заводская шкала читает удлинение пробега как воду вместо воздуха.');
  label(s, 'ДОЛИ ОБЪЁМА ПРИ θ = 0,25 (ρs = 2,63 г/см³)', MX, 1.98, 5.2);
  const by = lin(0, 1, 6.25, 2.55);
  [[1.35, [0.513, 0.25, 0.237]], [1.52, [0.578, 0.25, 0.172]]].forEach(([rho, v], k) => {
    const x = 1.45 + k * 2.1; let acc = 0;
    [['твёрдая фаза', 'B9AE9F'], ['вода', '8FCCD4'], ['воздух', 'F1EEEA']].forEach(([nm, col], q) => {
      const y1 = by(acc + v[q]), y0 = by(acc);
      R(s, x, y1, 1.3, y0 - y1 - 0.02, col, q === 2 ? { line:{ color:'D9D3CB', width:0.75 } } : {});
      T(s, String(v[q].toFixed(3)).replace('.', ','), x, y1, 1.3, y0 - y1, { fontSize:14, bold:true, align:'center', valign:'middle' });
      if (k === 1) T(s, nm, x + 1.4, (y0 + y1) / 2 - 0.14, 1.4, 0.28, { fontSize:12.5, color:INK2, valign:'middle' });
      acc += v[q];
    });
    T(s, String(rho).replace('.', ',') + ' г/см³', x - 0.2, 6.32, 1.7, 0.3, { fontSize:13.5, bold:true, align:'center' });
  });
  ['Модель: √ε смеси = φтв·√4,7 + θ·√80 + φвозд·√1',
   'Уплотнение 1,35 → 1,52: твёрдого больше на 0,17 / 2,63 ≈ 0,065 — вместо воздуха',
   'Пробег растёт: 0,065 × (√4,7 − 1) = 0,065 × 1,17 ≈ 0,076',
   'Датчик читает это как воду вместо воздуха: 0,076 / (√80 − 1) = 0,076 / 7,94 ≈ 0,0095'].forEach((t, k) => {
    const y = 2.0 + k * 0.85;
    T(s, String(k + 1), 6.75, y, 0.4, 0.6, { fontSize:24, bold:true, color:TEAL, fontFace:'Cambria' });
    T(s, t, 7.25, y + 0.02, RX - 7.25, 0.75, { fontSize:14.5, bold:k === 3 });
    hline(s, 6.75, y + 0.78, RX);
  });
  T(s, [{ text:'Прогноз: ≈ +0,0095 ', options:{ bold:true, color:INK } }, { text:'(при ρs 2,60–2,67: +0,0094…+0,0096)', options:{ color:INK2 } }], 6.75, 5.5, RX - 6.75, 0.4, { fontSize:17 });
  T(s, [{ text:'Скрыто: ', options:{ bold:true, color:TERRA } }, { text:'та же θ, та же проводимость, нет потерь в пути' }], 6.75, 6.0, RX - 6.75, 0.4, { fontSize:15 });
}
// B4 набухание
{
  const s = addBackup('Набухание и пересчёт эталона', 'Показывать, если спрашивают о набухании. Почва C при 1,52 г/см³ набухла: подъём ≈ 25 мм, объёмная деформация 8,3 %, θmax 0,460 > исходной пористости 0,415. Эталон считали по реальному объёму; дополнительной ошибки не нашли, но сравнение смешано с плотностью и ECe, а почва поднималась выше штанг.');
  const draw = (x, title, swell) => {
    T(s, title, x, 1.98, 3.6, 0.35, { fontSize:16, bold:true, align:'center' });
    R(s, x + 0.6, 2.7, 2.4, 3.2, SOIL[2], { objectName:'soil' });
    if (swell) R(s, x + 0.6, 2.42, 2.4, 0.28, 'A57A68', { objectName:'swollen-top' });
    s.addShape(pres.ShapeType.rect, { x:x + 0.6, y:2.7, w:2.4, h:3.2, line:{ color:HX.slate, width:1.5 }, objectName:'column' });
    [1.4, 2.13].forEach(dx => R(s, x + dx, 2.7, 0.07, 3.2, 'D7DBDC', { objectName:'rod' }));
    T(s, swell ? 'почва выше штанг на ≈ 25 мм' : 'почва = длина штанг', x, 6.02, 3.6, 0.3, { fontSize:13.5, align:'center', color:swell ? TERRA : INK2, bold:!!swell });
  };
  draw(MX, 'Сухая: 300 мм', false); draw(4.6, 'Влажная: ≈ 325 мм', true);
  [['≈ 25 мм', 'подъём поверхности'], ['8,3 %', 'объёмная деформация'], ['0,460 > 0,415', 'θmax больше исходной пористости']].forEach((r, k) => {
    const y = 2.0 + k * 1.15;
    if (k) hline(s, 8.95, y - 0.12, RX);
    T(s, r[0], 8.95, y, RX - 8.95, 0.55, { fontSize:28, bold:true, color:TERRA, valign:'middle' });
    T(s, r[1], 8.95, y + 0.55, RX - 8.95, 0.35, { fontSize:13.5, color:INK2 });
  });
  T(s, 'Делить воду на прежний объём — значит завысить эталон и занизить видимую ошибку датчика. Эталон считали по реальному объёму: дополнительной ошибки не нашли.',
    8.95, 5.45, RX - 8.95, 1.35, { fontSize:13.5 });
}
// B5 ловушка терминов
{
  const s = addBackup('Ловушка терминов «глина»', 'Показывать, если спрашивают о сравнении с паспортом и зарубежными работами. Физическая глина по Качинскому — частицы < 0,01 мм; глина по USDA и в паспорте — < 0,002 мм. В группе A условие паспорта (< 30 % глины) выполнено заведомо; для B и C проверить нельзя.');
  label(s, 'ВЕРХНЯЯ ГРАНИЦА РАЗМЕРА ЧАСТИЦ, ММ (В МАСШТАБЕ)', MX, 1.98, 8);
  [['глина по Качинскому < 0,001', 0.001, 'B9B2AA'], ['«clay» по USDA и в паспорте < 0,002', 0.002, HX.teal], ['физическая глина по Качинскому < 0,01', 0.01, HX.terra]].forEach((b, k) => {
    const y = 2.45 + k * 0.82, w = b[1] * 1000;
    R(s, MX, y, w, 0.55, b[2], { objectName:'size-bar' });
    if (w > 5) T(s, b[0], MX + 0.2, y, 8, 0.55, { fontSize:16, bold:true, valign:'middle', color:WH });
    else T(s, b[0], MX + w + 0.2, y, 8, 0.55, { fontSize:16, bold:true, valign:'middle' });
  });
  hline(s, MX, 5.0, RX);
  T(s, [{ text:'Ловушка перевода', options:{ bold:true, color:TERRA, breakLine:true } }, { text:'30 % «clay» из паспорта и 30 % физической глины — разные почвы. «Лёгкая глина» по Качинскому ≠ «clay» по USDA.' }],
    MX, 5.2, 5.6, 1.6, { fontSize:16 });
  T(s, [{ text:'Что из этого следует', options:{ bold:true, color:TEAL, breakLine:true } }, { text:'В группе A физической глины ≤ 20 %, значит, и USDA-глины ≤ 20 %: условие паспорта выполнено. Для B и C проверить нельзя.' }],
    6.95, 5.2, RX - 6.95, 1.6, { fontSize:16 });
}
// B6–B8 таблицы
{
  const s = addBackup('Статистика ошибки по вариантам', K3.backup[0].note);
  minTable(s, ['Вариант', 'ECe, дСм/м', 'MB', 'RMSE', 'Δθ min / max', '95 % ДИ для MB', 'p (MB = 0)', 'J_nl'], [
    ['A1-n', '0,35', '+0,004', '0,007', '−0,002 / +0,013', '−0,0003…+0,0083', '0,066', '−0,263'],
    ['A2-n', '0,28', '+0,002', '0,006', '−0,004 / +0,011', '−0,0023…+0,0063', '0,32', '−0,268'],
    ['B1-n', '0,52', '+0,007', '0,011', '−0,003 / +0,019', '+0,0006…+0,0134', '0,035', '−0,118'],
    ['B2-n', '0,45', '+0,006', '0,009', '−0,004 / +0,016', '+0,0009…+0,0111', '0,025', '0,000'],
    ['C1-n', '0,61', '+0,010', '0,014', '−0,001 / +0,027', '+0,0026…+0,0174', '0,014', '−0,306'],
    ['C2-n', '0,49', '+0,008', '0,012', '−0,002 / +0,022', '+0,0013…+0,0147', '0,025', '−0,225']],
    [1.25, 1.35, 1.15, 1.15, 2.0, 2.45, 1.3, 1.183], { rowH:[0.45, 0.52, 0.52, 0.52, 0.52, 0.52, 0.52], name:'table-III' });
  T(s, 'м³/м³; MB и RMSE — по средним трёх колонок, n = 10 уровней влажности. ДИ и p описательные (уровни — состояния одних колонок). J_nl < 0 — правосторонняя асимметрия: крупные положительные ошибки во влажной части. Код варианта: группа, плотность (1 — 1,35; 2 — 1,52 г/см³), n — незасолённая.',
    MX, 5.75, CW, 1.05, { fontSize:13, color:INK2 });
}
{
  const s = addBackup('Дисперсионный анализ ошибки', K3.backup[1].note);
  minTable(s, ['Источник изменчивости', 'SS, (м³/м³)²', 'df', 'MS, (м³/м³)²', 'F', 'p', 'η²'], [
    ['Почва (T)', '3,63·10⁻⁴', '2', '1,82·10⁻⁴', '2,74', '0,074', '0,091'],
    ['Плотность (ρb)', '4,17·10⁻⁵', '1', '4,17·10⁻⁵', '0,63', '0,43', '0,010'],
    ['Взаимодействие T × ρb', '3,3·10⁻⁶', '2', '1,7·10⁻⁶', '0,03', '0,98', '0,001'],
    ['Внутри вариантов', '3,58·10⁻³', '54', '6,63·10⁻⁵', '—', '—', '0,898'],
    ['Всего', '3,99·10⁻³', '59', '—', '—', '—', '1,000']],
    [3.6, 1.9, 0.9, 1.9, 1.2, 1.15, 1.183], { rowH:[0.45, 0.55, 0.55, 0.55, 0.55, 0.55], leftFirst:true, name:'table-IV' });
  T(s, 'N = 60 значений (6 вариантов × 10 уровней, средние трёх колонок); SS восстановлены по MB и RMSE вариантов. Внутривариантный член включает рост ошибки с влажностью. 20 000 симуляций округления: p для почвы 0,04–0,12, η² 0,07–0,11. Эффект плотности −0,0017, 95 % ДИ −0,006…+0,003.',
    MX, 5.3, CW, 1.3, { fontSize:13.5, color:INK2 });
}
{
  const s = addBackup('Свойства почв в вариантах', K3.backup[2].note);
  minTable(s, ['Вариант', 'Группа (физ. глина, %)', 'ρs, г/см³', 'MH, %', 'ρb, г/см³', 'σ1:5, дСм/м', 'ECe, дСм/м'], [
    ['A1-n', 'A (10–20)', '2,67 ± 0,01', '2,1 ± 0,2', '1,35 ± 0,02', '0,044 ± 0,003', '0,35 ± 0,02'],
    ['A2-n', 'A (10–20)', '2,67 ± 0,01', '2,3 ± 0,2', '1,52 ± 0,02', '0,036 ± 0,002', '0,28 ± 0,02'],
    ['B1-n', 'B (40–50)', '2,63 ± 0,01', '6,5 ± 0,3', '1,35 ± 0,03', '0,066 ± 0,004', '0,52 ± 0,03'],
    ['B2-n', 'B (40–50)', '2,63 ± 0,01', '6,8 ± 0,4', '1,52 ± 0,03', '0,058 ± 0,004', '0,45 ± 0,03'],
    ['C1-n', 'C (50–65)', '2,60 ± 0,02', '11,4 ± 0,6', '1,35 ± 0,03', '0,078 ± 0,005', '0,61 ± 0,04'],
    ['C2-n', 'C (50–65)', '2,60 ± 0,02', '11,1 ± 0,6', '1,52 ± 0,03', '0,062 ± 0,004', '0,49 ± 0,03']],
    [1.3, 2.4, 1.6, 1.5, 1.7, 1.75, 1.583], { rowH:[0.45, 0.52, 0.52, 0.52, 0.52, 0.52, 0.52], name:'table-II' });
  T(s, 'n = 3, среднее ± SD. Все ECe ниже 1 дСм/м и в 6,6–14 раз ниже порога засоления 4 дСм/м. В плотных вариантах ECe ниже на 0,07–0,12 дСм/м. Почва C2-n набухала при θ > 0,40.',
    MX, 5.75, CW, 0.9, { fontSize:13.5, color:INK2 });
}
// B9 ограничения
{
  const s = addBackup('Ограничения исследования', K3.backup[3].note);
  const half = Math.ceil(K3.limits.length / 2);
  [K3.limits.slice(0, half), K3.limits.slice(half)].forEach((col, j) => {
    const x = MX + j * 6.05;
    col.forEach((t, i) => {
      const y = 2.0 + i * 0.92;
      T(s, String(j * half + i + 1), x, y, 0.45, 0.5, { fontSize:20, bold:true, color:TEAL, fontFace:'Cambria' });
      T(s, t, x + 0.5, y + 0.03, 5.3, 0.82, { fontSize:14 });
    });
  });
}
}

// скругляем концы и стыки у ломаных (custGeom): так линии графиков выглядят аккуратнее
async function patchXml(file) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    let xml = await zip.file(name).async('string');
    xml = xml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, sp => {
      if (!sp.includes('<a:custGeom>')) return sp;
      return sp.replace(/<a:ln( w="\d+")?>([\s\S]*?)<\/a:ln>/, (m, w, inner) => {
        const cut = inner.search(/<a:headEnd|<a:tailEnd/);
        const body = cut < 0 ? inner + '<a:round/>' : inner.slice(0, cut) + '<a:round/>' + inner.slice(cut);
        return `<a:ln${w || ''} cap="rnd">${body}</a:ln>`;
      });
    });
    zip.file(name, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' }));
}

(async () => {
  await prepIcons(ICON_LIST);
  const ill = await ILL.illustrations();
  build(ill);
  const out = path.resolve(process.argv[2] || 'deck6.pptx');
  await pres.writeFile({ fileName:out });
  await applyTheme(out, THEME);
  await patchXml(out);
  const total = secs.reduce((a, b) => a + b, 0);
  console.log('written', out, '; speech', mmss(total), '; with transitions', mmss(total + 4 * K.slides.length));
})().catch(e => { console.error(e); process.exit(1); });
