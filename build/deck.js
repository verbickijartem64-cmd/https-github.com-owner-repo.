const pptxgen = require('pptxgenjs');
const path = require('path');
const { applyTheme } = require(process.env.PPTX_SKILL + '/scripts/apply_theme.js');
const { slides: TALK } = require('./content.js');

const THEME = {
  name: 'CS616 Chernozem',
  headFontFace: 'Cambria',
  bodyFontFace: 'Calibri',
  colors: { dk1:'2A1E16', lt1:'FFFFFF', dk2:'5B4636', lt2:'F1F3F2',
    accent1:'1C7C8C', accent2:'D9A441', accent3:'B8552F', accent4:'8C6A4F', accent5:'6E7F80', accent6:'4E8A5B',
    hlink:'1C7C8C', folHlink:'8C6A4F' }
};
const H = THEME.colors;                 // hex values, only for charts
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';            // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = 'Датчики CS616: систематическая погрешность заводской калибровки';
pres.author = 'Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е.';
const C = pres.SchemeColor;
const DK=C.text1, MUTED=C.text2, WH=C.background1, LT=C.background2;
const TEAL=C.accent1, OCHRE=C.accent2, TERRA=C.accent3, BROWN=C.accent4, SLATE=C.accent5, GREEN=C.accent6;
const FOOT = 'CS616 · погрешность заводской калибровки · Кубанский ГАУ';

pres.defineSlideMaster({ title:'LIGHT', background:{ color: WH }, margin:[0.5,0.6,0.6,0.6],
  objects:[
    { placeholder:{ options:{ name:'title', type:'title', x:0.6, y:0.35, w:12.1, h:0.95, fontSize:30, bold:true, color:DK, align:'left', valign:'middle', margin:0 }, text:'' } },
    { text:{ text:FOOT, options:{ x:0.6, y:7.05, w:9, h:0.3, fontSize:10, color:MUTED, margin:0, isTextBox:true } } } ],
  slideNumber:{ x:12.0, y:7.05, w:0.73, h:0.3, fontSize:10, color:MUTED, align:'right' } });
pres.defineSlideMaster({ title:'DARK', background:{ color: DK }, margin:[0.5,0.6,0.6,0.6],
  objects:[
    { placeholder:{ options:{ name:'title', type:'title', x:0.6, y:0.35, w:12.1, h:0.95, fontSize:30, bold:true, color:WH, align:'left', valign:'middle', margin:0 }, text:'' } },
    { text:{ text:FOOT, options:{ x:0.6, y:7.05, w:9, h:0.3, fontSize:10, color:OCHRE, margin:0, isTextBox:true } } } ],
  slideNumber:{ x:12.0, y:7.05, w:0.73, h:0.3, fontSize:10, color:OCHRE, align:'right' } });
pres.defineSlideMaster({ title:'COVER', background:{ color: DK }, margin:[0.5,0.6,0.6,0.6],
  objects:[
    { placeholder:{ options:{ name:'title', type:'title', x:0.7, y:1.3, w:7.2, h:2.3, fontSize:44, bold:true, color:WH, align:'left', valign:'top', margin:0 }, text:'' } } ] });

// ---------- helpers ----------
const fmt = n => String(n).replace('.', ',');
function T(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({ x, y, w, h, fontSize:16, color:DK, margin:0, valign:'top', isTextBox:true }, o));
}
function card(s, x, y, w, h, fill = LT, name = 'card', o = {}) {
  s.addShape(pres.ShapeType.roundRect, Object.assign({ x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, rectRadius:0.08, objectName:name }, o));
}
function circ(s, x, y, d, fill, label, fs = 18, color = WH) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w:d, h:d, fill:{ color: fill }, line:{ type:'none' }, objectName:'badge' });
  s.addText(label, { x, y, w:d, h:d, align:'center', valign:'middle', fontSize:fs, bold:true, color, margin:0, isTextBox:true });
}
function arrowDown(s, x, y, w = 0.4, h = 0.35, fill = SLATE) {
  s.addShape(pres.ShapeType.downArrow, { x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, objectName:'arrow' });
}
function arrowRight(s, x, y, w = 0.4, h = 0.4, fill = SLATE) {
  s.addShape(pres.ShapeType.rightArrow, { x, y, w, h, fill:{ color: fill }, line:{ type:'none' }, objectName:'arrow' });
}
function addSlide(i, master = 'LIGHT', section) {
  const sl = pres.addSlide({ masterName: master, sectionTitle: section });
  const d = TALK[i - 1];
  if (master !== 'COVER' || true) sl.addText(d.title, { placeholder:'title' });
  const words = d.text.join(' ').split(/\s+/).length;
  const mins = Math.max(1, Math.round(words / 140 * 2) / 2);
  sl.addNotes(`[≈ ${fmt(mins)} мин${d.core ? '' : ' · можно пропустить при нехватке времени'}]\n\n` + d.text.map(t => t.replace(/^▸ ?/, '')).join('\n\n'));
  return sl;
}
const chartBase = { catAxisLabelColor:H.dk2, valAxisLabelColor:H.dk2, catAxisLabelFontSize:12, valAxisLabelFontSize:12,
  catAxisLabelFontFace:'+mn-lt', valAxisLabelFontFace:'+mn-lt', legendFontFace:'+mn-lt', legendFontSize:12, legendColor:H.dk2,
  valGridLine:{ color:'D9DEDC', size:0.5 }, catGridLine:{ style:'none' }, dataLabelFontFace:'+mn-lt' };

pres.addSection({ title:'Зачем и как устроен датчик' });
// ---------- 1. cover ----------
{
  const s = addSlide(1, 'COVER', 'Зачем и как устроен датчик');
  T(s, 'Систематическая погрешность заводской калибровки датчиков влажности CS616 в незасолённых почвах разного гранулометрического состава', 0.7, 3.3, 7.3, 1.35, { fontSize:18, color:OCHRE, italic:true });
  T(s, 'Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е.', 0.7, 5.05, 7.3, 0.35, { fontSize:14, color:WH, bold:true });
  T(s, 'Кубанский государственный аграрный университет имени И. Т. Трубилина · Краснодар, 2026', 0.7, 5.5, 7.3, 0.6, { fontSize:14, color:WH });
  // probe over three soils
  const sx = 9.0;
  [['A', TEAL, 'чернозём обыкновенный'], ['B', BROWN, 'чернозём типичный'], ['C', TERRA, 'лугово-чернозёмная']].forEach(([l, col, nm], k) => {
    card(s, sx, 2.25 + k * 1.35, 3.6, 1.2, col, 'soil-' + l);
    T(s, l, sx + 0.25, 2.25 + k * 1.35, 0.6, 1.2, { fontSize:36, bold:true, color:WH, valign:'middle' });
    T(s, nm, sx + 0.95, 2.25 + k * 1.35, 1.35, 1.2, { fontSize:12, color:WH, valign:'middle' });
  });
  s.addShape(pres.ShapeType.roundRect, { x:sx + 2.4, y:1.2, w:0.9, h:0.65, fill:{ color: WH }, line:{ type:'none' }, rectRadius:0.06, objectName:'probe-head' });
  T(s, 'CS616', sx + 2.4, 1.2, 0.9, 0.65, { fontSize:12, bold:true, align:'center', valign:'middle' });
  [2.55, 3.15].forEach(dx => s.addShape(pres.ShapeType.rect, { x:sx + dx - 0.04 + 0.0, y:1.85, w:0.08, h:4.3, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
}

// ---------- 2. cost of 1% ----------
{
  const s = addSlide(2, 'LIGHT');
  const rows = [['0,01', '1 л воды на 100 л почвы', 'один процент объёмной влажности'], ['10 мм', 'воды в метровом слое', 'слой воды, которого «не хватает» или «лишний»'], ['100 м³', 'воды на гектар', 'то же самое, в единицах поливной нормы']];
  rows.forEach((r, k) => {
    const y = 1.6 + k * 1.75;
    card(s, 0.6, y, 6.6, 1.55, LT, 'stat-' + k);
    T(s, r[0], 0.85, y, 2.6, 1.55, { fontSize:44, bold:true, color:TEAL, valign:'middle' });
    T(s, [{ text:r[1], options:{ bold:true, fontSize:18, breakLine:true } }, { text:r[2], options:{ fontSize:14, color:MUTED } }], 3.55, y + 0.1, 3.5, 1.35, { valign:'middle' });
  });
  card(s, 7.6, 1.6, 5.1, 5.0, DK, 'passport-card');
  T(s, 'Паспорт датчика', 7.95, 1.9, 4.4, 0.4, { fontSize:16, color:OCHRE, bold:true });
  T(s, '±0,025', 7.95, 2.35, 4.4, 1.0, { fontSize:60, bold:true, color:WH });
  T(s, 'до 25 мм воды в метровом слое, которые прибор может «придумать» или «не заметить»', 7.95, 3.45, 4.4, 1.2, { fontSize:18, color:WH });
  T(s, [{ text:'Но паспорт получен в основном на супесях и более лёгких почвах, а на виноградниках края физической глины бывает ', options:{ color:WH } }, { text:'до 65 %', options:{ bold:true, color:OCHRE } }], 7.95, 4.8, 4.4, 1.6, { fontSize:18 });
}

// ---------- 3. sensor measures time ----------
{
  const s = addSlide(3, 'LIGHT');
  // soil block + rods
  card(s, 0.6, 1.6, 5.6, 2.3, BROWN, 'soil-block');
  T(s, 'почва', 2.8, 3.45, 2, 0.35, { fontSize:14, color:WH });
  s.addShape(pres.ShapeType.roundRect, { x:0.95, y:1.6, w:1.3, h:0.5, fill:{ color: DK }, line:{ type:'none' }, rectRadius:0.05, objectName:'probe-head' });
  T(s, 'датчик', 0.95, 1.6, 1.3, 0.5, { fontSize:12, color:WH, align:'center', valign:'middle' });
  [1.2, 2.0].forEach(x => s.addShape(pres.ShapeType.rect, { x, y:2.1, w:0.07, h:1.6, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
  s.addShape(pres.ShapeType.line, { x:2.6, y:2.2, w:3.2, h:0, line:{ color: WH, width:2.5, dashType:'dash', endArrowType:'triangle' }, objectName:'pulse-out' });
  s.addShape(pres.ShapeType.line, { x:5.8, y:3.0, w:-3.2, h:0, line:{ color: OCHRE, width:2.5, dashType:'dash', endArrowType:'triangle' }, objectName:'pulse-back' });
  T(s, 'импульс идёт по штангам и возвращается', 2.6, 2.4, 3.4, 0.5, { fontSize:14, color:WH });
  // permittivity chart
  s.addChart(pres.charts.BAR, [{ name:'Диэлектрическая проницаемость', labels:['Воздух', 'Минералы почвы', 'Вода'], values:[1, 4.7, 80] }], Object.assign({}, chartBase, {
    x:0.6, y:4.1, w:5.6, h:2.8, barDir:'bar', chartColors:[H.accent1], showValue:true, dataLabelFormatCode:'0.0', dataLabelColor:H.dk1, dataLabelFontSize:14, dataLabelPosition:'outEnd',
    showTitle:true, title:'Диэлектрическая проницаемость: чем больше, тем сильнее «тормозит» волну', titleFontSize:14, titleColor:H.dk2, titleFontFace:'+mn-lt', showLegend:false, valAxisHidden:true, valGridLine:{ style:'none' }, barGapWidthPct:40, catAxisOrientation:'maxMin' }));
  // chain
  const chain = ['Больше воды в почве', 'Волна идёт медленнее', 'Период τ длиннее'];
  chain.forEach((c, k) => {
    const y = 1.6 + k * 1.1;
    card(s, 6.8, y, 5.9, 0.8, k === 2 ? TEAL : LT, 'chain-' + k);
    T(s, c, 7.1, y, 5.3, 0.8, { fontSize:18, bold:true, color:k === 2 ? WH : DK, valign:'middle' });
    if (k < 2) arrowDown(s, 9.55, y + 0.78, 0.4, 0.3);
  });
  card(s, 6.8, 5.05, 5.9, 1.85, DK, 'formula');
  T(s, 'Заводской «переводчик» время → влажность', 7.1, 5.2, 5.3, 0.35, { fontSize:14, color:OCHRE, bold:true });
  T(s, 'θ = −0,0663 − 0,0063·τ + 0,0007·τ²', 7.1, 5.65, 5.3, 0.55, { fontSize:22, bold:true, color:WH });
  T(s, 'τ — период в микросекундах; θ — объёмная влажность', 7.1, 6.3, 5.3, 0.4, { fontSize:14, color:WH });
}

// ---------- 4. two hidden assumptions ----------
{
  const s = addSlide(4, 'LIGHT');
  const cards = [['1', 'Почва для волны — изолятор', 'Сигнал замедляется только из-за воды. Всё остальное — «прозрачно».'], ['2', 'Переводчик универсален', 'Уравнение, настроенное на одни почвы, верно и для других.']];
  cards.forEach((c, k) => {
    const x = 0.6 + k * 6.2;
    card(s, x, 1.6, 5.9, 2.3, LT, 'assumption-' + k);
    circ(s, x + 0.3, 1.9, 0.75, k ? TERRA : TEAL, c[0], 24);
    T(s, c[1], x + 1.3, 1.85, 4.4, 0.85, { fontSize:22, bold:true, valign:'middle' });
    T(s, c[2], x + 0.3, 2.85, 5.3, 0.9, { fontSize:16 });
  });
  T(s, 'Границы стандартной шкалы по паспорту', 0.6, 4.2, 8, 0.4, { fontSize:16, bold:true, color:MUTED });
  [['< 0,5 дСм/м', 'объёмная электропроводность почвы'], ['< 1,55 г/см³', 'плотность сложения'], ['< 30 %', 'глины (частицы мельче 0,002 мм)']].forEach((c, k) => {
    const x = 0.6 + k * 4.1;
    card(s, x, 4.7, 3.9, 1.35, WH, 'limit-' + k, { line:{ color: SLATE, width:1.25 } });
    T(s, c[0], x + 0.25, 4.8, 3.4, 0.7, { fontSize:28, bold:true, color:TEAL, valign:'middle' });
    T(s, c[1], x + 0.25, 5.45, 3.4, 0.5, { fontSize:14, color:MUTED });
  });
  T(s, [{ text:'Наш вопрос: не «сломается ли датчик за границей», а ', options:{} }, { text:'в какую сторону, насколько и при какой влажности он начнёт ошибаться', options:{ bold:true, color:TERRA } }], 0.6, 6.25, 12.1, 0.65, { fontSize:18 });
}

// ---------- 5. clay mechanism ----------
{
  const s = addSlide(5, 'LIGHT');
  // surface-area picture
  card(s, 0.6, 1.6, 5.7, 3.6, LT, 'surface-card');
  T(s, 'Одна и та же масса', 0.85, 1.7, 5.2, 0.4, { fontSize:14, color:MUTED, bold:true });
  s.addShape(pres.ShapeType.rect, { x:1.0, y:2.5, w:1.5, h:1.5, fill:{ color: SLATE }, line:{ type:'none' }, objectName:'sand-grain' });
  T(s, 'кубик сахара / песок', 0.85, 4.1, 2.0, 0.6, { fontSize:14, align:'center' });
  for (let i = 0; i < 6; i++) for (let j = 0; j < 6; j++)
    s.addShape(pres.ShapeType.rect, { x:3.55 + i * 0.25, y:2.5 + j * 0.25, w:0.2, h:0.2, fill:{ color: TERRA }, line:{ type:'none' }, objectName:'clay-particle' });
  T(s, 'пудра / глина', 3.35, 4.1, 2.0, 0.6, { fontSize:14, align:'center' });
  T(s, 'смектит: 600–800 м²/г — на порядки больше поверхности, чем у песка', 0.85, 4.65, 5.2, 0.5, { fontSize:14, bold:true, color:TERRA });
  // chain
  const steps = ['Ионы в растворе и на поверхности глин проводят ток', 'Импульс теряет энергию и слабеет', 'Порог срабатывания достигается позже («тихий будильник»)', 'Период τ длиннее → прибор «видит» больше воды'];
  steps.forEach((t, k) => {
    const y = 1.6 + k * 1.0;
    card(s, 6.9, y, 5.8, 0.75, k === 3 ? TEAL : LT, 'step-' + k);
    T(s, t, 7.15, y, 5.3, 0.75, { fontSize:16, bold:k === 3, color:k === 3 ? WH : DK, valign:'middle' });
    if (k < 3) arrowDown(s, 9.6, y + 0.74, 0.36, 0.26);
  });
  card(s, 0.6, 5.45, 12.1, 1.35, WH, 'counter-card', { line:{ color: OCHRE, width:1.5 } });
  T(s, [{ text:'Но есть и противодействие. ', options:{ bold:true, color:TERRA } }, { text:'Связанная с поверхностью вода замедляет волну слабее свободной и тянет показания вниз. Какая сила победит, решает эксперимент, а не рассуждение. Наше предположение: завышение, растущее с влажностью.', options:{} }], 0.9, 5.55, 11.5, 1.15, { fontSize:16, valign:'middle' });
}

// ---------- 6. terms trap ----------
{
  const s = addSlide(6, 'LIGHT');
  T(s, 'Размер частиц, мм (в масштабе)', 0.6, 1.55, 6, 0.35, { fontSize:14, bold:true, color:MUTED });
  const scale = 900 / 0.01 / 100; // inches per mm: 9 in for 0.01 mm
  [['глина по Качинскому < 0,001', 0.001, SLATE], ['«clay» по USDA и в паспорте < 0,002', 0.002, TEAL], ['физическая глина по Качинскому < 0,01', 0.01, TERRA]].forEach((b, k) => {
    const y = 2.0 + k * 0.95;
    s.addShape(pres.ShapeType.rect, { x:0.6, y, w:b[1] * scale, h:0.6, fill:{ color: b[2] }, line:{ type:'none' }, objectName:'size-bar' });
    if (b[1] > 0.005) T(s, b[0], 0.85, y, 8, 0.6, { fontSize:16, bold:true, valign:'middle', color:WH });
    else T(s, b[0], 0.6 + b[1] * scale + 0.2, y, 7, 0.6, { fontSize:16, bold:true, valign:'middle' });
  });
  card(s, 0.6, 5.0, 5.9, 1.8, LT, 'trap-card');
  T(s, [{ text:'Ловушка перевода', options:{ bold:true, color:TERRA, breakLine:true } }, { text:'30 % «clay» из паспорта и 30 % физической глины — разные почвы. «Лёгкая глина» по Качинскому ≠ «clay» по USDA.', options:{} }], 0.85, 5.1, 5.4, 1.6, { fontSize:16 });
  card(s, 6.8, 5.0, 5.9, 1.8, LT, 'consequence-card');
  T(s, [{ text:'Что из этого следует', options:{ bold:true, color:TEAL, breakLine:true } }, { text:'В группе A физической глины ≤ 20 %, значит и USDA-глины ≤ 20 %: условие паспорта выполнено. Для B и C проверить его нельзя.', options:{} }], 7.05, 5.1, 5.4, 1.6, { fontSize:16 });
}

pres.addSection({ title:'Вопрос и эксперимент' });
// ---------- 7. hypotheses ----------
{
  const s = addSlide(7, 'LIGHT', 'Вопрос и эксперимент');
  const hs = [['H1', 'Ошибка растёт с содержанием физической глины', 'Опровергнет: ошибка одинакова в лёгких и тяжёлых почвах', TEAL], ['H2', 'Ошибка нарастает с влажностью, нелинейно', 'Опровергнет: ошибка постоянна или растёт строго линейно', TERRA], ['H3', 'Плотность влияет, а сила влияния зависит от почвы', 'Опровергнет: плотность не меняет ошибку или меняет её одинаково во всех почвах', BROWN]];
  hs.forEach((h, k) => {
    const y = 1.55 + k * 1.6;
    card(s, 0.6, y, 12.1, 1.4, LT, 'h-card-' + k);
    circ(s, 0.9, y + 0.33, 0.75, h[3], h[0], 20);
    T(s, h[1], 2.0, y + 0.12, 10.4, 0.7, { fontSize:22, bold:true, valign:'middle' });
    T(s, h[2], 2.0, y + 0.82, 10.4, 0.5, { fontSize:16, color:MUTED, italic:true });
  });
  T(s, 'Каждое предсказание запрещает часть возможных результатов — поэтому его можно проверить', 0.6, 6.45, 12.1, 0.4, { fontSize:16, bold:true, color:TEAL });
}

// ---------- 8. experiment ----------
{
  const s = addSlide(8, 'LIGHT');
  // design grid
  T(s, 'Схема опыта: 3 почвы × 2 плотности', 0.6, 1.5, 6, 0.35, { fontSize:14, bold:true, color:MUTED });
  const cols = [['A', '10–20 % физ. глины', TEAL], ['B', '40–50 %', BROWN], ['C', '50–65 %', TERRA]];
  cols.forEach((c, k) => {
    T(s, [{ text:c[0], options:{ bold:true, fontSize:20, breakLine:true } }, { text:c[1], options:{ fontSize:12 } }], 1.9 + k * 1.7, 1.95, 1.6, 0.7, { align:'center', color:c[2] === TEAL ? TEAL : DK });
    [1.35, 1.52].forEach((d, r) => {
      card(s, 1.9 + k * 1.7, 2.75 + r * 0.85, 1.6, 0.75, c[2], 'cell');
      T(s, fmt(d) + ' г/см³', 1.9 + k * 1.7, 2.75 + r * 0.85, 1.6, 0.75, { fontSize:14, bold:true, color:WH, align:'center', valign:'middle' });
    });
  });
  T(s, 'плотность', 0.6, 2.75, 1.2, 0.75, { fontSize:14, color:MUTED, valign:'middle' });
  T(s, 'все почвы незасолённые: ECe 0,28–0,61 дСм/м', 1.9, 4.5, 5.0, 0.4, { fontSize:14, color:MUTED });
  // arithmetic
  card(s, 7.3, 1.5, 5.4, 1.45, DK, 'count-card');
  T(s, '6 × 3 × 10 = 180', 7.55, 1.55, 4.9, 0.8, { fontSize:36, bold:true, color:WH, valign:'middle' });
  T(s, 'вариантов × колонок × ступеней влажности', 7.55, 2.35, 4.9, 0.5, { fontSize:14, color:OCHRE });
  // exam analogy
  [['Датчик с заводским уравнением', 'студент', TEAL], ['Термостатно-весовой метод', 'ключ с ответами (неопределённость 0,005)', TERRA]].forEach((r, k) => {
    const y = 3.15 + k * 0.95;
    card(s, 7.3, y, 5.4, 0.8, LT, 'exam-' + k);
    T(s, [{ text:r[0], options:{ bold:true, fontSize:16, breakLine:true } }, { text:'= ' + r[1], options:{ fontSize:14, color:r[2] === TEAL ? TEAL : TERRA, bold:true } }], 7.5, y, 5.0, 0.8, { valign:'middle' });
  });
  // process
  const proc = ['Насыщение снизу, 48 ч', 'Ступенчатое иссушение при 20 °C', 'Ожидание равновесия 8–36 ч', 'Снимок: период τ + масса колонки'];
  proc.forEach((t, k) => {
    const x = 0.6 + k * 3.1;
    card(s, x, 5.35, 2.85, 1.35, k === 3 ? TEAL : LT, 'proc-' + k);
    T(s, t, x + 0.2, 5.35, 2.45, 1.35, { fontSize:16, bold:true, color:k === 3 ? WH : DK, valign:'middle' });
    if (k < 3) arrowRight(s, x + 2.88, 5.85, 0.24, 0.35);
  });
  T(s, 'Колонка 152 мм × 300 мм, три слоя уплотнения. Каждое решение убирает одну постороннюю причину.', 0.6, 6.78, 12, 0.3, { fontSize:12, color:MUTED });
}

// ---------- 9. target ----------
{
  const s = addSlide(9, 'LIGHT');
  const cx = 3.4, cy = 4.25;
  [2.3, 1.55, 0.8].forEach((r, k) => s.addShape(pres.ShapeType.ellipse, { x:cx - r, y:cy - r, w:2 * r, h:2 * r, fill:{ color: k % 2 ? WH : LT }, line:{ color: SLATE, width:1 }, objectName:'target-ring' }));
  s.addShape(pres.ShapeType.ellipse, { x:cx - 0.07, y:cy - 0.07, w:0.14, h:0.14, fill:{ color: DK }, line:{ type:'none' }, objectName:'bullseye' });
  const gx = cx + 0.75, gy = cy - 0.55;
  [[0, 0], [0.3, -0.2], [-0.25, 0.3], [0.2, 0.35], [-0.3, -0.15], [0.45, 0.05], [-0.05, -0.4]].forEach(([dx, dy]) =>
    s.addShape(pres.ShapeType.ellipse, { x:gx + dx - 0.09, y:gy + dy - 0.09, w:0.18, h:0.18, fill:{ color: TERRA }, line:{ type:'none' }, objectName:'shot' }));
  s.addShape(pres.ShapeType.line, { x:cx, y:cy, w:gx - cx, h:gy - cy, line:{ color: DK, width:2.5, endArrowType:'triangle' }, objectName:'mb-arrow' });
  T(s, 'MB — сдвиг центра кучки', 1.0, 6.65, 4.8, 0.35, { fontSize:14, bold:true, color:TERRA, align:'center' });
  // right
  card(s, 6.9, 1.6, 5.8, 1.5, DK, 'identity');
  T(s, 'RMSE² = MB² + разброс²', 7.15, 1.6, 5.3, 1.0, { fontSize:30, bold:true, color:WH, valign:'middle' });
  T(s, 'типичный промах складывается из сдвига и разброса', 7.15, 2.55, 5.3, 0.45, { fontSize:14, color:OCHRE });
  card(s, 6.9, 3.3, 5.8, 1.3, LT, 'mb0');
  T(s, [{ text:'MB ≈ 0: ', options:{ bold:true, color:TEAL } }, { text:'ошибки случайны, выше и ниже, при усреднении гасятся' }], 7.15, 3.3, 5.3, 1.3, { fontSize:16, valign:'middle' });
  card(s, 6.9, 4.8, 5.8, 1.3, LT, 'mbpos');
  T(s, [{ text:'MB > 0: ', options:{ bold:true, color:TERRA } }, { text:'датчик систематически завышает, усреднение не помогает; это опасно для порога полива' }], 7.15, 4.8, 5.3, 1.3, { fontSize:16, valign:'middle' });
  T(s, 'Критерий приемлемости: RMSE ≤ 0,02', 6.9, 6.3, 5.8, 0.45, { fontSize:20, bold:true, color:TEAL });
}

pres.addSection({ title:'Результаты' });
// ---------- 10. result 1 ----------
{
  const s = addSlide(10, 'LIGHT', 'Результаты');
  const labels = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2'];
  s.addChart([
    { type:pres.charts.BAR, data:[{ name:'MB (сдвиг)', labels, values:[0.004, 0.002, 0.007, 0.006, 0.010, 0.008] }, { name:'RMSE', labels, values:[0.007, 0.006, 0.011, 0.009, 0.014, 0.012] }],
      options:{ barDir:'col', chartColors:[H.accent4, H.accent1], barGapWidthPct:60, showValue:true, dataLabelFormatCode:'0.000', dataLabelFontSize:11, dataLabelColor:H.dk1, dataLabelPosition:'outEnd' } },
    { type:pres.charts.LINE, data:[{ name:'Критерий 0,02', labels, values:[0.02, 0.02, 0.02, 0.02, 0.02, 0.02] }],
      options:{ chartColors:[H.accent3], lineSize:2, lineDash:['dash'], lineDataSymbol:'none' } }
  ], Object.assign({}, chartBase, { x:0.6, y:1.5, w:7.6, h:5.3, valAxisMinVal:0, valAxisMaxVal:0.025, valAxisLabelFormatCode:'0.000', showLegend:true, legendPos:'b',
    showCatAxisTitle:true, catAxisTitle:'вариант: группа (A, B, C) и плотность (1 — 1,35; 2 — 1,52 г/см³)', catAxisTitleFontSize:12, catAxisTitleColor:H.dk2, catAxisTitleFontFace:'+mn-lt', valAxisTitle:'м³/м³', showValAxisTitle:true, valAxisTitleFontSize:12, valAxisTitleColor:H.dk2 }));
  const k = [['0,006–0,014', 'RMSE во всех шести вариантах: ниже критерия 0,02', TEAL], ['1 : 1,5 : 2', 'рост RMSE от группы A к B и C (0,0065; 0,0100; 0,0130)', BROWN], ['MB > 0', 'везде завышение; в A неотличимо от нуля, в B и C +0,006…+0,010', TERRA]];
  k.forEach((c, i) => {
    const y = 1.5 + i * 1.8;
    card(s, 8.6, y, 4.1, 1.6, LT, 'kpi-' + i);
    T(s, c[0], 8.85, y + 0.1, 3.7, 0.6, { fontSize:28, bold:true, color:c[2], valign:'middle' });
    T(s, c[1], 8.85, y + 0.75, 3.7, 0.8, { fontSize:14 });
  });
}

// ---------- 11. result 2 ----------
{
  const s = addSlide(11, 'LIGHT');
  const X = [0.05, 0.10, 0.20, 0.30, 0.38, 0.40, 0.44, 0.48];
  s.addChart(pres.charts.SCATTER, [
    { name:'θ', values:X },
    { name:'A (10–20 % физ. глины)', values:[0, 0.001, 0.003, 0.005, 0.010, 0.011, null, null] },
    { name:'B (40–50 %)', values:[0.001, 0.002, 0.005, 0.009, 0.017, 0.019, null, null] },
    { name:'C (50–65 %)', values:[0.001, 0.002, 0.006, 0.012, 0.022, 0.027, null, null] },
    { name:'C: состояния вне анализа', values:[null, null, null, null, null, 0.027, 0.031, 0.038] },
    { name:'Паспортный предел 0,025', values:X.map(() => 0.025) }
  ], Object.assign({}, chartBase, { x:0.6, y:1.5, w:8.0, h:5.35, lineSize:2.25, lineDataSymbolSize:8, chartColors:[H.accent1, H.accent4, H.accent3, 'E3A98F', H.accent5],
    valAxisMinVal:0, valAxisMaxVal:0.04, valAxisMajorUnit:0.005, catAxisMinVal:0, catAxisMaxVal:0.5, catAxisMajorUnit:0.1, valAxisLabelFormatCode:'General',
    showLegend:true, legendPos:'b', showCatAxisTitle:true, catAxisTitle:'влажность по эталону θ, м³/м³', catAxisTitleFontSize:12, catAxisTitleColor:H.dk2, catAxisTitleFontFace:'+mn-lt',
    showValAxisTitle:true, valAxisTitle:'ошибка Δθ, м³/м³', valAxisTitleFontSize:12, valAxisTitleColor:H.dk2, valAxisTitleFontFace:'+mn-lt' }));
  T(s, 'Прирост ошибки при θ от 0,20 до 0,40', 8.9, 1.5, 3.8, 0.4, { fontSize:14, bold:true, color:MUTED });
  [['A', '+0,008', TEAL], ['B', '+0,014', BROWN], ['C', '+0,021', TERRA]].forEach((r, k) => {
    const y = 1.95 + k * 0.85;
    card(s, 8.9, y, 3.8, 0.7, LT, 'inc-' + k);
    circ(s, 9.05, y + 0.1, 0.5, r[2], r[0], 16);
    T(s, r[1] + ' м³/м³', 9.75, y, 2.8, 0.7, { fontSize:22, bold:true, valign:'middle' });
  });
  card(s, 8.9, 4.6, 3.8, 2.25, DK, 'ratio-card');
  T(s, '2,6×', 9.15, 4.65, 3.3, 0.9, { fontSize:44, bold:true, color:OCHRE, valign:'middle' });
  T(s, 'во столько раз ошибка в C прирастает сильнее, чем в A. Пустые состояния (0,44 и 0,48) показывают +0,031 и +0,038: кривая не выходит на плато', 9.15, 5.55, 3.3, 1.25, { fontSize:14, color:WH });
}

// ---------- 12. variance partition ----------
{
  const s = addSlide(12, 'LIGHT');
  s.addChart(pres.charts.DOUGHNUT, [{ name:'Доля изменчивости ошибки', labels:['Почва 9,1 %', 'Плотность 1,0 %', 'Взаимодействие 0,1 %', 'Внутри вариантов 89,8 % (рост с влажностью)'], values:[9.1, 1.0, 0.1, 89.8] }],
    { x:0.6, y:1.5, w:5.8, h:5.3, holeSize:58, chartColors:[H.accent3, H.accent4, H.accent2, 'C9D1CF'], showLegend:true, legendPos:'b', legendFontSize:12, legendFontFace:'+mn-lt', legendColor:H.dk2, showPercent:false, showValue:false, dataBorder:{ pt:1, color:'FFFFFF' } });
  card(s, 6.8, 1.5, 5.9, 1.55, LT, 'between');
  T(s, [{ text:'≈ 10 % между вариантами', options:{ bold:true, fontSize:20, color:TERRA, breakLine:true } }, { text:'то, что мы хотим «услышать»: влияние почвы и плотности', options:{ fontSize:14 } }], 7.05, 1.5, 5.4, 1.55, { valign:'middle' });
  card(s, 6.8, 3.2, 5.9, 1.55, LT, 'within');
  T(s, [{ text:'≈ 90 % внутри вариантов', options:{ bold:true, fontSize:20, color:SLATE, breakLine:true } }, { text:'«музыка на концерте»: почти целиком рост ошибки с влажностью', options:{ fontSize:14 } }], 7.05, 3.2, 5.4, 1.55, { valign:'middle' });
  card(s, 6.8, 4.9, 5.9, 1.9, DK, 'verdict');
  T(s, 'Почва: F(2; 54) = 2,74; p = 0,074', 7.05, 5.0, 5.4, 0.55, { fontSize:20, bold:true, color:WH, valign:'middle' });
  T(s, [{ text:'«Не значимо» ≠ «эффекта нет». ', options:{ bold:true, color:OCHRE } }, { text:'Тест говорит лишь, что на таком шумовом фоне эффект не слышен', options:{ color:WH } }], 7.05, 5.6, 5.4, 1.1, { fontSize:14 });
}

// ---------- 13. same-checkpoint comparison ----------
{
  const s = addSlide(13, 'LIGHT');
  const lv = ['0,05', '0,10', '0,20', '0,30', '0,38', '0,40'];
  const rows = [['A', ['0,000', '0,001', '0,003', '0,005', '0,010', '0,011'], TEAL], ['B', ['0,001', '0,002', '0,005', '0,009', '0,017', '0,019'], BROWN], ['C', ['0,001', '0,002', '0,006', '0,012', '0,022', '0,027'], TERRA]];
  const hdr = { bold:true, color:WH, fill:{ color: DK }, align:'center', valign:'middle', fontSize:14 };
  const tbl = [[{ text:'θ, м³/м³', options:Object.assign({}, hdr, { align:'left' }) }].concat(lv.map(v => ({ text:v, options:hdr })))];
  rows.forEach(r => tbl.push([{ text:'Группа ' + r[0], options:{ bold:true, fontSize:16, color:WH, fill:{ color: r[2] }, valign:'middle' } }].concat(r[1].map(v => ({ text:v, options:{ fontSize:16, align:'center', valign:'middle', color:DK, fill:{ color: LT } } })))));
  tbl.push([{ text:'Порядок', options:{ bold:true, fontSize:14, color:MUTED, valign:'middle' } }].concat(lv.map((v, i) => ({ text:i < 2 ? 'A < B = C' : 'A < B < C', options:{ bold:true, fontSize:14, color:TEAL, align:'center', valign:'middle' } }))));
  s.addTable(tbl, { x:0.6, y:1.6, w:12.1, colW:[2.2, 1.65, 1.65, 1.65, 1.65, 1.65, 1.65], rowH:[0.5, 0.65, 0.65, 0.65, 0.55], border:{ type:'solid', pt:1, color:'FFFFFF' }, margin:0.05, objectName:'checkpoints' });
  T(s, 'Ошибка Δθ в шести «контрольных отметках» влажности, плотность 1,35 г/см³: порядок не нарушен ни разу', 0.6, 4.68, 12.1, 0.3, { fontSize:14, color:MUTED, italic:true });
  card(s, 0.6, 5.15, 3.9, 1.7, DK, 'anova2');
  T(s, [{ text:'Двухфакторный анализ, влажность — блок', options:{ fontSize:14, color:OCHRE, bold:true, breakLine:true } }, { text:'F(2; 10) = 6,92\np = 0,013', options:{ fontSize:22, color:WH, bold:true } }], 0.85, 5.15, 3.5, 1.7, { valign:'middle' });
  card(s, 4.7, 5.15, 3.9, 1.7, DK, 'friedman');
  T(s, [{ text:'Критерий Фридмана: только порядок, без допущений о форме', options:{ fontSize:14, color:OCHRE, bold:true, breakLine:true } }, { text:'χ²(2) = 11,27\np = 0,004', options:{ fontSize:22, color:WH, bold:true } }], 4.95, 5.15, 3.5, 1.7, { valign:'middle' });
  card(s, 8.8, 5.15, 3.9, 1.7, WH, 'caveat', { line:{ color: OCHRE, width:1.5 } });
  T(s, [{ text:'Граница вывода', options:{ bold:true, color:TERRA, breakLine:true } }, { text:'одна почва на группу: порядок устойчив, но перенос на другие почвы — гипотеза', options:{} }], 9.05, 5.15, 3.4, 1.7, { fontSize:14, valign:'middle' });
}

// ---------- 14. clay or salts ----------
{
  const s = addSlide(14, 'LIGHT');
  const mb = [0.004, 0.002, 0.007, 0.006, 0.010, 0.008];
  // regression line is drawn as a second scatter group
  const MH = [2.1, 2.3, 6.5, 6.8, 11.4, 11.1], EC = [0.35, 0.28, 0.52, 0.45, 0.61, 0.49];
  const mkChart = (x0, xvals, a, b, xt, ttl, minx, maxx, major, f, lx) => {
    s.addChart(pres.charts.SCATTER, [{ name:'x', values:xvals.concat(lx) }, { name:'шесть вариантов', values:mb.concat([null, null]) }, { name:'аппроксимация', values:xvals.map(() => null).concat(lx.map(v => a + b * v)) }
    ], Object.assign({}, chartBase, { x:x0, y:1.5, w:5.95, h:3.9, lineSize:2, lineDataSymbolSize:11, chartColors:[H.accent3, H.accent1], showLegend:false, valAxisMinVal:0, valAxisMaxVal:0.012, valAxisMajorUnit:0.002, valAxisLabelFormatCode:'General',
      catAxisMinVal:minx, catAxisMaxVal:maxx, catAxisMajorUnit:major, catAxisLabelFormatCode:f, showCatAxisTitle:true, catAxisTitle:xt, catAxisTitleFontSize:12, catAxisTitleColor:H.dk2, catAxisTitleFontFace:'+mn-lt',
      showTitle:true, title:ttl, titleFontSize:14, titleColor:H.dk1, titleFontFace:'+mn-lt', showValAxisTitle:true, valAxisTitle:'MB, м³/м³', valAxisTitleFontSize:12, valAxisTitleColor:H.dk2, valAxisTitleFontFace:'+mn-lt' }));
  };
  mkChart(0.6, MH, 0.0017, 0.00066, 'MH, %', 'Поверхность (гигроскопичность): R² = 0,88', 0, 12, 2, '0', [0, 12]);
  mkChart(6.75, EC, -0.0044, 0.0235, 'ECe, дСм/м', 'Электропроводность вытяжки: R² = 0,96', 0, 0.7, 0.1, '0.0', [0.25, 0.65]);
  card(s, 0.6, 5.6, 4.0, 1.25, LT, 'tied');
  T(s, [{ text:'r = 0,88', options:{ bold:true, fontSize:24, color:TERRA, breakLine:true } }, { text:'между MH и ECe: подозреваемые всегда вместе', options:{ fontSize:14 } }], 0.85, 5.6, 3.6, 1.25, { valign:'middle' });
  card(s, 4.8, 5.6, 4.0, 1.25, LT, 'fragile');
  T(s, [{ text:'p = 0,017', options:{ bold:true, fontSize:24, color:TEAL, breakLine:true } }, { text:'ECe значима в модели с обоими, но вывод хрупкий; три почвы, а не шесть', options:{ fontSize:14 } }], 5.05, 5.6, 3.6, 1.25, { valign:'middle' });
  card(s, 9.0, 5.6, 3.7, 1.25, DK, 'need');
  T(s, 'Нужны почвы, где глина и соли не идут вместе', 9.2, 5.6, 3.3, 1.25, { fontSize:16, bold:true, color:WH, valign:'middle' });
}

// ---------- 15. density ----------
{
  const s = addSlide(15, 'LIGHT');
  s.addChart(pres.charts.BAR, [
    { name:'Прогноз CRIM', labels:['A', 'B', 'C'], values:[9.5, 9.5, 9.5] },
    { name:'Наблюдалось', labels:['A', 'B', 'C'], values:[-2, -1, -2] } ],
    Object.assign({}, chartBase, { x:0.6, y:1.5, w:5.9, h:5.3, barDir:'col', chartColors:[H.accent5, H.accent3], barGapWidthPct:60, showValue:true, dataLabelFontSize:14, dataLabelColor:H.dk1, dataLabelPosition:'outEnd', dataLabelFormatCode:'+0.0;−0.0',
      valAxisMinVal:-4, valAxisMaxVal:12, valAxisMajorUnit:2, showLegend:true, legendPos:'b', showCatAxisTitle:true, catAxisTitle:'группа', catAxisTitleFontSize:12, catAxisTitleColor:H.dk2, catAxisTitleFontFace:'+mn-lt',
      showValAxisTitle:true, valAxisTitle:'сдвиг показаний при 1,52 против 1,35 г/см³, ×10⁻³ м³/м³', valAxisTitleFontSize:12, valAxisTitleColor:H.dk2, valAxisTitleFontFace:'+mn-lt', catAxisLabelPos:'low' }));
  card(s, 6.9, 1.5, 5.8, 1.35, DK, 'ci');
  T(s, [{ text:'95 % ДИ эффекта плотности: −0,006…+0,003', options:{ bold:true, fontSize:18, color:WH, breakLine:true } }, { text:'исключает прогноз +0,0095', options:{ fontSize:14, color:OCHRE } }], 7.15, 1.5, 5.3, 1.35, { valign:'middle' });
  T(s, 'Прогноз = связка допущений. Какое подвело?', 6.9, 3.05, 5.8, 0.4, { fontSize:16, bold:true, color:TEAL });
  [['«При той же влажности»', 'диапазоны влажности вариантов разные, а ошибка растёт с влажностью'], ['«При той же проводимости»', 'у плотных вариантов ECe ниже на 0,07–0,12 дСм/м; с ECe вместо MH коэффициент плотности +0,0004, p = 0,51']].forEach((r, k) => {
    const y = 3.55 + k * 1.55;
    card(s, 6.9, y, 5.8, 1.4, LT, 'assump-' + k);
    circ(s, 7.1, y + 0.45, 0.5, TERRA, '×', 20);
    T(s, [{ text:r[0], options:{ bold:true, fontSize:16, breakLine:true } }, { text:r[1], options:{ fontSize:14 } }], 7.8, y, 4.75, 1.4, { valign:'middle' });
  });
  T(s, 'Малый реальный эффект плотности нельзя ни подтвердить, ни исключить', 6.9, 6.7, 5.8, 0.3, { fontSize:12, color:MUTED, italic:true });
}

// ---------- 16. swelling ----------
{
  const s = addSlide(16, 'LIGHT');
  const draw = (x, title, soilH, rodNote) => {
    T(s, title, x, 1.5, 3.8, 0.4, { fontSize:16, bold:true, align:'center' });
    s.addShape(pres.ShapeType.rect, { x:x + 0.7, y:5.55 - soilH, w:2.4, h:soilH, fill:{ color: TERRA }, line:{ type:'none' }, objectName:'soil' });
    s.addShape(pres.ShapeType.rect, { x:x + 0.7, y:2.35, w:2.4, h:3.2, fill:{ type:'none' }, line:{ color: SLATE, width:2 }, objectName:'column' });
    [1.4, 2.4].forEach(dx => s.addShape(pres.ShapeType.rect, { x:x + dx, y:2.35, w:0.07, h:3.2, fill:{ color: OCHRE }, line:{ type:'none' }, objectName:'rod' }));
    T(s, rodNote, x, 5.65, 3.8, 0.4, { fontSize:14, align:'center', color:MUTED });
  };
  draw(0.6, 'Сухая: 300 мм', 3.2, 'почва = длина штанг');
  draw(4.7, 'Влажная: ≈ 325 мм', 3.2 + 0.0, '');
  // swollen: draw extra 0.27 inch? scale 3.2in = 300mm -> 25mm = 0.27in
  s.addShape(pres.ShapeType.rect, { x:5.4, y:2.08, w:2.4, h:0.27, fill:{ color: TERRA }, line:{ type:'none' }, objectName:'swollen-top' });
  T(s, 'почва выше штанг на ≈ 25 мм', 4.7, 5.65, 3.8, 0.4, { fontSize:14, align:'center', color:TERRA, bold:true });
  [['≈ 25 мм', 'подъём поверхности'], ['8,3 %', 'объёмная деформация'], ['0,460 > 0,415', 'θmax больше исходной пористости']].forEach((r, k) => {
    const y = 1.5 + k * 1.3;
    card(s, 8.9, y, 3.8, 1.15, LT, 'swell-' + k);
    T(s, r[0], 9.1, y + 0.05, 3.4, 0.6, { fontSize:26, bold:true, color:TERRA, valign:'middle' });
    T(s, r[1], 9.1, y + 0.65, 3.4, 0.4, { fontSize:14 });
  });
  card(s, 0.6, 6.15, 12.1, 0.8, WH, 'swell-note', { line:{ color: OCHRE, width:1.5 } });
  T(s, 'Если делить воду на прежний объём, эталон завышен, а ошибка датчика выглядит заниженной. Поэтому эталон считали по реальному объёму: дополнительной ошибки от набухания не нашли.', 0.85, 6.15, 11.6, 0.8, { fontSize:14, valign:'middle' });
}

pres.addSection({ title:'Что делать и куда двигаться' });
// ---------- 17. traffic light ----------
{
  const s = addSlide(17, 'LIGHT', 'Что делать и куда двигаться');
  const rows = [[GREEN, 'Лёгкие, 10–20 % физ. глины', 'заводская калибровка без поправки'], [OCHRE, 'Тяжёлые, 40–65 %', 'RMSE ≤ 0,02, но завышение 0,006–0,010, до 0,027 во влажной. Для порогов вблизи наименьшей влагоёмкости: калибровка под почву или 1–2 весовые проверки за сезон'], [TERRA, 'Постоянная поправка', 'не годится: ошибка растёт с влажностью'], [SLATE, 'Серая зона', '20–40 % физ. глины не проверяли; в поле нужна температурная поправка']];
  rows.forEach((r, k) => {
    const y = 1.5 + k * 1.35;
    card(s, 0.6, y, 6.9, 1.2, LT, 'tl-' + k);
    s.addShape(pres.ShapeType.ellipse, { x:0.8, y:y + 0.25, w:0.7, h:0.7, fill:{ color: r[0] }, line:{ type:'none' }, objectName:'light' });
    T(s, [{ text:r[1], options:{ bold:true, fontSize:16, breakLine:true } }, { text:r[2], options:{ fontSize:14, color:MUTED } }], 1.75, y, 5.6, 1.2, { valign:'middle' });
  });
  const lv = ['0,05', '0,10', '0,20', '0,30', '0,38', '0,40'];
  s.addChart(pres.charts.LINE, [
    { name:'Ошибка в группе C', labels:lv, values:[0.001, 0.002, 0.006, 0.012, 0.022, 0.027] },
    { name:'Остаток после вычитания +0,010', labels:lv, values:[-0.009, -0.008, -0.004, 0.002, 0.012, 0.017] } ],
    Object.assign({}, chartBase, { x:7.8, y:1.5, w:4.9, h:4.4, chartColors:[H.accent3, H.accent1], lineSize:2.5, lineDataSymbolSize:7, valAxisMinVal:-0.01, valAxisMaxVal:0.03, valAxisMajorUnit:0.01, valAxisLabelFormatCode:'0.000',
      showLegend:true, legendPos:'b', showTitle:true, title:'Постоянная поправка +0,010 в группе C', titleFontSize:14, titleColor:H.dk1, titleFontFace:'+mn-lt',
      showCatAxisTitle:true, catAxisTitle:'θ, м³/м³', catAxisTitleFontSize:12, catAxisTitleColor:H.dk2, catAxisTitleFontFace:'+mn-lt', catAxisLabelPos:'low' }));
  card(s, 7.8, 6.05, 4.9, 0.85, DK, 'scale-note');
  T(s, '0,010 ≈ 10 мм воды в метровом слое, 0,027 ≈ 27 мм', 8.0, 6.05, 4.5, 0.85, { fontSize:14, bold:true, color:WH, valign:'middle' });
}

// ---------- 18. next steps ----------
{
  const s = addSlide(18, 'LIGHT');
  const hd = { bold:true, color:WH, fill:{ color: DK }, fontSize:14, valign:'middle' };
  const cell = (t, o = {}) => ({ text:t, options:Object.assign({ fontSize:15, valign:'middle', color:DK, fill:{ color: LT } }, o) });
  const rowsT = [
    [{ text:'Вопрос, на который данных нет', options:hd }, { text:'Решающий эксперимент', options:hd }, { text:'Как читать результат', options:hd }],
    [cell('Глина или соли?', { bold:true }), cell('Развести их: высокая глина при низкой солёности и наоборот; ступенчатое засоление одной почвы (часть сделана, 1–4 дСм/м, отдельная статья)'), cell('Ошибка идёт за солью → механизм в растворе; за поверхностью → в проводимости глин')],
    [cell('Можно ли учесть проводимость в поправке? (идея, не результат)', { bold:true, color:TERRA }), cell('Мерить рядом объёмную электропроводность почвы и строить поправку по двум входам: τ и ЭП'), cell('Если остаток исчезает → поправка не требует калибровки под каждую почву')],
    [cell('Насколько устойчив эффект почвы?', { bold:true }), cell('Переанализ уже имеющихся 180 значений смешанной моделью: колонка — случайный эффект'), cell('Эффект сохраняется → снимается главное ограничение статистики')],
    [cell('Верно ли это для класса почв, а не для одной почвы?', { bold:true }), cell('Больше почв (в т. ч. 20–40 % физ. глины), минералогия глин, ветвь увлажнения, температура, поле'), cell('Каждый шаг превращает «в изученных почвах» в «для класса почв»')]
  ];
  s.addTable(rowsT, { x:0.6, y:1.5, w:12.1, colW:[3.4, 4.9, 3.8], rowH:[0.5, 1.3, 1.1, 1.0, 1.1], border:{ type:'solid', pt:1.5, color:'FFFFFF' }, margin:0.1, objectName:'next-steps' });
  T(s, 'Идея помечена как гипотеза: объёмную электропроводность в этом опыте мы не измеряли', 0.6, 6.62, 12.1, 0.3, { fontSize:12, color:MUTED, italic:true });
}

// ---------- 19. conclusions ----------
{
  const s = addSlide(19, 'DARK');
  const tk = [['1', 'Годится, но не бесплатно', 'RMSE 0,006–0,014 < 0,02 везде. Но в почвах с 40–65 % физ. глины датчик завышает и ошибка растёт с влажностью'], ['2', 'Доказывали устойчивость, а не только p', 'Порядок групп сохранялся на каждом уровне влажности, два разных критерия согласны. Слабые места показаны'], ['3', 'Гипотеза подтверждена частично', 'Почва и рост с влажностью: да. Плотность: нет. Механизм (глина или соли): не разделён']];
  tk.forEach((t, k) => {
    const x = 0.6 + k * 4.1;
    card(s, x, 1.55, 3.9, 3.5, '3D2E22', 'take-' + k);
    circ(s, x + 0.3, 1.8, 0.7, OCHRE, t[0], 22, DK);
    T(s, t[1], x + 0.3, 2.7, 3.3, 0.9, { fontSize:20, bold:true, color:WH });
    T(s, t[2], x + 0.3, 3.65, 3.3, 1.3, { fontSize:14, color:WH });
  });
  T(s, 'Границы применимости', 0.6, 5.25, 6, 0.35, { fontSize:14, bold:true, color:OCHRE });
  T(s, 'одна почва на группу · лаборатория, 20 °C, только иссушение · стенки колонки · разброс между датчиками ±0,005…0,015 · неопределённость эталона 0,005 · статистика по сводным данным вариантов', 0.6, 5.65, 12.1, 0.7, { fontSize:14, color:WH });
  T(s, 'Работа — карта: где датчик надёжен, где нужна проверка и какой эксперимент закроет пробел', 0.6, 6.4, 12.1, 0.5, { fontSize:18, bold:true, color:OCHRE });
}


// per-series styling that pptxgenjs cannot express: hide connecting line / markers, dash a limit line
async function patchSeries(file) {
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(require('fs').readFileSync(file));
  const rules = { 'шесть вариантов':{ noLine:true }, 'Паспортный предел 0,025':{ dash:'dash', noMarker:true }, 'аппроксимация':{ noMarker:true }, 'C: состояния вне анализа':{ dash:'sysDash' } };
  for (const name of Object.keys(zip.files).filter(n => /^ppt\/charts\/chart\d+\.xml$/.test(n))) {
    let xml = await zip.file(name).async('string');
    xml = xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ser => {
      const m = ser.match(/<c:tx>[\s\S]*?<c:v>([^<]*)<\/c:v>/); const r = m && rules[m[1]]; if (!r) return ser;
      if (r.noLine) ser = ser.replace(/(<c:spPr>[\s\S]*?)<a:ln[\s\S]*?<\/a:ln>/, '$1<a:ln w="28575"><a:noFill/></a:ln>');
      if (r.dash) ser = ser.replace(/(<c:spPr>[\s\S]*?<a:ln[^>]*>[\s\S]*?)<a:prstDash val="solid"\/>/, `$1<a:prstDash val="${r.dash}"/>`);
      if (r.noMarker) ser = ser.replace(/<c:marker>[\s\S]*?<\/c:marker>/, '<c:marker><c:symbol val="none"/></c:marker>');
      return ser;
    });
    zip.file(name, xml);
  }
  require('fs').writeFileSync(file, await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE' }));
}

(async () => {
  const out = path.resolve(process.argv[2] || 'deck.pptx');
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  await patchSeries(out);
  console.log('written', out);
})();
