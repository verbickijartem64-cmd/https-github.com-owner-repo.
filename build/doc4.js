// Версия 4: текст доклада на 10 минут.
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  AlignmentType, Footer, PageNumber, LevelFormat } = require('docx');
const K = require('./content4.json');
const K3 = require('./content3.js');

const FONT = 'Calibri', HEAD = 'Cambria';
const COL = { dk:'2A1E16', teal:'1C7C8C', terra:'B8552F', muted:'5B4636', lt:'F1F3F2' };
const wc = s => s.split(/\s+/).filter(Boolean).length;
const WPM = 140, SWITCH = 4;
const secs = K.slides.map(d => Math.round(wc(d.text.join(' ')) / WPM * 60));
const words = K.slides.map(d => wc(d.text.join(' ')));
const mmss = t => `${Math.floor(t / 60)}:${String(t % 60).padStart(2, '0')}`;
const cum = n => secs.slice(0, n).reduce((a, b) => a + b, 0) + SWITCH * n;
const totalWords = words.reduce((a, b) => a + b, 0);
const totalSpeech = secs.reduce((a, b) => a + b, 0);

const run = (t, o = {}) => new TextRun(Object.assign({ text:t, font:FONT, size:24 }, o));
const p = (children, o = {}) => new Paragraph(Object.assign({ spacing:{ after:140, line:312 } }, o, { children: typeof children === 'string' ? [run(children)] : children }));
const h1 = (t, br = true) => new Paragraph({ heading:HeadingLevel.HEADING_1, pageBreakBefore:br, spacing:{ before:120, after:200 }, children:[new TextRun({ text:t, font:HEAD, size:34, bold:true, color:COL.dk })] });
const h2 = t => new Paragraph({ heading:HeadingLevel.HEADING_2, keepNext:true, spacing:{ before:300, after:80 }, children:[new TextRun({ text:t, font:HEAD, size:27, bold:true, color:COL.teal })] });
const bullet = children => new Paragraph({ numbering:{ reference:'bul', level:0 }, spacing:{ after:90, line:290 }, children: typeof children === 'string' ? [run(children, { size:22 })] : children });
const bd = { style:BorderStyle.SINGLE, size:4, color:'C9D1CF' };
const borders = { top:bd, bottom:bd, left:bd, right:bd };
function cell(text, w, o = {}) {
  return new TableCell({ width:{ size:w, type:WidthType.DXA }, borders, margins:{ top:60, bottom:60, left:100, right:100 },
    shading:o.fill ? { fill:o.fill, type:ShadingType.CLEAR, color:'auto' } : undefined, verticalAlign:'center',
    children:[new Paragraph({ alignment:o.align, children:[new TextRun({ text:String(text), font:FONT, size:o.size || 20, bold:o.bold, color:o.color })] })] });
}
function table(cols, rows, opt = {}) {
  const W = cols.reduce((a, b) => a + b, 0);
  return new Table({ width:{ size:W, type:WidthType.DXA }, columnWidths:cols,
    rows:rows.map((r, i) => new TableRow({ tableHeader:i === 0, cantSplit:true,
      children:r.map((t, k) => cell(t, cols[k], i === 0 ? { fill:COL.dk, color:'FFFFFF', bold:true } : Object.assign({ fill:i % 2 ? COL.lt : undefined }, opt.boldFirst && k === 0 ? { bold:true } : {}, opt.boldLast && k === r.length - 1 ? { bold:true } : {}))) })) });
}

const body = [];
body.push(new Paragraph({ spacing:{ after:80 }, children:[new TextRun({ text:'Текст доклада · 10 минут', font:HEAD, size:48, bold:true, color:COL.dk })] }));
body.push(new Paragraph({ spacing:{ after:80 }, children:[new TextRun({ text:'Когда датчик влажности «льстит» почве', font:HEAD, size:30, bold:true, color:COL.teal })] }));
body.push(new Paragraph({ spacing:{ after:160 }, children:[new TextRun({ text:'Систематическая погрешность заводской калибровки датчиков влажности CS616 в незасолённых почвах разного гранулометрического состава', font:FONT, size:24, italics:true, color:COL.terra })] }));
body.push(p([run('Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е. · Кубанский государственный аграрный университет имени И. Т. Трубилина', { size:21, color:COL.muted })]));

body.push(h1('1. Регламент и хронометраж', false));
body.push(p([run('Речь — ', {}), run(`${totalWords} слов, ≈ ${mmss(totalSpeech)} при темпе 140 слов в минуту`, { bold:true }), run(`; с паузами на смену 12 слайдов (≈ 4 с каждая) — ≈ ${mmss(totalSpeech + SWITCH * 12)}. Запас до 10:00 оставлен на естественные паузы и реакцию зала.`)]));
body.push(p('Колонка «К концу слайда» — контрольные точки: если вы отстаёте больше чем на 20 секунд, сократите следующий слайд до его главной мысли (она выделена в начале каждого слайда). Запасные слайды 13–21 показывают только при вопросах — во время доклада они не используются.'));
body.push(table([600, 5300, 1000, 1200, 1300], [['№', 'Слайд', 'Слов', 'Время', 'К концу слайда']].concat(K.slides.map((d, i) => [String(d.n), d.title, String(words[i]), `≈ ${secs[i]} с`, mmss(cum(i + 1))])), { boldLast:true }));

body.push(h1('2. Текст по слайдам'));
K.slides.forEach((d, i) => {
  body.push(h2(`Слайд ${d.n}. ${d.title}`));
  body.push(p([run(`≈ ${secs[i]} с · к концу слайда ${mmss(cum(i + 1))}`, { size:20, color:COL.muted })], { spacing:{ after:60 } }));
  body.push(p([run('Главная мысль: ', { bold:true, color:COL.terra }), run(d.takeaway, { bold:true })]));
  d.text.forEach(t => body.push(p(t)));
});

body.push(h1('3. Вопросы и ответы'));
body.push(p('Ответы рассчитаны на 20–40 секунд. Для вопросов о числах и методике держите наготове запасные слайды (раздел 4).'));
K.qa.forEach((q, i) => {
  body.push(new Paragraph({ keepNext:true, spacing:{ before:200, after:80 }, children:[run(`Вопрос ${i + 1}. ${q.q}`, { bold:true })] }));
  body.push(p([run('Ответ. ', { italics:true, bold:true }), run(q.a)]));
});

body.push(h1('4. Запасные слайды: когда показывать'));
[
  ['13', 'Шесть теорий и типы связей', 'если спрашивают, на чём стоит работа и как связаны теории'],
  ['14', 'Допущения нашего опыта', 'вопросы о надёжности эталона, статистики, колонок'],
  ['15', 'Расчёт прогноза модели смеси', 'если просят объяснить, откуда +0,0095'],
  ['16', 'Набухание и пересчёт эталона', 'вопросы о набухающих почвах'],
  ['17', 'Ловушка терминов «глина»', 'сравнение с паспортом и зарубежными работами'],
  ['18', 'Статистика ошибки по вариантам', 'точные MB, RMSE, доверительные интервалы, p'],
  ['19', 'Дисперсионный анализ ошибки', 'вопросы о p = 0,074 и долях изменчивости'],
  ['20', 'Свойства почв', 'засолённость, плотность частиц, гигроскопичность'],
  ['21', 'Ограничения исследования', 'перенос на поле, методика, статистика']
].forEach(r => body.push(bullet([run(`Слайд ${r[0]}. ${r[1]}: `, { bold:true, size:22 }), run(r[2], { size:22 })])));

body.push(h1('5. Для подготовки (не произносится)'));
body.push(h2('Насколько твёрд каждый вывод'));
body.push(table([2300, 2300, 1400, 1050, 2350], [['Утверждение', 'Чем подтверждено', 'Тип связи', 'Твёрдость', 'Что мешает сказать сильнее']].concat(K3.ledger), { boldFirst:true }));
body.push(h2('Аналогии и где они ломаются'));
body.push(table([2300, 3500, 3600], [['Аналогия', 'Что она объясняет', 'Где перестаёт работать']].concat(K3.analogies), { boldFirst:true }));
body.push(h2('Что в докладе добавлено сверх статьи'));
[
  'Аналогии — пояснения, не данные; у каждой есть граница (таблица выше).',
  'Перевод 0,01 м³/м³ в 10 мм воды в метровом слое и 100 м³/га — пересчёт единиц; тождество RMSE² = MB² + разброс² — математическое.',
  'Схема срабатывания по порогу на слайде 4 — иллюстрация, а не измерение.',
  'Линии на слайде 7 построены по табл. V статьи (6 уровней, плотность 1,35), а дисперсионный анализ — по всем 60 значениям.',
  '«(1/3)⁶ ≈ 1 из 729» — интуиция логики рангового теста при допущении независимости уровней, а не самостоятельный тест.',
  'Остаток после вычитания постоянной поправки +0,010 на слайде 10 — арифметика по табл. V.',
  'Совет «проверять в двух точках» — следствие из роста ошибки с влажностью; в поле не проверялся, и это сказано на слайде.',
  'Прогнозы «если — то» сформулированы по обсуждению статьи; поправка по двум входам — гипотеза.'
].forEach(t => body.push(bullet(t)));
body.push(h2('Что подтвердить у авторов'));
[
  'Состояния вне анализа (θ = 0,44 и 0,48; Δθ = +0,031 и +0,038) взяты из итоговой статьи и везде названы «не вошедшими в анализ» (слайды 6 и 9). Если авторы их не подтвердят — убрать.',
  'Названия почв и их подтипы; название и дата конференции на титульном слайде.'
].forEach(t => body.push(bullet(t)));

const doc = new Document({
  creator:'Кубанский ГАУ', title:'Текст доклада: датчики CS616 (10 минут)', description:'Версия 4',
  styles:{ default:{ document:{ run:{ font:FONT, size:24 } } },
    paragraphStyles:[
      { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal', quickFormat:true, run:{ size:34, bold:true, font:HEAD }, paragraph:{ spacing:{ before:240, after:200 }, outlineLevel:0 } },
      { id:'Heading2', name:'Heading 2', basedOn:'Normal', next:'Normal', quickFormat:true, run:{ size:27, bold:true, font:HEAD }, paragraph:{ spacing:{ before:300, after:80 }, outlineLevel:1 } } ] },
  numbering:{ config:[{ reference:'bul', levels:[{ level:0, format:LevelFormat.BULLET, text:'•', alignment:AlignmentType.LEFT, style:{ paragraph:{ indent:{ left:540, hanging:270 } } } }] }] },
  sections:[{ properties:{ page:{ size:{ width:11906, height:16838 }, margin:{ top:1134, bottom:1134, left:1247, right:1247 } } },
    footers:{ default:new Footer({ children:[new Paragraph({ alignment:AlignmentType.CENTER, children:[new TextRun({ children:[PageNumber.CURRENT], font:FONT, size:20, color:COL.muted })] })] }) },
    children:body }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2] || 'talk4.docx', b); console.log('ok; words', totalWords, 'speech', mmss(totalSpeech), 'with switches', mmss(totalSpeech + SWITCH * 12)); });
