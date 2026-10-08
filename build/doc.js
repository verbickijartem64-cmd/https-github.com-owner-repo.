const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType, Footer, PageNumber, LevelFormat } = require('docx');
const { slides, qa, analogies } = require('./content.js');

const wc = s => s.text.join(' ').split(/\s+/).length;
const mins = s => Math.max(1, Math.round(wc(s) / 140 * 2) / 2);
const fmt = n => String(n).replace('.', ',');
const total = slides.reduce((a, s) => a + wc(s), 0);
const coreWords = slides.filter(s => s.core).reduce((a, s) => a + wc(s), 0);

const FONT = 'Calibri', HEAD = 'Cambria';
const p = (text, o = {}) => new Paragraph(Object.assign({ spacing:{ after:140, line:300 } }, o, { children: Array.isArray(text) ? text : [new TextRun({ text, font:FONT, size:24 })] }));
const run = (t, o = {}) => new TextRun(Object.assign({ text:t, font:FONT, size:24 }, o));
const h1 = t => new Paragraph({ heading:HeadingLevel.HEADING_1, spacing:{ before:360, after:200 }, children:[new TextRun({ text:t, font:HEAD, size:34, bold:true, color:'2A1E16' })] });
const h2 = t => new Paragraph({ heading:HeadingLevel.HEADING_2, spacing:{ before:300, after:120 }, keepNext:true, children:[new TextRun({ text:t, font:HEAD, size:28, bold:true, color:'1C7C8C' })] });
const bullet = (children) => new Paragraph({ numbering:{ reference:'bul', level:0 }, spacing:{ after:100, line:290 }, children });

const border = { style:BorderStyle.SINGLE, size:4, color:'C9D1CF' };
const borders = { top:border, bottom:border, left:border, right:border };
function cell(text, w, o = {}) {
  return new TableCell({ width:{ size:w, type:WidthType.DXA }, borders, margins:{ top:70, bottom:70, left:110, right:110 },
    shading: o.fill ? { fill:o.fill, type:ShadingType.CLEAR, color:'auto' } : undefined, verticalAlign:'center',
    children:[new Paragraph({ alignment:o.align, children:[new TextRun({ text, font:FONT, size:o.size || 21, bold:o.bold, color:o.color })] })] });
}
function table(cols, rows, head = true) {
  const W = cols.reduce((a, b) => a + b, 0);
  return new Table({ width:{ size:W, type:WidthType.DXA }, columnWidths:cols,
    rows: rows.map((r, i) => new TableRow({ tableHeader: head && i === 0, cantSplit:true, children: r.map((t, k) => cell(t, cols[k], head && i === 0 ? { fill:'2A1E16', color:'FFFFFF', bold:true } : { fill: i % 2 ? 'F1F3F2' : undefined })) })) });
}

const body = [];
body.push(new Paragraph({ spacing:{ after:120 }, children:[new TextRun({ text:'Текст доклада', font:HEAD, size:48, bold:true, color:'2A1E16' })] }));
body.push(new Paragraph({ spacing:{ after:240 }, children:[new TextRun({ text:'Систематическая погрешность заводской калибровки датчиков влажности CS616 в незасолённых почвах разного гранулометрического состава', font:FONT, size:26, italics:true, color:'B8552F' })] }));
body.push(p([run('Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е. Кубанский государственный аграрный университет имени И. Т. Трубилина', { size:21, color:'5B4636' })]));

body.push(h1('1. Как устроен доклад'));
body.push(p('Доклад рассчитан на слушателей из смежных областей: инженеров-мелиораторов, агрономов, физиков, статистиков. Он строится как история в трёх действиях: зачем (слайды 1–7), как мы это проверяли и что нашли (8–16), что с этим делать и куда двигаться (17–19). Каждый специальный термин и каждое скрытое допущение объясняется при первом появлении: сначала на бытовой аналогии, затем строго, с числами из статьи.'));
body.push(p([run('Три правила, которым подчинён текст. ', { bold:true }), run('Первое: любое число называется вместе со смыслом («0,01 — это литр воды на сто литров почвы»). Второе: доказательство показывается как рассуждение, с указанием, что именно оно доказывает и чего не доказывает. Третье: у каждой аналогии есть граница, которую можно назвать вслух (см. раздел 4).')]));
body.push(h2('Хронометраж'));
body.push(p(`Полный доклад — около ${Math.round(total / 140)} мин при спокойном темпе (140 слов в минуту), ${total} слов. Если регламент короче, пропустите слайды 6 и 16 (помечены «можно пропустить»): доклад сократится до ≈ ${Math.round(coreWords / 140)} мин. Если нужно ещё короче, объединяйте слайды 12 и 13 в один рассказ (парадокс «не значимо, но порядок устойчив») и сокращайте слайд 18 до двух строк таблицы.`));
const trows = [['№', 'Слайд', 'Время', 'Статус']].concat(slides.map(s => [String(s.n), s.title, '≈ ' + fmt(mins(s)) + ' мин', s.core ? 'основной' : 'можно пропустить']));
body.push(table([700, 5900, 1200, 1560], trows));

body.push(h1('2. Текст по слайдам'));
slides.forEach(s => {
  body.push(h2(`Слайд ${s.n}. ${s.title} (≈ ${fmt(mins(s))} мин)`));
  s.text.forEach(t => body.push(p(t)));
});

body.push(h1('3. Вероятные вопросы и ответы'));
body.push(p('Ответы даны развёрнуто: их можно произнести целиком или сократить до первого-двух предложений.', {}));
qa.forEach((q, i) => {
  body.push(new Paragraph({ keepNext:true, spacing:{ before:200, after:80 }, children:[new TextRun({ text:`Вопрос ${i + 1}. ${q[0]}`, font:FONT, size:24, bold:true })] }));
  body.push(p([run('Ответ. ', { italics:true, bold:true }), run(q[1])]));
});

body.push(h1('4. Аналогии и где они ломаются'));
body.push(p('Аналогии помогают войти в идею, но не заменяют доказательство. Если слушатель начнёт «доказывать» через аналогию, остановите его на границе, указанной в третьем столбце.'));
body.push(table([2300, 3300, 3760], [['Аналогия', 'Что она объясняет', 'Где перестаёт работать']].concat(analogies)));

body.push(h1('5. Примечания докладчику'));
[
  [run('Что добавлено в текст сверх статьи. ', { bold:true }), run('Аналогии (раздел 4), перевод влажности в миллиметры воды (0,01 = 10 мм в метровом слое = 100 м³/га: это пересчёт единиц), формула RMSE² = MB² + разброс² (математическое тождество) и «остаток после вычитания постоянной поправки +0,010» на слайде 17 (расчёт по табл. V статьи: 0,027 − 0,010 и т. д.). Всё остальное взято из итоговой статьи.')],
  [run('Состояния вне анализа (θ = 0,44 и 0,48; Δθ = +0,031 и +0,038). ', { bold:true }), run('В прежней версии презентации эти значения не использовались: они считались неподтверждёнными. В итоговой статье они приведены в аннотации, примечании к табл. V, рекомендациях и выводах, поэтому теперь показаны на слайде 11 пустыми маркерами и в тексте всегда названы «не вошедшими в анализ». Если авторы их не подтвердят, уберите соответствующие фразы в слайдах 11 и 17 (две строки в тексте).')],
  [run('Названия почв и титул. ', { bold:true }), run('Чернозём обыкновенный, чернозём типичный и лугово-чернозёмная почва взяты из статьи; подтипы и систему классификации авторы ещё должны подтвердить. На титульном слайде добавьте название и дату конференции, если требует оргкомитет.')],
  [run('Слабые места, которые лучше назвать самому. ', { bold:true }), run('Одна почва на группу; дисперсионный анализ и регрессии выполнены по сводным статистикам вариантов, а не по сырым колоночным значениям; граница между группами A и B чувствительна к округлению (p в 28 % и 85 % пересчётов для A1 и B1); коэффициент при плотности в регрессии устойчив, а его p-значение нет; влияние стенок колонки не измерялось. Раздел 3 содержит готовые ответы.')],
  [run('Слайды и редактирование. ', { bold:true }), run('Диаграммы в презентации — настоящие диаграммы PowerPoint (правятся через «Изменить данные»). Полный текст каждого слайда продублирован в заметках докладчика, с хронометражем.')]
].forEach(c => body.push(bullet(c)));

const doc = new Document({
  creator:'Кубанский ГАУ', title:'Текст доклада: CS616',
  styles:{ default:{ document:{ run:{ font:FONT, size:24 } } },
    paragraphStyles:[
      { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal', quickFormat:true, run:{ size:34, bold:true, font:HEAD }, paragraph:{ spacing:{ before:360, after:200 }, outlineLevel:0 } },
      { id:'Heading2', name:'Heading 2', basedOn:'Normal', next:'Normal', quickFormat:true, run:{ size:28, bold:true, font:HEAD }, paragraph:{ spacing:{ before:300, after:120 }, outlineLevel:1 } } ] },
  numbering:{ config:[{ reference:'bul', levels:[{ level:0, format:LevelFormat.BULLET, text:'•', alignment:AlignmentType.LEFT, style:{ paragraph:{ indent:{ left:540, hanging:270 } } } }] }] },
  sections:[{ properties:{ page:{ size:{ width:11906, height:16838 }, margin:{ top:1134, bottom:1134, left:1247, right:1247 } } },
    footers:{ default:new Footer({ children:[new Paragraph({ alignment:AlignmentType.CENTER, children:[new TextRun({ children:[PageNumber.CURRENT], font:FONT, size:20, color:'5B4636' })] })] }) },
    children: body }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2] || 'talk.docx', b); console.log('ok', total, 'words'); });
