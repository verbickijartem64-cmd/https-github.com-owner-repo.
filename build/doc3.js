const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  AlignmentType, Footer, PageNumber, LevelFormat, PageBreak } = require('docx');
const K = require('./content3.js');

const FONT = 'Calibri', HEAD = 'Cambria';
const COL = { dk:'2A1E16', teal:'1C7C8C', terra:'B8552F', muted:'5B4636', lt:'F1F3F2', ochre:'D9A441' };
const wc = s => s.split(/\s+/).filter(Boolean).length;
const WPM = 145;
const fullMin = s => wc(s.text.join(' ')) / WPM;
const shortMin = s => wc(s.short) / WPM;
const fmt = n => String(n).replace('.', ',');
const mm = x => fmt(Math.max(0.5, Math.round(x * 2) / 2));
const ss = x => `${Math.max(5, Math.round(x * 60 / 5) * 5)} с`;
const KIND = { core:'основной', deep:'углубление', optional:'можно пропустить' };

const run = (t, o = {}) => new TextRun(Object.assign({ text:t, font:FONT, size:23 }, o));
const p = (children, o = {}) => new Paragraph(Object.assign({ spacing:{ after:130, line:300 } }, o, { children: typeof children === 'string' ? [run(children)] : children }));
const h1 = t => new Paragraph({ heading:HeadingLevel.HEADING_1, pageBreakBefore:true, spacing:{ before:120, after:200 }, children:[new TextRun({ text:t, font:HEAD, size:34, bold:true, color:COL.dk })] });
const h2 = t => new Paragraph({ heading:HeadingLevel.HEADING_2, keepNext:true, spacing:{ before:280, after:100 }, children:[new TextRun({ text:t, font:HEAD, size:27, bold:true, color:COL.teal })] });
const h3 = t => new Paragraph({ keepNext:true, spacing:{ before:200, after:80 }, children:[run(t, { bold:true, size:24, color:COL.dk })] });
const bullet = children => new Paragraph({ numbering:{ reference:'bul', level:0 }, spacing:{ after:90, line:290 }, children: typeof children === 'string' ? [run(children)] : children });
const boxed = (label, text, fill) => new Paragraph({ spacing:{ before:60, after:160, line:290 }, shading:{ type:ShadingType.CLEAR, color:'auto', fill },
  border:{ left:{ style:BorderStyle.SINGLE, size:18, color:COL.teal, space:8 } }, indent:{ left:200, right:120 },
  children:[run(label, { bold:true, color:COL.teal }), run(text)] });

const bd = { style:BorderStyle.SINGLE, size:4, color:'C9D1CF' };
const borders = { top:bd, bottom:bd, left:bd, right:bd };
function cell(text, w, o = {}) {
  return new TableCell({ width:{ size:w, type:WidthType.DXA }, borders, margins:{ top:60, bottom:60, left:100, right:100 },
    shading: o.fill ? { fill:o.fill, type:ShadingType.CLEAR, color:'auto' } : undefined, verticalAlign:'center',
    children:[new Paragraph({ alignment:o.align, children:[new TextRun({ text, font:FONT, size:o.size || 20, bold:o.bold, color:o.color })] })] });
}
function table(cols, rows, opt = {}) {
  const W = cols.reduce((a, b) => a + b, 0);
  return new Table({ width:{ size:W, type:WidthType.DXA }, columnWidths:cols,
    rows: rows.map((r, i) => new TableRow({ tableHeader:i === 0, cantSplit:true,
      children: r.map((t, k) => cell(String(t), cols[k], i === 0 ? { fill:COL.dk, color:'FFFFFF', bold:true } : Object.assign({ fill: i % 2 ? COL.lt : undefined }, (opt.boldFirst && k === 0) ? { bold:true } : {}, opt.cellOpt ? opt.cellOpt(i, k, t) : {}))) })) });
}
const spacer = () => new Paragraph({ spacing:{ after:80 }, children:[] });

const S = K.slides;
const full = S.reduce((a, s) => a + fullMin(s), 0);
const main = S.reduce((a, s) => a + (s.kind === 'core' ? fullMin(s) : s.kind === 'deep' ? shortMin(s) : 0), 0);
const express = S.reduce((a, s) => a + (s.kind === 'optional' ? 0 : shortMin(s)), 0);

const body = [];
// ---------- title ----------
body.push(new Paragraph({ spacing:{ after:100 }, children:[new TextRun({ text:'Текст доклада', font:HEAD, size:52, bold:true, color:COL.dk })] }));
body.push(new Paragraph({ spacing:{ after:100 }, children:[new TextRun({ text:'Когда датчик влажности «льстит» почве', font:HEAD, size:32, bold:true, color:COL.teal })] }));
body.push(new Paragraph({ spacing:{ after:200 }, children:[new TextRun({ text:'Систематическая погрешность заводской калибровки датчиков влажности CS616 в незасолённых почвах разного гранулометрического состава', font:FONT, size:25, italics:true, color:COL.terra })] }));
body.push(p([run('Вербицкий А. Ю., Приходько И. А., Болдырева Л. М., Кесафоти Х. Е. · Кубанский государственный аграрный университет имени И. Т. Трубилина', { size:21, color:COL.muted })]));
body.push(p([run('Версия 3. ', { bold:true }), run('Доклад рассчитан на специалистов смежных областей: мелиораторов, агрономов, физиков, метрологов, статистиков. Каждый термин и каждое скрытое допущение объясняются при первом появлении — сначала на аналогии, затем строго, с числами из итоговой статьи. Отдельно показано, как устроены доказательства, на чём они держатся и насколько они твёрды.')]));

// ---------- 1. how to use ----------
body.push(new Paragraph({ heading:HeadingLevel.HEADING_1, spacing:{ before:300, after:200 }, children:[new TextRun({ text:'1. Как пользоваться этим документом', font:HEAD, size:34, bold:true, color:COL.dk })] }));
body.push(p('Презентация содержит 25 основных и 4 запасных слайда. У каждого основного слайда есть полный текст и короткая версия в одно-два предложения; оба текста есть и здесь, и в заметках докладчика в PowerPoint. Слайды трёх типов: основные, углубления (их можно рассказать коротко) и необязательные (их можно пропустить).'));
body.push(h2('Три маршрута под разный регламент'));
body.push(table([2000, 5400, 2000], [
  ['Маршрут', 'Что делать', 'Время речи'],
  ['Полный', 'все 25 слайдов, полный текст', `≈ ${Math.round(full)} мин`],
  ['Основной', 'основные слайды — полностью; углубления (7, 12, 18, 22) — короткая версия; слайды 8 и 20 пропустить', `≈ ${Math.round(main)} мин`],
  ['Экспресс', 'все слайды, кроме 8 и 20, — только короткая версия', `≈ ${Math.round(express)} мин`]
], { boldFirst:true }));
body.push(p([run('Время рассчитано по темпу 145 слов в минуту без пауз и переходов; на смену слайдов добавьте 3–5 секунд на каждый. Запасные слайды 26–29 показывают только при вопросах.', { size:20, color:COL.muted })]));
body.push(h2('Хронометраж по слайдам'));
body.push(table([600, 4700, 1700, 1250, 1150], [['№', 'Слайд', 'Тип', 'Полностью', 'Коротко']].concat(S.map(s => [String(s.n), s.title, KIND[s.kind], '≈ ' + mm(fullMin(s)) + ' мин', '≈ ' + ss(shortMin(s))]))));

// ---------- 2. one-page map ----------
body.push(h1('2. Карта доклада на одной странице'));
body.push(p('Главная мысль каждого слайда — то, что слушатель должен унести, даже если всё остальное забудет.'));
body.push(table([600, 1500, 7300], [['№', 'Часть', 'Главная мысль']].concat(S.map(s => [String(s.n), s.act || 'Титул', s.takeaway]))));

// ---------- 3. talk ----------
body.push(h1('3. Текст по слайдам'));
S.forEach(s => {
  body.push(h2(`Слайд ${s.n}. ${s.title}`));
  body.push(p([run(`≈ ${mm(fullMin(s))} мин · ${KIND[s.kind]}${s.act ? ' · часть «' + s.act + '»' : ''}`, { size:20, color:COL.muted })], { spacing:{ after:80 } }));
  body.push(p([run('Главная мысль: ', { bold:true, color:COL.terra }), run(s.takeaway, { bold:true })]));
  s.text.forEach(t => body.push(p(t)));
  body.push(boxed('Если времени мало: ', s.short, COL.lt));
});

// ---------- 4. logic of proof ----------
body.push(h1('4. Логика доказательства'));
body.push(p('Этот раздел — для подготовки к вопросам и для слушателей, которые хотят пройти доказательства на глубоком уровне. Он показывает, на чём держится каждый вывод и как его можно опровергнуть.'));
body.push(h2('4.1. Насколько твёрдо каждое утверждение'));
body.push(table([2300, 2300, 1400, 1050, 2350], [['Утверждение', 'Чем подтверждено', 'Тип связи', 'Твёрдость', 'Что мешает сказать сильнее']].concat(K.ledger),
  { boldFirst:true, cellOpt:(i, k, t) => k === 3 ? { bold:true, color: t === 'высокая' ? COL.teal : t === 'средняя' ? COL.muted : COL.terra } : {} }));
body.push(h2('4.2. Четыре типа связей между теориями'));
body.push(p('Работа соединяет шесть теорий: физику почв, коллоидную химию глин, теорию диэлектрических смесей, электродинамику, метрологию и статистику. В конце цепочки — решение мелиоратора. Для каждого вывода полезно понимать, какого рода связь он использует: от этого зависит, как вывод может оказаться ложным и как его проверить.'));
body.push(table([1900, 2600, 2500, 2400], [['Тип связи', 'Пример в работе', 'Как подводит', 'Как проверить']].concat(K.linkTypes), { boldFirst:true }));
body.push(h2('4.3. Почему «не значимо» не значит «эффекта нет»'));
body.push(p('Дисперсионный анализ отвечает на вопрос, выделяется ли разброс между группами на фоне разброса внутри групп. В нашем опыте разброс внутри вариантов на 90 % состоит из закономерного роста ошибки с влажностью, а не из случайного шума. Поэтому даже реальное различие почв, на которое приходится 9 % изменчивости, тонет в этом фоне: F(2; 54) = 2,74; p = 0,074.'));
body.push(p([run('Правильная формулировка: ', { bold:true }), run('при таком способе сравнения эффект почвы статистически не выделяется. '), run('Неправильная: ', { bold:true }), run('«почва не влияет». Отсутствие доказательства — не доказательство отсутствия. Кроме того, суммы квадратов восстановлены по сводным статистикам вариантов, поэтому диагностика остатков была невозможна, а p-значение чувствительно к округлению (в симуляциях 0,04–0,12).')]));
body.push(h2('4.4. Почему сравнение на равной влажности убедительнее'));
body.push(p('Блоковый план убирает из сравнения то, что общее для всех групп на данном уровне влажности: в каждом блоке сравниваются три группы при одной и той же влажности. Тогда рост ошибки с влажностью перестаёт быть «шумом», и остаётся вопрос, устойчив ли порядок групп от блока к блоку. На всех шести уровнях A < B ≤ C. Двухфакторный анализ без повторностей даёт F(2; 10) = 6,92; p = 0,013, ранговый критерий Фридмана — χ²(2) = 11,27; p = 0,004.'));
body.push(p('Ранговый критерий важен как проверка устойчивости: он не использует величины ошибок и не предполагает нормальности. Интуиция, которую можно дать залу: если бы почва не влияла и уровни были независимы, группа A оказалась бы наименьшей на всех шести уровнях с вероятностью (1/3)⁶ ≈ 1/729. Ограничение: уровни — последовательные состояния одних колонок, а каждая группа — одна почва. Поэтому вывод описательный и относится к изученным почвам.'));
body.push(h2('4.5. Как читать несбывшийся прогноз'));
body.push(p('Прогноз модели смеси (+0,0095) выводится не из одной модели, а из модели вместе со вспомогательными допущениями: та же влажность, та же проводимость, нет потерь сигнала. В философии науки это называют тезисом Дюэма — Куайна: опыт всегда проверяет связку, а не отдельную теорию. Логика простая: если из связки следует прогноз, а прогноз не подтвердился, то ложна хотя бы одна её часть.'));
body.push(p('В нашем опыте две части связки нарушены по данным: ECe у плотных вариантов ниже на 0,07–0,12 дСм/м, а диапазоны влажности вариантов различаются. Если в регрессии заменить гигроскопичность на ECe, коэффициент при плотности становится +0,0004 (p = 0,51). Поэтому корректный вывод — не «модель смеси неверна», а «в проводящих почвах её прогноз без учёта проводимости неприменим». Решающая проверка — уплотнение одной почвы при выровненной ECe.'));
body.push(h2('4.6. Как посчитать прогноз CRIM самостоятельно'));
[
  'Модель: √ε смеси = φтв·√4,7 + θ·√80 + φвозд·√1, где φ — доли объёма фаз. Смысл: время пробега складывается из времён через каждую фазу (скорость волны обратна √ε).',
  'Уплотнение с 1,35 до 1,52 г/см³ при той же влажности увеличивает долю твёрдой фазы на 0,17 / ρs; при ρs = 2,63 г/см³ — на ≈ 0,065, и ровно на столько же уменьшается доля воздуха.',
  'Каждая доля, перешедшая из воздуха в твёрдую фазу, увеличивает √ε смеси на √4,7 − 1 ≈ 1,17. Итого +0,065 × 1,17 ≈ +0,076.',
  'Заводская шкала читает любое увеличение √ε как воду, вытеснившую воздух; каждая доля воды даёт √80 − 1 ≈ 7,94. Кажущийся прирост влажности: 0,076 / 7,94 ≈ +0,0095 (при ρs 2,60–2,67 — от +0,0094 до +0,0096, как в статье).',
  'Наблюдалось: −0,001…−0,002 (95 % ДИ эффекта плотности −0,006…+0,003).'
].forEach(t => body.push(bullet(t)));

// ---------- 5. Q&A ----------
body.push(h1('5. Вероятные вопросы и ответы'));
body.push(p('Ответы даны развёрнуто: их можно произнести целиком или сократить до первого-двух предложений. Для вопросов о числах держите наготове запасные слайды 26–29.'));
K.qa.forEach((q, i) => {
  body.push(h3(`Вопрос ${i + 1}. ${q[0]}`));
  body.push(p([run('Ответ. ', { italics:true, bold:true }), run(q[1])]));
});

// ---------- 6. analogies ----------
body.push(h1('6. Аналогии и где они ломаются'));
body.push(p('Аналогия помогает войти в идею, но не заменяет доказательство. Если слушатель начинает «доказывать» через аналогию, остановите его на границе, указанной в третьем столбце.'));
body.push(table([2300, 3500, 3600], [['Аналогия', 'Что она объясняет', 'Где перестаёт работать']].concat(K.analogies), { boldFirst:true }));

// ---------- 7. glossary ----------
body.push(h1('7. Глоссарий для специалистов смежных областей'));
body.push(table([2500, 3700, 3200], [['Термин', 'Простыми словами', 'В нашей работе']].concat(K.glossary), { boldFirst:true }));

// ---------- 8. notes ----------
body.push(h1('8. Примечания докладчику'));
body.push(h2('Что в докладе добавлено сверх статьи'));
[
  'Аналогии (раздел 6) — авторские пояснения, не данные.',
  'Перевод влажности в миллиметры воды и кубометры на гектар — пересчёт единиц (0,01 м³/м³ = 10 мм в метровом слое = 100 м³/га).',
  'Тождество RMSE² = MB² + разброс² — математическое тождество для MB и RMSE (разброс со знаменателем n).',
  'Схемы на слайдах 3 и 6 — иллюстрации, а не измерения; на слайде 3 подписаны реальные значения табл. V для почвы C (эталон 0,40 → датчик 0,427; 0,30 → 0,312).',
  '«Облака» на слайде 15 и линии на слайде 16 построены по табл. V (6 уровней, плотность 1,35), а не по всем 60 значениям, по которым считался дисперсионный анализ.',
  'Интуиция «(1/3)⁶ ≈ 1 из 729» на слайде 16 — пояснение логики рангового теста при допущении независимости уровней, а не самостоятельный тест.',
  'Пошаговый расчёт прогноза CRIM и доли фаз при θ = 0,25 и ρs = 2,63 на слайде 18 — пересчёт по формуле статьи; итог (+0,0095) совпадает со статьёй (+0,0094…+0,0096).',
  'Остаток после вычитания постоянной поправки +0,010 на слайде 21 — арифметика по табл. V.',
  'Совет «проверять в двух точках влажности» — следствие из результата о росте ошибки с влажностью; в поле он не проверялся, и на слайде это сказано.',
  'Прогнозы «если — то» на слайде 23 сформулированы по обсуждению статьи; поправка по двум входам — гипотеза, объёмная электропроводность в опыте не измерялась.',
  'Шкала твёрдости выводов на слайде 24 — оценка по ограничениям, перечисленным в статье.'
].forEach(t => body.push(bullet(t)));
body.push(h2('Что нужно подтвердить у авторов'));
[
  'Состояния вне анализа (θ = 0,44 и 0,48; Δθ = +0,031 и +0,038) приведены в итоговой статье в аннотации, примечании к табл. V, рекомендациях и выводах, поэтому показаны на слайде 14 бледными точками и везде названы «не вошедшими в анализ». Если авторы их не подтвердят, уберите эти точки и фразы на слайдах 14 и 24, в таблице аналогий и в ответе на вопрос 10.',
  'Названия почв (чернозём обыкновенный, чернозём типичный, лугово-чернозёмная почва) взяты из статьи; подтипы и система классификации требуют подтверждения авторов.',
  'На титульный слайд добавьте название и дату конференции, если этого требует оргкомитет.'
].forEach(t => body.push(bullet(t)));
body.push(h2('Слабые места, которые лучше назвать самому'));
[
  'Одна почва на группу: выводы для классов почв — экстраполяция.',
  'Дисперсионный анализ и регрессии выполнены по сводным статистикам вариантов, а не по 180 колоночным значениям.',
  'Граница между группами A и B неустойчива к округлению: в симуляциях вывод для A1 менялся в 28 % случаев, для B1 — в 15 %.',
  'Коэффициент при плотности в регрессии устойчив (−0,0022…−0,0013), а его p-значение нет (p < 0,05 лишь в 76 % симуляций).',
  'Влияние стенок колонки и отбора микрокернов не измерялось; систематическая часть неопределённости эталона (0,005) соизмерима со сдвигом в лёгких почвах.'
].forEach(t => body.push(bullet(t)));
body.push(h2('Запасные слайды'));
K.backup.forEach((b, i) => body.push(bullet([run(`Слайд ${26 + i}. ${(t => t[0].toUpperCase() + t.slice(1))(b.title.replace('Запасной: ', ''))}. `, { bold:true }), run(b.note)])));

const doc = new Document({
  creator:'Кубанский ГАУ', title:'Текст доклада: датчики CS616', description:'Версия 3',
  styles:{ default:{ document:{ run:{ font:FONT, size:23 } } },
    paragraphStyles:[
      { id:'Heading1', name:'Heading 1', basedOn:'Normal', next:'Normal', quickFormat:true, run:{ size:34, bold:true, font:HEAD }, paragraph:{ spacing:{ before:240, after:200 }, outlineLevel:0 } },
      { id:'Heading2', name:'Heading 2', basedOn:'Normal', next:'Normal', quickFormat:true, run:{ size:27, bold:true, font:HEAD }, paragraph:{ spacing:{ before:280, after:100 }, outlineLevel:1 } } ] },
  numbering:{ config:[{ reference:'bul', levels:[{ level:0, format:LevelFormat.BULLET, text:'•', alignment:AlignmentType.LEFT, style:{ paragraph:{ indent:{ left:540, hanging:270 } } } }] }] },
  sections:[{ properties:{ page:{ size:{ width:11906, height:16838 }, margin:{ top:1134, bottom:1134, left:1247, right:1247 } } },
    footers:{ default:new Footer({ children:[new Paragraph({ alignment:AlignmentType.CENTER, children:[new TextRun({ children:[PageNumber.CURRENT], font:FONT, size:20, color:COL.muted })] })] }) },
    children: body }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2] || 'talk3.docx', b); console.log('ok; full', full.toFixed(1), 'main', main.toFixed(1), 'express', express.toFixed(1)); });
