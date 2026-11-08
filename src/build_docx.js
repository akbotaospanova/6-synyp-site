// Usage: node build_docx.js lessons.json sizes.json outdir
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, BorderStyle, PageBreak, ExternalHyperlink, VerticalAlign, Footer,
} = require('docx');

const lessons = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const sizes = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
const outdir = process.argv[4];
fs.mkdirSync(outdir, { recursive: true });

const FONT = 'Times New Roman';
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 680; // A4, 1.2 cm
const TW = PAGE_W - 2 * MARGIN; // 10546
const BLUE = '0B5A7A', GOLD = 'C98A0B', LIGHT = 'E8F3F6', GOLDL = 'FBF1DC';
const border = { style: BorderStyle.SINGLE, size: 4, color: '7A9AA6' };
const borders = { top: border, bottom: border, left: border, right: border };
const yt = q => 'https://www.youtube.com/results?search_query=' + encodeURIComponent(q);

function mk(sz) {
  const half = Math.round(sz * 2);
  const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: half, ...o });
  const p = (children, o = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children], spacing: { before: 0, after: 20, line: 228 }, ...o });
  const txt = (t, o = {}) => p(run(t, o.run || {}), o.para || {});
  // label: value paragraph
  const lv = (label, value) => p([run(label, { bold: true }), run(value)]);
  const cell = (paras, w, o = {}) => new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    margins: { top: 30, bottom: 30, left: 70, right: 70 },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    verticalAlign: o.va || VerticalAlign.TOP,
    columnSpan: o.span,
    children: paras,
  });
  return { run, p, txt, lv, cell, half };
}

function lessonBlocks(L, sz) {
  const { run, p, txt, lv, cell } = mk(sz);
  const out = [];
  // Title
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
    children: [run(`Сынып сағатының қысқа мерзімді жоспары (ҚМЖ) №${L.n}`, { bold: true, color: BLUE, size: Math.round(sz * 2) + 6 })],
  }));
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD, space: 2 } },
    children: [run(`«Адал азамат» бірыңғай тәрбие бағдарламасы · ${L.qr} тоқсан · ${L.m} — ${L.v.toLowerCase()} айы`, { italics: true, color: '555555' })],
  }));

  const C1 = 2300, C2 = TW - C1;
  const row = (label, paras) => new TableRow({ cantSplit: false, children: [cell([txt(label, { run: { bold: true, color: BLUE } })], C1, { fill: LIGHT }), cell(paras, C2)] });
  const b1lab = L.b1[0], b2lab = L.b2[0];
  out.push(new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: [C1, C2],
    rows: [
      row('Сынып сағаты', [txt(`6-сынып · Күні: ${L.date} (${L.week}) · Айдың құндылығы: ${L.v.toLowerCase()}`)]),
      row('Сынып жетекшісі', [txt('______________________   Қатысқаны: ____   Қатыспағаны: ____')]),
      row('Негізгі тақырып', [txt(L.t.includes('«') ? L.t : `«${L.t}»`, { run: { bold: true } })]),
      row('Құндылықтарды жүйелі кіріктіру', [
        txt('• ' + b1lab), txt('• ' + b2lab), txt(`• Жеке қауіпсіздік — «${L.s}»`),
        p([run('Қосымша түсіндірме жұмыстары: ', { italics: true }), run(L.x.join('; ') + '.')]),
      ]),
      row('Мақсаты', [txt(L.goal)]),
      row('Күтілетін нәтижелер', [lv('Біледі: ', L.kn), lv('Түсінеді: ', L.un), lv('Қолданады: ', L.ap)]),
      row('Әдіс-тәсілдер', [txt(L.meth)]),
      row('Аптаның дәйексөзі (эпиграф)', [txt(`«${L.q}»`, { run: { italics: true } })]),
    ],
  }));

  out.push(new Paragraph({ spacing: { before: 60, after: 30 }, children: [run('Сынып сағатының барысы', { bold: true, color: BLUE, size: Math.round(sz * 2) + 2 })] }));

  const W = [1400, 4850, 1600, 1450, TW - 1400 - 4850 - 1600 - 1450];
  const H = ['Кезең, уақыты', 'Педагогтің әрекеті', 'Оқушының әрекеті', 'Бағалау', 'Ресурстар'];
  const head = new TableRow({ tableHeader: true, children: H.map((h, i) => cell([txt(h, { run: { bold: true, color: 'FFFFFF' }, para: { alignment: AlignmentType.CENTER } })], W[i], { fill: BLUE, va: VerticalAlign.CENTER })) });
  const short = s => s.split(' — ')[0].replace(/^Конституция және /, 'Конституция, ');
  const r = (stage, teacher, student, assess, res) => new TableRow({ children: [
    cell([txt(stage, { run: { bold: true } })], W[0], { fill: GOLDL }),
    cell(teacher, W[1]), cell([txt(student)], W[2]), cell([txt(assess)], W[3]), cell([txt(res)], W[4]),
  ] });
  const taskP = (name, body) => p([run(name + ': ', { bold: true, color: GOLD }), run(body)]);
  const rows = [head,
    r('Басы\n(5 мин)'.replace('\n', ' '), L.st.map((s, i) => txt(`${i + 1}. ${s}`)), 'Жаттығуға қатысады, пікір айтады, болжам жасайды.', 'Ауызша мадақтау', 'Сайт: «Кіріспе», бейнеүзінді'),
    r(`${short(b1lab)} (9 мин)`, [txt(L.b1[1]), taskP(L.b1[2], L.b1[3])], 'Тыңдайды, Конституция мәтінімен жұмыс істейді, топта жоба құрып, 1 минут қорғайды.', L.b1[4], 'Сайт: «Тақырып», «3 тапсырма»; ватман, маркер'),
    r(`${short(b2lab)} (8 мин)`, [txt(L.b2[1]), taskP(L.b2[2], L.b2[3])], 'Жұппен талдайды, шешім ұсынады, нәтижені таныстырады.', L.b2[4], 'Сайт: «3 тапсырма»; карточкалар'),
    r('Қауіпсіздік сабағы (10 мин)', [txt(L.b3[0]), taskP(L.b3[1], L.b3[2])], 'Жағдайларды талдайды, алгоритм құрады, сахналайды.', L.b3[3] + ' «Екі жұлдыз, бір тілек».', 'Сайт: «Қауіпсіздік», «3 тапсырма»; жағдай карточкалары'),
    r('Танымдық ойын (7 мин)', [taskP(L.gm[0], L.gm[1])], 'Командамен сұрақ таңдайды, ақылдасады, жауап береді; телефон/интерактивті тақтада ойнайды.', 'Командалар ұпайы, онлайн ойын нәтижесі', 'Сайт: «Ойындар» (викторина, шын/жалған, сәйкестендір, әріп шашу, дөңгелек)'),
    r('Соңы (6 мин)', [
      taskP('Рефлексия «' + L.rf[0].replace(/[«»]/g, '') + '»', L.rf[1]),
      p([run('Үй тапсырмасы (таңдау бойынша): ', { bold: true }), run(L.hw.join(' '))]),
      txt('Педагог қорытындысы: аптаның дәйексөзіне оралып, тақырыптың «Адал азамат» құндылығымен байланысын түйіндейді.'),
    ], 'Рефлексия жасайды, кері байланыс береді, үй тапсырмасын таңдайды.', 'Өзін-өзі бағалау', 'Сайт: «Кері байланыс», «Үй тапсырмасы»'),
  ];
  out.push(new Table({ width: { size: TW, type: WidthType.DXA }, columnWidths: W, rows }));

  out.push(new Paragraph({ spacing: { before: 40, after: 0 }, children: [] }));
  const vids = L.vid.map((v, i) => [new ExternalHyperlink({ link: yt(v[1]), children: [run(v[0], { color: '1155CC', underline: {} })] }), run(i < L.vid.length - 1 ? '; ' : '.')]).flat();
  out.push(new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: [C1, C2],
    rows: [
      row('С деңгейлі оқушыға тапсырма', [txt(L.c)]),
      row('ЕБҚ оқушымен жұмыс', [txt(L.oop)]),
      row('Пәнаралық байланыс', [txt(L.mp)]),
      row('Бейнематериалдар', [p(vids)]),
      row('Сынып жетекшісінің рефлексиясы', [txt('Мақсатқа жетті ме? ____________  Не сәтті болды? ____________  Нені өзгерту керек? ____________')]),
    ],
  }));
  return out;
}

function doc(children, sz) {
  const { run } = mk(sz);
  return new Document({
    styles: { default: { document: { run: { font: FONT, size: Math.round(sz * 2) } } } },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, footer: 300 } } },
      children,
    }],
  });
}

(async () => {
  const all = [];
  for (const L of lessons) {
    const sz = sizes[String(L.n)] || 10;
    const blocks = lessonBlocks(L, sz);
    const nn = String(L.n).padStart(2, '0');
    const buf = await Packer.toBuffer(doc(blocks, sz));
    fs.writeFileSync(path.join(outdir, `QMZh_${nn}.docx`), buf);
    if (all.length) all.push(new Paragraph({ children: [new PageBreak()] }));
    all.push(...lessonBlocks(L, sz));
  }
  if (process.argv[5] === 'all') {
    const buf = await Packer.toBuffer(doc(all, 10));
    fs.writeFileSync(path.join(outdir, 'QMZh_6-synyp_34_synyp_sagaty.docx'), buf);
  }
  console.log('done');
})();
