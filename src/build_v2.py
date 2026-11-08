import json, re, sys, os
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

L_ALL = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'lessons.json')))
OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)
FONT = 'Times New Roman'
SZ = 12


def clean(s):
    s = re.sub(r'^(Мини-дәріс|Әңгіме)(\s*\([^)]*\))?:\s*', '', s)
    s = re.sub(r'^Қауіпсіздік сабағы:\s*', '', s)
    s = re.sub(r'^«[^»]*»(\s*\([^)]*\))?\.?\s*', '', s)
    return s


def sentences(s):
    parts = re.split(r'(?<=[.!?])\s+(?=[А-ЯӘҒҚҢӨҰҮҺІЁA-Z«0-9])', clean(s))
    return [p[0].upper() + p[1:] for p in parts if p.strip()]


def sub(t):
    p = t.split(' — ')
    return ' — '.join(p[1:]) if len(p) > 1 else t


def head(t):
    return t.split(' — ')[0]


def set_cell_width(cell, cm):
    cell.width = Cm(cm)
    tcPr = cell._tc.get_or_add_tcPr()
    w = tcPr.find(qn('w:tcW'))
    if w is None:
        w = OxmlElement('w:tcW'); tcPr.append(w)
    w.set(qn('w:w'), str(int(cm * 567))); w.set(qn('w:type'), 'dxa')



def grid(table, widths):
    table.autofit = False
    tbl = table._tbl
    g = tbl.tblGrid
    for i, gc in enumerate(g.findall(qn('w:gridCol'))):
        gc.set(qn('w:w'), str(int(widths[i] * 567)))
    tblPr = tbl.tblPr
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tblPr.append(lay)
    tw = tblPr.find(qn('w:tblW'))
    if tw is None:
        tw = OxmlElement('w:tblW'); tblPr.append(tw)
    tw.set(qn('w:w'), str(int(sum(widths) * 567))); tw.set(qn('w:type'), 'dxa')

def para(cell_or_doc, text, bold=False, italic=False, align=None, first=False, size=SZ):
    """text may contain **bold** spans."""
    if first and hasattr(cell_or_doc, 'paragraphs') and cell_or_doc.paragraphs and not cell_or_doc.paragraphs[0].text:
        p = cell_or_doc.paragraphs[0]
    else:
        p = cell_or_doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(0); pf.line_spacing = 1.0
    if align: p.alignment = align
    for i, chunk in enumerate(re.split(r'\*\*', text)):
        if not chunk: continue
        r = p.add_run(chunk)
        r.font.name = FONT; r.font.size = Pt(size)
        r._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
        r.bold = bold or (i % 2 == 1)
        r.italic = italic
    return p


def cellw(cell, lines, **kw):
    first = True
    for ln in lines:
        para(cell, ln, first=first, **kw); first = False


def build(L, doc):
    q = L['qz']
    # ---------- header table ----------
    t = doc.add_table(rows=0, cols=3)
    t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    W = [6.2, 5.6, 6.2]

    def row2(label, lines, bold_label=False, bold_val=False):
        r = t.add_row().cells
        a = r[0]; b = r[1].merge(r[2])
        cellw(a, [label], bold=bold_label)
        cellw(b, lines, bold=bold_val)
        return r

    row2('Сынып сағаты', [f'№{L["n"]}'])
    row2('Күні', [f'{L["date"]} ж. ({L["week"]})'])
    row2('Сынып жетекші', [''])
    r = t.add_row().cells
    cellw(r[0], ['Сынып: 6 «__» сынып']); cellw(r[1], ['Қатысушылар саны:']); cellw(r[2], ['Қатыспағандар саны:'])
    row2('Тәрбие сағатының тақырыбы', [L['t']], bold_label=True, bold_val=True)
    goals = [g.strip() for g in L['goal'].rstrip('.').split(';') if g.strip()]
    goals = [g[0].upper() + g[1:] + ';' for g in goals]
    goals.append(f'{L["v"]} қасиеттерін дамыту.')
    row2('Тәрбие сағатының мақсаты:', goals, bold_label=True)
    row2('Қауіпсіздік сабағының тақырыбы', [f'«{L["s"]}»'], bold_label=True)
    row2('Қауіпсіздік сабағының мақсаты',
         [f'Оқушыларды «{L["s"].split(" (")[0].rstrip("?")}» тақырыбы бойынша қауіпті жағдайларды тануға, қауіпсіздік ережелерін сақтауға және қажет кезде дұрыс көмекке жүгінуге үйрету'])
    row2('Құндылықтарды дарыту', [f'{L["v"]} айы'], bold_val=True)
    row2('Апта дәйексөзі', [L['q']])
    grid(t, W)
    for row in t.rows:
        tcs = row._tr.findall(qn('w:tc'))
        if len(tcs) == 2:
            set_cell_width(row.cells[0], W[0]); set_cell_width(row.cells[1], W[1] + W[2])
        else:
            for i in range(3): set_cell_width(row.cells[i], W[i])

    para(doc, '')
    para(doc, 'Сабақтың барысы:', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(doc, '')

    # ---------- process table ----------
    T = doc.add_table(rows=1, cols=4)
    T.style = 'Table Grid'; T.alignment = WD_TABLE_ALIGNMENT.CENTER
    CW = [2.8, 7.7, 4.5, 3.0]
    for i, h in enumerate(['Кезеңі', 'Мұғалім әрекеті', 'Оқушы әрекеті', 'Ресурстар']):
        cellw(T.rows[0].cells[i], [h], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)

    def stage(name, teacher, student, res):
        c = T.add_row().cells
        cellw(c[0], [name], bold=True)
        cellw(c[1], teacher); cellw(c[2], student); cellw(c[3], res)

    st = L['st']
    v = L['vid']
    # I
    stage('I. Кіріспе бөлім (10 мин)', [
        '– Оқушыларды жылы жүзбен сәлемдеседі. Психологиялық жағымды ахуал орнату мақсатында:',
        st[0],
        f'– Аптаның дәйексөзін бірге оқиды: **«{L["q"]}»** – «Бұл сөзді қалай түсінесіңдер?» деп сұрайды.',
        f'– {st[1]}',
        f'– {st[2]}',
        f'– Осылайша оқушылардың жауабы арқылы сабақтың тақырыбы анықталады: **«{L["t"]}»**.',
        '– Сұрақтар арқылы оқушылардың бастапқы білімін анықтайды:',
        *[f'{i+1}) «{x[0]}»' for i, x in enumerate(q[:3])],
    ], [
        '– Мұғаліммен бірге шеңбер құрып, жаттығуға белсенді қатысады, жылы тілектер айтады.',
        '– Дәйексөзді бірге оқып, мағынасын өз сөзімен түсіндіреді.',
        '– Бейнеүзіндіні қызығушылықпен тамашалап, әсерін бөліседі.',
        '– Проблемалық сұраққа болжам айтып, өз ойларын ортаға салады.',
        '– Бастапқы сұрақтарға жауап береді:',
        *[f'– «{x[1][x[2]]}»' for x in q[:3]],
        '– Сабақ тақырыбын өздері анықтап, дәптерлеріне жазып алады.',
    ], ['Интерактивті тақта, презентация,', f'бейнеролик «{v[0][0]}»', '«Адал азамат — 6-сынып» сайты'])

    # II
    b1, b2 = L['b1'], L['b2']
    teach2 = [f'**1. {head(b1[0])}: «{sub(b1[0]).strip("«»")}»**', '– Мұғалім түсіндіреді:']
    teach2 += ['– ' + s for s in sentences(b1[1])]
    teach2 += [f'**{b1[2]}** (топтық жұмыс):', b1[3], f'Дескриптор: {b1[4].replace("Дескрипторлар: ", "")}']
    teach2 += [f'**2. {head(b2[0])}: «{sub(b2[0]).strip("«»")}»**', '– Мұғалім әңгімелейді:']
    teach2 += ['– ' + s for s in sentences(b2[1])]
    teach2 += [f'**{b2[2]}** (жұптық жұмыс):', b2[3], f'Дескриптор: {b2[4].replace("Дескрипторлар: ", "")}']
    teach2 += [f'**Танымдық ойын «{L["gm"][0].strip("«»")}»:**', L['gm'][1],
               '– Әр дұрыс жауап үшін оқушыға жұлдызша беріледі. Соңында мұғалім жауаптарды қорытындылайды.']
    terms = ', '.join(m[0] for m in L['mt'])
    stud2 = [
        '– Мұғалімнің түсіндіруін зейінмен тыңдайды.',
        f'– Негізгі ұғымдарды дәптерге жазып алады: {terms}.',
        '– Өз пікірлерін білдіріп, сұрақтарға еркін жауап береді.',
        '– Топтарға бөлініп, 1-тапсырманы орындайды; ватманға жазып, сынып алдында қорғайды.',
        '– Жұппен 2-тапсырманы орындап, нәтижесін таныстырады.',
        '– Бір-бірінің жұмысын дескриптор бойынша бағалайды.',
        '– Ойынға белсене қатысады, командамен ақылдасып жауап береді:',
        *[f'– «{x[1][x[2]]}»' for x in q[3:6]],
    ]
    stage('II. Негізгі бөлім (20 мин)', teach2, stud2,
          [f'Видео «{v[1][0]}»,', 'ватман, маркер, стикерлер,', 'жағдай карточкалары,', 'бағалау парағы,', 'сайттағы ойындар'])

    # III
    b3 = L['b3']
    sname = L['s'].split(' (')[0]
    teach3 = [f'**Сабақтың мақсаты:** «{sname}» тақырыбы бойынша қауіпті жағдайларды тануға және дұрыс әрекет етуге үйрету.',
              '**Міндеті:** оқушыларды күнделікті өмірде қауіпсіздіктің негізгі қағидаларын қолдануға үйрету.',
              '**Сабақтың барысы:**',
              f'**Кіріспе.** Мұғалім сұрақ қояды: «{sname}» Оқушылардың өмірлік тәжірибесін тыңдап, ықтимал қауіптерді талқылайды.',
              '**Негізгі бөлім.**']
    teach3 += ['– ' + s for s in sentences(b3[0])]
    teach3 += [f'**{b3[1]}:**', b3[2], f'Дескриптор: {b3[3].replace("Дескрипторлар: ", "")}',
               '**Қорытынды.** Қауіпсіздік ережелерін бірге қайталап, жадынаманы бекітеді. «Екі жұлдыз, бір тілек» әдісімен бағалайды.']
    stage('III. Қауіпсіздік сабағы (10 мин)', teach3, [
        '– Мұғалімнің сұрағына жауап беріп, өз тәжірибесімен бөліседі.',
        '– Қауіпсіздік ережелерін тыңдап, жадынамаға жазады.',
        '– Топпен жағдайларды талдап, дұрыс әрекет алгоритмін құрады.',
        '– Бір жағдайды сахналап көрсетеді.',
        '– Ережелерді бірге қайталайды.',
    ], ['Инфографика,', f'бейнежазба «{v[2][0]}»,', 'плакат, жағдай карточкалары'])

    # IV
    facts = [x[0] for x in L['tf'] if x[1]]
    hw = L['hw']
    stage('IV. Қорытынды бөлім (5 мин)', [
        '– Сабақтың басты ойын бекітеді:',
        *[f'«{f}»' for f in facts[:3]],
        f'– Рефлексия **«{L["rf"][0].strip("«»")}»**: {L["rf"][1].replace(" Сайттағы кері байланыс толтырылады.", "")}',
        f'– **«{L["q"]}»** дәйексөзін бірлесіп оқиды.',
        '– **Үй тапсырмасы (таңдау бойынша):**',
        *hw,
        f'– **С деңгейлі оқушыға:** {L["c"]}',
        f'– **ЕБҚ оқушымен жұмыс:** {L["oop"]}',
    ], [
        '– Өз ойларын айтады, сабақтан алған әсерлерін бөліседі.',
        '– Рефлексия жасайды, стикерге жазады.',
        '– Дәйексөзді бірге оқиды.',
        '– Үй тапсырмасының деңгейін таңдап, күнделікке жазады.',
    ], ['Стикерлер, плакат,', 'дәйексөз парағы,', f'бейне «{v[3][0]}»'])
    grid(T, CW)
    for row in T.rows:
        for i, c in enumerate(row.cells):
            set_cell_width(c, CW[i])


def new_doc():
    d = Document()
    s = d.sections[0]
    s.page_width = Cm(21); s.page_height = Cm(29.7)
    s.left_margin = s.right_margin = Cm(1.5); s.top_margin = s.bottom_margin = Cm(1.5)
    st = d.styles['Normal']; st.font.name = FONT; st.font.size = Pt(SZ)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    # remove the default empty paragraph later by using it
    return d


allD = new_doc()
for i, L in enumerate(L_ALL):
    d = new_doc()
    build(L, d)
    d.save(os.path.join(OUT, f'QMZh2_{L["n"]:02d}.docx'))
    if i > 0:
        allD.add_page_break()
    build(L, allD)
allD.save(os.path.join(OUT, 'QMZh2_all.docx'))
print('ok')
