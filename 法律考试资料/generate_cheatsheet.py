# -*- coding: utf-8 -*-
# Generate a SIMPLE bilingual (Chinese + Russian) agency-law exam cheat sheet.

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


CN_FONT = '宋体'
CN_FONT_H = '黑体'
RU_FONT = 'Times New Roman'

doc = Document()

normal = doc.styles['Normal']
normal.font.name = CN_FONT
normal.font.size = Pt(11)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), CN_FONT)

RED = RGBColor(0xC0, 0x39, 0x2B)
BLUE = RGBColor(0x1F, 0x3A, 0x6E)
GRAY = RGBColor(0x55, 0x55, 0x55)
RU_COLOR = RGBColor(0x2E, 0x5A, 0x88)


def set_cn(run, font=CN_FONT):
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)


def set_ru(run):
    # Cyrillic/Latin font, keep eastAsia for any stray CJK
    run.font.name = RU_FONT
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), CN_FONT)
    rFonts.set(qn('w:ascii'), RU_FONT)
    rFonts.set(qn('w:hAnsi'), RU_FONT)


def shade_cell(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_cn(r, CN_FONT_H)
    r.font.size = Pt(19)
    r.font.bold = True
    r.font.color.rgb = BLUE


def subtitle(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_ru(r)
    r.font.size = Pt(11)
    r.font.color.rgb = GRAY


def h1(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_cn(r, CN_FONT_H)
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), '1F3A6E')
    pPr.append(shd)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)


def cn_para(text, bold=False, color=None, bullet=False):
    p = doc.add_paragraph(style='List Bullet' if bullet else None)
    r = p.add_run(text)
    set_cn(r)
    r.font.size = Pt(11)
    r.font.bold = bold
    if color:
        r.font.color.rgb = color
    p.paragraph_format.space_after = Pt(1)
    return p


def ru_para(text, bullet=False):
    # Russian explanation line, slightly indented + colored italic.
    p = doc.add_paragraph(style='List Bullet' if bullet else None)
    r = p.add_run('RU: ' + text)
    set_ru(r)
    r.font.size = Pt(10)
    r.font.italic = True
    r.font.color.rgb = RU_COLOR
    p.paragraph_format.space_after = Pt(5)
    if not bullet:
        p.paragraph_format.left_indent = Cm(0.5)
    return p


def make_table(headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    for i, htext in enumerate(headers):
        hdr[i].text = ''
        rr = hdr[i].paragraphs[0].add_run(htext)
        set_cn(rr, CN_FONT_H)
        rr.font.size = Pt(10)
        rr.font.bold = True
        rr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade_cell(hdr[i], '1F3A6E')
    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        fill = 'EEF2F8' if ri % 2 == 0 else 'FFFFFF'
        for i, ctext in enumerate(row):
            cells[i].text = ''
            for j, line in enumerate(str(ctext).split('\n')):
                pp = cells[i].paragraphs[0] if j == 0 else cells[i].add_paragraph()
                rr = pp.add_run(line)
                # Russian lines in tables flagged with prefix '~'
                if line.startswith('~'):
                    rr.text = line[1:]
                    set_ru(rr)
                    rr.font.size = Pt(9)
                    rr.font.italic = True
                    rr.font.color.rgb = RU_COLOR
                else:
                    set_cn(rr)
                    rr.font.size = Pt(9.5)
            shade_cell(cells[i], fill)
    if widths:
        for i, w in enumerate(widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)
    return table


LQ = '\u300c'
RQ = '\u300d'

# ========================== 封面 ==========================
title('代理法 · 考试速记（简版）')
subtitle('Шпаргалка по праву представительства (代理) — кратко, с пояснениями на русском')

# ========================== 1. 什么是代理 ==========================
h1('一、什么是代理？')
cn_para('代理 = 代理人按本人的【授权】，以【本人名义】与第三人做法律行为，后果【直接归本人】。', bold=True)
ru_para('Представительство (代理) — представитель (代理人) по полномочию (授权) принципала '
        '(本人) совершает сделку с третьим лицом ОТ ИМЕНИ принципала; права и обязанности '
        'возникают напрямую у принципала.')
cn_para('三方结构：本人 — 代理人 — 第三人（相对人）。', bold=True, color=RED)
ru_para('Всегда ТРИ стороны: принципал — представитель — третье лицо. Если сторон только две — это не представительство.')
cn_para('四个要件：①有代理权 ②以本人名义 ③在权限内 ④实施法律行为。', color=BLUE)
ru_para('Четыре условия: 1) есть полномочие; 2) от имени принципала; 3) в пределах полномочий; 4) совершается сделка.')

# ========================== 2. 三个易混概念 ==========================
h1('二、三个容易混的概念（重点区分）')
make_table(
    ['', '代理人 Представитель', '使者 Посланник', '冒名 Выдача себя за другого'],
    [
        ['自己拿主意？', '是，独立决定内容\n~Да, решает сам', '否，只传话\n~Нет, только передаёт', '冒用别人名字\n~Использует чужое имя'],
        ['要行为能力？', '要\n~Нужна дееспособность', '不要\n~Не нужна', '—'],
        ['几方关系？', '三方\n~Три стороны', '本质两方\n~По сути две', '两方\n~Две стороны — НЕ представительство'],
    ],
    widths=[3.0, 4.2, 4.2, 4.6],
)
cn_para('口诀：使者「传话」、代理「做主」、冒名「顶替」。', bold=True, color=RED)
ru_para('Запомни: посланник ПЕРЕДАЁТ, представитель РЕШАЕТ, самозванец ПОДМЕНЯЕТ.')

# ========================== 3. 判断步骤 ==========================
h1('三、判断三步法：是不是代理？什么代理？')
cn_para('① 有没有三方结构？没有 → 冒名行为，不是代理。', bullet=True)
ru_para('Шаг 1: Есть ли три стороны? Нет → это самозванство, не представительство.', bullet=True)
cn_para('② 以谁的名义？以自己名义 → 间接代理；以本人名义 → 直接代理。', bullet=True)
ru_para('Шаг 2: От чьего имени? От своего → косвенное; от имени принципала → прямое.', bullet=True)
cn_para('③ 有没有代理权？没有 → 看是「无权代理」还是「表见代理」。', bullet=True)
ru_para('Шаг 3: Есть ли полномочие? Нет → проверяем: представительство без полномочий или видимое.', bullet=True)

# ========================== 4. 代理权类型 ==========================
h1('四、代理权的类型')
make_table(
    ['类型', '关键词 / 怎么认'],
    [
        ['法定代理', '依法律产生（如监护人）\n~Законное: по закону (опекун и т.п.)'],
        ['意定（委托）代理', '本人授权产生，口头或书面\n~По воле: по доверенности/поручению'],
        ['明示授权 express', '明确委托（口头或书面）\n~Прямое: явно поручил'],
        ['默示授权 implied', '没明说，但本人习惯默许其代理\n~Подразумеваемое: молчаливо допускал'],
        ['表见/不容否认 estoppel', '无权，但让第三人有理由相信有权，本人知情不否认\n~Видимое: третье лицо обоснованно верит в полномочие'],
        ['客观必需 of necessity', '紧急情况、为本人利益、联系不上本人\n~По необходимости: экстренная ситуация'],
        ['追认 ratification', '本来无权，本人事后承认\n~Одобрение: принципал одобрил постфактум'],
    ],
    widths=[5.0, 11.0],
)

# ========================== 5. 无权代理 vs 表见代理 ==========================
h1('五、无权代理 vs 表见代理（最常考）')
make_table(
    ['', '无权代理 Без полномочий', '表见代理 Видимое'],
    [
        ['是什么', '没有代理权就代理\n~Действует без полномочий', '没权，但第三人有理由相信他有权\n~Нет полномочий, но видимость есть'],
        ['效力', '效力待定，本人追认才有效\n~Действует только после одобрения', '直接有效，约束本人\n~Сразу действительно для принципала'],
        ['保护谁', '保护本人\n~Защищает принципала', '保护善意第三人、交易安全\n~Защищает добросовестное третье лицо'],
        ['法条', '民法典171条', '民法典172条'],
    ],
    widths=[2.6, 6.7, 6.7],
)
cn_para('★ 伪造公章、假冒名义（如老干妈案）→ 不可归责于本人 → 不构成表见代理。', bold=True, color=RED)
ru_para('Важно: подделка печати/подписи (дело «Лаоганьма») → вину нельзя возложить на принципала → видимого представительства НЕТ.')

# ========================== 6. 直接 vs 间接代理 ==========================
h1('六、直接代理 vs 间接代理')
cn_para('直接代理：以本人名义 → 合同直接约束本人。', bullet=True)
ru_para('Прямое: от имени принципала → договор связывает принципала напрямую.', bullet=True)
cn_para('间接代理：以代理人自己名义。第三人是否知道代理关系是关键：', bullet=True)
ru_para('Косвенное: от имени представителя. Ключ — знало ли третье лицо о представительстве:', bullet=True)
cn_para('　— 第三人知道（民法典925条）→ 合同直接约束本人和第三人。', color=BLUE)
ru_para('Знало (ст. 925) → договор напрямую связывает принципала и третье лицо.')
cn_para('　— 第三人不知道（民法典926条）→ 先约束代理人；后可披露：本人有介入权 / 第三人有选择权。', color=BLUE)
ru_para('Не знало (ст. 926) → сначала связывает представителя; затем раскрытие: у принципала право вступить, у третьего лица право выбора.')

# ========================== 7. 终止与义务 ==========================
h1('七、终止 & 主要义务')
cn_para('终止：期限届满 / 双方同意 / 本人撤回；本人或代理人死亡、破产、丧失行为能力。', bullet=True)
ru_para('Прекращение: истёк срок / соглашение / отзыв; смерть, банкротство, недееспособность стороны.', bullet=True)
cn_para('对外终止须通知第三人，否则不能对抗第三人。', bullet=True)
ru_para('Прекращение действует против третьих лиц только после уведомления.', bullet=True)
cn_para('代理人义务：勤勉、忠实（不得自己代理/双方代理）、保密、报账、原则上不得转委托。', bullet=True)
ru_para('Обязанности представителя: добросовестность, лояльность (нельзя сделка с самим собой), '
        'тайна, отчётность, обычно нельзя передоверие.', bullet=True)
cn_para('本人义务：付报酬、偿还费用、让代理人查账。', bullet=True)
ru_para('Обязанности принципала: вознаграждение, возмещение расходов, доступ к счетам.', bullet=True)

# ========================== 8. 答题模板 ==========================
h1('八、答题模板（直接套用）')
cn_para('模板A — 是否构成代理：', bold=True, color=RED)
cn_para('① 三方结构？② 以本人名义？③ 有无代理权？④ 结论：是/否代理人。')
ru_para('Шаблон A — есть ли представительство: 1) три стороны? 2) от имени принципала? 3) есть полномочие? 4) вывод.')
cn_para('模板B — 代理类型：', bold=True, color=RED)
cn_para('先定法系 → 大陆法（法定/意定）或英美法（明示/默示/表见/必需/追认）→ 指出关键事实 → 定类型。')
ru_para('Шаблон B — тип: сначала правовая система, затем сопоставить с типами по ключевому факту.')
cn_para('模板C — 无权/表见代理：', bold=True, color=RED)
cn_para('① 确认无代理权 → ② 检验表见四要件（权利外观？善意无过失？可归责本人？）→ ③ 结论与后果。')
ru_para('Шаблон C: 1) нет полномочия → 2) проверить 4 условия видимого представительства → 3) вывод и последствия.')
cn_para('模板D — 间接代理：', bold=True, color=RED)
cn_para('① 以自己名义订约 → ② 第三人是否知情（925/926条）→ ③ 介入权/选择权 → ④ 结论。')
ru_para('Шаблон D: 1) от своего имени → 2) знало ли третье лицо (ст.925/926) → 3) право вступить/выбора → 4) вывод.')

# ========================== 9. 典型案例答案 ==========================
h1('九、典型案例·标准答案')
cases = [
    ('外运公司代办出口·付费是否合理？',
     '不合理。货代装船即完成职责，运费应由货主付；外运公司无付款义务。',
     'Необоснованно. Экспедитор выполнил долг, погрузив товар; платить должен грузовладелец.'),
    ('皮货案·R是否有代理权？',
     '皮货非易腐物，无客观必需代理权，R越权出售应赔偿。',
     'Меха не скоропортящиеся → нет права по необходимости → R отвечает за превышение.'),
    ('老干妈案·是否成立代理/付费？',
     '伪造公章冒充经理，不可归责老干妈 → 非表见代理，属无权代理；协议不生效，老干妈不付费。',
     'Подделка печати → не видимое, а без полномочий; договор не действует, платить не нужно.'),
    ('长期签单餐费·辞职后被拒付？',
     '经理长期签单、公司结清 → 表见代理，公司应付款，事后可向其追偿。',
     'Видимое представительство → компания платит, потом может взыскать с работника.'),
    ('委托卖房·代理人自己买',
     '自己代理，滥用代理权；本人不追认 → 无效，应退房。',
     'Сделка с самим собой → недействительна без одобрения → вернуть квартиру.'),
    ('生鲜+转委托C公司·C有无代理权？',
     '紧急且联系不上本人，为本人利益必须转委托 → C具备代理权。',
     'Экстренно и нельзя связаться → передоверие допустимо → у C есть полномочие.'),
    ('机械设备·乙名义、丙不知情、乙因甲未付款',
     '依926条乙披露后，丙可选甲或乙主张。答案：C（可选甲或乙）。',
     'Ст.926: после раскрытия третье лицо выбирает принципала или представителя. Ответ: C.'),
]
for cn_q, cn_a, ru_a in cases:
    p = doc.add_paragraph()
    r = p.add_run('▸ ' + cn_q)
    set_cn(r, CN_FONT_H)
    r.font.size = Pt(10.5)
    r.font.bold = True
    r.font.color.rgb = BLUE
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(1)
    pa = doc.add_paragraph()
    r1 = pa.add_run('答：')
    set_cn(r1)
    r1.font.bold = True
    r1.font.color.rgb = RED
    r1.font.size = Pt(10.5)
    r2 = pa.add_run(cn_a)
    set_cn(r2)
    r2.font.size = Pt(10.5)
    pa.paragraph_format.space_after = Pt(1)
    pr = doc.add_paragraph()
    r3 = pr.add_run('RU: ' + ru_a)
    set_ru(r3)
    r3.font.size = Pt(9.5)
    r3.font.italic = True
    r3.font.color.rgb = RU_COLOR
    pr.paragraph_format.space_after = Pt(4)

doc.save('/workspace/法律考试资料/代理法_考试速记_中俄简版.docx')
print('OK saved')
