# -*- coding: utf-8 -*-
"""
КОМПАКТНАЯ версия для запоминания (背诵版) — закрытый экзамен.
Тема: 对外贸易立法管理.
Принцип: максимум структуры/чисел/мнемоник, минимум воды.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CN = "Microsoft YaHei"
RU = "Calibri"

C_TITLE = RGBColor(0x1F, 0x3A, 0x5F)
C_H1 = RGBColor(0x0B, 0x52, 0x94)
C_KEY = RGBColor(0xA1, 0x10, 0x10)      # красный — учить наизусть
C_MEM = RGBColor(0x7A, 0x10, 0x80)      # фиолетовый — мнемоника
C_GREY = RGBColor(0x66, 0x66, 0x66)


def shade(cell, hexc):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexc)
    tcPr.append(shd)


def sr(run, font=CN, size=10.5, bold=False, color=None, italic=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    rf.set(qn("w:eastAsia"), font)
    rf.set(qn("w:ascii"), font)
    rf.set(qn("w:hAnsi"), font)


def para(doc, runs, align=None, sa=2, sb=0, ind=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(sa)
    p.paragraph_format.space_before = Pt(sb)
    if ind is not None:
        p.paragraph_format.left_indent = Cm(ind)
    for text, opt in runs:
        sr(p.add_run(text), **opt)
    return p


def h1(doc, tag, cn):
    p = para(doc, [(f"{tag}　{cn}", dict(font=CN, size=15, bold=True, color=C_H1))], sb=8, sa=2)
    pPr = p._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "8")
    b.set(qn("w:space"), "1"); b.set(qn("w:color"), "0B5294")
    pb.append(b); pPr.append(pb)


def line(doc, cn, key=True, ind=0.5):
    """Основная строка для зубрёжки."""
    para(doc, [("▸ ", dict(font=CN, size=10.5, bold=True, color=C_KEY)),
               (cn, dict(font=CN, size=10.5, bold=key, color=C_KEY if key else RGBColor(0x22, 0x22, 0x22)))],
         sa=2, ind=ind)


def sub(doc, cn, ind=1.0):
    para(doc, [("· " + cn, dict(font=CN, size=10, color=RGBColor(0x33, 0x33, 0x33)))], sa=1, ind=ind)


def mem(doc, txt):
    """Мнемоника /口诀."""
    para(doc, [("🧠 口诀: ", dict(font=CN, size=10, bold=True, color=C_MEM)),
               (txt, dict(font=CN, size=10, bold=True, color=C_MEM))], sa=3, ind=0.5)


def table(doc, headers, rows, hdr_color="0B5294", widths=None, sizes=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    for c, h in zip(t.rows[0].cells, headers):
        shade(c, hdr_color)
        sr(c.paragraphs[0].add_run(h), font=RU, size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    for row in rows:
        cells = t.add_row().cells
        for i, (cell, val) in enumerate(zip(cells, row)):
            key = (i == 0)
            sz = sizes[i] if sizes else 9.5
            sr(cell.paragraphs[0].add_run(val), font=CN, size=sz, bold=key,
               color=C_KEY if key else RGBColor(0x22, 0x22, 0x22))
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t


# ===================================================================
doc = Document()
for s in doc.sections:
    s.left_margin = Cm(1.8); s.right_margin = Cm(1.8)
    s.top_margin = Cm(1.5); s.bottom_margin = Cm(1.5)
n = doc.styles["Normal"]; n.font.name = CN; n.font.size = Pt(10.5)
n.element.rPr.rFonts.set(qn("w:eastAsia"), CN)

para(doc, [("对外贸易立法管理 · 背诵版", dict(font=CN, size=20, bold=True, color=C_TITLE))],
     align=WD_ALIGN_PARAGRAPH.CENTER, sa=1)
para(doc, [("Сжатый конспект для закрытого экзамена · красное = учить наизусть · 🧠 = мнемоника",
            dict(font=RU, size=9.5, italic=True, color=C_GREY))],
     align=WD_ALIGN_PARAGRAPH.CENTER, sa=8)

# ---------------- 0. 调控手段 ----------------
h1(doc, "0", "对外贸易调控手段（3种）")
line(doc, "法律手段＝基础；经济手段＝为主；行政手段＝为辅。")
sub(doc, "法律：权威性·统一性·严肃性·规范性。")
sub(doc, "经济：汇率·税收·信贷·价格（间接）。")
sub(doc, "行政：命令·制度程序（直接）——登记·商检·配额。")
mem(doc, "「法基础、经为主、行为辅」")

# ---------------- 1. 概述 ----------------
h1(doc, "1", "法律调控的必要性（3条）")
line(doc, "①社会主义市场经济的客观要求（保障经济+行政手段）。")
line(doc, "②与国际贸易通行规则接轨（先是法律规则接轨）。")
line(doc, "③激烈竞争中保护国家和企业利益（应对反倾销等）。")
mem(doc, "「市场·接轨·保护」＝内·外·争")

h1(doc, "2", "外贸立法发展（4阶段）")
table(doc, ["阶段·年份", "关键词"],
      [("①1949—1977", "计划经济·仅货物·等级低·多内部文件"),
       ("②1978—1991", "改革开放·范围拓宽到技术/服务贸易"),
       ("③1992—2000", "市场经济·直接→间接调控·700多项·衔接国际"),
       ("④2001入世至今", "废改立·统一完备透明·清理+修改+提高透明度")],
      hdr_color="1F6F43", sizes=[10, 9.5])
mem(doc, "年份：49 / 78 / 92 / 2001入世")

h1(doc, "3", "立法体系的构成（渊源）")
line(doc, "国内法（按效力降序）：宪法＞法律＞行政法规＞地方性法规。")
sub(doc, "下位法不得与上位法相抵触。")
line(doc, "国际法：①国际条约（双边+多边）②国际贸易惯例。")
sub(doc, "已与140多国签双边条约；1971复联后参加100多个条约。")
sub(doc, "例：Incoterms / FOB（货物装船、风险随之转移）。")
mem(doc, "国内「宪—法—行—地」；国际「条约+惯例」")

# ---------------- 2. 外贸法 ----------------
h1(doc, "4", "《对外贸易法》关键数字")
line(doc, "1994.7.1 实施＝第一部基本法；2004.7.1 修订施行。")
line(doc, "结构：11章 70条；处于外贸立法体系核心地位。")
line(doc, "适用范围：地域(单独关税区不适用—港澳)·对人·时间。")
mem(doc, "「94首部、04修订、11章70条、港澳不适用」")

para(doc, [("六大基本原则：", dict(font=CN, size=10.5, bold=True, color=C_H1))], sb=4, sa=2)
table(doc, ["原则", "记忆点"],
      [("①统一制度", "首要原则·全国关境内统一制定+实施"),
       ("②鼓励发展", "对外开放国策·国民利益最大化"),
       ("③公平自由秩序", "国家统一管理下、法律范围内的公平自由"),
       ("④货物技术自由进出口", "必要限度内自由（不损国家安全/公共利益）"),
       ("⑤发展服务贸易", "依条约逐步·市场准入+国民待遇"),
       ("⑥平等互利·互惠对等", "互惠=互给优惠；对等=同等待遇/可报复")],
      hdr_color="1F6F43", sizes=[10, 9.5])
mem(doc, "「统一·鼓励·公平·自由·服务·平等」")

h1(doc, "5", "2004年修订（8点）")
line(doc, "1.允许自然人从事外贸。")
line(doc, "2.经营权：审批 → 登记备案。")
line(doc, "3.增加国营贸易管理。")
line(doc, "4.增加自动许可（仅备案、用于监测）。")
line(doc, "5.增加“与贸易有关的知识产权保护”（WTO三大支柱之一）。")
line(doc, "6.完善贸易救济：新增“对外贸易调查”章。")
line(doc, "7.完善法律责任（刑事/行政处罚、从业禁止）。")
line(doc, "8.新增外贸监测与公共服务（预警机制）。")
mem(doc, "重点记：自然人 + 审批→备案 + 知识产权 + 调查章")

# ---------------- 3. 货物 ----------------
h1(doc, "6", "货物贸易立法")
line(doc, "核心：《货物进出口管理条例》+配套规章（许可证·配额·自动许可·国营贸易·特殊货物）。")
line(doc, "主要环节：①进出口商品检验②海关③外汇管理。")
line(doc, "维护秩序「两反一保」(WTO允许、保护国内产业)：")
table(doc, ["条例", "章条", "针对"],
      [("反倾销条例", "6章59条", "低价倾销 → 反倾销税"),
       ("反补贴条例", "6章58条", "外国政府补贴"),
       ("保障措施条例", "5章34条", "进口数量激增")],
      hdr_color="1F6F43", sizes=[10, 9.5, 9.5])
mem(doc, "「反倾59、反补58、保障34」；2001.12颁布、2004修改")

# ---------------- 4. 技术 ----------------
h1(doc, "7", "技术贸易立法")
line(doc, "核心法规：《技术进出口管理条例》(2001.12)。")
line(doc, "技术分类→管理形式：")
sub(doc, "禁止/限制 → 目录管理；限制 → 许可证管理；自由 → 合同登记管理。")
line(doc, "技术进出口5形式：专利权转让·专利申请权转让·专利实施许可·技术秘密转让·技术服务等。")
line(doc, "知识产权三法：商标法·专利法·著作权法。")
mem(doc, "「目录(禁/限)·许可证(限)·登记(自由)」")

# ---------------- 5. 服务 ----------------
h1(doc, "8", "服务贸易立法")
line(doc, "趋势「三位一体」：货物+技术+服务。")
line(doc, "体系：以《外贸法》为最高层次基本法 + 行业性法律为主体。")
line(doc, "GATS 4种提供方式：跨境交付·境外消费·商业存在·自然人流动。")
sub(doc, "问题：我国主要只规范“商业存在”，其余三种规定少。")
mem(doc, "GATS四方式「跨境·境外·商业·自然人」")

# ---------------- 6. 跨境电商 ----------------
h1(doc, "9", "跨境电子商务（3文件+年份）")
line(doc, "2012：开始改革试点。")
line(doc, "2013：《支持跨境电商零售出口意见》(小批量·多频次；经营主体3类)。")
line(doc, "2016：《促进跨境电商健康快速发展指导意见》(首部指导性文件·11条)。")
line(doc, "零售进口：税收政策通知 + 进口商品清单。")
mem(doc, "「12试点、13出口、16指导意见(首部)」")

# ---------------- 速记总表 ----------------
h1(doc, "★", "终极速记（一眼扫）")
table(doc, ["考点", "答案"],
      [("三手段", "法(基础)·经(主)·行(辅)"),
       ("必要性3", "市场经济·国际接轨·竞争保护"),
       ("4阶段", "49-77 / 78-91 / 92-2000 / 2001入世"),
       ("国内法渊源", "宪法>法律>行政法规>地方性法规"),
       ("外贸法日期", "1994.7.1实施 / 2004.7.1修订"),
       ("外贸法结构", "11章70条·核心地位"),
       ("不适用地区", "单独关税区：香港、澳门"),
       ("六原则首要", "实行全国统一的对外贸易制度"),
       ("经营权改革", "审批 → 登记备案；准许自然人"),
       ("两反一保章条", "反倾59 / 反补58 / 保障34（01颁04改）"),
       ("技术管理", "目录(禁限)·许可证(限)·合同登记(自由)"),
       ("知识产权三法", "商标·专利·著作权"),
       ("GATS四方式", "跨境交付·境外消费·商业存在·自然人流动"),
       ("跨境电商", "2012试点·2013出口·2016指导意见")],
      hdr_color="A11010", sizes=[10, 9.5])

# ---------------- мини-словарь ----------------
h1(doc, "词", "核心词汇（中→俄）")
terms = [
    ("调控手段", "инструмент регулирования"),
    ("接轨", "состыковка с междунар. нормами"),
    ("法律渊源", "источники права"),
    ("行政法规", "административный регламент"),
    ("地方性法规", "местные нормативные акты"),
    ("相抵触", "противоречить вышестоящему акту"),
    ("国际惯例", "международный обычай"),
    ("关境", "таможенная территория"),
    ("单独关税区", "отдельная таможенная территория"),
    ("溯及力", "обратная сила закона"),
    ("审批", "разрешительный порядок"),
    ("登记备案", "регистрация и постановка на учёт"),
    ("国营贸易", "государственная торговля"),
    ("配额", "квота"),
    ("许可证", "лицензия"),
    ("知识产权", "интеллектуальная собственность"),
    ("最惠国待遇", "режим наибольшего благоприятствования"),
    ("国民待遇", "национальный режим"),
    ("互惠", "взаимные льготы"),
    ("对等", "паритет / взаимность (ответные меры)"),
    ("反倾销", "антидемпинг"),
    ("反补贴", "компенсационные меры"),
    ("保障措施", "защитные меры"),
    ("损害", "ущерб"),
    ("目录管理", "управление по каталогу"),
    ("合同登记", "регистрация контракта"),
    ("三位一体", "триединство (товары+техн.+услуги)"),
    ("跨境交付", "трансграничная поставка"),
    ("境外消费", "потребление за рубежом"),
    ("商业存在", "коммерческое присутствие"),
    ("自然人流动", "перемещение физлиц"),
    ("跨境电商", "трансгран. э-коммерция"),
    ("零售出口", "розничный экспорт"),
    ("出口退税", "возврат налога при экспорте"),
]
t = doc.add_table(rows=0, cols=4)
t.style = "Table Grid"
half = (len(terms) + 1) // 2
for i in range(half):
    cells = t.add_row().cells
    left = terms[i]
    sr(cells[0].paragraphs[0].add_run(left[0]), font=CN, size=9.5, bold=True, color=RGBColor(0x11, 0x11, 0x11))
    sr(cells[1].paragraphs[0].add_run(left[1]), font=RU, size=9, color=RGBColor(0x22, 0x22, 0x22))
    j = i + half
    if j < len(terms):
        right = terms[j]
        sr(cells[2].paragraphs[0].add_run(right[0]), font=CN, size=9.5, bold=True, color=RGBColor(0x11, 0x11, 0x11))
        sr(cells[3].paragraphs[0].add_run(right[1]), font=RU, size=9, color=RGBColor(0x22, 0x22, 0x22))

para(doc, [("加油！", dict(font=CN, size=12, bold=True, color=C_TITLE))],
     align=WD_ALIGN_PARAGRAPH.CENTER, sb=8)

out = "/workspace/对外贸易立法管理_背诵版.docx"
doc.save(out)
print("saved:", out)
