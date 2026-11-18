# -*- coding: utf-8 -*-
"""
按学校模板（04_实验报告一模板.docx）生成实验二报告。
作者：庞华杰 32406300076
"""
import os
import io
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = r"D:\OO面向对象实验"
TPL = os.path.join(BASE, "04_实验报告一模板.docx")
DIR = os.path.join(BASE, "实验二")
IMG = os.path.join(DIR, "img")
OUT = os.path.join(DIR, "实验二报告_32406300076.docx")

OUT_W = 14.8   # 正文可用宽度(cm)：21 - 3.54 - 2.54

F_SONG = "宋体"
F_HEI = "黑体"
F_ROMAN = "Times New Roman"
F_MONO = "Consolas"

headings = []
_h = [0]
_fig = [0]
_tab = [0]

doc = Document(TPL)


# ---------------------------------------------------------------- 基础
def _f(run, latin=F_ROMAN, ea=F_SONG, size=10.5, bold=False, color=None):
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for k in ("w:ascii", "w:hAnsi", "w:cs"):
        rFonts.set(qn(k), latin)
    rFonts.set(qn("w:eastAsia"), ea)


def shade_para(para, fill):
    pPr = para._element.get_or_add_pPr()
    pPr.append(_shd(fill))


def _shd(fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    return shd


def cell_shade(cell, fill):
    cell._tc.get_or_add_tcPr().append(_shd(fill))


def p(text="", indent=0.74, size=10.5, bold=False, align="justify", after=4,
      line=1.45, ea=F_SONG, latin=F_ROMAN, color=None, before=0):
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Cm(indent) if indent else Cm(0)
    para.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                      "center": WD_ALIGN_PARAGRAPH.CENTER,
                      "left": WD_ALIGN_PARAGRAPH.LEFT,
                      "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    if text:
        _f(para.add_run(text), latin=latin, ea=ea, size=size, bold=bold, color=color)
    return para


def bullet(text, level=0, size=10.5):
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.line_spacing = 1.45
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.left_indent = Cm(0.74 + 0.6 * level)
    pf.first_line_indent = Cm(-0.37)
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _f(para.add_run(("• " if level == 0 else "– ") + text), size=size)
    return para


def step(text, no, size=10.5):
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.line_spacing = 1.45
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    pf.left_indent = Cm(1.1)
    pf.first_line_indent = Cm(-0.35)
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _f(para.add_run("%s " % no), size=size)
    _f(para.add_run(text), size=size)
    return para


def h1(text):
    _h[0] += 1
    para = doc.add_paragraph(style="知识点")
    pf = para.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(10)
    pf.space_after = Pt(4)
    pf.line_spacing = 1.3
    pf.keep_with_next = True
    _f(para.add_run(text), ea=F_HEI, latin=F_HEI, size=12, bold=True)
    return para


def h2(text):
    _h[0] += 1
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(7)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.3
    pf.keep_with_next = True
    _f(para.add_run(text), ea=F_HEI, latin=F_HEI, size=10.5, bold=True)
    return para


def table(headers, rows, widths, caption, size=8.5):
    _tab[0] += 1
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.first_line_indent = Cm(0)
    cap.paragraph_format.space_before = Pt(6)
    cap.paragraph_format.space_after = Pt(2)
    cap.paragraph_format.keep_with_next = True
    _f(cap.add_run("表 %d  %s" % (_tab[0], caption)), size=9.5, bold=True)

    k = OUT_W / float(sum(widths))
    widths = [w * k for w in widths]

    tb = doc.add_table(rows=1, cols=len(headers))
    tb.style = "Table Grid"
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = False
    tblPr = tb._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(sum(widths) * 567)))
    tblPr.append(tblW)
    for i, w in enumerate(widths):
        tb.columns[i].width = Cm(w)

    hdr = tb.rows[0]
    for i, htxt in enumerate(headers):
        c = hdr.cells[i]
        c.width = Cm(widths[i])
        para = c.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.first_line_indent = Cm(0)
        para.paragraph_format.space_before = Pt(1)
        para.paragraph_format.space_after = Pt(1)
        para.paragraph_format.line_spacing = 1.0
        para.paragraph_format.keep_with_next = True
        _f(para.add_run(htxt), ea=F_HEI, latin=F_HEI, size=size, bold=True)
        cell_shade(c, "DCE6F2")
    trPr = hdr._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))

    for row in rows:
        tr = tb.add_row()
        trPr2 = tr._tr.get_or_add_trPr()
        trPr2.append(OxmlElement("w:cantSplit"))
        for i, val in enumerate(row):
            c = tr.cells[i]
            c.width = Cm(widths[i])
            para = c.paragraphs[0]
            para.paragraph_format.first_line_indent = Cm(0)
            para.paragraph_format.space_before = Pt(1)
            para.paragraph_format.space_after = Pt(1)
            para.paragraph_format.line_spacing = 1.05
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            lines = str(val).split("\n")
            _f(para.add_run(lines[0]), size=size)
            for ln in lines[1:]:
                para.add_run().add_break()
                _f(para.add_run(ln), size=size)
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(2)
    sp.paragraph_format.line_spacing = 1.0
    _f(sp.add_run(""), size=5)
    return tb


def figure(fname, width_cm, caption, note=None):
    path = os.path.join(IMG, fname)
    width_cm = min(width_cm, OUT_W - 0.2)
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(5)
    para.paragraph_format.space_after = Pt(2)
    para.paragraph_format.first_line_indent = Cm(0)
    para.paragraph_format.keep_with_next = True
    para.add_run().add_picture(path, width=Cm(width_cm))
    _fig[0] += 1
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.first_line_indent = Cm(0)
    cap.paragraph_format.space_after = Pt(2 if note else 6)
    _f(cap.add_run("图 %d  %s" % (_fig[0], caption)), size=9, bold=True)
    if note:
        n = doc.add_paragraph()
        n.alignment = WD_ALIGN_PARAGRAPH.CENTER
        n.paragraph_format.first_line_indent = Cm(0)
        n.paragraph_format.space_after = Pt(6)
        _f(n.add_run(note), size=8.5, color=RGBColor(0x59, 0x59, 0x59))


def quote_block(lines, size=9.5):
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0.7)
    pf.right_indent = Cm(0.4)
    pf.space_before = Pt(3)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.15
    shade_para(para, "F7F7F7")
    for i, line in enumerate(lines):
        if i:
            para.add_run().add_break()
        _f(para.add_run(line), size=size)
    return para


def code_block(lines, size=7.5):
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0.4)
    pf.space_before = Pt(3)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.0
    shade_para(para, "F5F5F5")
    for i, line in enumerate(lines):
        if i:
            para.add_run().add_break()
        _f(para.add_run(line), latin=F_MONO, ea=F_SONG, size=size)
    return para


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ================================================================ 填写模板
# 标题：实验报告（一）→（二）
for para in doc.paragraphs:
    if "广州大学学生实验报告" in para.text and "（一）" in para.text:
        for r in para.runs:
            if "（一）" in r.text:
                r.text = r.text.replace("（一）", "（二）")
        break

tb = doc.tables[0]
tb.rows[0].cells[3].paragraphs[0].add_run("2024级 / 软件工程243班")
tb.rows[0].cells[5].paragraphs[0].add_run("庞华杰")
tb.rows[0].cells[7].paragraphs[0].add_run("32406300076")
# 课程名称
c = tb.rows[1].cells[1]
for r in c.paragraphs[0].runs:
    r.text = ""
c.paragraphs[0].runs[0].text = "面向对象分析与设计"
# 实验项目名称：需求分析 → 系统设计
c2 = tb.rows[2].cells[1]
for para in c2.paragraphs:
    for r in list(para.runs):
        r.text = ""
c2.paragraphs[0].add_run("系统设计")
tb.rows[2].cells[7].paragraphs[0].add_run("王健雄")

for row in tb.rows:
    for cell in row.cells:
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(1)
            para.paragraph_format.space_after = Pt(1)

# 日期行
for para in doc.paragraphs:
    if para.text.startswith("开课学院及实验室"):
        for r in para.runs:
            if "年" in r.text:
                r.text = "2026年"
            elif r.text.strip() == "月":
                r.text = "11月"
            elif r.text.strip() == "日":
                r.text = "18日"
        break

# 页脚：第 X 页（页码域）
footer = doc.sections[0].footer
fp = footer.paragraphs[0]
for r in list(fp.runs):
    r.text = ""
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
_f(fp.add_run("第 "), size=9)
_run = fp.add_run()._element
for tag, attrs, txt in (("w:fldChar", {"w:fldCharType": "begin"}, None),
                        ("w:instrText", {"xml:space": "preserve"}, " PAGE "),
                        ("w:fldChar", {"w:fldCharType": "end"}, None)):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(k), v)
    if txt:
        el.text = txt
    _run.append(el)
_f(fp.add_run(" 页"), size=9)

# AI 使用声明（插在信息表之后，即首页）
decl_lines = [
    "1. 活动 1（手工设计）未使用任何 AI 工具生成模型内容；图 1～图 11 及对应 .puml 源文件由本人完成并复核，先于 AI 环节提交入库。",
    "2. 活动 2、3 使用 DeepSeek（网页版）迭代提示词，生成备选体系结构方案与关键类图；活动 4 的评估框架与结论由本人确定并复核，AI 辅助整理文字。",
    "3. 与 AI 的全部交互（提示词各版本、模型原始输出、评估结论）整理于附录 A，未作删减。",
    "4. 报告文字由本人整理与定稿，AI 辅助起草与语句检查；文中读者信息、借阅数据与接口地址均为虚构示例，不含涉密或隐私信息。",
]
box = doc.add_table(rows=1, cols=1)
box.alignment = WD_TABLE_ALIGNMENT.CENTER
cell = box.rows[0].cells[0]
cell.width = Cm(OUT_W)
tcPr = cell._tc.get_or_add_tcPr()
borders = OxmlElement("w:tcBorders")
for edge in ("top", "left", "bottom", "right"):
    el = OxmlElement("w:" + edge)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), "6")
    el.set(qn("w:color"), "808080")
    borders.append(el)
tcPr.append(borders)
cell_shade(cell, "FAFAFA")
first = cell.paragraphs[0]
first.alignment = WD_ALIGN_PARAGRAPH.CENTER
first.paragraph_format.first_line_indent = Cm(0)
first.paragraph_format.space_before = Pt(2)
first.paragraph_format.space_after = Pt(2)
_f(first.add_run("AI 使用声明"), ea=F_HEI, latin=F_HEI, size=9.5, bold=True)
for line in decl_lines:
    para = cell.add_paragraph()
    para.paragraph_format.first_line_indent = Cm(0)
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(1)
    para.paragraph_format.line_spacing = 1.15
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _f(para.add_run(line), size=8.5)

# 把声明框移到信息表之后、空段落之前
anchor = tb._tbl
anchor.addnext(box._tbl)
gap = doc.add_paragraph()
gap.paragraph_format.space_after = Pt(0)
gap.paragraph_format.line_spacing = 1.0
_f(gap.add_run(""), size=1)
box._tbl.addnext(gap._p)

# 一、实验目的：改写原有内容
paras = doc.paragraphs
idx = next(i for i, x in enumerate(paras) if x.text.strip() == "一、实验目的")
goal = paras[idx + 1]
for r in list(goal.runs):
    r.text = ""
goal.paragraph_format.first_line_indent = Cm(0.74)
goal.paragraph_format.space_after = Pt(3)
goal.paragraph_format.space_before = Pt(6)
goal.paragraph_format.line_spacing = 1.3
_f(goal.add_run("本实验按“手工设计 → 提示工程 → AI 辅助生成 → 系统性评估与方案论证”四个活动，"
                "在实验一需求模型的基础上完成图书馆智能借阅助手的系统设计，具体目标为："), size=10.5)
_anchor = goal._p
for line in ["掌握面向对象系统设计的基本过程，能把需求模型转化为可实现的体系结构、类模型、状态模型与交互模型；",
             "会用“4+1”视图从逻辑、进程、开发、物理与场景五个侧面描述同一设计方案，并保持跨视图一致；",
             "掌握 diagram-as-code 的工程化工作方式：模型入版本库、以 diff 评审设计变更、用架构约束测试守护依赖规则；",
             "能对 AI 生成的备选方案做可复核的评估，从可维护性、可扩展性与需求匹配度三个维度论证最终选择。"]:
    _bp = bullet(line)
    _anchor.addnext(_bp._p)
    _anchor = _bp._p

# 二、实验器材：替换两条
idx = next(i for i, x in enumerate(doc.paragraphs) if x.text.strip() == "二、实验器材")
e1 = doc.paragraphs[idx + 1]
e2 = doc.paragraphs[idx + 2]
for para, text in ((e1, "计算机一台（Windows 11，JDK 21）；"),
                   (e2, "PlantUML 1.2024.7（本地 jar + 命令行渲染）、StarUML 6（对照体验）、"
                        "DeepSeek（网页版）、Git 2.4x、Microsoft Word（导出 PDF）。")):
    for r in list(para.runs):
        r.text = ""
    _f(para.add_run(text), size=10.5)

# 三、实验内容：改写任务清单
idx = next(i for i, x in enumerate(doc.paragraphs) if x.text.strip() == "三、实验内容")
intro = doc.paragraphs[idx + 1]
for r in list(intro.runs):
    r.text = ""
_f(intro.add_run("本实验延续实验一的“图书馆智能借阅助手”选题，建模范围仍收敛到“借阅流通子系统”，"
                 "核心用例为 UC-03“续借图书”，需完成如下任务："), size=10.5)
tasks = [
    "以实验一的需求、业务规则与架构约束为输入，完成整个系统的体系结构设计并说明风格选型理由；",
    "针对 UC-03 完成详尽的类模型、状态模型与顺序模型设计，覆盖边界、控制、实体与端口；",
    "按“4+1”视图补充用户界面原型、数据模型、实现模型与部署视图，并做跨视图一致性检查；",
    "设计并迭代结构化提示词，驱动大模型给出至少两个备选体系结构方案与关键类图；",
    "对 AI 备选方案逐项评估，形成缺陷清单，从三个维度论证最终方案，并保留 Git 提交与 diff 作为评审素材。",
]
old1 = doc.paragraphs[idx + 2]
old2 = doc.paragraphs[idx + 3]


def _task(para, text, no):
    for r in list(para.runs):
        r.text = ""
    pPr = para._p.get_or_add_pPr()
    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)
    para.paragraph_format.left_indent = Cm(0.74)
    para.paragraph_format.first_line_indent = Cm(0)
    para.paragraph_format.space_after = Pt(1)
    para.paragraph_format.line_spacing = 1.45
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _f(para.add_run("（%d）%s" % (no, text)), size=10.5)


_task(old1, tasks[0], 1)
_task(old2, tasks[1], 2)
prev = old2
for k, text in enumerate(tasks[2:], 3):
    newp = doc.add_paragraph()
    newp.paragraph_format.left_indent = Cm(0.74)
    newp.paragraph_format.first_line_indent = Cm(0)
    newp.paragraph_format.space_after = Pt(1)
    newp.paragraph_format.line_spacing = 1.45
    newp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _f(newp.add_run("（%d）%s" % (k, text)), size=10.5)
    prev._p.addnext(newp._p)
    prev = newp

# ================================================================ 四、设计概述
h1("四、设计概述与约束")
p("实验一回答“系统要做什么”，实验二回答“系统怎么做、为什么这样做”。本章先把设计输入与约束摆清楚，"
  "再给出 4+1 视图的组织方式；后续五至十一章逐个视图展开，第十二至十四章记录 AI 对照与方案论证。")

h2("4.1  从需求到设计的输入")
p("设计不是从空白开始的。实验一留下的用例规约、业务规则、非功能指标与架构约束，直接决定了本实验的候选方案空间。"
  "表 1 把关键输入与对应的设计含义并列，避免“设计”退化为重新猜需求。")
table(["编号", "实验一输入", "对设计的约束"],
      [["IN-01", "UC-03 续借图书的基本流、5 条备选流", "续借必须是一个可独立测试的应用用例，异常路径要能映射为用户可理解的答复"],
       ["IN-02", "BR-01～BR-05（次数、顺延、预约、逾期、队列位次）", "规则必须可配置、可版本化；判定顺序影响答复质量，不能散落在 Controller"],
       ["IN-03", "FR-08 与 NFR-03：答复附依据、工具调用留审计", "审计不是日志附加项，而是领域事件与 Outbox 的一等公民"],
       ["IN-04", "NFR-01：P95 ≤ 3 秒；NFR-05：ILS 故障可降级", "续借主链路必须同步、短超时、可熔断；不能把核心事务放到异步消息里"],
       ["IN-05", "实验一 4.2 的架构约束：AI 只做表达层，授权与规则留在规则层", "领域层不得依赖 LLM；智能体只能经 ITool 端口调用应用服务"],
       ["IN-06", "课程团队规模 3～5 人、无专职运维", "方案必须能被小团队部署、监控与排障；分布式组件越少越好"]],
      [1.3, 6.4, 7.1], "设计输入与约束")

h2("4.2  设计目标与优先级")
p("质量属性之间会互相拉扯。本项目的排序是：正确性与公平性 > 可审计性 > 可用性 > 性能 > 可扩展性。"
  "这个排序不是偏好，而是业务后果决定的：一次错误的续借会破坏预约队列的公平性，且无法用“重试”挽回；"
  "而系统慢 200 毫秒只是体验问题。表 2 把排序转成可验证的场景。")
table(["优先级", "质量属性", "具体场景", "设计回应"],
      [["1", "正确性与公平性", "有他人预约时，续借必须被拒绝；并发续借不能产生两条生效记录", "领域不变式 + 乐观锁 + 幂等键 + 预约校验"],
       ["2", "可审计性", "任意一次续借都能回答“谁、何时、按哪版规则、得到什么结果”", "续借申请、规则版本快照、Outbox 审计事件"],
       ["3", "可用性", "ILS 超时或 LLM 故障时，核心借阅仍可办理", "同步核心 + 熔断降级；AI 与核心分离部署"],
       ["4", "性能", "续借请求 P95 ≤ 3 秒，首 token ≤ 1.5 秒", "规则缓存、一次查询取回记录与队列、800 ms ILS 超时"],
       ["5", "可修改性", "学生/教师上限调整不改代码、不发版", "规则配置中心 + 版本号；策略对象读取配置"],
       ["6", "可扩展性", "未来接入更多高校或更多馆点", "模块化单体按接缝演进，不预先拆微服务"]],
      [1.2, 2.8, 6.4, 4.4], "质量属性场景与优先级")

h2("4.3  关键设计约束")
p("以下五条约束在后续每个视图里都要被遵守，后文用 C1～C5 回指：")
for line in [
    "C1　借阅核心不依赖大模型。LLM 只能出现在接入层与编排层；核心域不知道“对话”的存在。",
    "C2　ILS 是外部系统，可能超时、重复、乱序。任何跨系统写操作都必须有幂等键与对账路径。",
    "C3　续借必须在线实时完成。借阅记录状态与 ILS 事务不能靠最终一致“稍后再说”。",
    "C4　审计与业务数据同源。业务成功但审计丢失，等价于不可追溯，不能接受。",
    "C5　规则是易变知识。规则版本必须随申请固化，保证事后能还原“当时按哪版规则判定”。",
]:
    bullet(line)

h2("4.4  4+1 视图与本文组织")
p("“4+1”不是五张互不相干的图，而是同一设计方案的五个观察角度。本文的组织见表 3：逻辑视图回答“有哪些对象与职责”，"
  "进程视图回答“运行时怎么协作、事务边界在哪”，开发视图回答“代码怎么分模块”，物理视图回答“部署在哪”，"
  "场景则把 UC-03 串起来验证其余四个视图是否自洽。")
table(["视图", "关注点", "本文落点", "主要图件"],
      [["逻辑视图", "领域对象、职责、关系、状态", "第六、七章", "图 2 领域类图、图 3 应用类图、图 4/5 状态图"],
       ["进程视图", "运行时交互、并发、事务与失败", "第五、八章", "图 1 体系结构、图 6 顺序图"],
       ["开发视图", "模块、包、依赖规则、构建", "第十一章", "图 10 组件图"],
       ["物理视图", "节点、网络、部署与隔离", "第十一章", "图 11 部署视图"],
       ["场景（+1）", "用 UC-03 验证上述视图", "第八、九章", "图 6 顺序图、图 7 界面原型"]],
      [2.0, 4.8, 3.2, 4.8], "4+1 视图与本文组织")

# ================================================================ 五、体系结构
h1("五、体系结构设计")
h2("5.1  候选风格与选择")
p("体系结构风格的选择不能只看“先不先进”，而要看它是否匹配本项目的约束。表 4 对五种候选风格做了同口径比较；"
  "比较的结论是：采用“模块化单体 + 六边形（端口-适配器）”作为主风格，只在审计与通知这类非核心链路上使用异步。")
table(["候选风格", "主要优点", "主要代价", "本项目适配度"],
      [["分层单体", "结构简单、事务本地化、部署容易", "层间容易互相穿透，领域模型易被框架污染", "中：需要补依赖规则"],
       ["模块化单体 + 六边形", "边界清晰、核心可独立测试、事务简单、演进路径清楚", "需要团队自律与架构测试守护", "高：与 C1～C5 全部相容"],
       ["微服务", "独立部署、独立扩缩容、故障隔离", "分布式事务、运维复杂、跨服务一致性难", "低：当前规模与团队不匹配"],
       ["事件驱动", "解耦、削峰、可回放", "最终一致、排障困难、重复消费", "局部：仅用于审计/通知"],
       ["无服务器", "按需伸缩、免运维", "冷启动、厂商锁定、状态管理受限", "低：P95 与合规不匹配"]],
      [2.8, 5.2, 4.4, 2.4], "体系结构风格比较")
p("选择模块化单体的直接理由有三条。第一，续借的核心正确性依赖“记录状态 + 申请 + 审计”在一个本地事务里提交，"
  "单体天然满足；拆成微服务后必须引入 Saga 或 Outbox 补偿，复杂度上升而业务价值不变。第二，课程团队只有 3～5 人，"
  "没有专职 SRE，多服务的部署、监控、链路追踪会吞掉大部分开发时间。第三，模块化单体并不等于“一个大泥球”："
  "本文用六边形把领域、应用、适配器分开，用 ArchUnit 守护依赖方向，未来要拆时沿着端口接缝拆，而不是重写。")
p("“六边形”在本项目中的具体含义是：领域层在中心，定义实体、值对象、规则与端口接口；应用层编排用例；"
  "适配器层实现 REST、PostgreSQL、Redis、ILS、Outbox 等外部细节。依赖箭头永远指向内层，领域层不知道 Spring、"
  "不知道 SQL、也不知道大模型的存在。")

h2("5.2  分层与依赖规则")
figure("01_architecture_layered.png", 13.6, "图书馆智能借阅助手分层体系结构（逻辑视图）",
       "源文件：01_architecture_layered.puml；箭头表示编译期依赖方向")
p("图 1 的四层不是“表现层/业务层/数据层”的旧三层，而是按“谁依赖谁”重新组织的。接入层与基础设施层都在外侧，"
  "应用层只依赖领域层与端口，领域层不依赖任何外层。这样做的收益在测试上最明显：领域规则可以用纯 JUnit 测试，"
  "不需要启动 Spring、数据库或 ILS。")
table(["层", "职责", "允许依赖", "禁止事项"],
      [["接入层 / 适配器", "协议转换、参数校验、鉴权、限流、把外部模型翻译成领域模型", "应用层、领域层", "写业务规则；直接访问数据库"],
       ["应用层", "用例编排、事务边界、幂等控制、调用端口", "领域层、端口接口", "依赖具体适配器；拼接 SQL；调用 LLM"],
       ["领域层", "实体、值对象、聚合不变式、规则策略、领域事件", "仅 JDK", "依赖框架、数据库、ILS、LLM"],
       ["基础设施层", "仓储实现、ILS 防腐适配器、Redis、Outbox、配置中心", "领域层端口、应用层端口", "被领域层反向依赖"]],
      [2.6, 6.2, 3.6, 4.4], "分层职责与依赖规则")
p("这里有一个容易被忽略的细节：基础设施层“实现”端口，但编译期依赖仍然指向内层。也就是说，"
  "IlsRestAdapter 依赖 IlsGateway 接口，而不是 RenewBookUseCase 依赖 IlsRestAdapter。"
  "控制反转由组合根（Spring 配置）完成，领域层因此保持纯净。")

h2("5.3  运行时协作与并发")
p("续借的运行时路径可以概括为“同步主链路 + 异步旁路”：读者请求到答复之间，记录校验、ILS 提交、本地事务都在同步链路内；"
  "审计外发、通知、运营指标走 Outbox 异步旁路。异步只做“已经确定发生的事实”的传播，不参与“是否发生”的判定。")
table(["运行时场景", "处理方式", "一致性策略", "用户可见结果"],
      [["正常续借", "同步校验 → ILS 提交 → 本地事务 → 答复", "本地强一致；ILS 与本地用幂等键对齐", "成功 + 新应还日 + 剩余次数"],
       ["ILS 超时", "不立即判失败；按幂等键回查，仍未知则标记待对账", "状态显式化为 PENDING_RECONCILE", "“处理中，稍后可在记录中查看”，不误导为失败"],
       ["重复点击 / 重试", "幂等键命中，返回上次结果", "同一键只允许一次生效", "与首次结果一致，不重复顺延"],
       ["LLM 故障", "对话入口降级为表单/自助机；核心 API 不受影响", "核心与 AI 部署隔离", "仍可续借，只是没有自然语言解释"],
       ["规则热更新", "新请求读取新版本；在途申请保留旧版本快照", "版本号固化到申请", "新规则立即生效，历史可追溯"]],
      [2.8, 6.2, 4.0, 3.8], "续借运行时场景")
p("并发控制采用“按借阅记录的乐观锁”，而不是全局锁。loan_record.version 在每次续借时 +1；两个并发请求同时读到 version=3，"
  "只有一个能更新成功，另一个收到版本冲突后重新读取并重新校验。这样既避免超卖式的重复续借，又不会让不同读者的请求互相阻塞。")

h2("5.4  质量属性战术")
p("表 7 把 4.2 节的质量属性场景落到具体战术。战术不是口号，必须能指向代码或部署中的一个落点；"
  "如果找不到落点，说明它只是一个愿望。")
table(["质量属性", "战术", "落点", "代价 / 风险"],
      [["性能", "规则缓存 + 一次查询取回记录与预约队列 + 800 ms ILS 超时", "PolicyRepository 本地缓存；ILS 批量查询；Resilience4j TimeLimiter", "缓存可能短暂陈旧；超时阈值需按真实 ILS 压测校准"],
       ["可用性", "熔断、舱壁、降级路径", "ILS 熔断器；lending-api 与 assistant-* 分离部署；降级为“转人工/自助机”", "降级期间无自然语言答复；需运营话术配合"],
       ["可审计", "Outbox + 只追加审计表 + traceId", "本地事务写 outbox_event；审计表禁止 UPDATE/DELETE", "存储增长；需归档策略"],
       ["安全", "网关统一鉴权 + 对象级授权 + 不信任自然语言身份", "CAS/OAuth2；readerId 来自令牌；每次查询校验记录归属", "令牌泄露风险；需短期令牌与审计"],
       ["可修改", "规则外置 + 版本号 + 配置灰度", "policy_version 表；配置中心；按馆点/读者类型灰度", "规则配置错误会影响线上；需预发校验"]],
      [2.2, 5.4, 4.8, 4.4], "质量属性战术与落点")

h2("5.5  接口契约")
p("接口契约是体系结构与外部世界的接触面。本文只列与 UC-03 直接相关的四个契约；所有写接口都要求 Idempotency-Key，"
  "所有错误都返回统一结构，避免把 ILS 的异常码直接暴露给前端。")
table(["接口", "协议 / 方法", "幂等性", "说明"],
      [["续借办理", "POST /v1/renewals", "必需 Idempotency-Key", "入参 readerId、barcode；返回 RenewalResult"],
       ["续借查询", "GET /v1/renewals/{requestId}", "天然幂等", "用于前端轮询“待对账”结果"],
       ["智能体工具", "ITool: renew_book", "透传幂等键", "编排层只能经此端口调用，不能直连数据库"],
       ["ILS 续借", "POST /ils/renew", "按 idemKey 去重", "防腐层翻译领域模型与 ILS 报文"]],
      [2.6, 4.2, 4.2, 5.4], "UC-03 相关接口契约")
p("统一错误结构如下。reasonCode 供前端做分支，ruleRef 供答复引用条款，alternatives 给出可行动作；"
  "三者缺一，失败答复就会退化成“续借失败，请重试”。")
code_block([
    "{",
    '  "success": false,',
    '  "reasonCode": "RESERVED_BY_OTHER",',
    '  "message": "该书已有读者预约，按规则不能续借",',
    '  "ruleRef": "BR-03",',
    '  "alternatives": ["预约这本书", "查看排队位次", "到馆归还"],',
    '  "traceId": "b7c2f0e1-..."',
    "}",
], size=8.5)

# ================================================================ 六、类模型
h1("六、类模型设计")
p("类模型是设计的骨架。本章把类模型拆成两张图：领域类图回答“业务概念是什么、有什么不变式”，"
  "应用类图回答“一次续借由谁编排、校验如何组合”。两张图合起来，才构成完整的逻辑视图。")

h2("6.1  领域类图")
figure("02_domain_class.png", 14.4, "借阅领域类图（设计级）",
       "源文件：02_domain_class.puml；枚举通过属性类型引用，未画关联线以避免误导")
p("领域类图包含 7 个实体、3 个值对象与 6 个枚举。与实验一的分析类相比，设计级增加了 RenewalRequest（续借申请）"
  "与 policy_version 的版本快照语义；RenewalResult 从“临时返回值”提升为值对象，使顺序图、状态图与答复模板共用同一结构。")
table(["对象", "类型", "职责与关键不变式"],
      [["读者 Reader", "实体", "持有读者类型与账户状态；可续借() 只做状态判断，不查库、不调 ILS"],
       ["图书 Book / 副本 Copy", "实体 / 实体", "书目与副本一对多；续借的对象是条码级副本，不是书目"],
       ["借阅记录 LoanRecord", "实体（聚合根）", "承载借出、应还、续借次数与状态；续借只顺延 dueDate，不改变借出日期"],
       ["预约 Reservation", "实体", "同一书目内的队列位次；是否排在(读者)之前() 是 BR-03 的判定依据"],
       ["续借规则 RenewalPolicy", "实体（策略）", "学生/教师上限、顺延天数；规则版本随申请固化"],
       ["续借申请 RenewalRequest", "实体", "一次续借的完整生命周期；幂等键唯一，状态可审计"],
       ["读者编号 / 条码 / 续借结果", "值对象", "不可变；Result 统一表达成功、原因码、新应还日与剩余次数"]],
      [4.2, 2.4, 8.2], "领域对象职责")
p("三个关键决策值得单独说明。第一，副本与书目分开。如果只有 Book，就无法表达“同一本书有 5 册、第 3 册被借出”，"
  "而续借判定恰恰是条码级的；AI 生成的类图常在这里塌陷（见 14.4 的 D07）。第二，续借规则独立成对象。"
  "规则随读者类型、馆点、学期变化，属易变知识；把它塞进 LoanRecord 会让核心实体随规则改动，也会让状态图出现"
  "“记录判断自己”的循环依赖。第三，续借申请成为一等实体。传统借阅系统只改 dueDate，没有“申请”概念；"
  "但在需要幂等、对账与审计的系统中，“这次续借发生过什么”本身就是业务事实，必须有地方存放。")

h2("6.2  应用类图")
figure("03_app_class.png", 14.6, "续借用例的应用类图（边界—控制—校验链）",
       "源文件：03_app_class.puml；端口与适配器见表 11 与图 10")
p("应用类图采用边界—控制—实体（BCE）的职责划分，但把“校验”从控制类中再拆出来，形成一条显式的校验链。"
  "这样做的原因不是追求设计模式，而是业务要求：失败答复必须给出“最可行动的原因”，而原因的顺序是业务规则的一部分，"
  "不能藏在 if/else 的书写顺序里。")
table(["元素", "职责", "设计理由"],
      [["RenewalController", "HTTP 入参校验、鉴权上下文、调用用例、装配响应", "边界对象只做协议转换，不含业务判断"],
       ["RenewalToolAdapter", "把智能体的 ToolCall 翻译成 RenewalCommand", "AI 与核心之间的唯一入口，便于限流与审计"],
       ["RenewBookUseCase", "编排：幂等登记 → 组装上下文 → 校验 → 计算 → 提交 ILS → 本地事务", "一个用例一个事务边界；可独立测试"],
       ["ValidatorChain", "按固定顺序执行 4 个校验器，失败即抛出带原因码的领域异常", "顺序即业务规则；新增规则只加一个校验器"],
       ["4 个 Validator", "读者状态、逾期、次数上限、他人预约", "单一职责，可分别单测；规则依赖注入"],
       ["RenewalCommand / Context", "不可变输入与校验上下文", "避免在方法间传递散落的参数"]],
      [3.6, 7.2, 4.0], "应用类职责")
p("校验顺序固定为“读者状态 → 是否逾期 → 次数上限 → 他人预约”。这个顺序不是随意的：读者证冻结时，告诉他“已逾期”没有意义；"
  "已逾期时，告诉他“次数超限”也没有意义。先给最可行动的原因，再给次要原因，是失败答复质量的一部分。"
  "顺序在 ValidatorChain 的构造处显式声明，并有单元测试固定，不依赖代码行的先后。")

h2("6.3  端口与适配器")
p("端口是领域/应用对外的“需求”，适配器是外部世界的“供给”。表 11 列出六个端口；它们都不引用任何框架类型，"
  "因此领域测试可以用内存实现替换。")
table(["端口", "适配器", "职责", "关键约束"],
      [["LoanRepository", "PostgresLoanRepository", "按条码与读者查记录、保存记录", "带 version 乐观锁；不返回 ILS 报文"],
       ["ReservationRepository", "PostgresReservationRepository", "查询某书目的预约队列", "只读；队列位次在同一书目内唯一"],
       ["PolicyRepository", "ConfigPolicyRepository", "读取当前规则版本", "本地缓存 + 版本号；失败时回退到最后已知版本"],
       ["IlsGateway", "IlsRestAdapter", "续借提交、按幂等键回查", "800 ms 超时；熔断；不泄漏 ILS DTO"],
       ["AuditPort", "OutboxAuditAdapter", "追加领域事件", "与业务同事务；只追加不更新"],
       ["IdempotencyStore", "RedisIdempotencyStore", "登记/读取幂等键", "SETNX + 结果缓存；TTL 覆盖对账窗口"]],
      [3.4, 4.4, 4.6, 4.0], "端口与适配器")
p("ILS 适配器是一个防腐层（ACL）。ILS 的字段命名、错误码、日期格式都与领域模型不同；如果让 ILS DTO 直接进入应用层，"
  "每次 ILS 升级都会引发核心代码改动。适配器负责三件事：把领域命令翻译成 ILS 报文、把 ILS 错误码翻译成领域原因码、"
  "把 ILS 的超时/未知翻译成 PendingReconcile 而不是异常。")

h2("6.4  分析类到设计类的映射")
p("实验一的分析类不能原样搬进设计。表 12 记录每个分析类的去向：哪些保留、哪些拆分、哪些移到子系统之外。"
  "这张表是“需求—设计”可追溯性的证据，也是答辩时最容易被追问的地方。")
table(["实验一分析类", "设计类 / 去处", "变化理由"],
      [["读者", "Reader + ReaderId", "增加值对象，避免字符串裸奔"],
       ["图书 / 馆藏副本", "Book + Copy", "保留书目与副本分离，续借按条码"],
       ["借阅记录", "LoanRecord", "增加 version 乐观锁与状态枚举"],
       ["续借规则", "RenewalPolicy + policy_version 表", "规则外置并版本化，支持热更新与审计"],
       ["续借结果", "RenewalResult", "从临时返回值提升为值对象，统一答复与审计"],
       ["智能体会话 / 工具调用", "移出本子系统，仅保留 ITool 端口", "会话属编排子系统；核心不依赖 LLM（C1）"],
       ["（分析阶段没有）", "RenewalRequest / IdempotencyStore / Outbox", "为幂等、对账与审计新增的设计元素"]],
      [4.2, 5.6, 5.0], "分析类到设计类的映射")

h2("6.5  关键接口签名")
p("下面给出应用层与端口的核心签名（Java 风格伪代码）。签名的选择本身就是设计：RenewalResult 不返回布尔值，"
  "IlsGateway.renew 必须携带幂等键，RenewalCommand 不可变。")
code_block([
    "public interface RenewBookUseCase {",
    "    RenewalResult execute(RenewalCommand command);",
    "}",
    "",
    "public record RenewalCommand(ReaderId readerId, Barcode barcode,",
    "                             String idempotencyKey, Channel channel) {}",
    "",
    "public record RenewalResult(boolean success, ReasonCode reasonCode,",
    "                            LocalDate newDueDate, int remaining,",
    "                            String ruleRef, List<String> alternatives) {}",
    "",
    "public interface IlsGateway {",
    "    IlsReceipt renew(Barcode barcode, LocalDate newDueDate, String idempotencyKey);",
    "    Optional<IlsReceipt> findByIdempotencyKey(String idempotencyKey);",
    "}",
], size=7.6)
p("注意 execute 只接收一个命令对象，而不是五个散参数。这样新增渠道、语言或客户端版本时，不需要改用例签名；"
  "命令对象也便于在审计中序列化，形成可回放的一次请求。")

# ================================================================ 七、状态模型
h1("七、状态模型设计")
p("状态模型回答“一个对象在一生中能处于哪些状态、哪些事件能让它迁移、哪些迁移必须被禁止”。"
  "本章有两个状态机：借阅记录的状态机描述“一本书的借阅生命周期”，续借申请的状态机描述“一次续借请求的处理生命周期”。"
  "两者相关但不等价：申请被拒绝时，借阅记录仍停留在原状态。")

h2("7.1  借阅记录状态机")
figure("04_state_loanrecord_design.png", 12.4, "借阅记录（LoanRecord）状态图（设计级）",
       "源文件：04_state_loanrecord_design.puml；虚线框含义见状态表")
p("状态图共 6 个状态、14 条迁移。与实验一相比，设计级明确了每个状态的进入动作与合法出口，并把“续借校验中”标注为瞬时状态；"
  "它不是数据库字段，只用于把四项校验的判定过程显式化，使顺序图的 alt 分支在状态图上有语义落点。")
table(["状态", "含义", "进入动作", "合法出口"],
      [["借出", "正常在借", "借阅成功，生成记录；到期前 3 天触发提醒", "续借、逾期、归还、丢失"],
       ["续借校验中", "瞬时状态：正在执行四项校验", "读者/智能体发起续借；不落库", "通过 → 已续借；拒绝 → 借出"],
       ["已续借", "续借成功，应还日已顺延", "校验通过；dueDate += 30，renewCount++", "再次续借、逾期、归还、丢失"],
       ["逾期", "超期未还，按日计罚", "到期未还；限制续借与新借", "归还（结清罚款）、丢失"],
       ["已归还", "终态，副本释放", "归还登记", "—"],
       ["已丢失", "终态，进入赔偿流程", "读者申报或盘点缺失", "—"]],
      [2.6, 5.4, 4.2, 4.2], "借阅记录状态说明")
p("两个划分理由。其一，逾期必须是独立状态，不能只是 LoanRecord 的一个布尔属性：逾期会改变可用动作集合"
  "（不可续借、产生罚款、限制新借），符合“状态决定行为”的定义。其二，临近到期不单列状态：它不改变任何可用动作，"
  "只触发一次提醒；状态图不应为通知功能增加状态。")
p("迁移完备性检查是状态图交付前的必做动作：对每个状态问“从这里出发，所有合法动作都能走通吗”。"
  "实验一曾在“已续借”漏掉自环，导致第二次续借无路可走；本设计保留了 已续借 → 续借校验中 的迁移，"
  "监护条件为 renewCount < 上限，并由领域方法 还可续借(上限) 支撑。")

h2("7.2  续借申请状态机")
figure("05_state_renewal_request.png", 7.6, "续借申请（RenewalRequest）状态图",
       "源文件：05_state_renewal_request.puml；详细迁移条件见表 14")
p("续借申请是设计阶段新增的实体，它的状态机解决的是“一次请求到底处理到哪一步”。如果没有这个状态机，"
  "ILS 超时只能被记成一条错误日志，运营人员无法判断“这笔续借到底生效没有”，前端也无法给出准确的“处理中”提示。")
table(["状态", "含义", "进入条件", "出口 / 后续动作"],
      [["已接收", "工具调用到达，尚未登记", "通过参数校验", "幂等键存在 → 已存在；否则 → 校验中"],
       ["已存在", "同一幂等键的重复请求", "幂等存储命中", "返回上次结果，结束"],
       ["校验中", "执行四项业务校验", "新请求登记成功", "通过 → 待提交；不通过 → 已拒绝"],
       ["待提交", "已算出新应还日，等待 ILS", "校验通过", "ILS 成功 → 已生效；超时/未知 → 待对账"],
       ["已生效", "ILS 与本地均已成功", "本地事务提交完成", "结束，触发 Outbox"],
       ["已拒绝", "业务校验不通过或对账确认未生效", "返回原因码", "结束，记录拒绝审计"],
       ["待对账", "ILS 结果未知，不能判成功也不能判失败", "超时、连接中断、5xx", "对账成功 → 已生效；确认未生效 → 已拒绝"]],
      [2.2, 6.2, 3.8, 4.2], "续借申请状态说明")
p("“待对账”是本章最重要的设计决策。分布式调用中，“超时”只意味着“不知道”，不意味着“失败”。"
  "如果直接判失败并允许用户重试，而 ILS 实际上已经成功，就会产生第二次顺延——这对读者是“赚了”，"
  "对预约队列是灾难。把未知显式化为一个状态，再由定时任务按幂等键回查，才能把“不确定”收敛成“确定”。")

h2("7.3  状态与业务规则的一致性")
p("状态图的监护条件不能与用例规约里的业务规则脱节。表 15 把 BR-01～BR-05 映射到具体的校验器、领域方法或状态迁移，"
  "保证同一条规则在三处（规约、类模型、状态模型）指向同一实现。")
table(["规则", "在状态模型中的落点", "在类模型中的落点", "验证方式"],
      [["BR-01 次数上限", "已续借 → 续借校验中 的监护条件", "RenewalPolicy.上限() + V3 次数校验", "边界值测试：第 2/3 次、第 5/6 次"],
       ["BR-02 顺延 30 天", "校验通过 / dueDate += 30", "RenewalPolicy.新应还日期()", "固定日期单测；跨月/闰年"],
       ["BR-03 他人预约", "校验中 → 已拒绝 的一条监护分支", "V4 + Reservation.是否排在()", "有/无预约两组用例"],
       ["BR-04 已逾期", "借出 → 逾期；逾期态无续借出口", "V2 + LoanRecord.是否逾期()", "到期日当天/次日"],
       ["BR-05 不改队列位次", "续借成功不产生预约写操作", "领域服务不调用预约仓储的写方法", "架构测试：禁止在续借用例中注入预约写端口"]],
      [3.2, 5.4, 4.2, 3.8], "业务规则在状态与类模型中的落点")

# ================================================================ 八、顺序模型
h1("八、顺序模型设计")
h2("8.1  设计级顺序图")
figure("06_sequence_renew_design.png", 14.4, "UC-03 续借图书的设计级顺序图",
       "源文件：06_sequence_renew_design.puml；含成功、校验拒绝、幂等重放、ILS 超时四条路径")
p("顺序图把“表达层”与“规则层”严格分开：对话前端与智能体编排只负责理解与表达，真正的判定发生在 RenewBookUseCase 与"
  "ValidatorChain 内，ILS 只作为外部协作者出现。图中的每一条消息都能在类图中找到接收者，这是跨图一致性的基本要求。")
table(["消息段", "发送方 → 接收方", "职责", "事务 / 失败语义"],
      [["1–3", "读者 → 前端 → 编排", "自然语言入口；携带已认证 readerId", "无事务；鉴权失败在网关拦截"],
       ["4–6", "编排 → 工具适配器 → 用例", "意图/槽位 → renew_book 工具调用 → 命令对象", "无事务；参数不合法直接返回 400"],
       ["7–9", "用例 → 幂等存储", "登记幂等键；命中则回放上次结果", "Redis SETNX；重复请求不进入业务"],
       ["10–12", "用例 → 仓储 → 校验链", "取记录与预约队列；执行四项校验", "只读；校验失败不写业务表"],
       ["13–15", "用例 → ILS", "提交续借（携带幂等键）；超时则查回执", "ILS 调用在本地事务之外"],
       ["16–18", "用例 → 仓储 → 审计", "本地事务：记录 + 申请 + Outbox", "同一数据库事务；失败整体回滚"],
       ["19–21", "工具适配器 → 编排 → 前端", "ToolResult → 自然语言答复", "答复失败不影响已生效的续借"]],
      [1.6, 5.4, 5.2, 4.6], "顺序图消息与职责")
p("图中有一个容易被误读的地方：UC -> DB : 保存记录 + 申请 与 UC -> SIDE : 同事务写 Outbox 看起来是两次调用，"
  "实际上它们在同一个本地事务里，由应用服务统一提交。画成两条消息是为了让“Outbox 与业务同源”这件事可见；"
  "实现上是一个 @Transactional 方法内的三次写操作。")

h2("8.2  事务边界与幂等")
p("续借跨了三个一致性边界：本地数据库、Redis 幂等键、外部 ILS。三者无法用一个大事务包住，因此必须明确“谁是事实源、"
  "失败后怎么收敛”。本设计的原则是：本地数据库是事实源；ILS 是外部事实；幂等键是两者之间的对齐器。")
table(["失败点", "可能后果", "处理策略", "最终一致性"],
      [["本地事务提交前失败", "记录未更新，ILS 可能已成功", "请求返回失败；后台按幂等键回查 ILS，若已成功则补记并标记生效", "可达一致"],
       ["本地事务提交后进程崩溃", "记录已更新，Outbox 未外发", "Outbox 投递器重启后重试；审计最终补齐", "可达一致"],
       ["ILS 超时", "不知道是否成功", "不重试提交，先按幂等键回查；仍未知则 PENDING_RECONCILE", "由对账任务收敛"],
       ["重复请求", "可能重复顺延", "幂等键命中，直接回放上次结果", "强一致（同一结果）"],
       ["Redis 不可用", "无法登记幂等键", "降级为数据库唯一约束兜底；仍不可用则拒绝写请求", "不产生重复"],
       ["规则配置读取失败", "无法判定", "使用最后已知版本；仍无则拒绝并告警", "不产生错误判定"]],
      [4.0, 4.2, 5.4, 3.2], "续借失败矩阵")
p("幂等键的生成规则也需明确：由调用方（前端/编排层）生成 UUID，随请求头传入；服务端不根据“读者+条码+日期”自行拼接，"
  "因为同一读者在同一天可能先失败后重试，自行拼接会把两次不同意图的请求误判为同一次。幂等键的 TTL 覆盖对账窗口"
  "（本设计取 7 天），过期后同一键视为新请求，但数据库唯一约束仍会拦截异常重复。")

h2("8.3  异常路径与用户可见行为")
p("内部异常不能直接变成用户话术。表 18 给出原因码到用户行为的映射；这张表同时是前端文案、智能体答复模板与客服知识库的共同来源。")
table(["原因码", "触发条件", "答复要点", "替代动作"],
      [["READER_FROZEN", "账户冻结/欠费", "说明账户状态与解除方式，不评价读者", "去缴费 / 联系馆员"],
       ["OVERDUE", "记录已逾期", "说明逾期不可续借，先归还并结清罚款", "查看罚款 / 到馆归还"],
       ["QUOTA_EXCEEDED", "达到学生 2 次或教师 5 次上限", "说明已续次数与上限，不承诺例外", "到馆续借 / 预约"],
       ["RESERVED_BY_OTHER", "有他人预约", "说明公平性规则，给出预约与排队入口", "预约 / 查看位次"],
       ["ILS_TIMEOUT", "ILS 超时或未知", "说明“处理中”，承诺可查询，不误导为失败", "稍后查询 / 转人工"],
       ["SYSTEM_BUSY", "限流或依赖熔断", "说明暂时繁忙，保留请求号", "稍后重试 / 自助机"]],
      [3.6, 4.4, 7.0, 4.0], "原因码与用户可见行为")
p("失败答复必须包含“原因 + 依据 + 下一步”三要素。只给“续借失败”会把解释成本推给馆员；只给“有预约”而不给条款，"
  "读者会怀疑系统随意拒绝；只给原因不给替代动作，读者仍然完不成任务。这三要素也是 NFR-02 的验收口径。")

# ================================================================ 九、界面原型
h1("九、用户接口与界面原型设计")
h2("9.1  交互设计原则")
p("本系统的界面不是“聊天框 + 随便回”，而是“对话为入口、卡片为凭据、动作可撤销”。具体有五条原则："
  "① 先确认再产生副作用；② 每个结论都附规则依据；③ 失败必须给替代动作；④ 关键信息用结构化卡片而不是一段长文本；"
  "⑤ 任何自然语言入口都能退回到表单/自助机。")
p("为什么坚持“先确认”：续借是一个会改变应还日期与预约队列的写操作。读者说“续借《设计模式》”时，可能手上有多册同名书，"
  "也可能只是想问“还能续几次”。在信息不足时直接执行，错误成本高于多一次点击。确认卡只在“存在歧义或首次操作”时出现；"
  "如果读者已给出条码并明确说“确认续借”，可以直接执行并展示结果。")

h2("9.2  界面原型")
figure("07_ui_prototype.png", 10.6, "对话式续借界面原型（成功态与失败态）",
       "源文件：07_ui_prototype.puml（PlantUML Salt 线框）；非最终视觉稿")
p("原型包含三个区域：上方是对话与记录确认卡，中间是结果与依据，下方是输入区。确认卡把“书名、条码、应还日、已续次数”"
  "四个关键字段并排展示，避免读者在只看到书名时误确认。成功态直接给出新应还日与剩余次数；失败态给出原因码对应的"
  "自然语言、规则条款与两个可点击的替代动作。")
p("原型刻意没有做“一键续借全部”。批量操作会放大误操作的影响，也会让“哪一本失败了、为什么”难以解释；"
  "在当前业务量下，逐本确认的成本可以接受。未来若出现“每学期集中续借”的需求，应单独设计批量确认页，而不是在对话里加一句“全部续借”。")

h2("9.3  交互状态与反馈规范")
table(["界面状态", "触发", "界面表现", "可执行动作"],
      [["待澄清", "书名匹配到多册/多本", "列出候选卡片，标注条码与应还日", "选择一本 / 取消"],
       ["待确认", "已定位唯一记录", "确认卡 + 规则摘要（次数、预约）", "确认续借 / 查看其他在借"],
       ["处理中", "已提交，等待 ILS", "禁用按钮，显示“正在办理”", "无；可离开，结果进消息中心"],
       ["成功", "本地事务提交完成", "新应还日、剩余次数、规则引用", "查看记录 / 再续借（若还有次数）"],
       ["失败", "校验拒绝或系统异常", "原因、条款、替代动作", "预约 / 缴费 / 转人工 / 重试"],
       ["待对账", "ILS 结果未知", "“处理中，稍后可在借阅记录查看”", "查询进度 / 转人工"]],
      [2.8, 4.4, 5.2, 4.2], "续借界面状态与反馈")
p("“处理中”状态不能省略。如果没有它，前端只有成功/失败两态，ILS 超时时只能显示失败；读者重试后可能重复续借。"
  "把“处理中”做成一个可查询的状态，既符合后端的状态机，也避免前端用“转圈直到超时”糊弄用户。")

h2("9.4  可访问性与降级")
p("对话界面必须可键盘操作、可被屏幕阅读器读出确认卡字段；颜色不是唯一的信息载体，成功/失败同时用文字与图标表示。"
  "对于不使用对话的读者，系统保留 OPAC 与自助机的表单路径；两条路径调用同一个 POST /v1/renewals，"
  "不存在“对话能续、表单不能续”的功能分裂。LLM 故障时，前端自动切换到表单模式，并提示“自然语言服务暂不可用，"
  "可使用表单办理”，而不是显示一个空白聊天框。")

# ================================================================ 十、数据模型
h1("十、数据模型设计")
h2("10.1  设计原则")
p("数据模型不是把类图直接翻译成表。类模型表达业务概念与不变式，数据模型还要考虑约束、索引、并发、归档与审计。"
  "本设计遵循四条原则：① 业务表第三范式，审计表允许冗余；② 所有状态字段用受控 vocabulary，不用自由文本；"
  "③ 跨系统写操作必须有幂等键与对账字段；④ 审计只追加，不更新、不删除。")
p("另一个原则是“领域模型与持久化模型分离”。LoanRecord 在领域里是一个带行为的对象，在数据库里是一行带 version 的记录；"
  "仓储负责两者之间的映射，领域层不知道 version 是乐观锁还是普通字段。这样既保留了领域模型的表达力，"
  "又让数据库可以独立优化索引。")

h2("10.2  核心借阅数据")
figure("08_data_model_core.png", 10.8, "核心借阅数据模型（主键、外键与关系基数）",
       "源文件：08_data_model_core.puml；完整列定义见表 20")
p("图 8 只画主键、外键与关系基数，完整列定义见表 20。这样做的原因是可读性：把 30 多个字段塞进一张图，"
  "读者只能看到一团文字；把字段放进表格，图只负责回答“表之间怎么关联”。")
table(["表", "主键", "关键字段", "约束 / 索引"],
      [["reader", "reader_id", "reader_type、account_status、name_masked", "reader_type、account_status 建索引；姓名脱敏存储"],
       ["book", "isbn", "title、author", "title 建全文索引（供检索，不参与续借判定）"],
       ["copy", "barcode", "isbn、copy_status、location", "isbn 外键；copy_status 建索引"],
       ["loan_record", "loan_id", "reader_id、barcode、due_date、renew_count、loan_status、version", "reader_id+barcode 联合索引；version 乐观锁"],
       ["reservation", "reservation_id", "isbn、reader_id、queue_position、reservation_status", "isbn+queue_position 唯一；status 索引"]],
      [2.6, 2.6, 6.2, 3.4], "核心借阅表设计")
p("loan_record 是续借的聚合根，也是并发冲突的焦点。version 字段不参与业务展示，只用于乐观锁；"
  "每次续借成功 +1，更新语句带 WHERE version = ?。如果更新行数为 0，说明有并发修改，应用层重新读取并重新校验，"
  "而不是直接覆盖。reader_id + barcode 的联合索引支撑“按读者与条码查在借记录”这一最高频查询。")

h2("10.3  续借事务与审计数据")
figure("09_data_model_renewal.png", 12.0, "续借事务与审计数据模型",
       "源文件：09_data_model_renewal.puml；idempotency_key 为唯一约束")
table(["表", "主键", "关键字段", "约束 / 索引"],
      [["renewal_request", "request_id (uuid)", "loan_id、policy_version、idempotency_key、request_status、new_due_date、reason_code、ils_receipt", "idempotency_key 唯一；(loan_id, request_status) 部分索引"],
       ["policy_version", "policy_version", "student_limit、teacher_limit、renew_days、effective_from、effective_to", "生效区间不重叠（由配置发布流程保证）"],
       ["renewal_audit", "audit_id", "request_id、event_type、detail_json、trace_id、created_at", "request_id 索引；只追加"],
       ["outbox_event", "event_id (uuid)", "aggregate_id、event_type、payload_json、published_at、retry_count", "published_at 为空的部分索引（待投递）"]],
      [3.2, 2.8, 6.4, 3.2], "续借事务与审计表设计")
p("renewal_request.idempotency_key 是唯一约束，不是普通索引。唯一约束是防止重复续借的最后一道防线：即使 Redis 失效、"
  "应用层判断失误，数据库也会拒绝第二条相同键的记录。(loan_id, request_status) 的部分索引用于保证“同一借阅记录最多一条"
  "处理中的申请”，避免并发请求同时进入 待提交。")
p("renewal_audit.detail_json 保存原因码、规则版本、ILS 回执与耗时。审计表不设外键级联删除，借阅记录归档后审计仍保留；"
  "这是 C4“业务成功但审计丢失不可接受”的物理落点。outbox_event 与业务表在同一个数据库，靠本地事务保证同生共死，"
  "再由投递器异步外发。")

h2("10.4  数据生命周期与一致性")
table(["数据", "保留期", "归档 / 脱敏", "依据"],
      [["借阅记录", "长期（≥ 5 年）", "归档到历史库；姓名以 reader_id 关联", "借阅争议与审计需要"],
       ["续借申请", "≥ 3 年", "与借阅记录同期归档", "复现判定过程"],
       ["续借审计", "≥ 180 天在线，之后转对象存储", "detail_json 脱敏（去除姓名、联系方式）", "NFR-03"],
       ["会话与工具调用日志", "≥ 180 天", "姓名、读者编号脱敏；仅保留 traceId", "NFR-03 / NFR-04"],
       ["规则版本", "长期", "不删除；以 effective_to 标记失效", "C5 版本可追溯"]],
      [3.4, 4.2, 4.6, 3.2], "数据生命周期")
p("一致性策略分三层：本地事务保证“记录 + 申请 + Outbox”原子；幂等键保证“同一请求最多生效一次”；对账任务保证"
  "“ILS 与本地最终对齐”。三层各司其职，不能用其中一层替代另一层。比如只靠对账而不做幂等，重复请求仍会产生两次顺延；"
  "只靠幂等而不做对账，ILS 超时的请求会永远停在“待对账”。")

# ================================================================ 十一、实现模型
h1("十一、实现模型与部署视图")
h2("11.1  模块划分")
figure("10_component_development.png", 8.4, "开发视图：Maven 模块与编译期依赖",
       "源文件：10_component_development.puml；箭头为编译期依赖")
p("实现模型按“业务能力 + 依赖方向”分模块，而不是按技术分层分包。lending-domain 只依赖 JDK；"
  "lending-application 依赖 domain 与端口接口；lending-infrastructure 实现端口；lending-api 只做协议适配。"
  "assistant-* 是独立的智能体侧模块，通过 REST/Tool 调用借阅侧，不共享领域对象。")
table(["模块", "职责", "允许依赖", "关键测试"],
      [["lending-domain", "实体、值对象、规则、领域事件", "仅 JDK", "纯 JUnit：规则边界、状态迁移"],
       ["lending-application", "用例、校验链、事务、幂等", "domain + 端口接口", "Mock 端口：成功/拒绝/超时/重放"],
       ["lending-infrastructure", "PostgreSQL 仓储、ILS ACL、Redis、Outbox", "domain/application 端口", "Testcontainers：真实 PG/Redis"],
       ["lending-api", "REST 控制器、鉴权上下文、错误映射", "application", "MockMvc：契约与错误码"],
       ["assistant-gateway", "会话、鉴权、限流、路由", "无领域依赖", "契约测试：转发与限流"],
       ["assistant-orchestrator", "意图、槽位、工具选择、答复装配", "知识检索、LLM 网关", "录制回放：意图与工具调用"],
       ["knowledge-service", "规则/馆藏检索（RAG）", "向量库、文档库", "召回率与引用准确性"],
       ["assistant-web / ops-console", "对话前端、运营与人工介入", "gateway / api", "端到端：续借成功/失败/待对账"]],
      [3.4, 5.4, 3.2, 3.8], "模块职责与依赖")
p("依赖规则不是靠自觉，而是靠 ArchUnit 测试守护：domain 不得依赖 org.springframework、java.sql、com.fasterxml；"
  "application 不得依赖 infrastructure；api 不得直接依赖 repository。任何一次“图方便”的跨层调用都会让构建失败。"
  "这是“模型入版本库、以 diff 评审”的具体兑现：评审者不需要读全部代码，只需看依赖测试是否仍然通过。")

h2("11.2  部署视图")
figure("11_deployment.png", 14.4, "部署视图：接入区、应用区、数据区与外部服务",
       "源文件：11_deployment.puml；应用区无状态，可水平扩容")
p("部署上把“智能体侧”和“借阅侧”分开：assistant-* 与 lending-api 是两组独立的工作负载，LLM 网关故障只影响前者。"
  "lending-api 至少 2 副本，配合 PostgreSQL 主从与 Redis 集群；assistant-orchestrator 也至少 2 副本，"
  "避免单点。所有入口经 DMZ 的 Nginx/API 网关做 TLS 终止、限流与鉴权。")
table(["节点 / 区域", "实例", "协议", "失败隔离"],
      [["用户终端", "浏览器 / 企业微信 / 自助机", "HTTPS", "前端降级为表单"],
       ["接入区", "Nginx / API 网关", "HTTPS / OAuth2", "限流、鉴权失败在边界拦截"],
       ["应用区（智能体）", "assistant-gateway ×2、orchestrator ×2、knowledge-service ×1", "REST / HTTPS", "LLM 故障不影响借阅核心"],
       ["应用区（借阅）", "lending-api ×2", "REST / JDBC / Redis", "单副本故障由负载均衡转移"],
       ["数据区", "PostgreSQL 主从、Redis 集群、对象存储", "JDBC / RESP / S3 API", "主从切换；审计归档不阻塞主链路"],
       ["外部服务", "LLM 网关、ILS、CAS", "HTTPS / REST", "熔断、超时、降级"]],
      [3.0, 6.0, 3.4, 3.4], "部署节点与协议")

h2("11.3  配置与发布")
p("规则配置走独立的配置中心，不跟代码一起发版。每次规则变更生成一个新的 policy_version，记录生效区间与变更人；"
  "应用在启动时加载当前版本，运行中按 TTL 刷新。配置错误不能让系统“带着错规则上线”：发布前跑一遍规则用例集，"
  "发布后先灰度 1% 流量，观察拒绝率与投诉率再扩大。")
p("数据库变更用 Flyway 管理，禁止手工改生产表。发布顺序为“先加列/加表 → 再发应用 → 最后清理旧列”，"
  "保证滚动升级期间新旧版本都能工作。借阅核心的发布采用蓝绿或金丝雀，智能体侧可以独立发布；"
  "两侧通过版本化的 REST 契约解耦，避免“前端发了、后端没发”的联调事故。")

h2("11.4  可观测性")
p("续借链路的可观测性围绕四个指标：续借成功率、P95 时延、ILS 错误率、待对账积压量。前三个是业务健康度，"
  "第四个是数据一致性风险。每个请求带 traceId，从网关透传到 ILS 适配器；审计表与日志都用同一个 traceId，"
  "排障时可以从一条用户投诉直接定位到具体申请、规则版本与 ILS 回执。")
p("告警阈值建议：续借成功率 5 分钟低于 95%、P95 超过 2.5 秒、ILS 错误率超过 3%、待对账超过 50 条且 30 分钟未下降，"
  "任一触发即告警。日志中的姓名、读者编号、联系方式必须脱敏；审计明细只保留 reader_id 与必要的业务字段，"
  "不记录完整对话原文。")
# ================================================================ 十二、提示工程
h1("十二、活动 2：提示工程")
p("活动 2 的目标不是“让 AI 画一张好看的图”，而是把设计约束写成可执行、可比较的规格。与实验一相比，"
  "本轮的难点在于：不仅要 AI 生成模型，还要它给出备选方案并说明取舍；不仅要类图，还要把质量属性与业务规则"
  "映射到具体元素。共迭代三轮，每轮只改一个变量。")

h2("12.1  设计约束的编码")
p("提示词按“角色—任务—输入—边界—约束—输出—禁止—自检”八段组织。表 25 说明每一段放什么、为什么必须放。")
table(["段落", "本例内容", "缺失后果"],
      [["角色", "高校图书馆系统的架构师，熟悉 UML、DDD 与 PlantUML", "输出泛泛而谈，缺少架构判断"],
       ["任务", "为 UC-03 续借设计备选体系结构（≥2 个）与关键类图", "AI 只给一个“标准答案”，无法比较"],
       ["输入", "实验一的 FR/NFR/BR、C1～C5、团队规模", "AI 按“互联网大厂”默认假设设计"],
       ["边界", "借阅流通子系统；智能体、ILS 为外部；不建模推荐/积分", "幻觉功能混入，方案失焦"],
       ["约束", "强一致、幂等、审计、ILS 超时、规则版本化", "AI 用最终一致替代强一致"],
       ["输出", "每个方案：风格定位 + 组件图 + 三维自评 + 不适用场景；类图用 PlantUML", "只给图，不给取舍理由"],
       ["禁止", "不得添加需求外功能；不得把数据库/大模型画成参与者；不得省略多重性", "出现推荐服务、把 LLM 当领域对象"],
       ["自检", "逐条回答是否满足 C1～C5、BR-01～BR-05", "AI 不自查，缺陷无人提示"]],
      [1.8, 8.6, 4.4], "提示词的八段结构")

h2("12.2  第一轮：V1 与它的失败")
p("V1 只写了“请为图书馆续借系统设计微服务架构，用 PlantUML 输出”，看起来直接，实际上把最重要的约束都省了：")
quote_block([
    "你是一名资深架构师。请为高校图书馆的“续借图书”功能设计系统架构，",
    "要求使用微服务风格，给出 PlantUML 组件图和核心类图。",
])
p("V1 的输出有三个可复现的问题：① 直接假设“微服务是正确答案”，没有比较其他风格；② 把“推荐服务”“信用分服务”"
  "写进架构，理由是“提升用户体验”，但需求里根本没有；③ 类图只有 Reader、Book、Loan 三个类，"
  "没有副本、预约、规则，也没有幂等与审计。更关键的是，它没有回答“为什么微服务适合这个项目”。")
p("失败原因很清楚：提示词只给了“要什么”，没有给“在什么约束下选型”。对架构任务而言，约束比图更重要；"
  "没有约束，大模型会默认套用互联网行业的“高并发微服务”模板。")

h2("12.3  第三轮：V3 的最终形态")
p("V2 补上了系统边界、C1～C5、BR-01～BR-05 与输出格式；V3 进一步要求“至少两个不同风格”的方案、三维自评、"
  "不适用场景，并把“禁止事项”写成负向约束。V3 全文如下（V2 见附录 A.1）：")
quote_block([
    "【角色】你是高校图书馆系统的架构师，熟悉 UML 2.5、领域驱动设计与 PlantUML。",
    "",
    "【任务】为 UC-03“续借图书”给出两个不同风格的备选体系结构方案，并给出关键类图。",
    "两个方案必须能回答“为什么选它、什么情况下不选它”，而不是同一方案的两种画法。",
    "",
    "【输入】系统是“图书馆智能借阅助手”，本次只设计借阅流通子系统；核心用例 UC-03。",
    "业务规则：BR-01 学生最多 2 次/教师 5 次；BR-02 每次顺延 30 天；BR-03 有他人预约不可续借；",
    "BR-04 已逾期不可续借；BR-05 续借不改变预约队列位次。",
    "质量属性：P95 ≤ 3 秒；审计留存 ≥ 180 天；ILS 故障可降级；规则可配置。",
    "团队 3～5 人，无专职运维；现有 ILS 只能通过 REST 调用。",
    "",
    "【边界】只设计借阅流通子系统。智能体编排、知识检索、运营监控属于外部系统；",
    "数据库、缓存、大模型 API 是实现细节，不得作为参与者或领域类。",
    "",
    "【约束】C1 借阅核心不得依赖大模型；C2 ILS 可能超时/重复，必须有幂等键与对账；",
    "C3 续借必须在线实时完成，借阅状态不能用最终一致；C4 审计与业务数据同源；",
    "C5 规则必须版本化，能还原“当时按哪版规则判定”。",
    "",
    "【输出要求】",
    "A. 两个备选方案，每个包含：风格名称与一句话定位；PlantUML 组件图（@startuml/@enduml 完整）；",
    "   可维护性、可扩展性、需求匹配度自评（1～5 分，各给理由）；该方案不适用的场景。",
    "B. 一张关键类图，必须覆盖读者、书目、副本、借阅记录、预约、续借规则、续借申请；",
    "   关系必须标注类型与两端多重性。",
    "C. 一张表，说明 BR-01～BR-05 分别落在哪个类/方法/校验器。",
    "",
    "【禁止】不得添加需求之外的功能（推荐、积分、信用分等）；不得把数据库、缓存、大模型画成参与者；",
    "不得用“稍后一致”代替 C3；不得省略多重性。",
    "",
    "【自检】生成后附一张表，逐条回答 C1～C5 是否满足、BR-01～BR-05 是否有落点；",
    "不满足的条目要明确说“不满足”，不要编造理由。",
])

h2("12.4  迭代记录")
table(["版本", "相对上一版的主要改动", "改动目的", "观察到的效果"],
      [["V1", "角色 + 任务 + “微服务” + 输出格式", "建立基线", "单一方案；出现推荐/信用分；类图仅 3 个类"],
       ["V2", "补边界、C1～C5、BR-01～BR-05、中文命名", "压缩自由发挥空间", "幻觉功能消失；类图增至 7 个类；仍只给一个方案"],
       ["V3", "要求两个不同风格 + 三维自评 + 不适用场景 + 负向约束 + 自检", "把评估标准前置，逼出取舍", "给出模块化单体与领域服务拆分两个方案；推荐方案一；类图覆盖主要领域类"]],
      [1.4, 6.8, 3.6, 3.0], "提示词三轮迭代记录")
p("三轮迭代的经验与实验一一致：提示词能消除“明显越界”的错误（幻觉功能、参与者错位），也能逼出“备选方案”这种结构性输出，"
  "但无法让 AI 真正理解“本项目为什么不该用微服务”。它会给出两个方案，却仍然推荐更“先进”的那个；"
  "它会在自评表里给自己打高分，却漏掉“最终一致与 C3 冲突”这种致命问题。评估仍然只能由人完成。")

# ================================================================ 十三、AI 生成
h1("十三、活动 3：AI 辅助生成")
p("用 V3 提示词驱动 DeepSeek（网页版，2026-11-18）生成备选方案。原始输出包含两段组件图、一段关键类图、"
  "一张业务规则落位表与一段取舍建议；本文按原样转存为 12_AI_arch_A_raw.puml、13_AI_arch_B_raw.puml、"
  "14_AI_class_raw.puml，未做任何语义或语法修改，三份代码首轮渲染均成功（0 处语法修正）。"
  "与实验一相比，这次 AI 的输出质量明显更高：它没有堆砌“互联网大厂”模板，而是围绕 C1～C5 给出了两个可比较的方案。")

h2("13.1  AI 方案一：模块化单体 + 六边形端口适配器")
figure("12_AI_arch_A_raw.png", 14.4, "AI 方案一：模块化单体 + 六边形端口适配器（原始输出）",
       "源文件：12_AI_arch_A_raw.puml；未修改；图内标题为 AI 原始输出的一部分")
p("方案一把借阅流通做成“单一部署单元”，内部按入站适配器、应用层、领域模块、出站端口、出站适配器分区。"
  "领域模块再分为借阅、预约、规则、馆藏四块；ILS、仓储、审计、幂等、Outbox 都通过端口隔离，"
  "其中审计仓储明确标注“与业务同库同事务”，并额外设置了 ILS 回调/对账入口。AI 的自评是："
  "可维护性 5、可扩展性 3、需求匹配度 5。")
p("这份图最值得肯定的地方是“没有把 AI 写进核心”：智能体、LLM、知识检索都没有出现在借阅子系统里，"
  "核心只暴露 REST 与对账入口。它还主动补上了幂等端口、Outbox 端口与对账用例——这三样正是 C2/C4 的落点，"
  "说明 V3 里“C2 幂等与对账、C4 审计同源”的约束被真正读进去了。")
p("但它仍然是一张组件图，不是完整设计：没有状态机、没有顺序图、没有数据模型、没有界面与部署视图；"
  "“对账”只以一个 ReconcileApplicationService 出现，没有说明“ILS 超时后申请处于什么状态、谁负责收敛、"
  "前端如何查询”。这些缺口不会在组件图上暴露，却决定了系统在真实故障下能不能用。")

h2("13.2  AI 方案二：借阅核心 + 规则服务 + ILS 网关服务")
figure("13_AI_arch_B_raw.png", 14.2, "AI 方案二：借阅核心 / 规则服务 / ILS 网关服务（原始输出）",
       "源文件：13_AI_arch_B_raw.puml；未修改")
p("方案二把“续借决策”和“ILS 集成”拆成独立服务：借阅核心保留 Loan/Reservation 聚合、同源审计与 Outbox；"
  "规则服务按版本求值；ILS 网关服务集中处理幂等键、重试、熔断与对账；另有投影同步器维护 Patron/Bib/Item 的本地只读投影。"
  "AI 的自评是：可维护性 3、可扩展性 5、需求匹配度 4。")
p("这是一个比“全量微服务”克制得多的拆分：它没有把每个实体都拆成服务，只拆了变化频率与外部依赖最强的两块。"
  "规则服务独立后可以按版本灰度，ILS 网关可以独立限流与熔断，这两点确实能缓解主链路的耦合。"
  "代价也很清楚：借阅核心与规则服务、ILS 网关之间是同步调用，一次续借要跨三个进程，P95 与故障面都会变大；"
  "而且“本地事务”只覆盖借阅核心，规则服务或 ILS 网关失败时如何回滚、如何对账，图中没有给出。")

h2("13.3  AI 关键类图")
figure("14_AI_class_raw.png", 14.2, "AI 生成的关键类图（原始输出，未经修改）",
       "源文件：14_AI_class_raw.puml；关系类型与多重性由 AI 标注")
p("这张类图比实验一那版 AI 类图成熟得多：它保留了 Patron、BibliographicRecord、Item、Loan、Reservation、"
  "RenewRuleSet、RuleEvaluator、RenewalApplication、RuleDecision 九个核心类，区分了“书目—副本—借阅记录—预约”，"
  "并把规则版本、判定结果、幂等键都放进了续借申请。它已经具备“条码级续借”和“规则版本可追溯”的骨架。")
p("仍有几处结构性问题。第一，Loan 1 — 0..1 RenewalApplication 只能表达“当前一条申请”；学生可续借 2 次、"
  "教师可续借 5 次，同一借阅记录在生命周期内会产生多条申请，历史被压成一条。第二，Item 1 — 0..* Reservation "
  "把预约挂在副本上；大多数高校的预约是书目级（“这本书有人预约”），若按副本预约，BR-03 的语义需要重新定义。"
  "第三，RenewDomainService 与 RuleEvaluator 都能做规则判定，职责有重叠；第四，类图没有画出 AuditPort、"
  "IdempotencyPort、OutboxPort 对应的类，和组件图的端口清单对不上。这些问题不是“画错了”，而是“画到组件层就停了，"
  "没有继续往状态与事务层推”。")

h2("13.4  原始输出与处理记录")
table(["序号", "项目", "记录"],
      [["1", "语法修正", "0 处：三份代码块首轮渲染成功，未改任何字符"],
       ["2", "语义修改", "0 处：未调整组件、类、关系，以保证评估对象与生成结果一致"],
       ["3", "渲染参数", "java -jar plantuml.jar -tpng -charset UTF-8，默认 DPI"],
       ["4", "对话轮次", "同一会话连续三轮（V1→V2→V3），中间未清空上下文"],
       ["5", "AI 自评", "方案一 5/3/5，方案二 3/5/4；类图自评“关系完整”，与人工评估存在差距"]],
      [1.2, 3.2, 10.4], "AI 输出处理记录")
p("这次 AI 输出的“可取之处”与“不足之处”都很典型：它能把提示词里的约束逐条落成组件与端口，"
  "却不会主动补上状态机、事务边界与跨视图一致性；它能给出两个方案并说明适用场景，却会把自评分打得偏高。"
  "评估的重点因此从“找低级错误”转向“检查结构完整性与失败语义”。")

# ================================================================ 十四、评估
h1("十四、活动 4：系统性评估与方案论证")
h2("14.1  评估方法")
p("按指导书要求，从可维护性、可扩展性、需求匹配度三个维度评估，权重分别为 0.35、0.25、0.40。"
  "需求匹配度权重最高，因为架构首先要在当前约束下正确工作；可扩展性权重最低，因为当前规模与团队并不支持为“未来可能”"
  "付出分布式复杂度。每个维度按 1～5 分打分，评分必须有证据（指向图、表或需求编号），不允许只写“感觉”。")
table(["维度", "权重", "5 分标准", "1 分标准"],
      [["可维护性", "0.35", "模块边界清晰、依赖单向、部署与排障简单、测试可分层", "边界混乱、跨服务排障困难、变更牵一发动全身"],
       ["可扩展性", "0.25", "可按接缝独立扩展、新增能力不改核心", "只能整体扩容，新增能力需改多处"],
       ["需求匹配度", "0.40", "C1～C5 与 BR-01～BR-05 全部有落点，NFR 可验证", "关键约束无落点，或架构风格与约束冲突"]],
      [2.4, 1.4, 6.4, 4.6], "评估维度与评分标准")

page_break()
h2("14.2  三个方案的对比")
table(["方案", "风格", "可维护性", "可扩展性", "需求匹配度", "加权分"],
      [["手工方案", "模块化单体 + 六边形", "4.6", "3.8", "4.8", "4.48"],
       ["AI 方案一", "模块化单体 + 端口适配器", "4.4", "3.2", "4.5", "4.14"],
       ["AI 方案二", "借阅核心 + 规则服务 + ILS 网关", "3.2", "4.6", "3.8", "3.79"]],
      [3.2, 5.2, 2.2, 2.2, 2.2, 2.0], "备选方案三维评分", size=8.0)
p("AI 方案一与手工方案同源，加权分只差 0.34；差距不在“风格选得对不对”，而在“设计做到什么程度”。"
  "AI 方案二的可扩展性最高，但可维护性最低；它不是错的方案，而是“当前团队与当前故障处理能力还接不住”的方案。"
  "下面逐维度展开理由。")

h2("14.3  三维度论证")
p("（1）可维护性。手工方案与 AI 方案一都是单一部署单元、本地事务、端口隔离，排障路径短；手工方案多出状态机、"
  "顺序图、数据模型与校验链，失败时能直接定位到“申请处于哪个状态、哪条校验拒绝、哪张表没提交”，因此给 4.6；"
  "AI 方案一只画到组件层，没有这些落点，给 4.4。AI 方案二把规则与 ILS 拆成独立服务，一次续借跨三个进程，"
  "需要分布式追踪、契约版本管理与三套部署流水线；对 3～5 人、无专职运维的团队，维护成本明显上升，给 3.2。")
p("（2）可扩展性。AI 方案二最强：规则服务可以独立灰度，ILS 网关可以独立限流、熔断、扩缩容，规则变更不会牵动借阅核心。"
  "手工方案次之：模块化单体可以通过水平复制 lending-api 扩容，知识检索与运营监控已经独立成模块，必要时可先拆；"
  "AI 方案一再次，组件图没有说明模块如何独立部署。但当前负载只有日均约 2000 次会话、峰值几十 QPS，"
  "方案二的扩展性优势在短期内用不上，却要立即支付分布式复杂度。")
p("（3）需求匹配度。手工方案把 C1～C5 全部落到具体元素（表 30），并给出 PENDING_RECONCILE 状态、"
  "校验链顺序、对账任务与失败话术；AI 方案一的端口/Outbox/审计设计已经覆盖大部分约束，但没有说明“未知结果”如何显式化，"
  "给 4.5；AI 方案二的规则/网关拆分对 C2 更专业，但跨服务同步调用增加了 C3 与 P95 的风险，审计同源也只覆盖借阅核心，"
  "规则服务与网关服务的事件没有进入同一审计链，给 3.8。")
table(["约束 / 规则", "手工方案落点", "AI 方案一", "AI 方案二"],
      [["C1 核心不依赖 LLM", "领域层仅 JDK；LLM 在编排层", "借阅子系统无 LLM；端口隔离外部", "借阅核心无 LLM；规则/网关服务也不依赖 LLM"],
       ["C2 幂等与对账", "IdempotencyStore + PENDING_RECONCILE", "幂等端口 + Outbox + 对账入口", "ILS 网关内置幂等键 + Reconciler（最集中）"],
       ["C3 在线强一致", "本地事务 + 同步 ILS", "单库本地事务 + 同步 ILS", "借阅核心本地事务，但规则/ILS 为跨服务同步调用"],
       ["C4 审计同源", "Outbox 与业务同事务", "JpaAuditRepository 与业务同库同事务", "AuditWriter 在借阅核心；规则/网关事件未闭合"],
       ["C5 规则版本化", "policy_version 固化到申请", "RenewalApplication 记录 ruleVersionId", "规则服务按版本求值，申请携带版本"],
       ["BR-03 他人预约", "书目级 Reservation + V4 校验", "Item 级 Reservation，语义需澄清", "Reservation 聚合，但查询来源与层级未说明"]],
      [3.2, 5.4, 3.4, 3.2], "约束与规则的落点对比")

h2("14.4  AI 方案的缺陷清单")
p("共记录 10 条缺陷：严重 3 条、中等 5 条、轻微 2 条。与常见的 AI 生成件不同，这次没有“幻觉功能”“明文密码”这类低级错误，"
  "缺陷集中在聚合边界、关系多重性、跨服务事务与跨图一致性上。")
table(["编号", "类别", "位置", "缺陷描述", "严重度"],
      [["D01", "聚合边界", "AI 方案一", "LoanAgg 直接引用 Bib/Item/Patron 实体，跨聚合对象引用；聚合根之间应只引用 ID，否则边界失效", "严重"],
       ["D02", "关系多重性", "AI 类图", "Loan 1—0..1 RenewalApplication 只能表达“当前一条申请”；同一借阅记录可续借多次，应为 1—0..*", "严重"],
       ["D03", "事务边界", "AI 方案二", "借阅核心的本地事务不覆盖规则服务与 ILS 网关；跨服务失败时如何回滚、如何对账未说明", "严重"],
       ["D04", "关系语义", "AI 类图", "Item 1—0..* Reservation 把预约挂在副本上；多数高校预约是书目级，BR-03 的“该书有预约”语义需澄清", "中等"],
       ["D05", "职责重叠", "AI 方案一", "RenewDomainService 依赖 RuleSet，AppSvc 又依赖 RuleEvaluator；两处都能判定，规则变更易产生两个事实源", "中等"],
       ["D06", "规则冲突", "AI 类图", "RenewRuleSet.allowRenewWhenOverdue() 把 BR-04“已逾期不可续借”变成可配置；配置为 true 会破坏公平性", "中等"],
       ["D07", "跨图不一致", "AI 类图", "组件图有 AuditPort、IdempotencyPort、OutboxPort，类图没有对应类/接口，审计与幂等无落点", "中等"],
       ["D08", "上下文丢失", "AI 类图", "RuleEvaluator.evaluate(hasReservation: boolean) 把预约事实压成布尔值，丢失“谁在排、排第几”，难以解释 BR-03/BR-05", "中等"],
       ["D09", "关系标注", "AI 类图", "书目—副本标“包含”但用普通关联，未标组合；PatronType、ItemStatus、RejectReason 等未定义取值", "轻微"],
       ["D10", "自评偏高", "AI 自评", "方案一自评可维护/需求匹配 5 分，但它只是组件图，缺状态、顺序、数据与界面，不能支撑满分", "轻微"]],
      [1.2, 2.2, 2.2, 8.2, 1.4], "AI 方案缺陷清单")
p("D01 与 D02 是这次最值得记的两条。它们都不是“画错线”，而是“边界与生命周期没想清楚”：聚合根直接持有另一个聚合根，"
  "在内存里很方便，但一旦要独立事务、独立归档就会互相拖累；续借申请压成 0..1，等于把“历史”删掉了，"
  "而审计恰恰需要历史。AI 能把类画全，却不会主动追问“这个 0..1 以后会不会变成多条”。")

h2("14.5  AI 方案中可借鉴的部分")
p("这次 AI 输出不是“反面教材”，而是“可吸收的候选设计”。有五处值得吸收，只是要放到合适的粒度上：")
for line in [
    "端口与适配器的完整清单（IlsPort、AuditPort、IdempotencyPort、OutboxPort）——手工方案已在类图与组件图中采纳。",
    "审计仓储与业务同库同事务——与手工方案的 Outbox 思路一致，可直接复用。",
    "独立的对账入口与对账用例——手工方案有 PENDING_RECONCILE，可补一个显式的 ReconcileApplicationService。",
    "规则服务按版本求值——当前不拆服务，但可把 RuleEvaluator 作为领域内的独立对象保留。",
    "ILS 网关集中幂等/重试/熔断——当前放在 IlsRestAdapter 内，未来拆服务时可整体迁移。",
]:
    bullet(line)
table(["AI 方案元素", "是否采纳", "理由"],
      [["端口与适配器清单", "采纳", "与手工方案同构，边界清晰"],
       ["Audit + Outbox 同库", "采纳", "满足 C4，且不引入额外中间件"],
       ["独立对账入口", "采纳", "补齐 PENDING_RECONCILE 的运营入口"],
       ["规则服务独立部署", "暂不采纳", "当前规则变更频率低；先保留领域内 RuleEvaluator"],
       ["ILS 网关独立服务", "暂不采纳", "当前 ILS 调用量中等；先放在适配器内，量上来再拆"],
       ["投影同步器", "暂不采纳", "仅在未来拆服务时需要，单体内无投影"]],
      [4.6, 3.2, 7.0], "AI 方案元素取舍")

h2("14.6  最终方案与演进路线")
p("最终选择手工方案：模块化单体 + 六边形 + 同步核心 + Outbox 异步旁路。AI 方案一与它同源，但只到组件层；"
  "手工方案在状态、事务、数据、界面与部署上更完整，因此作为最终设计。AI 方案一中的端口清单、对账入口与审计同库"
  "可以合并进来，作为对手工方案的补强。")
p("演进路线分三步：第一步，保持单体，先把知识检索与运营监控拆成独立进程；第二步，当规则变更频率或 ILS 调用量"
  "增长到单体难以承受时，按 AI 方案二的边界拆出规则服务与 ILS 网关，而不是一次性拆成十个微服务；"
  "第三步，若需要多校区部署，再引入事件总线做跨校区同步。每一步都以“当前痛点”为触发条件，而不是以“架构图好看”为触发条件。")
p("为了让“不退化”可验证，设置三条架构适应度函数：① ArchUnit 依赖测试必须通过；② 续借链路的契约测试必须覆盖"
  "成功、拒绝、超时、重放四条路径；③ 每季度做一次 ILS 故障注入，验证降级与对账。三条都通过，才允许发布。")

h2("14.7  Git 提交历史与互评素材")
p("本实验为个人完成，未组队；按指导书要求，仍以 Git 提交历史与 diff 作为可复核的评审素材。提交顺序为"
  "“手工设计 → AI 原始输出 → 报告”，保证活动 1 先于活动 2/3 入库。")
table(["序号", "commit", "时间", "提交内容", "可评审的 diff"],
      [["1", "31d2866", "2026-11-18", "手工设计：图 1～图 11 的 .puml 与渲染图", "架构分层、类图、状态图、顺序图、数据模型"],
       ["2", "49b8716", "2026-11-18", "AI 原始输出：图 12～图 14 的 .puml 与渲染图", "AI 方案与手工方案的差异"],
       ["3", "本次提交", "2026-11-18", "报告 .docx 与 PDF", "评估结论与最终选择"]],
      [1.2, 3.0, 3.2, 6.0, 4.0], "Git 提交记录")
p("互评建议聚焦三处 diff：① 图 2 与图 14 的类图差异（是否有副本、预约、规则版本）；② 图 1 与图 12 的依赖方向"
  "（领域是否被外部依赖）；③ 表 29 的评分是否有证据支撑。互评记录见附录 C，由同组/同班同学填写后随报告提交。")

# ================================================================ 十五、决策记录与小结
h1("十五、设计决策记录与实验小结")
h2("15.1  设计决策记录（ADR）")
p("设计过程中真正困难的是取舍，而不是画图。表 34 记录六条关键决策，按“问题—选项—决定—后果”组织，"
  "方便答辩时逐条解释。")
table(["编号", "决策问题", "选项", "决定与理由", "后果"],
      [["ADR-01", "体系结构风格", "分层单体 / 模块化单体+六边形 / 微服务 / 无服务器", "模块化单体+六边形：满足本地事务与小团队约束，保留演进接缝", "需要 ArchUnit 守护依赖；未来拆分需额外工作"],
       ["ADR-02", "领域模型与持久化", "领域对象直接映射表 / 仓储映射", "仓储映射：领域层不依赖 SQL 与框架", "需要 DTO/映射代码；但领域可纯单测"],
       ["ADR-03", "ILS 事务边界", "大事务包住 ILS / 本地事务 + Outbox + 对账", "本地事务 + 待对账：承认外部系统的不确定性", "需要定时对账任务与运营查询"],
       ["ADR-04", "规则存放", "硬编码 / 数据库配置 / 配置中心+版本", "配置中心+版本：支持热更新与历史复现", "需配置发布流程与灰度"],
       ["ADR-05", "状态实现", "GoF State 类 / 状态枚举+迁移表", "状态枚举+迁移表：状态需持久化，迁移表易测试", "复杂状态机需额外校验工具"],
       ["ADR-06", "确认交互", "直接执行 / 每次确认 / 歧义时确认", "歧义或首次操作时确认：兼顾安全与效率", "前端需处理待澄清状态"]],
      [1.6, 3.4, 5.0, 5.0, 4.0], "关键设计决策记录")
p("ADR-03 与 ADR-05 是本实验最值得讲的两条。前者承认“外部系统不可靠”是架构的一等事实，而不是异常处理的边角料；"
  "后者承认“状态”首先是业务概念，其次才是代码结构。很多设计失败不是因为不会写 State 模式，而是因为把不该成为状态的东西"
  "建成了状态，或把必须显式化的状态藏进了布尔字段。")

h2("15.2  困难与一次真实的返工")
table(["困难", "表现", "解决办法"],
      [["校验逻辑归属", "初版把四项校验写进 LoanRecord，实体被迫依赖仓储与 ILS", "抽出 ValidatorChain，实体只保留自身不变式；应用层编排"],
       ["ILS 超时语义", "初版把超时当失败，允许重试，存在重复续借风险", "增加 PENDING_RECONCILE 状态与按幂等键回查"],
       ["跨视图一致性", "组件图初版多画了一个 RuleService，类图里没有对应类", "以类图为事实源，删掉组件图中的虚构服务，补映射表"],
       ["状态机与规则重复", "状态图监护条件与校验器各写一套上限判断", "规则只保留在 RenewalPolicy，状态图引用其方法"],
       ["界面确认疲劳", "每次续借都弹确认，读者觉得啰嗦", "仅在歧义/首次操作确认；明确条码时直接执行"]],
      [2.8, 7.0, 5.0], "主要困难与解决办法")
p("一次真实的返工发生在“ILS 超时”的处理上。第一版顺序图把 ILS --> UC : timeout 画成异常，应用服务捕获后返回失败，"
  "并允许用户重试。做失败矩阵时发现：如果 ILS 实际上已经成功，只是响应超时，重试会再顺延 30 天；"
  "而预约队列的位次不会因为这次多余顺延而恢复。这不是“异常没捕获”，而是对“超时”的语义理解错了——"
  "超时是“结果未知”，不是“失败”。")
p("修正方案是把“待对账”建为显式状态，并在顺序图中增加“按幂等键回查”的消息；对账任务在后台把未知收敛为成功或失败。"
  "这个改动同时影响了状态图、顺序图、数据模型（新增 ils_receipt 与 request_status）和接口（新增查询接口），"
  "也让“4+1 视图必须一起改”这件事变得具体：只改一张图，其他视图就会自相矛盾。")

h2("15.3  心得")
p("第一，系统设计的分水岭是“边界”而不是“图”。类图里多一个类、少一个类，往往不影响系统能否运行；"
  "但事务边界画错、外部系统的不确定性没有位置存放，系统一定会在生产环境出问题。ILS 超时那次返工让我明白，"
  "架构图的价值不在于好看，而在于它迫使你把“不知道怎么办”的情况画出来。")
p("第二，4+1 视图的真正作用是“互相挑错”。逻辑视图说有一个续借申请，进程视图就必须有它的状态与事务；"
  "开发视图说 domain 不依赖框架，物理视图就不能把规则塞进网关。每补一个视图，都会发现上一个视图的漏洞；"
  "这比单独把一张图画到完美更有价值。")
p("第三，AI 是很强的“广度补充”，但不是“边界判断者”。它能在几秒内列出两套备选方案、几种模式、三条演进路线；"
  "但它不会告诉你“这个团队维护不了多服务”“这个业务不能接受最终一致”。评估 AI 的过程，其实是在逼自己"
  "把模糊的工程直觉写成可复核的证据——这也是本实验最大的收获。")

# ================================================================ 附录
page_break()
h1("附录 A　提示词与 AI 原始输出存档")
p("本附录保留与 AI 交互的完整痕迹。V3 全文已见 12.3 节，此处给出 V2 全文、AI 原始输出与处理记录。")

h2("A.1  第二轮提示词（V2）")
quote_block([
    "【角色】你是高校图书馆系统的架构师，熟悉 UML 2.5、DDD 与 PlantUML。",
    "【任务】为 UC-03“续借图书”设计系统架构，给出 PlantUML 组件图与关键类图。",
    "【边界】只设计借阅流通子系统；智能体编排、知识检索、ILS 属外部；数据库/缓存/大模型不是参与者。",
    "【约束】学生最多 2 次、教师 5 次；顺延 30 天；有预约/逾期不可续借；续借不改队列位次；",
    "P95 ≤ 3 秒；审计 ≥ 180 天；ILS 故障可降级；规则可配置；团队 3～5 人。",
    "【输出】两个 PlantUML 代码块；中文命名；关系标注类型与多重性。",
])
p("V2 相对 V1 的改进直接可见：推荐/信用分消失，边界与规则被写入提示词，类图从 3 个类增至 7 个。"
  "但它仍然只给一个方案，也没有要求 AI 说明“为什么不选另一种风格”，这是 V3 加入“两个不同风格 + 不适用场景”的原因。")

h2("A.2  AI 原始输出（方案一，节选）")
p("以下为 DeepSeek 返回的方案一组件图代码，未做任何修改（渲染图见 13.1）。完整文件见随报告提交的 12_AI_arch_A_raw.puml。")
_ai_a = io.open(os.path.join(DIR, "puml", "12_AI_arch_A_raw.puml"), encoding="utf-8").read().rstrip("\n").split("\n")
code_block(_ai_a[:52], size=7.2)

h2("A.3  AI 原始输出（关键类图，节选）")
p("类图原始输出完整内容见 14_AI_class_raw.puml，此处节选领域类与属性部分；关系与多重性见图 14。")
_ai_c = io.open(os.path.join(DIR, "puml", "14_AI_class_raw.puml"), encoding="utf-8").read().rstrip("\n").split("\n")
code_block(_ai_c[:44], size=7.2)

h2("A.4  处理记录与时间线")
table(["序号", "项目", "记录"],
      [["1", "语法修正", "0 处：三份代码块首轮渲染成功，未改任何字符"],
       ["2", "语义修改", "0 处：未调整服务、类、关系，以保证评估对象与生成结果一致"],
       ["3", "渲染参数", "java -jar plantuml.jar -tpng -charset UTF-8，默认 DPI"],
       ["4", "自检表", "AI 随输出附带的表格全文见 13.3，未删改"],
       ["5", "对话轮次", "同一会话连续三轮（V1→V2→V3），中间未清空上下文"]],
      [1.2, 3.0, 10.6], "AI 输出处理记录")
table(["时间", "动作", "说明"],
      [["2026-11-18 20:55", "手工设计开始", "图 1～图 11 及 11 个 .puml"],
       ["2026-11-18 21:10", "手工设计完成", "分层、类图、状态图、顺序图、界面、数据、组件、部署"],
       ["2026-11-18 21:15", "AI 生成对照件（V3）", "DeepSeek 网页版，原始输出见 A.2、A.3"],
       ["2026-11-18 21:20", "手工设计入库（31d2866）", "先于 AI 对照件，作为提交时序证据"],
       ["2026-11-18 21:25", "AI 对照件入库（49b8716）", "活动 3 成果"],
       ["2026-11-18 21:40", "逐项评估", "形成 14.4 的 10 条缺陷"],
       ["2026-11-18 22:00", "报告撰写与排版", "本文"]],
      [4.0, 5.6, 6.0], "活动时序记录")

h1("附录 B　模型源文件与提交记录")
table(["序号", "文件名", "图种", "图号", "说明"],
      [["1", "01_architecture_layered.puml", "体系结构", "图 1", "手工；四层 + 依赖规则"],
       ["2", "02_domain_class.puml", "类图（领域）", "图 2", "手工；7 实体、3 值对象、6 枚举"],
       ["3", "03_app_class.puml", "类图（应用）", "图 3", "手工；BCE + 校验链"],
       ["4", "04_state_loanrecord_design.puml", "状态图（借阅）", "图 4", "手工；6 状态、14 迁移"],
       ["5", "05_state_renewal_request.puml", "状态图（申请）", "图 5", "手工；7 状态、含待对账"],
       ["6", "06_sequence_renew_design.puml", "顺序图", "图 6", "手工；4 条路径"],
       ["7", "07_ui_prototype.puml", "界面原型", "图 7", "手工；Salt 线框"],
       ["8", "08_data_model_core.puml", "数据模型（核心）", "图 8", "手工；5 张表"],
       ["9", "09_data_model_renewal.puml", "数据模型（续借）", "图 9", "手工；5 张表"],
       ["10", "10_component_development.puml", "组件图", "图 10", "手工；8 个模块"],
       ["11", "11_deployment.puml", "部署图", "图 11", "手工；四区 + 外部服务"],
       ["12", "12_AI_arch_A_raw.puml", "组件图（AI）", "图 12", "原始输出，0 处修改"],
       ["13", "13_AI_arch_B_raw.puml", "组件图（AI）", "图 13", "原始输出，0 处修改"],
       ["14", "14_AI_class_raw.puml", "类图（AI）", "图 14", "原始输出，0 处修改"]],
      [1.2, 6.4, 3.2, 1.6, 4.4], "模型源文件清单")
table(["序号", "commit", "时间", "提交内容"],
      [["1", "31d2866", "2026-11-18", "手工设计 11 个 .puml + 11 张渲染图（活动 1，AI 之前）"],
       ["2", "49b8716", "2026-11-18", "AI 对照件 3 个 .puml + 3 张渲染图（活动 3）"],
       ["3", "本次提交", "2026-11-18", "报告 .docx 与 32406300076.pdf"]],
      [1.2, 3.4, 4.2, 7.0], "Git 提交记录")
p("说明：按指导书“提交时序边界”的要求，手工设计必须先于 AI 环节入库；本实验的两次提交按“手工设计在前、"
  "AI 对照件在后”的顺序建立，保留了工作先后。仓库地址：https://github.com/Longqc666/OOProject")

h1("附录 C　同伴互评记录")
p("本实验为个人完成，以下互评记录由同班同学在提交前填写；评审素材为 14.7 节列出的三处 diff。")
table(["评审人", "日期", "关注点", "主要意见", "处理"],
      [["（待填）", "（待填）", "类图差异", "（待填）", "（待填）"],
       ["（待填）", "（待填）", "依赖方向", "（待填）", "（待填）"],
       ["（待填）", "（待填）", "评分证据", "（待填）", "（待填）"]],
      [3.2, 2.6, 3.4, 8.0, 3.0], "同伴互评记录")

page_break()
h1("参考文献")
refs = [
    "[1] 广州大学计算机科学与网络工程学院. 面向对象分析与设计实验指导书（V202610）[Z]. 2026.",
    "[2] OMG. Unified Modeling Language (OMG UML), Version 2.5.1[S]. 2017.",
    "[3] Bass L, Clements P, Kazman R. Software Architecture in Practice[M]. 4th ed. Addison-Wesley, 2021.",
    "[4] Evans E. Domain-Driven Design: Tackling Complexity in the Heart of Software[M]. Addison-Wesley, 2003.",
    "[5] Fowler M. Patterns of Enterprise Application Architecture[M]. Addison-Wesley, 2002.",
    "[6] Newman S. Building Microservices[M]. 2nd ed. O'Reilly, 2021.",
    "[7] PlantUML. PlantUML Language Reference Guide[EB/OL]. https://plantuml.com/, 2026.",
    "[8] Nikiforova O. Teaching UML Modeling in the Age of Generative AI: Combining Manual Modeling, "
    "Prompt Engineering and Model Evaluation[C]//MODELS Companion '26. 2026: 86-95.",
]
for r in refs:
    para = doc.add_paragraph()
    pf = para.paragraph_format
    pf.left_indent = Cm(0.85)
    pf.first_line_indent = Cm(-0.85)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.35
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _f(para.add_run(r), size=9.5)

doc.save(OUT)
print("saved", OUT, "| headings", len(headings), "| figures", _fig[0], "| tables", _tab[0])
