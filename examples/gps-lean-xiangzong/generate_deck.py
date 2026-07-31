#!/usr/bin/env python3
"""向总汇报 6 页重建 — 内容与用户提供的 6 张原稿截图一一对应。"""

from __future__ import annotations

import html
import json
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parent
SLIDES_DIR = ROOT / "slides"
OUT_DIR = ROOT / "output"

# 贴近原稿：白底 + 橙色强调 + 深蓝标题（避免紫/奶油陶土/报章风）
NAVY = "#1B3A5F"
NAVY_DEEP = "#0F2740"
ORANGE = "#E87A2E"
ORANGE_SOFT = "#FFF4EB"
ORANGE_LINE = "#F0C9A8"
BG = "#FFFFFF"
CARD = "#FFFFFF"
TEXT = "#2C3E50"
MUTED = "#5D6D7E"
LINE = "#E8EEF2"
SOFT_BG = "#F7F9FB"

FONT = "Source Han Sans SC, PingFang SC, Microsoft YaHei, Noto Sans SC, sans-serif"


def esc(s: str) -> str:
    return html.escape(s)


def wrap(text_in: str, max_chars: int) -> list[str]:
    lines, cur = [], ""
    for ch in text_in:
        cur += ch
        if len(cur) >= max_chars and ch in "，。；、）」》 ":
            lines.append(cur.strip())
            cur = ""
        elif len(cur) >= max_chars + 3:
            lines.append(cur)
            cur = ""
    if cur.strip():
        lines.append(cur.strip())
    return lines or [text_in]


def svg_wrap(body: str, bg: str = BG) -> str:
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <defs>
    <linearGradient id="navyGrad" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="{NAVY_DEEP}"/>
      <stop offset="100%" stop-color="{NAVY}"/>
    </linearGradient>
    <filter id="sh" x="-4%" y="-4%" width="108%" height="116%">
      <feDropShadow dx="0" dy="1" stdDeviation="3" flood-color="#1B3A5F" flood-opacity="0.08"/>
    </filter>
  </defs>
  <rect width="1280" height="720" fill="{bg}"/>
{body}
</svg>
'''


def T(x, y, content, size=16, fill=TEXT, weight=500, anchor="start"):
    return (
        f'<text x="{x}" y="{y}" fill="{fill}" font-size="{size}" font-weight="{weight}" '
        f'font-family="{FONT}" text-anchor="{anchor}">{esc(content)}</text>'
    )


def TB(x, y, lines, size=13, fill=TEXT, weight=400, lh=1.45):
    parts = [
        f'<text x="{x}" y="{y}" fill="{fill}" font-size="{size}" font-weight="{weight}" '
        f'font-family="{FONT}">'
    ]
    for i, line in enumerate(lines):
        dy = 0 if i == 0 else int(size * lh)
        parts.append(f'<tspan x="{x}" dy="{dy}">{esc(line)}</tspan>')
    parts.append("</text>")
    return "\n".join(parts)


def R(x, y, w, h, fill=CARD, stroke=None, r=12, shadow=False):
    s = f' stroke="{stroke}" stroke-width="1.5"' if stroke else ' stroke="none"'
    f = ' filter="url(#sh)"' if shadow else ""
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" ry="{r}" fill="{fill}"{s}{f}/>'


# ───────────────── Slide 1: 痛点（对应图1） ─────────────────

def slide_01():
    pains = [
        ("01", "落地不均衡，标杆与常态两极分化",
         "标杆产线成果显著，但多数产线仍停留在表层 5S 整改；深层次浪费与工艺瓶颈未触及，存在「重展示、轻落地、难持续」的形式化问题。"),
        ("02", "改善偏单点，系统性攻坚能力不足",
         "改善多为碎片化整改，对拉晶高能耗、切片断线损耗等核心技术痛点攻坚不够，难以从根本上拉动良率与成本等关键指标。"),
        ("03", "全员氛围不足，自主改善意识薄弱",
         "推进仍依赖管理推动与专班督导，一线员工多为「被动执行」模式，尚未形成「主动找浪费、主动提改善、主动创价值」的精益文化。"),
        ("04", "机制落地不彻底，考核未闭环",
         "精益落地与人才晋升、基地评价挂钩常流于形式，缺乏量化指标支撑与清晰奖惩机制，导致管理关注度不一、基层执行驱动力不足。"),
        ("05", "精益信息化缺失，改善难沉淀",
         "改善经验难沉淀、难复制，数据与现场脱节，缺少统一数字化看板与闭环系统，难以支撑精益从「运动式」走向「常态化」。"),
    ]
    # 左列 01-03，右列 04-05 + 总结
    left = pains[:3]
    right = pains[3:]
    body = f'''
  {T(56, 52, "当前精益推进现存真实痛点", 30, NAVY, 800)}
  {R(56, 64, 80, 4, ORANGE, r=2)}
'''
    for i, (num, title, desc) in enumerate(left):
        y = 95 + i * 195
        body += f'''
  {R(56, y, 580, 175, CARD, ORANGE_LINE, shadow=True)}
  {R(56, y, 8, 175, ORANGE, r=0)}
  {T(84, y + 38, num, 22, ORANGE, 800)}
  {T(84, y + 72, title, 17, NAVY, 700)}
  {TB(84, y + 102, wrap(desc, 30), 13, MUTED, 400, 1.45)}
'''
    for i, (num, title, desc) in enumerate(right):
        y = 95 + i * 195
        body += f'''
  {R(660, y, 560, 175, CARD, ORANGE_LINE, shadow=True)}
  {R(660, y, 8, 175, ORANGE, r=0)}
  {T(688, y + 38, num, 22, ORANGE, 800)}
  {T(688, y + 72, title, 17, NAVY, 700)}
  {TB(688, y + 102, wrap(desc, 28), 13, MUTED, 400, 1.45)}
'''
    body += f'''
  {R(660, 485, 560, 175, SOFT_BG, ORANGE_LINE, shadow=True)}
  {TB(688, 525, wrap("精益生产本质是消除浪费与持续改善；实践中仍面临从表层整改到深层次攻坚的多重执行挑战。", 28), 13, TEXT, 500, 1.45)}
  {TB(688, 595, wrap("闭环机制与数据沉淀是精益从「运动式」走向「常态化」的关键，破除形式主义、回归价值创造。", 28), 13, MUTED, 400, 1.45)}
'''
    return svg_wrap(body)


# ───────────────── Slide 2: 筑基→势起（对应图2） ─────────────────

def slide_02():
    body = f'''
  {T(56, 42, "筑基 →「势起」", 28, NAVY, 800)}
  {R(56, 52, 64, 4, ORANGE, r=2)}
  {TB(56, 82, [
      "自 2025 年 9 月以来，GPS 精益体系已从「探索与单点整改」迈入「体系起势与机制落地」阶段。",
      "2026 年为关键推进年：精益作为「第二生产力」，以「453」框架启航精益转型，践行「真善美」价值主张。"
  ], 13, MUTED, 400, 1.4)}

  {T(56, 145, "四级架构", 15, ORANGE, 700)}
'''
    levels = [
        ("决策层 · 战略引领", "确立精益战略，审批重大课题，提供核心资源与政策支持"),
        ("推进层 · 体系赋能", "体系规划、方法导入、项目督导与专业人才培养"),
        ("执行层 · 攻坚落地", "基地一把手主责，目标分解与资源协同，确保课题出成果"),
        ("落地层 · 现场改善", "班组长带一线严守 SOP，参与现场改善提案"),
    ]
    for i, (t, d) in enumerate(levels):
        x = 56 + i * 305
        body += f'''
  {R(x, 160, 290, 100, CARD, LINE, shadow=True)}
  {R(x, 160, 290, 34, NAVY, r=12)}
  <rect x="{x}" y="182" width="290" height="12" fill="{NAVY}"/>
  {T(x + 145, 182, t, 13, "#fff", 700, "middle")}
  {TB(x + 14, 214, wrap(d, 14), 12, MUTED, 400, 1.35)}
'''
    body += f'  {T(56, 290, "五大体系", 15, ORANGE, 700)}\n'
    systems = [
        ("01 对标研学", "被动学习 → 主动对标"),
        ("02 标杆复制", "单点优势 → 全域标准化"),
        ("03 人才育成", "基础普及 → 精英赋能"),
        ("04 考核管控", "柔性要求 → 刚性闭环"),
        ("05 价值生产力", "精益锚定核心课题"),
    ]
    for i, (t, d) in enumerate(systems):
        x = 56 + i * 244
        body += f'''
  {R(x, 305, 230, 78, ORANGE_SOFT if i == 4 else CARD, ORANGE if i == 4 else LINE, shadow=True)}
  {T(x + 16, 335, t, 14, NAVY, 700)}
  {T(x + 16, 360, d, 12, MUTED, 400)}
'''
    body += f'  {T(56, 415, "三阶段路径（2026 下半年）", 15, ORANGE, 700)}\n'
    phases = [
        ("7–8 月 · 筑基巩固期", "补短板、定标准、全员培训，夯实现场基础管理规范化"),
        ("9–10 月 · 攻坚突破期", "专项改善课题，解决核心经营痛点与瓶颈"),
        ("11–12 月 · 价值兑现期", "固化成果、建立长效机制，验证经济效益与效率提升"),
    ]
    for i, (t, d) in enumerate(phases):
        x = 56 + i * 405
        body += f'''
  {R(x, 430, 390, 110, CARD, LINE, shadow=True)}
  {T(x + 18, 465, t, 15, NAVY, 700)}
  {TB(x + 18, 495, wrap(d, 20), 12, MUTED, 400, 1.4)}
'''
    body += f'''
  {R(56, 560, 1168, 110, NAVY_DEEP, r=14)}
  {T(84, 605, "453 框架 = 四级架构 × 五大体系 × 三阶段路径", 18, ORANGE, 700)}
  {T(84, 640, "以体系起势驱动机制落地，打造支撑公司高质量发展的「第二生产力」", 14, "#A8C5D4", 400)}
'''
    return svg_wrap(body)


# ───────────────── Slide 3: 五大体系 01-04（对应图3） ─────────────────

def slide_03():
    blocks = [
        ("01", "对标研学体系", "从「被动学习」到「对标赶超」",
         ["三维对标：构建「横向内部基地 + 纵向工序 + 制造标杆」网络，识别绩效差距。",
          "量化整改：聚焦能耗、良率、OEE 等核心 KPI，形成月度差距清单与整改计划，确保改善见效。"]),
        ("02", "标杆复制体系", "从「单点标杆」到「全域标准化」",
         ["标准化拆解：将标杆车间/产线成功经验沉淀为 SOP 与点检标准，跨工序、跨产线复制。",
          "标杆升级：从基础 5S 走向提质增效、降本与零缺陷，打造行业领先标杆。"]),
        ("03", "人才育成体系", "从「基础普及」到「精英赋能」",
         ["三级梯队：精益骨干（执行层）+ 内训师（传播层）+ 精益专班（决策层）。",
          "利益深度绑定：精益认证与职业晋升通道、绩效薪酬激励挂钩，形成「学会—会用—受益」正循环。"]),
        ("04", "考核管控体系", "从「柔性要求」到「刚性闭环」",
         ["制度化考核：精益改善指标全面纳入公司正式考核，对基地负责人与核心干部量化评价，避免形式主义。",
          "红黑榜：部门负责人作为第一责任人，公开红黑榜展示改善结果，奖惩清晰、闭环管理。"]),
    ]
    body = f'''
  {T(56, 42, "五大体系", 28, NAVY, 800)}
  {R(56, 52, 64, 4, ORANGE, r=2)}
  {T(56, 82, "管理支撑系统：对标 · 复制 · 育成 · 考核", 14, MUTED, 400)}
'''
    positions = [(56, 110), (668, 110), (56, 400), (668, 400)]
    for (num, title, sub, bullets), (x, y) in zip(blocks, positions):
        body += f'''
  {R(x, y, 556, 265, CARD, LINE, shadow=True)}
  {R(x, y, 556, 72, NAVY, r=12)}
  <rect x="{x}" y="{y+52}" width="556" height="20" fill="{NAVY}"/>
  {T(x + 24, y + 32, f"{num}  {title}", 18, "#fff", 800)}
  {T(x + 24, y + 58, sub, 12, ORANGE, 600)}
'''
        yy = y + 105
        for b in bullets:
            body += TB(x + 24, yy, wrap("· " + b, 28), 13, MUTED, 400, 1.4)
            yy += 16 + 13 * 1.4 * len(wrap(b, 28))
    return svg_wrap(body)


# ───────────────── Slide 4: 价值生产力（对应图4） ─────────────────

def slide_04():
    topics = [
        ("拉晶硅损优化", "厘清硅损价值链浪费点；降低切割与磨削余量；提升硅料来源利用效率。"),
        ("切片损耗管控", "攻克断线、碎片难点；提升硅片切割良率与优品率。"),
        ("组件质量攻坚", "管控层压缺陷、隐裂、虚焊及电池片碎片率；保障产品可靠性与寿命。"),
        ("设备效能跃升", "全面夯实 TPM 运行模式；提升 OEE 指标；保障设备与产线连续稳定运行。"),
        ("精益数字化+", "以标准化为基础，全局拉通顶层数字化模型；整合 DMS、TPMS、QMS 等系统。"),
    ]
    body = f'''
  {T(56, 40, "五大体系", 22, NAVY, 800)}
  {T(56, 72, "＞ 价值生产力体系：以精益锚定核心课题、助力价值增长", 16, ORANGE, 700)}
  {R(56, 84, 64, 3, ORANGE, r=2)}

  {R(56, 110, 360, 430, "url(#navyGrad)", r=16)}
  {T(84, 170, "05", 36, ORANGE, 800)}
  {T(84, 220, "聚焦五大", 26, "#fff", 800)}
  {T(84, 258, "核心课题", 26, "#fff", 800)}
  {T(84, 300, "突破生产瓶颈", 16, "#A8C5D4", 500)}
  {TB(84, 360, wrap("推动精益改善转化为产能增量、良率增量、效益增量，打造「第二生产力」引擎。", 16), 13, "#A8C5D4", 400, 1.45)}
'''
    for i, (title, desc) in enumerate(topics):
        y = 110 + i * 86
        body += f'''
  {R(440, y, 784, 76, CARD, LINE, shadow=True)}
  {R(440, y, 8, 76, ORANGE, r=0)}
  {T(468, y + 30, f"0{i+1}  {title}", 17, NAVY, 800)}
  {T(468, y + 56, desc, 12, MUTED, 400)}
'''
    body += f'''
  {R(56, 560, 1168, 110, ORANGE_SOFT, ORANGE, r=14)}
  {T(84, 605, "＞ 实现管理与产能双增长", 16, ORANGE, 700)}
  {TB(84, 635, ["推动精益改善转化为实实在在的 产能增量、良率增量、效益增量，打造支撑公司业务规模化、高质量发展的「第二生产力」引擎"], 13, TEXT, 500, 1.4)}
'''
    return svg_wrap(body)


# ───────────────── Slide 5: 三个阶段（对应图5） ─────────────────

def slide_05():
    body = f'''
  {T(56, 38, "三个阶段", 26, NAVY, 800)}
  {R(56, 48, 64, 4, ORANGE, r=2)}

  {R(56, 70, 570, 78, CARD, LINE, shadow=True)}
  {T(76, 100, "2026 上半年｜筑基起势（0→1）", 15, NAVY, 700)}
  {T(76, 128, "建体系、树标杆、统一组织认知", 12, MUTED, 400)}
  {R(654, 70, 570, 78, NAVY, r=12)}
  {T(674, 100, "2026 下半年｜全域放大与价值兑现（1→N）", 15, "#fff", 700)}
  {T(674, 128, "标杆复制、痛点攻坚、机制固化、产能效益兑现", 12, "#A8C5D4", 400)}

  {T(56, 175, "实施路径：三阶段推进，层层递进", 14, ORANGE, 700)}
'''
    phases = [
        ("01", "7–8 月", "标杆认证 · 统一执行标准",
         ["标杆车间/产线认证；8 月明确复制进度，全工序夯实生产管理基础。",
          "系统输出工具方法论；识别各基地短板，形成各工序精益标准手册。",
          "启动组件 DMS，实现 Andon 在线响应与问题闭环。"],
         "统一执行标准，消除产线基础差距，实现标准化作业"),
        ("02", "9–10 月", "专项攻坚 · 文化激活",
         ["拉晶专项：降硅损、断线、碎片与能耗。",
          "加速组件「改善周」；黑带人才聚焦日常管理、TPM 与品质改善。",
          "落地人才赋能体系；绩效与改善成果、人才培养硬挂钩。"],
         "攻克生产瓶颈，激发组织内生动力，形成全员参与氛围"),
        ("03", "11–12 月", "成果固化 · 价值兑现",
         ["开展半年度精益人才认证与项目成果发布。",
          "复盘全年精益改善成果，建设标准体系与最佳实践案例库。",
          "量化降本、提质、增效的经济价值，完成价值闭环。"],
         "形成长效精益管理机制，管理成果固化、价值清晰可见"),
    ]
    for i, (num, time, theme, bullets, goal) in enumerate(phases):
        x = 56 + i * 405
        head = ORANGE if i == 1 else NAVY
        body += f'''
  {R(x, 195, 390, 465, CARD, LINE, shadow=True)}
  {R(x, 195, 390, 88, head, r=12)}
  <rect x="{x}" y="263" width="390" height="20" fill="{head}"/>
  {T(x + 20, 230, f"{num}  {time}", 13, "#fff" if i == 1 else ORANGE, 700)}
  {T(x + 20, 262, theme, 16, "#fff", 800)}
'''
        yy = 310
        for b in bullets:
            lines = wrap("· " + b, 18)
            body += TB(x + 18, yy, lines, 12, MUTED, 400, 1.35)
            yy += 10 + int(12 * 1.35 * len(lines))
        body += f'''
  {R(x + 14, 560, 362, 80, ORANGE_SOFT, r=10)}
  {TB(x + 26, 588, wrap("目标：" + goal, 18), 12, NAVY, 600, 1.35)}
'''
    return svg_wrap(body)


# ───────────────── Slide 6: 真善美（对应图6） ─────────────────

def slide_06():
    cols = [
        ("守「真」", "数据驱动、过程可控",
         "拒绝「假」改善，以真实数据与追求极致为改善基石。",
         ["全流程可追溯：损耗/OEE 实时采集，量化收益、精准定位问题。",
          "SPC 底线：工艺参数设限，异常自动预警。",
          "根因分析：杜绝经验拍脑袋，以真实缺陷为依据，拒绝主观归因。"],
         "可信改善", "客观数据支撑，拒绝假成果，让改善可见"),
        ("行「善」", "守正创新、系统攻坚",
         "立足生产实际，创造能解决真问题的有效创新。",
         ["夯实基础：推进 TPM、SMED；吃透工艺逻辑，闭环到一线。",
          "攻坚痛点：聚焦拉晶、切片、组件核心瓶颈，全员改善。",
          "一地一策：适配不同基地特征定制策略。"],
         "有效改善", "系统突破核心痛点，让改善有价值"),
        ("求「美」", "极致现场、零缺陷品质",
         "持续精进，追求现场管理与产品品质的双重极致。",
         ["现场之美：全域 6S 与目视化，消除「隐形浪费」。",
          "品质之美：防错、巡检、EL 检测，严防不良流出。",
          "零缺陷：严控批次稳定性，降低客诉损失。"],
         "长效改善", "成果标准化、持续精进，让改善成日常"),
    ]
    body = f'''
  {T(56, 38, "精益全域落地｜以「真善美」锚定改善底层逻辑", 22, NAVY, 800)}
  {R(56, 48, 64, 4, ORANGE, r=2)}
  {T(56, 78, "所有精益课题、现场改善、体系运行统一对标公司核心价值主张，构建可信、有效、长效的改善闭环", 12, MUTED, 400)}
'''
    for i, (title, defn, lead, bullets, outcome, foot) in enumerate(cols):
        x = 56 + i * 405
        head = NAVY if i != 1 else "#0D8A7A"
        soft = ORANGE_SOFT if i != 1 else "#E8F6F3"
        accent = ORANGE if i != 1 else "#0D8A7A"
        body += f'''
  {R(x, 100, 390, 560, CARD, LINE, shadow=True)}
  {R(x, 100, 390, 100, head, r=12)}
  <rect x="{x}" y="180" width="390" height="20" fill="{head}"/>
  {T(x + 22, 140, title, 26, "#fff", 800)}
  {T(x + 22, 175, defn, 13, "#A8C5D4", 500)}
  {TB(x + 20, 230, wrap(lead, 20), 12, TEXT, 500, 1.35)}
'''
        yy = 290
        for b in bullets:
            lines = wrap("· " + b, 18)
            body += TB(x + 18, yy, lines, 12, MUTED, 400, 1.35)
            yy += 8 + int(12 * 1.35 * len(lines))
        body += f'''
  {R(x + 14, 545, 362, 95, soft, r=10)}
  {T(x + 28, 580, outcome, 15, accent, 800)}
  {TB(x + 28, 608, wrap(foot, 18), 12, NAVY, 500, 1.35)}
'''
    return svg_wrap(body)


SLIDES = [
    ("slide-01.svg", "当前精益推进现存真实痛点", slide_01),
    ("slide-02.svg", "筑基→势起｜453总框架", slide_02),
    ("slide-03.svg", "五大体系｜01–04", slide_03),
    ("slide-04.svg", "价值生产力体系｜五大核心课题", slide_04),
    ("slide-05.svg", "三个阶段", slide_05),
    ("slide-06.svg", "真善美底层逻辑", slide_06),
]


# ───────────────── PPTX（6 页，内容同截图） ─────────────────

def set_font(run, size=12, bold=False, color=RGBColor(0x2C, 0x3E, 0x50)):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Microsoft YaHei"
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
    if ea is None:
        ea = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
    ea.set("typeface", "Microsoft YaHei")


def box(slide, l, t, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1.25)
    try:
        sh.adjustments[0] = 0.08
    except Exception:
        pass
    return sh


def tb(slide, l, t, w, h, text, size=14, bold=False, color=RGBColor(0x2C, 0x3E, 0x50), align=PP_ALIGN.LEFT):
    b = slide.shapes.add_textbox(l, t, w, h)
    tf = b.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_font(run, size=size, bold=bold, color=color)
    return b


def build_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    N = RGBColor(0x1B, 0x3A, 0x5F)
    O = RGBColor(0xE8, 0x7A, 0x2E)
    W = RGBColor(0xFF, 0xFF, 0xFF)
    M = RGBColor(0x5D, 0x6D, 0x7E)
    TCOL = RGBColor(0x2C, 0x3E, 0x50)
    L = RGBColor(0xE8, 0xEE, 0xF2)
    Soft = RGBColor(0xFF, 0xF4, 0xEB)
    BG_c = RGBColor(0xFF, 0xFF, 0xFF)
    Teal = RGBColor(0x0D, 0x8A, 0x7A)

    # --- P1 痛点 ---
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c)
    tb(s, Inches(0.45), Inches(0.25), Inches(12), Inches(0.45), "当前精益推进现存真实痛点", 26, True, N)
    pains = [
        ("01 落地不均衡，标杆与常态两极分化",
         "标杆产线成果显著，但多数产线仍停留在表层5S整改；深层次浪费与工艺瓶颈未触及，存在「重展示、轻落地、难持续」的形式化问题。"),
        ("02 改善偏单点，系统性攻坚能力不足",
         "改善多为碎片化整改，对拉晶高能耗、切片断线损耗等核心技术痛点攻坚不够，难以从根本上拉动良率与成本等关键指标。"),
        ("03 全员氛围不足，自主改善意识薄弱",
         "推进仍依赖管理推动与专班督导，一线员工多为「被动执行」模式，尚未形成「主动找浪费、主动提改善、主动创价值」的精益文化。"),
        ("04 机制落地不彻底，考核未闭环",
         "精益落地与人才晋升、基地评价挂钩常流于形式，缺乏量化指标支撑与清晰奖惩机制，导致管理关注度不一、基层执行驱动力不足。"),
        ("05 精益信息化缺失，改善难沉淀",
         "改善经验难沉淀、难复制，数据与现场脱节，缺少统一数字化看板与闭环系统，难以支撑精益从「运动式」走向「常态化」。"),
    ]
    for i, (title, desc) in enumerate(pains):
        col, row = (0, i) if i < 3 else (1, i - 3)
        x = Inches(0.4 + col * 6.45)
        y = Inches(0.9 + row * 1.9)
        box(s, x, y, Inches(6.2), Inches(1.75), W, L)
        tb(s, x + Inches(0.2), y + Inches(0.15), Inches(5.8), Inches(0.35), title, 14, True, O)
        tb(s, x + Inches(0.2), y + Inches(0.55), Inches(5.8), Inches(1.05), desc, 12, False, M)
    box(s, Inches(6.85), Inches(4.7), Inches(6.2), Inches(2.35), Soft, O)
    tb(s, Inches(7.1), Inches(4.95), Inches(5.7), Inches(1.9),
       "精益生产本质是消除浪费与持续改善；实践中仍面临从表层整改到深层次攻坚的多重执行挑战。\n"
       "闭环机制与数据沉淀是精益从「运动式」走向「常态化」的关键，破除形式主义、回归价值创造。",
       12, False, TCOL)

    # --- P2 筑基势起 ---
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c)
    tb(s, Inches(0.45), Inches(0.2), Inches(12), Inches(0.4), "筑基 →「势起」", 24, True, N)
    tb(s, Inches(0.45), Inches(0.65), Inches(12.4), Inches(0.7),
       "自2025年9月以来，GPS精益体系已从「探索与单点整改」迈入「体系起势与机制落地」阶段。"
       "2026年为关键推进年：精益作为「第二生产力」，以「453」框架启航精益转型，践行「真善美」价值主张。",
       12, False, M)
    tb(s, Inches(0.45), Inches(1.4), Inches(4), Inches(0.3), "四级架构", 13, True, O)
    for i, (a, b) in enumerate([
        ("决策层·战略引领", "确立战略/审批课题/资源政策支持"),
        ("推进层·体系赋能", "体系规划/方法导入/督导与育人"),
        ("执行层·攻坚落地", "基地主责/目标分解/资源协同"),
        ("落地层·现场改善", "班组SOP/现场提案改善"),
    ]):
        x = Inches(0.45 + i * 3.2)
        box(s, x, Inches(1.75), Inches(3.05), Inches(1.15), W, L)
        tb(s, x + Inches(0.12), Inches(1.9), Inches(2.8), Inches(0.3), a, 13, True, N)
        tb(s, x + Inches(0.12), Inches(2.3), Inches(2.8), Inches(0.45), b, 11, False, M)
    tb(s, Inches(0.45), Inches(3.15), Inches(4), Inches(0.3), "五大体系", 13, True, O)
    for i, name in enumerate(["01对标研学", "02标杆复制", "03人才育成", "04考核管控", "05价值生产力"]):
        x = Inches(0.45 + i * 2.55)
        box(s, x, Inches(3.5), Inches(2.4), Inches(0.7), Soft if i == 4 else W, O if i == 4 else L)
        tb(s, x, Inches(3.65), Inches(2.4), Inches(0.4), name, 13, True, N, PP_ALIGN.CENTER)
    tb(s, Inches(0.45), Inches(4.45), Inches(8), Inches(0.3), "三阶段路径（2026下半年）", 13, True, O)
    for i, (t, d) in enumerate([
        ("7–8月 筑基巩固期", "补短板、定标准、全员培训"),
        ("9–10月 攻坚突破期", "专项课题，解决核心痛点瓶颈"),
        ("11–12月 价值兑现期", "固化成果、长效机制、效益验证"),
    ]):
        x = Inches(0.45 + i * 4.25)
        box(s, x, Inches(4.85), Inches(4.05), Inches(1.2), W, L)
        tb(s, x + Inches(0.15), Inches(5.0), Inches(3.7), Inches(0.3), t, 13, True, N)
        tb(s, x + Inches(0.15), Inches(5.4), Inches(3.7), Inches(0.4), d, 12, False, M)
    box(s, Inches(0.45), Inches(6.3), Inches(12.4), Inches(0.85), N)
    tb(s, Inches(0.7), Inches(6.5), Inches(12), Inches(0.5),
       "453框架 = 四级架构 × 五大体系 × 三阶段路径　｜　打造「第二生产力」", 14, True, W)

    # --- P3 五大体系01-04 ---
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c)
    tb(s, Inches(0.45), Inches(0.2), Inches(12), Inches(0.4), "五大体系", 24, True, N)
    tb(s, Inches(0.45), Inches(0.6), Inches(12), Inches(0.3), "管理支撑系统：对标 · 复制 · 育成 · 考核", 12, False, M)
    details = [
        ("01 对标研学体系", "从「被动学习」到「对标赶超」",
         "三维对标：横向内部基地+纵向工序+制造标杆，识别绩效差距。\n量化整改：聚焦能耗、良率、OEE等核心KPI，月度差距清单与整改计划。"),
        ("02 标杆复制体系", "从「单点标杆」到「全域标准化」",
         "标准化拆解：标杆车间/产线经验沉淀为SOP与点检标准，跨线复制。\n标杆升级：从基础5S走向提质增效、降本与零缺陷，打造行业领先标杆。"),
        ("03 人才育成体系", "从「基础普及」到「精英赋能」",
         "三级梯队：精益骨干（执行）+内训师（传播）+精益专班（决策）。\n利益绑定：认证与晋升通道、绩效激励挂钩，形成学会—会用—受益正循环。"),
        ("04 考核管控体系", "从「柔性要求」到「刚性闭环」",
         "制度化考核：精益指标纳入正式考核，基地与核心干部量化评价。\n红黑榜：部门负责人主责，公开展示改善结果，奖惩清晰、闭环管理。"),
    ]
    for i, (title, sub, body_txt) in enumerate(details):
        col, row = i % 2, i // 2
        x, y = Inches(0.4 + col * 6.5), Inches(1.05 + row * 3.05)
        box(s, x, y, Inches(6.25), Inches(2.85), W, L)
        box(s, x, y, Inches(6.25), Inches(0.85), N)
        tb(s, x + Inches(0.2), y + Inches(0.15), Inches(5.8), Inches(0.3), title, 15, True, W)
        tb(s, x + Inches(0.2), y + Inches(0.48), Inches(5.8), Inches(0.3), sub, 11, False, O)
        tb(s, x + Inches(0.2), y + Inches(1.1), Inches(5.8), Inches(1.5), body_txt, 12, False, M)

    # --- P4 价值生产力 ---
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c)
    tb(s, Inches(0.45), Inches(0.2), Inches(12), Inches(0.35), "五大体系", 20, True, N)
    tb(s, Inches(0.45), Inches(0.55), Inches(12), Inches(0.35),
       "＞ 价值生产力体系：以精益锚定核心课题、助力价值增长", 14, True, O)
    box(s, Inches(0.4), Inches(1.05), Inches(4.0), Inches(4.7), N)
    tb(s, Inches(0.7), Inches(1.5), Inches(3.4), Inches(0.4), "05", 28, True, O)
    tb(s, Inches(0.7), Inches(2.2), Inches(3.4), Inches(1.0), "聚焦五大\n核心课题", 24, True, W)
    tb(s, Inches(0.7), Inches(3.5), Inches(3.4), Inches(0.4), "突破生产瓶颈", 14, False, RGBColor(0xA8, 0xC5, 0xD4))
    tb(s, Inches(0.7), Inches(4.2), Inches(3.4), Inches(1.2),
       "推动精益改善转化为产能增量、良率增量、效益增量，打造「第二生产力」引擎。", 12, False, RGBColor(0xA8, 0xC5, 0xD4))
    topics = [
        "01 拉晶硅损优化 — 厘清硅损价值链浪费点；降低切割与磨削余量；提升硅料利用效率。",
        "02 切片损耗管控 — 攻克断线、碎片难点；提升硅片切割良率与优品率。",
        "03 组件质量攻坚 — 管控层压缺陷、隐裂、虚焊及碎片率；保障可靠性与寿命。",
        "04 设备效能跃升 — 夯实TPM运行模式；提升OEE；保障设备与产线连续稳定运行。",
        "05 精益数字化+ — 以标准化为基础，拉通顶层数字化模型；整合DMS、TPMS、QMS。",
    ]
    for i, line in enumerate(topics):
        y = Inches(1.05 + i * 0.95)
        box(s, Inches(4.65), y, Inches(8.25), Inches(0.85), W, L)
        tb(s, Inches(4.9), y + Inches(0.22), Inches(7.8), Inches(0.5), line, 12, True, N)
    box(s, Inches(0.4), Inches(6.0), Inches(12.5), Inches(1.15), Soft, O)
    tb(s, Inches(0.7), Inches(6.2), Inches(12), Inches(0.3), "＞ 实现管理与产能双增长", 14, True, O)
    tb(s, Inches(0.7), Inches(6.55), Inches(12), Inches(0.45),
       "推动精益改善转化为实实在在的产能增量、良率增量、效益增量，打造支撑公司业务规模化、高质量发展的「第二生产力」引擎。",
       12, False, TCOL)

    # --- P5 三阶段 ---
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c)
    tb(s, Inches(0.45), Inches(0.15), Inches(12), Inches(0.35), "三个阶段", 22, True, N)
    box(s, Inches(0.4), Inches(0.6), Inches(6.1), Inches(0.85), W, L)
    tb(s, Inches(0.6), Inches(0.75), Inches(5.7), Inches(0.3), "2026上半年｜筑基起势（0→1）", 13, True, N)
    tb(s, Inches(0.6), Inches(1.1), Inches(5.7), Inches(0.25), "建体系、树标杆、统一组织认知", 11, False, M)
    box(s, Inches(6.75), Inches(0.6), Inches(6.1), Inches(0.85), N)
    tb(s, Inches(6.95), Inches(0.75), Inches(5.7), Inches(0.3), "2026下半年｜全域放大与价值兑现（1→N）", 13, True, W)
    tb(s, Inches(6.95), Inches(1.1), Inches(5.7), Inches(0.25), "标杆复制、痛点攻坚、机制固化、产能效益兑现", 11, False, RGBColor(0xA8, 0xC5, 0xD4))
    tb(s, Inches(0.45), Inches(1.6), Inches(10), Inches(0.3), "实施路径：三阶段推进，层层递进", 12, True, O)
    stages = [
        ("01  7–8月", "标杆认证 · 统一执行标准",
         "· 标杆车间/产线认证；8月明确复制进度，全工序夯实管理基础。\n"
         "· 系统输出工具方法论；识别各基地短板，形成工序精益标准手册。\n"
         "· 启动组件DMS，实现Andon在线响应与问题闭环。",
         "目标：统一执行标准，消除产线基础差距，实现标准化作业"),
        ("02  9–10月", "专项攻坚 · 文化激活",
         "· 拉晶专项：降硅损、断线、碎片与能耗。\n"
         "· 加速组件「改善周」；黑带人才聚焦日常管理、TPM与品质改善。\n"
         "· 落地人才赋能体系；绩效与改善成果、人才培养硬挂钩。",
         "目标：攻克生产瓶颈，激发组织内生动力，形成全员参与氛围"),
        ("03  11–12月", "成果固化 · 价值兑现",
         "· 开展半年度精益人才认证与项目成果发布。\n"
         "· 复盘全年精益改善成果，建设标准体系与最佳实践案例库。\n"
         "· 量化降本、提质、增效的经济价值，完成价值闭环。",
         "目标：形成长效精益管理机制，管理成果固化、价值清晰可见"),
    ]
    for i, (t, theme, acts, goal) in enumerate(stages):
        x = Inches(0.4 + i * 4.3)
        box(s, x, Inches(2.0), Inches(4.1), Inches(5.1), W, L)
        box(s, x, Inches(2.0), Inches(4.1), Inches(1.15), O if i == 1 else N)
        tb(s, x + Inches(0.2), Inches(2.15), Inches(3.7), Inches(0.3), t, 12, True, W)
        tb(s, x + Inches(0.2), Inches(2.55), Inches(3.7), Inches(0.35), theme, 14, True, W)
        tb(s, x + Inches(0.2), Inches(3.4), Inches(3.7), Inches(2.4), acts, 11, False, M)
        box(s, x + Inches(0.15), Inches(5.9), Inches(3.8), Inches(1.0), Soft)
        tb(s, x + Inches(0.3), Inches(6.05), Inches(3.5), Inches(0.75), goal, 11, True, N)

    # --- P6 真善美 ---
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c)
    tb(s, Inches(0.4), Inches(0.15), Inches(12.5), Inches(0.35),
       "精益全域落地｜以「真善美」锚定改善底层逻辑", 18, True, N)
    tb(s, Inches(0.4), Inches(0.55), Inches(12.5), Inches(0.35),
       "所有精益课题、现场改善、体系运行统一对标公司核心价值主张，构建可信、有效、长效的改善闭环", 11, False, M)
    cols = [
        (N, "守「真」", "数据驱动、过程可控",
         "全流程可追溯：损耗/OEE实时采集，量化收益。\nSPC底线：工艺参数设限，异常自动预警。\n根因分析：以真实缺陷为依据，拒绝主观归因。",
         "可信改善：客观数据支撑，拒绝假成果，让改善可见"),
        (Teal, "行「善」", "守正创新、系统攻坚",
         "夯实基础：推进TPM、SMED，闭环到一线。\n攻坚痛点：聚焦拉晶、切片、组件核心瓶颈。\n一地一策：适配不同基地特征定制策略。",
         "有效改善：系统突破核心痛点，让改善有价值"),
        (N, "求「美」", "极致现场、零缺陷品质",
         "现场之美：全域6S与目视化，消除隐形浪费。\n品质之美：防错、巡检、EL，严防不良流出。\n零缺陷：严控批次稳定性，降低客诉损失。",
         "长效改善：成果标准化、持续精进，让改善成日常"),
    ]
    for i, (c, title, defn, body_txt, foot) in enumerate(cols):
        x = Inches(0.4 + i * 4.3)
        box(s, x, Inches(1.05), Inches(4.1), Inches(6.0), W, L)
        box(s, x, Inches(1.05), Inches(4.1), Inches(1.2), c)
        tb(s, x + Inches(0.2), Inches(1.25), Inches(3.7), Inches(0.4), title, 20, True, W)
        tb(s, x + Inches(0.2), Inches(1.75), Inches(3.7), Inches(0.3), defn, 12, False, RGBColor(0xA8, 0xC5, 0xD4))
        tb(s, x + Inches(0.2), Inches(2.55), Inches(3.7), Inches(2.8), body_txt, 12, False, M)
        box(s, x + Inches(0.15), Inches(5.6), Inches(3.8), Inches(1.2), Soft)
        tb(s, x + Inches(0.3), Inches(5.8), Inches(3.5), Inches(0.85), foot, 12, True, N)

    out = OUT_DIR / "GPS精益_筑基势起_向总汇报.pptx"
    prs.save(str(out))
    print("wrote", out)
    return out


def write_html(labels):
    tpl_path = ROOT.parents[1] / "skills/_shared/assets/preview-template.html"
    template = tpl_path.read_text(encoding="utf-8")
    html_out = (
        template.replace("{{TITLE}}", "GPS精益｜筑基→势起（6页）")
        .replace("{{LOGO}}", "GPS")
        .replace("{{ACCENT_COLOR}}", ORANGE)
        .replace("{{SLIDES_JSON}}", json.dumps(labels, ensure_ascii=False))
    )
    (OUT_DIR / "index.html").write_text(html_out, encoding="utf-8")


def write_meta():
    (ROOT / "outline.json").write_text(
        json.dumps(
            {
                "title": "GPS精益｜筑基→势起",
                "approved": True,
                "pages": [{"index": i + 1, "title": label, "file": name} for i, (name, label, _) in enumerate(SLIDES)],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (ROOT / "README.md").write_text(
        """# GPS 精益｜筑基→「势起」（6页 · 对齐原稿）

严格按用户提供的 6 张原稿截图内容重建，不多页、不改叙事。

## 页面对照

| 页 | 内容 |
|----|------|
| 01 | 当前精益推进现存真实痛点 |
| 02 | 筑基→「势起」（453总框架） |
| 03 | 五大体系 01–04 |
| 04 | 价值生产力体系 · 五大核心课题 |
| 05 | 三个阶段 |
| 06 | 真善美底层逻辑 |

## 交付物

- `output/slide-01.svg` … `slide-06.svg`
- `output/index.html`
- `output/GPS精益_筑基势起_向总汇报.pptx`

```bash
python generate_deck.py
```
""",
        encoding="utf-8",
    )


def main():
    SLIDES_DIR.mkdir(parents=True, exist_ok=True)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    # 清理旧的 7–10 页
    for p in list(OUT_DIR.glob("slide-*.svg")) + list(SLIDES_DIR.glob("slide-*.svg")):
        if p.name > "slide-06.svg":
            p.unlink()
            print("removed", p)
    labels = []
    for name, label, fn in SLIDES:
        content = fn()
        (SLIDES_DIR / name).write_text(content, encoding="utf-8")
        (OUT_DIR / name).write_text(content, encoding="utf-8")
        labels.append({"file": name, "label": label})
        print("wrote", name)
    write_html(labels)
    write_meta()
    build_pptx()
    print("DONE — 6 slides")


if __name__ == "__main__":
    main()
