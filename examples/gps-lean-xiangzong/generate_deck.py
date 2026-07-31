#!/usr/bin/env python3
"""向总汇报 6 页 — 模板 B：精密青石灰（内容不变，视觉全换）。"""

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

# ── Template B: Precision Slate-Teal（区别于上一套橙白商务）──
INK = "#0F172A"
INK2 = "#1E293B"
TEAL = "#0F766E"
TEAL_SOFT = "#CCFBF1"
TEAL_MID = "#14B8A6"
BG = "#F1F5F9"
PAPER = "#FFFFFF"
TEXT = "#0F172A"
MUTED = "#64748B"
LINE = "#CBD5E1"
RULE = "#94A3B8"

# 标题用衬线感栈，正文无衬线 — 形成与旧模板不同的气质
HEAD = "Merriweather, Source Han Serif SC, Noto Serif SC, Songti SC, serif"
BODY = "Source Sans 3, PingFang SC, Microsoft YaHei, Noto Sans SC, sans-serif"


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
    <pattern id="bp" width="24" height="24" patternUnits="userSpaceOnUse">
      <path d="M24 0H0V24" fill="none" stroke="{TEAL}" stroke-width="0.4" opacity="0.07"/>
    </pattern>
    <linearGradient id="inkGrad" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="{INK}"/>
      <stop offset="100%" stop-color="{INK2}"/>
    </linearGradient>
  </defs>
  <rect width="1280" height="720" fill="{bg}"/>
  <rect width="1280" height="720" fill="url(#bp)"/>
{body}
</svg>
'''


def T(x, y, content, size=16, fill=TEXT, weight=500, anchor="start", font=None):
    fam = font or BODY
    return (
        f'<text x="{x}" y="{y}" fill="{fill}" font-size="{size}" font-weight="{weight}" '
        f'font-family="{fam}" text-anchor="{anchor}">{esc(content)}</text>'
    )


def TB(x, y, lines, size=13, fill=TEXT, weight=400, lh=1.45, font=None):
    fam = font or BODY
    parts = [
        f'<text x="{x}" y="{y}" fill="{fill}" font-size="{size}" font-weight="{weight}" '
        f'font-family="{fam}">'
    ]
    for i, line in enumerate(lines):
        dy = 0 if i == 0 else int(size * lh)
        parts.append(f'<tspan x="{x}" dy="{dy}">{esc(line)}</tspan>')
    parts.append("</text>")
    return "\n".join(parts)


def R(x, y, w, h, fill=PAPER, stroke=None, r=4):
    s = f' stroke="{stroke}" stroke-width="1"' if stroke else ' stroke="none"'
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" ry="{r}" fill="{fill}"{s}/>'


def rule(x, y, w, color=TEAL):
    return f'<rect x="{x}" y="{y}" width="{w}" height="2" fill="{color}"/>'


def num_badge(x, y, num):
    return f'''
  <circle cx="{x}" cy="{y}" r="18" fill="{TEAL}"/>
  {T(x, y + 5, num, 13, "#fff", 700, "middle")}
'''


# ───────────────── Slide 1: 痛点 ─────────────────

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
    body = f'''
  {R(0, 0, 1280, 72, INK, r=0)}
  {T(48, 46, "当前精益推进现存真实痛点", 26, "#fff", 700, font=HEAD)}
  {T(1180, 46, "01 / 06", 12, TEAL_MID, 600, "end")}
'''
    for i, (num, title, desc) in enumerate(pains):
        col, row = (0, i) if i < 3 else (1, i - 3)
        x = 40 + col * 620
        y = 100 + row * 165
        body += f'''
  {R(x, y, 600, 148, PAPER, LINE, r=4)}
  {rule(x, y, 600, TEAL)}
  {num_badge(x + 36, y + 36, num)}
  {T(x + 68, y + 42, title, 15, INK, 700)}
  {TB(x + 24, y + 72, wrap(desc, 30), 12, MUTED, 400, 1.4)}
'''
    body += f'''
  {R(660, 430, 580, 230, INK, r=4)}
  {T(688, 475, "关键判断", 14, TEAL_MID, 700)}
  {TB(688, 515, wrap("精益生产本质是消除浪费与持续改善；实践中仍面临从表层整改到深层次攻坚的多重执行挑战。", 28), 13, "#E2E8F0", 400, 1.45)}
  {TB(688, 595, wrap("闭环机制与数据沉淀是精益从「运动式」走向「常态化」的关键，破除形式主义、回归价值创造。", 28), 13, "#94A3B8", 400, 1.45)}
'''
    return svg_wrap(body)


# ───────────────── Slide 2: 筑基→势起 ─────────────────

def slide_02():
    body = f'''
  {R(0, 0, 1280, 72, INK, r=0)}
  {T(48, 46, "筑基 →「势起」", 26, "#fff", 700, font=HEAD)}
  {T(1180, 46, "02 / 06", 12, TEAL_MID, 600, "end")}
  {TB(48, 100, [
      "自 2025 年 9 月以来，GPS 精益体系已从「探索与单点整改」迈入「体系起势与机制落地」阶段。",
      "2026 年为关键推进年：精益作为「第二生产力」，以「453」框架启航精益转型，践行「真善美」价值主张。"
  ], 13, MUTED, 400, 1.4)}

  {T(48, 165, "四级架构", 13, TEAL, 700)}
  {rule(48, 175, 48)}
'''
    levels = [
        ("决策层 · 战略引领", "确立精益战略，审批重大课题，提供核心资源与政策支持"),
        ("推进层 · 体系赋能", "体系规划、方法导入、项目督导与专业人才培养"),
        ("执行层 · 攻坚落地", "基地一把手主责，目标分解与资源协同，确保课题出成果"),
        ("落地层 · 现场改善", "班组长带一线严守 SOP，参与现场改善提案"),
    ]
    for i, (t, d) in enumerate(levels):
        x = 48 + i * 305
        body += f'''
  {R(x, 190, 290, 105, PAPER, LINE, r=4)}
  {R(x, 190, 290, 6, TEAL if i == 0 else INK, r=0)}
  {T(x + 14, 225, t, 13, INK, 700)}
  {TB(x + 14, 250, wrap(d, 14), 11, MUTED, 400, 1.35)}
'''
    body += f'''
  {T(48, 325, "五大体系", 13, TEAL, 700)}
  {rule(48, 335, 48)}
'''
    systems = [
        ("01 对标研学", "被动学习 → 主动对标"),
        ("02 标杆复制", "单点优势 → 全域标准化"),
        ("03 人才育成", "基础普及 → 精英赋能"),
        ("04 考核管控", "柔性要求 → 刚性闭环"),
        ("05 价值生产力", "精益锚定核心课题"),
    ]
    for i, (t, d) in enumerate(systems):
        x = 48 + i * 242
        body += f'''
  {R(x, 350, 228, 72, TEAL_SOFT if i == 4 else PAPER, TEAL if i == 4 else LINE, r=4)}
  {T(x + 14, 378, t, 13, INK, 700)}
  {T(x + 14, 402, d, 11, MUTED, 400)}
'''
    body += f'''
  {T(48, 450, "三阶段路径（2026 下半年）", 13, TEAL, 700)}
  {rule(48, 460, 48)}
'''
    phases = [
        ("7–8 月 · 筑基巩固期", "补短板、定标准、全员培训，夯实现场基础管理规范化"),
        ("9–10 月 · 攻坚突破期", "专项改善课题，解决核心经营痛点与瓶颈"),
        ("11–12 月 · 价值兑现期", "固化成果、建立长效机制，验证经济效益与效率提升"),
    ]
    for i, (t, d) in enumerate(phases):
        x = 48 + i * 405
        body += f'''
  {R(x, 475, 390, 95, PAPER, LINE, r=4)}
  {T(x + 16, 508, t, 14, INK, 700)}
  {TB(x + 16, 535, wrap(d, 20), 11, MUTED, 400, 1.35)}
'''
    body += f'''
  {R(48, 595, 1184, 90, INK, r=4)}
  {T(72, 635, "453 框架 = 四级架构 × 五大体系 × 三阶段路径", 16, TEAL_MID, 700)}
  {T(72, 662, "以体系起势驱动机制落地，打造支撑公司高质量发展的「第二生产力」", 12, "#94A3B8", 400)}
'''
    return svg_wrap(body)


# ───────────────── Slide 3: 五大体系 01-04 ─────────────────

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
  {R(0, 0, 1280, 72, INK, r=0)}
  {T(48, 46, "五大体系", 26, "#fff", 700, font=HEAD)}
  {T(200, 46, "管理支撑：对标 · 复制 · 育成 · 考核", 13, "#94A3B8", 400)}
  {T(1180, 46, "03 / 06", 12, TEAL_MID, 600, "end")}
'''
    positions = [(40, 100), (660, 100), (40, 400), (660, 400)]
    for (num, title, sub, bullets), (x, y) in zip(blocks, positions):
        body += f'''
  {R(x, y, 580, 270, PAPER, LINE, r=4)}
  {R(x, y, 72, 270, INK, r=0)}
  {T(x + 36, y + 145, num, 22, TEAL_MID, 800, "middle")}
  {T(x + 96, y + 48, title, 18, INK, 700, font=HEAD)}
  {T(x + 96, y + 78, sub, 12, TEAL, 600)}
  {rule(x + 96, y + 92, 40)}
'''
        yy = y + 120
        for b in bullets:
            body += TB(x + 96, yy, wrap("· " + b, 28), 12, MUTED, 400, 1.4)
            yy += 14 + 12 * 1.4 * len(wrap(b, 28))
    return svg_wrap(body)


# ───────────────── Slide 4: 价值生产力 ─────────────────

def slide_04():
    topics = [
        ("拉晶硅损优化", "厘清硅损价值链浪费点；降低切割与磨削余量；提升硅料来源利用效率。"),
        ("切片损耗管控", "攻克断线、碎片难点；提升硅片切割良率与优品率。"),
        ("组件质量攻坚", "管控层压缺陷、隐裂、虚焊及电池片碎片率；保障产品可靠性与寿命。"),
        ("设备效能跃升", "全面夯实 TPM 运行模式；提升 OEE 指标；保障设备与产线连续稳定运行。"),
        ("精益数字化+", "以标准化为基础，全局拉通顶层数字化模型；整合 DMS、TPMS、QMS 等系统。"),
    ]
    body = f'''
  {R(0, 0, 1280, 72, INK, r=0)}
  {T(48, 42, "五大体系", 18, "#94A3B8", 500)}
  {T(48, 66, "价值生产力体系：以精益锚定核心课题、助力价值增长", 16, TEAL_MID, 600)}
  {T(1180, 46, "04 / 06", 12, TEAL_MID, 600, "end")}

  {R(40, 100, 340, 450, "url(#inkGrad)", r=4)}
  {T(68, 160, "05", 40, TEAL_MID, 800)}
  {T(68, 220, "聚焦五大", 24, "#fff", 700, font=HEAD)}
  {T(68, 258, "核心课题", 24, "#fff", 700, font=HEAD)}
  {T(68, 300, "突破生产瓶颈", 14, "#94A3B8", 400)}
  {TB(68, 360, wrap("推动精益改善转化为产能增量、良率增量、效益增量，打造「第二生产力」引擎。", 14), 12, "#CBD5E1", 400, 1.45)}
'''
    for i, (title, desc) in enumerate(topics):
        y = 100 + i * 90
        body += f'''
  {R(400, y, 840, 80, PAPER, LINE, r=4)}
  {R(400, y, 6, 80, TEAL, r=0)}
  {T(430, y + 32, f"0{i+1}  {title}", 16, INK, 700)}
  {T(430, y + 58, desc, 12, MUTED, 400)}
'''
    body += f'''
  {R(40, 575, 1200, 105, TEAL_SOFT, TEAL, r=4)}
  {T(68, 615, "实现管理与产能双增长", 15, TEAL, 700)}
  {TB(68, 645, ["推动精益改善转化为实实在在的 产能增量、良率增量、效益增量，打造支撑公司业务规模化、高质量发展的「第二生产力」引擎"], 12, INK, 500, 1.35)}
'''
    return svg_wrap(body)


# ───────────────── Slide 5: 三个阶段 ─────────────────

def slide_05():
    body = f'''
  {R(0, 0, 1280, 72, INK, r=0)}
  {T(48, 46, "三个阶段", 26, "#fff", 700, font=HEAD)}
  {T(1180, 46, "05 / 06", 12, TEAL_MID, 600, "end")}

  {R(40, 95, 590, 78, PAPER, LINE, r=4)}
  {T(60, 125, "2026 上半年｜筑基起势（0→1）", 14, INK, 700)}
  {T(60, 152, "建体系、树标杆、统一组织认知", 12, MUTED, 400)}
  {R(650, 95, 590, 78, INK, r=4)}
  {T(670, 125, "2026 下半年｜全域放大与价值兑现（1→N）", 14, "#fff", 700)}
  {T(670, 152, "标杆复制、痛点攻坚、机制固化、产能效益兑现", 12, "#94A3B8", 400)}

  {T(40, 200, "实施路径：三阶段推进，层层递进", 13, TEAL, 700)}
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
        x = 40 + i * 410
        head = TEAL if i == 1 else INK
        body += f'''
  {R(x, 220, 390, 450, PAPER, LINE, r=4)}
  {R(x, 220, 390, 86, head, r=4)}
  <rect x="{x}" y="280" width="390" height="26" fill="{head}"/>
  {T(x + 20, 255, f"{num}  {time}", 12, TEAL_MID if i != 1 else "#fff", 700)}
  {T(x + 20, 285, theme, 15, "#fff", 700)}
'''
        yy = 335
        for b in bullets:
            lines = wrap("· " + b, 18)
            body += TB(x + 18, yy, lines, 12, MUTED, 400, 1.35)
            yy += 10 + int(12 * 1.35 * len(lines))
        body += f'''
  {R(x + 14, 575, 362, 75, TEAL_SOFT, r=4)}
  {TB(x + 26, 600, wrap("目标：" + goal, 18), 11, INK, 600, 1.35)}
'''
    return svg_wrap(body)


# ───────────────── Slide 6: 真善美 ─────────────────

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
  {R(0, 0, 1280, 72, INK, r=0)}
  {T(48, 42, "精益全域落地｜以「真善美」锚定改善底层逻辑", 18, "#fff", 700, font=HEAD)}
  {T(48, 64, "统一对标公司核心价值主张，构建可信、有效、长效的改善闭环", 11, "#94A3B8", 400)}
  {T(1180, 46, "06 / 06", 12, TEAL_MID, 600, "end")}
'''
    for i, (title, defn, lead, bullets, outcome, foot) in enumerate(cols):
        x = 40 + i * 410
        head = INK if i != 1 else TEAL
        soft = TEAL_SOFT
        body += f'''
  {R(x, 100, 390, 560, PAPER, LINE, r=4)}
  {R(x, 100, 390, 100, head, r=4)}
  <rect x="{x}" y="180" width="390" height="20" fill="{head}"/>
  {T(x + 22, 145, title, 24, "#fff", 700, font=HEAD)}
  {T(x + 22, 175, defn, 12, "#94A3B8" if i != 1 else "#CCFBF1", 500)}
  {TB(x + 20, 230, wrap(lead, 20), 12, TEXT, 500, 1.35)}
'''
        yy = 290
        for b in bullets:
            lines = wrap("· " + b, 18)
            body += TB(x + 18, yy, lines, 12, MUTED, 400, 1.35)
            yy += 8 + int(12 * 1.35 * len(lines))
        body += f'''
  {R(x + 14, 545, 362, 95, soft, r=4)}
  {T(x + 28, 580, outcome, 15, TEAL, 700)}
  {TB(x + 28, 608, wrap(foot, 18), 12, INK, 500, 1.35)}
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


# ───────────────── PPTX ─────────────────

def set_font(run, size=12, bold=False, color=RGBColor(0x0F, 0x17, 0x2A)):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Microsoft YaHei"
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
    if ea is None:
        ea = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
    ea.set("typeface", "Microsoft YaHei")


def box(slide, l, t, w, h, fill, line=None, radius=True):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(shape, l, t, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    if radius:
        try:
            sh.adjustments[0] = 0.04
        except Exception:
            pass
    return sh


def tb(slide, l, t, w, h, text, size=14, bold=False, color=RGBColor(0x0F, 0x17, 0x2A), align=PP_ALIGN.LEFT):
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
    N = RGBColor(0x0F, 0x17, 0x2A)
    Teal = RGBColor(0x0F, 0x76, 0x6E)
    TealSoft = RGBColor(0xCC, 0xFB, 0xF1)
    W = RGBColor(0xFF, 0xFF, 0xFF)
    M = RGBColor(0x64, 0x74, 0x8B)
    Tcol = RGBColor(0x0F, 0x17, 0x2A)
    L = RGBColor(0xCB, 0xD5, 0xE1)
    BG_c = RGBColor(0xF1, 0xF5, 0xF9)

    # P1
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c, radius=False)
    box(s, 0, 0, prs.slide_width, Inches(0.75), N, radius=False)
    tb(s, Inches(0.45), Inches(0.2), Inches(11), Inches(0.4), "当前精益推进现存真实痛点", 22, True, W)
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
        y = Inches(1.0 + row * 1.85)
        box(s, x, y, Inches(6.2), Inches(1.7), W, L)
        tb(s, x + Inches(0.2), y + Inches(0.15), Inches(5.8), Inches(0.35), title, 13, True, Teal)
        tb(s, x + Inches(0.2), y + Inches(0.55), Inches(5.8), Inches(1.0), desc, 11, False, M)
    box(s, Inches(6.85), Inches(4.7), Inches(6.2), Inches(2.35), N)
    tb(s, Inches(7.1), Inches(4.95), Inches(5.7), Inches(1.9),
       "关键判断\n精益生产本质是消除浪费与持续改善；实践中仍面临从表层整改到深层次攻坚的多重执行挑战。\n闭环机制与数据沉淀是精益从「运动式」走向「常态化」的关键，破除形式主义、回归价值创造。",
       12, False, W)

    # P2
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c, radius=False)
    box(s, 0, 0, prs.slide_width, Inches(0.75), N, radius=False)
    tb(s, Inches(0.45), Inches(0.2), Inches(12), Inches(0.4), "筑基 →「势起」", 22, True, W)
    tb(s, Inches(0.45), Inches(0.9), Inches(12.4), Inches(0.65),
       "自2025年9月以来，GPS精益体系已从「探索与单点整改」迈入「体系起势与机制落地」阶段。"
       "2026年为关键推进年：精益作为「第二生产力」，以「453」框架启航精益转型，践行「真善美」价值主张。",
       11, False, M)
    tb(s, Inches(0.45), Inches(1.6), Inches(4), Inches(0.25), "四级架构", 12, True, Teal)
    for i, (a, b) in enumerate([
        ("决策层·战略引领", "确立战略/审批课题/资源政策支持"),
        ("推进层·体系赋能", "体系规划/方法导入/督导与育人"),
        ("执行层·攻坚落地", "基地主责/目标分解/资源协同"),
        ("落地层·现场改善", "班组SOP/现场提案改善"),
    ]):
        x = Inches(0.45 + i * 3.2)
        box(s, x, Inches(1.95), Inches(3.05), Inches(1.1), W, L)
        tb(s, x + Inches(0.12), Inches(2.1), Inches(2.8), Inches(0.3), a, 12, True, N)
        tb(s, x + Inches(0.12), Inches(2.5), Inches(2.8), Inches(0.4), b, 10, False, M)
    tb(s, Inches(0.45), Inches(3.25), Inches(4), Inches(0.25), "五大体系", 12, True, Teal)
    for i, name in enumerate(["01对标研学", "02标杆复制", "03人才育成", "04考核管控", "05价值生产力"]):
        x = Inches(0.45 + i * 2.55)
        box(s, x, Inches(3.55), Inches(2.4), Inches(0.65), TealSoft if i == 4 else W, Teal if i == 4 else L)
        tb(s, x, Inches(3.7), Inches(2.4), Inches(0.35), name, 12, True, N, PP_ALIGN.CENTER)
    tb(s, Inches(0.45), Inches(4.45), Inches(8), Inches(0.25), "三阶段路径（2026下半年）", 12, True, Teal)
    for i, (t, d) in enumerate([
        ("7–8月 筑基巩固期", "补短板、定标准、全员培训"),
        ("9–10月 攻坚突破期", "专项课题，解决核心痛点瓶颈"),
        ("11–12月 价值兑现期", "固化成果、长效机制、效益验证"),
    ]):
        x = Inches(0.45 + i * 4.25)
        box(s, x, Inches(4.8), Inches(4.05), Inches(1.1), W, L)
        tb(s, x + Inches(0.15), Inches(4.95), Inches(3.7), Inches(0.3), t, 12, True, N)
        tb(s, x + Inches(0.15), Inches(5.35), Inches(3.7), Inches(0.35), d, 11, False, M)
    box(s, Inches(0.45), Inches(6.2), Inches(12.4), Inches(0.9), N)
    tb(s, Inches(0.7), Inches(6.45), Inches(12), Inches(0.45),
       "453框架 = 四级架构 × 五大体系 × 三阶段路径　｜　打造「第二生产力」", 13, True, W)

    # P3
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c, radius=False)
    box(s, 0, 0, prs.slide_width, Inches(0.75), N, radius=False)
    tb(s, Inches(0.45), Inches(0.2), Inches(12), Inches(0.4), "五大体系", 22, True, W)
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
        box(s, x, y, Inches(0.7), Inches(2.85), N, radius=False)
        tb(s, x + Inches(0.1), y + Inches(1.15), Inches(0.5), Inches(0.4), f"0{i+1}", 14, True, Teal, PP_ALIGN.CENTER)
        tb(s, x + Inches(0.9), y + Inches(0.25), Inches(5.1), Inches(0.35), title, 14, True, N)
        tb(s, x + Inches(0.9), y + Inches(0.65), Inches(5.1), Inches(0.3), sub, 11, False, Teal)
        tb(s, x + Inches(0.9), y + Inches(1.15), Inches(5.1), Inches(1.4), body_txt, 11, False, M)

    # P4
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c, radius=False)
    box(s, 0, 0, prs.slide_width, Inches(0.75), N, radius=False)
    tb(s, Inches(0.45), Inches(0.15), Inches(12), Inches(0.25), "五大体系", 14, False, RGBColor(0x94, 0xA3, 0xB8))
    tb(s, Inches(0.45), Inches(0.4), Inches(12), Inches(0.3), "价值生产力体系：以精益锚定核心课题、助力价值增长", 14, True, Teal)
    box(s, Inches(0.4), Inches(1.05), Inches(4.0), Inches(4.6), N)
    tb(s, Inches(0.7), Inches(1.5), Inches(3.4), Inches(0.4), "05", 28, True, Teal)
    tb(s, Inches(0.7), Inches(2.2), Inches(3.4), Inches(1.0), "聚焦五大\n核心课题", 22, True, W)
    tb(s, Inches(0.7), Inches(3.5), Inches(3.4), Inches(0.35), "突破生产瓶颈", 12, False, RGBColor(0x94, 0xA3, 0xB8))
    tb(s, Inches(0.7), Inches(4.1), Inches(3.4), Inches(1.1),
       "推动精益改善转化为产能增量、良率增量、效益增量，打造「第二生产力」引擎。", 11, False, RGBColor(0xCB, 0xD5, 0xE1))
    topics = [
        "01 拉晶硅损优化 — 厘清硅损价值链浪费点；降低切割与磨削余量；提升硅料利用效率。",
        "02 切片损耗管控 — 攻克断线、碎片难点；提升硅片切割良率与优品率。",
        "03 组件质量攻坚 — 管控层压缺陷、隐裂、虚焊及碎片率；保障可靠性与寿命。",
        "04 设备效能跃升 — 夯实TPM运行模式；提升OEE；保障设备与产线连续稳定运行。",
        "05 精益数字化+ — 以标准化为基础，拉通顶层数字化模型；整合DMS、TPMS、QMS。",
    ]
    for i, line in enumerate(topics):
        y = Inches(1.05 + i * 0.92)
        box(s, Inches(4.65), y, Inches(8.25), Inches(0.82), W, L)
        tb(s, Inches(4.9), y + Inches(0.2), Inches(7.8), Inches(0.5), line, 11, True, N)
    box(s, Inches(0.4), Inches(5.9), Inches(12.5), Inches(1.2), TealSoft, Teal)
    tb(s, Inches(0.7), Inches(6.1), Inches(12), Inches(0.3), "实现管理与产能双增长", 13, True, Teal)
    tb(s, Inches(0.7), Inches(6.45), Inches(12), Inches(0.45),
       "推动精益改善转化为实实在在的产能增量、良率增量、效益增量，打造支撑公司业务规模化、高质量发展的「第二生产力」引擎。",
       11, False, Tcol)

    # P5
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c, radius=False)
    box(s, 0, 0, prs.slide_width, Inches(0.75), N, radius=False)
    tb(s, Inches(0.45), Inches(0.2), Inches(12), Inches(0.4), "三个阶段", 22, True, W)
    box(s, Inches(0.4), Inches(0.95), Inches(6.1), Inches(0.8), W, L)
    tb(s, Inches(0.6), Inches(1.05), Inches(5.7), Inches(0.3), "2026上半年｜筑基起势（0→1）", 12, True, N)
    tb(s, Inches(0.6), Inches(1.4), Inches(5.7), Inches(0.25), "建体系、树标杆、统一组织认知", 10, False, M)
    box(s, Inches(6.75), Inches(0.95), Inches(6.1), Inches(0.8), N)
    tb(s, Inches(6.95), Inches(1.05), Inches(5.7), Inches(0.3), "2026下半年｜全域放大与价值兑现（1→N）", 12, True, W)
    tb(s, Inches(6.95), Inches(1.4), Inches(5.7), Inches(0.25), "标杆复制、痛点攻坚、机制固化、产能效益兑现", 10, False, RGBColor(0x94, 0xA3, 0xB8))
    tb(s, Inches(0.45), Inches(1.95), Inches(10), Inches(0.25), "实施路径：三阶段推进，层层递进", 12, True, Teal)
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
        box(s, x, Inches(2.35), Inches(4.1), Inches(4.7), W, L)
        box(s, x, Inches(2.35), Inches(4.1), Inches(1.1), Teal if i == 1 else N)
        tb(s, x + Inches(0.2), Inches(2.5), Inches(3.7), Inches(0.3), t, 11, True, W)
        tb(s, x + Inches(0.2), Inches(2.9), Inches(3.7), Inches(0.35), theme, 13, True, W)
        tb(s, x + Inches(0.2), Inches(3.7), Inches(3.7), Inches(2.2), acts, 11, False, M)
        box(s, x + Inches(0.15), Inches(5.9), Inches(3.8), Inches(0.95), TealSoft)
        tb(s, x + Inches(0.3), Inches(6.05), Inches(3.5), Inches(0.7), goal, 10, True, N)

    # P6
    s = prs.slides.add_slide(blank)
    box(s, 0, 0, prs.slide_width, prs.slide_height, BG_c, radius=False)
    box(s, 0, 0, prs.slide_width, Inches(0.75), N, radius=False)
    tb(s, Inches(0.4), Inches(0.15), Inches(12.5), Inches(0.3),
       "精益全域落地｜以「真善美」锚定改善底层逻辑", 16, True, W)
    tb(s, Inches(0.4), Inches(0.45), Inches(12.5), Inches(0.25),
       "所有精益课题、现场改善、体系运行统一对标公司核心价值主张，构建可信、有效、长效的改善闭环", 10, False, RGBColor(0x94, 0xA3, 0xB8))
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
        box(s, x, Inches(1.0), Inches(4.1), Inches(6.05), W, L)
        box(s, x, Inches(1.0), Inches(4.1), Inches(1.15), c)
        tb(s, x + Inches(0.2), Inches(1.2), Inches(3.7), Inches(0.4), title, 18, True, W)
        tb(s, x + Inches(0.2), Inches(1.7), Inches(3.7), Inches(0.3), defn, 11, False, RGBColor(0x94, 0xA3, 0xB8))
        tb(s, x + Inches(0.2), Inches(2.5), Inches(3.7), Inches(2.8), body_txt, 11, False, M)
        box(s, x + Inches(0.15), Inches(5.55), Inches(3.8), Inches(1.25), TealSoft)
        tb(s, x + Inches(0.3), Inches(5.75), Inches(3.5), Inches(0.9), foot, 11, True, N)

    out = OUT_DIR / "GPS精益_筑基势起_向总汇报.pptx"
    prs.save(str(out))
    print("wrote", out)
    return out


def write_html(labels):
    tpl_path = ROOT.parents[1] / "skills/_shared/assets/preview-template.html"
    template = tpl_path.read_text(encoding="utf-8")
    html_out = (
        template.replace("{{TITLE}}", "GPS精益｜筑基→势起（模板B）")
        .replace("{{LOGO}}", "GPS")
        .replace("{{ACCENT_COLOR}}", TEAL)
        .replace("{{SLIDES_JSON}}", json.dumps(labels, ensure_ascii=False))
    )
    (OUT_DIR / "index.html").write_text(html_out, encoding="utf-8")


def write_meta():
    (ROOT / "outline.json").write_text(
        json.dumps(
            {
                "title": "GPS精益｜筑基→势起",
                "template": "precision-slate-teal",
                "approved": True,
                "pages": [{"index": i + 1, "title": label, "file": name} for i, (name, label, _) in enumerate(SLIDES)],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (ROOT / "README.md").write_text(
        """# GPS 精益｜筑基→「势起」（6页）

内容对齐原稿 6 张截图；当前视觉为 **模板 B：精密青石灰**（slate + teal）。

## 模板说明
- 上一套：橙白商务圆角卡片
- 当前套：石板黑顶栏 + 青绿强调 + 蓝图底纹 + 小圆角学术风

## 页面
01 痛点 · 02 筑基势起 · 03 五大体系01-04 · 04 价值生产力 · 05 三阶段 · 06 真善美

```bash
python generate_deck.py
```
""",
        encoding="utf-8",
    )


def main():
    SLIDES_DIR.mkdir(parents=True, exist_ok=True)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
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
    print("DONE — template B (precision slate-teal), 6 slides")


if __name__ == "__main__":
    main()
