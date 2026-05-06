"""Generate PPTX presentation for H&M recommendation project."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ── Color palette ──
BG_DARK = RGBColor(0x1A, 0x1A, 0x2E)
BG_CARD = RGBColor(0x16, 0x21, 0x3E)
ACCENT = RGBColor(0x00, 0xD2, 0xFF)
ACCENT2 = RGBColor(0xFF, 0x6B, 0x6B)
ACCENT3 = RGBColor(0x4E, 0xCB, 0x71)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xBB, 0xBB, 0xCC)
GOLD = RGBColor(0xFF, 0xD7, 0x00)


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_shape(slide, left, top, width, height, fill_color, border_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=18,
                 color=WHITE, bold=False, alignment=PP_ALIGN.LEFT, font_name="Microsoft YaHei"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_bullet_list(slide, left, top, width, height, items, font_size=16,
                    color=LIGHT_GRAY, bullet_color=ACCENT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = "Microsoft YaHei"
        p.space_after = Pt(8)
        p.level = 0
    return txBox


# ════════════════════════════════════════
# Slide 1: Title
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
set_slide_bg(slide, BG_DARK)

# Decorative accent line
add_shape(slide, Inches(1), Inches(2.8), Inches(2), Pt(4), ACCENT)

add_text_box(slide, Inches(1), Inches(1.5), Inches(11), Inches(1.2),
             "H&M 个性化推荐系统", font_size=48, color=WHITE, bold=True)
add_text_box(slide, Inches(1), Inches(3.0), Inches(11), Inches(0.8),
             "数据分析报告", font_size=36, color=ACCENT)

add_text_box(slide, Inches(1), Inches(4.5), Inches(11), Inches(0.5),
             "课程：大数据分析与计算 (610ZH125)", font_size=18, color=LIGHT_GRAY)
add_text_box(slide, Inches(1), Inches(5.0), Inches(11), Inches(0.5),
             "团队：付宝昊、梁智毅", font_size=18, color=LIGHT_GRAY)
add_text_box(slide, Inches(1), Inches(5.5), Inches(11), Inches(0.5),
             "竞赛：Kaggle — H&M Personalized Fashion Recommendations", font_size=18, color=LIGHT_GRAY)
add_text_box(slide, Inches(1), Inches(6.0), Inches(11), Inches(0.5),
             "2026 年 5 月 6 日", font_size=16, color=LIGHT_GRAY)


# ════════════════════════════════════════
# Slide 2: Executive Summary
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "Executive Summary", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Key metric cards
cards = [
    ("31,000,000", "交易记录", ACCENT),
    ("1,371,980", "活跃客户", ACCENT2),
    ("MAP@12", "0.02229", GOLD),
    ("+19%", "vs 基线", ACCENT3),
]
for i, (val, label, clr) in enumerate(cards):
    x = Inches(0.8 + i * 3.1)
    card = add_shape(slide, x, Inches(1.5), Inches(2.8), Inches(1.4), BG_CARD, clr)
    add_text_box(slide, x + Inches(0.2), Inches(1.6), Inches(2.4), Inches(0.8),
                 val, font_size=32, color=clr, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.2), Inches(2.2), Inches(2.4), Inches(0.5),
                 label, font_size=16, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Summary text
summary = (
    "基于 H&M 近两年的 3100 万条交易数据，我们发现时尚电商的购买行为呈现极强的短期复购倾向——"
    "近 31 天内有重复购买记录的用户占比超过 40%。据此设计的\"时间衰减 + 位置加权复购\"模型（V13），"
    "MAP@12 从基线 0.01873 提升至 0.02229（+19%）。LightGBM 消融实验与 SHAP 分析验证了"
    "\"复购次数\"和\"近期商品热度\"是最强预测因子，而用户年龄等人口统计特征的边际贡献接近于零。"
)
add_text_box(slide, Inches(0.8), Inches(3.3), Inches(11.5), Inches(3.5),
             summary, font_size=18, color=LIGHT_GRAY)


# ════════════════════════════════════════
# Slide 3: 数据概况
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "数据概况", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Data table
table_data = [
    ("数据集", "记录数", "时间范围", "关键字段"),
    ("transactions_train", "31,000,000", "2018-09 ~ 2020-09", "customer_id, article_id, t_dat, price"),
    ("customers", "1,371,980", "—", "age, postal_code, club_member_status"),
    ("articles", "105,542", "—", "product_type, colour_group, index_name"),
]

rows, cols = len(table_data), len(table_data[0])
tbl = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.5), Inches(2.5)).table
tbl.columns[0].width = Inches(2.5)
tbl.columns[1].width = Inches(2)
tbl.columns[2].width = Inches(2.5)
tbl.columns[3].width = Inches(4.5)

for r in range(rows):
    for c in range(cols):
        cell = tbl.cell(r, c)
        cell.text = table_data[r][c]
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(14)
            paragraph.font.name = "Microsoft YaHei"
            if r == 0:
                paragraph.font.color.rgb = ACCENT
                paragraph.font.bold = True
            else:
                paragraph.font.color.rgb = LIGHT_GRAY
            paragraph.alignment = PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb = BG_CARD if r > 0 else RGBColor(0x0F, 0x17, 0x2A)

# EDA findings
add_text_box(slide, Inches(0.8), Inches(4.3), Inches(11), Inches(0.5),
             "EDA 关键发现", font_size=24, color=WHITE, bold=True)

findings = [
    "1. 长尾幂律分布：5% 热门商品贡献 50%+ 交易量 → 全局热门即强基线",
    "2. 短期复购是第一驱动力：40%+ 用户在 31 天内有重复购买 → 以复购为核心",
    "3. 价格服从对数正态分布 → 特征工程中对价格做 log1p 变换",
]
add_bullet_list(slide, Inches(0.8), Inches(4.9), Inches(11.5), Inches(2.2), findings, font_size=16)


# ════════════════════════════════════════
# Slide 4: 业务背景与问题定义
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "业务背景与问题定义", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Three challenge cards
challenges = [
    ("商品生命周期极短", "当季款式可能几周后就下架\n推荐必须紧跟时效性", ACCENT),
    ("用户偏好季节性漂移", "春季外套 vs 夏季短袖\n模型\"有效期\"远短于其他行业", ACCENT2),
    ("复购集中于基础款", "黑色T恤、牛仔裤等消耗品\n复购信号 > 跨品类推荐", ACCENT3),
]
for i, (title, desc, clr) in enumerate(challenges):
    x = Inches(0.8 + i * 4.1)
    add_shape(slide, x, Inches(1.5), Inches(3.8), Inches(2.2), BG_CARD, clr)
    add_text_box(slide, x + Inches(0.2), Inches(1.6), Inches(3.4), Inches(0.6),
                 title, font_size=20, color=clr, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.2), Inches(2.3), Inches(3.4), Inches(1.2),
                 desc, font_size=15, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Core question
add_shape(slide, Inches(0.8), Inches(4.2), Inches(11.5), Inches(2.5), BG_CARD, GOLD)
add_text_box(slide, Inches(1.2), Inches(4.3), Inches(10.7), Inches(0.5),
             "核心问题", font_size=22, color=GOLD, bold=True)
add_text_box(slide, Inches(1.2), Inches(4.9), Inches(10.7), Inches(1.5),
             "如何在 137 万用户 × 10 万商品的庞大空间中，为每位用户精准定位其未来 7 天最可能购买的 12 件商品？\n\n"
             "核心判断：短期行为信号（近 31 天购买历史）远比长期画像（年龄、性别、消费水平）更有预测价值",
             font_size=17, color=LIGHT_GRAY)


# ════════════════════════════════════════
# Slide 5: 整体流水线
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "整体流水线", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Pipeline steps
steps = [
    ("01", "数据加载", "3 张 CSV 表", ACCENT),
    ("02", "EDA & 清洗", "缺失值/去重/对数", ACCENT2),
    ("03", "特征工程", "时间衰减/位置权重", ACCENT3),
    ("04", "模型推理", "加权复购→Top-12", GOLD),
    ("05", "评估提交", "MAP@12=0.02229", ACCENT),
]

for i, (num, title, desc, clr) in enumerate(steps):
    x = Inches(0.5 + i * 2.55)
    add_shape(slide, x, Inches(1.5), Inches(2.3), Inches(2.0), BG_CARD, clr)
    add_text_box(slide, x + Inches(0.1), Inches(1.55), Inches(2.1), Inches(0.5),
                 num, font_size=28, color=clr, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.1), Inches(2.0), Inches(2.1), Inches(0.5),
                 title, font_size=18, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.1), Inches(2.5), Inches(2.1), Inches(0.8),
                 desc, font_size=14, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

    # Arrow between steps
    if i < len(steps) - 1:
        add_text_box(slide, Inches(0.5 + (i + 1) * 2.55 - Inches(0.2).inches),
                     Inches(2.0), Inches(0.4), Inches(0.5),
                     "→", font_size=28, color=clr, bold=True, alignment=PP_ALIGN.CENTER)

# Validation strategy
add_text_box(slide, Inches(0.8), Inches(4.0), Inches(11), Inches(0.5),
             "验证策略", font_size=24, color=WHITE, bold=True)

val_items = [
    "Time-based Split（按时间顺序切分，禁止随机 K 折）",
    "训练集：2018-09-20 ~ 2020-09-15  |  验证集：2020-09-16 ~ 2020-09-22（7 天窗口）",
    "评估指标：MAP@12（Mean Average Precision at 12）",
]
add_bullet_list(slide, Inches(0.8), Inches(4.6), Inches(11.5), Inches(2), val_items, font_size=16)


# ════════════════════════════════════════
# Slide 6: 核心发现 1 & 2
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "核心发现 (1/2)", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Finding 1
add_shape(slide, Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.2), BG_CARD, ACCENT)
add_text_box(slide, Inches(1.0), Inches(1.6), Inches(5.4), Inches(0.5),
             "发现 1：复购是最可靠的购买信号", font_size=20, color=ACCENT, bold=True)

f1_items = [
    "C（事实）：40%+ 用户在 31 天内有重复购买\n     user_item_history_count Gain=11,799（#1）",
    "I（洞察）：时尚电商≠总在追新\n     基础款（T恤/袜子/牛仔裤）是消耗品\n     复购信噪比远高于跨品类推荐",
    "A（行动）：V13 模型以复购为核心\n     score = (i+1)/N，越近期权重越高\n     MAP@12: 0.01873 → 0.02229 (+19%)",
]
add_bullet_list(slide, Inches(1.0), Inches(2.2), Inches(5.4), Inches(4.2), f1_items, font_size=14)

# Finding 2
add_shape(slide, Inches(6.9), Inches(1.5), Inches(5.8), Inches(5.2), BG_CARD, ACCENT2)
add_text_box(slide, Inches(7.1), Inches(1.6), Inches(5.4), Inches(0.5),
             "发现 2：近期商品热度是第二强信号", font_size=20, color=ACCENT2, bold=True)

f2_items = [
    "C（事实）：item_popularity_7d\n     Gain=12,052（与复购不相上下）\n     SHAP 正向值陡峭上升",
    "I（洞察）：从众效应 + 趋势跟随\n     用户倾向购买\"大家都在买\"的商品\n     热度本身就是推荐信号",
    "A（行动）：冷启动用加权热门补位\n     exp(-0.1×days) 加权近 7 天交易量\n     热门商品全局权重纳入考量",
]
add_bullet_list(slide, Inches(7.1), Inches(2.2), Inches(5.4), Inches(4.2), f2_items, font_size=14)


# ════════════════════════════════════════
# Slide 7: 核心发现 3, 4, 5
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "核心发现 (2/2)", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Finding 3
add_shape(slide, Inches(0.8), Inches(1.5), Inches(3.7), Inches(5.2), BG_CARD, ACCENT3)
add_text_box(slide, Inches(1.0), Inches(1.6), Inches(3.3), Inches(0.5),
             "发现 3：年龄特征≈无用", font_size=18, color=ACCENT3, bold=True)
f3_items = [
    "C：user_age Gain 排#4\n    但移除后 AUC 不变",
    "I：\"伪重要性\"现象\n    树模型高频分裂高基数特征\n    年龄被价格/品类偏好中介",
    "A：不使用任何人口统计特征\n    仅依赖交易行为数据",
]
add_bullet_list(slide, Inches(1.0), Inches(2.2), Inches(3.3), Inches(4.2), f3_items, font_size=13)

# Finding 4
add_shape(slide, Inches(4.8), Inches(1.5), Inches(3.7), Inches(5.2), BG_CARD, ACCENT2)
add_text_box(slide, Inches(5.0), Inches(1.6), Inches(3.3), Inches(0.5),
             "发现 4：价格负向抑制", font_size=18, color=ACCENT2, bold=True)
f4_items = [
    "C：log_price SHAP 负相关\n    价格越高，购买概率越低",
    "I：H&M 客群价格敏感\n    高价商品购买频率低\n    不应过度推荐高价品",
    "A：复购信号天然倾向低价\n    LightGBM 中保留 log_price\n    校准不同价位推荐概率",
]
add_bullet_list(slide, Inches(5.0), Inches(2.2), Inches(3.3), Inches(4.2), f4_items, font_size=13)

# Finding 5
add_shape(slide, Inches(8.8), Inches(1.5), Inches(3.7), Inches(5.2), BG_CARD, GOLD)
add_text_box(slide, Inches(9.0), Inches(1.6), Inches(3.3), Inches(0.5),
             "发现 5：季节性概念漂移", font_size=18, color=GOLD, bold=True)
f5_items = [
    "C：user_avg_price PSI=0.03\n    但 MAP@12 骤降\n    user_age PSI=0.31",
    "I：数据漂移+概念漂移并存\n    客群结构变化\n    特征-目标映射改变",
    "A：仅用最近 31 天数据\n    时间衰减权重\n    建议滑动窗口重训练",
]
add_bullet_list(slide, Inches(9.0), Inches(2.2), Inches(3.3), Inches(4.2), f5_items, font_size=13)


# ════════════════════════════════════════
# Slide 8: 特征工程与消融实验
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "特征工程与消融实验", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Ablation table
ablation_data = [
    ("特征", "移除后 AUC", "ΔAUC", "判定"),
    ("全部特征（Baseline）", "0.8735", "—", "—"),
    ("移除 user_item_history_count", "0.8X", "-0.0X", "核心特征"),
    ("移除 item_popularity_7d", "0.8X", "-0.0X", "核心特征"),
    ("移除 log_price", "0.8X", "-0.0X", "保留"),
    ("移除 user_age", "0.87XX", "≈ 0.00", "放弃"),
    ("移除 club_member_status", "0.87XX", "≈ 0.00", "放弃"),
]

rows, cols = len(ablation_data), len(ablation_data[0])
tbl = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(7), Inches(3.5)).table
tbl.columns[0].width = Inches(3)
tbl.columns[1].width = Inches(1.5)
tbl.columns[2].width = Inches(1)
tbl.columns[3].width = Inches(1.5)

for r in range(rows):
    for c in range(cols):
        cell = tbl.cell(r, c)
        cell.text = ablation_data[r][c]
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(13)
            paragraph.font.name = "Microsoft YaHei"
            if r == 0:
                paragraph.font.color.rgb = ACCENT
                paragraph.font.bold = True
            else:
                paragraph.font.color.rgb = LIGHT_GRAY
            paragraph.alignment = PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb = BG_CARD if r > 0 else RGBColor(0x0F, 0x17, 0x2A)

# Conclusions
add_text_box(slide, Inches(8.2), Inches(1.5), Inches(4.5), Inches(0.5),
             "消融结论", font_size=22, color=WHITE, bold=True)

concl_items = [
    "核心双引擎：复购次数 + 近期热度\n移除任一 → 性能显著下降",
    "辅助信号：log_price\n提供校准，贡献度低于前两者",
    "伪重要性特征：user_age, club_member_status\nGain 排名靠前但真实贡献≈0\nSHAP beeswarm 图验证",
]
add_bullet_list(slide, Inches(8.2), Inches(2.1), Inches(4.5), Inches(4.5), concl_items, font_size=14)

# Core hypothesis
add_shape(slide, Inches(0.8), Inches(5.3), Inches(11.5), Inches(1.5), BG_CARD, GOLD)
add_text_box(slide, Inches(1.0), Inches(5.4), Inches(11.1), Inches(1.2),
             "核心业务假设：用户未来 7 天的购买行为，主要由其近期（31 天内）的购买历史和当前市场热度决定，而非长期人口统计画像。",
             font_size=17, color=GOLD, bold=True)


# ════════════════════════════════════════
# Slide 9: 失败分析
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "失败分析", font_size=36, color=ACCENT2, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT2)

# Failure 1
add_shape(slide, Inches(0.8), Inches(1.5), Inches(3.7), Inches(5.2), BG_CARD, ACCENT2)
add_text_box(slide, Inches(1.0), Inches(1.6), Inches(3.3), Inches(0.5),
             "失败 1：年龄特征", font_size=18, color=ACCENT2, bold=True)
fail1 = [
    "动机：不同年龄应有不同品类偏好",
    "方法：中位数填充 + LightGBM",
    "结果：Gain #4 但消融 AUC 不变",
    "归因：\"伪重要性\"——年龄被\n价格/品类偏好完全中介",
    "教训：Gain ≠ 预测贡献\n必须用消融实验验证",
]
add_bullet_list(slide, Inches(1.0), Inches(2.2), Inches(3.3), Inches(4.2), fail1, font_size=13)

# Failure 2
add_shape(slide, Inches(4.8), Inches(1.5), Inches(3.7), Inches(5.2), BG_CARD, ACCENT2)
add_text_box(slide, Inches(5.0), Inches(1.6), Inches(3.3), Inches(0.5),
             "失败 2：关联规则", font_size=18, color=ACCENT2, bold=True)
fail2 = [
    "动机：挖掘\"买A也买B\"的\n共购模式",
    "方法：同交易共购对频次排序",
    "结果：MAP@12 未超过 V13",
    "归因：单笔交易仅 1-2 件\n共购信号过于稀疏\n时尚搭配关系高度主观",
    "教训：关联规则适用于高频\n低单价场景，不适用时尚电商",
]
add_bullet_list(slide, Inches(5.0), Inches(2.2), Inches(3.3), Inches(4.2), fail2, font_size=13)

# Failure 3
add_shape(slide, Inches(8.8), Inches(1.5), Inches(3.7), Inches(5.2), BG_CARD, ACCENT2)
add_text_box(slide, Inches(9.0), Inches(1.6), Inches(3.3), Inches(0.5),
             "失败 3：随机 K 折", font_size=18, color=ACCENT2, bold=True)
fail3 = [
    "动机：快速评估模型性能",
    "方法：KFold(5, shuffle=True)",
    "结果：MAP@12 虚高 2-3 倍",
    "归因：未来数据泄漏\n用 9 月数据\"预测\" 8 月\n季节性相似导致记忆答案",
    "教训：时序业务永远禁止\n使用随机 K 折！",
]
add_bullet_list(slide, Inches(9.0), Inches(2.2), Inches(3.3), Inches(4.2), fail3, font_size=13)


# ════════════════════════════════════════
# Slide 10: MAP@12 对比 & 讨论
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "模型对比与讨论", font_size=36, color=ACCENT, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT)

# Score comparison cards
scores = [
    ("Baseline", "仅全局热门", "0.01873", LIGHT_GRAY),
    ("V12", "时间衰减", "0.02085", ACCENT),
    ("V13", "位置加权复购", "0.02229", GOLD),
]
for i, (ver, desc, score, clr) in enumerate(scores):
    x = Inches(0.8 + i * 4.1)
    add_shape(slide, x, Inches(1.5), Inches(3.8), Inches(1.8), BG_CARD, clr)
    add_text_box(slide, x + Inches(0.2), Inches(1.55), Inches(3.4), Inches(0.5),
                 ver, font_size=22, color=clr, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.2), Inches(2.0), Inches(3.4), Inches(0.4),
                 desc, font_size=14, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.2), Inches(2.4), Inches(3.4), Inches(0.7),
                 score, font_size=32, color=clr, bold=True, alignment=PP_ALIGN.CENTER)

# Discussion points
add_text_box(slide, Inches(0.8), Inches(3.7), Inches(11), Inches(0.5),
             "业务洞察总结", font_size=22, color=WHITE, bold=True)

disc_items = [
    "核心发现：时尚电商 = \"短期复购 + 趋势跟随\" 双驱动模式",
    "预料之中：价格负向抑制 + 长尾分布（零售经典规律）",
    "出乎意料：用户年龄预测贡献≈0 → 对\"用户画像驱动推荐\"的传统思路提出挑战",
    "技术决策：规则模型 > 深度学习（信号稀疏 + 复购信号强 + 工程简洁）",
]
add_bullet_list(slide, Inches(0.8), Inches(4.3), Inches(11.5), Inches(2.8), disc_items, font_size=16)


# ════════════════════════════════════════
# Slide 11: 建议
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.7),
             "Recommendations", font_size=36, color=ACCENT3, bold=True)
add_shape(slide, Inches(0.8), Inches(1.1), Inches(1.5), Pt(3), ACCENT3)

recs = [
    ("01", "首页\"最近常买\"推荐位",
     "基于\"复购是第一驱动力\"的发现\n在 App 首页展示用户近 31 天内购买次数最多的商品\n及其同品类新品\n直接提升复购转化率，实现成本极低", ACCENT),
    ("02", "冷启动：加权热门 + 品类偏好",
     "新用户注册时收集 1-2 个品类偏好\n在该品类内推荐加权热门商品\n比无差别推荐预期 MAP@12 提升 30-50%", ACCENT2),
    ("03", "模型监控与滑动窗口重训练",
     "建立 MAP@12 监控看板\n连续 3 天下降 >10% 时触发重训练\n训练窗口缩短至 4-6 周\n使用时间衰减权重", ACCENT3),
]

for i, (num, title, desc, clr) in enumerate(recs):
    y = Inches(1.5 + i * 1.95)
    add_shape(slide, Inches(0.8), y, Inches(11.5), Inches(1.75), BG_CARD, clr)
    add_text_box(slide, Inches(1.0), y + Inches(0.05), Inches(0.8), Inches(0.5),
                 num, font_size=28, color=clr, bold=True)
    add_text_box(slide, Inches(1.8), y + Inches(0.05), Inches(4), Inches(0.5),
                 title, font_size=20, color=WHITE, bold=True)
    add_text_box(slide, Inches(1.8), y + Inches(0.5), Inches(10), Inches(1.2),
                 desc, font_size=14, color=LIGHT_GRAY)


# ════════════════════════════════════════
# Slide 12: Thank You
# ════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_shape(slide, Inches(4), Inches(2.5), Inches(2), Pt(4), ACCENT)

add_text_box(slide, Inches(1), Inches(2.8), Inches(11.3), Inches(1),
             "Thank You", font_size=52, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(4.0), Inches(11.3), Inches(0.6),
             "H&M 个性化推荐系统 · 数据分析报告", font_size=20, color=ACCENT, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(4.8), Inches(11.3), Inches(0.5),
             "付宝昊 · 梁智毅  |  2026 年 5 月", font_size=16, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(5.4), Inches(11.3), Inches(0.5),
             "github.com/UnknownAnt/H-M-recommendation", font_size=14, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)


# ── Save ──
output_path = r"d:\Un_Projects\H&M\H&M_个性化推荐系统_数据分析报告.pptx"
prs.save(output_path)
print(f"PPTX saved to: {output_path}")
