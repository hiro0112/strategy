# -*- coding: utf-8 -*-
"""
MeetMemo 事業提案資料（PowerPoint）を生成するスクリプト。
これまでに作成した以下3本のMarkdownレポートの内容を統合してスライド化する。
  - 事業検証レポート_MeetMemo_2026-09-13.md
  - 戦略設計_議事録SaaS・業務効率化_2026-09-13.md
  - 事業計画書_MeetMemo_草案_2026-09-13.md
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
import copy

# ------------------------------------------------------------------
# 基本設定
# ------------------------------------------------------------------
FONT_JP = "Yu Gothic UI"

NAVY = RGBColor(0x16, 0x24, 0x3E)
NAVY_LIGHT = RGBColor(0x24, 0x3B, 0x63)
TEAL = RGBColor(0x00, 0x9E, 0x8D)
AMBER = RGBColor(0xE8, 0xA0, 0x2E)
CORAL = RGBColor(0xE0, 0x5D, 0x5D)
GRAY_TEXT = RGBColor(0x4A, 0x4F, 0x58)
LIGHT_BG = RGBColor(0xF5, 0xF7, 0xFA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ROW_ALT = RGBColor(0xEE, 0xF2, 0xF6)
BORDER_GRAY = RGBColor(0xD9, 0xDE, 0xE4)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

section_no = 0  # 章番号カウンタ


def add_slide():
    return prs.slides.add_slide(BLANK)


def set_bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    # 最背面へ
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)
    return bg


def add_textbox(slide, l, t, w, h, text, size=18, color=GRAY_TEXT, bold=False,
                 align=PP_ALIGN.LEFT, font=FONT_JP, anchor=None, italic=False,
                 line_spacing=None):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    if anchor:
        tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.bold = bold
        r.font.italic = italic
        r.font.name = font
    return box


def add_bullets(slide, l, t, w, h, items, size=15, color=GRAY_TEXT, font=FONT_JP,
                 space_after=10, line_spacing=1.08, marker_color=None):
    """items: list of (text, level) or plain str（level=0）"""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if isinstance(item, tuple):
            text, level = item
        else:
            text, level = item, 0
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        marker = "● " if level == 0 else "‐ "
        r = p.add_run()
        r.text = marker + text
        r.font.size = Pt(size if level == 0 else size - 1)
        r.font.color.rgb = (marker_color if (marker_color and level == 0) else color)
        r.font.name = font
        r.font.bold = (level == 0)
        if level > 0:
            p.level = 0
            box2 = None
    return box


def header(slide, no, section, title, accent=TEAL):
    """章番号バー付きのスライドヘッダーを描画"""
    set_bg(slide, WHITE)
    # 上部の帯
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.12))
    bar.fill.solid(); bar.fill.fore_color.rgb = accent; bar.line.fill.background()
    bar.shadow.inherit = False

    # 章番号ラベル
    add_textbox(slide, Inches(0.55), Inches(0.32), Inches(2.5), Inches(0.4),
                f"{no:02d}  {section}", size=13, color=accent, bold=True)
    # タイトル
    add_textbox(slide, Inches(0.5), Inches(0.62), Inches(11.5), Inches(0.75),
                title, size=26, color=NAVY, bold=True)
    # 区切り線
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.35),
                                   Inches(12.2), Pt(1.4))
    line.fill.solid(); line.fill.fore_color.rgb = BORDER_GRAY; line.line.fill.background()
    line.shadow.inherit = False
    # フッター
    add_textbox(slide, Inches(0.55), Inches(7.12), Inches(6), Inches(0.3),
                "MeetMemo 事業提案資料", size=9, color=RGBColor(0xA9,0xAE,0xB6))
    add_textbox(slide, Inches(11.5), Inches(7.12), Inches(1.3), Inches(0.3),
                str(no), size=9, color=RGBColor(0xA9,0xAE,0xB6), align=PP_ALIGN.RIGHT)
    return


def styled_table(slide, l, t, w, h, headers, rows, col_widths=None,
                  header_bg=NAVY, font_size=12.5, header_size=13,
                  highlight_col=None, highlight_color=None):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    gshape = slide.shapes.add_table(n_rows, n_cols, l, t, w, h)
    table = gshape.table
    if col_widths:
        for i, cw in enumerate(col_widths):
            table.columns[i].width = cw
    # ヘッダー
    for c, htext in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = ""
        tf = cell.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = htext
        r.font.size = Pt(header_size); r.font.bold = True
        r.font.color.rgb = WHITE; r.font.name = FONT_JP
        cell.fill.solid(); cell.fill.fore_color.rgb = header_bg
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_top = Pt(4); cell.margin_bottom = Pt(4)
    # 本体
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.text = ""
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
            r = p.add_run(); r.text = str(val)
            r.font.size = Pt(font_size); r.font.name = FONT_JP
            bold_hl = (highlight_col is not None and ci == highlight_col)
            r.font.bold = bold_hl
            r.font.color.rgb = TEAL if bold_hl else RGBColor(0x33,0x38,0x40)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_top = Pt(3); cell.margin_bottom = Pt(3)
            cell.margin_left = Pt(6); cell.margin_right = Pt(6)
            base = ROW_ALT if ri % 2 == 0 else WHITE
            if highlight_col is not None and ci == highlight_col:
                base = RGBColor(0xE3, 0xF5, 0xF2)
            cell.fill.solid(); cell.fill.fore_color.rgb = base
    # 罫線を細く（デフォルトのテーマ枠のまま。簡易処理のため省略）
    return table


# ====================================================================
# Slide 1: 表紙
# ====================================================================
s = add_slide()
set_bg(s, NAVY)
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.55), SLIDE_W, Inches(0.06))
band.fill.solid(); band.fill.fore_color.rgb = TEAL; band.line.fill.background(); band.shadow.inherit = False

add_textbox(s, Inches(0.9), Inches(2.55), Inches(9), Inches(0.5),
            "事業提案資料", size=18, color=RGBColor(0x9F, 0xB4, 0xD4), bold=True)
add_textbox(s, Inches(0.85), Inches(2.95), Inches(11), Inches(1.4),
            "MeetMemo", size=60, color=WHITE, bold=True)
add_textbox(s, Inches(0.9), Inches(4.75), Inches(11), Inches(0.6),
            "会議録音をAIが自動で議事録化し、タスクを抽出して担当者に通知するSaaS",
            size=17, color=RGBColor(0xD7, 0xDF, 0xEA))
add_textbox(s, Inches(0.9), Inches(6.7), Inches(6), Inches(0.4),
            "2026年9月14日", size=13, color=RGBColor(0x8C, 0x9B, 0xB5))
add_textbox(s, Inches(0.9), Inches(6.7), Inches(11), Inches(0.4),
            "対象：課題検証／PEST・SWOT・3C分析／事業計画", size=13,
            color=RGBColor(0x8C, 0x9B, 0xB5), align=PP_ALIGN.RIGHT)

# ====================================================================
# Slide 2: 目次
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "AGENDA", "目次")
agenda_left = [
    "01  エグゼクティブサマリー",
    "02  課題と解決策",
    "03  外部環境分析（PEST）",
    "04  市場規模（TAM・SAM・SOM）",
    "05  3C分析",
]
agenda_right = [
    "06  SWOT分析",
    "07  競合比較",
    "08  差別化ポイント",
    "09  ビジネスモデル",
    "10  実行ロードマップ／財務計画",
]
add_bullets(s, Inches(0.9), Inches(1.9), Inches(5.6), Inches(4.5), agenda_left,
            size=18, color=NAVY, space_after=22)
add_bullets(s, Inches(6.9), Inches(1.9), Inches(5.6), Inches(4.5), agenda_right,
            size=18, color=NAVY, space_after=22)

# ====================================================================
# Slide 3: エグゼクティブサマリー
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "SUMMARY", "エグゼクティブサマリー")

cards = [
    ("課題", "議事録作成に年間約320時間。負担感67%に対し\nDX進捗はわずか1.4%というギャップ", CORAL),
    ("解決策", "録音→AI議事録化→タスク自動抽出→\n担当者へ自動通知を一気通貫で提供", TEAL),
    ("市場", "TAM 約1,030億円 ／ SAM 約400億円\nSOM 約6億円ARR（3〜5年後目標）", AMBER),
    ("収益モデル", "月額3,000円／ユーザーの定額課金\n中小企業向けの低価格・シンプル設計", NAVY_LIGHT),
]
cx = 0.55
cw = 2.95
for i, (label, body, color) in enumerate(cards):
    x = Inches(cx + i * (cw + 0.15))
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.65), Inches(cw), Inches(2.15))
    card.fill.solid(); card.fill.fore_color.rgb = LIGHT_BG
    card.line.color.rgb = color; card.line.width = Pt(1.5)
    card.shadow.inherit = False
    top = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.65), Inches(cw), Inches(0.42))
    top.fill.solid(); top.fill.fore_color.rgb = color; top.line.fill.background(); top.shadow.inherit = False
    tf = top.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = label; r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = FONT_JP
    add_textbox(s, x + Inches(0.15), Inches(2.25), Inches(cw - 0.3), Inches(1.5),
                body, size=12.5, color=RGBColor(0x33,0x38,0x40), line_spacing=1.2)

add_textbox(s, Inches(0.55), Inches(4.15), Inches(12.2), Inches(0.35),
            "MeetMemoが狙うポジション", size=14, color=NAVY, bold=True)
summary_bullets = [
    "Notta・Otter.ai・LINE WORKS AiNote等の既存競合は「文字起こし・要約」が中心で、タスクの自動アサイン・担当者通知は弱い／非対応",
    "従業員50名以下の中小企業に特化し、人手不足下での「タスクの抜け漏れ防止」という実行支援価値を訴求",
    "デジタル化・AI導入補助金2026の活用や低価格設計により、中小企業のIT予算制約を乗り越えて導入障壁を下げる",
    "初年度150社→3年目1,800社を目標に、Year3で単年度黒字化（ARR約2.5億円）を計画",
]
add_bullets(s, Inches(0.55), Inches(4.55), Inches(12.2), Inches(2.4), summary_bullets,
            size=14, color=GRAY_TEXT, space_after=10, marker_color=TEAL)

# ====================================================================
# Slide 4: 課題
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "PROBLEM", "課題：議事録作成とタスク管理の負担")

stat_cards = [
    ("約320時間", "議事録作成に費やす\n年間平均時間／人"),
    ("67%", "議事録作成に\n負担を感じている割合"),
    ("70%", "AIサポートツールの\n活用を希望する割合"),
    ("1.4%", "実際にDXが進んでいる\n現場の割合"),
]
cw2 = 2.85
for i, (num, label) in enumerate(stat_cards):
    x = Inches(0.55 + i * (cw2 + 0.2))
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.65), Inches(cw2), Inches(1.75))
    card.fill.solid(); card.fill.fore_color.rgb = NAVY
    card.line.fill.background(); card.shadow.inherit = False
    add_textbox(s, x, Inches(1.78), Inches(cw2), Inches(0.75), num, size=30, bold=True,
                color=AMBER if i in (0,3) else WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s, x + Inches(0.15), Inches(2.55), Inches(cw2 - 0.3), Inches(0.75), label,
                size=12, color=RGBColor(0xD7,0xDF,0xEA), align=PP_ALIGN.CENTER, line_spacing=1.1)

problem_bullets = [
    "60〜90分の会議1本の議事録化に2〜3時間、長いケースでは1日がかりという回答も複数存在",
    "中小企業は人手不足が深刻（正社員の人手不足感50.6%、2026年4月時点）で、少人数で会議運営から実行フォローまでを担う必要がある",
    "AI議事録ツール自体の利用率は63.3%まで進んでいるが、「特定の会議でしか使わない」限定利用が6割超という指摘があり、文字起こし・要約の先＝タスク実行支援が手つかずのまま",
    "→ 「言った・言わない」「やる予定を忘れていた」といったタスクの抜け漏れが、限られた人員体制の中小企業にとって経営リスクになっている",
]
add_bullets(s, Inches(0.55), Inches(3.75), Inches(12.2), Inches(3), problem_bullets,
            size=14.5, color=GRAY_TEXT, space_after=12, marker_color=CORAL)

# ====================================================================
# Slide 5: 解決策
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "SOLUTION", "解決策：議事録化からタスク通知までを一気通貫")

steps = [
    ("1", "自動録音・\n文字起こし", "Zoom/Teams/Google Meet等の\n会議音声をAIが自動録音・文字起こし", TEAL),
    ("2", "AI議事録化", "発言内容を要約し、決定事項・\n論点を整理した議事録を自動生成", TEAL),
    ("3", "タスク自動抽出", "「誰が・何を・いつまでに」を\nAIが会議中の発言から自動抽出", AMBER),
    ("4", "担当者へ自動通知", "Slack／メール等で担当者本人に\n即時通知し、期日リマインドも自動化", CORAL),
]
bw = 2.75
gap = 0.35
startx = 0.55
for i, (no, title, body, color) in enumerate(steps):
    x = Inches(startx + i * (bw + gap))
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(2.0), Inches(bw), Inches(2.6))
    box.fill.solid(); box.fill.fore_color.rgb = LIGHT_BG
    box.line.color.rgb = color; box.line.width = Pt(1.5); box.shadow.inherit = False
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(bw/2 - 0.32), Inches(2.25), Inches(0.64), Inches(0.64))
    circ.fill.solid(); circ.fill.fore_color.rgb = color; circ.line.fill.background(); circ.shadow.inherit = False
    tf = circ.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = no; r.font.size = Pt(22); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_JP
    add_textbox(s, x + Inches(0.1), Inches(3.05), Inches(bw - 0.2), Inches(0.6), title,
                size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER, line_spacing=1.05)
    add_textbox(s, x + Inches(0.15), Inches(3.7), Inches(bw - 0.3), Inches(0.85), body,
                size=11, color=GRAY_TEXT, align=PP_ALIGN.CENTER, line_spacing=1.15)
    if i < 3:
        arrow = s.shapes.add_shape(MSO_SHAPE.CHEVRON, x + Inches(bw + 0.03), Inches(2.95), Inches(0.3), Inches(0.5))
        arrow.fill.solid(); arrow.fill.fore_color.rgb = BORDER_GRAY; arrow.line.fill.background(); arrow.shadow.inherit = False

add_textbox(s, Inches(0.55), Inches(5.15), Inches(12.2), Inches(1.4),
            "議事録作成の工数削減だけでなく、「タスクの抜け漏れ」を防ぎ、限られた人員での実行力を底上げする。\n"
            "既存の文字起こし・要約ツールが手をつけていない「会議後の実行支援」を提供価値の中心に据える。",
            size=14.5, color=NAVY, line_spacing=1.3)

# ====================================================================
# Slide 6: PEST分析
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "PEST", "外部環境分析（PEST分析）")

pest = [
    ("Political｜政治", TEAL, [
        "「デジタル化・AI導入補助金2026」で導入費用を最大1/2補助（上限450万円）",
        "2026年4月、個人情報保護法改正案を閣議決定（課徴金制度の導入等）",
    ]),
    ("Economic｜経済", AMBER, [
        "人手不足倒産441件（2025年度、3年連続過去最多）",
        "DX予算確保の難しさが課題の24.9〜26.2%、特に20人以下企業は「何から始めるか分からない」が最多",
    ]),
    ("Social｜社会", CORAL, [
        "雇用型テレワーカー比率26.0%でハイブリッドワークが定着",
        "議事録DX進捗はわずか1.4%と、大きな未開拓領域が残存",
    ]),
    ("Technological｜技術", NAVY_LIGHT, [
        "LLMの軽量化・低コスト化（MoE等）が進み、推論コストを抑えた高精度化が進展",
        "Slack/Teams等との通知API連携が一般化し、実装障壁が低下",
    ]),
]
qw, qh = Inches(5.95), Inches(2.35)
positions = [(0.55, 1.65), (6.75, 1.65), (0.55, 4.15), (6.75, 4.15)]
for (label, color, items), (px, py) in zip(pest, positions):
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(py), qw, qh)
    box.fill.solid(); box.fill.fore_color.rgb = LIGHT_BG
    box.line.color.rgb = color; box.line.width = Pt(1.25); box.shadow.inherit = False
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(px), Inches(py), qw, Inches(0.42))
    bar.fill.solid(); bar.fill.fore_color.rgb = color; bar.line.fill.background(); bar.shadow.inherit = False
    tf = bar.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = "  " + label; r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = FONT_JP
    add_bullets(s, Inches(px + 0.2), Inches(py + 0.55), qw - Inches(0.4), qh - Inches(0.65),
                items, size=12, color=RGBColor(0x33,0x38,0x40), space_after=8, marker_color=color)

# ====================================================================
# Slide 7: 市場規模 TAM/SAM/SOM
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "MARKET SIZE", "市場規模（TAM・SAM・SOM）")

funnel = [
    ("TAM", "約1,030億円", "Web会議を利用する国内全企業（約172万社）", 9.6, TEAL),
    ("SAM", "約400億円", "IT投資意欲のある中堅・中小企業（約67万社）", 6.6, AMBER),
    ("SOM", "約6億円（3〜5年後ARR目標）", "現実的に獲得可能なシェア（約1万社）", 3.6, CORAL),
]
cy = 1.65
for label, amount, desc, w_in, color in funnel:
    x = Inches((13.333 - w_in) / 2)
    box = s.shapes.add_shape(MSO_SHAPE.TRAPEZOID if label == "SOM" else MSO_SHAPE.RECTANGLE,
                              x, Inches(cy), Inches(w_in), Inches(1.05))
    box.fill.solid(); box.fill.fore_color.rgb = color; box.line.fill.background(); box.shadow.inherit = False
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = f"{label}　{amount}"
    r.font.size = Pt(19); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_JP
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = desc
    r2.font.size = Pt(11.5); r2.font.color.rgb = RGBColor(0xFF,0xFF,0xFF); r2.font.name = FONT_JP
    cy += 1.2

add_textbox(s, Inches(0.55), Inches(5.35), Inches(12.2), Inches(0.35),
            "市場の成長トレンド", size=14, color=NAVY, bold=True)
trend_bullets = [
    "国内音声認識市場：2023年度に前年度比21.0%増の約150億円（ITR）、2025年度に約244億円へ（矢野経済研究所、CAGR16.4%）",
    "国内生成AI市場：2024年の約1,016億円から2028年に約8,028億円へ約8倍に拡大予測",
    "グローバルAI Meeting Assistant市場：CAGR25.8%で2025年約34.7億ドル→2033年約214.8億ドルへ拡大予測",
]
add_bullets(s, Inches(0.55), Inches(5.75), Inches(12.2), Inches(1.5), trend_bullets,
            size=12.5, color=GRAY_TEXT, space_after=6, marker_color=TEAL)

# ====================================================================
# Slide 8: 3C分析
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "3C", "3C分析")

three_c = [
    ("Customer｜市場・顧客", TEAL, [
        "従業員50名以下の中小企業（国内336.5万社、Web会議導入率51.0%）",
        "議事録の負担感67%・DX進捗1.4%という大きなギャップ層",
        "人手不足＋予算制約から、低価格・低学習コストのツールへの受容性が高い",
    ]),
    ("Competitor｜競合", CORAL, [
        "Notta・Otter.ai・LINE WORKS AiNoteはいずれも文字起こし・要約が中心",
        "Fireflies.aiは公式に「タスク自動アサイン・リマインド手段がない」と明言",
        "Zoom／Teams等大手プラットフォームのAI要約標準搭載も間接競合化",
    ]),
    ("Company｜自社", NAVY_LIGHT, [
        "議事録化＋タスク自動抽出＋担当者通知を一気通貫で提供",
        "従業員50名以下に特化したシンプル・低価格設計（3,000円/ユーザー/月）",
        "補助金対象ツール登録によりさらに価格競争力を強化",
    ]),
]
cw3 = 3.95
for i, (label, color, items) in enumerate(three_c):
    x = Inches(0.55 + i * (cw3 + 0.2))
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.65), Inches(cw3), Inches(4.9))
    box.fill.solid(); box.fill.fore_color.rgb = LIGHT_BG
    box.line.color.rgb = color; box.line.width = Pt(1.25); box.shadow.inherit = False
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.65), Inches(cw3), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = color; bar.line.fill.background(); bar.shadow.inherit = False
    tf = bar.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = label; r.font.size = Pt(14.5); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = FONT_JP
    add_bullets(s, x + Inches(0.2), Inches(2.35), Inches(cw3 - 0.4), Inches(4.1),
                items, size=12.5, color=RGBColor(0x33,0x38,0x40), space_after=14, marker_color=color)

# ====================================================================
# Slide 9: SWOT分析
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "SWOT", "SWOT分析")

swot = [
    ("Strengths｜強み", TEAL, [
        "タスク自動抽出＋担当者通知の一気通貫提供（競合の空白地帯）",
        "50名以下企業に特化したシンプルUI・低価格設計",
        "後発ゆえ最新の低コストLLM／音声認識を前提に開発可能",
    ]),
    ("Weaknesses｜弱み", CORAL, [
        "新規参入のためブランド認知・導入実績が不足",
        "資金力・開発リソースで先行競合に劣る",
        "音声認識精度・誤タスク抽出対策など作り込みがこれから",
    ]),
    ("Opportunities｜機会", AMBER, [
        "［政治］補助金2026で導入費用を最大1/2補助",
        "［経済］人手不足倒産増加で効率化ニーズ拡大",
        "［社会］ハイブリッドワーク定着、DX進捗1.4%の未開拓市場",
        "［技術］LLM低コスト化で開発優位性",
    ]),
    ("Threats｜脅威", NAVY_LIGHT, [
        "［政治］個人情報保護法改正で規制対応コスト増",
        "［経済］中小企業のIT予算確保の難しさ",
        "［社会］大手プラットフォームのAI機能無償化",
        "［技術］技術キャッチアップの速さによる同質化",
    ]),
]
positions2 = [(0.55, 1.65), (6.75, 1.65), (0.55, 4.3), (6.75, 4.3)]
qh2 = Inches(2.5)
for (label, color, items), (px, py) in zip(swot, positions2):
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(px), Inches(py), qw, qh2)
    box.fill.solid(); box.fill.fore_color.rgb = LIGHT_BG
    box.line.color.rgb = color; box.line.width = Pt(1.25); box.shadow.inherit = False
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(px), Inches(py), qw, Inches(0.4))
    bar.fill.solid(); bar.fill.fore_color.rgb = color; bar.line.fill.background(); bar.shadow.inherit = False
    tf = bar.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = "  " + label; r.font.size = Pt(13.5); r.font.bold = True
    r.font.color.rgb = WHITE; r.font.name = FONT_JP
    add_bullets(s, Inches(px + 0.2), Inches(py + 0.52), qw - Inches(0.4), qh2 - Inches(0.6),
                items, size=11, color=RGBColor(0x33,0x38,0x40), space_after=5, marker_color=color)

# ====================================================================
# Slide 10: 競合比較
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "COMPETITION", "競合比較")

headers = ["項目", "Notta", "Otter.ai", "LINE WORKS\nAiNote", "MeetMemo（自社）"]
rows = [
    ["料金目安", "2,508円〜\n／ユーザー／月", "約4,500円\n／ユーザー／月", "月19,800円〜\n（チーム・100時間）", "3,000円\n／ユーザー／月"],
    ["文字起こし・要約", "○", "○（AI Chat有）", "○（精度90%以上）", "○"],
    ["タスク自動抽出", "△", "△", "情報限定的", "◎"],
    ["担当者への自動通知", "×", "×", "×（公開情報上）", "◎"],
    ["ターゲット", "個人〜チーム全般", "個人〜チーム\n（海外中心）", "LINE WORKS\n利用企業", "中小企業\n（50名以下）特化"],
]
col_widths = [Inches(2.2), Inches(2.15), Inches(2.15), Inches(2.35), Inches(2.45)]
styled_table(s, Inches(0.55), Inches(1.7), Inches(11.3), Inches(3.1),
             headers, rows, col_widths=col_widths, highlight_col=4, font_size=12, header_size=12.5)

add_textbox(s, Inches(0.55), Inches(5.15), Inches(11.3), Inches(1.7),
            "既存競合は文字起こし・要約の精度競争が中心で、会議後の「タスク実行支援」を主軸に据えたプレイヤーは少ない。\n"
            "MeetMemoはこの空白地帯を主戦場とし、料金面でもユーザー数×定額のシンプルな設計で中小企業の予算制約に対応する。",
            size=13.5, color=NAVY, line_spacing=1.3)

# ====================================================================
# Slide 11: 差別化ポイント
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "DIFFERENTIATION", "差別化ポイント（自社が勝てる3つの軸）")

diffs = [
    ("軸1", "タスク自動化への特化", "会議の「その後」まで自動化。競合が\n手薄なタスク抽出・担当者通知を主軸に", TEAL),
    ("軸2", "中小企業向け低価格設計", "ユーザー数×定額3,000円で予算化しやすく、\n従量課金の不透明さを回避", AMBER),
    ("軸3", "補助金活用による価格優位", "デジタル化・AI導入補助金2026の対象登録で\n実質導入コストをさらに引き下げ", CORAL),
]
bw3 = 3.85
for i, (no, title, body, color) in enumerate(diffs):
    x = Inches(0.55 + i * (bw3 + 0.25))
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.85), Inches(bw3), Inches(3.4))
    card.fill.solid(); card.fill.fore_color.rgb = LIGHT_BG
    card.line.color.rgb = color; card.line.width = Pt(1.5); card.shadow.inherit = False
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.25), Inches(2.1), Inches(0.7), Inches(0.7))
    badge.fill.solid(); badge.fill.fore_color.rgb = color; badge.line.fill.background(); badge.shadow.inherit = False
    tf = badge.text_frame; p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = no; r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_JP
    add_textbox(s, x + Inches(0.2), Inches(3.0), Inches(bw3 - 0.4), Inches(0.7), title,
                size=16.5, bold=True, color=NAVY, line_spacing=1.1)
    add_textbox(s, x + Inches(0.2), Inches(3.75), Inches(bw3 - 0.4), Inches(1.35), body,
                size=12.5, color=GRAY_TEXT, line_spacing=1.25)

# ====================================================================
# Slide 12: ビジネスモデル
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "BUSINESS MODEL", "ビジネスモデル")

add_textbox(s, Inches(0.55), Inches(1.6), Inches(5.7), Inches(0.35), "収益モデル", size=15, bold=True, color=NAVY)
biz_left = [
    "月額課金（サブスクリプション）：3,000円／ユーザー／月",
    "契約単位：企業単位（平均利用ユーザー数5〜8名／社を想定）",
    "支払サイクル：月額／年額（年払いで割引を想定）",
    "デジタル化・AI導入補助金2026の対象ツール登録を目指し、実質導入コストを引き下げ",
]
add_bullets(s, Inches(0.55), Inches(2.0), Inches(5.7), Inches(2.7), biz_left,
            size=13, color=GRAY_TEXT, space_after=12, marker_color=TEAL)

add_textbox(s, Inches(6.75), Inches(1.6), Inches(5.7), Inches(0.35), "コスト構造", size=15, bold=True, color=NAVY)
biz_right = [
    "変動費：音声認識・LLM API利用料（会議時間・ユーザー数に比例）",
    "変動費：Slack／メール等の通知連携API費用",
    "固定費：開発・エンジニア人件費、営業・マーケティング費",
    "固定費：クラウドインフラ基盤費、カスタマーサポート費",
]
add_bullets(s, Inches(6.75), Inches(2.0), Inches(5.7), Inches(2.7), biz_right,
            size=13, color=GRAY_TEXT, space_after=12, marker_color=AMBER)

headers = ["比較軸", "MeetMemo", "参考：競合水準"]
rows = [
    ["個人〜チーム向け単価", "3,000円／ユーザー／月",
     "Notta 2,508円〜／ユーザー／月\nOtter.ai Business 約4,500円／ユーザー／月"],
    ["契約形態", "ユーザー課金でシンプル・予測可能",
     "LINE WORKS AiNoteは時間従量制\n（チーム月19,800円〜、100時間まで）"],
]
styled_table(s, Inches(0.55), Inches(5.0), Inches(11.9), Inches(1.8), headers, rows,
             col_widths=[Inches(3.0), Inches(4.4), Inches(4.5)], highlight_col=1, font_size=12)

# ====================================================================
# Slide 13: 実行ロードマップ
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "ROADMAP", "実行ロードマップ（最初の1年）")

phases = [
    ("Q1", "1〜3ヶ月目", "MVP開発・検証", [
        "コア機能（録音〜議事録化〜タスク抽出〜通知）をMVPとして開発",
        "中小企業5〜10社にPoC協力を依頼し精度・効果を検証",
    ], TEAL),
    ("Q2", "4〜6ヶ月目", "ベータ版・限定リリース", [
        "PoCフィードバックを反映しベータ版をリリース",
        "限定招待制で30〜50社規模の有料トライアル開始",
        "補助金対象ツール登録申請を開始",
    ], AMBER),
    ("Q3", "7〜9ヶ月目", "正式ローンチ", [
        "正式版を一般提供開始、料金プランを確定",
        "SaaS比較サイト・商工会議所等への掲載・連携",
    ], CORAL),
    ("Q4", "10〜12ヶ月目", "拡販・体制強化", [
        "有料導入150社（期末）を目標に営業強化",
        "活用事例・口コミの収集とシード資金調達の検討",
    ], NAVY_LIGHT),
]
bw4 = 2.95
for i, (q, period, title, items, color) in enumerate(phases):
    x = Inches(0.55 + i * (bw4 + 0.15))
    top = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.65), Inches(bw4), Inches(0.85))
    top.fill.solid(); top.fill.fore_color.rgb = color; top.line.fill.background(); top.shadow.inherit = False
    tf = top.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = q; r.font.size = Pt(20); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_JP
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = period; r2.font.size = Pt(10.5); r2.font.color.rgb = WHITE; r2.font.name = FONT_JP

    body = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(2.5), Inches(bw4), Inches(3.4))
    body.fill.solid(); body.fill.fore_color.rgb = LIGHT_BG
    body.line.color.rgb = color; body.line.width = Pt(1.25); body.shadow.inherit = False
    add_textbox(s, x + Inches(0.15), Inches(2.62), Inches(bw4 - 0.3), Inches(0.55), title,
                size=13.5, bold=True, color=NAVY, line_spacing=1.1)
    add_bullets(s, x + Inches(0.15), Inches(3.2), Inches(bw4 - 0.3), Inches(2.6), items,
                size=10.8, color=RGBColor(0x33,0x38,0x40), space_after=8, marker_color=color, line_spacing=1.15)

# ====================================================================
# Slide 14: 財務計画
# ====================================================================
s = add_slide(); section_no += 1
header(s, section_no, "FINANCIALS", "財務計画（初年度〜3年目の収支概算）")

headers = ["項目", "Year1", "Year2", "Year3"]
rows = [
    ["有料導入社数（期末）", "150社", "600社", "1,800社"],
    ["年間売上高（概算）", "約1,350万円", "約6,750万円", "約2億4,840万円"],
    ["総コスト（人件費・開発・マーケ等）", "約3,400万円", "約7,900万円", "約1億5,100万円"],
    ["営業損益（概算）", "約▲2,050万円", "約▲1,150万円", "約＋9,740万円"],
]
styled_table(s, Inches(0.55), Inches(1.65), Inches(7.3), Inches(2.7), headers, rows,
             col_widths=[Inches(3.1), Inches(1.4), Inches(1.4), Inches(1.4)], font_size=12.5)

# チャート（売上高・総コストの推移）
chart_data = CategoryChartData()
chart_data.categories = ["Year1", "Year2", "Year3"]
chart_data.add_series("年間売上高（万円）", (1350, 6750, 24840))
chart_data.add_series("総コスト（万円）", (3400, 7900, 15100))
gframe = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(8.1), Inches(1.65),
                             Inches(4.75), Inches(3.1), chart_data)
chart = gframe.chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False
chart.legend.font.size = Pt(10)
chart.legend.font.name = FONT_JP
cat_ax = chart.category_axis
cat_ax.tick_labels.font.size = Pt(10.5); cat_ax.tick_labels.font.name = FONT_JP
val_ax = chart.value_axis
val_ax.tick_labels.font.size = Pt(9.5); val_ax.tick_labels.font.name = FONT_JP
val_ax.has_major_gridlines = False
colors = [TEAL, RGBColor(0xC9, 0xCE, 0xD6)]
for i, series in enumerate(chart.plots[0].series):
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = colors[i]

add_textbox(s, Inches(0.55), Inches(4.55), Inches(12.2), Inches(0.35),
            "資金計画の考え方", size=14, color=NAVY, bold=True)
fin_bullets = [
    "Year1〜2は開発・PoC・初期営業投資が先行し赤字を見込むため、シード〜シリーズA相当の外部資金調達（目安3,000万〜5,000万円）を想定",
    "Year3で単年度黒字化を見込み、SOM（3〜5年後ARR約6億円）に向けた成長軌道に乗せることを目指す",
    "補助金対象ツール登録や年払い比率の向上により、キャッシュフローの改善を図る",
]
add_bullets(s, Inches(0.55), Inches(4.95), Inches(12.2), Inches(2), fin_bullets,
            size=13, color=GRAY_TEXT, space_after=8, marker_color=TEAL)

# ====================================================================
# Slide 15: まとめ
# ====================================================================
s = add_slide(); section_no += 1
set_bg(s, NAVY)
bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, Inches(0.12))
bar.fill.solid(); bar.fill.fore_color.rgb = TEAL; bar.line.fill.background(); bar.shadow.inherit = False

add_textbox(s, Inches(0.6), Inches(0.55), Inches(6), Inches(0.4), "10  SUMMARY", size=13, color=TEAL, bold=True)
add_textbox(s, Inches(0.55), Inches(0.85), Inches(11), Inches(0.75), "まとめ・Next Action", size=28, color=WHITE, bold=True)

summary2 = [
    "課題仮説・市場規模・競合分析のいずれからも、「タスク自動抽出＋担当者通知」という機能ギャップは事業機会として妥当性が高い",
    "中小企業のIT予算制約に対しては、低価格の定額課金モデルと補助金活用の両面でアプローチする",
    "Year1はPoC・MVP開発を優先し、Year3での単年度黒字化（ARR約2.5億円）を目指すロードマップを設定",
]
add_bullets(s, Inches(0.6), Inches(1.95), Inches(11.6), Inches(2.1), summary2,
            size=15.5, color=RGBColor(0xE3, 0xE8, 0xF0), space_after=14, marker_color=TEAL)

next_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(4.35), Inches(12.1), Inches(2.3))
next_box.fill.solid(); next_box.fill.fore_color.rgb = NAVY_LIGHT
next_box.line.fill.background(); next_box.shadow.inherit = False
add_textbox(s, Inches(0.9), Inches(4.55), Inches(11), Inches(0.4), "Next Action", size=15, bold=True, color=AMBER)
next_actions = [
    "ターゲット企業（従業員10〜300名規模）へのインタビュー・PoC実施による価格受容性の検証",
    "Slack／Teams／メール通知機能を含むMVPの開発着手",
    "デジタル化・AI導入補助金2026の対象ツール登録に向けた準備",
]
add_bullets(s, Inches(0.9), Inches(5.0), Inches(11.4), Inches(1.5), next_actions,
            size=13.5, color=WHITE, space_after=8, marker_color=AMBER)

out_path = r"C:\Users\hfuka\OneDrive\ドキュメント\0.AI-claude\strategy\MeetMemo_事業提案資料_2026-09-14.pptx"
prs.save(out_path)
print("saved:", out_path)
