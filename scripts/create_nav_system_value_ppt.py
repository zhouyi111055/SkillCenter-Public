"""Generate an editable blue/white business PPTX from the provided reference image.

The slide is intentionally built from PowerPoint-native editable objects (text boxes,
rectangles, lines, arrows, tables, and simple vector icons) instead of using a full-slide
screenshot background.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path("outputs/nav-system-value/nav_system_value_editable.pptx")

SLIDE_W = 13.333
SLIDE_H = 7.5

NAVY = RGBColor(0, 45, 105)
DARK_NAVY = RGBColor(0, 31, 78)
BLUE = RGBColor(0, 86, 184)
MID_BLUE = RGBColor(31, 104, 191)
LIGHT_BLUE = RGBColor(231, 241, 255)
PALE_BLUE = RGBColor(246, 250, 255)
YELLOW = RGBColor(255, 211, 35)
GREEN = RGBColor(35, 168, 72)
GREY = RGBColor(115, 130, 150)
LIGHT_GREY = RGBColor(226, 233, 242)
BLACK = RGBColor(19, 29, 47)
WHITE = RGBColor(255, 255, 255)
ORANGE = RGBColor(245, 146, 44)
RED = RGBColor(225, 54, 67)
FONT = "Yu Gothic"


def add_text(slide, text, x, y, w, h, size=14, color=BLACK, bold=False,
             align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, font=FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_rect(slide, x, y, w, h, fill=WHITE, line=LIGHT_GREY, radius=False, width=1,
             transparency=0):
    shape_type = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.fill.transparency = transparency
    shp.line.color.rgb = line
    shp.line.width = Pt(width)
    return shp


def add_line(slide, x1, y1, x2, y2, color=LIGHT_GREY, width=1.0):
    ln = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color
    ln.line.width = Pt(width)
    return ln


def add_section_title(slide, n, title, x, y, w):
    add_rect(slide, x, y, w, 0.34, fill=RGBColor(248, 251, 255), line=RGBColor(238, 243, 250), radius=False)
    add_line(slide, x, y, x, y + 0.34, BLUE, 2.0)
    add_text(slide, f"{n}. {title}", x + 0.1, y + 0.06, w - 0.15, 0.24, size=13, color=DARK_NAVY, bold=True)


def add_card(slide, x, y, w, h):
    return add_rect(slide, x, y, w, h, fill=WHITE, line=RGBColor(205, 218, 235), radius=True, width=1)


def icon_circle_check(slide, cx, cy, r=0.11, color=GREEN):
    circ = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(cx-r), Inches(cy-r), Inches(2*r), Inches(2*r))
    circ.fill.solid(); circ.fill.fore_color.rgb = color
    circ.line.color.rgb = color
    add_text(slide, "✓", cx-r*0.55, cy-r*0.75, 2*r, 2*r, size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)


def icon_target(slide, x, y, s=0.32):
    for i, scale in enumerate([1, .66, .32]):
        c = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x+s*(1-scale)/2), Inches(y+s*(1-scale)/2), Inches(s*scale), Inches(s*scale))
        c.fill.background(); c.line.color.rgb = MID_BLUE; c.line.width = Pt(1.6)
    add_line(slide, x+s*.2, y+s*.8, x+s*.83, y+s*.17, MID_BLUE, 1.7)
    arr = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RIGHT_TRIANGLE, Inches(x+s*.71), Inches(y+s*.08), Inches(s*.18), Inches(s*.18))
    arr.rotation = 45; arr.fill.solid(); arr.fill.fore_color.rgb = MID_BLUE; arr.line.color.rgb = MID_BLUE


def icon_clock(slide, x, y, s=0.32):
    c = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x), Inches(y), Inches(s), Inches(s))
    c.fill.background(); c.line.color.rgb = MID_BLUE; c.line.width = Pt(1.5)
    add_line(slide, x+s/2, y+s/2, x+s/2, y+s*.22, MID_BLUE, 1.5)
    add_line(slide, x+s/2, y+s/2, x+s*.72, y+s*.66, MID_BLUE, 1.5)


def icon_bars(slide, x, y, s=0.34):
    for i, h in enumerate([.13, .22, .31]):
        add_rect(slide, x+i*s*.28, y+s-h, s*.15, h, fill=MID_BLUE, line=MID_BLUE)
    add_line(slide, x, y+s*.72, x+s*.8, y+s*.16, MID_BLUE, 1.8)
    tri = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RIGHT_TRIANGLE, Inches(x+s*.68), Inches(y+s*.06), Inches(s*.16), Inches(s*.16))
    tri.rotation = 315; tri.fill.solid(); tri.fill.fore_color.rgb = MID_BLUE; tri.line.color.rgb = MID_BLUE


def icon_person(slide, x, y, s=0.34):
    head = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x+s*.31), Inches(y), Inches(s*.28), Inches(s*.28))
    head.fill.background(); head.line.color.rgb = MID_BLUE; head.line.width = Pt(1.4)
    body = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ARC, Inches(x+s*.1), Inches(y+s*.22), Inches(s*.7), Inches(s*.55))
    body.line.color.rgb = MID_BLUE; body.line.width = Pt(1.4)
    add_line(slide, x+s*.78, y+s*.17, x+s*.98, y+s*.04, MID_BLUE, 1.3)
    add_line(slide, x+s*.92, y+s*.04, x+s*.92, y+s*.22, MID_BLUE, 1.3)


def icon_clipboard(slide, x, y, s=0.33):
    add_rect(slide, x+s*.18, y+s*.11, s*.58, s*.77, fill=WHITE, line=MID_BLUE, radius=True, width=1.4)
    add_rect(slide, x+s*.32, y, s*.3, s*.17, fill=WHITE, line=MID_BLUE, radius=True, width=1.4)
    for i in range(3):
        add_line(slide, x+s*.3, y+s*(.35+i*.15), x+s*.65, y+s*(.35+i*.15), MID_BLUE, .9)


def left_icon(slide, kind, x, y):
    if kind == "folder":
        add_rect(slide, x, y+0.05, .24, .16, fill=WHITE, line=MID_BLUE, radius=True, width=1.4)
        add_rect(slide, x+.03, y, .12, .08, fill=WHITE, line=MID_BLUE, radius=True, width=1.4)
    elif kind == "process":
        for i in range(3):
            add_rect(slide, x, y+i*.08, .22, .035, fill=WHITE, line=MID_BLUE, radius=True, width=1.2)
            add_line(slide, x+.05, y+i*.08+.035, x+.05, y+i*.08+.065, MID_BLUE, 1.0)
        add_line(slide, x+.16, y+.04, x+.25, y+.04, MID_BLUE, 1.0)
        add_line(slide, x+.16, y+.12, x+.25, y+.12, MID_BLUE, 1.0)
    elif kind == "clipboard":
        icon_clipboard(slide, x-.02, y-.02, .33)
    elif kind == "compass":
        c = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x), Inches(y), Inches(.25), Inches(.25))
        c.fill.background(); c.line.color.rgb = MID_BLUE; c.line.width = Pt(1.4)
        add_line(slide, x+.07, y+.18, x+.19, y+.06, MID_BLUE, 1.3)
    else:
        icon_circle_check(slide, x+.12, y+.12, .13, GREEN)


def add_mock_ui(slide, x, y, w, h):
    add_card(slide, x, y, w, h)
    add_rect(slide, x, y, .78, h, fill=DARK_NAVY, line=DARK_NAVY, radius=True)
    for i, item in enumerate(["ホーム", "案件一覧", "工程一覧", "チェックポイント", "ツール一覧", "テンプレート", "社内ナレッジ", "管理メニュー"]):
        add_text(slide, "▣  " + item, x+.12, y+.12+i*.28, .55, .12, size=4.9, color=WHITE, bold=(i==0))
    add_text(slide, "チェックポイント実施", x+.95, y+.12, w-.95, .17, size=6.3, color=BLACK, bold=True)
    add_text(slide, "5-3　テスト環境構築・動作確認", x+.95, y+.39, 2.3, .15, size=5.7, color=BLACK, bold=True)
    add_text(slide, "案件名　A案件_基幹システム刷新　　 工程　結合テスト　　 担当者　田中 花子　　 期限　2025/05/18", x+.95, y+.67, 3.95, .12, size=4.6, color=BLACK)
    add_rect(slide, x+w-.95, y+.65, .62, .07, fill=MID_BLUE, line=MID_BLUE, radius=True)
    add_text(slide, "65%", x+w-.28, y+.62, .18, .09, size=4.6, color=BLACK)
    add_rect(slide, x+1.03, y+.9, 1.25, h-1.1, fill=RGBColor(250,252,255), line=LIGHT_GREY)
    for i, lab in enumerate(["確認内容", "実施手順", "使用ツール", "標準・テンプレート", "注意事項", "完了確認"]):
        yy = y+1.08+i*.35
        c = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(x+1.1), Inches(yy), Inches(.12), Inches(.12))
        c.fill.solid(); c.fill.fore_color.rgb = GREEN if i==5 else MID_BLUE; c.line.color.rgb = c.fill.fore_color.rgb
        add_text(slide, str(i+1), x+1.125, yy+.018, .07, .06, size=3.9, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, lab, x+1.28, yy-.01, .78, .1, size=4.7, color=BLACK, bold=i in [0,1])
        add_text(slide, "▣", x+2.12, yy-.01, .08, .08, size=5.2, color=GREEN if i<4 else RED)
    add_rect(slide, x+2.43, y+.9, w-2.6, .73, fill=WHITE, line=LIGHT_GREY, radius=True)
    add_text(slide, "STEP1  確認内容", x+2.55, y+1.0, .8, .14, size=5.4, color=BLACK, bold=True)
    add_rect(slide, x+3.26, y+1.0, .18, .11, fill=RGBColor(228,250,235), line=GREEN, radius=True)
    add_text(slide, "完了", x+3.29, y+1.015, .1, .05, size=3.5, color=GREEN, bold=True)
    add_text(slide, "作業内容　　 結合テストに必要なテスト環境を構築し、システムが正常に起動・動作することを確認する。", x+2.55, y+1.23, w-2.85, .12, size=4.1, color=BLACK)
    add_text(slide, "確認観点　　 ・テスト環境が設計書通りに構築されていること\n　　　　　　 ・各ミドルウェアが正常に起動していること\n　　　　　　 ・基本機能が実行できること", x+2.55, y+1.44, w-2.85, .35, size=4.0, color=BLACK)
    for i, lab in enumerate(["STEP2  実施手順", "STEP3  使用ツール", "STEP4  標準・テンプレート", "STEP5  注意事項", "STEP6  完了確認"]):
        yy = y+1.74+i*.28
        add_rect(slide, x+2.43, yy, w-2.6, .21, fill=WHITE, line=LIGHT_GREY, radius=True)
        add_text(slide, lab, x+2.55, yy+.055, 1.4, .08, size=4.7, color=BLACK, bold=True)
        add_text(slide, "完了" if i<4 else "未完了", x+w-.95, yy+.055, .25, .08, size=3.7, color=GREEN if i<4 else ORANGE, bold=True)
        add_text(slide, "⌄", x+w-.32, yy+.04, .09, .08, size=6, color=BLUE, bold=True)
    for i, (txt, col) in enumerate([("スキップする", MID_BLUE), ("差戻しにする", ORANGE), ("完了する", BLUE)]):
        add_rect(slide, x+3.45+i*.95, y+h-.24, .7, .15, fill=WHITE if i<2 else BLUE, line=col, radius=True)
        add_text(slide, txt, x+3.55+i*.95, y+h-.205, .5, .06, size=3.6, color=col if i<2 else WHITE, bold=True, align=PP_ALIGN.CENTER)


def build():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, fill=WHITE, line=WHITE)
    add_rect(slide, 0, 0, SLIDE_W, .88, fill=NAVY, line=NAVY)
    add_rect(slide, 0, .77, SLIDE_W, .08, fill=RGBColor(222, 234, 250), line=RGBColor(222, 234, 250), transparency=18)
    add_text(slide, "IT開発向けナビ型システムの価値と効果", 0, .12, SLIDE_W, .34, size=24, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, "作業を迷わず、漏れなく、効率的に。標準化と品質向上を実現する「実行ナビ型プラットフォーム」", 0, .56, SLIDE_W, .22, size=13.5, color=YELLOW, bold=True, align=PP_ALIGN.CENTER)

    # 1. left and center
    add_section_title(slide, 1, "システムの特徴（ナビ型の強み）", .12, .98, 2.85)
    flow_y = [1.50, 2.08, 2.70, 3.31, 3.92]
    labels = ["案件 から", "工程 へ", "チェックポイント で", "作業をナビゲート し", "完了までをサポート"]
    kinds = ["folder", "process", "clipboard", "compass", "check"]
    for i, yy in enumerate(flow_y):
        left_icon(slide, kinds[i], .34, yy)
        add_text(slide, labels[i], .78, yy+.04, 1.45, .16, size=9.4, color=DARK_NAVY, bold=True)
        if i < len(flow_y)-1:
            add_line(slide, .46, yy+.34, .46, flow_y[i+1]-.1, MID_BLUE, 1.2)
    add_mock_ui(slide, 2.32, 1.32, 6.18, 3.55)

    # 2. right effects
    add_section_title(slide, 2, "このシステムで得られる効果", 8.72, .98, 4.15)
    effects = [
        (icon_target, "作業の抜け漏れ防止", "チェックポイントごとに必要な内容を順番にナビゲート。\n未実施の項目は完了できない仕組みで漏れを防止。"),
        (icon_clock, "作業効率の向上", "手順・ツール・テンプレートへワンクリックでアクセス。\n探す時間を削減し、作業に集中できる。"),
        (icon_bars, "品質の標準化・向上", "標準手順・注意事項を全員が同じ順番で実施。\n属人化を排除し、品質のばらつきを低減。"),
        (icon_person, "教育・引継ぎの効率化", "新人でも迷わず作業できる構成。\n教育コストを削減し、早期戦力化を実現。"),
        (icon_clipboard, "実施履歴の可視化", "誰が・いつ・何を実施したかを自動記録。\n管理・監査・品質改善に活用できる。"),
    ]
    for i, (fn, title, body) in enumerate(effects):
        yy = 1.40 + i*.70
        add_rect(slide, 8.72, yy, 4.15, .60, fill=RGBColor(249,251,255), line=RGBColor(249,251,255), radius=True)
        fn(slide, 8.88, yy+.11, .33)
        add_text(slide, title, 9.36, yy+.10, 2.5, .16, size=9.3, color=DARK_NAVY, bold=True)
        add_text(slide, body, 9.36, yy+.32, 3.15, .22, size=6.8, color=BLACK)

    # 3. table
    add_card(slide, .12, 4.98, 4.65, 2.48)
    add_section_title(slide, 3, "従来の課題と解決", .12, 4.98, 4.65)
    add_rect(slide, .22, 5.40, 4.45, .25, fill=RGBColor(242,246,252), line=LIGHT_GREY)
    add_text(slide, "よくある課題（従来）", .62, 5.47, 1.25, .09, size=6.8, color=DARK_NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, "本システムでの解決", 2.92, 5.47, 1.25, .09, size=6.8, color=DARK_NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_line(slide, 2.26, 5.40, 2.26, 7.37, LIGHT_GREY, 1)
    rows = [
        ("手順や資料が分散している", "必要な情報を1画面に集約"),
        ("最新版がどれかわからない", "常に最新の標準を一元管理"),
        ("ツールの起動が面倒", "画面から直接起動"),
        ("作業の抜け漏れが発生する", "順番ナビ＋未完了ロックで防止"),
        ("属人化していて引継ぎが大変", "標準化された手順で誰でも実施"),
        ("実施状況の把握が困難", "履歴の自動記録で可視化"),
    ]
    for i, (a, b) in enumerate(rows):
        yy = 5.66+i*.28
        add_line(slide, .22, yy-.02, 4.67, yy-.02, LIGHT_GREY, .7)
        add_text(slide, "▣", .31, yy+.04, .12, .08, size=7, color=MID_BLUE, bold=True)
        add_text(slide, a, .58, yy+.04, 1.55, .09, size=6.5, color=BLACK)
        add_text(slide, "→", 2.25, yy+.035, .20, .08, size=9, color=BLUE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, b, 2.55, yy+.04, 1.85, .09, size=6.5, color=BLACK)

    # 4. features
    add_card(slide, 4.86, 4.98, 3.88, 2.48)
    add_section_title(slide, 4, "主な機能", 4.86, 4.98, 3.88)
    features = [
        ("工程一覧表示機能", "開発フロー全体を一覧で表示し、対象工程へ迷わずアクセス可能"),
        ("工程別チェックポイント表示機能", "工程ごとに実施すべきチェックポイントを一覧表示"),
        ("ツール連携機能（案内＋起動）", "使用ツールの案内と直接起動で作業効率を向上"),
        ("標準手順参照機能", "標準手順書や運用ルール、補足資料をすぐに参照可能"),
        ("マスタ設定機能", "工程・チェックポイント・ツール等を柔軟に設定・更新可能"),
    ]
    for i, (t, b) in enumerate(features):
        yy = 5.50+i*.36
        circ = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(5.02), Inches(yy-.02), Inches(.26), Inches(.26))
        circ.fill.solid(); circ.fill.fore_color.rgb = LIGHT_BLUE; circ.line.color.rgb = LIGHT_BLUE
        add_text(slide, ["▦", "▤", "🔧", "↗", "≋"][i], 5.055, yy+.015, .18, .09, size=8, color=MID_BLUE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, t, 5.40, yy-.02, 2.45, .12, size=8.1, color=DARK_NAVY, bold=True)
        add_text(slide, b, 5.40, yy+.13, 2.95, .11, size=5.6, color=BLACK)

    # 5. utilization points
    add_card(slide, 8.82, 4.98, 4.10, 2.48)
    add_section_title(slide, 5, "利用率を高めるポイント", 8.82, 4.98, 4.10)
    points = [
        "作業開始の入口にする（毎日必ず開く導線に）",
        "ツール起動と直結し、作業を中断させない",
        "未完了ロックや必須確認で確実に実施させる",
        "新人教育や品質向上に役立つことを現場が実感できる",
        "現場の声を取り入れ、使いやすく継続的に改善する",
    ]
    for i, txt in enumerate(points):
        yy = 5.48+i*.40
        icon_circle_check(slide, 9.06, yy+.08, .075, GREEN)
        add_text(slide, txt, 9.26, yy+.02, 3.25, .13, size=7.4, color=BLACK)
        if i < 4:
            add_line(slide, 8.96, yy+.30, 12.75, yy+.30, LIGHT_GREY, .7)

    # footer
    add_rect(slide, 0, 7.08, SLIDE_W, .42, fill=NAVY, line=NAVY)
    add_text(slide, "「見るだけのシステム」ではなく、「実際に作業を導き、完了まで支援するシステム」へ", .25, 7.15, 8.4, .18, size=12.7, color=YELLOW, bold=True)
    add_text(slide, "標準化・効率化・品質向上を同時に実現し、プロジェクトの成功に貢献します。", .58, 7.36, 7.45, .12, size=10.2, color=WHITE, bold=True)
    footer_icons = [("品質向上", "▱"), ("効率化", "↗"), ("属人化解消", "♙")]
    for i, (lab, sym) in enumerate(footer_icons):
        cx = 9.82 + i*1.08
        circ = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Inches(cx-.37), Inches(6.96), Inches(.74), Inches(.74))
        circ.fill.solid(); circ.fill.fore_color.rgb = WHITE; circ.line.color.rgb = WHITE
        add_text(slide, sym, cx-.14, 7.08, .28, .16, size=16, color=MID_BLUE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, lab, cx-.30, 7.46, .60, .10, size=7, color=DARK_NAVY, bold=True, align=PP_ALIGN.CENTER)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
