"""
データセンター警備提案書 - au Starlink通信冗長化ソリューション
PowerPoint生成スクリプト
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import pptx.oxml.ns as nsmap
from lxml import etree
import copy

# ========== カラー定義（コーポレート：赤×白） ==========
C_BG_DARK   = RGBColor(0xFF, 0xFF, 0xFF)   # 背景（白）
C_BG_MID    = RGBColor(0xFF, 0xF5, 0xF5)   # カード背景（薄赤白）
C_PRIMARY   = RGBColor(0xCC, 0x00, 0x00)   # メイン赤
C_ACCENT    = RGBColor(0x99, 0x00, 0x00)   # 濃赤（アクセント）
C_GREEN     = RGBColor(0xCC, 0x00, 0x00)   # ※緑→赤に統一
C_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT     = RGBColor(0x33, 0x33, 0x33)   # 本文テキスト（ほぼ黒）
C_MUTED     = RGBColor(0x88, 0x88, 0x88)   # サブテキスト（グレー）
C_BORDER    = RGBColor(0xE0, 0xC0, 0xC0)   # 枠線（薄赤）
C_RED_SOFT  = RGBColor(0xCC, 0x00, 0x00)   # 問題点（赤）
C_CARD_BG   = RGBColor(0xFF, 0xFA, 0xFA)   # カード背景（ほぼ白）

W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

blank_layout = prs.slide_layouts[6]  # 完全ブランク

# ========== ユーティリティ ==========

def add_rect(slide, l, t, w, h, fill=None, line_color=None, line_w=Pt(1)):
    shape = slide.shapes.add_shape(1, l, t, w, h)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.line.width = 0
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = line_w
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, l, t, w, h,
             size=Pt(14), bold=False, color=C_LIGHT,
             align=PP_ALIGN.LEFT, wrap=True, italic=False):
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = wrap
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb

def add_textbox_multiline(slide, lines, l, t, w, h,
                           size=Pt(13), bold=False, color=C_LIGHT,
                           align=PP_ALIGN.LEFT, line_spacing=Pt(6)):
    """lines: list of (text, bold, color) tuples or plain strings"""
    txb = slide.shapes.add_textbox(l, t, w, h)
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        if isinstance(item, str):
            txt, b, c = item, bold, color
        else:
            txt, b, c = item
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = align
        p.space_before = line_spacing
        run = p.add_run()
        run.text = txt
        run.font.size = size
        run.font.bold = b
        run.font.color.rgb = c
    return txb

def fill_slide_bg(slide, color=C_BG_DARK):
    add_rect(slide, 0, 0, W, H, fill=color)
    # 全スライド共通: 上部に赤帯
    add_rect(slide, 0, 0, W, Inches(1.1), fill=C_PRIMARY)

def section_bar(slide, l=Inches(0.6), t=Inches(1.35), h=Inches(0.38)):
    """タイトル左のアクセントバー（赤背景上では不要なので非表示相当）"""
    pass  # ヘッダー帯で代替

def page_num(slide, cur, total):
    add_text(slide, f"{cur} / {total}",
             W - Inches(1.5), H - Inches(0.45), Inches(1.3), Inches(0.3),
             size=Pt(11), color=C_MUTED, align=PP_ALIGN.RIGHT)

def slide_title_bar(slide, title):
    # 赤ヘッダー帯の上にタイトルを白で表示
    add_text(slide, title,
             Inches(0.6), Inches(0.2), Inches(12.1), Inches(0.7),
             size=Pt(22), bold=True, color=C_WHITE)
    # 帯下の区切り線
    add_rect(slide, Inches(0.5), Inches(1.15), Inches(12.3), Pt(1.5),
             fill=C_BORDER)

def card_box(slide, l, t, w, h, title, body,
             title_color=C_PRIMARY, body_size=Pt(12.5),
             icon=""):
    add_rect(slide, l, t, w, h, fill=C_CARD_BG, line_color=C_BORDER, line_w=Pt(1.2))
    # カード上部に細い赤ライン
    add_rect(slide, l, t, w, Inches(0.045), fill=C_PRIMARY)
    ty = t + Inches(0.18)
    if icon:
        add_text(slide, icon, l + Inches(0.18), ty, Inches(0.5), Inches(0.4),
                 size=Pt(20), color=C_LIGHT)
        tx = l + Inches(0.72)
        tw = w - Inches(0.9)
    else:
        tx = l + Inches(0.22)
        tw = w - Inches(0.44)
    add_text(slide, title, tx, ty, tw, Inches(0.35),
             size=Pt(14), bold=True, color=title_color)
    add_text(slide, body, tx, ty + Inches(0.38), tw, h - Inches(0.7),
             size=body_size, color=C_LIGHT, wrap=True)

def highlight_box(slide, l, t, w, h, title, body):
    add_rect(slide, l, t, w, h,
             fill=RGBColor(0xFF, 0xF0, 0xF0), line_color=C_PRIMARY, line_w=Pt(1.5))
    add_rect(slide, l, t, Inches(0.06), h, fill=C_PRIMARY)
    add_text(slide, title, l + Inches(0.18), t + Inches(0.1), w - Inches(0.28), Inches(0.35),
             size=Pt(14), bold=True, color=C_PRIMARY)
    add_text(slide, body, l + Inches(0.18), t + Inches(0.46), w - Inches(0.28), h - Inches(0.55),
             size=Pt(12.5), color=C_LIGHT, wrap=True)

def metric_box(slide, l, t, w, h, value, label):
    add_rect(slide, l, t, w, h, fill=C_CARD_BG, line_color=C_BORDER, line_w=Pt(1.2))
    add_rect(slide, l, t, w, Inches(0.045), fill=C_PRIMARY)
    add_text(slide, value, l, t + Inches(0.15), w, Inches(0.6),
             size=Pt(32), bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)
    add_text(slide, label, l, t + Inches(0.75), w, Inches(0.55),
             size=Pt(11.5), color=C_MUTED, align=PP_ALIGN.CENTER, wrap=True)

# ========================================================
# スライド 1: タイトル
# ========================================================
sl = prs.slides.add_slide(blank_layout)
# タイトルスライドは上半分赤・下半分白
add_rect(sl, 0, 0, W, H, fill=C_WHITE)
add_rect(sl, 0, 0, W, Inches(4.2), fill=C_PRIMARY)
add_rect(sl, 0, H - Inches(0.08), W, Inches(0.08), fill=C_PRIMARY)

# バッジ
add_rect(sl, Inches(0.8), Inches(0.55), Inches(1.9), Inches(0.32),
         fill=RGBColor(0xAA, 0x00, 0x00), line_color=None)
add_text(sl, "PROPOSAL 2026", Inches(0.82), Inches(0.57), Inches(1.86), Inches(0.28),
         size=Pt(11), bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# メインタイトル（赤帯上→白テキスト）
add_text(sl, "災害に強いデータセンター警備",
         Inches(0.8), Inches(1.05), Inches(11.5), Inches(0.9),
         size=Pt(40), bold=True, color=C_WHITE)
add_text(sl, "au Starlink 通信冗長化ソリューション",
         Inches(0.8), Inches(1.95), Inches(11.5), Inches(0.7),
         size=Pt(28), bold=True, color=RGBColor(0xFF, 0xCC, 0xCC))

# サブタイトル（赤帯上→白）
add_text(sl,
         "衛星通信 × 地上回線のハイブリッド構成で\n「どんな災害でも途切れない警備体制」を実現します",
         Inches(0.8), Inches(2.75), Inches(10.5), Inches(1.1),
         size=Pt(16), color=RGBColor(0xFF, 0xEE, 0xEE), wrap=True)

# 白エリアのメタ情報
add_rect(sl, Inches(0.8), Inches(4.35), Inches(11.5), Pt(1.5), fill=C_BORDER)
for i, (lbl, val) in enumerate([
    ("提案日", "2026年6月25日"),
    ("提案先", "貴社データセンター担当者様"),
    ("作成者", "警備営業部"),
]):
    x = Inches(0.8) + i * Inches(3.9)
    add_text(sl, lbl, x, Inches(4.5), Inches(1.2), Inches(0.3),
             size=Pt(11), color=C_MUTED)
    add_text(sl, val, x, Inches(4.8), Inches(3.7), Inches(0.4),
             size=Pt(14), bold=True, color=C_LIGHT)

page_num(sl, 1, 13)

# ========================================================
# スライド 2: 目次
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "本日のご提案内容")

items = [
    ("⚠", "1. 現状課題の整理",       "データセンター警備における通信断リスクと既存体制の脆弱性を確認します"),
    ("🛰", "2. au Starlinkとは",      "低軌道衛星通信の仕組みとauサービスとしての信頼性をご説明します"),
    ("📊", "3. 他手段との比較",        "地上回線・LTE・従来衛星との比較表でStarlinkの優位性を示します"),
    ("🔒", "4. 警備への活用提案",      "通信冗長化の具体的な構成と警備オペレーションへの組み込み方法をご提示します"),
    ("📱", "5. スマホ警備ホットライン","au Starlink Wi-Fi経由でスマホを警備専用ホットラインとして活用します"),
    ("📋", "6. 導入ステップ・費用感",  "Starlink端末はオプション追加で対応可能。スモールスタートから段階的展開まで、ロードマップと費用をご案内します"),
    ("✅", "7. 導入効果",             "定量的な効果指標と期待されるビジネス価値をご紹介します"),
]

cols = 2
rows_per_col = 4
cw, ch = Inches(5.9), Inches(0.98)
gx, gy = Inches(0.25), Inches(0.22)
ox, oy = Inches(0.5), Inches(2.0)

for idx, (icon, title, desc) in enumerate(items):
    col = idx % 2
    row = idx // 2
    x = ox + col * (cw + gx)
    y = oy + row * (ch + gy)
    card_box(sl, x, y, cw, ch, title, desc, icon=icon, body_size=Pt(11.5))

page_num(sl, 2, 13)

# ========================================================
# スライド 3: 現状課題
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "データセンター警備が抱える「通信断」リスク")

problems = [
    ("🌀", "自然災害による通信断",
     "地震・台風・大雨による地上回線の切断。光ファイバーの物理損傷は復旧に数日〜数週間を要する。"),
    ("⚡", "停電・設備障害",
     "長時間停電時のUPS枯渇。通信機器の電源断により監視カメラ・センサーが失報するリスクがある。"),
    ("🔌", "単一キャリア依存",
     "1系統の地上回線のみでは、局舎障害・ケーブル切断で警備センターとの通信が完全孤立する。"),
]
cw, ch = Inches(3.9), Inches(1.55)
for i, (icon, ttl, body) in enumerate(problems):
    x = Inches(0.5) + i * (cw + Inches(0.22))
    card_box(sl, x, Inches(2.0), cw, ch, ttl, body,
             title_color=C_RED_SOFT, icon=icon, body_size=Pt(11.5))

highlight_box(sl,
    Inches(0.5), Inches(3.72), Inches(12.3), Inches(1.42),
    "⚡ 実際に起きた事例",
    "2024年能登半島地震では地上回線が広範囲で断絶し、最長6日間の通信不通が発生。\n"
    "警備の安否確認・指示系統が機能せず、不正侵入・盗難・火災の早期検知が困難になりました。")

for i, (val, lbl) in enumerate([
    ("72時間",  "大規模災害時の地上回線\n平均復旧待機時間"),
    ("約3倍",   "通信断時の不正侵入\n検知遅延リスク"),
    ("99.9%",  "顧客が求める警備システム\n稼働率（SLA水準）"),
]):
    metric_box(sl, Inches(0.5) + i * Inches(4.15), Inches(5.28), Inches(3.9), Inches(1.55), val, lbl)

page_num(sl, 3, 13)

# ========================================================
# スライド 4: au Starlinkとは
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "au Starlink とは ― 低軌道衛星通信の新時代")

# 左カラム
card_box(sl, Inches(0.5), Inches(2.0), Inches(5.8), Inches(1.6),
         "SpaceX Starlink × KDDI（au）",
         "SpaceXが展開する低軌道（LEO）衛星（高度約550km）をKDDI（au）が日本向け法人サービスとして提供。地上インフラに依存しない独立した通信経路を実現します。",
         icon="🛰", body_size=Pt(12))

card_box(sl, Inches(0.5), Inches(3.75), Inches(5.8), Inches(1.85),
         "法人向けサービスの特長",
         "・KDDIの法人サポート体制（SLA保証）\n・既存auモバイル回線とのセット割\n・導入・保守の一元窓口\n・暗号化通信によるセキュリティ対応",
         icon="📡", body_size=Pt(12))

# 右カラム: 構成図テキスト
add_rect(sl, Inches(6.6), Inches(2.0), Inches(6.2), Inches(5.1),
         fill=C_CARD_BG, line_color=C_BORDER, line_w=Pt(1.2))
add_text(sl, "通信経路イメージ",
         Inches(6.8), Inches(2.1), Inches(5.8), Inches(0.3),
         size=Pt(11), color=C_MUTED, align=PP_ALIGN.CENTER)

nodes = [
    ("🏢 データセンター（警備LAN）",        C_PRIMARY, Inches(2.3), Inches(0.4)),
    ("📡 Starlinkアンテナ（au提供）",        C_PRIMARY, Inches(2.3), Inches(0.4)),
    ("🛰  Starlink衛星群（高度550km LEO）",  C_PRIMARY, Inches(2.5), Inches(0.4)),
    ("🖥  auゲートウェイ（KDDI地上局）",     C_GREEN,   Inches(2.5), Inches(0.4)),
    ("🔒 警備センター（24H監視）",           C_GREEN,   Inches(2.3), Inches(0.4)),
]
ny = Inches(2.5)
for (txt, col, nw, nh) in nodes:
    nx = Inches(6.6) + (Inches(6.2) - nw) / 2
    add_rect(sl, nx, ny, nw, nh,
             fill=RGBColor(0xFF, 0xF8, 0xF8), line_color=col, line_w=Pt(1.2))
    add_text(sl, txt, nx, ny, nw, nh,
             size=Pt(12), color=col, align=PP_ALIGN.CENTER, bold=True)
    ny += nh + Inches(0.12)
    if ny < Inches(6.8):
        add_text(sl, "↕ 自動切替対応", Inches(6.6), ny - Inches(0.15),
                 Inches(6.2), Inches(0.2),
                 size=Pt(10), color=C_MUTED, align=PP_ALIGN.CENTER)

# スペック
for i, (val, lbl) in enumerate([
    ("20〜100 Mbps", "下り通信速度"),
    ("20〜40 ms",    "低遅延（LEO）"),
    ("99%+",         "au法人SLA可用性"),
]):
    metric_box(sl, Inches(0.5) + i * Inches(2.05), Inches(5.7), Inches(1.9), Inches(1.1), val, lbl)

page_num(sl, 4, 13)

# ========================================================
# スライド 5: 比較表
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "通信手段の比較 ― なぜ au Starlink が最適解か")

headers = ["比較項目", "地上光回線\n（現状）", "携帯LTE/5G\n（モバイル）", "従来衛星通信\n（静止軌道）", "au Starlink\n（提案）"]
col_w = [Inches(2.8), Inches(2.2), Inches(2.2), Inches(2.2), Inches(2.2)]
rows_data = [
    ("地震・台風時の耐障害性", "× ケーブル切断", "△ 基地局被災リスク", "○ 地上非依存", "◎ 最強の独立経路"),
    ("通信速度",              "◎ 1Gbps+",      "○ 50〜200Mbps",      "× 1〜5Mbps",   "◎ 20〜100Mbps"),
    ("遅延",                  "◎ 1〜5ms",      "○ 10〜40ms",         "× 500〜700ms", "◎ 20〜40ms"),
    ("HD映像監視",            "◎ 対応",         "○ 対応",              "× 遅延で不可",  "◎ リアルタイム"),
    ("法人サポート",          "○ あり",         "△ キャリア依存",      "△ 専門業者",    "◎ KDDI一元化"),
    ("冗長バックアップ適性",  "× 同経路リスク", "△ 集中障害リスク",    "○ 適合",        "◎ 最適（異経路）"),
]

row_h = Inches(0.52)
start_y = Inches(2.0)

# ヘッダー行
x = Inches(0.4)
for j, (hdr, cw) in enumerate(zip(headers, col_w)):
    bg = C_PRIMARY if j == 4 else RGBColor(0xCC, 0x00, 0x00)
    add_rect(sl, x, start_y, cw, row_h, fill=bg)
    add_text(sl, hdr, x + Inches(0.05), start_y + Inches(0.06),
             cw - Inches(0.1), row_h - Inches(0.1),
             size=Pt(11.5), bold=True,
             color=C_BG_DARK if j == 4 else C_PRIMARY,
             align=PP_ALIGN.CENTER, wrap=True)
    x += cw

# データ行
mark_colors = {
    "◎": C_GREEN, "○": C_PRIMARY, "△": RGBColor(0xFF, 0xAA, 0x44), "×": C_RED_SOFT
}
for ri, row in enumerate(rows_data):
    y = start_y + (ri + 1) * row_h
    row_bg = RGBColor(0xFF, 0xF5, 0xF5) if ri % 2 else C_CARD_BG
    x = Inches(0.4)
    for j, (cell, cw) in enumerate(zip(row, col_w)):
        highlight_col = j == 4
        bg = RGBColor(0xFF, 0xEE, 0xEE) if highlight_col else row_bg
        add_rect(sl, x, y, cw, row_h, fill=bg)
        fc = C_WHITE if j == 0 else C_LIGHT
        for mark, mc in mark_colors.items():
            if cell.startswith(mark):
                fc = mc
                break
        add_text(sl, cell, x + Inches(0.05), y + Inches(0.08),
                 cw - Inches(0.1), row_h - Inches(0.14),
                 size=Pt(11), color=fc, align=PP_ALIGN.CENTER, wrap=True)
        x += cw

highlight_box(sl,
    Inches(0.4), Inches(6.4), Inches(12.5), Inches(0.78),
    "結論",
    "au Starlink は「地上インフラに依存しない独立経路」と「HD映像転送に耐える速度・低遅延」を両立する、警備バックアップ通信として現時点で最も実用的な選択肢です。")

page_num(sl, 5, 13)

# ========================================================
# スライド 6: 活用シナリオ
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "具体的な活用シナリオ ― 災害発生時の通信フロー")

def flow_col(slide, x, y, title, title_color, steps):
    add_text(slide, title, x, y, Inches(5.8), Inches(0.35),
             size=Pt(14), bold=True, color=title_color)
    cy = y + Inches(0.42)
    for (icon, lbl, ttl, body) in steps:
        add_rect(slide, x, cy, Inches(0.5), Inches(0.5),
                 fill=title_color, line_color=None)
        add_text(slide, icon, x, cy, Inches(0.5), Inches(0.5),
                 size=Pt(18), align=PP_ALIGN.CENTER, color=C_WHITE)
        add_text(slide, lbl, x + Inches(0.58), cy, Inches(5.1), Inches(0.2),
                 size=Pt(10), color=C_MUTED)
        add_text(slide, ttl, x + Inches(0.58), cy + Inches(0.18), Inches(5.1), Inches(0.3),
                 size=Pt(13), bold=True, color=C_WHITE)
        add_text(slide, body, x + Inches(0.58), cy + Inches(0.46), Inches(5.1), Inches(0.45),
                 size=Pt(11.5), color=C_LIGHT, wrap=True)
        cy += Inches(1.0)

before_steps = [
    ("🌀", "発災",   "地震・台風発生",        "地上ケーブル切断・通信局舎損傷"),
    ("📵", "数分後", "通信完全断",            "カメラ映像・センサー情報が届かない"),
    ("⚠", "数日後", "施設の「監視空白」継続", "不正侵入・火災の検知不能。顧客報告が困難に"),
]
after_steps = [
    ("🌀", "発災",   "地震・台風発生",              "地上ケーブル切断を検知"),
    ("⚡", "数秒以内", "Starlinkへ自動フェイルオーバー", "ルーターが自動切替。警備通信が継続"),
    ("🛡", "継続",  "監視・通報機能が維持",         "映像・センサー・緊急通報がすべて継続稼働"),
]

flow_col(sl, Inches(0.5), Inches(2.0), "【現状】地上回線のみ ― 災害時", C_RED_SOFT, before_steps)
flow_col(sl, Inches(7.0), Inches(2.0), "【提案後】Starlink冗長化 ― 災害時", C_GREEN, after_steps)

# 中央の矢印
add_text(sl, "VS", Inches(6.0), Inches(3.4), Inches(0.9), Inches(0.9),
         size=Pt(22), bold=True, color=C_MUTED, align=PP_ALIGN.CENTER)

highlight_box(sl,
    Inches(0.5), Inches(5.95), Inches(12.3), Inches(0.85),
    "🔑 対応可能な警備システム",
    "映像監視（IP-CAM / HD・4K）　入退室管理システム　侵入センサー・火災報知連携　警備員携帯端末（PTT/VoIP）　緊急通報システム")

page_num(sl, 6, 13)

# ========================================================
# スライド 7: システム構成
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "推奨システム構成 ― ハイブリッド冗長化モデル")

# 左側: 平常時
add_text(sl, "【平常時：地上回線メイン】",
         Inches(0.5), Inches(2.05), Inches(5.5), Inches(0.3),
         size=Pt(12), color=C_MUTED)
normal_nodes = [
    ("🏢 データセンター 警備LAN",          C_PRIMARY),
    ("🔀 冗長化ルーター（自動切替機能）",   C_PRIMARY),
    ("🔌 地上光回線（既存インフラ）",       C_ACCENT),
    ("🖥  警備センター 監視システム",       C_GREEN),
]
ny = Inches(2.4)
for txt, col in normal_nodes:
    add_rect(sl, Inches(0.6), ny, Inches(5.2), Inches(0.5),
             fill=RGBColor(0xFF, 0xF8, 0xF8), line_color=col, line_w=Pt(1.2))
    add_text(sl, txt, Inches(0.65), ny + Inches(0.08),
             Inches(5.1), Inches(0.35), size=Pt(12), bold=True, color=col,
             align=PP_ALIGN.CENTER)
    ny += Inches(0.62)

# 中央
add_text(sl, "⟷\n障害発生時\n自動切替", Inches(6.0), Inches(3.2), Inches(1.1), Inches(1.0),
         size=Pt(12), color=C_MUTED, align=PP_ALIGN.CENTER, wrap=True)

# 右側: 障害時
add_text(sl, "【障害時：Starlinkへ自動切替】",
         Inches(7.2), Inches(2.05), Inches(5.5), Inches(0.3),
         size=Pt(12), color=C_GREEN)
failover_nodes = [
    ("🏢 データセンター 警備LAN",              C_PRIMARY),
    ("🔀 冗長化ルーター（フェイルオーバー発動）", C_PRIMARY),
    ("📡 au Starlink（衛星経由で接続維持）",    C_GREEN),
    ("🖥  警備センター 監視システム",           C_GREEN),
]
ny = Inches(2.4)
for txt, col in failover_nodes:
    add_rect(sl, Inches(7.3), ny, Inches(5.2), Inches(0.5),
             fill=RGBColor(0xFF, 0xF8, 0xF8), line_color=col, line_w=Pt(1.2))
    add_text(sl, txt, Inches(7.35), ny + Inches(0.08),
             Inches(5.1), Inches(0.35), size=Pt(12), bold=True, color=col,
             align=PP_ALIGN.CENTER)
    ny += Inches(0.62)

# 下部カード3つ
comps = [
    ("🔀 冗長化ルーター",  "WAN2ポート対応ルーター（Peplink/FortiGate等）により自動管理。切替時間は5〜30秒。"),
    ("📡 Starlink端末設置", "屋上または外壁に設置。直径約60cmのFlat HPアンテナで安定稼働。"),
    ("🔋 UPS連携",          "Starlink端末・ルーターをUPSに接続し、停電時も数時間の通信継続を確保。"),
]
cw = Inches(3.95)
for i, (ttl, body) in enumerate(comps):
    card_box(sl, Inches(0.5) + i * (cw + Inches(0.22)),
             Inches(5.75), cw, Inches(1.45), ttl, body, body_size=Pt(12))

page_num(sl, 7, 13)

# ========================================================
# スライド 8: スマホ警備ホットライン
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "スマホ警備ホットライン ― au Starlink Wi-Fi接続で常時つながる")

# 左: 構成図
add_rect(sl, Inches(0.5), Inches(2.0), Inches(5.8), Inches(5.2),
         fill=C_CARD_BG, line_color=C_BORDER, line_w=Pt(1.2))
add_text(sl, "ホットライン構成図",
         Inches(0.5), Inches(2.1), Inches(5.8), Inches(0.3),
         size=Pt(11), color=C_MUTED, align=PP_ALIGN.CENTER)

hot_nodes = [
    ("📱 警備員スマホ（au回線 or Wi-Fi）",  C_PRIMARY),
    ("📡 Starlinkアンテナ（データセンター）", C_PRIMARY),
    ("🛰  Starlink衛星（LEO 550km）",        C_PRIMARY),
    ("🖥  警備センター（ホットライン受付）",  C_GREEN),
    ("🚔 緊急対応チーム（現地急行）",        C_GREEN),
]
ny = Inches(2.55)
for txt, col in hot_nodes:
    add_rect(sl, Inches(0.9), ny, Inches(4.9), Inches(0.45),
             fill=RGBColor(0xFF, 0xF8, 0xF8), line_color=col, line_w=Pt(1.2))
    add_text(sl, txt, Inches(0.95), ny + Inches(0.06),
             Inches(4.8), Inches(0.33), size=Pt(12), bold=True, color=col,
             align=PP_ALIGN.CENTER)
    ny += Inches(0.62)

# 右: 機能一覧
add_text(sl, "ホットライン機能一覧",
         Inches(6.6), Inches(2.0), Inches(6.2), Inches(0.35),
         size=Pt(14), bold=True, color=C_WHITE)

features = [
    ("☎", "ワンタッチ緊急呼び出し", "ボタン一発で警備センターへ直接接続。操作不要。"),
    ("🔊", "VoIP音声通話",           "Starlink Wi-Fi経由で通話。au回線が不通でも継続。"),
    ("📢", "PTTグループ通話",         "現場警備員全員へ同時一斉通話で指示を伝達。"),
    ("📍", "GPS位置情報共有",         "リアルタイム位置をセンターへ送信。巡回管理が可能。"),
    ("📸", "映像・写真の即時送信",    "異常発見時に現場写真・動画をリアルタイム共有。"),
    ("💬", "テキスト状況報告",        "騒音環境でも文字で報告。ログとして記録・管理も可能。"),
]
fy = Inches(2.45)
for icon, title, body in features:
    add_text(sl, f"{icon} {title}",
             Inches(6.6), fy, Inches(6.2), Inches(0.28),
             size=Pt(13), bold=True, color=C_PRIMARY)
    add_text(sl, body,
             Inches(6.6), fy + Inches(0.28), Inches(6.2), Inches(0.3),
             size=Pt(11.5), color=C_LIGHT, wrap=True)
    add_rect(sl, Inches(6.6), fy + Inches(0.62), Inches(6.2), Pt(0.8), fill=C_BORDER)
    fy += Inches(0.72)

# 下部3カード
comps = [
    ("📶 二重通信経路",     "au LTE＋Starlink Wi-Fiの両方に対応。地上回線切断時もStarlink経由で通話・通報が継続。"),
    ("🔋 堅牢端末運用",    "タフネススマホ（防塵・防水・耐衝撃）＋予備バッテリーで24H稼働を実現。"),
    ("🛡 専用APNで保護",  "au法人専用APN（閉域網）適用により通話・映像データが外部インターネットに出ない設計。"),
]
cw = Inches(3.95)
for i, (ttl, body) in enumerate(comps):
    card_box(sl, Inches(0.5) + i * (cw + Inches(0.22)),
             Inches(6.05), cw, Inches(1.2), ttl, body, body_size=Pt(11.5))

page_num(sl, 8, 13)

# ========================================================
# スライド 9: 導入ロードマップ
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "導入ロードマップ ― 最短2ヶ月で稼働開始")

steps = [
    ("STEP 01", "現地調査\n要件定義",  "電波環境測定\n設置場所確認\n既存回線確認"),
    ("STEP 02", "設計\n機器選定",      "ルーター選定\n回線構成設計\nUPS容量計算"),
    ("STEP 03", "機器調達\n工事準備",  "au申込手続き\nルーター・配線\n資材手配"),
    ("STEP 04", "設置工事\n開通確認",  "アンテナ設置\n回線開通\nフェイルオーバーテスト"),
    ("STEP 05", "運用開始\n保守契約",  "警備員説明\n24H保守開始"),
]
sw = Inches(2.3)
sx = Inches(0.4)
for i, (num, ttl, body) in enumerate(steps):
    add_rect(sl, sx, Inches(2.0), sw, Inches(2.2),
             fill=C_CARD_BG, line_color=C_PRIMARY, line_w=Pt(1.5))
    add_text(sl, num, sx, Inches(2.08), sw, Inches(0.25),
             size=Pt(10), bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)
    add_text(sl, ttl, sx, Inches(2.32), sw, Inches(0.65),
             size=Pt(14), bold=True, color=C_WHITE, align=PP_ALIGN.CENTER, wrap=True)
    add_rect(sl, sx + Inches(0.2), Inches(2.98), sw - Inches(0.4), Pt(1), fill=C_BORDER)
    add_text(sl, body, sx + Inches(0.1), Inches(3.05), sw - Inches(0.2), Inches(1.1),
             size=Pt(11.5), color=C_LIGHT, align=PP_ALIGN.CENTER, wrap=True)
    sx += sw + Inches(0.17)
    if i < 4:
        add_text(sl, "→", sx - Inches(0.17) + Inches(0.02), Inches(2.9),
                 Inches(0.15), Inches(0.3),
                 size=Pt(18), bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)

highlight_box(sl,
    Inches(0.4), Inches(4.35), Inches(12.5), Inches(0.75),
    "💡 Starlink端末はオプション導入に対応",
    "既存の警備契約にStarlink通信バックアップをオプションとして追加可能。まず1棟・3ヶ月のトライアルから始め、効果を確認してから全館展開へ移行できます。")

# スケジュール / チェックポイント
card_box(sl, Inches(0.4), Inches(5.3), Inches(5.9), Inches(2.0),
         "📅 標準スケジュール",
         "Week 1-2：現地調査・要件定義\nWeek 3-4：設計・機器発注\nWeek 5-6：工事準備・資材搬入\nWeek 7-8：設置工事・開通テスト\nWeek 9〜：本番運用開始",
         body_size=Pt(12))
card_box(sl, Inches(6.6), Inches(5.3), Inches(5.9), Inches(2.0),
         "✅ 導入前チェックポイント",
         "□ 屋上または外壁の設置スペース確保\n□ 南〜西方向の空の視界（衛星仰角）\n□ 電源（AC100V）の引き込み\n□ データセンターの設備変更申請手続き\n□ 既存ルーターの置き換えまたは追加",
         body_size=Pt(12))

page_num(sl, 9, 13)

# ========================================================
# スライド 10: 費用感
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "概算費用 ― 標準モデル（データセンター1棟）")

# 初期費用テーブル
add_text(sl, "【初期費用】", Inches(0.5), Inches(2.0), Inches(5.9), Inches(0.3),
         size=Pt(13), bold=True, color=C_PRIMARY)
init_items = [
    ("Starlink端末（Flat High Performance）",  "約 150,000円"),
    ("冗長化ルーター（Peplink Balance等）",     "約 80,000〜150,000円"),
    ("UPS（無停電電源装置）",                  "約 50,000〜100,000円"),
    ("設置工事費（アンテナ・配線）",           "約 100,000〜200,000円"),
    ("設定・開通テスト費",                     "約 50,000円"),
    ("初期費用合計（目安）",                   "約 43〜65万円"),
]
ty = Inches(2.35)
for i, (item, price) in enumerate(init_items):
    is_total = i == len(init_items) - 1
    bg = RGBColor(0xFF, 0xEE, 0xEE) if is_total else (C_CARD_BG if i % 2 == 0 else RGBColor(0xFF, 0xF5, 0xF5))
    add_rect(sl, Inches(0.5), ty, Inches(4.0), Inches(0.4), fill=bg)
    add_rect(sl, Inches(4.5), ty, Inches(1.9), Inches(0.4), fill=bg)
    tc = C_PRIMARY if is_total else C_LIGHT
    add_text(sl, item, Inches(0.55), ty + Inches(0.06), Inches(3.9), Inches(0.3),
             size=Pt(12), bold=is_total, color=tc)
    add_text(sl, price, Inches(4.52), ty + Inches(0.06), Inches(1.85), Inches(0.3),
             size=Pt(12), bold=is_total, color=C_PRIMARY if is_total else C_LIGHT,
             align=PP_ALIGN.RIGHT)
    ty += Inches(0.42)

# 月額費用テーブル
add_text(sl, "【月額費用】", Inches(7.1), Inches(2.0), Inches(5.9), Inches(0.3),
         size=Pt(13), bold=True, color=C_PRIMARY)
monthly_items = [
    ("au Starlink 法人回線料",      "約 15,000〜25,000円"),
    ("保守・監視サービス料",        "約 10,000〜20,000円"),
    ("（既存回線は継続）",          "変更なし"),
    ("月額追加費用（目安）",        "約 2.5〜4.5万円"),
]
ty = Inches(2.35)
for i, (item, price) in enumerate(monthly_items):
    is_total = i == len(monthly_items) - 1
    bg = RGBColor(0xFF, 0xEE, 0xEE) if is_total else (C_CARD_BG if i % 2 == 0 else RGBColor(0xFF, 0xF5, 0xF5))
    add_rect(sl, Inches(7.1), ty, Inches(3.8), Inches(0.4), fill=bg)
    add_rect(sl, Inches(10.9), ty, Inches(1.9), Inches(0.4), fill=bg)
    tc = C_PRIMARY if is_total else C_LIGHT
    add_text(sl, item, Inches(7.15), ty + Inches(0.06), Inches(3.7), Inches(0.3),
             size=Pt(12), bold=is_total, color=tc)
    add_text(sl, price, Inches(10.92), ty + Inches(0.06), Inches(1.85), Inches(0.3),
             size=Pt(12), bold=is_total, color=C_PRIMARY if is_total else C_LIGHT,
             align=PP_ALIGN.RIGHT)
    ty += Inches(0.42)

highlight_box(sl,
    Inches(7.1), Inches(4.2), Inches(5.7), Inches(1.1),
    "💰 費用対効果",
    "通信断による損害賠償・SLA違約金リスクと比較すると、月額3〜5万円の追加投資でリスクを大幅に低減できます。")

highlight_box(sl,
    Inches(0.5), Inches(5.55), Inches(6.3), Inches(1.1),
    "📦 Starlink端末はオプション導入",
    "既存の警備サービス契約にStarlink通信バックアップをオプションとして追加する形態でも提供可能です。初期投資を抑えたトライアルからご検討いただけます。")

add_text(sl, "※上記は概算です。現地調査・設計後に正式な見積をご提示します。au法人向け特別価格適用の場合あり。",
         Inches(0.5), Inches(7.1), Inches(12.3), Inches(0.25),
         size=Pt(10), color=C_MUTED, italic=True)

page_num(sl, 10, 13)

# ========================================================
# スライド 11: 導入効果
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "導入による期待効果")

for i, (val, lbl) in enumerate([
    ("99.99%",    "警備通信システム\n稼働率（目標）"),
    ("<30秒",     "フェイルオーバー\n切替時間"),
    ("72時間",    "UPS組み合わせ時\n継続稼働目安"),
    ("0件",       "通信断による\n監視空白（目標）"),
]):
    metric_box(sl, Inches(0.5) + i * Inches(3.12), Inches(2.0), Inches(2.9), Inches(1.35), val, lbl)

# 定性・顧客メリット
eff_left = [
    "災害時の顧客への「継続稼働報告」が可能になり信頼性が向上",
    "警備員・警備センター間の通信断がなくなり指揮系統が維持される",
    "BCP対応訴求ポイントとして新規顧客への差別化に活用可能",
    "SLA保証の強化による契約単価の見直し（アップセル）機会の創出",
    "通信障害起因の損害賠償リスクを大幅に低減",
    "KDDI au法人パートナーシップによる安定したサポート体制",
]
eff_right = [
    "データセンターのBCP要件を満たす通信冗長化を証明できる",
    "金融・医療・行政等の厳格なコンプライアンス要件への対応強化",
    "入居テナントへ「災害時も監視継続」をアピール可能",
    "警備会社との長期契約の根拠となる付加価値の提供",
    "ISO27001・SOC2等のセキュリティ認証審査でのアドバンテージ",
    "スマホホットライン導入により警備員の現場対応力が向上",
]

add_text(sl, "定性的な効果", Inches(0.5), Inches(3.55), Inches(5.9), Inches(0.3),
         size=Pt(14), bold=True, color=C_WHITE)
add_text(sl, "顧客企業のメリット", Inches(6.9), Inches(3.55), Inches(5.9), Inches(0.3),
         size=Pt(14), bold=True, color=C_WHITE)

for i, txt in enumerate(eff_left):
    y = Inches(3.95) + i * Inches(0.5)
    add_text(sl, "✓", Inches(0.5), y, Inches(0.3), Inches(0.35),
             size=Pt(13), bold=True, color=C_GREEN)
    add_text(sl, txt, Inches(0.85), y, Inches(5.5), Inches(0.42),
             size=Pt(12.5), color=C_LIGHT, wrap=True)

for i, txt in enumerate(eff_right):
    y = Inches(3.95) + i * Inches(0.5)
    add_text(sl, "✓", Inches(6.9), y, Inches(0.3), Inches(0.35),
             size=Pt(13), bold=True, color=C_GREEN)
    add_text(sl, txt, Inches(7.25), y, Inches(5.5), Inches(0.42),
             size=Pt(12.5), color=C_LIGHT, wrap=True)

page_num(sl, 11, 13)

# ========================================================
# スライド 12: FAQ
# ========================================================
sl = prs.slides.add_slide(blank_layout)
fill_slide_bg(sl)
slide_title_bar(sl, "よくあるご質問（FAQ）")

faqs = [
    ("大雨や台風でStarlinkのアンテナは使えますか？",
     "Starlink Flat High Performanceは防水・耐候設計（IP56）で、強風（最大278km/h相当）にも耐える設計です。豪雨時は一時的な速度低下の可能性がありますが、補助としてLTE回線との3重化も対応可能です。"),
    ("既存の警備システムはそのまま使えますか？",
     "はい。冗長化ルーターをネットワーク出口に追加するだけで、既存の警備カメラ・センサー・入退室管理システムはそのまま継続利用できます。既存システムへの変更は原則不要です。"),
    ("セキュリティ（情報漏洩リスク）は問題ありませんか？",
     "Starlinkの通信はAES-256による暗号化が施されています。au法人向けでは専用APN（閉域網）＋IPsec VPNの構成で、通信がインターネットに出ない設計でご提案します。"),
    ("建物の許可が必要ですか？工事の規模はどの程度ですか？",
     "屋上への設置には建物管理者の承認が必要です。アンテナは約60cmで架台または固定ボルト設置。電源配線を含め、工事期間は一般的に1〜2日程度です。"),
]
faq_h = Inches(1.22)
for i, (q, a) in enumerate(faqs):
    y = Inches(2.0) + i * (faq_h + Inches(0.12))
    add_rect(sl, Inches(0.5), y, Inches(12.3), faq_h,
             fill=C_CARD_BG, line_color=C_BORDER, line_w=Pt(1.2))
    add_rect(sl, Inches(0.5), y, Inches(0.5), faq_h,
             fill=RGBColor(0x00, 0x28, 0x44), line_color=None)
    add_text(sl, "Q", Inches(0.5), y + Inches(0.38), Inches(0.5), Inches(0.4),
             size=Pt(14), bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)
    add_text(sl, q, Inches(1.1), y + Inches(0.08), Inches(11.5), Inches(0.35),
             size=Pt(13), bold=True, color=C_WHITE)
    add_text(sl, "A：" + a, Inches(1.1), y + Inches(0.45), Inches(11.5), Inches(0.68),
             size=Pt(12), color=C_LIGHT, wrap=True)

page_num(sl, 12, 13)

# ========================================================
# スライド 13: CTA
# ========================================================
sl = prs.slides.add_slide(blank_layout)
# CTA: 赤背景全面
add_rect(sl, 0, 0, W, H, fill=C_PRIMARY)
add_rect(sl, 0, 0, W, Inches(1.1), fill=RGBColor(0xAA, 0x00, 0x00))

add_text(sl, "NEXT STEP",
         Inches(5.8), Inches(0.22), Inches(1.7), Inches(0.3),
         size=Pt(11), bold=True, color=RGBColor(0xFF, 0xCC, 0xCC), align=PP_ALIGN.CENTER)

add_text(sl, "まずは無料の現地調査から始めましょう",
         Inches(1.0), Inches(1.3), Inches(11.3), Inches(0.8),
         size=Pt(34), bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

add_text(sl,
         "電波環境の測定・設置スペースの確認・概算見積まで\n無料にてご対応いたします。お気軽にお声がけください。",
         Inches(1.5), Inches(2.25), Inches(10.3), Inches(0.9),
         size=Pt(16), color=RGBColor(0xFF, 0xEE, 0xEE), align=PP_ALIGN.CENTER, wrap=True)

# CTAボタン風（白地に赤テキスト）
add_rect(sl, Inches(4.4), Inches(3.25), Inches(4.5), Inches(0.65),
         fill=C_WHITE, line_color=None)
add_text(sl, "無料現地調査を申し込む",
         Inches(4.4), Inches(3.33), Inches(4.5), Inches(0.5),
         size=Pt(16), bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)

add_rect(sl, Inches(0.5), Inches(4.2), Inches(12.3), Pt(1),
         fill=RGBColor(0xAA, 0x00, 0x00))

for i, (lbl, val, sub) in enumerate([
    ("PROPOSAL BY",     "警備営業部",          "データセンター警備推進チーム"),
    ("SERVICE PARTNER", "KDDI株式会社（au）",   "法人向けStarlinkサービス提供"),
    ("DATE",            "2026年6月25日",        "本資料の有効期限：3ヶ月"),
]):
    x = Inches(0.7) + i * Inches(4.15)
    add_rect(sl, x, Inches(4.4), Inches(3.9), Inches(1.6),
             fill=RGBColor(0xAA, 0x00, 0x00), line_color=None)
    add_text(sl, lbl, x + Inches(0.15), Inches(4.5), Inches(3.6), Inches(0.28),
             size=Pt(10), color=RGBColor(0xFF, 0xCC, 0xCC))
    add_text(sl, val, x + Inches(0.15), Inches(4.78), Inches(3.6), Inches(0.35),
             size=Pt(15), bold=True, color=C_WHITE)
    add_text(sl, sub, x + Inches(0.15), Inches(5.13), Inches(3.6), Inches(0.3),
             size=Pt(11), color=RGBColor(0xFF, 0xDD, 0xDD))

add_text(sl, "13 / 13",
         W - Inches(1.5), H - Inches(0.45), Inches(1.3), Inches(0.3),
         size=Pt(11), color=RGBColor(0xFF, 0xCC, 0xCC), align=PP_ALIGN.RIGHT)

# ========== 保存 ==========
out = "/home/user/test/datacenter_security_proposal.pptx"
prs.save(out)
print(f"Saved: {out}")
