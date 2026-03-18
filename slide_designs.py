"""
SlideAI Design System v5 — マネーフォワード資料スタイル準拠
トンマナ:
  - 白背景・ネイビー/ブルー系カラー + オレンジアクセント
  - タイトル下線 + ■マーク付きサマリー文
  - ネイビーヘッダーバー付きコンテンツボックス
  - 横フロー矢印図・2カラム比較・ツリー構造・論点リスト
  - フッター（左:会社名、右:ページ番号）、右上「For Discussion」
"""

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import lxml.etree as etree
from pptx.util import Pt

# ── カラーパレット（マネーフォワード資料準拠）──
NAVY    = RGBColor(0x00, 0x20, 0x60)  # 濃紺
BLUE1   = RGBColor(0x1F, 0x38, 0x64)  # 濃青
BLUE2   = RGBColor(0x2E, 0x74, 0xB5)  # 中青
BLUE3   = RGBColor(0x9D, 0xC3, 0xE6)  # 薄青
LBLUE   = RGBColor(0xBD, 0xD7, 0xEE)  # 水色背景
VLIGHT  = RGBColor(0xDA, 0xE8, 0xF5)  # 超薄青
ORANGE  = RGBColor(0xE8, 0x77, 0x22)  # オレンジ（MFアクセント）
RED     = RGBColor(0xC0, 0x00, 0x00)  # 赤（強調枠）
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
BGRAY   = RGBColor(0xF2, 0xF2, 0xF2)  # 背景薄グレー
LGRAY   = RGBColor(0xD9, 0xD9, 0xD9)
GRAY1   = RGBColor(0x26, 0x26, 0x26)
GRAY2   = RGBColor(0x40, 0x40, 0x40)
GRAY3   = RGBColor(0x80, 0x80, 0x80)
BLACK   = RGBColor(0x00, 0x00, 0x00)

def I(n): return int(Inches(n))
def P(n): return int(Pt(n))

def _r(slide, x, y, w, h, fill=WHITE, lc=None, lw=1.0, radius=0):
    sh = slide.shapes.add_shape(1, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if lc:
        sh.line.color.rgb = lc; sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    return sh

def _t(slide, x, y, w, h, text, sz=11, bold=False, col=None,
       align=PP_ALIGN.LEFT, italic=False, wrap=True):
    if not text: return None
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame; tf.word_wrap = wrap; tf.auto_size = None
    bp  = tf._txBody.find(qn('a:bodyPr'))
    if bp is not None:
        bp.set('lIns', '45720'); bp.set('rIns', '45720')
        bp.set('tIns', '22860'); bp.set('bIns', '22860')
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(sz); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = col or BLACK
    return txb

def _parse(b):
    for sep in ('：', ':'):
        sp = b.find(sep)
        if sp > 0: return b[:sp].strip(), b[sp+1:].strip()
    return '', b.strip()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 共通レイアウト部品
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _page_frame(slide, prs, context_label='', page_num=0, label='For Discussion'):
    """全スライド共通フレーム: 上部コンテキストバー、フッター"""
    W = int(prs.slide_width); H = int(prs.slide_height)
    # 上部細線
    _r(slide, 0, I(0.28), W, P(1), fill=NAVY)
    # コンテキストラベル（左上）
    if context_label:
        _t(slide, I(0.3), I(0.05), I(7), I(0.22),
           context_label, sz=8, col=GRAY2)
    # For Discussion（右上）
    _t(slide, W - I(2.5), I(0.05), I(2.3), I(0.22),
       label, sz=8, italic=True, col=GRAY2, align=PP_ALIGN.RIGHT)
    # フッター下線
    _r(slide, 0, H - I(0.35), W, P(0.5), fill=NAVY)
    # フッター左: 会社名
    _t(slide, I(0.3), H - I(0.3), I(4), I(0.22),
       'SlideAI Copy Right Reserved', sz=7, col=GRAY3)
    # フッター右: ページ番号
    if page_num:
        _t(slide, W - I(0.6), H - I(0.3), I(0.5), I(0.22),
           str(page_num), sz=8, col=GRAY2, align=PP_ALIGN.RIGHT)

def _slide_title(slide, prs, title, summary=''):
    """スライドタイトル: テキスト + 下線 + ■サマリー文"""
    W = int(prs.slide_width)
    # タイトルテキスト
    _t(slide, I(0.3), I(0.35), W - I(0.6), I(0.55),
       title, sz=20, bold=True, col=NAVY)
    # タイトル下線（全幅）
    _r(slide, I(0.3), I(0.95), W - I(0.6), P(1.5), fill=NAVY)
    # ■サマリー文
    if summary:
        _t(slide, I(0.3), I(1.0), W - I(0.6), I(0.45),
           '■ ' + summary, sz=10, col=GRAY1)

def _navy_header_box(slide, x, y, w, h, text, sz=11):
    """ネイビー背景の横帯ヘッダーボックス"""
    _r(slide, x, y, w, h, fill=NAVY)
    _t(slide, x, y + P(3), w, h - P(6),
       text, sz=sz, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

def _content_box(slide, x, y, w, h, header_text, items, header_sz=11, item_sz=9):
    """ネイビーヘッダー付きコンテンツボックス"""
    header_h = I(0.32)
    _navy_header_box(slide, x, y, w, header_h, header_text, sz=header_sz)
    _r(slide, x, y + header_h, w, h - header_h, fill=WHITE, lc=NAVY, lw=0.75)
    iy = y + header_h + I(0.1)
    avail_h = h - header_h - I(0.15)
    per = avail_h // max(len(items), 1)
    for item in items:
        lbl, det = _parse(item)
        if lbl:
            _t(slide, x + I(0.12), iy, w - I(0.2), min(per * 0.45, I(0.32)),
               '• ' + lbl, sz=item_sz, bold=True, col=BLUE1)
            if det:
                _t(slide, x + I(0.18), iy + I(0.3), w - I(0.26), per - I(0.3),
                   det, sz=item_sz - 1, col=GRAY1)
        else:
            _t(slide, x + I(0.12), iy, w - I(0.2), per,
               '• ' + det, sz=item_sz, col=GRAY1)
        iy += per


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 1. タイトルスライド（先頭固定）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_title(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)

    # 相手先社名（上部）
    client = body[0] if body else ''
    lbl, det = _parse(client)
    client_name = (lbl or client)[:20]
    _t(slide, I(0.7), I(1.2), W - I(1.4), I(0.6),
       client_name, sz=20, col=BLUE2)

    # タイトルブロック（ネイビー背景）
    _r(slide, I(0.5), I(2.2), W - I(1.0), I(1.1), fill=NAVY)
    # タグライン
    title_parts = headline.split('】')
    if '【' in headline and '】' in headline:
        tag = headline[headline.find('【'):headline.find('】')+1]
        main = headline[headline.find('】')+1:].strip()
        _t(slide, I(0.7), I(2.28), W - I(1.4), I(0.32),
           tag, sz=11, col=WHITE)
        _t(slide, I(0.7), I(2.62), W - I(1.4), I(0.55),
           main, sz=18, bold=True, col=WHITE)
    else:
        _t(slide, I(0.7), I(2.3), W - I(1.4), I(0.75),
           headline, sz=18, bold=True, col=WHITE)

    # 日付・提出元
    if len(body) > 1:
        _t(slide, I(0.7), I(3.55), I(4), I(0.3),
           body[1], sz=10, col=GRAY2)
    if len(body) > 2:
        _t(slide, I(0.7), I(3.88), I(4), I(0.3),
           body[2], sz=10, col=GRAY2)

    # 右下: 会社名ウォーターマーク
    _t(slide, W - I(3.0), H - I(0.7), I(2.8), I(0.35),
       'SMART STRATEGY', sz=16, bold=True, col=GRAY1, align=PP_ALIGN.RIGHT)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 2. 目次スライド
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_toc(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs, headline.split('　')[0] if '　' in headline else '')
    _slide_title(slide, prs, headline)

    item_h = I(0.65)
    start_y = I(1.55)
    for i, b in enumerate(body[:8]):
        iy = start_y + i * (item_h + I(0.1))
        if iy + item_h > H - I(0.5): break
        _, text = _parse(b)
        text = b if not text else f'{i+1}. {text}'
        # ハイライトアイテム（active）は薄青背景
        is_active = '★' in b or '【現在】' in b
        bg = LBLUE if is_active else WHITE
        lc_col = BLUE2 if is_active else None
        _r(slide, I(0.5), iy, W - I(1.0), item_h - I(0.05),
           fill=bg, lc=lc_col if is_active else None, lw=0.5)
        clean = b.replace('★', '').replace('【現在】', '').strip()
        _t(slide, I(0.8), iy + I(0.12), W - I(1.6), item_h - I(0.2),
           f'{i+1}. {clean}', sz=14, col=BLUE2 if is_active else GRAY1)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 3. ツリーマップ（市場・構造分析）
# 左:ツリーマップ / 右:KPIカード
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_treemap(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs)
    _slide_title(slide, prs, headline,
                 body[0].split('：')[1] if body and '：' in body[0] else '')

    TY = I(1.5); TH = H - I(2.1); TW = I(5.5); TX = I(0.4)
    segs = [(b.split('：')[0] if '：' in b else b[:10]) for b in body[:6]]
    ratios = [0.42, 0.26, 0.16, 0.10, 0.04, 0.02]
    colors  = [NAVY, BLUE2, LBLUE, VLIGHT, BGRAY, LGRAY]
    tcs     = [WHITE, WHITE, BLUE1, BLUE1, GRAY2, GRAY2]

    lw = int(TW * 0.56)
    h0 = int(TH * ratios[0] / (ratios[0] + ratios[1]))
    h1 = TH - h0 - P(2)
    _r(slide, TX, TY, lw, h0, fill=colors[0])
    _t(slide, TX+I(0.1), TY+I(0.15), lw-I(0.2), h0-I(0.3),
       segs[0] if segs else '主要市場', sz=12, bold=True, col=tcs[0])
    _r(slide, TX, TY+h0+P(2), lw, h1, fill=colors[1])
    _t(slide, TX+I(0.1), TY+h0+P(2)+I(0.12), lw-I(0.2), h1-I(0.2),
       segs[1] if len(segs)>1 else 'サブ市場', sz=11, bold=True, col=tcs[1])
    rx2 = TX+lw+P(2); rw2 = TW-lw-P(2); cum = 0
    denom = sum(ratios[2:]) or 1
    for i,(r,c,tc) in enumerate(zip(ratios[2:], colors[2:], tcs[2:])):
        bh = int(TH*(r/denom)); bh = max(bh, I(0.5))
        if cum+bh > TH: bh = TH-cum
        _r(slide, rx2, TY+cum, rw2, bh-P(1), fill=c, lc=LGRAY, lw=0.3)
        lbl = segs[i+2] if i+2 < len(segs) else f'セグメント{i+3}'
        _t(slide, rx2+I(0.08), TY+cum+I(0.06), rw2-I(0.1), bh, lbl, sz=9, col=tc)
        cum += bh
    _t(slide, TX, TY+TH+I(0.06), TW, I(0.22),
       '市場セグメント構造（面積＝重要度）', sz=8, col=GRAY3)

    KX = I(6.2); KY = I(1.5); KW = I(3.8)
    for i, b in enumerate(body[:5]):
        lbl, det = _parse(b); ky = KY + i * I(1.05)
        _r(slide, KX, ky, KW, I(0.92), fill=BGRAY, lc=LGRAY, lw=0.3)
        _r(slide, KX, ky, P(4), I(0.92), fill=NAVY)
        _t(slide, KX+I(0.15), ky+I(0.05), KW-I(0.2), I(0.3),
           lbl, sz=9, bold=True, col=NAVY)
        _t(slide, KX+I(0.15), ky+I(0.38), KW-I(0.2), I(0.5),
           det, sz=9, col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 4. フロー図（プロセス・ステップ）
# 上部: 矢印フロー / 下部: 2分割詳細
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_flow(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs)
    summary = body[0].split('：')[1] if body and '：' in body[0] else ''
    _slide_title(slide, prs, headline, summary)

    n = min(len(body), 5)
    if n == 0: return

    # 上部: ネイビー背景の横フロー
    FY = I(1.55); FH = I(1.0)
    _r(slide, I(0.3), FY, W - I(0.6), FH, fill=NAVY)
    step_w = int((W - I(0.8)) / n)

    for i, b in enumerate(body[:n]):
        lbl, det = _parse(b)
        bx = I(0.4) + i * step_w; sw = step_w - I(0.05)
        # ステップボックス（白枠）
        _r(slide, bx, FY + I(0.08), sw - I(0.08), FH - I(0.16),
           fill=WHITE, lc=BLUE3, lw=0.5)
        # ステップ番号（丸）
        _r(slide, bx + I(0.08), FY + I(0.12), I(0.38), I(0.38), fill=ORANGE)
        _t(slide, bx + I(0.08), FY + I(0.1), I(0.38), I(0.38),
           f'①②③④⑤'[i], sz=11, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        _t(slide, bx + I(0.5), FY + I(0.12), sw - I(0.6), I(0.35),
           lbl or b[:12], sz=9, bold=True, col=NAVY)
        if det:
            _t(slide, bx + I(0.5), FY + I(0.5), sw - I(0.6), I(0.38),
               det[:20], sz=8, col=GRAY2)
        # 矢印（テキスト）
        if i < n - 1:
            _t(slide, bx + sw - I(0.1), FY + I(0.38), I(0.18), I(0.28),
               '▶', sz=12, col=BLUE3, align=PP_ALIGN.CENTER)

    # 下部: 詳細カード（2分割）
    left_items = body[:3]; right_items = body[3:6]
    DY = FY + FH + I(0.3); DH = H - DY - I(0.45)
    cw = int((W - I(0.9)) / 2)
    for ci, (items, title) in enumerate([(left_items, '課題の構造'), (right_items, '解決アプローチ')]):
        cx = I(0.4) + ci * (cw + I(0.1))
        _r(slide, cx, DY, cw, DH, fill=NAVY)
        _t(slide, cx, DY + I(0.08), cw, I(0.28),
           title, sz=10, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        iy = DY + I(0.4)
        for item in items:
            lbl, det = _parse(item)
            _r(slide, cx + I(0.1), iy, cw - I(0.2), I(0.55), fill=WHITE)
            _t(slide, cx + I(0.18), iy + I(0.04), cw - I(0.35), I(0.25),
               lbl or item[:18], sz=9, bold=True, col=NAVY)
            if det:
                _t(slide, cx + I(0.18), iy + I(0.3), cw - I(0.35), I(0.22),
                   det[:28], sz=8, col=GRAY2)
            iy += I(0.65)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 5. 2カラム比較（背景・目的 / 左右）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_twocol(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs)
    summary = body[0].split('：')[1] if body and '：' in body[0] else ''
    _slide_title(slide, prs, headline, summary)

    mid_items = len(body) // 2
    left_items  = body[:mid_items]
    right_items = body[mid_items:]

    CY = I(1.55); CH = H - CY - I(0.45); cw = int((W - I(0.9)) / 2)
    header_h = I(0.38)

    for ci, (items, col_title) in enumerate([(left_items, '背景'), (right_items, '目的')]):
        cx = I(0.4) + ci * (cw + I(0.1))
        # ヘッダー
        _r(slide, cx, CY, cw, header_h, fill=NAVY)
        _t(slide, cx, CY + I(0.06), cw, header_h - I(0.1),
           col_title, sz=13, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        # コンテンツエリア
        _r(slide, cx, CY + header_h, cw, CH - header_h, fill=WHITE, lc=NAVY, lw=0.75)
        iy = CY + header_h + I(0.15)
        for item in items:
            lbl, det = _parse(item)
            display = lbl or item
            # 大項目
            _r(slide, cx + I(0.12), iy, I(0.22), I(0.22), fill=NAVY)
            _t(slide, cx + I(0.42), iy, cw - I(0.55), I(0.3),
               display[:30], sz=10, bold=True, col=BLUE1)
            if det:
                _t(slide, cx + I(0.55), iy + I(0.28), cw - I(0.65), I(0.42),
                   '• ' + det, sz=9, col=GRAY1)
            iy += I(0.75)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 6. 比較テーブル（◎○△マトリクス）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_table(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs)
    summary = body[0].split('：')[1] if body and '：' in body[0] else ''
    _slide_title(slide, prs, headline, summary)

    cols = ['評価軸', '自社', '競合A', '競合B']
    col_x = [I(0.4), I(2.85), I(5.3), I(7.75)]
    col_w = [I(2.35), I(2.35), I(2.35), I(2.35)]
    HY = I(1.55); RH = I(0.78)

    for ci, (lbl, cx, cw) in enumerate(zip(cols, col_x, col_w)):
        bg = NAVY if ci == 0 else BLUE2
        _r(slide, cx, HY, cw - P(2), I(0.55), fill=bg)
        _t(slide, cx, HY+I(0.09), cw-P(2), I(0.38),
           lbl, sz=11, bold=True, col=WHITE, align=PP_ALIGN.CENTER)

    patterns = [('◎','△','○'),('○','◎','△'),('◎','○','△'),('○','△','◎'),('◎','△','◎'),('△','◎','○')]
    sc_col = {'◎': RGBColor(0x00,0x80,0x00), '○': BLUE2, '△': ORANGE}

    for ri, b in enumerate(body):
        ry = HY + I(0.55) + P(2) + ri * (RH + P(1))
        if ry + RH > H - I(0.45): break
        lbl, det = _parse(b)
        bg = VLIGHT if ri % 2 == 0 else WHITE
        for cx, cw in zip(col_x, col_w):
            _r(slide, cx, ry, cw-P(2), RH, fill=bg, lc=LGRAY, lw=0.3)
        _t(slide, col_x[0]+I(0.1), ry+I(0.08), col_w[0]-I(0.15), RH-I(0.1),
           lbl or b[:16], sz=10, bold=True, col=BLUE1)
        if det:
            _t(slide, col_x[0]+I(0.1), ry+I(0.38), col_w[0]-I(0.15), RH-I(0.4),
               det[:22], sz=8, col=GRAY3)
        pat = patterns[ri % len(patterns)]
        for sc, cx, cw in zip(pat, col_x[1:], col_w[1:]):
            _t(slide, cx, ry+I(0.06), cw-P(2), RH-I(0.1),
               sc, sz=16, bold=True, col=sc_col.get(sc, GRAY2), align=PP_ALIGN.CENTER)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 7. ガントチャート＋収益グラフ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_gantt(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs)
    summary = body[0].split('：')[1] if body and '：' in body[0] else ''
    _slide_title(slide, prs, headline, summary)

    GX=I(0.4); GY=I(1.55); GW=I(5.8); GH=H-I(2.1)
    phases=['M1-2','M3-4','M5-6','M7-9','M10-12','Y2-3']
    pw=GW//len(phases)
    _r(slide, GX, GY, GW, I(0.4), fill=NAVY)
    for i, ph in enumerate(phases):
        _t(slide, GX+i*pw, GY+I(0.06), pw, I(0.28),
           ph, sz=8, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        if i > 0: _r(slide, GX+i*pw, GY, P(0.5), GH, fill=GRAY3)

    task_defs = [(0,2,NAVY),(0,3,BLUE2),(1,1,ORANGE),(2,4,BLUE3),(2,1,RED),(4,2,NAVY)]
    lw2 = int(GW*0.30); ROW = int((GH-I(0.42))//max(len(body[:6]),1))
    for ri, b in enumerate(body[:6]):
        lbl, _ = _parse(b); ry = GY+I(0.42)+ri*ROW
        start, dur, col = task_defs[ri%len(task_defs)]
        _r(slide, GX, ry, GW, ROW-P(1), fill=BGRAY)
        _r(slide, GX, ry, P(0.5), ROW-P(1), fill=col)
        _t(slide, GX+I(0.08), ry+P(4), lw2-I(0.1), ROW-P(8),
           lbl or f'タスク{ri+1}', sz=8, col=GRAY1)
        bx = GX+lw2+start*pw; bw = min(dur*pw-P(2), GW-lw2-start*pw-P(2))
        if bw > 0: _r(slide, bx, ry+P(4), bw, ROW-P(8), fill=col)

    # 右: 積み上げ棒グラフ
    BX=I(6.5); BY=I(1.55); BW=I(3.2); BH=GH-I(0.4)
    _t(slide, BX, BY, BW, I(0.28), '売上推移（万円/月）', sz=8, bold=True, col=GRAY2)
    periods=['6ヶ月','12ヶ月','24ヶ月']
    vals=[[300,100,50],[500,180,120],[800,280,220]]
    bc=[NAVY,BLUE2,ORANGE]; bw2=(BW-I(0.4))//3; mv=1400
    for bi,(period,vl) in enumerate(zip(periods,vals)):
        bx2=BX+I(0.2)+bi*(bw2+I(0.1)); cum=0; total=sum(vl)
        for v,c in zip(vl,bc):
            bh2=int((v/mv)*(BH-I(1.0))); by2=BY+BH-I(0.5)-cum-bh2
            _r(slide,bx2,by2,bw2-P(1),bh2,fill=c); cum+=bh2
        _t(slide,bx2,BY+BH-I(0.42),bw2,I(0.36),period,sz=7,col=GRAY2,align=PP_ALIGN.CENTER)
        _t(slide,bx2,BY+BH-I(0.78),bw2,I(0.34),f'{total}万',sz=8,bold=True,col=NAVY,align=PP_ALIGN.CENTER)
    for i,(lbl,c) in enumerate(zip(['体験料','F&B','その他'],bc)):
        _r(slide,BX+I(0.2)+i*I(1.0),BY+BH+I(0.05),I(0.16),I(0.16),fill=c)
        _t(slide,BX+I(0.4)+i*I(1.0),BY+BH+I(0.03),I(0.75),I(0.2),lbl,sz=7,col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 8. 論点リスト（主要論点スタイル）
# オレンジ丸ラベル + ブレットリスト
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_detail(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs)
    summary = body[0].split('：')[1] if body and '：' in body[0] else ''
    _slide_title(slide, prs, headline, summary)

    # ヘッダーバー（論点タイトル）
    _navy_header_box(slide, I(0.4), I(1.55), W - I(0.8), I(0.38),
                     '主要論点', sz=11)

    # bodyをグループ化（各bodyエントリを独立グループとして扱う）
    items_to_show = body[:6]
    grp_h = int((H - I(2.55)) / max(len(items_to_show), 1))
    gy = I(2.05)

    for gi, b in enumerate(items_to_show):
        cy = gy + gi * grp_h
        lbl, det = _parse(b)
        group_title = lbl if lbl else b[:14]
        detail_text = det if det else ''

        # オレンジ丸 + グループタイトル
        _r(slide, I(0.4), cy + I(0.04), I(1.5), grp_h - I(0.08), fill=VLIGHT, lc=LBLUE, lw=0.5)
        _r(slide, I(0.48), cy + I(0.1), I(0.48), I(0.48), fill=ORANGE)
        _t(slide, I(0.48), cy + I(0.08), I(0.48), I(0.48),
           str(gi+1), sz=12, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        _t(slide, I(0.42), cy + I(0.62), I(1.6), grp_h - I(0.72),
           group_title[:12], sz=8, bold=True, col=NAVY, align=PP_ALIGN.CENTER)

        # 詳細テキスト（右側）
        _r(slide, I(2.05), cy + I(0.04), W - I(2.45), grp_h - I(0.08),
           fill=WHITE, lc=LGRAY, lw=0.3)
        if detail_text:
            # 詳細を箇条書きに分割
            bullets = detail_text.split('・') if '・' in detail_text else [detail_text]
            ih = (grp_h - I(0.15)) // max(len(bullets[:3]), 1)
            iy = cy + I(0.06)
            for bullet in bullets[:3]:
                if bullet.strip():
                    _r(slide, I(2.12), iy + I(0.1), I(0.16), I(0.16), fill=NAVY)
                    _t(slide, I(2.35), iy + I(0.04), W - I(2.7), ih - I(0.06),
                       bullet.strip()[:50], sz=9, col=GRAY1)
                    iy += ih


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 9. まとめ（末尾固定）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_summary(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _r(slide, 0, 0, W, H, fill=WHITE)
    _page_frame(slide, prs)
    _slide_title(slide, prs, headline)

    kpis = body[:3]; kw = (W - I(1.0)) // 3; ky = I(1.55); kh = I(1.6)
    for i, b in enumerate(body[:3]):
        lbl, val = _parse(b)
        kx = I(0.4) + i * (kw + I(0.1))
        _r(slide, kx, ky, kw, kh, fill=BGRAY, lc=LGRAY, lw=0.5)
        _r(slide, kx, ky, kw, P(4), fill=[NAVY, BLUE2, ORANGE][i])
        _t(slide, kx+I(0.12), ky+I(0.1), kw-I(0.24), I(0.32),
           lbl, sz=10, bold=True, col=[NAVY,BLUE2,ORANGE][i])
        _t(slide, kx+I(0.12), ky+I(0.48), kw-I(0.24), I(0.65),
           val[:18], sz=12, bold=True, col=NAVY, wrap=True)
        rest = val[18:]
        if rest:
            _t(slide, kx+I(0.12), ky+I(1.18), kw-I(0.24), I(0.38),
               rest[:22], sz=8, col=GRAY2)

    remaining = body[3:]
    _t(slide, I(0.4), I(3.4), I(9.8), I(0.32),
       'Next Actions', sz=12, bold=True, col=NAVY)
    _r(slide, I(0.4), I(3.74), I(9.8), P(1.5), fill=BLUE2)
    half = (len(remaining)+1)//2; cw2 = (W-I(1.0))//2
    for i, b in enumerate(remaining[:8]):
        lbl, det = _parse(b)
        ci = i//half; ri = i%half
        cx = I(0.4)+ci*(cw2+I(0.2)); cy = I(3.92)+ri*I(0.85)
        if cy+I(0.75) > H-I(0.35): break
        _r(slide, cx, cy, I(0.38), I(0.38),
           fill=[NAVY,BLUE2,ORANGE][ci%3])
        _t(slide, cx, cy, I(0.38), I(0.38),
           '✓', sz=11, bold=True, col=WHITE, align=PP_ALIGN.CENTER)
        _t(slide, cx+I(0.48), cy, cw2-I(0.56), I(0.28),
           lbl or f'アクション{i+1}', sz=10, bold=True, col=BLUE1)
        _t(slide, cx+I(0.48), cy+I(0.30), cw2-I(0.56), I(0.45),
           det, sz=10, col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ルーティング（コンテンツ優先・重複回避）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DESIGN_NAMES = {
    slide_title:   'title',
    slide_toc:     'toc',
    slide_treemap: 'treemap',
    slide_flow:    'flow',
    slide_twocol:  'twocol',
    slide_table:   'table',
    slide_gantt:   'gantt',
    slide_detail:  'detail',
    slide_summary: 'summary',
}

DESIGN_CANDIDATES = {
    'データ':    [slide_treemap, slide_detail, slide_gantt],
    '比較':      [slide_table,   slide_twocol, slide_detail],
    'プロセス':  [slide_flow,    slide_gantt,  slide_detail],
    '関係性':    [slide_flow,    slide_treemap, slide_detail],
    '組織':      [slide_table,   slide_detail,  slide_flow],
    'テキスト':  [slide_detail,  slide_twocol,  slide_treemap],
    '市場':      [slide_treemap, slide_detail,  slide_table],
    '競合':      [slide_table,   slide_twocol,  slide_detail],
    '計画':      [slide_gantt,   slide_flow,    slide_detail],
    '収益':      [slide_gantt,   slide_detail,  slide_treemap],
    'フロー':    [slide_flow,    slide_gantt,   slide_detail],
    'ステップ':  [slide_flow,    slide_gantt,   slide_detail],
    'サービス':  [slide_flow,    slide_detail,  slide_treemap],
    '背景':      [slide_twocol,  slide_detail,  slide_flow],
    '目的':      [slide_twocol,  slide_detail,  slide_flow],
    '論点':      [slide_detail,  slide_twocol,  slide_table],
    '目次':      [slide_toc,     slide_detail,  slide_flow],
}

ALL_MIDDLE = [slide_treemap, slide_flow, slide_twocol, slide_table, slide_gantt, slide_detail]

def get_design_fn(purpose, content_type, page, total, used_history=None):
    if used_history is None: used_history = []
    if page == 1: return slide_title
    if page == total: return slide_summary
    avoid = set((used_history[-3:] if len(used_history) >= 3 else used_history) + ['title','summary','toc'])

    candidates = []
    for key in [content_type, purpose]:
        if not key: continue
        for kw, fns in DESIGN_CANDIDATES.items():
            if kw in key:
                candidates.extend(fns)
                break
    seen = set(); unique = []
    for fn in candidates:
        n = DESIGN_NAMES.get(fn,'')
        if n not in seen and n not in ('title','summary'):
            seen.add(n); unique.append(fn)
    for fn in unique:
        if DESIGN_NAMES.get(fn,'') not in avoid: return fn
    for fn in ALL_MIDDLE:
        if DESIGN_NAMES.get(fn,'') not in avoid: return fn
    counts = {DESIGN_NAMES[fn]: used_history.count(DESIGN_NAMES[fn]) for fn in ALL_MIDDLE}
    least = min(counts, key=counts.get)
    for fn in ALL_MIDDLE:
        if DESIGN_NAMES[fn] == least: return fn
    return slide_detail
