"""
SlideAI Design System v4 — コンテンツ優先・デザイン非重複
設計原則:
  - 構成設計段階でスライドタイプを決定（content_typeとpurposeから）
  - 7種類のデザインプールから選択、直近3枚と同じものを使わない
  - 先頭=タイトル・末尾=まとめ は固定
  - ページ数無制限に対応（8枚以上でも全て異なるレイアウトを維持）
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

# ── カラーパレット ──
NAVY   = RGBColor(0x1A, 0x28, 0x4E)
BLUE   = RGBColor(0x2E, 0x72, 0xB5)
LBLUE  = RGBColor(0xBE, 0xD8, 0xF0)
TEAL   = RGBColor(0x00, 0x6E, 0x7A)
LTEAL  = RGBColor(0xCC, 0xEA, 0xED)
ORANGE = RGBColor(0xE2, 0x6C, 0x1A)
GREEN  = RGBColor(0x1C, 0x7A, 0x48)
RED    = RGBColor(0xBE, 0x2E, 0x2E)
PURPLE = RGBColor(0x6B, 0x3A, 0x9A)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BGRAY  = RGBColor(0xF3, 0xF5, 0xF8)
LGRAY  = RGBColor(0xE8, 0xEA, 0xEE)
GRAY1  = RGBColor(0x1A, 0x1A, 0x1A)
GRAY2  = RGBColor(0x44, 0x44, 0x44)
GRAY3  = RGBColor(0x88, 0x88, 0x88)
DIVIDER = RGBColor(0xCC, 0xCC, 0xCC)
PALETTE = [BLUE, TEAL, ORANGE, GREEN, PURPLE, RED, RGBColor(0x8B,0x6B,0x10)]

def I(n): return int(Inches(n))
def P(n): return int(Pt(n))

def _r(slide, x, y, w, h, fill=WHITE, lc=None, lw=0.5):
    sh = slide.shapes.add_shape(1, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if lc: sh.line.color.rgb = lc; sh.line.width = Pt(lw)
    else:   sh.line.fill.background()
    return sh

def _t(slide, x, y, w, h, text, sz=11, bold=False, col=None,
       align=PP_ALIGN.LEFT, italic=False, wrap=True):
    if not text: return None
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame; tf.word_wrap = wrap; tf.auto_size = None
    bp  = tf._txBody.find(qn('a:bodyPr'))
    if bp is not None:
        bp.set('lIns','45720'); bp.set('rIns','45720')
        bp.set('tIns','22860'); bp.set('bIns','22860')
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(sz); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = col or GRAY1
    return txb

def _header(slide, prs, title):
    W = int(prs.slide_width)
    _t(slide, I(0.5), I(0.22), W-I(1.0), I(0.62), title, sz=21, bold=True, col=NAVY)
    _r(slide, I(0.5), I(0.9), W-I(1.0), P(2), fill=BLUE)

def _parse(b):
    for sep in ('：', ':'):
        sp = b.find(sep)
        if sp > 0: return b[:sp].strip(), b[sp+1:].strip()
    return '', b.strip()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 1. TITLE — タイトル（先頭固定）
# 左: 概要テキスト / 右: ピラミッド構造図
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_title(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _t(slide, I(0.5), I(0.35), I(6.2), I(1.6), headline, sz=24, bold=True, col=NAVY)
    _r(slide, I(0.5), I(2.1), P(3), I(4.2), fill=BLUE)
    for i, b in enumerate(body[:5]):
        lbl, det = _parse(b)
        by = I(2.2) + i * I(0.85)
        if lbl:
            _t(slide, I(0.8), by, I(5.5), I(0.3), lbl, sz=10, bold=True, col=BLUE)
        _t(slide, I(0.8), by+I(0.3), I(5.5), I(0.5), det, sz=10, col=GRAY2)
    # ピラミッド
    layers = ['実行計画', '事業モデル', 'ビジョン']
    lc2 = [LBLUE, BLUE, NAVY]; tc2 = [GRAY1, WHITE, WHITE]
    rx=I(7.0); rh=I(1.4); rw=I(3.2)
    for i, (lbl, fc, tc3) in enumerate(zip(layers, lc2, tc2)):
        li = len(layers)-1-i; shrink=li*I(0.42)
        bx=rx+shrink; bw=rw-shrink*2
        _r(slide, bx, I(1.4)+i*(rh-P(2)), bw, rh-P(4), fill=fc)
        _t(slide, bx, I(1.4)+i*(rh-P(2))+I(0.38), bw, rh-I(0.5),
           lbl, sz=11, bold=True, col=tc3, align=PP_ALIGN.CENTER)
    _t(slide, I(7.0), I(1.4)+3*(rh-P(2))+P(4), I(3.2), I(0.28),
       '事業構造', sz=8, col=GRAY3, align=PP_ALIGN.CENTER)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 2. TREEMAP — 市場・構造分析
# 左: ツリーマップ（面積=重要度） / 右: KPIカード
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_treemap(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _header(slide, prs, headline)
    TX=I(0.5); TY=I(1.3); TW=I(5.5); TH=I(5.2)
    segs = [(b.split('：')[0] if '：' in b else b[:10]) for b in body[:6]]
    ratios=[0.42,0.26,0.16,0.10,0.04,0.02]
    colors=[BLUE,TEAL,LBLUE,LTEAL,LGRAY,BGRAY]
    tcs=[WHITE,WHITE,GRAY1,GRAY1,GRAY2,GRAY2]
    lw=int(TW*0.56)
    h0=int(TH*(ratios[0]/(ratios[0]+ratios[1])))
    h1=TH-h0-P(2)
    _r(slide,TX,TY,lw,h0,fill=colors[0])
    _t(slide,TX+I(0.12),TY+I(0.15),lw-I(0.2),h0-I(0.3),
       segs[0] if segs else '主要市場',sz=13,bold=True,col=tcs[0])
    _r(slide,TX,TY+h0+P(2),lw,h1,fill=colors[1])
    _t(slide,TX+I(0.12),TY+h0+P(2)+I(0.12),lw-I(0.2),h1-I(0.2),
       segs[1] if len(segs)>1 else 'サブ市場',sz=11,bold=True,col=tcs[1])
    rx2=TX+lw+P(2); rw2=TW-lw-P(2); cum=0
    denom=sum(ratios[2:])
    for i,(r,c,tc) in enumerate(zip(ratios[2:],colors[2:],tcs[2:])):
        bh=int(TH*(r/denom)); bh=max(bh,I(0.5))
        if cum+bh>TH: bh=TH-cum
        _r(slide,rx2,TY+cum,rw2,bh-P(1),fill=c)
        lbl=segs[i+2] if i+2<len(segs) else f'セグメント{i+3}'
        _t(slide,rx2+I(0.08),TY+cum+I(0.06),rw2-I(0.1),bh,lbl,sz=9,col=tc)
        cum+=bh
    _t(slide,TX,TY+TH+I(0.06),TW,I(0.25),'■ 市場セグメント構造（面積＝重要度）',sz=8,col=GRAY3)
    KX=I(6.3); KY=I(1.3); KW=I(4.0)
    for i,b in enumerate(body[:5]):
        lbl,det=_parse(b); ky=KY+i*I(1.08)
        col=PALETTE[i%len(PALETTE)]
        _r(slide,KX,ky,KW,I(0.96),fill=BGRAY,lc=LGRAY,lw=0.3)
        _r(slide,KX,ky,P(4),I(0.96),fill=col)
        _t(slide,KX+I(0.18),ky+I(0.06),KW-I(0.22),I(0.32),lbl,sz=10,bold=True,col=col)
        _t(slide,KX+I(0.18),ky+I(0.42),KW-I(0.22),I(0.5),det,sz=10,col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 3. FLOW — プロセス・ステップ
# 上: 矢印フロー / 下: 詳細2列カード
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_flow(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _header(slide, prs, headline)
    n=min(len(body),6); FY=I(1.3); FH=I(1.8)
    if n==0: return
    step_w=int((W-I(1.0))/n)
    for i,b in enumerate(body[:n]):
        lbl,det=_parse(b)
        bx=I(0.5)+i*step_w; sw=step_w-I(0.12)
        col=PALETTE[i%len(PALETTE)]
        _r(slide,bx,FY,sw,FH,fill=col)
        _r(slide,bx,FY,I(0.42),I(0.42),fill=WHITE)
        _t(slide,bx,FY,I(0.42),I(0.42),str(i+1),sz=13,bold=True,col=col,align=PP_ALIGN.CENTER)
        _t(slide,bx+I(0.05),FY+I(0.5),sw-I(0.1),I(0.5),
           lbl or det[:8],sz=10,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        if det and det!=lbl:
            _t(slide,bx+I(0.05),FY+I(1.05),sw-I(0.1),I(0.65),
               det[:25],sz=8,col=RGBColor(0xDD,0xEE,0xFF),align=PP_ALIGN.CENTER)
        if i<n-1:
            _t(slide,bx+sw+I(0.01),FY+I(0.7),I(0.1),I(0.4),'▶',sz=10,col=GRAY3,align=PP_ALIGN.CENTER)
    half=(n+1)//2; cw=int((W-I(1.2))/2); DY=FY+FH+I(0.3)
    for i,b in enumerate(body[:n]):
        lbl,det=_parse(b); col=PALETTE[i%len(PALETTE)]
        ci=i//half; ri=i%half
        cx=I(0.5)+ci*(cw+I(0.2)); cy=DY+ri*I(1.0)
        _r(slide,cx,cy,cw,I(0.88),fill=BGRAY,lc=LGRAY,lw=0.3)
        _r(slide,cx,cy,P(4),I(0.88),fill=col)
        _t(slide,cx+I(0.18),cy+I(0.05),cw-I(0.22),I(0.3),lbl or f'ステップ{i+1}',sz=9,bold=True,col=col)
        _t(slide,cx+I(0.18),cy+I(0.38),cw-I(0.22),I(0.46),det,sz=9,col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 4. TABLE — 比較マトリクス（◎○△）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_table(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _header(slide, prs, headline)
    cols=['評価軸','自社','競合A','競合B']
    cw=[I(2.4),I(2.4),I(2.4),I(2.4)]
    cx=[I(0.5),I(3.0),I(5.5),I(7.95)]
    HY=I(1.3); RH=I(0.82)
    for ci,(lbl,x,w) in enumerate(zip(cols,cx,cw)):
        bg=NAVY if ci==0 else BLUE
        _r(slide,x,HY,w-P(2),I(0.62),fill=bg)
        _t(slide,x,HY+I(0.1),w-P(2),I(0.5),lbl,sz=11,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
    patterns=[('◎','△','○'),('○','◎','△'),('◎','○','△'),('○','△','◎'),('◎','△','◎'),('△','◎','○')]
    sc={'◎':GREEN,'○':BLUE,'△':ORANGE}
    for ri,b in enumerate(body):
        ry=HY+I(0.62)+P(2)+ri*(RH+P(1))
        if ry+RH>H-I(0.3): break
        lbl,det=_parse(b); bg=BGRAY if ri%2==0 else WHITE
        for x,w in zip(cx,cw):
            _r(slide,x,ry,w-P(2),RH,fill=bg,lc=DIVIDER,lw=0.3)
        _t(slide,cx[0]+I(0.1),ry+I(0.08),cw[0]-I(0.2),RH-I(0.1),lbl or b[:14],sz=10,bold=True,col=NAVY)
        if det:
            _t(slide,cx[0]+I(0.1),ry+I(0.4),cw[0]-I(0.2),RH-I(0.4),det[:22],sz=8,col=GRAY3)
        pat=patterns[ri%len(patterns)]
        for s,x,w in zip(pat,cx[1:],cw[1:]):
            _t(slide,x,ry+I(0.08),w-P(2),RH-I(0.1),s,sz=16,bold=True,col=sc.get(s,GRAY2),align=PP_ALIGN.CENTER)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 5. GANTT — ロードマップ＋収益グラフ
# 左: ガントチャート / 右: 積み上げ棒グラフ
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_gantt(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _header(slide, prs, headline)
    GX=I(0.5); GY=I(1.3); GW=I(6.0); GH=H-I(1.8)
    phases=['M1-2','M3-4','M5-6','M7-9','M10-12','Y2-3']
    pw=GW//len(phases)
    _r(slide,GX,GY,GW,I(0.42),fill=NAVY)
    for i,ph in enumerate(phases):
        _t(slide,GX+i*pw,GY+I(0.05),pw,I(0.32),ph,sz=8,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        if i>0: _r(slide,GX+i*pw,GY,P(0.5),GH,fill=RGBColor(0x88,0x88,0x88))
    task_defs=[(0,2,BLUE),(0,3,TEAL),(1,1,ORANGE),(2,4,GREEN),(2,1,RED),(4,2,BLUE)]
    lw2=int(GW*0.30); ROW=int((GH-I(0.45))//max(len(body[:6]),1))
    for ri,b in enumerate(body[:6]):
        lbl,_=_parse(b); ry=GY+I(0.45)+ri*ROW
        start,dur,col=task_defs[ri%len(task_defs)]
        _r(slide,GX,ry,GW,ROW-P(1),fill=BGRAY)
        _t(slide,GX+I(0.05),ry+P(3),lw2-I(0.1),ROW-P(6),lbl or f'タスク{ri+1}',sz=8,col=GRAY1)
        bx=GX+lw2+start*pw; bw=min(dur*pw-P(2),GW-lw2-start*pw-P(2))
        if bw>0: _r(slide,bx,ry+P(4),bw,ROW-P(8),fill=col)
    _t(slide,GX,GY+GH+I(0.05),GW,I(0.2),'■ 実施ロードマップ',sz=8,col=GRAY3)
    BX=I(7.0); BY=I(1.3); BW=I(3.2); BH=GH-I(0.5)
    _t(slide,BX,BY,BW,I(0.3),'■ 売上推移（万円/月）',sz=8,bold=True,col=GRAY2)
    periods=['6ヶ月','12ヶ月','24ヶ月']
    vals=[[300,100,50],[500,180,120],[800,280,220]]
    bc=[BLUE,TEAL,ORANGE]; bw2=(BW-I(0.4))//3; mv=1400
    for bi,(period,vl) in enumerate(zip(periods,vals)):
        bx2=BX+I(0.2)+bi*(bw2+I(0.1)); cum=0; total=sum(vl)
        for v,c in zip(vl,bc):
            bh2=int((v/mv)*(BH-I(1.0))); by2=BY+BH-I(0.5)-cum-bh2
            _r(slide,bx2,by2,bw2-P(1),bh2,fill=c); cum+=bh2
        _t(slide,bx2,BY+BH-I(0.42),bw2,I(0.38),period,sz=7,col=GRAY2,align=PP_ALIGN.CENTER)
        _t(slide,bx2,BY+BH-I(0.82),bw2,I(0.36),f'{total}万',sz=8,bold=True,col=NAVY,align=PP_ALIGN.CENTER)
    for i,(lbl,c) in enumerate(zip(['体験料','F&B','その他'],bc)):
        _r(slide,BX+I(0.2)+i*I(1.0),BY+BH+I(0.08),I(0.18),I(0.18),fill=c)
        _t(slide,BX+I(0.42)+i*I(1.0),BY+BH+I(0.06),I(0.78),I(0.22),lbl,sz=7,col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 6. DETAIL — 詳細説明（ポイントリスト＋KPI）
# 左: 番号付きポイント / 右: 縦KPIカード
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_detail(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _header(slide, prs, headline)
    LW=I(6.2); LY=I(1.3); rh=I(0.88)
    for i,b in enumerate(body[:6]):
        lbl,det=_parse(b); col=PALETTE[i%len(PALETTE)]
        cy=LY+i*rh
        if cy+rh>H-I(0.3): break
        _r(slide,I(0.5),cy,LW,rh-P(3),fill=BGRAY,lc=LGRAY,lw=0.3)
        _r(slide,I(0.5),cy,P(4),rh-P(3),fill=col)
        _r(slide,I(0.68),cy+I(0.1),I(0.32),I(0.32),fill=col)
        _t(slide,I(0.68),cy+I(0.04),I(0.32),I(0.32),str(i+1),sz=10,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        _t(slide,I(1.12),cy+I(0.04),LW-I(0.7),I(0.3),lbl or b[:18],sz=10,bold=True,col=col)
        _t(slide,I(1.12),cy+I(0.38),LW-I(0.7),I(0.46),det,sz=10,col=GRAY2)
    RX=I(7.0); RY=I(1.3); RW=I(3.2)
    kpis=body[:3]; kh=int((H-I(1.8))//max(len(kpis),1))
    for i,b in enumerate(kpis):
        lbl,val=_parse(b); col=PALETTE[i%len(PALETTE)]
        ky=RY+i*kh
        _r(slide,RX,ky,RW,kh-P(4),fill=BGRAY,lc=LGRAY,lw=0.3)
        _r(slide,RX,ky,RW,P(4),fill=col)
        _t(slide,RX+I(0.15),ky+I(0.1),RW-I(0.3),I(0.35),lbl,sz=10,bold=True,col=col)
        words=val.split('、') if '、' in val else [val]
        line1=words[0][:16]; rest=val[len(line1):]
        _t(slide,RX+I(0.15),ky+I(0.5),RW-I(0.3),I(0.5),line1,sz=13,bold=True,col=NAVY,wrap=True)
        if rest:
            _t(slide,RX+I(0.15),ky+I(1.05),RW-I(0.3),I(0.45),rest[:20],sz=9,col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 7. SUMMARY — まとめ（末尾固定）
# 上: KPI3列 / 下: アクションリスト
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def slide_summary(slide, prs, headline, body):
    W, H = int(prs.slide_width), int(prs.slide_height)
    _header(slide, prs, headline)
    kpis=body[:3]; kw=(W-I(1.2))//3; ky=I(1.3); kh=I(1.75)
    for i,b in enumerate(kpis):
        lbl,val=_parse(b); col=PALETTE[i]
        kx=I(0.5)+i*(kw+I(0.1))
        _r(slide,kx,ky,kw,kh,fill=BGRAY,lc=LGRAY,lw=0.5)
        _r(slide,kx,ky,kw,P(4),fill=col)
        _t(slide,kx+I(0.15),ky+I(0.1),kw-I(0.3),I(0.35),lbl,sz=10,bold=True,col=col)
        short=val[:18] if val else ''
        _t(slide,kx+I(0.15),ky+I(0.5),kw-I(0.3),I(0.72),short,sz=12,bold=True,col=NAVY,wrap=True)
        rest=val[18:] if len(val)>18 else ''
        _t(slide,kx+I(0.15),ky+I(1.28),kw-I(0.3),I(0.4),rest,sz=8,col=GRAY2)
    remaining=body[3:]
    _t(slide,I(0.5),I(3.3),I(9.8),I(0.36),'Next Actions',sz=12,bold=True,col=NAVY)
    _r(slide,I(0.5),I(3.68),I(9.8),P(1.5),fill=BLUE)
    half=(len(remaining)+1)//2; cw2=(W-I(1.2))//2
    for i,b in enumerate(remaining[:8]):
        lbl,det=_parse(b); col=PALETTE[i%len(PALETTE)]
        ci=i//half; ri=i%half
        cx=I(0.5)+ci*(cw2+I(0.2)); cy=I(3.88)+ri*I(0.88)
        if cy+I(0.8)>H-I(0.15): break
        _r(slide,cx,cy,I(0.42),I(0.42),fill=col)
        _t(slide,cx,cy,I(0.42),I(0.42),'✓',sz=12,bold=True,col=WHITE,align=PP_ALIGN.CENTER)
        _t(slide,cx+I(0.52),cy,cw2-I(0.6),I(0.3),lbl or f'アクション{i+1}',sz=10,bold=True,col=col)
        _t(slide,cx+I(0.52),cy+I(0.32),cw2-I(0.6),I(0.48),det,sz=10,col=GRAY2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# コンテンツタイプ → デザイン候補マッピング
# 各タイプに「最適デザイン」と「代替デザイン」を定義
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# デザイン名（重複チェックに使用）
DESIGN_NAMES = {
    slide_title:   'title',
    slide_treemap: 'treemap',
    slide_flow:    'flow',
    slide_table:   'table',
    slide_gantt:   'gantt',
    slide_detail:  'detail',
    slide_summary: 'summary',
}

# コンテンツタイプ/purposeキーワード → デザイン優先順位リスト
# [最適, 第2候補, 第3候補, ...]
DESIGN_CANDIDATES = {
    'データ':    [slide_treemap, slide_detail, slide_gantt],
    '比較':      [slide_table,   slide_treemap, slide_detail],
    'プロセス':  [slide_flow,    slide_gantt,   slide_detail],
    '関係性':    [slide_flow,    slide_treemap, slide_detail],
    '組織':      [slide_table,   slide_detail,  slide_flow],
    'テキスト':  [slide_detail,  slide_flow,    slide_treemap],
    '地図':      [slide_treemap, slide_detail,  slide_table],
    # purposeキーワードマッチ
    '市場':      [slide_treemap, slide_detail,  slide_table],
    '競合':      [slide_table,   slide_treemap, slide_detail],
    '比較':      [slide_table,   slide_treemap, slide_detail],
    'スケジュール': [slide_gantt, slide_flow,   slide_detail],
    'ロードマップ': [slide_gantt, slide_flow,   slide_detail],
    '計画':      [slide_gantt,   slide_flow,    slide_detail],
    '収益':      [slide_gantt,   slide_detail,  slide_treemap],
    'フロー':    [slide_flow,    slide_gantt,   slide_detail],
    'ステップ':  [slide_flow,    slide_gantt,   slide_detail],
    'サービス':  [slide_flow,    slide_detail,  slide_treemap],
    'まとめ':    [slide_summary, slide_detail,  slide_flow],
    '結論':      [slide_summary, slide_detail,  slide_flow],
    'アクション':[slide_summary, slide_detail,  slide_flow],
}

# 全デザインのローテーション用リスト（先頭・末尾固定分を除く）
ALL_DESIGNS = [slide_treemap, slide_flow, slide_table, slide_gantt, slide_detail]

def get_design_fn(purpose: str, content_type: str, page: int, total: int,
                  used_history: list = None) -> callable:
    """
    コンテンツ内容から最適なデザイン関数を選択。
    used_history: 直近使用したデザイン名のリスト（重複回避に使用）
    """
    if used_history is None:
        used_history = []

    # 先頭・末尾は固定
    if page == 1:
        return slide_title
    if page == total:
        return slide_summary

    # 直近N枚で使われたデザイン名セット（Nは小さいほど変化豊か）
    avoid = set(used_history[-3:]) if len(used_history) >= 3 else set(used_history)
    # summary と title は常に避ける（中盤では使わない）
    avoid.update(['title', 'summary'])

    # 候補リストを収集（content_type → purpose → デフォルト順）
    candidates = []
    for key in [content_type, purpose]:
        if not key: continue
        for kw, fns in DESIGN_CANDIDATES.items():
            if kw in key:
                candidates.extend(fns)
                break

    # candidates に重複があれば除去、順序維持
    seen = set(); unique_candidates = []
    for fn in candidates:
        name = DESIGN_NAMES.get(fn, '')
        if name not in seen and name not in ('title', 'summary'):
            seen.add(name); unique_candidates.append(fn)

    # 候補から「避けるデザイン」を除いた最初のものを選択
    for fn in unique_candidates:
        name = DESIGN_NAMES.get(fn, '')
        if name not in avoid:
            return fn

    # 候補が全て避ける対象の場合、全デザインから選択
    for fn in ALL_DESIGNS:
        name = DESIGN_NAMES.get(fn, '')
        if name not in avoid:
            return fn

    # 最終フォールバック: 最も使われていないデザイン
    counts = {DESIGN_NAMES[fn]: used_history.count(DESIGN_NAMES[fn]) for fn in ALL_DESIGNS}
    least = min(counts, key=counts.get)
    for fn in ALL_DESIGNS:
        if DESIGN_NAMES[fn] == least:
            return fn

    return slide_detail
