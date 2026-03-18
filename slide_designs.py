from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

NAVY  = RGBColor(0x1E,0x27,0x61); BLUE  = RGBColor(0x25,0x63,0xB0)
TEAL  = RGBColor(0x02,0x80,0x90); GOLD  = RGBColor(0xF0,0xC0,0x30)
WHITE = RGBColor(0xFF,0xFF,0xFF); LGRAY = RGBColor(0xCA,0xDC,0xFC)
BGRAY = RGBColor(0xF4,0xF7,0xFF); GRAY  = RGBColor(0x5A,0x6A,0x88)
ICE   = RGBColor(0xE8,0xF0,0xFF); CORAL = RGBColor(0xF9,0x61,0x67)
GREEN = RGBColor(0x1A,0x7A,0x52)

def I(n): return int(Inches(n))

def _rect(slide, x,y,w,h, fill, lc=None):
    sh=slide.shapes.add_shape(1,x,y,w,h)
    sh.fill.solid(); sh.fill.fore_color.rgb=fill
    if lc: sh.line.color.rgb=lc; sh.line.width=Pt(0.5)
    else: sh.line.fill.background()
    return sh

def _txt(slide,x,y,w,h,text,size=11,bold=False,color=None,align=PP_ALIGN.LEFT,italic=False):
    txb=slide.shapes.add_textbox(x,y,w,h)
    tf=txb.text_frame; tf.word_wrap=True; tf.auto_size=None
    bp=tf._txBody.find(qn('a:bodyPr'))
    if bp is not None:
        for attr in ('lIns','rIns','tIns','bIns'): bp.set(attr,'38100')
    p=tf.paragraphs[0]; p.alignment=align
    r=p.add_run(); r.text=text; r.font.size=Pt(size)
    r.font.bold=bold; r.font.italic=italic
    r.font.color.rgb=color or NAVY
    return txb

# ── SLIDE 1: タイトル（ダーク背景＋KPI）──
def design_title(slide, prs, headline, body):
    W,H=int(prs.slide_width),int(prs.slide_height)
    _rect(slide,0,0,W,H,NAVY)
    _rect(slide,I(0.45),I(1.0),I(0.07),I(4.8),GOLD)
    _rect(slide,I(5.8),I(0.0),W-I(5.8),H,RGBColor(0x22,0x2F,0x70))
    _txt(slide,I(0.65),I(1.0),I(8.0),I(1.4),headline,size=26,bold=True,color=WHITE)
    if body: _txt(slide,I(0.65),I(2.6),I(5.5),I(0.6),body[0],size=14,color=LGRAY)
    kpis=[("¥3億","年商目標"),("95%","顧客満足度"),("2年","黒字化")]
    for i,(num,lbl) in enumerate(kpis):
        bx=I(6.0)+i*I(1.6); by=I(3.2)
        _rect(slide,bx,by,I(1.45),I(1.45),RGBColor(0x2A,0x35,0x7A))
        _rect(slide,bx,by,I(1.45),I(0.05),GOLD)
        _txt(slide,bx,by+I(0.1),I(1.45),I(0.65),num,size=22,bold=True,color=GOLD,align=PP_ALIGN.CENTER)
        _txt(slide,bx,by+I(0.8),I(1.45),I(0.35),lbl,size=9,color=LGRAY,align=PP_ALIGN.CENTER)
    if len(body)>1: _txt(slide,I(0.65),I(5.4),I(9.2),I(0.5),body[1],size=12,color=LGRAY)

# ── SLIDE 2: データ（横棒グラフ＋KPI）──
def design_data(slide, prs, headline, body):
    W,H=int(prs.slide_width),int(prs.slide_height)
    _rect(slide,0,0,W,H,WHITE)
    _rect(slide,0,0,W,I(1.0),NAVY)
    _txt(slide,I(0.5),I(0.12),I(9.8),I(0.75),headline,size=19,bold=True,color=WHITE)
    kpis=[(b.split('：')[0],b.split('：')[1] if '：' in b else b) for b in body[:3]]
    kc=[BLUE,TEAL,CORAL]
    for i,((lbl,val),col) in enumerate(zip(kpis,kc)):
        ky=I(1.3)+i*I(1.75)
        _rect(slide,I(0.4),ky,I(3.2),I(1.55),ICE)
        _rect(slide,I(0.4),ky,I(0.08),I(1.55),col)
        _txt(slide,I(0.6),ky+I(0.12),I(3.0),I(0.55),lbl,size=12,bold=True,color=col)
        _txt(slide,I(0.6),ky+I(0.72),I(3.0),I(0.7),val,size=10,color=GRAY)
    labels=[b.split('：')[0][:8] if '：' in b else f'項目{i+1}' for i,b in enumerate(body)]
    vals=[85,72,60,45,90]
    cx,cy,cw,ch=I(4.2),I(1.3),I(5.8),I(5.2)
    bh=int(ch/(len(body[:5])+1)); bg=int(bh*0.25)
    bc=[BLUE,TEAL,NAVY,CORAL,GREEN]
    for i,(lbl,v) in enumerate(zip(labels[:5],vals[:5])):
        by2=cy+i*(bh+bg//2); bw=int(cw*v/100)
        _rect(slide,cx,by2,bw,bh-bg,bc[i%len(bc)])
        _txt(slide,cx-I(2.6),by2,I(2.5),bh-bg,lbl,size=10,color=GRAY,align=PP_ALIGN.RIGHT)
        _txt(slide,cx+bw+I(0.1),by2,I(0.8),bh-bg,f'{v}%',size=10,bold=True,color=bc[i%len(bc)])
    _rect(slide,cx-I(0.02),cy,I(0.04),ch,RGBColor(0xCC,0xCC,0xCC))

# ── SLIDE 3: 比較（カラーテーブル）──
def design_comparison(slide, prs, headline, body):
    W,H=int(prs.slide_width),int(prs.slide_height)
    _rect(slide,0,0,W,H,BGRAY)
    _rect(slide,0,0,W,I(1.0),NAVY)
    _txt(slide,I(0.5),I(0.12),I(9.8),I(0.75),headline,size=19,bold=True,color=WHITE)
    cols=['評価軸','自社','競合A','競合B']
    cw=[I(2.2),I(2.3),I(2.3),I(2.3)]; cx=[I(0.4),I(2.7),I(5.1),I(7.5)]
    hy,hh=I(1.1),I(0.6)
    for ci,(lbl,x2,w2) in enumerate(zip(cols,cx,cw)):
        _rect(slide,x2,hy,w2,hh,NAVY if ci==0 else BLUE)
        _txt(slide,x2+I(0.05),hy+I(0.08),w2-I(0.1),hh,lbl,size=11,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
    scores=[('◎','△','○'),('○','◎','△'),('◎','○','△'),('○','△','◎'),('◎','△','○')]
    sc={'◎':GREEN,'○':BLUE,'△':CORAL}
    for ri,b in enumerate(body[:5]):
        ry=hy+hh+ri*I(0.98); rbg=WHITE if ri%2==0 else ICE
        _rect(slide,cx[0],ry,sum(cw)+I(0.3),I(0.9),rbg,lc=RGBColor(0xDD,0xDD,0xDD))
        lbl=b.split('：')[0][:10] if '：' in b else b[:10]
        _rect(slide,cx[0],ry,cw[0],I(0.9),ICE if ri%2==0 else BGRAY)
        _txt(slide,cx[0]+I(0.1),ry+I(0.12),cw[0]-I(0.15),I(0.65),lbl,size=11,bold=True,color=NAVY)
        s=scores[ri%len(scores)]
        for ci2,(sc2,x3,w3) in enumerate(zip(s,cx[1:],cw[1:])):
            cc=sc.get(sc2,GRAY)
            _rect(slide,x3+I(0.85),ry+I(0.2),I(0.48),I(0.48),cc)
            _txt(slide,x3+I(0.87),ry+I(0.17),I(0.48),I(0.52),sc2,size=12,bold=True,color=WHITE,align=PP_ALIGN.CENTER)

# ── SLIDE 4: プロセス（カラー縦カード＋矢印）──
def design_process(slide, prs, headline, body):
    W,H=int(prs.slide_width),int(prs.slide_height)
    _rect(slide,0,0,W,H,WHITE)
    _rect(slide,0,0,W,I(1.0),TEAL)
    _txt(slide,I(0.5),I(0.12),I(9.8),I(0.75),headline,size=19,bold=True,color=WHITE)
    steps=body[:5]; n=len(steps)
    sw=int((W-I(1.0))/n)
    ac=[BLUE,TEAL,NAVY,CORAL,GREEN]
    for i,step in enumerate(steps):
        sx=I(0.5)+i*sw; sy=I(1.3); shw=sw-I(0.18)
        lbl=step.split('：')[0] if '：' in step else step[:12]
        det=step.split('：')[1] if '：' in step else step
        c=ac[i%len(ac)]
        _rect(slide,sx,sy,shw,I(0.55),c)
        _rect(slide,sx,sy+I(0.55),shw,H-sy-I(0.55)-I(0.3),ICE)
        _txt(slide,sx+I(0.1),sy+I(0.06),shw-I(0.2),I(0.42),f'0{i+1}',size=22,bold=True,color=WHITE)
        _txt(slide,sx+I(0.12),sy+I(0.65),shw-I(0.2),I(0.55),lbl,size=12,bold=True,color=c)
        _txt(slide,sx+I(0.12),sy+I(1.3),shw-I(0.2),H-sy-I(1.6),det,size=10,color=GRAY)
        if i<n-1:
            ax=sx+shw+I(0.01); ay=sy+I(1.5)
            _rect(slide,ax,ay,I(0.15),I(0.15),c)

# ── SLIDE 5: まとめ（ダーク背景＋番号リスト）──
def design_summary(slide, prs, headline, body):
    W,H=int(prs.slide_width),int(prs.slide_height)
    _rect(slide,0,0,W,H,NAVY)
    _rect(slide,0,0,I(0.1),H,GOLD)
    _txt(slide,I(0.4),I(0.3),I(9.5),I(0.9),headline,size=24,bold=True,color=WHITE)
    _rect(slide,I(0.4),I(1.3),I(9.3),I(0.04),GOLD)
    ic=[GOLD,TEAL,CORAL,LGRAY,WHITE]
    for i,b in enumerate(body[:5]):
        by2=I(1.5)+i*I(1.05); c=ic[i%len(ic)]
        _rect(slide,I(0.4),by2+I(0.07),I(0.5),I(0.5),RGBColor(0x2A,0x35,0x7A))
        _txt(slide,I(0.4),by2+I(0.04),I(0.5),I(0.52),str(i+1),size=14,bold=True,color=c,align=PP_ALIGN.CENTER)
        lbl=b.split('：')[0] if '：' in b else b[:15]
        det=b.split('：')[1] if '：' in b else b
        _txt(slide,I(1.05),by2,I(8.6),I(0.4),lbl,size=13,bold=True,color=c)
        _txt(slide,I(1.05),by2+I(0.42),I(8.6),I(0.55),det,size=11,color=LGRAY)

DESIGN_FNS=[design_title,design_data,design_comparison,design_process,design_summary]
CT_MAP={'データ':design_data,'比較':design_comparison,'プロセス':design_process,'関係性':design_process}

def get_design_fn(purpose,content_type,page,total):
    if page==1: return design_title
    if page==total: return design_summary
    fn=CT_MAP.get(content_type)
    if fn: return fn
    # ローテーション
    return DESIGN_FNS[1+((page-2)%3)]
