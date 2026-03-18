"""
SlideAI Backend — FastAPI + python-pptx
Kimuraさんの手順通りに実装:
  1. 構成設計 (Claude API)
  2. テンプレート選択 (メタデータ検索)
  3. 空スライドを枚数分作成
  4. テンプレートスライドをコピーして適用
  5. テキスト流し込み
  6. 視覚QA (Claude Vision)
  7. 修正ループ
"""

import os, io, json, base64, re, copy, tempfile, shutil, logging
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import anthropic

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("slideai")

BASE_DIR = Path(__file__).parent

# ── テンプレートファイルマップ ──
TEMPLATE_FILES = {
    "DTC_PowerLibrary_2013":  BASE_DIR / "DTC.pptx",
    "McKinsey":               BASE_DIR / "McKinsey.pptx",
    "BCG":                    BASE_DIR / "BCG.pptx",
    "PowerLibrary_2009":      BASE_DIR / "PowerLibrary2009.pptx",
}

# ── メタデータ読み込み ──
with open(BASE_DIR / "metadata.json", encoding="utf-8") as f:
    ALL_METADATA = json.load(f)

app = FastAPI(title="SlideAI API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ════════════════════════════════════════
# リクエスト/レスポンス型
# ════════════════════════════════════════
class GenerateRequest(BaseModel):
    topic: str
    page_count: int = 5
    notes: str = ""
    api_key: str

class SlideSpec(BaseModel):
    page: int
    title: str
    purpose: str
    content_type: str
    key_points: list[str]
    visual_hint: str = ""
    headline: str = ""
    body: list[str] = []
    data_note: str = ""

# ════════════════════════════════════════
# Claude API ヘルパー
# ════════════════════════════════════════
def call_claude(api_key: str, prompt: str, system: str = "", max_tokens: int = 2000) -> str:
    client = anthropic.Anthropic(api_key=api_key)
    kwargs = dict(
        model="claude-sonnet-4-20250514",
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}]
    )
    if system:
        kwargs["system"] = system
    msg = client.messages.create(**kwargs)
    return msg.content[0].text

def call_claude_vision(api_key: str, prompt: str, image_b64: str, media_type: str = "image/jpeg") -> str:
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": image_b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    )
    return msg.content[0].text

def parse_json(text: str):
    cleaned = re.sub(r"```json\s*", "", text)
    cleaned = re.sub(r"```\s*", "", cleaned).strip()
    return json.loads(cleaned)

# ════════════════════════════════════════
# Step 1: 構成設計
# ════════════════════════════════════════
def design_structure(api_key: str, topic: str, page_count: int, notes: str) -> list[dict]:
    prompt = f"""プレゼンテーション構成を設計してください。

テーマ: {topic}
ページ数: {page_count}
追加指示: {notes or 'なし'}

要件:
- 1ページ目は必ずタイトルスライド
- 最終ページはまとめ/結論スライド
- 各スライドのkey_pointsは5〜8個、各ポイントは30文字以上の具体的な内容
- 数値・事例・具体例を必ず含める
- content_typeは テキスト/データ/比較/プロセス/関係性/組織 から選ぶ

JSONのみ返答（コードブロック不要）:
[{{"page":1,"title":"スライドタイトル","purpose":"目的","content_type":"テキスト","key_points":["ポイント1（30文字以上）","ポイント2","ポイント3","ポイント4","ポイント5"],"visual_hint":"グラフ|テーブル|フロー|箇条書き|マトリクス"}}]"""

    raw = call_claude(api_key, prompt, max_tokens=2000)
    return parse_json(raw)

# ════════════════════════════════════════
# Step 2: テンプレート選択
# ════════════════════════════════════════
TYPE_KEYWORD_MAP = {
    "タイトル": "テキスト", "表紙": "テキスト", "概要": "テキスト",
    "課題": "テキスト", "背景": "テキスト", "現状": "テキスト",
    "まとめ": "テキスト", "結論": "テキスト", "サマリー": "テキスト",
    "市場": "データ", "規模": "データ", "売上": "データ",
    "推移": "データ", "数値": "データ", "KPI": "データ",
    "比較": "比較", "競合": "比較", "マトリクス": "比較", "評価": "比較",
    "プロセス": "プロセス", "フロー": "プロセス", "ステップ": "プロセス",
    "ロードマップ": "プロセス", "スケジュール": "プロセス", "ガント": "プロセス",
    "組織": "組織", "チーム": "組織", "体制": "組織",
    "関係": "関係性", "構造": "関係性", "モデル": "関係性",
}

def select_template(purpose: str, content_type: str) -> Optional[dict]:
    """メタデータからベストマッチのテンプレートスライドを返す"""
    combined = (purpose + " " + content_type).lower()

    # content_typeで絞り込み
    candidates = [m for m in ALL_METADATA if m.get("content_type") == content_type]

    # なければ全体から
    if not candidates:
        candidates = ALL_METADATA

    # キーワードスコアリング
    def score(m):
        text = " ".join([
            m.get("title", ""),
            m.get("category", ""),
            m.get("use_case", ""),
            " ".join(m.get("keywords", []))
        ]).lower()
        return sum(1 for w in combined.split() if len(w) > 1 and w in text)

    ranked = sorted(candidates, key=score, reverse=True)
    return ranked[0] if ranked else None

# ════════════════════════════════════════
# Step 3: テンプレートスライドをコピー
# ════════════════════════════════════════
def copy_slide_from_template(src_prs: Presentation, slide_idx: int, dst_prs: Presentation) -> any:
    """テンプレートのスライドをdst_prsにコピーして返す"""
    from pptx.oxml.ns import qn
    import lxml.etree as etree

    src_slide = src_prs.slides[slide_idx]

    # スライドレイアウトを追加（blank使用）
    blank_layout = dst_prs.slide_layouts[6]  # Blank layout
    new_slide = dst_prs.slides.add_slide(blank_layout)

    # XML要素をコピー（図形・テキストボックスなど）
    src_sp_tree = src_slide.shapes._spTree
    dst_sp_tree = new_slide.shapes._spTree

    # 既存のプレースホルダーをクリア
    for elem in dst_sp_tree.findall(qn("p:sp")):
        dst_sp_tree.remove(elem)

    # ソーススライドの全要素をコピー
    for elem in src_sp_tree:
        tag = elem.tag
        # spTree自体の属性系は除く
        if tag in [qn("p:sp"), qn("p:grpSp"), qn("p:cxnSp"), qn("p:pic"), qn("p:graphicFrame")]:
            dst_sp_tree.append(copy.deepcopy(elem))

    # 背景をコピー
    if src_slide.background.fill.type is not None:
        bg = new_slide.background
        bg.fill.solid()
        try:
            src_fill = src_slide.background.fill
            if src_fill.fore_color.type is not None:
                bg.fill.fore_color.rgb = src_fill.fore_color.rgb
        except Exception:
            pass

    return new_slide

# ════════════════════════════════════════
# Step 4 & 5: テキスト流し込み
# ════════════════════════════════════════
def inject_content(slide, spec: dict) -> None:
    """スライドのテキストボックスにコンテンツを流し込む"""
    headline = spec.get("headline") or spec.get("title", "")
    body     = spec.get("body") or spec.get("key_points", [])

    shapes = list(slide.shapes)
    text_shapes = [s for s in shapes if s.has_text_frame]

    if not text_shapes:
        return

    # ソート: 上→下、左→右
    text_shapes.sort(key=lambda s: (s.top, s.left))

    # 最初のテキストボックス → タイトル/ヘッドライン
    if text_shapes:
        _set_text(text_shapes[0], headline, bold=True)

    # 残りのテキストボックス → ボディ
    body_shapes = text_shapes[1:]
    if not body_shapes:
        return

    if len(body_shapes) == 1:
        # 1つのテキストボックスに全部入れる
        _set_text_list(body_shapes[0], body)
    else:
        # 複数テキストボックスに分散
        per_box = max(1, len(body) // len(body_shapes))
        for i, shape in enumerate(body_shapes):
            chunk_start = i * per_box
            chunk_end   = chunk_start + per_box if i < len(body_shapes) - 1 else len(body)
            chunk = body[chunk_start:chunk_end]
            if chunk:
                _set_text_list(shape, chunk)

def _set_text(shape, text: str, bold: bool = False) -> None:
    """テキストフレームにテキストをセット（既存スタイルを維持）"""
    try:
        tf = shape.text_frame
        tf.word_wrap = True
        if tf.paragraphs:
            para = tf.paragraphs[0]
            if para.runs:
                run = para.runs[0]
                run.text = text
                if bold:
                    run.font.bold = True
            else:
                run = para.add_run()
                run.text = text
                if bold:
                    run.font.bold = True
            # 残りの段落をクリア
            for p in tf.paragraphs[1:]:
                for r in p.runs:
                    r.text = ""
        else:
            para = tf.add_paragraph()
            run = para.add_run()
            run.text = text
    except Exception as e:
        log.warning(f"_set_text error: {e}")

def _set_text_list(shape, items: list[str]) -> None:
    """テキストフレームにリストをセット"""
    try:
        tf = shape.text_frame
        tf.word_wrap = True
        # 既存段落をクリア
        for para in tf.paragraphs:
            for run in para.runs:
                run.text = ""
        # 最初の段落を使い回し、足りなければ追加
        for i, item in enumerate(items):
            if i < len(tf.paragraphs):
                para = tf.paragraphs[i]
                if para.runs:
                    para.runs[0].text = item
                else:
                    para.add_run().text = item
            else:
                new_para = tf.add_paragraph()
                new_para.add_run().text = item
    except Exception as e:
        log.warning(f"_set_text_list error: {e}")

# ════════════════════════════════════════
# Step 5b: テンプレートなしスライドを新規作成
# ════════════════════════════════════════
def create_slide_from_scratch(dst_prs: Presentation, spec: dict) -> any:
    """テンプレートが使えない場合、python-pptxでスライドを新規作成"""
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor

    slide_layout = dst_prs.slide_layouts[6]  # Blank
    slide = dst_prs.slides.add_slide(slide_layout)

    W = int(dst_prs.slide_width)
    H = int(dst_prs.slide_height)
    MX = int(Inches(0.45))
    MY = int(Inches(0.35))
    TH = int(Inches(0.75))
    CY = int(MY + TH + Inches(0.15))
    CW = int(W - MX * 2)
    CH = int(H - CY - Inches(0.25))

    headline = spec.get("headline") or spec.get("title", "")
    body     = spec.get("body") or spec.get("key_points", [])
    ct       = spec.get("content_type", "テキスト")

    NAVY  = RGBColor(0x1E, 0x3A, 0x5F)
    BLUE  = RGBColor(0x25, 0x63, 0xB0)
    LGRAY = RGBColor(0xCA, 0xCE, 0xD8)
    LIGHT = RGBColor(0xEE, 0xF3, 0xFB)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    GRAY  = RGBColor(0x5A, 0x64, 0x78)

    def add_rect(x, y, w, h, color):
        shape = slide.shapes.add_shape(1, x, y, w, h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.fill.background()
        return shape

    def add_textbox(x, y, w, h, text, size=12, bold=False, color=None, wrap=True, align=PP_ALIGN.LEFT):
        txb = slide.shapes.add_textbox(x, y, w, h)
        tf  = txb.text_frame
        tf.word_wrap = wrap
        tf.auto_size = None
        para = tf.paragraphs[0]
        para.alignment = align
        run = para.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color or NAVY
        return txb

    # タイトル（全スライド共通）
    add_textbox(MX, MY, CW, TH, headline, size=22, bold=True, color=NAVY)
    add_rect(MX, int(MY + TH), CW, 25000, LGRAY)

    # ヘルパー: Inches を int に変換
    def I(n): return int(Inches(n))

    if ct == "プロセス":
        steps = body[:5]
        n = len(steps) or 3
        step_w = int((CW - I(0.18) * (n-1)) / n)
        accent_colors = [BLUE, RGBColor(0x1A,0x7A,0x52), RGBColor(0xD4,0x88,0x0A), RGBColor(0x4B,0x8E,0xE8), RGBColor(0xC0,0x39,0x2B)]
        for i, step in enumerate(steps):
            sx = int(MX + i * (step_w + I(0.18)))
            ac = accent_colors[i % len(accent_colors)]
            add_rect(sx, CY, step_w, CH, LIGHT)
            add_rect(sx, CY, step_w, 70000, ac)
            add_textbox(sx + I(0.15), CY + I(0.15), step_w - I(0.3), I(0.5),
                        f"0{i+1}", size=26, bold=True, color=ac)
            sp = step.find("：")
            head = step[:sp] if sp > 0 else step[:min(len(step), 18)]
            detail = step[sp+1:] if sp > 0 else step[18:]
            add_textbox(sx + I(0.15), CY + I(0.8), step_w - I(0.3), I(0.65),
                        head, size=12, bold=True, color=NAVY)
            if detail:
                add_textbox(sx + I(0.15), CY + I(1.55), step_w - I(0.3), max(I(0.3), CH - I(1.8)),
                            detail, size=10, color=GRAY)

    elif ct == "比較":
        cols = min(len(body), 4)
        label_w = I(1.5)
        cell_w  = int((CW - label_w) / max(cols, 1))
        rows = ["自社", "競合A", "競合B"]
        row_h = int(CH / (len(rows) + 1))
        # ヘッダー
        add_rect(MX, CY, CW, row_h, NAVY)
        add_textbox(MX, CY, label_w, row_h, "評価軸", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        for ci, label in enumerate(body[:cols]):
            add_textbox(MX + label_w + ci * cell_w, CY, cell_w, row_h,
                        label, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # 行
        eval_data = [
            ["◎ 特化・高品質", "◎ 競争力高い", "◎ 独自差別化", "◎ ブランド力"],
            ["△ 差別化薄い", "○ 標準品質", "△ 革新性低い", "○ 認知度高い"],
            ["○ 低価格訴求", "△ 品質ばらつき", "○ 立地優位", "△ サービス浅い"],
        ]
        for ri, row_label in enumerate(rows):
            ry = int(CY + (ri + 1) * row_h)
            fill = LIGHT if ri % 2 == 0 else WHITE
            add_rect(MX, ry, CW, row_h, fill)
            add_rect(MX, ry, label_w, row_h, BLUE if ri == 0 else NAVY)
            add_textbox(MX, ry, label_w, row_h, row_label, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            for ci in range(cols):
                eval_text = eval_data[ri][ci] if ci < len(eval_data[ri]) else "－"
                add_textbox(int(MX + label_w + ci * cell_w + I(0.08)), int(ry + I(0.06)),
                            int(cell_w - I(0.16)), int(row_h - I(0.12)),
                            eval_text, size=10, color=NAVY if ri == 0 else GRAY, align=PP_ALIGN.CENTER)

    else:
        # デフォルト: 2列テキスト
        pts  = body[:8]
        half = (len(pts) + 1) // 2
        col_w = int((CW - I(0.2)) / 2)
        item_h = int(min(CH / max(half, 1) - I(0.08), I(1.4)))
        for col_idx, group in enumerate([pts[:half], pts[half:]]):
            gx = int(MX + col_idx * (col_w + I(0.2)))
            for i, pt in enumerate(group):
                iy = int(CY + i * (item_h + I(0.08)))
                ac = [BLUE, RGBColor(0x1A,0x7A,0x52), RGBColor(0xD4,0x88,0x0A), RGBColor(0x4B,0x8E,0xE8)][i % 4]
                add_rect(gx, iy, col_w, item_h, LIGHT)
                add_rect(gx, iy, I(0.05), item_h, ac)
                sp = pt.find("：")
                if sp > 0:
                    add_textbox(gx + I(0.12), iy + I(0.06), col_w - I(0.2), I(0.32),
                                pt[:sp], size=11, bold=True, color=NAVY)
                    add_textbox(gx + I(0.12), iy + I(0.40), col_w - I(0.2), max(I(0.2), item_h - I(0.5)),
                                pt[sp+1:], size=10, color=GRAY)
                else:
                    add_textbox(gx + I(0.12), iy + I(0.08), col_w - I(0.2), max(I(0.2), item_h - I(0.16)),
                                pt, size=11, color=RGBColor(0x1A, 0x2A, 0x44))

    return slide

# ════════════════════════════════════════
# Step 6: 視覚QA (Claude Vision)
# ════════════════════════════════════════
def qa_slides_with_vision(api_key: str, prs: Presentation, tmp_dir: Path) -> list[dict]:
    """各スライドをPNG化してClaude Visionでチェック"""
    issues = []

    # LibreOfficeでPDF変換
    pptx_path = tmp_dir / "check.pptx"
    pdf_path  = tmp_dir / "check.pdf"
    prs.save(str(pptx_path))

    ret = os.system(f"libreoffice --headless --convert-to pdf '{pptx_path}' --outdir '{tmp_dir}' 2>/dev/null")
    if ret != 0 or not pdf_path.exists():
        log.warning("PDF conversion failed, skipping QA")
        return []

    # PDFを画像に変換
    ret = os.system(f"pdftoppm -jpeg -r 100 '{pdf_path}' '{tmp_dir}/slide' 2>/dev/null")
    if ret != 0:
        log.warning("pdftoppm failed, skipping QA")
        return []

    slide_images = sorted(tmp_dir.glob("slide-*.jpg"))
    if not slide_images:
        slide_images = sorted(tmp_dir.glob("slide*.jpg"))

    for i, img_path in enumerate(slide_images):
        with open(img_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        prompt = """このスライドを視覚的に確認してください。
以下の問題があれば報告してください（なければ「問題なし」と返答）:
- テキストがボックスからはみ出している
- テキストが非常に小さくて読めない
- 要素が重なっている
- 空のプレースホルダーが残っている
- 明らかなレイアウト崩れ

問題があれば「問題あり: [内容]」、なければ「問題なし」のみ返答。"""

        try:
            result = call_claude_vision(api_key, prompt, img_b64)
            if "問題あり" in result:
                issues.append({"slide_index": i, "issue": result})
                log.info(f"Slide {i+1} QA issue: {result}")
        except Exception as e:
            log.warning(f"Vision QA failed for slide {i+1}: {e}")

    return issues

# ════════════════════════════════════════
# Step 7: コンテンツ生成
# ════════════════════════════════════════
def generate_content(api_key: str, plan: list[dict], topic: str, notes: str) -> list[dict]:
    prompt = f"""各スライドの詳細コンテンツを生成してください。

テーマ: {topic}
追加指示: {notes or 'なし'}

構成:
{chr(10).join(f"Page{s['page']}: {s['title']}（{s['purpose']}）" for s in plan)}

条件:
- headlineは20〜45文字のインパクトある1文（数値・体言止めを使う）
- bodyは6〜8個のポイント、「見出し：詳細説明」形式（各40文字以上）
- 具体的な数値・事例・根拠を必ず含める
- data_noteはデータスライドのみ「65,78,85,79,92,98」形式で数値を記載

JSONのみ返答（コードブロック不要）:
[{{"page":1,"headline":"メインメッセージ","body":["見出し1：詳細説明40文字以上","見出し2：...","見出し3：...","見出し4：...","見出し5：...","見出し6：..."],"data_note":""}}]"""

    raw = call_claude(api_key, prompt, max_tokens=3000)
    return parse_json(raw)

# ════════════════════════════════════════
# メイン生成エンドポイント
# ════════════════════════════════════════
@app.post("/generate")
async def generate(req: GenerateRequest):
    log.info(f"Generate request: topic='{req.topic}', pages={req.page_count}")
    tmp_dir = Path(tempfile.mkdtemp())

    try:
        # Step 1: 構成設計
        log.info("Step 1: Designing structure...")
        plan = design_structure(req.api_key, req.topic, req.page_count, req.notes)
        log.info(f"  → {len(plan)} slides planned")

        # Step 3: コンテンツ生成
        log.info("Step 3: Generating content...")
        contents = generate_content(req.api_key, plan, req.topic, req.notes)
        content_map = {c["page"]: c for c in contents}

        # 出力pptxを作成（スライドサイズはDTCテンプレートに合わせる）
        ref_prs = Presentation(str(TEMPLATE_FILES["DTC_PowerLibrary_2013"]))
        dst_prs = Presentation()
        dst_prs.slide_width  = ref_prs.slide_width   # 10.83"
        dst_prs.slide_height = ref_prs.slide_height  # 7.5"

        # テンプレートファイルをキャッシュ
        loaded_templates = {}

        slide_info = []

        for slide_spec in plan:
            page = slide_spec["page"]
            ct   = slide_spec.get("content_type", "テキスト")
            spec = {**slide_spec, **(content_map.get(page, {}))}

            # Step 2: テンプレート選択
            log.info(f"Step 2: Selecting template for page {page} ({ct})...")
            tmpl_meta = select_template(
                slide_spec.get("purpose", "") + " " + slide_spec.get("visual_hint", ""),
                ct
            )

            used_template = False

            if tmpl_meta:
                src_name = tmpl_meta.get("source", "")
                slide_num = tmpl_meta.get("slide_num", 1) - 1  # 0-indexed
                tmpl_file = TEMPLATE_FILES.get(src_name)

                if tmpl_file and tmpl_file.exists():
                    # テンプレートファイルをロード（キャッシュ）
                    if src_name not in loaded_templates:
                        loaded_templates[src_name] = Presentation(str(tmpl_file))
                    src_prs = loaded_templates[src_name]

                    if 0 <= slide_num < len(src_prs.slides):
                        log.info(f"  → Using {src_name} slide {slide_num+1}: {tmpl_meta.get('title','')}")
                        # Step 3: テンプレートスライドをコピー
                        new_slide = copy_slide_from_template(src_prs, slide_num, dst_prs)
                        # Step 4: テキスト流し込み
                        inject_content(new_slide, spec)
                        used_template = True
                        slide_info.append({
                            "page": page,
                            "template_source": src_name,
                            "template_slide": slide_num + 1,
                            "template_title": tmpl_meta.get("title", ""),
                            "used_template": True
                        })

            if not used_template:
                # Step 4: テンプレートなし → 新規作成
                log.info(f"  → Creating from scratch for page {page}")
                create_slide_from_scratch(dst_prs, spec)
                slide_info.append({
                    "page": page,
                    "template_source": None,
                    "used_template": False
                })

        # Step 5: 一時保存
        pptx_path = tmp_dir / "output.pptx"
        dst_prs.save(str(pptx_path))

        # Step 6: 視覚QA
        log.info("Step 6: Visual QA...")
        qa_issues = qa_slides_with_vision(req.api_key, dst_prs, tmp_dir)
        if qa_issues:
            log.info(f"  → {len(qa_issues)} issues found (logged, auto-fix not yet implemented)")

        # Step 7: 完了 → バイナリ返却
        with open(pptx_path, "rb") as f:
            pptx_bytes = f.read()

        fname = f"{req.topic[:20]}_{__import__('datetime').date.today()}.pptx"
        return StreamingResponse(
            io.BytesIO(pptx_bytes),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f'attachment; filename="{fname}"',
                "X-Slide-Count": str(len(plan)),
                "X-QA-Issues": str(len(qa_issues)),
            }
        )

    except Exception as e:
        log.error(f"Generation failed: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.get("/health")
def health():
    return {"status": "ok", "templates": {k: v.exists() for k, v in TEMPLATE_FILES.items()}}

@app.get("/")
def root():
    return {"message": "SlideAI API", "docs": "/docs"}
