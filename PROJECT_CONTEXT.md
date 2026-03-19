# SlideAI バックエンド — 開発コンテキスト引き継ぎ

## プロジェクト概要
テーマ入力 → AI構成設計 → スライド生成 → pptxダウンロードの全自動化システム

## 本番URL
- **API**: https://web-production-8895a.up.railway.app
- **GitHub**: https://github.com/shin-bike/slideai-backend
- **フロントエンド**: `frontend.html`（ローカルで開く）

---

## システム構成

### バックエンド（Railway）
```
slideai_backend/
├── main.py              # FastAPI バックエンド（メイン）
├── slide_designs.py     # スライドデザインシステム v7
├── requirements.txt
├── Procfile / railway.json
├── metadata.json        # テンプレートメタデータ（未使用）
└── DTC.pptx / BCG.pptx / McKinsey.pptx / PowerLibrary2009.pptx
```

### 処理フロー
1. `design_structure()` → Claude API でスライド構成を設計（JSON）
2. `generate_content()` → Claude API でコンテンツ生成（**5ページずつ分割**）
3. `create_slide_from_scratch()` → `slide_designs.py` の関数でpptx生成
4. pptxバイナリを返却

---

## slide_designs.py v7（現在の最新）

### デザインコンセプト
- **マネーフォワード資料スタイル**を再現
- 白背景・ネイビー/ブルー系カラー + オレンジアクセント
- **シャドウなし**（XMLレベルで除去）
- **塗りつぶしは最小限**（ネイビーはヘッダー帯のみ、基本は白地+枠線）
- **内容に合わせて動的に図を描画**（項目数・ステップ数に自動対応）

### 7種類のデザイン関数

| 関数 | レイアウト | 動的対応 |
|------|-----------|---------|
| `slide_title` | ネイビーブロック＋クライアント名 | - |
| `slide_treemap` | ツリーマップ（面積＝重要度）＋KPIカード | セグメント数2〜6に対応 |
| `slide_flow` | 上段:プロセスフロー矢印図 / 下段:2カラム詳細 | 3〜8ステップに対応、幅・フォント自動調整 |
| `slide_twocol` | 背景/目的の2カラム比較 | 項目数に応じて行高動的計算 |
| `slide_table` | ◎○△スコアマトリクス | 行数3〜8に対応、フォント自動調整 |
| `slide_gantt` | ガントチャート＋積み上げ棒グラフ | タスク数3〜8に対応 |
| `slide_detail` | 論点リスト（オレンジ丸ラベル＋ブレット） | 4項目以下=1列、5〜8項目=2列に自動切替 |
| `slide_summary` | KPI3列＋アクションリスト | KPI数・アクション数に動的対応 |

### デザインルーティング（`get_design_fn`）
- `page==1` → `slide_title`（固定）
- `page==total` → `slide_summary`（固定）
- `content_type` / `purpose` キーワードから最適デザインを選択
- `used_history`（直近3枚）で同じデザインの連続を回避

### コンテンツタイプ → デザインマッピング
```python
'データ'   → treemap > detail > gantt
'比較'     → table   > twocol > detail
'プロセス' → flow    > gantt  > detail
'市場'     → treemap > detail > table
'競合'     → table   > twocol > detail
'計画'     → gantt   > flow   > detail
'収益'     → gantt   > detail > treemap
```

---

## main.py の重要な実装

### JSON生成の分割対応（v最新）
```python
def generate_content(...):
    CHUNK_SIZE = 5  # 5ページずつ分割
    chunks = [plan[i:i+CHUNK_SIZE] for i in range(0, len(plan), CHUNK_SIZE)]
    for chunk in chunks:
        tokens = max(2000, len(chunk) * 700 + 500)
        raw = call_claude(api_key, prompt, max_tokens=tokens)
        all_results.extend(parse_json(raw))
```

### parse_json（途中切れ修復機能付き）
```python
def parse_json(text):
    # 正常パース → 失敗時は最後の完全な}までで修復 → それも失敗ならエラー
```

### create_slide_from_scratch
```python
def create_slide_from_scratch(dst_prs, spec, used_history=None):
    from slide_designs import get_design_fn, DESIGN_NAMES
    fn = get_design_fn(purpose, ct, page, total, used_history or [])
    fn(slide, dst_prs, headline, body)
    if used_history is not None:
        used_history.append(DESIGN_NAMES.get(fn, 'detail'))
```

### generateエンドポイント
```python
used_history = []
for slide_spec in plan:
    spec = {**slide_spec, **(content_map.get(page, {})), "total_pages": len(plan)}
    create_slide_from_scratch(dst_prs, spec, used_history)
```

---

## カラーパレット（v7）
```python
NAVY   = RGBColor(0x00, 0x20, 0x60)  # 濃紺（ヘッダー帯）
BLUE1  = RGBColor(0x1F, 0x38, 0x64)  # 濃青
BLUE2  = RGBColor(0x2E, 0x74, 0xB5)  # 中青
BLUE3  = RGBColor(0x9D, 0xC3, 0xE6)  # 薄青
LBLUE  = RGBColor(0xBD, 0xD7, 0xEE)  # 水色（左ブロック）
ORANGE = RGBColor(0xE8, 0x77, 0x22)  # オレンジ（丸ラベル・アクセント）
```

---

## 共通フレーム（全スライド共通）
```
上部: 細いネイビー線
右上: "For Discussion" (イタリック)
下部: フッター線
左下: "SlideAI Copy Right Reserved"
```

---

## 解決済みの主要バグ
1. JSONが途中で切れる → 5ページ分割生成 + parse_json修復ロジック
2. テキストが重なる → top=Noneのプレースホルダー問題解決（テンプレートコピー廃止）
3. シャドウが出る → XML直接操作でeffectLstを除去
4. 同じデザインが連続 → used_historyで直近3枚を回避
5. ツリーマップ構造 → セグメント数に応じて動的レイアウト変更

---

## 改善要望・TODO
- [ ] フロー図の矢印をより本物に近く（ペンタゴン型等）
- [ ] ツリーマップの右列ラベルが見切れる場合の対応
- [ ] ガントの収益グラフで実際の数値を本文から抽出する精度向上
- [ ] 10〜15ページ以上でのデザインバリエーション確保
