# -*- coding: utf-8 -*-
"""
robot_pptx_core.py
ロボット甲子園 アイデアシート → プレゼンPowerPoint 生成コア

SalesNexus AI（旧 名刺管理）の research_core.py に相当する共有ロジック。
Streamlit 非依存。app から import して使う。

ワークフロー:
  1) ocr_idea_sheet()        手書きアイデアシートを Gemini 2.5 Flash で構造化
  2) generate_presentation() OCR結果を Ollama gpt-oss:120b-cloud でスライド原稿化
  3) build_pptx()            python-pptx で白背景・非ビジネス調のスライド生成

依存: google-generativeai, ollama, python-pptx, pymupdf(任意), Pillow
"""

import io
import os
import json
import re
import hashlib
from datetime import datetime

# ----------------------------------------------------------------------------
# 定数
# ----------------------------------------------------------------------------
# コスト/濫用対策のサーバー側上限（クライアントからは変更不可）
MAX_INPUT_CHARS = 8000        # 生成に渡す入力テキストの最大文字数
MAX_OUTPUT_TOKENS = 2048      # 生成の最大出力トークン

OCR_MODEL = "gemini-2.5-flash"            # 手書きOCR（クラウド・高精度・安価）
OCR_VISION_MODEL = "qwen3-vl:235b-cloud"  # 手書きOCR代替（Ollamaビジョン。ローカルは qwen3-vl:8b）
GEN_MODEL_DEFAULT = "gpt-oss:20b-cloud"    # 原稿生成（軽量級・負荷レベル1。重いなら120bに変更可）

# 配色（学生向け・元気だが白背景を生かしたクリーンな配色）
COL_INK = "1F2937"      # 本文（濃いグレー）
COL_MUTE = "6B7280"     # 補足
COL_BG_CARD = "F3F4F6"  # 中に置く薄いカード（※フチには触れさせない）
COL_A1 = "EA5A2D"       # アクセント1（オレンジ）
COL_A2 = "2563EB"       # アクセント2（ブルー）
COL_A3 = "16A34A"       # アクセント3（グリーン）
ACCENTS = [COL_A1, COL_A2, COL_A3]

# Windows標準の親しみやすい日本語フォント
FONT_HEAD = "Meiryo"
FONT_BODY = "Meiryo"


# ============================================================================
# 1. OCR（Gemini 2.5 Flash）
# ============================================================================
def _img_bytes_to_part(image_bytes, mime="image/png"):
    return {"mime_type": mime, "data": image_bytes}


OCR_PROMPT = """あなたは手書きアンケート/アイデアシートを読み取る専門OCRです。
画像は、ロボット系イベントで高校生・高専生が手書きしたロボットのアイデアシートです。
すべての手書き文字・図のキャプションを読み取り、次のJSONだけを出力してください（前置き・説明・```不要）。
読み取れない項目は空文字 "" にしてください。推測で創作しないこと。

{
  "robot_name": "ロボットの名称",
  "school": "学校名",
  "grade": "学年・組",
  "author": "氏名",
  "impression": "産業用ロボットを見た感想（あれば）",
  "usage_steps": ["使い方の箇条書きを1項目ずつ"],
  "shape_note": "形・大きさの説明（図のキャプション含む）",
  "motivation": "なぜこの作業をさせようと思ったか",
  "target_demand": "誰が使う・需要に関する記述",
  "price": "想定価格"
}
"""


def ocr_idea_sheet(image_bytes, gemini_api_key, mime="image/png"):
    """手書きアイデアシート画像を Gemini 2.5 Flash で構造化JSONに変換して dict を返す。"""
    import google.generativeai as genai
    genai.configure(api_key=gemini_api_key)
    model = genai.GenerativeModel(OCR_MODEL)
    resp = model.generate_content(
        [OCR_PROMPT, _img_bytes_to_part(image_bytes, mime)],
        generation_config={"temperature": 0.1},
    )
    return _safe_json(resp.text, default=_empty_idea())


def ocr_idea_sheet_ollama(image_bytes, model=OCR_VISION_MODEL, host=None):
    """手書きアイデアシート画像を Ollama のビジョンモデル(qwen3-vl)で構造化JSONに変換。
    Geminiキーが使えないときの代替経路。images にバイト列を渡す。"""
    client = _ollama_client(host)
    resp = client.chat(
        model=model,
        messages=[{"role": "user", "content": OCR_PROMPT, "images": [image_bytes]}],
        options={"temperature": 0.1},
    )
    return _safe_json(resp["message"]["content"], default=_empty_idea())


def run_ocr(image_bytes, *, engine="gemini", mime="image/png",
            gemini_api_key=None, vision_model=OCR_VISION_MODEL, host=None):
    """OCRエンジンのディスパッチャ。engine='gemini' または 'ollama'。"""
    if engine == "ollama":
        return ocr_idea_sheet_ollama(image_bytes, model=vision_model, host=host)
    return ocr_idea_sheet(image_bytes, gemini_api_key, mime=mime)


def _empty_idea():
    return {
        "robot_name": "", "school": "", "grade": "", "author": "",
        "impression": "", "usage_steps": [], "shape_note": "",
        "motivation": "", "target_demand": "", "price": "",
    }


# ============================================================================
# 2. スライド原稿生成（Ollama gpt-oss:120b-cloud）
# ============================================================================
def _ollama_client(host=None):
    """Ollamaクライアントを返す。
    OLLAMA_API_KEY があれば ollama.com に直接接続（ローカルOllama不要・クラウド/サーバー向け）。
    それ以外は OLLAMA_HOST かローカル(localhost:11434)を使用。"""
    import ollama
    api_key = os.environ.get("OLLAMA_API_KEY")
    host = host or os.environ.get("OLLAMA_HOST")
    if api_key:
        return ollama.Client(host=host or "https://ollama.com",
                             headers={"Authorization": "Bearer " + api_key})
    if host:
        return ollama.Client(host=host)
    return ollama
GEN_SYSTEM = """あなたは、ロボット甲子園に出場する高校生・高専生のアイデアを、
聴衆（同世代の学生・先生・審査員）にワクワクして伝えるプレゼン原稿に仕上げる編集者です。
かたいビジネス文書ではなく、前向きで分かりやすい、はずむような言葉を使ってください。
専門用語は噛み砕き、1行は短く。出力はJSONのみ。"""

GEN_PROMPT_TMPL = """次のアイデアシート情報から、8枚構成のプレゼン原稿をJSONで作ってください。
情報が空の項目は、シートの他の記述から自然に補ってよいが、事実をねつ造しないこと。
各ポイントは「短い見出し」と「その説明文」をセットにすること（見出しだけで終わらせない）。

# アイデアシート情報
{idea_json}

# 出力JSON（このスキーマ厳守・```やコメント禁止）
{{
  "title": "表紙の大見出し（ロボット名を主役に、短く力強く）",
  "subtitle": "表紙のサブコピー（ひとことで魅力を伝える）",
  "team_line": "学校名・学年・氏名をまとめた1行",
  "slides": [
    {{"key": "concept", "title": "このロボットってなに？", "lead": "1行で核心",
      "points": [{{"head": "短い見出し(15字以内)", "desc": "その内容を説明する1〜2文(40〜70字)"}}]}},
    {{"key": "why",     "title": "なんで作りたいの？",   "lead": "課題を一言で", "points": [{{"head":"","desc":""}}]}},
    {{"key": "how",     "title": "どうやって動く？",     "lead": "仕組みを一言で", "points": [{{"head":"","desc":""}}]}},
    {{"key": "shape",   "title": "形と大きさ",           "lead": "見た目のポイント", "points": [{{"head":"","desc":""}}]}},
    {{"key": "who",     "title": "だれの役に立つ？",     "lead": "使う人・場面", "points": [{{"head":"","desc":""}}]}},
    {{"key": "price",   "title": "おねだんと未来",       "lead": "価格の考え方", "points": [{{"head":"","desc":""}}]}}
  ]
}}
各スライドの points は3項目（how は4項目まで可）。見出しは体言中心で簡潔に、説明文は中高生にも伝わるやさしい言葉で。"""


def generate_presentation(idea_data, model=GEN_MODEL_DEFAULT, host=None):
    """OCR結果dict → スライド原稿dict。Ollama を使用。
    コスト爆発を防ぐためサーバー側で入力文字数と出力トークンに上限を設ける。"""
    client = _ollama_client(host)
    idea_json = json.dumps(idea_data, ensure_ascii=False, indent=2)
    idea_json = idea_json[:MAX_INPUT_CHARS]   # 入力サイズの上限（巨大入力でのコスト爆発を防止）
    prompt = GEN_PROMPT_TMPL.format(idea_json=idea_json)
    resp = client.chat(
        model=model,
        messages=[
            {"role": "system", "content": GEN_SYSTEM},
            {"role": "user", "content": prompt},
        ],
        options={"temperature": 0.6, "num_predict": MAX_OUTPUT_TOKENS},  # 出力トークン上限
    )
    text = resp["message"]["content"]
    return _safe_json(text, default=_fallback_content(idea_data))


def _fallback_content(idea):
    """AI生成が失敗した場合に、OCR結果から最低限の原稿を組み立てる。"""
    name = idea.get("robot_name") or "わたしたちのロボット"

    def pts(items):
        return [{"head": str(x), "desc": ""} for x in items if str(x).strip()]

    return {
        "title": name,
        "subtitle": "アイデアをカタチに",
        "team_line": " ".join(x for x in [idea.get("school", ""), idea.get("grade", ""), idea.get("author", "")] if x),
        "slides": [
            {"key": "concept", "title": "このロボットってなに？", "lead": idea.get("impression", "")[:30],
             "points": pts([name])},
            {"key": "why", "title": "なんで作りたいの？", "lead": "", "points": pts(_split(idea.get("motivation", "")))},
            {"key": "how", "title": "どうやって動く？", "lead": "",
             "points": pts(idea.get("usage_steps", []) or _split(idea.get("shape_note", "")))},
            {"key": "shape", "title": "形と大きさ", "lead": "", "points": pts(_split(idea.get("shape_note", "")))},
            {"key": "who", "title": "だれの役に立つ？", "lead": "", "points": pts(_split(idea.get("target_demand", "")))},
            {"key": "price", "title": "おねだんと未来", "lead": "",
             "points": pts([idea.get("price", "")] if idea.get("price") else [])},
        ],
    }


# ============================================================================
# 3. PowerPoint 生成（python-pptx）
# ============================================================================
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
# フチに触れさせない安全マージン（複合機が端まで印刷できないため）
SAFE = Inches(0.55)


def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _set_white_bg(slide):
    """スライド背景を必ず白に。"""
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = _rgb("FFFFFF")


def _no_line(shape):
    shape.line.fill.background()


def _txt(slide, l, t, w, h, lines, *, size, color, bold=False, font=FONT_BODY,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, leading=1.12, space_after=4):
    """テキストボックス追加。lines は str か (str, opts) のリスト。"""
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    if isinstance(lines, str):
        lines = [lines]
    for i, ln in enumerate(lines):
        opts = {}
        if isinstance(ln, tuple):
            ln, opts = ln
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = opts.get("align", align)
        p.line_spacing = leading
        p.space_after = Pt(opts.get("space_after", space_after))
        p.space_before = Pt(0)
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(opts.get("size", size))
        r.font.bold = opts.get("bold", bold)
        r.font.name = opts.get("font", font)
        r.font.color.rgb = _rgb(opts.get("color", color))
        _set_cjk_font(r, opts.get("font", font))
    return tb


def _set_cjk_font(run, font_name):
    """日本語が確実にそのフォントで出るよう east-asian 指定を入れる。"""
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", font_name)


def _badge(slide, l, t, d, text, color):
    """連番などを入れる色付き円バッジ。"""
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, d, d)
    c.fill.solid()
    c.fill.fore_color.rgb = _rgb(color)
    _no_line(c)
    tf = c.text_frame
    tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size = Pt(16); r.font.bold = True; r.font.name = FONT_HEAD
    r.font.color.rgb = _rgb("FFFFFF")
    _set_cjk_font(r, FONT_HEAD)
    return c


def _card(slide, l, t, w, h, fill=COL_BG_CARD):
    """角丸の薄いカード（中に置くだけ。端には触れさせない）。"""
    rr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    rr.fill.solid()
    rr.fill.fore_color.rgb = _rgb(fill)
    _no_line(rr)
    try:
        rr.adjustments[0] = 0.06
    except Exception:
        pass
    return rr


def _title(slide, text, accent):
    """各スライド共通のタイトル（アクセント色のドット＋見出し）。下線は引かない。"""
    dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, SAFE, Inches(0.62), Inches(0.20), Inches(0.20))
    dot.fill.solid(); dot.fill.fore_color.rgb = _rgb(accent); _no_line(dot)
    _txt(slide, SAFE + Inches(0.34), Inches(0.5), SLIDE_W - SAFE * 2 - Inches(0.34), Inches(0.7),
         text, size=30, color=COL_INK, bold=True, font=FONT_HEAD)


def _add_picture_fit(slide, img_bytes, l, t, max_w, max_h):
    """画像を枠内にアスペクト比維持で収め、中央寄せで配置。"""
    from PIL import Image
    bio = io.BytesIO(img_bytes)
    iw, ih = Image.open(bio).size
    ratio = min(max_w / iw, max_h / ih)
    w = int(iw * ratio); h = int(ih * ratio)
    l2 = l + (max_w - w) // 2
    t2 = t + (max_h - h) // 2
    slide.shapes.add_picture(io.BytesIO(img_bytes), l2, t2, width=w, height=h)


# 薄いアクセント色（白に混ぜたティント）
_TINTS = {COL_A1: "FCEDE6", COL_A2: "E7EEFD", COL_A3: "E6F4EC"}


def _norm_point(p):
    """points要素を (見出し, 説明) に正規化。文字列・dictどちらも許容。"""
    if isinstance(p, dict):
        return (str(p.get("head") or p.get("text") or "").strip(),
                str(p.get("desc") or "").strip())
    return (str(p).strip(), "")


def _point_card(slide, l, t, w, h, n, head, desc, accent):
    """番号バッジ＋見出し＋説明文の1枚カード（薄い背景・角丸）。"""
    _card(slide, l, t, w, h, fill="F5F6F8")
    bd = Inches(0.5)
    bx = l + Inches(0.26)
    if desc:
        by = t + Inches(0.24)
        head_anchor = MSO_ANCHOR.TOP
    else:
        by = t + (h - bd) // 2          # 説明が無い時は縦中央
        head_anchor = MSO_ANCHOR.MIDDLE
    _badge(slide, bx, by, bd, str(n), accent)
    tx = bx + bd + Inches(0.28)
    tw = l + w - tx - Inches(0.3)
    if desc:
        _txt(slide, tx, t + Inches(0.2), tw, Inches(0.42),
             head, size=17, color=COL_INK, bold=True, font=FONT_HEAD, leading=1.05)
        _txt(slide, tx, t + Inches(0.66), tw, h - Inches(0.78),
             desc, size=12.5, color=COL_MUTE, font=FONT_BODY, leading=1.2)
    else:
        _txt(slide, tx, t, tw, h, head, size=17, color=COL_INK, bold=True,
             font=FONT_HEAD, anchor=MSO_ANCHOR.MIDDLE, leading=1.1)


def _cards_column(slide, points, l, t, w, accent, bottom=Inches(7.0), max_n=4):
    """縦方向にカードを敷き詰める。"""
    pts = [_norm_point(p) for p in points if _norm_point(p)[0]][:max_n]
    if not pts:
        return
    n = len(pts)
    gap = Inches(0.22)
    avail = bottom - t
    ch = int((avail - gap * (n - 1)) / n)
    ch = max(min(ch, Inches(1.55)), Inches(0.9))
    for i, (head, desc) in enumerate(pts):
        cy = t + (ch + gap) * i
        _point_card(slide, l, cy, w, ch, i + 1, head, desc, accent)


def _render_split(slide, points, l_text_w, body_top, img_bytes, accent, max_n=3):
    """左テキスト / 右画像 のスプリット描画（どの本文スライドでも使える）。"""
    _cards_column(slide, points, SAFE, body_top, l_text_w, accent, max_n=max_n)
    pic_l = SAFE + l_text_w + Inches(0.35)
    pic_w = SLIDE_W - pic_l - SAFE
    pic_h = Inches(7.0) - body_top
    _card(slide, pic_l, body_top, pic_w, pic_h, fill=COL_BG_CARD)
    _add_picture_fit(slide, img_bytes, pic_l + Inches(0.2), body_top + Inches(0.2),
                     pic_w - Inches(0.4), pic_h - Inches(0.4))


# 画像の自動振り分け優先順（写真リストだけ渡された場合）
_AUTO_KEYS = ["shape", "concept", "how", "who", "why"]


def _photos_to_map(photos):
    m = {}
    for img, key in zip(photos, _AUTO_KEYS):
        m[key] = img
    return m


def build_pptx(content, photos=None, output_path="robot_presentation.pptx",
               price_value="", images_by_key=None):
    """
    content       : generate_presentation() の戻り dict
    images_by_key : {本文スライドkey: 画像bytes} の対応（優先）。
    photos        : 画像bytesのリスト（後方互換。指定優先順で各ページへ自動配分）。
    ※ ギャラリーページは作らない。割り当てられた画像は各ページにスプリット配置。
    """
    if images_by_key is None:
        images_by_key = _photos_to_map(photos or [])
    images_by_key = {k: v for k, v in images_by_key.items() if v}
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    # ---- 表紙（必ず白背景・ビジネス調を避ける） ----
    s = prs.slides.add_slide(blank)
    _set_white_bg(s)
    # 上部にカラフルな丸を3つ（端には触れない位置に）— 競技会らしい遊び心
    for i, col in enumerate(ACCENTS):
        d = Inches(0.5)
        c = s.shapes.add_shape(MSO_SHAPE.OVAL,
                               SAFE + Inches(0.0) + Inches(0.7) * i, Inches(1.05), d, d)
        c.fill.solid(); c.fill.fore_color.rgb = _rgb(col); _no_line(c)
    # カテゴリタグ（角丸・中置き）
    tag = _card(s, SAFE, Inches(1.9), Inches(3.4), Inches(0.5), fill=COL_INK)
    _txt(s, SAFE, Inches(1.97), Inches(3.4), Inches(0.36),
         "ロボット甲子園 アイデア提案", size=14, color="FFFFFF", bold=True,
         font=FONT_HEAD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # 大見出し
    _txt(s, SAFE, Inches(2.75), SLIDE_W - SAFE * 2, Inches(2.0),
         content.get("title", ""), size=54, color=COL_INK, bold=True, font=FONT_HEAD)
    # サブコピー（アクセント色）
    _txt(s, SAFE, Inches(4.7), SLIDE_W - SAFE * 2, Inches(0.8),
         content.get("subtitle", ""), size=22, color=COL_A1, bold=True, font=FONT_HEAD)
    # チーム情報
    _txt(s, SAFE, Inches(6.35), SLIDE_W - SAFE * 2, Inches(0.6),
         content.get("team_line", ""), size=16, color=COL_MUTE, font=FONT_BODY)

    # ---- 本文スライド ----
    slides = content.get("slides", [])
    for idx, sl in enumerate(slides):
        s = prs.slides.add_slide(blank)
        _set_white_bg(s)
        accent = ACCENTS[idx % len(ACCENTS)]
        _title(s, sl.get("title", ""), accent)

        key = sl.get("key", "")
        lead = sl.get("lead", "")
        points = sl.get("points") or []
        img = images_by_key.get(key)

        # リード文（共通）
        if lead:
            _txt(s, SAFE, Inches(1.55), SLIDE_W - SAFE * 2, Inches(0.6),
                 lead, size=19, color=accent, bold=True, font=FONT_HEAD)
            body_top = Inches(2.35)
        else:
            body_top = Inches(1.95)

        # ---- レイアウト分岐 ----
        if key == "price":
            # 左に金額パネル / 右に補足カード（価格ページは画像より金額を主役に）
            panel_w = Inches(4.55)
            panel_h = Inches(7.0) - body_top
            price_txt = price_value or _first_price(points)
            _card(s, SAFE, body_top, panel_w, panel_h, fill=_TINTS.get(accent, COL_BG_CARD))
            _txt(s, SAFE + Inches(0.4), body_top + Inches(0.5), panel_w - Inches(0.8), Inches(0.5),
                 "想定価格", size=15, color=accent, bold=True, font=FONT_HEAD)
            _txt(s, SAFE + Inches(0.4), body_top + Inches(1.15), panel_w - Inches(0.8), panel_h - Inches(1.5),
                 price_txt, size=30, color=COL_INK, bold=True, font=FONT_HEAD, leading=1.15)
            _cards_column(s, points, SAFE + panel_w + Inches(0.55), body_top,
                          SLIDE_W - (SAFE + panel_w + Inches(0.55)) - SAFE, accent, max_n=3)

        elif img:
            # 画像が割り当てられたページ：左テキスト / 右画像
            _render_split(s, points, Inches(5.85), body_top, img, accent, max_n=3)

        else:
            # 全幅カード（見出し＋説明）
            max_n = 4 if key == "how" else 3
            _cards_column(s, points, SAFE, body_top, SLIDE_W - SAFE * 2, accent, max_n=max_n)

    # ---- 裏表紙（白・シンプル） ----
    s = prs.slides.add_slide(blank)
    _set_white_bg(s)
    _txt(s, SAFE, Inches(3.0), SLIDE_W - SAFE * 2, Inches(1.0),
         "ありがとうございました！", size=40, color=COL_INK, bold=True,
         font=FONT_HEAD, align=PP_ALIGN.CENTER)
    _txt(s, SAFE, Inches(4.1), SLIDE_W - SAFE * 2, Inches(0.6),
         content.get("team_line", ""), size=16, color=COL_MUTE,
         font=FONT_BODY, align=PP_ALIGN.CENTER)

    prs.save(output_path)
    return output_path


# ============================================================================
# 補助
# ============================================================================
def _safe_json(text, default=None):
    """LLM出力から最初のJSONオブジェクトを安全に取り出す。```除去にも対応。"""
    if not text:
        return default
    t = text.strip()
    t = re.sub(r"^```(?:json)?", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    try:
        return json.loads(t)
    except Exception:
        m = re.search(r"\{.*\}", t, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
    return default


def _first_price(points):
    """価格ポイント群から代表的な金額文字列を1つ取り出す。"""
    for p in points:
        head, desc = _norm_point(p)
        for t in (head, desc):
            if any(k in t for k in ("円", "万", "¥", "価格")):
                return t
    return _norm_point(points[0])[0] if points else ""


def _split(text):
    """説明文を箇条書きへ簡易分割。"""
    if not text:
        return []
    parts = re.split(r"[。\n・]", text)
    return [p.strip() for p in parts if p.strip()][:4]


def pdf_first_page_png(pdf_bytes, dpi=200):
    """PDFの1ページ目をPNG bytesにして返す（OCR入力用）。pymupdf必須。"""
    import fitz
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png")


# ============================================================================
# ベータ版PowerPointの取り込み
# ============================================================================
def pptx_extract_media_images(pptx_bytes, min_kb=12):
    """ベータ版PPTに埋め込まれた写真（jpg/png等）だけを取り出して bytes のリストで返す。
    スライド全体の画像化ではなく“中の写真”を拾うので、実機写真などの素材抽出に向く。
    min_kb 未満（アイコン等）は除外。"""
    import zipfile
    out = []
    seen = set()
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z:
        names = [n for n in z.namelist() if n.startswith("ppt/media/")]
        names.sort(key=_natural_key)
        for n in names:
            ext = os.path.splitext(n)[1].lower()
            if ext not in (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff"):
                continue
            data = z.read(n)
            if len(data) < min_kb * 1024:        # 小さすぎる画像（装飾・アイコン）は除外
                continue
            key = hashlib.md5(data).hexdigest()
            if key in seen:
                continue
            seen.add(key)
            out.append(data)
    return out


def pptx_extract_text(pptx_bytes):
    """各スライドのテキストを取り出して [{'slide':n, 'text':...}] を返す（python-pptxのみ・常に動作）。"""
    from pptx import Presentation as _P
    prs = _P(io.BytesIO(pptx_bytes))
    out = []
    for i, slide in enumerate(prs.slides, 1):
        chunks = []
        for sh in slide.shapes:
            if sh.has_text_frame:
                t = "\n".join(p.text for p in sh.text_frame.paragraphs if p.text.strip())
                if t.strip():
                    chunks.append(t.strip())
        out.append({"slide": i, "text": "\n".join(chunks)})
    return out


def _natural_key(name):
    m = re.search(r"(\d+)", os.path.basename(name))
    return int(m.group(1)) if m else 0


def pptx_to_slide_images(pptx_bytes):
    """ベータ版PPTの各スライドをPNG bytesのリストにして返す。
    Windowsの PowerPoint(COM) を優先し、無ければ LibreOffice(soffice) を使用。
    どちらも無ければ RuntimeError。"""
    import tempfile
    tmpdir = tempfile.mkdtemp(prefix="draftppt_")
    src = os.path.join(tmpdir, "draft.pptx")
    with open(src, "wb") as f:
        f.write(pptx_bytes)

    # 1) PowerPoint COM（Windows・PowerPointインストール済みなら最も確実）
    try:
        return _pptx_images_powerpoint(src, tmpdir)
    except Exception:
        pass
    # 2) LibreOffice 経由（soffice + pdftoppm が必要）
    try:
        return _pptx_images_soffice(src, tmpdir)
    except Exception as e:
        raise RuntimeError(
            "スライド画像化には PowerPoint(COM) または LibreOffice が必要です。"
            f"（詳細: {e}）画面キャプチャの貼り付けもご利用ください。")


def _pptx_images_powerpoint(src, tmpdir):
    import glob
    import comtypes.client
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    try:
        powerpoint.Visible = 1
    except Exception:
        pass
    pres = powerpoint.Presentations.Open(src, WithWindow=False)
    out_dir = os.path.join(tmpdir, "png")
    os.makedirs(out_dir, exist_ok=True)
    try:
        pres.Export(out_dir, "PNG")  # 全スライドをPNG出力（言語によりファイル名が異なる）
    finally:
        pres.Close()
        powerpoint.Quit()
    files = sorted(glob.glob(os.path.join(out_dir, "*.PNG")) +
                   glob.glob(os.path.join(out_dir, "*.png")), key=_natural_key)
    return [open(p, "rb").read() for p in files]


def _pptx_images_soffice(src, tmpdir):
    import glob
    import subprocess
    import shutil
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("soffice not found")
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, src],
                   check=True, timeout=120,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = os.path.join(tmpdir, "draft.pdf")
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        subprocess.run([pdftoppm, "-png", "-r", "150", pdf, os.path.join(tmpdir, "slide")],
                       check=True, timeout=120)
        files = sorted(glob.glob(os.path.join(tmpdir, "slide-*.png")), key=_natural_key)
        return [open(p, "rb").read() for p in files]
    # pdftoppm が無ければ pymupdf でPDF→PNG
    import fitz
    doc = fitz.open(pdf)
    return [doc[i].get_pixmap(dpi=150).tobytes("png") for i in range(len(doc))]

