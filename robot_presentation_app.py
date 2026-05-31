# -*- coding: utf-8 -*-
"""
robot_presentation_app.py
ロボット甲子園 アイデアシート → プレゼン自動生成アプリ（Streamlit）

ワークフロー:
  ① 手書きアイデアシート(PDF/画像/画面キャプチャ)を入力 → OCR
  ② 元画像を見ながら抽出内容を確認・修正
  ③ 写真(ファイル or 画面キャプチャ)を追加 …ベータ版PPTのスライド画像も取り込み可
  ④ Ollama で原稿生成 → PowerPoint(白表紙)をダウンロード

起動: start_robot.bat（ポート 8523）
"""

import io
import os
import json
import hashlib
import sqlite3
import streamlit as st
import robot_pptx_core as core

# 画面キャプチャ貼り付け（任意コンポーネント）
try:
    from streamlit_paste_button import paste_image_button
    HAS_PASTE = True
except Exception:
    HAS_PASTE = False

st.set_page_config(page_title="ロボット甲子園 プレゼン作成", page_icon="🤖", layout="wide")

SETTINGS_DB = "settings.db"
SETTINGS_FILE = "robot_settings.json"   # 旧形式（あれば自動移行）
DEFAULTS = {"gemini_api_key": "", "ocr_engine": "Gemini 2.5 Flash",
            "vision_model": core.OCR_VISION_MODEL,
            "gen_model": core.GEN_MODEL_DEFAULT, "ollama_host": ""}


# ---------------------------------------------------------------- 設定ストレージ（SQLite）
def _db():
    conn = sqlite3.connect(SETTINGS_DB)
    conn.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
    return conn


def load_settings():
    cfg = dict(DEFAULTS)
    conn = _db()
    rows = conn.execute("SELECT key, value FROM settings").fetchall()
    conn.close()
    if rows:
        for k, v in rows:
            cfg[k] = v
    elif os.path.exists(SETTINGS_FILE):     # 旧JSONからの移行（初回のみ）
        try:
            cfg.update(json.load(open(SETTINGS_FILE, encoding="utf-8")))
            save_settings(cfg)
        except Exception:
            pass
    # 環境変数があれば最優先で上書き（クラウドのSecrets運用）。
    # 環境変数由来のキーは UI に出さない（DOM/F12漏洩を防ぐ）ためフラグで記録。
    cfg["_gemini_from_env"] = bool(os.environ.get("GEMINI_API_KEY"))
    cfg["_ollama_key_from_env"] = bool(os.environ.get("OLLAMA_API_KEY"))
    if os.environ.get("GEMINI_API_KEY"):
        cfg["gemini_api_key"] = os.environ["GEMINI_API_KEY"]
    if os.environ.get("OLLAMA_HOST"):
        cfg["ollama_host"] = os.environ["OLLAMA_HOST"]
    return cfg


def save_settings(s):
    conn = _db()
    conn.executemany("INSERT INTO settings(key, value) VALUES(?, ?) "
                     "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                     [(k, "" if v is None else str(v)) for k, v in s.items()
                      if not k.startswith("_")])
    conn.commit()
    conn.close()


def pil_to_png(img):
    b = io.BytesIO()
    try:
        img.save(b, "PNG")
    except Exception:
        img.convert("RGB").save(b, "PNG")
    return b.getvalue()


def h(b):
    return hashlib.md5(b).hexdigest()


cfg = load_settings()

# ---------------------------------------------------------------- セッション初期化
st.session_state.setdefault("idea", None)
st.session_state.setdefault("sheet_img", None)     # OCR対象のPNG bytes
st.session_state.setdefault("sheet_mime", "image/png")
st.session_state.setdefault("pasted_photos", [])   # 貼り付けで追加した写真
st.session_state.setdefault("draft_slides", [])    # ベータ版PPTから取り込んだスライド画像
st.session_state.setdefault("draft_text", [])      # ベータ版PPTのテキスト

st.title("ロボット甲子園 プレゼン作成")
st.caption("アイデアシートを読み取り → 画像を見ながら修正 → 白表紙のPowerPointに出力")


# ---------------------------------------------------------------- アクセス制限（パスワード）
def _check_access():
    """APP_PASSWORD が設定されていれば、一致するまでアプリ本体を表示しない。
    公開URLの第三者による無断利用（API課金の浪費）を防ぐ。"""
    expected = os.environ.get("APP_PASSWORD")
    if not expected:
        return True  # 未設定なら制限なし（ローカル等）
    if st.session_state.get("_authed"):
        return True
    st.info("このアプリは関係者用です。合言葉を入力してください。")
    pw = st.text_input("合言葉", type="password", key="_pw")
    if pw:
        if pw == expected:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("合言葉が違います。")
    return False


if not _check_access():
    st.stop()

# ---------------------------------------------------------------- サイドバー
with st.sidebar:
    st.header("設定")
    cfg["ocr_engine"] = st.radio(
        "OCRエンジン（手書き）",
        ["Gemini 2.5 Flash", "Ollama (qwen3-vl)"],
        index=0 if cfg.get("ocr_engine", "Gemini 2.5 Flash") == "Gemini 2.5 Flash" else 1,
    )
    # Gemini APIキー：環境変数で設定済みなら入力欄を出さない（DOM/F12漏洩防止）
    if cfg.get("_gemini_from_env"):
        st.success("Gemini APIキー：環境変数で設定済み")
    else:
        cfg["gemini_api_key"] = st.text_input(
            "Gemini APIキー（Gemini選択時）", value=cfg.get("gemini_api_key", ""), type="password")
    cfg["vision_model"] = st.text_input(
        "ビジョンモデル（Ollama選択時）", value=cfg.get("vision_model", core.OCR_VISION_MODEL))
    cfg["gen_model"] = st.text_input(
        "原稿生成モデル（Ollama）", value=cfg.get("gen_model", core.GEN_MODEL_DEFAULT))
    # Ollamaキー：環境変数のみで扱い、UIには一切出さない
    if cfg.get("_ollama_key_from_env"):
        st.success("Ollama APIキー：環境変数で設定済み")
    # Ollamaホスト：環境変数で設定済みなら表示のみ
    if os.environ.get("OLLAMA_HOST"):
        st.caption(f"Ollamaホスト：{cfg.get('ollama_host')}（環境変数）")
    else:
        cfg["ollama_host"] = st.text_input(
            "Ollamaホスト（空欄=localhost）", value=cfg.get("ollama_host", ""))

    # 設定変更の自動保存（※APIキーなど秘密は保存対象から除外）
    persist = {k: v for k, v in cfg.items()
               if k in ("ocr_engine", "vision_model", "gen_model")}
    if persist != st.session_state.get("_saved_cfg"):
        save_settings(persist)
        st.session_state["_saved_cfg"] = dict(persist)

    if not HAS_PASTE:
        st.warning("画面キャプチャ貼り付けは streamlit-paste-button 未導入です。"
                   "setup_robot.bat を再実行してください。")

use_ollama = cfg.get("ocr_engine", "").startswith("Ollama")

# ============================================================
# ① アイデアシートを入力
# ============================================================
st.subheader("① アイデアシートを入力")
in1, in2 = st.columns(2)
with in1:
    sheet = st.file_uploader("ファイル（PDF / JPG / PNG）",
                             type=["pdf", "jpg", "jpeg", "png"], key="sheet_file")
with in2:
    if HAS_PASTE:
        st.caption("画面キャプチャ（Win+Shift+S）→ 下のボタンで貼り付け")
        spr = paste_image_button("📋 画面キャプチャを貼り付け", key="sheet_paste")
    else:
        spr = None

# 入力ソースを解決して session_state へ（ファイル優先）
if sheet is not None:
    raw = sheet.getvalue()
    if sheet.type == "application/pdf":
        try:
            st.session_state.sheet_img = core.pdf_first_page_png(raw)
            st.session_state.sheet_mime = "image/png"
        except Exception as e:
            st.error(f"PDFの画像化に失敗（pymupdf未導入の可能性）: {e}")
    else:
        st.session_state.sheet_img = raw
        st.session_state.sheet_mime = sheet.type
elif spr is not None and spr.image_data is not None:
    st.session_state.sheet_img = pil_to_png(spr.image_data)
    st.session_state.sheet_mime = "image/png"

if st.session_state.sheet_img:
    st.image(st.session_state.sheet_img, caption="読み取り対象", width=420)
    if st.button("OCR実行", type="primary"):
        if not use_ollama and not cfg.get("gemini_api_key"):
            st.error("Gemini APIキーを設定してください。または OCRエンジンを Ollama に切り替えてください。")
        else:
            with st.spinner("手書き文字を読み取っています…"):
                try:
                    st.session_state.idea = core.run_ocr(
                        st.session_state.sheet_img,
                        engine="ollama" if use_ollama else "gemini",
                        mime=st.session_state.sheet_mime,
                        gemini_api_key=cfg.get("gemini_api_key"),
                        vision_model=cfg.get("vision_model", core.OCR_VISION_MODEL),
                        host=cfg.get("ollama_host") or None,
                    )
                    st.success("読み取り完了。元画像を見ながら確認してください。")
                except Exception as e:
                    st.error(f"OCRエラー: {e}")

# ============================================================
# ② 画像を見ながら確認・修正
# ============================================================
idea = st.session_state.idea
if idea:
    st.divider()
    st.subheader("② 画像を見ながら確認・修正")
    img_col, form_col = st.columns([1, 1], gap="large")

    with img_col:
        if st.session_state.sheet_img:
            st.image(st.session_state.sheet_img, use_container_width=True)
        else:
            st.info("元画像がありません")

    with form_col:
        idea["robot_name"] = st.text_input("ロボットの名称", idea.get("robot_name", ""))
        cc1, cc2 = st.columns(2)
        with cc1:
            idea["school"] = st.text_input("学校名", idea.get("school", ""))
            idea["grade"] = st.text_input("学年・組", idea.get("grade", ""))
        with cc2:
            idea["author"] = st.text_input("氏名", idea.get("author", ""))
            idea["price"] = st.text_input("想定価格", idea.get("price", ""))
        idea["impression"] = st.text_area("感想", idea.get("impression", ""), height=68)
        idea["motivation"] = st.text_area("なぜ作りたいか", idea.get("motivation", ""), height=68)
        idea["target_demand"] = st.text_area("誰の役に立つ", idea.get("target_demand", ""), height=68)
        idea["shape_note"] = st.text_area("形・大きさの説明", idea.get("shape_note", ""), height=68)
        steps_text = st.text_area("使い方（1行ずつ）", "\n".join(idea.get("usage_steps", [])), height=110)
        idea["usage_steps"] = [s.strip() for s in steps_text.splitlines() if s.strip()]

    # 全幅で拡大表示（手書きの細部確認用）
    if st.session_state.sheet_img:
        with st.expander("🔍 元画像を大きく表示"):
            st.image(st.session_state.sheet_img, use_container_width=True)

    st.session_state.idea = idea

    # ============================================================
    # ③ 写真を追加（ファイル or 画面キャプチャ）
    # ============================================================
    st.divider()
    st.subheader("③ 写真・スライド画像を追加（任意）")
    st.caption("1枚目は「形と大きさ」スライドに入ります。複数枚はギャラリーに並びます。"
               "ベータ版PPTのスライドは画面キャプチャで貼り付けると取り込めます。")
    p1, p2 = st.columns(2)
    with p1:
        pics = st.file_uploader("写真ファイル（複数可）",
                                type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="photo_files")
    with p2:
        if HAS_PASTE:
            ppr = paste_image_button("📋 キャプチャを貼り付けて追加", key="photo_paste")
            if ppr is not None and ppr.image_data is not None:
                png = pil_to_png(ppr.image_data)
                if h(png) not in {h(x) for x in st.session_state.pasted_photos}:
                    st.session_state.pasted_photos.append(png)
                    st.success("貼り付けた画像を追加しました")
        if st.session_state.pasted_photos and st.button("貼り付け画像をクリア"):
            st.session_state.pasted_photos = []

    # ベータ版PowerPointの取り込み
    st.markdown("**ベータ版のPowerPointを取り込む**")
    st.caption("仮に作った .pptx をアップロードして取り込みます。"
               "「中の写真だけ抽出」は実機写真などの素材だけを拾うので、各ページへの振り分けに向きます。")
    draft = st.file_uploader("ベータ版 PowerPoint（.pptx）", type=["pptx"], key="draft_pptx")
    mode = st.radio("取り込み方", ["中の写真だけ抽出（推奨）", "各スライドを画像化"], horizontal=True)
    if draft is not None and st.button("PPTを取り込む"):
        data = draft.getvalue()
        with st.spinner("取り込み中…"):
            try:
                if mode.startswith("中の写真"):
                    st.session_state.draft_slides = core.pptx_extract_media_images(data)
                    msg = "枚の写真"
                else:
                    st.session_state.draft_slides = core.pptx_to_slide_images(data)
                    msg = "枚のスライド画像"
                st.success(f"{len(st.session_state.draft_slides)} {msg}を取り込みました")
            except Exception as e:
                st.session_state.draft_slides = []
                st.warning(f"取り込みに失敗しました: {e}")
        try:
            st.session_state.draft_text = core.pptx_extract_text(data)
        except Exception:
            st.session_state.draft_text = []
    if st.session_state.draft_slides:
        st.image(st.session_state.draft_slides, width=120)
        if st.button("取り込んだ画像をクリア"):
            st.session_state.draft_slides = []
            st.session_state.draft_text = []
    if st.session_state.draft_text:
        with st.expander("ベータ版PPTの文章（コピーして各欄に活用できます）"):
            for d in st.session_state.draft_text:
                if d["text"].strip():
                    st.markdown(f"**Slide {d['slide']}**")
                    st.text(d["text"])

    # ---- 画像をどのページに置くか割り当て（ギャラリーは作らない）----
    uploaded = [p.getvalue() for p in pics] if pics else []
    pool = uploaded + st.session_state.pasted_photos + st.session_state.draft_slides
    images_by_key = {}
    if pool:
        st.markdown("**画像をページに割り当て**")
        st.caption("各画像を置きたいページを選んでください（同じページは先に選んだ画像が使われます）。"
                   "選ばなかった画像は使いません。ギャラリーページは作りません。")
        labels = ["使わない", "このロボットってなに？", "なんで作りたいの？",
                  "どうやって動く？", "形と大きさ", "だれの役に立つ？"]
        label_to_key = {
            "このロボットってなに？": "concept", "なんで作りたいの？": "why",
            "どうやって動く？": "how", "形と大きさ": "shape", "だれの役に立つ？": "who",
        }
        # 既定の自動割り当て（形と大きさ→なに→どうやって→だれ→なんで の順）
        auto_order = ["形と大きさ", "このロボットってなに？", "どうやって動く？",
                      "だれの役に立つ？", "なんで作りたいの？"]
        ncol = 4
        rows = (len(pool) + ncol - 1) // ncol
        assigned_labels = {}
        for r in range(rows):
            cols = st.columns(ncol)
            for c in range(ncol):
                i = r * ncol + c
                if i >= len(pool):
                    break
                with cols[c]:
                    st.image(pool[i], use_container_width=True)
                    default = auto_order[i] if i < len(auto_order) else "使わない"
                    sel = st.selectbox("配置先", labels,
                                       index=labels.index(default), key=f"assign_{i}")
                    assigned_labels[i] = sel
        # 先勝ちでマップ化
        for i in range(len(pool)):
            sel = assigned_labels.get(i, "使わない")
            if sel != "使わない":
                k = label_to_key[sel]
                if k not in images_by_key:
                    images_by_key[k] = pool[i]
        used = len(images_by_key)
        st.caption(f"プール {len(pool)} 枚 / 配置 {used} 枚"
                   f"（ファイル {len(uploaded)}＋貼り付け {len(st.session_state.pasted_photos)}"
                   f"＋PPT {len(st.session_state.draft_slides)}）")
    st.session_state.images_by_key = images_by_key

    # ============================================================
    # ④ プレゼンを生成
    # ============================================================
    st.divider()
    st.subheader("④ プレゼンを生成")
    if st.button("PowerPointを作る", type="primary"):
        with st.spinner("スライド原稿を作成中（Ollama）…"):
            try:
                content = core.generate_presentation(
                    idea, model=cfg["gen_model"], host=cfg.get("ollama_host") or None)
            except Exception as e:
                st.warning(f"AI生成に失敗したため、シート内容から簡易生成します: {e}")
                content = core._fallback_content(idea)
        with st.spinner("PowerPointを組み立て中…"):
            out = core.build_pptx(content, output_path="robot_presentation.pptx",
                                  price_value=idea.get("price", ""),
                                  images_by_key=st.session_state.get("images_by_key", {}))
        with open(out, "rb") as f:
            st.download_button(
                "📥 PowerPointをダウンロード", f.read(),
                file_name=f"{idea.get('robot_name','robot') or 'robot'}_提案.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary")
        with st.expander("生成された原稿（確認用）"):
            st.json(content)
