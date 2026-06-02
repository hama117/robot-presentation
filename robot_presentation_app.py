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
# 生成結果（DL後も残るよう session_state に保持）
st.session_state.setdefault("normal_result", None)   # {"bytes":..., "filename":..., "content":...}
st.session_state.setdefault("sato_result", None)     # 佐藤方式
st.session_state.setdefault("biz_result", None)      # business方式
# プロジェクト復元時に使う：ファイル割り当ての復元データ（インポート時に流し込む）
st.session_state.setdefault("restored_images_by_key", None)
st.session_state.setdefault("restored_sato_images_by_key", None)
st.session_state.setdefault("restored_biz_images_by_key", None)
st.session_state.setdefault("restored_cover_image", None)
st.session_state.setdefault("restored_uploaded", [])   # ファイルアップロード復元用
# プロジェクト名（保存時のファイル名にも使用）
st.session_state.setdefault("project_name", "")

st.title("ロボット甲子園 プレゼン作成")
st.caption("アイデアシートを読み取り → 画像を見ながら修正 → 白表紙のPowerPointに出力")


# ============================================================
# プロジェクトの保存・読み込み（タイトル下に常時表示）
# ============================================================
with st.expander("💾 プロジェクトを保存／📂 読み込み（編集中の状態と生成済みPPTをまとめて1ファイルに）", expanded=False):
    st.caption("**保存**：いまの編集状態（アイデアシート画像・OCR結果・写真・割り当て・表紙アイコン・生成済みPPT）を"
               "1つの.zipにまとめてダウンロードします。"
               "**読み込み**：以前保存した.zipをアップロードすれば、その時点から続きを編集できます。")

    pcol1, pcol2 = st.columns(2)

    # ---- 保存 ----
    with pcol1:
        st.markdown("**📥 プロジェクトを保存（エクスポート）**")
        if st.session_state.idea or st.session_state.normal_result or st.session_state.sato_result:
            # 現在の編集状態を集めて save_project_zip に渡す
            try:
                cur_idea = st.session_state.idea or {}
                robot_nm = (cur_idea.get("robot_name") if cur_idea else "") or "robot"
                state_to_save = {
                    "robot_name": robot_nm,
                    "idea": cur_idea,
                    "sheet_img": st.session_state.sheet_img,
                    "sheet_mime": st.session_state.sheet_mime,
                    "pasted_photos": st.session_state.pasted_photos,
                    "draft_slides": st.session_state.draft_slides,
                    "draft_text": st.session_state.draft_text,
                    "uploaded_files_data": st.session_state.get("uploaded_files_data", []),
                    "cover_image": st.session_state.get("cover_image_cached"),
                    "images_by_key": st.session_state.get("images_by_key", {}),
                    "sato_images_by_key": st.session_state.get("sato_images_by_key", {}),
                    "biz_images_by_key": st.session_state.get("biz_images_by_key", {}),
                    "normal_pptx_bytes": (st.session_state.normal_result or {}).get("bytes"),
                    "sato_pptx_bytes": (st.session_state.sato_result or {}).get("bytes"),
                    "biz_pptx_bytes": (st.session_state.biz_result or {}).get("bytes"),
                    "normal_content": (st.session_state.normal_result or {}).get("content"),
                    "sato_content": (st.session_state.sato_result or {}).get("content"),
                    "biz_content": (st.session_state.biz_result or {}).get("content"),
                    "sato_research": (st.session_state.sato_result or {}).get("research", ""),
                    "biz_research": (st.session_state.biz_result or {}).get("research", ""),
                }
                zip_bytes = core.save_project_zip(state_to_save)
                from datetime import datetime
                ts = datetime.now().strftime("%Y%m%d_%H%M")
                st.download_button(
                    "📥 プロジェクトを.zipでダウンロード",
                    zip_bytes,
                    file_name=f"{robot_nm}_project_{ts}.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True, key="dl_project")
                st.caption(f"サイズ: 約 {len(zip_bytes) / 1024:.0f} KB")
            except Exception as e:
                st.error(f"プロジェクト保存エラー: {e}")
        else:
            st.info("アイデアシートを読み取ると、ここから保存できるようになります。")

    # ---- 読み込み ----
    with pcol2:
        st.markdown("**📂 プロジェクトを読み込み（インポート）**")
        proj_file = st.file_uploader(
            "以前保存した .zip を選択", type=["zip"], key="project_zip_uploader")
        if proj_file is not None:
            if st.button("📂 このプロジェクトを読み込む", key="btn_load_project",
                         type="primary", use_container_width=True):
                try:
                    restored = core.load_project_zip(proj_file.getvalue())
                    # session_state に流し込み
                    st.session_state.idea = restored.get("idea") or None
                    st.session_state.sheet_img = restored.get("sheet_img")
                    st.session_state.sheet_mime = restored.get("sheet_mime", "image/png")
                    st.session_state.pasted_photos = restored.get("pasted_photos", []) or []
                    st.session_state.draft_slides = restored.get("draft_slides", []) or []
                    st.session_state.draft_text = restored.get("draft_text", []) or []
                    # ファイル割り当て・表紙画像は次回再描画時に復元するため一時保存
                    st.session_state.restored_images_by_key = restored.get("images_by_key") or {}
                    st.session_state.restored_sato_images_by_key = restored.get("sato_images_by_key") or {}
                    st.session_state.restored_biz_images_by_key = restored.get("biz_images_by_key") or {}
                    st.session_state.restored_cover_image = restored.get("cover_image")
                    st.session_state.restored_uploaded = restored.get("uploaded_files_data", []) or []
                    # 生成結果も復元
                    rn = restored.get("robot_name", "robot") or "robot"
                    if restored.get("normal_pptx_bytes"):
                        st.session_state.normal_result = {
                            "bytes": restored["normal_pptx_bytes"],
                            "filename": f"{rn}_提案.pptx",
                            "content": restored.get("normal_content") or {},
                        }
                    else:
                        st.session_state.normal_result = None
                    if restored.get("sato_pptx_bytes"):
                        st.session_state.sato_result = {
                            "bytes": restored["sato_pptx_bytes"],
                            "filename": f"{rn}_発表_佐藤方式.pptx",
                            "content": restored.get("sato_content") or {},
                            "research": restored.get("sato_research", ""),
                        }
                    else:
                        st.session_state.sato_result = None
                    if restored.get("biz_pptx_bytes"):
                        st.session_state.biz_result = {
                            "bytes": restored["biz_pptx_bytes"],
                            "filename": f"{rn}_提案書_business方式.pptx",
                            "content": restored.get("biz_content") or {},
                            "research": restored.get("biz_research", ""),
                        }
                    else:
                        st.session_state.biz_result = None
                    st.session_state.project_name = rn
                    st.success(f"プロジェクト「{rn}」を読み込みました（保存日時: {restored.get('saved_at', '不明')}）")
                    st.rerun()
                except Exception as e:
                    st.error(f"読み込みに失敗しました: {e}")

    # ---- 逆輸入：PowerPointで手編集したPPTをプロジェクトに戻す ----
    st.markdown("---")
    st.markdown("**🔄 PowerPointで手編集したPPTを取り込む（逆輸入）**")
    st.caption("PowerPointで直接編集したファイルを、プロジェクトに戻して保存できます。"
               "次回エクスポートする.zipに反映されます。")
    rcol1, rcol2, rcol3 = st.columns(3)
    with rcol1:
        edited_normal = st.file_uploader(
            "①従来方式の編集済みPPTX", type=["pptx"], key="reimport_normal_uploader")
        if edited_normal is not None and st.button(
                "①従来方式に取り込む", key="btn_reimport_normal", use_container_width=True):
            edited_bytes = edited_normal.getvalue()
            cur = st.session_state.normal_result or {}
            cur_idea = st.session_state.idea or {}
            rn = (cur_idea.get("robot_name") or st.session_state.project_name or "robot")
            st.session_state.normal_result = {
                "bytes": edited_bytes,
                "filename": cur.get("filename", f"{rn}_提案.pptx"),
                "content": cur.get("content") or {},
            }
            st.success("①従来方式に手編集済みPPTを取り込みました。")
            st.rerun()
    with rcol2:
        edited_sato = st.file_uploader(
            "②佐藤先生方式の編集済みPPTX", type=["pptx"], key="reimport_sato_uploader")
        if edited_sato is not None and st.button(
                "②佐藤方式に取り込む", key="btn_reimport_sato", use_container_width=True):
            edited_bytes = edited_sato.getvalue()
            cur = st.session_state.sato_result or {}
            cur_idea = st.session_state.idea or {}
            rn = (cur_idea.get("robot_name") or st.session_state.project_name or "robot")
            st.session_state.sato_result = {
                "bytes": edited_bytes,
                "filename": cur.get("filename", f"{rn}_発表_佐藤方式.pptx"),
                "content": cur.get("content") or {},
                "research": cur.get("research", ""),
            }
            st.success("②佐藤方式に手編集済みPPTを取り込みました。")
            st.rerun()
    with rcol3:
        edited_biz = st.file_uploader(
            "③business方式の編集済みPPTX", type=["pptx"], key="reimport_biz_uploader")
        if edited_biz is not None and st.button(
                "③business方式に取り込む", key="btn_reimport_biz", use_container_width=True):
            edited_bytes = edited_biz.getvalue()
            cur = st.session_state.biz_result or {}
            cur_idea = st.session_state.idea or {}
            rn = (cur_idea.get("robot_name") or st.session_state.project_name or "robot")
            st.session_state.biz_result = {
                "bytes": edited_bytes,
                "filename": cur.get("filename", f"{rn}_提案書_business方式.pptx"),
                "content": cur.get("content") or {},
                "research": cur.get("research", ""),
            }
            st.success("③business方式に手編集済みPPTを取り込みました。")
            st.rerun()


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
    cfg["ocr_engine"] = "Gemini 2.5 Flash"   # OCRはGemini固定
    # Gemini APIキー：環境変数で設定済みなら入力欄を出さない（DOM/F12漏洩防止）
    if cfg.get("_gemini_from_env"):
        st.success("Gemini APIキー：環境変数で設定済み")
    else:
        cfg["gemini_api_key"] = st.text_input(
            "Gemini APIキー", value=cfg.get("gemini_api_key", ""), type="password")
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
    persist = {k: v for k, v in cfg.items() if k in ("gen_model",)}
    if persist != st.session_state.get("_saved_cfg"):
        save_settings(persist)
        st.session_state["_saved_cfg"] = dict(persist)

    if not HAS_PASTE:
        st.warning("画面キャプチャ貼り付けは streamlit-paste-button 未導入です。"
                   "setup_robot.bat を再実行してください。")

use_ollama = False   # OCRはGemini固定（後方互換のため残置）


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
        if not cfg.get("gemini_api_key"):
            st.error("Gemini APIキーが設定されていません（環境変数 GEMINI_API_KEY を確認してください）。")
        else:
            with st.spinner("手書き文字を読み取っています…"):
                try:
                    st.session_state.idea = core.run_ocr(
                        st.session_state.sheet_img,
                        engine="gemini",
                        mime=st.session_state.sheet_mime,
                        gemini_api_key=cfg.get("gemini_api_key"),
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
            st.image(st.session_state.sheet_img, width='stretch')
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
            st.image(st.session_state.sheet_img, width='stretch')

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
    # プロジェクト復元から来たアップロード写真をプールに加える
    restored_uploaded = st.session_state.get("restored_uploaded", [])
    # 次回プロジェクト保存時に拾えるよう、現在のアップロード一覧をsessionに保存
    st.session_state.uploaded_files_data = uploaded + restored_uploaded
    pool = uploaded + restored_uploaded + st.session_state.pasted_photos + st.session_state.draft_slides

    images_by_key = {}
    sato_images_by_key = {}
    biz_images_by_key = {}

    # プロジェクト復元時の割り当てヒント：bytes同士を見比べてindexを逆引きできるようにマップ
    restored_imap = st.session_state.get("restored_images_by_key") or {}
    restored_satomap = st.session_state.get("restored_sato_images_by_key") or {}
    restored_bizmap = st.session_state.get("restored_biz_images_by_key") or {}

    if pool:
        st.markdown("**画像をページに割り当て**")
        st.caption("各画像について、**①従来方式 / ②佐藤先生方式 / ③business方式** それぞれの配置先を選んでください。"
                   "同じページに複数選んだ場合は先に選んだ画像が使われます。"
                   "「使わない」を選んだ方式では、その画像は表示されません。")

        # 従来方式：従来のキー（concept/why/how/shape/who）
        labels_conv = ["使わない", "このロボットってなに？", "なんで作りたいの？",
                       "どうやって動く？", "形と大きさ", "だれの役に立つ？"]
        label_to_key_conv = {
            "このロボットってなに？": "concept", "なんで作りたいの？": "why",
            "どうやって動く？": "how", "形と大きさ": "shape", "だれの役に立つ？": "who",
        }
        key_to_label_conv = {v: k for k, v in label_to_key_conv.items()}
        auto_conv = ["形と大きさ", "このロボットってなに？", "どうやって動く？",
                     "だれの役に立つ？", "なんで作りたいの？"]

        # 佐藤方式：スライドindexにマップ（0=起, 1=承, 2=転, 3=転続き, 4=結準備, 5=結）
        labels_sato = ["使わない",
                       "起：表紙・つかみ",
                       "承：このロボットってなに？",
                       "転：見どころ（実演）⭐",
                       "転続き：どうやって動く？",
                       "結準備：だれの役に立つ？",
                       "結:まとめ"]
        label_to_idx_sato = {
            "起：表紙・つかみ": 0,
            "承：このロボットってなに？": 1,
            "転：見どころ（実演）⭐": 2,
            "転続き：どうやって動く？": 3,
            "結準備：だれの役に立つ？": 4,
            "結:まとめ": 5,
        }
        idx_to_label_sato = {v: k for k, v in label_to_idx_sato.items()}
        auto_sato = ["転：見どころ（実演）⭐", "転続き：どうやって動く？",
                     "起：表紙・つかみ", "承：このロボットってなに？", "結準備：だれの役に立つ？"]

        # business方式：スライドindexにマップ（0=problem, 1=solution, 2=how, 3=impact, 4=why, 5=next）
        labels_biz = ["使わない",
                      "01 課題と市場背景",
                      "02 ソリューション概要",
                      "03 仕組みと差別化",
                      "04 活用シーンと効果⭐",
                      "05 今、なぜ我々か",
                      "06 次のアクション"]
        label_to_idx_biz = {
            "01 課題と市場背景": 0,
            "02 ソリューション概要": 1,
            "03 仕組みと差別化": 2,
            "04 活用シーンと効果⭐": 3,
            "05 今、なぜ我々か": 4,
            "06 次のアクション": 5,
        }
        idx_to_label_biz = {v: k for k, v in label_to_idx_biz.items()}
        # business方式の自動配分（impact → how → solution → problem → why の順）
        auto_biz = ["04 活用シーンと効果⭐", "03 仕組みと差別化",
                    "02 ソリューション概要", "01 課題と市場背景", "05 今、なぜ我々か"]

        # プール内bytesから「復元時の割り当て」を逆引きするマップ
        def _restored_label_for_pool_idx(pool_bytes_obj, side):
            """side: 'conv' / 'sato' / 'biz'。プール内画像bytesが復元データで何に割り当てられていたか。"""
            if side == "conv":
                target_map = restored_imap
                label_map = key_to_label_conv
            elif side == "sato":
                target_map = restored_satomap
                label_map = idx_to_label_sato
            else:
                target_map = restored_bizmap
                label_map = idx_to_label_biz
            for k, b in target_map.items():
                if b == pool_bytes_obj:
                    if side == "conv":
                        return label_map.get(k, None)
                    else:
                        return label_map.get(int(k), None)
            return None

        # 画像を3列で並べ、各画像にドロップダウンを3つ重ねる
        ncol = 3
        rows = (len(pool) + ncol - 1) // ncol
        for r in range(rows):
            cols = st.columns(ncol)
            for c in range(ncol):
                i = r * ncol + c
                if i >= len(pool):
                    break
                with cols[c]:
                    st.image(pool[i], width='stretch')
                    # ① 従来方式
                    restored_c = _restored_label_for_pool_idx(pool[i], "conv")
                    default_c = restored_c if restored_c else (auto_conv[i] if i < len(auto_conv) else "使わない")
                    if default_c not in labels_conv:
                        default_c = "使わない"
                    sel_c = st.selectbox(
                        "① 従来方式の配置先", labels_conv,
                        index=labels_conv.index(default_c), key=f"assign_conv_{i}")
                    if sel_c != "使わない":
                        k = label_to_key_conv[sel_c]
                        if k not in images_by_key:
                            images_by_key[k] = pool[i]
                    # ② 佐藤方式
                    restored_s = _restored_label_for_pool_idx(pool[i], "sato")
                    default_s = restored_s if restored_s else (auto_sato[i] if i < len(auto_sato) else "使わない")
                    if default_s not in labels_sato:
                        default_s = "使わない"
                    sel_s = st.selectbox(
                        "② 佐藤方式の配置先", labels_sato,
                        index=labels_sato.index(default_s), key=f"assign_sato_{i}")
                    if sel_s != "使わない":
                        idx_s = label_to_idx_sato[sel_s]
                        if idx_s not in sato_images_by_key:
                            sato_images_by_key[idx_s] = pool[i]
                    # ③ business方式
                    restored_b = _restored_label_for_pool_idx(pool[i], "biz")
                    default_b = restored_b if restored_b else (auto_biz[i] if i < len(auto_biz) else "使わない")
                    if default_b not in labels_biz:
                        default_b = "使わない"
                    sel_b = st.selectbox(
                        "③ business方式の配置先", labels_biz,
                        index=labels_biz.index(default_b), key=f"assign_biz_{i}")
                    if sel_b != "使わない":
                        idx_b = label_to_idx_biz[sel_b]
                        if idx_b not in biz_images_by_key:
                            biz_images_by_key[idx_b] = pool[i]

        used_c = len(images_by_key)
        used_s = len(sato_images_by_key)
        used_b = len(biz_images_by_key)
        st.caption(f"プール {len(pool)} 枚 ／ ①従来 {used_c} 枚 ／ ②佐藤 {used_s} 枚 ／ ③business {used_b} 枚配置"
                   f"（内訳：ファイル {len(uploaded) + len(restored_uploaded)}"
                   f"＋貼り付け {len(st.session_state.pasted_photos)}"
                   f"＋PPT {len(st.session_state.draft_slides)}）")
    st.session_state.images_by_key = images_by_key
    st.session_state.sato_images_by_key = sato_images_by_key
    st.session_state.biz_images_by_key = biz_images_by_key

    # ============================================================
    # ④ プレゼンを生成
    # ============================================================
    st.divider()
    st.subheader("④ プレゼンを生成")
    imap = st.session_state.get("images_by_key", {})
    sato_imap = st.session_state.get("sato_images_by_key", {})
    biz_imap = st.session_state.get("biz_images_by_key", {})
    # 佐藤・business方式：明示割り当てが無ければ photos で自動配分に回すためのリスト
    sato_photos_fallback = [v for v in imap.values() if v]
    biz_photos_fallback = [v for v in imap.values() if v]

    # 表紙アイコン/写真（佐藤方式・business方式で使用。任意。未指定なら既定装飾）
    st.markdown("**表紙アイコン／写真（任意）**")
    st.caption("②佐藤先生方式・③business方式の表紙に置く画像をアップロードできます。"
               "未指定の場合は既定装飾が入ります（①従来方式の表紙には影響しません）。")
    cover_file = st.file_uploader("表紙画像（JPG/PNG）",
                                   type=["jpg", "jpeg", "png"], key="cover_icon_file")
    if HAS_PASTE:
        cpr = paste_image_button("📋 表紙画像を貼り付け", key="cover_paste")
    else:
        cpr = None
    # 表紙画像のbytesを確定（優先順位: 新規アップロード > 貼り付け > 復元値）
    cover_bytes = None
    if cover_file is not None:
        cover_bytes = cover_file.getvalue()
    elif cpr is not None and cpr.image_data is not None:
        cover_bytes = pil_to_png(cpr.image_data)
    elif st.session_state.restored_cover_image:
        # プロジェクト復元から来た表紙画像
        cover_bytes = st.session_state.restored_cover_image
        st.caption("📂 復元されたプロジェクトの表紙画像を使用中")
    # 次回プロジェクト保存用にキャッシュ
    st.session_state.cover_image_cached = cover_bytes
    if cover_bytes:
        st.image(cover_bytes, caption="表紙に配置されます", width=180)

    st.divider()

    # ───────────────────────────────────────────────────────────
    # 生成方式1：従来（アイデアシート準拠）
    # ───────────────────────────────────────────────────────────
    st.markdown("### ① アイデアシート方式で作る")
    st.caption("アイデアシートの内容に沿って8枚構成のプレゼン原稿を作成します。")
    if st.button("① アイデアシート方式で作る", type="primary",
                 key="btn_normal", use_container_width=True):
        with st.spinner("スライド原稿を作成中（Ollama）…"):
            try:
                content = core.generate_presentation(
                    idea, model=cfg["gen_model"], host=cfg.get("ollama_host") or None)
            except Exception as e:
                st.warning(f"AI生成に失敗したため、シート内容から簡易生成します: {e}")
                content = core._fallback_content(idea)
        with st.spinner("PowerPointを組み立て中…"):
            out = core.build_pptx(content, output_path="robot_presentation.pptx",
                                  price_value=idea.get("price", ""), images_by_key=imap)
        with open(out, "rb") as f:
            pptx_bytes = f.read()
        # 生成結果を session_state に保存（DL後もボタンや原稿が消えないように）
        st.session_state.normal_result = {
            "bytes": pptx_bytes,
            "filename": f"{idea.get('robot_name','robot') or 'robot'}_提案.pptx",
            "content": content,
        }
        st.success("生成が完了しました。下のボタンからダウンロードできます。")

    # 生成結果が残っていれば、常に表示（DL後も消えない）
    if st.session_state.normal_result:
        res = st.session_state.normal_result
        st.download_button(
            "📥 ダウンロード（アイデアシート方式）",
            res["bytes"],
            file_name=res["filename"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary", key="dl_normal",
            use_container_width=True)
        with st.expander("生成された原稿（確認用）"):
            st.json(res["content"])
        if st.button("🗑️ ①の結果をクリア", key="clear_normal"):
            st.session_state.normal_result = None
            st.rerun()

    st.divider()

    # ───────────────────────────────────────────────────────────
    # 生成方式2：佐藤先生方式（6フェーズ・Gemini Webリサーチ込み）
    # ───────────────────────────────────────────────────────────
    st.markdown("### ② 佐藤先生方式で作る（6フェーズ）")
    st.caption("佐藤先生の口頭発表指導（起／承／転／転続き／結準備／結 の6フェーズ）に従い、"
               "各スライドに「役割（存在意義）」と本文段落・カード説明を入れて文章量を確保します。"
               "「転」スライドは『ここで実演します』の静止画＋プレースホルダで出力するので、"
               "完成後にご自身で動画に差し替えてください。")
    use_research = st.checkbox(
        "🔎 Gemini 2.5 Flash + Google検索で背景を調査して文章を厚くする",
        value=True, key="sato_use_research",
        help="課題の統計、類似事例、技術トレンドなどをWebから取り込み、各スライドの本文に反映します。"
             "Gemini APIキーが必要（OCRと同じキーを使用）。")
    if st.button("② 佐藤先生方式で作る（6フェーズ）", type="primary",
                 key="btn_sato", use_container_width=True):
        gem_key = cfg.get("gemini_api_key", "") if use_research else None
        # 1) Webリサーチ（任意）
        research_text = ""
        if gem_key:
            with st.spinner("Gemini 2.5 Flash が Google で背景情報を調査中…"):
                try:
                    research_text = core.gemini_research_topic(idea, gem_key)
                except Exception as e:
                    st.info(f"リサーチをスキップ（{e}）。シート情報のみで生成します。")
                    research_text = ""
        elif use_research and not gem_key:
            st.info("Gemini APIキーが未設定のため、リサーチは省略します。")
        # 2) スライド原稿生成
        with st.spinner("6フェーズの原稿を作成中（Ollama）…"):
            try:
                content = core.generate_presentation_sato(
                    idea, model=cfg["gen_model"],
                    host=cfg.get("ollama_host") or None,
                    research_text=research_text or None)
            except Exception as e:
                st.warning(f"AI生成に失敗したため、シート内容から簡易生成します: {e}")
                content = core._fallback_content_sato(idea)
        # 3) PowerPoint組み立て（表紙画像があれば渡す）
        # 画像割り当て：明示割り当て(sato_imap)があればそれを使い、無ければ photos で自動配分
        with st.spinner("PowerPointを組み立て中…"):
            if sato_imap:
                out = core.build_pptx_sato(
                    content, images_by_key=sato_imap,
                    output_path="robot_presentation_sato.pptx",
                    cover_image=cover_bytes)
            else:
                out = core.build_pptx_sato(
                    content, photos=sato_photos_fallback,
                    output_path="robot_presentation_sato.pptx",
                    cover_image=cover_bytes)
        with open(out, "rb") as f:
            pptx_bytes = f.read()
        # 生成結果を session_state に保存（DL後もボタンや原稿が消えないように）
        st.session_state.sato_result = {
            "bytes": pptx_bytes,
            "filename": f"{idea.get('robot_name','robot') or 'robot'}_発表_佐藤方式.pptx",
            "content": content,
            "research": research_text or "",
        }
        st.success("生成が完了しました。下のボタンからダウンロードできます。")

    # 生成結果が残っていれば、常に表示（DL後も消えない）
    if st.session_state.sato_result:
        res = st.session_state.sato_result
        st.download_button(
            "📥 ダウンロード（佐藤先生方式）",
            res["bytes"],
            file_name=res["filename"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary", key="dl_sato",
            use_container_width=True)
        with st.expander("生成された原稿（確認用）"):
            st.json({k: v for k, v in res["content"].items() if k != "_research"})
        if res["research"]:
            with st.expander("🔎 Gemini調査結果（参考リサーチ）"):
                st.markdown(res["research"])
        if st.button("🗑️ ②の結果をクリア", key="clear_sato"):
            st.session_state.sato_result = None
            st.rerun()

    st.divider()

    # ───────────────────────────────────────────────────────────
    # 生成方式3：business調（提案書スタイル・Gemini Webリサーチ込み）
    # ───────────────────────────────────────────────────────────
    st.markdown("### ③ business調で作る（提案書）")
    st.caption("審査員・投資家・技術者向けのbusiness調（提案書スタイル）で7枚構成のプレゼンを作成します。"
               "Problem → Solution → How → Impact → Why → Next の構成で、KPI数値・出典付きの"
               "「事業提案」として仕上げます。佐藤方式と並行して使い分けてください。")
    use_research_biz = st.checkbox(
        "🔎 Gemini 2.5 Flash + Google検索で課題の統計・市場データを取り込む",
        value=True, key="biz_use_research",
        help="課題の規模感を示す統計や、競合・関連市場の数値をWebから取り込み、KPI欄に反映します。"
             "business調は数値が命なので、強くおすすめ。Gemini APIキーが必要（OCRと同じキー）。")
    if st.button("③ business調で作る（提案書）", type="primary",
                 key="btn_biz", use_container_width=True):
        gem_key = cfg.get("gemini_api_key", "") if use_research_biz else None
        # 1) Webリサーチ
        research_text = ""
        if gem_key:
            with st.spinner("Gemini 2.5 Flash が Google で市場データを調査中…"):
                try:
                    research_text = core.gemini_research_topic(idea, gem_key)
                except Exception as e:
                    st.info(f"リサーチをスキップ（{e}）。シート情報のみで生成します。")
                    research_text = ""
        elif use_research_biz and not gem_key:
            st.info("Gemini APIキーが未設定のため、リサーチは省略します。")
        # 2) 原稿生成
        with st.spinner("business調の原稿を作成中（Ollama）…"):
            try:
                content = core.generate_presentation_biz(
                    idea, model=cfg["gen_model"],
                    host=cfg.get("ollama_host") or None,
                    research_text=research_text or None)
            except Exception as e:
                st.warning(f"AI生成に失敗したため、シート内容から簡易生成します: {e}")
                content = core._fallback_content_biz(idea)
        # 3) PPT生成
        with st.spinner("PowerPointを組み立て中…"):
            if biz_imap:
                out = core.build_pptx_biz(
                    content, images_by_key=biz_imap,
                    output_path="robot_presentation_biz.pptx",
                    cover_image=cover_bytes)
            else:
                out = core.build_pptx_biz(
                    content, photos=biz_photos_fallback,
                    output_path="robot_presentation_biz.pptx",
                    cover_image=cover_bytes)
        with open(out, "rb") as f:
            pptx_bytes = f.read()
        st.session_state.biz_result = {
            "bytes": pptx_bytes,
            "filename": f"{idea.get('robot_name','robot') or 'robot'}_提案書_business方式.pptx",
            "content": content,
            "research": research_text or "",
        }
        st.success("生成が完了しました。下のボタンからダウンロードできます。")

    # 生成結果が残っていれば、常に表示
    if st.session_state.biz_result:
        res = st.session_state.biz_result
        st.download_button(
            "📥 ダウンロード（business方式）",
            res["bytes"],
            file_name=res["filename"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary", key="dl_biz",
            use_container_width=True)
        with st.expander("生成された原稿（確認用）"):
            st.json({k: v for k, v in res["content"].items() if k != "_research"})
        if res["research"]:
            with st.expander("🔎 Gemini調査結果（参考リサーチ）"):
                st.markdown(res["research"])
        if st.button("🗑️ ③の結果をクリア", key="clear_biz"):
            st.session_state.biz_result = None
            st.rerun()
