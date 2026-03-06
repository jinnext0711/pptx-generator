"""
pptx-generator - Streamlit GUIアプリ
"""
import io
import json
import os
import sys
from datetime import datetime, date

import streamlit as st

# プロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
from src.builder import PresentationBuilder
from src.data_loader import load_from_file, load_from_text

# ページ設定
st.set_page_config(
    page_title="資料メーカー",
    page_icon="📊",
    layout="wide",
)

# カスタムCSS
st.markdown("""
<style>
    /* カラーパレットのプレビュー */
    .color-swatch {
        display: inline-block;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        border: 2px solid #ddd;
        vertical-align: middle;
        margin-right: 6px;
    }
    /* スライドプレビューカード */
    .slide-preview {
        background: #fff;
        border: 1px solid #ccc;
        border-radius: 4px;
        padding: 8px;
        margin-bottom: 10px;
        aspect-ratio: 16 / 9;
        overflow: hidden;
        position: relative;
        font-family: 'Meiryo', sans-serif;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .slide-preview .slide-num {
        position: absolute;
        top: 3px;
        left: 6px;
        font-size: 9px;
        color: #999;
    }
    .slide-preview .slide-badge {
        position: absolute;
        top: 3px;
        right: 6px;
        font-size: 8px;
        background: rgba(0,0,0,0.08);
        padding: 1px 5px;
        border-radius: 3px;
        color: #666;
    }
    .slide-preview .slide-title {
        font-size: 10px;
        font-weight: bold;
        margin-top: 16px;
        margin-bottom: 4px;
        padding-bottom: 3px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .slide-preview .slide-body {
        font-size: 8px;
        line-height: 1.5;
        color: #555;
    }
    .slide-preview .slide-body-item {
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .slide-preview .two-col {
        display: flex;
        gap: 6px;
    }
    .slide-preview .col-half {
        flex: 1;
        min-width: 0;
    }
    .slide-preview .col-title {
        font-size: 8px;
        font-weight: bold;
        margin-bottom: 2px;
    }
    .slide-preview .key-msg {
        text-align: center;
        font-size: 11px;
        font-weight: bold;
        margin: 8px 0;
        padding: 6px 4px;
    }
    .slide-preview .comp-label {
        text-align: center;
        color: #fff;
        font-size: 8px;
        font-weight: bold;
        padding: 2px 0;
        border-radius: 2px;
        margin-bottom: 3px;
    }
    .slide-preview .comp-arrow {
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        font-weight: bold;
        padding: 0 2px;
    }
    .slide-preview-title-slide {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        height: 100%;
        text-align: center;
    }
    .slide-preview-title-slide .main-title {
        font-size: 12px;
        font-weight: bold;
    }
    .slide-preview-title-slide .sub-title {
        font-size: 9px;
        margin-top: 4px;
    }
    .slide-preview-section {
        display: flex;
        align-items: center;
        justify-content: center;
        height: 100%;
        text-align: center;
        font-weight: bold;
        font-size: 11px;
    }
</style>
""", unsafe_allow_html=True)

# スライドタイプの定義（日本語ラベル→内部名）
SLIDE_TYPE_MAP = {
    "コンテンツ（箇条書き）": "content",
    "2カラム": "two_column",
    "キーメッセージ": "key_message",
    "Before / After 比較": "comparison",
    "テーブル": "table",
    "グラフ": "chart",
    "セクション区切り": "section",
    "画像": "image",
}
# 逆引き
SLIDE_TYPE_LABELS = {v: k for k, v in SLIDE_TYPE_MAP.items()}


def load_color_schemes():
    """カラースキームの一覧と詳細を読み込む"""
    color_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.COLOR_DIR)
    schemes = {}
    for filename in sorted(os.listdir(color_dir)):
        if filename.endswith(".json"):
            with open(os.path.join(color_dir, filename), "r", encoding="utf-8") as f:
                data = json.load(f)
            schemes[os.path.splitext(filename)[0]] = data
    return schemes


def load_templates():
    """テンプレートの一覧と詳細を読み込む"""
    template_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), config.TEMPLATE_DIR)
    templates = {}
    for filename in sorted(os.listdir(template_dir)):
        if filename.endswith(".json"):
            with open(os.path.join(template_dir, filename), "r", encoding="utf-8") as f:
                data = json.load(f)
            templates[os.path.splitext(filename)[0]] = data
    return templates


def build_data_from_form(title, subtitle, author, slides_data):
    """フォーム入力からデータdictを構築"""
    data = {
        "title": title,
        "date": str(date.today()),
        "slides": []
    }

    if author:
        data["author"] = author

    # タイトルスライド
    data["slides"].append({
        "type": "title",
        "title": title,
        "subtitle": subtitle,
    })

    # 各スライド
    for s in slides_data:
        data["slides"].append(s)

    return data


def render_slide_editor(index, slide=None):
    """スライド編集UIを描画"""
    slide = slide or {}

    col1, col2 = st.columns([1, 3])

    with col1:
        current_type = slide.get("type", "content")
        default_label = SLIDE_TYPE_LABELS.get(current_type, "コンテンツ（箇条書き）")
        default_idx = list(SLIDE_TYPE_MAP.keys()).index(default_label)

        selected_type = st.selectbox(
            "種別",
            list(SLIDE_TYPE_MAP.keys()),
            index=default_idx,
            key=f"type_{index}"
        )
        slide_type = SLIDE_TYPE_MAP[selected_type]

    with col2:
        slide_title = st.text_input(
            "タイトル",
            value=slide.get("title", ""),
            key=f"title_{index}"
        )

    result = {"type": slide_type, "title": slide_title}

    # === タイプ別の入力フィールド ===

    if slide_type == "content":
        body_text = st.text_area(
            "内容（1行ごとに箇条書き項目）",
            value="\n".join(slide.get("body", [])) if isinstance(slide.get("body"), list) else slide.get("body", ""),
            height=120,
            key=f"body_{index}"
        )
        result["body"] = [line.strip() for line in body_text.split("\n") if line.strip()]

    elif slide_type == "two_column":
        c_left, c_right = st.columns(2)
        with c_left:
            result["left_title"] = st.text_input(
                "左カラム見出し",
                value=slide.get("left_title", ""),
                key=f"left_title_{index}"
            )
            left_text = st.text_area(
                "左カラム内容（1行ごと）",
                value="\n".join(slide.get("left_body", [])) if isinstance(slide.get("left_body"), list) else slide.get("left_body", ""),
                height=100,
                key=f"left_body_{index}"
            )
            result["left_body"] = [l.strip() for l in left_text.split("\n") if l.strip()]
        with c_right:
            result["right_title"] = st.text_input(
                "右カラム見出し",
                value=slide.get("right_title", ""),
                key=f"right_title_{index}"
            )
            right_text = st.text_area(
                "右カラム内容（1行ごと）",
                value="\n".join(slide.get("right_body", [])) if isinstance(slide.get("right_body"), list) else slide.get("right_body", ""),
                height=100,
                key=f"right_body_{index}"
            )
            result["right_body"] = [l.strip() for l in right_text.split("\n") if l.strip()]

    elif slide_type == "key_message":
        result["message"] = st.text_input(
            "キーメッセージ（中央に大きく表示）",
            value=slide.get("message", ""),
            key=f"message_{index}",
            placeholder="例: 売上は前年比120%で推移"
        )
        body_text = st.text_area(
            "根拠・補足（1行ごと）",
            value="\n".join(slide.get("body", [])) if isinstance(slide.get("body"), list) else slide.get("body", ""),
            height=80,
            key=f"keybody_{index}"
        )
        result["body"] = [line.strip() for line in body_text.split("\n") if line.strip()]

    elif slide_type == "comparison":
        c_left, c_right = st.columns(2)
        with c_left:
            result["before_title"] = st.text_input(
                "Beforeラベル",
                value=slide.get("before_title", "Before"),
                key=f"before_title_{index}"
            )
            before_text = st.text_area(
                "Before内容（1行ごと）",
                value="\n".join(slide.get("before_items", [])) if isinstance(slide.get("before_items"), list) else slide.get("before_items", ""),
                height=100,
                key=f"before_items_{index}"
            )
            result["before_items"] = [l.strip() for l in before_text.split("\n") if l.strip()]
        with c_right:
            result["after_title"] = st.text_input(
                "Afterラベル",
                value=slide.get("after_title", "After"),
                key=f"after_title_{index}"
            )
            after_text = st.text_area(
                "After内容（1行ごと）",
                value="\n".join(slide.get("after_items", [])) if isinstance(slide.get("after_items"), list) else slide.get("after_items", ""),
                height=100,
                key=f"after_items_{index}"
            )
            result["after_items"] = [l.strip() for l in after_text.split("\n") if l.strip()]

    elif slide_type == "table":
        st.caption("ヘッダー（カンマ区切り）")
        headers_text = st.text_input(
            "ヘッダー",
            value=", ".join(slide.get("headers", [])),
            key=f"headers_{index}",
            label_visibility="collapsed"
        )
        result["headers"] = [h.strip() for h in headers_text.split(",") if h.strip()]

        st.caption("データ行（1行に1レコード、カンマ区切り）")
        rows_text = st.text_area(
            "データ行",
            value="\n".join([", ".join(row) for row in slide.get("rows", [])]),
            height=100,
            key=f"rows_{index}",
            label_visibility="collapsed"
        )
        result["rows"] = [
            [c.strip() for c in row.split(",")]
            for row in rows_text.split("\n") if row.strip()
        ]

    elif slide_type == "chart":
        chart_types = {"棒グラフ": "bar", "折れ線グラフ": "line", "円グラフ": "pie"}
        selected_chart = st.selectbox(
            "グラフタイプ",
            list(chart_types.keys()),
            key=f"chart_type_{index}"
        )
        result["chart_type"] = chart_types[selected_chart]

        categories_text = st.text_input(
            "カテゴリ（カンマ区切り）",
            value=", ".join(slide.get("categories", [])),
            key=f"categories_{index}"
        )
        result["categories"] = [c.strip() for c in categories_text.split(",") if c.strip()]

        st.caption("データ系列（系列名: 値1, 値2, ... を1行ずつ）")
        series_default = ""
        for s in slide.get("series", []):
            vals = ", ".join(str(v) for v in s.get("values", []))
            series_default += f"{s.get('name', '')}: {vals}\n"
        series_text = st.text_area(
            "データ系列",
            value=series_default.strip(),
            height=80,
            key=f"series_{index}",
            label_visibility="collapsed"
        )
        series_list = []
        for line in series_text.split("\n"):
            if ":" in line:
                name, vals_str = line.split(":", 1)
                vals = []
                for v in vals_str.split(","):
                    v = v.strip()
                    try:
                        vals.append(float(v))
                    except ValueError:
                        pass
                series_list.append({"name": name.strip(), "values": vals})
        result["series"] = series_list

    elif slide_type == "section":
        pass  # タイトルのみ

    elif slide_type == "image":
        result["image_path"] = st.text_input(
            "画像ファイルパス",
            value=slide.get("image_path", ""),
            key=f"image_{index}"
        )

    return result


def _escape_html(text):
    """HTMLエスケープ"""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def render_slide_preview_html(slide_data, slide_num, scheme):
    """スライドデータからHTMLプレビューを生成"""
    stype = slide_data.get("type", "content")
    title = _escape_html(slide_data.get("title", ""))
    accent = scheme.get("accent_color", "4472C4")
    heading_c = scheme.get("heading_color", "1F4E79")
    text_c = scheme.get("text_color", "333333")
    bg_c = scheme.get("bg_color", "FFFFFF")
    accent_light = scheme.get("accent_light", "D6E4F0")
    section_bg = scheme.get("section_bg", accent)
    section_text = scheme.get("section_text", "FFFFFF")
    type_label = SLIDE_TYPE_LABELS.get(stype, stype)

    # タイトルスライド
    if stype == "title":
        subtitle = _escape_html(slide_data.get("subtitle", ""))
        return f"""
        <div class="slide-preview" style="background:#{bg_c};">
            <span class="slide-num">{slide_num}</span>
            <span class="slide-badge">タイトル</span>
            <div class="slide-preview-title-slide">
                <div style="width:40%;height:2px;background:#{accent};margin-bottom:6px;"></div>
                <div class="main-title" style="color:#{heading_c};">{title}</div>
                <div class="sub-title" style="color:#{accent};">{subtitle}</div>
            </div>
        </div>"""

    # セクション区切り
    if stype == "section":
        return f"""
        <div class="slide-preview" style="background:#{section_bg};">
            <span class="slide-num" style="color:rgba(255,255,255,0.5);">{slide_num}</span>
            <span class="slide-badge" style="background:rgba(255,255,255,0.2);color:rgba(255,255,255,0.7);">セクション</span>
            <div class="slide-preview-section" style="color:#{section_text};">{title}</div>
        </div>"""

    # コンテンツ（箇条書き）
    if stype == "content":
        body = slide_data.get("body", [])
        if isinstance(body, str):
            body = [body]
        items_html = "".join(
            f'<div class="slide-body-item">&bull; {_escape_html(item)}</div>' for item in body[:5]
        )
        if len(body) > 5:
            items_html += '<div class="slide-body-item" style="color:#aaa;">...</div>'
        return f"""
        <div class="slide-preview" style="background:#{bg_c};">
            <span class="slide-num">{slide_num}</span>
            <span class="slide-badge">{type_label}</span>
            <div class="slide-title" style="color:#{heading_c};border-bottom:2px solid #{accent};">{title}</div>
            <div class="slide-body" style="color:#{text_c};">{items_html}</div>
        </div>"""

    # 2カラム
    if stype == "two_column":
        lt = _escape_html(slide_data.get("left_title", ""))
        rt = _escape_html(slide_data.get("right_title", ""))
        lb = slide_data.get("left_body", [])
        rb = slide_data.get("right_body", [])
        left_items = "".join(f'<div class="slide-body-item">&bull; {_escape_html(i)}</div>' for i in lb[:3])
        right_items = "".join(f'<div class="slide-body-item">&bull; {_escape_html(i)}</div>' for i in rb[:3])
        return f"""
        <div class="slide-preview" style="background:#{bg_c};">
            <span class="slide-num">{slide_num}</span>
            <span class="slide-badge">2カラム</span>
            <div class="slide-title" style="color:#{heading_c};border-bottom:2px solid #{accent};">{title}</div>
            <div class="two-col">
                <div class="col-half">
                    <div class="col-title" style="color:#{accent};">{lt}</div>
                    <div class="slide-body" style="color:#{text_c};">{left_items}</div>
                </div>
                <div style="width:1px;background:#{accent_light};"></div>
                <div class="col-half">
                    <div class="col-title" style="color:#{accent};">{rt}</div>
                    <div class="slide-body" style="color:#{text_c};">{right_items}</div>
                </div>
            </div>
        </div>"""

    # キーメッセージ
    if stype == "key_message":
        msg = _escape_html(slide_data.get("message", ""))
        body = slide_data.get("body", [])
        items_html = "".join(f'<div class="slide-body-item">&bull; {_escape_html(i)}</div>' for i in body[:3])
        return f"""
        <div class="slide-preview" style="background:#{bg_c};">
            <span class="slide-num">{slide_num}</span>
            <span class="slide-badge">Key Message</span>
            <div class="slide-title" style="color:#{heading_c};border-bottom:2px solid #{accent};">{title}</div>
            <div class="key-msg" style="color:#{accent};">{msg}</div>
            <div style="width:50%;height:1px;background:#{accent_light};margin:0 auto 4px;"></div>
            <div class="slide-body" style="color:#{text_c};text-align:center;">{items_html}</div>
        </div>"""

    # Before/After比較
    if stype == "comparison":
        bt = _escape_html(slide_data.get("before_title", "Before"))
        at = _escape_html(slide_data.get("after_title", "After"))
        bi = slide_data.get("before_items", [])
        ai = slide_data.get("after_items", [])
        b_items = "".join(f'<div class="slide-body-item">&bull; {_escape_html(i)}</div>' for i in bi[:3])
        a_items = "".join(f'<div class="slide-body-item">&bull; {_escape_html(i)}</div>' for i in ai[:3])
        return f"""
        <div class="slide-preview" style="background:#{bg_c};">
            <span class="slide-num">{slide_num}</span>
            <span class="slide-badge">比較</span>
            <div class="slide-title" style="color:#{heading_c};border-bottom:2px solid #{accent};">{title}</div>
            <div class="two-col">
                <div class="col-half">
                    <div class="comp-label" style="background:#E74C3C;">{bt}</div>
                    <div class="slide-body" style="color:#{text_c};">{b_items}</div>
                </div>
                <div class="comp-arrow" style="color:#{accent};">→</div>
                <div class="col-half">
                    <div class="comp-label" style="background:#27AE60;">{at}</div>
                    <div class="slide-body" style="color:#{text_c};">{a_items}</div>
                </div>
            </div>
        </div>"""

    # テーブル
    if stype == "table":
        headers = slide_data.get("headers", [])
        rows = slide_data.get("rows", [])
        header_bg = scheme.get("table_header_bg", accent)
        header_text = scheme.get("table_header_text", "FFFFFF")
        if headers:
            hdr = "".join(f'<th style="background:#{header_bg};color:#{header_text};padding:1px 3px;font-size:7px;">{_escape_html(h)}</th>' for h in headers[:5])
            row_html = ""
            for r_idx, row in enumerate(rows[:3]):
                cells = "".join(f'<td style="padding:1px 3px;font-size:7px;">{_escape_html(str(c))}</td>' for c in row[:5])
                row_html += f"<tr>{cells}</tr>"
            return f"""
            <div class="slide-preview" style="background:#{bg_c};">
                <span class="slide-num">{slide_num}</span>
                <span class="slide-badge">テーブル</span>
                <div class="slide-title" style="color:#{heading_c};border-bottom:2px solid #{accent};">{title}</div>
                <table style="width:100%;border-collapse:collapse;margin-top:2px;"><tr>{hdr}</tr>{row_html}</table>
            </div>"""

    # グラフ
    if stype == "chart":
        chart_type = slide_data.get("chart_type", "bar")
        chart_label = {"bar": "棒", "line": "折線", "pie": "円"}.get(chart_type, chart_type)
        return f"""
        <div class="slide-preview" style="background:#{bg_c};">
            <span class="slide-num">{slide_num}</span>
            <span class="slide-badge">グラフ</span>
            <div class="slide-title" style="color:#{heading_c};border-bottom:2px solid #{accent};">{title}</div>
            <div style="text-align:center;margin-top:8px;">
                <div style="display:inline-block;width:60%;height:40px;border:1px dashed #{accent};border-radius:4px;display:flex;align-items:center;justify-content:center;color:#{accent};font-size:9px;">{chart_label}グラフ</div>
            </div>
        </div>"""

    # 画像
    if stype == "image":
        img_path = _escape_html(slide_data.get("image_path", ""))
        return f"""
        <div class="slide-preview" style="background:#{bg_c};">
            <span class="slide-num">{slide_num}</span>
            <span class="slide-badge">画像</span>
            <div class="slide-title" style="color:#{heading_c};border-bottom:2px solid #{accent};">{title}</div>
            <div style="text-align:center;margin-top:8px;">
                <div style="display:inline-block;width:50%;height:35px;background:#f0f0f0;border:1px dashed #ccc;border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:8px;color:#999;">🖼 {img_path}</div>
            </div>
        </div>"""

    # フォールバック
    return f"""
    <div class="slide-preview" style="background:#{bg_c};">
        <span class="slide-num">{slide_num}</span>
        <span class="slide-badge">{type_label}</span>
        <div class="slide-title" style="color:#{heading_c};">{title}</div>
    </div>"""


def main():
    st.title("📊 資料メーカー")
    st.caption("PowerPointプレゼンテーションを簡単に自動生成")

    # サイドバー: テンプレート＆カラー選択＆プレビュー
    with st.sidebar:
        st.header("⚙️ 設定")

        # テンプレート選択
        templates = load_templates()
        template_options = {v.get("description", k): k for k, v in templates.items()}
        selected_template_label = st.selectbox(
            "テンプレート",
            list(template_options.keys()),
        )
        selected_template = template_options[selected_template_label]

        st.markdown("---")

        # カラースキーム選択
        st.subheader("🎨 カラースキーム")
        color_schemes = load_color_schemes()

        color_options = {}
        for name, scheme in color_schemes.items():
            accent = scheme.get("accent_color", "4472C4")
            desc = scheme.get("description", name)
            color_options[f"#{accent} {desc}"] = name

        selected_color_label = st.radio(
            "カラーを選択",
            list(color_options.keys()),
            format_func=lambda x: x.split(" ", 1)[1] if " " in x else x,
            label_visibility="collapsed"
        )
        selected_color = color_options[selected_color_label]

        scheme = color_schemes[selected_color]
        accent = scheme.get("accent_color", "4472C4")
        secondary = scheme.get("secondary_color", accent)
        bg = scheme.get("bg_color", "FFFFFF")
        text_c = scheme.get("text_color", "333333")

        st.markdown(
            f"""
            <div style="display: flex; gap: 4px; margin-top: 4px;">
                <div class="color-swatch" style="background-color: #{accent};" title="アクセント"></div>
                <div class="color-swatch" style="background-color: #{secondary};" title="セカンダリ"></div>
                <div class="color-swatch" style="background-color: #{bg}; border-color: #ccc;" title="背景"></div>
                <div class="color-swatch" style="background-color: #{text_c};" title="テキスト"></div>
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown("---")

        # === サイドバー内プレビュー ===
        st.subheader("📋 スライドプレビュー")

    # メインエリア: タブ切り替え
    tab_manual, tab_file, tab_json = st.tabs(["✏️ 手動入力", "📁 ファイル読み込み", "📄 JSON直接入力"])

    # === タブ1: 手動入力 ===
    with tab_manual:
        pres_title = st.text_input("プレゼンタイトル", value="", placeholder="例: Q4 業績報告")
        col_sub, col_author = st.columns(2)
        with col_sub:
            pres_subtitle = st.text_input("サブタイトル", value="", placeholder="例: 2025年度第4四半期")
        with col_author:
            pres_author = st.text_input("作成者", value="", placeholder="例: 営業部")

        st.markdown("---")

        # スライド管理
        st.subheader("スライド構成")

        if "slides" not in st.session_state:
            st.session_state.slides = [{"type": "content", "title": "", "body": []}]

        # スライド管理ボタン
        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            if st.button("+ 追加"):
                st.session_state.slides.append({"type": "content", "title": "", "body": []})
                st.experimental_rerun()
        with col_b:
            if len(st.session_state.slides) > 1:
                if st.button("- 削除"):
                    st.session_state.slides.pop()
                    st.experimental_rerun()
        with col_c:
            if st.button("複製"):
                import copy
                if st.session_state.slides:
                    st.session_state.slides.append(copy.deepcopy(st.session_state.slides[-1]))
                    st.experimental_rerun()
        with col_d:
            if len(st.session_state.slides) > 1:
                move_idx = st.number_input(
                    "移動対象",
                    min_value=1,
                    max_value=len(st.session_state.slides),
                    value=len(st.session_state.slides),
                    key="move_idx",
                    help="移動するスライド番号"
                )

        # 並び替えボタン
        if len(st.session_state.slides) > 1:
            col_up, col_down, _ = st.columns([1, 1, 4])
            with col_up:
                if st.button("↑ 上へ"):
                    idx = st.session_state.get("move_idx", len(st.session_state.slides)) - 1
                    if 0 < idx < len(st.session_state.slides):
                        st.session_state.slides[idx - 1], st.session_state.slides[idx] = \
                            st.session_state.slides[idx], st.session_state.slides[idx - 1]
                        st.experimental_rerun()
            with col_down:
                if st.button("↓ 下へ"):
                    idx = st.session_state.get("move_idx", len(st.session_state.slides)) - 1
                    if 0 <= idx < len(st.session_state.slides) - 1:
                        st.session_state.slides[idx], st.session_state.slides[idx + 1] = \
                            st.session_state.slides[idx + 1], st.session_state.slides[idx]
                        st.experimental_rerun()

        # 各スライドの編集
        slides_data = []
        for i in range(len(st.session_state.slides)):
            slide_label = SLIDE_TYPE_LABELS.get(
                st.session_state.slides[i].get("type", "content"), "コンテンツ"
            )
            with st.expander(f"スライド {i + 1} - {slide_label}", expanded=(i < 2)):
                slide_result = render_slide_editor(i, st.session_state.slides[i])
                slides_data.append(slide_result)
                st.session_state.slides[i] = slide_result

        st.markdown("---")

        # 生成ボタン
        if st.button("🚀 PowerPointを生成", type="primary", use_container_width=True, key="gen_manual"):
            if not pres_title:
                st.warning("プレゼンタイトルを入力してください。")
            else:
                data = build_data_from_form(pres_title, pres_subtitle, pres_author, slides_data)
                _generate_and_download(data, selected_template, selected_color, pres_title)

        # サイドバーにプレビュー描画
        with st.sidebar:
            title_slide = {
                "type": "title",
                "title": pres_title or "タイトル",
                "subtitle": pres_subtitle
            }
            st.markdown(
                render_slide_preview_html(title_slide, 1, scheme),
                unsafe_allow_html=True
            )
            for i, slide in enumerate(st.session_state.get("slides", [])):
                st.markdown(
                    render_slide_preview_html(slide, i + 2, scheme),
                    unsafe_allow_html=True
                )

    # === タブ2: ファイル読み込み ===
    with tab_file:
        st.markdown("JSON / CSV / テキストファイルをアップロードして自動変換")
        uploaded = st.file_uploader(
            "ファイルを選択",
            type=["json", "csv", "txt", "md"],
            key="file_upload"
        )

        if uploaded:
            import tempfile
            ext = os.path.splitext(uploaded.name)[1]
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False, mode="wb") as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = tmp.name

            try:
                data = load_from_file(tmp_path)
                st.success(f"読み込み完了: {len(data.get('slides', []))} スライド")

                with st.expander("データプレビュー", expanded=False):
                    st.json(data)

                # スライドプレビューをサイドバーに表示
                with st.sidebar:
                    for i, slide in enumerate(data.get("slides", [])):
                        st.markdown(
                            render_slide_preview_html(slide, i + 1, scheme),
                            unsafe_allow_html=True
                        )

                file_title = data.get("title", uploaded.name)

                if st.button("🚀 PowerPointを生成", type="primary", use_container_width=True, key="gen_file"):
                    _generate_and_download(data, selected_template, selected_color, file_title)
            except Exception as e:
                st.error(f"ファイルの読み込みに失敗しました: {e}")
            finally:
                os.unlink(tmp_path)

    # === タブ3: JSON直接入力 ===
    with tab_json:
        st.markdown("JSONを直接編集してスライドを定義")

        sample_json = json.dumps({
            "title": "サンプルプレゼンテーション",
            "slides": [
                {"type": "title", "title": "サンプルプレゼンテーション", "subtitle": "自動生成デモ"},
                {"type": "key_message", "title": "結論", "message": "売上は前年比120%で成長", "body": ["EC事業が牽引", "新規顧客の獲得が好調"]},
                {"type": "two_column", "title": "分析", "left_title": "強み", "left_body": ["ブランド力", "技術力"], "right_title": "課題", "right_body": ["人材確保", "コスト最適化"]},
                {"type": "content", "title": "アクションプラン", "body": ["採用強化", "DX推進", "海外展開"]},
            ]
        }, ensure_ascii=False, indent=2)

        json_input = st.text_area(
            "JSONデータ",
            value=sample_json,
            height=400,
            key="json_input"
        )

        # JSONプレビューをサイドバーに表示
        try:
            preview_data = json.loads(json_input)
            with st.sidebar:
                for i, slide in enumerate(preview_data.get("slides", [])):
                    st.markdown(
                        render_slide_preview_html(slide, i + 1, scheme),
                        unsafe_allow_html=True
                    )
        except (json.JSONDecodeError, TypeError):
            pass

        if st.button("🚀 PowerPointを生成", type="primary", use_container_width=True, key="gen_json"):
            try:
                data = json.loads(json_input)
                if "slides" not in data:
                    st.warning("'slides' キーが見つかりません。")
                else:
                    json_title = data.get("title", "プレゼンテーション")
                    _generate_and_download(data, selected_template, selected_color, json_title)
            except json.JSONDecodeError as e:
                st.error(f"JSONの解析に失敗しました: {e}")


def _generate_and_download(data, template_name, color_name, title):
    """プレゼンテーションを生成してダウンロードボタンを表示"""
    with st.spinner("生成中..."):
        builder = PresentationBuilder(
            template_name=template_name,
            color_name=color_name
        )
        builder.build(data)

        # メモリ上にファイルを保存
        buffer = io.BytesIO()
        builder.prs.save(buffer)
        buffer.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{title}_{timestamp}.pptx"

    st.success(f"生成完了！ ({len(data.get('slides', []))} スライド)")

    st.download_button(
        label="📥 ダウンロード",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        type="primary",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
