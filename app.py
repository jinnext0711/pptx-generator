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
    .color-swatch.selected {
        border: 3px solid #333;
        box-shadow: 0 0 4px rgba(0,0,0,0.3);
    }
    /* スライドプレビューカード */
    .slide-card {
        background: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 8px;
    }
    .slide-type-badge {
        background: #e9ecef;
        border-radius: 4px;
        padding: 2px 8px;
        font-size: 0.8em;
        color: #495057;
    }
</style>
""", unsafe_allow_html=True)


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
    slide_types = {
        "コンテンツ（箇条書き）": "content",
        "テーブル": "table",
        "グラフ": "chart",
        "セクション区切り": "section",
        "画像": "image",
    }

    col1, col2 = st.columns([1, 3])

    with col1:
        current_type = slide.get("type", "content")
        # 日本語ラベルに変換
        type_labels = {v: k for k, v in slide_types.items()}
        default_label = type_labels.get(current_type, "コンテンツ（箇条書き）")
        default_idx = list(slide_types.keys()).index(default_label)

        selected_type = st.selectbox(
            "種別",
            list(slide_types.keys()),
            index=default_idx,
            key=f"type_{index}"
        )
        slide_type = slide_types[selected_type]

    with col2:
        slide_title = st.text_input(
            "タイトル",
            value=slide.get("title", ""),
            key=f"title_{index}"
        )

    result = {"type": slide_type, "title": slide_title}

    # タイプ別の入力フィールド
    if slide_type == "content":
        body_text = st.text_area(
            "内容（1行ごとに箇条書き項目）",
            value="\n".join(slide.get("body", [])) if isinstance(slide.get("body"), list) else slide.get("body", ""),
            height=120,
            key=f"body_{index}"
        )
        result["body"] = [line.strip() for line in body_text.split("\n") if line.strip()]

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


def main():
    st.title("📊 資料メーカー")
    st.caption("PowerPointプレゼンテーションを簡単に自動生成")

    # サイドバー: テンプレート＆カラー選択
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

        # カラーパレット表示
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

        # 選択中のカラーのプレビュー
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

    # メインエリア: タブ切り替え
    tab_manual, tab_file, tab_json = st.tabs(["✏️ 手動入力", "📁 ファイル読み込み", "📄 JSON直接入力"])

    # === タブ1: 手動入力 ===
    with tab_manual:
        col_left, col_right = st.columns([2, 1])

        with col_left:
            pres_title = st.text_input("プレゼンタイトル", value="", placeholder="例: Q4 業績報告")
            c1, c2 = st.columns(2)
            with c1:
                pres_subtitle = st.text_input("サブタイトル", value="", placeholder="例: 2025年度第4四半期")
            with c2:
                pres_author = st.text_input("作成者", value="", placeholder="例: 営業部")

        with col_right:
            st.markdown("##### プレビュー情報")
            st.info(
                f"**テンプレート**: {selected_template_label}\n\n"
                f"**カラー**: {scheme.get('description', selected_color)}"
            )

        st.markdown("---")

        # スライド管理
        st.subheader("スライド構成")

        if "slides" not in st.session_state:
            st.session_state.slides = [{"type": "content", "title": "", "body": []}]

        # スライド追加・削除ボタン
        col_add, col_remove = st.columns(2)
        with col_add:
            if st.button("+ スライド追加"):
                st.session_state.slides.append({"type": "content", "title": "", "body": []})
                st.experimental_rerun()
        with col_remove:
            if len(st.session_state.slides) > 1:
                if st.button("- 最後を削除"):
                    st.session_state.slides.pop()
                    st.experimental_rerun()

        # 各スライドの編集
        slides_data = []
        for i in range(len(st.session_state.slides)):
            with st.expander(f"スライド {i + 1}", expanded=(i < 3)):
                slide_result = render_slide_editor(i, st.session_state.slides[i])
                slides_data.append(slide_result)
                # セッションに反映
                st.session_state.slides[i] = slide_result

        st.markdown("---")

        # 生成ボタン
        if st.button("🚀 PowerPointを生成", type="primary", use_container_width=True, key="gen_manual"):
            if not pres_title:
                st.warning("プレゼンタイトルを入力してください。")
            else:
                data = build_data_from_form(pres_title, pres_subtitle, pres_author, slides_data)
                _generate_and_download(data, selected_template, selected_color, pres_title)

    # === タブ2: ファイル読み込み ===
    with tab_file:
        st.markdown("JSON / CSV / テキストファイルをアップロードして自動変換")
        uploaded = st.file_uploader(
            "ファイルを選択",
            type=["json", "csv", "txt", "md"],
            key="file_upload"
        )

        if uploaded:
            # 一時ファイルに保存して読み込み
            import tempfile
            ext = os.path.splitext(uploaded.name)[1]
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False, mode="wb") as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = tmp.name

            try:
                data = load_from_file(tmp_path)
                st.success(f"読み込み完了: {len(data.get('slides', []))} スライド")

                # プレビュー表示
                with st.expander("データプレビュー", expanded=True):
                    st.json(data)

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
            "title": "サンプル",
            "slides": [
                {"type": "title", "title": "サンプル", "subtitle": "テスト"},
                {"type": "content", "title": "項目", "body": ["内容1", "内容2"]},
            ]
        }, ensure_ascii=False, indent=2)

        json_input = st.text_area(
            "JSONデータ",
            value=sample_json,
            height=300,
            key="json_input"
        )

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
