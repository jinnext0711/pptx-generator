"""
スライド種別ごとの生成関数
"""
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData

from src.style import apply_font, hex_to_rgb, COLORS, LAYOUT


def add_title_slide(prs, data, style):
    """タイトルスライドを追加"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白レイアウト

    # 背景にアクセントカラーの帯
    accent = hex_to_rgb(style.get("accent_color", "4472C4"))

    # タイトルテキストボックス
    left = Inches(1.0)
    top = Inches(2.5)
    width = Inches(11.3)
    height = Inches(1.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = data.get("title", "")
    apply_font(run, style.get("font_name", "Meiryo"),
               style.get("title_size_pt", 28), COLORS["dark"], bold=True)

    # サブタイトル
    subtitle = data.get("subtitle", "")
    if subtitle:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        p2.space_before = Pt(12)
        run2 = p2.add_run()
        run2.text = subtitle
        apply_font(run2, style.get("font_name", "Meiryo"),
                   style.get("subtitle_size_pt", 18), accent)

    # 日付・作成者
    author = data.get("author", "")
    date = data.get("date", "")
    if author or date:
        info_text = " | ".join(filter(None, [author, date]))
        p3 = tf.add_paragraph()
        p3.alignment = PP_ALIGN.CENTER
        p3.space_before = Pt(24)
        run3 = p3.add_run()
        run3.text = info_text
        apply_font(run3, style.get("font_name", "Meiryo"), 12, COLORS["dark"])

    # アクセントライン
    line_left = Inches(4.0)
    line_top = Inches(2.2)
    line_width = Inches(5.3)
    line = slide.shapes.add_shape(
        1, line_left, line_top, line_width, Pt(4)  # 1 = Rectangle
    )
    line.fill.solid()
    line.fill.fore_color.rgb = accent
    line.line.fill.background()


def add_content_slide(prs, data, style):
    """箇条書きコンテンツスライドを追加"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = hex_to_rgb(style.get("accent_color", "4472C4"))

    # スライドタイトル
    _add_slide_title(slide, data.get("title", ""), style, accent)

    # 本文（箇条書き）
    body = data.get("body", [])
    if isinstance(body, str):
        body = [body]

    left = LAYOUT["margin_left"]
    top = LAYOUT["margin_top"]
    width = LAYOUT["content_width"]
    height = LAYOUT["content_height"]
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    font_name = style.get("font_name", "Meiryo")
    body_size = style.get("body_size_pt", 14)

    for i, item in enumerate(body):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(8)
        p.level = 0

        # 箇条書きマーカー
        run = p.add_run()
        run.text = f"● {item}"
        apply_font(run, font_name, body_size, COLORS["dark"])


def add_table_slide(prs, data, style):
    """テーブルスライドを追加"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = hex_to_rgb(style.get("accent_color", "4472C4"))

    _add_slide_title(slide, data.get("title", ""), style, accent)

    headers = data.get("headers", [])
    rows = data.get("rows", [])
    if not headers and not rows:
        return

    num_rows = len(rows) + 1  # ヘッダー行含む
    num_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if num_cols == 0:
        return

    left = LAYOUT["margin_left"]
    top = LAYOUT["margin_top"]
    width = LAYOUT["content_width"]
    height = Inches(0.4) * num_rows

    table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = table_shape.table

    font_name = style.get("font_name", "Meiryo")

    # ヘッダー行
    for j, header in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = str(header)
        # ヘッダーの背景色
        cell.fill.solid()
        cell.fill.fore_color.rgb = accent
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                apply_font(run, font_name, 12, COLORS["white"], bold=True)

    # データ行
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            if j >= num_cols:
                break
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            # 交互の行色
            if i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = COLORS["light_gray"]
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    apply_font(run, font_name, 11, COLORS["dark"])


def add_chart_slide(prs, data, style):
    """グラフスライドを追加"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = hex_to_rgb(style.get("accent_color", "4472C4"))

    _add_slide_title(slide, data.get("title", ""), style, accent)

    chart_type_str = data.get("chart_type", "bar")
    categories = data.get("categories", [])
    series_list = data.get("series", [])

    if not categories or not series_list:
        return

    # チャートタイプのマッピング
    chart_type_map = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE,
        "pie": XL_CHART_TYPE.PIE,
    }
    chart_type = chart_type_map.get(chart_type_str, XL_CHART_TYPE.COLUMN_CLUSTERED)

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for s in series_list:
        chart_data.add_series(s.get("name", ""), s.get("values", []))

    left = LAYOUT["margin_left"]
    top = LAYOUT["margin_top"]
    width = LAYOUT["content_width"]
    height = LAYOUT["content_height"]

    chart_shape = slide.shapes.add_chart(chart_type, left, top, width, height, chart_data)

    # グラフのフォント設定
    chart = chart_shape.chart
    chart.has_legend = len(series_list) > 1
    if chart.has_legend:
        chart.legend.include_in_layout = False


def add_image_slide(prs, data, style):
    """画像スライドを追加"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = hex_to_rgb(style.get("accent_color", "4472C4"))

    _add_slide_title(slide, data.get("title", ""), style, accent)

    image_path = data.get("image_path", "")
    if not image_path:
        return

    try:
        # 画像を中央配置
        left = Inches(2.0)
        top = LAYOUT["margin_top"]
        height = LAYOUT["content_height"]
        slide.shapes.add_picture(image_path, left, top, height=height)
    except Exception:
        # 画像が見つからない場合はテキストで代替
        txBox = slide.shapes.add_textbox(
            LAYOUT["margin_left"], LAYOUT["margin_top"],
            LAYOUT["content_width"], Inches(1.0)
        )
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f"[画像が見つかりません: {image_path}]"
        apply_font(run, style.get("font_name", "Meiryo"), 14, COLORS["dark"])


def add_section_slide(prs, data, style):
    """セクション区切りスライドを追加"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    accent = hex_to_rgb(style.get("accent_color", "4472C4"))

    # 背景塗りつぶし
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = accent

    # セクションタイトル（中央配置、白文字）
    left = Inches(1.0)
    top = Inches(2.8)
    width = Inches(11.3)
    height = Inches(1.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = data.get("title", "")
    apply_font(run, style.get("font_name", "Meiryo"),
               style.get("heading_size_pt", 24), COLORS["white"], bold=True)


def _add_slide_title(slide, title_text, style, accent_color):
    """各スライドのタイトルを追加するヘルパー"""
    if not title_text:
        return

    left = LAYOUT["margin_left"]
    top = LAYOUT["title_top"]
    width = LAYOUT["content_width"]
    height = LAYOUT["title_height"]

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title_text
    apply_font(run, style.get("font_name", "Meiryo"),
               style.get("heading_size_pt", 24), COLORS["dark"], bold=True)

    # タイトル下のアクセントライン
    line = slide.shapes.add_shape(
        1, left, top + height - Pt(4), Inches(3.0), Pt(3)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = accent_color
    line.line.fill.background()


# スライドタイプ名と関数のマッピング
SLIDE_TYPES = {
    "title": add_title_slide,
    "content": add_content_slide,
    "table": add_table_slide,
    "chart": add_chart_slide,
    "image": add_image_slide,
    "section": add_section_slide,
}
