"""
デフォルト設定
"""

# 出力設定
OUTPUT = {
    "dir": "output",
    "default_filename": None,  # Noneの場合はタイムスタンプで自動生成
}

# デフォルトスタイル設定
STYLE = {
    "font_name": "Meiryo",       # 日本語対応フォント
    "font_name_en": "Calibri",   # 英語フォント
    "title_size_pt": 28,
    "subtitle_size_pt": 18,
    "heading_size_pt": 24,
    "body_size_pt": 14,
    "accent_color": "4472C4",    # 青系アクセント
    "bg_color": "FFFFFF",        # 背景色
    "text_color": "333333",      # テキスト色
}

# テンプレートディレクトリ
TEMPLATE_DIR = "templates"

# スライドサイズ（ワイドスクリーン 16:9）
SLIDE_WIDTH_INCHES = 13.333
SLIDE_HEIGHT_INCHES = 7.5
