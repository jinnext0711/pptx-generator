"""
色・フォント・レイアウトの定数とヘルパー関数
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_str):
    """16進数文字列をRGBColorに変換"""
    hex_str = hex_str.lstrip("#")
    return RGBColor(
        int(hex_str[0:2], 16),
        int(hex_str[2:4], 16),
        int(hex_str[4:6], 16),
    )


# プリセットカラー
COLORS = {
    "primary": RGBColor(0x44, 0x72, 0xC4),
    "secondary": RGBColor(0xED, 0x7D, 0x31),
    "dark": RGBColor(0x33, 0x33, 0x33),
    "light_gray": RGBColor(0xF2, 0xF2, 0xF2),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "black": RGBColor(0x00, 0x00, 0x00),
}

# レイアウト定数
LAYOUT = {
    "margin_left": Inches(0.8),
    "margin_top": Inches(1.5),
    "content_width": Inches(11.7),
    "content_height": Inches(5.0),
    "title_top": Inches(0.3),
    "title_height": Inches(1.0),
}


def apply_font(run, font_name, size_pt, color=None, bold=False):
    """テキストランにフォント設定を適用"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        if isinstance(color, str):
            run.font.color.rgb = hex_to_rgb(color)
        else:
            run.font.color.rgb = color
