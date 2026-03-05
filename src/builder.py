"""
プレゼンテーション組み立てのコアロジック
"""
import json
import os

from pptx import Presentation
from pptx.util import Inches

import config
from src.slide_factory import SLIDE_TYPES


class PresentationBuilder:
    """プレゼンテーション組み立てクラス"""

    def __init__(self, template_name=None, style_config=None):
        """
        Args:
            template_name: テンプレート名（templates/配下のJSONファイル名）
            style_config: スタイル設定（省略時はconfig.pyのデフォルト）
        """
        self.style = dict(config.STYLE)
        self.template_config = None
        self.prs = None

        # テンプレート読み込み
        if template_name:
            self._load_template(template_name)

        # スタイル上書き
        if style_config:
            self.style.update(style_config)

    def _load_template(self, template_name):
        """テンプレートJSONを読み込む"""
        template_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            config.TEMPLATE_DIR,
            f"{template_name}.json"
        )

        if not os.path.exists(template_path):
            print(f"警告: テンプレート '{template_name}' が見つかりません。デフォルト設定を使用します。")
            return

        with open(template_path, "r", encoding="utf-8") as f:
            self.template_config = json.load(f)

        # テンプレートのスタイルを適用
        if "style" in self.template_config:
            self.style.update(self.template_config["style"])

    def build(self, data):
        """データからプレゼンテーションを組み立てる

        Args:
            data: 統一形式のdict
                {
                    "title": "...",
                    "slides": [{"type": "...", ...}, ...]
                }

        Returns:
            pptx.Presentation オブジェクト
        """
        self.prs = Presentation()

        # スライドサイズ設定（ワイドスクリーン 16:9）
        width = config.SLIDE_WIDTH_INCHES
        height = config.SLIDE_HEIGHT_INCHES
        if self.template_config:
            width = self.template_config.get("slide_width_inches", width)
            height = self.template_config.get("slide_height_inches", height)
        self.prs.slide_width = Inches(width)
        self.prs.slide_height = Inches(height)

        # タイトルスライドのデータ補完
        slides = data.get("slides", [])
        if slides and slides[0].get("type") == "title":
            # トップレベルの情報をタイトルスライドに引き継ぐ
            if "author" not in slides[0] and "author" in data:
                slides[0]["author"] = data["author"]
            if "date" not in slides[0] and "date" in data:
                slides[0]["date"] = data["date"]

        # 各スライドを生成
        for slide_data in slides:
            slide_type = slide_data.get("type", "content")
            factory_fn = SLIDE_TYPES.get(slide_type)

            if factory_fn:
                factory_fn(self.prs, slide_data, self.style)
            else:
                print(f"警告: 未対応のスライドタイプ '{slide_type}' をスキップしました。")

        return self.prs

    def save(self, filepath):
        """ファイルに保存

        Args:
            filepath: 保存先パス

        Returns:
            保存先の絶対パス
        """
        if self.prs is None:
            raise RuntimeError("build()を先に実行してください。")

        # 出力ディレクトリ作成
        output_dir = os.path.dirname(filepath)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        self.prs.save(filepath)
        return os.path.abspath(filepath)

    @staticmethod
    def list_templates():
        """利用可能なテンプレート一覧を返す"""
        template_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            config.TEMPLATE_DIR
        )

        templates = []
        if not os.path.exists(template_dir):
            return templates

        for filename in sorted(os.listdir(template_dir)):
            if filename.endswith(".json"):
                filepath = os.path.join(template_dir, filename)
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                templates.append({
                    "name": os.path.splitext(filename)[0],
                    "description": data.get("description", ""),
                })

        return templates
