"""
builderとslide_factoryのテスト
"""
import json
import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.builder import PresentationBuilder
from src.data_loader import load_from_file, load_from_text


class TestPresentationBuilder:
    """PresentationBuilderのテスト"""

    def test_build_basic(self):
        """基本的なプレゼンテーション生成"""
        data = {
            "title": "テスト",
            "slides": [
                {"type": "title", "title": "テストタイトル", "subtitle": "サブタイトル"},
                {"type": "content", "title": "内容", "body": ["項目1", "項目2"]},
            ]
        }
        builder = PresentationBuilder()
        prs = builder.build(data)
        assert len(prs.slides) == 2

    def test_build_all_slide_types(self):
        """全スライドタイプの生成"""
        data = {
            "title": "全タイプテスト",
            "slides": [
                {"type": "title", "title": "タイトル"},
                {"type": "content", "title": "コンテンツ", "body": ["A", "B"]},
                {"type": "section", "title": "セクション"},
                {"type": "table", "title": "テーブル",
                 "headers": ["A", "B"], "rows": [["1", "2"]]},
                {"type": "chart", "title": "グラフ", "chart_type": "bar",
                 "categories": ["Q1", "Q2"], "series": [{"name": "S1", "values": [10, 20]}]},
            ]
        }
        builder = PresentationBuilder()
        prs = builder.build(data)
        assert len(prs.slides) == 5

    def test_save_file(self):
        """ファイル保存"""
        data = {
            "title": "保存テスト",
            "slides": [{"type": "title", "title": "テスト"}]
        }
        builder = PresentationBuilder()
        builder.build(data)

        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            path = f.name

        try:
            saved = builder.save(path)
            assert os.path.exists(saved)
            assert os.path.getsize(saved) > 0
        finally:
            os.unlink(path)

    def test_template_loading(self):
        """テンプレート読み込み（default_colorでblueが適用される）"""
        builder = PresentationBuilder(template_name="default")
        assert builder.style["accent_color"] == "4472C4"

    def test_color_scheme_loading(self):
        """カラースキーム指定"""
        builder = PresentationBuilder(color_name="red")
        assert builder.style["accent_color"] == "C0392B"

    def test_color_overrides_template_default(self):
        """CLI指定のカラーがテンプレートのdefault_colorを上書き"""
        # pitchテンプレートのdefault_colorはredだが、greenを指定
        builder = PresentationBuilder(template_name="pitch", color_name="green")
        assert builder.style["accent_color"] == "27AE60"

    def test_template_default_color(self):
        """テンプレートのdefault_colorが適用される"""
        # pitchテンプレートはdefault_color: "red"
        builder = PresentationBuilder(template_name="pitch")
        assert builder.style["accent_color"] == "C0392B"

    def test_dark_color_bg(self):
        """ダークカラースキームの背景色設定"""
        builder = PresentationBuilder(color_name="dark")
        assert builder.style["bg_color"] == "2C3E50"

    def test_list_templates(self):
        """テンプレート一覧"""
        templates = PresentationBuilder.list_templates()
        names = [t["name"] for t in templates]
        assert "default" in names
        assert "report" in names
        assert "pitch" in names

    def test_list_colors(self):
        """カラースキーム一覧"""
        colors = PresentationBuilder.list_colors()
        names = [c["name"] for c in colors]
        assert "blue" in names
        assert "red" in names
        assert "dark" in names
        assert "green" in names

    def test_two_column_slide(self):
        """2カラムスライドの生成"""
        data = {
            "title": "2カラムテスト",
            "slides": [
                {"type": "two_column", "title": "比較分析",
                 "left_title": "現状", "left_body": ["課題1", "課題2"],
                 "right_title": "提案", "right_body": ["解決策1", "解決策2"]},
            ]
        }
        builder = PresentationBuilder()
        prs = builder.build(data)
        assert len(prs.slides) == 1

    def test_key_message_slide(self):
        """キーメッセージスライドの生成"""
        data = {
            "title": "KMテスト",
            "slides": [
                {"type": "key_message", "title": "結論",
                 "message": "売上は前年比120%",
                 "body": ["EC事業が牽引", "新規顧客の獲得"]},
            ]
        }
        builder = PresentationBuilder()
        prs = builder.build(data)
        assert len(prs.slides) == 1

    def test_comparison_slide(self):
        """Before/After比較スライドの生成"""
        data = {
            "title": "比較テスト",
            "slides": [
                {"type": "comparison", "title": "改善効果",
                 "before_title": "Before", "before_items": ["手作業", "3日かかる"],
                 "after_title": "After", "after_items": ["自動化", "30分で完了"]},
            ]
        }
        builder = PresentationBuilder()
        prs = builder.build(data)
        assert len(prs.slides) == 1

    def test_all_new_slide_types_combined(self):
        """新旧スライドタイプを組み合わせた生成"""
        data = {
            "title": "統合テスト",
            "slides": [
                {"type": "title", "title": "タイトル", "subtitle": "サブ"},
                {"type": "key_message", "title": "結論", "message": "成長中"},
                {"type": "two_column", "title": "分析",
                 "left_title": "A", "left_body": ["a1"],
                 "right_title": "B", "right_body": ["b1"]},
                {"type": "comparison", "title": "変化",
                 "before_items": ["旧"], "after_items": ["新"]},
                {"type": "content", "title": "詳細", "body": ["x"]},
                {"type": "table", "title": "表",
                 "headers": ["H"], "rows": [["1"]]},
                {"type": "section", "title": "セクション"},
            ]
        }
        builder = PresentationBuilder(template_name="default", color_name="blue")
        prs = builder.build(data)
        assert len(prs.slides) == 7


class TestDataLoader:
    """data_loaderのテスト"""

    def test_load_json(self):
        """JSON読み込み"""
        data = {
            "title": "JSONテスト",
            "slides": [{"type": "title", "title": "テスト"}]
        }
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8"
        ) as f:
            json.dump(data, f, ensure_ascii=False)
            path = f.name

        try:
            result = load_from_file(path)
            assert result["title"] == "JSONテスト"
            assert len(result["slides"]) == 1
        finally:
            os.unlink(path)

    def test_load_csv(self):
        """CSV読み込み"""
        csv_content = "名前,値\nA,100\nB,200\n"
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".csv", delete=False, encoding="utf-8"
        ) as f:
            f.write(csv_content)
            path = f.name

        try:
            result = load_from_file(path)
            # タイトルスライド + テーブルスライド
            assert len(result["slides"]) == 2
            assert result["slides"][1]["type"] == "table"
            assert result["slides"][1]["headers"] == ["名前", "値"]
        finally:
            os.unlink(path)

    def test_load_text(self):
        """テキスト読み込み"""
        text = """# はじめに
自己紹介です

# アジェンダ
項目A
項目B"""
        result = load_from_text(text, "テストプレゼン")
        # タイトル + 2つのコンテンツスライド
        assert result["title"] == "テストプレゼン"
        assert len(result["slides"]) == 3
        assert result["slides"][1]["title"] == "はじめに"

    def test_unsupported_format(self):
        """未対応形式のエラー"""
        with tempfile.NamedTemporaryFile(suffix=".xyz", delete=False) as f:
            path = f.name

        try:
            with pytest.raises(ValueError):
                load_from_file(path)
        finally:
            os.unlink(path)
