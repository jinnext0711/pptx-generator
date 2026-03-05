"""
pptx-generator - PowerPointプレゼンテーション自動生成ツール
CLIエントリポイント
"""
import argparse
import os
import sys
from datetime import datetime

import config
from src.builder import PresentationBuilder
from src.data_loader import load_from_file, load_from_url


def main():
    parser = argparse.ArgumentParser(
        description="PowerPointプレゼンテーション自動生成ツール"
    )
    parser.add_argument(
        "--input", "-i",
        help="入力データファイルパス（JSON, CSV, テキスト）"
    )
    parser.add_argument(
        "--template", "-t",
        default="default",
        help="テンプレート名（デフォルト: default）"
    )
    parser.add_argument(
        "--output", "-o",
        help="出力ファイルパス（デフォルト: output/YYYYMMDD_HHMMSS.pptx）"
    )
    parser.add_argument(
        "--title",
        help="プレゼンテーションタイトル（入力データのタイトルを上書き）"
    )
    parser.add_argument(
        "--url",
        help="データ取得元URL"
    )
    parser.add_argument(
        "--list-templates",
        action="store_true",
        help="利用可能なテンプレート一覧を表示"
    )

    args = parser.parse_args()

    # テンプレート一覧表示
    if args.list_templates:
        templates = PresentationBuilder.list_templates()
        if not templates:
            print("テンプレートが見つかりません。")
        else:
            print("利用可能なテンプレート:")
            for t in templates:
                print(f"  {t['name']:15s} - {t['description']}")
        return

    # 入力データの取得
    if args.input:
        if not os.path.exists(args.input):
            print(f"エラー: ファイルが見つかりません: {args.input}")
            sys.exit(1)
        print(f"ファイルを読み込んでいます: {args.input}")
        data = load_from_file(args.input)
    elif args.url:
        print(f"URLからデータを取得しています: {args.url}")
        data = load_from_url(args.url)
    else:
        print("エラー: --input または --url を指定してください。")
        parser.print_help()
        sys.exit(1)

    # タイトル上書き
    if args.title:
        data["title"] = args.title
        # タイトルスライドも更新
        if data.get("slides") and data["slides"][0].get("type") == "title":
            data["slides"][0]["title"] = args.title

    # 出力パスの決定
    if args.output:
        output_path = args.output
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(config.OUTPUT["dir"], f"{timestamp}.pptx")

    # プレゼンテーション生成
    print(f"テンプレート '{args.template}' でプレゼンテーションを生成しています...")
    builder = PresentationBuilder(template_name=args.template)
    builder.build(data)

    saved_path = builder.save(output_path)
    print(f"生成完了: {saved_path}")


if __name__ == "__main__":
    main()
