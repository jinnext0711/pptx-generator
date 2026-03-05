# pptx-generator - Claude Code プロジェクト設定

## 1. プロジェクト概要
- **何のプロジェクトか**: PowerPointプレゼンテーション自動生成ツール
- **主な目的**: 各種データ（JSON, CSV, テキスト, URL）からPPTXファイルを自動生成

## 2. 技術スタック
- **言語・ランタイム**: Python 3
- **主要ライブラリ**: python-pptx, requests

## 3. ディレクトリ構造
| パス | 役割 |
|------|------|
| main.py | CLIエントリポイント |
| config.py | デフォルト設定 |
| src/ | コアモジュール |
| templates/ | テンプレート定義（JSON） |
| output/ | 生成ファイル出力先 |
| tests/ | テスト |

## 4. 開発でよく使うコマンド
- **実行**: `python main.py --input data.json --template default`
- **テスト**: `python -m pytest tests/`

## 5. コーディング規約・ルール
- コメントは日本語
- Pythonファイル名はスネークケース（Pythonのimport制約による）
- 非Pythonファイル名はケバブケース
