"""
データ入力処理（JSON, CSV, テキスト, URL）
すべての入力を統一dict形式に変換する
"""
import csv
import json
import os
from datetime import date


def load_from_file(filepath):
    """ファイルからデータを読み込み、統一形式のdictを返す"""
    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".json":
        return _load_json(filepath)
    elif ext == ".csv":
        return _load_csv(filepath)
    elif ext in (".txt", ".md"):
        return _load_text(filepath)
    else:
        raise ValueError(f"未対応のファイル形式です: {ext}")


def load_from_url(url):
    """URLからデータを取得し、統一形式のdictを返す"""
    import requests
    response = requests.get(url, timeout=30)
    response.raise_for_status()

    content_type = response.headers.get("content-type", "")

    if "json" in content_type or url.endswith(".json"):
        data = response.json()
        return _normalize(data)
    elif "csv" in content_type or url.endswith(".csv"):
        lines = response.text.strip().split("\n")
        return _parse_csv_lines(lines, os.path.basename(url))
    else:
        # テキストとして処理
        return _parse_text(response.text, os.path.basename(url))


def load_from_text(text, title="プレゼンテーション"):
    """テキスト文字列からデータを生成"""
    return _parse_text(text, title)


def _load_json(filepath):
    """JSONファイルを読み込む"""
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    return _normalize(data)


def _load_csv(filepath):
    """CSVファイルを読み込み、テーブルスライドとして構成"""
    filename = os.path.splitext(os.path.basename(filepath))[0]

    with open(filepath, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        return {"title": filename, "slides": []}

    headers = rows[0]
    data_rows = rows[1:]

    return _parse_csv_data(headers, data_rows, filename)


def _load_text(filepath):
    """テキスト/Markdownファイルを読み込む"""
    filename = os.path.splitext(os.path.basename(filepath))[0]

    with open(filepath, "r", encoding="utf-8") as f:
        text = f.read()

    return _parse_text(text, filename)


def _parse_csv_data(headers, data_rows, title):
    """CSVデータをスライド構成に変換（15行ごとにページ分割）"""
    max_rows_per_slide = 15
    slides = [
        {"type": "title", "title": title, "date": str(date.today())}
    ]

    for i in range(0, len(data_rows), max_rows_per_slide):
        chunk = data_rows[i:i + max_rows_per_slide]
        page_num = i // max_rows_per_slide + 1
        total_pages = (len(data_rows) + max_rows_per_slide - 1) // max_rows_per_slide

        slide_title = title
        if total_pages > 1:
            slide_title = f"{title} ({page_num}/{total_pages})"

        slides.append({
            "type": "table",
            "title": slide_title,
            "headers": headers,
            "rows": chunk,
        })

    return {"title": title, "slides": slides}


def _parse_csv_lines(lines, title):
    """CSV文字列行をパース"""
    reader = csv.reader(lines)
    rows = list(reader)
    if not rows:
        return {"title": title, "slides": []}
    return _parse_csv_data(rows[0], rows[1:], title)


def _parse_text(text, title):
    """テキストをスライド構成に変換
    # で始まる行 → スライドタイトル
    それ以降の行 → 箇条書き
    空行 → スライド区切り
    """
    slides = [
        {"type": "title", "title": title, "date": str(date.today())}
    ]

    blocks = text.strip().split("\n\n")

    for block in blocks:
        lines = block.strip().split("\n")
        if not lines:
            continue

        slide_title = ""
        body_lines = []

        for line in lines:
            line = line.strip()
            if not line:
                continue
            if line.startswith("# "):
                slide_title = line[2:].strip()
            elif line.startswith("## "):
                slide_title = line[3:].strip()
            else:
                # Markdownの箇条書きマーカーを除去
                if line.startswith("- ") or line.startswith("* "):
                    line = line[2:]
                body_lines.append(line)

        if slide_title or body_lines:
            slides.append({
                "type": "content",
                "title": slide_title,
                "body": body_lines,
            })

    return {"title": title, "slides": slides}


def _normalize(data):
    """JSONデータの正規化（必須フィールドの補完）"""
    if "title" not in data:
        data["title"] = "プレゼンテーション"
    if "slides" not in data:
        data["slides"] = []
    if "date" not in data:
        data["date"] = str(date.today())
    return data
