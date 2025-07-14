# ===============================================
# Excel画像・テキストデータを整形してJSON/Excel/画像出力するスクリプト
# ===============================================
# このスクリプトは以下のような流れで処理を行います：
# 1. Excelファイル（.xlsx）の中に埋め込まれた画像を取り出す
# 2. Excelの表形式データを整形し、Noごとの質問・回答にまとめる
# 3. 各質問・回答を1つのJSONファイルに変換、関連画像のファイルパスも記録
# 4. 全体をまとめた整形済みExcelファイルも出力
# 5. 実行環境情報（requirements.txt, Pythonバージョン）も一緒に保存

# 各シートで同じ形式（列名「分類」「項目」「No.」「質問」「回答」）を前提としています。
# シートごとに表形式が異なったり、質問が「No.」ではなく別の方法で識別されていると、別途調整が必要です。


import zipfile
import os
import xml.etree.ElementTree as ET
import shutil
from pathlib import Path
import json
import pandas as pd

# パス設定
input_xlsx_file = "FAQ抽出検討_トヨタ紡織_中村.xlsx"
json_output_dir = Path("json_output")
image_output_dir = json_output_dir / "image"
json_output_dir.mkdir(parents=True, exist_ok=True)
image_output_dir.mkdir(parents=True, exist_ok=True)


# === 列番号 → A1形式 ===
def colnum_to_excel_col(col_num):
    result = ""
    while col_num >= 0:
        result = chr(col_num % 26 + ord("A")) + result
        col_num = col_num // 26 - 1
    return result


# === 画像抽出 & 保存 ===
image_map = {}
with zipfile.ZipFile(input_xlsx_file, "r") as z:
    namelist = z.namelist()

    # --- 画像が貼られているシート情報の抽出 --
    # 画像の貼り付け位置情報は "drawing.xml" にあるが、それが「どのシート（Sheet1など）に貼られているのか」は別ファイルに書かれている。
    # そのため、まずは "xl/workbook.xml" から、内部的なシートのID（r:id）と、実際の「シート名」を対応づける辞書を作成する。
    # xl/workbook.xml から sheetId → シート名 を取得
    sheet_id_to_name = {}
    with z.open("xl/workbook.xml") as f:
        tree = ET.parse(f)
        ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        for sheet in tree.findall(".//main:sheets/main:sheet", ns):
            # それぞれの <sheet> タグから r:id と name（人が見るシート名）を取り出す
            r_id = sheet.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]
            sheet_id_to_name[r_id] = sheet.attrib["name"]

    # === workbook.xml.rels から rId → 対応するシートファイル名を取得する ===
    rId_to_sheet_file = {}

    # Excelファイルの内部構造では、全体の構成情報（workbook.xml）と、
    # 実際のシートファイル（sheet1.xml など）の関連情報が、リレーションファイル（.rels）に定義されている

    with z.open("xl/_rels/workbook.xml.rels") as f:
        tree = ET.parse(f)

        # 各 <Relationship> タグに、rId とリンク先ファイルの情報が入っている
        for rel in tree.getroot():
            # この Relationship が 'worksheet'（＝シート）であれば処理する
            if rel.attrib["Type"].endswith("/worksheet"):
                rId_to_sheet_file[rel.attrib["Id"]] = rel.attrib["Target"].split("/")[
                    -1
                ]

    # === 各シートファイルに紐づく drawing.xml を特定 ===
    sheet_to_drawing = {}

    for f in namelist:
        # シート用のリレーションファイルだけ対象とする
        if f.startswith("xl/worksheets/_rels/") and f.endswith(".xml.rels"):

            # リレーションファイル（sheetX.xml.rels）を開く
            with z.open(f) as rels_file:
                tree = ET.parse(rels_file)

                # 各 <Relationship> タグをチェック
                for rel in tree.getroot():
                    # Type 属性が drawing（図形情報）に関するものであれば処理する
                    if rel.attrib["Type"].endswith("/drawing"):
                        # このリレーションが属している sheet ファイル名を取得
                        sheet_file = f.split("/")[-1].replace(".rels", "")
                        # sheet ファイルに対応する drawing.xml を記録しておく
                        sheet_to_drawing[sheet_file] = rel.attrib["Target"].split("/")[
                            -1
                        ]

    # === drawing.xml.rels から rId と画像ファイル名の対応を取得 ===
    drawing_to_rId_image = {}

    for f in namelist:
        # drawing.xml.rels ファイルのみ対象にする（画像と drawing.xml の紐づけが書かれている）
        if f.startswith("xl/drawings/_rels/") and f.endswith(".xml.rels"):
            drawing_file = f.split("/")[-1].replace(".rels", "")
            drawing_to_rId_image[drawing_file] = {}

            # drawing.xml.rels を XML として読み込み
            with z.open(f) as rels_file:
                tree = ET.parse(rels_file)
                for rel in tree.getroot():

                    # rel タグの Type 属性が image の場合（画像ファイルとのリンク）
                    if rel.attrib["Type"].endswith("/image"):
                        drawing_to_rId_image[drawing_file][rel.attrib["Id"]] = (
                            os.path.basename(rel.attrib["Target"])
                        )

    # --- drawing.xml（図形定義ファイル）を解析して、画像の貼り付け位置と画像ファイルを保存 --- #
    for drawing_xml in [
        f
        for f in namelist
        if f.startswith("xl/drawings/drawing") and f.endswith(".xml")
    ]:
        drawing_name = drawing_xml.split("/")[-1]

        # drawingファイルと対応するシート名を取得
        sheet_name = next(
            (
                sheet_id_to_name[rId]
                for rId, sheet_file in rId_to_sheet_file.items()
                if sheet_to_drawing.get(sheet_file) == drawing_name
            ),
            None,
        )
        if not sheet_name:
            continue

        # drawing.xml を開いて XML ツリーとして読み込む
        with z.open(drawing_xml) as f:
            tree = ET.parse(f)
            ns = {
                "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            }

            # drawing内にある各画像・図形の定義（twoCellAnchor）を1つずつ処理
            for anchor in tree.findall("xdr:twoCellAnchor", ns):

                # 画像の左上にあるセルの位置（列番号・行番号）を取得
                col = int(anchor.find("xdr:from/xdr:col", ns).text)
                row = int(anchor.find("xdr:from/xdr:row", ns).text)
                cell_name = f"{colnum_to_excel_col(col)}{row + 1}"

                # 画像データの参照（<a:blip> 要素）を取得
                blip = anchor.find(".//a:blip", ns)
                if blip is not None:

                    # 画像の埋め込みID（rId）を取得
                    rId = blip.attrib.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                    )

                    # rId から対応する画像ファイル名を取得
                    image_file = drawing_to_rId_image.get(drawing_name, {}).get(rId)
                    if image_file:
                        media_path = f"xl/media/{image_file}"

                        # 実際にその画像ファイルが含まれていれば保存処理を行う
                        if media_path in namelist:
                            save_name = f"{sheet_name}_{cell_name}_{image_file}"
                            save_path = image_output_dir / save_name

                            # zipから画像ファイルを読み取り → 保存フォルダに書き出し
                            with (
                                z.open(media_path) as img_file,
                                open(save_path, "wb") as out_file,
                            ):
                                shutil.copyfileobj(img_file, out_file)

                            # 保存した画像パスを、画像の貼り付け行番号ごとに記録（後でNo.とマッチさせる用）
                            image_map.setdefault((sheet_name, row + 1), []).append(
                                str(save_path)
                            )
                            print(f"画像保存: {save_name}")

# 複数シートの場合を想定して結合
excel_path = Path(input_xlsx_file)
excel_file = pd.read_excel(excel_path, sheet_name=None)
df_list = []
for sheet_name, df in excel_file.items():
    df["シート名"] = sheet_name
    df_list.append(df)
df_all = pd.concat(df_list, ignore_index=True)


# === セル結合処理 & 元行保持 ===
def melt_merged_rows(df):
    """Excel でセル結合された行を想定し、質問と回答を結合して 1 行にまとめる。

    Parameters
    ----------
    df : pd.DataFrame

    Returns
    -------
    pd.DataFrame
        結合・集約後のデータフレーム
    """
    # 1) 結合セルの下方向への値コピー
    df_filled = df.copy()
    df_filled[["分類", "項目", "No."]] = df_filled[["分類", "項目", "No."]].ffill()

    # 2) 「新しい質問が始まる行」を識別するためのグループ ID
    # 3) 同じ grp 内で 質問・回答 を連結
    df_filled["grp"] = df["分類"].notna().cumsum()
    df_filled["original_row"] = df_filled.index + 2
    return (
        df_filled.groupby("grp", sort=False)
        .agg(
            {
                "シート名": "first",
                "分類": "first",
                "項目": "first",
                "No.": "first",
                "質問": lambda s: "\n".join(s.dropna().astype(str)),
                "回答": lambda s: "\n".join(s.dropna().astype(str)),
                "original_row": list,
            }
        )
        .reset_index(drop=True)
        .astype({"No.": int})
    )


organized_df = melt_merged_rows(df_all)

# === JSON出力 ===
for row in organized_df.to_dict(orient="records"):
    sheet = row["シート名"]
    no = row["No."]
    original_rows = row.pop("original_row")
    images = []
    for (img_sheet, img_row), paths in image_map.items():
        if img_sheet == sheet and img_row in original_rows:
            images.extend(paths)
    row["image_urls"] = images

    json_name = f"{excel_path.stem}_{sheet}_QA_{no}.json"
    with open(json_output_dir / json_name, "w", encoding="utf-8") as f:
        json.dump(row, f, ensure_ascii=False, indent=2)

# === 整形済みExcel保存 ===
# excel_output_path = json_output_dir / "QA整形結果.xlsx"
# organized_df.drop(columns=["original_row"]).to_excel(excel_output_path, index=False)

# print("JSON / Excel すべて出力完了")
