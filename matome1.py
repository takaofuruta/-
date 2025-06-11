import streamlit as st
from PyPDF2 import PdfReader
import openpyxl
import re
import unicodedata

# Streamlitタイトルと説明
st.title("PDF to Excel Processing App")
st.write("Upload multiple PDF files and process them to update an Excel template!")

# ファイルアップロード
uploaded_files = st.file_uploader("Upload your PDF files (e.g., 図面データ.pdf, 面積表 図面.pdf)", accept_multiple_files=True, type=["pdf"])

# エクセルテンプレートのファイル名
excel_template = "建築工事届.xlsx"

if uploaded_files:
    # エクセルテンプレートを読み込む
    try:
        workbook = openpyxl.load_workbook(excel_template)
        sheet_name2 = "第三種換気"

        if sheet_name2 not in workbook.sheetnames:
            st.error(f"Sheet '{sheet_name2}' not found in the template!")
        else:
            sheet = workbook[sheet_name2]

        # 各PDFファイルを処理
        for uploaded_file in uploaded_files:
            pdf_reader = PdfReader(uploaded_file)
            full_text = pdf_reader.pages[0].extract_text()

            # 前処理: 空白や改行を削除
            processed_text = re.sub(r"\s+", "", full_text)

            if "図面データ.pdf" in uploaded_file.name:
                # 図面データの処理: 建築主情報を抽出
                name_start = processed_text.find("建築主:") + len("建築主:")
                name_end = processed_text.find("〒") - 1
                name = processed_text[name_start:name_end].strip()

                zip_start = processed_text.find("〒") + 1
                zip_end = processed_text.find("-", zip_start) + 5
                zip_code = processed_text[zip_start:zip_end].strip()

                address_start = processed_text.find("住所") + len("住所")
                address_end = processed_text.find("電話番号", address_start)
                address = processed_text[address_start:address_end].strip()

                phone_start = processed_text.find("電話番号") + len("電話番号")
                phone_end = processed_text.find("-", phone_start) + 9
                phone = processed_text[phone_start:phone_end].strip()

                # エクセルに転記
                sheet["I16"] = name  # 建築主名
                sheet["I17"] = zip_code  # 郵便番号
                sheet["I18"] = address  # 住所
                sheet["I19"] = phone  # 電話番号

                st.write(f"Processed 図面データ: Added 建築主情報 (Name: {name}, Zip: {zip_code}, Address: {address}, Phone: {phone}) to Excel.")

            elif "面積表 図面.pdf" in uploaded_file.name:
                # 面積表データの処理
                table_pattern = r"部屋名.*?合計"
                tables = re.findall(table_pattern, processed_text)

                # 各表のデータを抽出
                room_pattern = r"(玄関|階段|トイレ|ＬＤＫ|洗面脱衣室|洋室|廊下|サービスルーム)[^\d]*([\d.]+)"
                ceiling_heights = {
                    "玄関": 2.58,
                    "ホール": 2.4,
                    "階段": 2.875,
                    "洗面脱衣室": 2.4,
                    "洋室": 2.4,
                    "トイレ": 2.4,
                    "廊下": 2.4,
                    "サービスルーム": 2.4,
                    "ＬＤＫ": 2.4
                }

                floor_data = {}
                for table in tables:
                    # 「玄関」の有無で1階・2階を判別
                    if "玄関" in table:
                        floor_name = "1階"
                    else:
                        floor_name = "2階"

                    matches = re.findall(room_pattern, table)
                    room_counts = {}
                    floor_areas = {}
                    for room, area in matches:
                        room_counts[room] = room_counts.get(room, 0) + 1
                        unique_room = f"{room}{room_counts[room]}" if room_counts[room] > 1 else room
                        floor_areas[unique_room] = float(area)

                    if floor_name not in floor_data:
                        floor_data[floor_name] = floor_areas
                    else:
                        floor_data[floor_name].update(floor_areas)

                # エクセルに転記
                ground_floor_areas = floor_data.get("1階", {})
                for i, (room, area) in enumerate(ground_floor_areas.items(), start=8):
                    sheet.cell(row=i, column=2, value=room)
                    sheet.cell(row=i, column=4, value=area)
                    room_base_name = re.sub(r"\d+$", "", room)
                    sheet.cell(row=i, column=5, value=ceiling_heights.get(room_base_name, "-"))

                first_floor_areas = floor_data.get("2階", {})
                for i, (room, area) in enumerate(first_floor_areas.items(), start=19):
                    sheet.cell(row=i, column=2, value=room)
                    sheet.cell(row=i, column=4, value=area)
                    room_base_name = re.sub(r"\d+$", "", room)
                    sheet.cell(row=i, column=5, value=ceiling_heights.get(room_base_name, "-"))

                st.write("Processed 面積表: Room data added to Excel.")

        # 処理済みエクセルファイルを保存
        output_file = "Processed_建築工事届.xlsx"
        workbook.save(output_file)

        # ダウンロードボタンを表示
        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Processed Excel",
                data=file,
                file_name=output_file
            )

        st.success("All processing completed successfully!")

    except Exception as e:
        st.error(f"An error occurred: {e}")
