import streamlit as st
import re
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text
import openpyxl
from openpyxl.styles import Alignment
import unicodedata

# Streamlitタイトルと説明
st.title("PDF to Excel Processor")
st.write("複数のPDFファイルをアップロードし、それぞれの処理を実行して結果をExcelに転記します。")

# ファイルアップロード
uploaded_files = st.file_uploader("PDFファイルをアップロードしてください（例: 図面データ.pdf, 面積表 図面.pdf）", accept_multiple_files=True, type=["pdf"])

# エクセルテンプレートのパス
excel_path = "建築工事届.xlsx"

if uploaded_files:
    # エクセルテンプレートを読み込み
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet_name_main = "建築工事届（別記第40号様式）"
        sheet_name_aux = "第三種換気"

        if sheet_name_main not in workbook.sheetnames or sheet_name_aux not in workbook.sheetnames:
            st.error(f"シート名 '{sheet_name_main}' または '{sheet_name_aux}' が見つかりません！")
        else:
            ws_main = workbook[sheet_name_main]
            ws_aux = workbook[sheet_name_aux]

            for uploaded_file in uploaded_files:
                if "図面データ.pdf" in uploaded_file.name:
                    # 図面データの処理（処理内容はそのまま）
                    text = extract_text(uploaded_file)

                    name_start = text.find("建築主:") + len("建築主:")
                    name_end = text.find("〒") - 1
                    name = text[name_start:name_end]

                    yubinbango_start = text.find("〒") + 1
                    yubinbango_end = text.find("-", yubinbango_start)
                    yubinbango = text[yubinbango_start:yubinbango_end]

                    hyphen_index = text.find("-")
                    yubinbango_last4 = text[hyphen_index + 1: hyphen_index + 5]

                    zip_pos = text.find("〒")
                    next_line_start = text.find("\n", zip_pos) + 1
                    address_end = text.find("建築場所（地名地番）")
                    top_address = text[next_line_start:address_end].strip()

                    place_start = text.find("建築場所（地名地番）")
                    place_line_start = text.find("\n", place_start) + 1
                    place_line_end = text.find("\n", place_line_start)
                    place_line = text[place_line_start:place_line_end].strip()

                    ken_index = place_line.find("県")
                    ken_name = place_line[:ken_index + 1] if ken_index != -1 else ""

                    start_pos = text.find("建築場所（地名地番）")
                    end_pos = text.find("建築場所（住居表示）")
                    start_pos += len("建築場所（地名地番）")

                    address = text[start_pos:end_pos].strip()

                    lines = text.splitlines()
                    kouji_line = ""
                    for line in lines:
                        line_stripped = line.strip()
                        if line_stripped.endswith("工 事"):
                            kouji_line = line_stripped
                            break

                    # データをエクセルに転記
                    ken_cells = ["A12", "J72"]
                    for ken_cell in ken_cells:
                        ws_main[ken_cell] = ken_name
                    ws_main["I16"] = name
                    ws_main["I17"] = yubinbango
                    ws_main["O17"] = yubinbango_last4
                    ws_main["I18"] = top_address
                    ws_main["O72"] = address
                    ws_main["J85"] = kouji_line

                    st.write("図面データ.pdfの処理が完了しました。")

                elif "面積表 図面.pdf" in uploaded_file.name:
                    # 面積表 図面データの処理（処理内容はそのまま）
                    reader = PdfReader(uploaded_file)
                    full_text = reader.pages[0].extract_text()
                    processed_text = re.sub(r"\s+", "", full_text)

                    table_pattern = r"部屋名.*?合計"
                    tables = re.findall(table_pattern, processed_text)

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
                        "ＬＤＫ": 2.4,
                    }

                    floor_data = {}
                    for table in tables:
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

                        floor_data[floor_name] = floor_areas

                    ground_floor_areas = floor_data.get("1階", {})
                    for i, (room, area) in enumerate(ground_floor_areas.items(), start=8):
                        ws_aux.cell(row=i, column=2, value=room)
                        ws_aux.cell(row=i, column=4, value=area)
                        room_base_name = re.sub(r"\d+$", "", room)
                        ws_aux.cell(row=i, column=5, value=ceiling_heights.get(room_base_name, "-"))

                    first_floor_areas = floor_data.get("2階", {})
                    for i, (room, area) in enumerate(first_floor_areas.items(), start=19):
                        ws_aux.cell(row=i, column=2, value=room)
                        ws_aux.cell(row=i, column=4, value=area)
                        room_base_name = re.sub(r"\d+$", "", room)
                        ws_aux.cell(row=i, column=5, value=ceiling_heights.get(room_base_name, "-"))

                    st.write("面積表 図面.pdfの処理が完了しました。")

            # 処理済みエクセルファイルの保存
            output_file = "Processed_建築工事届.xlsx"
            workbook.save(output_file)

            # ダウンロードボタン
            with open(output_file, "rb") as file:
                st.download_button(label="処理済みExcelファイルをダウンロード", data=file, file_name=output_file)

            st.success("すべての処理が完了しました！")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
