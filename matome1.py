import streamlit as st
from PyPDF2 import PdfReader
import openpyxl
import re
import unicodedata

# Streamlit UI部分
st.title("PDF to Excel Processor")
uploaded_files = st.file_uploader("複数のPDFファイルをアップロードしてください", accept_multiple_files=True, type=["pdf"])
excel_template = "建築工事届.xlsx"

if uploaded_files:
    workbook = openpyxl.load_workbook(excel_template)
    sheet_name = "第三種換気"

    # シートのチェック
    if sheet_name not in workbook.sheetnames:
        st.error(f"シート '{sheet_name}' が見つかりません！")
    else:
        sheet = workbook[sheet_name]

    # 各PDFを処理
    for uploaded_file in uploaded_files:
        pdf_reader = PdfReader(uploaded_file)
        pdf_text = pdf_reader.pages[0].extract_text()

        # 前処理
        processed_text = re.sub(r"\s+", "", pdf_text)

        # 処理分岐
        if "図面データ.pdf" in uploaded_file.name:
            # 図面データの処理
                from pdfminer.high_level import extract_text
            text=extract_text("図面データ.pdf")
            name_start=text.find("建築主:")+len("建築主:")
            name_end=text.find("〒")-1
            name=text[name_start:name_end]
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
            import re

            lines = text.splitlines()

            # 面積項目リスト（順番通り）
            area_labels = [
                "敷地面積",       # → site_area
                "1階床面積",
                "2階床面積",
                "延床面積",       # → total_floor_area
                "建築面積",
                "建蔽率",
                "容積率"
            ]

            numeric_values = []
            for line in lines:
                cleaned = line.strip().replace("㎡", "").replace("％", "")
                if re.fullmatch(r"[\d.]+", cleaned):
                    numeric_values.append(cleaned)

            site_area = numeric_values[0]             
            total_floor_area = numeric_values[3]      

            import openpyxl
            from openpyxl.styles import Alignment

            excel_path = "建築工事届.xlsx"
            sheet_name = "建築工事届（別記第40号様式）"
            sheet_name2 = "第三種換気"

            wb = openpyxl.load_workbook(excel_path)

            # シート1: "建築工事届（別記第40号様式）"
            ws = wb[sheet_name]

            ken_cells = ["A12", "J72"]
            for ken_cell in ken_cells:
                ws[ken_cell] = ken_name
            ws["I16"] = name
            ws["I17"] = yubinbango
            ws["O17"] = yubinbango_last4
            ws["I18"] = top_address
            ws["O72"] = address
            ws["J85"] = kouji_line
            ws["K92"] = total_floor_area
            ws["S109"] = site_area

            # シート2: "第三種換気"
            ws2 = wb[sheet_name2]  # シートを取得

            # B1セルに値を入力
            ws2["B1"] = kouji_line
            # B1セルを右揃えに設定 (こちらでシート"ws2"を正しく指定)
            ws2["B1"].alignment = Alignment(horizontal="right", vertical="center")

            # 保存処理
            wb.save(excel_path)

        elif "面積表　図面.pdf" in uploaded_file.name:
            # 面積表の処理
            table_pattern = r"部屋名.*?合計"
            tables = re.findall(table_pattern, processed_text)

            # 各表の処理
            room_pattern = r"(玄関|階段|トイレ|ＬＤＫ|洗面脱衣室|洋室|廊下|サービスルーム)[^\d]*([\d.]+)"
            floor_data = {"1階": {}, "2階": {}}

            for table in tables:
                if "玄関" in table:
                    floor_name = "1階"
                else:
                    floor_name = "2階"

                matches = re.findall(room_pattern, table)
                for room, area in matches:
                    floor_data[floor_name][room] = float(area)

            # エクセル転記
            for i, (room, area) in enumerate(floor_data["1階"].items(), start=8):
                sheet.cell(row=i, column=2, value=room)
                sheet.cell(row=i, column=4, value=area)

            for i, (room, area) in enumerate(floor_data["2階"].items(), start=19):
                sheet.cell(row=i, column=2, value=room)
                sheet.cell(row=i, column=4, value=area)

            st.write("面積表の処理完了！")

    # 処理済みエクセルファイルを保存＆ダウンロード
    output_file = "処理済み_建築工事届.xlsx"
    workbook.save(output_file)
    with open(output_file, "rb") as file:
        st.download_button(label="処理済みエクセルをダウンロード", data=file, file_name=output_file)

st.info("全ての処理が完了しました！")
