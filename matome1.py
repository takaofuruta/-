import streamlit as st
import pandas as pd
import openpyxl
import os
import re
import unicodedata
from PyPDF2 import PdfReader
from pdfminer.high_level import extract_text
from openpyxl.styles import Alignment

# Streamlit UI
st.title("建築データ処理アプリ")
st.write("『図面データ.pdf』『面積表 図面.pdf』をアップロードして処理を行い、Excelを更新します。")

# PDFファイルのアップロード（特定の2つのファイルのみ）
uploaded_files = st.file_uploader("PDFをアップロード（2つのファイルのみ選択可）", type=["pdf"], accept_multiple_files=True)

# Excelファイルのテンプレート
excel_template = "建築工事届.xlsx"

# ファイルの処理開始
if uploaded_files:
    updated_excel = "処理済_建築工事届.xlsx"

    # Excelの元データを保持
    if os.path.exists(excel_template):
        wb = openpyxl.load_workbook(excel_template)
    else:
        st.error(f"テンプレートのExcel ({excel_template}) が見つかりません！")
        st.stop()

    # PDFの処理ロジック
    for uploaded_file in uploaded_files:
        pdf_name = uploaded_file.name
        st.write(f"📂 アップロードされたファイル: {pdf_name}")

        try:
            reader = PdfReader(uploaded_file)
            extracted_text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])

            if "図面データ.pdf" in pdf_name:
                st.write(f"⚙ {pdf_name} → **図面データの処理開始**")

                # 建築主名の抽出
                name_start = extracted_text.find("建築主:") + len("建築主:")
                name_end = extracted_text.find("〒") - 1
                name = extracted_text[name_start:name_end]

                # 郵便番号の抽出
                yubinbango_start = extracted_text.find("〒") + 1
                yubinbango_end = extracted_text.find("-", yubinbango_start)
                yubinbango = extracted_text[yubinbango_start:yubinbango_end]

                # 建築場所の抽出
                place_start = extracted_text.find("建築場所（地名地番）")
                place_line_start = extracted_text.find("\n", place_start) + 1
                place_line_end = extracted_text.find("\n", place_line_start)
                address = extracted_text[place_line_start:place_line_end].strip()

                # Excelの更新（建築工事届（別記第40号様式））
                sheet_name1 = "建築工事届（別記第40号様式）"
                ws1 = wb[sheet_name1]
                ws1["I16"] = name
                ws1["I17"] = yubinbango
                ws1["O72"] = address

            elif "面積表 図面.pdf" in pdf_name:
                st.write(f"⚙ {pdf_name} → **面積表の処理開始**")

                # 面積抽出
                numeric_values = [line.strip().replace("㎡", "").replace("％", "")
                                  for line in extracted_text.splitlines() if re.fullmatch(r"[\d.]+", line.strip())]

                site_area = numeric_values[0]  # 敷地面積
                total_floor_area = numeric_values[3]  # 延床面積

                # Excelの更新（面積情報の転記）
                sheet_name2 = "第三種換気"
                ws2 = wb[sheet_name2]
                ws2["K92"] = total_floor_area
                ws2["S109"] = site_area

            else:
                st.warning(f"⚠ {pdf_name} は対象外のPDFです。スキップします。")

        except Exception as e:
            st.error(f"❌ PDF処理中にエラーが発生: {e}")

    # Excelファイルの保存
    wb.save(updated_excel)
    st.success(f"✅ Excelファイル ({updated_excel}) を更新しました！")

    # ダウンロードボタンの追加
    with open(updated_excel, "rb") as f:
        st.download_button(label="📥 処理済みExcelをダウンロード", data=f, file_name=updated_excel, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
