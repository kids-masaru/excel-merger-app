import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os
import tempfile
from openpyxl.utils.exceptions import InvalidFileException
from copy import copy
import base64
import io
import traceback

# アプリのタイトルとスタイルの設定
st.set_page_config(page_title="Excelシートマージツール（マクロ対応版）", layout="wide")

# アプリのタイトルと説明
st.title("Excelシートマージツール（マクロ対応版）")
st.markdown("""
### このツールの機能
- template.xlsm と 利用明細書 のExcelファイルを受け取ります
- テンプレートの最初の2つのシートはそのまま保持します
- 利用明細書の全シートをSheet1, Sheet2...という名前でテンプレートに追加します
- マクロは保持されます（出力は.xlsm形式）
""")

# ファイルのダウンロード用関数
def get_binary_file_downloader_html(bin_data, file_label='File', file_name='file.xlsm'):
    bin_str = base64.b64encode(bin_data).decode()
    href = f'<a href="data:application/vnd.ms-excel.sheet.macroEnabled.12;base64,{bin_str}" download="{file_name}">📥 {file_label}</a>'
    return href

def load_template(file_content):
    """template.xlsm を読み込む関数"""
    try:
        file_stream = io.BytesIO(file_content)
        template_wb = load_workbook(file_stream, keep_vba=True)
        if len(template_wb.sheetnames) < 2:
            st.error("エラー: テンプレートには最低2つのシートが必要です")
            return None
        return template_wb
    except Exception as e:
        st.error(f"テンプレートファイルの読み込みエラー: {str(e)}")
        st.error(traceback.format_exc())
        return None

def load_additional_data(file_content):
    """利用明細書のExcelファイルを読み込む関数"""
    try:
        file_stream = io.BytesIO(file_content)
        data_wb = load_workbook(file_stream, keep_vba=True)
        return data_wb
    except Exception as e:
        st.error(f"利用明細書の読み込みエラー: {str(e)}")
        st.error(traceback.format_exc())
        return None

def merge_workbooks(template_wb, data_wb):
    """
    利用明細書のシートをテンプレートワークブックに追加する関数
    テンプレートのシートはそのままで、追加シートは「Sheet1」から連番で命名
    """
    try:
        template_sheet_names = template_wb.sheetnames.copy()
        st.write(f"既存テンプレートシート（維持）: {', '.join(template_sheet_names)}")
        
        # マッピング情報は内部で保持（表示はしない）
        sheet_name_mapping = {}
        sheet_number = 1
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_sheets = len(data_wb.sheetnames)
        for idx, sheet_name in enumerate(data_wb.sheetnames):
            progress = int((idx / total_sheets) * 100)
            progress_bar.progress(progress)
            status_text.text(f"処理中... {sheet_name} ({idx+1}/{total_sheets})")
            
            source_sheet = data_wb[sheet_name]
            new_sheet_name = f"Sheet{sheet_number}"
            sheet_number += 1
            
            while new_sheet_name in template_wb.sheetnames:
                new_sheet_name = f"Sheet{sheet_number}"
                sheet_number += 1
            
            sheet_name_mapping[sheet_name] = new_sheet_name
            
            target_sheet = template_wb.create_sheet(title=new_sheet_name)
            
            for i, col in enumerate(source_sheet.columns, 1):
                col_letter = openpyxl.utils.get_column_letter(i)
                if col_letter in source_sheet.column_dimensions:
                    target_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
            
            for i, row in enumerate(source_sheet.rows, 1):
                if i in source_sheet.row_dimensions:
                    target_sheet.row_dimensions[i].height = source_sheet.row_dimensions[i].height
            
            for row in source_sheet.rows:
                for cell in row:
                    if isinstance(cell, openpyxl.cell.cell.Cell):
                        target_cell = target_sheet.cell(row=cell.row, column=cell.column)
                        target_cell.value = cell.value
                        
                        if cell.has_style:
                            target_cell.font = copy(cell.font)
                            target_cell.border = copy(cell.border)
                            target_cell.fill = copy(cell.fill)
                            target_cell.number_format = cell.number_format
                            target_cell.protection = copy(cell.protection)
                            target_cell.alignment = copy(cell.alignment)
            
            for merged_range in source_sheet.merged_cells.ranges:
                target_sheet.merge_cells(str(merged_range))
        
        progress_bar.progress(100)
        status_text.text("処理完了！")
        
        return template_wb
    
    except Exception as e:
        st.error(f"ワークブックのマージエラー: {str(e)}")
        st.error(traceback.format_exc())
        return None

def main():
    st.sidebar.title("使い方")
    st.sidebar.markdown("""
    1. template.xlsm をアップロード (.xlsm推奨)
    2. 利用明細書をアップロード
    3. 「処理開始」ボタンをクリック
    4. 処理完了後、結果ファイルをダウンロード
    
    **注意点:**
    - テンプレートには最低2つのシートが必要です
    - マクロは保持されます
    - 出力ファイルは.xlsm形式です
    """)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("デバッグ情報")
    st.sidebar.info("アプリが正常に読み込まれました")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. template.xlsm")
        template_file = st.file_uploader("template.xlsm をアップロード", type=["xlsm", "xlsx"])
    
    with col2:
        st.subheader("2. 利用明細書")
        data_file = st.file_uploader("利用明細書をアップロード", type=["xlsx", "xlsm", "xls"])
    
    st.markdown("---")
    
    if template_file is not None and data_file is not None:
        st.info("ファイルがアップロードされました。「処理開始」ボタンをクリックしてください。")
        
        if st.button("処理開始", key="process_button", help="クリックしてファイルをマージします"):
            st.info("処理を開始します。しばらくお待ちください...")
            
            try:
                template_content = template_file.read()
                data_content = data_file.read()
                
                template_wb = load_template(template_content)
                if template_wb is not None:
                    data_wb = load_additional_data(data_content)
                    if data_wb is not None:
                        merged_wb = merge_workbooks(template_wb, data_wb)
                        if merged_wb is not None:
                            try:
                                output_buffer = io.BytesIO()
                                merged_wb.save(output_buffer)
                                output_buffer.seek(0)
                                
                                st.success("マージが完了しました！")
                                
                                output_filename = "merged_excel.xlsm"
                                st.markdown(
                                    get_binary_file_downloader_html(
                                        output_buffer.getvalue(), 
                                        "結果ファイルをダウンロード", 
                                        output_filename
                                    ), 
                                    unsafe_allow_html=True
                                )
                                
                                st.subheader("処理の概要")
                                st.markdown(f"""
                                - テンプレートの元のシート数: {len(template_wb.sheetnames) - len(data_wb.sheetnames)}
                                - 追加されたシート数: {len(data_wb.sheetnames)}
                                - 最終シート数: {len(merged_wb.sheetnames)}
                                """)
                            except Exception as e:
                                st.error(f"ファイル保存エラー: {str(e)}")
                                st.error(traceback.format_exc())
            except Exception as e:
                st.error(f"予期せぬエラーが発生しました: {str(e)}")
                st.error(traceback.format_exc())
    else:
        st.info("template.xlsm と 利用明細書 をアップロードしてください。")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"アプリケーションエラー: {str(e)}")
        st.error(traceback.format_exc())
