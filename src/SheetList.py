# カレントディレクトリのエクセルファイルからシート名を全て出力する
# --使い方--
# ・「SheetList.py」で実行
# ・カレントディレクトリに存在するxlsxおよびxlsmファイルから、シート名を出力する。

import glob
import openpyxl

# カレントディレクトリのファイル一覧を取得
files = glob.glob("*")

for file in files:
    # xlsとxlsbはopenpyxlでは対応していない
    if file.endswith(('xlsx', 'xlsm')):
        print(file + 'のシート一覧:')
        wb = openpyxl.load_workbook(file)
        sheets = wb.sheetnames
        for sheet in sheets:
            print('\t' + sheet)