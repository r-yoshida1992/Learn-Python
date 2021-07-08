# 引数で指定したディレクトリ配下のエクセルファイルからシート名を全て出力する
# --使い方--
# ・「SheetList.py」で実行
# ・指定のディレクトリに存在するxlsxおよびxlsmファイルから、シート名を出力する。

import glob
import openpyxl
import sys

# [真の場合] if [条件式] else [偽の場合]
files = [glob.glob(sys.argv[1] + "\*") if len(sys.argv) > 1 else glob.glob("*")]

for file in files:
    # xlsとxlsbはopenpyxlでは対応していない
    if file.endswith(('xlsx', 'xlsm')):
        print(file + 'のシート一覧:')
        wb = openpyxl.load_workbook(file)
        sheets = wb.sheetnames
        for sheet in sheets:
            print('\t' + sheet)
