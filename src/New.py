# 空のExcelファイルを作成する
# --使い方--
# ・「New.py hoge」で実行
# ・hoge.xlsxがカレントディレクトリに作成される。
# ・既に同名ファイルが存在していた場合上書きしてしまうので、注意 TODO ←上書きしないように修正する。

import sys
import openpyxl

# コマンドラインで引数が渡されているかを検査
if len(sys.argv) > 1:
    wb = openpyxl.Workbook()
    fileName = sys.argv[1] + '.xlsx'
    wb.save(fileName)
    print(fileName + 'を作成しました。')
else:
    # 引数の指定がない
    print('ファイル名を指定して下さい。')