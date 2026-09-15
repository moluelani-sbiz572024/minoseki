#
# 管理簿の編集
#
import openpyxl
#
from config import IO_DATA, KWD
from excel import operator


def 受払在庫管理簿に転記する(商品マスタ, 入出荷テーブル, 本当の最終行を取得する):
    """
    入出荷テーブルの、まだ転記していないデータを、
    受払在庫管理簿の各商品シートに書き込む
    """
    
    # 受払在庫管理簿を、書き込み用に開く
    受払wb = openpyxl.load_workbook(IO_DATA['受払在庫管理簿'])

    # 入出荷テーブルを、書き込み用に開く(転記済みフラグを立てるため)
    入出荷wb = openpyxl.load_workbook(IO_DATA['商品別売上'])
    入出荷ws = 入出荷wb['入出荷テーブル']

    # 「受払転記済み」列(L列)の見出しが無ければ、追加する
    if 入出荷ws.cell(row=1, column=12).value != '受払転記済み':
        入出荷ws.cell(row=1, column=12).value = '受払転記済み'

    for インデックス, 行 in 入出荷テーブル.iterrows():
        エクセル行番号 = インデックス + 2  # pandasは0始まり、Excelは2行目からデータなので+2

        既に転記済み = 入出荷ws.cell(row=エクセル行番号, column=12).value == '済'
        if 既に転記済み:
            continue

        商品名 = 行['商品名']

        if 商品名 not in 受払wb.sheetnames:
            受払ws = _新しい商品シートを作る(受払wb, 商品名)
            print(f"「{商品名}」の新しいシートを作りました。")
        else:
            受払ws = 受払wb[商品名]

        最終行 = 本当の最終行を取得する(受払ws)
        新しい行番号 = 最終行 + 1

        受払ws.cell(row=新しい行番号, column=1).value = 行['入荷日']
        受払ws.cell(row=新しい行番号, column=1).number_format = 'm"月"d"日"'

        受払ws.cell(row=新しい行番号, column=2).value = 行['出荷日']
        受払ws.cell(row=新しい行番号, column=2).number_format = 'm"月"d"日"'

        受払ws.cell(row=新しい行番号, column=3).value = 商品名
        受払ws.cell(row=新しい行番号, column=5).value = 行['仕入れ数量']
        受払ws.cell(row=新しい行番号, column=6).value = 行['出荷数量']
        受払ws.cell(row=新しい行番号, column=7).value = \
            operator.商品マスタから仕入れ先を取得する(商品マスタ, 行['商品キーID'])
        受払ws.cell(row=新しい行番号, column=8).value = 行['仕入れ単価']
        受払ws.cell(row=新しい行番号, column=9).value = 行['出荷先']

        # 転記済みフラグを立てる
        入出荷ws.cell(row=エクセル行番号, column=12).value = '済'

    受払wb.save(IO_DATA['受払在庫管理簿'])
    入出荷wb.save(IO_DATA['商品別売上'])
    
    print("受払在庫管理簿への転記が完了しました")

#-----
# プライベート関数
#-----

def _新しい商品シートを作る(受払wb, 商品名):
    """
    ひな形をコピーして、商品名シートを上之保ゆずエール(瓶)の右隣・ひな形の左隣に作る
    """
    ひな形 = 受払wb['ひな形']
    新シート = 受払wb.copy_worksheet(ひな形)
    新シート.title = 商品名  # ← タブの名前(シート名)に、商品名を設定

    # 1行目のタイトルにも、商品名を書く
    新シート.cell(row=1, column=1).value = 商品名

    # 3行目以降の、ひな形に入ってるサンプルデータを消す
    for r in range(3, 新シート.max_row + 1):
        for c in range(1, 13):
            新シート.cell(row=r, column=c).value = None

    # 新シートは、今「一番最後」にいる。
    # 「ひな形」の位置まで移動させることで、
    # ひな形の"直前"に割り込む形になる
    新シートの今の位置 = 受払wb.sheetnames.index(商品名)
    ひな形の位置 = 受払wb.sheetnames.index('ひな形')
    受払wb.move_sheet(商品名, offset=ひな形の位置 - 新シートの今の位置)

    return 新シート