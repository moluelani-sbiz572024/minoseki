#
# テーブル更新
#
import pandas as pd
import openpyxl
#
from config import IO_DATA, KWD


def 仕入れ入出荷テーブルに書き込む(仕入れデータ, 入出荷テーブル, 本当の最終行を取得する):
    """
    """
    wb = openpyxl.load_workbook(IO_DATA['商品別売上'])
    ws = wb['入出荷テーブル']
    最終行 = 本当の最終行を取得する(ws)
    取引キー = 最終行 - 1

    for インデックス, 行 in 仕入れデータ.iterrows():
        既に存在する = (
            _同じ日付か(入出荷テーブル['入荷日'], 行['入荷日']) &
            _同じ値か(入出荷テーブル['商品キーID'], 行['商品キーID']) &
            _同じ値か(入出荷テーブル['仕入れ数量'], 行['入荷数量']) &
            _同じ値か(入出荷テーブル['仕入れ単価'], 行['仕入れ単価'])
        ).any()

        if 既に存在する:
            continue

        新しい行番号 = 最終行 + 1
        取引キー += 1
        ws.cell(row=新しい行番号, column=1).value = 取引キー
        ws.cell(row=新しい行番号, column=2).value = 行['入荷日']
        ws.cell(row=新しい行番号, column=2).number_format = 'm"月"d"日"'
        ws.cell(row=新しい行番号, column=4).value = 行['商品キーID']
        ws.cell(row=新しい行番号, column=5).value = 行['商品名']
        ws.cell(row=新しい行番号, column=7).value = 行['入荷数量']
        ws.cell(row=新しい行番号, column=9).value = 行['仕入れ単価']
        ws.cell(row=新しい行番号, column=11).value = "済"

        最終行 = 新しい行番号

    wb.save(IO_DATA['商品別売上'])


def 出荷入出荷テーブルに書き込む(出荷データ, 入出荷テーブル, 本当の最終行を取得する):
    """
    """
    wb = openpyxl.load_workbook(IO_DATA['商品別売上'])
    ws = wb['入出荷テーブル']
    最終行 = 本当の最終行を取得する(ws)
    取引キー = 最終行 - 1

    for インデックス, 行 in 出荷データ.iterrows():
        既に存在する = (
            _同じ日付か(入出荷テーブル['出荷日'], 行['出荷日']) &
            _同じ値か(入出荷テーブル['商品キーID'], 行['商品キーID']) &
            _同じ値か(入出荷テーブル['出荷数量'], 行['出荷数量']) &
            _同じ値か(入出荷テーブル['出荷先'], 行['出荷先'])
        ).any()

        if 既に存在する:
            continue

        新しい行番号 = 最終行 + 1
        取引キー += 1
        ws.cell(row=新しい行番号, column=1).value = 取引キー
        ws.cell(row=新しい行番号, column=3).value = 行['出荷日']
        ws.cell(row=新しい行番号, column=3).number_format = 'm"月"d"日"'
        ws.cell(row=新しい行番号, column=4).value = 行['商品キーID']
        ws.cell(row=新しい行番号, column=5).value = 行['商品名']
        ws.cell(row=新しい行番号, column=8).value = 行['出荷数量']
        ws.cell(row=新しい行番号, column=10).value = 行['出荷先']
        ws.cell(row=新しい行番号, column=11).value = "済"

        最終行 = 新しい行番号

    wb.save(IO_DATA['商品別売上'])

#-----
# プライベート関数
#-----

def _同じ値か(列, 値):
    """
    列と値を比較する関数：
    普通の == だと、空欄(NaN)同士を比べたとき、「違う」と判定されてしまうので、
    「両方とも空欄」の場合も「同じ」として扱うように判定する。
    """
    if pd.isna(値):
        return 列.isna()
    else:
        return 列 == 値


def _同じ日付か(列, 値):
    """
    日付を比較する関数：
    「ただの数字として保存された日付」と「本物の日付」など、
    見た目の形が違っても、同じ日なら「同じ」として扱う。
    """
    if pd.isna(値):
        return 列.isna()

    値_正規化 = pd.to_datetime(値)
    列_正規化 = pd.to_datetime(列, errors='coerce')

    return 列_正規化 == 値_正規化