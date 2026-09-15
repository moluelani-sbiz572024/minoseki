#
# ツールアプリ
#
import pandas as pd
#
from config import IO_DATA, KWD
from service.update_table import 仕入れ入出荷テーブルに書き込む, 出荷入出荷テーブルに書き込む
from service.edit_master import 受払在庫管理簿に転記する
from service.lock_forms import フォームをロックする
from excel import operator


def メイン処理():
    """
    商品別売上データを元に受払在庫管理簿に転記する
    """

    # 商品別売上データの読み込み
    商品マスタ = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['商品マスタ'])
    仕入れフォーム = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['仕入れフォーム'])
    入出荷テーブル = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['入出荷テーブル'])

    # 販売登録フォームは、ファイルがまだ書き換わってないうちに、先に読み込んでおく
    # (仕入れ処理の中でopenpyxlがファイルを保存し直すと、数式の計算済みの値が消えてしまうため)
    販売登録フォーム縦持ち = operator.販売登録フォームを整形する()

    # 商品マスタ・仕入れフォームのキー列に自動採番し、ロックをかける
    フォームをロックする()

    # ロックの処理でファイルが書き換わったので、商品マスタ・仕入れフォームを読み込み直す
    商品マスタ = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['商品マスタ'])
    仕入れフォーム = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['仕入れフォーム'])

    # 仕入れ処理
    結合仕入れデータ = pd.merge(仕入れフォーム, 商品マスタ, on=KWD['商品キーID'], how=KWD['左方向'])
    仕入れ入出荷テーブルに書き込む(結合仕入れデータ, 入出荷テーブル, operator.本当の最終行を取得する)

    # 出荷処理(販売登録フォーム整形)
    商品名なし商品マスタ = 商品マスタ.drop(columns=['商品名'])
    結合販売データ = pd.merge(販売登録フォーム縦持ち, 商品名なし商品マスタ, on=KWD['商品キーID'], how=KWD['左方向'])
    出荷入出荷テーブルに書き込む(結合販売データ, 入出荷テーブル, operator.本当の最終行を取得する)

    print('入出荷テーブルへの書き込みが完了しました')

    # 入出荷テーブルを読み込み直す（仕入れ・出荷の書き込みが反映された、最新の状態にするため）
    入出荷テーブル = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['入出荷テーブル'])

    # 受払在庫管理簿への転記
    受払在庫管理簿に転記する(商品マスタ, 入出荷テーブル, operator.本当の最終行を取得する)


# 実行エントリポイント
if __name__ == '__main__':
    メイン処理()