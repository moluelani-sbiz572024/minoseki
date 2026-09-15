#
# Excel操作
#
import datetime
import pandas as pd
#
from config import IO_DATA, KWD


def 本当の最終行を取得する(ws):
    """
    見た目は空でも書式だけ残っている行を無視して、
    本当にデータが入っている最後の行番号を求める
    """
    for r in range(ws.max_row, 0, -1):
        行に何かある = any(ws.cell(row=r, column=c).value is not None for c in range(1, 12))
        if 行に何かある:
            return r
    
    return 1  # 何もデータが無ければ、見出し行(1行目)を返す


def 販売登録フォームを整形する():
    """
    販売登録フォームを縦持ちに変換する
    """
    # 1行目(出荷先の情報)だけを、見出しなしで読み込む
    出荷先行 = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['販売登録フォーム'], header=None, nrows=1)

    # 2行目を見出しとして、3行目以降のデータを読み込む
    販売登録フォーム = pd.read_excel(IO_DATA['商品別売上'], sheet_name=KWD['販売登録フォーム'], header=1)

    # 固定列(商品キーID・欠品・ブルワリー・商品名・品目・期首在庫)の名前を取得
    固定列 = 販売登録フォーム.columns[:6].tolist()

    # 日付の列だけを取り出す(「合計」など、日付じゃない列は除外する)
    日付列 = [列名 for 列名 in 販売登録フォーム.columns[6:] if isinstance(列名, (datetime.datetime, pd.Timestamp))]

    # 縦持ちに変換する
    縦持ち = 販売登録フォーム.melt(
        id_vars=固定列,
        value_vars=日付列,
        var_name='出荷日',
        value_name='出荷数量'
    )

    # 出荷数量が空欄の行は、削除する
    縦持ち = 縦持ち.dropna(subset=['出荷数量'])

    # 各日付列に対応する「出荷先」の対応表を作る
    出荷先対応表 = {}
    for i, 列名 in enumerate(日付列):
        列番号 = 販売登録フォーム.columns.get_loc(列名)
        出荷先対応表[列名] = 出荷先行.iloc[0, 列番号]

    # 縦持ちデータに、出荷先の列を追加する
    縦持ち['出荷先'] = 縦持ち['出荷日'].map(出荷先対応表)

    return 縦持ち


def 商品マスタから仕入れ先を取得する(商品マスタ, 商品キーID):
    """
    商品キーIDから商品マスタの仕入れ先を取得する
    """
    該当行 = 商品マスタ[商品マスタ['商品キーID'] == 商品キーID]

    if not 該当行.empty:
        return 該当行.iloc[0]['仕入れ先']

    return None