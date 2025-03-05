
import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import glob

CWID = os.getenv("USERNAME")
Difftime = 15

# Read and filter data
d1 = pd.read_csv(f"Files/{CWID}/p_ausgabe.txt", encoding="cp932", sep="\t", dtype={8: str})
d2 = d1[~d1['ARBPL'].str.contains("QYA2100|QYA2101") & ~d1['PLTXT'].str.contains("安定性")]
d3 = d2[['PLNNR', 'PLTXT', 'ARBPL', 'WERKS_ARBPL', 'STEUSCHL', 'VORGTXT', 'LFD_MMNR', 'MKMNR', 'STMWERK', 'STPRBVERF']]
d3.loc[:, 'MKMNR'] = d3['MKMNR'].astype(str)  # .locを使用してMKMNR列を文字列型に変換
d4 = d3[~d3['MKMNR'].str.contains("Y0201OR") & (d3['MKMNR'] != "") & d3['VORGTXT'].isin([
    "Microbiological Laboratory", "Chemical Laboratory", "Informative Data", "Microbioloical Laboratory",
    "Chemical Laboratory-PAM", "Chemical Laboratory-QT3", "Chemical Laboratory-BIOP"])]
d4.loc[:, 'MKMNR'] = pd.Categorical(d4['MKMNR'], ordered=True)  # .locを使用してMKMNR列をカテゴリ型に変換
d4 = d4.copy()  # d4のコピーを作成して警告を回避
d4.loc[:, 'num'] = d4.groupby('PLNNR').cumcount() + 1  # .locを使用してnum列を追加
d5 = pd.read_csv(f"Files/{CWID}/mat_ausgabe.txt", encoding="cp932", sep="\t")[['PLNNR', 'MATNR', 'WERK']]
d5['MATNR'] = d5['MATNR'].str[-8:]
d6 = pd.merge(d5, d4, how='outer').dropna(subset=['MATNR', 'MKMNR']).drop_duplicates()
d6.to_csv(f"Files/{CWID}/table.txt", encoding="cp932", index=False, sep="\t")

# Read and process export data
# カラム名を変換する関数
def rename_columns(df):
    new_columns = {}
    for col in df.columns:
        if 'Text' in col:
            new_columns[col] = 'product'
        elif 'オブジェクト' in col:
            new_columns[col] = 'product'
        else:
            new_columns[col] = col
    df.rename(columns=new_columns, inplace=True)

    if 'product' not in df.columns:
        print("No columns renamed to 'product'. Press any key to exit.")
        input()  # ユーザーが何かキーを押すまで待機
        sys.exit()  # スクリプトを終了
    return(df)

dataj1 = pd.read_csv(f"Files/{CWID}/export1.txt", skiprows=1, encoding="cp932", sep="\t")
dataj1.columns = dataj1.columns.str.replace(' ', '')
dataj1 =rename_columns(dataj1)
dataj2 = pd.read_csv(f"Files/{CWID}/export2.txt", skiprows=1, encoding="cp932", sep="\t")
dataj2.columns = dataj2.columns.str.replace(' ', '')
dataj2 = rename_columns(dataj2)

dataj3 = dataj1.copy()
dataj4 = dataj2.copy()

dataj3 = dataj3[~dataj3.apply(lambda row: row.astype(str).str.contains('データ').any(), axis=1)]
dataj4 = dataj4[~dataj4.apply(lambda row: row.astype(str).str.contains('データ').any(), axis=1)]

dataj5 = pd.concat([dataj3, dataj4], ignore_index=True)

dataj6 = dataj5[['品目', 'product', 'QMatAu']].drop_duplicates()
if dataj6.empty:
    print("更新する製剤はありませんでした")
    input()  # ユーザーが何かキーを押すまで待機
    exit()  # スクリプトを終了

dataj6.columns = ['MATNR', 'product', 'QMatAu']
dataj6['MATNR'] = dataj6['MATNR'].apply(lambda x: f"{int(x):08d}")

data1 = pd.merge(dataj6, d6, how='left').dropna()
data1['product'] = data1['product'].str.replace(",", ".")

if data1.shape[0] == 0:
    raise Exception("更新する製剤がありませんでした")
else:
    data1['date1'] = np.where(data1['product'].isin(["BASC-100", "LANTHANUM CARBONATE"]),
                              (datetime.now() - timedelta(days=1200)).strftime("%Y/%m/%d"), "")
    data1['date2'] = np.where(data1['product'].isin(["BASC-100", "LANTHANUM CARBONATE"]),
                              datetime.now().strftime("%Y/%m/%d"), "")

    data2 = data1[['MATNR', 'QMatAu', 'MKMNR', 'date1', 'date2', 'WERK', 'num', 'product']]
    data2.to_csv(f"Files/{CWID}/table2.csv", index=False, quotechar='"', encoding="cp932")

    data3 = data1[['MATNR', 'product']].drop_duplicates().groupby('MATNR').head(1).reset_index(drop=True)
    data3.to_csv(f"Files/{CWID}/table3.csv", index=False, quotechar='"', encoding="cp932")

    print("以下の製剤を更新します")
    print(data3)

    pickup = []
    for i, matnr in enumerate(data3['MATNR']):
        files = glob.glob(f"rawdata/results/*{matnr}*")
        if files:
            if (datetime.now() - datetime.fromtimestamp(os.path.getmtime(files[0]))).total_seconds() / 3600 > Difftime:
                pickup.append(i)
        else:
            pickup.append(i)

    data4 = data3.iloc[pickup]
    data4.to_csv(f"Files/{CWID}/table5.csv", index=False, quotechar='"', encoding="cp932")

    data5 = data2[data2['MATNR'].isin(data4['MATNR'])]
    data5.to_csv(f"Files/{CWID}/table4.csv", index=False, quotechar='"', encoding="cp932", na_rep="")

    print("以下をSAPからダウンロードします")
    print(data4)

    pickup_q = []
    for i, matnr in enumerate(data3['MATNR']):
        files = glob.glob(f"rawdata/qa33/*{matnr}*")
        if files:
            if (datetime.now() - datetime.fromtimestamp(os.path.getmtime(files[0]))).total_seconds() / 3600 > Difftime:
                pickup_q.append(i)
        else:
            pickup_q.append(i)

    data6 = data3.iloc[pickup_q].drop_duplicates()
    data6.to_csv(f"Files/{CWID}/table6.csv", index=False, quotechar='"', encoding="cp932")