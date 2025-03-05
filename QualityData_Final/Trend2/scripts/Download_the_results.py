import pandas as pd
import os
import numpy as np
import subprocess
import time
from datetime import datetime

# CWIDの値を設定
CWID = os.getenv("USERNAME")

# CSVファイルを読み込む
d1 = pd.read_csv(f"Files/{CWID}/table4.csv", encoding="cp932") #種類

# MATNR列の値を8桁のゼロ埋め文字列にフォーマット
d1['MATNR'] = d1['MATNR'].apply(lambda x: f"{int(x):08d}")

d1['date1'] = d1['date1'].replace(np.nan,"")
d1['date2'] = d1['date2'].replace(np.nan,"")

unique_matnr = d1['MATNR'].unique()  # Call the method with parentheses

# Loop through each unique MATNR
for matnr in unique_matnr:
    # 現在の日時と時刻を取得
    now = datetime.now()

    # フォーマットして出力
    formatted_time = now.strftime("%Y-%m-%d %H:%M:%S")
    print("Current time:", formatted_time)
    print(matnr)
    # Filter d1 by MATNR
    filtered_d1 = d1[d1['MATNR'] == matnr]
    print(filtered_d1['product'].unique().tolist())
    # Output the MKMNR column to a text file
    clip_path = f"Files/{CWID}/clip.txt"
    filtered_d1['MKMNR'].to_csv(clip_path, index=False, header=False)
    while True:
        if os.path.exists(clip_path):
            break
        else:
            time.sleep(1)  # 1秒待機してから再度チェック

    # Get unique values of date1 and date2
    unique_date1 = filtered_d1['date1'].unique()
    unique_date2 = filtered_d1['date2'].unique()
    
    # コマンドを構築
    command = [
        "cscript",
        "/nologo",
        "scripts\\Download_of_Inspectionlot_of_Results2.vbs",
        matnr,
        ', '.join(map(str, unique_date1)),
        ', '.join(map(str, unique_date2))
    ]

    # コマンドを実行
    subprocess.run(command, shell=True)

    while True:
        if os.path.exists(clip_path):
            os.remove(clip_path)
            time.sleep(1)  # 1秒待機してから再度チェック
        else:
            break