import pandas as pd
import os
import difflib
import re
import shutil
import sys

CWID = os.getenv("USERNAME")

print("CSVファイルを作成")


# カラム名を変換する関数
def rename_columns(df):
    new_columns = {}
    for col in df.columns:
        if "Text" in col:
            new_columns[col] = "product"
        elif "オブジェクト" in col:
            new_columns[col] = "product"
        else:
            new_columns[col] = col
    df.rename(columns=new_columns, inplace=True)

    if "product" not in df.columns:
        print("No columns renamed to 'product'. Press any key to exit.")
        input()  # ユーザーが何かキーを押すまで待機
        sys.exit()  # スクリプトを終了
    return df


df = pd.read_csv(f"Files/{CWID}/table2.csv", encoding="CP932")
codes = df["MATNR"].apply(lambda x: f"{int(x):08d}").unique()

# エクセルファイルの読み込み
excel_file_path = f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/Trend2/AdjustResultIndividually.xlsx"
excel_data = pd.read_excel(excel_file_path)

for code in codes:
    dt1 = pd.read_csv(
        "rawdata/results/code" + code + ".txt",
        encoding="CP932",
        sep="\t",
        low_memory=False,
    )
    print(code)
    dt1.dropna(axis=1, how="all", inplace=True)
    dt1.columns = dt1.columns.str.replace(" ", "")
    dt1["ERGEBNIS"] = dt1["ERGEBNIS"].replace("---", "")
    dt11 = dt1[
        ["PRUEFLOS", "MATNR", "HSDAT", "CHARG", "KURZTEXT", "ERGEBNIS", "VERWMERKM"]
    ]

    filtered_df = df[df["MATNR"] == int(code)]
    mkmnr_order = filtered_df["MKMNR"].unique().tolist()

    # mkmnr_orderの値がdt11のVERWMERKM列に存在するか確認
    valid_mkmnr_order = [x for x in mkmnr_order if x in dt11["VERWMERKM"].values]
    invalid_mkmnr_order = [
        x for x in dt11["VERWMERKM"].unique() if x not in mkmnr_order
    ]

    # if len(invalid_mkmnr_order) > 0:
    #     print("invalid mkmnr error")
    #     print(len(invalid_mkmnr_order))
    #     input()

    order = valid_mkmnr_order + invalid_mkmnr_order
    dt12 = dt11.set_index("VERWMERKM").loc[order].reset_index()

    # ピボット前のKURZTEXTの順番を取得
    unique_order = dt12["KURZTEXT"].unique()

    # ピボット処理
    dt13 = dt12.pivot(
        index=["PRUEFLOS", "MATNR", "HSDAT", "CHARG"],
        columns="KURZTEXT",
        values="ERGEBNIS",
    )

    # 列名を保持するために再インデックス
    dt13 = dt13.reindex(columns=unique_order)
    dt13.reset_index(inplace=True)
    # 列名の変更
    dt13 = dt13.rename(
        columns={
            "PRUEFLOS": "Inspection lot",
            "CHARG": "ロット",
            "HSDAT": "製造日",
            "MATNR": "品目",
        }
    )

    if any("REMARK" in col for col in dt1.columns):
        dt14 = dt1[
            ["PRUEFLOS", "MATNR", "HSDAT", "CHARG", "KURZTEXT", "REMARK", "VERWMERKM"]
        ]
        dt15 = dt14.set_index("VERWMERKM").loc[order].reset_index()
        dt15["KURZTEXT"] = dt15["KURZTEXT"] + "_Remark"

        # ピボット前のKURZTEXTの順番を取得
        unique_order = dt15["KURZTEXT"].unique()
        # ピボット処理
        dt16 = dt15.pivot(
            index=["PRUEFLOS", "MATNR", "HSDAT", "CHARG"],
            columns="KURZTEXT",
            values="REMARK",
        )

        # 列名を保持するために再インデックス
        dt16 = dt16.reindex(columns=unique_order)
        dt16.reset_index(inplace=True)
        # 列名の変更
        dt16 = dt16.rename(
            columns={
                "PRUEFLOS": "Inspection lot",
                "CHARG": "ロット",
                "HSDAT": "製造日",
                "MATNR": "品目",
            }
        )
        dt16.dropna(axis=1, how="all", inplace=True)
        dt17 = pd.merge(dt13, dt16, how="outer")
    else:
        dt17 = dt13

    # エクセルデータの読み込みに対して処理を行う
    for index, row in excel_data.iterrows():
        inspection_lot = row["Inspection lot"]
        lot = row["ロット"]
        item = row["品目"]
        colname = row["colname"]
        value = row["value"]

        # dt17の該当する行を見つける
        matching_rows = dt17[
            (dt17["品目"] == item)
            & (dt17["ロット"] == lot)
            & (dt17["Inspection lot"] == inspection_lot)
        ]

        # 一致する行がある場合、その行のcolname列の値をvalueに置換
        if not matching_rows.empty:
            dt17.loc[matching_rows.index, colname] = value

    # 溶出試験
    # 溶出試験及びRemarkが含まれる列を特定
    columns_to_process = [
        col for col in dt17.columns if "溶出率" in col and "Remark" in col
    ]

    # 各列に対して処理を行う
    for col in columns_to_process:
        print(f"Processing column: {col}")  # 処理の列名を表示

        # 平均、最小値、最大値と個々の値を抽出する関数
        def extract_values(cell):
            if isinstance(cell, str):  # セルが文字列の場合のみ処理
                # ()がある場合のパターン
                match_with_parentheses = re.match(
                    r"(?:Ave\.?|Av|ave)?(\d+)\((\d+(?:[.,]\d+){5,11})\)", cell
                )
                if match_with_parentheses:
                    avg_value = int(match_with_parentheses.group(1))
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_with_parentheses.group(2))
                    ]
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # ()がない場合のパターン
                match_without_parentheses = re.match(r"(\d+(?:[.,]\d+){5,11})", cell)
                if match_without_parentheses:
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_without_parentheses.group(1))
                    ]
                    avg_value = sum(individual_values) // len(individual_values)
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # ()が前にある場合のパターン
                match_parentheses_first = re.match(
                    r"\((\d+(?:[.,]\d+){5,11})\)Ave\.?(\d+)", cell
                )
                if match_parentheses_first:
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_parentheses_first.group(1))
                    ]
                    avg_value = int(match_parentheses_first.group(2))
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # ()のみの場合のパターン
                match_only_parentheses = re.match(r"\((\d+(?:[.,]\d+){5,11})\)", cell)
                if match_only_parentheses:
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_only_parentheses.group(1))
                    ]
                    avg_value = sum(individual_values) // len(individual_values)
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # []のみの場合のパターン
                match_square_brackets = re.match(r"\[(\d+(?:[.,]\s?\d+){5,11})\]", cell)
                if match_square_brackets:
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_square_brackets.group(1))
                    ]
                    avg_value = sum(individual_values) // len(individual_values)
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # av.XX []の場合のパターン
                match_av_square_brackets = re.match(
                    r"av\.?(\d+)\s?\[(\d+(?:[.,]\s?\d+){5,11})\]", cell
                )
                if match_av_square_brackets:
                    avg_value = int(match_av_square_brackets.group(1))
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_av_square_brackets.group(2))
                    ]
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # AV:XX ()の場合のパターン
                match_av_colon_parentheses = re.match(
                    r"AV:?(\d+)\s?\((\d+(?:[.,]\d+){5,11})\)", cell
                )
                if match_av_colon_parentheses:
                    avg_value = int(match_av_colon_parentheses.group(1))
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_av_colon_parentheses.group(2))
                    ]
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # Ave:XX ()の場合のパターン
                match_ave_colon_parentheses = re.match(
                    r"Ave:?(\d+)\s?\((\d+(?:[.,]\d+){5,11})\)", cell
                )
                if match_ave_colon_parentheses:
                    avg_value = int(match_ave_colon_parentheses.group(1))
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_ave_colon_parentheses.group(2))
                    ]
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # Ave;XX ()の場合のパターン
                match_ave_semicolon_parentheses = re.match(
                    r"Ave;?(\d+)\s?\((\d+(?:[.,]\d+){5,11})\)", cell
                )
                if match_ave_semicolon_parentheses:
                    avg_value = int(match_ave_semicolon_parentheses.group(1))
                    individual_values = [
                        int(x)
                        for x in re.split(
                            r"[.,]", match_ave_semicolon_parentheses.group(2)
                        )
                    ]
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # XX ()の場合のパターン
                match_number_parentheses = re.match(
                    r"(\d+)\s?\((\d+(?:[.,]\d+){5,11})\)", cell
                )
                if match_number_parentheses:
                    avg_value = int(match_number_parentheses.group(1))
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_number_parentheses.group(2))
                    ]
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # AvXX(XX,XX,...)のように最後に)がない場合のパターン
                match_missing_parentheses = re.match(
                    r"(?:Ave\.?|Av|ave)?(\d+)\((\d+(?:[.,]\d+){5,11})$", cell
                )
                if match_missing_parentheses:
                    avg_value = int(match_missing_parentheses.group(1))
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_missing_parentheses.group(2))
                    ]
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values
                # ,と.が混在している場合のパターン
                match_mixed_delimiters = re.match(r"(\d+(?:[.,]\d+){5,11})", cell)
                if match_mixed_delimiters:
                    individual_values = [
                        int(x)
                        for x in re.split(r"[.,]", match_mixed_delimiters.group(1))
                    ]
                    avg_value = sum(individual_values) // len(individual_values)
                    min_value = min(individual_values)
                    max_value = max(individual_values)
                    return avg_value, min_value, max_value, individual_values

            return None, None, None, None  # それ以外の場合Noneを返す

        # 平均、最小値、最大値と個々の値を格納するリスト
        avg_values = []
        min_values = []
        max_values = []
        individual_values_list = [[] for _ in range(12)]  # 最大12個の個々の値に対応

        # 各セルに対して処理を行う
        for cell in dt17[col]:
            avg_value, min_value, max_value, individual_values = extract_values(cell)
            if individual_values is not None:
                avg_values.append(avg_value)
                min_values.append(min_value)
                max_values.append(max_value)
                for i in range(len(individual_values)):
                    individual_values_list[i].append(individual_values[i])
                # 個々の値が2個未満の場合、Noneを追加
                for i in range(len(individual_values), 12):
                    individual_values_list[i].append(None)
            else:
                avg_values.append(None)
                min_values.append(None)
                max_values.append(None)
                for i in range(12):
                    individual_values_list[i].append(None)

        colname = col.replace("_Remark", "")

        # ()がある行が1つでもある場合のみ平均の列を作成
        if any(avg is not None for avg in avg_values):
            dt17[f"{colname}_平均"] = avg_values
        for i in range(12):
            dt17[f"{colname}_個々_{i+1}"] = individual_values_list[i]

        # 最小値と最大値の列を作成
        dt17[f"{colname}_最小値"] = min_values
        dt17[f"{colname}_最大値"] = max_values

    dt17.dropna(axis=1, how="all", inplace=True)
    dt2 = pd.read_csv(
        "rawdata/qa33/q_" + code + ".txt", encoding="CP932", sep="\t", skiprows=1
    )
    dt2.dropna(axis=1, how="all", inplace=True)
    dt2 = rename_columns(dt2)
    dt2.columns = dt2.columns.str.replace(" ", "")
    dt2 = dt2.rename(
        columns={
            "検査ロット": "Inspection lot",
            "product": "品名",
            "品目": "品目コード",
        }
    )

    if "使用決定日" in dt2.columns:
        dt2["判定日"] = dt2["使用決定日"].fillna(dt2["コード日"])
    else:
        dt2["判定日"] = dt2["コード日"]
    dt21 = dt2[["Inspection lot", "品目コード", "品名", "ロット", "判定日"]]
    dt3 = pd.merge(dt21, dt17, how="right")
    selected_columns = [
        "判定日",
        "製造日",
        "品目コード",
        "品名",
        "Inspection lot",
        "ロット",
    ] + list(dt3.columns[7:])
    dt4 = dt3[selected_columns]
    dt5 = dt4.sort_values(by="判定日", ascending=True)

    table = pd.read_excel("List.xlsx")
    path = f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/dash/TestResults"

    name = dt5["品名"].value_counts().idxmax()
    name = name.replace("/", " ")
    print(name)
    filename = os.path.join(path, f"{name}_code{code}.csv")
    if int(code) not in table["品目コード"].values:
        # コードが見つからなかった場合
        new_row = pd.DataFrame({"品目コード": [int(code)], "品名": name})
        table = pd.concat([table, new_row], ignore_index=True)
        # List.xlsxに上書き保存
        table.to_excel("List.xlsx", index=False)
        dt5.to_csv(filename, index=False, encoding="CP932")
    else:
        # コードが一致する行を見つける
        matched_row = table[table["品目コード"] == int(code)]

        # 一致する行の'team'列の値を取得
        team = matched_row["team"].values[0]
        # pathと'team'の値を結合
        if pd.isna(team):
            dt5.to_csv(filename, index=False, encoding="CP932")
        else:
            if os.path.exists(filename):
                # ファイルを削除
                os.remove(filename)
            filename = os.path.join(path, str(team), f"{name}_code{code}.csv")
            # ディレクトリが存在しない場合作成
            os.makedirs(os.path.dirname(filename), exist_ok=True)
            dt5.to_csv(filename, index=False, encoding="CP932")

            # 新しいファイルパスの先頭部分
#            new_base_path = (
#                "//Bjpyoks0062/byl_psj$/QLC/004_試験・出荷判定2_トレンド・P2R試験結果"
#            )

            # 既存のファイルパスの先頭部分を新しいパスに置換
#            new_output_file_path = filename.replace(
#                f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/dash/TestResults",
#                new_base_path,
#            ).replace(".csv", ".xlsx")

            # ファイルを保存するディレクトリが存在しない場合作成
#            new_output_directory = os.path.dirname(new_output_file_path)
#            os.makedirs(new_output_directory, exist_ok=True)

#            # 新しいファイルパスにデータフレームをExcelファイルとして保存
#            dt5.to_excel(new_output_file_path, index=False)


def filter_and_merge_excel(input_excel_path, search_directory, check_csv_path):
    # エクセルファイルを読み込む
    df = pd.read_excel(input_excel_path)

    # merge列に値がある行をフィルタリング
    filtered_df = df[df["merge"].notna()]

    # merge値に対して操作を行う
    for merge_value in filtered_df["merge"].unique():
        # 現在のmerge値に対応する行をフィルタリング
        current_df = filtered_df[filtered_df["merge"] == merge_value]

        # フィルタリングされた行の品目コードを取得し、文字列に変換
        item_codes = current_df["品目コード"].astype(str).tolist()

        # 保持したCSVファイルを読み込む
        check_df = pd.read_csv(check_csv_path, encoding="CP932")

        # 品目列の名前を確認（仮に'品目'とする）
        if "MATNR" in check_df.columns:
            # 品目列に品目コードが含まれているか確認
            check_df["MATNR"] = check_df["MATNR"].astype(str)  # 品目列を文字列に変換
            matching_items = check_df[check_df["MATNR"].isin(item_codes)]

            if not matching_items.empty:
                print(f"Matching items found for merge value '{merge_value}':")

                # 抽出されたファイルを保存するリスト
                extracted_files = []
                output_directories = set()

                # 保持されたディレクトリからファイルを再帰的に抽出
                for root, dirs, files in os.walk(search_directory):
                    for file in files:
                        if file.endswith(".csv") and any(
                            item_code in file for item_code in item_codes
                        ):
                            extracted_files.append(os.path.join(root, file))
                            output_directories.add(root)

                # 抽出されたファイルをマージ
                merged_df = pd.DataFrame()
                for file in extracted_files:
                    temp_df = pd.read_csv(file, encoding="CP932")
                    merged_df = pd.concat([merged_df, temp_df], ignore_index=True)

                # 判定日列を基準に並び替える
                if "判定日" in merged_df.columns:
                    merged_df["判定日"] = pd.to_datetime(
                        merged_df["判定日"], errors="coerce"
                    )
                    merged_df = merged_df.sort_values(by="判定日")

                # 出力ディレクトリにファイルを保存
                for output_directory in output_directories:
                    output_file_path = os.path.join(
                        output_directory, f"{merge_value}_merged.csv"
                    )
                    merged_df.to_csv(output_file_path, index=False, encoding="CP932")
                    print(f"Output file created: {output_file_path}")

                    # 新しいファイルパスの先頭部分
#                    new_base_path = "//Bjpyoks0062/byl_psj$/QLC/004_試験・出荷判定2_トレンド・P2R試験結果"

                    # 既存のファイルパスの先頭部分を新しいパスに置換
#                    new_output_file_path = output_file_path.replace(
#                        f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/dash/TestResults",
#                        new_base_path,
#                    ).replace(".csv", ".xlsx")

                    # ファイルを保存するディレクトリが存在しない場合作成
#                    new_output_directory = os.path.dirname(new_output_file_path)
#                    os.makedirs(new_output_directory, exist_ok=True)

                    # 新しいファイルパスにデータフレームをExcelファイルとして保存
#                    merged_df.to_excel(new_output_file_path, index=False)


# 使用例
input_excel_path = (
    f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/Trend2/List.xlsx"
)
search_directory = (
    f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/dash/TestResults"
)
check_csv_path = f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/Trend2/Files/{CWID}/table2.csv"

filter_and_merge_excel(input_excel_path, search_directory, check_csv_path)


def copy_files(src_dir, dst_dir):
    # ソースディレクトリが存在するか確認
    if not os.path.exists(src_dir):
        print(f"ソースディレクトリ {src_dir} が存在しません。")
        return

    # 目的ディレクトリが存在しない場合作成
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)

    # ソースディレクトリ内の全ファイルを取得
    files = os.listdir(src_dir)

    for file in files:
        # ファイルのフルパスを作成
        src_file = os.path.join(src_dir, file)
        dst_file = os.path.join(dst_dir, file)

        # CSVファイルを読み込む
        df = pd.read_csv(src_file, encoding="CP932")

        # inspection lot列の値が90000000000以下ものに絞り込む
        filtered_df = df[df["Inspection lot"] <= 890000000000]

        # 絞り込んだデータが存在する場合のみコピー
        if not filtered_df.empty:
            filtered_df.to_csv(dst_file, index=False, encoding="CP932")
            print(f"{file} をコピーしました。")
        else:
            print(f"{file} は条件に合致するデータがありません。")


# 使用例
copy_files(
    f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/dash/TestResults/SFP",
    f"C:/Users/{CWID}/Bayer/PRC public - General/製造Dashboard/QC",
)
copy_files(
    f"C:/Users/{CWID}/OneDrive - Bayer/Desktop/QualityData/dash/TestResults/RAWM",
    f"C:/Users/{CWID}/Bayer/PRC public - General/製造Dashboard/CoA",
)
