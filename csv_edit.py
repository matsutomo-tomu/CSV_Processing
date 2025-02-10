import pandas as pd
import os

def collect_csv_data(sheet_name):
    """シート用のCSVデータを収集する関数"""
    combined_data = []
    print(f"--- {sheet_name} のデータを入力してください ---")
    
    while True:
        # CSVファイルのパス入力
        csv_file = input(f"{sheet_name}に記入するCSVファイルのパスを入力してください（終了するには 'end' と入力）: ").strip().strip('"')
        if csv_file.lower() == "end":
            break

        # ファイル存在チェック
        if not os.path.isfile(csv_file):
            print(f"エラー: CSVファイルが見つかりません -> {csv_file}")
            continue

        try:
            # CSVを読み込み
            df = pd.read_csv(csv_file)
            combined_data.append(df)
            # 1行分の空白を挿入
            empty_row = pd.DataFrame([[""] * len(df.columns)], columns=df.columns)
            combined_data.append(empty_row)

            print(f"CSVファイルを追加しました -> {csv_file}")
        except Exception as e:
            print(f"エラー: ファイルの処理中に問題が発生しました -> {e}")
    
    return combined_data

# 保存するExcelファイル名
output_excel = input("出力するExcelファイルの名前を入力してください（拡張子は省略可）: ").strip().strip('"')
if not output_excel.endswith(".xlsx"):
    output_excel += ".xlsx"

# 保存先フォルダを指定
save_dir = input("Excelファイルを保存するフォルダのパスを入力してください（未入力の場合は現在のフォルダ）: ").strip().strip('"')
if not save_dir:
    save_dir = os.getcwd()
elif not os.path.isdir(save_dir):
    print(f"エラー: 指定したフォルダが存在しません -> {save_dir}")
    exit()

output_path = os.path.join(save_dir, output_excel)

# シート1のデータを収集
sheet1_data = collect_csv_data("シート1")

# シート2のデータを収集
sheet2_data = collect_csv_data("シート2")

# Excelファイルに書き出し
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    if sheet1_data:
        sheet1_df = pd.concat(sheet1_data, ignore_index=True)
        sheet1_df.to_excel(writer, index=False, sheet_name="シート1")
    if sheet2_data:
        sheet2_df = pd.concat(sheet2_data, ignore_index=True)
        sheet2_df.to_excel(writer, index=False, sheet_name="シート2")

print(f"Excelファイルが作成されました: {output_path}")
