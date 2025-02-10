import pandas as pd
import os

# ディレクトリ取得
script_dir = os.path.dirname(os.path.abspath(__file__))

# ファイルパス入力
csv_file1 = input("1つ目のCSVファイルのパスを入力してください: ").strip().strip('"')
csv_file2 = input("2つ目のCSVファイルのパスを入力してください: ").strip().strip('"')

# ファイル存在チェック
if not os.path.isfile(csv_file1):
    print(f"エラー: ファイルが見つかりません -> {csv_file1}")
    exit()
if not os.path.isfile(csv_file2):
    print(f"エラー: ファイルが見つかりません -> {csv_file2}")
    exit()

# データフレーム読み込み
df1 = pd.read_csv(csv_file1)
df2 = pd.read_csv(csv_file2)

# 出力ファイル名指定
output_excel = input("出力するExcelファイル名 入力 : ").strip().strip('"')
if not output_excel.endswith(".xlsx"):  # 拡張子後付け
    output_excel += ".xlsx"

# 保存先パス
xlsx_file = input("保存先のパス入力 → ")
output_path = os.path.join(xlsx_file, output_excel)

# Excelファイル作成
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='Sheet1', index=False)
    df2.to_excel(writer, sheet_name='Sheet2', index=False)

print(f"Excelファイルが作成されました: {output_path}")
