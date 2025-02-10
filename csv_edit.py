import pandas as pd
import os

# 保存するExcelファイル名
output_excel = input("出力するExcelファイルの名前を入力してください（拡張子は省略可）: ").strip().strip('"')
if not output_excel.endswith(".xlsx"):
    output_excel += ".xlsx"

# Excel Writerを準備
output_path = os.path.join(os.getcwd(), output_excel)
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    sheet_count = 1 
    
    while True:
        # 1つ目のCSVファイルのパス入力
        csv_file1 = input(f"1つ目のCSVファイルのパスを入力してください（終了するには 'end' と入力）: ").strip().strip('"')
        
        # "end"を入力して終了
        if csv_file1.lower() == "end":
            break
        
        # 2つ目のCSVファイルのパス入力
        csv_file2 = input(f"2つ目のCSVファイルのパスを入力してください: ").strip().strip('"')
        
        # ファイル存在チェック
        if not os.path.isfile(csv_file1):
            print(f"エラー: 1つ目のCSVファイルが見つかりません -> {csv_file1}")
            continue
        if not os.path.isfile(csv_file2):
            print(f"エラー: 2つ目のCSVファイルが見つかりません -> {csv_file2}")
            continue
        
        try:
            # CSV読み込み
            df1 = pd.read_csv(csv_file1)
            df2 = pd.read_csv(csv_file2)
            
            # シート名を付けてExcelに書き込む
            df1.to_excel(writer, sheet_name=f"Sheet{sheet_count}_Data1", index=False)
            df2.to_excel(writer, sheet_name=f"Sheet{sheet_count}_Data2", index=False)
            print(f"CSVファイルのデータをExcelに追加しました:\n 1つ目 -> {csv_file1}\n 2つ目 -> {csv_file2}")
            
            # シート番号を更新
            sheet_count += 1
        except Exception as e:
            print(f"エラー: CSVファイルの処理中に問題が発生しました -> {e}")

print(f"Excelファイルが作成されました: {output_path}")
