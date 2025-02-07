import os
import pandas as pd
import glob
import sys

def main():
    search_in_excelfiles(search_text)

def search_in_excelfiles(search_text):

    # カレントディレクトリ、input/outputディレクトリのPathを変数に代入
    currentdir = os.getcwd()
    inputdir = os.path.join(currentdir,"06_Search_in_ExcelFiles\\in")
    ourputdir = os.path.join(currentdir,"06_Search_in_ExcelFiles\\out")

    # Excelファイルの拡張子を持つ全ファイルを対象とする。"**"とrecursiveオプションでサブディレクトリも走査対象とする
    Excelfiles = glob.glob(os.path.join(inputdir,"**","*.xls*"),recursive=True)
    # HitしたExcelファイルを格納するためのリスト初期設定
    found_files = []

    # Excelファイル単位のループ
    for file in Excelfiles:
        try:
            excelfile = pd.ExcelFile(file)
            for sheetname in excelfile.sheet_names:
                df = excelfile.parse(sheetname,dtype=str)
                if df.astype(str).apply(lambda x: x.str.contains(search_text,na=False)).any().any():
                    found_files.append(file)
                    break
        except Exception as e :
                print(f"エラー: {file} を処理中に問題が発生しました: {e}", file=sys.stderr)

    if found_files:
        print("検索対象文字列を含むブック：")
        for f in found_files:
            print(f)
    else:
        print("検索対象文字列を含むExcelブック無し")

if __name__ == "__main__":
    search_text = input("検索対象文字列を入力：").strip()
    main()