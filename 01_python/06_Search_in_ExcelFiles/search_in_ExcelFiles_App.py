import tkinter as tk
from tkinter import messagebox
import os
import glob
import pandas as pd

###############################################################################
# inputdirに格納されているExcelファイルに検索対象となる文字列が存在するかチェックする
###############################################################################

# 検索処理
# inputdir内のExcelファイルに検索対象文字列が存在するかチェック
def search_in_excelfiles(search_text):

    # カレントディレクトリ、inputディレクトリのPathを変数に代入
    currentdir = os.getcwd()
    inputdir = os.path.join(currentdir,"06_Search_in_ExcelFiles\\in")
    # Excelファイルの拡張子を持つ全ファイルを対象とする。"**"とrecursiveオプションでサブディレクトリも走査対象とする
    Excelfiles = glob.glob(os.path.join(inputdir,"**","*.xls*"),recursive=True)
    # HitしたExcelファイルを格納するためのリスト初期設定
    found_files = []

    # Excelファイル単位のループ
    for file in Excelfiles:
        try:
            excelfile = pd.ExcelFile(file)
            for sheetname in excelfile.sheet_names:
                # 検索対象文字列を探索。探索結果をリストに格納
                df = excelfile.parse(sheetname,dtype=str)
                if df.astype(str).apply(lambda x: x.str.contains(search_text,na=False)).any().any():
                    found_files.append(file)
                    break
        except Exception as e :
                return "処理中に問題が発生しました"
    # 検索対象場が見つからない場合のmsg出力
    return found_files if found_files else ["検索対象文字列を含むExcelブック無し"]

# 検索結果の取得とGUIへの出力
# 入力フォーム、出力結果を表示するウィンドウを定義するメソッド
def on_register():
    # 入力フォームから値を取得
    target_value = entry_target_value.get()
    # 前段で定義した文字列探索メソッド呼び出してリストに格納
    vallist = search_in_excelfiles(target_value)
    
    # 文字列探索用リストの要素単位でループ
    for value in vallist:
        # フォームに入力されているかチェック。入力なければエラー表示
        if not (value):
            messagebox.showwarning("警告", "検索対象文字列を入力してください")
            return
        
        vallist = search_in_excelfiles(target_value)
        text_area.delete("1.0",tk.END)
        # テキストエリアに検索結果を表示。複数ある場合は改行する
        for value in vallist:
            text_area.insert(tk.END,f"{value}\n")

    # 実行成功のメッセージボックスを表示
    messagebox.showinfo("情報", "実行完了")

# GUIに関する設定
# アプリケーションに関する設定用変数appを初期化する。
app = tk.Tk()
app.title("ExcelファイルGrepツール")

# 出力コンソールのサイズを設定。出力される結果に合わせてウィンドウサイズを自動調整するようにした
app.columnconfigure(0,weight=1)
app.columnconfigure(1,weight=1)
app.rowconfigure(8,weight=1)

# 入力フォームを設定
label_target_value = tk.Label(app, text="検索対象文字列", anchor=tk.CENTER)
label_target_value.grid(row=0, column=0)
entry_target_value = tk.Entry(app)
entry_target_value.grid(row=0, column=1)

# 実行ボタンを作成、クリックイベントも設定
button_register = tk.Button(app, text="実行", command=on_register)
button_register.grid(row=7, column=0, columnspan=2)

# テキストエリアにスクロールバーを設定
frame = tk.Frame(app)
frame.grid(row=8,column=0,columnspan=2,sticky="nsew")

# 結果出力用にテキストエリアを作成
text_area = tk.Text(frame,wrap=tk.WORD)
text_area.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
# テキストエリアのスクロールバー配置場所を設定
scrollbar = tk.Scrollbar(frame,command=text_area.yview)
scrollbar.pack(side=tk.RIGHT,fill=tk.Y)
text_area.config(yscrollcommand=scrollbar.set)

app.mainloop()