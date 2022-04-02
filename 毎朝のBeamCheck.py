# coding: utf-8

import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
import pandas as pd
import subprocess                     #←エクセルを開くためだけ
import datetime                       #←今日の日付を取得
from glob import glob


#####  関数  #####     関数はGUIより前に定義

#  スピンボックスからCSVへ登録用関数
def regist_today():
    value_4E=sp_4E.get()           #sp_に変更
    value_6E=sp_6E.get()
    value_9E=sp_9E.get()
    value_12E=sp_12E.get()
    value_15E=sp_15E.get()
    value_6X=sp_6X.get()
    value_10X=sp_10X.get()
    value_4X=sp_4X.get()
    value_cal=sp_cal.get()
    filepath = glob('checkmate.csv')
    df_dailyQC =pd.read_csv('checkmate.csv',encoding="shift-jis",index_col=0)
    df_dailyQC[today]=value_4E,value_6E,value_9E,value_12E,value_15E,value_6X,value_10X,value_4X,value_cal
    df_dailyQC.to_csv('checkmate.csv',encoding="shift-jis")
    # mode='a', header=False)

# 登録確認
def ask_yn():
    if tk.messagebox.askyesno("確認", "登録しますか？"):
        value_4E=sp_4E.get()           #sp_に変更
        value_6E=sp_6E.get()
        value_9E=sp_9E.get()
        value_12E=sp_12E.get()
        value_15E=sp_15E.get()
        value_6X=sp_6X.get()
        value_10X=sp_10X.get()
        value_4X=sp_4X.get()
        value_cal=sp_cal.get()
        filepath = glob('checkmate.csv')
        df_dailyQC =pd.read_csv('checkmate.csv',encoding="shift-jis",index_col=0)
        df_dailyQC[today]=value_4E,value_6E,value_9E,value_12E,value_15E,"",value_6X,value_10X,value_4X,value_cal
        df_dailyQC.to_csv('checkmate.csv',encoding="shift-jis")
        root.destroy()
        subprocess.Popen(r'毎朝のBeamCheck',shell=True)

    
#メインウィンドウ　GUI 
root=tk.Tk()
root.geometry("1350x650")
root.title('毎朝のビームチェック！')

today = datetime.date.today()

#リニアックトラブルシューティングを開くだけ=================================================
EXCEL = r'リニアックトラブルシューティング.xlsx'
subprocess.Popen(EXCEL,shell=True)
#===========================================================================================

#ene=['4E','6E','9E','12E','15E','6X','10X','4X','校正']


#データフレーム用のフレーム―――――――――――――――――――――――――――――――――
frame_DataFrame = ttk.Labelframe(root,text = "Time Series",padding = 10)
#frame_DataFrame.grid(row =1,column =0,padx = 10,pady = 10)
frame_DataFrame.place(x="20",y="30")


###　CSVからツリービューに読み出し　###　最終行から7行抽出　---------------------------------
def main():
    _df1 =pd.read_csv('checkmate.csv',encoding="shift-jis",index_col=0)
    _df =_df1.fillna('')              #####空欄を空白で埋める
    df =_df.iloc[0:10,-8:] #dfは最終列から７列分取り出し
    tree=ttk.Treeview(frame_DataFrame,show="headings")
    
    style = ttk.Style()        ###スタイル###
    style.configure("Treeview", font=('Meiryo UI', 14,), rowheight=40)
    style.configure("Treeview.Heading",font=('Meiryo UI', 11,'bold','italic','underline'),foreground="blue")
    
    tree['column']= ("Beam",) +tuple(df)
    tree.heading("Beam", text="Beam")
    tree.column("Beam",width=60,anchor='center')
    for c in df:
        tree.heading(c,text=c)
        tree.column(c,anchor=tk.CENTER,width = 120, stretch = False)
    
    for i,row in enumerate(df.itertuples()):
        #tree.insert("","end",tags=i,values=row)
        tree.insert("","end",tags=i,values=row)
    
    tree.pack()

    
    #tree.bind('<<TreeviewSelect>>',selected)
    
    #tree.mainloop()

if __name__ == '__main__':
    main() 

   
#ここまで――――――――――――――――――――――――――――――――――――――


#登録ボタン用のフレーム―――――――――――――――――――――――――――――
frame_regist = ttk.Labelframe(root,text = "Registration",padding = 10)
#frame_regist.grid(row=3,column =2,columnspan=2,padx = 5,pady = 5)
frame_regist.place(x="1175",y="525")


###　登録ボタン　###
save_button = tk.Button(frame_regist, text = "登　録",command=ask_yn)
save_button.grid(row =1,column =0,padx = 10,pady = 10,ipady = 5)

###　終了ボタン　###
ex_button = tk.Button(frame_regist, text = "終  了",command=lambda:root.destroy())
ex_button.grid(row =1,column =1,padx = 10,pady = 10,ipady = 5)


#ここまで――――――――――――――
#========================================================================================
#  スピンボックスのフレーム ―-
#frame_spinbox = ttk.Labelframe(root,text = "Today's Output",padding = 10)
#frame_spinbox.grid(row =1,column =2,sticky=tk.W)
#frame_spinbox.place(x="1060",y="106")

#  スピンボックスのフレーム 電子線―
frame_E= ttk.Labelframe(root,text = "E-ray",padding = 10)
#frame_spinbox.grid(row =1,column =2,sticky=tk.W)
frame_E.place(x="1075",y="60")
#  スピンボックスのフレーム Ｘ線--
frame_X = ttk.Labelframe(root,text = "X-ray",padding = 10)
#frame_spinbox.grid(row =1,column =2,sticky=tk.W)
frame_X.place(x="1075",y="290")
#  スピンボックスのフレーム 校正―
frame_cal = ttk.Labelframe(root,text = "キャリブレーション",padding = 10)
#frame_spinbox.grid(row =1,column =2,sticky=tk.W)
frame_cal.place(x="1075",y="430")
#========================================================================================

#  ラベルフレーム
#frame_label= ttk.Labelframe(root,text ="LQ")
#frame_label.place(x="25",y="110")
###　ラベル　###
label=ttk.Label(root,text ='校正したら△▽クリック↑',foreground="blue")
label.place(x=1060,y=500)

###  4E  ###
spinbox_4E = tk.StringVar()
spinbox_4E.set(100.0)
sp_4E = ttk.Spinbox(frame_E,textvariable = spinbox_4E,from_=-80, to =120,increment = 0.1, width = 7,font=("",18))
sp_4E.grid(column=1)
###  6E  ###
spinbox_6E = tk.StringVar()
spinbox_6E.set(100.0)
sp_6E = ttk.Spinbox(frame_E,textvariable = spinbox_6E,from_=-80, to =120,increment = 0.1 ,width = 7,font=("",18))
sp_6E.grid(column=1)
###  9E  ###
spinbox_9E = tk.StringVar()
spinbox_9E.set(100.0)
sp_9E = ttk.Spinbox(frame_E,textvariable = spinbox_9E,from_=-80, to =120,increment = 0.1 ,width = 7,font=("",18))
sp_9E.grid(column=1)
###  12E  ###
spinbox_12E = tk.StringVar()
spinbox_12E.set(100.0)
sp_12E = ttk.Spinbox(frame_E,textvariable = spinbox_12E,from_=-80, to =120,increment = 0.1 ,width = 7,font=("",18))
sp_12E.grid(column=1)
###  15E  ###
spinbox_15E = tk.StringVar()
spinbox_15E.set(100.0)
sp_15E = ttk.Spinbox(frame_E,textvariable = spinbox_15E,from_=-80, to =120,increment = 0.1 ,width = 7,font=("",18))
sp_15E.grid(column=1)
###  6X  ###
spinbox_6X = tk.StringVar()
spinbox_6X.set(100.0)
sp_6X = ttk.Spinbox(frame_X,textvariable = spinbox_6X,from_=-80, to =120,increment = 0.1 ,width = 7,font=("",18))
sp_6X.grid(column=1)
###  10X  ###
spinbox_10X = tk.StringVar()
spinbox_10X.set(100.0)
sp_10X = ttk.Spinbox(frame_X,textvariable = spinbox_10X,from_=-80, to =120,increment = 0.1 ,width = 7,font=("",18))
sp_10X.grid(column=1)
###  4X  ###
spinbox_4X = tk.StringVar()
spinbox_4X.set(100.0)
sp_4X = ttk.Spinbox(frame_X,textvariable = spinbox_4X,from_=-80, to =120,increment = 0.1 ,width = 7,font=("",18))
sp_4X.grid(column=1)

###  Cal  ###
cal=["実施"]
spinbox_cal = tk.StringVar()
#spinbox_cal.set("-")
sp_cal = ttk.Spinbox(frame_cal,textvariable = spinbox_cal,width = 7,font=("",18),values=cal,)
sp_cal.grid(column=1)
#ここまで――――――――――――――――――――――――――――――――――――――――――――――――――



###　キャリブのチェックボックス　###
#chk = tk.Checkbutton(root,text="キャリブレーション実施",font=("",12))
#chk.place(x="890",y="50")

root.mainloop()