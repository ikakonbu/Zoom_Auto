
import tkinter as tk #GUIライブラリ
import tkinter.ttk as ttk #GUIライブラリのラジオボタン用
from tkinter import messagebox #メッセージボックスを出す
import sys, math, time, datetime #いろいろ。
from datetime import date
import subprocess #url開くのに必要
from threading import Thread #並列実行でメッセージを出す
import configparser #confファイルを読む
import platform #実行OSを検知

import os
os.chdir(os.path.dirname(os.path.abspath(__file__))) #カレントディレクトリをファイルのあるディレクトリに設定(画像ファイルをよむため)

ini = configparser.ConfigParser() #設定ファイルを読み込み
ini.read('setting.conf')
classdata = configparser.ConfigParser()
classdata.read('class.conf','UTF-8')



#ここから下は概ねウインドウの設定
# メインウィンドウを定義
root = tk.Tk()
root.title('群大Zoom自動実行クライアント')
root.geometry("410x480")
if platform.system() != 'Windows':
	root.geometry("450x570")
root.resizable(0,0)
root.configure(bg="#FDF9F1")

#タブの設定
nb = ttk.Notebook(width=400, height=470)
tab1 = tk.Frame(nb, bg="#FDF9F1")
tab2 = tk.Frame(nb, bg="#FDF9F1")
nb.add(tab1, text=' 表示 ', padding=5)
nb.add(tab2, text=' 設定 ', padding=5)
nb.pack(expand=1, fill='both')

#現在時刻を表示するラベル
label = tk.Label(tab1, fg='#310D04', bg="#FDF9F1", text=str(datetime.datetime.today().second), anchor="w", font=("M+ 2p",20))
label.place(x=-100, y=0)

#画像を読み込んでおく
inclass = tk.PhotoImage(file="images/講義中.png")
endclass =  tk.PhotoImage(file="images/講義終了.png")
startclass = tk.PhotoImage(file="images/講義開始.png")
noclass = tk.PhotoImage(file="images/空きコマ.png")
holiday = tk.PhotoImage(file="images/休日.png")
nothing = tk.PhotoImage(file="images/なにもない.png")
goschool = tk.PhotoImage(file="images/登校日.png")

#画像はキャンバスの上に読み込む必要があるのでキャンバスを設定
canvas = tk.Canvas(tab1, bg="#FDF9F1", width=400, height=400, highlightbackground="#FDF9F1")
canvas.place(x=0, y=42)
canvas.create_image(0, 0, image=nothing, anchor=tk.NW)

kyuzitu = 0 #マニュアルで休日を設定された時に使用
imagetimer = [1,-1] #指定した時間たったら画像に変更するフラグ
oldimage = 6 #1個前の画像を保存しておく
nowimage = 6 #現在の状態
jobid=None #並列実行時に使う
daylist = [ #プログラム⇔設定ファイルのデータの変換用
["月", 0],
["火", 1],
["水", 2],
["木", 3],
["金", 4]]
classtime = ['8:35','10:15','12:35','14:15','15:55']#各時限の5分前の時刻を記入しておく

def imageprint(u): #画像を変更するときの関数
    if u==1: canvas.create_image(0, 0, image=inclass, anchor=tk.NW)
    if u==2:
        canvas.create_image(0, 0, image=endclass, anchor=tk.NW)
        imagetimer[0]=6
        imagetimer[1]=6 #講義終了後5分で戻す
    if u==3: canvas.create_image(0, 0, image=startclass, anchor=tk.NW)
    if u==4: canvas.create_image(0, 0, image=noclass, anchor=tk.NW)
    if u==5: canvas.create_image(0, 0, image=holiday, anchor=tk.NW)
    if u==6: canvas.create_image(0, 0, image=nothing, anchor=tk.NW)
    if u==7: canvas.create_image(0, 0, image=goschool, anchor=tk.NW)

def btn_click(x): #休日モードのボタンを押されたときの関数
    global kyuzitu,oldimage,nowimage
    if kyuzitu == 1:
        kyuzitu=0
        btn['text'] = "休日モードにする"
        imageprint(oldimage)
        print(oldimage)
        nowimage = oldimage
        btn.pack(x=290, y=5)
        if jobId is not None and x==1:  # 休日モードを外した際に一回メインルーチンを回して確認
            tab1.after_cancel(jobId)
            jobId=None
            loop()

    if kyuzitu == 0:
        kyuzitu=1
        btn['text'] = "休日モードを外す"
        oldimage = nowimage
        imageprint(5)
        btn.pack(x=290, y=5)




#休日モードのボタンの配置
btn = tk.Button(tab1, fg='#310D04', bg='#FDF3E3', text='休日モードにする', font=("M+ 2p",8), command = lambda: btn_click(1))
btn.place(x=290, y=5)

kmode = tk.BooleanVar()
if ini['setting']['setting1'] == "True":
    kmode.set(True)
else:
    kmode.set(False)

kmode_box = tk.Checkbutton(tab2, variable=kmode, text='休日モードが1日で切れるようにする', bg='#FDF9F1', font=("M+ 2p",12))
kmode_box.place(x=20, y=40)  #休日モードを消せる設定とチェック状態を格納する変数を宣言

settime = tk.Entry(tab2, width=10, font=("M+ 2p",12))
settime.place(x=210, y=90)
settime.insert(tk.END,str(ini['setting']['setting2']))
label2 = tk.Label(tab2, text = "休日モードを外す時間", bg="#FDF9F1" ,font=("M+ 2p",12,"bold"))
label2.place(x=20,y=90) #いつ消すかをマニュアルで設定できるようにここに書いておく




#マニュアルでZoomを開くところの表示設定
label3 = tk.Label(tab2, text = "マニュアル実行", bg="#FDF9F1" ,font=("M+ 2p",12,"bold"))
label3.place(x=20,y=250)
label4 = tk.Label(tab2, height=2, text = "        　 曜日の　　　　時限 ", bg="#f7bc7c" ,font=("M+ 2p",12))
label4.place(x=30,y=295)
if platform.system() != 'Windows':
    label4.place(x=50,y=300)

Comfont = ("M+ 2p" , '16')

manual_day_text=tk.StringVar()
manual_day = ttk.Combobox(tab2, font=Comfont, values=("月","火","水","木","金"), textvariable=manual_day_text, state="readonly", width=2 )
manual_day.current(0)
manual_day.place(x=35, y=304)

manual_time_text=tk.StringVar()
manual_time = ttk.Combobox(tab2, font=Comfont, values=("1","2","3","4","5"), textvariable=manual_time_text, state="readonly", width=2 )
manual_time.current(0)
manual_time.place(x=150, y=304)
if platform.system() != 'Windows':
    manual_time.place(x=135, y=304)

manual_btn = tk.Button(tab2, fg='#310D04', bg='#FDF3E3', text='    実行    ', font=("M+ 2p",14), command = lambda: manual_do(manual_day_text.get(),manual_time_text.get()))
manual_btn.place(x=270, y=297)



#学籍番号によって動作を変えるハイブリッド授業用設定
label4 = tk.Label(tab2, text = "学籍番号設定(ハイブリッド用)", bg="#FDF9F1" ,font=("M+ 2p",12,"bold"))
label4.place(x=20,y=150)

#ラジオボタンの状態を持っておく変数
radio_num = tk.StringVar()
#設定ファイルから前回の設定を再現
radio_num.set(ini['setting']['setting3'])

#ttkのパーツはstyleで設定しないとフォントとか変わらないからここでstyleを決める
style = ttk.Style()
style.configure("TRadiobutton",font=("M+ 2p",12))

rb1 = ttk.Radiobutton(tab2, text='奇数',value='1',variable=radio_num)
rb1.place(x=50, y=190)

rb2 = ttk.Radiobutton(tab2, text='偶数',value='2',variable=radio_num,)
rb2.place(x=130, y=190)

rb3 = ttk.Radiobutton(tab2, text='設定なし',value='3',variable=radio_num)
rb3.place(x=240, y=190)


#メッセージボックス出力関数
def printmessage(x,y):
    ret = messagebox.showinfo(x,y)

#終了時に警告&設定書き込みする部分
def on_closing():
    if messagebox.askokcancel("修了確認", "終了しますか？"):
        config = configparser.RawConfigParser()
        section1 = 'setting'
        config.add_section("setting")
        config.set(section1, 'setting1', str(kmode.get()))
        config.set(section1, 'setting2', settime.get())
        config.set(section1, 'setting3', radio_num.get())
        file =  open('setting.conf', 'w')
        config.write(file)
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)


#マニュアルでZoomを開く関数
def manual_do(doday, dotime):
    if classdata[doday+"曜日のID"][dotime] != "aki":
        murl = 'zoommtg:\"//zoom.us/join?confno=' + classdata[doday+"曜日のID"][dotime] + '&pwd=' + classdata[doday+"曜日のpass"][dotime] + "\""
        if platform.system()=='Windows':
            subprocess.Popen('start ' + murl , shell=True) #WIndowsコマンドとしてコマンドプロンプトに渡して実行させる
        else:
            subprocess.Popen('open ' + murl , shell=True) #Mac環境下の場合はこっち

        if classdata[doday + "曜日の説明"][dotime] != "なし":
            Thread(target=printmessage , args=('説明',classdata[doday + "曜日の説明"][dotime])).start()
    else:
        Thread(target=printmessage , args=("あきこま", "その時間は空きコマだよ")).start()





# メインルーチン
def loop():

    jobid = tab1.after(60000, loop) #1分ごとに定期的にこいつを動かして表示を更新
    t = datetime.datetime.today()
    nowtime = str(t.hour) + ":" + str(format(t.minute,'02')) #現在時刻を取る
    nowday = t.weekday()#曜日も取得
    weeknum = datetime.date.today().isocalendar()[1] #週番号をとる
    Hybrid_setting = int(radio_num.get()) #ハイブリッドの設定してるか取る

    #学籍番号の設定があるときはここで投稿日か判定
    if weeknum%2 == 0:
        if nowday%2 ==0:
            if Hybrid_setting==1: Hybrid_flag=1
            if Hybrid_setting==2: Hybrid_flag=0
        else:
            if Hybrid_setting==1: Hybrid_flag=0
            if Hybrid_setting==2: Hybrid_flag=1
    else:
        if nowday%2 ==0:
            if Hybrid_setting==1: Hybrid_flag=0
            if Hybrid_setting==2: Hybrid_flag=1
        else:
            if Hybrid_setting==1: Hybrid_flag=1
            if Hybrid_setting==2: Hybrid_flag=0

    if Hybrid_setting == 3:  #ハイブリッドの設定してなかったら全蒸し
        Hybrid_flag = 0

    if kmode.get()==True and kyuzitu==1 and nowtime == settime.get() :
        Thread(target=printmessage , args=("注意","休日モードを解除しました")).start()
        btn_click(0)


    if kyuzitu == 0: #休日モードはoff?
        if nowday<5: #平日？休日？
            if Hybrid_flag == 0:
                for i in range(5):
                    if nowtime == str(classtime[i]):           #該当する次官になった？
                        if classdata[str(daylist[nowday][0]) + "曜日のID"][str(i+1)] != 'aki':          #空きコマは無視する
                            url = 'zoommtg:\"//zoom.us/join?confno=' + classdata[daylist[nowday][0]+"曜日のID"][str(i+1)] + '&pwd=' +  classdata[daylist[nowday][0]+"曜日のpass"][str(i+1)] + "\""    #URLスキームの形に変形
                            if platform.system()=='Windows':
                                subprocess.Popen('start ' + url , shell=True) #WIndowsコマンドとしてコマンドプロンプトに渡して実行させる
                            else:
                                subprocess.Popen('open ' + url , shell=True) #Mac環境下の場合はこっち

                            if classdata[str(daylist[nowday][0]) + "曜日の説明"][str(i+1)] != "なし":
                                Thread(target=printmessage , args=('説明',classdata[str(daylist[nowday][0]) + "曜日の説明"][str(i+1)])).start()

                            imageprint(3)
                            imagetimer[0]=1
                            imagetimer[1]=6
                        else:
                            imageprint(4)
                            imagetimer[0]=4
                            imagetimer[1]=6
            else:
                imageprint(7) #登校日の表示
        else:
            imageprint(5) #休日の表示

    label['text'] = "現在時刻   " + str(datetime.datetime.today().hour) + ":" + str(datetime.datetime.today().minute).zfill(2)
    label.place(x=-0, y=0) #時刻表示を更新

    if imagetimer[1] > 0: imagetimer[1] = imagetimer[1]-1
    if imagetimer[1] == 0:
        imagetimer[1] = -1
        imageprint(imagetimer[0]) #タイマーで予約された画像があれば指定した時間に表示
        if imagetimer[0]==1: #講義が始まったら90分後に終了画像を出すように予約
            imagetimer[0]=2
            imagetimer[1]=90
        if imagetimer[0]==4: #休講だったら90分後になにもない状態にする
            imagetimer[0]=6
            imagetimer[1]=90

    if kyuzitu == 1: imageprint(5)


# 最初の起動
loop()

# メインループ
root.mainloop()
