
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
root.iconbitmap('images/icon.ico')
root.title('群大Zoom自動実行クライアント')
root.geometry("410x480")
if platform.system() != 'Windows':
	root.geometry("450x570")
root.resizable(0,0)
root.configure(bg="#FDF9F1")

#フォント設定置き場
Comfont = ("M+ 2p" , '16')
stdfont = ("M+ 2p" , '12')
boldfont = ("M+ 2p", '12', "bold")
btnfont = ("M+ 2p" , '8')
timefont = ("M+ 2p" , '20')

#タブの設定
nb = ttk.Notebook(width=400, height=470)
tab1 = tk.Frame(nb, bg="#FDF9F1")
tab2 = tk.Frame(nb, bg="#FDF9F1")
nb.add(tab1, text=' 表示 ', padding=5)
nb.add(tab2, text=' 設定 ', padding=5)
nb.pack(expand=1, fill='both')

#現在時刻を表示するラベル
label = tk.Label(tab1, fg='#310D04', bg="#FDF9F1", text=str(datetime.datetime.today().second), anchor="w", font=timefont)
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
jobid=None #並列実行時に使う
zoom_started=[0,0,0,0,0] #すでにzoomが立ち上がっているかの判定
daylist = [ #プログラム⇔設定ファイルのデータの変換用
["月", 0],
["火", 1],
["水", 2],
["木", 3],
["金", 4]]
nowtime = datetime.datetime.now()
classtime = [datetime.datetime(nowtime.year,nowtime.month,nowtime.day,8,40,0),
             datetime.datetime(nowtime.year,nowtime.month,nowtime.day,10,20,0),
             datetime.datetime(nowtime.year,nowtime.month,nowtime.day,12,40,0),
             datetime.datetime(nowtime.year,nowtime.month,nowtime.day,14,20,0),
             datetime.datetime(nowtime.year,nowtime.month,nowtime.day,16,0,0)]#各時限の開始時刻を記入しておく

def imageprint(u): #画像を変更するときの関数
    if u==1: canvas.create_image(0, 0, image=inclass, anchor=tk.NW)
    if u==2: canvas.create_image(0, 0, image=endclass, anchor=tk.NW)
    if u==3: canvas.create_image(0, 0, image=startclass, anchor=tk.NW)
    if u==4: canvas.create_image(0, 0, image=noclass, anchor=tk.NW)
    if u==5: canvas.create_image(0, 0, image=holiday, anchor=tk.NW)
    if u==6: canvas.create_image(0, 0, image=nothing, anchor=tk.NW)
    if u==7: canvas.create_image(0, 0, image=goschool, anchor=tk.NW)

def btn_click(x): #休日モードのボタンを押されたときの関数
    global kyuzitu
    if kyuzitu == 1:
        kyuzitu=0
        btn['text'] = "休日モードにする"
        btn.pack(x=290, y=5)
        if jobId is not None and x==1:  # 休日モードを外した際に一回メインルーチンを回して確認
            tab1.after_cancel(jobId)
            jobId=None
            loop()

    if kyuzitu == 0:
        kyuzitu=1
        btn['text'] = "休日モードを外す"
        imageprint(5)
        btn.pack(x=290, y=5)
 
def update_classtime(): #授業時間の日時がずれると正しく計算できなくなるので,日付変わったら更新する関数
    nowtime = datetime.datetime.now()
    classtime = [datetime.datetime(nowtime.year,nowtime.month,nowtime.day,8,40,0,0),
                datetime.datetime(nowtime.year,nowtime.month,nowtime.day,10,20,0,0),
                datetime.datetime(nowtime.year,nowtime.month,nowtime.day,12,40,0,0),
                datetime.datetime(nowtime.year,nowtime.month,nowtime.day,14,20,0,0),
                datetime.datetime(nowtime.year,nowtime.month,nowtime.day,16,0,0,0)]


#休日モードのボタンの配置
btn = tk.Button(tab1, fg='#310D04', bg='#FDF3E3', text='休日モードにする', font=btnfont, command = lambda: btn_click(1))
btn.place(x=290, y=5)

kmode = tk.BooleanVar()
if ini['setting']['setting1'] == "True":
    kmode.set(True)
else:
    kmode.set(False)

kmode_box = tk.Checkbutton(tab2, variable=kmode, text='休日モードが1日で切れるようにする', bg='#FDF9F1', font=stdfont)
kmode_box.place(x=20, y=40)  #休日モードを消せる設定とチェック状態を格納する変数を宣言

settime = tk.Entry(tab2, width=10, font=stdfont)
settime.place(x=210, y=90)
settime.insert(tk.END,str(ini['setting']['setting2']))
label2 = tk.Label(tab2, text = "休日モードを外す時間", bg="#FDF9F1" ,font=boldfont)
label2.place(x=20,y=90) #いつ消すかをマニュアルで設定できるようにここに書いておく




#マニュアルでZoomを開くところの表示設定
label3 = tk.Label(tab2, text = "マニュアル実行", bg="#FDF9F1" ,font=boldfont)
label3.place(x=20,y=250)
label4 = tk.Label(tab2, height=2, text = "        　曜日の 　　　　　時限 ", bg="#f7bc7c" ,font=stdfont)
label4.place(x=20,y=295)
if platform.system() != 'Windows':
    label4.place(x=40,y=300)

manual_day_text=tk.StringVar()
manual_day = ttk.Combobox(tab2, font=Comfont, values=("月","火","水","木","金"), textvariable=manual_day_text, state="readonly", width=2 )
manual_day.current(0)
manual_day.place(x=25, y=304)

manual_time_text=tk.StringVar()
manual_time = ttk.Combobox(tab2, font=Comfont, values=("1-2","3-4","5-6","7-8","9-10"), textvariable=manual_time_text, state="readonly", width=4 )
manual_time.current(0)
manual_time.place(x=130, y=304)
if platform.system() != 'Windows':
    manual_time.place(x=115, y=304)

tab1.option_add("*TCombobox*Listbox*Font", stdfont)
manual_btn = tk.Button(tab2, fg='#310D04', bg='#FDF3E3', text='    実行    ', font=("M+ 2p",14), command = lambda: manual_do(manual_day_text.get(),str(int((int(manual_time_text.get()[0])+1)/2)) ))
manual_btn.place(x=270, y=297)



#学籍番号によって動作を変えるハイブリッド授業用設定
label4 = tk.Label(tab2, text = "学籍番号設定(ハイブリッド用)", bg="#FDF9F1" ,font=boldfont)
label4.place(x=20,y=150)

#ラジオボタンの状態を持っておく変数
radio_num = tk.StringVar()
#設定ファイルから前回の設定を再現
radio_num.set(ini['setting']['setting3'])

#ttkのパーツはstyleで設定しないとフォントとか変わらないからここでstyleを決める
style = ttk.Style()
style.configure("TRadiobutton",font=stdfont)

rb1 = ttk.Radiobutton(tab2, text='奇数',value='1',variable=radio_num)
rb1.place(x=50, y=190)

rb2 = ttk.Radiobutton(tab2, text='偶数',value='2',variable=radio_num,)
rb2.place(x=130, y=190)

rb3 = ttk.Radiobutton(tab2, text='設定なし',value='3',variable=radio_num)
rb3.place(x=240, y=190)


config_btn = tk.Button(tab2, fg='#595959', bg='#FDF5F1', text=' 設定ファイルを開く ', font=("M+ 2p",10), command = lambda: config_open())
config_btn.place(x=250,y=390)

#メッセージボックス出力関数
def printmessage(x,y):
    ret = messagebox.showinfo(x,y)

def config_open():
	subprocess.Popen('start ' + '.\class.conf', shell=True)
	Thread(target=printmessage , args=("注意", "変更した設定は再起動後に有効になります")).start()

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

    jobid = root.after(60000, loop) #1分ごとに定期的にこいつを動かして表示を更新
    nowtime = datetime.datetime.now()
    nowday = nowtime.weekday()#曜日も取得
    weeknum = datetime.date.today().isocalendar()[1] #週番号をとる
    Hybrid_setting = int(radio_num.get()) #ハイブリッドの設定してるか取る
    if(nowtime.hour==0 and nowtime.minute==0): update_classtime() #日付変わったら授業時間のdatetimeを更新

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

    if kmode.get()==True and kyuzitu==1 and str(nowtime.hour) + ":" + nowtime.minutes.zfill(2) == settime.get() :
        Thread(target=printmessage , args=("注意","休日モードを解除しました")).start()
        btn_click(0)


    if kyuzitu == 0: #休日モードはoff?
        if nowday<5: #平日？休日？
            if Hybrid_flag == 0:
                #初期画像表示
                imageprint(6)
                for i in range(5):
                    #現在時刻との差分をとる
                    sabun_sign=1
                    sabuntemp = nowtime - classtime[i]   #timedelta型で-の値は自動的に正に直されるので、自分で符号を判定
                    if sabuntemp.days==-1: sabun_sign=-1 #授業時間と現在時刻との差分(正負ありの秒数)を生成
                    sabun = abs(sabuntemp).seconds*sabun_sign


                    #授業5分前~授業中(差分-300~5400秒)になった？
                    if sabun >= -300 and sabun < 5400 and zoom_started[i]==0:           
                        zoom_started[i]=1

                        #空きコマでなければZoom実行
                        if classdata[str(daylist[nowday][0]) + "曜日のID"][str(i+1)] != 'aki':
                            url = 'zoommtg:\"//zoom.us/join?confno=' + classdata[daylist[nowday][0]+"曜日のID"][str(i+1)] + '&pwd=' +  classdata[daylist[nowday][0]+"曜日のpass"][str(i+1)] + "\""    #URLスキームの形に変形
                            if platform.system()=='Windows':
                                subprocess.Popen('start ' + url , shell=True) #WIndowsコマンドとしてコマンドプロンプトに渡して実行させる
                            else:
                                subprocess.Popen('open ' + url , shell=True) #Mac環境下の場合はこっち

                            #クラスの説明があれば表示
                            if classdata[str(daylist[nowday][0]) + "曜日の説明"][str(i+1)] != "なし":
                                Thread(target=printmessage , args=('説明',classdata[str(daylist[nowday][0]) + "曜日の説明"][str(i+1)])).start()

                            imageprint(3) #5分後に1番の画像を表示
                        else:
                            imageprint(4) #5分後に4版を表示
                    else:
                        zoom_started[i]=0


                    #授業開始時(差分0~60秒)に授業中の画像を表示
                    if sabun >= 0 and sabun <= 60: 
                        if classdata[str(daylist[nowday][0]) + "曜日のID"][str(i+1)] != 'aki': imageprint(1)


                    #いままでやってた授業が終わったら(差分90分~91分)お疲れ様の画像出してあげる
                    if sabun >= 5400 and sabun <= 5460: 
                         if classdata[str(daylist[nowday][0]) + "曜日のID"][str(i+1)] != 'aki': imageprint(2)

            else:
                imageprint(7) #登校日の表示
        else:
            imageprint(5) #休日の表示

    label['text'] = "現在時刻   " + str(nowtime.hour) + ":" + str(nowtime.minute).zfill(2)
    label.place(x=-0, y=0) #時刻表示を更新
    


# 最初の起動
loop()

# メインループ
root.mainloop()
