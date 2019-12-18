# coding: UTF-8

import sys
import os
import subprocess
import random
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import cm

#グローバル変数
words = 0 #単語数の初期値
testlist = [] #テスト参照リストの初期値
problems = 0 #取得した問題数
order = '' #ランダムか昇順か
problemslist = [] #問題番号を格納するリスト
rangestart = 0 #範囲開始の変数
rangeend = 0 #範囲終了の変数

def get_file():
    global words
    global testlist

    #root.withdraw()
    fTyp=[('Text File','.txt')]
    iDir = os.path.abspath(os.path.dirname(__file__))#'/home/ユーザ名/Desktop/'
    #1つのファイルを選択する,pass取得
    titleText = 'テストファイルを選択'
    testfile = filedialog.askopenfilename(filetypes = fTyp, title = titleText, initialdir = iDir)
    testfilepath = testfile
    #pathを取得できてる→ファイルを開ける
    #print(testfilepath)

    words, testlist = set_file(testfilepath) #ここで戻り値を受け取る

    labelwords['text'] = ('単語数：%s' % words )
    if words <= 20:
        cb1['values'] = ('全単語','10','20')
    else : cb1['values'] = ('0','10','20')

    cb2['values'] = ('昇順','ランダム')


    textrangestart.delete(0,tk.END)
    textrangeend.delete(0,tk.END)

    textrangestart.insert(tk.END,'1')
    textrangeend.insert(tk.END,'%s'%words) #範囲終了初期値


    #print(words)
    #print(testlist)

def make_test():
    #ここにテスト作成の処理を入れる
    #テストの携帯は何で決める？
    #欲しいのは範囲の開始、終了、問題数、ランダムかそうでないか

    global problems #問題数
    global rangestart #範囲開始
    global rangeend #範囲終了
    global order #並べ方
    global problemslist #順番のリスト
    #このグローバル変数たちは引っ張ってくることができるのか
    #ここから関数を2つ作っていく。1つはA4、1つはB5
    if int(problems) <= 10:
        pdf_B5()

    elif int(problems)  <= 20:
        pdf_A4()

    pdfpath = os.path.abspath('Test.pdf')
    print(pdfpath)

    process = open_pdf("%s"%pdfpath, page=1)
    process.wait()

    print('OK2')

def open_pdf(filename, page=1):
    return subprocess.Popen(["C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe", '/A', 'page={:d}'.format(page), filename], stdout=subprocess.PIPE)


def pdf_B5():
    global ploblemlist
    global testlist

    pdfFile = canvas.Canvas('./Test.pdf')
    pdfFile.saveState()
    #後でテストファイルを開く処理入れる

    #テストを捨てようかと思ったけど、これを更新し続ければよくね

    pdfFile.setAuthor('Croma')
    pdfFile.setTitle('テスト作成')
    pdfFile.setSubject('Test')

    #この下をif文でコントロールする→A4,B5をグローバル変数にでもするか
    #→ファイルを読み取って、それを初期値とするシークバープログラム
    # A4
    #pdfFile.setPageSize((21.0*cm, 29.7*cm))
    # B5
    pdfFile.setPageSize((18.2*cm, 25.7*cm))

    #タイトル
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    pdfFile.setFont('HeiseiKakuGo-W5', 28)
    pdfFile.drawString(6.6*cm, 23.6*cm, '単語テスト')

    #問題番号
    pdfFile.setFont('HeiseiKakuGo-W5',12)

    Hightstart = 21.6
    for i in range(0, 10, 1):
        pdfFile.drawString(1*cm, Hightstart*cm, '%s.'% str(i+1))
        Hightstart -= 2.15

    #英語
    Hightstart = 21.63 #2回宣言しないと一回使うとなくなる
    for i in range(0,len(problemslist),1):
        pdfFile.drawString(2.2*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][1])
        Hightstart -= 2.15

    #日本語
    Hightstart = 21.5
    for i in range(0,len(problemslist),1):
        pdfFile.drawString(11*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][2])
        Hightstart -= 2.15

    pdfFile.showPage() #1ページ目確定

    #この下からまた同じように書き始める
    pdfFile.setFont('HeiseiKakuGo-W5', 28)
    pdfFile.drawString(6.6*cm, 23.6*cm, '単語テスト')

    #問題番号
    pdfFile.setFont('HeiseiKakuGo-W5',12)
    pdfFile.setLineWidth(1)

    Hightstart = 21.6
    for i in range(0, 10, 1):
        pdfFile.drawString(1*cm, Hightstart*cm, '%s.'% str(i+1))
        Hightstart -= 2

    #英語
    Hightstart = 21.63 #2回宣言しないと一回使うとなくなる
    for i in range(0,len(problemslist),1):
        pdfFile.drawString(2*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][1])
        Hightstart -= 2


    #線
    Hightstart = 21.6
    for i in range(0,len(problemslist),1):
        pdfFile.line(11*cm, Hightstart*cm, 16.5*cm, Hightstart*cm)
        Hightstart -= 2


    pdfFile.showPage() #2ページ目確定

    pdfFile.save()

def pdf_A4():
    global problemslist
    global testlist

    Hightstart = 0

    pdfFile = canvas.Canvas('./Test.pdf')
    pdfFile.saveState()

    pdfFile.setAuthor('Croma')
    pdfFile.setTitle('テスト作成')
    pdfFile.setSubject('Test')

    # A4
    pdfFile.setPageSize((21.0*cm, 29.7*cm))

    #タイトル
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    pdfFile.setFont('HeiseiKakuGo-W5', 30)
    pdfFile.drawString(7.9*cm, 28.1*cm, '単語テスト')

    #問題番号
    pdfFile.setFont('HeiseiKakuGo-W5',12)

    Hightstart = 26.5
    for i in range(0, 20, 1):
        pdfFile.drawString(1*cm, Hightstart*cm, '%s.'% str(i+1))
        Hightstart -= 1.31

    #英語
    Hightstart = 26.6 #2回宣言しないと一回使うとなくなる
    for i in range(0,len(problemslist),1):
        pdfFile.drawString(3*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][1])
        Hightstart -= 1.31

    #日本語
    Hightstart = 26.5
    for i in range(0,len(problemslist),1):
        pdfFile.drawString(14*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][2])
        Hightstart -= 1.31

    pdfFile.showPage() #1ページ目確定

    #この下からまた同じように書き始める
    #タイトル
    pdfFile.setFont('HeiseiKakuGo-W5', 30)
    pdfFile.drawString(7.9*cm, 28.1*cm, '単語テスト')

    #問題番号
    pdfFile.setFont('HeiseiKakuGo-W5',12)
    pdfFile.setLineWidth(1)

    Hightstart = 26.5
    for i in range(0, 20, 1):
        pdfFile.drawString(1*cm, Hightstart*cm, '%s.'% str(i+1))
        Hightstart -= 1.31

    #英語
    Hightstart = 26.6 #2回宣言しないと一回使うとなくなる
    for i in range(0,len(problemslist),1):
        pdfFile.drawString(3*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][1])
        Hightstart -= 1.31


    #線
    Hightstart = 26.5
    for i in range(0,len(problemslist),1):
        pdfFile.line(14*cm, Hightstart*cm, 19.4*cm, Hightstart*cm)
        Hightstart -= 1.31

    pdfFile.showPage() #2ページ目確定

    pdfFile.save()
    print('OK')

def make_flashcard():
    global problemslist
    global testlist

    pdfFile = canvas.Canvas('./Flashcard.pdf')
    pdfFile.saveState()
    #後でテストファイルを開く処理入れる

    #テストを捨てようかと思ったけど、これを更新し続ければよくね

    pdfFile.setAuthor('Croma')
    pdfFile.setTitle('テスト作成')
    pdfFile.setSubject('Flashcard')

    #この下をif文でコントロールする→A4,B5をグローバル変数にでもするか
    #→ファイルを読み取って、それを初期値とするシークバープログラム
    # A4
    pdfFile.setPageSize((21.0*cm, 29.7*cm))
    # B5
    # pdfFile.setPageSize((18.2*cm, 25.7*cm))

    #タイトル
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
    pdfFile.setFont('HeiseiKakuGo-W5', 65)
    pdfFile.drawString(2.28*cm, 14.25*cm, '%s'%testlist[0][0])

    pdfFile.showPage() #1ページ目確定

    #単語番号
    pdfFile.setFont('HeiseiKakuGo-W5',30)

    Hightstart = 26.6 #2回宣言しないと一回使うとなくなる
    for i in range(0,len(problemslist),1):

        pdfFile.drawString(0.5*cm, Hightstart*cm, '%s.'% str(i+1))

        pdfFile.drawString(3*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][1])

        pdfFile.drawString(10*cm, Hightstart*cm,'%s'%testlist[problemslist[i]][2])

        Hightstart -= 3

        if i % 9 == 8 :
            #print('OK')
            pdfFile.showPage() #1ページ目確定
            pdfFile.setFont('HeiseiKakuGo-W5',30)
            Hightstart = 26.6

    pdfFile.save()

    print('OK')

    pdfpath = os.path.abspath('Flashcard.pdf')
    print(pdfpath)

    process = open_pdf("%s"%pdfpath, page=1)
    process.wait()

    print('OK2')



def set_file(x):
    #path = testfilepath
    f = open(x,encoding='utf-8')
    pretestlist = f.readlines()
    testlist = []

    for i in pretestlist:
        testlist.append(i.strip().split('|'))

    words = len(testlist) - 1 #ファイル内の単語数

    return words, testlist

def rand_ints_nodup(a,b,k):
    ns = []
    while len(ns) < k:
        n = random.randint(a,b)
        if not n in ns:
            ns.append(n)

    return ns

def cb1_selected(event):
    global problems

    problems = hmp.get() #問題数取得
    if problems == '全単語' :
        problems = words
    #print('取得した問題数は：%s' % problems)

    #quit()

def cb2_seleted(event):
    global order
    global rangestart
    global rangeend
    global problemslist

    rangestart = textrangestart.get()
    rangeend = textrangeend.get()

    order = orders.get()
    print('選択法は：%s' % order)

    #ここの中にリストの処理を書いていこう

    if order == '昇順' :
        for i in range(int(rangestart), int(rangeend)+1, 1):
            problemslist.append(i)

        print(problemslist)
        print(len(problemslist))

    if order == 'ランダム' :
        problemslist = rand_ints_nodup(int(rangestart),int(rangeend),int(problems))

        print(problemslist)


#ウィンドウを作成
root = tk.Tk() #これ自体でウィンドウはできてる
root.title("テストツール")
root.geometry('760x420')


#ボタン、パーツ類
#現状配置をいじれていないからそのまんまの順番になってる
get_file_Button = tk.Button(root,text='テストファイルを取得')
get_file_Button["command"] = get_file
get_file_Button.pack()

labelwords = tk.Label(root, text=u'') #単語数表示
labelwords.pack()

#ここに単語の範囲入れる
labelrangestart = tk.Label(root, text=u'範囲開始:')
labelrangestart.pack()

textrangestart = tk.Entry(root)
textrangestart.insert(tk.END,'')
textrangestart.pack()

labelrangeend = tk.Label(root, text=u'範囲終了:')
labelrangeend.pack()

textrangeend = tk.Entry(root)
textrangeend.insert(tk.END,'')#'%s'%words)
textrangeend.pack()

hmp = StringVar() #how many problems 問題数
cb1 = ttk.Combobox(root, textvariable=hmp)
cb1.bind('<<ComboboxSelected>>', cb1_selected)

cb1['values'] = ('%s'%words)
cb1.set('%s'%words)
cb1.pack()

orders = StringVar()
cb2 = ttk.Combobox(root, textvariable=orders)
cb2.bind('<<ComboboxSelected>>', cb2_seleted)
#ComboboxSelectedはしっかりそういう名前にしとかないとinter動かん

cb2['values'] = ('')#('昇順','ランダム')
cb2.set('順番')
cb2.pack()


make_test_Button = tk.Button(root,text='テスト作成')
make_test_Button['command'] = make_test
make_test_Button.pack()


make_flashcard_Button = tk.Button(root,text='単語帳を作成')
make_flashcard_Button['command'] = make_flashcard
make_flashcard_Button.pack()

root.mainloop()#ここがループすることによって動いている、故最後
