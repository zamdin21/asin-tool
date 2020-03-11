# -*- coding: utf-8 -*-
import math
import time
import datetime
import os
import re
import sys
import pyperclip
import asyncio
import aiohttp
import requests
import subprocess
from pymongo import MongoClient
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from tkinter import filedialog
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import ProcessPoolExecutor
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from voluptuous import Schema, Match
from func_timeout import func_timeout, FunctionTimedOut

USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0'
headers = {'User-Agent': USER_AGENT}
cookies = dict(cookies_are='working')
semaphore = asyncio.Semaphore(10)  # 最大同時ダウンロード数を3に制限するためのセマフォを作成
# ASINが10桁の大文字アルファベットと数字のみで構成されているかチェックするためのスキーマを作成
schema = Schema(Match(r'[A-Z\d]{10}'))
# pass.txtから各パスを取得
try:
    with open('pass.txt', 'r') as r:
        all_pass_text = r.read()
        ASIN_PATH = ''
        MONGO_DB_PATH = ''
        CHROME_PATH = ''
        path_dict = {'ASIN_PATH': ASIN_PATH,
                     'MONGO_DB_PATH': MONGO_DB_PATH, 'CHROME_PATH': CHROME_PATH}
        for name, var in path_dict.items():
            var = re.search(r'{0}.*「(.+)」'.format(name),
                            all_pass_text)
            if not var:
                print('「pass.txt」内の{0}が空欄でないか確認してください。'.format(name))
                sys.exit()
            path_dict['{0}'.format(name)] = var.group(1)
except FileNotFoundError:
    print('「pass.txt」が実行フォルダ内に存在しているか確認してください')
    sys.exit()

# TAB1 関数 ----------------------------------------------------------------

class Asinfetch():
    "AmazonのすべてのページからASINを抽出するクラス"

    def __init__(self, url):
        self.url = url
        self.asin_list = []

    async def requests_url_perpage(self):
        "各ページのURLをコルーチン関数にリクエストする"
        # ページの最大数を取得するために、最初に一度リクエストする
        response = requests.get(self.url, headers=headers, cookies=cookies)
        soup = BeautifulSoup(response.content, 'lxml')
        print(soup.title.text)
        # セラーリサーチの場合はプライム関係なく取得
        if not 'marketplaceID' in self.url:
            if not 'Amazonプライム対象商品' in soup.title.text:
                messagebox.showinfo(
                    'Page Error', u'プライムで絞り込んでいないか、ロボット対策の可能性があります。プライムチェックしている場合は一度ブラウザを再起動して再度実行してください。それでもダメな場合は、検索キーワードを変えてください。')
                m_label_1_1.set('')
                return 0
        try:
            self.page_num = int(soup.find_all(class_='a-disabled')[1].text)
        except Exception:
            try:
                # たまにpagnDisabledクラス属性がないページがある。その場合、pagnLinkから一番後ろのものをとる
                self.page_num = int(soup.find_all(
                    'span', class_='pagnLink')[-1].text)
            except Exception:
                self.page_num = 1
        # カテゴリーとキーワードをセット
        category_fetch = soup.find('option', selected='selected').text
        category_text.set(category_fetch)
        keyword_fetch = soup.find('input', id='twotabsearchtextbox')['value']
        keyword.set(re.sub(r'　', ' ', keyword_fetch))  # 全角スペースを半角スペースに直してセット
        urls = []
        for i in range(self.page_num):
            urls.append(self.url + '&page={0}'.format(i + 1))
        # セッションオブジェクト作成
        async with aiohttp.ClientSession() as session:
            self.session = session
            coroutines = []
            for i, page_url in enumerate(urls):
                coroutine = self.scrape_amazon_page(i, page_url)
                coroutines.append(coroutine)
            # コルーチンを完了した順に返す
            for coroutine in asyncio.as_completed(coroutines):
                await coroutine

    async def scrape_amazon_page(self, i, page_url):
        "AmazonのページからASINを抽出する"
        with await semaphore:
            await asyncio.sleep(1)
            # 非同期にリクエストを送り、レスポンスヘッダを取得する
            response = await self.session.get(page_url, headers=headers)
            soup = BeautifulSoup(await response.read(), 'lxml')
            product_list = soup.find_all(class_='s-result-item')
            print('Processing {0}/{1} page...'.format(i + 1, self.page_num))
            print('商品数：{0}'.format(len(product_list)))
            for product_one in product_list:
                try:
                    # 予約というテキストが検出した場合は予約商品なのでリストに追加しない
                    reserve = product_one.find_all(
                        text=re.compile('予約'), limit=1)
                    if reserve:
                        continue
                except:
                    pass
                try:
                    # スポンサープロダクトチェックがoffの場合、スポンサープロダクトを除外する
                    if sponsor_product_check.get() == 'off':
                        sponsor = product_one.find_all(
                            text=re.compile('スポンサー'), limit=1)
                        if sponsor:
                            continue
                except:
                    pass
                # 禁止ワードが入っている商品を除外
                banword_str = banword.get('1.0', 'end -1c')
                if not banword_str == '':
                    banword_list = banword_str.split('\n')
                    ban_flag = False
                    try:
                        title_txt = product_one.select_one(
                            'h2[data-attribute]').text
                        for bantxt in banword_list:
                            if bantxt in title_txt:
                                ban_flag = True
                                continue
                        if ban_flag:
                            continue
                    except:
                        pass
                try:
                    price_text = product_one.find(
                        'span', class_='a-price-whole').text
                    price = int(
                        ''.join(re.findall(r'[\d]+', price_text)))  # 数字だけをリストに抽出して繋げる。int化
                    if price == 0:  # たまに価格が０のものがある（原因不明）
                        continue
                except:
                    continue  # s-price属性がないものがある（恐らくAmazonChoice製品）
                try:
                    prime_notation = product_one.find('i', attrs={'aria-label':True}).get('aria-label')
                except AttributeError:
                    print('Not get ASIN.', file=sys.stderr)
                    return
                    prime_notation = ''
                if price_updown.get() == r'\500↓':
                    if price <= 500 and 'プライム' in prime_notation:  # 「prime」表記があるものだけ抽出（あわせ買い、ただの配送料無料を除外）
                        self.asin_append(product_one)
                elif price_updown.get() == r'\500↑':
                    if price > 500 and 'プライム' in prime_notation:
                        self.asin_append(product_one)
                elif price_updown.get() == 'ALL':
                    if  'プライム' in prime_notation:
                        self.asin_append(product_one)

    def asin_append(self, product_one):
        "asin_listにasinを追加する"
        try:
            asin = product_one.get('data-asin')
            self.asin_list.append(asin)
        except:
            print('Asin not get.', file=sys.stderr)
            pass
            """
            try:
                schema(asin)
            except Exception:
                print('Schema Error! {}'.format(i))
            """


def asinfetchBtn_clicked():
    "コルーチンに分けてASIN抽出"
    url = tab1_url_text.get()
    if not url:
        messagebox.showinfo('URL None', u'URLを入力してください')
        return
    if not re.match(r'https://www.amazon.co.jp', url):
        messagebox.showinfo('URL Error', u'AmazonのRLを入力してください')
        return
    if price_updown.get() == 'Price':
        messagebox.showinfo('Price Error', u'Priceを選択してください')
        return
    fetch_ins = Asinfetch(url)  # Asinfetchオブジェクト作成
    start_time = time.time()
    # イベントループ取得
    # イベントループで()を実行し、完了するまで待つ
    loop = asyncio.get_event_loop()
    loop.run_until_complete(fetch_ins.requests_url_perpage())
    end_time = round(time.time() - start_time, 2)
    print('')
    print('総ASIN数：{0}'.format(len(fetch_ins.asin_list)))
    str_asin_list = '\n'.join(fetch_ins.asin_list)
    txt1.delete('1.0', 'end -1c')
    txt1.insert('1.0', str_asin_list)
    pyperclip.copy(str_asin_list)  # クリップボードにコピー
    messagebox.showinfo(
        'Success', '総ASIN数: {0}\n経過時間： {1}秒'.format(len(fetch_ins.asin_list), end_time))
    if asin_save_check.get() == 'on' and not category_text.get() == 'カテゴリー' and fetch_ins.asin_list:
        splitBtn_clicked(url)
        saveBtn_clicked_first()


def asinfetchBtn_clicked_first():
    "Fetchボタンを押したとき、スレッドを分けることでバックグラウンド処理"
    # m_label_1_1.set('ASIN抽出中...')
    asinfetchBtn_clicked()
    """
    excuter = ThreadPoolExecutor(max_workers=1)
    excuter.submit(asinfetchBtn_clicked)
    loop.run_in_executor(executor, asinfetchBtn_clicked)
    loop.run_forever()
    loop.close()
    """


def referenceBtn_clicked():
    "ファイル参照"
    fTyp = [("txt", "txt")]
    #iDir = os.path.abspath(os.path.dirname(__file__))
    iDir = path_dict['ASIN_PATH']
    filepath = filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
    file1.set(filepath)


def saveBtn_clicked_first():
    "Saveボタンを押したとき、スレッドを分けることでバックグラウンド処理"
    m_label_1_1.set('MongoDBに保存中...')
    excuter = ThreadPoolExecutor(max_workers=1)
    excuter.submit(saveBtn_clicked)


def saveBtn_clicked():
    "MongoDBにASINを保存"
    try:
        db = client.ASIN
    except Exception as e:
        print('例外args:', e.args)
        messagebox.showinfo('DBError', 'MongoDBに接続できません。エラー内容をご確認ください')
        m_label_1_1.set('')
    pbval = 0
    start_time = time.time()
    text_path = file1.get()
    if text_path == '':
        messagebox.showinfo('Not File path', u'ファイルパスを入力してください')
        return
    category = category_text.get()
    if category == 'カテゴリー':
        messagebox.showinfo('Category None', u'カテゴリーを選択してください')
        return
    try:
        with open(text_path) as r:
            asin_list = r.readlines()
    except FileNotFoundError:
        messagebox.showinfo('Not File', u'存在するファイルパスを入力してください')
        return
    duplicate_asin_list = []
    success_asin_list = []
    insert_list = []
    collection = db[category]
    key = re.sub(r'[　、\s,]+', '+', keyword.get()).strip()
    if not key:
        key = "None"
    # 新しいASINの場合はMongoDBに追加
    for asin in asin_list:
        duplicate_flag = False
        newline_pos = asin.find('\n')  # 改行コード削除
        if newline_pos > 0:
            asin = asin[:newline_pos]
        if collection.find_one({'asin': asin}):    # ASIN重複チェック(無いならNoneが返る)
            duplicate_asin_list.append(asin)
            duplicate_flag = True
        if not duplicate_flag:
            insert_list.append({'key': key, 'asin': asin})
            success_asin_list.append(asin)
        pbval = pbval + 1
    if insert_list:
        collection.insert_many(insert_list)  # １度でまとめて挿入した方が断然早い
        collection.create_index('asin')  # インデックスを張る
    elapased_time = time.time() - start_time

    messagebox.showinfo('Processed', 'DBに挿入したASIN数：{0}\n重複したASIN数　　：{1}\n\n処理時間：{2} 秒'.format(
        len(success_asin_list), len(duplicate_asin_list), round(elapased_time, 2)))
    a = '\n'.join(success_asin_list)
    if asin_duplicate_check.get() == 'off':  # 重複チェックoffの場合のみ、重複していないASINのみテキスト欄に挿入し、クリップボードにコピー
        txt1.delete('1.0', 'end -1c')
        txt1.insert('1.0', a)
        pyperclip.copy(a)
    m_label_1_1.set('完了')
    time.sleep(3)
    m_label_1_1.set('')


def exitBtn_clicked():
    "ウィンドウ閉じる"
    root.quit()


def dammy():
    "未実装ボタン用"
    pass


def asin_copyBtn_clicked():
    "テキストボックスのASINをコピー"
    str_asin_list = txt1.get('1.0', 'end -1c')
    pyperclip.copy(str_asin_list)
    asin_list = str_asin_list.split('\n')
    messagebox.showinfo('Copy', 'コピーしたASIN総数：' + str(len(asin_list)))


def deleteBtn_clicked():
    "テキストボックスのASINを削除"
    txt1.delete('1.0', 'end -1c')


def delete_banword_Btn_clicked():
    "テキストボックスのASINを削除"
    banword.delete('1.0', 'end -1c')


def goto_amazon():
    url = "https://www.amazon.co.jp/"
    # Pathオブジェクトではsubprocessでエラーになるため、strに変換
    subprocess.Popen(
        [str(Path(r'{0}'.format(path_dict['CHROME_PATH'])).joinpath('chrome.exe')), url])


def tab1_deleteBtn_clicked():
    "ファイル、カテゴリ、キーワード欄の削除"
    file1.set('')
    category_text.set('カテゴリー')
    keyword.set('')
    tab1_url_text.set('')


def splitBtn_clicked(url=''):
    "ASINを分割してファイル出力"
    asin_list = txt1.get('1.0', 'end -1c')
    if asin_list == '':
        messagebox.showinfo('Input Error', u'ASINを入力してください')
        return
    category = category_text.get()
    if category == 'カテゴリー':
        messagebox.showinfo('Category None', u'カテゴリーを選択してください')
        return
    updown = price_updown.get()
    if updown == 'Price':
        messagebox.showinfo('Input Error', u'Priceを選択してください')
        return
    elif updown == r'\500↑':
        updown_name = '500円以上'
    elif updown == r'\500↓':
        updown_name = '500円以下'
    elif updown == r'ALL':
        updown_name = 'ALL'
    max_nums = max_num.get()
    if max_nums == 'Num':
        messagebox.showinfo('Input None', u'Numを選択してください')
        return
    if 'marketplaceID' in url:
        save_path = Path(r'{0}'.format(path_dict['ASIN_PATH'])).joinpath('{0}\{1}\{2}'.format(
            updown_name, 'SHOP', category_text.get().strip()))
    else:
        save_path = Path(r'{0}'.format(path_dict['ASIN_PATH'])).joinpath('{0}\{1}\{2}'.format(
            updown_name, category_text.get().strip(), re.sub(r'[　、\s,]+', '+', keyword.get())))
    save_path.mkdir(exist_ok=True, parents=True)
    path = save_path.joinpath('asin_{}.txt'.format(len(asin_list.split('\n'))))
    file1.set(save_path)
    with path.open('w', encoding='utf-8')as f:
        f.write(asin_list)
    asin_list = asin_list.split('\n')
    repeat_num = len(asin_list) / max_nums
    for i in range(math.ceil(repeat_num)):
        split_asin_list = asin_list[max_nums *
                                    i:max_nums * (i + 1)]
        str_asin = '\n'.join(split_asin_list)
        path = save_path.joinpath(str(i + 1) + '.txt')
        with path.open('w', encoding='utf-8') as f:
            f.write(str_asin)


def first_connect_monodb():
    "ツール起動時にMongoDBからショップリストを取得する。その際MongoDBが起動していなかった場合に、自動で立ち上げる"
    "失敗した際は１回だけ再試行する"
    connect_flag = False
    try:
        retailer_list = func_timeout(3, get_shoplist)  # 3秒でタイムアウトする
    except FunctionTimedOut:
        print('MongoDBを起動します。')
        connect_mongodb_clicked()
        connect_flag = True
        retailer_list = []

    if retailer_list:
        return retailer_list
    else:
        if not connect_flag:
            print('MongoDBに接続できないので、MongoDBを起動します。')
            connect_mongodb_clicked()
        else:
            print('Connecting...')
        time.sleep(1)
        try:
            retailer_list = func_timeout(3, get_shoplist)  # 3秒でタイムアウトする
        except:
            print('有在庫タブのショップリストを取得できませんでした。直す場合は本ツールを再起動してください。')
        if retailer_list:
            print('Success.')
            return retailer_list
        else:
            print('有在庫タブのショップリストを取得できませんでした。直す場合は本ツールを再起動してください。')


def get_shoplist():
    "MongoDBからショップリストを取得"
    db = client.Purchase
    collection = db.ShopList
    retailer_list = [data['ShopName'] for data in collection.find()]
    return retailer_list


def get_creditlist():
    "MongoDBからクレジットカードリストを取得"
    db = client.Purchase
    collection = db.CreditCard
    creditcard_list = [data['CreditCard'] for data in collection.find()]
    return creditcard_list


def connect_mongodb_clicked():
    "新しくコマンドプロンプトを立ち上げ、MongoDBに接続"
    subprocess.Popen(
        ["start", "cmd", "/k", 'mongod --dbpath {0}'.format(path_dict['MONGO_DB_PATH'])], shell=True)
    p = subprocess.Popen(
        ['mongo'], shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, encoding='utf-8')
    with open('cmd.txt', 'r') as f:
        print(p.communicate(f.read())[0])


def tab1_exit_btn_clicked():
    "ツールを閉じる"
    root.quit()


# TAB2 関数 ----------------------------------------------------------------


class mycalendar(Frame):
    "カレンダー作成クラス"

    def __init__(self, master, command_name, **kwargs):
        "初期化メソッド"
        Frame.__init__(self, master, **kwargs)
        self.master = master
        self.command_name = command_name
        # 現在の日付を取得
        now = datetime.datetime.now()
        # 現在の年と月を属性に追加
        self.year = now.year
        self.month = now.month

        # frame_top部分の作成
        frame_top = Frame(self)
        frame_top.pack(pady=5)
        self.previous_month = Button(frame_top, text="<")
        self.previous_month.bind("<1>", self.change_month)
        self.previous_month.pack(side="left", padx=10)
        self.current_year = Label(frame_top, text=self.year, font=("", 18))
        self.current_year.pack(side="left")
        self.current_month = Label(
            frame_top, text=self.month, font=("", 18))
        self.current_month.pack(side="left")
        self.next_month = Button(frame_top, text=">")
        self.next_month.bind("<1>", self.change_month)
        self.next_month.pack(side="left", padx=10)

        # frame_week部分の作成
        frame_week = Frame(self)
        frame_week.pack()
        self.cal_y, self.cal_x = [3, 3]
        width_weekly = 5
        button_mon = Label(frame_week, text="Mon",
                           width=width_weekly, anchor=CENTER)
        button_mon.grid(column=0, row=0, ipady=self.cal_y, ipadx=self.cal_x)
        button_tue = Label(frame_week, text="Tue",
                           width=width_weekly, anchor=CENTER)
        button_tue.grid(column=1, row=0, ipady=self.cal_y, ipadx=self.cal_x)
        button_wed = Label(frame_week, text="Wed",
                           width=width_weekly, anchor=CENTER)
        button_wed.grid(column=2, row=0, ipady=self.cal_y, ipadx=self.cal_x)
        button_thu = Label(frame_week, text="Thu",
                           width=width_weekly, anchor=CENTER)
        button_thu.grid(column=3, row=0, ipady=self.cal_y, ipadx=self.cal_x)
        button_fri = Label(frame_week, text="Fri",
                           width=width_weekly, anchor=CENTER)
        button_fri.grid(column=4, row=0, ipady=self.cal_y, ipadx=self.cal_x)
        button_sta = Label(frame_week, text="Sat",
                           width=width_weekly, anchor=CENTER, foreground="blue")
        button_sta.grid(column=5, row=0, ipady=self.cal_y, ipadx=self.cal_x)
        button_san = Label(frame_week, text="San",
                           width=width_weekly, anchor=CENTER, foreground="red")
        button_san.grid(column=6, row=0, ipady=self.cal_y, ipadx=self.cal_x)

        # frame_calendar部分の作成
        self.frame_calendar = Frame(self)
        self.frame_calendar.pack(padx=2, pady=2)

        # 日付部分を作成するメソッドの呼び出し
        self.create_calendar(self.year, self.month)

    def create_calendar(self, year, month):
        "指定した年(year),月(month)のカレンダーウィジェットを作成する"

        # ボタンがある場合には削除する（初期化）
        try:
            for key, item in self.day.items():
                item.destroy()
        except:
            pass

        # calendarモジュールのインスタンスを作成
        import calendar
        cal = calendar.Calendar()
        # 指定した年月のカレンダーをリストで返す
        days = cal.monthdayscalendar(year, month)
        # 日付スタイル
        style_cBtn = Style()
        style_cBtn.configure(
            'myday.TButton', font=('Helvetica', 14)
        )
        # 日付ボタンを格納する変数をdict型で作成
        self.day = {}
        # for文を用いて、日付ボタンを生成
        for i in range(0, 42):
            c = i - (7 * int(i / 7))
            r = int(i / 7)
            try:
                # 日付が0でなかったら、ボタン作成
                if days[r][c] != 0:
                    self.day[i] = Button(
                        self.frame_calendar, text=days[r][c], style='myday.TButton', width=4, command=self.input_day(days[r][c]))
                    self.day[i].grid(
                        column=c, row=r, ipady=self.cal_y, ipadx=self.cal_x)
            except:
                # 月によっては、i=41まで日付がないため、日付がないiのエラー回避が必要
                break

    def change_month(self, event):
        "押されたラベルを判定し、月の計算"
        if event.widget["text"] == "<":
            self.month -= 1
        else:
            self.month += 1
        # 月が0、13になったときの処理
        if self.month == 0:
            self.year -= 1
            self.month = 12
        elif self.month == 13:
            self.year += 1
            self.month = 1
        # frame_topにある年と月のラベルを変更する
        self.current_year["text"] = self.year
        self.current_month["text"] = self.month
        # 日付部分を作成するメソッドの呼び出し
        self.create_calendar(self.year, self.month)

    def input_day(self, day):
        "カレンダーのクリックした日付をテキストボックスに挿入。commandの仕様上、関数内関数として定義"
        def input_main():
            if self.command_name == 'period_first':
                period_text.set(
                    '{}/{}/{}'.format(self.year, self.month, day))
            if self.command_name == 'period_second':
                period_text2.set(
                    '{}/{}/{}'.format(self.year, self.month, day))
            if self.command_name == 'buyday':
                buyday_text.set(
                    '{}/{}/{}'.format(self.year, self.month, day))
            self.master.destroy()
        return input_main


def create_calender_window(command_name):
    "カレンダーウィンドウ作成"
    root_calender = Tk()
    root_calender.title('Calender')
    calender = mycalendar(root_calender, command_name)
    calender.pack()
    root_calender.mainloop()


def tab2_saveBtn_clicked_first():
    "Saveボタンを押したとき、スレッドを分けることでバックグラウンド処理"
    excuter = ThreadPoolExecutor(max_workers=1)
    excuter.submit(tab2_saveBtn_clicked)


def tab2_saveBtn_clicked():
    "MongoDBに購入情報を保存"
    OrderNumber = tab2_order_number_text.get().strip()  # 前後の空白は削除
    Date = buyday_text.get().strip()
    Retailer = retailer_list_text.get().strip()
    ProductName = tab2_main_text['product_name'].get().strip()
    Asin = tab2_main_text['asin'].get().strip()
    Num = tab2_main_text['num'].get()
    BuyPrice = tab2_main_text['buy_price'].get().strip()
    MarketPrice = tab2_main_text['market_price'].get().strip()
    Expenses = tab2_main_text['expenses'].get().strip()
    RealCost = tab2_main_text['real_cost'].get().strip()
    BenefitPlans = tab2_main_text['benefit_plans'].get().strip()
    Breakeven_point = tab2_main_text['breakeven_point'].get().strip()
    Point1 = (tab2_shop_list[0].get().strip(),
              tab2_point_text[0].get().strip())
    Point2 = (tab2_shop_list[1].get().strip(),
              tab2_point_text[1].get().strip())
    Point3 = (tab2_shop_list[2].get().strip(),
              tab2_point_text[2].get().strip())
    Credit = credit_text.get()
    Shipping = shipping_method_text.get()
    Memo = tab2_memo_text.get()
    buyinfo_list = [(OrderNumber, '受注番号'), (Date, '購入日'), (Retailer, '購入店舗'),
                    (ProductName, '商品名'), (Asin, 'ASIN'), (Num,
                                                           '個数'), (BuyPrice, '購入単価'), (MarketPrice, '購入時相場'),
                    (Expenses, '販売手数料'), (RealCost, '実質仕入値'), (BenefitPlans,
                                                               '利益予定額'), (Credit, 'Credit Card'),
                    (Breakeven_point, '損益分岐点')]
    for i in buyinfo_list:
        if not i[0]:
            messagebox.showinfo('Input Error', '{}を入力してください'.format(i[1]))
            return
    Point_list = [Point1, Point2, Point3]
    for p in Point_list:
        if p[1]:
            if not p[0] or p[0] == 'Select Shop':
                messagebox.showinfo('Input Error', 'ポイント獲得ショップを入力してください')
                return
    db = client.Purchase
    now_date = datetime.date.today()
    collection_name = str(now_date.year) + '/' + \
        str(now_date.month)  # 年月毎にコレクション作成
    collection = db[collection_name]
    # 新しい受注番号の場合はMongoDBに追加(追加チェックボタンがオンの場合は追加する)
    if collection.find_one({'OrderNumber': OrderNumber}) and add_ordernum_text.get() == 'off':
        messagebox.showinfo('Duplicate', '既に登録されている受注番号です')
        return
    else:
        collection.insert_one({
            'OrderNumber': OrderNumber,
            'Date': Date,
            'Retailer': Retailer,
            'ProductName': ProductName,
            'Asin': Asin,
            'Num': Num,
            'BuyPrice': BuyPrice,
            'MarketPrice': MarketPrice,
            'Expenses': Expenses,
            'RealCost': RealCost,
            'BenefitPlans': BenefitPlans,
            'Breakeven_point': Breakeven_point,
            'Credit': Credit,
            'Shipping': Shipping,
            'Point1': Point1,
            'Point2': Point2,
            'Point3': Point3,
            'Memo': Memo,
        })
        collection.create_index('OrderNumber')  # インデックスを張る
        messagebox.showinfo('Success', '登録しました')


def fetch_amazon_info_first():
    "Fetchボタンを押したとき、スレッドを分けることでバックグラウンド処理"
    asin = tab2_main_text['asin'].get().strip()
    if not asin:
        messagebox.showinfo('Input Error', 'ASINを入力してください')
        return
    elif not re.search(r'[A-Z\d]{10}', asin):
        messagebox.showinfo('Input Error', '正しいASIN（大文字英数字10桁）を入力してください')
        return
    # もしすでに1度抽出しており、販売予定額と損益分岐点に値が入っていたら、単純に送料、販売手数料と損益分岐点の計算だけする
    if tab2_main_text['market_price'].get() and tab2_main_text['breakeven_point'].get():
        breakeven_point = int(tab2_main_text['market_price'].get(
        ).replace(',', '')) - int(tab2_main_text['postage'].get().replace(',', '')) - int(tab2_main_text['expenses'].get().replace(',', ''))
        tab2_main_text['breakeven_point'].set('{:,}'.format(breakeven_point))
        return

    print('Start fetch 「{0}」 from amazon.'.format(asin))
    excuter = ThreadPoolExecutor(max_workers=1)
    excuter.submit(fetch_amazon_info, asin)


def fetch_amazon_info(asin):
    "Amazonの商品ページから価格と商品名を抽出"
    max_retries = 3  # 最大3回リトライする
    retries = 0  # 現在のリトライ回数を示す変数
    while True:
        start_time = time.time()
        m_label_2_2.set('商品名、カート価格を抽出中...')
        USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0'
        headers = {'User-Agent': USER_AGENT}
        url = 'https://www.amazon.co.jp/dp/{}/'.format(asin)
        html_data = requests.get(url, headers=headers)
        soup = BeautifulSoup(html_data.content, 'lxml')
        time.sleep(1)
        try:
            market_price = soup.find('span', id='priceblock_ourprice').text
            # 商品価格が取得できた場合は、FBAシミュレーターに処理を渡す
            print('Amazonカート価格：{}円'.format(market_price))
            market_price = re.search(
                r'[\d,]+', market_price).group()  # エンマークを取り除く
            product_name = soup.find('span', id='productTitle').text.strip()
            tab2_main_text['market_price'].set(market_price)
            tab2_main_text['product_name'].set(product_name)
            fetch_from_FBASimulater(asin, start_time)
            break
        except:
            # 商品価格が取得できない場合、指数関数的なwaitを取りリトライ。
            retries += 1
            if retries >= max_retries:
                # リトライ回数の上限を超えた場合は例外を発生させる
                if soup.title.text == 'Amazon CAPTCHA':
                    print('ロボット対策のページが開いており、情報を取得できません。\nChromeのキャッシュを削除してください')
                    break
                else:
                    print('Too many retries.')
                    print('{0}'.format(soup.title.text))
                    break
            print('商品価格が取得できなかったので、リトライします。（現在のリトライ回数：{0}回'.format(retries))
            wait = 2**(retries - 1)  # 2の0乗は1
            print('Waiting {0}seconds...'.format(wait))
            time.sleep(wait)  # ウェイトを取る


def fetch_from_FBASimulater(asin, start_time):
    "FBAシミュレーターから手数料を抽出"
    market_price = int(tab2_main_text['market_price'].get().replace(',', ''))
    postage = int(tab2_main_text['postage'].get().replace(',', ''))
    if not market_price:
        messagebox.showinfo('Amazonから販売価格を取得できませんでした。一度ツールを再起動し、再度試してみてください。')
        m_label_2_2.set('')
        return

    m_label_2_2.set('販売手数料を抽出中...')
    options = Options()
    # options.binary_location = '/usr/bin/google-chrome'    #Ubuntu用
    options.add_argument('--headless')
    options.add_argument('--window-size=1280,1024')
    # カレントディレクトリ内のchromedriver.exeを参照
    driver = webdriver.Chrome(
        executable_path=os.path.join(os.getcwd(), 'chromedriver.exe'), chrome_options=options)
    # FBAシミュレーターURL
    driver.get(
        'https://sellercentral.amazon.co.jp/hz/fba/profitabilitycalculator/index?lang=ja_JP')
    print('Get fba simulator page.')
    input_element = driver.find_element_by_id('search-string')
    input_element.send_keys(asin)
    input_element.send_keys(Keys.RETURN)
    print('Send asin and enter.')
    # たまにSASIN入力後の画面に、商品を選択する画面が出てくるときがある。
    # tkinterのメッセージがYes or Noの2つだけなので2つから選択。。
    try:
        select_page = driver.find_element_by_css_selector('h4')
        if '商品を' in select_page.text:
            time.sleep(1)
            select_list = driver.find_elements_by_css_selector('li.product p')
            answer = messagebox.askyesno(
                'Product select', '商品が複数見つかりました。以下から選択してください。\n\nはい： {0}\n\nいいえ： {1}'.format(select_list[0].text, select_list[1].text))
            if answer:
                button = driver.find_element_by_css_selector(
                    'button[value="0"]')
            else:
                button = driver.find_element_by_css_selector(
                    'button[value="1"]')
            button.click()
    except:
        pass

    wait = WebDriverWait(driver, 10)  # 10秒でタイムアウトするWebDriverWaitオブジェクトを作成
    print('Loading search product...')
    wait.until(EC.visibility_of_all_elements_located((
        By.CSS_SELECTOR, '#afn-pricing')))
    input_element = driver.find_element_by_id('afn-pricing')
    input_element.send_keys('{0}'.format(market_price))
    input_element = driver.find_element_by_id('afn-fees-inbound-delivery')
    input_element.send_keys('{0}'.format(postage))
    input_element = driver.find_element_by_id('afn-cost-of-goods')
    input_element.send_keys('1000')  # 商品原価は手数料に提供しないので、適当に1000と入力
    # aria-hidden属性があるdiv要素が消えるまで待つ（ロード中のオーバーレイ状態の時のみ出現する属性）
    wait.until(EC.invisibility_of_element_located((
        By.CSS_SELECTOR, 'div[aria-hidden]')))
    print('Waiting load...')
    time.sleep(1)
    button = driver.find_element_by_id('update-fees-link-announce')
    button.click()
    print('Sucess calculate button clicked.')
    # 読み込み中が消えるまで待つ（disabled属性）
    wait.until(EC.invisibility_of_element_located((
        By.CSS_SELECTOR, 'button[disabled]')))
    shipment_fees = driver.find_element_by_id('afn-selling-fees').text
    fulfillment_fees = driver.find_element_by_id(
        'afn-amazon-fulfillment-fees').text
    print('出荷手数料：{0},FBA手数料：{1}'.format(
        shipment_fees, fulfillment_fees))
    # driver.save_screenshot('screenshot.png')    #スクリーンショット
    driver.quit()
    # 総販売手数料　＝　販売手数料に消費税を掛けたもの＋FBA出荷手数料
    expenses = round(int(shipment_fees) * 1.08) + int(fulfillment_fees)
    tab2_main_text['expenses'].set('{:,}'.format(expenses))
    # 損益分岐点　＝　販売予定額　－　総販売手数料
    postage = int(tab2_main_text['postage'].get())
    breakeven_point = market_price - (postage + expenses)
    tab2_main_text['breakeven_point'].set('{:,}'.format(breakeven_point))
    elapased_time = time.time() - start_time
    print('Sucsessed. Passed time is {}ms'.format(round(elapased_time, 2)))
    m_label_2_2.set('完了    Time: {}ms'.format(round(elapased_time, 2)))
    time.sleep(3)
    m_label_2_2.set('')


def cal_nesesary_info():
    cal = {
        'num': (tab2_main_text['num'].get(), '個数'),
        'buy_price': (tab2_main_text['buy_price'].get().strip(), '購入単価'),
        'market_price': (tab2_main_text['market_price'].get().strip(), '販売予定額'),
        'expenses': (tab2_main_text['expenses'].get().strip(), '販売手数料')
    }
    return cal


def total_amount_cal():
    "売上総額、利益予定額、ポイントを自動計算"
    cal = cal_nesesary_info()
    for name, text in cal.items():
        if not text[0]:
            messagebox.showinfo(
                'Input Error', '{}を入力してください'.format(text[1]))
            return
        if not name == 'num':
            for i in re.findall(r'\D+', text[0]):
                if not i == ',':
                    messagebox.showinfo(
                        'Input Error', '{}に正しい数値を入力してください'.format(text[1]))
                    return
            cal[name] = int(text[0].replace(',', ''))
        else:
            cal[name] = text[0]

    # 還元率に値が入っていたら獲得ポイントも計算
    Point = []
    reduction = [tab2_reduction_rate[i].get() for i in range(3)]
    for i, reduction_one in enumerate(reduction):
        point = tab2_point_text[i].get().strip()
        if reduction_one:
            point = round(cal['buy_price']) * cal['num'] * reduction_one / 100
            tab2_point_text[i].set('{:,}'.format(round(point)))
            Point.append(point)
        elif point:
            point = int(point.replace(',', ''))
            tab2_point_text[i].set('{:,}'.format(point))
            Point.append(point)
    realcost = round(cal['num'] * cal['buy_price'] - sum(Point))
    amount = round(cal['num'] * cal['market_price'] + sum(Point))
    tab2_main_text['real_cost'].set('{:,}'.format(realcost))
    tab2_main_text['benefit_plans'].set('{:,}'.format(
        amount - round((cal['buy_price'] + cal['expenses']) * cal['num'])))
    for name, text in cal.items():
        tab2_main_text[name].set('{:,}'.format(text))    # 個数には￥をつけない


def retailer_resister():
    "新しい購入店舗をMongoDBに登録"
    db = client.Purchase
    collection = db.ShopList
    shop_name = retailer_list_text.get()
    if collection.find_one({'ShopName': shop_name}):
        messagebox.showinfo('Duplicate', '既に登録されている購入店舗です')
        return
    else:
        collection.insert_one({'ShopName': shop_name})
        messagebox.showinfo('Success', '登録しました')
        retailer_list.append(shop_name)
        retailer_list_om['values'] = retailer_list


def retailer_delete():
    "既にMongoDBに登録されている購入店舗を削除"
    db = client.Purchase
    collection = db.ShopList
    shop_name = retailer_list_text.get()
    if not collection.find_one({'ShopName': shop_name}):
        messagebox.showinfo('Duplicate', '登録されていない購入店舗です')
        return
    else:
        collection.remove({'ShopName': shop_name})
        messagebox.showinfo('Success', '削除しました')
        retailer_list.remove(shop_name)
        retailer_list_om['values'] = retailer_list


def credit_resister():
    "新しいクレジットカードをMongoDBに登録"
    db = client.Purchase
    collection = db.CreditCard
    credit_name = credit_text.get()
    if collection.find_one({'CreditCard': credit_name}):
        messagebox.showinfo('Duplicate', '既に登録されている購入店舗です')
        return
    else:
        collection.insert_one({'CreditCard': credit_name})
        messagebox.showinfo('Success', '登録しました')
        credit_list.append(credit_name)
        credit_cb['values'] = credit_list


def credit_delete():
    "既にMongoDBに登録されているクレジットカード削除"
    db = client.Purchase
    collection = db.CreditCard
    credit_name = credit_text.get()
    if not collection.find_one({'CreditCard': credit_name}):
        messagebox.showinfo('Duplicate', '登録されていない購入店舗です')
        return
    else:
        collection.remove({'CreditCard': credit_name})
        messagebox.showinfo('Success', '削除しました')
        credit_list.remove(credit_name)
        credit_cb['values'] = credit_list


def tab2_delete_all():
    "全ての項目を削除"
    tab2_order_number_text.set('')
    retailer_list_text.set('')
    tab2_main_text['product_name'].set('')
    tab2_main_text['asin'].set('')
    tab2_main_text['num'].set(1)
    tab2_main_text['buy_price'].set('')
    tab2_main_text['market_price'].set('')
    tab2_main_text['expenses'].set('')
    tab2_main_text['real_cost'].set('')
    tab2_main_text['benefit_plans'].set('')
    tab2_main_text['breakeven_point'].set('')
    for i in range(3):
        if i == 0:
            tab2_shop_list[i].set('Select Shop')
        else:
            tab2_shop_list[i].set('')
        tab2_point_text[i].set('')
        tab2_reduction_rate[i].set(0)
    credit_text.set('Credit Card')
    add_ordernum_text.set('off')
    shipping_method_text.set('自己発送')


def tab2_delete_fromDB():
    "MongoDBから現在の受注番号のデータを削除"
    db = client.Purchase
    now_date = datetime.date.today()
    collection_name = str(now_date.year) + '/' + \
        str(now_date.month)  # 年月毎にコレクション作成
    collection = db[collection_name]
    order_number = tab2_order_number_text.get().strip()
    if collection.find_one({'OrderNumber': order_number}):
        ask = messagebox.askokcancel(
            'Confirm', '受注番号：{0} をデータベースから削除しますか？'.format(order_number))
        if ask:
            collection.remove(
                {'OrderNumber': order_number})
            messagebox.showinfo('Success', '削除しました')
    else:
        messagebox.showinfo('DBError', '登録されていない受注番号です')


# TAB3 関数 ----------------------------------------------------------------


class StockDisplay(Frame):
    "保存した有在庫を表示するクラス"

    def __init__(self, master, **kwargs):
        "初期化メソッド"
        Frame.__init__(self, master, **kwargs)
        self.master = master
        self.display_stock()

    def display_stock(self):
        # ツリービューの作成
        tree = Treeview(self.master)
        # 列インデックスの作成
        tree["columns"] = (1, 2, 3)
        # 表スタイルの設定(headingsはツリー形式ではない、通常の表形式)
        tree["show"] = "headings"
        # 各列の設定(インデックス,オプション(今回は幅を指定))
        tree.column(1, width=75)
        tree.column(2, width=75)
        tree.column(3, width=100)
        # 各列のヘッダー設定(インデックス,テキスト)
        tree.heading(1, text="日付")
        tree.heading(2, text="内訳")
        tree.heading(3, text="金額")

        # レコードの作成
        # 1番目の引数-配置場所（ツリー形式にしない表設定ではブランクとする）
        # 2番目の引数-end:表の配置順序を最下部に配置
        #             (行インデックス番号を指定することもできる)
        # 3番目の引数-values:レコードの値をタプルで指定する
        tree.insert("", "end", values=("2017/5/1", "食費", 3500))
        tree.insert("", "end", values=("2017/5/10", "光熱費", 7800))
        tree.insert("", "end", values=("2017/5/10", "住宅費", 64000))

        # ツリービューの配置
        tree.pack()


def create_display_window():
    "カレンダーウィンドウ作成"
    root_display = Tk()
    root_display.title('Calender')
    display_stock = StockDisplay(root_display)
    display_stock.pack()
    root_display.mainloop()

# TAB4 関数 ----------------------------------------------------------------


" メイン処理--------------------------------------------------------------------------"


if __name__ == '__main__':
    # MongoClient接続
    client = MongoClient('localhost', 27017)
    # rootの作成
    root = Tk()
    root.title('Buyer Tool')
    root.resizable(0, 0)

    # ノートブックの作成
    nb = Notebook()

    # タブを作成
    tab1 = Frame(nb)
    tab2 = Frame(nb)
    tab3 = Frame(nb)
    tab4 = Frame(nb)
    tab1.grid(row=0, column=0, columnspan=6)
    tab2.grid(row=0, column=0, columnspan=6)
    tab3.grid(row=0, column=0, columnspan=6)
    tab4.grid(row=0, column=0, columnspan=6)

    nb.add(tab1, text='　ASIN　', padding=3)
    nb.add(tab2, text=' 有在庫 　', padding=3)
    nb.add(tab3, text='　分析　', padding=3)
    nb.add(tab4, text='　トイザらス　', padding=3)
    nb.pack(fill='both')

    "TAB1 ASIN --------------------------------------------------------------"

    # frame1_1の作成
    frame1_1 = Frame(tab1, padding=(2, 8))
    frame1_1.grid(row=1, column=0, columnspan=6)

    # 「URL」ラベルの作成
    tab1_url_label = Label(frame1_1, text='URL>>', width=6)
    tab1_url_label.grid(row=0, column=0, padx=5, sticky=E)

    # URL入力ボックスの作成
    tab1_url_text = StringVar()
    tab1_url_entry = Entry(
        frame1_1, textvariable=tab1_url_text, width=52)
    tab1_url_entry.grid(row=0, column=1, columnspan=4, sticky=EW, padx=0)

    # Amazon開くボタンの作成
    tab1_Amazon_button = Button(
        frame1_1, text='Amazon', command=goto_amazon, takefocus=False)
    tab1_Amazon_button.grid(row=0, column=5, padx=5)

    # Frame_categoryの作成
    Frame_category = Frame(frame1_1)
    Frame_category.grid(row=1, column=0, columnspan=2)

    # カテゴリーの幅が変わらないための空ラベル
    label_ex = Label(Frame_category, text='', width=22)
    label_ex.grid(row=0, column=0, padx=5)

    # カテゴリー選択ボタンの作成
    optionList = ['カテゴリー', 'すべてのカテゴリー', 'TVゲーム', 'パソコン・周辺機器', 'おもちゃ', 'ホビー', '家電&カメラ', '楽器', 'スポーツ&アウトドア', '車&バイク', 'DIY・工具・ガーデン', '文房具・オフィス用品',
                  'ホーム&キッチン', 'ペット用品', 'ドラッグストア', 'ビューティー', 'ラグジュアリービューティー', 'ベビー&マタニティ', 'ファッション', '服&ファッション小物', 'シューズ&バッグ', '腕時計', 'ジュエリー']
    category_text = StringVar()
    category_cb = OptionMenu(
        Frame_category, category_text, *optionList)
    category_cb.grid(row=0, column=0,  columnspan=2,
                     padx=5, pady=2, sticky="ew")

    # キーワードラベルの作成
    label1 = Label(frame1_1, text='Keyword:')
    label1.grid(row=1, column=2, padx=0, pady=2)

    # キーワード入力欄の作成
    keyword = StringVar()
    keyword_entry = Entry(
        frame1_1, textvariable=keyword, width=24)
    keyword_entry.grid(row=1, column=3, columnspan=2, padx=2, pady=2)

    # Deleteボタンの作成
    delete_button = Button(
        frame1_1, text='Delete', command=tab1_deleteBtn_clicked, takefocus=False)
    delete_button.grid(row=1, column=5, padx=0, pady=2)

    # Frame_fetch_optionsの作成
    Frame_fetch_option = Frame(frame1_1)
    Frame_fetch_option.grid(row=2, column=0, columnspan=7, sticky=E)

    # ASIN保存チェックラベル
    asin_save_check_label = Label(
        Frame_fetch_option, text='分割＆保存', anchor=E, width=10)
    asin_save_check_label.grid(row=0, column=0, padx=0, pady=3, sticky=E)

    # ASIN保存チェック
    asin_save_check = StringVar()
    asin_save_check.set('on')
    asin_save_check_box = Checkbutton(
        Frame_fetch_option, variable=asin_save_check, onvalue='on', offvalue='off', takefocus=False, width=2)
    asin_save_check_box.grid(row=0, column=1, padx=3, pady=3, sticky=E)

    # スポンサープロダクトチェックラベル
    sponsor_product_label = Label(
        Frame_fetch_option, text='SP', anchor=E, width=4)
    sponsor_product_label.grid(row=0, column=2, padx=0, pady=3, sticky=E)

    # スポンサープロダクトチェック
    sponsor_product_check = StringVar()
    sponsor_product_check.set('on')
    sponsor_product_check_box = Checkbutton(
        Frame_fetch_option, variable=sponsor_product_check, onvalue='on', offvalue='off', takefocus=False, width=2)
    sponsor_product_check_box.grid(row=0, column=3, padx=3, pady=3, sticky=E)

    # 重複チェックラベル
    asin_duplicate_check_label = Label(
        Frame_fetch_option, text='重複無視', anchor=E, width=8)
    asin_duplicate_check_label.grid(row=0, column=4, padx=0, pady=3, sticky=E)

    # 重複チェック
    asin_duplicate_check = StringVar()
    asin_duplicate_check.set('off')
    asin_duplicate_check_box = Checkbutton(
        Frame_fetch_option, variable=asin_duplicate_check, onvalue='on', offvalue='off', takefocus=False, width=2)
    asin_duplicate_check_box.grid(row=0, column=5, padx=3, pady=3, sticky=E)

    # カテゴリーの幅が変わらないための空ラベル
    price_updown_label_dammy = Label(Frame_fetch_option, text='', width=12)
    price_updown_label_dammy.grid(row=0, column=6, padx=3, sticky=E)

    # 金額上限
    price_updown_list = ('Price', r'\500↑', r'\500↓', 'ALL')
    price_updown = StringVar()
    price_updown_cb = OptionMenu(
        Frame_fetch_option, price_updown, *price_updown_list)
    price_updown_cb.grid(row=0, column=6, padx=3, sticky=E)

    # 抽出ボタンの作成
    tab1_fetch_button = Button(
        Frame_fetch_option, text='Fetch', command=asinfetchBtn_clicked_first, takefocus=False)
    tab1_fetch_button.grid(row=0, column=7, padx=5, sticky=E)

    # frame1_2の作成
    frame1_2 = Frame(tab1, relief="ridge", padding=8)
    frame1_2.grid(row=2, column=0,  columnspan=3, pady=1)

    # splitラベルの作成
    label2 = Label(frame1_2, text='【ASIN分割ツール】 ')
    label2.grid(row=0, column=0, padx=0, pady=2)

    # Text
    txt1 = Text(frame1_2, width=12, height=8)
    txt1.grid(row=1, rowspan=4, column=0, padx=5)

    # Scrollbar
    scrollbar1 = Scrollbar(
        frame1_2,
        orient=VERTICAL,
        command=txt1.yview)
    txt1['yscrollcommand'] = scrollbar1.set
    scrollbar1.grid(row=1, rowspan=4, column=1, sticky=N + S)

    # Countラベルの作成
    label2 = Label(frame1_2, text='分割数 ')
    label2.grid(row=0, column=2, padx=0, pady=2)

    # Count
    optionList = (400, 100, 200, 300, 400,
                  500, 600, 700, 800, 900, 1000)
    max_num = IntVar()
    max_num.set(400)
    max_num_cb = OptionMenu(
        frame1_2, max_num, *optionList)
    max_num_cb.grid(row=1, column=2, padx=3, sticky="ew")

    # Copyボタンの作成
    asin_copy_button = Button(frame1_2, text='Copy',
                              command=asin_copyBtn_clicked)
    asin_copy_button.grid(row=2, column=2, pady=0, padx=3)

    # Deleteボタンの作成
    Delete_button = Button(frame1_2, text='Delete',
                           command=deleteBtn_clicked)
    Delete_button.grid(row=3, column=2, padx=3)

    # Splitボタンの作成
    split_button = Button(frame1_2, text='Split',
                          command=splitBtn_clicked)
    split_button.grid(row=4, column=2, padx=3)

    # frame1_3の作成
    frame1_3 = Frame(tab1, relief="ridge", padding=8)
    frame1_3.grid(row=2, column=3,  columnspan=3, pady=1)

    # 禁止ワードラベルの作成
    banword_label = Label(frame1_3, text='【禁止ワード 】')
    banword_label.grid(row=0, column=0, padx=0, pady=2)

    # 禁止ワードボックス
    banword = Text(frame1_3, width=12, height=8)
    banword.grid(row=1, rowspan=4, column=0, padx=5)

    # 禁止ワードスクロールバー
    banword_scrollbar = Scrollbar(
        frame1_3,
        orient=VERTICAL,
        command=banword.yview)
    banword['yscrollcommand'] = banword_scrollbar.set
    banword_scrollbar.grid(row=1, rowspan=4, column=1, sticky=N + S)

    # ダミーラベルの作成
    blank_txt_tab1_t = StringVar()
    blank_txt_tab1 = Label(frame1_3, textvariable=blank_txt_tab1_t, width=14)
    blank_txt_tab1.grid(row=0, column=2, padx=0, sticky=W)

    # ダミーラベルの作成
    blank_txt_tab1_t = StringVar()
    blank_txt_tab1 = Label(frame1_3, textvariable=blank_txt_tab1_t, width=14)
    blank_txt_tab1.grid(row=1, column=2, padx=0, sticky=W)

    # ダミーラベルの作成
    blank_txt_tab1_t = StringVar()
    blank_txt_tab1 = Label(frame1_3, textvariable=blank_txt_tab1_t, width=14)
    blank_txt_tab1.grid(row=2, column=2, pady=3, padx=0, sticky=W)

    # banword_Deleteボタンの作成
    banword_delete = Button(frame1_3, text='Delete',
                            command=delete_banword_Btn_clicked)
    banword_delete.grid(row=3, column=2, padx=3)

    # banword_Saveボタンの作成
    banword_save = Button(frame1_3, text='Save',
                          command=dammy)
    banword_save.grid(row=4, column=2, pady=0, padx=3)

    # frame_referenceの作成
    frame_reference = Frame(tab1, padding=(2, 8))
    frame_reference.grid(row=3, column=0, columnspan=6,  pady=0)

    # ラベルの作成
    # 「ファイル」ラベルの作成
    label_file = Label(frame_reference, text='ファイル>>')
    label_file.grid(row=0, column=0, padx=5)

    # 参照ファイルパス表示ラベルの作成
    file1 = StringVar()
    file1_entry = Entry(
        frame_reference, textvariable=file1, width=35)
    file1_entry.grid(row=0, column=1, columnspan=3)

    # 参照ボタンの作成
    button1 = Button(
        frame_reference, text='Reference', command=referenceBtn_clicked)
    button1.grid(row=0, column=4, padx=5)

    # Saveボタンの作成
    save_asin_toMongo = Button(frame_reference, text='Save',
                               command=saveBtn_clicked_first)
    save_asin_toMongo.grid(row=0, column=5, padx=0, pady=2)

    # frame1_4の作成
    frame1_4 = Frame(tab1, padding=(2, 0))
    frame1_4.grid(row=4, column=0, columnspan=6,  pady=0)

    m_label_1 = Label(frame1_4,
                      text='  Massage------------>')
    m_label_1.grid(row=0, column=0, padx=5, pady=0, sticky=W)

    # 進捗ラベルの作成
    m_label_1_1 = StringVar()
    m_label_1_1_label = Label(frame1_4, textvariable=m_label_1_1, width=22)
    m_label_1_1_label.grid(row=0, column=1, padx=5, sticky=W)

    # ダミーラベルの作成
    blank_txt_tab1_t = StringVar()
    blank_txt_tab1 = Label(frame1_4, textvariable=blank_txt_tab1_t, width=14)
    blank_txt_tab1.grid(row=0, column=4, padx=0, sticky=W)

    # Exit接続ボタンの作成
    tab1_exit_btn = Button(frame1_4, text='Exit',
                           command=tab1_exit_btn_clicked)
    tab1_exit_btn.grid(row=0, column=5, padx=0, pady=2, sticky=E)

    " TAB2 有在庫 ---------------------------------------------------------------"

    # frame2_1の作成
    frame2_1 = Frame(tab2, padding=(2, 8))
    frame2_1.grid(row=0, column=0, columnspan=7)
    num_length = 12

    # frame_orderの作成
    frame_order = Frame(frame2_1, padding=(0, 0))
    frame_order.grid(row=0, column=0, padx=3, pady=0, sticky=W, columnspan=7)

    # 購入日ラベル
    buyday_label = Label(frame_order, text='  　 購入日：', anchor=W)
    buyday_label.grid(row=0, column=0, padx=3, pady=3, sticky=W)

    # 購入日入力欄
    buyday_text = StringVar()

    # 今日の日付をセットしておく
    buyday_text.set(str(datetime.date.today()).replace('-', '/'))
    buyday_entry = Entry(
        frame_order, textvariable=buyday_text, width=num_length - 1, takefocus=False, justify=CENTER)
    buyday_entry.grid(row=0, column=1, pady=3, padx=0,  sticky=W)

    # カレンダーボタンの作成
    calender_button = Button(frame_order, text='<',
                             command=lambda: create_calender_window('buyday'), width=2, takefocus=False)
    calender_button.grid(row=0, column=3, padx=3, sticky=W)

    # 受注番号ラベル
    tab2_order_number = Label(
        frame_order, text='　  受注番号：', anchor=W, takefocus=False)
    tab2_order_number.grid(row=0, column=4, padx=0, pady=3, sticky=W)

    # 受注番号入力欄
    tab2_order_number_text = StringVar()
    tab2_order_number_entry = Entry(
        frame_order, textvariable=tab2_order_number_text, width=16)
    tab2_order_number_entry.grid(row=0, column=5, padx=0, pady=3, sticky=W)

    #  追加ラベル
    add_ordernum_label = Label(frame_order, text='追加', anchor=W, width=4)
    add_ordernum_label.grid(row=0, column=6, padx=1, pady=3, sticky=E)

    # 追加チェックボックス
    add_ordernum_text = StringVar()
    add_ordernum_text.set('off')
    add_ordernum = Checkbutton(
        frame_order, variable=add_ordernum_text, onvalue='on', offvalue='off', takefocus=False, width=4)
    add_ordernum.grid(row=0, column=7, padx=1, pady=3, sticky=E)

    # frame_retailerの作成
    frame_retailer = Frame(frame2_1, padding=(0, 0))
    frame_retailer.grid(row=1, column=0, padx=3,
                        pady=0, sticky=W, columnspan=6)
    # 購入店舗名ラベル
    retailer_label = Label(frame_retailer, text='   購入店舗：', anchor=W)
    retailer_label.grid(row=0, column=0, padx=3, pady=3, sticky=W)

    # MongoDBから購入店舗リスト取得
    retailer_list = first_connect_monodb()

    # 購入店舗名Combobox
    retailer_list_text = StringVar()
    retailer_list_om = Combobox(
        frame_retailer, textvariable=retailer_list_text, values=retailer_list, width=40, takefocus=False)
    retailer_list_om.grid(row=0, column=1, padx=0, pady=3, sticky="ew")

    # 新しい購入店舗をMongoDBに登録するボタン
    retailer_resister_btn = Button(
        frame_retailer, text='登録', command=retailer_resister, width=6, takefocus=False)
    retailer_resister_btn.grid(row=0, column=2, padx=4, sticky=EW)

    # 既にMongoDBに登録されている購入店舗を削除するボタン
    retailer_delete_btn = Button(
        frame_retailer, text='削除', command=retailer_delete, width=6, takefocus=False)
    retailer_delete_btn.grid(row=0, column=3, padx=0, sticky=EW)

    # 商品名以下
    tab2_label_list = [
        ('product_name', '商品名：', 60),
        ('asin', 'ASIN：', num_length),
        ('market_price', '販売予定額：', num_length),
        ('postage', 'FBA納品送料：', num_length),
        ('expenses', '販売手数料：', num_length),
        ('breakeven_point', '損益分岐点：', num_length),
        ('num', '個数：', 3),
        ('buy_price', '購入単価：', num_length),
        ('real_cost', '実質仕入値：', num_length),
        ('benefit_plans', '利益予定額：', num_length)]
    tab2_label = {}
    tab2_main_text = {}
    tab2_main_entry = {}
    i = 0
    for key, value, length in tab2_label_list:
        # ラベルの作成
        tab2_label[key] = Label(
            frame2_1, text=value, anchor=E, width=num_length)
        tab2_label[key].grid(row=i + 2, column=0, padx=3, pady=3)
        # 個数欄のみスピンボックス
        if key == 'num':
            tab2_main_text['num'] = IntVar()
            tab2_main_text['num'].set(1)
            num_spinbox = Spinbox(
                frame2_1, increment=1.0, from_=0, to=10000, width=5, justify=RIGHT, textvariable=tab2_main_text['num'])
            num_spinbox.grid(
                row=i + 2, column=1, columnspan=3,  pady=3, sticky=W)
            i += 1
            continue
        if key == 'product_name':
            direct = LEFT
        else:
            direct = RIGHT
        tab2_main_text[key] = StringVar()
        tab2_main_entry[key] = Entry(
            frame2_1, textvariable=tab2_main_text[key], justify=direct, width=length, takefocus=False)
        tab2_main_entry[key].grid(
            row=i + 2, column=1, columnspan=3,  pady=3, sticky=W)
        i += 1
    # 送料固定
    tab2_main_text['postage'].set('0')
    # frame2_2の作成
    frame2_2 = Frame(frame2_1, padding=(0, 0))
    frame2_2.grid(row=3, column=3, rowspan=4, pady=0, sticky=E)

    # ポイントラベルの作成
    tab2_shop_text = StringVar()
    tab2_shop_text.set('【Point】　ショップ名　 　還元率　  獲得ポイント')
    tab2_shop_label = Label(
        frame2_2, textvariable=tab2_shop_text, justify=RIGHT)
    tab2_shop_label.grid(row=0, column=0, columnspan=4,
                         padx=5, pady=2, sticky=W)

    # ポイントリスト
    shopList = ['Select Shop', 'LINEショッピング',
                'Yahoo!ショッピング', '楽天市場', 'ビックカメラ', 'ドン・キホーテ']
    tab2_label_ex = []
    tab2_shop_list = []
    tab2_shop_list_om = []
    tab2_point_text = []
    tab2_point_entry = []
    tab2_p_label = []
    tab2_reduction_rate = []
    tab2_reduction_rate_spin = []
    for i in range(3):
        if not i == 0:
            shopList.pop(0)
            shopList.insert(0, '　　　　　')
        # カテゴリーの幅が変わらないための空ラベル
        tab2_label_ex.append(Label(frame2_2, text='', width=20))
        tab2_label_ex[i].grid(row=i + 1, column=0, padx=5)
        # ショップ名
        tab2_shop_list.append(StringVar())
        tab2_shop_list[i].set('')
        tab2_shop_list_om.append(OptionMenu(
            frame2_2, tab2_shop_list[i], *shopList))
        tab2_shop_list_om[i].grid(row=i + 1, column=0, padx=3, sticky=E)
        # 還元率
        tab2_reduction_rate.append(DoubleVar())
        tab2_reduction_rate_spin.append(Spinbox(
            frame2_2, increment=1, from_=0, to=100, width=4, justify=RIGHT, textvariable=tab2_reduction_rate[i]))
        tab2_reduction_rate_spin[i].grid(
            row=i + 1, column=1, padx=0, pady=3, sticky=E)
        # ポイント
        tab2_point_text.append(StringVar())
        tab2_point_entry.append(Entry(
            frame2_2, textvariable=tab2_point_text[i], justify=RIGHT, width=num_length, takefocus=False))
        tab2_point_entry[i].grid(
            row=i + 1, column=2, padx=3, pady=3, sticky=E)
        # Pラベル
        tab2_p_label.append(Label(frame2_2, text='P', anchor=E, width=1))
        tab2_p_label[i].grid(row=i + 1, column=3, padx=0, pady=3, sticky=E)

    # frame2_3の作成
    frame2_3 = Frame(frame2_1, padding=(0, 0))
    frame2_3.grid(row=7, column=0, columnspan=6, pady=0, sticky=E)

    # クレカラベル
    tab2_credit_label = Label(frame2_3, text='クレカ：', anchor=E)
    tab2_credit_label.grid(row=0, column=0, padx=0, pady=0, sticky=E)

    # MongoDBからクレジットカードリスト取得
    #CreditList = ['Credit Card', 'Amazon MASTER 6473', '三井 MASTER 1091','三井 VISA 4138', 'Orico MASTER 5508', 'EPOS VISA 5785']
    credit_list = get_creditlist()

    credit_text = StringVar()
    credit_cb = Combobox(
        frame2_3, textvariable=credit_text, values=credit_list, width=21, takefocus=False)
    credit_cb.grid(row=0, column=1,  columnspan=2,
                   padx=0, pady=3, sticky="w")

    # クレジットカードをMongoDBに登録するボタン
    credit_resister_btn = Button(
        frame2_3, text='登録', command=credit_resister, width=4, takefocus=False)
    credit_resister_btn.grid(row=0, column=3, padx=3, sticky=EW)

    # 既にMongoDBに登録されているクレジットカードを削除するボタン
    credit_delete_btn = Button(
        frame2_3, text='削除', command=credit_delete, width=4, takefocus=False)
    credit_delete_btn.grid(row=0, column=4, padx=0, sticky=EW)

    # frame2_4の作成
    frame2_4 = Frame(frame2_1, padding=(0, 0))
    frame2_4.grid(row=8, column=0, columnspan=5, pady=0, sticky=E)

    # メモラベル
    tab2_memo_label = Label(frame2_4, text='メモ：', anchor=E)
    tab2_memo_label.grid(row=0, column=0, padx=0, pady=0, sticky=E)

    # メモ入力欄
    tab2_memo_text = StringVar()
    tab2_memo_entry = Entry(
        frame2_4, textvariable=tab2_memo_text, width=36)
    tab2_memo_entry.grid(row=0, column=1, columnspan=2, pady=0)

    # frame2_5の作成
    frame2_5 = Frame(frame2_1, padding=(0, 0))
    frame2_5.grid(row=9, column=3, columnspan=5, rowspan=2, pady=0, sticky=E)

    # 下部ボタン
    tab2_under_button_list = [
        ('Fetch', fetch_amazon_info_first),
        ('Calc', total_amount_cal),
        ('DelTxt', tab2_delete_all),
        ('DelDB', tab2_delete_fromDB),
        ('Save', tab2_saveBtn_clicked_first),
        ('Exit', exitBtn_clicked),
    ]
    tab2_under_button = []
    row_n = 0
    column_n = 0
    for i, fn in enumerate(tab2_under_button_list):
        # Caluculateボタンの作成
        if i >= 1:
            row_n = 1
            column_n = i - 1
        tab2_under_button.append(Button(frame2_5, text=fn[0], width=6,
                                        command=fn[1], takefocus=False))
        tab2_under_button[i].grid(
            row=row_n, column=column_n, padx=3, pady=0, sticky=E)

    #  直送か、代理かラベル
    shipping_method_label = Label(frame2_5, text='代理業者に直送', anchor=E)
    shipping_method_label.grid(
        row=0, column=2, columnspan=2, padx=0, pady=3, sticky=E)

    # 直送か、代理か
    shipping_method_text = StringVar()
    shipping_method_text.set('自己発送')
    shipping_method = Checkbutton(
        frame2_5, variable=shipping_method_text, onvalue='直送', offvalue='自己発送', takefocus=False, width=4)
    shipping_method.grid(row=0, column=4, padx=0, pady=3, sticky=W)

    # frame2_6の作成
    frame2_6 = Frame(frame2_1, padding=(2, 0))
    frame2_6.grid(row=11, column=3, columnspan=5,  pady=0)

    m_label_2 = Label(frame2_6,
                      text='  Massage---->')
    m_label_2.grid(row=0, column=0, padx=5, pady=0, sticky=W)

    # 進捗ラベルの作成
    m_label_2_2 = StringVar()
    m_label_2_2_label = Label(frame2_6, textvariable=m_label_2_2, width=25)
    m_label_2_2_label.grid(row=0, column=3, padx=5, sticky=W)

    # frame3_1の作成
    frame3_1 = Frame(tab3, padding=(2, 8))
    frame3_1.grid(row=1, column=0, columnspan=6)

    # 期間ラベル
    period_label = Label(frame3_1, text='期間：', anchor=W)
    period_label.grid(row=0, column=0, padx=3, pady=3, sticky=W)

    # 期間入力欄
    period_text = StringVar()

    # 今月のの１日をセットしておく
    period_str_first = str(datetime.date.today()
                           ).replace('-', '/')[:-2] + '01'
    period_text.set(period_str_first)
    period_entry = Entry(
        frame3_1, textvariable=period_text, width=num_length - 1, takefocus=False, justify=CENTER)
    period_entry.grid(row=0, column=1, pady=3, padx=0,  sticky=W)

    # カレンダーボタンの作成
    calender_button = Button(frame3_1, text='<',
                             command=lambda: create_calender_window('period_first'), width=2, takefocus=False)
    calender_button.grid(row=0, column=2, padx=3, sticky=W)

    # 期間ラベル2
    period_label2 = Label(frame3_1, text='～', anchor=W)
    period_label2.grid(row=0, column=3, padx=3, pady=3, sticky=W)

    # 期間入力欄
    period_text2 = StringVar()

    # 今日の日付をセットしておく
    period_text2.set(str(datetime.date.today()).replace('-', '/'))
    period_entry2 = Entry(
        frame3_1, textvariable=period_text2, width=num_length - 1, takefocus=False, justify=CENTER)
    period_entry2.grid(row=0, column=4, pady=3, padx=0,  sticky=W)

    # カレンダーボタンの作成
    calender_button = Button(frame3_1, text='<',
                             command=lambda: create_calender_window('period_second'), width=2, takefocus=False)
    calender_button.grid(row=0, column=5, padx=3, sticky=W)

    # 表示ボタンの作成
    tab3_test = Button(
        frame3_1, text='表示', command=create_display_window)
    tab3_test.grid(row=0, column=6, padx=5)

    """Tab4"""
    # frame4_1の作成
    frame4_1 = Frame(tab4, padding=(2, 8))
    frame4_1.grid(row=0, column=0, columnspan=6)

    # 「URL」ラベルの作成
    tab4_url_label = Label(frame4_1, text='URL>>', width=6)
    tab4_url_label.grid(row=0, column=0, padx=5, sticky=E)

    # URL入力ボックスの作成
    tab4_url_text = StringVar()
    tab4_url_entry = Entry(
        frame4_1, textvariable=tab1_url_text, width=52)
    tab4_url_entry.grid(row=0, column=1, columnspan=4, sticky=EW, padx=0)

    # Amazon開くボタンの作成
    tab4_Amazon_button = Button(
        frame4_1, text='抽出', command=goto_amazon, takefocus=False)
    tab4_Amazon_button.grid(row=0, column=5, padx=5)

    # frame4_2の作成
    frame4_2 = Frame(tab4, relief="ridge", padding=8)
    frame4_2.grid(row=1, column=0,  columnspan=2, pady=1)

    # splitラベルの作成
    label4 = Label(frame4_2, text='【ASINリスト】 ')
    label4.grid(row=0, column=0, padx=0, pady=2)

    # Text
    txt4 = Text(frame4_2, width=12, height=8)
    txt4.grid(row=1, rowspan=4, column=0, padx=5)

    # Scrollbar
    scrollbar4 = Scrollbar(
        frame4_2,
        orient=VERTICAL,
        command=txt4.yview)
    txt4['yscrollcommand'] = scrollbar4.set
    scrollbar4.grid(row=1, rowspan=4, column=1, sticky=N + S)

    root.mainloop()
