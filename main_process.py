import sys
import time
import re
import csv

# マルチスレッド処理に使用
import threading

# GUIの作り込みに使用
import tkinter
import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import *

# PDF関連のライブラリ
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams
from io import StringIO

# WEB操作に使用するライブラリ
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# モジュールのインポート
from web_driver import WebDriver
from gif_player import TkGif
from config_loader import LoadConfig
from adjust_data_info import adjust_date_info
from reply_manager import reply_mode, download_mode, click_matched_title
from move_manager import reset_move, change_move


class MainProcess:
    """
    メインプロセスを管理するクラス。
    - サイボウズメールの自動処理（返信・ダウンロード）を実行。
    - PDF解析、GUIの操作、WebDriver管理などを行う。
    """
    def __init__(self):
        """
        クラスの初期化処理。
        設定ファイルの読み込み、WebDriverの準備、GUIのセットアップを行う。
        """
        # 設定ファイルの読み込み
        config_loader = LoadConfig()
        self.config = config_loader.config
        self.config.optionxform = str
        self.config.read_string(config_loader.config_strings(config_loader.USER_INFO_CONFIG_ENC_PATH))

        # 使用ユーザー判別
        self.user = config_loader.identify_user()

        # chromeのアップデートを実行
        self.webdriver = WebDriver(self.user)
        self.webdriver.update_webdriver()
        self.chrome_service = self.webdriver.chrome_service
        self.options = self.webdriver.options

        # サイボウズにログインするIDとPASSの辞書をiniファイルから取得
        self.id_dict = config_loader.id_dict
        self.pass_dict = config_loader.pass_dict
        self.last_name_dict = config_loader.last_name_dict

        # アイコンに使用するGIFのPATH
        self.penguin_standby = r"C:\path\to\penguin_standby.gif"
        self.penguin_download = r"C:\path\to\penguin_download.gif"

        # 変数の初期化
        self.flg = None
        self.split_text = None
        self.body_text = None
        self.max_dict = None
        self.cus_name = None
        self.min_id = None
        self.min_pass = None
        self.mail_title_sub_flg = None
        self.download_mail_add = None
        self.mail_index_title = None
        self.detected_mail = 0

        # self.PIC_name = None
        # self.max_ratio = None
        # self.fixed_flg = None
        # self.attached_pass = None
        # self.attached_pass_index = None

        # 辞書
        self.ratio_dict = {}
        self.distance_id_dict = {}
        self.distance_pass_dict = {}
        self.distance_pic_dict = {}
        self.distance_from_dict = {}

        self.file = ""

        # メインウィンドウの生成
        self.root = TkinterDnD.Tk()
        self.root.update()
        self.root.geometry('225x220')
        self.root.attributes("-topmost", True)
        self.root.title("ぺんぎん")
        self.iconfile = r"C:\path\to\penguin.ico"
        self.root.iconbitmap(default=self.iconfile)

        # ドラッグ＆ドロップ対応のラベル
        self.text = tk.StringVar(self.root)
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack()
        self.label = tk.Label(self.main_frame)
        self.label.pack()
        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind("<<Drop>>", self.drop)
        # ウィジェットの配置
        self.label.grid(row=0, column=0, padx=10)

        # GIFアニメーションのセットアップ
        self.gif_player = TkGif(self.penguin_standby, self.label)
        self.gif_player2 = TkGif(self.penguin_download, self.label)
        self.gif_player.play()

        # ラジオボタンの設定（返信モード・ダウンロードモード）
        self.var = tkinter.IntVar()
        self.var.set(0)
        self.rdo1 = tkinter.Radiobutton(self.root, value=0, variable=self.var, text="返信")
        self.rdo2 = tkinter.Radiobutton(self.root, value=1, variable=self.var, text="ダウンロード")
        self.rdo1.pack()
        self.rdo2.pack()

        # スレッドとイベントのセットアップ
        self.thread_flg = False
        self.Event = threading.Event()
        self.thread1 = threading.Thread(target=self.auto_reply)

        # メインループ開始
        self.root.mainloop()

    def drop(self, event):
        """
        PDFファイルをドロップしたときの処理。
        - PDFを解析し、必要なデータを抽出。

        Args:
            event (tk.Event): ドロップイベント。

        Returns:
            None
        """
        # ドロップされたファイルのパスを取得し、解析処理を開始
        self.text.set(event.data)  # ドロップされたファイルパスをGUIにセット
        self.file = event.data.strip("{}")  # ファイルパスの余分な{}を削除
        fp = open(self.file, "rb")  # PDFファイルをバイナリモードで開く
        outfp = StringIO()
        rmgr = PDFResourceManager()
        lprms = LAParams()
        device = TextConverter(rmgr, outfp, laparams=lprms)
        iprtr = PDFPageInterpreter(rmgr, device)

        # PDFファイルの各ページを解析してテキストを抽出
        try:
            for page in PDFPage.get_pages(fp):
                iprtr.process_page(page) # 各ページを処理
        except Exception as e:  # エラーが発生した場合に内容を出力
            print(e)

        # 抽出したテキストを取得し、後処理
        self.body_text = outfp.getvalue()  # テキストデータを取得
        outfp.close()
        device.close()
        fp.close()

        # テキストを分割してリストに格納
        self.split_text = self.body_text.split()

        # スレッドを管理し、自動返信プロセスを開始
        if not self.thread_flg:
            self.thread_flg = True  # スレッド実行中のフラグを設定
            self.root.after(200, # 少し待機してからスレッドを開始
                       self.thread1.start(),
                       )
        else:
            # スレッドを新しく作成して再開始
            self.thread1 = threading.Thread(target=self.auto_reply)
            self.root.after(200,
                       self.thread1.start(),
                       )

    def replaced_contents(self):
        """
        PDFから抽出したテキストデータを整形し、不要な文字や情報を除去する。
        - 漢字、ひらがな、カタカナを除去。
        - 特定の特殊文字（例: <, >, "）を削除。
        - 整形後のテキストをリスト形式で返却。

        Returns:
            dict: 整形されたデータを格納した辞書。
                  - "first_letter_in_replaced_split": 整形後のリストの先頭要素。
                  - "replaced_split": 整形後のテキストを分割したリスト。
                  - "replaced_split_for_name": 名前用に整形したリスト。
        """
        # 正規表現パターンの設定
        kanji = u'[一-龥]'
        hira_kata = u'[ぁ-んァ-ン]'

        # 漢字とひらがな・カタカナを削除
        replaced_body = re.sub(kanji, "", self.body_text)
        replaced_body = re.sub(hira_kata, "", replaced_body)

        # 特定の特殊文字を削除
        replaced_text = (
            replaced_body.replace("<", "").replace(">", "")
                         .replace('"', "").replace("", "")
        )

        # テキストを分割してリスト化
        replaced_split = replaced_text.split()
        replaced_split_for_name = replaced_text.split()

        # 整形後のデータを辞書形式で格納
        replaced_data = {
            "first_letter_in_replaced_split": replaced_split[0],
            "replaced_split": replaced_split,
            "replaced_split_for_name": replaced_split_for_name
                         }

        return replaced_data

    def process_distance_id_and_pass(self):
        """
        PDFから解析した内容を元に、メールアドレス（ID）、パスワード、担当者名を特定する。
        - テキスト内に含まれるIDとパスワードを検出し、それらの位置を基に距離を計算。
        - 最も近いID、パスワード、および担当者名を特定して辞書形式で返却。

        Returns:
            dict: 検出した情報を辞書形式で返却。
                  - "detected_id": 検出したID。
                  - "detected_pass": 検出したパスワード。
                  - "N.S_PIC": 検出した担当者名。
        """
        # 辞書をリセット
        self.distance_id_dict = {}  # IDの距離情報を保持する辞書
        self.distance_pass_dict = {} # パスワードの距離情報を保持する辞書

        try:
            # IDを探索して距離を計算
            for IDS in self.id_dict.keys():
                if IDS in self.replaced_contents()["first_letter_in_replaced_split"]:  # IDが先頭のテキストに含まれるか確認
                    self.download_mail_add = IDS  # 検出したメールアドレスを保存

                    # テキスト内の「To:」からIDまでの距離を計算
                    distance_to = (
                            self.replaced_contents()["replaced_split"].index(IDS) +
                            self.replaced_contents()["replaced_split"].index("To:")
                                   )
                    self.distance_id_dict[distance_to] = self.id_dict[IDS] # 距離をキーにIDを格納
                    self.distance_pic_dict[distance_to] = self.last_name_dict[IDS]  # 距離をキーに担当者名を格納

            # パスワードを探索して距離を計算
            for PASS in self.pass_dict.keys():
                if PASS in self.replaced_contents()["first_letter_in_replaced_split"]:  # パスワードが先頭のテキストに含まれるか確認

                    # テキスト内の「To:」からパスワードまでの距離を計算
                    distance_pass = (
                            self.replaced_contents()["replaced_split"].index(PASS) +
                            self.replaced_contents()["replaced_split"].index("To:")
                                    )
                    self.distance_pass_dict[distance_pass] = self.pass_dict[PASS]  # 距離をキーにパスワードを格納

            # 最短距離のID、パスワード、担当者名を特定
            self.min_id = min(self.distance_id_dict.keys())  # IDの最短距離を取得
            self.min_pass = min(self.distance_pass_dict.keys()) # パスワードの最短距離を取得
            min_pic = min(self.distance_pic_dict.keys())  # 担当者名の最短距離を取得

            # 特定した情報を辞書形式で返却
            return {
                "detected_id": self.distance_id_dict[self.min_id],  # 検出したID
                "detected_pass": self.distance_pass_dict[self.min_pass],  # 検出したパスワード
                "N.S_PIC": self.distance_pic_dict[min_pic]  # 検出した担当者名
            }
        except ValueError as e:
            # エラーが発生した場合の処理
            messagebox.showinfo("エラー発生", f"{e}")
            reset_move(self)
            self.Event.clear()
            self.root.quit()

    def login_to_cybozu(self, driver, wait):
        """
        サイボウズにログインし、メール一覧ページを開く。
        - ユーザー名とパスワードを入力してログインを試行。
        - ログイン後、メールリストが表示されるまで待機。
        - ログインに失敗した場合、エラーメッセージを表示してアプリを終了。

        Args:
            driver (webdriver.Chrome): Selenium WebDriverオブジェクト。
            wait (WebDriverWait): 要素がロードされるまでの待機を管理するオブジェクト。

        Returns:
            None
        """

        # サイボウズのログインページにアクセス
        driver.get('https://cybozu.com')

        # ログインフォームの要素を取得
        user_name = driver.find_element(By.ID, "username-:0-text")
        password = driver.find_element(By.ID, "password-:1-text")
        login = driver.find_element(By.CLASS_NAME, "login-button")

        # 担当者名を出力
        print("日鋼ス担当者：", self.process_distance_id_and_pass()["N.S_PIC"])

        # ユーザー名とパスワードを入力
        try:
            user_name.send_keys(self.process_distance_id_and_pass()["detected_id"])
            password.send_keys(self.process_distance_id_and_pass()["detected_pass"])
        except Exception as e:
            messagebox.showinfo("メールアドレス不明", f"ログインできません！プログラムを終了します。\n {e}")
            driver.quit()
            sys.exit()

        # ログイン後のタブを切り替える
        handle_array = driver.window_handles
        driver.switch_to.window(handle_array[0])

        # ログインボタンをクリック
        login.click()

        # メールリストのロードを待機
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//ul[@class='list_item_container']")))
        except Exception as e:
            # メールリストがロードされない場合の処理
            try:
                # ビューボタンをクリックし、メールリストが見える状態にする
                view_change_button = wait.until(EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "#mainmenu > div.menuOpen > form:nth-child(3) > span > a")))
                view_change_button.click()
                wait.until(EC.presence_of_element_located((By.XPATH, "//ul[@class='list_item_container']")))
            except Exception as e:
                # それでも失敗した場合はエラーメッセージを表示して終了
                messagebox.showinfo("読み取りエラー", f"メールが読み取れません！アプリケーションを終了します。\n{e}")
                self.root.quit()

    def set_reply_sentence(self, driver):
        """
        メールの返信文を作成する。
        - PDF内の情報を元に送付者の苗字を特定。
        - 担当者が自身の場合と不在の場合で異なる返信文を作成。

        Args:
            driver (webdriver.Chrome): Selenium WebDriverオブジェクト。

        Returns:
            str: 作成された返信文。
        """
        # 変数のリセット
        min_from = None  # 最も近い送付者の距離を初期化

        # CSVファイルから送付者情報を取得
        with open(r"C:\path\to\reply_name.csv", "r") as f:
            for row in csv.DictReader(f):
                # PDFテキストに送付者のメールアドレスが含まれているかチェック
                if row["E-mail"] in self.replaced_contents()["replaced_split_for_name"]:
                    # 距離を計算して辞書に格納
                    distance_from = self.replaced_contents()["replaced_split_for_name"].index(row["E-mail"]) + \
                                    self.replaced_contents()["replaced_split_for_name"].index("From:")
                    self.distance_from_dict[distance_from] = row["苗字"]
                    print(self.distance_from_dict)
                else:
                    # 見つからなかった場合、送付者名を空にする
                    self.cus_name = ""

        # 自身が担当者の場合の処理
        if self.process_distance_id_and_pass()["N.S_PIC"] == self.user:
            try:
                # 最短距離の送付者を特定
                min_from = min(self.distance_from_dict.keys())
            except ValueError:
                # エラー処理：送付者が登録されていない場合
                reset_move(self)
                self.Event.clear()
                messagebox.showinfo("エラー発生",
                                    "Reply_nameに登録がないメールアドレスです！登録してから再度実行してください！")
                messagebox.showinfo("強制終了", "アプリケーションを終了します。")
                driver.quit()
                self.root.quit()

            # 送付者の苗字を取得
            self.cus_name = self.distance_from_dict[min_from]
            print("メール送付者の苗字は" + self.cus_name + "です。")

            # 返信文を作成
            sentence = (
                f"{self.cus_name} 様\n\n"
                "お世話になっております。\n"
                f"ABC Companyの{self.user}です。\n\n"
                "標題の件、返信致します。\n\n"
                "以上、ご確認宜しくお願い致します。\n\n"
            )
            return sentence

        # 他の担当者が不在の場合の処理
        else:
            try:
                # 最短距離の送付者を特定
                min_from = min(self.distance_from_dict.keys())
            except ValueError:
                reset_move(self)
                self.Event.clear()
                messagebox.showinfo("エラー発生",
                                    "Reply_nameに登録がないメールアドレスです！登録してから再度実行してください！")
                messagebox.showinfo("強制終了", "アプリケーションを終了します。")
                driver.quit()
                self.root.quit()

            # 送付者の苗字を取得
            self.cus_name = self.distance_from_dict[min_from]
            absence_user = self.process_distance_id_and_pass()["N.S_PIC"]
            print("メール送付者の苗字は" + self.cus_name + "です。")
            print("不在担当者は" + absence_user + "です。")

            # 不在対応の返信文を作成
            sentence = (
                f"{self.cus_name} 様\n\n"
                "お世話になっております。\n"
                f"ABC Companyの{self.user}と申します。\n\n"
                f"{absence_user}の代理にてご対応致します。\n\n"
                "標題の件、返信させて頂きますので、\n"
                "以上、ご確認お願い致します。\n\n"
            )
            return sentence

    def auto_reply(self):
        """
        自動返信またはメールのダウンロードを行うメイン処理。
        - アニメーション変更
        - サイボウズへのログイン
        - メールタイトルのマッチングと返信/ダウンロード処理

        Args:
            None

        Returns:
            None
        """
        # アニメーションの変更（ダウンロード開始を通知）
        change_move(self)

        # フラグ変数の初期化
        self.mail_title_sub_flg = False  # メールタイトルが空の場合のフラグ
        self.flg = None  # 処理完了フラグ

        # driverとwaitをセットし、サイボウズへのログインを実行
        driver = webdriver.Chrome(service=self.chrome_service, options=self.options)
        wait = WebDriverWait(driver, 3)  # 最大3秒待機
        self.login_to_cybozu(driver, wait)  # サイボウズへのログイン処理

        # メール情報の日付やタイトルを調整
        date_info = adjust_date_info(self, driver)["date_info"]
        date_info2 = adjust_date_info(self, driver)["date_info2"]
        mail_index_title = adjust_date_info(self, driver)["mail_index_title"]
        received_time = adjust_date_info(self, driver)["received_time"]
        lag_9h = adjust_date_info(self, driver)["lag_9h"]
        lag_9h_plus_23h_more = adjust_date_info(self, driver)["lag_9h_plus_23h_more"]

        # メールリストを取得
        mail_title = driver.find_elements(By.XPATH, "//ul[@class='list_item_container']")

        # メールタイトルの整形
        self.mail_index_title = mail_index_title
        fixed_mail_index_title = (self.mail_index_title.replace("�", "").replace("　", "").replace(" ", "")
                                  .replace(",", ".").replace("－", "−").replace("㈱", "").replace("〜", "～"))
        fixed_mail_index_title = re.sub(r'[!-~]', "", fixed_mail_index_title)
        fixed_mail_index_title_nosub = fixed_mail_index_title.replace("㈱", "")
        split_text_join = ",".join(self.split_text)
        split_text_join_fixed = split_text_join.replace(",", "")

        # メールタイトルが空の場合の処理
        if fixed_mail_index_title == "":
            self.mail_title_sub_flg = True
            print("メールタイトルに文字列がありません！")
            print(fixed_mail_index_title_nosub)
            fixed_mail_index_title = fixed_mail_index_title_nosub

        # デバッグ用ログ出力
        print("メールタイトル：" + fixed_mail_index_title)
        print("受信時刻：" + received_time)
        print("9時間の時差の場合:", lag_9h)
        print("9時間時差+23時以上の場合:", lag_9h_plus_23h_more)
        print("受信日:" + date_info)
        print("時差で日をまたいだ受信日:", date_info2)
        print("OCRしたPDF：" + split_text_join_fixed)

        # 検索に適合したメールのカウントを初期化
        self.detected_mail = 0

        # 条件に合うメールを探す
        click_matched_title(self,driver, mail_title, date_info, date_info2, received_time, fixed_mail_index_title,
                                 lag_9h, lag_9h_plus_23h_more, split_text_join_fixed)

        # メールが見つからない場合、次のページを探索
        if self.detected_mail == 0:
            while not self.flg:
                # 次のページボタンをクリック
                next_button = driver.find_element(By.XPATH, "//span[@class = 'icon_btn arrow_right_icon']")
                next_button.click()
                time.sleep(0.5)

                try:
                    # 新しいページのメールリストを取得
                    mail_title = driver.find_elements(By.XPATH, "//ul[@class='list_item_container']")
                except Exception as e:
                    # メールが見つからなければアプリケーションを終了
                    messagebox.showinfo("強制終了", f"アプリケーションを終了します！\n{e}")
                    driver.quit()
                    self.root.quit()

                # 条件に合うメールを再探索
                click_matched_title(self,driver, mail_title, date_info, date_info2, received_time,
                                         fixed_mail_index_title, lag_9h, lag_9h_plus_23h_more, split_text_join_fixed)

        # 返信モードの場合
        if self.var.get() == 0:
            reply_mode(self, driver, wait)  # メール返信処理

        # ダウンロードモード
        elif self.var.get() == 1:
            download_mode(self, driver, wait, split_text_join_fixed)  # メール添付ファイルのダウンロード処理


if __name__ == "__main__":
    app = MainProcess()
