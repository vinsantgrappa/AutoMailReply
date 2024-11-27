import time
from tkinter import messagebox
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchWindowException
import pyautogui
import pyperclip
import re
import difflib
from find_editable_contents import find_editable_element

# モジュールのインストール
from move_manager import reset_move


def reply_mode(instance, driver, wait):
    """
    メール返信モードを処理する関数。
    - 自動返信対象のメールを特定し、返信ウィンドウに自動入力を行う。

    Args:
        instance (object): 呼び出し元クラスのインスタンス。
        driver (selenium.webdriver.Chrome): Selenium WebDriverオブジェクト。
        wait (selenium.webdriver.support.ui.WebDriverWait): 要素が読み込まれるまで待機するためのオブジェクト。

    Returns:
        None
    """

    if instance.detected_mail > 1:
        # 複数の候補が見つかった場合の処理
        print(f"複数のメールが見つかりました。件数: {instance.detected_mail}")
        instance.max_dict.click() # 最も一致したメールをクリック
        time.sleep(0.8)

        # 返信ボタンをクリック
        reply_button = wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "全員に返信")))
        reply_button.click()

        # 新しい返信ウィンドウに切り替え
        handle_array = driver.window_handles
        driver.close()
        driver.switch_to.window(handle_array[1])
        time.sleep(0.8)

        # 返信メッセージを入力
        editable_element = find_editable_element(driver)
        if editable_element:
            print("返信入力欄を見つけました。返信メッセージを入力します。")
            editable_element.send_keys(instance.set_reply_sentence(driver))

        # 使用した変数をリセット
        instance.cus_name = ""
        instance.detected_mail = 0
        instance.ratio_dict = {}
        instance.distance_pic_dict = {}
        instance.distance_from_dict = {}
        instance.flg = None

        # アニメーションをリセット
        reset_move(instance)
        instance.Event.clear()

    else:
        # マッチするメールが1件の場合
        print("1件のメールが見つかりました。返信を開始します。")
        instance.max_dict.click() # 一致したメールをクリック
        time.sleep(0.5)

        # 返信ボタンをクリック
        reply_button = wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "全員に返信")))
        reply_button.click()

        # 返信ウィンドウに切り替え
        instance.flg = None
        handle_array = driver.window_handles
        driver.close()
        driver.switch_to.window(handle_array[1])
        time.sleep(0.8)

        # 返信メッセージを入力
        editable_element = find_editable_element(driver)
        if editable_element:
            print("返信入力欄を見つけました。返信メッセージを入力します。")
            editable_element.send_keys(instance.set_reply_sentence(driver))

        # 使用した変数をリセット
        instance.detected_mail = 0
        instance.ratio_dict = {}
        instance.distance_pic_dict = {}
        instance.distance_from_dict = {}
        instance.cus_name = ""

        # アニメーションをリセット
        reset_move(instance)
        instance.Event.clear()

def download_mode(instance, driver, wait, split_text_join_fixed):
    """
        メール添付ファイルのダウンロードモードを処理する関数。
        - 添付ファイルのリンクをクリックしてダウンロード。
        - 特定の条件（暗号化PDFや特定のリンク）にも対応。

        Args:
            instance (object): 呼び出し元クラスのインスタンス。
                               必須プロパティ:
                               - cus_name (str): 返信先の顧客名。
                               - detected_mail (int): 見つかったメールの数。
                               - attached_pass (str): 添付ファイルのパスワード。
            driver (selenium.webdriver.Chrome): Selenium WebDriverオブジェクト。
            wait (selenium.webdriver.support.ui.WebDriverWait): 要素が読み込まれるまで待機するためのオブジェクト。
            split_text_join_fixed (str): メール本文を結合して整形した文字列。

        Returns:
            None
        """

    # 暗号化ファイルの処理
    if "暗号化" in split_text_join_fixed:
        instance.max_dict.click()
        time.sleep(0.8)

        # パスワードをメール本文から取得
        try:
            if "【パスワード】" in split_text_join_fixed:
                instance.attached_pass_index = instance.split_text.index("【パスワード】") + 1
            elif "Password:" in split_text_join_fixed:
                instance.attached_pass_index = instance.split_text.index("Password:") + 1
        except ValueError:
            messagebox.showinfo("エラー", "パスワードのPDFが添付されていません！")
            instance.root.quit()
        instance.attached_pass = instance.split_text[instance.attached_pass_index]
        instance.attached_pass = instance.attached_pass.replace("＝", "=")

        # ダウンロードリンクをクリックして保存
        if "本メールの添付ファイルは暗号化されております。お客様メールアドレスと後ほどお送りするパスワードにて下記URLからダウンロードし保存してください。" in split_text_join_fixed:
            link_url_index1 = instance.split_text.index("URL】") + 1
            link_url_index2 = instance.split_text.index("URL】") + 2
            link_url = instance.split_text[link_url_index1] + instance.split_text[link_url_index2]
            print("url:" + link_url)
            download_link = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, f"a[href='{link_url}']")))
            download_link.click()
            time.sleep(1)

            # ダウンロードフォームに情報を入力
            pyautogui.press("tab")
            pyautogui.press("tab")
            pyperclip.copy(instance.download_mail_add)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.press("tab")
            pyperclip.copy(instance.attached_pass)
            pyautogui.hotkey("ctrl", "v")
            pyautogui.press("tab")
            pyautogui.press("enter")

            # 使用した変数やフラグをリセット
            instance.detected_mail = 0
            instance.ratio_dict = {}
            instance.distance_pic_dict = {}
            instance.distance_from_dict = {}
            instance.cus_name = ""
            instance.attached_pass = ""

            # GIFアニメーションをリセット
            reset_move(instance)
            instance.Event.clear()

    if instance.cus_name == "モリなび":
        print("クリックするXPATH", instance.max_dict)
        instance.max_dict.click()
        time.sleep(0.8)
        morynavi_add = wait.until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='view_mail']/table[2]/tbody/tr[2]/td/tt/a[1]")))
        morynavi_add.click()
        handle_array = driver.window_handles
        driver.switch_to.window(handle_array[1])
        morynavi_id = wait.until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='main']/div[1]/form/div[1]/div[1]/div/input")))
        morynavi_pass = wait.until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='main']/div[1]/form/div[1]/div[2]/div/input")))
        morynavi_id.send_keys(instance.morynavi_id_input)
        morynavi_pass.send_keys(instance.morynavi_pass_input)
        morynavi_button = wait.until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='main']/div[1]/form/div[3]/button")))
        morynavi_button.click()
        time.sleep(1)
        driver.execute_script("window.scrollTo(document.body.scrollWidth, 0)")
        morynavi_print = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[@id='main']/div/div[2]/div[1]/div[2]/div[1]/div/div/div[1]/button")))
        time.sleep(2)
        morynavi_print.click()
        driver.switch_to.window(handle_array[1])
        time.sleep(1)

        # 変数、辞書ををリセット
        instance.cus_name = ""
        instance.detected_mail = 0
        instance.ratio_dict = {}
        instance.distance_pic_dict = {}
        instance.distance_from_dict = {}
        instance.flg = None
        reset_move(instance)
        instance.Event.clear()
        driver.quit()

    else:
        instance.max_dict.click()
        download_flg = True
        try:
            child_num = 1
            while download_flg:
                try:
                    download_link = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,
                                                                               "#view_mail > table:nth-child(2) > tbody > tr.view_content > td > div.vr_viewContentsAttach > div:nth-child("
                                                                               + str(child_num) + ") > div > a")))
                    download_link.click()
                    child_num += 1

                except Exception as e:
                    print(f"エラー発生。リセットを行います。\n{e}")

                    # 変数、辞書をリセット
                    instance.detected_mail = 0
                    instance.ratio_dict = {}
                    instance.distance_pic_dict = {}
                    instance.distance_from_dict = {}
                    instance.cus_name = ""
                    reset_move(instance)
                    instance.Event.clear()
                    download_flg = False

        except NoSuchWindowException:
            messagebox.showinfo("エラー", "添付ファイルが見つかりませんね、、、一回落とすね")
            reset_move(instance)
            instance.Event.clear()
            instance.root.quit()


def click_matched_title(instance, driver, mail_title, date_info,
                        date_info2, received_time, fixed_mail_index_title,
                        lag_9h, lag_9h_plus_23h_more, split_text_join_fixed):
    """
    メールリストから条件に一致するタイトルをクリックし、本文を解析してマッチング率を計算する。

    Args:
        instance (object): 呼び出し元クラスのインスタンス。
        driver (selenium.webdriver.Chrome): Selenium WebDriverオブジェクト。
        mail_title (list): メールタイトル要素のリスト。
        date_info (str): 日付情報。
        date_info2 (str): 修正後の日付情報（翌日扱い）。
        received_time (str): メール受信時刻。
        fixed_mail_index_title (str): 整形済みのメールタイトル文字列。
        lag_9h (str): 日本時間に修正した時刻。
        lag_9h_plus_23h_more (str): 翌日扱いの修正時刻。
        split_text_join_fixed (str): 整形済みのメール本文データ。

    Returns:
        None
    """
    for i, title in enumerate(mail_title):
        print(f"処理中のメール: {i + 1}/{len(mail_title)}")

        # メールタイトルを整形
        str_title = title.get_attribute("outerHTML")
        replace_title = (
            str_title.replace("㈱", "").replace("　", "").replace(" ", "")
            .replace("�", "").replace(",", ".").replace("－", "−")
        )

        # 特定の記号や漢字を除去して文字列を整形
        kanji_pattern = re.compile(r'[\u4e00-\u9fff]+')
        replace_title_only_kanji = kanji_pattern.sub("", replace_title)
        replace_title_no_mark = re.sub(r'[!-~]', "", replace_title)
        replace_title_no_mark = re.sub(
            r'[\u2460-\u2473\u3240-\u324F\u3251-\u325F\u3260-\u327F\u3280-\u328F\u3290-\u329F]', "",
            replace_title_no_mark)
        split_title = replace_title_no_mark.split()

        # フラグが立っている場合は漢字を使用
        if instance.mail_title_sub_flg:
            split_title = replace_title_only_kanji

        # 条件1: 日付または受信時刻と一致し、タイトルが一致する場合
        if (date_info in replace_title or received_time in replace_title) and fixed_mail_index_title in split_title:
            title.click()
            time.sleep(0.8)

            # メール本文と受信時刻を取得
            try:
                body = driver.find_element(By.XPATH, "//tt")
                re_time = driver.find_element(By.XPATH, "//tr[@class='view_date']")
            except Exception as e:
                print(f"{e}\n HTML構造が異なる場合の代替要素を取得")
                body = driver.find_element(By.XPATH, "//*[@id='view_mail']")
                re_time = driver.find_element(By.XPATH, "//tr[@class='view_date']")
            str_body = body.get_attribute("outerHTML")
            str_re_time = re_time.get_attribute("outerHTML")

            # 本文の不要なタグを除去
            str_fixed_body = str_body.replace("<tt>", "")
            str_fixed_body = str_fixed_body.replace("<br>", "")
            str_fixed_body = str_fixed_body.replace("</tt>", "")
            pattern = r'\s'
            str_fixed_body = re.sub(pattern, "", str_fixed_body)
            last_occurrence = str_fixed_body.rfind('<spanlang="EN-US"')
            str_fixed_body = str_fixed_body[:last_occurrence]

            # 受信時刻が一致する場合
            if received_time in str_re_time or lag_9h in str_re_time or lag_9h_plus_23h_more in str_re_time:
                instance.detected_mail += 1
                print(f"一致するメールを発見（{instance.detected_mail}件目）")
                print(f"メールタイトル: {fixed_mail_index_title}")

                # 本文のマッチング率を計算
                s = difflib.SequenceMatcher(None, split_text_join_fixed, str_fixed_body).ratio()
                instance.ratio_dict[title] = s
                print("マッチング率：", s)
                instance.max_dict = max(instance.ratio_dict, key=instance.ratio_dict.get)
                instance.flg = True

        # 条件2: 修正後の日付（翌日扱い）が一致する場合
        elif (date_info2 in replace_title or lag_9h in replace_title) and fixed_mail_index_title in split_title:
            print(f"修正後の日付（{date_info2}または{lag_9h}）に一致するメールを確認中: {title}")
            title.click()
            time.sleep(0.8)

            try:
                # メール本文と受信時刻の要素を取得
                body = driver.find_element(By.XPATH, "//tt")
                re_time = driver.find_element(By.XPATH, "//tr[@class='view_date']")
            except Exception as e:
                print(f"{e}\n HTML構造が異なる場合の代替要素を取得")
                body = driver.find_element(By.XPATH, "//*[@id='view_mail']")
                re_time = driver.find_element(By.XPATH, "//tr[@class='view_date']")

            # メール本文と時刻を取得し整形
            str_body = body.get_attribute("outerHTML")
            str_re_time = re_time.get_attribute("outerHTML")
            str_fixed_body = str_body.replace("<tt>", "")
            str_fixed_body = str_fixed_body.replace("<br>", "")
            str_fixed_body = str_fixed_body.replace("</tt>", "")

            # 条件: 受信時刻が一致する場合
            if received_time in str_re_time or lag_9h in str_re_time or lag_9h_plus_23h_more in str_re_time:
                instance.detected_mail += 1
                print(f"一致するメールを発見（{instance.detected_mail}件目）")

                # 本文のマッチング率を計算
                ratio_s = difflib.SequenceMatcher(None, split_text_join_fixed, str_fixed_body).ratio()
                instance.ratio_dict[title] = ratio_s
                print(f"マッチング率: {ratio_s}")
                instance.max_dict = max(instance.ratio_dict, key=instance.ratio_dict.get)
                instance.flg = True

        # 条件3: 翌日扱いで23時以降の時刻が一致する場合
        elif (lag_9h_plus_23h_more in replace_title) and fixed_mail_index_title in split_title:
            print(f"翌日扱いで23時以降の条件（{lag_9h_plus_23h_more}）に一致するメールを確認中: {title}")
            title.click()
            time.sleep(0.8)

            try:
                # メール本文と受信時刻の要素を取得
                body = driver.find_element(By.XPATH, "//tt")
                re_time = driver.find_element(By.XPATH, "//tr[@class='view_date']")
            except Exception as e:
                # フォールバック処理
                print(f"{e}\n HTML構造が異なる場合の代替要素を取得")
                body = driver.find_element(By.XPATH, "//*[@id='view_mail']")
                re_time = driver.find_element(By.XPATH, "//tr[@class='view_date']")

            # メール本文と時刻を取得し整形
            str_body = body.get_attribute("outerHTML")
            str_re_time = re_time.get_attribute("outerHTML")
            str_fixed_body = str_body.replace("<tt>", "")
            str_fixed_body = str_fixed_body.replace("<br>", "")
            str_fixed_body = str_fixed_body.replace("</tt>", "")

            # 条件: 受信時刻が一致する場合
            if received_time in str_re_time or lag_9h in str_re_time or lag_9h_plus_23h_more in str_re_time:
                instance.detected_mail += 1
                print(f"一致するメールを発見（{instance.detected_mail}件目）")

                # 本文のマッチング率を計算
                s = difflib.SequenceMatcher(None, split_text_join_fixed, str_fixed_body).ratio()
                instance.ratio_dict[title] = s
                print(f"マッチング率: {s}")
                instance.max_dict = max(instance.ratio_dict, key=instance.ratio_dict.get)
                instance.flg = True

        # 条件に一致しない場合
        else:
            print("条件に一致しないためスキップ")
            print(f"抽出されたタイトル: {split_title}")