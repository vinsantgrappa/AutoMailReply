from tkinter import messagebox

def adjust_date_info(instance, driver):
    """
    メール情報（日時やタイトル）を解析し、調整済みのデータを返す。
    - 時刻を9時間進めることで、タイムゾーンを日本時間に修正。
    - 日付が23時を超えた場合は翌日として処理し、曜日を更新する。
    - メールのタイトルや受信時刻も取得・整形して返す。

    Args:
        instance (object): 呼び出し元クラスのインスタンス。
                           主に以下のプロパティが利用される:
                           - split_text (list): メールの内容を分割した文字列リスト。
                           - config (configparser.ConfigParser): 設定データを保持するオブジェクト。
        driver (selenium.webdriver.Chrome): Selenium WebDriverオブジェクト。

    Returns:
        dict: 以下の情報を格納した辞書:
            - subject (int): メールの件名のインデックス。
            - date_info (str): 日付情報 (月/日 曜日) の文字列。
            - date_info2 (str): 修正後の日付情報 (翌日/曜日) の文字列。
            - mail_index_title (str): メール件名のテキスト。
            - received_time (str): メール受信時刻。
            - lag_9h (str): 日本時間に修正した時刻。
            - lag_9h_plus_23h_more (str): 修正後、翌日扱いの時刻。
    """

    # 初期化：処理中の値を格納する変数をリセット
    subject = None
    num_months = None
    kanji_week = None
    str_date2 = None
    kanji_week2 = None
    lag_9h = None
    lag_9h_plus_23h_more = None
    fixed_flg = False # 翌日に日付がずれる場合にTrueになるフラグ

    # メール件名のインデックスを取得
    try:
        subject = instance.split_text.index("Subject:") + 1
    except Exception as e:
        # 件名が見つからない場合、エラー表示
        messagebox.showinfo("検索エラー", f"該当するメールがありません！\n{e}")

    # "Date:"を基準に日時情報を解析
    Date = instance.split_text.index("Date:")
    date = instance.split_text.index("Date:") + 2  # 日付部分のインデックス
    months = instance.split_text.index("Date:") + 3  # 月部分のインデックス
    week = instance.split_text.index("Date:") + 1  # 曜日部分のインデックス
    times = instance.split_text.index("Date:") + 5  # 時刻部分のインデックス
    str_time = instance.split_text[times]  # 時刻の文字列（例: "14:30"）

    # 時間を9時間進めて、日本時間に調整
    fixed = int(str_time[:2]) + 9
    if fixed > 23:
        fixed2 = fixed - 24
        fixed_flg = True
    else:
        fixed2 = "***" # 翌日扱いでない場合のダミー値

    # 修正後の時刻文字列を作成
    fixed_str_time = str(fixed) + str_time[2:]    # "HH:MM"形式を修正
    fixed_str_time2 = str(fixed2) + str_time[2:]  # 翌日扱いの場合の時刻

    # 時刻文字列を整形（"0"で始まる場合の処理）
    if str_time[0] == "0":
        zero = str_time[1:5] # "09:00" → "9:00"
    else:
        zero = str_time[:5]
    received_time = zero # 元の受信時刻

    # 時刻が "0" で始まる場合は先頭の "0" を削除
    if fixed_str_time[0] == "0":
        zero2 = fixed_str_time[1:5]  # "09:30" → "9:30"
    # 時刻が "HH:MM:SS" 形式の場合は先頭4文字だけ取得
    elif len(fixed_str_time) == 7:
        zero2 = fixed_str_time[:4]  # 秒部分を削除
    # 通常の "HH:MM" の場合はそのまま取得
    else:
        zero2 = fixed_str_time[:5]

    # 同じ処理を翌日扱いの時刻にも適用
    if fixed_str_time2[0] == "0":
        zero3 = fixed_str_time2[1:5]
    elif len(fixed_str_time2) == 7:
        zero3 = fixed_str_time2[:4]
    else:
        zero3 = fixed_str_time2[:5]

    lag_9h = zero2 # 9時間進めた結果の時刻
    lag_9h_plus_23h_more = zero3 # 翌日扱いの時刻

    # 不要な":"を除去
    if ":" in lag_9h[5:]:
        lag_9h = lag_9h[5:].replace(":", "")
    elif ":" in lag_9h_plus_23h_more[5:]:
        lag_9h_plus_23h_more = lag_9h_plus_23h_more[5:].replace(":", "")

    # 月、日付、曜日を変換用辞書で取得
    str_month = instance.split_text[months]
    str_date = instance.split_text[date]
    str_week = instance.split_text[week]
    month = dict(instance.config.items('month'))  # 月の変換辞書
    date_dict = dict(instance.config.items('date'))  # 日付の変換辞書
    week_dict = dict(instance.config.items('week'))  # 曜日の変換辞書
    fixed_week = dict(instance.config.items('fixed_week'))  # 翌日の曜日変換辞書

    # 月、日、曜日を辞書で対応する値に変換
    if str_month in month.keys():
        num_months = month[str_month]
    if str_week in week_dict.keys():
        kanji_week = week_dict[str_week]
    if str_date in date_dict.keys():
        str_date = date_dict[str_date]

    # 翌日に日付がずれる場合の処理
    if fixed_flg:
        try:
            kanji_week2 = fixed_week[str_week] # 翌日の曜日
            str_date2 = str(int(str_date) + 1)  # 翌日の日付 ！！日付が存在するかしないかを判定するコードを追記する必要がある。！！
            fixed_flg = None
        except KeyError as e:
            messagebox.showinfo("エラー発生", str(e)) # 辞書にない場合のエラー表示
            instance.root.quit()
            driver.quit()
    else:
        try:
            str_date = str_date
            kanji_week = kanji_week
            kanji_week2 = "***"  # 翌日の情報がない場合のダミー値
            str_date2 = "***"
        except UnboundLocalError as e:
            messagebox.showinfo("エラー発生", str(e))  # 未定義変数のエラー
            instance.root.quit()
            driver.quit()

    # 整形したデータを辞書で返す
    return {
        "subject": subject,
        "date_info": num_months + "/" + str_date + kanji_week,
        "date_info2": num_months + "/" + str_date2 + kanji_week2,
        "mail_index_title": " ".join(instance.split_text[subject:Date]),
        "received_time": received_time,
        "lag_9h": lag_9h,
        "lag_9h_plus_23h_more": lag_9h_plus_23h_more
    }

