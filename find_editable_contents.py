from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def find_editable_element(driver):
    """
    編集可能な要素（tinymceエディタやtextarea）を探す関数。
    - AjaxFrame、Data_ifr、または直接のtextareaエレメントを順に検索。
    受信したメールによってエレメントが異なる。
    返信文を入力するためのエレメントを探すのに使用する。

    Args:
        driver (selenium.webdriver.Chrome): Selenium WebDriverオブジェクト。

    Returns:
        WebElement: 見つかった編集可能な要素。
        None: 要素が見つからなかった場合。
    """
    wait_time = 1  # 各要素の探索に使用する待機時間

    # Step 1: 'AjaxFrame'内に'tinymce'エディタを探す
    try:
        iframe_ajax = WebDriverWait(driver, wait_time).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "AjaxFrame"))
        )
        print("AjaxFrameに切り替えました")
        # 'tinymce'エディタを探す
        tinymce_body = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.ID, "tinymce"))
        )
        print("AjaxFrame内でtinymceエディタを発見")
        return tinymce_body
    except TimeoutException:
        print("AjaxFrame内にtinymceが見つからないため、次に進みます")
        driver.switch_to.default_content()  # メインコンテンツに戻る

    # Step 2: 'Data_ifr'に切り替えて'tinymce'を探す
    try:
        iframe_data = WebDriverWait(driver, wait_time).until(
            EC.frame_to_be_available_and_switch_to_it((By.ID, "Data_ifr"))
        )
        print("Data_ifrに切り替えました")
        tinymce_body = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.ID, "tinymce"))
        )
        print("Data_ifr内でtinymceエディタを発見")
        return tinymce_body
    except TimeoutException:
        print("Data_ifr内にもtinymceが見つからないため、次に進みます")
        driver.switch_to.default_content()  # メインコンテンツに戻る

    # Step 3: 'textarea' (ID: 'Data') を直接探す
    try:
        data_textarea = WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.ID, "Data"))
        )
        print("textareaエレメント(Data)を発見しました")
        return data_textarea
    except TimeoutException:
        print("textareaエレメントも見つかりませんでした")

    # 要素が見つからない場合はNoneを返す
    return None
