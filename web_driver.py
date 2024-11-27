import sys
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome import service as fs

from subprocess import CREATE_NO_WINDOW
from tkinter import messagebox
import os
import socket
import ctypes
import pygetwindow as gw
import psutil

class WebDriver:
    """
    WebDriverを管理するクラス。
    - ChromeDriverのプロセス管理
    - アプリの多重起動防止
    - アプリウィンドウの管理
    """
    def __init__(self, user):
        """
        クラスの初期化。ChromeDriverの設定とアプリウィンドウの多重起動チェックを行う。

        Args:
            user (str): 使用しているユーザー名。
        """

        # 起動済みのアプリウィンドウハンドルを取得（タイトル名で検索）
        self.app_window_handle = self.get_existing_app_window_handle("ぱいぷー")
        self.PORT = 12345
        self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.options = Options()
        self.chrome_install = ChromeDriverManager().install()
        self.chrome_folder = os.path.dirname(self.chrome_install)
        self.chromedriver_path = os.path.join(self.chrome_folder, "chromedriver.exe")
        self.chrome_service = fs.Service(executable_path=self.chromedriver_path)
        self.chrome_service.creation_flags = CREATE_NO_WINDOW
        self.user = user

    def kill_existing_chromedriver(self):
        """
        既存のChromeDriverプロセスを終了させる。
        """
        for process in psutil.process_iter(attrs=["pid", "name"]):
            process_info = process.as_dict(attrs=["pid", "name"])
            if process_info.get("name") == "chromedriver.exe":
                print(f"Killing leftover ChromeDriver process: {process_info['pid']}")
                process.kill()  # プロセスを強制終了

    def get_existing_app_window_handle(self, title):
        """
        指定されたタイトルのウィンドウハンドルを取得。

        Args:
            title (str): ウィンドウのタイトル。

        Returns:
            int or None: ウィンドウハンドル（存在しない場合はNone）。
        """
        windows = gw.getWindowsWithTitle(title)
        if windows:
            return windows[0]._hWnd  # 最初に見つかったウィンドウのハンドルを返す
        return None

    def bring_app_to_front(self, window_handle):
        """
        指定されたウィンドウを最前面に表示。

        Args:
            window_handle (int): ウィンドウハンドル。
        """
        SW_RESTORE = 9 # ウィンドウを復元する定数
        ctypes.windll.user32.ShowWindow(window_handle, SW_RESTORE)  # 最小化されたウィンドウを復元
        ctypes.windll.user32.SetForegroundWindow(window_handle)  # ウィンドウを最前面に表示

    def update_webdriver(self):
        """
        ChromeDriverを更新し、多重起動を防止。
        """
        try:
            # ポートバインドを試行（多重起動を防止するため）
            self.s.bind(("localhost", self.PORT))
            self.s.listen(1)

            # ChromeDriverの設定オプション
            self.options.add_experimental_option("detach", True)
            self.options.add_experimental_option("excludeSwitches", ["enable-logging"])

            # ChromeDriverの既存プロセスを終了
            self.kill_existing_chromedriver()  # ChromeDriverの既存プロセスを確認して削除

            # chromeドライバーのexeファイルが正しく認識されないケースもあるため、pathを正しく指定する必要あり。
            messagebox.showinfo("ドライバー更新完了", self.user + "さん！ ChromeDriverの更新が正常に完了しました！")
        except socket.error:
            self.bring_app_to_front(self.app_window_handle)
            sys.exit()
