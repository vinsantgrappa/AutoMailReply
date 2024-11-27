import os
import configparser
from Crypto.Cipher import AES
import base64

class LoadConfig:
    """
    暗号化された設定ファイルを読み込んでデコードし、
    必要なユーザー情報やログイン情報を管理するクラス。
    """
    def __init__(self):
        """
        クラスの初期化メソッド。
        各種パス、設定ファイル、ログイン情報を初期化して読み込む。
        """
        self.user = None  # 現在のユーザー名

        # 暗号化設定ファイルのパス
        self.USER_INFO_CONFIG_ENC_PATH = r"C:\path\to\user_info_config.enc"

        # 設定ファイルを読み込む
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        self.config.read_string(self.config_strings(self.USER_INFO_CONFIG_ENC_PATH))

        # 設定ファイルから取得する情報
        self.users_info_dict = dict(self.config.items('users_info_dict'))
        self.file_directory = os.environ["USERNAME"]

        # サイボウズにログインするIDとPASSの辞書をiniファイルから取得
        self.id_dict = dict(self.config.items('id_dict'))
        self.pass_dict = dict(self.config.items('pass_dict'))
        self.last_name_dict = dict(self.config.items('last_name_dict'))

        # サイボウズにログインするIDとPASSの辞書をiniファイルから取得
        self.id_dict = dict(self.config.items('id_dict'))
        self.pass_dict = dict(self.config.items('pass_dict'))
        self.last_name_dict = dict(self.config.items('last_name_dict'))

    def identify_user(self):
        """
        現在のユーザー名から該当するユーザー情報を特定する。

        Returns:
            str: 特定されたユーザー名。
        """
        for user_info in self.users_info_dict.keys():
            if user_info in self.file_directory:  # ファイルディレクトリ名と一致するか確認
                self.user = self.users_info_dict[user_info]

                return self.user

    def config_strings(self, file_path):
        """
        暗号化された設定ファイルを復号化して設定文字列を取得する。

        Args:
            file_path (str): 暗号化設定ファイルのパス。

        Returns:
            str: 復号化された設定データ。
        """
        config_string = ""
        ENCRYPTED_FILE_PATH = file_path
        master_key_base64 = os.getenv('MASTER_KEY')  # 環境変数からマスターキーを取得

        if not master_key_base64:
            raise ValueError("MASTER_KEY is not set in environment variables.")  # エラー：マスターキー未設定

        # BASE64エンコードされたマスターキーをデコード
        master_key = base64.b64decode(master_key_base64)

        # 暗号化ファイルを読み込み
        with open(ENCRYPTED_FILE_PATH, "rb") as dict_enc_file:
            nonce = dict_enc_file.read(16)  # 暗号化時のノンス
            tag = dict_enc_file.read(16)  # 認証タグ
            ciphertext = dict_enc_file.read()  # 暗号化されたデータ

        # AES EAXモードで復号化
        cipher = AES.new(master_key, AES.MODE_EAX, nonce=nonce)

        try:
            config_dict_data = cipher.decrypt_and_verify(ciphertext, tag)  # 復号化と認証を実行
            config_string = config_dict_data.decode('shift_jis')  # shift_jisでデコード
        except ValueError as e:
            print("復号化に失敗しました:", e)
        except UnicodeDecodeError as e:
            print("デコードに失敗しました:", e)

        return config_string

    def run(self):
        """
        クラスのメイン処理を実行。
        現在のユーザーを特定する。
        """
        self.identify_user()


if __name__ == "__main__":
    # クラスを実行
    loader = LoadConfig()
    loader.run()