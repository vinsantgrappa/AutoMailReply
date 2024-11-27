import threading
import tkinter as tk
from PIL import Image, ImageTk


class GifPlayer(threading.Thread):
    """
    GIFアニメーションを再生するクラス。
    threading.Threadを継承してGIFの再生処理を非同期で実行。
    """
    def __init__(self, path: str, label: tk.Label):
        """
        コンストラクタ。

        :param path: 再生するGIFファイルのパス。
        :param label: GIFを表示するTkinterのラベルウィジェット。
        """
        super().__init__(daemon=True)  # スレッドをデーモンとして実行（メイン終了時にスレッドも終了）
        self._please_stop = False  # 停止フラグ
        self.path = path  # GIFファイルのパス
        self.label = label  # 表示先のTkinterラベル
        self.duration = []  # 各フレームの表示間隔（ミリ秒単位）
        self.frame_index = 0  # 現在のフレームインデックス
        self.frames = []  # GIFの全フレームを格納するリスト
        self.last_frame_index = None  # 最後のフレームインデックス

        # GIFのフレームデータを読み込む
        self.load_frames()

    def load_frames(self):
        """
        GIFファイルの全フレームを読み込み、framesリストに格納。
        フレームの表示間隔もdurationリストに保存。
        """
        if isinstance(self.path, str):
            img = Image.open(self.path)  # PILを使用してGIFファイルを開く
            frames = []  # フレームデータを格納するリスト
            duration = []  # フレームごとの表示間隔を格納するリスト
        else:
            # パスが文字列でない場合は何もしない
            return

        frame_index = 0  # フレームカウント用
        is_not_duration = False  # 表示間隔が未指定かどうかのフラグ

        try:
            # GIFの各フレームを順次読み込む
            while True:
                frames.append(ImageTk.PhotoImage(img.copy()))  # フレームをImageTkオブジェクトとして保存
                img.seek(frame_index)  # 次のフレームへ
                frame_index += 1
                try:
                    # フレームの表示間隔を取得してリストに追加
                    duration.append(img.info['duration'])
                except KeyError:
                    # duration情報がない場合
                    is_not_duration = True  # フラグをON
        except EOFError:
            # 最後のフレームに到達したら終了
            if is_not_duration:
                # durationが指定されていない場合、デフォルト値を設定
                duration = [60] * len(frames)
            self.frames = frames  # 全フレームを保存
            self.duration = duration  # 全フレームの表示間隔を保存
            self.last_frame_index = frame_index - 1  # 最後のフレームインデックスを保存

    def run(self):
        """
        スレッドが開始されたときに呼ばれるメソッド。
        再生を開始する。
        """
        self.next_frame()

    def next_frame(self):
        """
        次のフレームを表示する。
        """
        if not self._please_stop:
            # configでフレーム変更
            self.label.configure(image=self.frames[self.frame_index])
            self.frame_index += 1
            # 最終フレームになったらフレームを０に戻す
            if self.frame_index > self.last_frame_index:
                self.frame_index = 0
        # durationミリ秒あとにコールバック
        self.label.after(self.duration[self.frame_index], self.next_frame)

    def stop(self):
        """
        GIF再生を停止する。
        """
        self._please_stop = True  # 停止フラグをONにする


class TkGif():
    """
    TkinterラベルでGIFを再生するための簡易ラッパークラス。
    """
    def __init__(self, path, label: tk.Label) -> None:
        """
        コンストラクタ。

        :param path: 再生するGIFファイルのパス。
        :param label: GIFを表示するTkinterのラベルウィジェット。
        """
        self.path = path  # GIFファイルのパス
        self.label = label  # 表示先のラベルウィジェット


    def play(self):
        """
        GIFアニメーションの再生を開始する。
        """
        self.player = GifPlayer(self.path, self.label)
        self.player.start()  # GifPlayerスレッドを開始

    def stop_loop(self):
        """
        GIFアニメーションのループを停止する。
        """
        self.player.stop()  # GifPlayerの停止フラグを設定

