from gif_player import TkGif

def reset_move(instance):
    """
    現在のGIFアニメーションを停止し、スタンバイ状態のアニメーションに切り替える。

    Args:
        instance (object): 呼び出し元クラスのインスタンス。
                           主に以下の属性を利用:
                           - gif_player2: 現在再生中のGIFアニメーション。
                           - pipoo_standby: スタンバイ状態のGIFパス。
                           - label: GIFを表示するラベルウィジェット。
    """
    # 現在再生中のGIFアニメーションを停止
    instance.gif_player2.stop_loop()

    # スタンバイ状態のGIFアニメーションをセット
    instance.gif_player = TkGif(instance.pipoo_standby, instance.label)

    # スタンバイアニメーションを再生
    instance.gif_player.play()

    # GUIを更新
    instance.root.update()


def change_move(instance):
    """
    現在のGIFアニメーションを停止し、ダウンロード状態のアニメーションに切り替える。

    Args:
        instance (object): 呼び出し元クラスのインスタンス。
                           主に以下の属性を利用:
                           - gif_player: 現在再生中のGIFアニメーション。
                           - pipoo_download: ダウンロード状態のGIFパス。
                           - label: GIFを表示するラベルウィジェット。
    """
    # 現在再生中のGIFアニメーションを停止
    instance.gif_player.stop_loop()

    # ダウンロード状態のGIFアニメーションをセット
    instance.gif_player2 = TkGif(instance.pipoo_download, instance.label)

    # ダウンロードアニメーションを再生
    instance.gif_player2.play()

    # GUIを更新
    instance.root.update()

