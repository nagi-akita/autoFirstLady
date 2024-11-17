import pyautogui
import pygetwindow as gw
import time
import boto3

# 定数
NINMEI_BTN_X = 296
NINMEI_BTN_Y = 177
CLOSE_BTN_X = 318
CLOSE_BTN_Y = 46
LIST_BTN_X = 335
LIST_BTN_Y = 613

KENZO_X = 76
KENZO_Y = 498
KAGAKU_X = 187
KAGAKU_Y = 497
NAIMU_X = 307
NAIMU_Y = 498
BOEI_X = 307
BOEI_Y = 350
SENRYAKU_X = 187
SENRYAKU_Y = 350
FUJIN_X = 74
FUJIN_Y = 350

# S3のバケット名とファイル名
bucket_name = 'lastwar468fl'
file_name = 'flflg.txt'

# S3クライアントを作成
s3 = boto3.client('s3')

# 任命条件を確認するための関数
def should_appoint_fujin():
    try:
        # S3からflflg.txtを取得
        response = s3.get_object(Bucket=bucket_name, Key=file_name)
        content = response['Body'].read().decode('utf-8').strip()

        # flflg.txtの内容が '1' なら任命
        return content == '1'
    except Exception as e:
        # 何らかのエラーが発生した場合は任命を行う
        return True

# flflg.txtの値を0にリセットする関数
def reset_flflg_to_zero():
    try:
        # flflg.txtを"0"に更新してS3にアップロード
        s3.put_object(Bucket=bucket_name, Key=file_name, Body='0', ContentType='text/plain')
    except Exception as e:
        print(f"flflg.txtの更新に失敗しました: {e}")

# 無限ループ
try:
    while True:
        # BlueStacksのウィンドウを取得
        window = gw.getWindowsWithTitle('BlueStacks App Player')[0]

        if window:
            window.activate()

        # ウィンドウの左上座標とサイズを取得
        window_x, window_y = window.topleft

        # 任命メソッドの定義
        def ninmei(swipe_cnt, ninmei_cnt):
            # 一覧
            time.sleep(2)
            pyautogui.moveTo(window_x + LIST_BTN_X, window_y + LIST_BTN_Y)
            pyautogui.click()

            # スワイプの開始座標 (BlueStacksのウィンドウ内)
            start_x = window_x + 330
            start_y = window_y + 200  # スワイプ開始点 (画面下部)

            # スワイプする距離
            swipe_distance = 400

            # スワイプ速度 (大きな数値ほど遅い)
            swipe_duration = 0.5

            # 下方向にスワイプを実行
            for _ in range(swipe_cnt):
                # マウスを指定の開始位置に移動し、ドラッグ操作でスワイプ
                time.sleep(1)  # スワイプ間の待機時間
                pyautogui.moveTo(start_x, start_y)
                pyautogui.dragTo(start_x, start_y + swipe_distance, duration=swipe_duration)

            # 任命
            for _ in range(ninmei_cnt):
                time.sleep(2)
                pyautogui.moveTo(window_x + NINMEI_BTN_X, window_y + NINMEI_BTN_Y)
                pyautogui.click()
            
            # 閉じる
            for _ in range(2):
                time.sleep(1)
                pyautogui.moveTo(window_x + CLOSE_BTN_X, window_y + CLOSE_BTN_Y)
                pyautogui.click()        


        # 夫人の任命判定
        if should_appoint_fujin():
            # 夫人
            time.sleep(2)
            pyautogui.moveTo(window_x + FUJIN_X, window_y + FUJIN_Y)
            pyautogui.click()

            # 任命
            ninmei(1, 1)

            # 任命後、flflg.txtの値を0にリセット
            reset_flflg_to_zero()

        # 建造
        time.sleep(2)
        pyautogui.moveTo(window_x + KENZO_X, window_y + KENZO_Y)
        pyautogui.click()

        # 任命
        ninmei(7, 3)
        # 建造

        # 科学
        time.sleep(2)
        pyautogui.moveTo(window_x + KAGAKU_X, window_y + KAGAKU_Y)
        pyautogui.click()

        # 任命
        ninmei(3, 3)
        # 科学

        # 内務
        time.sleep(2)
        pyautogui.moveTo(window_x + NAIMU_X, window_y + NAIMU_Y)
        pyautogui.click()

        # 任命
        ninmei(3, 3)
        # 内務

        # 防衛
        time.sleep(2)
        pyautogui.moveTo(window_x + BOEI_X, window_y + BOEI_Y)
        pyautogui.click()

        # 任命
        ninmei(3, 3)
        # 防衛

        # 戦略
        time.sleep(2)
        pyautogui.moveTo(window_x + SENRYAKU_X, window_y + SENRYAKU_Y)
        pyautogui.click()

        # 任命
        ninmei(1, 1)
        # 戦略

except KeyboardInterrupt:
    print("操作を中断しました。")
