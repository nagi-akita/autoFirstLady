import boto3
import pyautogui
import pygetwindow as gw
import subprocess
import time
import os
import cv2
import numpy as np
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage

# AWS S3 設定
bucket_name = 'lastwar468fl'
s3_client = boto3.client('s3')

# スクリーンショットを保存するディレクトリ
screenshot_dir = 'screenshots'
list_file_path = os.path.join(screenshot_dir, 'list.txt')

# 基準となるスクリーンショットのパス
CORRECT_SCREENSHOT_PATH = os.path.join(screenshot_dir, "screenshot-correct.png")

# 類似度の閾値（この値未満の場合、迷子と判定）
SIMILARITY_THRESHOLD = 0.5

# 連続で閾値未満だった回数をカウント
consecutive_low_similarity_count = 0

# 最後にメールを送信した時点でのカウント値
last_email_sent_count = 0

# ディレクトリがなければ作成
if not os.path.exists(screenshot_dir):
    os.makedirs(screenshot_dir)

# Gmailのアカウント情報
smtp_server = 'smtp.gmail.com'
smtp_port = 587  # TLSを使用するポート
# gmail_user = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@gmail.com'  # 送信元メールアドレス エラー通知するならgmailで可能
# gmail_password = 'xxxxxxxxxxxxxxxxxx'  # アプリパスワード

# 送信先メールアドレス
# recipient_email = 'xxxxxxxxx@xxxxxxxxxxxxx.com'

# 定数
NINMEI_BTN_X = 296
NINMEI_BTN_Y = 177
CLOSE_BTN_X = 318
CLOSE_BTN_Y = 46
LIST_BTN_X = 335
LIST_BTN_Y = 613
KICK_BTN_X = 330
KICK_BTN_Y = 177
OK_BTN_X = 220
OK_BTN_Y = 450

KENZO_X = 76
KENZO_Y = 498
KAGAKU_X = 187
# KAGAKU_Y = 497
KAGAKU_Y = 580
NAIMU_X = 307
NAIMU_Y = 498
BOEI_X = 307
BOEI_Y = 350
SENRYAKU_X = 187
SENRYAKU_Y = 380
FUJIN_X = 74
FUJIN_Y = 350

# 画像類似度計算関数
def calculate_image_similarity(img1_path, img2_path):
    """
    2つの画像の類似度を計算します。
    
    Args:
        img1_path (str): 比較する画像1のパス
        img2_path (str): 比較する画像2のパス
        
    Returns:
        float: 類似度スコア（0～1、1が完全一致）
    """
    try:
        # 画像を読み込む（日本語パス対応）
        with open(img1_path, 'rb') as f:
            img_array = np.asarray(bytearray(f.read()), dtype=np.uint8)
            img1 = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
        
        with open(img2_path, 'rb') as f:
            img_array = np.asarray(bytearray(f.read()), dtype=np.uint8)
            img2 = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
        
        if img1 is None or img2 is None:
            print("画像の読み込みに失敗しました")
            return 0.0
        
        # 画像サイズを統一
        if img1.shape != img2.shape:
            img2 = cv2.resize(img2, (img1.shape[1], img1.shape[0]))
        
        # グレースケール変換
        gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        
        # ヒストグラム比較
        hist_scores = []
        for i in range(3):  # BGR各チャンネル
            hist1 = cv2.calcHist([img1], [i], None, [256], [0, 256])
            hist2 = cv2.calcHist([img2], [i], None, [256], [0, 256])
            
            # ヒストグラムを正規化
            cv2.normalize(hist1, hist1, 0, 1, cv2.NORM_MINMAX)
            cv2.normalize(hist2, hist2, 0, 1, cv2.NORM_MINMAX)
            
            # 比較（相関法）
            score = cv2.compareHist(hist1, hist2, cv2.HISTCMP_CORREL)
            hist_scores.append(score)
        
        hist_score = sum(hist_scores) / 3
        
        # 構造的類似性指標（SSIM）
        try:
            from skimage.metrics import structural_similarity as ssim
            ssim_score, _ = ssim(gray1, gray2, full=True)
            ssim_score = max(0, ssim_score)  # SSIMは-1～1なので、負の値は0に
        except ImportError:
            # scikit-imageがインストールされていない場合
            ssim_score = 0.5  # デフォルト値
        
        # 重み付き平均（scikit-imageがない場合はヒストグラムのみ）
        if 'ssim' in locals():
            combined_score = 0.5 * hist_score + 0.5 * ssim_score
        else:
            combined_score = hist_score
        
        return combined_score
    
    except Exception as e:
        print(f"類似度計算エラー: {str(e)}")
        return 0.0

# メール送信（画像添付機能付き）
def send_email(subject, body, image_path=None):
    try:
        # メールのヘッダーと本文を作成
        msg = MIMEMultipart()
        msg['From'] = gmail_user
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # メール本文を追加
        msg.attach(MIMEText(body, 'plain'))
        
        # 画像を添付（指定されている場合）
        if image_path and os.path.exists(image_path):
            with open(image_path, 'rb') as img_file:
                img_data = img_file.read()
                image = MIMEImage(img_data, name=os.path.basename(image_path))
                msg.attach(image)

        # SMTPサーバーに接続してメールを送信
        # with smtplib.SMTP(smtp_server, smtp_port) as server:
            # server.starttls()  # TLS暗号化を開始
            # server.login(gmail_user, gmail_password)  # Gmailにログイン
            # server.send_message(msg)  # メールを送信

        print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email: {e}")

# メールの件名と本文を設定して送信
send_email("副大統領スクリーンショット起動", "スクリーンショットのバッチが起動されました。")

# 古いスクリーンショットの削除
def delete_old_screenshots():
    now = datetime.now()
    cutoff_time = now - timedelta(days=1)

    for filename in os.listdir(screenshot_dir):
        # 基準画像は削除しない
        if filename == "screenshot-correct.png":
            continue
            
        filepath = os.path.join(screenshot_dir, filename)
        # ファイルがpng形式であるか確認
        if os.path.isfile(filepath) and filename.endswith('.png'):
            file_time = datetime.fromtimestamp(os.path.getmtime(filepath))
            if file_time < cutoff_time:
                os.remove(filepath)
                print(f'Deleted old screenshot: {filename}')

# 任命条件を確認するための関数
def should_appoint_fujin():
    try:
        # S3からvpflg.txtを取得
        response = s3_client.get_object(Bucket=bucket_name, Key='vpflg.txt')
        content = response['Body'].read().decode('utf-8').strip()
        return content  # 0, 1, 2, 3 のいずれかを返す
    except Exception as e:
        print(f"vpflg.txtの取得に失敗しました: {e}")
        return '0'  # デフォルトで何もしない

# vpflg.txtの値を0にリセットする関数
def reset_vpflg_to_zero():
    try:
        s3_client.put_object(Bucket=bucket_name, Key='vpflg.txt', Body='0', ContentType='text/plain')
    except Exception as e:
        print(f"vpflg.txtの更新に失敗しました: {e}")

# 割り込み任命処理関数
def interrupt_appointment_handler(window_x, window_y):
    """
    vpflg.txt の値に基づいて割り込み処理を行う:
    0: 何もしない
    1: 副大統領を任命
    2: キックを実行
    3: 強制再起動
    4: メンテナンスモード
    """
    flflg_value = should_appoint_fujin()
    if flflg_value == '1':  # 夫人の任命
        print(f"副大統領の任命")
        time.sleep(2)
        pyautogui.moveTo(window_x + FUJIN_X, window_y + FUJIN_Y)
        pyautogui.click()
        ninmei(1, 1)
        reset_vpflg_to_zero()
    elif flflg_value == '2':  # キックの実行
        print(f"キックの実行")
        time.sleep(2)
        pyautogui.moveTo(window_x + FUJIN_X, window_y + FUJIN_Y)
        pyautogui.click()
        time.sleep(3)
        pyautogui.moveTo(window_x + LIST_BTN_X, window_y + LIST_BTN_Y)
        pyautogui.click()
        time.sleep(3)
        pyautogui.moveTo(window_x + KICK_BTN_X, window_y + KICK_BTN_Y)
        pyautogui.click()
        time.sleep(3)
        pyautogui.moveTo(window_x + OK_BTN_X, window_y + OK_BTN_Y)
        pyautogui.click()
        reset_vpflg_to_zero()
        # 閉じる
        
        for _ in range(2):
            time.sleep(1)
            pyautogui.moveTo(window_x + CLOSE_BTN_X, window_y + CLOSE_BTN_Y)
            pyautogui.click()
    elif flflg_value == '3':  # 強制再起動
        print("BlueStacksを強制再起動します。")
        restart_bluestacks()
        reset_vpflg_to_zero()
        return  # メインループの最初から再開
    elif flflg_value == '4':  # メンテナンスモード
        print("メンテナンスモードに入ります。")
        maintenance_mode(window_x, window_y)
        return  # メインループの最初から再開

# 任命メソッドの定義
def ninmei(swipe_cnt, ninmei_cnt):
    # BlueStacksのウィンドウを取得
    window = gw.getWindowsWithTitle('BlueStacks App Player')[0]

    if window:
        window.activate()

    # ウィンドウの左上座標とサイズを取得
    window_x, window_y = window.topleft

    # 一覧
    time.sleep(2)
    pyautogui.moveTo(window_x + LIST_BTN_X, window_y + LIST_BTN_Y)
    pyautogui.click()

    # スワイプの開始座標
    start_x = window_x + 330
    start_y = window_y + 200  # スワイプ開始点

    # スワイプする距離
    swipe_distance = 400

    # スワイプ速度
    swipe_duration = 0.5

    # 下方向にスワイプを実行
    for _ in range(swipe_cnt):
        time.sleep(1)
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

# スクリーンショット取得
def capture_screenshot(window, screenshot_dir):
    global consecutive_low_similarity_count
    
    if window is None or not window.visible or window.width <= 0 or window.height <= 0:
        send_email("Warning: BlueStacks Not Found or Invalid Window Size", "Screenshot skipped: BlueStacks is not running or window size is invalid.")
        print("BlueStacks not running or invalid window size detected. Restarting BlueStacks.")
        restart_bluestacks()
        return False
    elif window.width <= window.height:
        screenshot = pyautogui.screenshot(region=(window.left, window.top + 40, window.width - 40, window.height - 40))

        # タイトルとして「YYYY年MM月DD日HH時MM分」を取得
        title = time.strftime('%Y年%m月%d日%H時%M分')

        # ファイル名にタイムスタンプを使用
        timestamp = time.strftime('%Y%m%d-%H%M%S')
        screenshot_filename = f'screenshot-{timestamp}.png'
        screenshot_path = os.path.join(screenshot_dir, screenshot_filename)

        # スクリーンショットを保存
        screenshot.save(screenshot_path)

        # 基準画像と比較（基準画像が存在する場合）
        if os.path.exists(CORRECT_SCREENSHOT_PATH):
            try:
                # メンテナンスモード時は類似度チェックをスキップ
                flflg_value = should_appoint_fujin()
                if flflg_value != '4':
                    # 類似度を計算
                    similarity = calculate_image_similarity(CORRECT_SCREENSHOT_PATH, screenshot_path)
                    print(f"画像類似度: {similarity:.4f}")
                    
                    # 類似度が閾値未満の場合
                    if similarity < SIMILARITY_THRESHOLD:
                        print(f"類似度が閾値未満です: {similarity:.4f} < {SIMILARITY_THRESHOLD}")
                        consecutive_low_similarity_count += 1
                        
                        global last_email_sent_count
                        
                        # 二回連続で閾値未満の場合、迷子の可能性ありとしてメール送信
                        if consecutive_low_similarity_count >= 2 and consecutive_low_similarity_count - last_email_sent_count >= 2:
                            alert_message = f"迷子の可能性あり！画像類似度が低下しています。類似度: {similarity:.4f}"
                            print(alert_message)
                            send_email(
                                subject="【警告】迷子の可能性あり", 
                                body=f"{alert_message}\n\n現在時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n連続検出回数: {consecutive_low_similarity_count}回", 
                                image_path=screenshot_path
                            )
                            # 最後にメールを送信した時点でのカウント値を更新
                            last_email_sent_count = consecutive_low_similarity_count
                        
                        # 10回連続で閾値未満の場合、再度メール送信
                        elif consecutive_low_similarity_count - last_email_sent_count >= 10:
                            alert_message = f"迷子の可能性あり！画像類似度が低下し続けています。類似度: {similarity:.4f}"
                            print(alert_message)
                            send_email(
                                subject="【警告】迷子の可能性あり", 
                                body=f"{alert_message}\n\n現在時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n連続検出回数: {consecutive_low_similarity_count}回", 
                                image_path=screenshot_path
                            )
                            # 最後にメールを送信した時点でのカウント値を更新
                            last_email_sent_count = consecutive_low_similarity_count
                    else:
                        # 類似度が閾値以上の場合、カウンターをリセット
                        consecutive_low_similarity_count = 0
                        last_email_sent_count = 0
                else:
                    print("メンテナンスモード中: 画像類似度チェックをスキップします")
            except Exception as e:
                print(f"画像類似度計算中にエラーが発生しました: {e}")
        else:
            # 基準画像が存在しない場合、現在のスクリーンショットを基準画像として保存
            print(f"基準画像が見つかりません。現在のスクリーンショットを基準画像として保存します: {CORRECT_SCREENSHOT_PATH}")
            screenshot.save(CORRECT_SCREENSHOT_PATH)

        # S3にアップロード
        s3_client.upload_file(screenshot_path, bucket_name, screenshot_filename)
        s3_url = f'https://{bucket_name}.s3.amazonaws.com/{screenshot_filename}'

        # list.txt にタイトルとS3 URLを記載（新しいものを先頭に追加）
        with open(list_file_path, 'r+', encoding='utf-8') as file:
            lines = file.readlines()
            lines.insert(0, f'{title} {s3_url}\n')
            file.seek(0)
            file.writelines(lines[:10])
            file.truncate()

            # S3バケット内の古いファイルを削除
            if len(lines) > 10:
                old_filename = lines[-1].strip().split()[-1].split('/')[-1]
                s3_client.delete_object(Bucket=bucket_name, Key=old_filename)
                print(f'Deleted old screenshot from S3: {old_filename}')
            
            s3_client.upload_file(list_file_path, bucket_name, 'list.txt')
            
        return True
    else:
        send_email("Warning: Invalid Window Size", "Screenshot skipped: Width is greater than height.")
        print("Invalid window size detected. Restarting BlueStacks.")
        restart_bluestacks()
        return False



# メンテナンスモード処理
def maintenance_mode(window_x, window_y):
    """
    メンテナンスモード中の処理:
    1. 10秒に一度 coordinates.txt を参照
    2. coordinates.txt に座標があれば、Bluestacksのその位置をクリック
    3. クリックした後 coordinates.txt の内容を空にする
    4. クリックした3秒後にスクリーンショットを撮り、S3にアップロード
    5. coordinates.txt の中に座標がなければ vpflg.txt を参照
    6. vpflg.txt の値が4であれば1の処理に戻る
    7. vpflg.txt の値が4以外の場合はメインループを再開
    """
    # 2秒待機
    time.sleep(5)

    print("メンテナンスモード処理を開始します。")
    
    # メンテナンスモード開始時に最新のスクリーンショットを撮影
    windows = gw.getWindowsWithTitle('BlueStacks App Player')
    if windows:
        window = windows[0]
        window.activate()
        # ウィンドウの左上座標を取得
        window_x, window_y = window.topleft
        print(f"BlueStacksウィンドウの座標: ({window_x}, {window_y})")
        print("メンテナンスモード開始時のスクリーンショットを撮影します。")
        capture_screenshot(window, screenshot_dir)
    
    while True:
        try:
            # BlueStacksのウィンドウを取得して必ずアクティブにする
            windows = gw.getWindowsWithTitle('BlueStacks App Player')
            if not windows:
                print("BlueStacksウィンドウが見つかりません。")
                time.sleep(10)
                continue
                
            window = windows[0]
            window.activate()
            
            # ウィンドウの左上座標を取得し直す
            window_x, window_y = window.topleft
            
            # coordinates.txt を参照
            try:
                response = s3_client.get_object(Bucket=bucket_name, Key='coordinates.txt')
                coordinates = response['Body'].read().decode('utf-8').strip()
                
                # 座標があればクリック
                if coordinates:
                    print(f"座標を検出: {coordinates}")
                    try:
                        x, y = map(int, coordinates.split(','))
                        
                        # クリック
                        pyautogui.moveTo(window_x + x, window_y + y)
                        pyautogui.click()
                        print(f"座標 ({x}, {y}) をクリックしました。")
                        
                        # coordinates.txt の内容を空にする
                        s3_client.put_object(Bucket=bucket_name, Key='coordinates.txt', Body='', ContentType='text/plain')
                        print("coordinates.txtの内容を空にしました。")
                        
                        # 3秒待機
                        time.sleep(3)
                        
                        # スクリーンショットを撮影
                        capture_screenshot(window, screenshot_dir)
                        print("スクリーンショットを撮影しました。")
                    except ValueError:
                        print(f"無効な座標形式: {coordinates}")
                        s3_client.put_object(Bucket=bucket_name, Key='coordinates.txt', Body='', ContentType='text/plain')
                else:
                    # 座標がなければvpflg.txtを参照
                    flflg_value = should_appoint_fujin()
                    if flflg_value != '4':
                        print("メンテナンスモードを終了します。")
                        return  # メンテナンスモードを終了
            except Exception as e:
                print(f"coordinates.txtの参照に失敗しました: {e}")
                
            # 5秒待機
            time.sleep(5)
            
        except Exception as e:
            print(f"メンテナンスモード中にエラーが発生しました: {e}")
            time.sleep(10)

# BlueStacks再起動
def restart_bluestacks():
    try:
        try:
            subprocess.call(["taskkill", "/F", "/IM", "HD-Player.exe"])
        except subprocess.SubprocessError:
            print("Failed to terminate BlueStacks process. Continuing with restart.")
        send_email("BlueStacks Restart", "BlueStacks is being restarted due to an issue.")
        time.sleep(5)  # 少し待機してから再起動
        subprocess.Popen([
            "C:\\Program Files\\BlueStacks_nxt\\HD-Player.exe", 
            "--instance", "Pie64", 
            "--cmd", "launchAppWithBsx", 
            "--package", "com.fun.lastwar.gp", 
            "--source", "desktop_shortcut"
        ])
        print("BlueStacks restarted.")
        time.sleep(120)  # 再起動後の安定性確保のため待機   
    except Exception as e:
        print(f"Failed to restart BlueStacks: {e}")
        time.sleep(60)  # エラーが発生した場合は60秒待機
        restart_bluestacks()
    
# メインループ
try:
    while True:
        try:
            delete_old_screenshots()

            windows = gw.getWindowsWithTitle('BlueStacks App Player')
            if not windows:
                send_email("Error: BlueStacks Window Not Found", "BlueStacks may have crashed.")
                restart_bluestacks()
                continue

            window = windows[0]
            # スクリーンショットの取得
            time.sleep(2)
            if not capture_screenshot(window, screenshot_dir):
                continue

            window.activate()

            # ウィンドウの左上座標とサイズを取得
            window_x, window_y = window.topleft

            # 初期化処理
#            time.sleep(2)
#            pyautogui.moveTo(window_x + 337, window_y + 150) # キャンペーンの×ボタン
#            pyautogui.click()
#            time.sleep(2)
#            pyautogui.moveTo(window_x + 29, window_y + 55) # プロフィール
#            pyautogui.click()
#            time.sleep(2)
#            pyautogui.moveTo(window_x + 245, window_y + 498) # 468ボタン
#            pyautogui.click()


#pyautogui.moveTo(window_x+337,window_y+81)
#pyautogui.click()
#time.sleep(2)
#pyautogui.moveTo(window_x+251,window_y+122)
#pyautogui.click()
#time.sleep(2)
#pyautogui.moveTo(window_x+245,window_y+498)
#pyautogui.click()
#time.sleep(2)


            # 各処理の間に割り込み任命処理を挿入
            if interrupt_appointment_handler(window_x, window_y) is not None:
                continue

            # 建造
            time.sleep(2)
            pyautogui.moveTo(window_x + KENZO_X, window_y + KENZO_Y)
            pyautogui.click()
            ninmei(7, 3)

            if interrupt_appointment_handler(window_x, window_y) is not None:
                continue

            # 科学
            time.sleep(2)
            pyautogui.moveTo(window_x + KAGAKU_X, window_y + KAGAKU_Y)
            pyautogui.click()
            ninmei(7, 3)

            if interrupt_appointment_handler(window_x, window_y) is not None:
                continue

            # 内務
            time.sleep(2)
            pyautogui.moveTo(window_x + NAIMU_X, window_y + NAIMU_Y)
            pyautogui.click()
            ninmei(3, 3)

            if interrupt_appointment_handler(window_x, window_y) is not None:
                continue

            # 防衛
            time.sleep(2)
            pyautogui.moveTo(window_x + BOEI_X, window_y + BOEI_Y)
            pyautogui.click()
            ninmei(3, 3)

            if interrupt_appointment_handler(window_x, window_y) is not None:
                continue

            # 戦略
            time.sleep(2)
            pyautogui.moveTo(window_x + SENRYAKU_X, window_y + SENRYAKU_Y)
            pyautogui.click()
            ninmei(1, 1)
            time.sleep(1)
        except Exception as e:
            print(f"予期しないエラーが発生しました: {e}")
            send_email("Error: Unexpected Exception", f"An unexpected error occurred: {e}")
            time.sleep(60)  # 1分待機
            continue  # ループを最初から再開
                        
        
except KeyboardInterrupt:
    print("操作を中断しました。")
