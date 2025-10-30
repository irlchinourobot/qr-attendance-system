# -*- coding: utf-8 -*-
import os
import datetime
import jwt
import gspread
import qrcode
import io
from flask import Flask, render_template, request, redirect, session, url_for, send_file, jsonify
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from haversine import haversine, Unit # 距離計算ライブラリ
# -----------------------------------------------------------------------------
# 初期設定
# -----------------------------------------------------------------------------
app = Flask(__name__)
# 環境変数からFlaskのシークレットキーを読み込む (セッション管理に必要)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'default-flask-secret-key')

# --- 設定項目 (環境変数から読み込む) ---
# Google Cloudの認証情報ファイルへのパス (RenderのSecret Filesで設定)
# Google OAuth 2.0 クライアント情報
CLIENT_SECRETS_FILE = 'client_secrets.json'
# スプレッドシートを操作するためのサービスアカウント認証情報 (RenderのSecret Filesで設定)
SERVICE_ACCOUNT_FILE = 'credentials.json'
# Google APIが要求する権限の範囲
SCOPES = ['https://www.googleapis.com/auth/userinfo.profile', 'https://www.googleapis.com/auth/userinfo.email', 'openid']
# Googleスプレッドシートの名前 (環境変数で設定)
SPREADSHEET_NAME = os.environ.get('SPREADSHEET_NAME', '研究室出欠記録')
# JWTトークンの暗号化に使う秘密鍵 (環境変数で設定)
JWT_SECRET = os.environ.get('JWT_SECRET', 'default-jwt-secret-key')


# ★★★ IPアドレス制限の設定 ★★★
# ここに許可したいIPアドレスの「前方部分」をリストで指定します。
# このIPからのアクセスはGPS認証が免除されます。
ALLOWED_IP_PREFIXES = [
    '127.0.0.1', # ローカル開発用
     '133.7.7.240', # 例: 大学のネットワーク1
    # '203.0.113.',   # 例: 大学のネットワーク2
]

# ★★★ GPS設定 ★★★
# 教室の緯度・経度を設定
CLASSROOM_LAT = 36.0760254  # 例: 福井大学の緯度
CLASSROOM_LON = 136.2129435 # 例: 福井大学の経度
# 判定を許可する半径 (メートル)
MAX_DISTANCE_METERS = 100 

# -----------------------------------------------------------------------------
# ルーティング
# -----------------------------------------------------------------------------

# --- メインページ (QRコード表示) ---
@app.route('/')
def index():
    # 'qr_display' モードで index.html を表示
    return render_template('index.html', mode='qr_display')

# --- 動的なQRコード画像生成 ---
@app.route('/qr_image.png')
def qr_image():
    try:
        jst = datetime.timezone(datetime.timedelta(hours=9))
        now = datetime.datetime.now(jst)
        # ★★★ 有効期限を10分に延長 ★★★
        expiration_time = now + datetime.timedelta(minutes=10)
        
        payload = {'exp': expiration_time.timestamp(), 'iat': now.timestamp()}
        token = jwt.encode(payload, JWT_SECRET, algorithm='HS256')
        
        qr_url = url_for('attend', token=token, _external=True)

        qr_img = qrcode.make(qr_url)
        img_io = io.BytesIO()
        qr_img.save(img_io, 'PNG')
        img_io.seek(0)
        return send_file(img_io, mimetype='image/png')
    except Exception as e:
        print(f"Error generating QR code: {e}")
        return "Error", 500

# --- 打刻処理の開始 (IP/GPS分岐) ---
@app.route('/attend')
def attend():
    token = request.args.get('token')
    if not token:
        return render_template('index.html', mode='error', message="トークンがありません。"), 400

    try:
        jwt.decode(token, JWT_SECRET, algorithms=['HS256'])
        
        client_ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        is_ip_allowed = any(client_ip.startswith(prefix) for prefix in ALLOWED_IP_PREFIXES)

        if is_ip_allowed:
            print(f"IP address {client_ip} is allowed. Skipping GPS check.")
            flow = Flow.from_client_secrets_file(
                CLIENT_SECRETS_FILE, scopes=SCOPES, redirect_uri=url_for('callback', _external=True)
            )
            authorization_url, state = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true',
                prompt='consent select_account'
            )
            session['state'] = state
            # ★★★ v2と同様のサーバーサイドリダイレクトに変更 ★★★
            return redirect(authorization_url)
        else:
            print(f"IP address {client_ip} is not allowed. Proceeding to GPS check.")
            # ★★★ GPS確認モードで表示 ★★★
            return render_template('index.html', mode='gps_check', token=token)

    except jwt.ExpiredSignatureError:
        return render_template('index.html', mode='error', message="QRコードの有効期限が切れています。ページを更新して再試行してください。"), 403
    except jwt.InvalidTokenError:
        return render_template('index.html', mode='error', message="無効なQRコードです。"), 403
    except Exception as e:
        print(f"Attend error: {e}")
        return render_template('index.html', mode='error', message="サーバーエラーが発生しました。"), 500

# --- 位置情報を検証し、Google認証へ進むAPI ---
@app.route('/verify_location', methods=['POST'])
def verify_location():
    data = request.get_json()
    token = data.get('token')
    lat = data.get('latitude')
    lon = data.get('longitude')

    if not all([token, lat, lon]):
        return jsonify({'success': False, 'message': 'データが不足しています。'})

    try:
        jwt.decode(token, JWT_SECRET, algorithms=['HS256'])
        
        user_location = (lat, lon)
        classroom_location = (CLASSROOM_LAT, CLASSROOM_LON)
        distance = haversine(user_location, classroom_location, unit=Unit.METERS)
        
        if distance <= MAX_DISTANCE_METERS:
            flow = Flow.from_client_secrets_file(
                CLIENT_SECRETS_FILE, scopes=SCOPES, redirect_uri=url_for('callback', _external=True)
            )
            authorization_url, state = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true',
                prompt='consent select_account'
            )
            session['state'] = state
            return jsonify({'success': True, 'redirect_url': authorization_url})
        else:
            return jsonify({'success': False, 'message': f'教室から {int(distance)}m 離れています。教室に入ってから再試行してください。'})

    except (jwt.ExpiredSignatureError, jwt.InvalidTokenError):
        return jsonify({'success': False, 'message': 'QRコードの有効期限が切れました。最初のページに戻って更新してください。'})
    except Exception as e:
        print(f"Verify location error: {e}")
        return jsonify({'success': False, 'message': 'サーバーでエラーが発生しました。'})

# --- Google認証後のコールバック処理 ---
@app.route('/callback')
def callback():
    try:
        flow = Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE, scopes=SCOPES, state=session['state'],
            redirect_uri=url_for('callback', _external=True)
        )
        flow.fetch_token(authorization_response=request.url)
        credentials = flow.credentials
        
        userinfo_service = build('oauth2', 'v2', credentials=credentials)
        user_info = userinfo_service.userinfo().get().execute()
        
        email = user_info.get('email')
        name = user_info.get('name')
        
        jst = datetime.timezone(datetime.timedelta(hours=9))
        timestamp = datetime.datetime.now(jst).strftime('%Y-%m-%d %H:%M:%S')

        gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
        spreadsheet = gc.open(SPREADSHEET_NAME)
        sheet = spreadsheet.sheet1
        
        # 常に2行目に新しいデータを挿入する (1行目はヘッダーと仮定)
        sheet.insert_row([timestamp, email, name], 2)

        MAX_ROWS_WITH_HEADER = 10001 
        
        try:
            current_row_count = sheet.row_count
            
            if current_row_count > MAX_ROWS_WITH_HEADER:
                print(f"Row count {current_row_count} exceeded limit {MAX_ROWS_WITH_HEADER}. Starting backup (A:H)...")
                
                # 1. バックアップ対象のデータを取得 (A列からH列まで)
                start_backup_row_index = MAX_ROWS_WITH_HEADER + 1
                
                # ★★★ 修正点: A列からH列まで取得 ★★★
                old_data = sheet.get(f'A{start_backup_row_index}:H{current_row_count}')
                
                # 2. ヘッダー行を取得 (A列からH列まで)
                # ★★★ 修正点: A1からH1までを明示的に取得 ★★★
                header = sheet.get('A1:H1')[0] 
                
                # 3. 新しいバックアップシートを作成
                backup_sheet_name = f"Backup_{now.strftime('%Y%m%d_%H%M')}"
                
                # ★★★ 修正点: 列数を8に変更 ★★★
                backup_sheet = spreadsheet.add_worksheet(
                    title=backup_sheet_name, 
                    rows=len(old_data) + 1, # 必要な行数+ヘッダー
                    cols=8  # A～H列の8列
                )
                
                # 4. ヘッダーとデータを新シートに書き込む (A列～H列)
                # ★★★ 修正点: 書き込み範囲を明示 (A1からHの最後まで) ★★★
                backup_sheet.update(
                    f'A1:H{len(old_data) + 1}', 
                    [header] + old_data,
                    value_input_option='USER_ENTERED' # 関数の場合、関数としてペースト
                )
                
                # 5. 元シートの古いデータを削除 (バックアップ成功後)
                sheet.delete_rows(start_backup_row_index, current_row_count)
                
                print(f"Backup successful. Moved {len(old_data)} rows (A:H) to sheet '{backup_sheet_name}'.")

        except Exception as e_backup:
            # バックアップ処理でエラーが起きても、打刻自体は成功している
            print(f"Error during sheet backup process: {e_backup}")
        # --- バックアップ処理ここまで ---
        
        # ★★★ 成功モードで表示 ★★★
        return render_template('index.html', mode='success', message="打刻が完了しました！")

    except Exception as e:
        print(f"Callback error: {e}")
        # ★★★ エラーモードで表示 ★★★
        return render_template('index.html', mode='error', message="エラーが発生しました。スプレッドシートへの記録に失敗した可能性があります。"), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
