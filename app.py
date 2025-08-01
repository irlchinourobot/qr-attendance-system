import os
import json
from datetime import datetime, timedelta, timezone

import jwt
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from flask import Flask, redirect, url_for, session, request, render_template, abort, jsonify
import qrcode

# -----------------------------------------------------------------------------
# 初期設定
# -----------------------------------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'default-flask-secret-key')
JWT_SECRET = os.environ.get('JWT_SECRET', 'default-jwt-secret-key')
CRON_SECRET = os.environ.get('CRON_SECRET', 'default-cron-secret-key')

# Google OAuth 2.0 クライアント情報
CLIENT_SECRETS_FILE = 'client_secrets.json'
SCOPES = ['https://www.googleapis.com/auth/userinfo.profile', 'https://www.googleapis.com/auth/userinfo.email', 'openid']

# Googleスプレッドシートの設定
GSPREAD_CREDENTIALS_FILE = 'credentials.json'
SPREADSHEET_NAME = '研究室出欠記録' # ★★★ あなたが作成したスプレッドシートの名前に書き換えてください ★★★

# -----------------------------------------------------------------------------
# QRコード生成ロジック (app.pyに統合)
# -----------------------------------------------------------------------------
def generate_attendance_qr_code():
    """JWTトークンを含んだQRコードを生成し、static/today_qr.pngに保存します。"""
    try:
        # RenderのサーバーURLは動的に取得しないため、環境変数から読み込むか、固定で設定
        # デプロイ後にRenderのURLをここに設定する
        base_url = os.environ.get('RENDER_EXTERNAL_URL', 'http://127.0.0.1:5000')
        qr_code_file_path = os.path.join('static', 'today_qr.png')

        jst = timezone(timedelta(hours=+9), 'JST')
        now = datetime.now(jst)
        expiration_time = now.replace(hour=23, minute=59, second=59, microsecond=0)
        
        payload = {'iat': now, 'exp': expiration_time, 'iss': 'qr_attendance_system'}
        token = jwt.encode(payload, JWT_SECRET, algorithm='HS256')
        url = f"{base_url}/attend?token={token}"
        
        img = qrcode.make(url)
        
        dir_name = os.path.dirname(qr_code_file_path)
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)
            
        img.save(qr_code_file_path)
        print(f"QRコードが正常に生成されました。URL: {url}")
        return True, url
    except Exception as e:
        print(f"QRコードの生成中にエラーが発生しました: {e}")
        return False, str(e)

# -----------------------------------------------------------------------------
# ヘルパー関数 (変更なし)
# -----------------------------------------------------------------------------
def get_spreadsheet():
    try:
        gc = gspread.service_account(filename=GSPREAD_CREDENTIALS_FILE)
        spreadsheet = gc.open(SPREADSHEET_NAME)
        return spreadsheet.sheet1
    except Exception as e:
        print(f"スプレッドシートへのアクセス中にエラーが発生しました: {e}")
        return None

def record_attendance(user_info):
    worksheet = get_spreadsheet()
    if not worksheet: return False
    try:
        jst = timezone(timedelta(hours=+9), 'JST')
        timestamp = datetime.now(jst).strftime('%Y-%m-%d %H:%M:%S')
        email = user_info.get('email', 'N/A')
        name = user_info.get('name', 'N/A')
        worksheet.append_row([timestamp, email, name])
        print(f"記録完了: {timestamp}, {email}, {name}")
        return True
    except Exception as e:
        print(f"出席の記録中にエラーが発生しました: {e}")
        return False

# -----------------------------------------------------------------------------
# Flaskルーティング
# -----------------------------------------------------------------------------
@app.route('/')
def index():
    import time
    qr_url = url_for('static', filename='today_qr.png', t=time.time())
    return render_template('index.html', qr_code_url=qr_url)

# ★★★ 新しいルート：Cronジョブ用のQR生成トリガー ★★★
@app.route('/cron/generate-qr')
def trigger_qr_generation():
    # クエリパラメータで秘密のキーをチェック
    if request.args.get('secret') != CRON_SECRET:
        return "Unauthorized", 401
    
    success, message = generate_attendance_qr_code()
    if success:
        return jsonify({"status": "success", "url": message}), 200
    else:
        return jsonify({"status": "error", "message": message}), 500

# attend, login, callback, process_attendance, success, logoutルートは変更なし
@app.route('/attend')
def attend():
    token = request.args.get('token')
    if not token: return "エラー: トークンがありません。", 400
    try:
        jwt.decode(token, JWT_SECRET, algorithms=['HS256'])
        session['attendance_pending'] = True
        return redirect(url_for('login'))
    except jwt.ExpiredSignatureError: return "エラー: このQRコードの有効期限が切れています。", 403
    except jwt.InvalidTokenError: return "エラー: 無効なQRコードです。", 403

@app.route('/login')
def login():
    flow = Flow.from_client_secrets_file(CLIENT_SECRETS_FILE, scopes=SCOPES, redirect_uri=url_for('callback', _external=True))
    authorization_url, state = flow.authorization_url()
    session['state'] = state
    return redirect(authorization_url)

@app.route('/callback')
def callback():
    if 'state' not in session or session['state'] != request.args.get('state'): abort(500)
    flow = Flow.from_client_secrets_file(CLIENT_SECRETS_FILE, scopes=SCOPES, state=session['state'], redirect_uri=url_for('callback', _external=True))
    flow.fetch_token(authorization_response=request.url)
    credentials = flow.credentials
    session['credentials'] = {'token': credentials.token, 'refresh_token': credentials.refresh_token, 'token_uri': credentials.token_uri, 'client_id': credentials.client_id, 'client_secret': credentials.client_secret, 'scopes': credentials.scopes}
    if session.get('attendance_pending'):
        session.pop('attendance_pending', None)
        return redirect(url_for('process_attendance'))
    return redirect(url_for('success'))

@app.route('/process_attendance')
def process_attendance():
    if 'credentials' not in session: return redirect(url_for('login'))
    try:
        creds = Credentials(**session['credentials'])
        user_info_service = build('oauth2', 'v2', credentials=creds)
        user_info = user_info_service.userinfo().get().execute()
        if record_attendance(user_info):
            return redirect(url_for('success'))
        else:
            return "エラー: スプレッドシートへの記録に失敗しました。", 500
    except Exception as e:
        print(f"出席処理中にエラーが発生しました: {e}")
        return "エラー: 処理中に問題が発生しました。", 500

@app.route('/success')
def success():
    return """<html><head><title>成功</title></head><body><h1>打刻が完了しました！</h1></body></html>"""

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

# -----------------------------------------------------------------------------
# 実行 (gunicornが使うため、この部分はローカル実行時のみ使われる)
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    # ローカルテスト用に、初回QRコードを生成
    generate_attendance_qr_code()
    app.run(debug=True, host='0.0.0.0', port=5000)

