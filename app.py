# -*- coding: utf-8 -*-
import os
import datetime
import jwt
import gspread
import qrcode
import io
from flask import Flask, render_template, request, redirect, session, url_for, send_file
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

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

# --- メインページ ---
@app.route('/')
def index():
    """
    QRコード画像を表示するメインページ。
    HTML側で /qr_image.png を読み込むことで、動的にQRコードが生成される。
    """
    return render_template('index.html')

# --- 動的なQRコード画像生成 ---
@app.route('/qr_image.png')
def qr_image():
    """
    アクセスされるたびに、その日有効なQRコード画像を動的に生成して返す。
    ファイルとしては保存しない。
    """
    try:
        # 日本時間のタイムゾーンを設定
        jst = datetime.timezone(datetime.timedelta(hours=9))
        
        # 今日の終わりの時刻（有効期限）を計算
        now = datetime.datetime.now(jst)
        expiration_time = now.replace(hour=23, minute=59, second=59, microsecond=0)
        
        # JWTトークンのペイロードを作成
        payload = {
            'exp': expiration_time,
            'iat': now
        }
        
        # JWTトークンを生成
        token = jwt.encode(payload, JWT_SECRET, algorithm='HS256')
        
        # RenderサーバーのURLを自動で取得し、QRコードに含めるURLを生成
        base_url = request.host_url
        qr_url = f"{base_url}attend?token={token}"

        # QRコードを生成
        qr_img = qrcode.make(qr_url)
        
        # 画像をファイルに保存せず、メモリ上のバッファに保存
        img_io = io.BytesIO()
        qr_img.save(img_io, 'PNG')
        img_io.seek(0)
        
        # メモリ上の画像データを、PNG画像としてブラウザに返す
        return send_file(img_io, mimetype='image/png')

    except Exception as e:
        print(f"Error generating QR code: {e}")
        # エラーが発生した場合は、プレースホルダー画像を返すかエラーメッセージを表示
        return "Error generating QR code", 500


# --- 打刻処理の開始 ---
@app.route('/attend')
def attend():
    """
    QRコードからアクセスされた際の最初の受付窓口。
    トークンを検証し、Google認証へリダイレクトする。
    """
    token = request.args.get('token')
    if not token:
        return "エラー: トークンがありません。", 400

    try:
        # トークンの検証
        jwt.decode(token, JWT_SECRET, algorithms=['HS256'])
        
        # Google認証フローを初期化
        flow = Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE,
            scopes=SCOPES,
            redirect_uri=url_for('callback', _external=True)
        )
        authorization_url, state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true'
        )
        session['state'] = state
        return redirect(authorization_url)

    except jwt.ExpiredSignatureError:
        return "エラー: このQRコードの有効期限が切れています。", 403
    except jwt.InvalidTokenError:
        return "エラー: 無効なQRコードです。", 403
    except Exception as e:
        print(f"Attend error: {e}")
        return "サーバーエラーが発生しました。", 500

# --- Google認証後のコールバック処理 ---
@app.route('/callback')
def callback():
    """
    Google認証成功後に呼び出される。
    ユーザー情報を取得し、スプレッドシートに記録する。
    """
    try:
        flow = Flow.from_client_secrets_file(
            CLIENT_SECRETS_FILE,
            scopes=SCOPES,
            state=session['state'],
            redirect_uri=url_for('callback', _external=True)
        )
        flow.fetch_token(authorization_response=request.url)
        
        credentials = flow.credentials
        
        # ユーザー情報を取得
        userinfo_service = build('oauth2', 'v2', credentials=credentials)
        user_info = userinfo_service.userinfo().get().execute()
        
        email = user_info.get('email')
        name = user_info.get('name')
        
        # スプレッドシートに記録
        jst = datetime.timezone(datetime.timedelta(hours=9))
        timestamp = datetime.datetime.now(jst).strftime('%Y-%m-%d %H:%M:%S')

        gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
        sheet = gc.open(SPREADSHEET_NAME).sheet1
        sheet.append_row([timestamp, email, name])
        
        return "<h1>打刻が完了しました！</h1><p>このページを閉じてください。</p>"

    except Exception as e:
        print(f"Callback error: {e}")
        return "エラーが発生しました。スプレッドシートへの記録に失敗した可能性があります。", 500

if __name__ == '__main__':
    # 開発環境で実行する際のポートとデバッグ設定
    app.run(host='0.0.0.0', port=5000, debug=True)
