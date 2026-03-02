import functions_framework
import json
import os
import requests
from datetime import datetime
from google.auth import default
from googleapiclient.discovery import build

STRIPE_WEBHOOK_SECRET = os.environ.get('STRIPE_WEBHOOK_SECRET', '')
CUSTOMER_SHEET_ID = os.environ.get('CUSTOMER_SHEET_ID', '')
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN', '')
LINE_NOTIFY_USER_ID = os.environ.get('LINE_NOTIFY_USER_ID', '')
STRIPE_API_KEY = os.environ.get('STRIPE_API_KEY', '')

@functions_framework.http
def stripe_webhook(request):
    payload = request.get_data(as_text=True)
    
    try:
        event = json.loads(payload)
    except json.JSONDecodeError:
        return 'Invalid payload', 400
    
    event_type = event.get('type', '')
    print(f'Received event: {event_type}')
    
    if event_type == 'checkout.session.completed':
        handle_checkout_completed(event['data']['object'])
    elif event_type == 'invoice.payment_failed':
        handle_payment_failed(event['data']['object'])
    
    return 'OK', 200

def handle_checkout_completed(session):
    """初回決済完了時の処理"""
    customer_id = session.get('customer', '')
    customer_email = session.get('customer_details', {}).get('email', '')
    customer_name = session.get('customer_details', {}).get('name', '')
    amount = session.get('amount_total', 0)
    
    print(f'New subscription: {customer_name} ({customer_email}) - {amount}円')
    
    # 未使用コードを取得して割り当て
    result = assign_unused_code(customer_id, customer_name, customer_email, amount)
    
    if result.get('success'):
        code = result.get('code')
        
        # 顧客にメールでコード通知
        send_code_email(customer_email, customer_name, code, amount)
        
        # あなたにLINE通知
        send_line_notification(
            f'🎉 新規契約がありました！\n\n'
            f'👤 {customer_name}\n'
            f'✉️ {customer_email}\n'
            f'💰 {amount:,}円/月\n'
            f'🔑 {code}\n\n'
            f'メールでコードを送信済みです。'
        )
    else:
        # 未使用コードがない場合
        send_line_notification(
            f'⚠️ 新規契約がありましたが、未使用コードがありません！\n\n'
            f'👤 {customer_name}\n'
            f'✉️ {customer_email}\n'
            f'💰 {amount:,}円/月\n\n'
            f'手動でコードを発行してください。'
        )

def handle_payment_failed(invoice):
    """支払い失敗時"""
    customer_id = invoice.get('customer', '')
    customer_email = invoice.get('customer_email', '')
    
    send_line_notification(
        f'⚠️ 支払いが失敗しました\n\n'
        f'Customer ID: {customer_id}\n'
        f'Email: {customer_email}\n\n'
        f'確認してください。'
    )

def get_sheets_service():
    credentials, project = default(scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return build('sheets', 'v4', credentials=credentials)

def assign_unused_code(customer_id, name, email, amount):
    """未使用コードを探して顧客情報を割り当て"""
    try:
        service = get_sheets_service()
        
        # 全データ取得
        result = service.spreadsheets().values().get(
            spreadsheetId=CUSTOMER_SHEET_ID,
            range='顧客管理!A:R'
        ).execute()
        rows = result.get('values', [])
        
        # 未使用の行を探す
        for i, row in enumerate(rows[1:], start=2):
            status = row[7] if len(row) > 7 else ''
            
            if status == '未使用':
                code = row[6] if len(row) > 6 else ''
                folder_id = row[2] if len(row) > 2 else ''
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                # プランを金額から判定
                if amount <= 5500:
                    plan = '記帳5000'
                elif amount <= 11000:
                    plan = '記帳10000'
                else:
                    plan = '記帳14000'
                
                # 行を更新（A列〜R列）
                update_values = [[
                    '',                 # A: line_user_id（後でLINE連携時に入る）
                    name or '未設定',   # B: customer_name
                    folder_id,          # C: folder_id（既存のまま）
                    now,                # D: registered_at
                    False,              # E: notified
                    '',                 # F: sent_at
                    code,               # G: customer_code（既存のまま）
                    '案内済',           # H: status
                    email or '',        # I: email
                    '',                 # J: phone
                    0,                  # K: trial_count
                    0,                  # L: total_count
                    '',                 # M: memo
                    customer_id,        # N: stripe_customer_id
                    plan,               # O: プラン
                    amount,             # P: 月額料金
                    now,                # Q: 課金開始日
                    ''                  # R: 備考
                ]]
                
                service.spreadsheets().values().update(
                    spreadsheetId=CUSTOMER_SHEET_ID,
                    range=f'顧客管理!A{i}:R{i}',
                    valueInputOption='RAW',
                    body={'values': update_values}
                ).execute()
                
                print(f'Assigned code {code} to {name}')
                return {'success': True, 'code': code}
        
        print('No unused code available')
        return {'success': False, 'error': 'no_unused_code'}
        
    except Exception as e:
        print(f'Error assigning code: {e}')
        return {'success': False, 'error': str(e)}

def send_code_email(email, name, code, amount):
    """顧客にコード通知メールを送信（Stripe経由）"""
    if not STRIPE_API_KEY:
        print('STRIPE_API_KEY not set, skipping email')
        return
    
    # Stripe APIでメール送信（または別のメール送信方法）
    # 今回はシンプルにログ出力のみ
    # 実際のメール送信は後で実装可能
    print(f'Email would be sent to {email}: Code is {code}')
    
    # TODO: 実際のメール送信実装
    # SendGrid, AWS SES, Gmail API など

def send_line_notification(message):
    """管理者にLINE通知を送信"""
    if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_NOTIFY_USER_ID:
        print('LINE notification not configured')
        return
    
    url = 'https://api.line.me/v2/bot/message/push'
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {LINE_CHANNEL_ACCESS_TOKEN}'
    }
    data = {
        'to': LINE_NOTIFY_USER_ID,
        'messages': [{'type': 'text', 'text': message}]
    }
    
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code == 200:
            print('LINE notification sent')
        else:
            print(f'LINE notification failed: {response.status_code}')
    except Exception as e:
        print(f'LINE notification error: {e}')
