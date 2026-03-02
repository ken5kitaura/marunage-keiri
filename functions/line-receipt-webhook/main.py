import json
import hashlib
import hmac
import base64
import os
import re
import requests
from datetime import datetime
from flask import Flask, request
from google.auth import default
from googleapiclient.discovery import build

app = Flask(__name__)

# MK（まるなげ経理）チャネル
MK_LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET', '')
MK_LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN', '')
# KZ（絆パートナーズ）チャネル
KZ_LINE_CHANNEL_SECRET = os.environ.get('KZ_LINE_CHANNEL_SECRET', '')
KZ_LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('KZ_LINE_CHANNEL_ACCESS_TOKEN', '')

# 後方互換のためのエイリアス（チャネル判定前のデフォルト）
LINE_CHANNEL_SECRET = MK_LINE_CHANNEL_SECRET
LINE_CHANNEL_ACCESS_TOKEN = MK_LINE_CHANNEL_ACCESS_TOKEN

CUSTOMER_SHEET_ID = os.environ.get('CUSTOMER_SHEET_ID', '')
KZ_CUSTOMER_SHEET_ID = os.environ.get('KZ_CUSTOMER_SHEET_ID', '')
GAS_UPLOAD_URL = os.environ.get('GAS_UPLOAD_URL', '')
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')

# 管理者LINE ID
ADMIN_USER_ID = 'U6980f7c583babed09518d986f704e959'

# チャネル設定マップ: destination (Bot User ID) → チャネル情報
# ※ 初回デプロイ後にログからBot User IDを確認して設定する
MK_BOT_USER_ID = os.environ.get('MK_BOT_USER_ID', '')
KZ_BOT_USER_ID = os.environ.get('KZ_BOT_USER_ID', '')

CHANNELS = {
    'MK': {
        'secret': MK_LINE_CHANNEL_SECRET,
        'access_token': MK_LINE_CHANNEL_ACCESS_TOKEN,
        'name': 'まるなげ経理',
    },
    'KZ': {
        'secret': KZ_LINE_CHANNEL_SECRET,
        'access_token': KZ_LINE_CHANNEL_ACCESS_TOKEN,
        'name': '絆パートナーズ経理',
    },
}

def resolve_channel(body_text):
    """
    Webhookリクエストのbodyからチャネルを判定する。
    destinationフィールド（Bot User ID）で判定し、
    不明な場合は署名検証で特定を試みる。
    戻り値: チャネルキー ('MK' or 'KZ')
    """
    try:
        data = json.loads(body_text)
        destination = data.get('destination', '')
    except Exception:
        destination = ''

    if destination and MK_BOT_USER_ID and destination == MK_BOT_USER_ID:
        return 'MK'
    if destination and KZ_BOT_USER_ID and destination == KZ_BOT_USER_ID:
        return 'KZ'

    # Bot User IDが未設定 or 不明な場合、destinationをログに出力（初回設定用）
    if destination:
        print(f'[channel] Unknown destination: {destination} — Bot User IDの環境変数を設定してください')

    # デフォルトはMK
    return 'MK'

def get_channel_config(channel_key):
    """チャネルキーからチャネル設定を取得"""
    return CHANNELS.get(channel_key, CHANNELS['MK'])

@app.route('/', methods=['GET', 'POST'])
def line_webhook():
    if request.method == 'GET':
        return 'OK', 200
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)

    # チャネル判定（destination から Bot User ID で特定）
    channel_key = resolve_channel(body)
    print(f'[webhook] channel={channel_key}')

    if not verify_signature(body, signature, channel_key):
        return 'Invalid signature', 403
    try:
        events = json.loads(body).get('events', [])
        for event in events:
            if event['type'] == 'message':
                msg_type = event['message']['type']
                if msg_type == 'image':
                    handle_image_message(event, channel_key)
                elif msg_type == 'file':
                    handle_file_message(event, channel_key)
                elif msg_type == 'text':
                    handle_text_message(event, channel_key)
            elif event['type'] == 'follow':
                handle_follow_event(event, channel_key)
    except Exception as e:
        print(f'Error processing event: {e}')
    return 'OK', 200

def verify_signature(body, signature, channel_key=None):
    """
    署名検証。channel_keyが指定されていればそのチャネルのシークレットで検証。
    channel_keyがNoneの場合は全チャネルのシークレットで試行。
    """
    if channel_key:
        secret = CHANNELS.get(channel_key, {}).get('secret', '')
        if not secret:
            return True
        hash = hmac.new(secret.encode('utf-8'), body.encode('utf-8'), hashlib.sha256).digest()
        expected = base64.b64encode(hash).decode('utf-8')
        return hmac.compare_digest(signature, expected)

    # channel_key未定の場合、全チャネルのシークレットで検証を試行
    for key, ch in CHANNELS.items():
        secret = ch.get('secret', '')
        if not secret:
            continue
        hash = hmac.new(secret.encode('utf-8'), body.encode('utf-8'), hashlib.sha256).digest()
        expected = base64.b64encode(hash).decode('utf-8')
        if hmac.compare_digest(signature, expected):
            return True

    return False

def get_sheets_service():
    credentials, project = default(scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return build('sheets', 'v4', credentials=credentials)

def get_sheet_id_for_channel(channel_key):
    """チャネルに応じた顧客管理シートIDを返す"""
    if channel_key == 'KZ':
        return KZ_CUSTOMER_SHEET_ID
    return CUSTOMER_SHEET_ID

def get_status_col_for_channel(channel_key):
    """チャネルに応じたステータス列インデックスを返す（0始まり）"""
    if channel_key == 'KZ':
        return 5  # F列 (KZ: 6番目, 0始まりで5)
    return 7  # H列 (MK: 8番目, 0始まりで7)

def get_customer_info(user_id, channel_key='MK'):
    """顧客情報を取得（folder_id, customer_name, status）"""
    try:
        service = get_sheets_service()
        sheet_id = get_sheet_id_for_channel(channel_key)
        status_col = get_status_col_for_channel(channel_key)
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range='顧客管理!A:M'
        ).execute()
        rows = result.get('values', [])
        for row in rows[1:]:
            if len(row) >= 1 and row[0] == user_id:
                folder_id = row[2] if len(row) > 2 else ''
                customer_name = row[1] if len(row) > 1 else ''
                status = row[status_col] if len(row) > status_col else ''
                return {
                    'folder_id': folder_id,
                    'customer_name': customer_name,
                    'status': status,
                    'exists': True
                }
        return {'exists': False}
    except Exception as e:
        print(f'Error getting customer info: {e}')
        return {'exists': False}

def register_new_user(user_id, channel_key='MK'):
    """新規ユーザーを「お試し」として登録"""
    try:
        service = get_sheets_service()
        sheet_id = get_sheet_id_for_channel(channel_key)
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        if channel_key == 'KZ':
            # KZ: A=LINE ID, B=顧客名, C=フォルダID, D=登録日, E=コード, F=ステータス
            values = [[user_id, '未登録', '', now, '', 'お試し']]
            append_range = '顧客管理!A:F'
        else:
            # MK: A=LINE ID, B=顧客名, C=フォルダID, D=登録日, E=?, F=?, G=コード, H=ステータス
            values = [[user_id, '未登録', '', now, False, '', '', 'お試し']]
            append_range = '顧客管理!A:H'
        service.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range=append_range,
            valueInputOption='RAW',
            body={'values': values}
        ).execute()
        print(f'New user registered [{channel_key}]: {user_id}')
        return True
    except Exception as e:
        print(f'Error registering user [{channel_key}]: {e}')
        return False

def update_trial_count(user_id, channel_key='MK'):
    """お試し送信回数をカウントアップ"""
    try:
        service = get_sheets_service()
        sheet_id = get_sheet_id_for_channel(channel_key)
        # KZはカウント列が異なる可能性があるが、同じロジックを使用
        # MK: K列(index 10), KZ: 同様にK列を使用（なければスキップ）
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range='顧客管理!A:L'
        ).execute()
        rows = result.get('values', [])
        for i, row in enumerate(rows[1:], start=2):
            if len(row) >= 1 and row[0] == user_id:
                current_count = int(row[10]) if len(row) > 10 and row[10] else 0
                service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f'顧客管理!K{i}',
                    valueInputOption='RAW',
                    body={'values': [[current_count + 1]]}
                ).execute()
                return current_count + 1
        return 0
    except Exception as e:
        print(f'Error updating trial count [{channel_key}]: {e}')
        return 0

def handle_follow_event(event, channel_key='MK'):
    """友達追加イベント - 登録はせず、ウェルカムメッセージのみ送信"""
    user_id = event['source'].get('userId', '')
    if user_id:
        config = get_channel_config(channel_key)
        service_name = config['name']
        # 登録はしない（画像を送った時点で登録する）
        reply_message(event['replyToken'],
            f'{service_name}へようこそ！📸\n\n'
            '領収書の写真を送るだけで\n'
            '記帳業務をすべてお任せいただけます。\n\n'
            'まずは試しに1枚送ってみてください！',
            channel_key)

def handle_file_message(event, channel_key='MK'):
    """PDFファイルなどの処理"""
    message_id = event['message']['id']
    file_name = event['message'].get('fileName', '')
    user_id = event['source'].get('userId', 'unknown')

    # PDFのみ対応
    if not file_name.lower().endswith('.pdf'):
        reply_message(event['replyToken'],
            '⚠️ 対応していないファイル形式です。\n\n'
            '画像（JPG, PNG）またはPDFを\n'
            'お送りください。',
            channel_key)
        return

    # 顧客情報を取得
    customer_info = get_customer_info(user_id, channel_key)

    # 未登録ユーザーは「お試し」として登録
    if not customer_info.get('exists'):
        register_new_user(user_id, channel_key)
        customer_info = {'status': 'お試し', 'folder_id': '', 'customer_name': ''}

    status = customer_info.get('status', 'お試し')
    folder_id = customer_info.get('folder_id', '')

    # ファイルをダウンロード
    file_content = download_content_from_line(message_id, channel_key)
    if not file_content:
        reply_message(event['replyToken'], '❌ ファイルの取得に失敗しました。\nもう一度お試しください。', channel_key)
        return

    # Gemini で分類
    classification = classify_document_with_gemini(file_content, 'application/pdf')

    # 分類結果に応じて処理
    process_classified_document(event, classification, file_content, folder_id, status, user_id, file_name, channel_key)

def handle_image_message(event, channel_key='MK'):
    message_id = event['message']['id']
    user_id = event['source'].get('userId', 'unknown')

    # 顧客情報を取得
    customer_info = get_customer_info(user_id, channel_key)

    # 未登録ユーザーは「お試し」として登録
    if not customer_info.get('exists'):
        register_new_user(user_id, channel_key)
        customer_info = {'status': 'お試し', 'folder_id': '', 'customer_name': ''}

    status = customer_info.get('status', 'お試し')
    folder_id = customer_info.get('folder_id', '')

    # 画像をダウンロード
    image_content = download_content_from_line(message_id, channel_key)
    if not image_content:
        reply_message(event['replyToken'], '❌ 画像の取得に失敗しました。\nもう一度お試しください。', channel_key)
        return

    # Gemini で分類 + OCR
    classification = classify_document_with_gemini(image_content, 'image/jpeg')

    # 分類結果に応じて処理
    filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{message_id}.jpg"
    process_classified_document(event, classification, image_content, folder_id, status, user_id, filename, channel_key)

def process_classified_document(event, classification, content, folder_id, status, user_id, filename, channel_key='MK'):
    """分類結果に応じてドキュメントを処理"""

    category = classification.get('category', 'unknown')

    # ==== クレカ売上票 ====
    if category == 'credit_slip':
        reply_message(event['replyToken'],
            '⚠️ クレジットカード売上票です\n\n'
            '二重計上を防ぐため、これは保存しません。\n'
            'レシート本体をお送りください📸',
            channel_key)
        return

    # ==== 不明・その他 ====
    if category == 'unknown':
        reply_message(event['replyToken'],
            '⚠️ 認識できませんでした\n\n'
            '以下のいずれかをお送りください：\n'
            '・レシート/領収書\n'
            '・通帳',
            channel_key)
        return

    # ==== 通帳 ====
    if category == 'passbook':
        if folder_id:
            # 通帳フォルダに保存（フォルダ構成: 親/通帳/）
            passbook_folder_id = get_or_create_subfolder(folder_id, '通帳')
            if passbook_folder_id:
                upload_via_gas(content, filename, passbook_folder_id)

        if status == '契約済':
            reply_message(event['replyToken'],
                '✅ 通帳を受け取りました\n\n'
                '担当者が確認いたします。',
                channel_key)
        else:
            update_trial_count(user_id, channel_key)
            reply_message(event['replyToken'],
                '✅ 通帳を受け取りました\n\n'
                '引き続き、領収書や通帳を\nお送りください📸',
                channel_key)
        return

    # ==== レシート/領収書 ====
    if category == 'receipt':
        if folder_id:
            # レシートフォルダに保存（folder_idは既に「領収書」フォルダ）
            upload_via_gas(content, filename, folder_id)

        date_str = classification.get('extracted_data', {}).get('date', '')
        store_name = classification.get('extracted_data', {}).get('store_name', '')
        amount = classification.get('extracted_data', {}).get('amount', '')

        if status == '契約済':
            reply_message(event['replyToken'],
                '✅ 領収書を受け取りました\n\n'
                '担当者が確認のうえ記帳いたします。\n'
                '引き続きよろしくお願いいたします。',
                channel_key)
        else:
            update_trial_count(user_id, channel_key)

            result_lines = ['✅ 領収書を受け取りました\n']
            if date_str:
                result_lines.append(f'📅 {date_str}')
            if store_name:
                result_lines.append(f'🏪 {store_name}')
            if amount:
                result_lines.append(f'💰 {amount}円')

            result_lines.append('\nご利用の流れはこれだけです。')
            result_lines.append('LINEで領収書を送るだけ。')
            result_lines.append('\nご契約後は担当者がすべて確認し、')
            result_lines.append('会計ソフト用のデータとしてお届けします。')
            result_lines.append('\n▼ お申込みはこちら')
            result_lines.append('全プラン最終チェック付き')
            result_lines.append('')
            result_lines.append('📦 ミニマル（月30行まで）')
            result_lines.append('月5,000円')
            result_lines.append('https://buy.stripe.com/6oUeVdeuA2W0abA7aO3Ru00')
            result_lines.append('')
            result_lines.append('⭐ ライト（月100行まで）← おすすめ')
            result_lines.append('月10,000円')
            result_lines.append('https://buy.stripe.com/dRm3cvaek7cgerQ9iW3Ru01')
            result_lines.append('')
            result_lines.append('📦 スタンダード（月200行まで）')
            result_lines.append('月14,000円')
            result_lines.append('https://buy.stripe.com/28EaEXgCI2W06ZoeDg3Ru02')
            result_lines.append('')
            result_lines.append('※超過分は20円/行')

            reply_message(event['replyToken'], '\n'.join(result_lines), channel_key)
        return

def classify_document_with_gemini(content, mime_type):
    """Gemini で書類を分類 + データ抽出"""
    if not GEMINI_API_KEY:
        print('GEMINI_API_KEY not set')
        return {'category': 'unknown', 'error': 'config'}
    
    try:
        url = f'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}'
        
        content_base64 = base64.b64encode(content).decode('utf-8')
        
        prompt = '''この画像を分類してください。

【分類カテゴリ】
1. receipt（レシート/領収書）
2. passbook（通帳）
3. credit_slip（クレジットカード売上票）
4. unknown（上記以外/判別不能）

【判別ポイント】

■ passbook（通帳）の特徴：
- 「普通預金」「当座預金」などのヘッダー
- 年月日/お支払金額/お預り金額/差引残高 の列構成
- 銀行の通帳形式

■ credit_slip（クレカ売上票）の特徴：
- 「クレジット売上票」「カード売上票」「CREDIT」の文言
- カード番号の一部（****1234）
- 承認番号、カード会社名（VISA, Mastercard, JCB等）
- 「お客様控え」の文言
- レシートと一緒に出てくる決済控え

■ receipt（レシート/領収書）の特徴：
- 店名、日付、商品明細、合計金額
- 「領収書」「レシート」の文言
- credit_slipに該当しないもの

【回答形式】
JSON形式で回答してください：

{
  "category": "receipt" | "passbook" | "credit_slip" | "unknown",
  "confidence": 0.0〜1.0,
  "reason": "判定理由を簡潔に",
  "extracted_data": {
    // categoryに応じたデータ
  }
}

■ receipt の場合の extracted_data:
{
  "date": "2026年2月15日",
  "store_name": "セブンイレブン",
  "amount": "1234"
}

■ passbook の場合の extracted_data:
{
  "bank_name": "三菱UFJ銀行",
  "date_range": "2026/01/05〜01/27",
  "latest_balance": "451277"
}

■ credit_slip / unknown の場合:
extracted_data は空オブジェクト {}
'''

        payload = {
            'contents': [{
                'parts': [
                    {'text': prompt},
                    {
                        'inline_data': {
                            'mime_type': mime_type,
                            'data': content_base64
                        }
                    }
                ]
            }],
            'generationConfig': {
                'temperature': 0.1,
                'maxOutputTokens': 1000
            }
        }
        
        response = requests.post(url, json=payload, timeout=30)
        
        if response.status_code != 200:
            print(f'Gemini API error: {response.status_code} {response.text}')
            return {'category': 'unknown', 'error': 'api'}
        
        result = response.json()
        text = result.get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
        
        # JSONを抽出
        json_match = re.search(r'\{[\s\S]*\}', text)
        if json_match:
            classification = json.loads(json_match.group())
            print(f'Classification result: {classification}')
            return classification
        
        print(f'Failed to parse Gemini response: {text}')
        return {'category': 'unknown', 'error': 'parse'}
        
    except Exception as e:
        print(f'Gemini classification error: {e}')
        return {'category': 'unknown', 'error': str(e)}

def get_or_create_subfolder(parent_folder_id, subfolder_name):
    """親フォルダ内にサブフォルダを取得または作成"""
    try:
        credentials, project = default(scopes=['https://www.googleapis.com/auth/drive'])
        service = build('drive', 'v3', credentials=credentials)
        
        # 親フォルダの親（MKxxxフォルダ）を取得
        parent_file = service.files().get(fileId=parent_folder_id, fields='parents').execute()
        mk_folder_id = parent_file.get('parents', [None])[0]
        
        if not mk_folder_id:
            print('Could not find MKxxx folder')
            return None
        
        # サブフォルダを検索
        query = f"'{mk_folder_id}' in parents and name='{subfolder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        results = service.files().list(q=query, fields='files(id, name)').execute()
        files = results.get('files', [])
        
        if files:
            return files[0]['id']
        
        # なければ作成
        folder_metadata = {
            'name': subfolder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [mk_folder_id]
        }
        folder = service.files().create(body=folder_metadata, fields='id').execute()
        print(f'Created subfolder: {subfolder_name} ({folder["id"]})')
        return folder['id']
        
    except Exception as e:
        print(f'Error getting/creating subfolder: {e}')
        return None

# ============================================================
# 顧客コードの正規化・検証
# ============================================================

def get_sheet_id_for_code(customer_code):
    """コードのプレフィックスに応じて顧客管理シートIDを返す"""
    if customer_code and customer_code.startswith('KZ'):
        return KZ_CUSTOMER_SHEET_ID
    return CUSTOMER_SHEET_ID

def get_code_col_for_prefix(customer_code):
    """コードのプレフィックスに応じて顧客コード列番号を返す（0始まり）"""
    if customer_code and customer_code.startswith('KZ'):
        return 4  # E列 (KZ: 5番目, 0始まりで4)
    return 6  # G列 (MK: 7番目, 0始まりで6)

def get_status_col_for_prefix(customer_code):
    """コードのプレフィックスに応じてステータス列番号を返す（1始まり、Sheets API用）"""
    if customer_code and customer_code.startswith('KZ'):
        return 'F'  # KZ: F列
    return 'H'  # MK: H列

def get_line_col_for_prefix(customer_code):
    """コードのプレフィックスに応じてLINE ID列を返す（1始まり、Sheets API用）"""
    return 'A'  # 両方ともA列

def normalize_customer_code(text):
    """
    顧客コードを正規化
    - 全角→半角変換
    - 大文字変換
    - 前後スペース除去
    """
    if not text:
        return None
    
    # 前後スペース除去
    text = text.strip()
    
    # 全角→半角変換（英数字）
    normalized = ''
    for char in text:
        code = ord(char)
        # 全角英字 A-Z (0xFF21-0xFF3A) → 半角
        if 0xFF21 <= code <= 0xFF3A:
            normalized += chr(code - 0xFF21 + ord('A'))
        # 全角英字 a-z (0xFF41-0xFF5A) → 半角
        elif 0xFF41 <= code <= 0xFF5A:
            normalized += chr(code - 0xFF41 + ord('a'))
        # 全角数字 0-9 (0xFF10-0xFF19) → 半角
        elif 0xFF10 <= code <= 0xFF19:
            normalized += chr(code - 0xFF10 + ord('0'))
        else:
            normalized += char
    
    # 大文字変換
    normalized = normalized.upper()
    
    return normalized

def is_valid_customer_code_format(code):
    """
    顧客コードの形式を検証
    有効: MK + 3桁の数字 (MK001, MK123 など)
    有効: KZ + 3桁の数字 (KZ001, KZ123 など)
    """
    if not code:
        return False
    
    # 正規表現: MKまたはKZ + 3桁数字
    pattern = r'^(MK|KZ)\d{3}$'
    return bool(re.match(pattern, code))

def customer_code_exists(customer_code):
    """
    顧客コードが顧客管理シートに存在するか確認
    戻り値: {'exists': True/False, 'row_index': int, 'row_data': list}
    """
    try:
        service = get_sheets_service()
        sheet_id = get_sheet_id_for_code(customer_code)
        code_col = get_code_col_for_prefix(customer_code)
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range='顧客管理!A:M'
        ).execute()
        rows = result.get('values', [])
        
        for i, row in enumerate(rows[1:], start=2):  # 1行目はヘッダー
            row_code = row[code_col] if len(row) > code_col else ''
            if row_code.upper() == customer_code:
                return {
                    'exists': True,
                    'row_index': i,
                    'row_data': row
                }
        
        return {'exists': False}
        
    except Exception as e:
        print(f'Error checking customer code: {e}')
        return {'exists': False, 'error': str(e)}

# ============================================================
# テキストメッセージ処理
# ============================================================

def handle_text_message(event, channel_key='MK'):
    text = event['message']['text'].strip()
    user_id = event['source'].get('userId', 'unknown')

    # 「お試し」の場合
    if text == 'お試し':
        customer_info = get_customer_info(user_id, channel_key)
        if not customer_info.get('exists'):
            register_new_user(user_id, channel_key)
        reply_message(event['replyToken'],
            '📝 お試し利用ですね！\n\n'
            'さっそく領収書の写真を送ってみてください📸\n'
            '読み取り結果をお返しします。',
            channel_key)
        return

    # 「契約済み」の場合
    if text == '契約済み':
        code_prefix = 'KZ' if channel_key == 'KZ' else 'MK'
        reply_message(event['replyToken'],
            '✅ ご利用ありがとうございます！\n\n'
            '▼ 既にコードをお持ちの方\n'
            '【顧客コード】を入力してください\n'
            f'例: {code_prefix}001\n\n'
            '▼ これからお申込みの方\n'
            '全プラン最終チェック付き\n\n'
            '📦 ミニマル（月30行まで）\n'
            '月5,000円\n'
            'https://buy.stripe.com/6oUeVdeuA2W0abA7aO3Ru00\n\n'
            '⭐ ライト（月100行まで）← おすすめ\n'
            '月10,000円\n'
            'https://buy.stripe.com/dRm3cvaek7cgerQ9iW3Ru01\n\n'
            '📦 スタンダード（月200行まで）\n'
            '月14,000円\n'
            'https://buy.stripe.com/28EaEXgCI2W06ZoeDg3Ru02\n\n'
            '※超過分は20円/行\n\n'
            '決済後、メールで届くコードを\n'
            'このLINEで入力してください。',
            channel_key)
        return

    # 顧客コードの可能性をチェック（MKで始まる、KZで始まる、または全角版）
    normalized_code = normalize_customer_code(text)

    if normalized_code and (normalized_code.startswith('MK') or normalized_code.startswith('KZ')):
        # 形式チェック
        if not is_valid_customer_code_format(normalized_code):
            code_prefix = 'KZ' if channel_key == 'KZ' else 'MK'
            reply_message(event['replyToken'],
                '⚠️ 顧客コードの形式が正しくありません\n\n'
                f'正しい形式: {code_prefix} + 3桁の数字\n'
                f'例: {code_prefix}001, {code_prefix}123\n\n'
                'コードをご確認のうえ、再度入力してください。',
                channel_key)
            return

        # 存在チェック
        code_check = customer_code_exists(normalized_code)

        if not code_check.get('exists'):
            reply_message(event['replyToken'],
                f'⚠️ 顧客コード「{normalized_code}」は登録されていません\n\n'
                'コードをご確認のうえ、再度入力してください。\n\n'
                'お申込みがまだの方はこちら:\n'
                'https://marunagekeiri.com',
                channel_key)
            return

        # 紐付け処理
        result = link_user_with_customer_code(user_id, normalized_code, channel_key)

        if result.get('success'):
            customer_name = result.get('customer_name', '')
            reply_message(event['replyToken'],
                f'✅ {customer_name}様、ようこそ！\n\n'
                '設定が完了しました。\n'
                '領収書を送ってください📸',
                channel_key)
        elif result.get('already_linked_other'):
            reply_message(event['replyToken'],
                '⚠️ このコードは既に別のアカウントで使用されています\n\n'
                'お心当たりがない場合は、サポートまでご連絡ください。',
                channel_key)
        elif result.get('already_linked'):
            reply_message(event['replyToken'],
                '✅ 既に設定済みです。\n\n'
                '領収書を送ってください📸',
                channel_key)
        else:
            reply_message(event['replyToken'],
                '⚠️ エラーが発生しました。\n'
                'お手数ですが、サポートまでご連絡ください。',
                channel_key)
        return

    # ヘルプ
    if text in ['ヘルプ', 'help', '使い方']:
        help_text = ('📸 領収書の送り方\n\n'
                    '1. このトークに領収書の写真を送信\n'
                    '2. 自動で受け付けられます\n\n'
                    '複数枚ある場合は1枚ずつ送ってください。')
        reply_message(event['replyToken'], help_text, channel_key)
    elif text == '状態確認':
        customer_info = get_customer_info(user_id, channel_key)
        if customer_info.get('exists') and customer_info.get('folder_id'):
            name = customer_info.get('customer_name', '')
            status = customer_info.get('status', '')
            reply_message(event['replyToken'], f'✅ ご登録済み\n👤 {name}\n📋 {status}', channel_key)
        else:
            reply_message(event['replyToken'], '📋 お試し利用中です。\n\nサービス詳細はこちら\nhttps://marunagekeiri.com', channel_key)
    else:
        reply_message(event['replyToken'],
            '📸 領収書の写真を送ってください\n\n'
            '💡「ヘルプ」で使い方を確認できます',
            channel_key)

# ============================================================
# LINE紐付け処理
# ============================================================

def link_user_with_customer_code(user_id, customer_code, channel_key='MK'):
    """顧客コードでLINE IDを紐付け"""
    try:
        service = get_sheets_service()
        sheet_id = get_sheet_id_for_code(customer_code)
        code_col = get_code_col_for_prefix(customer_code)
        status_col_letter = get_status_col_for_prefix(customer_code)
        name_col = 1  # B列（0始まり）- MK/KZ共通
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range='顧客管理!A:M'
        ).execute()
        rows = result.get('values', [])

        for i, row in enumerate(rows[1:], start=2):
            row_code = row[code_col] if len(row) > code_col else ''
            row_line_id = row[0] if len(row) > 0 else ''
            row_name = row[name_col] if len(row) > name_col else ''

            if row_code.upper() == customer_code:
                # 既に別のユーザーが紐付いている場合
                if row_line_id and row_line_id != user_id:
                    return {'success': False, 'already_linked_other': True}

                # 既に同じユーザーが紐付いている場合
                if row_line_id == user_id:
                    return {'success': True, 'already_linked': True, 'customer_name': row_name}

                # 新規紐付け
                service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f'顧客管理!A{i}',
                    valueInputOption='RAW',
                    body={'values': [[user_id]]}
                ).execute()

                # ステータスを「契約済」に更新
                service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f'顧客管理!{status_col_letter}{i}',
                    valueInputOption='RAW',
                    body={'values': [['契約済']]}
                ).execute()

                print(f'Linked user {user_id} with customer code {customer_code} [{channel_key}]')

                # 同じLINE IDの「お試し」行を削除
                delete_trial_rows_for_user(user_id, target_row=i, channel_key=channel_key)

                # フォルダ名を変更
                row_folder_id = row[2] if len(row) > 2 else ''
                if row_folder_id and row_name:
                    rename_customer_folder(row_folder_id, f'{customer_code}_{row_name}')

                # 管理者に通知
                config = get_channel_config(channel_key)
                send_admin_notification(
                    f'✅ LINE連携完了 [{config["name"]}]\n\n'
                    f'👤 {row_name}\n'
                    f'🔑 {customer_code}\n\n'
                    f'領収書の受付を開始しました。'
                )
                return {'success': True, 'customer_name': row_name}

        # ここには来ないはず（事前にcustomer_code_existsでチェック済み）
        return {'success': False, 'not_found': True}

    except Exception as e:
        print(f'Error linking user with customer code: {e}')
        return {'success': False, 'error': str(e)}

def delete_trial_rows_for_user(user_id, target_row, channel_key='MK'):
    """
    同じLINE IDの「お試し」行を削除
    target_row: 紐付け先の行（この行は削除しない）
    """
    try:
        service = get_sheets_service()
        spreadsheet_id = get_sheet_id_for_channel(channel_key)
        status_col = get_status_col_for_channel(channel_key)
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range='顧客管理!A:M'
        ).execute()
        rows = result.get('values', [])

        # 削除対象の行を収集（後ろから削除するため逆順）
        rows_to_delete = []

        for i, row in enumerate(rows[1:], start=2):
            if i == target_row:
                continue  # 紐付け先の行はスキップ

            row_line_id = row[0] if len(row) > 0 else ''
            row_status = row[status_col] if len(row) > status_col else ''

            # 同じLINE IDで「お試し」ステータスの行
            if row_line_id == user_id and row_status == 'お試し':
                rows_to_delete.append(i)

        if not rows_to_delete:
            return

        # スプレッドシートのシートIDを取得
        spreadsheet = service.spreadsheets().get(
            spreadsheetId=spreadsheet_id
        ).execute()
        sheet_gid = None
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == '顧客管理':
                sheet_gid = sheet['properties']['sheetId']
                break

        if sheet_gid is None:
            print('Could not find sheet ID for 顧客管理')
            return

        # 後ろの行から削除（インデックスがずれないように）
        rows_to_delete.sort(reverse=True)

        requests_list = []
        for row_index in rows_to_delete:
            requests_list.append({
                'deleteDimension': {
                    'range': {
                        'sheetId': sheet_gid,
                        'dimension': 'ROWS',
                        'startIndex': row_index - 1,  # 0-indexed
                        'endIndex': row_index
                    }
                }
            })

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests_list}
        ).execute()

        print(f'Deleted {len(rows_to_delete)} trial row(s) for user {user_id} [{channel_key}]')

    except Exception as e:
        print(f'Error deleting trial rows: {e}')

# ============================================================
# ユーティリティ関数
# ============================================================

def download_content_from_line(message_id, channel_key='MK'):
    """LINEから画像/ファイルをダウンロード"""
    config = get_channel_config(channel_key)
    url = f'https://api-data.line.me/v2/bot/message/{message_id}/content'
    headers = {'Authorization': f'Bearer {config["access_token"]}'}
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.content
        print(f'Download failed [{channel_key}]: {response.status_code}')
        return None
    except Exception as e:
        print(f'Error downloading content [{channel_key}]: {e}')
        return None

def upload_via_gas(content, filename, folder_id):
    try:
        content_base64 = base64.b64encode(content).decode('utf-8')
        payload = {
            'image': content_base64,
            'filename': filename,
            'folderId': folder_id
        }
        response = requests.post(GAS_UPLOAD_URL, json=payload, timeout=30)
        result = response.json()
        if result.get('success'):
            print(f'File uploaded via GAS: {result.get("fileId")}')
            return result.get('fileId')
        else:
            print(f'GAS upload error: {result.get("error")}')
            return None
    except Exception as e:
        print(f'Error uploading via GAS: {e}')
        return None

def reply_message(reply_token, text, channel_key='MK'):
    config = get_channel_config(channel_key)
    url = 'https://api.line.me/v2/bot/message/reply'
    headers = {'Content-Type': 'application/json', 'Authorization': f'Bearer {config["access_token"]}'}
    data = {'replyToken': reply_token, 'messages': [{'type': 'text', 'text': text}]}
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code != 200:
            print(f'Reply failed [{channel_key}]: {response.status_code} {response.text}')
    except Exception as e:
        print(f'Error sending reply [{channel_key}]: {e}')

def send_admin_notification(message, channel_key='MK'):
    """管理者にLINE通知を送信（MKチャネル経由で送信）"""
    # 管理者通知は常にMKチャネルから送信
    config = get_channel_config('MK')
    url = 'https://api.line.me/v2/bot/message/push'
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {config["access_token"]}'
    }
    data = {
        'to': ADMIN_USER_ID,
        'messages': [{'type': 'text', 'text': message}]
    }

    try:
        requests.post(url, headers=headers, json=data)
    except Exception as e:
        print(f'Admin notification error: {e}')

def rename_customer_folder(folder_id, new_name):
    """Google Driveのフォルダ名を変更"""
    try:
        credentials, project = default(scopes=['https://www.googleapis.com/auth/drive'])
        service = build('drive', 'v3', credentials=credentials)
        
        file = service.files().get(fileId=folder_id, fields='parents').execute()
        parent_id = file.get('parents', [None])[0]
        
        if parent_id:
            service.files().update(
                fileId=parent_id,
                body={'name': new_name}
            ).execute()
            print(f'Folder renamed to: {new_name}')
            return True
    except Exception as e:
        print(f'Error renaming folder: {e}')
    return False

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
