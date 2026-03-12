"""
請求書自動印刷スクリプト
- IMAPでメール取得（前月1日〜末日）
- Claude APIで請求書判定
- PDF添付を自動印刷
- Slackにレポート送信
"""

import imaplib
import email
from email.header import decode_header
import os
import sys
import json
import subprocess
import tempfile
import calendar
import datetime
from pathlib import Path

# ===== 設定 =====
IMAP_HOST = "imap.lolipop.jp"
IMAP_PORT = 993
IMAP_USER = "invoice@y-kyo.com"
IMAP_PASS = os.environ.get("INVOICE_IMAP_PASS", "")

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# 必ず届くはずの送信元（未着チェック用）
REQUIRED_SENDERS = {
    "yabu@k--ing.com": "(有)ケーイング（藪様）",
    "sato@toshu-llc.co.jp": "合同会社トーシュー（佐藤様）",
}

# Slack通知
SLACK_WEBHOOK_URL = os.environ.get("INVOICE_SLACK_WEBHOOK", "")

# 印刷設定
SCRIPT_DIR = Path(__file__).parent
SUMATRA_PATH = SCRIPT_DIR / "SumatraPDF" / "SumatraPDF.exe"
PRINTER_NAME = "Brother DCP-J528N Printer (2 コピー)"

# 処理済みメール記録
PROCESSED_FILE = SCRIPT_DIR / "processed.json"

# ドライランモード（--dry-run で有効）
DRY_RUN = "--dry-run" in sys.argv


def get_target_period():
    """対象期間（当月1日〜実行日）を返す"""
    today = datetime.date.today()
    first_day = datetime.date(today.year, today.month, 1)
    last_day = today
    return first_day, last_day


# ===== ユーティリティ =====

def decode_mime_header(header_value):
    """MIMEエンコードされたヘッダーをデコードする"""
    if not header_value:
        return ""
    decoded_parts = decode_header(header_value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)


def get_email_body(msg):
    """メール本文をテキストで取得する"""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                body += payload.decode(charset, errors="replace")
            elif content_type == "text/html" and not body:
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                body += payload.decode(charset, errors="replace")
    else:
        payload = msg.get_payload(decode=True)
        charset = msg.get_content_charset() or "utf-8"
        body = payload.decode(charset, errors="replace")
    return body[:3000]  # Claude APIに送る量を制限


def extract_pdf_attachments(msg):
    """メールからPDF添付ファイルを抽出する"""
    attachments = []
    if not msg.is_multipart():
        return attachments
    for part in msg.walk():
        content_disposition = str(part.get("Content-Disposition", ""))
        if "attachment" not in content_disposition and "inline" not in content_disposition:
            continue
        filename = part.get_filename()
        if filename:
            filename = decode_mime_header(filename)
        if not filename:
            continue
        if filename.lower().endswith(".pdf"):
            data = part.get_payload(decode=True)
            if data:
                attachments.append({"filename": filename, "data": data})
    return attachments


def load_processed():
    """処理済みMessage-IDを読み込む"""
    if PROCESSED_FILE.exists():
        with open(PROCESSED_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_processed(processed):
    """処理済みMessage-IDを保存する"""
    with open(PROCESSED_FILE, "w", encoding="utf-8") as f:
        json.dump(processed, f, ensure_ascii=False, indent=2)


# ===== IMAP =====

def fetch_emails(first_day, last_day):
    """IMAPから指定期間のメールを取得する"""
    print(f"IMAP接続中: {IMAP_HOST}:{IMAP_PORT}")
    print(f"対象期間: {first_day} 〜 {last_day}")
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(IMAP_USER, IMAP_PASS)
    mail.select("INBOX")

    # IMAP SINCE/BEFORE で期間指定（BEFOREは「その日より前」なので+1日）
    since_str = first_day.strftime("%d-%b-%Y")
    before_date = last_day + datetime.timedelta(days=1)
    before_str = before_date.strftime("%d-%b-%Y")

    status, message_ids = mail.search(None, f'(SINCE {since_str} BEFORE {before_str})')
    if status != "OK" or not message_ids[0]:
        print("対象メールなし")
        mail.logout()
        return []

    emails = []
    for msg_id in message_ids[0].split():
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            continue
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        emails.append(msg)

    mail.logout()
    print(f"{len(emails)}件のメールを取得")
    return emails


# ===== Claude API 判定 =====

def classify_email(sender, subject, body, attachment_names):
    """Claude APIでメールが請求書かどうか判定する"""
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # 送信元のマッピング情報を構築
    sender_info_lines = "\n".join(
        f"  - {addr} → {name}" for addr, name in REQUIRED_SENDERS.items()
    )

    prompt = f"""以下のメールを分析し、請求書（invoice）に関するメールかどうか判定してください。

【既知の取引先一覧】
{sender_info_lines}

送信者: {sender}
件名: {subject}
本文:
{body}

添付ファイル名: {', '.join(attachment_names) if attachment_names else 'なし'}

以下の厳密なJSON形式のみで回答してください（余計なテキスト不要）:
{{
  "is_invoice": true または false,
  "confidence": 0.0〜1.0の数値,
  "company_name": "送信元の会社名（上記の取引先一覧に該当すればその名称を使うこと）",
  "invoice_summary": "請求内容の要約（金額があれば含める）",
  "reason": "判定理由を1行で"
}}"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=500,
        temperature=0,
        messages=[{"role": "user", "content": prompt}],
    )

    text = response.content[0].text.strip()
    # JSON部分を抽出
    if "```" in text:
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()

    return json.loads(text)


# ===== 印刷 =====

def print_pdf(pdf_path):
    """SumatraPDFでPDFを印刷する"""
    if DRY_RUN:
        print(f"  [DRY RUN] 印刷スキップ: {pdf_path}")
        return True

    if not SUMATRA_PATH.exists():
        print(f"  エラー: SumatraPDFが見つかりません: {SUMATRA_PATH}")
        return False

    cmd = [str(SUMATRA_PATH), "-print-to-default", "-silent", str(pdf_path)]
    if PRINTER_NAME:
        cmd = [str(SUMATRA_PATH), "-print-to", PRINTER_NAME, "-silent", str(pdf_path)]

    try:
        result = subprocess.run(cmd, capture_output=True, timeout=120)
        if result.returncode == 0:
            print(f"  印刷成功: {pdf_path.name}")
            return True
        else:
            print(f"  印刷失敗 (code {result.returncode}): {result.stderr.decode(errors='replace')}")
            return False
    except subprocess.TimeoutExpired:
        print(f"  印刷タイムアウト: {pdf_path.name}")
        return False


# ===== Slack通知 =====

def send_slack_report(results, missing_senders, period_str=""):
    """Slackにレポートを送信する"""
    import requests

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    prefix = "[DRY RUN] " if DRY_RUN else ""

    lines = [f"📬 *{prefix}請求書自動印刷レポート* ({now})"]
    lines.append(f"📅 *対象期間: {period_str}*\n")

    if not results and not missing_senders:
        lines.append("新しいメールはありませんでした。")
    else:
        printed = [r for r in results if r.get("printed")]
        skipped = [r for r in results if not r.get("is_invoice")]
        errors = [r for r in results if r.get("error")]

        # 【印刷したファイル】を会社ごとにまとめる
        if printed:
            # 会社名→ファイル名リストを集約
            company_files = {}
            for r in printed:
                name = r["company_name"]
                if name not in company_files:
                    company_files[name] = []
                company_files[name].extend(r.get("filenames", []))

            total_files = sum(len(files) for files in company_files.values())
            lines.append(f"*【印刷したファイル】*")
            lines.append(f"合計 {total_files} ファイル\n")
            for company, files in company_files.items():
                lines.append(f"✅ *{company}* — {len(files)}ファイル")
                for fn in files:
                    lines.append(f"　　・{fn}")

        # スキップしたメール
        if skipped:
            lines.append(f"\n*【スキップ】*")
            for r in skipped:
                lines.append(f"⏭️ *{r['sender']}* — 請求書ではないと判定")
                lines.append(f"　　理由: {r['reason']}")

        # エラー
        if errors:
            lines.append(f"\n*【エラー】*")
            for r in errors:
                lines.append(f"❌ ({r['sender']}): {r['error']}")

    # 必須送信者の未着チェック
    if missing_senders:
        lines.append("\n⚠️ *以下の取引先から請求書が届いていません:*")
        for addr, name in missing_senders.items():
            lines.append(f"　　• {name} ({addr})")

    payload = {"text": "\n".join(lines)}

    try:
        resp = requests.post(SLACK_WEBHOOK_URL, json=payload, timeout=30)
        if resp.status_code == 200:
            print("Slack通知送信完了")
        else:
            print(f"Slack通知失敗: {resp.status_code} {resp.text}")
    except Exception as e:
        print(f"Slack通知エラー: {e}")


# ===== メイン処理 =====

def main():
    print("=" * 50)
    print(f"請求書自動印刷 開始 ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M')})")
    if DRY_RUN:
        print("*** ドライランモード（印刷しません）***")
    print("=" * 50)

    # 環境変数チェック
    missing_env = []
    if not IMAP_PASS:
        missing_env.append("INVOICE_IMAP_PASS")
    if not ANTHROPIC_API_KEY:
        missing_env.append("ANTHROPIC_API_KEY")
    if not SLACK_WEBHOOK_URL:
        missing_env.append("INVOICE_SLACK_WEBHOOK")
    if missing_env:
        print(f"エラー: 環境変数が未設定です: {', '.join(missing_env)}")
        sys.exit(1)

    # 対象期間（前月）
    first_day, last_day = get_target_period()
    period_str = f"{first_day.year}年{first_day.month}月"
    print(f"対象月: {period_str}")

    # 処理済みリスト読み込み
    processed = load_processed()

    # メール取得
    try:
        emails = fetch_emails(first_day, last_day)
    except Exception as e:
        error_msg = f"IMAP接続エラー: {e}"
        print(error_msg)
        send_slack_report([{"sender": "SYSTEM", "error": error_msg}], {}, period_str)
        sys.exit(1)

    results = []
    confirmed_senders = set()  # 届いた必須送信者

    for msg in emails:
        message_id = msg.get("Message-ID", "")

        # 処理済みならスキップ
        if message_id in processed:
            # 必須送信者チェックのため差出人だけ確認
            sender = email.utils.parseaddr(msg.get("From", ""))[1].lower()
            if sender in REQUIRED_SENDERS:
                confirmed_senders.add(sender)
            continue

        sender = email.utils.parseaddr(msg.get("From", ""))[1].lower()
        sender_display = decode_mime_header(msg.get("From", ""))
        subject = decode_mime_header(msg.get("Subject", ""))
        date_str = msg.get("Date", "")

        print(f"\n--- 処理中: {sender} / {subject} ---")

        if sender in REQUIRED_SENDERS:
            confirmed_senders.add(sender)

        # 添付PDF抽出
        pdf_attachments = extract_pdf_attachments(msg)
        attachment_names = [a["filename"] for a in pdf_attachments]

        # メール本文取得
        body = get_email_body(msg)

        # Claude APIで判定
        result = {
            "sender": sender,
            "sender_display": sender_display,
            "subject": subject,
            "date": date_str,
            "filenames": attachment_names,
        }

        try:
            classification = classify_email(sender, subject, body, attachment_names)
            result["is_invoice"] = classification.get("is_invoice", False)
            result["confidence"] = classification.get("confidence", 0)
            result["company_name"] = classification.get("company_name", sender)
            result["invoice_summary"] = classification.get("invoice_summary", "")
            result["reason"] = classification.get("reason", "")
            print(f"  判定: {'請求書' if result['is_invoice'] else '請求書ではない'} (確信度: {result['confidence']})")
        except Exception as e:
            result["error"] = f"Claude API判定エラー: {e}"
            result["is_invoice"] = False
            print(f"  判定エラー: {e}")
            results.append(result)
            continue  # 判定失敗はprocessedに入れない（次回再処理）

        # 請求書なら印刷
        if result["is_invoice"] and pdf_attachments:
            result["printed"] = True
            print_success = True
            with tempfile.TemporaryDirectory() as tmpdir:
                for att in pdf_attachments:
                    pdf_path = Path(tmpdir) / att["filename"]
                    with open(pdf_path, "wb") as f:
                        f.write(att["data"])
                    if not print_pdf(pdf_path):
                        print_success = False
                        result["error"] = f"印刷失敗: {att['filename']}"
            if not print_success:
                result["printed"] = False
        elif result["is_invoice"] and not pdf_attachments:
            result["printed"] = False
            result["error"] = "請求書と判定されたがPDF添付なし"
            print("  警告: PDF添付ファイルなし")
        else:
            result["printed"] = False

        results.append(result)

        # 処理済みに記録
        processed[message_id] = {
            "sender": sender,
            "subject": subject,
            "date": date_str,
            "is_invoice": result["is_invoice"],
            "processed_at": datetime.datetime.now().isoformat(),
        }

    # 処理済みリスト保存
    save_processed(processed)

    # 必須送信者の未着チェック
    missing_senders = {}
    for addr, name in REQUIRED_SENDERS.items():
        if addr not in confirmed_senders:
            missing_senders[addr] = name

    # Slack通知
    send_slack_report(results, missing_senders, period_str)

    print("\n完了!")


if __name__ == "__main__":
    main()
