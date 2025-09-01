import streamlit as st
from datetime import datetime
import imaplib
import email
from email.header import decode_header

st.title("メール自動読み上げアプリ")

gmail_user = st.text_area("メールアドレスを入力してください")
gmail_pass = st.text_area("メールアドレスのアプリパスワードを入力してください")
def fetch_latest_mail(user, password):
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(user, password)
        mail.select('inbox')
        # プロモーションカテゴリのみ検索
        result, data = mail.search(None, 'X-GM-RAW', 'category:promotions')
        mail_ids = data[0].split()
        if not mail_ids:
            return None, None
        latest_id = mail_ids[-1]
        result, msg_data = mail.fetch(latest_id, '(RFC822)')
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        subject, encoding = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8", errors="ignore")
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                disp = str(part.get("Content-Disposition"))
                if ctype == "text/plain" and "attachment" not in disp:
                    charset = part.get_content_charset() or "utf-8"
                    body = part.get_payload(decode=True).decode(charset, errors="ignore")
                    break
        else:
            charset = msg.get_content_charset() or "utf-8"
            body = msg.get_payload(decode=True).decode(charset, errors="ignore")
        return subject, body
    except Exception as e:
        st.error(f"メール取得エラー: {e}")
        return None, None

def say(sender, massage):
    if massage:
        message = f"{sender}さんから{massage}と送られました。"
        # ページ表示時に自動で読み上げ
        st.components.v1.html(f"""
            <script>
                var msg = new SpeechSynthesisUtterance("{message}");
                window.speechSynthesis.speak(msg);
            </script>
        """, height=0)
    now = datetime.now()
    st.write(f"投稿日: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    st.write(f"{sender}さんから{massage}と送られました。")



if gmail_user and gmail_pass:
    subject, body = fetch_latest_mail(gmail_user, gmail_pass)
    if subject is not None:
        # 本文が「広告」だけの場合のみ表示
        if body.strip() == "広告":
            st.write("広告")
        else:
            say(subject, body)
    else:
        st.write("まだメールが届いていません。")
else:
    st.info("Gmailアドレスとアプリパスワードを入力してください。")