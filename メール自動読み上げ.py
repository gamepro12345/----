import streamlit as st
from datetime import datetime
import imaplib
import email
from email.header import decode_header, make_header
import quopri, base64, re, json

def remove_unreadable(text):
    # 日本語・英数字・句読点・スペースのみ残す（スラッシュ等も除去）
    return re.sub(r'[^\u3040-\u30FF\u4E00-\u9FFF\uFF10-\uFF19\uFF21-\uFF3A\uFF41-\uFF5A\u0020-\u007E。、．，・！？\n\r]', '', text)


st.title("メール自動読み上げアプリ")

gmail_user = st.text_area("メールアドレスを入力してください")
gmail_pass = st.text_input("メールアドレスのアプリパスワードを入力してください", type="password")

def _decode_mime(s):
    if s is None:
        return ""
    try:
        # "=?UTF-8?...?=" 形式をまとめてデコード
        return str(make_header(decode_header(s)))
    except Exception:
        return s

def _html_to_text(html: str) -> str:
    # 超簡易: タグ除去 & 余分な空白整形（必要ならBeautifulSoupに置換可）
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", "", html)
    text = re.sub(r"(?is)<br\s*/?>", "\n", text)
    text = re.sub(r"(?is)</p\s*>", "\n\n", text)
    text = re.sub(r"(?is)<.*?>", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)
    return text.strip()

def _get_best_body(msg: email.message.Message) -> str:
    """
    text/plain を優先し、無ければ text/html をプレーンテキスト化。
    """
    plain = None
    html = None

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if "attachment" in disp:
                continue
            try:
                payload = part.get_payload(decode=True)
            except Exception:
                payload = None
            charset = part.get_content_charset() or "utf-8"

            if payload is None:
                # デコードできなかった場合の保険
                payload = part.get_payload()
                if isinstance(payload, str):
                    data = payload
                else:
                    data = ""
            else:
                try:
                    data = payload.decode(charset, errors="ignore")
                except Exception:
                    data = payload.decode("utf-8", errors="ignore")

            if ctype == "text/plain" and plain is None:
                plain = data
            elif ctype == "text/html" and html is None:
                html = data
    else:
        ctype = msg.get_content_type()
        payload = msg.get_payload(decode=True)
        charset = msg.get_content_charset() or "utf-8"
        if payload is None:
            data = msg.get_payload() if isinstance(msg.get_payload(), str) else ""
        else:
            try:
                data = payload.decode(charset, errors="ignore")
            except Exception:
                data = payload.decode("utf-8", errors="ignore")

        if ctype == "text/plain":
            plain = data
        elif ctype == "text/html":
            html = data

    if plain and plain.strip():
        return plain.strip()
    if html and html.strip():
        return _html_to_text(html)
    return ""

def fetch_latest_mail(user, password):
    """
    Gmail IMAPからプロモーションカテゴリ最新1通を取得。
    """
    mail = None
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(user, password)
        mail.select('inbox')
        # Gmail拡張検索。必要なら 'UTF-8' を明示して安定させる
        # result, data = mail.search('UTF-8', 'X-GM-RAW', 'category:promotions')
        result, data = mail.search(None, 'X-GM-RAW', 'category:promotions')
        if result != "OK":
            return None
        mail_ids = data[0].split()
        if not mail_ids:
            return None
        latest_id = mail_ids[-1]
        result, msg_data = mail.fetch(latest_id, '(RFC822)')
        if result != "OK":
            return None
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)
        subject = _decode_mime(msg.get("Subject"))
        from_ = _decode_mime(msg.get("From"))
        body = _get_best_body(msg)
        return {
            "subject": subject,
            "from": from_,
            "body": body
        }
    except Exception as e:
        st.error(f"メール取得エラー: {e}")
        return None
    finally:
        try:
            if mail is not None:
                mail.logout()
        except:
            pass

def speak_component(text_to_say: str):
    """
    ページ表示時に自動で発話するHTMLを埋め込む。
    """
    safe = json.dumps(text_to_say)  # JS文字列として安全にエスケープ
    st.components.v1.html(f"""
        <script>
            try {{
                const utter = new SpeechSynthesisUtterance({safe});
                window.speechSynthesis.cancel();
                window.speechSynthesis.speak(utter);
            }} catch (e) {{
                alert('読み上げに失敗しました: ' + e);
            }}
        </script>
    """, height=0)

# ...既存のコード...

if gmail_user and gmail_pass:
    data = fetch_latest_mail(gmail_user, gmail_pass)
    if data is None:
        st.write("まだメールが届いていないか、取得に失敗しました。")
    else:
        subject = data["subject"] or "(件名なし)"
        from_ = data["from"] or "(差出人不明)"
        body = data["body"]

        st.write(f"**差出人**: {from_}")
        st.write(f"**件名**: {subject}")
        st.write("**本文（先頭）**:")
        st.write((body[:500] + "…") if len(body) > 500 else (body or "(本文なし)"))

        # 読み上げ用テキスト（本文が無ければ件名だけでも）
        to_read = f"差出人: {from_}。件名: {subject}。本文: {body}" if body else f"差出人: {from_}。件名: {subject}。"
        to_read = remove_unreadable(to_read)  # ← ここで記号除去
        speak_component(to_read)

        now = datetime.now()
        st.caption(f"取得時刻: {now.strftime('%Y-%m-%d %H:%M:%S')}")
else:
    st.info("Gmailアドレスとアプリパスワードを入力してください。")