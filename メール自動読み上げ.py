import streamlit as st
import os
from datetime import datetime
import imaplib
import email
from email.header import decode_header, make_header
import quopri, base64, re, json

def remove_unreadable(text):
    # 日本語・英数字・句読点・スペースのみ残す（スラッシュ等も除去）
    return re.sub(r'[^\u3040-\u30FF\u4E00-\u9FFF\uFF10-\uFF19\uFF21-\uFF3A\uFF41-\uFF5A\u0020-\u007E。、．，・！？\n\r]', '', text)


st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@700&display=swap');
    h1 {
        font-family: 'Noto Sans JP', sans-serif !important;
        color: #2c3e50;
        letter-spacing: 2px;
    }
    </style>
""", unsafe_allow_html=True)

# 変更: タイトルと画像を横並びで表示（元の st.title(...) と不要な st. 行を置換）
col_title, col_img = st.columns([3, 1])
with col_title:
    st.title("メール自動読み上げアプリ")
with col_img:
    # unnamed.jpg をこのスクリプトと同じフォルダで探す
    img_path = os.path.join(os.path.dirname(__file__), "unnamed.jpg") if "__file__" in globals() else "unnamed.jpg"
    if os.path.exists(img_path):
        st.image(img_path, width=140)
    else:
        # 画像が無ければ空白を維持（必要なら代替テキストを表示）
        st.write("")

st.write("このアプリはあなたがお使いのメールの最新メールを取得し、内容を自動で読み上げます。")
st.write("メールアドレスとアプリパスワードを入れればOKです。")
st.write("アプリパスワードの取得方法は[こちら](https://support.google.com/accounts/answer/185833?hl=ja)を参照してください。")
st.write("また、カテゴリごとに最新メールを選択できます。")
gmail_user = st.text_input("メールアドレスを入力してください")
gmail_pass = st.text_input("メールアドレスのアプリパスワードを入力してください", type="password")

category = st.selectbox(
    "読むメールの種類を選んでください",
    ("すべて", "メイン", "広告")
)
def fetch_mails(user, password, category="広告", num=10):
    """
    Gmail IMAPからカテゴリ最新num件を取得。
    """
    mails = []
    mail = None
    try:
        imap_host = get_imap_host(user)  # 変更: ホストを決定
        mail = imaplib.IMAP4_SSL(imap_host)
        mail.login(user, password)
        mail.select('inbox')
        if category == "すべて":
            result, data = mail.search(None, 'ALL')
        elif category == "メイン":
            result, data = mail.search(None, 'X-GM-RAW', 'category:primary')
        else:
            result, data = mail.search(None, 'X-GM-RAW', 'category:promotions')
        if result != "OK":
            return []
        mail_ids = data[0].split()
        if not mail_ids:
            return []
        # 最新num件のIDを取得
        latest_ids = mail_ids[-num:]
        for mail_id in reversed(latest_ids):
            result, msg_data = mail.fetch(mail_id, '(RFC822)')
            if result != "OK":
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            subject = _decode_mime(msg.get("Subject"))
            from_ = _decode_mime(msg.get("From"))
            body = _get_best_body(msg)
            mails.append({
                "subject": subject,
                "from": from_,
                "body": body
            })
        return mails
    except Exception as e:
        st.error(f"メール取得エラー: {e}")
        return []
    finally:
        try:
            if mail is not None:
                mail.logout()
        except:
            pass

def remove_unreadable(text):
    # URLを除去
    text = re.sub(r'https?://\S+|www\.\S+', '', text)
    # 日本語・英数字・句読点・スペースのみ残す（記号は除去）
    text = re.sub(r'[^0-9A-Za-z\u3040-\u30FF\u4E00-\u9FFF。、．，・！？\s\n\r]', '', text)
    return text

def _decode_mime(s):
    if s is None:
        return ""
    try:
        # "=?UTF-8?...?=" 形式をまとめてデコード
        return str(make_header(decode_header(s)))
    except Exception:
        return s

# 追加: メールアドレスから適切なIMAPホストを決定するヘルパー
def get_imap_host(user_email: str) -> str:
    """
    user_email のドメインに応じてIMAPホストを返す。
    - ドメインが 'gmail.com' を末尾に含む場合は Gmail の公式ホストを使う（test.6765884.gmail.com 等に対応）
    - それ以外は簡易的に 'imap.<domain>' を返す（必要なら設定UIを追加してください）
    """
    try:
        domain = user_email.split('@')[-1].lower()
    except Exception:
        domain = ""
    if domain.endswith("gmail.com"):
        return "imap.gmail.com"
    if domain:
        return f"imap.{domain}"
    return "imap.gmail.com"

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


def fetch_latest_mail(user, password, category="広告"):
    """
    Gmail IMAPからカテゴリ最新1通を取得。
    """
    mail = None
    try:
        imap_host = get_imap_host(user)  # 変更: ホストを決定
        mail = imaplib.IMAP4_SSL(imap_host)
        mail.login(user, password)
        mail.select('inbox')
        # ▼カテゴリごとに検索条件を切り替え
        if category == "すべて":
            result, data = mail.search(None, 'ALL')
        elif category == "メイン":
            result, data = mail.search(None, 'X-GM-RAW', 'category:primary')
        else:  # "広告"
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
    ページに再生ボタンを追加し、自動で押す試みを行う。
    - 可視ボタン（ユーザーが押せる）と隠しボタン（自動クリック対象）を作る
    - どちらのクリックでも短い音声を再生してから SpeechSynthesis を行う
    - 自動クリックは element.click(), MouseEvent dispatch, 繰り返し試行 を行う
    """
    safe = json.dumps(text_to_say)  # JS文字列として安全にエスケープ

    # 簡易サイレントWAV（短いヘッダのみ）を data URI として使う
    silent_wav = "data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAESsAACJWAAACABAAZGF0YQAAAAA="

    # 変更: height を 0 -> 140 にしてボタンを表示可能にする
    st.components.v1.html(f"""
        <script>
        (async function(){{
            const text = {safe};
            const silent = "{silent_wav}";

            function waitForBody(timeout = 2000) {{
                return new Promise((resolve) => {{
                    if (document && document.body) return resolve(document.body);
                    let resolved = false;
                    const onReady = () => {{
                        if (document && document.body) {{
                            cleanup();
                            resolved = true;
                            resolve(document.body);
                        }}
                    }};
                    const cleanup = () => {{
                        document.removeEventListener('DOMContentLoaded', onReady);
                        document.removeEventListener('readystatechange', onReady);
                    }};
                    document.addEventListener('DOMContentLoaded', onReady);
                    document.addEventListener('readystatechange', onReady);

                    let waited = 0;
                    const iv = setInterval(() => {{
                        if (document && document.body) {{
                            clearInterval(iv);
                            cleanup();
                            if (!resolved) {{
                                resolved = true;
                                resolve(document.body);
                            }}
                        }}
                        waited += 100;
                        if (waited > timeout) {{
                            clearInterval(iv);
                            cleanup();
                            if (!resolved) {{
                                resolved = true;
                                resolve(null);
                            }}
                        }}
                    }}, 100);
                }});
            }}

            function doSpeakOnce() {{
                try {{
                    // 無音音声を再生してオーディオをアンロック
                    const audio = document.createElement('audio');
                    audio.src = silent;
                    audio.autoplay = true;
                    audio.playsInline = true;
                    audio.muted = true;
                    audio.style.width = "1px";
                    audio.style.height = "1px";
                    audio.style.opacity = "0";
                    audio.style.position = "fixed";
                    audio.style.left = "-9999px";
                    // append は安全に行う
                    try {{ (document.body || document.documentElement || document).appendChild(audio); }} catch(e) {{ /* ignore */ }}

                    // 再生試行（Promise拒否は無視）
                    try {{ audio.play().catch(()=>{{}}); }} catch(e) {{ }}

                    // 少し待ってからSpeechSynthesis
                    setTimeout(function() {{
                        try {{
                            const utter = new SpeechSynthesisUtterance(text);
                            window.speechSynthesis.cancel();
                            window.speechSynthesis.speak(utter);
                        }} catch (e) {{
                            console.warn('speak failed', e);
                        }}
                    }}, 200);

                }} catch (e) {{
                    console.warn('doSpeakOnce error', e);
                }}
            }}

            // ボタン作成
            try {{
                const body = await waitForBody();
                const container = body || document.documentElement || document;

                // 可視ボタン（ユーザーが押せる）
                const btn = document.createElement('button');
                btn.id = 'streamlit_speak_button';
                btn.textContent = '読み上げ再生';
                btn.style.position = 'relative';
                btn.style.display = 'inline-block';
                btn.style.margin = '8px';
                btn.style.padding = '8px 12px';
                btn.style.borderRadius = '6px';
                btn.style.background = '#4CAF50';
                btn.style.color = '#fff';
                btn.style.border = 'none';
                btn.style.boxShadow = '0 2px 6px rgba(0,0,0,0.2)';
                btn.style.cursor = 'pointer';

                // 隠しボタン（自動クリック用）
                const hidden = document.createElement('button');
                hidden.id = 'streamlit_speak_button_hidden';
                hidden.style.position = 'fixed';
                hidden.style.left = '-9999px';
                hidden.style.opacity = '0';
                hidden.style.width = '1px';
                hidden.style.height = '1px';

                const handler = function(e) {{
                    e && e.preventDefault && e.preventDefault();
                    doSpeakOnce();
                }};

                btn.addEventListener('click', handler);
                hidden.addEventListener('click', handler);

                try {{ container.appendChild(btn); }} catch(e) {{ /* ignore */ }}
                try {{ container.appendChild(hidden); }} catch(e) {{ /* ignore */ }}

                // 自動クリックを複数手法で試す
                const tryAutoClick = () => {{
                    try {{ hidden.click(); }} catch(e){{}}
                    try {{ hidden.dispatchEvent(new MouseEvent('click', {{bubbles:true, cancelable:true, view:window}})); }} catch(e){{}}
                    try {{ btn.click(); }} catch(e){{}}
                }};

                // 即時、アニメフレーム、タイマーで繰り返し試行（計1秒程度）
                tryAutoClick();
                requestAnimationFrame(() => tryAutoClick());
                setTimeout(() => tryAutoClick(), 50);
                setTimeout(() => tryAutoClick(), 200);
                setTimeout(() => tryAutoClick(), 500);

                // もし自動クリックで動作しなければ、ボタンは残してユーザーが押せるようにする
            }} catch (e) {{
                console.warn('speak_component init error', e);
            }}
        }})();
        </script>
    """, height=140)

if gmail_user and gmail_pass:
    mails = fetch_mails(gmail_user, gmail_pass, category, num=10)
    if not mails:
        st.write("まだメールが届いていないか、取得に失敗しました。")
    else:
        # 件名一覧を表示して選択
        subjects = [f"{i+1}. {remove_unreadable(m['subject'])}" for i, m in enumerate(mails)]
        selected = st.selectbox("読み上げるメールを選んでください", subjects)
        idx = subjects.index(selected)
        mail = mails[idx]
        subject = mail["subject"] or "(件名なし)"
        from_ = mail["from"] or "(差出人不明)"
        body = mail["body"]

        from_masked = re.sub(r'<.*?>', '<***>', from_)

        st.write(f"**差出人**: {from_masked}")
        st.write(f"**件名**: {subject}")
        st.write("**本文（先頭）**:")
        st.write((body[:500] + "…") if len(body) > 500 else (body or "(本文なし)"))

        to_read = f"差出人: {from_masked}。件名: {subject}。本文: {body}" if body else f"差出人: {from_masked}。件名: {subject}。"
        to_read = remove_unreadable(to_read)
        speak_component(to_read)

        now = datetime.now()
        st.caption(f"取得時刻: {now.strftime('%Y-%m-%d %H:%M:%S')}")
else:
    st.info("Gmailアドレスとアプリパスワードを入力してください。")