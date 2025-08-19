import streamlit as st
from datetime import datetime

st.title("メール自動読み上げアプリ")

# 5秒ごとに自動リロード（JavaScriptで実現）
st.components.v1.html("""
    <script>
        setTimeout(function(){
            window.location.reload();
        }, 5000);
    </script>
""", height=0)

email = st.text_input("あなたのメールアドレスを入力してください")

def say(sender, massage):   
    if massage != "":
        message = f"{sender}さんから{massage}と送られました。"
        st.components.v1.html(f"""
            <script>
                var msg = new SpeechSynthesisUtterance("{message}");
                window.speechSynthesis.speak(msg);
            </script>
        """, height=0)
    now = datetime.now()
    st.write(f"投稿日: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    st.write(f"{sender}さんから{massage}と送られました。")

say("ミスター世間クレーマー", "円安が収まらない。")