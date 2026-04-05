def render_speech_synthesis_button(text, key):
    """テキストを読み上げるスピーカーボタンを表示する"""
    if not text:
        return
    
    # JavaScriptによる読み上げロジック
    # クリーンアップ（改行などの除去）
    clean_text = text.replace("'", "\\'").replace("\n", " ")
    
    html_code = f"""
    <button id="btn-{key}" style="
        background: none;
        border: 1px solid #ddd;
        border-radius: 5px;
        padding: 2px 8px;
        cursor: pointer;
        font-size: 16px;
        margin-top: 5px;
        color: #555;
    " onclick="speak_{key}()" title="読み上げる">🔊</button>
    
    <script>
    function speak_{key}() {{
        const btn = document.getElementById('btn-{key}');
        
        // 再生中の場合は停止
        if (window.speechSynthesis.speaking) {{
            window.speechSynthesis.cancel();
            btn.innerText = '🔊';
            return;
        }}
        
        // iOS Safari対策: 一度空のcancelを呼ぶことで音声エンジンを強制的にアクティブにする
        window.speechSynthesis.cancel();
        
        // 少し遅延を入れてから発話させる（iOS対策）
        setTimeout(() => {{
            const uttr = new SpeechSynthesisUtterance('{clean_text}');
            uttr.lang = 'ja-JP';
            uttr.rate = 1.1;
            
            uttr.onstart = () => {{ btn.innerText = '⏹'; btn.style.color = '#dc3545'; }};
            uttr.onend = () => {{ btn.innerText = '🔊'; btn.style.color = '#555'; }};
            uttr.onerror = (e) => {{
                console.error("SpeechSynthesisError:", e);
                btn.innerText = '🔊'; 
                btn.style.color = '#555'; 
            }};
            
            window.speechSynthesis.speak(uttr);
        }}, 50);
    }}
    </script>
    """
    components.html(html_code, height=45)

def render_voice_input_button(key_prefix):
    """音声入力ボタンを表示し、結果をセッション状態に返す"""
    # Streamlitのセッション状態との橋渡し用hidden field
    input_key = f"{key_prefix}_voice_input_result"
    
    html_code = f"""
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <button id="mic-btn-{key_prefix}" style="
            background-color: #f0f2f6;
            border: 1px solid #dcdfe6;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            cursor: pointer;
            font-size: 20px;
            display: flex;
            justify-content: center;
            align-items: center;
            transition: all 0.3s;
        " onclick="startRecognition()">🎤</button>
        <span id="status-{key_prefix}" style="margin-left: 10px; font-size: 14px; color: #666;"></span>
    </div>

    <script>
    function startRecognition() {{
        const btn = document.getElementById('mic-btn-{key_prefix}');
        const status = document.getElementById('status-{key_prefix}');
        const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
        
        recognition.lang = 'ja-JP';
        recognition.interimResults = false;
        recognition.maxAlternatives = 1;

        recognition.onstart = () => {{
            btn.style.backgroundColor = '#ff4b4b';
            btn.style.color = 'white';
            status.innerText = '音声を認識中... 話してください';
        }};

        recognition.onspeechend = () => {{
            recognition.stop();
        }};

        recognition.onresult = (event) => {{
            const result = event.results[0][0].transcript;
            status.innerText = '認識完了: ' + result;
            
            // 親ウィンドウ（Streamlit）の隠し入力フィールドに値をセットして送信
            // ただしStreamlitの仕様上、直接セットしても反応しない場合があるため、
            // カスタムイベントや特定のDOM操作が必要
            window.parent.postMessage({{
                type: 'streamlit:set_component_value',
                value: result,
                key: '{input_key}'
            }}, '*');
            
            // 簡易的な方法として、ブラウザのプロンプト等で値を渡すことも可能だが、
            // ここではStreamlitのセッション更新を待つ
            setTimeout(() => {{
                // ページ全体にメッセージを送る
                const event = new CustomEvent('voiceInput', {{ detail: result }});
                window.parent.document.dispatchEvent(event);
            }}, 500);
        }};

        recognition.onerror = (event) => {{
            if (event.error === 'not-allowed') {{
                status.innerText = 'マイク権限エラー: 設定で許可するか、AndroidはHTTPS通信が必要です';
            }} else {{
                status.innerText = 'エラーが発生しました: ' + event.error;
            }}
            btn.style.backgroundColor = '#f0f2f6';
            btn.style.color = 'black';
        }};

        recognition.onend = () => {{
            btn.style.backgroundColor = '#f0f2f6';
            btn.style.color = 'black';
        }};

        recognition.start();
    }}
    </script>
    """
    
    # 認識結果を受け取るためのコンポーネント
    # 注意: Streamlit公式のiframeから親への通信は制限があるため、
    # 実際にはURLパラメータや、カスタムコンポーネントライブラリなしでは少し工夫が必要。
    # ここでは、認識されたテキストを一時的に表示し、ユーザーが確認して送信できるスタイルにするか、
    # あるいは直接セッションに書き込むための「隠しボタン」的なアプローチをとる。
    
    components.html(html_code, height=60)
    
    # 結果を受け取るための実験的な仕組み
    # (実際には st.chat_input に自動で流し込むのはJSのセキュリティ制約上難しいため、
    # 音声認識されたテキストを通知として表示し、それを入力欄に反映させるガイドを出すのが現実的)
    return None

# ---------- ページUIの実装 ----------