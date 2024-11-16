import threading
import keyboard
import pyperclip
import pyautogui
import tkinter as tk
import requests
import json
import time

# OLLAMA API 的 URL 和模型名稱
OLLAMA_API_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "llama3.2:3b"

# 定義 system prompt，用於指示 AI 翻譯和潤飾句子
SYSTEM_PROMPT = ("Only give me the sentence.Please translate the following text into fluent English and refine the sentence "
                 "to ensure grammatical accuracy. If it's already in English, improve it.")

# Function to query the Ollama API for a completion
def query_ollama(input_text, stream=True):
    headers = {
        "Content-Type": "application/json"
    }

    payload = {
        "model": MODEL_NAME,
        "prompt": f"{SYSTEM_PROMPT}\n\n{input_text}",
        "stream": stream  # Enable/disable streaming based on the argument
    }

    try:
        # Send request to Ollama API
        response = requests.post(OLLAMA_API_URL, headers=headers, data=json.dumps(payload), stream=stream)
        
        # Check for non-200 HTTP response codes
        if response.status_code != 200:
            return f"API error, status code: {response.status_code}"

        # If streaming is enabled
        if stream:
            full_response = ""
            # Iterate through streamed lines of JSON
            for line in response.iter_lines():
                if line:
                    try:
                        data = json.loads(line.decode('utf-8'))
                        full_response += data.get('response', '')
                        if data.get('done', False):
                            break
                    except json.JSONDecodeError:
                        return f"Failed to parse line: {line}"
            return full_response

        # If streaming is disabled, parse the full response
        else:
            data = response.json()
            return data.get('response', 'AI response is empty.')

    except requests.RequestException as e:
        return f"Request failed: {str(e)}"


# 顯示 loading 畫面
def show_loading_popup(window, loading_label):
    loading_label.config(text="Loading... Please wait.", font=("Arial", 14))
    window.update()

# 顯示確認彈出視窗，並在 AI 運算前先顯示 loading 畫面
def show_confirmation_popup(selected_text):
    def apply_result(rewritten_text):
        # 使用者確認後，貼上結果並覆蓋選取範圍
        pyperclip.copy(rewritten_text)
        keyboard.press_and_release('ctrl+v')
        window.destroy()

    def cancel():
        # 使用者取消操作，關閉視窗
        window.destroy()

    # 建立彈出視窗
    window = tk.Tk()
    window.title("AI 重寫")
    window.geometry("400x200")
    window.attributes("-topmost", True)  # 強制視窗置頂
    window.focus_force()  # 強制焦點到視窗

    loading_label = tk.Label(window, text="Initializing...", font=("Arial", 12))
    loading_label.pack()

    # 顯示 loading 並模擬運算過程
    window.after(1000, show_loading_popup, window, loading_label)

    # 使用背景執行來處理 AI 請求，避免阻塞
    def process_ai():
        rewritten_text = query_ollama(selected_text)

        # 更新視窗內容為 AI 回傳的結果
        loading_label.config(text=rewritten_text)

        # 加入確認和取消按鈕
        confirm_button = tk.Button(window, text="Confirm (Ctrl+Y)", command=lambda: apply_result(rewritten_text))
        confirm_button.pack()

        cancel_button = tk.Button(window, text="Cancel (Esc)", command=cancel)
        cancel_button.pack()

        # 允許快捷鍵操作
        window.bind('<Control-y>', lambda event: apply_result(rewritten_text))
        window.bind('<Escape>', lambda event: cancel())

    threading.Thread(target=process_ai).start()  # 使用線程來處理 AI 請求

    window.mainloop()

# 偵測到快捷鍵後的動作
def handle_hotkey():
    # 使用 pyautogui 模擬 Ctrl+C 來複製當前選取的內容
    pyautogui.hotkey('ctrl', 'c')

    # 延遲一段時間，確保剪貼簿更新
    time.sleep(0.5)  # 增加等待時間，確保剪貼簿有足夠時間更新

    max_attempts = 5
    prev_text = None
    selected_text = ""

    for _ in range(max_attempts):
        selected_text = pyperclip.paste()
        if selected_text != prev_text and selected_text:  # 確認內容有更新
            break
        prev_text = selected_text
        time.sleep(0.1)  # 每次嘗試間隔 0.1 秒

    print(f"選取的文字: {selected_text}")

    if selected_text:
        show_confirmation_popup(selected_text)

# 背景偵測快捷鍵的函數
def hotkey_listener():
    print("開始監聽熱鍵...")
    keyboard.add_hotkey('ctrl+alt+m', handle_hotkey)
    keyboard.wait()  # 持續等待，讓程式持續偵測

# 使用線程來執行快捷鍵監聽
if __name__ == "__main__":
    hotkey_thread = threading.Thread(target=hotkey_listener, daemon=True)
    hotkey_thread.start()
    
    print("程式在背景運行，等待快捷鍵觸發...")
    keyboard.wait()  # 保持主程式運行
