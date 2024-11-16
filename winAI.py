import threading
import keyboard
import pyperclip
import pyautogui
import tkinter as tk
from tkinter import ttk, messagebox
import requests
import json
import time
import sys
from PIL import Image, ImageTk
import os
import configparser
import queue
import ctypes
import win32clipboard
from tkinter import font

# Windows DPI awareness
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

class ModernUI(ttk.Frame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.master = master
        self.setup_styles()
        self.create_widgets()
        
    def setup_styles(self):
        # 創建現代風格
        style = ttk.Style()
        style.theme_use('clam')
        
        # 定義顏色
        self.colors = {
            'primary': '#3B82F6',  # Blue-600
            'primary_hover': '#2563EB',  # Blue-700
            'secondary': '#F3F4F6',  # Gray-100
            'background': '#FFFFFF',
            'text': '#1F2937',  # Gray-800
            'border': '#E5E7EB',  # Gray-200
            'error': '#EF4444',  # Red-500
            'success': '#10B981'  # Green-500
        }
        
        # 配置全局字體
        self.default_font = ('Segoe UI', 10)
        self.header_font = ('Segoe UI', 12, 'bold')
        
        # 配置各種元素樣式
        style.configure('Modern.TFrame',
                       background=self.colors['background'])
        
        style.configure('Card.TFrame',
                       background=self.colors['background'],
                       relief='flat',
                       borderwidth=0)
        
        style.configure('Modern.TLabel',
                       background=self.colors['background'],
                       foreground=self.colors['text'],
                       font=self.default_font)
        
        style.configure('Header.TLabel',
                       background=self.colors['background'],
                       foreground=self.colors['text'],
                       font=self.header_font)
        
        # 按鈕樣式
        style.configure('Primary.TButton',
                       padding=(20, 10),
                       background=self.colors['primary'],
                       foreground='white',
                       font=self.default_font)
        
        style.map('Primary.TButton',
                 background=[('active', self.colors['primary_hover']),
                           ('disabled', self.colors['border'])])
        
        style.configure('Secondary.TButton',
                       padding=(20, 10),
                       background=self.colors['secondary'],
                       foreground=self.colors['text'],
                       font=self.default_font)
        
        # 進度條樣式
        style.configure('Modern.Horizontal.TProgressbar',
                       background=self.colors['primary'],
                       troughcolor=self.colors['secondary'],
                       thickness=6)

    def create_widgets(self):
        self.configure(style='Card.TFrame')
        self.grid(sticky='nsew', padx=20, pady=20)
        
        # 配置grid權重
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # 標題區域
        header_frame = ttk.Frame(self, style='Card.TFrame')
        header_frame.grid(row=0, column=0, sticky='ew', pady=(0, 20))
        
        ttk.Label(header_frame,
                 text="WinAI Translator",
                 style='Header.TLabel').grid(row=0, column=0, sticky='w')
        
        ttk.Label(header_frame,
                 text="Ctrl + Alt + M to translate selected text",
                 style='Modern.TLabel').grid(row=1, column=0, sticky='w')
        
        # 原文區域
        ttk.Label(self,
                 text="Original Text",
                 style='Modern.TLabel').grid(row=1, column=0, sticky='w', pady=(0, 5))
        
        # 創建文本框容器
        text_container = ttk.Frame(self, style='Card.TFrame')
        text_container.grid(row=2, column=0, sticky='nsew', pady=(0, 15))
        text_container.grid_columnconfigure(0, weight=1)
        
        self.original_text = tk.Text(text_container,
                                   height=5,
                                   wrap='word',
                                   font=self.default_font,
                                   relief='flat',
                                   padx=10,
                                   pady=10)
        self.original_text.grid(row=0, column=0, sticky='nsew')
        self.original_text.configure(bg=self.colors['secondary'])
        
        # 添加滾動條
        original_scrollbar = ttk.Scrollbar(text_container, orient='vertical', command=self.original_text.yview)
        original_scrollbar.grid(row=0, column=1, sticky='ns')
        self.original_text['yscrollcommand'] = original_scrollbar.set
        
        # 翻譯區域
        ttk.Label(self,
                 text="Translation",
                 style='Modern.TLabel').grid(row=3, column=0, sticky='w', pady=(0, 5))
        
        # 翻譯文本框容器
        translation_container = ttk.Frame(self, style='Card.TFrame')
        translation_container.grid(row=4, column=0, sticky='nsew', pady=(0, 15))
        translation_container.grid_columnconfigure(0, weight=1)
        
        self.translation_text = tk.Text(translation_container,
                                      height=5,
                                      wrap='word',
                                      font=self.default_font,
                                      relief='flat',
                                      padx=10,
                                      pady=10)
        self.translation_text.grid(row=0, column=0, sticky='nsew')
        self.translation_text.configure(bg=self.colors['secondary'])
        
        # 添加滾動條
        translation_scrollbar = ttk.Scrollbar(translation_container, orient='vertical', command=self.translation_text.yview)
        translation_scrollbar.grid(row=0, column=1, sticky='ns')
        self.translation_text['yscrollcommand'] = translation_scrollbar.set
        
        # 進度條
        self.progress = ttk.Progressbar(self,
                                      mode='indeterminate',
                                      style='Modern.Horizontal.TProgressbar')
        
        # 按鈕區域
        button_frame = ttk.Frame(self, style='Card.TFrame')
        button_frame.grid(row=6, column=0, sticky='e', pady=(20, 0))
        
        self.cancel_btn = ttk.Button(button_frame,
                                   text="Cancel",
                                   style='Secondary.TButton',
                                   command=self.master.destroy)
        self.cancel_btn.pack(side='right', padx=(10, 0))
        
        self.confirm_btn = ttk.Button(button_frame,
                                    text="Confirm",
                                    style='Primary.TButton',
                                    command=self.confirm)
        self.confirm_btn.pack(side='right')

    def confirm(self):
        text = self.translation_text.get('1.0', tk.END).strip()
        pyperclip.copy(text)
        keyboard.write(text)
        self.master.destroy()

    def set_original_text(self, text):
        self.original_text.delete('1.0', tk.END)
        self.original_text.insert('1.0', text)
        self.original_text.config(state='disabled')

    def set_translation(self, text):
        self.translation_text.delete('1.0', tk.END)
        self.translation_text.insert('1.0', text)
        
    def start_progress(self):
        self.progress.grid(row=5, column=0, sticky='ew', pady=(15, 0))
        self.progress.start(10)
        
    def stop_progress(self):
        self.progress.stop()
        self.progress.grid_remove()

class ModernWinAI:
    def __init__(self):
        self.setup_config()
        self.hotkey_active = True
        try:
            self.setup_hotkey()
            print("ModernWinAI initialized successfully!")
        except Exception as e:
            print(f"Failed to initialize hotkey: {e}")
            raise

    def setup_config(self):
        self.config = configparser.ConfigParser()
        self.config['API'] = {
            'url': 'http://localhost:11434/api/generate',
            'model': 'llama3.2:3b'
        }
        self.config['Hotkeys'] = {
            'translate': 'ctrl+alt+m',
            'confirm': 'ctrl+y',
            'cancel': 'escape'
        }

    def setup_hotkey(self):
        try:
            # 先嘗試移除所有現有的快捷鍵
            try:
                keyboard.unhook_all()
            except:
                pass

            # 使用 on_press 方法來實現快捷鍵功能
            def check_hotkey(e):
                if self.hotkey_active:
                    # 解析設定的快捷鍵組合
                    hotkey = self.config['Hotkeys']['translate'].lower()
                    keys = set(hotkey.split('+'))
                    
                    # 獲取當前按下的按鍵
                    pressed_keys = set()
                    if keyboard.is_pressed('ctrl'):
                        pressed_keys.add('ctrl')
                    if keyboard.is_pressed('alt'):
                        pressed_keys.add('alt')
                    if keyboard.is_pressed('m'):
                        pressed_keys.add('m')
                    
                    # 檢查是否匹配設定的快捷鍵組合
                    if pressed_keys == keys:
                        self.handle_hotkey()

            # 註冊按鍵監聽
            keyboard.on_press(check_hotkey)
            
            print(f"Hotkey {self.config['Hotkeys']['translate']} registered successfully!")
            
        except Exception as e:
            error_msg = f"Failed to setup hotkey: {str(e)}"
            print(error_msg)
            raise Exception(error_msg)

    def get_selected_text(self):
        """使用多種方法嘗試獲取選中的文字"""
        original_clipboard = None
        selected_text = ""
        
        try:
            # 保存原始剪貼簿內容
            try:
                original_clipboard = pyperclip.paste()
            except:
                pass

            # 方法1: 使用 pyperclip
            try:
                pyperclip.copy('')
                time.sleep(0.2)
                keyboard.send('ctrl+c')
                time.sleep(0.5)
                selected_text = pyperclip.paste()
            except Exception as e:
                print(f"Method 1 failed: {e}")

            # 方法2: 如果方法1失敗，直接嘗試訪問Windows剪貼簿
            if not selected_text:
                try:
                    win32clipboard.OpenClipboard()
                    selected_text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                except Exception as e:
                    print(f"Method 2 failed: {e}")

            # 方法3: 使用 pyautogui
            if not selected_text:
                try:
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(0.5)
                    selected_text = pyperclip.paste()
                except Exception as e:
                    print(f"Method 3 failed: {e}")

            if selected_text:
                selected_text = selected_text.strip()
                print(f"Successfully got selected text: {selected_text[:50]}...")
            
            return selected_text if selected_text else ""

        finally:
            try:
                if original_clipboard is not None:
                    time.sleep(0.2)
                    pyperclip.copy(original_clipboard)
            except Exception as e:
                print(f"Failed to restore clipboard: {e}")

    def handle_hotkey(self):
        try:
            print("Hotkey triggered, attempting to get selected text...")
            self.hotkey_active = False
            
            selected_text = self.get_selected_text()
            print(f"Selection result: {'Success' if selected_text else 'Failed'}")
            
            if selected_text:
                print(f"Selected text length: {len(selected_text)}")
                self.create_window(selected_text)
            else:
                print("No text was selected or copied")
                messagebox.showinfo("Notice", "Please select text first")
                
        except Exception as e:
            print(f"Error in handle_hotkey: {e}")
            messagebox.showerror("Error", f"Failed to process selection: {e}")
        finally:
            self.hotkey_active = True

    def create_window(self, selected_text=""):
        root = tk.Tk()
        root.title("WinAI Translator")
        
        # 設置窗口大小和位置
        window_width = 600
        window_height = 500
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 設置窗口樣式
        root.configure(bg='white')
        if sys.platform == 'win32':
            root.attributes('-alpha', 0.98)
        root.attributes('-topmost', True)
        
        # 綁定快捷鍵
        root.bind('<Escape>', lambda e: root.destroy())
        root.bind('<Control-y>', lambda e: self.ui.confirm())
        
        # 創建UI
        self.ui = ModernUI(root)
        if selected_text:
            self.ui.set_original_text(selected_text)
            self.start_translation(selected_text)
        
        root.mainloop()

    def start_translation(self, text):
        self.ui.start_progress()
        threading.Thread(target=self.translate_text, args=(text,), daemon=True).start()

    def translate_text(self, text):
        try:
            response = self.query_ollama(text)
            if response:
                self.ui.master.after(0, self.ui.set_translation, response)
        except Exception as e:
            self.ui.master.after(0, messagebox.showerror, "Error", str(e))
        finally:
            self.ui.master.after(0, self.ui.stop_progress)

    def query_ollama(self, text):
        headers = {"Content-Type": "application/json"}
        payload = {
            "model": self.config['API']['model'],
            "prompt": f"Please translate the following text into fluent English and refine the sentence to ensure grammatical accuracy. If it's already in English, improve it.\n\n{text}",
            "stream": False
        }

        try:
            response = requests.post(
                self.config['API']['url'], 
                headers=headers, 
                json=payload,
                timeout=30
            )
            response.raise_for_status()
            return response.json().get('response', '').strip()
        except requests.exceptions.RequestException as e:
            raise Exception(f"Translation service error: {str(e)}")

    def run(self):
        """主程序運行方法"""
        try:
            # 設置Windows任務欄圖標
            if sys.platform == 'win32':
                import ctypes
                myappid = 'mycompany.winai.translator.1.0'
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            
            print("ModernWinAI is running...")
            print(f"Press {self.config['Hotkeys']['translate']} to translate selected text")
            print(f"Press {self.config['Hotkeys']['confirm']} to confirm translation")
            print(f"Press {self.config['Hotkeys']['cancel']} to cancel")
            
            # 主循環
            while True:
                try:
                    time.sleep(1)
                except KeyboardInterrupt:
                    print("ModernWinAI shutting down...")
                    break
                except Exception as e:
                    print(f"Error in main loop: {e}")
                    time.sleep(5)
        finally:
            # 清理工作
            try:
                keyboard.unhook_all()
            except:
                pass

def show_startup_error(error_message):
    """顯示啟動錯誤對話框"""
    root = tk.Tk()
    root.withdraw()  # 隱藏主窗口
    messagebox.showerror("Startup Error", error_message)
    root.destroy()

def main():
    try:
        app = ModernWinAI()
        app.run()
    except Exception as e:
        error_message = f"Application failed to start: {e}"
        print(f"Critical error: {error_message}")
        show_startup_error(error_message)
        sys.exit(1)

if __name__ == "__main__":
    main()