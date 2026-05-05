import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import threading
import time
import sys

# Windows specific libraries (wrapped in try-except for cross-platform development/checking)
IS_WINDOWS = False
IMPORT_ERROR_MSG = ""

try:
    import win32gui
    import win32process
    import win32api
    import win32con
    import win32com.client
    import psutil
    import keyboard  # Switched from pynput to keyboard library
    IS_WINDOWS = True
except ImportError as e:
    IMPORT_ERROR_MSG = str(e)
    # On non-Windows environments for dev, this is expected.
    # On Windows, this means libraries are missing.
    print(f"[DEBUG] ライブラリのインポートに失敗しました: {e}")

CONFIG_FILE = "config.json"

class FocusApp:
    def __init__(self, root):
        self.root = root
        self.root.title("作業集中アプリ")
        self.root.geometry("300x200")

        self.target_process_path = None
        self.is_monitoring = False
        self.hotkey_listener = None

        self.load_config()
        self.setup_ui()

    def load_config(self):
        default_config = {"shortcut": "alt+p"}
        if os.path.exists(CONFIG_FILE):
            try:
                print(f"[DEBUG] 設定を読み込み中: {CONFIG_FILE}")
                with open(CONFIG_FILE, "r") as f:
                    self.config = json.load(f)
            except:
                print("[ERROR] 設定の読み込みに失敗しました。デフォルトを使用します。")
                self.config = default_config
        else:
            print("[DEBUG] 設定ファイルが見つかりません。新規作成します。")
            self.config = default_config
            self.save_config()

    def save_config(self):
        print(f"[DEBUG] 設定を保存中: {CONFIG_FILE}")
        with open(CONFIG_FILE, "w") as f:
            json.dump(self.config, f)

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(frame, text="停止中", font=("MS Gothic", 12))
        self.status_label.pack(pady=10)

        self.start_button = ttk.Button(frame, text="監視開始 (8秒後に捕捉)", command=self.start_countdown)
        self.start_button.pack(pady=5)

        self.stop_button = ttk.Button(frame, text="監視停止", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(pady=5)

        ttk.Label(frame, text=f"ショートカット: {self.config['shortcut']}").pack(pady=5)

    def start_countdown(self):
        print("[DEBUG] カウントダウン開始 (8秒)")
        self.start_button.config(state=tk.DISABLED)
        self.countdown_val = 8
        self.show_overlay()
        self.update_countdown()

    def show_overlay(self):
        self.overlay = tk.Toplevel(self.root)
        self.overlay.attributes("-topmost", True)
        self.overlay.overrideredirect(True) # Remove title bar

        # Center of screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        width, height = 400, 300
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.overlay.geometry(f"{width}x{height}+{x}+{y}")

        # Use a bright color for visibility
        self.overlay.configure(bg='black')
        self.countdown_label = tk.Label(self.overlay, text=str(self.countdown_val),
                                        font=("Arial", 120, "bold"), fg="white", bg="black")
        self.countdown_label.pack(expand=True)

        # Make background "transparent" or semi-transparent if Windows
        if IS_WINDOWS:
            self.overlay.attributes("-alpha", 0.7)

    def update_countdown(self):
        if self.countdown_val > 0:
            self.countdown_label.config(text=str(self.countdown_val))
            self.countdown_val -= 1
            self.root.after(1000, self.update_countdown)
        else:
            self.capture_and_start()

    def capture_and_start(self):
        print("[DEBUG] カウントダウン終了。対象アプリを捕捉します。")
        self.overlay.destroy()
        if not IS_WINDOWS:
            print(f"[ERROR] Windowsライブラリが利用不可です: {IMPORT_ERROR_MSG}")
            messagebox.showerror("エラー", f"Windowsライブラリの読み込みに失敗しました。\n\n詳細: {IMPORT_ERROR_MSG}\n\npip install -r requirements.txt を実行したか確認してください。")
            self.start_button.config(state=tk.NORMAL)
            return

        # Get current active window
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            print("[ERROR] アクティブなウィンドウが見つかりませんでした。")
            messagebox.showerror("エラー", "対象のウィンドウを取得できませんでした。")
            self.start_button.config(state=tk.NORMAL)
            return

        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            self.target_process_path = process.exe()

            print(f"[DEBUG] 対象を捕捉しました: {os.path.basename(self.target_process_path)} (PID: {pid})")
            print(f"[DEBUG] パス: {self.target_process_path}")

            self.is_monitoring = True
            self.status_label.config(text=f"監視中: {os.path.basename(self.target_process_path)}")
            self.stop_button.config(state=tk.NORMAL)

            # Start monitoring thread
            print("[DEBUG] 監視スレッドを開始します。")
            self.monitor_thread = threading.Thread(target=self.monitoring_loop, daemon=True)
            self.monitor_thread.start()

            # Start hotkey listener
            self.start_hotkey_listener()

        except Exception as e:
            print(f"[ERROR] 対象の捕捉中にエラーが発生しました: {e}")
            messagebox.showerror("エラー", f"プロセスの取得に失敗しました: {e}")
            self.start_button.config(state=tk.NORMAL)

    def start_hotkey_listener(self):
        # Use the keyboard library for more robust global hotkeys on Windows
        shortcut = self.config['shortcut']
        print(f"[DEBUG] グローバルホットキーを登録します: {shortcut}")
        try:
            def on_hotkey():
                print(f"[EVENT] グローバルホットキー ({shortcut}) を検知しました！")
                self.root.after(0, self.stop_monitoring)

            keyboard.add_hotkey(shortcut, on_hotkey, suppress=False)
            print("[DEBUG] ホットキーの登録に成功しました。")
        except Exception as e:
            print(f"[ERROR] ホットキーの登録に失敗しました: {e}")

    def stop_hotkey_listener(self):
        print("[DEBUG] 全てのホットキー設定を解除します。")
        try:
            keyboard.remove_all_hotkeys()
        except Exception as e:
            print(f"[ERROR] ホットキーの解除中にエラーが発生しました: {e}")

    def monitoring_loop(self):
        if IS_WINDOWS:
            import pythoncom
            pythoncom.CoInitialize()

        my_pid = os.getpid()
        print(f"[DEBUG] 監視ループを開始しました。自PID: {my_pid}")

        while self.is_monitoring:
            try:
                curr_hwnd = win32gui.GetForegroundWindow()
                if curr_hwnd:
                    _, pid = win32process.GetWindowThreadProcessId(curr_hwnd)
                    if pid == my_pid:
                        # Don't snatch focus if the user is interacting with this app
                        time.sleep(0.5)
                        continue

                    try:
                        curr_process = psutil.Process(pid)
                        curr_path = curr_process.exe()

                        if curr_path != self.target_process_path:
                            # Not the target process! Bring it back.
                            print(f"[EVENT] 非対象アプリへの切り替えを検知: {os.path.basename(curr_path)}")
                            self.bring_target_to_front()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        # Some system processes might be restricted
                        print("[DEBUG] プロセス情報にアクセスできません。呼び戻しを試みます。")
                        self.bring_target_to_front()
            except Exception as e:
                print(f"[ERROR] 監視ループ内でエラー: {e}")

            time.sleep(0.5) # Check every 0.5 seconds

    def bring_target_to_front(self):
        print("[DEBUG] 対象アプリを前面に呼び戻します...")
        # We need to find a window belonging to the target process path
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    p = psutil.Process(pid)
                    if p.exe() == self.target_process_path:
                        title = win32gui.GetWindowText(hwnd)
                        if title:
                            hwnds.append((hwnd, title))
                except:
                    pass
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)

        if hwnds:
            target = hwnds[0][0]
            print(f"[DEBUG] 前面に移動するウィンドウ: {hwnds[0][1]}")

            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('%')

                if win32gui.IsIconic(target):
                    win32gui.ShowWindow(target, win32con.SW_RESTORE)

                win32gui.SetForegroundWindow(target)
                print("[DEBUG] 呼び戻しに成功しました。")
            except Exception as e:
                print(f"[ERROR] 呼び戻しに失敗しました: {e}")
                try:
                    win32gui.ShowWindow(target, win32con.SW_SHOW)
                    win32gui.SetForegroundWindow(target)
                except:
                    pass

    def stop_monitoring(self):
        if not self.is_monitoring:
            return
        print("[DEBUG] 監視停止リクエストを受理しました。")
        self.is_monitoring = False
        self.status_label.config(text="停止中")
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.stop_hotkey_listener()
        print("[DEBUG] 監視を停止し、後処理を完了しました。")
        messagebox.showinfo("情報", "監視を解除しました。")

if __name__ == "__main__":
    root = tk.Tk()
    app = FocusApp(root)
    root.mainloop()
