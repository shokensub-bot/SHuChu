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
HISTORY_FILE = "history.json"

import re

class FocusApp:
    def __init__(self, root):
        self.root = root
        self.root.title("作業集中アプリ")
        self.root.geometry("400x300") # Slightly wider for new UI

        self.target_process_path = None
        self.is_monitoring = False
        self.hotkey_listener = None
        self.timer_end_time = None
        self.is_recording_active = False

        self.load_config()
        self.load_history()
        self.setup_ui()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_config(self):
        default_config = {
            "shortcut": "alt+p",
            "timer_enabled": False,
            "timer_minutes": 30,
            "auto_record_enabled": False
        }
        if os.path.exists(CONFIG_FILE):
            try:
                print(f"[DEBUG] 設定を読み込み中: {CONFIG_FILE}")
                with open(CONFIG_FILE, "r") as f:
                    loaded_config = json.load(f)
                    self.config = {**default_config, **loaded_config}
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

    def load_history(self):
        self.history = []
        if os.path.exists(HISTORY_FILE):
            try:
                print(f"[DEBUG] 履歴を読み込み中: {HISTORY_FILE}")
                with open(HISTORY_FILE, "r") as f:
                    self.history = json.load(f)
                    if not isinstance(self.history, list):
                        self.history = []
            except:
                print("[ERROR] 履歴の読み込みに失敗しました。")
                self.history = []
        else:
            print("[DEBUG] 履歴ファイルが見つかりません。")

    def save_history(self):
        print(f"[DEBUG] 履歴を保存中: {HISTORY_FILE}")
        with open(HISTORY_FILE, "w") as f:
            json.dump(self.history, f)

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(frame, text="停止中", font=("MS Gothic", 12))
        self.status_label.pack(pady=5)

        # Timer settings
        timer_frame = ttk.LabelFrame(frame, text="タイマーモード設定", padding="10")
        timer_frame.pack(fill=tk.X, pady=5)

        self.timer_enabled_var = tk.BooleanVar(value=self.config.get("timer_enabled", False))
        self.timer_radio_off = ttk.Radiobutton(timer_frame, text="無効", variable=self.timer_enabled_var, value=False, command=self.on_config_ui_change)
        self.timer_radio_off.pack(side=tk.LEFT, padx=5)
        self.timer_radio_on = ttk.Radiobutton(timer_frame, text="有効", variable=self.timer_enabled_var, value=True, command=self.on_config_ui_change)
        self.timer_radio_on.pack(side=tk.LEFT, padx=5)

        self.timer_widgets = []
        lbl_time = ttk.Label(timer_frame, text="時間:")
        lbl_time.pack(side=tk.LEFT, padx=(10, 2))
        self.timer_widgets.append(lbl_time)

        initial_val = self.minutes_to_hms(self.config.get("timer_minutes", 30))
        self.timer_minutes_var = tk.StringVar(value=initial_val)

        self.timer_combo = ttk.Combobox(timer_frame, textvariable=self.timer_minutes_var, width=8)
        self.update_history_display()
        self.timer_combo.pack(side=tk.LEFT)
        self.timer_combo.bind("<FocusOut>", lambda e: self.on_config_ui_change())
        self.timer_combo.bind("<<ComboboxSelected>>", lambda e: self.on_config_ui_change())
        self.timer_widgets.append(self.timer_combo)

        self.delete_hist_button = tk.Button(timer_frame, text="×", fg="red", command=self.delete_current_history,
                                            relief=tk.FLAT, font=("Arial", 10, "bold"))
        self.delete_hist_button.pack(side=tk.LEFT, padx=2)
        self.timer_widgets.append(self.delete_hist_button)

        lbl_min = ttk.Label(timer_frame, text="分")
        lbl_min.pack(side=tk.LEFT)
        self.timer_widgets.append(lbl_min)

        # Auto record settings
        self.auto_record_var = tk.BooleanVar(value=self.config.get("auto_record_enabled", False))
        self.auto_record_check = ttk.Checkbutton(frame, text="自動録画モード (win+alt+r)",
                                                 variable=self.auto_record_var, command=self.on_config_ui_change)
        self.auto_record_check.pack(pady=5)

        self.update_timer_ui_state()

        self.start_button = ttk.Button(frame, text="監視開始 (8秒後に捕捉)", command=self.start_countdown)
        self.start_button.pack(pady=5)

        self.stop_button = ttk.Button(frame, text="監視停止", command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.pack(pady=5)

        ttk.Label(frame, text=f"ショートカット: {self.config['shortcut']}").pack(pady=5)

    def minutes_to_hms(self, total_minutes):
        if total_minutes <= 0:
            return str(total_minutes)

        h = total_minutes // 60
        m = total_minutes % 60

        if h > 0 and m > 0:
            return f"{h}h{m}m"
        elif h > 0:
            return f"{h}h"
        else:
            return f"{m}m"

    def hms_to_minutes(self, s):
        s = s.strip().lower()
        if not s:
            return 0

        # Pattern for "1h30m", "1h30", "1h", "30m", "30"
        match = re.match(r'^(\d+h)?(\d+m?)?$', s)
        if not match:
            # Fallback to see if it's just a number
            try:
                return int(s)
            except ValueError:
                return 0

        h_str, m_str = match.groups()
        total = 0
        if h_str:
            total += int(h_str[:-1]) * 60
        if m_str:
            if m_str.endswith('m'):
                total += int(m_str[:-1])
            else:
                total += int(m_str)
        return total

    def update_history_display(self):
        # Sort history by time (minutes) and convert to HMS format
        sorted_history = sorted(list(set(self.history)))
        display_values = [self.minutes_to_hms(m) for m in sorted_history]
        self.timer_combo['values'] = display_values

    def delete_current_history(self):
        current_val = self.timer_minutes_var.get()
        minutes = self.hms_to_minutes(current_val)
        if minutes in self.history:
            self.history.remove(minutes)
            self.save_history()
            self.update_history_display()
            print(f"[DEBUG] 履歴から削除しました: {current_val} ({minutes}分)")

    def on_config_ui_change(self):
        self.config["timer_enabled"] = self.timer_enabled_var.get()
        self.config["auto_record_enabled"] = self.auto_record_var.get()
        try:
            val = self.hms_to_minutes(self.timer_minutes_var.get())
            if val > 0:
                self.config["timer_minutes"] = val
        except ValueError:
            pass
        self.save_config()
        self.update_timer_ui_state()

    def update_timer_ui_state(self):
        state = tk.NORMAL if self.timer_enabled_var.get() else tk.DISABLED
        for w in self.timer_widgets:
            try:
                w.config(state=state)
            except:
                pass

    def set_ui_state(self, state):
        """監視中/停止中のUI有効・無効切り替え"""
        self.start_button.config(state=state)
        self.timer_radio_off.config(state=state)
        self.timer_radio_on.config(state=state)
        self.auto_record_check.config(state=state)

        if state == tk.NORMAL:
            self.update_timer_ui_state()
        else:
            for w in self.timer_widgets:
                try:
                    w.config(state=tk.DISABLED)
                except:
                    pass

    def start_countdown(self):
        # Update config from UI variables to ensure they are in sync
        self.config["timer_enabled"] = self.timer_enabled_var.get()

        input_val = self.timer_minutes_var.get()
        minutes = self.hms_to_minutes(input_val)

        # Validate timer if enabled
        if self.config.get("timer_enabled"):
            if minutes <= 0:
                messagebox.showerror("エラー", "有効な時間を入力してください（1分以上の時間を指定してください）。")
                return

            self.config["timer_minutes"] = minutes

            # Update history
            if minutes not in self.history:
                print(f"[DEBUG] 履歴に新しい時間を追加します: {minutes}分")
                self.history.append(minutes)
                self.save_history()
                self.update_history_display()

        self.save_config()

        print(f"[DEBUG] カウントダウン開始 (8秒) - タイマーモード: {self.config['timer_enabled']} ({self.config.get('timer_minutes')}分)")
        self.set_ui_state(tk.DISABLED)
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

            # Start timer if enabled
            if self.config.get("timer_enabled"):
                minutes = self.config.get("timer_minutes", 30)
                self.timer_end_time = time.time() + (minutes * 60)
                print(f"[DEBUG] タイマーを開始します: {minutes}分 (終了予定: {time.strftime('%H:%M:%S', time.localtime(self.timer_end_time))})")

            # Start auto recording if enabled
            if self.config.get("auto_record_enabled"):
                print("[DEBUG] 自動録画を開始します (win+alt+r)")
                try:
                    keyboard.press_and_release('win+alt+r')
                    self.is_recording_active = True
                except Exception as re:
                    print(f"[ERROR] 録画開始キーの送信に失敗しました: {re}")

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
            # Check timer
            if self.timer_end_time and time.time() >= self.timer_end_time:
                print("[EVENT] タイマー時間が経過しました。")
                self.timer_end_time = None
                self.root.after(0, self.on_timer_complete)
                break

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

    def stop_monitoring(self, show_msg=True):
        if not self.is_monitoring:
            return
        print("[DEBUG] 監視停止リクエストを受理しました。")
        self.is_monitoring = False

        # Stop recording if active
        if self.is_recording_active:
            print("[DEBUG] 自動録画を停止します (win+alt+r)")
            try:
                keyboard.press_and_release('win+alt+r')
            except Exception as re:
                print(f"[ERROR] 録画停止キーの送信に失敗しました: {re}")
            self.is_recording_active = False

        # Reset timer
        if self.timer_end_time:
            print("[DEBUG] 実行中のタイマーをリセットします。")
            self.timer_end_time = None

        self.status_label.config(text="停止中")
        self.set_ui_state(tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.stop_hotkey_listener()
        print("[DEBUG] 監視を停止し、後処理を完了しました。")
        if show_msg:
            messagebox.showinfo("情報", "監視を解除しました。")

    def on_timer_complete(self):
        print("[EVENT] タイマー完了処理を開始します。")
        self.timer_end_time = None
        # Stop monitoring (including recording stop) without showing the dialog
        self.stop_monitoring(show_msg=False)

        # Show completion overlay
        self.show_completion_overlay()

    def on_closing(self):
        if self.is_monitoring:
            print("[DEBUG] 監視中にアプリが終了されます。")
            self.stop_monitoring()
        self.root.destroy()

    def show_completion_overlay(self):
        print("[DEBUG] 完了オーバーレイを表示します。")
        self.comp_overlay = tk.Toplevel(self.root)
        self.comp_overlay.attributes("-topmost", True)
        self.comp_overlay.attributes("-fullscreen", True)
        self.comp_overlay.configure(bg='black')
        self.comp_overlay.focus_force()

        label = tk.Label(
            self.comp_overlay,
            text="時間になりました！\nESCキーを押して戻ってください…",
            font=("MS Gothic", 40, "bold"),
            fg="white",
            bg="black",
            justify=tk.CENTER
        )
        label.pack(expand=True)

        self.comp_overlay.bind("<Escape>", lambda e: self.comp_overlay.destroy())

if __name__ == "__main__":
    root = tk.Tk()
    app = FocusApp(root)
    root.mainloop()
