import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import threading
import time
import sys

# Windows specific libraries (wrapped in try-except for cross-platform development/checking)
try:
    import win32gui
    import win32process
    import win32api
    import win32con
    import win32com.client
    import psutil
    import keyboard  # Switched from pynput to keyboard library
    IS_WINDOWS = True
except ImportError:
    IS_WINDOWS = False

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
                with open(CONFIG_FILE, "r") as f:
                    self.config = json.load(f)
            except:
                self.config = default_config
        else:
            self.config = default_config
            self.save_config()

    def save_config(self):
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
        self.overlay.destroy()
        if not IS_WINDOWS:
            messagebox.showinfo("情報", "Windows環境以外では動作しません。")
            self.start_button.config(state=tk.NORMAL)
            return

        # Get current active window
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            messagebox.showerror("エラー", "対象のウィンドウを取得できませんでした。")
            self.start_button.config(state=tk.NORMAL)
            return

        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            self.target_process_path = process.exe()
            self.target_hwnd = hwnd # Keep a reference to the main window

            self.is_monitoring = True
            self.status_label.config(text=f"監視中: {os.path.basename(self.target_process_path)}")
            self.stop_button.config(state=tk.NORMAL)

            # Start monitoring thread
            self.monitor_thread = threading.Thread(target=self.monitoring_loop, daemon=True)
            self.monitor_thread.start()

            # Start hotkey listener
            self.start_hotkey_listener()

        except Exception as e:
            messagebox.showerror("エラー", f"プロセスの取得に失敗しました: {e}")
            self.start_button.config(state=tk.NORMAL)

    def start_hotkey_listener(self):
        # Use the keyboard library for more robust global hotkeys on Windows
        try:
            keyboard.add_hotkey(self.config['shortcut'], lambda: self.root.after(0, self.stop_monitoring), suppress=False)
        except Exception as e:
            print(f"Hotkey listener error: {e}")

    def stop_hotkey_listener(self):
        try:
            keyboard.remove_all_hotkeys()
        except:
            pass

    def monitoring_loop(self):
        if IS_WINDOWS:
            import pythoncom
            pythoncom.CoInitialize()

        my_pid = os.getpid()

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
                            self.bring_target_to_front()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        # Some system processes might be restricted
                        self.bring_target_to_front()
            except Exception as e:
                print(f"Monitor error: {e}")

            time.sleep(0.5) # Check every 0.5 seconds

    def bring_target_to_front(self):
        # We need to find a window belonging to the target process path
        def callback(hwnd, hwnds):
            if win32gui.IsWindowVisible(hwnd):
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                try:
                    # Some processes might have multiple windows.
                    # We want to find the one that was likely the main window.
                    p = psutil.Process(pid)
                    if p.exe() == self.target_process_path:
                        # Exclude some common system windows or empty titles if possible
                        title = win32gui.GetWindowText(hwnd)
                        if title:
                            hwnds.append((hwnd, title))
                except:
                    pass
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)

        if hwnds:
            # For Clip Studio, there might be multiple windows.
            # We try to pick the one that looks like a main window (has title).
            # If we already have a target_hwnd from capture, try to see if it's still valid.
            target = hwnds[0][0]

            try:
                # To bring a window to front reliably on Windows:
                # 1. Send a dummy Alt key to unlock the SetForegroundWindow restriction
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('%')

                # 2. Try to restore if minimized
                if win32gui.IsIconic(target):
                    win32gui.ShowWindow(target, win32con.SW_RESTORE)

                # 3. Set foreground
                win32gui.SetForegroundWindow(target)
            except Exception as e:
                # Fallback
                try:
                    win32gui.ShowWindow(target, win32con.SW_SHOW)
                    win32gui.SetForegroundWindow(target)
                except:
                    pass

    def stop_monitoring(self):
        if not self.is_monitoring:
            return
        self.is_monitoring = False
        self.status_label.config(text="停止中")
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.stop_hotkey_listener()
        messagebox.showinfo("情報", "監視を解除しました。")

if __name__ == "__main__":
    root = tk.Tk()
    app = FocusApp(root)
    root.mainloop()
