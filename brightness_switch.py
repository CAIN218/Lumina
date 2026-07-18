import os
import subprocess
import keyboard
import tkinter as tk
from tkinter import messagebox
import time
import sys
import threading
import queue
import ctypes
import winshell
from win32com.client import Dispatch

import screen_brightness_control as sbc

# customtkinter（設定画面専用のUIライブラリ）はここでは読み込まない。
# バックグラウンド常駐時はホットキー監視しかしないため、
# 使わない重いUIアセットをメモリに乗せないよう open_setting_window() 内で遅延importする。

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "app_config.txt")
ICON_PATH = os.path.join(BASE_DIR, "app_icon.ico")

last_adjust_time = 0

# ----------------------------------------------------------------------
# デザイントークン — AudioSwitcherと同じグラファイト×カッパー配色を踏襲し、
# シリーズとして統一感を持たせている。
# ----------------------------------------------------------------------
COLOR_BG           = "#16151a"
COLOR_PANEL        = "#1d1b21"
COLOR_INPUT        = "#26232b"
COLOR_BORDER       = "#322e38"
COLOR_ACCENT       = "#d68a4c"
COLOR_ACCENT_HOVER = "#e6a05f"
COLOR_ACCENT_DIM   = "#5c4530"
COLOR_TEXT_MAIN    = "#f2ece1"
COLOR_TEXT_SUB     = "#8f8878"
COLOR_TEXT_FAINT   = "#5c5850"

FONT_DISPLAY = ("Segoe UI Semibold", 17)
FONT_LABEL   = ("Segoe UI", 10, "bold")
FONT_BODY    = ("Segoe UI", 10)
FONT_MONO    = ("Consolas", 11, "bold")

DEFAULT_HOTKEY_UP = "ctrl+alt+up"
DEFAULT_HOTKEY_DOWN = "ctrl+alt+down"
DEFAULT_STEP = "10"


def kill_previous_instances():
    """裏で既に動いている古いBrightSwitchのプロセスをすべて安全に終了させる"""
    try:
        current_pid = os.getpid()
        target_process = "BrightSwitch.exe"
        if getattr(sys, 'frozen', False):
            subprocess.run(
                f'taskkill /F /IM {target_process} /FI "PID ne {current_pid}"',
                shell=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
    except Exception:
        pass


def get_monitor_snapshot():
    """接続中のモニター一覧と現在の輝度を取得する。
    モニターによってはDDC/CI非対応で例外が出るため、1台ずつ個別にtryする。"""
    snapshot = []
    try:
        infos = sbc.list_monitors_info()
    except Exception as e:
        log_error("list_monitors_info", e)
        return snapshot

    for i, info in enumerate(infos):
        name = info.get("name") or info.get("model") or f"Monitor {i}"
        try:
            brightness = sbc.get_brightness(display=i)[0]
        except Exception as e:
            log_error(f"get_brightness[{i}]", e)
            brightness = None
        snapshot.append({"index": i, "name": name, "brightness": brightness})
    return snapshot


def set_monitor_brightness(index, value):
    """0-100にクランプしてからDDC/CI経由で設定する"""
    value = max(0, min(100, int(round(value))))
    sbc.set_brightness(value, display=index)
    return value


def load_config():
    """設定ファイルの読み込み"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
            if len(lines) >= 4:
                return lines[0], lines[1], lines[2], lines[3]
    return DEFAULT_HOTKEY_UP, DEFAULT_HOTKEY_DOWN, DEFAULT_STEP, "0"


def save_config(hotkey_up, hotkey_down, step, startup_val):
    """設定ファイルの保存とスタートアップ制御"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        f.write(f"{hotkey_up}\n{hotkey_down}\n{step}\n{startup_val}")
    manage_startup(startup_val)


def manage_startup(startup_val):
    """Windowsのスタートアップフォルダへの自動登録・解除（生スクリプト/exeハイブリッド対応）"""
    try:
        startup_dir = winshell.startup()
        shortcut_path = os.path.join(startup_dir, "BrightSwitch.lnk")

        if startup_val == "1":
            if getattr(sys, 'frozen', False):
                target = sys.executable
                arguments = "/background"
                icon_target = target
            else:
                pythonw_path = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
                target = pythonw_path if os.path.exists(pythonw_path) else sys.executable
                arguments = f'"{os.path.abspath(__file__)}" /background'
                icon_target = None

            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = target
            shortcut.Arguments = arguments
            shortcut.WorkingDirectory = BASE_DIR

            if icon_target:
                shortcut.IconLocation = icon_target

            shortcut.save()
        else:
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
    except Exception as e:
        tk.messagebox.showerror("Startup Error", f"Failed to update startup folder: {str(e)}")


class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class _MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", _RECT),
                ("rcWork", _RECT), ("dwFlags", ctypes.c_ulong)]


def _active_monitor_work_area():
    """マウスカーソルがあるモニターの作業領域（タスクバーを除く）を返す(AudioSwitcherと同じ考え方)"""
    pt = _POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    MONITOR_DEFAULTTONEAREST = 2
    hmonitor = ctypes.windll.user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
    info = _MONITORINFO()
    info.cbSize = ctypes.sizeof(_MONITORINFO)
    ctypes.windll.user32.GetMonitorInfoW(hmonitor, ctypes.byref(info))
    rc = info.rcWork
    return rc.left, rc.top, rc.right, rc.bottom


def _rounded_rect(cv, x1, y1, x2, y2, r, **kwargs):
    points = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return cv.create_polygon(points, smooth=True, **kwargs)


# 常駐中は1枚のToplevelを使い回し、Tclインタプリタの再生成を避ける(AudioSwitcherと同じ理由)
class _ToastManager:
    WIDTH, HEIGHT = 320, 68

    def __init__(self, root):
        self.root = root
        self.window = None
        self.canvas = None
        self.fade_job = None
        self.current_alpha = 0.0

    def _ensure_window(self):
        if self.window is not None:
            return
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.attributes("-alpha", 0.0)

        transparent_key = "#010203"
        toast.configure(bg=transparent_key)
        toast.attributes("-transparentcolor", transparent_key)

        canvas = tk.Canvas(toast, width=self.WIDTH, height=self.HEIGHT,
                            bg=transparent_key, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        toast.withdraw()

        self.window = toast
        self.canvas = canvas

    def _place_bottom_right(self):
        try:
            left, top, right, bottom = _active_monitor_work_area()
        except Exception:
            left, top = 0, 0
            right = self.window.winfo_screenwidth()
            bottom = self.window.winfo_screenheight()
        x = right - self.WIDTH - 6
        y = bottom - self.HEIGHT - 6
        self.window.geometry(f"{self.WIDTH}x{self.HEIGHT}+{x}+{y}")

    def show(self, brightness_percent):
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None

        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 16,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        canvas.create_rectangle(0, 12, 4, self.HEIGHT - 12, fill=COLOR_ACCENT, outline="")
        canvas.create_text(22, 20, anchor="w", text="BRIGHTNESS",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))
        canvas.create_text(22, 42, anchor="w", text=f"{brightness_percent}%",
                            fill=COLOR_TEXT_MAIN, font=("Segoe UI", 14, "bold"))

        # 明るさを表す横棒インジケーター
        bar_x1, bar_x2 = 130, self.WIDTH - 22
        bar_y1, bar_y2 = 36, 44
        _rounded_rect(canvas, bar_x1, bar_y1, bar_x2, bar_y2, 4,
                      fill=COLOR_INPUT, outline=COLOR_BORDER, width=1)
        fill_width = (bar_x2 - bar_x1) * max(0, min(100, brightness_percent)) / 100
        if fill_width > 2:
            _rounded_rect(canvas, bar_x1, bar_y1, bar_x1 + fill_width, bar_y2, 4,
                          fill=COLOR_ACCENT, outline="")

        self._place_bottom_right()
        self.window.deiconify()
        self._place_bottom_right()
        self.current_alpha = 0.0
        self._fade_in()

    def _fade_in(self):
        if self.current_alpha < 0.96:
            self.current_alpha += 0.12
            self.window.attributes("-alpha", min(self.current_alpha, 0.96))
            self.fade_job = self.root.after(15, self._fade_in)
        else:
            self.fade_job = self.root.after(2000, self._fade_out)

    def _fade_out(self):
        if self.current_alpha > 0.0:
            self.current_alpha -= 0.06
            self.window.attributes("-alpha", max(self.current_alpha, 0.0))
            self.fade_job = self.root.after(15, self._fade_out)
        else:
            self.window.withdraw()
            self.fade_job = None


_toast_queue = queue.Queue()


def log_error(context, exc):
    """サイレントに握りつぶさず、原因調査用にエラーをログファイルへ記録する"""
    try:
        log_path = os.path.join(BASE_DIR, "error.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {context}: {repr(exc)}\n")
            if exc is not None and exc.__traceback__ is not None:
                import traceback
                f.write("".join(traceback.format_tb(exc.__traceback__)))
    except Exception:
        pass


def _log_thread_exception(args):
    log_error(f"thread:{args.thread.name if args.thread else '?'}", args.exc_value)


threading.excepthook = _log_thread_exception


# AudioSwitcherで踏んだ不具合(低レベルキーボードフックのコールバック内で
# 重い処理をブロッキングで行うと、Windowsがそのフックを黙って無効化し
# ホットキーが反応しなくなることがある)を教訓に、ホットキーのコールバックは
# デバウンス判定とキューへの投入だけに留め、実際のDDC/CI通信は別スレッドの
# ワーカーに任せる。
_adjust_request_queue = queue.Queue()


def _make_hotkey_callback(direction):
    def _callback():
        global last_adjust_time
        current_time = time.time()
        if current_time - last_adjust_time < 0.35:
            return
        last_adjust_time = current_time
        _adjust_request_queue.put(direction)
    return _callback


brightness_up = _make_hotkey_callback(+1)
brightness_down = _make_hotkey_callback(-1)


def _adjust_worker():
    """キューに積まれた上げ/下げ要求を、フックコールバックの外側で実処理する。
    DDC/CI通信(I2C経由)は遅いモニターだと数十〜百msかかることがあるが、
    ここで時間を使ってもキーボードフックの応答性には影響しない。"""
    _, _, step_str, _ = load_config()
    while True:
        direction = _adjust_request_queue.get()
        try:
            _, _, step_str, _ = load_config()
            try:
                step = int(step_str)
            except ValueError:
                step = int(DEFAULT_STEP)

            snapshot = get_monitor_snapshot()
            if not snapshot:
                continue

            new_values = []
            for m in snapshot:
                if m["brightness"] is None:
                    continue
                target = m["brightness"] + direction * step
                try:
                    new_value = set_monitor_brightness(m["index"], target)
                    new_values.append(new_value)
                except Exception as e:
                    log_error(f"set_brightness[{m['index']}]", e)

            if new_values:
                avg = round(sum(new_values) / len(new_values))
                _toast_queue.put(avg)
        except Exception as e:
            log_error("adjust_worker", e)


def _rehook_watchdog(hotkey_up, hotkey_down, interval_sec=300):
    """他アプリのグローバルフックとの競合でkeyboardライブラリの内部状態が
    ズレてホットキーが反応しなくなる既知の問題への対策(AudioSwitcherと同じ)。"""
    while True:
        time.sleep(interval_sec)
        try:
            keyboard.unhook_all()
            keyboard.add_hotkey(hotkey_up, brightness_up, trigger_on_release=False)
            keyboard.add_hotkey(hotkey_down, brightness_down, trigger_on_release=False)
        except Exception as e:
            log_error("rehook_watchdog", e)


def start_backend():
    """バックグラウンドでのキーボード監視 + 常駐トースト用Tkループの開始"""
    hotkey_up, hotkey_down, _, _ = load_config()
    keyboard.add_hotkey(hotkey_up, brightness_up, trigger_on_release=False)
    keyboard.add_hotkey(hotkey_down, brightness_down, trigger_on_release=False)
    threading.Thread(target=_rehook_watchdog, args=(hotkey_up, hotkey_down), daemon=True).start()
    threading.Thread(target=_adjust_worker, daemon=True).start()

    root = tk.Tk()
    root.withdraw()
    root.report_callback_exception = lambda exc_type, exc_value, tb: log_error("tk_callback", exc_value)
    toast_manager = _ToastManager(root)

    def poll_toast_queue():
        try:
            while True:
                brightness_percent = _toast_queue.get_nowait()
                toast_manager.show(brightness_percent)
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_toast_queue", e)
        root.after(50, poll_toast_queue)

    root.after(50, poll_toast_queue)
    root.mainloop()


class _HotkeyRecorder:
    """1つのCTkEntryに対する「クリックしてキーを押すと記録」の挙動をまとめたヘルパー。
    設定画面でホットキー入力欄が2つ(明るく/暗く)必要になったため、
    AudioSwitcherではその場のクロージャで書いていたロジックをクラスに切り出した。"""

    def __init__(self, entry, initial_value):
        self.entry = entry
        self.saved_value = initial_value
        self.temp_value = initial_value
        self.is_recording = False
        self.current_keys = set()

        entry.configure(state="normal")
        entry.insert(0, initial_value)
        entry.configure(state="readonly")
        entry.bind("<FocusIn>", self._on_focus_in)
        entry.bind("<FocusOut>", self._on_focus_out)
        entry.bind("<KeyPress>", self._on_key_press)
        entry.bind("<KeyRelease>", self._on_key_release)

    def _set_text(self, text, color=None):
        self.entry.configure(state="normal")
        self.entry.delete(0, "end")
        self.entry.insert(0, text)
        if color:
            self.entry.configure(text_color=color)
        self.entry.configure(state="readonly")

    def _on_focus_in(self, _event):
        self.is_recording = True
        self.current_keys.clear()
        self._set_text("Press Keys...", color=COLOR_ACCENT)

    def _on_focus_out(self, _event):
        self.is_recording = False
        self.entry.configure(text_color=COLOR_TEXT_MAIN)
        if self.temp_value == "Press Keys..." or not self.temp_value:
            self.temp_value = self.saved_value
        self._set_text(self.temp_value)

    def _on_key_press(self, event):
        if not self.is_recording:
            return "break"
        if event.keysym == "Escape":
            self.temp_value = self.saved_value
            self.entry.master.focus_set()
            return "break"
        key_map = {
            "Control_L": "ctrl", "Control_R": "ctrl", "Alt_L": "alt", "Alt_R": "alt",
            "Shift_L": "shift", "Shift_R": "shift", "Win_L": "win", "Win_R": "win",
            "Up": "up", "Down": "down", "Left": "left", "Right": "right",
        }
        key_name = key_map.get(event.keysym, event.keysym.lower())
        if "caps" in key_name or "lock" in key_name or "num_" in key_name:
            return "break"
        self.current_keys.add(key_name)

        ordered = []
        for modifier in ["ctrl", "alt", "shift", "win"]:
            if modifier in self.current_keys:
                ordered.append(modifier)
        for k in sorted(self.current_keys):
            if k not in ["ctrl", "alt", "shift", "win"]:
                ordered.append(k)
        self.temp_value = "+".join(ordered)
        self._set_text(self.temp_value)
        return "break"

    def _on_key_release(self, _event):
        if not self.is_recording:
            return "break"
        if self.temp_value and self.temp_value != "Press Keys...":
            self.is_recording = False
            self.entry.master.focus_set()
        return "break"

    @property
    def value(self):
        return self.temp_value


def open_setting_window():
    """2重起動防止付き設定UI。モニター一覧をスライダーで直接調整でき、
    明るく/暗くの2つのホットキーとステップ幅を設定できる。"""
    import customtkinter as ctk

    kill_previous_instances()

    snapshot = get_monitor_snapshot()
    saved_up, saved_down, saved_step, saved_startup = load_config()

    ctk.set_appearance_mode("dark")

    root = ctk.CTk(fg_color=COLOR_BG)
    root.title("BrightSwitch")
    root.geometry("560x100")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    if os.path.exists(ICON_PATH):
        try:
            root.iconbitmap(ICON_PATH)
        except Exception:
            pass

    header = ctk.CTkFrame(root, fg_color="transparent")
    header.pack(fill="x", padx=30, pady=(30, 18))
    ctk.CTkLabel(header, text="BrightSwitch", font=FONT_DISPLAY,
                 text_color=COLOR_TEXT_MAIN, anchor="w").pack(fill="x")
    ctk.CTkLabel(header, text="モニターの明るさをワンショートカットで調整する。",
                 font=FONT_BODY, text_color=COLOR_TEXT_SUB, anchor="w").pack(fill="x", pady=(2, 0))

    if not snapshot:
        warn = ctk.CTkFrame(root, fg_color=COLOR_PANEL, corner_radius=14,
                             border_width=1, border_color=COLOR_BORDER)
        warn.pack(fill="x", padx=30, pady=(0, 18))
        ctk.CTkLabel(
            warn, text="DDC/CI対応モニターが見つかりませんでした。\nモニターのOSD設定でDDC/CIを有効にしてください。",
            font=FONT_BODY, text_color=COLOR_TEXT_SUB, justify="left"
        ).pack(padx=20, pady=16)

    # モニター一覧 + 明るさスライダー
    monitor_panel = ctk.CTkFrame(root, fg_color=COLOR_PANEL, corner_radius=14,
                                  border_width=1, border_color=COLOR_BORDER)
    monitor_panel.pack(fill="x", padx=30, pady=(0, 18))

    ctk.CTkLabel(monitor_panel, text="MONITORS", font=FONT_LABEL, text_color=COLOR_TEXT_SUB,
                 anchor="w").pack(fill="x", padx=20, pady=(16, 4))

    def make_slider_row(parent, monitor):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=20, pady=6)

        name = monitor["name"]
        label = ctk.CTkLabel(row, text=name, font=FONT_BODY, text_color=COLOR_TEXT_MAIN,
                              anchor="w", width=180)
        label.pack(side="left")

        value_label = ctk.CTkLabel(row, text="--", font=FONT_MONO, text_color=COLOR_ACCENT, width=44)
        value_label.pack(side="right")

        current = monitor["brightness"] if monitor["brightness"] is not None else 50

        def on_slider_release(_value, index=monitor["index"]):
            try:
                new_value = set_monitor_brightness(index, slider.get())
                value_label.configure(text=f"{new_value}%")
            except Exception as e:
                log_error("slider_set_brightness", e)

        slider = ctk.CTkSlider(
            row, from_=0, to=100, number_of_steps=100,
            fg_color=COLOR_INPUT, progress_color=COLOR_ACCENT,
            button_color=COLOR_TEXT_MAIN, button_hover_color=COLOR_ACCENT_HOVER,
        )
        slider.pack(side="left", fill="x", expand=True, padx=12)
        slider.set(current)
        value_label.configure(text=f"{current}%")
        slider.bind("<ButtonRelease-1>", lambda e: on_slider_release(slider.get()))

    if snapshot:
        for m in snapshot:
            make_slider_row(monitor_panel, m)
    ctk.CTkLabel(monitor_panel, text="", height=4).pack()

    # ホットキー・ステップ幅
    container = ctk.CTkFrame(root, fg_color=COLOR_PANEL, corner_radius=14,
                              border_width=1, border_color=COLOR_BORDER)
    container.pack(fill="both", expand=True, padx=30, pady=0)

    def section_label(parent, text):
        ctk.CTkLabel(parent, text=text, font=FONT_LABEL, text_color=COLOR_TEXT_SUB,
                     anchor="w").pack(fill="x", padx=24, pady=(18, 6))

    entry_kwargs = dict(
        font=FONT_MONO, justify="center", state="readonly",
        fg_color=COLOR_INPUT, text_color=COLOR_TEXT_MAIN,
        border_width=0, corner_radius=8,
    )

    hotkey_row = ctk.CTkFrame(container, fg_color="transparent")
    hotkey_row.pack(fill="x", padx=24, pady=(0, 6))

    up_col = ctk.CTkFrame(hotkey_row, fg_color="transparent")
    up_col.pack(side="left", fill="x", expand=True, padx=(0, 8))
    ctk.CTkLabel(up_col, text="明るく", font=("Segoe UI", 8), text_color=COLOR_TEXT_FAINT,
                 anchor="w").pack(fill="x")
    up_box = ctk.CTkFrame(up_col, fg_color=COLOR_INPUT, corner_radius=8,
                           border_width=1, border_color=COLOR_BORDER)
    up_box.pack(fill="x", pady=(2, 0))
    entry_up = ctk.CTkEntry(up_box, **entry_kwargs)
    entry_up.pack(fill="x", padx=2, pady=2, ipady=6)

    down_col = ctk.CTkFrame(hotkey_row, fg_color="transparent")
    down_col.pack(side="left", fill="x", expand=True, padx=(8, 0))
    ctk.CTkLabel(down_col, text="暗く", font=("Segoe UI", 8), text_color=COLOR_TEXT_FAINT,
                 anchor="w").pack(fill="x")
    down_box = ctk.CTkFrame(down_col, fg_color=COLOR_INPUT, corner_radius=8,
                             border_width=1, border_color=COLOR_BORDER)
    down_box.pack(fill="x", pady=(2, 0))
    entry_down = ctk.CTkEntry(down_box, **entry_kwargs)
    entry_down.pack(fill="x", padx=2, pady=2, ipady=6)

    recorder_up = _HotkeyRecorder(entry_up, saved_up)
    recorder_down = _HotkeyRecorder(entry_down, saved_down)

    ctk.CTkLabel(container, text="クリックしてから新しいキーの組み合わせを押すと変更できます",
                 font=("Segoe UI", 8), text_color=COLOR_TEXT_FAINT, anchor="w").pack(
        fill="x", padx=24, pady=(4, 14))

    section_label(container, "STEP (1回のホットキーで変化させる%)")
    step_var = tk.StringVar(value=saved_step)
    step_entry = ctk.CTkEntry(
        container, textvariable=step_var, font=FONT_MONO, justify="center",
        fg_color=COLOR_INPUT, text_color=COLOR_TEXT_MAIN, border_width=1,
        border_color=COLOR_BORDER, corner_radius=8,
    )
    step_entry.pack(fill="x", padx=24, pady=(0, 20))

    # ボトム
    bottom_bar = ctk.CTkFrame(root, fg_color="transparent")
    bottom_bar.pack(fill="x", padx=30, pady=(18, 30))

    startup_var = tk.StringVar(value=saved_startup)
    switch_startup = ctk.CTkSwitch(
        bottom_bar, text="Start on boot", variable=startup_var, onvalue="1", offvalue="0",
        font=FONT_BODY, text_color=COLOR_TEXT_SUB,
        fg_color=COLOR_INPUT, progress_color=COLOR_ACCENT,
        button_color=COLOR_TEXT_MAIN, button_hover_color=COLOR_ACCENT_HOVER,
    )
    switch_startup.pack(side="left", anchor="w")

    btn_frame = ctk.CTkFrame(bottom_bar, fg_color="transparent")
    btn_frame.pack(side="right")

    def validated_step():
        try:
            step = int(step_var.get())
            if step <= 0 or step > 100:
                raise ValueError
            return str(step)
        except ValueError:
            tk.messagebox.showwarning("Invalid Step", "STEPには1〜100の整数を入力してください。")
            return None

    def on_save_only():
        step = validated_step()
        if step is None:
            return
        save_config(recorder_up.value, recorder_down.value, step, startup_var.get())
        root.destroy()
        sys.exit(0)

    btn_save_only = ctk.CTkButton(
        btn_frame, text="Save only", command=on_save_only,
        fg_color="transparent", hover_color=COLOR_INPUT, text_color=COLOR_TEXT_SUB,
        border_width=1, border_color=COLOR_BORDER, corner_radius=8,
        font=FONT_LABEL, width=100, height=36,
    )
    btn_save_only.pack(side="left", padx=(0, 10))

    def on_save_and_run():
        step = validated_step()
        if step is None:
            return
        save_config(recorder_up.value, recorder_down.value, step, startup_var.get())
        root.destroy()
        start_backend()

    btn_save_run = ctk.CTkButton(
        btn_frame, text="Save & Run", command=on_save_and_run,
        fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color=COLOR_BG,
        corner_radius=8, font=FONT_LABEL, width=120, height=36,
    )
    btn_save_run.pack(side="right")

    root.update_idletasks()
    required_height = root.winfo_reqheight() + 12
    root.geometry(f"560x{required_height}")

    root.mainloop()


if __name__ == "__main__":
    try:
        if "/background" in sys.argv:
            if os.path.exists(CONFIG_FILE):
                start_backend()
            else:
                open_setting_window()
        else:
            open_setting_window()
    except SystemExit:
        raise
    except Exception as e:
        log_error("main", e)
        raise
