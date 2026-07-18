import os
import subprocess
import tkinter as tk
from tkinter import messagebox
import time
import sys
import threading
import queue
import ctypes
import math
import winshell
from win32com.client import Dispatch

import pystray
from PIL import Image, ImageDraw
import screen_brightness_control as sbc

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "app_config.txt")

# ----------------------------------------------------------------------
# デザイントークン — AudioSwitcherシリーズと同じグラファイト×カッパー配色。
# ----------------------------------------------------------------------
COLOR_BG            = "#16151a"
COLOR_PANEL         = "#1d1b21"
COLOR_TRACK         = "#26232b"
COLOR_BORDER        = "#322e38"
COLOR_ACCENT        = "#d68a4c"
COLOR_ACCENT_HOVER  = "#e6a05f"
COLOR_TEXT_MAIN     = "#f2ece1"
COLOR_TEXT_SUB      = "#8f8878"
COLOR_TEXT_FAINT    = "#5c5850"

FONT_TITLE = ("Segoe UI Semibold", 12)
FONT_BODY  = ("Segoe UI", 10)
FONT_MONO  = ("Consolas", 10, "bold")

DEFAULT_SETTINGS = {
    "startup": "0",
    "large_slider": "0",
    "show_toast_after": "1",
    "show_number": "1",
    "delay_until_release": "0",  # 「止まるまで変更を遅らせる」。0=ドラッグ中も逐次反映、1=離した時だけ反映
    "link_sliders": "0",         # 「一斉に動かせるようにする」
}


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


# ----------------------------------------------------------------------
# モニター制御
# ----------------------------------------------------------------------

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


# ----------------------------------------------------------------------
# 設定ファイル(key=value形式)
# ----------------------------------------------------------------------

def load_config():
    settings = dict(DEFAULT_SETTINGS)
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or "=" not in line:
                        continue
                    key, _, value = line.partition("=")
                    settings[key.strip()] = value.strip()
        except Exception as e:
            log_error("load_config", e)
    return settings


def save_config(settings):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            for key, value in settings.items():
                f.write(f"{key}={value}\n")
    except Exception as e:
        log_error("save_config", e)


def manage_startup(enabled):
    """Windowsのスタートアップフォルダへの自動登録・解除"""
    try:
        startup_dir = winshell.startup()
        shortcut_path = os.path.join(startup_dir, "BrightSwitch.lnk")

        if enabled:
            if getattr(sys, 'frozen', False):
                target = sys.executable
                arguments = ""
                icon_target = target
            else:
                pythonw_path = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
                target = pythonw_path if os.path.exists(pythonw_path) else sys.executable
                arguments = f'"{os.path.abspath(__file__)}"'
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
        log_error("manage_startup", e)


# ----------------------------------------------------------------------
# 幾何ヘルパー(AudioSwitcherシリーズ共通)
# ----------------------------------------------------------------------

class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class _MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", _RECT),
                ("rcWork", _RECT), ("dwFlags", ctypes.c_ulong)]


def _active_monitor_work_area():
    """マウスカーソルがあるモニターの作業領域(タスクバーを除く)を返す"""
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


# ----------------------------------------------------------------------
# トレイアイコンの生成(太陽のグリフ。小さいサイズでも潰れないよう
# 装飾を絞ったフラットデザイン)
# ----------------------------------------------------------------------

def render_tray_icon():
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    cx, cy = size / 2, size / 2
    core_r = 15
    draw.ellipse([cx - core_r, cy - core_r, cx + core_r, cy + core_r], fill=(242, 236, 225, 255))

    ray_inner, ray_outer, ray_width = core_r + 6, core_r + 15, 5
    for i in range(8):
        angle = math.radians(i * 45)
        x1 = cx + ray_inner * math.cos(angle)
        y1 = cy + ray_inner * math.sin(angle)
        x2 = cx + ray_outer * math.cos(angle)
        y2 = cy + ray_outer * math.sin(angle)
        draw.line([x1, y1, x2, y2], fill=(242, 236, 225, 255), width=ray_width)
    return img


# ----------------------------------------------------------------------
# モダンなカスタムスライダー(Canvas描画・角丸トラック・アクセント色)
# ----------------------------------------------------------------------

class ModernSlider:
    def __init__(self, canvas, x, y, width, height, value=50, on_change=None, on_release=None):
        self.canvas = canvas
        self.x, self.y, self.width, self.height = x, y, width, height
        self.value = value
        self.on_change = on_change
        self.on_release = on_release
        self._dragging = False
        self._items = []

        canvas.tag_bind("bg", "<Enter>", lambda e: None)  # no-op anchor for tag namespace safety

        self._hit_id = canvas.create_rectangle(
            x, y, x + width, y + height, fill="", outline="", tags="slider_hit"
        )
        canvas.tag_bind(self._hit_id, "<Button-1>", self._on_press)
        canvas.tag_bind(self._hit_id, "<B1-Motion>", self._on_drag)
        canvas.tag_bind(self._hit_id, "<ButtonRelease-1>", self._on_release)

        self.redraw()

    def _value_from_x(self, event_x):
        track_x1 = self.x + self._thumb_r
        track_x2 = self.x + self.width - self._thumb_r
        ratio = (event_x - track_x1) / max(1, (track_x2 - track_x1))
        return max(0, min(100, ratio * 100))

    @property
    def _thumb_r(self):
        return self.height / 2 + 3

    def _on_press(self, event):
        self._dragging = True
        self.value = self._value_from_x(event.x)
        self.redraw()
        if self.on_change:
            self.on_change(self.value, False)

    def _on_drag(self, event):
        if not self._dragging:
            return
        self.value = self._value_from_x(event.x)
        self.redraw()
        if self.on_change:
            self.on_change(self.value, False)

    def _on_release(self, event):
        self._dragging = False
        if self.on_release:
            self.on_release(self.value)

    def set_value(self, value, redraw=True):
        self.value = max(0, min(100, value))
        if redraw:
            self.redraw()

    def redraw(self):
        for item in self._items:
            self.canvas.delete(item)
        self._items.clear()

        track_y1 = self.y + self.height / 2 - 4
        track_y2 = self.y + self.height / 2 + 4
        track_x1 = self.x + self._thumb_r
        track_x2 = self.x + self.width - self._thumb_r

        self._items.append(_rounded_rect(
            self.canvas, track_x1, track_y1, track_x2, track_y2, 4,
            fill=COLOR_TRACK, outline=""
        ))

        fill_x = track_x1 + (track_x2 - track_x1) * (self.value / 100)
        if fill_x > track_x1 + 1:
            self._items.append(_rounded_rect(
                self.canvas, track_x1, track_y1, fill_x, track_y2, 4,
                fill=COLOR_ACCENT, outline=""
            ))

        thumb_cx = fill_x
        thumb_cy = self.y + self.height / 2
        r = self._thumb_r
        self._items.append(self.canvas.create_oval(
            thumb_cx - r, thumb_cy - r, thumb_cx + r, thumb_cy + r,
            fill=COLOR_TEXT_MAIN, outline=COLOR_ACCENT, width=2
        ))

        # ヒット領域(操作対象)を最前面に上げ直す
        self.canvas.tag_raise(self._hit_id)


# ----------------------------------------------------------------------
# クリックで開くスライダーパネル
# ----------------------------------------------------------------------

class SliderPanel:
    ROW_HEIGHT = 46
    PADDING = 18

    def __init__(self, root):
        self.root = root
        self.window = None
        self.canvas = None
        self.sliders = []
        self.value_labels = []
        self.snapshot = []

    def _apply_brightness(self, monitor_index, value, is_final, link_sliders):
        settings = load_config()
        delay_mode = settings.get("delay_until_release", "0") == "1"

        if delay_mode and not is_final:
            return  # 「止まるまで変更を遅らせる」がONの間はドラッグ中は反映しない

        try:
            new_value = set_monitor_brightness(monitor_index, value)
        except Exception as e:
            log_error("panel_set_brightness", e)
            return

        if link_sliders:
            self._sync_linked_sliders(monitor_index, new_value, is_final, delay_mode)

        if is_final and settings.get("show_toast_after", "1") == "1":
            from_toast = getattr(self.root, "_toast_manager", None)
            if from_toast:
                from_toast.show(new_value)

    def _sync_linked_sliders(self, source_index, source_value, is_final, delay_mode):
        for m, slider, label in zip(self.snapshot, self.sliders, self.value_labels):
            if m["index"] == source_index:
                continue
            slider.set_value(source_value)
            if label is not None:
                label.configure(text=f"{int(round(source_value))}%")
            if not delay_mode or is_final:
                try:
                    set_monitor_brightness(m["index"], source_value)
                except Exception as e:
                    log_error("panel_set_brightness_linked", e)

    def toggle(self):
        if self.window is not None and self.window.winfo_viewable():
            self.close()
        else:
            self.open()

    def open(self):
        self.close()
        settings = load_config()
        large = settings.get("large_slider", "0") == "1"
        show_number = settings.get("show_number", "1") == "1"
        link_sliders = settings.get("link_sliders", "0") == "1"

        self.snapshot = get_monitor_snapshot()

        row_h = self.ROW_HEIGHT + (14 if large else 0)
        slider_h = 18 if large else 12
        width = 340
        rows = max(1, len(self.snapshot))
        height = self.PADDING * 2 + rows * row_h + (30 if not self.snapshot else 0)

        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(bg=COLOR_BORDER)

        outer = tk.Frame(win, bg=COLOR_BORDER)
        outer.pack(padx=1, pady=1)
        canvas = tk.Canvas(outer, width=width, height=height, bg=COLOR_PANEL,
                            highlightthickness=0)
        canvas.pack()

        self.window = win
        self.canvas = canvas
        self.sliders = []
        self.value_labels = []

        if not self.snapshot:
            canvas.create_text(
                width / 2, height / 2, text="DDC/CI対応モニターが見つかりません",
                fill=COLOR_TEXT_SUB, font=FONT_BODY
            )
        else:
            for i, m in enumerate(self.snapshot):
                row_y = self.PADDING + i * row_h
                name = m["name"]
                display_name = name if len(name) <= 22 else name[:20] + "..."
                canvas.create_text(
                    self.PADDING, row_y, anchor="nw", text=display_name,
                    fill=COLOR_TEXT_MAIN, font=FONT_BODY
                )

                value_label = None
                label_w = 0
                if show_number:
                    label_w = 40
                    value_label = tk.Label(
                        win, text=f"{m['brightness'] or 0}%", font=FONT_MONO,
                        fg=COLOR_ACCENT, bg=COLOR_PANEL
                    )
                    canvas.create_window(
                        width - self.PADDING, row_y + 20, anchor="ne", window=value_label
                    )

                slider_y = row_y + 22
                slider_width = width - self.PADDING * 2 - label_w - (8 if show_number else 0)

                def on_change(value, is_final, index=m["index"], label=value_label):
                    if label is not None:
                        label.configure(text=f"{int(round(value))}%")
                    self._apply_brightness(index, value, is_final, link_sliders)

                def on_release(value, index=m["index"], label=value_label):
                    if label is not None:
                        label.configure(text=f"{int(round(value))}%")
                    self._apply_brightness(index, value, True, link_sliders)

                slider = ModernSlider(
                    canvas, self.PADDING, slider_y, slider_width, slider_h,
                    value=m["brightness"] or 0, on_change=on_change, on_release=on_release
                )
                self.sliders.append(slider)
                self.value_labels.append(value_label)

        win.update_idletasks()
        try:
            left, top, right, bottom = _active_monitor_work_area()
        except Exception:
            right, bottom = win.winfo_screenwidth(), win.winfo_screenheight()
        win_w = win.winfo_reqwidth()
        win_h = win.winfo_reqheight()
        x = right - win_w - 8
        y = bottom - win_h - 8
        win.geometry(f"{win_w}x{win_h}+{x}+{y}")

        win.bind("<FocusOut>", lambda e: self.root.after(120, self._close_if_unfocused))
        win.focus_force()

    def _close_if_unfocused(self):
        if self.window is None:
            return
        try:
            focused = self.window.focus_get()
        except Exception:
            focused = None
        if focused is None:
            self.close()

    def close(self):
        if self.window is not None:
            try:
                self.window.destroy()
            except Exception:
                pass
            self.window = None
            self.canvas = None
            self.sliders = []
            self.value_labels = []


# ----------------------------------------------------------------------
# トースト通知(スライダー操作後に現在値を表示。AudioSwitcherと同じ仕組み)
# ----------------------------------------------------------------------

class ToastManager:
    WIDTH, HEIGHT = 260, 60

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

    def _place(self):
        try:
            left, top, right, bottom = _active_monitor_work_area()
        except Exception:
            left, top = 0, 0
            right = self.window.winfo_screenwidth()
            bottom = self.window.winfo_screenheight()
        x = right - self.WIDTH - 8
        y = bottom - self.HEIGHT - 8
        self.window.geometry(f"{self.WIDTH}x{self.HEIGHT}+{x}+{y}")

    def show(self, brightness_percent):
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None

        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 14,
                      fill=COLOR_PANEL, outline=COLOR_BORDER, width=1)
        canvas.create_rectangle(0, 10, 4, self.HEIGHT - 10, fill=COLOR_ACCENT, outline="")
        canvas.create_text(20, 18, anchor="w", text="BRIGHTNESS",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))
        canvas.create_text(20, 38, anchor="w", text=f"{int(round(brightness_percent))}%",
                            fill=COLOR_TEXT_MAIN, font=("Segoe UI", 13, "bold"))

        self._place()
        self.window.deiconify()
        self._place()
        self.current_alpha = 0.0
        self._fade_in()

    def _fade_in(self):
        if self.current_alpha < 0.96:
            self.current_alpha += 0.14
            self.window.attributes("-alpha", min(self.current_alpha, 0.96))
            self.fade_job = self.root.after(15, self._fade_in)
        else:
            self.fade_job = self.root.after(1600, self._fade_out)

    def _fade_out(self):
        if self.current_alpha > 0.0:
            self.current_alpha -= 0.07
            self.window.attributes("-alpha", max(self.current_alpha, 0.0))
            self.fade_job = self.root.after(15, self._fade_out)
        else:
            self.window.withdraw()
            self.fade_job = None


# ----------------------------------------------------------------------
# トレイアイコン + 右クリックメニュー
#
# pystrayのアイコンはメッセージポンプを持つ独自のイベントループを回すため、
# Tkinterのmainloopとは別スレッドで動かす。トレイ側のコールバックは
# pystrayのスレッドで呼ばれるので、Tk操作はすべてキュー経由でmainloopに戻す
# (AudioSwitcherでホットキーのコールバックを軽量に保った時と同じ考え方)。
# ----------------------------------------------------------------------

_ui_action_queue = queue.Queue()


def _build_menu():
    settings = load_config()

    def is_on(key):
        return load_config().get(key, "0") == "1"

    def toggle(key):
        def _handler(icon, item):
            current = load_config()
            current[key] = "0" if current.get(key, "0") == "1" else "1"
            save_config(current)
            if key == "startup":
                manage_startup(current[key] == "1")
        return _handler

    def on_open_panel(icon, item):
        _ui_action_queue.put(("open_panel", None))

    def on_quit(icon, item):
        _ui_action_queue.put(("quit", None))

    return pystray.Menu(
        pystray.MenuItem("open", on_open_panel, default=True, visible=False),
        pystray.MenuItem("サインイン時に起動する", toggle("startup"), checked=lambda i: is_on("startup")),
        pystray.MenuItem("大きいスライダーを使う", toggle("large_slider"), checked=lambda i: is_on("large_slider")),
        pystray.MenuItem("調整後の明るさを表示する", toggle("show_toast_after"), checked=lambda i: is_on("show_toast_after")),
        pystray.MenuItem("現在の数字を表示する", toggle("show_number"), checked=lambda i: is_on("show_number")),
        pystray.MenuItem("止まるまで変更を遅らせる", toggle("delay_until_release"), checked=lambda i: is_on("delay_until_release")),
        pystray.MenuItem("一斉に動かせるようにする", toggle("link_sliders"), checked=lambda i: is_on("link_sliders")),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("終了", on_quit),
    )


def run_app():
    kill_previous_instances()

    settings = load_config()
    if not os.path.exists(CONFIG_FILE):
        save_config(settings)

    root = tk.Tk()
    root.withdraw()
    root.report_callback_exception = lambda exc_type, exc_value, tb: log_error("tk_callback", exc_value)

    toast_manager = ToastManager(root)
    root._toast_manager = toast_manager  # SliderPanelから参照するため
    panel = SliderPanel(root)

    tray_icon = pystray.Icon("BrightSwitch", render_tray_icon(), "BrightSwitch", _build_menu())

    def poll_ui_queue():
        try:
            while True:
                action, payload = _ui_action_queue.get_nowait()
                if action == "open_panel":
                    panel.toggle()
                elif action == "quit":
                    tray_icon.stop()
                    root.after(50, root.destroy)
        except queue.Empty:
            pass
        root.after(50, poll_ui_queue)

    root.after(50, poll_ui_queue)

    threading.Thread(target=tray_icon.run, daemon=True).start()
    root.mainloop()


if __name__ == "__main__":
    try:
        run_app()
    except SystemExit:
        raise
    except Exception as e:
        log_error("main", e)
        raise
