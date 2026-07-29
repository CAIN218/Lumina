import os
import subprocess
import tkinter as tk
import time
import sys
import threading
import queue
import ctypes
import math
import json
import urllib.request
import webbrowser
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

__version__ = "1.0.1"
GITHUB_REPO = "CAIN218/Lumina"

# ----------------------------------------------------------------------
# デザイントークン — 「Lumina(光)」らしい、深い紫紺の背景に暖色の
# グロー(光暈)が浮かぶ配色。AudioSwitcherシリーズの渋いグラファイト×
# カッパーとは意図的に differentiate し、単体ブランドとして今風に。
# ----------------------------------------------------------------------
COLOR_PANEL_BG      = "#1b1531"
COLOR_PANEL_BG_ALT  = "#231a3f"  # マスターフェーダー行のハイライト用
COLOR_PANEL_BORDER  = "#4a3d78"
COLOR_TRACK         = "#2c2450"
COLOR_ACCENT        = "#ffb454"
COLOR_ACCENT_HOVER  = "#ffcf8c"
COLOR_ACCENT_SOFT   = "#7c6bb8"  # サブアクセント(マスターの目盛りなどに使用)
COLOR_TEXT_MAIN     = "#f8f6ff"
COLOR_TEXT_SUB      = "#a79fd6"
COLOR_TEXT_FAINT    = "#655a99"
COLOR_TRANSPARENT_KEY = "#0a090f"  # ウィンドウの角丸を出すための透過キー色

FONT_TITLE = ("Segoe UI Semibold", 9)
FONT_BODY  = ("Segoe UI", 10)
FONT_MONO  = ("Consolas", 9, "bold")

DEFAULT_SETTINGS = {
    "startup": "0",
    "show_toast_after": "0",     # 「調整後の明るさを表示する」。既定でOFF
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


def _parse_version(tag):
    """'v1.2.3' -> (1, 2, 3)。数値以外が混ざっていても比較不能で落ちない
    ように、パースできない要素は0として扱う。"""
    parts = []
    for p in tag.lstrip("vV").split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts)


_update_queue = queue.Queue()

def _update_check_loop():
    """1日1回、GitHubの最新リリースを確認する。バックグラウンドの常駐
    ツールなので、失敗(オフライン等)は完全に握りつぶし通常動作に一切
    影響させない。新しいバージョンが見つかった時だけ、トースト表示用の
    キューに積んでメインスレッド側(poll_ui_queue)に伝える。"""
    while True:
        try:
            req = urllib.request.Request(
                f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
                headers={"Accept": "application/vnd.github+json", "User-Agent": "Lumina-update-check"},
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            latest_tag = data.get("tag_name", "")
            release_url = data.get("html_url") or f"https://github.com/{GITHUB_REPO}/releases/latest"
            if latest_tag and _parse_version(latest_tag) > _parse_version(__version__):
                _update_queue.put((latest_tag, release_url))
        except Exception as e:
            log_error("update_check", e)
        time.sleep(24 * 60 * 60)


def enable_dpi_awareness():
    """Per-Monitor DPIを意識させる。これを呼ばないとTkinterの描画がOSに
    よって拡大縮小され(見た目が滲む/荒くなる)、かつマルチモニター環境で
    GetCursorPos等の物理ピクセル座標とTk側の論理座標がズレて、パネルが
    意図しない位置(左上など)に出ることがある。exeとしてPyInstallerで
    固めた場合、マニフェストでのDPI宣言がないためコード側で明示する。"""
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception as e:
            log_error("enable_dpi_awareness", e)


def kill_previous_instances():
    """裏で既に動いている古いLuminaのプロセスをすべて安全に終了させる"""
    try:
        current_pid = os.getpid()
        target_process = "Lumina.exe"
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
# DDC/CI書き込みの非同期化(スライダーのドラッグを軽くする)
#
# DDC/CI通信はI2C経由で1回あたり数十〜百ms以上かかることがある。UIの
# ドラッグイベントの中で直接sbc.set_brightness()を呼ぶと、特に「一斉に
# 動かせるようにする」でモニター台数分まとめて呼んだ時に、ドラッグ操作
# そのものがカクつく(重い)。ここでは「各モニターについて直近の目標値
# だけを保持し、別スレッドが空き次第それを書き込む」方式にして、UI側の
# 描画・入力処理は一切ブロックしないようにする。
# ----------------------------------------------------------------------

_pending_lock = threading.Lock()
_pending_values = {}
_pending_wake = threading.Event()


def request_brightness(index, value):
    """このモニターの目標値を更新する(即座には書き込まない)"""
    with _pending_lock:
        _pending_values[index] = value
    _pending_wake.set()


def _brightness_writer_worker():
    while True:
        _pending_wake.wait()
        with _pending_lock:
            items = list(_pending_values.items())
            _pending_values.clear()
            _pending_wake.clear()
        for index, value in items:
            try:
                set_monitor_brightness(index, value)
            except Exception as e:
                log_error(f"brightness_writer[{index}]", e)


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
        shortcut_path = os.path.join(startup_dir, "Lumina.lnk")

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
# 幾何ヘルパー
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
    """マウスカーソルがあるモニターの作業領域(タスクバーを除く)を返す。
    失敗した場合は例外を投げる(呼び出し側でフォールバックする)。"""
    pt = _POINT()
    if not ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
        raise OSError("GetCursorPos failed")
    MONITOR_DEFAULTTONEAREST = 2
    hmonitor = ctypes.windll.user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
    if not hmonitor:
        raise OSError("MonitorFromPoint failed")
    info = _MONITORINFO()
    info.cbSize = ctypes.sizeof(_MONITORINFO)
    if not ctypes.windll.user32.GetMonitorInfoW(hmonitor, ctypes.byref(info)):
        raise OSError("GetMonitorInfoW failed")
    rc = info.rcWork
    if rc.right <= rc.left or rc.bottom <= rc.top:
        raise OSError(f"invalid work area rect: {rc.left},{rc.top},{rc.right},{rc.bottom}")
    return rc.left, rc.top, rc.right, rc.bottom


def _safe_work_area(fallback_window):
    """_active_monitor_work_area()が失敗した場合、少なくとも実際に存在する
    画面範囲に収まるようプライマリモニターの座標にフォールバックする
    (0,0起点で決め打ちにすると本当に左上に出てしまうため、それも避ける)。"""
    try:
        return _active_monitor_work_area()
    except Exception as e:
        log_error("work_area_fallback", e)
        w = fallback_window.winfo_screenwidth()
        h = fallback_window.winfo_screenheight()
        return 0, 0, w, h


def _rounded_rect(cv, x1, y1, x2, y2, r, **kwargs):
    points = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return cv.create_polygon(points, smooth=True, **kwargs)


def _lerp_color(c1, c2, t):
    """2つの#rrggbb色をtで線形補間する。透明度が使えないCanvas上で
    "薄いグロー"を表現するため、背景色に向かってブレンドする用途に使う。"""
    t = max(0.0, min(1.0, t))
    r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
    r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
    r = round(r1 + (r2 - r1) * t)
    g = round(g1 + (g2 - g1) * t)
    b = round(b1 + (b2 - b1) * t)
    return f"#{r:02x}{g:02x}{b:02x}"


def _ease_out_cubic(t):
    return 1 - (1 - t) ** 3


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
# モダンなカスタムスライダー(Canvas描画・細身のトラック・
# ドラッグ中はつまみの周りにグロー(光暈)が浮かぶ)
# ----------------------------------------------------------------------

class ModernSlider:
    TRACK_THICKNESS = 4
    THUMB_R = 7
    THUMB_R_ACTIVE = 8

    def __init__(self, canvas, x, y, width, height, value=50, on_change=None, on_release=None,
                 panel_bg=COLOR_PANEL_BG):
        self.canvas = canvas
        self.x, self.y, self.width, self.height = x, y, width, height
        self.value = value
        self.on_change = on_change
        self.on_release = on_release
        self.panel_bg = panel_bg
        self._dragging = False
        self._items = []

        self._hit_id = canvas.create_rectangle(
            x, y, x + width, y + height, fill="", outline=""
        )
        canvas.tag_bind(self._hit_id, "<Button-1>", self._on_press)
        canvas.tag_bind(self._hit_id, "<B1-Motion>", self._on_drag)
        canvas.tag_bind(self._hit_id, "<ButtonRelease-1>", self._on_release)

        self.redraw()

    def _value_from_x(self, event_x):
        track_x1 = self.x + self.THUMB_R_ACTIVE
        track_x2 = self.x + self.width - self.THUMB_R_ACTIVE
        ratio = (event_x - track_x1) / max(1, (track_x2 - track_x1))
        return max(0, min(100, ratio * 100))

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
        self.redraw()
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

        t = self.TRACK_THICKNESS / 2
        track_y1 = self.y + self.height / 2 - t
        track_y2 = self.y + self.height / 2 + t
        track_x1 = self.x + self.THUMB_R_ACTIVE
        track_x2 = self.x + self.width - self.THUMB_R_ACTIVE

        self._items.append(_rounded_rect(
            self.canvas, track_x1, track_y1, track_x2, track_y2, t,
            fill=COLOR_TRACK, outline=""
        ))

        fill_x = track_x1 + (track_x2 - track_x1) * (self.value / 100)
        if fill_x > track_x1 + 1:
            self._items.append(_rounded_rect(
                self.canvas, track_x1, track_y1, fill_x, track_y2, t,
                fill=COLOR_ACCENT, outline=""
            ))

        thumb_cx = fill_x
        thumb_cy = self.y + self.height / 2

        if self._dragging:
            # ドラッグ中だけ、つまみの周りにグロー(光暈)を薄く重ねて表示する。
            # Canvasは半透明塗りを扱えないため、背景色に向けて色をブレンド
            # した同心円を重ねることで擬似的なグローを表現している。
            for ring_r, mix in ((16, 0.72), (12, 0.5)):
                glow_color = _lerp_color(COLOR_ACCENT, self.panel_bg, mix)
                self._items.append(self.canvas.create_oval(
                    thumb_cx - ring_r, thumb_cy - ring_r, thumb_cx + ring_r, thumb_cy + ring_r,
                    fill=glow_color, outline=""
                ))

        r = self.THUMB_R_ACTIVE if self._dragging else self.THUMB_R
        fill = COLOR_ACCENT_HOVER if self._dragging else COLOR_TEXT_MAIN
        self._items.append(self.canvas.create_oval(
            thumb_cx - r, thumb_cy - r, thumb_cx + r, thumb_cy + r,
            fill=fill, outline=COLOR_ACCENT, width=2
        ))

        self.canvas.tag_raise(self._hit_id)


# ----------------------------------------------------------------------
# クリックで開くスライダーパネル
# ----------------------------------------------------------------------

class SliderPanel:
    ROW_HEIGHT = 44
    HEADER_HEIGHT = 34
    PADDING = 18
    WIDTH = 320
    CORNER_RADIUS = 16
    ANIM_DURATION_MS = 190
    ANIM_SLIDE_PX = 14

    def __init__(self, root):
        self.root = root
        self.window = None
        self.canvas = None
        self.sliders = []
        self.value_labels = []
        self.master_slider = None
        self.master_label = None
        self.snapshot = []
        self._anim_job = None
        self._anim_start = 0
        self._final_x = 0
        self._final_y = 0

    def _apply_brightness(self, monitor_index, value, is_final):
        """個別モニターのスライダーの値を反映する。「一斉に動かせるように
        する」がONの時でも、個別スライダーは常にそのモニターだけを独立して
        動かす(連動はマスターフェーダー経由でのみ行う)。"""
        settings = load_config()
        delay_mode = settings.get("delay_until_release", "0") == "1"

        if delay_mode and not is_final:
            return  # 「止まるまで変更を遅らせる」がONの間はドラッグ中は反映しない

        request_brightness(monitor_index, value)

        if is_final and settings.get("show_toast_after", "0") == "1":
            toast_manager = getattr(self.root, "_toast_manager", None)
            if toast_manager:
                toast_manager.show(value)

    def _apply_master(self, value, is_final):
        settings = load_config()
        delay_mode = settings.get("delay_until_release", "0") == "1"
        if delay_mode and not is_final:
            return

        for m, slider, label in zip(self.snapshot, self.sliders, self.value_labels):
            slider.set_value(value)
            if label is not None:
                label.configure(text=f"{int(round(value))}%")
            request_brightness(m["index"], value)

        if is_final and settings.get("show_toast_after", "0") == "1":
            toast_manager = getattr(self.root, "_toast_manager", None)
            if toast_manager:
                toast_manager.show(value)

    def toggle(self):
        if self.window is not None and self.window.winfo_viewable():
            self.close()
        else:
            self.open()

    def open(self):
        self.close()
        settings = load_config()
        show_number = settings.get("show_number", "1") == "1"
        link_sliders = settings.get("link_sliders", "0") == "1"

        self.snapshot = get_monitor_snapshot()
        has_master = link_sliders and len(self.snapshot) > 1

        row_h = self.ROW_HEIGHT
        width = self.WIDTH
        rows = max(1, len(self.snapshot))
        body_h = rows * row_h if self.snapshot else 50
        master_h = row_h + 6 if has_master else 0
        height = self.HEADER_HEIGHT + master_h + body_h + self.PADDING

        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.attributes("-alpha", 0.0)
        win.configure(bg=COLOR_TRANSPARENT_KEY)
        win.attributes("-transparentcolor", COLOR_TRANSPARENT_KEY)

        canvas = tk.Canvas(win, width=width, height=height, bg=COLOR_TRANSPARENT_KEY,
                            highlightthickness=0)
        canvas.pack()

        self.window = win
        self.canvas = canvas
        self.sliders = []
        self.value_labels = []
        self.master_slider = None
        self.master_label = None

        # 角丸パネル本体
        _rounded_rect(canvas, 0, 0, width, height, self.CORNER_RADIUS,
                      fill=COLOR_PANEL_BG, outline=COLOR_PANEL_BORDER, width=1.2)

        # ヘッダー(タイトル + 閉じるボタン)
        canvas.create_text(
            self.PADDING, self.HEADER_HEIGHT / 2, anchor="w", text="LUMINA",
            fill=COLOR_TEXT_SUB, font=("Segoe UI", 8, "bold")
        )
        close_cx, close_cy, close_r = width - self.PADDING, self.HEADER_HEIGHT / 2, 9
        close_hit = canvas.create_oval(
            close_cx - close_r, close_cy - close_r, close_cx + close_r, close_cy + close_r,
            fill="", outline=""
        )
        close_line1 = canvas.create_line(
            close_cx - 4, close_cy - 4, close_cx + 4, close_cy + 4,
            fill=COLOR_TEXT_SUB, width=1.4
        )
        close_line2 = canvas.create_line(
            close_cx - 4, close_cy + 4, close_cx + 4, close_cy - 4,
            fill=COLOR_TEXT_SUB, width=1.4
        )
        for item in (close_hit, close_line1, close_line2):
            canvas.tag_bind(item, "<Button-1>", lambda e: self.close())
        canvas.tag_bind(close_hit, "<Enter>", lambda e: (
            canvas.itemconfig(close_line1, fill=COLOR_TEXT_MAIN),
            canvas.itemconfig(close_line2, fill=COLOR_TEXT_MAIN),
        ))
        canvas.tag_bind(close_hit, "<Leave>", lambda e: (
            canvas.itemconfig(close_line1, fill=COLOR_TEXT_SUB),
            canvas.itemconfig(close_line2, fill=COLOR_TEXT_SUB),
        ))
        canvas.create_line(
            self.PADDING * 0.4, self.HEADER_HEIGHT, width - self.PADDING * 0.4, self.HEADER_HEIGHT,
            fill=COLOR_PANEL_BORDER
        )

        cursor_y = self.HEADER_HEIGHT

        if not self.snapshot:
            canvas.create_text(
                width / 2, self.HEADER_HEIGHT + body_h / 2,
                text="DDC/CI対応モニターが\n見つかりません",
                fill=COLOR_TEXT_SUB, font=FONT_BODY, justify="center"
            )
        else:
            if has_master:
                master_row_y = cursor_y + 8
                _rounded_rect(
                    canvas, self.PADDING * 0.4, cursor_y + 2, width - self.PADDING * 0.4,
                    cursor_y + master_h - 4, 10, fill=COLOR_PANEL_BG_ALT, outline=""
                )
                canvas.create_text(
                    self.PADDING, master_row_y, anchor="nw", text="すべてのモニター",
                    fill=COLOR_ACCENT, font=("Segoe UI", 10, "bold")
                )

                master_value = round(
                    sum((m["brightness"] or 0) for m in self.snapshot) / len(self.snapshot)
                )
                master_label = None
                label_w = 0
                if show_number:
                    label_w = 34
                    master_label = tk.Label(
                        win, text=f"{master_value}%", font=FONT_MONO,
                        fg=COLOR_ACCENT, bg=COLOR_PANEL_BG_ALT
                    )
                    canvas.create_window(
                        width - self.PADDING, master_row_y + 17, anchor="ne", window=master_label
                    )

                slider_width = width - self.PADDING * 2 - label_w - (6 if show_number else 0)

                def on_master_change(value, is_final):
                    if master_label is not None:
                        master_label.configure(text=f"{int(round(value))}%")
                    self._apply_master(value, is_final)

                def on_master_release(value):
                    if master_label is not None:
                        master_label.configure(text=f"{int(round(value))}%")
                    self._apply_master(value, True)

                self.master_slider = ModernSlider(
                    canvas, self.PADDING, master_row_y + 18, slider_width, 14,
                    value=master_value, on_change=on_master_change, on_release=on_master_release,
                    panel_bg=COLOR_PANEL_BG_ALT,
                )
                self.master_label = master_label
                cursor_y += master_h

            for i, m in enumerate(self.snapshot):
                row_y = cursor_y + i * row_h + 10
                name = m["name"]
                display_name = name if len(name) <= 24 else name[:22] + "..."
                canvas.create_text(
                    self.PADDING, row_y, anchor="nw", text=display_name,
                    fill=COLOR_TEXT_MAIN, font=FONT_BODY
                )

                value_label = None
                label_w = 0
                if show_number:
                    label_w = 34
                    value_label = tk.Label(
                        win, text=f"{m['brightness'] or 0}%", font=FONT_MONO,
                        fg=COLOR_ACCENT, bg=COLOR_PANEL_BG
                    )
                    canvas.create_window(
                        width - self.PADDING, row_y + 17, anchor="ne", window=value_label
                    )

                slider_y = row_y + 18
                slider_width = width - self.PADDING * 2 - label_w - (6 if show_number else 0)

                def on_change(value, is_final, index=m["index"], label=value_label):
                    if label is not None:
                        label.configure(text=f"{int(round(value))}%")
                    self._apply_brightness(index, value, is_final)

                def on_release(value, index=m["index"], label=value_label):
                    if label is not None:
                        label.configure(text=f"{int(round(value))}%")
                    self._apply_brightness(index, value, True)

                slider = ModernSlider(
                    canvas, self.PADDING, slider_y, slider_width, 14,
                    value=m["brightness"] or 0, on_change=on_change, on_release=on_release
                )
                self.sliders.append(slider)
                self.value_labels.append(value_label)

        win.update_idletasks()
        left, top, right, bottom = _safe_work_area(win)
        win_w = win.winfo_reqwidth()
        win_h = win.winfo_reqheight()
        self._final_x = max(left, right - win_w - 8)
        self._final_y = max(top, bottom - win_h - 8)
        win.geometry(f"{win_w}x{win_h}+{self._final_x}+{self._final_y + self.ANIM_SLIDE_PX}")

        win.bind("<FocusOut>", lambda e: self.root.after(120, self._close_if_unfocused))
        win.focus_force()

        self._anim_start = time.time()
        self._animate_open()

    def _animate_open(self):
        if self.window is None:
            return
        elapsed_ms = (time.time() - self._anim_start) * 1000
        t = min(1.0, elapsed_ms / self.ANIM_DURATION_MS)
        eased = _ease_out_cubic(t)

        alpha = eased
        y = round(self._final_y + self.ANIM_SLIDE_PX * (1 - eased))
        try:
            self.window.attributes("-alpha", alpha)
            self.window.geometry(f"+{self._final_x}+{y}")
        except Exception:
            return

        if t < 1.0:
            self._anim_job = self.root.after(8, self._animate_open)

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
        if self._anim_job is not None:
            try:
                self.root.after_cancel(self._anim_job)
            except Exception:
                pass
            self._anim_job = None
        if self.window is not None:
            try:
                self.window.destroy()
            except Exception:
                pass
            self.window = None
            self.canvas = None
            self.sliders = []
            self.value_labels = []
            self.master_slider = None
            self.master_label = None


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
        self.on_click = None
        self._hold_ms = 1600

    def _ensure_window(self):
        if self.window is not None:
            return
        toast = tk.Toplevel(self.root)
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)
        toast.attributes("-alpha", 0.0)

        toast.configure(bg=COLOR_TRANSPARENT_KEY)
        toast.attributes("-transparentcolor", COLOR_TRANSPARENT_KEY)

        canvas = tk.Canvas(toast, width=self.WIDTH, height=self.HEIGHT,
                            bg=COLOR_TRANSPARENT_KEY, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        # クリック可能なトースト(アップデート通知など)用。show()側で都度差し替える。
        canvas.bind("<Button-1>", lambda e: self.on_click() if self.on_click else None)
        toast.withdraw()

        self.window = toast
        self.canvas = canvas

    def _place(self):
        left, top, right, bottom = _safe_work_area(self.window)
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
                      fill=COLOR_PANEL_BG, outline=COLOR_PANEL_BORDER, width=1)
        canvas.create_rectangle(0, 10, 4, self.HEIGHT - 10, fill=COLOR_ACCENT, outline="")
        canvas.create_text(20, 18, anchor="w", text="LUMINA",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))
        canvas.create_text(20, 38, anchor="w", text=f"{int(round(brightness_percent))}%",
                            fill=COLOR_TEXT_MAIN, font=("Segoe UI", 13, "bold"))
        canvas.configure(cursor="")

        self.on_click = None
        self._hold_ms = 1600
        self._place()
        self.window.deiconify()
        self._place()
        self.current_alpha = 0.0
        self._fade_in()

    def show_update_available(self, latest_tag, release_url):
        """新しいバージョンが見つかった時の控えめな通知。クリックでリリース
        ページを開く。頻繁に出るものではないので、通常のトーストより長めに
        表示したままにする。"""
        self._ensure_window()
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None

        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 14,
                      fill=COLOR_PANEL_BG, outline=COLOR_PANEL_BORDER, width=1)
        canvas.create_rectangle(0, 10, 4, self.HEIGHT - 10, fill=COLOR_ACCENT, outline="")
        canvas.create_text(20, 18, anchor="w", text="UPDATE AVAILABLE",
                            fill=COLOR_ACCENT, font=("Segoe UI", 8, "bold"))
        canvas.create_text(20, 38, anchor="w", text=f"{latest_tag} — click to view",
                            fill=COLOR_TEXT_MAIN, font=("Segoe UI", 10, "bold"))
        canvas.configure(cursor="hand2")

        self.on_click = lambda: webbrowser.open(release_url)
        self._hold_ms = 8000
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
            self.fade_job = self.root.after(self._hold_ms, self._fade_out)

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

    def on_open_taskbar_settings(icon, item):
        # アプリ自身がタスクバーへの「常時表示」を強制するAPIはWindowsには
        # 無い(Vista以降、ユーザーがWindows設定側で明示的に選ぶ仕様)。
        # そのため設定画面を直接開くショートカットを用意している。
        try:
            os.startfile("ms-settings:taskbar")
        except Exception as e:
            log_error("open_taskbar_settings", e)

    def on_quit(icon, item):
        _ui_action_queue.put(("quit", None))

    return pystray.Menu(
        pystray.MenuItem("open", on_open_panel, default=True, visible=False),
        pystray.MenuItem("サインイン時に起動する", toggle("startup"), checked=lambda i: is_on("startup")),
        pystray.MenuItem("調整後の明るさを表示する", toggle("show_toast_after"), checked=lambda i: is_on("show_toast_after")),
        pystray.MenuItem("現在の数字を表示する", toggle("show_number"), checked=lambda i: is_on("show_number")),
        pystray.MenuItem("止まるまで変更を遅らせる", toggle("delay_until_release"), checked=lambda i: is_on("delay_until_release")),
        pystray.MenuItem("一斉に動かせるようにする", toggle("link_sliders"), checked=lambda i: is_on("link_sliders")),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("タスクバーの表示設定を開く...", on_open_taskbar_settings),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("終了", on_quit),
    )


def run_app():
    enable_dpi_awareness()
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

    tray_icon = pystray.Icon("Lumina", render_tray_icon(), "Lumina", _build_menu())

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

        try:
            while True:
                latest_tag, release_url = _update_queue.get_nowait()
                toast_manager.show_update_available(latest_tag, release_url)
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_update_queue", e)

        root.after(50, poll_ui_queue)

    root.after(50, poll_ui_queue)

    threading.Thread(target=tray_icon.run, daemon=True).start()
    threading.Thread(target=_brightness_writer_worker, daemon=True).start()
    threading.Thread(target=_update_check_loop, daemon=True).start()
    root.mainloop()


if __name__ == "__main__":
    try:
        run_app()
    except SystemExit:
        raise
    except Exception as e:
        log_error("main", e)
        raise
