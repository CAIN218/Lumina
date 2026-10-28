import os
import subprocess
import tkinter as tk
from tkinter import font as tkfont
import time
import sys
import threading
import queue
import ctypes
import math
import json
import tempfile
import zipfile
import winreg
import urllib.request
import urllib.parse
import hashlib
import urllib.error
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

__version__ = "1.1.0"
GITHUB_REPO = "CAIN218/Lumina"

# ----------------------------------------------------------------------
# デザイントークン — Yuga's lab のアプリアイコン(暗い背景・白い線・差し色1色)
# に合わせ、AudioSwitcherと同じ「暗い墨色×オフホワイト×差し色1色」の配色に
# 揃える。Luminaの差し色はアイコンの太陽と同じアンバー。
# ----------------------------------------------------------------------
COLOR_BG            = "#16151f"  # アイコンの背景色と同じ
COLOR_PANEL_BG      = "#1d1c26"  # パネル本体
COLOR_PANEL_BG_ALT  = "#25232f"  # マスターフェーダー行・ボタンのハイライト用
COLOR_PANEL_BORDER  = "#34313d"
COLOR_TRACK         = "#2e2c38"
COLOR_ACCENT        = "#e0a85a"  # アンバー(アイコンの太陽と同じ)
COLOR_ACCENT_HOVER  = "#ecbf80"
COLOR_ACCENT_DIM    = "#5c4528"  # アクセントの控えめ版
COLOR_TEXT_MAIN     = "#ece8e1"  # アイコンの白線に合わせたオフホワイト
COLOR_TEXT_SUB      = "#a29c90"
COLOR_TEXT_FAINT    = "#68645c"
COLOR_TRANSPARENT_KEY = "#010203"  # ウィンドウの角丸を出すための透過キー色(UIでは使わない色)

# ----------------------------------------------------------------------
# 表示言語 — 英語(既定)と日本語を右クリックメニューから切り替えられる。
# 画面上の文言はすべて L("English", "日本語") の形で両方を並べて書き、
# 表示する瞬間に選ぶ(パネルやメニューは開くたびに描き直すので即座に反映される)。
# ----------------------------------------------------------------------
LANG = "en"


def L(en, ja):
    """今の表示言語に合わせて、英語/日本語のどちらかを返す"""
    return ja if LANG == "ja" else en


# 画面には日本語と英数字が混ざる。Segoe UIには日本語の字形が無く、日本語部分
# だけ別フォントに置き換わって太さや高さが揃わないため、AudioSwitcherと同じく
# 両方を同じ字形で描ける「Yu Gothic UI」を優先し、無ければ順に代わりを使う。
FONT_CANDIDATES = ("Yu Gothic UI", "Meiryo UI", "Segoe UI", "Noto Sans CJK JP")
FONT_FAMILY = FONT_CANDIDATES[0]


def F(size, weight="normal"):
    """UI全体で使うフォント指定を返す(ファミリーはresolve_font_familyで決まる)"""
    return (FONT_FAMILY, size, weight)


def resolve_font_family(root):
    """インストールされているフォントを見て、使うファミリーを決める。
    Tkのルートが無いとフォント一覧を取れないので、ルート作成直後に呼ぶ。"""
    global FONT_FAMILY
    try:
        installed = set(tkfont.families(root))
    except Exception:
        installed = set()
    for name in FONT_CANDIDATES:
        if name in installed:
            FONT_FAMILY = name
            break


DEFAULT_SETTINGS = {
    "startup": "0",
    "show_toast_after": "0",     # 「調整後の明るさを表示する」。既定でOFF
    "show_number": "1",
    "delay_until_release": "0",  # 「止まるまで変更を遅らせる」。0=ドラッグ中も逐次反映、1=離した時だけ反映
    "link_sliders": "0",         # 「一斉に動かせるようにする」
    "auto_update": "0",          # 「自動アップデート」。AudioSwitcherと同じく既定でOFF(通知だけ出す)
    "language": "en",            # 表示言語。"en" / "ja"
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
_install_status_queue = queue.Queue()
_notes_queue = queue.Queue()

# 最後に見つかった新しいリリース。右クリックメニューの「vX.Y.Zに更新する」
# 項目の表示/非表示と、「自動アップデート」をONにした瞬間の即時更新に使う。
_update_state = {"info": None, "installing": False}

UPDATE_NOTICE_FILE = os.path.join(BASE_DIR, "update_notice.json")
_GITHUB_HEADERS = {"Accept": "application/vnd.github+json", "User-Agent": "Lumina-update-check"}


_CHECKSUM_ASSET = "SHA256SUMS"


def _is_github_release_url(url):
    """更新ファイルは必ずGitHubのリリースアセットから取る。それ以外のホスト・
    http:は(APIの応答が改ざんされていても)ダウンロードしない。"""
    try:
        p = urllib.parse.urlparse(url)
    except Exception:
        return False
    return p.scheme == "https" and p.hostname == "github.com" and p.path.lower().startswith(f"/{GITHUB_REPO.lower()}/releases/download/")


def _has_checksum_asset(assets):
    return any(a.get("name", "") == _CHECKSUM_ASSET for a in assets or [])


def _download_verified(url, dest_path, timeout):
    """urlをdest_pathに保存する。同じリリースのSHA256SUMSに載っているハッシュと
    一致しなければ保存したファイルを消して例外を投げる(実行も展開もしない)。"""
    if not _is_github_release_url(url):
        raise RuntimeError("untrusted update URL")
    headers = {"User-Agent": "Lumina-update-check"}
    sums_url = url.rsplit("/", 1)[0] + "/" + _CHECKSUM_ASSET
    with urllib.request.urlopen(urllib.request.Request(sums_url, headers=headers), timeout=10) as resp:
        sums = resp.read().decode("utf-8", "replace")
    name = url.rsplit("/", 1)[1]
    expected = None
    for line in sums.splitlines():
        parts = line.split()
        if len(parts) == 2 and parts[1].lstrip("*") == name:
            expected = parts[0].lower()
            break
    if not expected:
        raise RuntimeError("no checksum for update file")
    h = hashlib.sha256()
    with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=timeout) as resp, \
            open(dest_path, "wb") as f:
        for chunk in iter(lambda: resp.read(65536), b""):
            h.update(chunk)
            f.write(chunk)
    if h.hexdigest() != expected:
        try:
            os.remove(dest_path)
        except OSError:
            pass
        raise RuntimeError("update checksum mismatch")


def _github_get(path):
    req = urllib.request.Request(f"https://api.github.com/repos/{GITHUB_REPO}{path}", headers=_GITHUB_HEADERS)
    with urllib.request.urlopen(req, timeout=5) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _fetch_latest_release_info():
    """GitHubの最新リリースを確認する。__version__より新しい場合だけ
    (latest_tag, release_url, package_url, notes)を返し、そうでなければ
    Noneを返す。package_urlはポータブル版zip(Lumina-vX.Y.Z-win64.zip)の
    ダウンロードURL、notesはリリースの説明文(Markdown)。呼び出し側で例外を
    処理すること。"""
    data = _github_get("/releases/latest")
    latest_tag = data.get("tag_name", "")
    release_url = data.get("html_url") or ""
    # webbrowser.openに渡すので、このリポジトリのGitHubページ以外は使わない。
    if not release_url.lower().startswith(f"https://github.com/{GITHUB_REPO.lower()}/"):
        release_url = f"https://github.com/{GITHUB_REPO}/releases/latest"
    notes = data.get("body") or ""
    package_url = None
    # SHA256SUMSが無いリリースは自動更新せず、リリースページを開くフォール
    # バックに任せる(検証できないファイルは展開も実行もしない)。
    if _has_checksum_asset(data.get("assets")):
        for asset in data.get("assets", []) or []:
            name = asset.get("name", "").lower()
            if name.startswith("lumina") and name.endswith(".zip"):
                url = asset.get("browser_download_url")
                if _is_github_release_url(url or ""):
                    package_url = url
                break
    if latest_tag and _parse_version(latest_tag) > _parse_version(__version__):
        return (latest_tag, release_url, package_url, notes)
    return None


def _fetch_release_notes_worker():
    """右クリックメニューの「更新内容」用。今動いているバージョンのリリース
    ノートを取りに行き、まだ公開されていない(開発中の)バージョンなら最新
    リリースのものを代わりに出す。結果は_notes_queue経由でメインスレッドへ。"""
    try:
        try:
            data = _github_get(f"/releases/tags/v{__version__}")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
            data = _github_get("/releases/latest")
        _notes_queue.put(("ok", data.get("tag_name") or f"v{__version__}", data.get("body") or ""))
    except Exception as e:
        log_error("fetch_release_notes", e)
        _notes_queue.put(("error", f"v{__version__}", ""))


def _update_check_loop():
    """1日1回、GitHubの最新リリースを確認する。バックグラウンドの常駐
    ツールなので、失敗(オフライン等)は完全に握りつぶし通常動作に一切
    影響させない。新しいバージョンが見つかった時だけ、トースト表示用の
    キューに積んでメインスレッド側(poll_ui_queue)に伝える。"""
    while True:
        try:
            info = _fetch_latest_release_info()
            if info:
                _update_queue.put(info)
        except Exception as e:
            log_error("update_check", e)
        time.sleep(24 * 60 * 60)


def _write_update_notice(latest_tag, notes):
    """再起動後の新バージョンが「更新内容」を一度だけ表示できるよう、更新前の
    バージョンが自分で情報を書き残しておく。失敗しても更新自体は続ける。"""
    try:
        with open(UPDATE_NOTICE_FILE, "w", encoding="utf-8") as f:
            json.dump({"version": latest_tag, "notes": notes}, f, ensure_ascii=False)
    except Exception as e:
        log_error("write_update_notice", e)


def _read_pending_update_notice():
    """起動時に一度だけ呼ぶ。前回の更新時に書き残されたUPDATE_NOTICE_FILEが
    あれば(version, notes)を返してすぐ削除する(次回からは出さない)。"""
    if not os.path.exists(UPDATE_NOTICE_FILE):
        return None
    try:
        with open(UPDATE_NOTICE_FILE, "r", encoding="utf-8") as f:
            notice = json.load(f)
    except Exception as e:
        log_error("read_update_notice", e)
        notice = None
    try:
        os.remove(UPDATE_NOTICE_FILE)
    except Exception:
        pass
    if not notice:
        return None
    return notice.get("version", ""), notice.get("notes", "")


def _download_and_apply_update(package_url, release_url, latest_tag="", notes=""):
    """ポータブル版のzipをダウンロードし、中のLumina.exeを今のexeと入れ替えて
    再起動する。Luminaはインストーラーを持たない(zipを展開して置くだけの)配布
    形態なので、AudioSwitcherのようにインストーラーへ任せることができない。
    実行中のexeは自分では上書きできないため、小さなバッチを別プロセスで起動し、
    こちらが終了してロックが外れるのを待ってから差し替え・再起動させる。
    結果は_install_status_queue経由でメインスレッドに伝える。"""
    try:
        if not getattr(sys, "frozen", False):
            raise RuntimeError("self-update is only available in the packaged exe")
        work_dir = tempfile.mkdtemp(prefix="lumina-update-")
        zip_path = os.path.join(work_dir, "package.zip")
        _download_verified(package_url, zip_path, 60)

        new_exe = os.path.join(work_dir, "Lumina.exe")
        with zipfile.ZipFile(zip_path) as zf:
            member = next((n for n in zf.namelist() if os.path.basename(n).lower() == "lumina.exe"), None)
            if member is None:
                raise RuntimeError("Lumina.exe not found in the update package")
            with zf.open(member) as src, open(new_exe, "wb") as dst:
                dst.write(src.read())

        target = sys.executable
        # 値をそのままバッチに書き込むので、cmdが特別扱いする文字を含むパスでは
        # 差し替えを中止する(誤動作・コマンド注入の防止)。
        if any(c in target + new_exe for c in '%&^|<>!"'):
            raise RuntimeError("unsafe characters in path for update script")
        script = os.path.join(work_dir, "apply_update.cmd")
        # copyが成功する(=旧プロセスが終わってexeのロックが外れる)まで最大60秒
        # 待つ。それでも失敗した場合は差し替えずに元のexeをそのまま起動し直す。
        with open(script, "w", encoding="mbcs") as f:
            f.write(
                "@echo off\r\n"
                "set /a tries=0\r\n"
                ":retry\r\n"
                "ping -n 2 127.0.0.1 >nul\r\n"
                f'copy /y "{new_exe}" "{target}" >nul 2>&1 && goto launch\r\n'
                "set /a tries+=1\r\n"
                "if %tries% lss 60 goto retry\r\n"
                ":launch\r\n"
                f'start "" "{target}"\r\n'
            )
        _write_update_notice(latest_tag, notes)
        env = dict(os.environ)
        # PyInstaller(onefile)のexeから別のonefile exeを起動する時は、展開先
        # などの環境変数を引き継がないよう明示的にリセットさせる。
        env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
        subprocess.Popen(
            ["cmd", "/c", script], env=env, cwd=work_dir,
            creationflags=subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP,
        )
        _install_status_queue.put(("done", None))
    except Exception as e:
        log_error("apply_update", e)
        _install_status_queue.put(("error", release_url))


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

_known_unsupported_monitors = set()  # プロセスの寿命の間だけ保持(再起動すれば再挑戦する)


def get_monitor_snapshot():
    """接続中のモニター一覧と現在の輝度を取得する。
    モニターによってはDDC/CI非対応で例外が出るため、1台ずつ個別にtryする。
    一部のモニター(EIZOのEcoView競合等)は失敗するまでに数秒かかることが
    あり、毎回律儀に待つとパネルを開くたびに待たされてしまう。一度失敗
    したモニターはシリアル/uidで覚えておき、以降はこのプロセスが動いて
    いる間は読み取りを試さず即座に非対応扱いにする。"""
    snapshot = []
    try:
        infos = sbc.list_monitors_info()
    except Exception as e:
        log_error("list_monitors_info", e)
        return snapshot

    for i, info in enumerate(infos):
        name = info.get("name") or info.get("model") or f"Monitor {i}"
        key = info.get("serial") or info.get("uid") or name

        if key in _known_unsupported_monitors:
            snapshot.append({"index": i, "name": name, "brightness": None})
            continue

        try:
            brightness = sbc.get_brightness(display=i)[0]
        except Exception as e:
            log_error(f"get_brightness[{i}]", e)
            brightness = None
            _known_unsupported_monitors.add(key)
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
    """メインモニター(プライマリディスプレイ)の作業領域(タスクバーを除く)
    を返す。以前はマウスカーソルがあるモニターに追従していたが、常に
    メインモニターに表示してほしいという要望で固定した。失敗した場合は
    例外を投げる(呼び出し側でフォールバックする)。"""
    MONITOR_DEFAULTTOPRIMARY = 1
    hmonitor = ctypes.windll.user32.MonitorFromPoint(_POINT(0, 0), MONITOR_DEFAULTTOPRIMARY)
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
# トレイアイコンの生成
#
# Yuga's labのLuminaアイコン(暗い背景に白線のモニター＋アンバーの太陽)を、
# 通知領域の小さなサイズでも読めるよう線を太くしたグリフ。タスクバーの明暗に
# 合わせて線の色を切り替える(暗いタスクバー→白線、明るいタスクバー→墨色の線)。
# 白いタスクバーで白いアイコンが溶けて見えなくなる問題への対策。
# ----------------------------------------------------------------------

_PERSONALIZE_KEY = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"


def _taskbar_is_light():
    """タスクバー(スタートメニュー含む)がWindowsのライトテーマかどうかを
    レジストリから判定する。読み取りに失敗した場合はダーク扱いにする。"""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _PERSONALIZE_KEY) as key:
            value, _ = winreg.QueryValueEx(key, "SystemUsesLightTheme")
        return value == 1
    except Exception:
        return False


def render_tray_icon(light_taskbar=None, size=64):
    if light_taskbar is None:
        light_taskbar = _taskbar_is_light()
    if light_taskbar:
        line = (22, 21, 31, 255)    # COLOR_BG(アイコンの背景色)と同じ墨色
        sun = (196, 128, 40, 255)   # 明るい背景でも沈まないよう少し濃くしたアンバー
    else:
        line = (236, 232, 225, 255)  # COLOR_TEXT_MAIN(アイコンの白線)
        sun = (224, 168, 90, 255)    # COLOR_ACCENT

    ss = 4  # 縮小時にきれいなアンチエイリアスをかけるため、4倍で描いてから縮める
    S = size * ss
    k = S / 64
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    w = round(5.5 * k)

    # モニター(外形・スタンド・台座)
    draw.rounded_rectangle([3 * k, 7 * k, 61 * k, 47 * k], radius=7 * k, outline=line, width=w)
    draw.line([32 * k, 47 * k, 32 * k, 55 * k], fill=line, width=w)
    draw.rounded_rectangle([18 * k, 54 * k, 46 * k, 59.5 * k], radius=2.75 * k, fill=line)

    # 画面の中の太陽
    cx, cy, r = 32 * k, 27 * k, 7.5 * k
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=sun)
    ray_w = round(3.6 * k)
    for i in range(8):
        angle = math.radians(i * 45)
        x1, y1 = cx + 11.5 * k * math.cos(angle), cy + 11.5 * k * math.sin(angle)
        x2, y2 = cx + 14.5 * k * math.cos(angle), cy + 14.5 * k * math.sin(angle)
        draw.line([x1, y1, x2, y2], fill=sun, width=ray_w)
        for x, y in ((x1, y1), (x2, y2)):  # 線の端を丸くする
            draw.ellipse([x - ray_w / 2, y - ray_w / 2, x + ray_w / 2, y + ray_w / 2], fill=sun)
    return img.resize((size, size), Image.LANCZOS)


def _watch_taskbar_theme(on_change):
    """Windowsの「色」設定(ライト/ダーク)が切り替わったら、on_change(is_light)
    を呼ぶ。RegNotifyChangeKeyValueでレジストリの変更通知を待つので、ポーリング
    と違って何も起きていない間はCPUを使わない。通知APIが使えない場合だけ
    5秒おきの読み直しにフォールバックする。"""
    REG_NOTIFY_CHANGE_LAST_SET = 0x00000004
    last = _taskbar_is_light()
    while True:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _PERSONALIZE_KEY, 0,
                                winreg.KEY_NOTIFY | winreg.KEY_READ) as key:
                rc = ctypes.windll.advapi32.RegNotifyChangeKeyValue(
                    ctypes.c_void_p(int(key)), False, REG_NOTIFY_CHANGE_LAST_SET, None, False)
            if rc != 0:
                raise OSError(f"RegNotifyChangeKeyValue failed: {rc}")
        except Exception as e:
            log_error("watch_taskbar_theme", e)
            time.sleep(5)
        current = _taskbar_is_light()
        if current != last:
            last = current
            try:
                on_change(current)
            except Exception as e:
                log_error("taskbar_theme_changed", e)


def draw_lumina_glyph(canvas, cx, cy, size, line=COLOR_TEXT_MAIN, sun=COLOR_ACCENT):
    """トレイアイコンと同じモニター＋太陽のグリフをCanvasに描く(パネルの
    ヘッダーやトースト用)。sizeはグリフ全体の幅。"""
    k = size / 64
    x0, y0 = cx - size / 2, cy - size / 2
    w = max(1.5, 5 * k)
    _rounded_rect(canvas, x0 + 3 * k, y0 + 7 * k, x0 + 61 * k, y0 + 47 * k, 6 * k,
                  fill="", outline=line, width=w)
    canvas.create_line(x0 + 32 * k, y0 + 47 * k, x0 + 32 * k, y0 + 55 * k, fill=line, width=w)
    canvas.create_line(x0 + 20 * k, y0 + 57 * k, x0 + 44 * k, y0 + 57 * k, fill=line, width=w,
                       capstyle="round")
    scx, scy, r = x0 + 32 * k, y0 + 27 * k, 7.5 * k
    canvas.create_oval(scx - r, scy - r, scx + r, scy + r, fill=sun, outline="")
    for i in range(8):
        a = math.radians(i * 45)
        canvas.create_line(scx + 11.5 * k * math.cos(a), scy + 11.5 * k * math.sin(a),
                           scx + 15 * k * math.cos(a), scy + 15 * k * math.sin(a),
                           fill=sun, width=max(1.2, 3.4 * k), capstyle="round")


# ----------------------------------------------------------------------
# モダンなカスタムスライダー(Canvas描画・細身のトラック・
# ドラッグ中はつまみの周りにグロー(光暈)が浮かぶ)
# ----------------------------------------------------------------------

class ModernSlider:
    TRACK_THICKNESS = 4
    THUMB_R = 7
    THUMB_R_ACTIVE = 8

    def __init__(self, canvas, x, y, width, height, value=50, on_change=None, on_release=None,
                 panel_bg=COLOR_PANEL_BG, disabled=False):
        self.canvas = canvas
        self.x, self.y, self.width, self.height = x, y, width, height
        self.value = value
        self.on_change = on_change
        self.on_release = on_release
        self.panel_bg = panel_bg
        self.disabled = disabled
        self._dragging = False
        self._items = []

        self._hit_id = canvas.create_rectangle(
            x, y, x + width, y + height, fill="", outline=""
        )
        if not disabled:
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
        # 非対応モニターは操作できないことが一目でわかるよう、進捗の塗り
        # つぶしもグロー演出も出さず、つまみもトラック上に固定した薄い
        # 色のままにする(COLOR_ACCENTを一切使わない)。
        if not self.disabled and fill_x > track_x1 + 1:
            self._items.append(_rounded_rect(
                self.canvas, track_x1, track_y1, fill_x, track_y2, t,
                fill=COLOR_ACCENT, outline=""
            ))

        thumb_cx = fill_x
        thumb_cy = self.y + self.height / 2

        if self.disabled:
            r = self.THUMB_R
            self._items.append(self.canvas.create_oval(
                thumb_cx - r, thumb_cy - r, thumb_cx + r, thumb_cy + r,
                fill=self.panel_bg, outline=COLOR_TEXT_FAINT, width=2
            ))
            self.canvas.tag_raise(self._hit_id)
            return

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
    ROW_HEIGHT = 52
    HEADER_HEIGHT = 44
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
        self._snapshot_queue = queue.Queue()
        self._loading_token = 0

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
            if m["brightness"] is None:
                continue  # 非対応モニターはマスターフェーダーの対象外
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

    def _draw_shell(self, canvas, width, height, loading):
        """パネル本体(角丸背景・ヘッダー・閉じるボタン)を描く。中身
        (スライダー行)は含まない。open()の最初(読み込み中の仮表示)と、
        スナップショット取得後の本描画の両方から呼ぶため切り出している。"""
        _rounded_rect(canvas, 0, 0, width, height, self.CORNER_RADIUS,
                      fill=COLOR_PANEL_BG, outline=COLOR_PANEL_BORDER, width=1.2)

        draw_lumina_glyph(canvas, self.PADDING + 9, self.HEADER_HEIGHT / 2, 20)
        title = canvas.create_text(
            self.PADDING + 26, self.HEADER_HEIGHT / 2, anchor="w", text="Lumina",
            fill=COLOR_TEXT_MAIN, font=F(10, "bold")
        )
        bbox = canvas.bbox(title)
        canvas.create_text(
            (bbox[2] if bbox else self.PADDING + 80) + 6, self.HEADER_HEIGHT / 2 + 1, anchor="w",
            text=f"v{__version__}", fill=COLOR_TEXT_FAINT, font=F(8)
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
            self.PADDING, self.HEADER_HEIGHT, width - self.PADDING, self.HEADER_HEIGHT,
            fill=COLOR_PANEL_BORDER
        )
        if loading:
            canvas.create_text(
                width / 2, self.HEADER_HEIGHT + 30,
                text=L("Loading...", "読み込み中..."), fill=COLOR_TEXT_SUB, font=F(10)
            )

    def open(self):
        self.close()
        self._loading_token += 1
        token = self._loading_token

        # 応答が遅い/しないモニターがあると、get_monitor_snapshot()の
        # DDC/CI通信が数秒かかることがある(EIZO機で実測3秒台)。メイン
        # スレッドで同期的に待つとウィンドウ全体が固まり「開かない」ように
        # 見えてしまうため、実際の取得は必ず別スレッドで行う。
        width = self.WIDTH
        height = self.HEADER_HEIGHT + 60

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

        if self.snapshot:
            # 前回取得済みのデータがあれば、開閉を連打しても毎回「読み込み
            # 中」を経由しないよう、まずそれを使って即座に本描画する。
            # 値が変わっている可能性はあるので、裏では必ず再取得もする
            # (届き次第、表示中の内容を静かに更新する)。
            self._render_content(self.snapshot, initial=True)
        else:
            self._draw_shell(canvas, width, height, loading=True)
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

        def worker():
            snapshot = get_monitor_snapshot()
            self._snapshot_queue.put((token, snapshot))
        threading.Thread(target=worker, daemon=True).start()
        self._poll_snapshot()

    def _poll_snapshot(self):
        try:
            token, snapshot = self._snapshot_queue.get_nowait()
        except queue.Empty:
            if self.window is not None:
                self.root.after(50, self._poll_snapshot)
            return
        if token != self._loading_token or self.window is None:
            return  # 待っている間に閉じられた/開き直された結果は捨てる
        self._render_content(snapshot)

    def _render_content(self, snapshot, initial=False):
        settings = load_config()
        show_number = settings.get("show_number", "1") == "1"
        link_sliders = settings.get("link_sliders", "0") == "1"

        self.snapshot = snapshot
        has_master = link_sliders and len(self.snapshot) > 1

        row_h = self.ROW_HEIGHT
        width = self.WIDTH
        rows = max(1, len(self.snapshot))
        body_h = rows * row_h if self.snapshot else 50
        master_h = row_h + 8 if has_master else 0
        height = self.HEADER_HEIGHT + master_h + body_h + self.PADDING

        win = self.window
        canvas = self.canvas
        canvas.delete("all")
        canvas.configure(width=width, height=height)
        self.sliders = []
        self.value_labels = []
        self.master_slider = None
        self.master_label = None

        self._draw_shell(canvas, width, height, loading=False)

        cursor_y = self.HEADER_HEIGHT

        if not self.snapshot:
            canvas.create_text(
                width / 2, self.HEADER_HEIGHT + body_h / 2,
                text=L("No DDC/CI-capable\nmonitors found", "DDC/CI対応モニターが\n見つかりません"),
                fill=COLOR_TEXT_SUB, font=F(10), justify="center"
            )
        else:
            if has_master:
                master_row_y = cursor_y + 8
                _rounded_rect(
                    canvas, self.PADDING * 0.4, cursor_y + 2, width - self.PADDING * 0.4,
                    cursor_y + master_h - 4, 10, fill=COLOR_PANEL_BG_ALT, outline=""
                )
                canvas.create_text(
                    self.PADDING, master_row_y, anchor="nw", text=L("All Monitors", "すべてのモニター"),
                    fill=COLOR_ACCENT, font=F(9, "bold")
                )

                supported_brightness = [m["brightness"] for m in self.snapshot if m["brightness"] is not None]
                master_value = round(sum(supported_brightness) / len(supported_brightness)) if supported_brightness else 0
                master_label = None
                if show_number:
                    master_label = tk.Label(
                        win, text=f"{master_value}%", font=F(10, "bold"), pady=0,
                        fg=COLOR_ACCENT, bg=COLOR_PANEL_BG_ALT
                    )
                    canvas.create_window(
                        width - self.PADDING, master_row_y - 2, anchor="ne", window=master_label
                    )

                slider_width = width - self.PADDING * 2

                def on_master_change(value, is_final):
                    if master_label is not None:
                        master_label.configure(text=f"{int(round(value))}%")
                    self._apply_master(value, is_final)

                def on_master_release(value):
                    if master_label is not None:
                        master_label.configure(text=f"{int(round(value))}%")
                    self._apply_master(value, True)

                self.master_slider = ModernSlider(
                    canvas, self.PADDING, master_row_y + 24, slider_width, 14,
                    value=master_value, on_change=on_master_change, on_release=on_master_release,
                    panel_bg=COLOR_PANEL_BG_ALT,
                )
                self.master_label = master_label
                cursor_y += master_h

            for i, m in enumerate(self.snapshot):
                row_y = cursor_y + i * row_h + 10
                name = m["name"]
                display_name = name if len(name) <= 24 else name[:22] + "..."
                # brightnessがNone = get_monitor_snapshot()側でDDC/CI通信に
                # 失敗した(=このモニターは輝度調整に非対応)ということなので、
                # 名前を薄い色にし、スライダーも操作できない見た目にする。
                unsupported = m["brightness"] is None
                canvas.create_text(
                    self.PADDING, row_y, anchor="nw", text=display_name,
                    fill=COLOR_TEXT_FAINT if unsupported else COLOR_TEXT_MAIN, font=F(10)
                )

                value_label = None
                if unsupported:
                    value_label = tk.Label(
                        win, text=L("Unsupported", "非対応"), font=F(9), pady=0,
                        fg=COLOR_TEXT_FAINT, bg=COLOR_PANEL_BG
                    )
                    canvas.create_window(
                        width - self.PADDING, row_y - 2, anchor="ne", window=value_label
                    )
                elif show_number:
                    value_label = tk.Label(
                        win, text=f"{m['brightness'] or 0}%", font=F(10, "bold"), pady=0,
                        fg=COLOR_ACCENT, bg=COLOR_PANEL_BG
                    )
                    canvas.create_window(
                        width - self.PADDING, row_y - 2, anchor="ne", window=value_label
                    )

                slider_y = row_y + 24
                slider_width = width - self.PADDING * 2

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
                    value=m["brightness"] or 0, on_change=on_change, on_release=on_release,
                    disabled=unsupported,
                )
                self.sliders.append(slider)
                self.value_labels.append(value_label)

        win.update_idletasks()
        left, top, right, bottom = _safe_work_area(win)
        win_w = win.winfo_reqwidth()
        win_h = win.winfo_reqheight()
        self._final_x = max(left, right - win_w - 8)
        self._final_y = max(top, bottom - win_h - 8)
        if initial:
            # open()からの初回描画(キャッシュ済みデータで即座に本描画する
            # ケース)。まだフェード/スライドアニメーションが始まっていない
            # ので、_animate_open()が使うスライド分オフセットした位置に置く。
            win.geometry(f"{win_w}x{win_h}+{self._final_x}+{self._final_y + self.ANIM_SLIDE_PX}")
        else:
            # 読み込み中の仮表示、または表示中の内容をバックグラウンドで
            # 更新するケース。ウィンドウは既に表示済みなのでアニメーションは
            # やり直さず、サイズと位置(右下基準)だけ更新する。
            win.geometry(f"{win_w}x{win_h}+{self._final_x}+{self._final_y}")

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
    WIDTH, HEIGHT = 300, 68
    WHATSNEW_WIDTH = 360
    TEXT_X = 62

    def __init__(self, root):
        self.root = root
        self.window = None
        self.canvas = None
        self.fade_job = None
        self.current_alpha = 0.0
        self._slide = 0
        self.on_click = None
        self._hold_ms = 1600
        self._subtitle_id = None
        self._pending_release_url = None
        self.whatsnew_window = None
        self.whatsnew_canvas = None
        self.whatsnew_alpha = 0.0
        self.whatsnew_fade_job = None
        self.whatsnew_height = 200

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
        y = bottom - self.HEIGHT - 8 + round(self._slide)
        self.window.geometry(f"{self.WIDTH}x{self.HEIGHT}+{x}+{y}")

    def _draw_frame(self, header_text):
        """全トースト共通の土台(角丸パネル・左のアイコンバッジ・見出し)"""
        canvas = self.canvas
        canvas.delete("all")
        _rounded_rect(canvas, 1, 1, self.WIDTH - 1, self.HEIGHT - 1, 14,
                      fill=COLOR_PANEL_BG, outline=COLOR_PANEL_BORDER, width=1)
        cy = self.HEIGHT / 2
        _rounded_rect(canvas, 14, cy - 18, 50, cy + 18, 10, fill=COLOR_PANEL_BG_ALT, outline="")
        draw_lumina_glyph(canvas, 32, cy, 24)
        canvas.create_text(self.TEXT_X, cy - 11, anchor="w", text=header_text,
                           fill=COLOR_TEXT_SUB, font=F(8, "bold"))

    def _start(self, hold_ms):
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None
        self._hold_ms = hold_ms
        self.current_alpha = 0.0
        self._slide = 12
        self._place()
        self.window.deiconify()
        self._place()
        self._fade_in()

    def show(self, brightness_percent):
        self._ensure_window()
        self._draw_frame(L("Brightness", "明るさ"))
        canvas = self.canvas
        cy = self.HEIGHT / 2
        value = int(round(brightness_percent))
        canvas.create_text(self.TEXT_X, cy + 9, anchor="w", text=f"{value}%",
                           fill=COLOR_TEXT_MAIN, font=F(13, "bold"))
        # 現在値をミニバーでも示す
        bar_x1, bar_x2, bar_y = self.TEXT_X + 70, self.WIDTH - 22, cy + 10
        _rounded_rect(canvas, bar_x1, bar_y - 2, bar_x2, bar_y + 2, 2, fill=COLOR_TRACK, outline="")
        fill_x = bar_x1 + (bar_x2 - bar_x1) * max(0, min(100, value)) / 100
        if fill_x > bar_x1 + 2:
            _rounded_rect(canvas, bar_x1, bar_y - 2, fill_x, bar_y + 2, 2, fill=COLOR_ACCENT, outline="")
        canvas.configure(cursor="")
        self.on_click = None
        self._start(1600)

    def show_update_available(self, latest_tag, release_url, package_url=None, notes="", auto=False):
        """新しいバージョンが見つかった時の控えめな通知。exe版でzipがあれば
        クリック(自動アップデートONなら即座に)でダウンロードして入れ替える。
        それ以外(ソースから実行中・zipの無いリリース)はリリースページを開く。"""
        self._ensure_window()
        self._draw_frame(L("Update available", "アップデートがあります"))
        canvas = self.canvas
        can_install = bool(package_url) and getattr(sys, "frozen", False)
        auto = auto and can_install
        if auto:
            subtitle = L(f"{latest_tag} — updating automatically...", f"{latest_tag} — 自動で更新しています...")
        elif can_install:
            subtitle = L(f"{latest_tag} — click to update", f"{latest_tag} — クリックで更新")
        else:
            subtitle = L(f"{latest_tag} — click to view", f"{latest_tag} — クリックで詳細を表示")
        self._subtitle_id = canvas.create_text(self.TEXT_X, self.HEIGHT / 2 + 9, anchor="w", text=subtitle,
                                               fill=COLOR_TEXT_MAIN, font=F(10, "bold"))
        self._pending_release_url = release_url
        if auto:
            canvas.configure(cursor="")
            self.on_click = None
        elif can_install:
            canvas.configure(cursor="hand2")
            self.on_click = lambda: self.begin_update(latest_tag, release_url, package_url, notes)
        else:
            canvas.configure(cursor="hand2")
            self.on_click = lambda: webbrowser.open(release_url)
        self._start(8000)
        if auto:
            self.begin_update(latest_tag, release_url, package_url, notes)

    def begin_update(self, latest_tag, release_url, package_url, notes=""):
        """更新を開始する。Tkのメインスレッドを止めないよう、ダウンロードと
        展開は別スレッドで行い、結果は_install_status_queue経由で受け取る。"""
        if _update_state["installing"]:
            return
        _update_state["installing"] = True
        if self.window is None or not self.window.winfo_viewable():
            self.show_update_available(latest_tag, release_url, package_url, notes)
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
            self.fade_job = None
        self.on_click = None
        try:
            self.canvas.configure(cursor="")
            if self._subtitle_id is not None:
                self.canvas.itemconfigure(self._subtitle_id, text=L("Downloading...", "ダウンロードしています..."))
        except Exception:
            pass
        threading.Thread(target=_download_and_apply_update,
                         args=(package_url, release_url, latest_tag, notes), daemon=True).start()

    def finish_update(self, ok):
        """ダウンロード完了(成功/失敗)を受けて表示を更新する。失敗時は
        クリックでリリースページを開けるようにフォールバックする。"""
        if not ok:
            _update_state["installing"] = False
        try:
            if self._subtitle_id is not None and self.canvas is not None:
                if ok:
                    self.canvas.itemconfigure(self._subtitle_id, text=L("Updating — Lumina will restart automatically", "更新しています — 自動で再起動します"))
                else:
                    self.canvas.itemconfigure(self._subtitle_id, text=L("Update failed — click to view", "更新に失敗 — クリックで詳細を表示"))
                    self.canvas.configure(cursor="hand2")
                    url = self._pending_release_url
                    self.on_click = lambda: webbrowser.open(url)
        except Exception:
            pass
        if self.fade_job:
            self.root.after_cancel(self.fade_job)
        self.fade_job = self.root.after(5000, self._fade_out)

    def _fade_in(self):
        # フェードしながら下から少しせり上がる
        if self.current_alpha < 0.97:
            self.current_alpha += 0.12
            self._slide = max(0.0, self._slide * 0.7)
            self.window.attributes("-alpha", min(self.current_alpha, 0.97))
            self._place()
            self.fade_job = self.root.after(15, self._fade_in)
        else:
            self._slide = 0
            self._place()
            self.fade_job = self.root.after(self._hold_ms, self._fade_out)

    def _fade_out(self):
        if self.current_alpha > 0.0:
            self.current_alpha -= 0.08
            self.window.attributes("-alpha", max(self.current_alpha, 0.0))
            self.fade_job = self.root.after(15, self._fade_out)
        else:
            self.window.withdraw()
            self.fade_job = None

    # ------------------------------------------------------------------
    # 更新内容(リリースノート)のポップアップ。アップデート直後に一度だけ
    # 自動で出るほか、右クリックメニューの「更新内容」からいつでも開ける。
    # 読む時間が必要なので自動では消さず、クリックで閉じる。
    # ------------------------------------------------------------------

    def _ensure_whatsnew_window(self):
        if self.whatsnew_window is not None:
            return
        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.attributes("-alpha", 0.0)
        win.configure(bg=COLOR_TRANSPARENT_KEY)
        win.attributes("-transparentcolor", COLOR_TRANSPARENT_KEY)

        canvas = tk.Canvas(win, width=self.WHATSNEW_WIDTH, height=self.whatsnew_height,
                            bg=COLOR_TRANSPARENT_KEY, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        canvas.bind("<Button-1>", lambda e: self._dismiss_whats_new())
        win.withdraw()

        self.whatsnew_window = win
        self.whatsnew_canvas = canvas

    def _place_whats_new(self):
        left, top, right, bottom = _safe_work_area(self.whatsnew_window)
        x = right - self.WHATSNEW_WIDTH - 8
        y = bottom - self.whatsnew_height - 8
        self.whatsnew_window.geometry(f"{self.WHATSNEW_WIDTH}x{self.whatsnew_height}+{x}+{y}")

    def show_whats_new(self, version, notes, header=None, footer_url=None):
        if header is None:
            header = L("Updated", "アップデートしました")
        self._ensure_whatsnew_window()
        if self.whatsnew_fade_job:
            self.root.after_cancel(self.whatsnew_fade_job)
            self.whatsnew_fade_job = None

        canvas = self.whatsnew_canvas
        canvas.delete("all")
        W = self.WHATSNEW_WIDTH
        MAX_H = 440
        FOOTER_H = 48

        _rounded_rect(canvas, 18, 16, 54, 52, 10, fill=COLOR_PANEL_BG_ALT, outline="")
        draw_lumina_glyph(canvas, 36, 34, 24)
        canvas.create_text(66, 25, anchor="w", text=header, fill=COLOR_ACCENT, font=F(9, "bold"))
        canvas.create_text(66, 44, anchor="w", text=f"Lumina {version}",
                           fill=COLOR_TEXT_MAIN, font=F(13, "bold"))
        canvas.create_line(18, 66, W - 18, 66, fill=COLOR_PANEL_BORDER)

        # Markdownの見出し(#)と箇条書き(-)だけを簡単に見た目調整して並べる
        y = 78
        body = (notes or "").strip() or L("No release notes.", "リリースノートはありません。")
        for raw in body.splitlines():
            line = raw.strip()
            if not line:
                continue
            dot = None
            if line.startswith("#"):
                item = canvas.create_text(22, y + 6, anchor="nw", text=line.lstrip("#").strip(),
                                          fill=COLOR_ACCENT, font=F(9, "bold"))
            elif line.startswith(("- ", "* ")):
                dot = canvas.create_oval(24, y + 7, 29, y + 12, fill=COLOR_TEXT_SUB, outline="")
                item = canvas.create_text(38, y, anchor="nw", text=line[2:].replace("`", ""),
                                          fill=COLOR_TEXT_MAIN, font=F(9), width=W - 58)
            else:
                item = canvas.create_text(22, y, anchor="nw", text=line.replace("`", ""),
                                          fill=COLOR_TEXT_SUB, font=F(9), width=W - 44)
            bbox = canvas.bbox(item)
            bottom = bbox[3] if bbox else y + 16
            if bottom > MAX_H - FOOTER_H:
                # 入りきらない分は省略して、続きがあることだけ伝える
                canvas.delete(item)
                if dot is not None:
                    canvas.delete(dot)
                canvas.create_text(22, y, anchor="nw", text=L("…more in the release notes on GitHub", "…続きは GitHub のリリースノートへ"),
                                   fill=COLOR_TEXT_FAINT, font=F(9))
                y += 20
                break
            y = bottom + 5
        H = min(MAX_H, y + FOOTER_H)

        footer_y = H - 26
        btn = _rounded_rect(canvas, W - 98, footer_y - 14, W - 18, footer_y + 14, 8,
                            fill=COLOR_PANEL_BG_ALT, outline="")
        btn_text = canvas.create_text(W - 58, footer_y, text=L("Close", "閉じる"), fill=COLOR_TEXT_MAIN, font=F(9, "bold"))
        for item in (btn, btn_text):
            canvas.tag_bind(item, "<Enter>", lambda e: canvas.itemconfig(btn, fill=COLOR_PANEL_BORDER))
            canvas.tag_bind(item, "<Leave>", lambda e: canvas.itemconfig(btn, fill=COLOR_PANEL_BG_ALT))
        url = footer_url or f"https://github.com/{GITHUB_REPO}/releases"
        link = canvas.create_text(22, footer_y, anchor="w", text=L("View on GitHub", "GitHubで見る"),
                                  fill=COLOR_TEXT_SUB, font=F(9))
        canvas.tag_bind(link, "<Enter>", lambda e: canvas.itemconfig(link, fill=COLOR_ACCENT))
        canvas.tag_bind(link, "<Leave>", lambda e: canvas.itemconfig(link, fill=COLOR_TEXT_SUB))
        # リンクのクリックはウィンドウ全体の「クリックで閉じる」より先に処理し、
        # ブラウザを開いた上でポップアップも閉じる
        canvas.tag_bind(link, "<Button-1>", lambda e: webbrowser.open(url))
        panel = _rounded_rect(canvas, 1, 1, W - 1, H - 1, 16,
                              fill=COLOR_PANEL_BG, outline=COLOR_PANEL_BORDER, width=1)
        canvas.tag_lower(panel)
        canvas.configure(cursor="hand2", height=H)
        self.whatsnew_height = H

        self._place_whats_new()
        self.whatsnew_window.deiconify()
        self._place_whats_new()
        self.whatsnew_alpha = 0.0
        self._whatsnew_fade_in()

    def _dismiss_whats_new(self):
        if self.whatsnew_fade_job:
            self.root.after_cancel(self.whatsnew_fade_job)
            self.whatsnew_fade_job = None
        self._whatsnew_fade_out()

    def _whatsnew_fade_in(self):
        if self.whatsnew_alpha < 0.97:
            self.whatsnew_alpha += 0.12
            self.whatsnew_window.attributes("-alpha", min(self.whatsnew_alpha, 0.97))
            self.whatsnew_fade_job = self.root.after(15, self._whatsnew_fade_in)
        else:
            self.whatsnew_fade_job = None

    def _whatsnew_fade_out(self):
        if self.whatsnew_alpha > 0.0:
            self.whatsnew_alpha -= 0.08
            self.whatsnew_window.attributes("-alpha", max(self.whatsnew_alpha, 0.0))
            self.whatsnew_fade_job = self.root.after(15, self._whatsnew_fade_out)
        else:
            self.whatsnew_window.withdraw()
            self.whatsnew_fade_job = None


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
            elif key == "auto_update" and current[key] == "1":
                # ONにした時点で既に新しいバージョンが見つかっていれば、次の
                # 定期チェック(最大1日後)を待たずにすぐ更新する
                _ui_action_queue.put(("install_update", None))
        return _handler

    def on_open_panel(icon, item):
        _ui_action_queue.put(("open_panel", None))

    def on_show_notes(icon, item):
        _ui_action_queue.put(("show_notes", None))

    def on_install_update(icon, item):
        _ui_action_queue.put(("install_update", None))

    def update_label(item):
        info = _update_state["info"]
        return L(f"Update to {info[0]}", f"{info[0]} に更新する") if info else ""

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

    def set_language(lang):
        def _handler(icon, item):
            global LANG
            current = load_config()
            current["language"] = lang
            save_config(current)
            LANG = lang
            _ui_action_queue.put(("language_changed", None))
        return _handler

    return pystray.Menu(
        pystray.MenuItem("open", on_open_panel, default=True, visible=False),
        pystray.MenuItem(f"Lumina  v{__version__}", None, enabled=False),
        pystray.MenuItem(update_label, on_install_update, visible=lambda i: _update_state["info"] is not None),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(lambda i: L("Launch at sign-in", "サインイン時に起動する"), toggle("startup"), checked=lambda i: is_on("startup")),
        pystray.MenuItem(lambda i: L("Show brightness after adjusting", "調整後の明るさを表示する"), toggle("show_toast_after"), checked=lambda i: is_on("show_toast_after")),
        pystray.MenuItem(lambda i: L("Show current number", "現在の数字を表示する"), toggle("show_number"), checked=lambda i: is_on("show_number")),
        pystray.MenuItem(lambda i: L("Delay change until stopped", "止まるまで変更を遅らせる"), toggle("delay_until_release"), checked=lambda i: is_on("delay_until_release")),
        pystray.MenuItem(lambda i: L("Link sliders together", "一斉に動かせるようにする"), toggle("link_sliders"), checked=lambda i: is_on("link_sliders")),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(lambda i: L("Auto-update", "自動アップデート"), toggle("auto_update"), checked=lambda i: is_on("auto_update")),
        pystray.MenuItem(lambda i: L("What's New...", "更新内容..."), on_show_notes),
        pystray.MenuItem("Language / 言語", pystray.Menu(
            pystray.MenuItem("English", set_language("en"), checked=lambda i: LANG == "en", radio=True),
            pystray.MenuItem("日本語", set_language("ja"), checked=lambda i: LANG == "ja", radio=True),
        )),
        pystray.MenuItem(lambda i: L("Open taskbar display settings...", "タスクバーの表示設定を開く..."),
                         on_open_taskbar_settings),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(lambda i: L("Exit", "終了"), on_quit),
    )


def run_app():
    global LANG
    enable_dpi_awareness()
    kill_previous_instances()

    settings = load_config()
    LANG = "ja" if settings.get("language") == "ja" else "en"
    if not os.path.exists(CONFIG_FILE) or any(k not in settings for k in DEFAULT_SETTINGS):
        save_config(settings)

    root = tk.Tk()
    root.withdraw()
    root.report_callback_exception = lambda exc_type, exc_value, tb: log_error("tk_callback", exc_value)
    resolve_font_family(root)

    toast_manager = ToastManager(root)
    root._toast_manager = toast_manager  # SliderPanelから参照するため
    panel = SliderPanel(root)

    tray_icon = pystray.Icon("Lumina", render_tray_icon(), f"Lumina v{__version__}", _build_menu())

    def on_theme_changed(is_light):
        # pystrayのiconプロパティへの代入はスレッドセーフ(内部でトレイの
        # ウィンドウにメッセージを投げて差し替える)なので、監視スレッドから直接行う
        tray_icon.icon = render_tray_icon(is_light)

    def start_update(info):
        latest_tag, release_url, package_url, notes = info
        if package_url and getattr(sys, "frozen", False):
            toast_manager.begin_update(latest_tag, release_url, package_url, notes)
        else:
            webbrowser.open(release_url)

    def poll_ui_queue():
        try:
            while True:
                action, payload = _ui_action_queue.get_nowait()
                if action == "open_panel":
                    panel.toggle()
                elif action == "show_notes":
                    threading.Thread(target=_fetch_release_notes_worker, daemon=True).start()
                elif action == "install_update":
                    if _update_state["info"]:
                        start_update(_update_state["info"])
                elif action == "language_changed":
                    # 開いているパネルは閉じる(次に開いた時に新しい言語で描かれる)
                    panel.close()
                    tray_icon.update_menu()
                elif action == "quit":
                    tray_icon.stop()
                    root.after(50, root.destroy)
        except queue.Empty:
            pass

        try:
            while True:
                info = _update_queue.get_nowait()
                _update_state["info"] = info
                tray_icon.update_menu()
                auto = load_config().get("auto_update", "0") == "1"
                toast_manager.show_update_available(*info, auto=auto)
        except queue.Empty:
            pass
        except Exception as e:
            log_error("poll_update_queue", e)

        try:
            while True:
                status, _ = _install_status_queue.get_nowait()
                toast_manager.finish_update(status == "done")
                if status == "done":
                    # 差し替え用のバッチはこのプロセスの終了を待っているので、
                    # 「再起動します」の表示を少し見せてから終了する
                    root.after(1500, lambda: _ui_action_queue.put(("quit", None)))
        except queue.Empty:
            pass

        try:
            while True:
                status, version, notes = _notes_queue.get_nowait()
                if status == "ok":
                    toast_manager.show_whats_new(version, notes, header=L("What's New", "更新内容"))
                else:
                    toast_manager.show_whats_new(
                        f"v{__version__}",
                        L("Couldn't load the release notes.\nPlease check your internet connection.",
                          "更新内容を取得できませんでした。\nインターネット接続を確認してください。"),
                        header=L("What's New", "更新内容"))
        except queue.Empty:
            pass

        root.after(50, poll_ui_queue)

    root.after(50, poll_ui_queue)

    pending_notice = _read_pending_update_notice()
    if pending_notice:
        version, notes = pending_notice
        root.after(800, lambda: toast_manager.show_whats_new(version or f"v{__version__}", notes))

    threading.Thread(target=tray_icon.run, daemon=True).start()
    threading.Thread(target=_brightness_writer_worker, daemon=True).start()
    threading.Thread(target=_update_check_loop, daemon=True).start()
    threading.Thread(target=_watch_taskbar_theme, args=(on_theme_changed,), daemon=True).start()
    root.mainloop()


if __name__ == "__main__":
    try:
        run_app()
    except SystemExit:
        raise
    except Exception as e:
        log_error("main", e)
        raise
