# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

datas_sbc, binaries_sbc, hiddenimports_sbc = collect_all('screen_brightness_control')
datas_wmi, binaries_wmi, hiddenimports_wmi = collect_all('wmi')
datas_pystray, binaries_pystray, hiddenimports_pystray = collect_all('pystray')
datas_pil, binaries_pil, hiddenimports_pil = collect_all('PIL')

a = Analysis(
    ['brightness_switch.py'],
    pathex=[],
    binaries=binaries_sbc + binaries_wmi + binaries_pystray + binaries_pil,
    datas=datas_sbc + datas_wmi + datas_pystray + datas_pil,
    hiddenimports=[
        'screen_brightness_control',
        'wmi',
        'pystray',
        'pystray._win32',
        'PIL',
    ] + hiddenimports_sbc + hiddenimports_wmi + hiddenimports_pystray + hiddenimports_pil,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='BrightSwitch',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    # PyInstaller + UPX + Tkinterの組み合わせはtcl/tkのDLLが破損・不安定化する
    # 既知の問題があるため（AudioSwitcherで実機確認済み）、
    # Tcl/Tk関連DLLとpythonランタイム本体はUPX圧縮の対象から除外する。
    upx_exclude=['tcl86t.dll', 'tk86t.dll', '_tkinter*.pyd', 'python3*.dll', 'zlib1.dll'],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app_icon.ico'],
)
