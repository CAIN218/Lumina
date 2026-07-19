# Lumina

A resident monitor brightness controller for Windows.
Click the icon in the taskbar's notification area to adjust the brightness of every connected monitor directly, with a modern slider panel.

- Click the taskbar icon → a brightness slider panel for every monitor pops open right there
- Right-click the icon → a settings menu with checkable toggles for behavior and display options
- Adjust DDC/CI-capable monitors independently. Turning on "Link sliders together" adds a master fader
  that moves every monitor at once (individual sliders always keep moving independently either way)
- Optionally shows a toast notification with the current brightness after adjusting (off by default)

## Requirements

- Windows 10 / 11
- A DDC/CI-capable monitor (most external monitors connected via HDMI/DisplayPort support this).
  Laptop built-in displays should use Windows' own brightness control instead (no DDC/CI support).
- If your monitor doesn't support DDC/CI, the panel will show "No DDC/CI-capable monitors found."
  Some monitors let you enable DDC/CI from their OSD menu.

## Usage

### 1. Launch

Run `Lumina.exe` (or `python lumina.py`) and a sun icon appears in the taskbar's notification area. No window is shown.

### 2. Adjust brightness

**Left-click** the tray icon to open a panel with one slider per monitor, fading and sliding into place in the
bottom-right corner. Drag a slider to change brightness. Click the ✕ in the top-right of the panel, or click
outside it, to close it.

When "Link sliders together" is on, a master fader labeled "All Monitors" is added at the top of the panel.
Moving it brings every monitor to the same value. Individual monitor sliders always move independently
regardless of this setting.

### 3. Change settings

**Right-click** the tray icon to open the settings menu. Every option is a checkable toggle that applies
immediately (no save button needed).

| Option | Description |
|---|---|
| Launch at sign-in | Automatically start Lumina when Windows boots |
| Show brightness after adjusting | Show a toast notification when you release a slider (off by default) |
| Show current number | Display each monitor's brightness as a percentage in the panel |
| Delay change until stopped | When on, dragging doesn't apply until you release; when off, brightness updates live while dragging |
| Link sliders together | Adds a master fader to the panel that moves every monitor at once |
| Open taskbar display settings... | Shortcut straight to Windows' taskbar settings. See the note below |

### 4. Quitting

Right-click the tray icon and choose "Exit".

## About "always show" in the taskbar

Whether the notification-area icon is always visible or tucked away under the hidden-icons chevron (^) is
**managed entirely by Windows on a per-app basis — there's no API for an app to force this itself** (this has
been the case since Vista). The "Open taskbar display settings..." menu item takes you straight to the
relevant Windows settings page (Taskbar → Other system tray icons), where turning Lumina "On" makes it
always visible. Windows remembers that choice going forward.

## Config files

These files are created next to the exe:

- `app_config.txt` — the settings above, in `key=value` format
- `error.log` — Error log for troubleshooting

## Building from source

```
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm Lumina.spec
```

This produces `dist/Lumina.exe`.

## Out of scope for now

The following features, present in reference apps like Twinkle Tray, are not implemented:

- **Arrange sliders by physical monitor layout**: skipped because there was no reliable way found to map
  physical monitor position to DDC/CI enumeration order — better to omit it than map it incorrectly
- **Adjustable per-monitor range limits**: would need extra UI, skipped for now
- **Contrast control**: the library in use (`screen_brightness_control`) doesn't support it (possible via
  raw VCP calls, but that would need separate work)

## Design notes

Shares AudioSwitcher's design philosophy (resident process, lightweight callbacks, worker threads) but the
UI is rebuilt around a tray icon, with its own deep indigo/warm-glow color identity ("Lumina" = light)
instead of reusing AudioSwitcher's look.

- The tray icon (pystray) runs its own message loop, so it runs on a separate thread from Tkinter's
  mainloop. Tray-side callbacks (clicks, menu selections) fire on pystray's thread, so all Tk operations
  are routed back to the mainloop through a queue
- Panel sliders are a custom Canvas-drawn implementation (rounded track, accent color fill) rather than
  Tk's stock Scale widget. While dragging, a soft glow (concentric circles blended toward the background
  color) appears around the thumb
- The panel's rounded corners are achieved by keying the whole window to a transparent color and drawing a
  rounded polygon on top (a plain Tk Toplevel is always rectangular otherwise)
- Per-Monitor DPI Awareness is explicitly enabled. Without it, Tkinter's rendering gets upscaled/blurred by
  the OS, and on multi-monitor setups the physical-pixel coordinates from GetCursorPos etc. drift out of
  sync with Tk's logical coordinates — this was the cause of the panel occasionally appearing in the wrong
  place (e.g. top-left)
- DDC/CI communication (which can take tens to ~100ms over I2C) is never called directly from a slider's
  drag handler. Instead, each monitor's latest target value is held in memory and a worker thread writes it
  out whenever it's free (a coalescing queue). This keeps dragging itself always responsive, even when
  moving several monitors at once via the master fader.
