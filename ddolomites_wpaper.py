# -*- coding: utf-8 -*-
import os
import sys
import time
import threading
import ctypes
import datetime as dt
from pathlib import Path
from zoneinfo import ZoneInfo
import webbrowser
import requests
import pystray
from PIL import Image, ImageDraw, ImageTk
import tkinter as tk
from tkinter import ttk, messagebox
import winreg as reg  # 修改壁纸样式用

# ============================================================
#   DDolomitesWpaper 通用动态壁纸抓取器
#   Developer: Maguamale
# ============================================================
# ---------- Startup (Run key) ----------
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
APP_NAME = "DDolomitesWpaper"

def _current_executable_path() -> str:
    """
    返回写入 HKCU\\...\\Run 的命令行：
    - 打包(onefile)时：   "path\\to\\app.exe"
    - 开发模式(.py)时：   "path\\to\\python.exe" "path\\to\\script.py"
    """
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}"'
    else:
        python = sys.executable
        script = os.path.abspath(sys.argv[0])
        return f'"{python}" "{script}"'


def enable_startup(enable: bool) -> bool:
    """开/关开机自启（写 HKCU\\...\\Run），返回是否成功"""
    try:
        key = reg.OpenKey(reg.HKEY_CURRENT_USER, RUN_KEY, 0, reg.KEY_SET_VALUE)
        if enable:
            reg.SetValueEx(key, APP_NAME, 0, reg.REG_SZ, _current_executable_path())
        else:
            try:
                reg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass
        reg.CloseKey(key)
        log(f"Startup set to {enable}", also_status=True)
        return True
    except Exception as e:
        log(f"Startup toggle error: {e}", also_status=True)
        return False

def is_startup_enabled() -> bool:
    """检测是否已开启开机自启"""
    try:
        key = reg.OpenKey(reg.HKEY_CURRENT_USER, RUN_KEY, 0, reg.KEY_READ)
        val, _ = reg.QueryValueEx(key, APP_NAME)
        reg.CloseKey(key)
        return isinstance(val, str) and len(val) > 0
    except FileNotFoundError:
        return False
    except Exception:
        return False

# Windows 壁纸常量
SPI_SETDESKWALLPAPER = 20
SPIF_UPDATEINIFILE = 0x01
SPIF_SENDWININICHANGE = 0x02

# 罗马时区
ROME_TZ = ZoneInfo("Europe/Rome")

# 基础路径配置
USER_DESKTOP = Path(os.path.expanduser("~")) / "Desktop"
APP_DIR = Path(os.environ.get("LOCALAPPDATA", USER_DESKTOP)) / "DDolomitesWpaper"
APP_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = APP_DIR / "daemon.log"
CONFIG_FILE = APP_DIR / "config.txt"
IMAGE_PATH = USER_DESKTOP / "ddolomites_latest.jpg"

# 资源定位：兼容 PyInstaller --onefile（sys._MEIPASS）
def resource_path(rel_name: str) -> Path:
    """
    获取打包/开发两种模式下的资源绝对路径。
    - 打包(--onefile)时，PyInstaller 会把 --add-data 解压到 sys._MEIPASS。
    - 开发模式下，返回当前脚本目录下的相对路径。
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / rel_name
    return Path(__file__).parent / rel_name

# Logo 查找顺序：开发机绝对路径 → 打包资源(_MEIPASS) → EXE 同目录 → APP_DIR
def pick_logo_path() -> Path | None:
    candidates = [
        resource_path("logo.png"),
        (Path(sys.executable).parent / "logo.png") if getattr(sys, "frozen", False) else None,
        APP_DIR / "logo.png",
    ]
    for p in candidates:
        if p and p.exists():
            return p
    return None

stop_event = threading.Event()
status_lock = threading.Lock()
last_status = "Idle"

# -------------------- 配置管理 --------------------
DEFAULT_CONFIG = {
    "base_url": "https://www.megacam.at/webcam/6er-Sesselbahn/",
    "interval_min": "60",               # 默认 60 分钟
    "wallpaper_style": "stretch",          # fill / stretch / tile
}

def load_config():
    if not CONFIG_FILE.exists():
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    cfg = {}
    for line in CONFIG_FILE.read_text(encoding="utf-8").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            cfg[k.strip()] = v.strip()
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    return cfg

def save_config(cfg):
    CONFIG_FILE.write_text("\n".join(f"{k}={v}" for k, v in cfg.items()), encoding="utf-8")

config = load_config()

# -------------------- 日志与状态 --------------------
def log(msg, also_status=False):
    ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{ts}] {msg}\n")
    if also_status:
        set_status(msg)

def set_status(msg):
    global last_status
    with status_lock:
        last_status = msg

def get_status():
    with status_lock:
        return last_status

# -------------------- 时间逻辑 --------------------
def rome_now():
    return dt.datetime.now(dt.timezone.utc).astimezone(ROME_TZ)

def most_recent_slot_rome(interval_min: int) -> dt.datetime:
    """
    返回罗马时区“上一个整 interval_min 分钟”的时间戳（严格取上一槽）。
    例如：interval=10，当前 09:07 -> 09:00；当前正好 09:00 -> 08:50。
          interval=30，当前 09:05 -> 09:00；当前正好 09:30 -> 09:00。
          interval=60，当前 09:58 -> 09:00；当前正好 10:00 -> 09:00。
    """
    now_r = rome_now()
    # 为了在“恰好卡到边界”时取上一档，这里先往回退 1 秒再做 floor
    ref = now_r - dt.timedelta(seconds=1)
    minute_bucket = (ref.minute // interval_min) * interval_min
    floored = ref.replace(minute=minute_bucket, second=0, microsecond=0)
    return floored


# -------------------- 下载与壁纸 --------------------
def build_url(slot_dt: dt.datetime):
    base = config["base_url"].rstrip("/") + "/"
    return (
        f"{base}{slot_dt.year}/{slot_dt.month:02d}/{slot_dt.day:02d}/"
        f"{slot_dt.hour:02d}{slot_dt.minute:02d}_hu.jpg"
    )


def download_image_to_desktop():
    # 读取当前间隔；非法则视为 60
    try:
        interval = int(config.get("interval_min", "60"))
        if interval not in (10, 30, 60):
            interval = 60
    except ValueError:
        interval = 60

    # 从“上一个整 interval 分钟”的槽位开始尝试，最多回退 6 档
    set_status("Fetching image...")
    start_slot = most_recent_slot_rome(interval)

    for back in range(0, 6 + 1):
        slot = start_slot - dt.timedelta(minutes=interval * back)
        url = build_url(slot)
        log(f"Fetching: {url}")
        try:
            r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
            ctype = r.headers.get("Content-Type", "")
            if r.status_code == 200 and r.content and "image" in ctype.lower():
                tmp = IMAGE_PATH.with_suffix(".part")
                with open(tmp, "wb") as f:
                    f.write(r.content)
                tmp.replace(IMAGE_PATH)
                log(f"Saved to: {IMAGE_PATH}")
                return IMAGE_PATH
            else:
                log(f"Missed slot {slot.strftime('%Y-%m-%d %H:%M')} — HTTP {r.status_code}, ctype={ctype}")
        except Exception as e:
            log(f"Download error on {url}: {e}")

    return None


def apply_wallpaper_style(style: str):
    """
    style: 'fill' | 'stretch' | 'tile'
    对应注册表：
      - 填充 Fill:    WallpaperStyle=10, TileWallpaper=0
      - 拉伸 Stretch:  WallpaperStyle=2,  TileWallpaper=0
      - 平铺 Tile:    WallpaperStyle=0,  TileWallpaper=1
    """
    style = style.lower()
    wp_style, tile = "10", "0"  # 默认 fill
    if style == "stretch":
        wp_style, tile = "2", "0"
    elif style == "tile":
        wp_style, tile = "0", "1"

    try:
        key = reg.OpenKey(reg.HKEY_CURRENT_USER, r"Control Panel\Desktop", 0, reg.KEY_SET_VALUE)
        reg.SetValueEx(key, "WallpaperStyle", 0, reg.REG_SZ, wp_style)
        reg.SetValueEx(key, "TileWallpaper", 0, reg.REG_SZ, tile)
        reg.CloseKey(key)
        log(f"Applied wallpaper style: {style}")
    except Exception as e:
        log(f"Registry write error (style): {e}")

def set_wallpaper(image_path):
    # 先应用样式，再刷新壁纸
    apply_wallpaper_style(config.get("wallpaper_style", "fill"))
    try:
        ok = ctypes.windll.user32.SystemParametersInfoW(
            SPI_SETDESKWALLPAPER, 0, str(image_path),
            SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
        )
        log("Wallpaper updated." if ok else "Set wallpaper failed.", also_status=True)
        return bool(ok)
    except Exception as e:
        log(f"Set wallpaper error: {e}", also_status=True)
        return False

# -------------------- 主任务循环 --------------------
def run_once():
    set_status("Running...")
    img = download_image_to_desktop()
    if img:
        set_wallpaper(img)
    else:
        # 即便下载失败，也刷新样式到当前壁纸，避免用户以为没生效
        apply_wallpaper_style(config.get("wallpaper_style", "fill"))
        set_status("No image (check log).")

def worker_loop():
    log("DDolomitesWpaper daemon started.", also_status=True)
    run_once()
    while not stop_event.is_set():
        try:
            interval_min = int(config.get("interval_min", "60"))
        except ValueError:
            interval_min = 60
        set_status(f"Waiting {interval_min} minutes...")
        for _ in range(interval_min * 60):
            if stop_event.is_set():
                return
            time.sleep(1)
        run_once()

# -------------------- GUI 窗口工具 --------------------
def center_window(win: tk.Tk | tk.Toplevel, w: int, h: int):
    """把窗口居中到屏幕"""
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = int((sw - w) / 2)
    y = int((sh - h) / 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

# -------------------- GUI 窗口 --------------------
def load_logo_pil(size=(64,64)):
    p = pick_logo_path()
    if p:
        try:
            img = Image.open(p).convert("RGBA")
            if size:
                img = img.resize(size, Image.LANCZOS)
            return img
        except Exception:
            pass
    # fallback 内置简易图标
    size = size or (64,64)
    img = Image.new("RGBA", size, (32, 32, 36, 255))
    d = ImageDraw.Draw(img)
    d.polygon([(8,48),(24,28),(36,44),(48,24),(56,48)], fill=(200,200,210,255))
    d.ellipse((14,12,30,28), fill=(220,220,230,255))
    d.rectangle((0,0,size[0]-1,size[1]-1), outline=(90,90,98,255))
    return img

def load_logo_for_tk():
    pil = load_logo_pil(size=(64,64))
    return ImageTk.PhotoImage(pil) if pil else None

def open_url_window():
    root = tk.Tk()
    root.title("Change Base URL - DDolomitesWpaper")
    center_window(root, 520, 170)
    try:
        logo_img = load_logo_for_tk()
        if logo_img:
            root.iconphoto(False, logo_img)
    except Exception:
        pass

    ttk.Label(root, text="Enter new base URL:").pack(pady=8)
    url_var = tk.StringVar(value=config["base_url"])
    ttk.Entry(root, textvariable=url_var, width=70).pack(pady=5)

    def save():
        new_url = url_var.get().strip()
        if new_url:
            config["base_url"] = new_url
            save_config(config)
            log(f"Base URL changed to: {new_url}", also_status=True)
            messagebox.showinfo("Success", "URL updated. Fetching now.")
            threading.Thread(target=run_once, daemon=True).start()
            root.destroy()
        else:
            messagebox.showwarning("Invalid", "URL cannot be empty.")
    ttk.Button(root, text="Save", command=save).pack(pady=10)
    root.mainloop()

def open_interval_window():
    root = tk.Tk()
    root.title("Set Interval - DDolomitesWpaper")
    center_window(root, 320, 170)
    try:
        logo_img = load_logo_for_tk()
        if logo_img:
            root.iconphoto(False, logo_img)
    except Exception:
        pass

    ttk.Label(root, text="Select update interval (minutes):").pack(pady=8)
    combo_var = tk.StringVar(value=config["interval_min"])
    combo = ttk.Combobox(root, textvariable=combo_var, values=["10", "30", "60"], width=10, state="readonly")
    combo.pack(pady=5)

    def save_interval():
        sel = combo_var.get()
        if sel:
            config["interval_min"] = sel
            save_config(config)
            log(f"Interval changed to {sel} minutes", also_status=True)
            messagebox.showinfo("Success", f"Interval set to {sel} minutes")
            root.destroy()
    ttk.Button(root, text="Save", command=save_interval).pack(pady=10)
    root.mainloop()

def open_about_window():
    root = tk.Tk()
    root.title("About - DDolomitesWpaper")
    center_window(root, 420, 320)
    try:
        logo_img = load_logo_for_tk()
        if logo_img:
            root.iconphoto(False, logo_img)
    except Exception:
        pass

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="both", expand=True)

    # 显示 logo
    pil_logo = load_logo_pil(size=(96,96))
    if pil_logo:
        tk_logo = ImageTk.PhotoImage(pil_logo)
        lbl_logo = ttk.Label(frm, image=tk_logo)
        lbl_logo.image = tk_logo  # 防 GC
        lbl_logo.pack(pady=(0,8))

    ttk.Label(frm, text="DDolomitesWpaper", font=("Segoe UI", 14, "bold")).pack(pady=2)
    ttk.Label(frm, text="A configurable wallpaper updater", font=("Segoe UI", 10)).pack(pady=2)
    ttk.Separator(frm).pack(fill="x", pady=8)

    ttk.Label(frm, text="Developer: Maguamale").pack()
    mail_lbl = ttk.Label(frm, text="Email: zhengzh@email.com", foreground="#0066cc", cursor="hand2")
    mail_lbl.pack()
    mail_lbl.bind("<Button-1>", lambda e: webbrowser.open("mailto:zhengzh@email.com"))

    gh_lbl = ttk.Label(frm, text="GitHub: www.github.com/zhengzihao", foreground="#0066cc", cursor="hand2")
    gh_lbl.pack()
    gh_lbl.bind("<Button-1>", lambda e: webbrowser.open("https://www.github.com/zhengzihao"))

    ttk.Separator(frm).pack(fill="x", pady=10)
    ttk.Button(frm, text="Close", command=root.destroy).pack(pady=4)
    root.mainloop()

# -------------------- 托盘菜单 --------------------
def is_style(style):
    return lambda item: config.get("wallpaper_style", "fill").lower() == style

def set_style(style):
    config["wallpaper_style"] = style
    save_config(config)
    apply_wallpaper_style(style)
    ctypes.windll.user32.SystemParametersInfoW(
        SPI_SETDESKWALLPAPER, 0, str(IMAGE_PATH),
        SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
    )
    log(f"Wallpaper style set to: {style}", also_status=True)

def action_run_now(icon, item):
    threading.Thread(target=run_once, daemon=True).start()

def action_toggle_startup(icon, item):
    # 取当前状态，切换之
    cur = is_startup_enabled()
    ok = enable_startup(not cur)
    if ok:
        # 切换后刷新菜单勾选状态
        try:
            icon.update_menu()
        except Exception:
            pass

def action_open_log(icon, item):
    os.startfile(str(APP_DIR))

def action_open_image(icon, item):
    if IMAGE_PATH.exists():
        os.startfile(str(IMAGE_PATH))
    else:
        log("No image yet.")

def action_exit(icon, item):
    stop_event.set()
    icon.visible = False
    icon.stop()

def main():
    t = threading.Thread(target=worker_loop, daemon=True)
    t.start()

    image = load_logo_pil(size=(64,64))
    style_menu = pystray.Menu(
        pystray.MenuItem("Fill",   lambda icon, item: set_style("fill"),    checked=is_style("fill")),
        pystray.MenuItem("Stretch",lambda icon, item: set_style("stretch"), checked=is_style("stretch")),
        pystray.MenuItem("Tile",   lambda icon, item: set_style("tile"),    checked=is_style("tile")),
    )
    # 关键：开机自启菜单项（checked 动态显示当前状态）
    startup_item = pystray.MenuItem(
        "Auto Startup",
        action_toggle_startup,
        checked=lambda item: is_startup_enabled()
    )
    menu = pystray.Menu(
        pystray.MenuItem(lambda item: f"Status: {get_status()}", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Run now", action_run_now),
        pystray.MenuItem("Open latest image", action_open_image),
        pystray.MenuItem("Open log folder", action_open_log),
        startup_item,
        pystray.MenuItem("Change URL", lambda icon, item: threading.Thread(target=open_url_window,     daemon=True).start()),
        pystray.MenuItem("Set interval", lambda icon, item: threading.Thread(target=open_interval_window, daemon=True).start()),
        pystray.MenuItem("Wallpaper style", style_menu),

        pystray.Menu.SEPARATOR,
        pystray.MenuItem("About", lambda icon, item: threading.Thread(target=open_about_window, daemon=True).start()),
        pystray.MenuItem("Exit", action_exit),
    )

    icon = pystray.Icon("DDolomitesWpaper", image, "DDolomites Wallpaper", menu)
    icon.run()
    stop_event.set()
    t.join(timeout=5)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"Fatal error: {e}")
        time.sleep(2)
