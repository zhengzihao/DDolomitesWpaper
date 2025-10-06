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
#   作者：子豪定制版
# ============================================================

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

# Logo 搜索路径（优先用你给的绝对路径）
LOGO_CANDIDATES = [
    Path(r"C:\Users\zhengzh\PycharmProjects\typicalProject\logo.png"),
    Path(sys.executable).parent / "logo.png" if getattr(sys, "frozen", False) else None,
    APP_DIR / "logo.png",
]

stop_event = threading.Event()
status_lock = threading.Lock()
last_status = "Idle"

# -------------------- 配置管理 --------------------
DEFAULT_CONFIG = {
    "base_url": "https://www.megacam.at/webcam/6er-Sesselbahn/",
    "interval_min": "60",               # 默认 60 分钟
    "wallpaper_style": "fill",          # fill / stretch / tile
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

def most_recent_full_hour_rome():
    now_r = rome_now()
    floored = now_r.replace(minute=0, second=0, microsecond=0)
    return floored - dt.timedelta(hours=1)

def seconds_until_next_hour_at(minute=1):
    """每小时罗马时间 01 分执行（备用接口，如需按小时对齐可使用）"""
    now_r = rome_now()
    next_hour = (now_r.replace(minute=0, second=0, microsecond=0)
                 + dt.timedelta(hours=1))
    target = next_hour.replace(minute=minute, second=0, microsecond=0)
    return max(5, int((target - now_r).total_seconds()))

# -------------------- 下载与壁纸 --------------------
def build_url(rome_hour):
    base = config["base_url"].rstrip("/") + "/"
    return f"{base}{rome_hour.year}/{rome_hour.month:02d}/{rome_hour.day:02d}/{rome_hour.hour:02d}00_hu.jpg"

def download_image_to_desktop():
    rh = most_recent_full_hour_rome()
    url = build_url(rh)
    set_status("Fetching image...")
    log(f"Fetching: {url}")

    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        if r.status_code == 200 and r.content:
            with open(IMAGE_PATH, "wb") as f:
                f.write(r.content)
            log(f"Saved to: {IMAGE_PATH}")
            return IMAGE_PATH
        else:
            log(f"HTTP {r.status_code} or empty content.")
            return None
    except Exception as e:
        log(f"Download error: {e}")
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

# -------------------- GUI 窗口 --------------------
def open_url_window():
    def save():
        new_url = url_var.get().strip()
        if new_url:
            config["base_url"] = new_url
            save_config(config)
            log(f"Base URL changed to: {new_url}", also_status=True)
            messagebox.showinfo("Success", "URL updated successfully! Will fetch now.")
            # 立即运行一次以使用新 URL
            threading.Thread(target=run_once, daemon=True).start()
            root.destroy()
        else:
            messagebox.showwarning("Invalid", "URL cannot be empty.")
    root = tk.Tk()
    root.title("Change Base URL - DDolomitesWpaper")
    root.geometry("520x170")
    try:
        logo_img = load_logo_for_tk()
        if logo_img:
            root.iconphoto(False, logo_img)
    except Exception:
        pass
    ttk.Label(root, text="Enter new base URL:").pack(pady=8)
    url_var = tk.StringVar(value=config["base_url"])
    ttk.Entry(root, textvariable=url_var, width=70).pack(pady=5)
    ttk.Button(root, text="Save", command=save).pack(pady=10)
    root.mainloop()

def open_interval_window():
    def save_interval():
        sel = combo_var.get()
        if sel:
            config["interval_min"] = sel
            save_config(config)
            log(f"Interval changed to {sel} minutes", also_status=True)
            messagebox.showinfo("Success", f"Interval set to {sel} minutes")
            root.destroy()
    root = tk.Tk()
    root.title("Set Interval - DDolomitesWpaper")
    root.geometry("320x170")
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
    ttk.Button(root, text="Save", command=save_interval).pack(pady=10)
    root.mainloop()

def open_about_window():
    root = tk.Tk()
    root.title("About - DDolomitesWpaper")
    root.geometry("420x300")
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
        lbl_logo.image = tk_logo  # 防止被 GC
        lbl_logo.pack(pady=(0,8))

    ttk.Label(frm, text="DDolomitesWpaper", font=("Segoe UI", 14, "bold")).pack(pady=2)
    ttk.Label(frm, text="A configurable wallpaper updater", font=("Segoe UI", 10)).pack(pady=2)
    ttk.Separator(frm).pack(fill="x", pady=8)

    ttk.Label(frm, text="Developer: Maguamale").pack()
    mail_lbl = ttk.Label(frm, text="Email: zhengzh@email.com", foreground="#0066cc", cursor="hand2")
    mail_lbl.pack()
    mail_lbl.bind("<Button-1>", lambda e: webbrowser.open("mailto:zhengzh@email.com"))

    gh_lbl = ttk.Label(frm, text="GitHub: www.github.com/maguamale", foreground="#0066cc", cursor="hand2")
    gh_lbl.pack()
    gh_lbl.bind("<Button-1>", lambda e: webbrowser.open("https://www.github.com/maguamale"))

    ttk.Separator(frm).pack(fill="x", pady=10)
    ttk.Button(frm, text="Close", command=root.destroy).pack(pady=4)
    root.mainloop()

# -------------------- 托盘图标与资源 --------------------
def pick_logo_path():
    for p in LOGO_CANDIDATES:
        if p and p.exists():
            return p
    return None

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
    # 兜底内置图标
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

# -------------------- 菜单动作 --------------------
def action_run_now(icon, item):
    threading.Thread(target=run_once, daemon=True).start()

def action_open_log(icon, item):
    os.startfile(str(APP_DIR))

def action_open_image(icon, item):
    if IMAGE_PATH.exists():
        os.startfile(str(IMAGE_PATH))
    else:
        log("No image yet.")

def set_style(style):
    config["wallpaper_style"] = style
    save_config(config)
    # 立即应用到当前壁纸（即便不更换图片）
    apply_wallpaper_style(style)
    ctypes.windll.user32.SystemParametersInfoW(
        SPI_SETDESKWALLPAPER, 0, str(IMAGE_PATH),
        SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
    )
    log(f"Wallpaper style set to: {style}", also_status=True)

def is_style(style):
    return lambda item: config.get("wallpaper_style", "fill").lower() == style

def action_exit(icon, item):
    stop_event.set()
    icon.visible = False
    icon.stop()

# -------------------- 主程序入口 --------------------
def main():
    t = threading.Thread(target=worker_loop, daemon=True)
    t.start()

    image = load_logo_pil(size=(64,64))
    # 壁纸样式子菜单
    style_menu = pystray.Menu(
        pystray.MenuItem("Fill (填充)", lambda icon, item: set_style("fill"), checked=is_style("fill")),
        pystray.MenuItem("Stretch (拉伸)", lambda icon, item: set_style("stretch"), checked=is_style("stretch")),
        pystray.MenuItem("Tile (平铺)", lambda icon, item: set_style("tile"), checked=is_style("tile")),
    )

    menu = pystray.Menu(
        pystray.MenuItem(lambda item: f"Status: {get_status()}", None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Run now", action_run_now),
        pystray.MenuItem("Open log folder", action_open_log),
        pystray.MenuItem("Open latest image", action_open_image),
        pystray.MenuItem("Change URL", lambda icon, item: threading.Thread(target=open_url_window, daemon=True).start()),
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
