# 測試過程有衰減就會回測,回測的前兩秒不列入SPEC判斷
# 排除BT和touchpad的錯誤訊息
import os
import sys
import threading
import time
import collections
import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkfont
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta
import re
import math
import ctypes
import ctypes.wintypes as wintypes
from ctypes import cdll, c_int
from time import sleep
from queue import Queue, Empty

# ✅ Step 1: 新增 bisect（用於固定視窗 slicing）
import bisect

# === 新增: 匯入 Socket 與 N9010A 螢幕擷取函式 ===
import socket
from n9010a_capture import save_screenshot

# === 新增: UI 截圖（Tkinter 視窗）===
try:
    from PIL import ImageGrab  # Pillow
    PIL_AVAILABLE = True
except Exception:
    ImageGrab = None
    PIL_AVAILABLE = False

# ----------------------------------------------------
# === PyInstaller base dir（onefile/onefolder 相容） ===
from pathlib import Path


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        if hasattr(sys, "_MEIPASS"):
            return Path(sys._MEIPASS)
        return Path(sys.executable).parent
    return Path(__file__).parent


BASE_DIR = get_base_dir()
os.add_dll_directory(str(BASE_DIR))  # 讓 VNX_atten64.dll 能被載入

# === N9010A 預設連線參數 ===
N9010A_DEFAULT_IP = "169.254.103.141"
N9010A_PORT = 5025

# ----------------------------------------------------
# === VNX 衰減器 (Attenuator) CLI 初始化 ===
try:
    vnx = cdll.VNX_atten64
    vnx.fnLDA_SetTestMode(False)
    num_devices = vnx.fnLDA_GetNumDevices()
    if num_devices == 0:
        print("❌ 未偵測到任何 VNX 衰減器設備")
        sys.exit(1)

    DeviceArray = c_int * num_devices
    devices = DeviceArray()
    vnx.fnLDA_GetDevInfo(devices)
    devid = devices[0]
    vnx.fnLDA_InitDevice(devid)
    print(f"✅ 已初始化 VNX 裝置 (ID = {devid})，共偵測到 {num_devices} 台。")
except Exception as e:
    print("❌ 載入或初始化 VNX 衰減器失敗：", e)
    sys.exit(1)

# ----------------------------------------------------
# ✅ ✅ 重要修正：衰減設定改為「非阻塞」背景執行緒處理
# ----------------------------------------------------
current_atten = 0.0
target_atten = None   # ✅ 主邏輯用：最後一次「決策要設定」的衰減值（避免非同步延遲）
current_label = None
root = None

_atten_cmd_q: "Queue[float]" = Queue()
_atten_worker_started = False
_atten_worker_stop = False
_atten_last_sent = None
_atten_lock = threading.Lock()


def _atten_worker():
    global current_atten, _atten_worker_stop, _atten_last_sent
    while not _atten_worker_stop:
        try:
            target_db = _atten_cmd_q.get(timeout=0.2)
        except Empty:
            continue

        # 只保留最後一筆（快速連續變更時，避免堆一堆命令造成延遲）
        while True:
            try:
                target_db = _atten_cmd_q.get_nowait()
            except Empty:
                break

        try:
            db_int = int(float(target_db) * 20)  # 0.05 dB step
            vnx.fnLDA_SetAttenuationHR(devid, db_int)

            # ❗不要在主 loop sleep；背景 worker 可給極短暫 settle（可調小）
            sleep(0.02)

            with _atten_lock:
                current_atten = float(target_db)
                _atten_last_sent = current_atten

            # UI 更新一定要回主執行緒
            if root and current_label:
                def _ui_update():
                    if current_label:
                        current_label.config(text=f"{current_atten:.1f} dB")
                root.after(0, _ui_update)

        except Exception as e:
            print("⚠ 設定衰減失敗：", e)


def ensure_atten_worker():
    global _atten_worker_started
    if _atten_worker_started:
        return
    _atten_worker_started = True
    t = threading.Thread(target=_atten_worker, daemon=True)
    t.start()


def set_attenuation_db(db_value: float):
    """
    非阻塞：只把命令送到 queue，由背景 worker 處理。
    """
    ensure_atten_worker()
    try:
        _atten_cmd_q.put(float(db_value))
    except Exception as e:
        print("⚠ set_attenuation_db enqueue 失敗：", e)


# ✅✅✅ 新增：主邏輯用的衰減設定入口（避免 worker 更新延遲造成 stop/step 判斷飄移）
def set_attenuation_target(db_value: float):
    """
    ✅ 主邏輯用的衰減設定入口：
    - 立刻更新 target_atten（決策基準）
    - 再丟給 worker 去設定硬體（非阻塞）
    """
    global target_atten
    try:
        target_atten = float(db_value)
    except Exception:
        target_atten = db_value
    set_attenuation_db(db_value)


# ----------------------------------------------------
# === Win32 Raw Input 與 GUI 相關程式碼 ===

wintypes.WPARAM = ctypes.c_ulonglong
wintypes.LRESULT = ctypes.c_longlong
wintypes.LPARAM = ctypes.c_longlong
wintypes.HICON = wintypes.HANDLE
wintypes.HCURSOR = wintypes.HANDLE
wintypes.HBRUSH = wintypes.HANDLE

user32 = ctypes.windll.user32
user32.DefWindowProcW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM
]
user32.DefWindowProcW.restype = wintypes.LRESULT


class RAWINPUTHEADER(ctypes.Structure):
    _fields_ = [
        ('dwType', wintypes.DWORD),
        ('dwSize', wintypes.DWORD),
        ('hDevice', wintypes.HANDLE),
        ('wParam', wintypes.WPARAM),
    ]


class RAWMOUSE(ctypes.Structure):
    _fields_ = [
        ('usFlags', wintypes.USHORT),
        ('ulButtons', wintypes.ULONG),
        ('usButtonFlags', wintypes.USHORT),
        ('usButtonData', wintypes.USHORT),
        ('ulRawButtons', wintypes.ULONG),
        ('lLastX', ctypes.c_long),
        ('lLastY', ctypes.c_long),
        ('ulExtraInformation', wintypes.ULONG),
    ]


class RAWINPUT(ctypes.Structure):
    _fields_ = [
        ('header', RAWINPUTHEADER),
        ('data', RAWMOUSE),
    ]


class RAWINPUTDEVICE(ctypes.Structure):
    _fields_ = [
        ('usUsagePage', wintypes.USHORT),
        ('usUsage', wintypes.USHORT),
        ('dwFlags', wintypes.DWORD),
        ('hwndTarget', wintypes.HWND),
    ]


RIDEV_INPUTSINK = 0x00000100
RID_INPUT = 0x10000003
WM_INPUT = 0x00FF


def get_cursor_pos():
    pt = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    return (pt.x, pt.y)


def is_bluetooth_device(hdev) -> bool:
    RIDI_DEVICENAME = 0x20000007
    size = wintypes.UINT(0)
    user32.GetRawInputDeviceInfoW(hdev, RIDI_DEVICENAME, None, ctypes.byref(size))
    if size.value == 0:
        return False
    buf = ctypes.create_unicode_buffer(size.value)
    user32.GetRawInputDeviceInfoW(hdev, RIDI_DEVICENAME, buf, ctypes.byref(size))
    name = buf.value.lower()
    return 'bthenum' in name or 'bluetooth' in name


def is_touchpad_device(hdev) -> bool:
    RIDI_DEVICENAME = 0x20000007
    size = wintypes.UINT(0)
    user32.GetRawInputDeviceInfoW(hdev, RIDI_DEVICENAME, None, ctypes.byref(size))
    if size.value == 0:
        return False
    buf = ctypes.create_unicode_buffer(size.value)
    user32.GetRawInputDeviceInfoW(hdev, RIDI_DEVICENAME, buf, ctypes.byref(size))
    name = buf.value.lower()
    tp_keywords = (
        "touchpad", "precision touchpad", "synaptics", "elan", "etd",
        "hid-compliant touch pad", "i2c hid device"
    )
    return any(k in name for k in tp_keywords)


def get_device_pid_vid(hdev):
    RIDI_DEVICENAME = 0x20000007
    name_size = wintypes.UINT(0)
    user32.GetRawInputDeviceInfoW(hdev, RIDI_DEVICENAME, None, ctypes.byref(name_size))
    if name_size.value == 0:
        return None, None

    buf = ctypes.create_unicode_buffer(name_size.value)
    user32.GetRawInputDeviceInfoW(hdev, RIDI_DEVICENAME, buf, ctypes.byref(name_size))
    dev_name = buf.value
    m = re.search(r'VID_([0-9A-Fa-f]+)&PID_([0-9A-Fa-f]+)', dev_name)
    if m:
        return (m.group(1), m.group(2))
    return (None, None)


def make_wndclass():
    WNDPROCTYPE = ctypes.WINFUNCTYPE(
        wintypes.LRESULT,
        wintypes.HWND,
        wintypes.UINT,
        wintypes.WPARAM,
        wintypes.LPARAM
    )

    class WNDCLASS(ctypes.Structure):
        _fields_ = [
            ('style', wintypes.UINT),
            ('lpfnWndProc', WNDPROCTYPE),
            ('cbClsExtra', ctypes.c_int),
            ('cbWndExtra', ctypes.c_int),
            ('hInstance', wintypes.HINSTANCE),
            ('hIcon', wintypes.HICON),
            ('hCursor', wintypes.HCURSOR),
            ('hbrBackground', wintypes.HBRUSH),
            ('lpszMenuName', wintypes.LPCWSTR),
            ('lpszClassName', wintypes.LPCWSTR),
        ]

    return WNDCLASS, WNDPROCTYPE


# ----------------------------------------------------
# ====== VID/PID 對應資料庫載入（支援 JSON/CSV） ======
import json

DEVICE_DB = {}
DEVICE_DB_PATH = None


def _candidate_db_paths():
    paths = []
    paths.append(Path.cwd() / "device_db.json")
    paths.append(Path.cwd() / "device_db.csv")
    try:
        exe_dir = Path(sys.executable).parent
        paths.append(exe_dir / "device_db.json")
        paths.append(exe_dir / "device_db.csv")
    except Exception:
        pass
    paths.append(BASE_DIR / "device_db.json")
    paths.append(BASE_DIR / "device_db.csv")
    return paths


def load_device_db():
    db = {}
    loaded_path = None
    for p in _candidate_db_paths():
        if not p.exists():
            continue
        try:
            if p.suffix.lower() == ".json":
                with open(p, "r", encoding="utf-8") as f:
                    items = json.load(f)
                for it in items:
                    vid = str(it.get("vid", "")).strip().upper()
                    pid = str(it.get("pid", "")).strip().upper()
                    if len(vid) == 4 and len(pid) == 4:
                        db.setdefault((vid, pid), []).append({
                            "brand": (it.get("brand", "") or "").strip(),
                            "model": (it.get("model", "") or "").strip()
                        })
                loaded_path = str(p)
                break
            else:
                import csv as _csv
                with open(p, newline="", encoding="utf-8") as f:
                    r = _csv.DictReader(f)
                    for row in r:
                        vid = str(row.get("vid", "")).strip().upper()
                        pid = str(row.get("pid", "")).strip().upper()
                        if len(vid) == 4 and len(pid) == 4:
                            db.setdefault((vid, pid), []).append({
                                "brand": (row.get("brand", "") or "").strip(),
                                "model": (row.get("model", "") or "").strip()
                            })
                loaded_path = str(p)
                break
        except Exception as e:
            print(f"⚠ 載入裝置資料庫失敗：{p} -> {e}")
    return db, loaded_path


DEVICE_DB, DEVICE_DB_PATH = load_device_db()


def lookup_brand_model(vid: str, pid: str) -> str | None:
    entries = DEVICE_DB.get((str(vid).upper(), str(pid).upper()))
    if not entries:
        return None
    if len(entries) == 1:
        e = entries[0]
        return f"{e.get('brand', '').strip()} {e.get('model', '').strip()}".strip()
    parts = []
    for e in entries[:3]:
        parts.append(f"{e.get('brand', '').strip()} {e.get('model', '').strip()}".strip())
    s = " / ".join([p for p in parts if p])
    if len(entries) > 3:
        s += " +…"
    return s


# ----------------------------------------------------
# === 新增: 儀器 Socket 連線檢查 ===
def check_socket(host: str, port: int, timeout: float = 5.0) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(timeout)
        try:
            sock.connect((host, port))
            print(f"[Socket] 成功連線到 {host}:{port}")
            return True
        except socket.timeout:
            print(f"[Socket] 連線 {host}:{port} 逾時 ({timeout} 秒)。")
            return False
        except ConnectionRefusedError:
            print(f"[Socket] {host}:{port} 拒絕連線，請檢查儀器設定。")
            return False
        except socket.error as exc:
            print(f"[Socket] {host}:{port} 發生錯誤: {exc}")
            return False


def trigger_async_capture(atten_before: float) -> None:
    ip = n9010a_ip_var.get().strip() if n9010a_ip_var else ""
    if not ip:
        print("[螢幕擷取] 未設定 N9010A IP，跳過擷取。")
        return

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"fail_capture_{timestamp}_{atten_before:.1f}dB.png"
    output_path = Path.cwd() / filename

    def _worker() -> None:
        print(f"[螢幕擷取] 嘗試擷取 {ip} 畫面，儲存 {output_path}")
        try:
            success = save_screenshot(
                ip=ip,
                output_path=str(output_path),
                verbose=True,
            )
            if success:
                print(f"[螢幕擷取] 成功儲存: {output_path}")
            else:
                print("[螢幕擷取] 失敗，請檢查儀器連線設定。")
        except Exception as exc:
            print(f"[螢幕擷取] 擷取過程發生錯誤: {exc}")

    threading.Thread(target=_worker, daemon=True).start()


# ----------------------------------------------------
# === 新增: UI 視窗截圖（只用於 FAIL，存檔路徑同 Path.cwd()） ===
ui_fail_captured = False  # ✅ 每次測試只截一次 UI


def trigger_async_ui_capture(atten_before: float) -> None:
    """
    觸發 UI 視窗截圖（PNG）
    - 存檔路徑：Path.cwd()（與 N9010A 擷取相同）
    - 檔名：fail_ui_時間戳_衰減dB.png
    """
    global root

    if root is None:
        print("[UI截圖] root 尚未建立，跳過。")
        return

    if not PIL_AVAILABLE:
        print("[UI截圖] Pillow(ImageGrab) 未安裝或不可用，跳過 UI 截圖。")
        return

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"fail_ui_{timestamp}_{atten_before:.1f}dB.png"
    output_path = Path.cwd() / filename

    def _capture_on_ui_thread():
        try:
            root.update_idletasks()

            x = root.winfo_rootx()
            y = root.winfo_rooty()
            w = root.winfo_width()
            h = root.winfo_height()

            # 避免剛建立時 width/height 還是 1
            if w <= 2 or h <= 2:
                root.update()
                w = root.winfo_width()
                h = root.winfo_height()

            bbox = (x, y, x + w, y + h)

            def _worker_save():
                try:
                    img = ImageGrab.grab(bbox=bbox, all_screens=True)
                    img.save(str(output_path), "PNG")
                    print(f"[UI截圖] 成功儲存: {output_path}")
                except Exception as exc:
                    print(f"[UI截圖] 擷取/存檔失敗: {exc}")

            threading.Thread(target=_worker_save, daemon=True).start()

        except Exception as exc:
            print(f"[UI截圖] 取得視窗座標失敗: {exc}")

    try:
        root.after(0, _capture_on_ui_thread)
    except Exception as exc:
        print(f"[UI截圖] root.after 失敗: {exc}")


def trigger_ui_capture_on_first_fail(atten_before: float) -> None:
    """
    ✅ 只在「第一次 FAIL」截 UI
    """
    global ui_fail_captured
    if ui_fail_captured:
        return
    ui_fail_captured = True
    trigger_async_ui_capture(atten_before)


# ----------------------------------------------------
# === 全域狀態與資料結構 ===
monitoring = False
chart_start_time = None
ani = None
last_event_time = time.monotonic()
y_max_fixed = None
spec_line = None
active_spec_value = None
active_spec_is_manual = False

spec_hold_start = None
half_phase = False
half_step_time = None
spec_hold_duration = 10.0

ignored_devices = set()

show_traj_var = None
show_traj_enabled = False

show_rate_var = None
show_rate_enabled = False
show_rate_locked = False

# Auto test
auto_test_var = None
auto_test_locked = True
step_entry_widget = None
stop_entry_widget = None

prev_all_above = True
backtest_fail_count = 0
BACKTEST_MAX = 10

atten_var = None
step_var = None
stop_var = None
n9010a_ip_var = None

rate_frame = None
rate_labels = {}
start_label = None
stop_label = None
save_label = None

save_path = None

fig = None
ax_rate = None
ax_traj = None
canvas = None

start_time_str = None
stop_time_str = None

last_atten_change_time = None
atten_change_marks = []  # 記錄每次切換衰減時的 elapsed_time（秒）

# A 版：1 秒滑動視窗
device_event_times = {}
device_plot_rates = {}
plot_times = []
device_positions = {}
device_lines_rate = {}
device_colors = {}
device_vid_pid = {}
color_cycle = plt.rcParams['axes.prop_cycle'].by_key()['color']
next_color_index = 0

event_counter = {}
ms_report_data = {}

TIMEOUT = 5.0

# ✅ Step 2: 固定視窗額外秒數（前2秒 skip，不列入SPEC）
PLOT_WINDOW_EXTRA_SEC = 2.0

# ----------------------------------------------------
# ✅✅✅ 新增：Auto test 狀態機（「一 FAIL 立刻動作」）
# ----------------------------------------------------
AUTO_MODE_FULL = "FULL"
AUTO_MODE_BACKTEST = "BACKTEST"

auto_mode = AUTO_MODE_FULL

full_retry_count = 0
FULL_RETRY_MAX = 3

backtest_count = 0
BACKTEST_MAX = 10  # 保持一致（你原本也有）

pass_hold_start = None  # ✅ 2秒 skip 後開始，連續滿 spec_hold_duration 才算 PASS

initial_atten_at_start = None
full_step_direction = 1  # +1: full 往衰減增加走；-1: full 往衰減減少走

# ----------------------------------------------------
# === 方案B：SPEC 目標滑鼠選擇 + 高亮 ===
spec_target_var = None           # tk.StringVar
spec_target_menu = None          # OptionMenu widget
spec_target_map = {}             # display_text -> hdev
spec_target_hdev = None          # None=Auto(第一隻)

SPEC_TAG = "[SPEC] "
SPEC_BG = "#FFF2CC"
SPEC_FG = "#D32F2F"


def get_effective_spec_target():
    """
    回傳目前有效的 SPEC 目標 hdev：
    - 若使用者指定 -> 回指定
    - 若 Auto(None) -> 回第一隻偵測到的滑鼠
    """
    if spec_target_hdev is not None:
        return spec_target_hdev
    return next(iter(device_event_times.keys()), None)


def refresh_spec_target_visual():
    """
    方案B：更新 UI label 與 plot line 的高亮效果
    """
    target = get_effective_spec_target()

    # --- 更新左側 label 視覺 ---
    for hdev, lbl in rate_labels.items():
        is_target = (hdev == target)

        txt = lbl.cget("text")
        if txt.startswith(SPEC_TAG):
            txt = txt[len(SPEC_TAG):]

        if is_target:
            lbl.config(
                text=SPEC_TAG + txt,
                bg=SPEC_BG,
                fg=SPEC_FG,
                font=("Microsoft JhengHei", 12, "bold")
            )
        else:
            color = device_colors.get(hdev, "black")
            lbl.config(
                text=txt,
                bg=rate_frame.cget("bg") if rate_frame is not None else "SystemButtonFace",
                fg=color,
                font=("Microsoft JhengHei", 12)
            )

    # --- 更新 plot line 視覺 ---
    if show_rate_locked:
        for hdev, line in device_lines_rate.items():
            is_target = (hdev == target)
            if is_target:
                line.set_linewidth(3.0)
                line.set_alpha(1.0)
                line.set_zorder(10)
            else:
                line.set_linewidth(1.2)
                line.set_alpha(0.35)
                line.set_zorder(1)

        if canvas:
            canvas.draw_idle()


def add_device_to_spec_menu(hdev, dev_label: str):
    global spec_target_menu, spec_target_map, spec_target_var
    if spec_target_menu is None or spec_target_var is None:
        return

    display = f"{dev_label} | hdev={hdev}"
    if display in spec_target_map:
        return

    spec_target_map[display] = hdev

    try:
        menu = spec_target_menu["menu"]
        menu.add_command(label=display, command=lambda v=display: spec_target_var.set(v))
    except Exception as exc:
        print(f"[SPEC目標] 加入下拉選單失敗: {exc}")

    if str(spec_target_var.get()).startswith("(Auto)"):
        refresh_spec_target_visual()


def _apply_axis_for_spec(spec_val: int | None, *, from_manual: bool = False):
    global y_max_fixed, ax_rate, canvas
    if spec_val is None:
        y_max_fixed = None
        ax_rate.set_ylim(0, 1000)
        ax_rate.set_yticks(range(0, 1001, 500))
    else:
        if from_manual:
            y_max_fixed = int(math.ceil(spec_val / 0.85))
        else:
            if spec_val <= 1000:
                y_max_fixed = 1000
            elif spec_val <= 4000:
                y_max_fixed = 4000
            else:
                y_max_fixed = 8000

        step = max(1, y_max_fixed // 4)
        ax_rate.set_ylim(0, y_max_fixed)
        ax_rate.set_yticks(range(0, y_max_fixed + 1, step))

    if canvas:
        canvas.draw_idle()


def on_spec_button(val):
    global active_spec_value, active_spec_is_manual

    def _hide_manual():
        if spec_manual_entry.winfo_ismapped():
            spec_manual_entry.pack_forget()

    active_spec_value = None

    if val == "Manual":
        active_spec_is_manual = True
        if not spec_manual_entry.winfo_ismapped():
            spec_manual_entry.pack(side=tk.LEFT, padx=0)
        spec_selected_label.config(text="")
        txt = spec_manual_var.get().strip()
        manual_spec = int(txt) if txt.isdigit() else None
        active_spec_value = manual_spec
        _apply_axis_for_spec(manual_spec, from_manual=True)
        return

    if val in ("850", "3400", "6800"):
        active_spec_is_manual = False
        _hide_manual()
        active_spec_value = int(val)
        spec_selected_label.config(text=f"{active_spec_value} Hz")
        _apply_axis_for_spec(active_spec_value, from_manual=False)
        return

    if val == "None":
        active_spec_is_manual = False
        active_spec_value = 0
        _hide_manual()
        spec_selected_label.config(text="0 Hz")
        _apply_axis_for_spec(None)
        return

    _hide_manual()
    spec_selected_label.config(text="")
    _apply_axis_for_spec(None)


def on_manual_entry_change(*args):
    global active_spec_value, active_spec_is_manual
    txt = spec_manual_var.get().strip()
    active_spec_is_manual = True
    active_spec_value = int(txt) if txt.isdigit() else None
    _apply_axis_for_spec(active_spec_value, from_manual=True)


def focus_ring_button(parent, *, ring_color="#00C853", ring_thickness=4, **btn_kwargs):
    wrap = tk.Frame(
        parent,
        highlightbackground=ring_color,
        highlightcolor=ring_color,
        highlightthickness=0,
        bd=0
    )
    btn = tk.Button(wrap, **btn_kwargs)
    btn.pack(fill=tk.X)

    def on_focus_in(_):
        wrap.configure(highlightthickness=ring_thickness)

    def on_focus_out(_):
        wrap.configure(highlightthickness=0)

    btn.bind("<FocusIn>", on_focus_in)
    btn.bind("<FocusOut>", on_focus_out)
    btn.configure(takefocus=True, cursor="hand2")
    return btn, wrap


def focus_ring_checkbutton(parent, *, ring_color="#00C853", ring_thickness=3, **chk_kwargs):
    wrap = tk.Frame(
        parent,
        highlightbackground=ring_color,
        highlightcolor=ring_color,
        highlightthickness=0,
        bd=0
    )
    chk = tk.Checkbutton(wrap, **chk_kwargs)
    chk.pack(anchor='w')

    def on_focus_in(_):
        wrap.configure(highlightthickness=ring_thickness)

    def on_focus_out(_):
        wrap.configure(highlightthickness=0)

    chk.bind("<FocusIn>", on_focus_in)
    chk.bind("<FocusOut>", on_focus_out)
    chk.configure(takefocus=True, cursor="hand2")
    return chk, wrap


def _set_step_stop_state(enabled: bool):
    global step_entry_widget, stop_entry_widget
    state = "normal" if enabled else "disabled"
    if step_entry_widget is not None:
        step_entry_widget.config(state=state, disabledbackground="#e0e0e0")
    if stop_entry_widget is not None:
        stop_entry_widget.config(state=state, disabledbackground="#e0e0e0")


def setup_ui():
    global root, rate_frame, start_label, stop_label, save_label
    global spec_manual_entry, spec_manual_var, spec_selected_label
    global spec_line, fig, ax_rate, ax_traj, canvas
    global atten_var, step_var, stop_var, n9010a_ip_var
    global current_label
    global show_traj_var, show_traj_enabled
    global show_rate_var, show_rate_enabled
    global auto_test_var
    global step_entry_widget, stop_entry_widget
    global spec_target_var, spec_target_menu, spec_target_map, spec_target_hdev

    root = tk.Tk()
    root.title("多滑鼠 Report Rate & 軌跡 (A 版：事件視窗計數)")
    root.geometry("1000x700")
    root.protocol("WM_DELETE_WINDOW", on_app_close)

    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(family="Microsoft JhengHei", size=12)

    left = tk.Frame(root, width=520, bg="#ececec")
    left.pack(side=tk.LEFT, fill=tk.Y)
    left.pack_propagate(False)

    title_frame = tk.Frame(left)
    title_frame.pack(pady=5)
    tk.Label(
        title_frame,
        text="Report Rate (Hz)",
        font=("Microsoft JhengHei", 14, "bold")
    ).pack(side=tk.LEFT)

    rate_frame = tk.Frame(left)
    rate_frame.pack(pady=5)

    # === 方案B：SPEC 目標滑鼠下拉選單（放在 rate_frame 後、SPEC 前）===
    spec_target_map = {}
    spec_target_hdev = None

    target_frame = tk.Frame(left)
    target_frame.pack(fill=tk.X, pady=(0, 5))

    tk.Label(
        target_frame,
        text="SPEC目標滑鼠:",
        font=("Microsoft JhengHei", 12, "bold"),
        fg="#333"
    ).pack(side=tk.LEFT, padx=(0, 5))

    spec_target_var = tk.StringVar(value="(Auto) 第一隻偵測到的滑鼠")
    spec_target_menu = tk.OptionMenu(target_frame, spec_target_var, "(Auto) 第一隻偵測到的滑鼠")
    spec_target_menu.config(font=("Microsoft JhengHei", 11))
    spec_target_menu.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _on_spec_target_change(*_):
        global spec_target_hdev
        sel = spec_target_var.get()
        if sel.startswith("(Auto)"):
            spec_target_hdev = None
            print("[SPEC目標] Auto：使用第一隻偵測到的滑鼠")
        else:
            spec_target_hdev = spec_target_map.get(sel)
            print(f"[SPEC目標] 指定 hdev={spec_target_hdev}")
        refresh_spec_target_visual()

    spec_target_var.trace_add("write", _on_spec_target_change)

    # SPEC
    spec_frame = tk.Frame(left)
    spec_frame.pack(fill=tk.X, pady=5)

    tk.Label(
        spec_frame,
        text="SPEC",
        font=("Microsoft JhengHei", 12, "bold"),
        fg="red"
    ).pack(side=tk.LEFT, padx=(0, 10))

    btn_850, w850 = focus_ring_button(spec_frame, ring_color="#FF6D00", ring_thickness=4,
                                      text="1KHz", font=("Microsoft JhengHei", 12),
                                      command=lambda: on_spec_button("850"))
    w850.pack(side=tk.LEFT, padx=5)

    btn_3400, w3400 = focus_ring_button(spec_frame, ring_color="#FF6D00", ring_thickness=4,
                                        text="4KHz", font=("Microsoft JhengHei", 12),
                                        command=lambda: on_spec_button("3400"))
    w3400.pack(side=tk.LEFT, padx=5)

    btn_6800, w6800 = focus_ring_button(spec_frame, ring_color="#FF6D00", ring_thickness=4,
                                        text="8KHz", font=("Microsoft JhengHei", 12),
                                        command=lambda: on_spec_button("6800"))
    w6800.pack(side=tk.LEFT, padx=5)

    btn_none, wnone = focus_ring_button(spec_frame, ring_color="#FF6D00", ring_thickness=4,
                                        text="None", font=("Microsoft JhengHei", 12),
                                        command=lambda: on_spec_button("None"))
    wnone.pack(side=tk.LEFT, padx=5)

    btn_manual, wmanual = focus_ring_button(spec_frame, ring_color="#FF6D00", ring_thickness=4,
                                            text="Manual", font=("Microsoft JhengHei", 12),
                                            command=lambda: on_spec_button("Manual"))
    wmanual.pack(side=tk.LEFT, padx=5)

    spec_selected_label = tk.Label(
        spec_frame,
        text="",
        font=("Microsoft JhengHei", 12),
        fg="blue"
    )
    spec_selected_label.pack(side=tk.LEFT, padx=(5, 0))

    spec_manual_var = tk.StringVar()
    spec_manual_var.trace_add("write", on_manual_entry_change)
    spec_manual_entry = tk.Entry(
        spec_frame,
        textvariable=spec_manual_var,
        width=6,
        font=("Microsoft JhengHei", 12),
        bg="#FFFFCC"
    )

    # === 自動測試：在 SPEC 與 Attenuation 之間 ===
    auto_test_var = tk.BooleanVar(value=True)

    def on_toggle_auto_test():
        _set_step_stop_state(bool(auto_test_var.get()))

    chk_auto, chk_auto_wrap = focus_ring_checkbutton(
        left,
        ring_color="#FF6D00", ring_thickness=3,
        text="自動測試（自動調整衰減）",
        variable=auto_test_var,
        command=on_toggle_auto_test,
        font=("Microsoft JhengHei", 12)
    )
    chk_auto_wrap.pack(fill=tk.X, pady=5)

    # 衰減器參數
    atten_frame = tk.Frame(left)
    atten_frame.pack(fill=tk.X, pady=(0, 5))

    tk.Label(atten_frame, text="Attenuation (dB):", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(0, 5))
    atten_var = tk.StringVar()
    tk.Entry(atten_frame, textvariable=atten_var, width=6, font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT)

    tk.Label(atten_frame, text="Step:", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(15, 5))
    step_var = tk.StringVar()
    step_entry_widget = tk.Entry(atten_frame, textvariable=step_var, width=6, font=("Microsoft JhengHei", 12))
    step_entry_widget.pack(side=tk.LEFT)

    tk.Label(atten_frame, text="Stop:", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(15, 5))
    stop_var = tk.StringVar()
    stop_entry_widget = tk.Entry(atten_frame, textvariable=stop_var, width=6, font=("Microsoft JhengHei", 12))
    stop_entry_widget.pack(side=tk.LEFT)

    # 依初始勾選狀態，先套用
    _set_step_stop_state(bool(auto_test_var.get()))

    def on_set_atten():
        txt = atten_var.get().strip()
        try:
            db_val = float(txt)
            set_attenuation_target(db_val)
            global last_atten_change_time
            last_atten_change_time = time.monotonic()
        except ValueError:
            print("⚠ Attenuation: 請輸入有效浮點數")

    set_btn = tk.Button(
        atten_frame,
        text="Set",
        bg="#444", fg="white",
        font=("Microsoft JhengHei", 12),
        command=on_set_atten
    )
    set_btn.pack(side=tk.LEFT, padx=(17, 0))

    current_frame = tk.Frame(left)
    current_frame.pack(fill=tk.X, pady=(0, 5))
    tk.Label(current_frame, text="Current:", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(0, 5))
    current_label = tk.Label(current_frame, text=f"{current_atten:.1f} dB", font=("Microsoft JhengHei", 12, "bold"))
    current_label.pack(side=tk.LEFT)

    # N9010A IP + 檢查
    ip_frame = tk.Frame(left)
    ip_frame.pack(fill=tk.X, pady=(0, 5))
    tk.Label(ip_frame, text="N9010A IP:", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(0, 5))
    n9010a_ip_var = tk.StringVar(value=N9010A_DEFAULT_IP)
    tk.Entry(ip_frame, textvariable=n9010a_ip_var, width=16, font=("Microsoft JhengHei", 12)).pack(
        side=tk.LEFT, fill=tk.X, expand=True
    )

    def on_click_check_socket() -> None:
        ip = n9010a_ip_var.get().strip()
        if not ip:
            messagebox.showwarning("Socket 檢查", "請先輸入 N9010A IP 位址。")
            return
        if check_socket(ip, N9010A_PORT):
            messagebox.showinfo("Socket 檢查", f"{ip}:{N9010A_PORT} 連線成功。")
        else:
            messagebox.showwarning("Socket 檢查", f"{ip}:{N9010A_PORT} 連線失敗，請確認儀器狀態。")

    tk.Button(ip_frame, text="檢查連線", font=("Microsoft JhengHei", 12), command=on_click_check_socket).pack(
        side=tk.LEFT, padx=(10, 0)
    )

    # SPEC 持續秒數
    hold_frame = tk.Frame(left)
    hold_frame.pack(fill=tk.X, pady=(0, 5))
    tk.Label(hold_frame, text="SPEC持續秒數:", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(0, 5))
    spec_hold_var_local = tk.StringVar(value="10")

    def _validate_hold(*args):
        global spec_hold_duration
        txt = spec_hold_var_local.get().strip()
        if txt == "":
            return
        try:
            sec = float(txt)
            if sec < 1.0:
                raise ValueError
            spec_hold_duration = sec
        except ValueError:
            messagebox.showwarning("秒數錯誤", "請輸入 ≥1 的秒數")
            spec_hold_var_local.set("10")
            spec_hold_duration = 10.0

    spec_hold_var_local.trace_add("write", _validate_hold)
    tk.Entry(hold_frame, textvariable=spec_hold_var_local, width=5, font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT)

    # 勾選：回報率圖
    show_rate_var = tk.BooleanVar(value=False)

    def on_toggle_rate():
        global show_rate_enabled
        show_rate_enabled = show_rate_var.get()

    tk.Checkbutton(left, text="顯示回報率圖", variable=show_rate_var, command=on_toggle_rate,
                   font=("Microsoft JhengHei", 12)).pack(anchor='w', pady=5)

    # 勾選：軌跡圖
    show_traj_var = tk.BooleanVar(value=False)

    def on_toggle_traj():
        global show_traj_enabled
        show_traj_enabled = show_traj_var.get()

    tk.Checkbutton(left, text="顯示滑鼠軌跡圖", variable=show_traj_var, command=on_toggle_traj,
                   font=("Microsoft JhengHei", 12)).pack(anchor='w', pady=5)

    # START / STOP / Save
    tk.Button(left, text="START", bg="green", fg="white",
              font=("Microsoft JhengHei", 12, "bold"), command=start_monitoring).pack(fill=tk.X, pady=5)
    tk.Button(left, text="STOP", bg="red", fg="white",
              font=("Microsoft JhengHei", 12), command=stop_monitoring).pack(fill=tk.X, pady=5)
    tk.Button(left, text="Save Path", bg="blue", fg="white",
              font=("Microsoft JhengHei", 12), command=select_save).pack(fill=tk.X, pady=5)

    start_label = tk.Label(left, text="Start: N/A", font=("Microsoft JhengHei", 12))
    start_label.pack(pady=5)
    stop_label = tk.Label(left, text="Stop: N/A", font=("Microsoft JhengHei", 12))
    stop_label.pack(pady=5)
    save_label = tk.Label(left, text="Save Path: (auto)", wraplength=200, font=("Microsoft JhengHei", 12))
    save_label.pack(pady=5)

    # Matplotlib
    global fig, ax_rate, ax_traj, canvas, spec_line
    fig, (ax_rate, ax_traj) = plt.subplots(2, 1, figsize=(8, 8))
    fig.tight_layout(pad=3)

    ax_rate.set_title("各滑鼠 Report Rate Over Time",
                      fontdict={"family": "Microsoft JhengHei", "size": 14})
    ax_rate.set_xlabel("Time (s)", fontdict={"family": "Microsoft JhengHei", "size": 12})
    ax_rate.set_ylabel("Hz", fontdict={"family": "Microsoft JhengHei", "size": 12})
    ax_rate.set_ylim(0, 1000)
    ax_rate.set_yticks(range(0, 1001, 500))
    spec_line = ax_rate.axhline(y=0, color='red', linestyle='--', visible=False)
    ax_rate.set_visible(False)

    ax_traj.set_title("滑鼠軌跡（多裝置）",
                      fontdict={"family": "Microsoft JhengHei", "size": 14})
    ax_traj.set_xlabel("X", fontdict={"family": "Microsoft JhengHei", "size": 12})
    ax_traj.set_ylabel("Y", fontdict={"family": "Microsoft JhengHei", "size": 12})
    ax_traj.set_visible(False)

    canvas = FigureCanvasTkAgg(fig, root)
    canvas.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)


# ----------------------------------------------------
# === Raw Input Thread ===
def raw_input_loop():
    def wnd_proc(hwnd, msg, wParam, lParam):
        if msg == WM_INPUT:
            try:
                handle_raw_input(lParam)
            except Exception as e:
                print("⚠ WM_INPUT 處理失敗：", e)
        return user32.DefWindowProcW(hwnd, msg, wParam, lParam)

    WNDCLASS, WNDPROCTYPE = make_wndclass()
    hInst = ctypes.windll.kernel32.GetModuleHandleW(None)
    wc = WNDCLASS()
    wc.lpfnWndProc = WNDPROCTYPE(wnd_proc)
    wc.lpszClassName = "RawInputClassMulti"
    user32.RegisterClassW(ctypes.byref(wc))

    hwnd = user32.CreateWindowExW(
        0,
        wc.lpszClassName,
        "",
        0,
        0, 0, 0, 0,
        0, 0,
        hInst,
        None
    )

    devices = (RAWINPUTDEVICE * 1)()
    devices[0].usUsagePage = 0x01
    devices[0].usUsage = 0x02
    devices[0].dwFlags = RIDEV_INPUTSINK
    devices[0].hwndTarget = hwnd
    user32.RegisterRawInputDevices(devices, 1, ctypes.sizeof(RAWINPUTDEVICE))

    msg = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), hwnd, 0, 0) != 0:
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))


def handle_raw_input(lParam):
    size = wintypes.UINT()
    user32.GetRawInputData(lParam, RID_INPUT, None, ctypes.byref(size), ctypes.sizeof(RAWINPUTHEADER))
    buf = ctypes.create_string_buffer(size.value)
    user32.GetRawInputData(lParam, RID_INPUT, buf, ctypes.byref(size), ctypes.sizeof(RAWINPUTHEADER))
    raw = ctypes.cast(buf, ctypes.POINTER(RAWINPUT)).contents

    if raw.header.dwType != 0:
        return

    hdev = getattr(raw.header.hDevice, 'value', raw.header.hDevice)

    if hdev in ignored_devices:
        return
    if is_bluetooth_device(hdev) or is_touchpad_device(hdev):
        ignored_devices.add(hdev)
        return

    vid, pid = get_device_pid_vid(hdev)
    if not (vid and pid):
        return

    now = time.monotonic()
    if hdev not in device_event_times:
        initialize_new_device(hdev)

    event_counter[hdev] = event_counter.get(hdev, 0) + 1
    device_event_times[hdev].append(now)
    device_positions[hdev].append(get_cursor_pos())

    global last_event_time
    last_event_time = now


def initialize_new_device(hdev):
    global next_color_index
    n = max(0, len(plot_times))
    device_plot_rates[hdev] = [math.nan] * n

    device_event_times[hdev] = collections.deque()
    device_positions[hdev] = []
    event_counter[hdev] = 0
    ms_report_data[hdev] = []

    vid, pid = get_device_pid_vid(hdev)
    device_vid_pid[hdev] = (vid, pid)

    if vid and pid:
        name = lookup_brand_model(vid, pid)
        dev_label = f"{name} ({vid}:{pid})" if name else f"{vid}:{pid}"
    else:
        dev_label = "Unknown"

    color = color_cycle[next_color_index % len(color_cycle)]
    device_colors[hdev] = color
    next_color_index += 1

    def _init_ui_for_device():
        if show_rate_locked:
            line, = ax_rate.plot([], [], '-', color=color, label=dev_label)
            device_lines_rate[hdev] = line
            ax_rate.legend(loc='upper right')

        lbl = tk.Label(rate_frame, text=f"{dev_label} → 0 Hz",
                       font=("Microsoft JhengHei", 12), fg=color)
        lbl.pack(anchor='w')
        rate_labels[hdev] = lbl

        add_device_to_spec_menu(hdev, dev_label)
        refresh_spec_target_visual()

        if show_rate_locked:
            canvas.draw_idle()

    root.after(0, _init_ui_for_device)


def update_ms_report_rate(hdev, now_mono: float, now_dt: datetime) -> int:
    dq = device_event_times.get(hdev)
    if dq is None:
        return 0

    while dq and (now_mono - dq[0]) >= 1.0:
        dq.popleft()

    report_rate = len(dq)

    if hdev in rate_labels:
        vid, pid = device_vid_pid.get(hdev, (None, None))
        if vid and pid:
            name = lookup_brand_model(vid, pid)
            left_text = f"{name} ({vid}:{pid})" if name else f"{vid}:{pid}"
        else:
            left_text = "Unknown"

        prefix = SPEC_TAG if (rate_labels[hdev].cget("text").startswith(SPEC_TAG)) else ""
        rate_labels[hdev].config(text=f"{prefix}{left_text} → {report_rate} Hz")

    data_list = ms_report_data.setdefault(hdev, [])
    data_list.append((now_dt, report_rate, report_rate))
    cutoff = now_dt - timedelta(milliseconds=1100)
    while data_list and data_list[0][0] < cutoff:
        data_list.pop(0)

    return report_rate


def highlight_xticks_for_atten_changes(ax, marks, *, tol=0.25):
    if ax is None:
        return
    if not marks:
        for lbl in ax.get_xticklabels():
            lbl.set_color("black")
            lbl.set_fontweight("normal")
        return

    x0, x1 = ax.get_xlim()
    marks_in_view = [m for m in marks if (x0 - tol) <= m <= (x1 + tol)]
    for lbl in ax.get_xticklabels():
        txt = lbl.get_text().strip()
        try:
            xv = float(txt)
        except Exception:
            lbl.set_color("black")
            lbl.set_fontweight("normal")
            continue

        hit = any(abs(xv - m) <= tol for m in marks_in_view)
        if hit:
            lbl.set_color("red")
            lbl.set_fontweight("bold")
        else:
            lbl.set_color("black")
            lbl.set_fontweight("normal")


def update_plot(frame):
    global y_max_fixed, spec_line, active_spec_value
    global spec_hold_start
    global prev_all_above
    global half_phase, half_step_time
    global last_atten_change_time
    global backtest_fail_count, BACKTEST_MAX
    global auto_test_locked
    global target_atten

    # ✅ 新增：狀態機 globals
    global auto_mode, full_retry_count, backtest_count, pass_hold_start
    global initial_atten_at_start, full_step_direction

    if chart_start_time is None or not monitoring:
        return list(device_lines_rate.values())

    now = time.monotonic()
    now_ts = datetime.now()

    # ✅ 主邏輯衰減基準（避免 worker 更新延遲造成判斷飄）
    base_atten = target_atten if target_atten is not None else current_atten

    # 超過 TIMEOUT → 停止
    if now - last_event_time >= TIMEOUT:
        stop_monitoring()
        return list(device_lines_rate.values())

    elapsed_time = now - chart_start_time
    plot_times.append(elapsed_time)
    spec_value = active_spec_value

    current_max = 0
    for hdev in list(device_event_times.keys()):
        rate = update_ms_report_rate(hdev, now, now_ts)
        device_plot_rates.setdefault(hdev, [])
        device_plot_rates[hdev].append(rate)
        current_max = max(current_max, rate)

    if show_rate_locked:
        for hdev in list(device_event_times.keys()):
            if hdev not in device_lines_rate:
                color = device_colors.get(hdev, None) or color_cycle[len(device_lines_rate) % len(color_cycle)]
                vid, pid = device_vid_pid.get(hdev, (None, None))
                name = lookup_brand_model(vid, pid)
                dev_label = f"{name} ({vid}:{pid})" if name else (f"{vid}:{pid}" if vid and pid else "Unknown")
                line, = ax_rate.plot([], [], '-', color=color, label=dev_label)
                device_lines_rate[hdev] = line
                ax_rate.legend(loc='upper right')

        if y_max_fixed is None and elapsed_time >= 2.0:
            if current_max <= 1000:
                y_max_fixed = 1000
            elif current_max <= 4000:
                y_max_fixed = 4000
            else:
                y_max_fixed = 8000

        if y_max_fixed is not None:
            ax_rate.set_ylim(0, y_max_fixed)
        else:
            ax_rate.set_ylim(0, 1000)

        if spec_value not in (None, 0):
            spec_line.set_ydata([spec_value, spec_value])
            spec_line.set_visible(True)
        else:
            spec_line.set_visible(False)

        window_sec = float(spec_hold_duration) + PLOT_WINDOW_EXTRA_SEC
        if window_sec < 3.0:
            window_sec = 3.0

        if plot_times:
            if elapsed_time <= window_sec:
                x_left = 0.0
                x_right = window_sec
                i0 = 0
            else:
                x_left = elapsed_time - window_sec
                x_right = elapsed_time
                i0 = bisect.bisect_left(plot_times, x_left)

            x_view = plot_times[i0:]

            for hdev, line in device_lines_rate.items():
                y_all = device_plot_rates.get(hdev, [])
                y_view = y_all[i0:] if i0 < len(y_all) else []

                m = min(len(x_view), len(y_view))
                if m <= 0:
                    line.set_data([], [])
                else:
                    line.set_data(x_view[-m:], y_view[-m:])

            from matplotlib.ticker import MultipleLocator
            ax_rate.set_xlim(x_left, x_right)
            ax_rate.margins(x=0)
            ax_rate.xaxis.set_major_locator(MultipleLocator(1.0))

            highlight_xticks_for_atten_changes(ax_rate, atten_change_marks, tol=0.25)

        refresh_spec_target_visual()

    canvas.draw_idle()

    # SPEC skip（含：起始2秒、切衰減後2秒）
    skip_initial = (spec_value is not None and elapsed_time < 2.0)
    skip_after_change = (last_atten_change_time is not None and now - last_atten_change_time < 2.0)
    skip_spec_logic = (spec_value == 0) or skip_initial or skip_after_change

    # ✅ 如果在 skip 期：不累積 PASS
    if skip_spec_logic:
        pass_hold_start = None

    if not skip_spec_logic and spec_value is not None:
        target_hdev = get_effective_spec_target()
        if target_hdev is None or target_hdev not in device_event_times:
            return list(device_lines_rate.values())

        latest_rate = 0
        if ms_report_data.get(target_hdev):
            latest_rate = ms_report_data[target_hdev][-1][2]

        all_above = (latest_rate >= spec_value)

        # 固定衰減模式（Auto Test OFF）—保留原邏輯
        if not auto_test_locked:
            if not all_above:
                trigger_async_capture(base_atten)
                trigger_ui_capture_on_first_fail(base_atten)
                print(f"❌ [固定衰減模式][SPEC目標={target_hdev}] FAIL @ {base_atten:.2f} dB，停止測試")
                stop_monitoring()
                prev_all_above = all_above
                return list(device_lines_rate.values())
            else:
                if spec_hold_start is None:
                    spec_hold_start = now
                elif now - spec_hold_start >= spec_hold_duration:
                    print(f"✅ [固定衰減模式][SPEC目標={target_hdev}] PASS @ {base_atten:.2f} dB，停止測試")
                    stop_monitoring()
                    prev_all_above = all_above
                    return list(device_lines_rate.values())

            prev_all_above = all_above
            return list(device_lines_rate.values())

        # ------------------------------------------------
        # ✅✅✅ Auto test ON：一 FAIL 立刻動作 + 重測(3) + 回測(10/到初始)
        # ------------------------------------------------
        try:
            step_val = float(step_var.get().strip())
        except Exception:
            step_val = 0.0

        try:
            stop_val = float(stop_var.get().strip())
        except Exception:
            stop_val = base_atten

        half_step = step_val / 2.0 if step_val else 0.0

        if initial_atten_at_start is None:
            initial_atten_at_start = base_atten

        def _calc_next_backtest_atten(curr: float) -> float:
            move = -full_step_direction * half_step  # 往「初始衰減」方向
            if move == 0:
                return curr

            nxt = curr + move
            init_a = float(initial_atten_at_start)

            # clamp：不要越過初始衰減
            if full_step_direction == 1:
                if nxt < init_a:
                    nxt = init_a
            else:
                if nxt > init_a:
                    nxt = init_a
            return nxt

        # 1) FAIL：立刻動作
        if not all_above:
            pass_hold_start = None

            trigger_async_capture(base_atten)
            trigger_ui_capture_on_first_fail(base_atten)

            if auto_mode == AUTO_MODE_FULL:
                full_retry_count += 1
                print(f"❌ [FULL FAIL→同衰減重測][{full_retry_count}/{FULL_RETRY_MAX}] [SPEC目標={target_hdev}] @ {base_atten:.2f} dB")

                # ✅ 同衰減重測也要「前兩秒不判」
                last_atten_change_time = time.monotonic()

                if full_retry_count >= FULL_RETRY_MAX:
                    auto_mode = AUTO_MODE_BACKTEST
                    backtest_count = 0
                    print(f"➡️  重測已達 {FULL_RETRY_MAX} 次仍 FAIL → 進入 HALF step 回測（最多 {BACKTEST_MAX} 次 / 到初始衰減即停）")

                    nxt = _calc_next_backtest_atten(base_atten)
                    if abs(nxt - base_atten) <= 1e-9:
                        print("❌ [回測] half_step=0 或已到邊界無法回測 → 停止測試")
                        stop_monitoring()
                        prev_all_above = all_above
                        return list(device_lines_rate.values())

                    atten_change_marks.append(elapsed_time)
                    set_attenuation_target(nxt)
                    last_atten_change_time = time.monotonic()
                    backtest_count = 1
                    print(f"[回測 STEP#{backtest_count}][SPEC目標={target_hdev}] → {nxt:.2f} dB")

            else:
                if backtest_count >= BACKTEST_MAX:
                    print(f"❌ [回測 FAIL] 已達最大回測次數 {BACKTEST_MAX} → 停止測試")
                    stop_monitoring()
                    prev_all_above = all_above
                    return list(device_lines_rate.values())

                init_a = float(initial_atten_at_start)
                if abs(base_atten - init_a) <= 1e-6:
                    print(f"❌ [回測 FAIL] 已到達初始衰減 {init_a:.2f} dB → 停止測試")
                    stop_monitoring()
                    prev_all_above = all_above
                    return list(device_lines_rate.values())

                nxt = _calc_next_backtest_atten(base_atten)
                if abs(nxt - base_atten) <= 1e-9:
                    print("❌ [回測] 已到邊界/half_step=0 無法回測 → 停止測試")
                    stop_monitoring()
                    prev_all_above = all_above
                    return list(device_lines_rate.values())

                backtest_count += 1
                atten_change_marks.append(elapsed_time)
                set_attenuation_target(nxt)
                last_atten_change_time = time.monotonic()
                print(f"❌ [回測 FAIL→繼續回測 STEP#{backtest_count}][SPEC目標={target_hdev}] → {nxt:.2f} dB")

            prev_all_above = all_above
            return list(device_lines_rate.values())

        # 2) PASS：連續滿 spec_hold_duration 才算 PASS
        if pass_hold_start is None:
            pass_hold_start = now

        if (now - pass_hold_start) >= spec_hold_duration:
            if auto_mode == AUTO_MODE_BACKTEST:
                print(f"✅ [回測 PASS][SPEC目標={target_hdev}] @ {base_atten:.2f} dB → 停止測試")
                stop_monitoring()
                prev_all_above = all_above
                return list(device_lines_rate.values())

            # FULL 模式 PASS：繼續 full step
            full_retry_count = 0
            pass_hold_start = None

            if stop_val > base_atten:
                new_atten = min(stop_val, base_atten + step_val)
            else:
                new_atten = max(stop_val, base_atten - step_val)

            if new_atten != base_atten:
                atten_change_marks.append(elapsed_time)
                set_attenuation_target(new_atten)
                last_atten_change_time = time.monotonic()
                print(f"✅ [FULL PASS→全步STEP][SPEC目標={target_hdev}] → {new_atten:.2f} dB")
            else:
                print(f"✅ [到達STOP且PASS][SPEC目標={target_hdev}] @ {base_atten:.2f} dB (Stop={stop_val:.2f}) → 停止測試")
                stop_monitoring()
                prev_all_above = all_above
                return list(device_lines_rate.values())

        prev_all_above = all_above

    return list(device_lines_rate.values())


def start_monitoring():
    global monitoring, chart_start_time, ani, next_color_index, last_event_time, y_max_fixed, spec_line
    global start_time_str, stop_time_str
    global spec_hold_start
    global half_phase, half_step_time
    global last_atten_change_time
    global show_rate_locked
    global prev_all_above
    global backtest_fail_count
    global auto_test_locked
    global ui_fail_captured
    global target_atten

    # ✅ 新增：狀態機 globals
    global auto_mode, full_retry_count, backtest_count, pass_hold_start
    global initial_atten_at_start, full_step_direction

    ui_fail_captured = False

    half_phase = False
    half_step_time = None
    last_atten_change_time = None
    backtest_fail_count = 0

    # ✅ reset 狀態機
    auto_mode = AUTO_MODE_FULL
    full_retry_count = 0
    backtest_count = 0
    pass_hold_start = None
    initial_atten_at_start = None
    full_step_direction = 1

    show_rate_locked = bool(show_rate_var.get())
    auto_test_locked = bool(auto_test_var.get()) if auto_test_var is not None else True
    print(f"[Auto Test] {'ON' if auto_test_locked else 'OFF（固定衰減）'}")

    prev_all_above = True

    if ani is not None:
        ani.event_source.stop()
        ani = None

    device_event_times.clear()
    device_plot_rates.clear()
    device_positions.clear()
    plot_times.clear()
    atten_change_marks.clear()

    for lbl in rate_labels.values():
        lbl.destroy()
    rate_labels.clear()

    device_lines_rate.clear()
    device_colors.clear()
    next_color_index = 0
    event_counter.clear()
    ms_report_data.clear()
    device_vid_pid.clear()
    y_max_fixed = None

    now = time.monotonic()
    chart_start_time = now
    last_event_time = now
    monitoring = True
    spec_hold_start = None

    # ✅ 取得初始衰減 & full step 方向
    try:
        init_atten = float(atten_var.get().strip())
        initial_atten_at_start = init_atten
        set_attenuation_target(init_atten)

        try:
            stop_val = float(stop_var.get().strip())
        except Exception:
            stop_val = init_atten

        full_step_direction = 1 if stop_val > init_atten else -1

        # 起始也算「剛切換」：前2秒不判
        last_atten_change_time = time.monotonic()

    except Exception:
        pass

    start_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    stop_time_str = None
    start_label.config(text=time.strftime("Start: %Y-%m-%d %H:%M:%S", time.localtime()))
    stop_label.config(text="Stop: N/A")

    ax_rate.clear()
    ax_rate.set_title("各滑鼠 Report Rate Over Time",
                      fontdict={"family": "Microsoft JhengHei", "size": 14})
    ax_rate.set_xlabel("Time (s)", fontdict={"family": "Microsoft JhengHei", "size": 12})
    ax_rate.set_ylabel("Hz", fontdict={"family": "Microsoft JhengHei", "size": 12})
    spec_line = ax_rate.axhline(y=0, color='red', linestyle='--', visible=False)

    if show_rate_locked:
        ax_rate.set_visible(True)
        if active_spec_value in (None, 0):
            _apply_axis_for_spec(None)
        else:
            _apply_axis_for_spec(active_spec_value, from_manual=active_spec_is_manual)
    else:
        ax_rate.set_visible(False)

    canvas.draw_idle()
    refresh_spec_target_visual()

    ani = animation.FuncAnimation(fig, update_plot, interval=10, cache_frame_data=False)


def stop_monitoring():
    global monitoring, ani, stop_time_str
    if not monitoring:
        return
    monitoring = False

    stop_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    stop_label.config(text=time.strftime("Stop: %Y-%m-%d %H:%M:%S", time.localtime()))

    if ani is not None:
        ani.event_source.stop()
        ani = None

    save_results()


def on_app_close():
    global monitoring, ani, root, fig, _atten_worker_stop
    monitoring = False
    _atten_worker_stop = True

    if ani is not None:
        try:
            ani.event_source.stop()
        except Exception as exc:
            print(f"[關閉流程] 停止動畫時發生例外: {exc}")
        ani = None

    try:
        plt.close(fig)
    except Exception:
        pass
    fig = None

    root.quit()
    root.destroy()


def select_save():
    global save_path
    save_path = filedialog.asksaveasfilename(
        defaultextension='.csv',
        filetypes=[('CSV 檔', '*.csv')]
    )
    if save_path:
        save_label.config(text=f"Save Path: {save_path}")


def save_results():
    global save_path, start_time_str, stop_time_str

    if save_path:
        fn = save_path
    else:
        s = start_time_str if start_time_str else 'start'
        e = stop_time_str if stop_time_str else 'stop'
        fn = f"mouse_log_multi_{s}_{e}.csv"

    all_hdevs = list(device_plot_rates.keys())
    max_len = max((len(device_plot_rates[hdev]) for hdev in all_hdevs), default=0)

    header = []
    for _ in all_hdevs:
        header.extend(["Device", "VID", "PID", "Time(s)", "Hz", "X", "Y"])

    with open(fn, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)

        for row_idx in range(max_len):
            row = []
            for hdev in all_hdevs:
                vid, pid = device_vid_pid.get(hdev, (None, None))
                rates = device_plot_rates[hdev]
                poses = device_positions.get(hdev, [])
                t = plot_times[row_idx] if row_idx < len(plot_times) else ''
                r = rates[row_idx] if row_idx < len(rates) else ''
                if row_idx < len(poses):
                    x, y = poses[row_idx]
                elif poses:
                    x, y = poses[-1]
                else:
                    x, y = ('', '')
                row.extend([hdev, vid, pid, t, r, x, y])
            writer.writerow(row)

    save_label.config(text=f"✅ 已儲存到 {fn}")
    print(f"✅ 已儲存到 {fn}")


# ----------------------------------------------------
# === 主程式入口 ===
if __name__ == '__main__':
    setup_ui()
    ensure_atten_worker()
    threading.Thread(target=raw_input_loop, daemon=True).start()
    root.mainloop()
