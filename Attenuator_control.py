import os
import sys
import threading
import time
import collections
import csv
import tkinter as tk
from tkinter import filedialog
from tkinter import font as tkfont
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime
import re

import ctypes
import ctypes.wintypes as wintypes
from ctypes import cdll, c_int
from time import sleep

# ----------------------------------------------------
# === VNX 衰減器 (Attenuator) CLI 初始化 ===

os.add_dll_directory(os.getcwd())
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

def set_attenuation_db(db_value: float):
    """
    將衰減值 (dB) 設定到 VNX 裝置，並更新 global current_atten。
    """
    global current_atten
    try:
        db_int = int(db_value * 20)
        print(f"👉 設定衰減值為 {db_value} dB (乘以20後 = {db_int})")
        vnx.fnLDA_SetAttenuationHR(devid, db_int)
        sleep(0.5)
        current_atten = db_value
    except Exception as e:
        print("⚠ 設定衰減失敗：", e)

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
RID_INPUT       = 0x10000003
WM_INPUT        = 0x00FF

def get_cursor_pos():
    pt = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    return (pt.x, pt.y)

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
# === 全域狀態與資料結構 ===

monitoring = False
chart_start_time = None
ani = None
last_event_time = time.monotonic()
y_max_fixed = None
spec_line = None

active_spec_value = None    # 儲存 850/3400/6800 或 手動數字
spec_streak_start = None    # 連續滿足 SPEC 開始時間
at_stop_phase = False       # 是否已達 Stop 後進入觀察階段
at_stop_start = None        # Stop 階段開始時間

current_atten = None        # 當前衰減值

device_event_times = {}
device_smoothed_rate = {}
device_plot_rates = {}
plot_times = []
device_positions = {}
device_lines_rate = {}
device_colors = {}
device_vid_pid = {}
color_cycle = plt.rcParams['axes.prop_cycle'].by_key()['color']
next_color_index = 0

event_counter = {}
pre_event_counter = {}
ms_report_data = {}

TIMEOUT = 5.0  # 5 秒沒輸入就停止

root = None
rate_frame = None
rate_labels = {}
start_label = None
stop_label = None
save_label = None

spec_var = None
spec_manual_entry = None
spec_manual_var = None
spec_display_label = None
atten_var = None
step_var = None
stop_var = None
save_path = None

fig = None
ax_rate = None
ax_traj = None
canvas = None

start_time_str = None
stop_time_str = None

# ----------------------------------------------------
# === SPEC 下拉選單 callback ===

def on_spec_selected(val):
    """
    val = "SPEC" / "Manual" / "850" / "3400" / "6800" / "None"
    """
    global active_spec_value, spec_streak_start, at_stop_phase, at_stop_start

    # 重置所有連續計時
    spec_streak_start = None
    at_stop_phase = False
    at_stop_start = None

    if val == "SPEC" or val == "None":
        active_spec_value = None
        if spec_manual_entry.winfo_ismapped():
            spec_manual_entry.pack_forget()
        if spec_display_label.winfo_ismapped():
            spec_display_label.pack_forget()

    elif val == "Manual":
        active_spec_value = None
        if spec_display_label.winfo_ismapped():
            spec_display_label.pack_forget()
        spec_manual_var.set("")  # 清空
        if not spec_manual_entry.winfo_ismapped():
            spec_manual_entry.pack(side=tk.LEFT, padx=10)
        spec_var.set("Manual")

    else:
        try:
            sv = int(val)
            active_spec_value = sv
        except ValueError:
            active_spec_value = None

        if spec_manual_entry.winfo_ismapped():
            spec_manual_entry.pack_forget()
        spec_display_label.config(text=f"{val} Hz")
        if not spec_display_label.winfo_ismapped():
            spec_display_label.pack(side=tk.LEFT, padx=10)

        spec_var.set("SPEC")

def on_manual_entry_change(*args):
    global active_spec_value, spec_streak_start, at_stop_phase, at_stop_start
    txt = spec_manual_var.get().strip()
    if txt.isdigit():
        active_spec_value = int(txt)
    else:
        active_spec_value = None
    spec_streak_start = None
    at_stop_phase = False
    at_stop_start = None

# ----------------------------------------------------
# === Tkinter + Matplotlib GUI 初始化 ===

def setup_ui():
    global root, rate_frame, start_label, stop_label, save_label
    global spec_var, spec_manual_entry, spec_manual_var, spec_display_label, spec_line
    global fig, ax_rate, ax_traj, canvas
    global atten_var, step_var, stop_var

    root = tk.Tk()
    root.title("多滑鼠 Report Rate & 軌跡 (C# 算法移植版)")
    root.geometry("1000x700")

    # 1) 先把預設 Tkinter font 改成支援中文的 Microsoft JhengHei
    default_font = tkfont.nametofont("TkDefaultFont")
    default_font.configure(family="Microsoft JhengHei", size=12)

    left = tk.Frame(root)
    left.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

    # ─── 左側最上方：Report Rate (Hz) 字樣 ───
    tk.Label(left, text="Report Rate (Hz)", font=("Microsoft JhengHei", 14, "bold")).pack(pady=5)
    rate_frame = tk.Frame(left)
    rate_frame.pack(pady=5)

    # ─── SPEC 下拉選單 區塊 ───
    spec_var = tk.StringVar(value="SPEC")
    spec_manual_var = tk.StringVar()
    spec_manual_var.trace_add("write", on_manual_entry_change)

    spec_frame = tk.Frame(left)
    spec_frame.pack(fill=tk.X, pady=5)

    # (1) OptionMenu：字型也強制指定為 Microsoft JhengHei
    om = tk.OptionMenu(
        spec_frame,
        spec_var,
        "SPEC",
        "Manual",
        "850",
        "3400",
        "6800",
        "None",
        command=on_spec_selected
    )
    om.config(font=("Microsoft JhengHei", 12))
    om["menu"].config(font=("Microsoft JhengHei", 12))
    om.pack(side=tk.LEFT, anchor='w')

    # (2) 手動輸入 Entry (不一開始 pack，要選 Manual 才顯示)
    spec_manual_entry = tk.Entry(spec_frame, textvariable=spec_manual_var, width=6, font=("Microsoft JhengHei", 12))

    # (3) 選好 SPEC 後顯示的 Label (不一開始 pack)
    spec_display_label = tk.Label(spec_frame, text="", font=("Microsoft JhengHei", 12))

    # ─── 在 SPEC 與 START 之間插入 Attenuation 區塊 ───
    atten_frame = tk.Frame(left)
    atten_frame.pack(fill=tk.X, pady=(0, 5))

    tk.Label(atten_frame, text="Attenuation (dB):", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(0, 5))
    atten_var = tk.StringVar()
    atten_entry = tk.Entry(atten_frame, textvariable=atten_var, width=6, font=("Microsoft JhengHei", 12))
    atten_entry.pack(side=tk.LEFT)

    tk.Label(atten_frame, text="Step:", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(10, 5))
    step_var = tk.StringVar()
    step_entry = tk.Entry(atten_frame, textvariable=step_var, width=6, font=("Microsoft JhengHei", 12))
    step_entry.pack(side=tk.LEFT)

    tk.Label(atten_frame, text="Stop:", font=("Microsoft JhengHei", 12)).pack(side=tk.LEFT, padx=(10, 5))
    stop_var = tk.StringVar()
    stop_entry = tk.Entry(atten_frame, textvariable=stop_var, width=6, font=("Microsoft JhengHei", 12))
    stop_entry.pack(side=tk.LEFT)

    def on_set_atten():
        txt = atten_var.get().strip()
        try:
            db_val = float(txt)
            set_attenuation_db(db_val)
        except ValueError:
            print("⚠ Attenuation: 請輸入有效浮點數")

    tk.Button(
        atten_frame,
        text="Set",
        bg="#444", fg="white",
        font=("Microsoft JhengHei", 12),
        command=on_set_atten
    ).pack(side=tk.LEFT, padx=(10, 0))

    # ─── START / STOP / Save Path 按鈕 ───
    tk.Button(left, text="START", bg="green", fg="white", font=("Microsoft JhengHei", 12),
              command=start_monitoring).pack(fill=tk.X, pady=5)
    tk.Button(left, text="STOP", bg="red", fg="white", font=("Microsoft JhengHei", 12),
              command=stop_monitoring).pack(fill=tk.X, pady=5)
    tk.Button(left, text="Save Path", bg="blue", fg="white", font=("Microsoft JhengHei", 12),
              command=select_save).pack(fill=tk.X, pady=5)

    # ─── 開始/結束/儲存路徑 Label ───
    start_label = tk.Label(left, text="Start: N/A", font=("Microsoft JhengHei", 12))
    start_label.pack(pady=5)
    stop_label = tk.Label(left, text="Stop: N/A", font=("Microsoft JhengHei", 12))
    stop_label.pack(pady=5)
    save_label = tk.Label(left, text="Save Path: (auto)", wraplength=200, font=("Microsoft JhengHei", 12))
    save_label.pack(pady=5)

    # ─── 右側：Matplotlib Figure ───
    fig, (ax_rate, ax_traj) = plt.subplots(2, 1, figsize=(8, 8))
    fig.tight_layout(pad=3)

    # 回報率子圖設定
    ax_rate.set_title("各滑鼠 Report Rate Over Time", fontdict={"family":"Microsoft JhengHei", "size":14})
    ax_rate.set_xlabel("Time (s)", fontdict={"family":"Microsoft JhengHei", "size":12})
    ax_rate.set_ylabel("Hz", fontdict={"family":"Microsoft JhengHei", "size":12})
    ax_rate.set_ylim(0, 1000)
    ax_rate.set_yticks(range(0, 1001, 500))

    # SPEC 虛線 (預設隱藏)
    spec_line = ax_rate.axhline(y=0, color='red', linestyle='--', visible=False)

    # 軌跡子圖設定
    ax_traj.set_title("滑鼠軌跡（多裝置）", fontdict={"family":"Microsoft JhengHei", "size":14})
    ax_traj.set_xlabel("X", fontdict={"family":"Microsoft JhengHei", "size":12})
    ax_traj.set_ylabel("Y", fontdict={"family":"Microsoft JhengHei", "size":12})

    canvas = FigureCanvasTkAgg(fig, root)
    canvas.get_tk_widget().pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# ----------------------------------------------------
# === Raw Input Thread ===

def raw_input_loop():
    def wnd_proc(hwnd, msg, wParam, lParam):
        if msg == WM_INPUT:
            handle_raw_input(lParam)
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
        0,
        0,
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

    if raw.header.dwType == 0:  # RIM_TYPEMOUSE
        hdev = int(raw.header.hDevice)
        now = time.monotonic()
        if hdev not in device_event_times:
            initialize_new_device(hdev)

        event_counter[hdev] += 1
        device_event_times[hdev].append(now)
        pos = get_cursor_pos()
        device_positions[hdev].append(pos)

        global last_event_time
        last_event_time = now

def initialize_new_device(hdev):
    global next_color_index

    device_event_times[hdev] = collections.deque()
    device_smoothed_rate[hdev] = 0.0
    device_plot_rates[hdev] = [0.0] * len(plot_times)
    device_positions[hdev] = []

    event_counter[hdev] = 0
    pre_event_counter[hdev] = 0
    ms_report_data[hdev] = []

    vid, pid = get_device_pid_vid(hdev)
    device_vid_pid[hdev] = (vid, pid)
    if vid and pid:
        dev_label = f"{hdev}:{vid}:{pid}"
    else:
        dev_label = f"{hdev}:Unknown"

    color = color_cycle[next_color_index % len(color_cycle)]
    device_colors[hdev] = color
    next_color_index += 1

    line, = ax_rate.plot([], [], '-', color=color, label=dev_label)
    device_lines_rate[hdev] = line

    lbl = tk.Label(
        rate_frame,
        text=f"{dev_label} → 0 Hz",
        font=("Microsoft JhengHei", 12),
        fg=color
    )
    lbl.pack(anchor='w')
    rate_labels[hdev] = lbl

    ax_rate.legend(loc='upper right')

def update_ms_report_rate(hdev):
    now_dt = datetime.now()
    current_counter = event_counter[hdev]
    prev_counter = pre_event_counter[hdev]
    delta_count = current_counter - prev_counter

    data_list = ms_report_data[hdev]
    index = -1
    for i, (ts, cnt, r) in enumerate(data_list):
        elapsed_ms = (now_dt - ts).total_seconds() * 1000
        if elapsed_ms < 1000:
            index = i
            break

    sum_after = 0
    if index != -1:
        for (_, cnt, _) in data_list[index + 1:]:
            sum_after += cnt

    interp = 0
    if index > 0:
        ts_i, cnt_i, _ = data_list[index]
        ts_prev, cnt_prev, _ = data_list[index - 1]
        elapsed_i_ms = (now_dt - ts_i).total_seconds() * 1000
        delta_prev_ms = (ts_i - ts_prev).total_seconds() * 1000
        if delta_prev_ms > 0:
            interp = int(cnt_i * (1000 - elapsed_i_ms) / delta_prev_ms)

    report_rate = sum_after + delta_count + interp

    if hdev in rate_labels:
        vid, pid = device_vid_pid.get(hdev, (None, None))
        if vid and pid:
            rate_labels[hdev].config(text=f"{hdev}:{vid}:{pid} → {report_rate} Hz")
        else:
            rate_labels[hdev].config(text=f"{hdev}:Unknown → {report_rate} Hz")

    data_list.append((now_dt, delta_count, report_rate))
    pre_event_counter[hdev] = current_counter

    return report_rate

def update_plot(frame):
    global last_event_time, y_max_fixed, spec_line, active_spec_value
    global spec_streak_start, at_stop_phase, at_stop_start, current_atten

    # 如果還沒按 START 或者已經停止
    if chart_start_time is None or not monitoring:
        return list(device_lines_rate.values())

    now = time.monotonic()
    # 如果距離最後一次事件超過 TIMEOUT，就顯示最後一次 rate，然後停
    if now - last_event_time >= TIMEOUT:
        for hdev in event_counter.keys():
            if ms_report_data[hdev]:
                last_rate = ms_report_data[hdev][-1][2]
            else:
                last_rate = 0
            if hdev in rate_labels:
                vid, pid = device_vid_pid.get(hdev, (None, None))
                if vid and pid:
                    rate_labels[hdev].config(text=f"{hdev}:{vid}:{pid} → {last_rate} Hz")
                else:
                    rate_labels[hdev].config(text=f"{hdev}:Unknown → {last_rate} Hz")
        stop_monitoring()
        return list(device_lines_rate.values())

    elapsed_time = now - chart_start_time
    plot_times.append(elapsed_time)

    # 讀取目前的 SPEC 值
    spec_value = active_spec_value
    if spec_value is not None:
        spec_line.set_ydata([spec_value, spec_value])
        spec_line.set_visible(True)
    else:
        spec_line.set_visible(False)

    # 計算每隻滑鼠的報告率
    current_max = 0
    for hdev in device_event_times.keys():
        report_rate = update_ms_report_rate(hdev)
        device_plot_rates[hdev].append(report_rate)
        if report_rate > current_max:
            current_max = report_rate

    # 第一次過 2 秒後決定 Y 軸上限
    if elapsed_time >= 2 and y_max_fixed is None:
        if current_max <= 1000:
            y_max_fixed = 1000
        elif current_max <= 4000:
            y_max_fixed = 4000
        else:
            y_max_fixed = 8000

    # 設定 Y 軸範圍
    if y_max_fixed is not None:
        ax_rate.set_ylim(0, y_max_fixed)
        ax_rate.set_yticks(range(0, y_max_fixed + 1, 500))
    else:
        ax_rate.set_ylim(0, 1000)
        ax_rate.set_yticks(range(0, 1001, 500))

    # 設定 X 軸範圍 (最近 20 秒)
    window_start = max(0, elapsed_time - 20)
    window_end = window_start + 20
    ax_rate.set_xlim(window_start, window_end)
    start_int = int(window_start)
    end_int = int(window_end) + 1
    ax_rate.set_xticks(range(start_int, end_int, 1))

    # 更新每條折線的資料
    for hdev, line in device_lines_rate.items():
        xs = plot_times
        ys = device_plot_rates[hdev]
        line.set_data(xs, ys)

    # 重畫軌跡子圖
    ax_traj.clear()
    ax_traj.set_title("滑鼠軌跡（多裝置）", fontdict={"family":"Microsoft JhengHei", "size":14})
    ax_traj.set_xlabel("X", fontdict={"family":"Microsoft JhengHei", "size":12})
    ax_traj.set_ylabel("Y", fontdict={"family":"Microsoft JhengHei", "size":12})
    for hdev, positions in device_positions.items():
        if positions:
            xs, ys = zip(*positions)
            ax_traj.plot(xs, ys, '-', color=device_colors[hdev],
                         label=f"{hdev}:{device_vid_pid.get(hdev, ('??','??'))[0]}:{device_vid_pid.get(hdev, ('??','??'))[1]}")
    if any(len(positions) > 0 for positions in device_positions.values()):
        ax_traj.legend(loc='upper right')

    canvas.draw_idle()

    # ───────────────────────────────────────────────────
    # 只要有「任何一隻」滑鼠的最新 report_rate < SPEC，就停止
    if spec_value is not None and elapsed_time >= 2:
        for hdev in device_event_times.keys():
            if ms_report_data[hdev] and ms_report_data[hdev][-1][2] < spec_value:
                stop_monitoring()
                return list(device_lines_rate.values())

    # ───────────────────────────────────────────────────
    # 連續 5 秒滿足 SPEC 時，處理 Attenuation 遞減或 Stop 流程
    if spec_value is not None and elapsed_time >= 2:
        # 檢查所有滑鼠是否都 ≥ SPEC
        all_meet = True
        for hdev in device_event_times.keys():
            if not (ms_report_data[hdev] and ms_report_data[hdev][-1][2] >= spec_value):
                all_meet = False
                break

        if not all_meet:
            # 如果不都滿足 SPEC，重置 streak，並如果進入 Stop 階段則檢查是否需立即停止
            spec_streak_start = None
            if at_stop_phase:
                # 已在 Stop 階段，只要有不滿足 SPEC 就立即停止
                stop_monitoring()
                return list(device_lines_rate.values())
        else:
            # 所有滑鼠都 ≥ SPEC
            if not at_stop_phase:
                # 還沒到 Stop 值前，一般情況下檢查 5 秒
                if spec_streak_start is None:
                    spec_streak_start = now
                elif now - spec_streak_start >= 5:
                    # 連續 5 秒滿足 SPEC，開始遞減 Attenuation
                    try:
                        step_val = float(step_var.get().strip())
                    except:
                        step_val = 0
                    try:
                        stop_val = float(stop_var.get().strip())
                    except:
                        stop_val = None

                    if current_atten is not None and step_val > 0 and stop_val is not None:
                        new_atten = current_atten - step_val
                        if new_atten <= stop_val:
                            # 遞減後 <= Stop，設定到 Stop，進入 Stop 階段但不停止
                            set_attenuation_db(stop_val)
                            atten_var.set(str(stop_val))
                            at_stop_phase = True
                            at_stop_start = now
                        else:
                            # 正常遞減
                            set_attenuation_db(new_atten)
                            atten_var.set(str(new_atten))
                            spec_streak_start = now
                    else:
                        # 參數不足，重置 streak 以免無限迴圈
                        spec_streak_start = None
            else:
                # 已在 Stop 階段，只要連續 5 秒滿足 SPEC 就停止
                if at_stop_start is None:
                    at_stop_start = now
                elif now - at_stop_start >= 5:
                    stop_monitoring()
                    return list(device_lines_rate.values())

    return list(device_lines_rate.values())

def start_monitoring():
    global monitoring, chart_start_time, ani, next_color_index, last_event_time, y_max_fixed
    global spec_line, spec_streak_start, at_stop_phase, at_stop_start
    global start_time_str, stop_time_str

    if ani is not None:
        ani.event_source.stop()
        ani = None

    device_event_times.clear()
    device_smoothed_rate.clear()
    for h in list(device_plot_rates.keys()):
        device_plot_rates[h].clear()
    device_plot_rates.clear()
    device_positions.clear()
    plot_times.clear()
    for lbl in rate_labels.values():
        lbl.destroy()
    rate_labels.clear()
    device_lines_rate.clear()
    device_colors.clear()
    next_color_index = 0
    event_counter.clear()
    pre_event_counter.clear()
    ms_report_data.clear()
    device_vid_pid.clear()
    y_max_fixed = None
    spec_streak_start = None
    at_stop_phase = False
    at_stop_start = None

    now = time.monotonic()
    chart_start_time = now
    last_event_time = now
    monitoring = True

    start_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    stop_time_str = None

    start_label.config(text=time.strftime("Start: %Y-%m-%d %H:%M:%S", time.localtime()))
    stop_label.config(text="Stop: N/A")

    ax_rate.clear()
    ax_rate.set_title("各滑鼠 Report Rate Over Time", fontdict={"family":"Microsoft JhengHei", "size":14})
    ax_rate.set_xlabel("Time (s)", fontdict={"family":"Microsoft JhengHei", "size":12})
    ax_rate.set_ylabel("Hz", fontdict={"family":"Microsoft JhengHei", "size":12})
    ax_rate.set_ylim(0, 1000)
    ax_rate.set_yticks(range(0, 1001, 500))

    spec_line = ax_rate.axhline(y=0, color='red', linestyle='--', visible=False)

    ax_traj.clear()
    ax_traj.set_title("滑鼠軌跡（多裝置）", fontdict={"family":"Microsoft JhengHei", "size":14})
    ax_traj.set_xlabel("X", fontdict={"family":"Microsoft JhengHei", "size":12})
    ax_traj.set_ylabel("Y", fontdict={"family":"Microsoft JhengHei", "size":12})
    canvas.draw_idle()

    ani = animation.FuncAnimation(fig, update_plot, interval=10, cache_frame_data=False)

def stop_monitoring():
    global monitoring, ani, save_path, stop_time_str
    if not monitoring:
        return
    monitoring = False

    stop_time_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    stop_label.config(text=time.strftime("Stop: %Y-%m-%d %H:%M:%S", time.localtime()))

    if ani is not None:
        ani.event_source.stop()
        ani = None

    save_results()

def select_save():
    global save_path
    save_path = filedialog.asksaveasfilename(defaultextension='.csv',
                                             filetypes=[('CSV 檔', '*.csv')])
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
if __name__ == '__main__':
    setup_ui()
    threading.Thread(target=raw_input_loop, daemon=True).start()
    root.mainloop()
