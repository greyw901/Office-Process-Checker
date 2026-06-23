import os
import sys
import ctypes
import psutil
import datetime

TARGET_EXES = {
    # Microsoft Office
    "EXCEL.EXE",
    "WINWORD.EXE",
    "POWERPNT.EXE",
    "OUTLOOK.EXE",
    "ONENOTE.EXE",
    "MSACCESS.EXE",
    "MSPUB.EXE",
    # Adobe PDF
    "ACROBAT.EXE",
    "ACRORD32.EXE",
    # Foxit PDF
    "FOXITPDFEDITOR.EXE",
    "FOXITREADER.EXE",
}

def get_window_title(pid):
    EnumWindows        = ctypes.windll.user32.EnumWindows
    EnumWindowsProc    = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
    GetWindowThreadPID = ctypes.windll.user32.GetWindowThreadProcessId
    IsWindowVisible    = ctypes.windll.user32.IsWindowVisible
    GetWindowTextW     = ctypes.windll.user32.GetWindowTextW
    GetWindowTextLen   = ctypes.windll.user32.GetWindowTextLengthW

    titles = []

    def callback(hwnd, _):
        win_pid = ctypes.c_ulong()
        GetWindowThreadPID(hwnd, ctypes.byref(win_pid))
        if win_pid.value == pid and IsWindowVisible(hwnd):
            length = GetWindowTextLen(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                GetWindowTextW(hwnd, buf, length + 1)
                title = buf.value.strip()
                if title:
                    titles.append(title)
        return True

    EnumWindows(EnumWindowsProc(callback), 0)
    return max(titles, key=len) if titles else ""

def is_not_responding(pid):
    SMTO_ABORTIFHUNG = 0x0002
    WM_NULL = 0x0000
    EnumWindows        = ctypes.windll.user32.EnumWindows
    EnumWindowsProc    = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)
    GetWindowThreadPID = ctypes.windll.user32.GetWindowThreadProcessId
    IsWindowVisible    = ctypes.windll.user32.IsWindowVisible
    SendMessageTimeout = ctypes.windll.user32.SendMessageTimeoutW

    found_hung = []

    def callback(hwnd, _):
        win_pid = ctypes.c_ulong()
        GetWindowThreadPID(hwnd, ctypes.byref(win_pid))
        if win_pid.value == pid and IsWindowVisible(hwnd):
            result = ctypes.c_ulong()
            ret = SendMessageTimeout(
                hwnd, WM_NULL, 0, 0,
                SMTO_ABORTIFHUNG, 3000,
                ctypes.byref(result)
            )
            if ret == 0:
                found_hung.append(hwnd)
        return True

    EnumWindows(EnumWindowsProc(callback), 0)
    return len(found_hung) > 0

def format_duration(proc):
    try:
        delta = datetime.datetime.now() - datetime.datetime.fromtimestamp(proc.create_time())
        h, rem = divmod(int(delta.total_seconds()), 3600)
        m, s   = divmod(rem, 60)
        if h > 0:   return f"{h}h {m}m"
        elif m > 0: return f"{m}m {s}s"
        else:       return f"{s}s"
    except Exception:
        return "unknown"

def format_mb(proc):
    try:
        return f"{proc.memory_info().rss / (1024*1024):.0f} MB"
    except Exception:
        return "unknown"

def kill_proc(e):
    try:
        e['proc'].kill()
        title_str = e['title'] if e['title'] else "(no window)"
        print(f"  [KILLED] PID {e['pid']}  {e['name']}  {title_str}")
    except psutil.AccessDenied:
        print(f"  [FAILED] PID {e['pid']} -- Access denied. Run as Administrator.")
    except psutil.NoSuchProcess:
        print(f"  [SKIP]   PID {e['pid']} -- Process already gone.")

def main():
    os.chdir(os.path.dirname(os.path.abspath(sys.argv[0])))

    print("=" * 55)
    print("  Office / PDF Process Checker")
    print("=" * 55)
    print("Scanning processes...\n")

    responding     = []
    not_responding = []

    for proc in psutil.process_iter(['pid', 'name', 'create_time']):
        try:
            name = proc.info['name'].upper()
            if name not in TARGET_EXES:
                continue
            pid   = proc.info['pid']
            title = get_window_title(pid)
            hung  = is_not_responding(pid)
            entry = {
                "pid":      pid,
                "name":     proc.info['name'],
                "duration": format_duration(proc),
                "memory":   format_mb(proc),
                "title":    title,
                "proc":     proc,
            }
            (not_responding if hung else responding).append(entry)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    def print_entries(entries):
        for e in entries:
            title_str = f"[{e['title']}]" if e['title'] else "[no window]"
            print(f"  PID {e['pid']:<6}  {e['name']:<20}  {e['duration']:<12}  {e['memory']:<8}  {title_str}")

    if not responding and not not_responding:
        print("No Office or PDF processes found.")
        return

    if responding:
        print("[RUNNING NORMALLY]")
        print_entries(responding)
        print()

    if not_responding:
        print("[NOT RESPONDING]")
        print_entries(not_responding)
        print()

    all_entries = responding + not_responding
    name_map = {}
    for e in all_entries:
        name_map.setdefault(e['name'].upper(), []).append(e)
    pid_map = {str(e['pid']): e for e in all_entries}

    print("How to kill:")
    print("  By PID  -- kill specific process(es)   e.g.: 1808 12044")
    print("  By name -- kill all of that program     e.g.: Acrobat.exe")
    print("  Enter   -- do nothing and exit")
    print()
    raw = input("Input: ").strip()

    if not raw:
        print("No changes made.")
        return

    for token in raw.split():
        token_upper = token.upper()
        if token_upper in name_map:
            for e in name_map[token_upper]:
                kill_proc(e)
        elif token in pid_map:
            kill_proc(pid_map[token])
        else:
            print(f"  [SKIP] '{token}' -- not a recognised PID or process name")

try:
    main()
except Exception as e:
    print(f"\nUnexpected error: {e}")

input("\nPress Enter to close...")
