"""
Lists running Office and PDF processes, flags hung or dialog-blocked ones,
and brings their windows to the front or kills them.

Author: Yeoh Huey Sun
Email:  greyw901@gmail.com
GitHub: https://github.com/greyw901
"""
from __future__ import annotations

import ctypes
import datetime
import os
import platform
import re
import sys
import time
import traceback
from ctypes import wintypes
from dataclasses import dataclass, field

try:
    import psutil
except ImportError:  # reported on screen by entry(), not as a crash
    psutil = None


# ---------------------------------------------------------------- constants

TOOL_NAME = "Office Process Checker"
WIDTH = 56
KILL_WAIT_SECONDS = 5          # how long to wait for a killed process to exit
ONSCREEN_POS = (100, 100)      # where an off-screen window is moved to

# Executable name (upper case) -> program label. Dict order = list order.
TARGET_EXES = {
    "EXCEL.EXE":          "Microsoft Excel",
    "WINWORD.EXE":        "Microsoft Word",
    "POWERPNT.EXE":       "Microsoft PowerPoint",
    "OUTLOOK.EXE":        "Microsoft Outlook",
    "OLK.EXE":            "Microsoft Outlook (new)",
    "ONENOTE.EXE":        "Microsoft OneNote",
    "MSACCESS.EXE":       "Microsoft Access",
    "MSPUB.EXE":          "Microsoft Publisher",
    "ACROBAT.EXE":        "Adobe Acrobat",
    "ACROCEF.EXE":        "Adobe Acrobat (helper)",
    "ACRORD32.EXE":       "Adobe Reader",
    "RDRCEF.EXE":         "Adobe Reader (helper)",
    "FOXITPDFEDITOR.EXE": "Foxit PDF Editor",
    "FOXITPHANTOMPDF.EXE": "Foxit PhantomPDF",
    "FOXITPDFREADER.EXE": "Foxit PDF Reader",
    "FOXITREADER.EXE":    "Foxit Reader",
}

STATUS_HUNG = "NOT RESPONDING"
STATUS_DIALOG = "DIALOG OPEN"
STATUS_RUNNING = "running"
STATUS_NO_WINDOW = "no window"
STATUS_RANK = {STATUS_HUNG: 0, STATUS_DIALOG: 1, STATUS_RUNNING: 2, STATUS_NO_WINDOW: 3}


# ---------------------------------------------------------------- Win32 API

user32 = ctypes.WinDLL("user32", use_last_error=True)

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

user32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
user32.EnumWindows.restype = wintypes.BOOL
user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
user32.GetWindowThreadProcessId.restype = wintypes.DWORD
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
for _fn in ("IsWindowVisible", "IsWindowEnabled", "IsIconic", "IsHungAppWindow",
            "SetForegroundWindow"):
    getattr(user32, _fn).argtypes = [wintypes.HWND]
    getattr(user32, _fn).restype = wintypes.BOOL
user32.MonitorFromWindow.argtypes = [wintypes.HWND, wintypes.DWORD]
user32.MonitorFromWindow.restype = wintypes.HANDLE
user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                ctypes.c_int, ctypes.c_int, wintypes.UINT]
user32.SetWindowPos.restype = wintypes.BOOL
user32.ShowWindowAsync.argtypes = [wintypes.HWND, ctypes.c_int]
user32.ShowWindowAsync.restype = wintypes.BOOL
user32.GetForegroundWindow.argtypes = []
user32.GetForegroundWindow.restype = wintypes.HWND
user32.keybd_event.argtypes = [wintypes.BYTE, wintypes.BYTE, wintypes.DWORD, ctypes.c_size_t]
user32.keybd_event.restype = None

MONITOR_DEFAULTTONULL = 0
HWND_TOPMOST = wintypes.HWND(-1)
HWND_NOTOPMOST = wintypes.HWND(-2)
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SWP_ASYNCWINDOWPOS = 0x4000    # never block on a hung window's thread
SW_RESTORE = 9
VK_MENU = 0x12
KEYEVENTF_KEYUP = 0x0002
TITLE_BUFFER = 512


# ---------------------------------------------------------------- signals

class UserCancelled(Exception):
    """[00] pressed, or Ctrl+C."""


class RefreshList(Exception):
    """[R] pressed on the process list."""


# ---------------------------------------------------------------- data

@dataclass
class WindowInfo:
    hwnd: int
    title: str
    enabled: bool
    hung: bool
    minimized: bool
    offscreen: bool


@dataclass
class ProcInfo:
    pid: int
    name: str
    label: str
    created: float
    memory: str
    proc: "psutil.Process"
    windows: list[WindowInfo] = field(default_factory=list)

    @property
    def blocked(self) -> bool:
        # A modal dialog disables its owner window.
        return any(not w.enabled for w in self.windows)

    @property
    def status(self) -> str:
        if any(w.hung for w in self.windows):
            return STATUS_HUNG
        if self.blocked:
            return STATUS_DIALOG
        if not self.windows:
            return STATUS_NO_WINDOW
        return STATUS_RUNNING


# ---------------------------------------------------------------- scanning

def collect_windows() -> dict[int, list[WindowInfo]]:
    """One pass over all top-level visible windows, grouped by owning PID."""
    by_pid: dict[int, list[WindowInfo]] = {}

    def callback(hwnd: int, _lparam: int) -> bool:
        try:
            if not hwnd or not user32.IsWindowVisible(hwnd):
                return True
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            # GetWindowTextW reads cached text for other processes' windows,
            # so it never blocks on a hung window (GetWindowTextLength can).
            buf = ctypes.create_unicode_buffer(TITLE_BUFFER)
            user32.GetWindowTextW(hwnd, buf, TITLE_BUFFER)
            minimized = bool(user32.IsIconic(hwnd))
            offscreen = (not minimized
                         and not user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONULL))
            by_pid.setdefault(pid.value, []).append(WindowInfo(
                hwnd=hwnd,
                title=buf.value.strip(),
                enabled=bool(user32.IsWindowEnabled(hwnd)),
                hung=bool(user32.IsHungAppWindow(hwnd)),
                minimized=minimized,
                offscreen=offscreen,
            ))
        except Exception:
            pass  # one odd window must not stop the enumeration
        return True

    user32.EnumWindows(WNDENUMPROC(callback), 0)
    return by_pid


def format_age(created: float) -> str:
    secs = max(0, int(time.time() - created))
    d, rem = divmod(secs, 86400)
    h, rem = divmod(rem, 3600)
    m, s = divmod(rem, 60)
    if d:
        return f"{d}d {h}h"
    if h:
        return f"{h}h {m}m"
    if m:
        return f"{m}m {s}s"
    return f"{s}s"


def scan_processes() -> list[ProcInfo]:
    windows = collect_windows()
    found: list[ProcInfo] = []
    for proc in psutil.process_iter(["pid", "name", "create_time", "memory_info"]):
        name = proc.info["name"] or ""
        label = TARGET_EXES.get(name.upper())
        if label is None:
            continue
        mem = proc.info["memory_info"]
        found.append(ProcInfo(
            pid=proc.info["pid"],
            name=name,
            label=label,
            created=proc.info["create_time"] or time.time(),
            memory=f"{mem.rss / (1024 * 1024):.0f} MB" if mem else "unknown",
            proc=proc,
            windows=windows.get(proc.info["pid"], []),
        ))
    order = list(TARGET_EXES)
    found.sort(key=lambda p: (order.index(p.name.upper()), STATUS_RANK[p.status], p.created))
    return found


# ---------------------------------------------------------------- actions (workers: no printing)

def raise_targets(windows: list[WindowInfo]) -> list[WindowInfo]:
    """Off-screen windows first, then the dialog(s) or main window - last one gets focus."""
    if any(not w.enabled for w in windows):
        front = [w for w in windows if w.enabled]
    else:
        titled = [w for w in windows if w.title]
        front = [max(titled, key=lambda w: len(w.title))] if titled else windows[:1]
    extra = [w for w in windows if w.offscreen and w not in front]
    return extra + front


def raise_window(w: WindowInfo, focus: bool) -> list[str]:
    """Brings one window forward. Returns notes; raises OSError on API failure."""
    notes: list[str] = []
    if w.minimized:
        user32.ShowWindowAsync(w.hwnd, SW_RESTORE)
        notes.append("restored")
    if w.offscreen:
        x, y = ONSCREEN_POS
        if not user32.SetWindowPos(w.hwnd, None, x, y, 0, 0,
                                   SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS):
            raise ctypes.WinError(ctypes.get_last_error())
        notes.append("moved on-screen")
    # Topmost on/off lifts the window above others even if focus is refused.
    for z in (HWND_TOPMOST, HWND_NOTOPMOST):
        if not user32.SetWindowPos(w.hwnd, z, 0, 0, 0, 0,
                                   SWP_NOMOVE | SWP_NOSIZE | SWP_ASYNCWINDOWPOS):
            raise ctypes.WinError(ctypes.get_last_error())
    if focus:
        # A key event lets a background process take the foreground.
        user32.keybd_event(VK_MENU, 0, 0, 0)
        user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
        user32.SetForegroundWindow(w.hwnd)
    return notes


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


# ---------------------------------------------------------------- screen helpers

class Screen:
    """Prints lines and remembers them, so a screen can be redrawn on invalid input."""

    def __init__(self) -> None:
        self.lines: list[str] = []

    def out(self, line: str = "") -> None:
        print(line, flush=True)
        self.lines.append(line)

    def redraw(self) -> None:
        os.system("cls")
        for line in self.lines:
            print(line)


def clear() -> None:
    os.system("cls")


def banner() -> None:
    print("=" * WIDTH)
    print(TOOL_NAME)
    print("=" * WIDTH)
    print()


def options(lines: list[str]) -> None:
    print("-" * WIDTH)
    for line in lines:
        print(line)
    print()


def message_slot(message: str) -> None:
    if message:
        print(message)
        print()


def ask(prompt: str) -> str:
    raw = input(prompt).strip()
    if raw == "00":
        raise UserCancelled
    return raw


def proc_line(i: int, p: ProcInfo) -> str:
    return f'   {i}. "{p.name}"   PID {p.pid}   {p.status}'


def chosen_block(selected: list[ProcInfo]) -> None:
    print(f"Processes: {len(selected)}")
    for i, p in enumerate(selected, 1):
        print(proc_line(i, p))
    print()


def window_tags(w: WindowInfo, is_dialog: bool) -> str:
    tags = []
    if is_dialog:
        tags.append("dialog")
    if w.hung:
        tags.append("not responding")
    if w.offscreen:
        tags.append("off-screen")
    if w.minimized:
        tags.append("minimized")
    return f"   ({', '.join(tags)})" if tags else ""


# ---------------------------------------------------------------- screens

def scan_screen() -> list[ProcInfo]:
    clear()
    banner()
    print("Scanning ...", flush=True)
    return scan_processes()


def parse_selection(raw: str, procs: list[ProcInfo]) -> list[ProcInfo]:
    """Numbers, ranges and program names -> processes. ValueError carries the reason."""
    n = len(procs)
    picked: list[int] = []
    for token in re.split(r"[\s,]+", raw):
        if not token:
            continue
        m = re.fullmatch(r"(\d+)-(\d+)", token)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b:
                raise ValueError(f"range '{token}' must be ascending")
            if a < 1 or b > n:
                raise ValueError(f"range '{token}' is outside 1-{n}")
            picked.extend(range(a, b + 1))
        elif token.isdigit():
            k = int(token)
            if not 1 <= k <= n:
                hint = " (use the [n] number, not the PID)" if any(p.pid == k for p in procs) else ""
                raise ValueError(f"'{token}' is outside 1-{n}{hint}")
            picked.append(k)
        else:
            key = token.upper()
            if not key.endswith(".EXE"):
                key += ".EXE"
            hits = [i for i, p in enumerate(procs, 1) if p.name.upper() == key]
            if not hits:
                raise ValueError(f"'{token}' is not a number, range or listed program")
            picked.extend(hits)
    if not picked:
        raise ValueError("nothing selected")
    seen: set[int] = set()
    return [procs[k - 1] for k in picked if not (k in seen or seen.add(k))]


def select_screen(procs: list[ProcInfo]) -> list[ProcInfo] | None:
    """Process list. Returns the selection, None for Exit; raises RefreshList."""
    # Deviation: scanning runs before this selection screen - the scan result IS
    # what the user selects from, and scanning only reads, never changes anything.
    message = ""
    while True:
        clear()
        banner()
        message_slot(message)
        if not procs:
            print("No Office or PDF processes are running.")
            print()
            options(["[R] Refresh", "[0] Exit"])
        else:
            print("Select PROCESS:")
            last_label = None
            for i, p in enumerate(procs, 1):
                if p.label != last_label:
                    print()
                    print(f"   {p.label} - {p.name}")
                    last_label = p.label
                key = f"[{i}]"
                print(f"   {key:<5} PID {p.pid:<7} {format_age(p.created):>8}  "
                      f"{p.memory:>8}   {p.status}")
                for w in p.windows:
                    if w.title:
                        is_dialog = p.blocked and w.enabled
                        print(f'           "{w.title}"{window_tags(w, is_dialog)}')
            print()
            print("Pick by number (1 3), range (2-4) or program name (excel).")
            print()
            options(["[#] Select", "[R] Refresh", "[0] Exit"])

        raw = ask("Process: ")
        if raw == "0":
            return None
        if raw.upper() == "R":
            raise RefreshList
        if not procs:
            message = "Error: Invalid Input"
            continue
        try:
            return parse_selection(raw, procs)
        except ValueError as e:
            message = f"Error: Invalid Input - {e}"


def action_screen(selected: list[ProcInfo]) -> str | None:
    """Returns 'F' or 'K', None for Back."""
    message = ""
    while True:
        clear()
        banner()
        chosen_block(selected)
        message_slot(message)
        print("Select ACTION:")
        print()
        print("   Bring to front : shows the window or hidden dialog, nothing is lost")
        print("   Kill           : ends the process at once, unsaved work is lost")
        print()
        options(["[F] Bring to front", "[K] Kill", "[0] Back"])

        raw = ask("Action: ").upper()
        if raw == "0":
            return None
        if raw == "F":
            if not any(p.windows for p in selected):
                message = "Error: Invalid Input - selected processes have no window, only Kill applies"
                continue
            return "F"
        if raw == "K":
            return "K"
        message = "Error: Invalid Input"


def confirm_kill_screen(selected: list[ProcInfo]) -> bool:
    """True = continue, False = back."""
    message = ""
    dialogs = sum(1 for p in selected if p.status == STATUS_DIALOG)
    while True:
        clear()
        banner()
        chosen_block(selected)
        print("Action : Kill")
        print()
        message_slot(message)
        print(f"Kill {len(selected)} {'process' if len(selected) == 1 else 'processes'}? "
              "Unsaved work in them will be lost.")
        if dialogs:
            print(f"Processes with a dialog open: {dialogs} - "
                  "[0] Back and try Bring to front first.")
        print()
        options(["[Enter] Continue", "[0] Back", "[00] Cancel & exit"])

        raw = ask("Ans: ")
        if raw == "":
            return True
        if raw == "0":
            return False
        message = "Error: Invalid Input"


def failure(p: ProcInfo, step: str, e: BaseException, trace: str) -> dict:
    titles = [w.title for w in p.windows if w.title]
    return {
        "item": f'"{p.name}"   PID {p.pid}',
        "step": step,
        "error": f"{type(e).__name__} - {e}",
        "context": f"status = {p.status}, windows = {titles}, admin = {is_admin()}",
        "trace": trace,
    }


def run_kill(selected: list[ProcInfo], scr: Screen) -> list[dict]:
    failures: list[dict] = []
    killed: list[ProcInfo] = []
    skipped: list[tuple[ProcInfo, str]] = []
    sent: list[ProcInfo] = []

    scr.out("Killing ...")
    for p in selected:
        try:
            if not p.proc.is_running():      # also catches a reused PID
                skipped.append((p, "already exited"))
                continue
            p.proc.kill()
            sent.append(p)
        except psutil.NoSuchProcess:
            skipped.append((p, "already exited"))
        except psutil.AccessDenied as e:
            failures.append(failure(p, "killing", e, traceback.format_exc()))
        except Exception as e:
            failures.append(failure(p, "killing", e, traceback.format_exc()))

    if sent:
        scr.out("Waiting for exit ...")
        _gone, alive = psutil.wait_procs([p.proc for p in sent], timeout=KILL_WAIT_SECONDS)
        alive_pids = {a.pid for a in alive}
        for p in sent:
            if p.pid in alive_pids:
                e = TimeoutError(f"still running {KILL_WAIT_SECONDS}s after kill")
                failures.append(failure(p, "waiting for exit", e, ""))
            else:
                killed.append(p)

    scr.out()
    scr.out(f"Killed: {len(killed)}")
    for i, p in enumerate(killed, 1):
        scr.out(f'   {i}. "{p.name}"   PID {p.pid}')
    if skipped:
        scr.out()
        scr.out(f"Skipped: {len(skipped)}")
        for i, (p, why) in enumerate(skipped, 1):
            scr.out(f'   {i}. "{p.name}"   PID {p.pid}   :   {why}')
    return failures


def run_raise(selected: list[ProcInfo], scr: Screen) -> list[dict]:
    failures: list[dict] = []
    raised: list[tuple[ProcInfo, WindowInfo, list[str]]] = []
    skipped: list[tuple[ProcInfo, str]] = []

    scr.out("Raising windows ...")
    fresh = collect_windows()          # window handles may have changed since the scan
    last_hwnd = None
    with_windows = [p for p in selected if fresh.get(p.pid)]
    for p in selected:
        if not fresh.get(p.pid):
            skipped.append((p, "no window"))
    for pi, p in enumerate(with_windows):
        targets = raise_targets(fresh[p.pid])
        for wi, w in enumerate(targets):
            focus = pi == len(with_windows) - 1 and wi == len(targets) - 1
            try:
                notes = raise_window(w, focus)
                raised.append((p, w, notes))
                if focus:
                    last_hwnd = w.hwnd
            except Exception as e:
                failures.append(failure(p, f'raising window "{w.title}"', e, traceback.format_exc()))

    focus_ok = True
    if last_hwnd:
        time.sleep(0.3)
        focus_ok = user32.GetForegroundWindow() == last_hwnd

    scr.out()
    scr.out(f"Windows raised: {len(raised)}")
    for i, (p, w, notes) in enumerate(raised, 1):
        title = w.title or "(untitled)"
        extra = f"   ({', '.join(notes)})" if notes else ""
        scr.out(f'   {i}. "{title}"   PID {p.pid}{extra}')
    if skipped:
        scr.out()
        scr.out(f"Skipped: {len(skipped)}")
        for i, (p, why) in enumerate(skipped, 1):
            scr.out(f'   {i}. "{p.name}"   PID {p.pid}   :   {why}')
    if raised:
        scr.out()
        if not focus_ok:
            scr.out("Note: Windows kept focus here - click the flashing taskbar button.")
        scr.out("Close any dialog that appeared, then return to the list to check again.")
    return failures


def result_screen(selected: list[ProcInfo], action: str) -> bool:
    """Runs the action live. Returns True when the user exits."""
    scr = Screen()
    clear()
    for line in ["=" * WIDTH, TOOL_NAME, "=" * WIDTH, ""]:
        scr.out(line)
    scr.out(f"Processes: {len(selected)}")
    for i, p in enumerate(selected, 1):
        scr.out(proc_line(i, p))
    scr.out(f"Action : {'Kill' if action == 'K' else 'Bring to front'}")
    scr.out()
    scr.out("Processing")

    start = time.perf_counter()
    failures = run_kill(selected, scr) if action == "K" else run_raise(selected, scr)
    elapsed = time.perf_counter() - start

    if failures:
        scr.out()
        scr.out(f"Failed: {len(failures)}")
        for i, f in enumerate(failures, 1):
            scr.out(f"   {i}. {f['item']}   :   {f['error']}")
        if action == "K" and not is_admin() and any("AccessDenied" in f["error"] for f in failures):
            scr.out("   Access denied means the process runs as Administrator -")
            scr.out("   right-click this tool and choose Run as administrator.")
    scr.out()
    scr.out(f"Elapsed: {elapsed:.1f}s")
    scr.out()

    message = ""
    while True:
        message_slot(message)
        options(["[Enter] Back to list",
                 "[0] Save error log & exit" if failures else "[0] Exit"])
        raw = ask("Ans: ")
        if raw == "":
            return False
        if raw == "0":
            if failures:
                name = save_log(failures, selected, action)
                print()
                print(f'Saved: "{name}"')
                print()
                input("[Enter] to close...")
            return True
        message = "Error: Invalid Input"
        scr.redraw()


# ---------------------------------------------------------------- error log

def save_log(failures: list[dict], selected: list[ProcInfo] | None, action: str | None) -> str:
    base = f"error log_{datetime.date.today():%Y%m%d}"
    name = f"{base}.txt"
    n = 1
    while os.path.exists(name):
        name = f"{base}_{n}.txt"
        n += 1
    lines = [
        f"{TOOL_NAME} - error log",
        f"Run at        : {datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
        "Input folder  : (none - reads running processes)",
        f"Output folder : {os.getcwd()}",
        f"Python        : {sys.version.split()[0]} ({platform.architecture()[0]})",
        f"Windows       : {platform.version()}",
        f"psutil        : {getattr(psutil, '__version__', 'not installed')}",
        f"Administrator : {'yes' if is_admin() else 'no'}",
    ]
    if action:
        lines.append(f"Action        : {'Kill' if action == 'K' else 'Bring to front'}")
    if selected:
        lines.append(f"Processes: {len(selected)}")
        lines += [proc_line(i, p) for i, p in enumerate(selected, 1)]
    lines.append("")
    for f in failures:
        lines += [
            f["item"],
            f"  Step: {f['step']}",
            f"  Error: {f['error']}",
            f"  Context: {f['context']}",
            "  Traceback:",
        ]
        trace = f["trace"].strip() or "(none - not an exception)"
        lines += ["    " + t for t in trace.splitlines()]
        lines.append("")
    with open(name, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))
    return name


# ---------------------------------------------------------------- main

def main() -> None:
    procs = scan_screen()
    while True:
        try:
            selected = select_screen(procs)
        except RefreshList:
            procs = scan_screen()
            continue
        if selected is None:
            return

        action = None
        while True:
            action = action_screen(selected)
            if action != "K" or confirm_kill_screen(selected):
                break
        if action is None:
            continue

        if result_screen(selected, action):
            return
        procs = scan_screen()


def end_screen(message: str, failures: list[dict] | None = None) -> None:
    print()
    print(message)
    print()
    print("-" * WIDTH)
    if not failures:
        input("[Enter] to close...")
        return
    print("[Enter] to close...")
    print("[0] Save error log & close")
    print()
    while True:
        raw = input("Ans: ").strip()
        if raw == "0":
            name = save_log(failures, None, None)
            print()
            print(f'Saved: "{name}"')
            print()
            input("[Enter] to close...")
            return
        if raw == "":
            return


def entry() -> None:
    script_dir = os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, "frozen", False) else __file__
    ))
    os.chdir(script_dir)
    # A window title the console cannot encode must print as '?', not crash the tool.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    try:
        if psutil is None:
            clear()
            banner()
            end_screen("Error: psutil is not installed.\n"
                       "Run this once in Command Prompt:   pip install psutil")
            return
        main()
    except (UserCancelled, KeyboardInterrupt):
        print()
        print("Cancelled.")
        print()
        input("[Enter] to close...")
    except Exception as e:
        failure_record = {
            "item": f'"{TOOL_NAME}"',
            "step": "unexpected error",
            "error": f"{type(e).__name__} - {e}",
            "context": f"script folder = {script_dir}",
            "trace": traceback.format_exc(),
        }
        end_screen(f"Error: {type(e).__name__} - {e}", [failure_record])


if __name__ == "__main__":
    entry()
