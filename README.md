# Office Process Checker

A small Windows tool for the moment when double-clicking a PDF, Excel or
Word file does nothing — or an Office / PDF program has frozen.

It lists every running Office and PDF process, shows **why** each one
might be stuck, and lets you fix it — first the safe way, then the hard
way if needed.

## The problem

Most "the file won't open" cases are not a crash. The program is
**waiting for you to answer a message box you can't see** — hidden
behind other windows, or positioned off the edge of the screen. While it
waits, every new file you double-click is handed to that stuck instance
and silently dropped.

Task Manager doesn't help much here: it shows the program as running,
and the usual fix — End task — throws away any unsaved work in it.

## What it does

- Lists every running Office / PDF process, grouped by program, with
  problems first. Each process shows its PID, running time, memory, and
  the titles of **all** its open windows — so a hidden dialog becomes
  visible by name.
- Flags each process:
  - `NOT RESPONDING` — frozen (the same test Task Manager uses)
  - `DIALOG OPEN` — waiting on a message box; the box itself is tagged
    `(dialog)`, and `(off-screen)` if it's outside every monitor
  - `running` / `no window`
- Two actions on the processes you pick:
  - **Bring to front** — raises the dialog (or main window) above
    everything and moves off-screen windows back onto the screen.
    Nothing is closed, nothing is lost. Try this first.
  - **Kill** — ends the process, after a confirm screen. It checks the
    PID still belongs to the same process before killing, then waits to
    confirm it actually exited.
- After either action, return to the list to re-scan and check the
  problem is gone.

Pick processes by list number (`1 3`), range (`2-4`), or program name
(`excel`, `acrobat`) to select every process of that program.

### Programs covered

Excel, Word, PowerPoint, Outlook (classic and new), OneNote, Access,
Publisher, Adobe Acrobat / Reader (including their `AcroCEF` / `RdrCEF`
helper processes), and Foxit PDF Editor / Reader.

## Usage

```
pip install psutil
python "Office Process Checker.py"
```

Or double-click `Office Process Checker.py` in File Explorer.

To kill a program that runs as Administrator, run the tool as
Administrator too. Your normal Office programs never need this.

Full walkthrough, sample screens and troubleshooting:
[`office_process_checker_instruction.txt`](office_process_checker_instruction.txt).

## Requirements

- Windows (uses the Windows API through `ctypes`)
- Python 3.8+
- `psutil`

## License

MIT — see [LICENSE](LICENSE).

---
Author: Yeoh Huey Sun
Email: greyw901@gmail.com
GitHub: https://github.com/greyw901
