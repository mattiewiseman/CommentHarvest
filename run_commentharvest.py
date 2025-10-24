#!/usr/bin/env python3
import os
os.environ["TK_SILENCE_DEPRECATION"] = "1"  # hide macOS Tk banner

import subprocess
import sys
import threading
from pathlib import Path
from typing import Optional

# ---------- UI helpers ----------
def mac_dialog(title: str, message: str):
    try:
        # escape double quotes for AppleScript
        t = title.replace('"', r'\"')
        m = message.replace('"', r'\"')
        subprocess.run(
            ["osascript", "-e",
             f'display dialog "{m}" with title "{t}" buttons {{"OK"}} default button "OK" giving up after 10'],
            check=False
        )
    except Exception:
        pass

def pick_file_with_dialog() -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(
            title="Select a Word (.docx) file",
            filetypes=[("Word documents", "*.docx")],
        )
        root.update()
        root.destroy()
        return Path(path) if path else None
    except Exception:
        return None

def show_progress_window(start_text="Extracting comments…"):
    import tkinter as tk
    from tkinter import ttk
    win = tk.Tk()
    win.title("CommentHarvest")
    win.geometry("360x100")
    win.resizable(False, False)

    label = ttk.Label(win, text=start_text)
    label.pack(pady=(14, 8))

    bar = ttk.Progressbar(win, mode="indeterminate", length=300)
    bar.pack(pady=(0, 10))
    bar.start(10)  # smaller = faster spin

    # make sure closing the window doesn’t kill the process abruptly
    def disable_close():
        pass
    win.protocol("WM_DELETE_WINDOW", disable_close)

    return win, label, bar

# ---------- main ----------
def run_extractor(input_path: Path, output_path: Path) -> subprocess.CompletedProcess:
    cmd = [
        sys.executable, "-m", "src.extract_docx_comments",
        str(input_path), "-o", str(output_path),
        "--author", "--date",
    ]
    return subprocess.run(cmd, capture_output=True, text=True, check=False)

def main():
    # 1) input from drag&drop, else picker
    input_path: Optional[Path] = None
    if len(sys.argv) >= 2:
        input_path = Path(sys.argv[1]).expanduser()
    if not input_path:
        input_path = pick_file_with_dialog()
    if not input_path:
        mac_dialog("CommentHarvest", "No file selected.")
        return

    input_path = input_path.resolve()
    if not input_path.exists():
        mac_dialog("CommentHarvest", f"File not found:\n{input_path}")
        return

    output_path = Path.home() / "Desktop" / f"{input_path.stem}_comments.xlsx"

    # 2) show progress while extractor runs in a background thread
    try:
        import tkinter as tk
        from tkinter import ttk  # noqa: F401  (ensures ttk available)
        win, label, bar = show_progress_window()

        result_holder = {"result": None, "done": False}

        def worker():
            result_holder["result"] = run_extractor(input_path, output_path)
            result_holder["done"] = True

        th = threading.Thread(target=worker, daemon=True)
        th.start()

        def poll():
            if result_holder["done"]:
                bar.stop()
                win.destroy()
                r = result_holder["result"]
                if r is None or r.returncode != 0:
                    stdout = (r.stdout if r else "").strip()
                    stderr = (r.stderr if r else "").strip()
                    msg = "Extractor failed."
                    if stdout: msg += f"\n\nStdout:\n{stdout}"
                    if stderr: msg += f"\n\nStderr:\n{stderr}"
                    mac_dialog("CommentHarvest – Error", msg)
                    return
                try:
                    subprocess.run(["open", str(output_path)], check=False)
                finally:
                    mac_dialog("CommentHarvest",
                               f"Export complete ✅\n\nSaved to:\n{output_path}")
            else:
                win.after(120, poll)

        win.after(120, poll)
        win.mainloop()

    except Exception:
        # Fallback: run without progress UI
        r = run_extractor(input_path, output_path)
        if r.returncode != 0:
            mac_dialog("CommentHarvest – Error",
                       f"Extractor failed.\n\nStdout:\n{r.stdout}\n\nStderr:\n{r.stderr}")
            return
        try:
            subprocess.run(["open", str(output_path)], check=False)
        finally:
            mac_dialog("CommentHarvest",
                       f"Export complete ✅\n\nSaved to:\n{output_path}")

if __name__ == "__main__":
    main()
