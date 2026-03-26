import ctypes
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WM_HOTKEY = 0x0312
MOD_ALT = 0x0001
VK_P = 0x50
SW_RESTORE = 9
MUTEX_NAME = "WindowsHelperSingleInstance"


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", ctypes.c_void_p),
        ("message", ctypes.c_uint),
        ("wParam", ctypes.c_size_t),
        ("lParam", ctypes.c_ssize_t),
        ("time", ctypes.c_uint),
        ("pt_x", ctypes.c_long),
        ("pt_y", ctypes.c_long),
        ("lPrivate", ctypes.c_uint),
    ]


def create_single_instance_mutex():
    mutex = kernel32.CreateMutexW(None, True, MUTEX_NAME)
    already_running = kernel32.GetLastError() == 183
    return mutex, already_running


class WindowsHelperApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Windows Helper")
        self.root.geometry("420x310")
        self.root.minsize(420, 310)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        self.root.withdraw()

        self.hotkey_registered = False
        self.visible = False
        self.status_var = tk.StringVar(value="Hotkey: Alt+P")

        self._build_ui()
        self._register_hotkey()
        self.root.after(0, self.hide_window)

    def _build_ui(self):
        self.root.configure(bg="#f4f6f8")

        frame = tk.Frame(self.root, bg="#f4f6f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)

        title = tk.Label(
            frame,
            text="Windows Helper",
            font=("Segoe UI Semibold", 20),
            bg="#f4f6f8",
            fg="#1d2733",
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            frame,
            text="Hidden at startup. Press Alt+P from anywhere to open or hide it.",
            font=("Segoe UI", 10),
            bg="#f4f6f8",
            fg="#4d5a69",
            wraplength=370,
            justify="left",
        )
        subtitle.pack(anchor="w", pady=(6, 18))

        buttons = [
            ("Open Settings", "start ms-settings:"),
            ("Open Task Manager", "taskmgr"),
            ("Open Startup Folder", 'explorer.exe shell:startup'),
            ("Open Command Prompt", "start cmd"),
            ("Show IP Config", "ipconfig"),
            ("Lock PC", "rundll32.exe user32.dll,LockWorkStation"),
        ]

        for text, command in buttons:
            tk.Button(
                frame,
                text=text,
                font=("Segoe UI", 10),
                bg="#ffffff",
                fg="#17202a",
                relief="solid",
                bd=1,
                anchor="w",
                padx=12,
                command=lambda cmd=command: self.run_command(cmd),
            ).pack(fill="x", pady=4)

        footer = tk.Frame(frame, bg="#f4f6f8")
        footer.pack(fill="x", pady=(18, 0))

        tk.Label(
            footer,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            bg="#f4f6f8",
            fg="#566574",
        ).pack(side="left")

        tk.Button(
            footer,
            text="Hide",
            font=("Segoe UI", 10),
            command=self.hide_window,
            padx=12,
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            footer,
            text="Exit",
            font=("Segoe UI", 10),
            command=self.exit_app,
            padx=12,
        ).pack(side="right")

    def run_command(self, command):
        if command == "ipconfig":
            output = subprocess.run(
                ["ipconfig"],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                check=False,
            ).stdout.strip()
            messagebox.showinfo("IP Configuration", output or "No output returned.")
            return

        subprocess.Popen(command, shell=True)

    def _register_hotkey(self):
        if not user32.RegisterHotKey(None, 1, MOD_ALT, VK_P):
            raise RuntimeError("Could not register Alt+P. Another app may already be using it.")

        self.hotkey_registered = True
        thread = threading.Thread(target=self._hotkey_loop, daemon=True)
        thread.start()

    def _hotkey_loop(self):
        msg = MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            if msg.message == WM_HOTKEY:
                self.root.after(0, self.toggle_window)
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    def toggle_window(self):
        if self.visible:
            self.hide_window()
        else:
            self.show_window()

    def show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(200, lambda: self.root.attributes("-topmost", False))
        user32.ShowWindow(self.root.winfo_id(), SW_RESTORE)
        user32.SetForegroundWindow(self.root.winfo_id())
        self.visible = True
        self.status_var.set("Visible. Press Alt+P again to hide.")

    def hide_window(self):
        self.root.withdraw()
        self.visible = False
        self.status_var.set("Hidden. Press Alt+P to show.")

    def exit_app(self):
        if self.hotkey_registered:
            user32.UnregisterHotKey(None, 1)
        self.root.destroy()

    def run(self):
        self.root.mainloop()


def main():
    mutex, already_running = create_single_instance_mutex()
    if already_running:
        if mutex:
            kernel32.CloseHandle(mutex)
        messagebox.showinfo("Windows Helper", "Windows Helper is already running.")
        return

    try:
        app = WindowsHelperApp()
        app.run()
    finally:
        if mutex:
            kernel32.ReleaseMutex(mutex)
            kernel32.CloseHandle(mutex)


if __name__ == "__main__":
    if sys.platform != "win32":
        raise SystemExit("Windows Helper only runs on Windows.")
    main()
