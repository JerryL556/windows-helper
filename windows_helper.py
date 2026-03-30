import ctypes
import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

import pystray
from PIL import Image, ImageDraw


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WM_HOTKEY = 0x0312
WM_SYSCOMMAND = 0x0112
MOD_ALT = 0x0001
VK_P = 0x50
SW_RESTORE = 9
HWND_BROADCAST = 0xFFFF
SC_MONITORPOWER = 0xF170
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
        self.root.geometry("420x390")
        self.root.minsize(420, 390)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        self.hotkey_registered = False
        self.visible = True
        self.status_var = tk.StringVar(value="Visible. Use Alt+P or the tray icon.")
        self.tray_icon = None
        self.settings_window = None

        self._build_ui()
        self._setup_tray_icon()
        self._register_hotkey()

    def _build_ui(self):
        self.root.configure(bg="#f4f6f8")
        self._build_menu()

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
            text="Starts visible. Press Alt+P from anywhere, or use the tray icon, to show or hide it.",
            font=("Segoe UI", 10),
            bg="#f4f6f8",
            fg="#4d5a69",
            wraplength=370,
            justify="left",
        )
        subtitle.pack(anchor="w", pady=(6, 18))

        buttons = [
            ("Open Windows Settings", "start ms-settings:"),
            ("Open Task Manager", "taskmgr"),
            ("Open Downloads Folder", "downloads"),
            ("Open Startup Folder", 'explorer.exe shell:startup'),
            ("Open Command Prompt", "start cmd"),
            ("Show IP Config", "ipconfig"),
            ("Lock And Turn Off Screen", "lock_and_screen_off"),
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

    def _build_menu(self):
        menu_bar = tk.Menu(self.root)

        app_menu = tk.Menu(menu_bar, tearoff=0)
        app_menu.add_command(label="Show", command=self.show_window)
        app_menu.add_command(label="Hide", command=self.hide_window)
        app_menu.add_separator()
        app_menu.add_command(label="Exit", command=self.exit_app)
        menu_bar.add_cascade(label="App", menu=app_menu)

        tools_menu = tk.Menu(menu_bar, tearoff=0)
        tools_menu.add_command(label="Open Windows Settings", command=lambda: self.run_command("start ms-settings:"))
        tools_menu.add_command(label="Open Downloads Folder", command=lambda: self.run_command("downloads"))
        tools_menu.add_command(label="Show IP Config", command=lambda: self.run_command("ipconfig"))
        tools_menu.add_command(label="Lock And Turn Off Screen", command=lambda: self.run_command("lock_and_screen_off"))
        menu_bar.add_cascade(label="Tools", menu=tools_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label="Open Helper Settings", command=self.open_settings_window)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)

        self.root.config(menu=menu_bar)

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

        if command == "downloads":
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            subprocess.Popen(["explorer.exe", downloads_path])
            return

        if command == "lock_and_screen_off":
            user32.LockWorkStation()
            self.root.after(500, lambda: user32.SendMessageW(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, 2))
            return

        subprocess.Popen(command, shell=True)

    def open_settings_window(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.lift()
            self.settings_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        window.title("Helper Settings")
        window.geometry("360x220")
        window.minsize(360, 220)
        window.configure(bg="#f4f6f8")
        window.transient(self.root)
        self.settings_window = window
        window.bind("<Destroy>", self._on_settings_window_destroy, add="+")

        frame = tk.Frame(window, bg="#f4f6f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Helper Settings",
            font=("Segoe UI Semibold", 16),
            bg="#f4f6f8",
            fg="#1d2733",
        ).pack(anchor="w")

        tk.Label(
            frame,
            text="This panel gives you quick access to helper behavior and Windows settings.",
            font=("Segoe UI", 10),
            bg="#f4f6f8",
            fg="#4d5a69",
            wraplength=300,
            justify="left",
        ).pack(anchor="w", pady=(6, 16))

        tk.Button(
            frame,
            text="Open Windows Settings",
            font=("Segoe UI", 10),
            command=lambda: self.run_command("start ms-settings:"),
            padx=12,
        ).pack(fill="x", pady=4)

        tk.Button(
            frame,
            text="Hide To Tray",
            font=("Segoe UI", 10),
            command=self.hide_window,
            padx=12,
        ).pack(fill="x", pady=4)

        tk.Button(
            frame,
            text="Close Settings",
            font=("Segoe UI", 10),
            command=window.destroy,
            padx=12,
        ).pack(fill="x", pady=4)

        window.protocol("WM_DELETE_WINDOW", window.destroy)

    def _on_settings_window_destroy(self, event):
        if event.widget is self.settings_window:
            self.settings_window = None

    def _create_tray_image(self):
        image = Image.new("RGBA", (64, 64), (36, 103, 160, 255))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((6, 6, 58, 58), radius=12, fill=(36, 103, 160, 255))
        draw.rectangle((18, 18, 46, 26), fill=(255, 255, 255, 255))
        draw.rectangle((18, 30, 46, 38), fill=(215, 234, 249, 255))
        draw.rectangle((18, 42, 36, 46), fill=(255, 255, 255, 255))
        return image

    def _setup_tray_icon(self):
        menu = pystray.Menu(
            pystray.MenuItem("Show", self._tray_show_window, default=True),
            pystray.MenuItem("Hide", self._tray_hide_window),
            pystray.MenuItem("Exit", self._tray_exit_app),
        )
        self.tray_icon = pystray.Icon("windows_helper", self._create_tray_image(), "Windows Helper", menu)
        self.tray_icon.run_detached()

    def _tray_show_window(self, icon=None, item=None):
        self.root.after(0, self.show_window)

    def _tray_hide_window(self, icon=None, item=None):
        self.root.after(0, self.hide_window)

    def _tray_exit_app(self, icon=None, item=None):
        self.root.after(0, self.exit_app)

    def _register_hotkey(self):
        if not user32.RegisterHotKey(None, 1, MOD_ALT, VK_P):
            raise RuntimeError("Could not register Alt+P. Another app may already be using it.")

        self.hotkey_registered = True
        self.root.after(50, self._poll_hotkey_messages)

    def _poll_hotkey_messages(self):
        msg = MSG()
        PM_REMOVE = 0x0001
        while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, PM_REMOVE):
            if msg.message == WM_HOTKEY:
                self.toggle_window()
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        if self.hotkey_registered:
            self.root.after(50, self._poll_hotkey_messages)

    def toggle_window(self):
        if self.visible:
            self.hide_window()
        else:
            self.show_window()

    def show_window(self):
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(200, lambda: self.root.attributes("-topmost", False))
        user32.ShowWindow(self.root.winfo_id(), SW_RESTORE)
        user32.SetForegroundWindow(self.root.winfo_id())
        self.visible = True
        self.status_var.set("Visible. Press Alt+P or use the tray icon.")

    def hide_window(self):
        self.root.withdraw()
        self.visible = False
        self.status_var.set("Hidden to tray. Press Alt+P or tray Show.")

    def exit_app(self):
        if self.hotkey_registered:
            user32.UnregisterHotKey(None, 1)
            self.hotkey_registered = False
        if self.tray_icon is not None:
            self.tray_icon.stop()
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
