import ctypes
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

import pystray
from PIL import Image, ImageDraw


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WM_SYSCOMMAND = 0x0112
VK_MENU = 0x12
VK_P = 0x50
SW_RESTORE = 9
HWND_BROADCAST = 0xFFFF
SC_MONITORPOWER = 0xF170
MUTEX_NAME = "WindowsHelperSingleInstance"
DOWNLOADS_PATH = r"E:\下载"


def create_single_instance_mutex():
    # Use a named mutex so launching the packaged app twice shows one shared instance.
    mutex = kernel32.CreateMutexW(None, True, MUTEX_NAME)
    already_running = kernel32.GetLastError() == 183
    return mutex, already_running


class WindowsHelperApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Windows Helper")
        self.root.geometry("760x560")
        self.root.minsize(700, 520)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        self.hotkey_pressed = False
        self.visible = True
        self.status_var = tk.StringVar(value="Visible. Use Alt+P or the tray icon.")
        self.last_action_var = tk.StringVar(value="Ready")
        self.tray_icon = None
        self.settings_window = None

        self._build_ui()
        self._setup_tray_icon()
        self._start_hotkey_monitor()

    def _build_ui(self):
        self.root.configure(bg="#eef3f8")
        self._build_menu()

        frame = tk.Frame(self.root, bg="#eef3f8", padx=22, pady=20)
        frame.pack(fill="both", expand=True)

        header = tk.Frame(frame, bg="#1f3b5b", padx=20, pady=18)
        header.pack(fill="x")

        tk.Label(
            header,
            text="Windows Helper",
            font=("Segoe UI Semibold", 22),
            bg="#1f3b5b",
            fg="#ffffff",
        ).pack(anchor="w")

        tk.Label(
            header,
            text="Fast access to common Windows tools, folders, and system controls.",
            font=("Segoe UI", 10),
            bg="#1f3b5b",
            fg="#d8e4f0",
            wraplength=620,
            justify="left",
        ).pack(anchor="w", pady=(6, 12))

        meta = tk.Frame(header, bg="#1f3b5b")
        meta.pack(fill="x")

        tk.Label(
            meta,
            text="Hotkey: Alt+P",
            font=("Segoe UI Semibold", 10),
            bg="#dbe8f5",
            fg="#18314c",
            padx=10,
            pady=4,
        ).pack(side="left")

        tk.Label(
            meta,
            textvariable=self.last_action_var,
            font=("Segoe UI", 10),
            bg="#1f3b5b",
            fg="#ffffff",
        ).pack(side="right")

        content = tk.Frame(frame, bg="#eef3f8")
        content.pack(fill="both", expand=True, pady=(18, 14))
        content.grid_columnconfigure(0, weight=1, uniform="main")
        content.grid_columnconfigure(1, weight=1, uniform="main")
        content.grid_rowconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        self._build_section(
            content,
            0,
            0,
            "System",
            "Everyday Windows tools and controls.",
            [
                ("Open Windows Settings", "start ms-settings:", False),
                ("Open Task Manager", "taskmgr", False),
                ("Open Calculator", "calculator", True),
                ("Open Command Prompt", "start cmd", False),
            ],
        )

        self._build_section(
            content,
            0,
            1,
            "Folders And Info",
            "Jump to frequent locations or inspect network details.",
            [
                ("Open Downloads Folder", "downloads", True),
                ("Open Startup Folder", 'explorer.exe shell:startup', False),
                ("Show IP Config", "ipconfig", False),
            ],
        )

        self._build_section(
            content,
            1,
            0,
            "Power And Security",
            "Quick lock actions, separated from the rest of the shortcuts.",
            [
                ("Lock PC", "rundll32.exe user32.dll,LockWorkStation", False),
                ("Lock And Turn Off Screen", "lock_and_screen_off", False),
            ],
        )

        quick_panel = tk.Frame(content, bg="#ffffff", bd=1, relief="solid", padx=18, pady=18)
        quick_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0), pady=(10, 0))

        tk.Label(
            quick_panel,
            text="Helper Controls",
            font=("Segoe UI Semibold", 14),
            bg="#ffffff",
            fg="#1d2733",
        ).pack(anchor="w")

        tk.Label(
            quick_panel,
            text="Manage the app window and open the preferences panel.",
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#5b6978",
            wraplength=280,
            justify="left",
        ).pack(anchor="w", pady=(6, 14))

        tk.Button(
            quick_panel,
            text="Open Helper Settings",
            font=("Segoe UI Semibold", 10),
            bg="#245f94",
            fg="#ffffff",
            activebackground="#1f527f",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=self.open_settings_window,
        ).pack(fill="x", pady=(0, 8))

        tk.Button(
            quick_panel,
            text="Hide To Tray",
            font=("Segoe UI", 10),
            bg="#edf3f9",
            fg="#18314c",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=self.hide_window,
        ).pack(fill="x", pady=(0, 8))

        tk.Button(
            quick_panel,
            text="Exit Helper",
            font=("Segoe UI", 10),
            bg="#f7eaea",
            fg="#7a2525",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=self.exit_app,
        ).pack(fill="x")

        footer = tk.Frame(frame, bg="#dfe8f1", padx=16, pady=12)
        footer.pack(fill="x")

        tk.Label(
            footer,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            bg="#dfe8f1",
            fg="#425364",
        ).pack(side="left")

        tk.Label(
            footer,
            text="Tray ready",
            font=("Segoe UI Semibold", 9),
            bg="#dfe8f1",
            fg="#1f3b5b",
        ).pack(side="right")

    def _build_section(self, parent, row, column, title, description, actions):
        section = tk.Frame(parent, bg="#ffffff", bd=1, relief="solid", padx=18, pady=18)
        section.grid(
            row=row,
            column=column,
            sticky="nsew",
            padx=(0 if column == 0 else 10, 0),
            pady=(0 if row == 0 else 10, 0),
        )

        tk.Label(
            section,
            text=title,
            font=("Segoe UI Semibold", 14),
            bg="#ffffff",
            fg="#1d2733",
        ).pack(anchor="w")

        tk.Label(
            section,
            text=description,
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#5b6978",
            wraplength=280,
            justify="left",
        ).pack(anchor="w", pady=(6, 14))

        for label, command, primary in actions:
            button_bg = "#245f94" if primary else "#edf3f9"
            button_fg = "#ffffff" if primary else "#18314c"
            active_bg = "#1f527f" if primary else "#dce8f4"
            tk.Button(
                section,
                text=label,
                font=("Segoe UI Semibold" if primary else "Segoe UI", 10),
                bg=button_bg,
                fg=button_fg,
                activebackground=active_bg,
                activeforeground=button_fg,
                relief="flat",
                bd=0,
                anchor="w",
                padx=12,
                pady=10,
                command=lambda cmd=command: self.run_command(cmd),
            ).pack(fill="x", pady=4)

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
        tools_menu.add_command(label="Open Calculator", command=lambda: self.run_command("calculator"))
        tools_menu.add_command(label="Open Downloads Folder", command=lambda: self.run_command("downloads"))
        tools_menu.add_command(label="Show IP Config", command=lambda: self.run_command("ipconfig"))
        tools_menu.add_command(label="Lock And Turn Off Screen", command=lambda: self.run_command("lock_and_screen_off"))
        menu_bar.add_cascade(label="Tools", menu=tools_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label="Open Helper Settings", command=self.open_settings_window)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)

        self.root.config(menu=menu_bar)

    def run_command(self, command):
        if command == "calculator":
            subprocess.Popen(["calc.exe"])
            self._set_action_status("Opened Calculator")
            return

        if command == "ipconfig":
            output = subprocess.run(
                ["ipconfig"],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                check=False,
            ).stdout.strip()
            messagebox.showinfo("IP Configuration", output or "No output returned.")
            self._set_action_status("Displayed IP configuration")
            return

        if command == "downloads":
            subprocess.Popen(["explorer.exe", DOWNLOADS_PATH])
            self._set_action_status("Opened Downloads folder")
            return

        if command == "lock_and_screen_off":
            user32.LockWorkStation()
            self.root.after(500, lambda: user32.SendMessageW(HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, 2))
            self._set_action_status("Locked PC and turned off the screen")
            return

        subprocess.Popen(command, shell=True)
        if command == "start ms-settings:":
            self._set_action_status("Opened Windows Settings")
        elif command == "taskmgr":
            self._set_action_status("Opened Task Manager")
        elif command == 'explorer.exe shell:startup':
            self._set_action_status("Opened Startup folder")
        elif command == "start cmd":
            self._set_action_status("Opened Command Prompt")
        elif command == "rundll32.exe user32.dll,LockWorkStation":
            self._set_action_status("Locked PC")
        else:
            self._set_action_status("Ran command")

    def open_settings_window(self):
        if self.settings_window is not None and self.settings_window.winfo_exists():
            self.settings_window.lift()
            self.settings_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        window.title("Helper Settings")
        window.geometry("440x340")
        window.minsize(420, 320)
        window.configure(bg="#eef3f8")
        window.transient(self.root)
        self.settings_window = window
        window.bind("<Destroy>", self._on_settings_window_destroy, add="+")

        frame = tk.Frame(window, bg="#eef3f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)

        hero = tk.Frame(frame, bg="#1f3b5b", padx=18, pady=16)
        hero.pack(fill="x")

        tk.Label(
            hero,
            text="Helper Settings",
            font=("Segoe UI Semibold", 16),
            bg="#1f3b5b",
            fg="#ffffff",
        ).pack(anchor="w")

        tk.Label(
            hero,
            text="Quick controls for startup behavior, tray usage, and common helper actions.",
            font=("Segoe UI", 10),
            bg="#1f3b5b",
            fg="#d8e4f0",
            wraplength=360,
            justify="left",
        ).pack(anchor="w", pady=(6, 16))

        body = tk.Frame(frame, bg="#ffffff", bd=1, relief="solid", padx=18, pady=18)
        body.pack(fill="both", expand=True, pady=(14, 0))

        tk.Button(
            body,
            text="Open Windows Settings",
            font=("Segoe UI Semibold", 10),
            bg="#245f94",
            fg="#ffffff",
            activebackground="#1f527f",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=lambda: self.run_command("start ms-settings:"),
        ).pack(fill="x", pady=4)

        info_row = tk.Frame(body, bg="#ffffff")
        info_row.pack(fill="x", pady=(12, 10))

        tk.Label(
            info_row,
            text="Hotkey",
            font=("Segoe UI Semibold", 10),
            bg="#ffffff",
            fg="#1d2733",
        ).grid(row=0, column=0, sticky="w")

        tk.Label(
            info_row,
            text="Alt+P toggles the main window from anywhere.",
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#5b6978",
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))

        tk.Label(
            info_row,
            text="Window State",
            font=("Segoe UI Semibold", 10),
            bg="#ffffff",
            fg="#1d2733",
        ).grid(row=1, column=0, sticky="w", pady=(8, 0))

        tk.Label(
            info_row,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#5b6978",
            wraplength=240,
            justify="left",
        ).grid(row=1, column=1, sticky="w", padx=(12, 0), pady=(8, 0))

        tk.Button(
            body,
            text="Hide To Tray",
            font=("Segoe UI", 10),
            bg="#edf3f9",
            fg="#18314c",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=self.hide_window,
        ).pack(fill="x", pady=4)

        tk.Button(
            body,
            text="Close Settings",
            font=("Segoe UI", 10),
            bg="#edf0f5",
            fg="#18314c",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=window.destroy,
        ).pack(fill="x", pady=(10, 0))

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
        # Run the tray icon on its own thread so Tk's mainloop stays responsive.
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

    def _start_hotkey_monitor(self):
        # Polling avoids extra global hook dependencies for this lightweight helper.
        self.root.after(50, self._poll_hotkey_state)

    def _poll_hotkey_state(self):
        alt_pressed = bool(user32.GetAsyncKeyState(VK_MENU) & 0x8000)
        p_pressed = bool(user32.GetAsyncKeyState(VK_P) & 0x8000)
        combo_pressed = alt_pressed and p_pressed

        if combo_pressed and not self.hotkey_pressed:
            self.toggle_window()

        self.hotkey_pressed = combo_pressed
        self.root.after(50, self._poll_hotkey_state)

    def toggle_window(self):
        if self.visible:
            self.hide_window()
        else:
            self.show_window()

    def show_window(self):
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        # Briefly force topmost so the restored window reliably comes to the front.
        self.root.attributes("-topmost", True)
        self.root.after(200, lambda: self.root.attributes("-topmost", False))
        user32.ShowWindow(self.root.winfo_id(), SW_RESTORE)
        user32.SetForegroundWindow(self.root.winfo_id())
        self.visible = True
        self.status_var.set("Visible. Press Alt+P or use the tray icon.")
        self.last_action_var.set("Window restored")

    def hide_window(self):
        self.root.withdraw()
        self.visible = False
        self.status_var.set("Hidden to tray. Press Alt+P or tray Show.")
        self.last_action_var.set("Hidden to tray")

    def exit_app(self):
        if self.tray_icon is not None:
            self.tray_icon.stop()
        self.root.destroy()

    def _set_action_status(self, message):
        self.last_action_var.set(message)
        self.status_var.set(f"{message}. Press Alt+P or use the tray icon.")

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
