import ctypes
import ctypes.wintypes
import datetime
import json
from pathlib import Path
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

import pystray
from PIL import Image, ImageDraw


user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

VK_MENU = 0x12
VK_P = 0x50
VK_H = 0x48
VK_W = 0x57
VK_SPACE = 0x20
SW_RESTORE = 9
SWP_NOZORDER = 0x0004
SWP_NOACTIVATE = 0x0010
SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79
MUTEX_NAME = "WindowsHelperSingleInstance"
DOWNLOADS_PATH = r"E:\下载"
CONFIG_PATH = Path(__file__).with_name("windows_helper_config.json")
DEFAULT_CONFIG = {
    "hotkey": "Alt+P",
    "custom_shortcuts": [],
}
HOTKEY_OPTIONS = {
    "Alt+P": (VK_MENU, VK_P),
    "Alt+H": (VK_MENU, VK_H),
    "Alt+W": (VK_MENU, VK_W),
    "Alt+Space": (VK_MENU, VK_SPACE),
}


def create_single_instance_mutex():
    # Use a named mutex so launching the packaged app twice shows one shared instance.
    mutex = kernel32.CreateMutexW(None, True, MUTEX_NAME)
    already_running = kernel32.GetLastError() == 183
    return mutex, already_running


class WindowsHelperApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Windows Helper")
        self.root.geometry("980x720")
        self.root.minsize(760, 620)
        self.root.resizable(True, True)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        self.config = self._load_config()
        self.hotkey_pressed = False
        self.visible = True
        self.hotkey_label_var = tk.StringVar(value=f"Hotkey: {self._get_hotkey_label()}")
        self.status_var = tk.StringVar(value=f"Visible. Use {self._get_hotkey_label()} or the tray icon.")
        self.last_action_var = tk.StringVar(value="Ready")
        self.action_log = []
        self.tray_icon = None
        self.settings_window = None
        self.shortcut_window = None
        self.network_window = None
        self.log_window = None
        self.log_text = None
        self.custom_shortcuts_frame = None
        self.content_canvas = None
        self.content_frame = None
        self.content_window = None
        self.window_listbox = None
        self.window_targets = []
        self.blackout_window = None
        self.blackout_canvas = None
        self.blackout_text_id = None
        self.blackout_animation_job = None
        self.blackout_velocity = (6, 5)

        self._build_ui()
        self._setup_tray_icon()
        self._start_hotkey_monitor()

    def _load_config(self):
        config = {
            "hotkey": DEFAULT_CONFIG["hotkey"],
            "custom_shortcuts": list(DEFAULT_CONFIG["custom_shortcuts"]),
        }
        if not CONFIG_PATH.exists():
            return config

        try:
            saved = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return config

        if isinstance(saved, dict):
            if saved.get("hotkey") in HOTKEY_OPTIONS:
                config["hotkey"] = saved["hotkey"]
            shortcuts = saved.get("custom_shortcuts")
            if isinstance(shortcuts, list):
                config["custom_shortcuts"] = [
                    item
                    for item in shortcuts
                    if isinstance(item, dict) and item.get("label") and item.get("command")
                ]
        return config

    def _save_config(self):
        try:
            CONFIG_PATH.write_text(json.dumps(self.config, indent=2), encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Windows Helper", f"Could not save settings:\n{exc}")

    def _get_hotkey_label(self):
        hotkey = self.config.get("hotkey", DEFAULT_CONFIG["hotkey"])
        if hotkey not in HOTKEY_OPTIONS:
            hotkey = DEFAULT_CONFIG["hotkey"]
        return hotkey

    def _build_ui(self):
        self.root.configure(bg="#eef3f8")
        self._build_menu()

        frame = tk.Frame(self.root, bg="#eef3f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        header = tk.Frame(frame, bg="#1f3b5b", padx=20, pady=18)
        header.grid(row=0, column=0, sticky="ew")

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
            textvariable=self.hotkey_label_var,
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

        content_shell = tk.Frame(frame, bg="#eef3f8")
        content_shell.grid(row=1, column=0, sticky="nsew", pady=(14, 12))
        content_shell.grid_columnconfigure(0, weight=1)
        content_shell.grid_rowconfigure(0, weight=1)

        self.content_canvas = tk.Canvas(
            content_shell,
            bg="#eef3f8",
            highlightthickness=0,
            bd=0,
        )
        scrollbar = tk.Scrollbar(content_shell, orient="vertical", command=self.content_canvas.yview)
        self.content_canvas.configure(yscrollcommand=scrollbar.set)
        self.content_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.content_frame = tk.Frame(self.content_canvas, bg="#eef3f8")
        self.content_window = self.content_canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        self.content_frame.bind("<Configure>", self._on_content_configure)
        self.content_canvas.bind("<Configure>", self._on_canvas_configure)
        self.content_canvas.bind_all("<MouseWheel>", self._on_mousewheel, add="+")

        content = self.content_frame
        content.grid_columnconfigure(0, weight=1, uniform="main")
        content.grid_columnconfigure(1, weight=1, uniform="main")

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
            text="Black Out Screen",
            font=("Segoe UI Semibold", 10),
            bg="#111111",
            fg="#ffffff",
            activebackground="#000000",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=self.open_blackout_screen,
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

        self._build_misc_section(content, 2, 0, 2)
        self._build_window_resize_panel(content, 3, 0, 2)
        self.refresh_window_list()

        footer = tk.Frame(frame, bg="#dfe8f1", padx=16, pady=12)
        footer.grid(row=2, column=0, sticky="ew")

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

    def _build_misc_section(self, parent, row, column, columnspan):
        panel = tk.Frame(parent, bg="#ffffff", bd=1, relief="solid", padx=18, pady=18)
        panel.grid(
            row=row,
            column=column,
            columnspan=columnspan,
            sticky="nsew",
            pady=(10, 0),
        )
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_columnconfigure(1, weight=1)

        tk.Label(
            panel,
            text="Misc",
            font=("Segoe UI Semibold", 14),
            bg="#ffffff",
            fg="#1d2733",
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        tk.Label(
            panel,
            text="Custom shortcuts, hotkey settings, network diagnostics, and helper activity.",
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#5b6978",
            wraplength=620,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 14))

        buttons = [
            ("Manage Custom Shortcuts", self.open_shortcut_manager, True),
            ("Hotkey Settings", self.open_hotkey_settings, False),
            ("Network Diagnostics", self.open_network_diagnostics, False),
            ("Action Log", self.open_action_log, False),
        ]
        for index, (label, command, primary) in enumerate(buttons):
            button_bg = "#245f94" if primary else "#edf3f9"
            button_fg = "#ffffff" if primary else "#18314c"
            active_bg = "#1f527f" if primary else "#dce8f4"
            tk.Button(
                panel,
                text=label,
                font=("Segoe UI Semibold" if primary else "Segoe UI", 10),
                bg=button_bg,
                fg=button_fg,
                activebackground=active_bg,
                activeforeground=button_fg,
                relief="flat",
                bd=0,
                padx=12,
                pady=10,
                command=command,
            ).grid(row=2 + index // 2, column=index % 2, sticky="ew", padx=(0 if index % 2 == 0 else 8, 0), pady=4)

        self.custom_shortcuts_frame = tk.Frame(panel, bg="#ffffff")
        self.custom_shortcuts_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self._refresh_custom_shortcuts_panel()

    def _refresh_custom_shortcuts_panel(self):
        if self.custom_shortcuts_frame is None:
            return

        for widget in self.custom_shortcuts_frame.winfo_children():
            widget.destroy()

        shortcuts = self.config.get("custom_shortcuts", [])
        tk.Label(
            self.custom_shortcuts_frame,
            text="Custom Shortcuts",
            font=("Segoe UI Semibold", 10),
            bg="#ffffff",
            fg="#1d2733",
        ).pack(anchor="w")

        if not shortcuts:
            tk.Label(
                self.custom_shortcuts_frame,
                text="No custom shortcuts yet.",
                font=("Segoe UI", 10),
                bg="#ffffff",
                fg="#5b6978",
            ).pack(anchor="w", pady=(6, 0))
            return

        for index, item in enumerate(shortcuts):
            tk.Button(
                self.custom_shortcuts_frame,
                text=item["label"],
                font=("Segoe UI", 10),
                bg="#edf3f9",
                fg="#18314c",
                activebackground="#dce8f4",
                activeforeground="#18314c",
                relief="flat",
                bd=0,
                anchor="w",
                padx=12,
                pady=8,
                command=lambda item_index=index: self.run_custom_shortcut(item_index),
            ).pack(fill="x", pady=3)

    def _build_window_resize_panel(self, parent, row, column, columnspan):
        panel = tk.Frame(parent, bg="#ffffff", bd=1, relief="solid", padx=18, pady=18)
        panel.grid(
            row=row,
            column=column,
            columnspan=columnspan,
            sticky="nsew",
            pady=(10, 0),
        )

        header = tk.Frame(panel, bg="#ffffff")
        header.pack(fill="x")

        tk.Label(
            header,
            text="Window Resize",
            font=("Segoe UI Semibold", 14),
            bg="#ffffff",
            fg="#1d2733",
        ).pack(side="left")

        tk.Button(
            header,
            text="Refresh Windows",
            font=("Segoe UI", 10),
            bg="#edf3f9",
            fg="#18314c",
            relief="flat",
            bd=0,
            padx=12,
            pady=8,
            command=self.refresh_window_list,
        ).pack(side="right")

        tk.Label(
            panel,
            text="Select an open app window and force it to half the screen width and height, which is one quarter of the display area.",
            font=("Segoe UI", 10),
            bg="#ffffff",
            fg="#5b6978",
            wraplength=620,
            justify="left",
        ).pack(anchor="w", pady=(8, 14))

        list_frame = tk.Frame(panel, bg="#ffffff")
        list_frame.pack(fill="both", expand=True)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        self.window_listbox = tk.Listbox(
            list_frame,
            font=("Segoe UI", 10),
            activestyle="none",
            exportselection=False,
            height=8,
        )
        list_scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.window_listbox.yview)
        self.window_listbox.configure(yscrollcommand=list_scrollbar.set)
        self.window_listbox.grid(row=0, column=0, sticky="nsew")
        list_scrollbar.grid(row=0, column=1, sticky="ns")

        actions = tk.Frame(panel, bg="#ffffff")
        actions.pack(fill="x", pady=(14, 0))

        tk.Button(
            actions,
            text="Resize Selected Window To Quarter Screen",
            font=("Segoe UI Semibold", 10),
            bg="#245f94",
            fg="#ffffff",
            activebackground="#1f527f",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=self.resize_selected_window,
        ).pack(side="left")

        tk.Label(
            actions,
            text="The helper restores the target first, then resizes it.",
            font=("Segoe UI", 9),
            bg="#ffffff",
            fg="#5b6978",
        ).pack(side="right")

    def _on_content_configure(self, event):
        self.content_canvas.configure(scrollregion=self.content_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.content_canvas.itemconfigure(self.content_window, width=event.width)

    def _on_mousewheel(self, event):
        if not self.root.winfo_exists() or self.content_canvas is None:
            return
        widget = event.widget
        while widget is not None:
            if widget is self.content_canvas:
                self.content_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                return
            widget = getattr(widget, "master", None)

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
        menu_bar.add_cascade(label="Tools", menu=tools_menu)

        settings_menu = tk.Menu(menu_bar, tearoff=0)
        settings_menu.add_command(label="Open Helper Settings", command=self.open_settings_window)
        menu_bar.add_cascade(label="Settings", menu=settings_menu)

        self.root.config(menu=menu_bar)

    def refresh_window_list(self):
        windows = self._enumerate_resizeable_windows()
        self.window_targets = windows
        if self.window_listbox is None:
            return

        self.window_listbox.delete(0, tk.END)
        for item in windows:
            self.window_listbox.insert(tk.END, item["label"])

        if windows:
            self.window_listbox.selection_set(0)
            self.last_action_var.set(f"Found {len(windows)} open windows")
        else:
            self.last_action_var.set("No open windows available")

    def _enumerate_resizeable_windows(self):
        windows = []
        helper_hwnd = self.root.winfo_id()

        def callback(hwnd, lparam):
            if hwnd == helper_hwnd or not user32.IsWindowVisible(hwnd):
                return True

            title_length = user32.GetWindowTextLengthW(hwnd)
            if title_length <= 0:
                return True

            title_buffer = ctypes.create_unicode_buffer(title_length + 1)
            user32.GetWindowTextW(hwnd, title_buffer, title_length + 1)
            title = title_buffer.value.strip()
            if not title:
                return True

            rect = ctypes.wintypes.RECT()
            if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                return True

            width = rect.right - rect.left
            height = rect.bottom - rect.top
            if width <= 0 or height <= 0:
                return True

            windows.append(
                {
                    "hwnd": hwnd,
                    "title": title,
                    "label": f"{title} ({width}x{height})",
                }
            )
            return True

        user32.EnumWindows(EnumWindowsProc(callback), 0)
        windows.sort(key=lambda item: item["title"].lower())
        return windows

    def resize_selected_window(self):
        if self.window_listbox is None:
            return

        selection = self.window_listbox.curselection()
        if not selection:
            messagebox.showwarning("Window Resize", "Select a window first.")
            return

        target = self.window_targets[selection[0]]
        hwnd = target["hwnd"]

        if not user32.IsWindow(hwnd):
            messagebox.showwarning(
                "Window Resize",
                "That window is no longer available. Refresh the list and try again.",
            )
            self.refresh_window_list()
            return

        screen_width = user32.GetSystemMetrics(0)
        screen_height = user32.GetSystemMetrics(1)
        target_width = max(320, screen_width // 2)
        target_height = max(240, screen_height // 2)

        rect = ctypes.wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            messagebox.showerror("Window Resize", "Unable to read the selected window position.")
            return

        x = min(max(rect.left, 0), max(screen_width - target_width, 0))
        y = min(max(rect.top, 0), max(screen_height - target_height, 0))

        user32.ShowWindow(hwnd, SW_RESTORE)
        resized = user32.SetWindowPos(
            hwnd,
            None,
            x,
            y,
            target_width,
            target_height,
            SWP_NOZORDER | SWP_NOACTIVATE,
        )

        if not resized:
            messagebox.showerror("Window Resize", "Windows rejected the resize request for that app.")
            self._set_action_status("Quarter-screen resize failed")
            return

        self._set_action_status(f"Resized {target['title']} to quarter screen")
        self.refresh_window_list()

    def open_blackout_screen(self):
        if self.blackout_window is not None and self.blackout_window.winfo_exists():
            self._focus_blackout_screen()
            return

        overlay = tk.Toplevel(self.root, bg="#000000")
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)
        overlay.configure(cursor="none")

        x = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
        y = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
        width = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
        height = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
        overlay.geometry(f"{width}x{height}{x:+d}{y:+d}")

        overlay.bind("<KeyPress>", self.close_blackout_screen)
        overlay.bind("<Button>", lambda event: "break")
        overlay.bind("<ButtonRelease>", lambda event: "break")
        overlay.bind("<Motion>", lambda event: "break")
        overlay.bind("<MouseWheel>", lambda event: "break")
        overlay.bind("<Destroy>", self._on_blackout_destroy, add="+")

        canvas = tk.Canvas(
            overlay,
            bg="#000000",
            highlightthickness=0,
            bd=0,
        )
        canvas.pack(fill="both", expand=True)
        self.blackout_canvas = canvas
        self.blackout_text_id = canvas.create_text(
            width // 2,
            height // 2,
            text="Press anything to exit",
            font=("Segoe UI Semibold", 26),
            fill="#4e4e4e",
            justify="center",
        )

        self.blackout_window = overlay
        self._focus_blackout_screen()
        self._start_blackout_animation()
        self._set_action_status("Screen blackout active")

    def _start_blackout_animation(self):
        if self.blackout_canvas is None or self.blackout_text_id is None:
            return

        if self.blackout_animation_job is not None:
            self.blackout_canvas.after_cancel(self.blackout_animation_job)
            self.blackout_animation_job = None

        self.blackout_velocity = (6, 5)
        self._animate_blackout_text()

    def _animate_blackout_text(self):
        if (
            self.blackout_window is None
            or not self.blackout_window.winfo_exists()
            or self.blackout_canvas is None
            or self.blackout_text_id is None
        ):
            return

        self.blackout_canvas.update_idletasks()
        canvas_width = self.blackout_canvas.winfo_width()
        canvas_height = self.blackout_canvas.winfo_height()
        bounds = self.blackout_canvas.bbox(self.blackout_text_id)
        if bounds is None or canvas_width <= 0 or canvas_height <= 0:
            self.blackout_animation_job = self.blackout_canvas.after(30, self._animate_blackout_text)
            return

        left, top, right, bottom = bounds
        dx, dy = self.blackout_velocity

        if left + dx <= 0 or right + dx >= canvas_width:
            dx *= -1
        if top + dy <= 0 or bottom + dy >= canvas_height:
            dy *= -1

        self.blackout_velocity = (dx, dy)
        self.blackout_canvas.move(self.blackout_text_id, dx, dy)
        self.blackout_animation_job = self.blackout_canvas.after(30, self._animate_blackout_text)

    def _focus_blackout_screen(self):
        if self.blackout_window is None or not self.blackout_window.winfo_exists():
            return

        self.blackout_window.deiconify()
        self.blackout_window.lift()
        self.blackout_window.attributes("-topmost", True)
        self.blackout_window.focus_force()
        self.blackout_window.grab_set()

    def close_blackout_screen(self, event=None):
        if self.blackout_window is None or not self.blackout_window.winfo_exists():
            return

        if self.blackout_animation_job is not None and self.blackout_canvas is not None:
            self.blackout_canvas.after_cancel(self.blackout_animation_job)
            self.blackout_animation_job = None
        try:
            self.blackout_window.grab_release()
        except tk.TclError:
            pass
        self.blackout_window.destroy()

    def _on_blackout_destroy(self, event):
        if event.widget is not self.blackout_window:
            return
        self.blackout_window = None
        self.blackout_canvas = None
        self.blackout_text_id = None
        self.blackout_animation_job = None
        self._set_action_status("Screen blackout cleared")

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

    def run_custom_shortcut(self, index):
        shortcuts = self.config.get("custom_shortcuts", [])
        if index < 0 or index >= len(shortcuts):
            messagebox.showwarning("Custom Shortcut", "That shortcut is no longer available.")
            self._refresh_custom_shortcuts_panel()
            return

        shortcut = shortcuts[index]
        try:
            subprocess.Popen(shortcut["command"], shell=True)
        except OSError as exc:
            messagebox.showerror("Custom Shortcut", f"Could not run shortcut:\n{exc}")
            self._set_action_status("Custom shortcut failed")
            return

        self._set_action_status(f"Ran custom shortcut: {shortcut['label']}")

    def open_shortcut_manager(self):
        if self.shortcut_window is not None and self.shortcut_window.winfo_exists():
            self.shortcut_window.lift()
            self.shortcut_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        window.title("Custom Shortcuts")
        window.geometry("660x440")
        window.minsize(600, 380)
        window.configure(bg="#eef3f8")
        window.transient(self.root)
        self.shortcut_window = window
        window.bind("<Destroy>", self._on_shortcut_window_destroy, add="+")

        frame = tk.Frame(window, bg="#eef3f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        tk.Label(
            frame,
            text="Custom Shortcuts",
            font=("Segoe UI Semibold", 16),
            bg="#eef3f8",
            fg="#1d2733",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        listbox = tk.Listbox(frame, font=("Segoe UI", 10), activestyle="none", exportselection=False)
        listbox.grid(row=1, column=0, sticky="nsew", padx=(0, 12))

        form = tk.Frame(frame, bg="#ffffff", bd=1, relief="solid", padx=16, pady=16)
        form.grid(row=1, column=1, sticky="nsew")
        form.grid_columnconfigure(0, weight=1)

        label_var = tk.StringVar()
        command_var = tk.StringVar()

        tk.Label(form, text="Label", font=("Segoe UI Semibold", 10), bg="#ffffff", fg="#1d2733").grid(
            row=0, column=0, sticky="w"
        )
        tk.Entry(form, textvariable=label_var, font=("Segoe UI", 10)).grid(row=1, column=0, sticky="ew", pady=(4, 12))

        tk.Label(form, text="Command Or Path", font=("Segoe UI Semibold", 10), bg="#ffffff", fg="#1d2733").grid(
            row=2, column=0, sticky="w"
        )
        tk.Entry(form, textvariable=command_var, font=("Segoe UI", 10)).grid(
            row=3, column=0, sticky="ew", pady=(4, 12)
        )

        def refresh_list():
            listbox.delete(0, tk.END)
            for item in self.config.get("custom_shortcuts", []):
                listbox.insert(tk.END, item["label"])

        def load_selected(event=None):
            selection = listbox.curselection()
            if not selection:
                return
            item = self.config["custom_shortcuts"][selection[0]]
            label_var.set(item["label"])
            command_var.set(item["command"])

        def clear_form():
            listbox.selection_clear(0, tk.END)
            label_var.set("")
            command_var.set("")

        def save_shortcut():
            label = label_var.get().strip()
            command = command_var.get().strip()
            if not label or not command:
                messagebox.showwarning("Custom Shortcuts", "Enter both a label and a command.")
                return

            shortcut = {"label": label, "command": command}
            selection = listbox.curselection()
            if selection:
                self.config["custom_shortcuts"][selection[0]] = shortcut
                action = "Updated custom shortcut"
            else:
                self.config["custom_shortcuts"].append(shortcut)
                action = "Added custom shortcut"

            self._save_config()
            refresh_list()
            self._refresh_custom_shortcuts_panel()
            self._set_action_status(f"{action}: {label}")

        def delete_shortcut():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Custom Shortcuts", "Select a shortcut first.")
                return
            item = self.config["custom_shortcuts"].pop(selection[0])
            self._save_config()
            refresh_list()
            clear_form()
            self._refresh_custom_shortcuts_panel()
            self._set_action_status(f"Deleted custom shortcut: {item['label']}")

        def run_selected():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Custom Shortcuts", "Select a shortcut first.")
                return
            self.run_custom_shortcut(selection[0])

        listbox.bind("<<ListboxSelect>>", load_selected)
        refresh_list()

        tk.Button(
            form,
            text="Save Shortcut",
            font=("Segoe UI Semibold", 10),
            bg="#245f94",
            fg="#ffffff",
            activebackground="#1f527f",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=save_shortcut,
        ).grid(row=4, column=0, sticky="ew", pady=4)

        tk.Button(
            form,
            text="New Blank Shortcut",
            font=("Segoe UI", 10),
            bg="#edf3f9",
            fg="#18314c",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=clear_form,
        ).grid(row=5, column=0, sticky="ew", pady=4)

        tk.Button(
            form,
            text="Run Selected",
            font=("Segoe UI", 10),
            bg="#edf3f9",
            fg="#18314c",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=run_selected,
        ).grid(row=6, column=0, sticky="ew", pady=4)

        tk.Button(
            form,
            text="Delete Selected",
            font=("Segoe UI", 10),
            bg="#f7eaea",
            fg="#7a2525",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=delete_shortcut,
        ).grid(row=7, column=0, sticky="ew", pady=(12, 4))

    def _on_shortcut_window_destroy(self, event):
        if event.widget is self.shortcut_window:
            self.shortcut_window = None

    def open_hotkey_settings(self):
        window = tk.Toplevel(self.root)
        window.title("Hotkey Settings")
        window.geometry("360x300")
        window.minsize(340, 280)
        window.configure(bg="#eef3f8")
        window.transient(self.root)

        frame = tk.Frame(window, bg="#eef3f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text="Hotkey Settings",
            font=("Segoe UI Semibold", 16),
            bg="#eef3f8",
            fg="#1d2733",
        ).pack(anchor="w", pady=(0, 12))

        selected_hotkey = tk.StringVar(value=self._get_hotkey_label())
        for label in HOTKEY_OPTIONS:
            tk.Radiobutton(
                frame,
                text=label,
                variable=selected_hotkey,
                value=label,
                font=("Segoe UI", 10),
                bg="#eef3f8",
                fg="#1d2733",
                activebackground="#eef3f8",
                selectcolor="#ffffff",
            ).pack(anchor="w", pady=3)

        def save_hotkey():
            self.config["hotkey"] = selected_hotkey.get()
            self._save_config()
            self.hotkey_pressed = False
            self.hotkey_label_var.set(f"Hotkey: {self._get_hotkey_label()}")
            self._set_action_status(f"Hotkey changed to {self._get_hotkey_label()}")
            window.destroy()

        tk.Button(
            frame,
            text="Save Hotkey",
            font=("Segoe UI Semibold", 10),
            bg="#245f94",
            fg="#ffffff",
            activebackground="#1f527f",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            command=save_hotkey,
        ).pack(fill="x", pady=(18, 0))

    def open_network_diagnostics(self):
        if self.network_window is not None and self.network_window.winfo_exists():
            self.network_window.lift()
            self.network_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        window.title("Network Diagnostics")
        window.geometry("760x520")
        window.minsize(640, 420)
        window.configure(bg="#eef3f8")
        window.transient(self.root)
        self.network_window = window
        window.bind("<Destroy>", self._on_network_window_destroy, add="+")

        frame = tk.Frame(window, bg="#eef3f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        controls = tk.Frame(frame, bg="#eef3f8")
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        output = tk.Text(frame, font=("Consolas", 10), wrap="word", height=18)
        output.grid(row=1, column=0, sticky="nsew")

        def run_network_command(title, args):
            output.delete("1.0", tk.END)
            output.insert(tk.END, f"{title}\n\n")
            try:
                result = subprocess.run(
                    args,
                    capture_output=True,
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    timeout=20,
                    check=False,
                )
            except (OSError, subprocess.TimeoutExpired) as exc:
                output.insert(tk.END, f"Command failed: {exc}")
                self._set_action_status(f"{title} failed")
                return

            text = result.stdout.strip() or result.stderr.strip() or "No output returned."
            output.insert(tk.END, text)
            self._set_action_status(f"Ran network diagnostic: {title}")

        commands = [
            ("IP Config", ["ipconfig", "/all"]),
            ("Ping Localhost", ["ping", "127.0.0.1", "-n", "4"]),
            ("Ping DNS", ["ping", "1.1.1.1", "-n", "4"]),
            ("DNS Cache", ["ipconfig", "/displaydns"]),
        ]
        for label, args in commands:
            tk.Button(
                controls,
                text=label,
                font=("Segoe UI", 10),
                bg="#edf3f9",
                fg="#18314c",
                relief="flat",
                bd=0,
                padx=12,
                pady=8,
                command=lambda title=label, command_args=args: run_network_command(title, command_args),
            ).pack(side="left", padx=(0, 8))

        run_network_command("IP Config", ["ipconfig", "/all"])

    def _on_network_window_destroy(self, event):
        if event.widget is self.network_window:
            self.network_window = None

    def open_action_log(self):
        if self.log_window is not None and self.log_window.winfo_exists():
            self.log_window.lift()
            self.log_window.focus_force()
            self._refresh_action_log_text()
            return

        window = tk.Toplevel(self.root)
        window.title("Action Log")
        window.geometry("560x420")
        window.minsize(480, 340)
        window.configure(bg="#eef3f8")
        window.transient(self.root)
        self.log_window = window
        window.bind("<Destroy>", self._on_log_window_destroy, add="+")

        frame = tk.Frame(window, bg="#eef3f8", padx=18, pady=18)
        frame.pack(fill="both", expand=True)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        header = tk.Frame(frame, bg="#eef3f8")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tk.Label(
            header,
            text="Action Log",
            font=("Segoe UI Semibold", 16),
            bg="#eef3f8",
            fg="#1d2733",
        ).pack(side="left")

        tk.Button(
            header,
            text="Clear",
            font=("Segoe UI", 10),
            bg="#edf3f9",
            fg="#18314c",
            relief="flat",
            bd=0,
            padx=12,
            pady=8,
            command=self._clear_action_log,
        ).pack(side="right")

        self.log_text = tk.Text(frame, font=("Consolas", 10), wrap="word")
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self._refresh_action_log_text()

    def _on_log_window_destroy(self, event):
        if event.widget is self.log_window:
            self.log_window = None

    def _refresh_action_log_text(self):
        if self.log_window is None or not self.log_window.winfo_exists():
            return
        if not hasattr(self, "log_text") or self.log_text is None:
            return

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        entries = self.action_log or ["No actions recorded yet."]
        self.log_text.insert(tk.END, "\n".join(entries))
        self.log_text.configure(state="disabled")

    def _clear_action_log(self):
        self.action_log.clear()
        self._refresh_action_log_text()
        self._set_action_status("Cleared action log")

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
            text=f"{self._get_hotkey_label()} toggles the main window from anywhere.",
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
        modifier_key, action_key = HOTKEY_OPTIONS.get(self._get_hotkey_label(), HOTKEY_OPTIONS["Alt+P"])
        modifier_pressed = bool(user32.GetAsyncKeyState(modifier_key) & 0x8000)
        action_pressed = bool(user32.GetAsyncKeyState(action_key) & 0x8000)
        combo_pressed = modifier_pressed and action_pressed

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
        self._set_action_status("Window restored")

    def hide_window(self):
        self.root.withdraw()
        self.visible = False
        self._set_action_status("Hidden to tray")

    def exit_app(self):
        if self.tray_icon is not None:
            self.tray_icon.stop()
        self.root.destroy()

    def _set_action_status(self, message):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.action_log.append(f"{timestamp}  {message}")
        self.action_log = self.action_log[-50:]
        self.last_action_var.set(message)
        self.status_var.set(f"{message}. Press {self._get_hotkey_label()} or use the tray icon.")
        self._refresh_action_log_text()

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
