# falconUI_Tool.py
# 03/18 2025
import io
import os
import queue
import subprocess
import sys
from pathlib import Path
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# 確保控制台輸出使用 UTF-8
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

class FalconUIScriptBuilder:
    def __init__(self, root):
        self.root = root
        self.version = "1.0.34"
        self.root.title("FalconUI Script Builder")
        self.root.geometry("1200x700")
        self.root.minsize(900, 600)
        self.stop_on_error_var = tk.BooleanVar(value=True)
        # Initialize coord window reference
        self.coord_window = None
        self._drag_data = {}
        self.active_scrollable_canvas = None
        # Check if falconui.exe exists
        self.falconui_path = "falconCommand.exe"
        if not os.path.exists(self.falconui_path) and getattr(sys, "frozen", False):
            self.falconui_path = os.path.join(
                os.path.dirname(sys.executable), "falconCommand.exe"
            )

        # Set icon (if available)
        try:
            self.root.iconbitmap("falcon_icon.ico")
        except:
            pass

        # Initialize clipboard functionality
        self.clipboard = None

        # Set modern theme colors
        self.style = ttk.Style()

        # Main background color
        self.bg_color = "#f5f5f7"
        self.accent_color = "#007AFF"  # Blue accent color
        self.secondary_color = "#5AC8FA"  # Light blue
        self.success_color = "#34C759"  # Green
        self.warning_color = "#FF9500"  # Orange
        self.error_color = "#FF3B30"  # Red

        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure(
            "TButton",
            padding=(2, 2),
            relief="flat",
            background=self.accent_color,
            foreground="black",
            font=("Segoe UI", 9),
        )

        self.style.configure(
            "Secondary.TButton",
            padding=(2, 2),
            background="#f0f0f0",
            foreground="#333333",
        )

        self.style.configure(
            "Success.TButton",
            padding=(2, 2),
            background=self.success_color,
            foreground="black",
        )

        self.style.configure(
            "Warning.TButton",
            padding=(2, 2),
            background=self.warning_color,
            foreground="black",
        )

        self.style.configure(
            "Danger.TButton",
            padding=(2, 2),
            background=self.error_color,
            foreground="black",
        )

        self.style.configure(
            "Nav.TButton", padding=(2, 2), font=("Segoe UI", 9, "bold")
        )

        self.style.configure(
            "Command.TFrame", background="white", relief="solid", borderwidth=1
        )

        self.style.configure("TNotebook", background=self.bg_color, borderwidth=0)

        self.style.configure(
            "TNotebook.Tab",
            padding=[2, 2],
            font=("Segoe UI", 9, "bold"),
            background="#e0e0e0",
            foreground="#333333",
        )

        self.style.map(
            "TNotebook.Tab",
            background=[("selected", self.accent_color)],
            foreground=[("selected", "blue")],
        )

        self.style.map(
            "TButton",
            background=[("active", self.secondary_color)],
            foreground=[("active", "blue")],
        )

        # Initialize process and queue for execution logs
        self.current_process = None
        self.log_queue = queue.Queue()

        # Initialize autosave and backup settings
        self.autosave_interval = 5 * 60 * 1000  # 5 minutes in milliseconds
        self.backup_count = 5  # Number of backups to keep
        self.autosave_enabled = True
        self.backup_dir = os.path.join(os.path.expanduser("~"), ".falconui_backups")

        # Create backup directory if it doesn't exist
        if not os.path.exists(self.backup_dir):
            try:
                os.makedirs(self.backup_dir)
            except:
                self.backup_dir = os.path.dirname(os.path.abspath(__file__))

        # Create main layout with a slightly larger padding
        self.main_frame = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        # Left command panel
        self.command_frame = ttk.Frame(self.main_frame)
        self.main_frame.add(self.command_frame, weight=1)

        # Right editor
        self.editor_frame = ttk.Frame(self.main_frame)
        self.main_frame.add(self.editor_frame, weight=2)
        self.create_app_menu()
        self.create_system_menu()
        # Initialize command categories
        self.create_command_panel()

        # Initialize script editor
        self.create_script_editor()

        # Initialize status bar
        self.create_statusbar()

        # Initialize current script name
        self.current_script = "new_script.txt"
        self.is_script_modified = False
        self.update_title()

        # # Initialize debug state variables
        # self.debug_mode = False
        # self.current_debug_line = 0
        # self.debug_lines = []
        # self.debug_line_numbers = []
        # self.debug_step_event = threading.Event()
        # self.debug_skip_event = threading.Event()
        # self.debug_stop_event = threading.Event()

        # Start autosave timer
        self.schedule_autosave()

        # Bind close event
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_system_menu(self):
        """Create Windows-style system menu with standard keyboard shortcuts"""
        # Create system menu
        self.system_menu = tk.Menu(self.root, tearoff=0)
        self.system_menu.add_command(
            label="Restore(&R)", command=self.restore_window, underline=3
        )  # R - Restore
        self.system_menu.add_command(
            label="Move(&M)", command=self.move_window, underline=3
        )  # M - Move
        self.system_menu.add_command(
            label="Size(&S)", command=self.resize_window, underline=3
        )  # S - Size
        self.system_menu.add_command(
            label="Minimize(&N)", command=self.minimize_window, underline=4
        )  # N - Minimize
        self.system_menu.add_command(
            label="Maximize(&X)", command=self.maximize_window, underline=4
        )  # X - Maximize
        self.system_menu.add_separator()
        self.system_menu.add_command(
            label="Close(&C)", command=self.on_close, underline=3
        )  # C - Close

        # Bind Alt+Space shortcut
        self.root.bind("<Alt-space>", self.show_system_menu)

        # Bind standard Windows shortcuts for window operations
        self.root.bind("<Alt-F4>", lambda e: self.on_close())
        self.root.bind("<Alt-r>", lambda e: self.restore_window())
        self.root.bind("<Alt-m>", lambda e: self.move_window())
        self.root.bind("<Alt-s>", lambda e: self.resize_window())
        self.root.bind("<Alt-n>", lambda e: self.minimize_window())
        self.root.bind("<Alt-x>", lambda e: self.maximize_window())
        self.root.bind("<Alt-c>", lambda e: self.on_close())

    def show_system_menu(self, event=None):
        """Display the system menu"""
        # Get current window state to update menu item enable/disable status
        is_zoomed = self.root.wm_state() == "zoomed"
        is_iconic = self.root.wm_state() == "iconic"

        # Enable or disable menu items based on window state
        self.system_menu.entryconfig(
            "Restore(&R)", state="normal" if (is_zoomed or is_iconic) else "disabled"
        )
        self.system_menu.entryconfig(
            "Move(&M)", state="disabled" if (is_zoomed or is_iconic) else "normal"
        )
        self.system_menu.entryconfig(
            "Size(&S)", state="disabled" if (is_zoomed or is_iconic) else "normal"
        )
        self.system_menu.entryconfig(
            "Maximize(&X)", state="disabled" if is_zoomed else "normal"
        )
        self.system_menu.entryconfig(
            "Minimize(&N)", state="disabled" if is_iconic else "normal"
        )

        # Display menu at mouse position
        try:
            # If called by Alt+Space, display at window top left corner
            if event:
                x = self.root.winfo_rootx() + 5
                y = self.root.winfo_rooty() + 5
                self.system_menu.tk_popup(x, y)
            else:
                # Otherwise show at mouse position
                self.system_menu.tk_popup(
                    self.root.winfo_pointerx(), self.root.winfo_pointery()
                )
        finally:
            # Ensure menu is properly closed
            self.system_menu.grab_release()

    def restore_window(self):
        """Restore the window"""
        if self.root.wm_state() == "iconic":
            self.root.wm_deiconify()
        else:
            self.root.wm_state("normal")

    def move_window(self):
        """Move the window"""
        if self.root.wm_state() == "zoomed" or self.root.wm_state() == "iconic":
            return  # Cannot move in maximized or minimized state

        self.root.wm_attributes("-alpha", 0.8)  # Semi-transparent to help see position
        x, y = self.root.winfo_x(), self.root.winfo_y()

        # Create movement instructions
        move_label = tk.Label(
            self.root,
            text="Use arrow keys to move window, press Enter to confirm",
            font=("Segoe UI", 12, "bold"),
            fg="#FFFFFF",
            bg="#444444",
            relief="solid",
            bd=1,
            padx=10,
            pady=5,
        )
        move_label.place(relx=0.5, rely=0.5, anchor="center")
        move_label.focus_set()

        # Function to handle keyboard events
        def key_handler(e):
            nonlocal x, y
            key = e.keysym
            if key == "Left":
                x -= 10
            elif key == "Right":
                x += 10
            elif key == "Up":
                y -= 10
            elif key == "Down":
                y += 10
            elif key == "Return":  # Enter key to confirm
                move_label.destroy()
                self.root.wm_attributes("-alpha", 1.0)  # Restore transparency
                return "break"
            elif key == "Escape":  # ESC to cancel
                move_label.destroy()
                self.root.wm_attributes("-alpha", 1.0)  # Restore transparency
                return "break"

            self.root.geometry(f"+{x}+{y}")
            return "break"

        # Bind keyboard events to move_label
        move_label.bind("<Key>", key_handler)

    def resize_window(self):
        """Resize the window"""
        if self.root.wm_state() == "zoomed" or self.root.wm_state() == "iconic":
            return  # Cannot resize in maximized or minimized state

        self.root.wm_attributes("-alpha", 0.8)  # Semi-transparent to help see size
        width, height = self.root.winfo_width(), self.root.winfo_height()

        # Create resize instructions
        resize_label = tk.Label(
            self.root,
            text="Use arrow keys to resize window, press Enter to confirm",
            font=("Segoe UI", 12, "bold"),
            fg="#FFFFFF",
            bg="#444444",
            relief="solid",
            bd=1,
            padx=10,
            pady=5,
        )
        resize_label.place(relx=0.5, rely=0.5, anchor="center")
        resize_label.focus_set()

        # Function to handle keyboard events
        def key_handler(e):
            nonlocal width, height
            key = e.keysym
            if key == "Left":
                width -= 10
            elif key == "Right":
                width += 10
            elif key == "Up":
                height -= 10
            elif key == "Down":
                height += 10
            elif key == "Return":  # Enter key to confirm
                resize_label.destroy()
                self.root.wm_attributes("-alpha", 1.0)  # Restore transparency
                return "break"
            elif key == "Escape":  # ESC to cancel
                resize_label.destroy()
                self.root.wm_attributes("-alpha", 1.0)  # Restore transparency
                return "break"

            # Ensure window size is not smaller than minimum values
            width = max(width, 400)
            height = max(height, 300)

            self.root.geometry(f"{width}x{height}")
            return "break"

        # Bind keyboard events to resize_label
        resize_label.bind("<Key>", key_handler)

    def minimize_window(self):
        """Minimize the window"""
        self.root.wm_iconify()

    def maximize_window(self):
        """Maximize the window"""
        current_state = self.root.wm_state()
        if current_state == "zoomed":
            self.root.wm_state("normal")
        else:
            self.root.wm_state("zoomed")

    def _bind_to_mousewheel(self, canvas):
        """Bind mousewheel events to the canvas for scrolling"""
        # Set this canvas as the active scrollable one
        self.active_scrollable_canvas = canvas

        # Bind mousewheel events to the root window - one binding for all canvases
        if sys.platform.startswith("win"):
            self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        elif sys.platform.startswith("darwin"):  # macOS
            self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        else:  # Linux
            self.root.bind_all("<Button-4>", self._on_mousewheel)
            self.root.bind_all("<Button-5>", self._on_mousewheel)

    def _unbind_from_mousewheel(self, canvas):
        """Unbind mousewheel events from the canvas"""
        # Only unbind if this canvas is the active one
        if self.active_scrollable_canvas == canvas:
            # Clear the active canvas reference
            self.active_scrollable_canvas = None

            # Unbind all mousewheel events
            if sys.platform.startswith("win") or sys.platform.startswith("darwin"):
                self.root.unbind_all("<MouseWheel>")
            else:  # Linux
                self.root.unbind_all("<Button-4>")
                self.root.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        """Handle mousewheel scrolling"""
        # Only scroll if we have an active scrollable canvas
        if not self.active_scrollable_canvas:
            return

        # Get the canvas to scroll
        canvas = self.active_scrollable_canvas

        # Determine scroll direction
        if sys.platform.startswith("win"):
            delta = -1 * (event.delta // 120)
        elif sys.platform.startswith("darwin"):  # macOS
            delta = -1 * event.delta
        else:  # Linux
            if event.num == 4:
                delta = -1
            else:
                delta = 1

        # Perform the scroll
        canvas.yview_scroll(int(delta), "units")

        # Prevent event propagation
        return "break"

    def update_scrollregion(self, event, canvas):
        """Update the scroll region and check if scrolling is needed"""
        # Configure the scroll region based on the size of the contained frame
        canvas.configure(scrollregion=canvas.bbox("all"))

        # Check if scrolling is needed
        content_height = canvas.bbox("all")[3] if canvas.bbox("all") else 0
        visible_height = canvas.winfo_height()

        # Get the scrollbar for this canvas
        category_name = getattr(canvas, "category_name", "unknown")
        if (
            hasattr(self, "category_scrollbars")
            and category_name in self.category_scrollbars
        ):
            _, scrollbar = self.category_scrollbars[category_name]

            # Enable or disable scrolling based on content size
            if content_height > visible_height:
                # Content is taller than visible area - enable scrolling
                canvas.bind("<Enter>", lambda e, c=canvas: self._bind_to_mousewheel(c))
                canvas.bind(
                    "<Leave>", lambda e, c=canvas: self._unbind_from_mousewheel(c)
                )
                scrollbar.pack(side="right", fill="y")
            else:
                # Content fits in visible area - disable scrolling
                canvas.unbind("<Enter>")
                canvas.unbind("<Leave>")

                # If this is currently the active scrollable canvas, clear it
                if (
                    hasattr(self, "active_scrollable_canvas")
                    and self.active_scrollable_canvas == canvas
                ):
                    self.active_scrollable_canvas = None

                # Hide scrollbar
                scrollbar.pack_forget()

    def create_command_panel(self):
        # Command category title with improved styling
        header_frame = ttk.Frame(self.command_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        header_label = ttk.Label(
            header_frame,
            text="Command Categories",
            font=("Segoe UI", 12, "bold"),
            foreground="#333333",
            background=self.bg_color,
        )
        header_label.pack(side=tk.LEFT, padx=5)

        # Create category notebook
        self.command_notebook = ttk.Notebook(self.command_frame)
        self.command_notebook.pack(fill=tk.BOTH, expand=True)

        # Define command categories
        self.command_categories = {
            "Mouse Actions": [
                {
                    "name": "--click",
                    "description": "Click at specified coordinates",
                    "params": ["X(optional)", "Y(optional)"],
                },
                {
                    "name": "--double-click",
                    "description": "Double-click at coordinates",
                    "params": ["X(optional)", "Y(optional)"],
                },
                {
                    "name": "--right-click",
                    "description": "Right-click at coordinates",
                    "params": ["X(optional)", "Y(optional)"],
                },
                {
                    "name": "--fast-click",
                    "description": "Fast click multiple times",
                    "params": ["X(optional)", "Y(optional)", "COUNT", "DELAY"],
                },
                {
                    "name": "--moveto",
                    "description": "Move to coordinates",
                    "params": ["X", "Y", "DURATION(optional)"],
                },
                {
                    "name": "--relative-move",
                    "description": "Move relative to current position",
                    "params": ["X_OFFSET", "Y_OFFSET"],
                },
                {
                    "name": "--center-on-screen",
                    "description": "Move to screen center",
                    "params": [],
                },
                {
                    "name": "--scroll",
                    "description": "Scroll mouse wheel",
                    "params": ["AMOUNT", "X(optional)", "Y(optional)"],
                },
                {
                    "name": "--drag-to",
                    "description": "Drag to specified position",
                    "params": ["X", "Y", "DURATION(optional)"],
                },
                {
                    "name": "--position",
                    "description": "Get current coordinates",
                    "params": [],
                },
                {
                    "name": "--track-position",
                    "description": "Track mouse position",
                    "params": ["DURATION"],
                },
                {
                    "name": "--position-to-clipboard",
                    "description": "Copy mouse position to clipboard",
                    "params": [],
                },
            ],
            "Keyboard Input": [
                {"name": "--type", "description": "Type text", "params": ["TEXT"]},
                {"name": "--press", "description": "Press a key", "params": ["KEY"]},
                {
                    "name": "--clipboard-copy",
                    "description": "Copy selection to clipboard",
                    "params": [],
                },
                {
                    "name": "--clipboard-paste",
                    "description": "Paste from clipboard",
                    "params": [],
                },
                {
                    "name": "--clipboard-set",
                    "description": "Set clipboard content",
                    "params": ["TEXT"],
                },
                {
                    "name": "--clipboard-get",
                    "description": "Get clipboard content",
                    "params": [],
                },
            ],
            "System Functions": [
                {
                    "name": "--sleep",
                    "description": "Pause execution",
                    "params": ["SECONDS"],
                },
                {
                    "name": "--screenshot",
                    "description": "Take screenshot",
                    "params": ["FILENAME"],
                },
                {
                    "name": "--screen-size",
                    "description": "Get screen size",
                    "params": [],
                },
                {
                    "name": "--window-info",
                    "description": "Get window information",
                    "params": ["TITLE"],
                },
                {
                    "name": "--launch",
                    "description": "Launch application",
                    "params": ["APP_PATH"],
                },
                # Add the following entry to the "System Functions" list in the create_command_panel method
                {
                    "name": "--check-software",
                    "description": "Check if specified software is installed",
                    "params": ["SOFTWARE_NAME"],
                },
                {
                    "name": "--wait-until-installed",
                    "description": "Wait until software is installed",
                    "params": ["SOFTWARE_NAME", "--wait-time"],
                },
                {
                    "name": "--wait-until-process",
                    "description": "Wait until specific process is running",
                    "params": ["PROCESS_NAME","--wait-time"],
                },
                {
                    "name": "--wait-until-exist",
                    "description": "Wait until image/file/folder exists",
                    "params": ["PATH","--wait-time"],
                },
            ],
            "Image Recognition": [
                {
                    "name": "--click-image",
                    "description": "Click on image",
                    "params": ["IMAGE_PATH"],
                },
                {
                    "name": "--search-image",
                    "description": "Search image",
                    "params": ["IMAGE_PATH"],
                },
                {
                    "name": "--double-click-image",
                    "description": "Double-click on image",
                    "params": ["IMAGE_PATH"],
                },
                {
                    "name": "--right-click-image",
                    "description": "Right-click on image",
                    "params": ["IMAGE_PATH"],
                },
                {
                    "name": "--image-exists",
                    "description": "Check if image exists",
                    "params": ["IMAGE_PATH"],
                },
            ],
            "Advanced Features": [
                {
                    "name": "--run",
                    "description": "Execute command sequence",
                    "params": ["COMMAND (without --)", "COMMAND_ARGS", "--repeat N"],                    
                },
                {
                    "name": "--command-file",
                    "description": "Execute command file",
                    "params": ["FILE_PATH"],
                },
            ],
        }

        # Create pages for each category with improved styling
        for category_name, commands in self.command_categories.items():
            category_frame = ttk.Frame(self.command_notebook)
            self.command_notebook.add(category_frame, text=category_name)

            # Create scrollable area with improved styling
            canvas = tk.Canvas(
                category_frame, background="#ffffff", highlightthickness=0
            )
            scrollbar = ttk.Scrollbar(
                category_frame, orient="vertical", command=canvas.yview
            )
            scrollable_frame = ttk.Frame(canvas, style="TFrame")

            scrollable_frame.bind(
                "<Configure>",
                lambda e, canvas=canvas: self.update_scrollregion(e, canvas),
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            scrollbar.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)

            # Add command buttons with improved styling
            for i, command in enumerate(commands):
                cmd_frame = ttk.Frame(scrollable_frame, style="Command.TFrame")
                cmd_frame.pack(fill=tk.X, padx=8, pady=4, ipady=3)

                # Command name with better font and styling
                name_label = ttk.Label(
                    cmd_frame,
                    text=command["name"],
                    font=("Consolas", 10, "bold"),
                    foreground=self.accent_color,
                    background="white",
                )
                name_label.pack(anchor=tk.W, padx=10, pady=(5, 0))

                # Command description with improved styling
                desc_label = ttk.Label(
                    cmd_frame,
                    text=command["description"],
                    background="white",
                    foreground="#333333",
                    font=("Segoe UI", 9),
                )
                desc_label.pack(anchor=tk.W, padx=10)

                # Parameter list with improved styling
                if command["params"]:
                    params_text = "Parameters: " + ", ".join(command["params"])
                    params_label = ttk.Label(
                        cmd_frame,
                        text=params_text,
                        font=("Segoe UI", 8),
                        foreground="#777777",
                        background="white",
                    )
                    params_label.pack(anchor=tk.W, padx=10, pady=(0, 5))

                # Add button with improved styling
                add_btn = ttk.Button(
                    cmd_frame,
                    text="Add",
                    style="Secondary.TButton",
                    command=lambda cmd=command: self.add_command_to_script(cmd),
                )
                add_btn.pack(anchor=tk.E, padx=10, pady=3)

                # Bind click event for the entire frame using the helper handler
                cmd_frame.bind(
                    "<Button-1>",
                    lambda e, cmd=command: self.command_click_handler(e, cmd),
                )
                name_label.bind(
                    "<Button-1>",
                    lambda e, cmd=command: self.command_click_handler(e, cmd),
                )
                desc_label.bind(
                    "<Button-1>",
                    lambda e, cmd=command: self.command_click_handler(e, cmd),
                )
                if command["params"]:
                    params_label.bind(
                        "<Button-1>",
                        lambda e, cmd=command: self.command_click_handler(e, cmd),
                    )
                # Store canvas and scrollbar for this category (to enable/disable later)
                canvas.category_name = (
                    category_name  # Store category name for debugging
                )
                if not hasattr(self, "category_scrollbars"):
                    self.category_scrollbars = {}
                self.category_scrollbars[category_name] = (canvas, scrollbar)

    def command_click_handler(self, event, command):
        """Handle click event on a command and prevent further propagation"""
        self.add_command_to_script(command)
        return "break"

    def create_script_editor(self):
        # Script name and toolbar with improved styling
        toolbox_frame = ttk.Frame(self.editor_frame)
        toolbox_frame.pack(fill=tk.X, pady=(0, 15))

        script_name_frame = ttk.Frame(toolbox_frame)
        script_name_frame.pack(side=tk.LEFT, fill=tk.X, padx=5)

        script_name_label = ttk.Label(
            script_name_frame,
            text="Script Name:",
            font=("Segoe UI", 10, "bold"),
            foreground="#333333",
            background=self.bg_color,
        )
        script_name_label.pack(side=tk.LEFT, padx=5)

        self.script_name_var = tk.StringVar(value="new_script.txt")
        self.script_name_entry = ttk.Entry(
            script_name_frame, textvariable=self.script_name_var, font=("Segoe UI", 10)
        )
        self.script_name_entry.pack(side=tk.LEFT, fill=tk.X, padx=5)
        self.script_name_entry.bind("<KeyRelease>", lambda e: self.on_script_modified())

        # Add stop-on-error checkbox to the toolbox_frame

        stop_error_check = ttk.Checkbutton(
            script_name_frame,
            text="Stop on Error",
            variable=self.stop_on_error_var,
            command=self.update_stop_on_error_status,
        )
        stop_error_check.pack(side=tk.LEFT, padx=10)

        # Add coordinate display variables
        self.coord_x = tk.StringVar(value="X: 0")
        self.coord_y = tk.StringVar(value="Y: 0")

        # Add a frame for the coordinate display
        coord_frame = ttk.Frame(script_name_frame)
        coord_frame.pack(side=tk.LEFT, padx=5)

        # Show live coordinates button
        self.tracker_active = False
        self.track_button = ttk.Button(
            coord_frame,
            text="Show Live XY",
            command=self.toggle_coordinate_tracker,
            style="Secondary.TButton",
        )
        self.track_button.pack(side=tk.LEFT, padx=5)

        # Labels to display coordinates
        self.x_label = ttk.Label(
            coord_frame,
            textvariable=self.coord_x,
            font=("Consolas", 10, "bold"),
            background=self.bg_color,
        )
        self.x_label.pack(side=tk.LEFT, padx=3)

        self.y_label = ttk.Label(
            coord_frame,
            textvariable=self.coord_y,
            font=("Consolas", 10, "bold"),
            background=self.bg_color,
        )
        self.y_label.pack(side=tk.LEFT, padx=3)

        # Toolbar buttons with improved styling
        #tools_frame = ttk.Frame(toolbox_frame)
        #tools_frame.pack(side=tk.RIGHT, padx=5)
        tools_frame = ttk.Frame(self.editor_frame)
        tools_frame.pack(fill=tk.X, pady=5)
        
        # New button
        new_btn = ttk.Button(tools_frame, text="New", command=self.new_script)
        new_btn.pack(side=tk.LEFT, padx=3)

        # Open button
        open_btn = ttk.Button(tools_frame, text="Open", command=self.open_script)
        open_btn.pack(side=tk.LEFT, padx=3)

        # Save button
        save_btn = ttk.Button(
            tools_frame, text="Save", command=self.save_script, style="Success.TButton"
        )
        save_btn.pack(side=tk.LEFT, padx=3)

        # Validate button
        validate_btn = ttk.Button(
            tools_frame,
            text="Validate",
            command=self.validate_script,
            style="Secondary.TButton",
        )
        validate_btn.pack(side=tk.LEFT, padx=3)

        # Run button
        run_btn = ttk.Button(
            tools_frame, text="Run", command=self.run_script, style="Success.TButton"
        )
        run_btn.pack(side=tk.LEFT, padx=3)

        # Stop button (for stopping execution)
        self.stop_btn = ttk.Button(
            tools_frame,
            text="Stop",
            command=self.stop_execution,
            state=tk.DISABLED,
            style="Danger.TButton",
        )
        self.stop_btn.pack(side=tk.LEFT, padx=3)

        # Clear button
        clear_btn = ttk.Button(
            tools_frame,
            text="Clear",
            command=self.clear_script,
            style="Secondary.TButton",
        )
        clear_btn.pack(side=tk.LEFT, padx=3)

        # # Add debug tools frame
        # debug_frame = ttk.Frame(self.editor_frame)
        # debug_frame.pack(fill=tk.X, pady=5)

        # debug_label = ttk.Label(
        #     debug_frame, text="Debug Tools:", font=("Segoe UI", 9, "bold")
        # )
        # debug_label.pack(side=tk.LEFT, padx=5)

        # # Step-by-step execution button
        # self.step_btn = ttk.Button(
        #     debug_frame,
        #     text="Step Mode",
        #     command=self.run_step_by_step,
        #     style="Secondary.TButton",
        # )
        # self.step_btn.pack(side=tk.LEFT, padx=3)

        # # Next step button (initially disabled)
        # self.next_step_btn = ttk.Button(
        #     debug_frame,
        #     text="Next Step",
        #     command=self.execute_next_step,
        #     state=tk.DISABLED,
        # )
        # self.next_step_btn.pack(side=tk.LEFT, padx=3)

        # # Skip step button
        # self.skip_step_btn = ttk.Button(
        #     debug_frame,
        #     text="Skip Step",
        #     command=self.skip_current_step,
        #     state=tk.DISABLED,
        # )
        # self.skip_step_btn.pack(side=tk.LEFT, padx=3)

        # # Stop debug button
        # self.stop_debug_btn = ttk.Button(
        #     debug_frame,
        #     text="Stop Debug",
        #     command=self.stop_debug,
        #     state=tk.DISABLED,
        #     style="Danger.TButton",
        # )
        # self.stop_debug_btn.pack(side=tk.LEFT, padx=3)

        # # Line indicator in debug mode
        # self.current_line_var = tk.StringVar(value="")
        # current_line_label = ttk.Label(
        #     debug_frame, textvariable=self.current_line_var, font=("Consolas", 9)
        # )
        # current_line_label.pack(side=tk.RIGHT, padx=5)

        # Create a PanedWindow to divide the editor area vertically
        editor_paned = ttk.PanedWindow(self.editor_frame, orient=tk.VERTICAL)
        editor_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=3)

        # Script editor container with improved styling
        editor_container = ttk.LabelFrame(
            editor_paned,
            text="Script Content",
            padding=8,
        )
        editor_paned.add(editor_container, weight=4)

        # Script editor with improved font and colors
        self.script_editor = scrolledtext.ScrolledText(
            editor_container,
            wrap=tk.WORD,
            font=("Consolas", 11),
            background="#ffffff",
            foreground="#333333",
            insertbackground="#007AFF",  # Cursor color
            selectbackground="#C9E4FF",  # Selection background
            selectforeground="#333333",  # Selection text color
        )
        self.script_editor.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.script_editor.bind("<KeyRelease>", lambda e: self.on_script_modified())

        # Add shortcut key bindings - but with custom event handlers to prevent double actions
        self.script_editor.bind("<Control-s>", lambda e: self.save_script())
        self.script_editor.bind("<Control-o>", lambda e: self.open_script())
        self.script_editor.bind("<Control-n>", lambda e: self.new_script())

        # Override the built-in copy/paste/cut with our custom handlers
        self.script_editor.bind("<Control-c>", lambda e: self.copy_text())
        self.script_editor.bind("<Control-v>", self.paste_text)
        self.script_editor.bind("<Control-x>", lambda e: self.cut_text())

        # Disable built-in paste to avoid double paste
        self.script_editor.bind("<<Paste>>", lambda e: "break")

        # Set default content
        self.script_editor.insert(
            tk.END,
            "# Click commands on the left to add them here\n# Or edit script content directly\n\n",
        )

        # Execution log area with improved styling
        log_container = ttk.LabelFrame(editor_paned, text="Execution Log")
        editor_paned.add(log_container, weight=1)

        # Execution log text widget with improved styling
        self.log_text = scrolledtext.ScrolledText(
            log_container,
            wrap=tk.WORD,
            font=("Consolas", 10),
            background="#1E1E1E",  # Dark background for better contrast
            foreground="#00FF00",
            state=tk.DISABLED,
            height=15,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        # Enable standard replication
        self.log_text.bind("<Control-c>", self.copy_log_text)
        self.log_text.bind(
            "<Button-3>", self.show_log_context_menu
        )  # Add right-click menu
        # Configure text colors for the log - enhance these for better contrast
        # Simplified log text tag configuration
        # Log 顏色配置（無粗體／無斜體，保留對比與清晰度）
        self.log_text.tag_configure("header", foreground="#FFD700")                         # 金黃
        self.log_text.tag_configure("error", foreground="#FF4C4C", background="#330000")    # 紅底
        self.log_text.tag_configure("warning", foreground="#FFA500", background="#2F1B00")  # 橘底
        self.log_text.tag_configure("info", foreground="#CCCCCC")                           # 淡灰
        self.log_text.tag_configure("success", foreground="#00DD99")                        # 青綠
        self.log_text.tag_configure("command", foreground="#FFFFFF", background="#444488")  # 白字藍底
        self.log_text.tag_configure("action", foreground="#00BFFF")                         # 藍
        self.log_text.tag_configure("result", foreground="#87CEFA")                         # 淺藍
        self.log_text.tag_configure("wait", foreground="#FFFF66")                           # 黃

        # Log control frame
        log_controls = ttk.Frame(log_container)
        log_controls.pack(fill=tk.X, pady=(0, 5))

        # Clear log button with improved styling
        clear_log_btn = ttk.Button(
            log_controls,
            text="Clear Log",
            command=self.clear_log,
            style="Secondary.TButton",
        )
        clear_log_btn.pack(side=tk.RIGHT, padx=5)
        save_log_btn = ttk.Button(
            log_controls,
            text="Save Log",
            command=self.save_log,
            style="Secondary.TButton",
        )
        save_log_btn.pack(side=tk.RIGHT, padx=5)
        # Start periodic check for log messages
        self.check_log_queue()

    def save_log(self):
        """Save log to file"""
        file_path = filedialog.asksaveasfilename(
            title="Save Log",
            defaultextension=".log",
            filetypes=[
                ("Log Files", "*.log"),
                ("Text Files", "*.txt"),
                ("All Files", "*.*"),
            ],
        )

        if file_path:
            try:
                log_content = self.log_text.get("1.0", tk.END)
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(log_content)
                self.statusbar.config(text=f"Logs saved to: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot save log: {str(e)}")

    def copy_log_text(self, event=None):
        """ "Copy log text"""
        try:
            # Temporarily enable log_text for text selection operations
            self.log_text.config(state=tk.NORMAL)

            ## Check if any text is selected
            try:
                selected_text = self.log_text.get(tk.SEL_FIRST, tk.SEL_LAST)
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.statusbar.config(text="Selected log text copied")
            except tk.TclError:  # No text is selected
                # If no text is selected, copy all the content
                all_text = self.log_text.get("1.0", tk.END)
                self.root.clipboard_clear()
                self.root.clipboard_append(all_text)
                self.statusbar.config(text="All log text copied")

        finally:
            # disabled state
            self.log_text.config(state=tk.DISABLED)

        return "break"

    def show_log_context_menu(self, event):
        """Show log right-click menu"""
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Copy", command=lambda: self.copy_log_text())
        context_menu.add_command(
            label="Copy All", command=lambda: self.copy_all_log_text()
        )
        context_menu.add_separator()
        context_menu.add_command(label="Clear Log", command=self.clear_log)

        # Display the menu on the right click position
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

        return "break"

    def copy_all_log_text(self):
        """All log text copied"""
        try:
            all_text = self.log_text.get("1.0", tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(all_text)
            self.statusbar.config(text="All log text copied")
        except Exception as e:
            self.statusbar.config(text=f"Copy failed: {str(e)}")

    def update_stop_on_error_status(self):
        """Update status bar with stop-on-error setting"""
        if self.stop_on_error_var.get():
            self.statusbar.config(text="Script execution will stop on first error")
        else:
            self.statusbar.config(
                text="Script execution will continue even if errors occur"
            )

    def validate_script(self):
        """Validate the script by checking command syntax without executing"""
        script_content = self.script_editor.get("1.0", tk.END)
        lines = script_content.splitlines()

        # Clear the log
        self.clear_log()
        self.add_to_log("=== Script Validation ===\n\n", "header")

        valid = True
        errors = []
        line_number = 0

        for line in lines:
            line_number += 1
            line = line.strip()

            # Skip empty lines and comments
            if not line or line.startswith("#"):
                continue

            # Validate basic command syntax
            if not line.startswith("--"):
                errors.append((line_number, line, "Command must start with '--'"))
                valid = False
                continue

            # Split command and parameters
            parts = []
            cmd_parts = []
            in_quotes = False
            current_part = ""

            for char in line:
                if char == '"' and (not in_quotes):
                    in_quotes = True
                    current_part += char
                elif char == '"' and in_quotes:
                    in_quotes = False
                    current_part += char
                elif char == " " and not in_quotes:
                    if current_part:
                        parts.append(current_part)
                        current_part = ""
                else:
                    current_part += char

            if current_part:
                parts.append(current_part)

            if not parts:
                errors.append((line_number, line, "Empty command"))
                valid = False
                continue

            command = parts[0]
            parameters = parts[1:] if len(parts) > 1 else []

            # Check if the command exists in any category
            command_exists = False
            command_def = None

            for category in self.command_categories.values():
                for cmd in category:
                    if cmd["name"] == command:
                        command_exists = True
                        command_def = cmd
                        break
                if command_exists:
                    break

            if not command_exists:
                errors.append((line_number, line, f"Unknown command: {command}"))
                valid = False
                continue

            # Check parameter count if the command definition exists
            if command_def and "params" in command_def:
                required_params = [
                    p for p in command_def["params"] if not p.endswith("(optional)")
                ]
                if len(parameters) < len(required_params):
                    missing = len(required_params) - len(parameters)
                    errors.append(
                        (
                            line_number,
                            line,
                            f"Missing {missing} required parameter(s). Expected: {', '.join(required_params)}",
                        )
                    )
                    valid = False

        # Report validation results
        if valid:
            self.add_to_log(
                "Script validation successful! No syntax errors found.\n", "success"
            )
        else:
            self.add_to_log(f"Found {len(errors)} syntax error(s):\n\n", "error")
            for line_num, line_text, error_msg in errors:
                self.add_to_log(f"Line {line_num}: {line_text}\n", "command")
                self.add_to_log(f"[Error] {error_msg}\n\n", "error")

        return valid

    def create_statusbar(self):
        # Status bar with improved styling
        self.statusbar = ttk.Label(
            self.root,
            text="Ready",
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(10, 2),
            font=("Segoe UI", 7),
            background="#f0f0f0",
            foreground="#555555",
        )
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Update statusbar with version info at startup
        self.statusbar.config(text=f"FalconUI Script Builder v{self.version} - Ready")

    def show_about_dialog(self):
        """Show an About dialog with version information"""
        about_window = tk.Toplevel(self.root)
        about_window.title("About FalconUI Script Builder")
        about_window.geometry("400x300")
        about_window.resizable(False, False)

        # Make the dialog modal
        about_window.transient(self.root)
        about_window.grab_set()

        # Center the dialog on parent window
        about_window.geometry(
            "+%d+%d"
            % (
                self.root.winfo_rootx() + self.root.winfo_width() // 2 - 200,
                self.root.winfo_rooty() + self.root.winfo_height() // 2 - 150,
            )
        )

        # Add version information
        header_frame = ttk.Frame(about_window)
        header_frame.pack(fill=tk.X, pady=20)

        app_name = ttk.Label(
            header_frame,
            text="FalconUI Script Builder",
            font=("Segoe UI", 16, "bold"),
            foreground=self.accent_color,
            background=self.bg_color,
        )
        app_name.pack()

        version_label = ttk.Label(
            header_frame,
            text=f"Version {self.version}",
            font=("Segoe UI", 10),
            foreground="#333333",
            background=self.bg_color,
        )
        version_label.pack(pady=5)

        # Add more information
        info_frame = ttk.Frame(about_window)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        about_text = scrolledtext.ScrolledText(
            info_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 8),
            height=5,
            background="#ffffff",
        )
        about_text.pack(fill=tk.BOTH, expand=True)
        about_text.insert(
            tk.END,
            """FalconUI Script Builder is a visual tool for creating and running automation scripts with FalconCommand.

    © 2025 Quanta BU4 SW
    All rights reserved.

    For support: rola.jeng@quantatw.com
    """,
        )
        about_text.config(state=tk.DISABLED)

        # Add close button
        button_frame = ttk.Frame(about_window)
        button_frame.pack(fill=tk.X, pady=10)

        close_button = ttk.Button(
            button_frame,
            text="Close",
            command=about_window.destroy,
            style="Secondary.TButton",
        )
        close_button.pack(side=tk.RIGHT, padx=20, pady=10)

    def create_app_menu(self):
        """Create an application menu bar with Help option"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(
            label="New", command=self.new_script, accelerator="Ctrl+N"
        )
        file_menu.add_command(
            label="Open...", command=self.open_script, accelerator="Ctrl+O"
        )
        file_menu.add_command(
            label="Save", command=self.save_script, accelerator="Ctrl+S"
        )
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close, accelerator="Alt+F4")

        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Validate Script", command=self.validate_script)
        # tools_menu.add_command(
        #     label="Step-by-Step Debug", command=self.run_step_by_step
        # )
        tools_menu.add_checkbutton(
            label="Stop on Error",
            variable=self.stop_on_error_var,
            command=self.update_stop_on_error_status,
        )

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about_dialog)

    def toggle_coordinate_tracker(self):
        """Toggle the live coordinate tracker on/off with a floating window"""
        if self.tracker_active:
            # Turn off tracker
            self.tracker_active = False
            self.track_button.config(text="Show Live XY")
            self.statusbar.config(text="Coordinate tracking stopped")

            # Reset coordinate display
            self.coord_x.set("X: 0")
            self.coord_y.set("Y: 0")

            # Destroy floating window if it exists
            if hasattr(self, "coord_window") and self.coord_window is not None:
                self.coord_window.destroy()
                self.coord_window = None
        else:
            # Turn on tracker
            self.tracker_active = True
            self.track_button.config(text="Stop Tracking")
            self.statusbar.config(text="Tracking mouse coordinates in floating window")

            # Create floating coordinate window
            self.create_floating_coord_window()

            # Start the coordinate tracking
            self.update_coordinates()

    def create_floating_coord_window(self):
        """Create a small floating window to display coordinates"""
        # Create a new toplevel window
        self.coord_window = tk.Toplevel(self.root)
        self.coord_window.title("Mouse Coordinates")
        self.coord_window.geometry("200x70")  # Small window size

        # Make the window stay on top
        self.coord_window.attributes("-topmost", True)

        # Set window properties
        self.coord_window.overrideredirect(True)  # Remove window decorations
        self.coord_window.attributes("-alpha", 0.8)  # Make it slightly transparent

        # Set a border and background
        self.coord_window.configure(bg="#333333")

        # Add coordinate labels with improved visibility
        title_label = tk.Label(
            self.coord_window,
            text="Mouse Position",
            font=("Segoe UI", 10, "bold"),
            fg="#FFFFFF",
            bg="#333333",
        )
        title_label.pack(pady=(5, 0))

        # Create a frame for coordinates
        coord_frame = tk.Frame(self.coord_window, bg="#333333")
        coord_frame.pack(fill=tk.X, expand=True, pady=5)

        # X coordinate label with high contrast
        self.float_x_label = tk.Label(
            coord_frame,
            textvariable=self.coord_x,
            font=("Consolas", 12, "bold"),
            fg="#00FF00",  # Bright green for visibility
            bg="#333333",
            width=10,
        )
        self.float_x_label.pack(side=tk.LEFT, padx=10)

        # Y coordinate label with high contrast
        self.float_y_label = tk.Label(
            coord_frame,
            textvariable=self.coord_y,
            font=("Consolas", 12, "bold"),
            fg="#00FF00",  # Bright green for visibility
            bg="#333333",
            width=10,
        )
        self.float_y_label.pack(side=tk.LEFT, padx=10)

        # Add drag functionality
        title_label.bind("<ButtonPress-1>", self.start_drag)
        title_label.bind("<ButtonRelease-1>", self.stop_drag)
        title_label.bind("<B1-Motion>", self.do_drag)

        # Position the window in the top-right corner initially
        screen_width = self.root.winfo_screenwidth()
        self.coord_window.geometry(f"+{screen_width-220}+10")

        # Add close button (small x in corner)
        close_button = tk.Button(
            self.coord_window,
            text="×",
            font=("Arial", 12),
            command=self.toggle_coordinate_tracker,
            bg="#333333",
            fg="#FF6666",
            bd=0,
            highlightthickness=0,
        )
        close_button.place(x=180, y=2)

    def start_drag(self, event):
        """Start window drag operation"""
        self._drag_data = {"x": event.x, "y": event.y}

    def stop_drag(self, event):
        """Stop window drag operation"""
        self._drag_data = {}

    def do_drag(self, event):
        """Handle window drag operation"""
        if not hasattr(self, "_drag_data"):
            return

        delta_x = event.x - self._drag_data["x"]
        delta_y = event.y - self._drag_data["y"]
        x = self.coord_window.winfo_x() + delta_x
        y = self.coord_window.winfo_y() + delta_y
        self.coord_window.geometry(f"+{x}+{y}")

    def update_coordinates(self):
        """Update the coordinate display with current mouse position"""
        if not self.tracker_active:
            return

        try:
            # Get current mouse position (using root's winfo_pointerxy)
            x, y = self.root.winfo_pointerxy()

            # Update the displayed coordinates
            self.coord_x.set(f"X: {x}")
            self.coord_y.set(f"Y: {y}")

            # Schedule the next update
            self.root.after(50, self.update_coordinates)  # Faster updates (50ms)
        except Exception as e:
            self.statusbar.config(text=f"Error tracking coordinates: {str(e)}")
            self.tracker_active = False
            self.track_button.config(text="Show Live XY")
            if hasattr(self, "coord_window") and self.coord_window is not None:
                self.coord_window.destroy()
                self.coord_window = None

    def add_command_to_script(self, command):
        # Format command and its parameters
        cmd_line = command["name"]
        if command["params"]:
            for param in command["params"]:
                if param.startswith("--"):
                    cmd_line += f" {param}"
                else:
                    cmd_line += f" <{param}>"

        # Add to editor at current position or end
        try:
            pos = self.script_editor.index(tk.INSERT)
            line = int(pos.split(".")[0])

            # Check if current line is empty
            line_content = self.script_editor.get(f"{line}.0", f"{line}.end")

            if line_content.strip():
                # Not an empty line, insert new line then add command
                self.script_editor.insert(f"{line}.end", f"\n{cmd_line}")
            else:
                # Empty line, replace with command
                self.script_editor.delete(f"{line}.0", f"{line}.end")
                self.script_editor.insert(f"{line}.0", cmd_line)
        except:
            # Insertion failed, add to end
            current_text = self.script_editor.get("1.0", tk.END)
            if current_text.endswith("\n\n"):
                self.script_editor.insert(tk.END, cmd_line)
            elif current_text.endswith("\n"):
                self.script_editor.insert(tk.END, f"{cmd_line}")
            else:
                self.script_editor.insert(tk.END, f"\n{cmd_line}")

        # Update status
        self.statusbar.config(text=f"Added command: {command['name']}")
        self.on_script_modified()

    def new_script(self):
        if self.is_script_modified:
            response = messagebox.askyesnocancel(
                "Unsaved Changes", "There are unsaved changes. Do you want to save?"
            )
            if response is None:  # Cancel operation
                return
            if response:  # Yes, save
                if not self.save_script():
                    return  # Save cancelled or failed, don't continue with new

        self.script_editor.delete("1.0", tk.END)
        self.script_editor.insert(tk.END, "# New Script\n\n")
        self.script_name_var.set("new_script.txt")
        self.current_script = "new_script.txt"
        self.is_script_modified = False
        self.update_title()
        self.statusbar.config(text="Created new script")

    def open_script(self):
        if self.is_script_modified:
            response = messagebox.askyesnocancel(
                "Unsaved Changes", "There are unsaved changes. Do you want to save?"
            )
            if response is None:  # Cancel operation
                return
            if response:  # Yes, save
                if not self.save_script():
                    return  # Save cancelled or failed, don't continue with open

        file_path = filedialog.askopenfilename(
            title="Open Script",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
        )

        if file_path:
            try:
                # Try to open with different encodings
                try:
                    # First try UTF-8 with BOM
                    with open(file_path, "r", encoding="utf-8-sig") as file:
                        content = file.read()
                except UnicodeDecodeError:
                    try:
                        # If failed, try plain UTF-8
                        with open(file_path, "r", encoding="utf-8") as file:
                            content = file.read()
                    except UnicodeDecodeError:
                        # Finally try system default encoding
                        with open(file_path, "r") as file:
                            content = file.read()

                self.script_editor.delete("1.0", tk.END)
                self.script_editor.insert(tk.END, content)

                # Update script name
                script_name = os.path.basename(file_path)
                self.script_name_var.set(script_name)
                self.current_script = file_path
                self.is_script_modified = False
                self.update_title()

                self.statusbar.config(text=f"Opened: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open file: {str(e)}")

    def save_script(self):
        script_content = self.script_editor.get("1.0", tk.END)
        script_name = self.script_name_var.get()

        # If it's a new script or needs to be saved as
        if self.current_script == "new_script.txt" or not os.path.exists(
            self.current_script
        ):
            file_path = filedialog.asksaveasfilename(
                title="Save Script",
                initialfile=script_name,
                defaultextension=".txt",
                filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            )

            if not file_path:  # User cancelled
                return False

            self.current_script = file_path
            script_name = os.path.basename(file_path)
            self.script_name_var.set(script_name)

        try:
            # Use plain UTF-8 without BOM
            with open(self.current_script, "w", encoding="utf-8") as file:
                # Remove BOM if present at the beginning of the content
                if script_content.startswith("\ufeff"):
                    script_content = script_content[1:]
                file.write(script_content)

            self.is_script_modified = False
            self.update_title()
            self.statusbar.config(text=f"Saved: {self.current_script}")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Save failed: {str(e)}")
            return False

    def schedule_autosave(self):
        """Schedule the next autosave"""
        if self.autosave_enabled:
            self.root.after(self.autosave_interval, self.auto_save)

    def auto_save(self):
        """Automatically save the current script and create backups"""
        if self.is_script_modified and self.autosave_enabled:
            if self.current_script != "new_script.txt" and os.path.exists(
                self.current_script
            ):
                try:
                    # Regular autosave to the current file
                    script_content = self.script_editor.get("1.0", tk.END)

                    # Save to current file
                    with open(self.current_script, "w", encoding="utf-8") as file:
                        if script_content.startswith("\ufeff"):
                            script_content = script_content[1:]
                        file.write(script_content)

                    # Create a backup copy
                    timestamp = time.strftime("%Y%m%d-%H%M%S")
                    script_name = os.path.basename(self.current_script)
                    backup_name = f"{os.path.splitext(script_name)[0]}_{timestamp}{os.path.splitext(script_name)[1]}"
                    backup_path = os.path.join(self.backup_dir, backup_name)

                    # Save backup
                    with open(backup_path, "w", encoding="utf-8") as backup_file:
                        backup_file.write(script_content)

                    # Manage backup count - delete oldest backups if we have too many
                    script_base = os.path.splitext(script_name)[0]
                    backups = [
                        f
                        for f in os.listdir(self.backup_dir)
                        if f.startswith(script_base) and f != script_name
                    ]

                    if len(backups) > self.backup_count:
                        # Sort by timestamp (newest first)
                        backups.sort(reverse=True)
                        # Remove oldest backups
                        for old_backup in backups[self.backup_count :]:
                            try:
                                os.remove(os.path.join(self.backup_dir, old_backup))
                            except:
                                pass

                    self.is_script_modified = False
                    self.update_title()
                    self.statusbar.config(text=f"Auto-saved: {self.current_script}")

                except Exception as e:
                    self.statusbar.config(text=f"Auto-save failed: {str(e)}")

        # Schedule next autosave
        self.schedule_autosave()

    def add_to_log(self, message, tag=None):
        """Add text to the log with tag applied to the specific line only"""
        # Split the message into lines to apply tag to each line separately
        if not message:
            return
            
        self.log_queue.put((message, tag))

    def check_log_queue(self):
        """Check if there are messages in the queue and add them to the log text widget"""
        try:
            while True:
                message, tag = self.log_queue.get_nowait()
                # Temporarily enable text widget for modification
                self.log_text.config(state=tk.NORMAL)
                
                # Get the current end position before inserting
                start_index = self.log_text.index(tk.END)
                
                # Insert the message
                self.log_text.insert(tk.END, message)
                
                # Apply tag only to the newly inserted text
                if tag:
                    end_index = self.log_text.index(tk.END)
                    self.log_text.tag_add(tag, start_index, end_index)
                    
                self.log_text.see(tk.END)
                # Return to disabled state to prevent user edits
                self.log_text.config(state=tk.DISABLED)
                self.log_queue.task_done()
        except queue.Empty:
            # Schedule next check
            self.root.after(100, self.check_log_queue)

    def clear_log(self):
        """Clear the execution log display"""
        # Temporarily enable the text widget for clearing
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.add_to_log("Log cleared.\n", "info")

    def process_output(self, process, script_path):
        """Process the output from the script execution with enhanced error detection"""
        error_count = 0
        error_lines = []

        # Create a buffer to capture all output for logging
        output_buffer = io.StringIO()

        # Get screen resolution and try to get scale factor from falconCommand
        screen_width, screen_height = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        
        # Get scale factor if possible (simplified version for GUI app)
        try:
            import ctypes
            scale_factor = 1.0
            if sys.platform == 'win32':
                # This is a simplified approach for the GUI application
                try:
                    # Get DPI awareness
                    awareness = ctypes.c_int()
                    ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
                    
                    # Get DPI
                    dpi = ctypes.c_uint(0)
                    ctypes.windll.shcore.GetDpiForMonitor(0, 0, ctypes.byref(dpi), ctypes.byref(dpi))
                    
                    # Calculate scale factor (96 is the base DPI)
                    scale_factor = dpi.value / 96.0
                except:
                    # Fallback method
                    hdc = ctypes.windll.user32.GetDC(0)
                    dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
                    ctypes.windll.user32.ReleaseDC(0, hdc)
                    scale_factor = dpi / 96.0
        except:
            scale_factor = 1.0  # Default if we can't detect
        
        # Write header to the buffer with enhanced system information
        output_buffer.write(f"=== FalconUI Script Execution Log ===\n")
        output_buffer.write(f"Script: {script_path}\n")
        output_buffer.write(f"Date/Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        output_buffer.write(f"Screen Resolution: {screen_width}x{screen_height}\n")
        output_buffer.write(f"Display Scale Factor: {scale_factor:.2f} ({int(scale_factor*100)}%)\n")
        output_buffer.write(f"Stop on Error: {'Yes' if self.stop_on_error_var.get() else 'No'}\n")
        output_buffer.write("\n=== Execution Output ===\n\n")

        # Current command tracking
        current_command = None
        command_output = []

        while True:
            # Check if process has ended
            if process.poll() is not None:
                break

            # Get any available output
            output = process.stdout.readline()
            if output:
                # Write to log buffer
                output_buffer.write(output)

                # Group command outputs together
                if output.startswith("[command]"):
                    # If we were tracking a previous command, output its group now
                    if current_command and command_output:
                        self.add_command_group(current_command, command_output)
                        command_output = []
                    
                    # Start tracking new command
                    current_command = output.strip()
                    command_output = []
                else:
                    # Add to current command's output
                    if current_command:
                        command_output.append((output.strip(), self._determine_output_type(output)))
                    else:
                        # Regular line-by-line handling for non-grouped output
                        if "Error" in output or "Failed" in output or "Exception" in output:
                            error_count += 1
                            error_lines.append(output.strip())
                            timestamp = time.strftime("%H:%M:%S")
                            self.add_to_log(f"[{timestamp}] {output}", "error")
                        else:
                            tag = self._determine_output_type(output)
                            timestamp = time.strftime("%H:%M:%S")
                            self.add_to_log(f"[{timestamp}] {output}", tag)
            else:
                # No output available, give control back to GUI
                time.sleep(0.1)

        # Process remaining output and flush last command group
        remaining_output, stderr = process.communicate()
        if remaining_output:
            # Process the remaining output
            lines = remaining_output.splitlines()
            for line in lines:
                line_str = line + "\n" if not line.endswith("\n") else line
                output_buffer.write(line_str)
                
                if line_str.startswith("[command]"):
                    # If we were tracking a previous command, output its group now
                    if current_command and command_output:
                        self.add_command_group(current_command, command_output)
                        command_output = []
                    
                    # Start tracking new command
                    current_command = line_str.strip()
                    command_output = []
                else:
                    # Add to current command's output
                    if current_command:
                        command_output.append((line_str.strip(), self._determine_output_type(line_str)))
                    else:
                        # Regular handling
                        tag = self._determine_output_type(line_str)
                        self.add_to_log(line_str, tag)
        
        # Flush any remaining command group
        if current_command and command_output:
            self.add_command_group(current_command, command_output)

        # Process stderr if any
        if stderr:
            # Write stderr to log buffer
            output_buffer.write("\n=== ERRORS ===\n")
            output_buffer.write(stderr)

            # Process stderr for UI display
            for line in stderr.splitlines():
                line_str = line + "\n" if not line.endswith("\n") else line
                error_count += 1
                error_lines.append(line_str.strip())
                self.add_to_log(line_str, "error")

        # Get return code
        return_code = process.returncode

        # Add summary to log buffer
        output_buffer.write("\n=== Execution Summary ===\n")
        output_buffer.write(f"Return Code: {return_code}\n")
        output_buffer.write(f"Error Count: {error_count}\n")

        if error_count > 0:
            output_buffer.write("\n--- Error Summary ---\n")
            for i, error in enumerate(error_lines, 1):
                output_buffer.write(f"{i}. {error}\n")

        # Add completion timestamp
        output_buffer.write(
            f"\nExecution completed at: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
        )

        # Save the log to file
        try:
            # Create Falcon_Log directory if it doesn't exist
            today = time.strftime("%Y-%m-%d")
            log_dir = Path(f"C:/Falcon_Log/{today}")
            log_dir.mkdir(parents=True, exist_ok=True)

            # log_dir = os.path.join("C:/", "Falcon_Log")
            # os.makedirs(log_dir, exist_ok=True)

            # Extract original script name
            script_name = os.path.basename(script_path)
            script_name = os.path.splitext(script_name)[0]

            # Create timestamp
            timestamp = time.strftime("%Y%m%d_%H%M%S")

            # Create log file path
            log_filename = f"falconUI_{script_name}_{timestamp}.log"
            # log_path = os.path.join(log_dir, log_filename)
            log_path = log_dir / log_filename

            # Write log to file
            with open(log_path, "w", encoding="utf-8") as log_file:
                log_file.write(output_buffer.getvalue())

            # Add log file info to UI
            self.add_to_log(f"\nLog saved to: {log_path}\n", "info")

            # Update status
            if return_code == 0:
                self.statusbar.config(
                    text=f"Execution successful: {script_path} | Log: {log_path}"
                )
            else:
                self.statusbar.config(
                    text=f"Execution failed (code: {return_code}) with {error_count} errors | Log: {log_path}"
                )
        except Exception as e:
            self.add_to_log(f"\nError saving log file: {str(e)}\n", "error")
            self.statusbar.config(
                text=f"Execution completed but failed to save log: {str(e)}"
            )

        # Error summary in UI
        if error_count > 0:
            self.add_to_log(f"\n--- Error Summary ---\n", "error")
            self.add_to_log(f"Total errors: {error_count}\n", "error")
            for i, error in enumerate(error_lines[:5], 1):  # Show first 5 errors
                self.add_to_log(f"{i}. {error}\n", "error")
            if len(error_lines) > 5:
                self.add_to_log(
                    f"... and {len(error_lines) - 5} more errors\n", "error"
                )

        self.add_to_log(
            f"\n--- Process returned code: {return_code} ---\n",
            "success" if return_code == 0 else "error",
        )

        # Clean up
        self.current_process = None
        self.stop_btn.config(state=tk.DISABLED)
        self.root.deiconify()

    def _determine_output_type(self, line):
        """Determine the type of log output for appropriate styling"""
        line_str = line.strip()
        if "[command]" in line_str:
            return "command"
        elif "Error" in line_str or "Failed" in line_str or "Exception" in line_str:
            return "error"
        elif "[Warning]" in line_str:
            return "warning"
        elif "[V]" in line_str or "successfully" in line_str.lower():
            return "success"
        elif "Attempting to" in line_str or "Checking" in line_str:
            return "action"
        elif "Found" in line_str or "installed" in line_str:
            return "result"
        elif "Waiting" in line_str:
            return "wait"
        else:
            return "info"

    def add_command_group(self, command, outputs):
        """Add a visually separated command group to the log with timestamp"""
        # Get current timestamp
        timestamp = time.strftime("%H:%M:%S")
        
        # Add separator
        self.add_to_log("\n" + "=" * 50 + "\n", "info")
        
        # Add command with prominent styling and timestamp
        self.add_to_log(f"[{timestamp}] {command}\n", "command")
        
        # Add separator between command and outputs
        self.add_to_log("-" * 50 + "\n", "info")
        
        # Add each output with appropriate tag
        for output, tag in outputs:
            self.add_to_log(f"  {output}\n", tag)
        
        # Add bottom separator
        self.add_to_log("\n", "info")

    def run_script(self):
        # Disable the run button and enable stop button during execution
        self.stop_btn.config(state=tk.NORMAL)

        # Save script first
        if self.is_script_modified:
            response = messagebox.askyesno(
                "Unsaved Changes", "Script must be saved before running. Continue?"
            )
            if not response:
                self.stop_btn.config(state=tk.DISABLED)
                return
            if not self.save_script():
                self.stop_btn.config(state=tk.DISABLED)
                return

        # Validate the script before running
        if not self.validate_script():
            if not messagebox.askyesno(
                "Script Validation Failed",
                "The script contains syntax errors. Run anyway?",
            ):
                self.stop_btn.config(state=tk.DISABLED)
                return

        # Check if falconui.exe exists
        if not os.path.exists(self.falconui_path):
            file_path = filedialog.askopenfilename(
                title="Select falconCommand.exe location",
                filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")],
            )
            if not file_path:
                self.stop_btn.config(state=tk.DISABLED)
                return
            self.falconui_path = file_path

        try:
            # Execute script directly without creating a temporary copy
            cmd = [self.falconui_path, "--command-file", self.current_script]

            # Add stop-on-error flag if needed
            if hasattr(self, "stop_on_error_var") and self.stop_on_error_var.get():
                cmd.append("--stop-on-error")

            # Show execution message
            self.statusbar.config(text=f"Running: {self.current_script}")
            self.root.update()

            # Clear the log before running
            self.clear_log()
            self.add_to_log(
                f"=== Running script: {self.current_script} ===\n\n", "header"
            )

            # Minimize the main window
            self.root.iconify()

            # Use subprocess to execute with creationflags for non-blocking behavior
            startupinfo = None
            creationflags = 0
            if os.name == "nt":  # Windows
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                # This flag allows child processes to continue running after the parent is closed
                creationflags = subprocess.CREATE_NEW_PROCESS_GROUP

            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                encoding="utf-8",
                startupinfo=startupinfo,
                creationflags=creationflags,
            )

            # Store the process for potential stopping
            self.current_process = process

            # Start output processing in a separate thread to avoid blocking the GUI
            output_thread = threading.Thread(
                target=lambda: self.process_output(process, self.current_script)
            )
            output_thread.daemon = (
                True  # Thread will close when main application closes
            )
            output_thread.start()

        except Exception as e:
            self.add_to_log(f"Cannot execute script: {str(e)}\n", "error")
            self.statusbar.config(text=f"Execution failed: {str(e)}")
            self.stop_btn.config(state=tk.DISABLED)

    def run_step_by_step(self):
        """Run the script in step-by-step debug mode"""
        if self.is_script_modified:
            response = messagebox.askyesno(
                "Unsaved Changes", "Script must be saved before running. Continue?"
            )
            if not response:
                return
            if not self.save_script():
                return

        # Parse the script to extract executable lines
        script_content = self.script_editor.get("1.0", tk.END)
        lines = script_content.splitlines()

        self.debug_lines = []
        self.debug_line_numbers = []
        line_num = 0

        for line in lines:
            line_num += 1
            stripped = line.strip()
            if stripped and not stripped.startswith("#"):
                self.debug_lines.append(stripped)
                self.debug_line_numbers.append(line_num)

        if not self.debug_lines:
            messagebox.showinfo(
                "Debug Mode", "No executable commands found in the script."
            )
            return

        # Initialize debug state
        self.debug_mode = True
        self.current_debug_line = 0
        self.debug_step_event.clear()
        self.debug_skip_event.clear()
        self.debug_stop_event.clear()

        # Update UI for debug mode
        self.step_btn.config(state=tk.DISABLED)
        self.next_step_btn.config(state=tk.NORMAL)
        self.skip_step_btn.config(state=tk.NORMAL)
        self.stop_debug_btn.config(state=tk.NORMAL)

        # Disable normal run button during debug
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Run":
                widget.config(state=tk.DISABLED)
                break

        # Clear the log
        self.clear_log()
        self.add_to_log("=== Starting Step-by-Step Debug Mode ===\n\n", "header")
        self.add_to_log(f"Loaded {len(self.debug_lines)} executable commands\n", "info")
        self.add_to_log("Press 'Next Step' to execute the first command\n\n", "info")

        # Highlight the first command in the editor
        self.highlight_debug_line()

        # Start debug thread
        debug_thread = threading.Thread(target=self.debug_execution_thread)
        debug_thread.daemon = True
        debug_thread.start()

    def highlight_debug_line(self):
        """Highlight the current debug line in the editor"""
        # Clear previous highlights
        self.script_editor.tag_remove("current_line", "1.0", tk.END)

        if self.current_debug_line < len(self.debug_line_numbers):
            line_num = self.debug_line_numbers[self.current_debug_line]
            self.script_editor.tag_add(
                "current_line", f"{line_num}.0", f"{line_num}.end"
            )
            self.script_editor.tag_config(
                "current_line", background="#FFD700"
            )  # Gold highlight
            self.script_editor.see(f"{line_num}.0")  # Scroll to show the line

            # Update line indicator
            self.current_line_var.set(
                f"Line {line_num}: {self.debug_lines[self.current_debug_line]}"
            )
        else:
            self.current_line_var.set("Debug completed")

    def execute_next_step(self):
        """Execute the next step in debug mode"""
        if not self.debug_mode or self.current_debug_line >= len(self.debug_lines):
            return

        # Signal the debug thread to execute the next command
        self.debug_step_event.set()

    def skip_current_step(self):
        """Skip the current step in debug mode"""
        if not self.debug_mode or self.current_debug_line >= len(self.debug_lines):
            return

        # Signal the debug thread to skip the current command
        self.debug_skip_event.set()

    def stop_debug(self):
        """Stop the debug session"""
        if not self.debug_mode:
            return

        # Signal the debug thread to stop
        self.debug_stop_event.set()

        # Reset debug state
        self.exit_debug_mode()

    def exit_debug_mode(self):
        """Clean up after debug session ends"""
        self.debug_mode = False

        # Reset UI
        self.step_btn.config(state=tk.NORMAL)
        self.next_step_btn.config(state=tk.DISABLED)
        self.skip_step_btn.config(state=tk.DISABLED)
        self.stop_debug_btn.config(state=tk.DISABLED)

        # Re-enable normal run button
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button) and widget.cget("text") == "Run":
                widget.config(state=tk.NORMAL)
                break

        # Clear any highlights
        self.script_editor.tag_remove("current_line", "1.0", tk.END)
        self.current_line_var.set("")

        self.add_to_log("\n=== Debug Session Ended ===\n", "header")

    def debug_execution_thread(self):
        """Thread that handles the actual execution of commands in debug mode"""

        # Create a buffer to store all debug output
        debug_log = io.StringIO()

        # Write header to the log buffer
        debug_log.write(f"=== FalconUI Debug Session Log ===\n")
        debug_log.write(f"Script: {self.current_script}\n")
        debug_log.write(f"Date/Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        debug_log.write(
            f"Stop on Error: {'Yes' if self.stop_on_error_var.get() else 'No'}\n"
        )
        debug_log.write(f"Total Debug Lines: {len(self.debug_lines)}\n\n")

        try:
            while self.current_debug_line < len(self.debug_lines) and self.debug_mode:
                # Wait for user to press Next, Skip or Stop
                while not any(
                    [
                        self.debug_step_event.is_set(),
                        self.debug_skip_event.is_set(),
                        self.debug_stop_event.is_set(),
                    ]
                ):
                    time.sleep(0.1)

                # Check if stop was requested
                if self.debug_stop_event.is_set():
                    self.debug_stop_event.clear()
                    debug_log.write("\n=== Debug Session Stopped by User ===\n")
                    break

                # Check if skip was requested
                if self.debug_skip_event.is_set():
                    self.debug_skip_event.clear()
                    line = self.debug_lines[self.current_debug_line]
                    skip_msg = f"Skipping: {line}\n"
                    debug_log.write(skip_msg)
                    self.add_to_log(skip_msg, "warning")
                    self.current_debug_line += 1
                    self.highlight_debug_line()
                    continue

                # Next step requested
                if self.debug_step_event.is_set():
                    self.debug_step_event.clear()

                    # Get the current command
                    line = self.debug_lines[self.current_debug_line]
                    line_num = self.debug_line_numbers[self.current_debug_line]

                    exec_msg = f"Executing line {line_num}: {line}\n"
                    debug_log.write(exec_msg)
                    self.add_to_log(exec_msg, "command")

                    # Execute the command
                    try:
                        # Split the command into parts
                        parts = []
                        current_part = ""
                        in_quotes = False
                        i = 0

                        while i < len(line):
                            char = line[i]
                            if char == '"':
                                in_quotes = not in_quotes
                                current_part += char
                            elif char == " " and not in_quotes:
                                if current_part:
                                    parts.append(current_part)
                                    current_part = ""
                            else:
                                current_part += char
                            i += 1

                        if current_part:
                            parts.append(current_part)

                        if parts:
                            # Execute using subprocess
                            cmd = [self.falconui_path] + parts
                            process = subprocess.Popen(
                                cmd,
                                stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE,
                                universal_newlines=True,
                            )

                            stdout, stderr = process.communicate()

                            if stdout:
                                debug_log.write(stdout)
                                self.add_to_log(stdout, "info")

                            if stderr:
                                debug_log.write("\n--- Error Output ---\n")
                                debug_log.write(stderr)
                                self.add_to_log(stderr, "error")

                            if process.returncode == 0:
                                success_msg = f"Command completed successfully\n"
                                debug_log.write(success_msg)
                                self.add_to_log(success_msg, "success")
                            else:
                                error_msg = f"Command failed with exit code {process.returncode}\n"
                                debug_log.write(error_msg)
                                self.add_to_log(error_msg, "error")

                                # Check if we should stop on error
                                if (
                                    hasattr(self, "stop_on_error_var")
                                    and self.stop_on_error_var.get()
                                ):
                                    stop_msg = "Stopping debug execution due to error\n"
                                    debug_log.write(stop_msg)
                                    self.add_to_log(stop_msg, "error")
                                    break
                    except Exception as e:
                        error_msg = f"Error executing command: {str(e)}\n"
                        debug_log.write(error_msg)
                        self.add_to_log(error_msg, "error")

                    # Move to next line
                    self.current_debug_line += 1
                    self.highlight_debug_line()

            # Debug session completed or stopped
            if self.current_debug_line >= len(self.debug_lines) and self.debug_mode:
                completion_msg = "\n=== Debug Execution Completed ===\n"
                debug_log.write(completion_msg)
                self.add_to_log(completion_msg, "success")

            # Write completion timestamp
            debug_log.write(
                f"\nDebug session completed at: {time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            )

            # Save the debug log to file
            try:
                # Create Falcon_Log directory if it doesn't exist
                today = time.strftime("%Y-%m-%d")
                log_dir = os.path.join("C:/", "Falcon_Log", today)
                os.makedirs(log_dir, exist_ok=True)


                # Extract original script name
                script_name = os.path.basename(self.current_script)
                script_name = os.path.splitext(script_name)[0]

                # Create timestamp
                timestamp = time.strftime("%Y%m%d_%H%M%S")

                # Create log file path (add debug indicator)
                log_filename = f"{script_name}_debug_{timestamp}.log"
                #log_path = os.path.join(log_dir, log_filename)
                log_path = log_dir / log_filename

                # Write log to file
                with open(log_path, "w", encoding="utf-8") as log_file:
                    log_file.write(debug_log.getvalue())

                self.add_to_log(f"Debug log saved to: {log_path}\n", "info")
                self.statusbar.config(text=f"Debug session completed | Log: {log_path}")
            except Exception as e:
                self.add_to_log(f"Error saving debug log: {str(e)}\n", "error")
                self.statusbar.config(
                    text=f"Debug completed but failed to save log: {str(e)}"
                )

            # Exit debug mode
            self.root.after(0, self.exit_debug_mode)

        except Exception as e:
            error_msg = f"Error in debug thread: {str(e)}\n"
            debug_log.write(error_msg)
            self.add_to_log(error_msg, "error")

            # Try to save log even on error
            try:
                today = time.strftime("%Y-%m-%d")
                log_dir = os.path.join("C:/", "Falcon_Log", today)
                os.makedirs(log_dir, exist_ok=True)
                script_name = os.path.basename(self.current_script)
                script_name = os.path.splitext(script_name)[0]
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                log_filename = f"{script_name}_debug_error_{timestamp}.log"
                log_path = os.path.join(log_dir, log_filename)

                with open(log_path, "w", encoding="utf-8") as log_file:
                    log_file.write(debug_log.getvalue())
            except:
                pass

            self.root.after(0, self.exit_debug_mode)

    def stop_execution(self):
        """Stop the currently running process"""
        if self.current_process and self.current_process.poll() is None:
            try:
                # Try to terminate gracefully first
                self.current_process.terminate()
                self.add_to_log("\n--- Terminating process... ---\n", "error")

                # Use after() to check termination without blocking the GUI
                self.root.after(500, self.check_termination)
            except Exception as e:
                self.add_to_log(f"Error terminating process: {str(e)}\n", "error")
                self.statusbar.config(text=f"Error terminating process: {str(e)}")

    def check_termination(self):
        """Check if the process has terminated, else force kill"""
        if self.current_process and self.current_process.poll() is None:
            try:
                self.current_process.kill()
                self.add_to_log("--- Process forcefully killed ---\n", "error")
            except Exception as e:
                self.add_to_log(f"Error killing process: {str(e)}\n", "error")
            finally:
                self.current_process = None
                self.stop_btn.config(state=tk.DISABLED)
                self.statusbar.config(text="Process terminated")
        else:
            self.current_process = None
            self.stop_btn.config(state=tk.DISABLED)
            self.statusbar.config(text="Process terminated")

    def clear_script(self):
        if messagebox.askyesno(
            "Confirm", "Are you sure you want to clear the current script content?"
        ):
            self.script_editor.delete("1.0", tk.END)
            self.statusbar.config(text="Script content cleared")
            self.on_script_modified()

    def copy_text(self):
        """Copy selected text to clipboard"""
        try:
            selected_text = self.script_editor.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.statusbar.config(text="Copied selected text to clipboard")
        except tk.TclError:  # No text selected
            self.statusbar.config(text="No text selected")

    def paste_text(self, event=None):
        """Paste text from clipboard"""
        # Return 'break' to prevent default paste behavior
        try:
            clipboard_text = self.root.clipboard_get()
            # If text is selected, delete it first
            try:
                self.script_editor.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                pass  # No text selected, paste at cursor position

            self.script_editor.insert(tk.INSERT, clipboard_text)
            self.on_script_modified()
            self.statusbar.config(text="Text pasted")
        except tk.TclError:
            self.statusbar.config(text="No text in clipboard")

        # Return 'break' to prevent the default paste behavior
        return "break"

    def cut_text(self):
        """Cut selected text to clipboard"""
        try:
            selected_text = self.script_editor.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(selected_text)
            self.script_editor.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.on_script_modified()
            self.statusbar.config(text="Cut selected text")
        except tk.TclError:  # No text selected
            self.statusbar.config(text="No text selected")

    def on_script_modified(self):
        if not self.is_script_modified:
            self.is_script_modified = True
            self.update_title()

    def update_title(self):
        script_name = os.path.basename(self.current_script)
        if self.is_script_modified:
            self.root.title(
                f"FalconUI Script Builder v{self.version} - {script_name} *"
            )
        else:
            self.root.title(f"FalconUI Script Builder v{self.version} - {script_name}")

    def on_close(self):
        if self.is_script_modified:
            response = messagebox.askyesnocancel(
                "Unsaved Changes", "There are unsaved changes. Do you want to save?"
            )
            if response is None:  # Cancel close
                return
            if response:  # Yes, save
                if not self.save_script():
                    return  # Save cancelled or failed, don't close

        # Stop any running process
        if self.current_process and self.current_process.poll() is None:
            try:
                self.current_process.terminate()
                self.root.after(
                    500,
                    lambda: (
                        self.current_process.kill()
                        if self.current_process.poll() is None
                        else None
                    ),
                )
            except:
                pass

        self.root.destroy()


def main():
    root = tk.Tk()
    app = FalconUIScriptBuilder(root)

    # Bind global shortcut keys
    root.bind("<Control-s>", lambda e: app.save_script())
    root.bind("<Control-o>", lambda e: app.open_script())
    root.bind("<Control-n>", lambda e: app.new_script())
    root.bind("<Control-c>", lambda e: app.copy_text())
    root.bind("<Control-v>", lambda e: app.paste_text())
    root.bind("<Control-x>", lambda e: app.cut_text())

    root.mainloop()


if __name__ == "__main__":
    main()
