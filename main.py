import configparser
from ctypes import wintypes
import multiprocessing
import ctypes
import threading
import win32com.client
import customtkinter as ctk
from customtkinter import CTkImage  
from tkinter import messagebox, filedialog 
import tkinter.ttk as ttk
import json
import os
import time
from collections import defaultdict
import sys
from PIL import Image

if sys.platform == "win32":
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("com.example.AudioEditorApp")
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
data_dir = os.path.join(os.path.expanduser("~"), ".FolderLock&Hide")
os.makedirs(data_dir, exist_ok=True)
history_file = os.path.join(data_dir, "history.json")
MUTEX_NAME = "Global\\FolderProtectorMonitorMutex"

def resource_path(relative_path):

    try:
        base_path = sys._MEIPASS

    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_config():

    default = {"password": "1234", "protected_paths": []}

    if not os.path.exists(history_file):
        save_config(default)
        return default

    try:

        with open(history_file, "r") as f:
            return json.load(f)

    except:
        return default

def save_config(config):

    with open(history_file, "w") as f:
        json.dump(config, f, indent=4)

def password_dialog(folder_path, current_password, master=None):
    folder_name = os.path.basename(folder_path) or folder_path
    result = False
    temp_root = None

    if master is None:
        temp_root = ctk.CTk()
        temp_root.withdraw()
        master = temp_root
    dialog = ctk.CTkToplevel(master=master)
    dialog.title("Folder Access")
    dialog.resizable(False, False)
    window_width = 400
    window_height = 220
    dialog.update_idletasks()
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
    dialog.attributes('-topmost', True)
    dialog.update()
    dialog.attributes('-topmost', False)
    icon_path = resource_path(r"icons\icon.ico")
    icon_setter = WindowIconSetter("Folder Access", icon_path)

    def apply_icon_async():
        time.sleep(1)
        icon_setter.set_icon()
    threading.Thread(target=apply_icon_async, daemon=True).start()
    main_frame = ctk.CTkFrame(dialog)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    ctk.CTkLabel(main_frame, text=f"Folder: {folder_name}", font=("Segoe UI", 14, "bold"),
                 wraplength=350, justify="left").pack(pady=(0,10))
    ctk.CTkLabel(main_frame, text="Enter password to open this folder:", font=("Segoe UI", 12)).pack()
    pwd_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    pwd_frame.pack(pady=10)
    entry = ctk.CTkEntry(pwd_frame, show="*", width=250, font=("Arial", 12))
    entry.pack(side="left", padx=(0,5))
    entry.focus_set()
    show_password = CTkImage(light_image=Image.open(resource_path(r"icons\show_password.png")), size=(20,20))
    hide_password = CTkImage(light_image=Image.open(resource_path(r"icons\hide_password.png")), size=(20,20))
    show_var = False

    def toggle_show():
        nonlocal show_var

        if show_var:
            entry.configure(show="*")
            show_var = False
            eye_btn.configure(image=hide_password)
        else:
            entry.configure(show="")
            show_var = True
            eye_btn.configure(image=show_password)
    eye_btn = ctk.CTkButton(pwd_frame, text="", image=hide_password, width=40, command=toggle_show)
    eye_btn.pack(side="right")
    btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    btn_frame.pack(pady=10)

    def check_password():
        nonlocal result

        if entry.get() == current_password:
            result = True
            dialog.destroy()
        else:
            messagebox.showerror("Wrong Password", "Incorrect password. Folder will stay closed.")

    def on_cancel():
        nonlocal result
        result = False
        dialog.destroy()
    ctk.CTkButton(btn_frame, text="OK", command=check_password, width=100).pack(side="left", padx=5)
    ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel, width=100).pack(side="left", padx=5)
    dialog.bind("<Return>", lambda event: check_password())
    dialog.bind("<Escape>", lambda event: on_cancel())
    dialog.grab_set()
    dialog.wait_window()

    if temp_root:
        temp_root.destroy()
    return result

def run_monitor():
    ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if ctypes.windll.kernel32.GetLastError() == 183:
        return
    root = ctk.CTk()
    root.withdraw()
    root.update_idletasks()
    window_allowed_path = {}
    grace_paths = defaultdict(float)
    pending_folders = set()
    GRACE_SECONDS = 1.5
    shell = win32com.client.Dispatch("Shell.Application")

    def is_explicitly_protected(path, protected_roots):
        path = os.path.normpath(path)
        return path in protected_roots

    def show_password_dialog(folder_path, password):

        if folder_path not in pending_folders:
            return
        pending_folders.discard(folder_path)

        if password_dialog(folder_path, password, master=root):
            for w in shell.Windows():

                try:

                    if w.Document.Folder.Self.Path == folder_path:
                        w.Quit()

                except:
                    pass
            shell.Explore(folder_path)
            grace_paths[folder_path] = time.time()
        else:
            for w in shell.Windows():

                try:

                    if w.Document.Folder.Self.Path == folder_path:
                        w.Quit()

                except:
                    pass

    def check_windows():
        now = time.time()
        config = load_config()
        protected_roots = set(config.get("protected_paths", []))
        password = config.get("password", "1234")
        for path in list(grace_paths.keys()):

            if now - grace_paths[path] > GRACE_SECONDS:
                del grace_paths[path]
        current_windows = list(shell.Windows())
        current_hwnds = set()
        for window in current_windows:

            try:
                hwnd = window.HWND
                current_hwnds.add(hwnd)
                folder_path = window.Document.Folder.Self.Path

                if not is_explicitly_protected(folder_path, protected_roots):

                    if hwnd in window_allowed_path:
                        del window_allowed_path[hwnd]
                    continue

                if hwnd in window_allowed_path and window_allowed_path[hwnd] == folder_path:
                    continue

                if folder_path in grace_paths:
                    window_allowed_path[hwnd] = folder_path
                    continue

                if folder_path in pending_folders:
                    window.Quit()
                    continue
                window.Quit()
                pending_folders.add(folder_path)
                root.after(0, lambda fp=folder_path, pw=password: show_password_dialog(fp, pw))

            except Exception:
                pass
        for hwnd in list(window_allowed_path.keys()):

            if hwnd not in current_hwnds:
                del window_allowed_path[hwnd]
        root.after(100, check_windows)
    root.after(100, check_windows)
    root.mainloop()

class WindowIconSetter:

    def __init__(self, window_title, icon_path):
        self.window_title = window_title
        self.icon_path = os.path.abspath(icon_path) if icon_path else None
        self.user32 = ctypes.windll.user32
        self.shell32 = ctypes.windll.shell32

    def find_window_by_title(self, title):
        hwnd = self.user32.FindWindowW(None, title)

        if hwnd:
            return hwnd
        windows = []
        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def enum_windows_callback(hwnd, lParam):
            length = self.user32.GetWindowTextLengthW(hwnd) + 1
            buffer = ctypes.create_unicode_buffer(length)
            self.user32.GetWindowTextW(hwnd, buffer, length)

            if title in buffer.value:
                windows.append(hwnd)
            return True
        self.user32.EnumWindows(enum_windows_callback, 0)

        if windows:
            return windows[0]
        return None

    def set_icon(self):

        if not self.icon_path or not os.path.exists(self.icon_path):
            return False
        hwnd = None
        for i in range(50):
            hwnd = self.find_window_by_title(self.window_title)

            if hwnd:
                break
            time.sleep(0.1)

        if not hwnd:
            return False

        try:
            LR_LOADFROMFILE = 0x10
            IMAGE_ICON = 1
            hicon_small = self.user32.LoadImageW(
                0,
                self.icon_path,
                IMAGE_ICON,
                16, 16,
                LR_LOADFROMFILE
            )
            hicon_large = self.user32.LoadImageW(
                0,
                self.icon_path,
                IMAGE_ICON,
                32, 32,
                LR_LOADFROMFILE
            )
            WM_SETICON = 0x80
            ICON_SMALL = 0
            ICON_BIG = 1

            if hicon_small:
                self.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)

            if hicon_large:
                self.user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_large)
            return True

        except Exception as e:
            return False

class FolderProtectorGUI(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Folder Locker")
        self.geometry("800x700")
        self.resizable(False, False)
        self.config = load_config()
        self.protected_paths = set(self.config.get("protected_paths", []))
        self.data_dir = os.path.join(os.path.expanduser("~"), ".FolderLock&Hide")
        os.makedirs(self.data_dir, exist_ok=True)
        self.config_file = os.path.join(self.data_dir, "config.ini")
        self.password = self.config.get("password", "1234")
        self.setup_ui()
        self.load_window_geometry()
        self.refresh_treeview()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def load_window_geometry(self):

        if os.path.exists(self.config_file):
            config = configparser.ConfigParser()
            config.read(self.config_file)

            if "Geometry" in config:
                geometry = config["Geometry"].get("size", "")
                state = config["Geometry"].get("state", "normal")

                if geometry:
                    self.geometry(geometry)
                    self.update_idletasks()
                    self.update()

                if state == "zoomed":
                    self.state("zoomed")
                elif state == "iconic":
                    self.iconify()

    def save_window_geometry(self):
        config = configparser.ConfigParser()
        config["Geometry"] = {
            "size": self.geometry(),
            "state": self.state()
        }

        with open(self.config_file, "w") as f:
            config.write(f)

    def on_close(self):
        self.save_window_geometry()
        self.destroy()

    def setup_ui(self):
        self.lock = CTkImage(
            light_image=Image.open(resource_path(r"icons\lock.png")),
            size=(30, 30)
        )
        header = ctk.CTkLabel(self,image=self.lock, text="  FOLDER PASSWORD PROTECTOR",compound='left', font=("Segoe UI", 24, "bold"))
        header.pack(pady=(20,15))
        self.tabview = ctk.CTkTabview(self, width=750, height=550)
        self.tabview.pack(padx=20, pady=10, fill="both", expand=True)
        self.protect_tab = self.tabview.add("Protected Folders")
        self.settings_tab = self.tabview.add("Settings")
        self.protect_tab_icon = CTkImage(
            light_image=Image.open(resource_path(r"icons\protected.png")),
            size=(20, 20)
        )
        self.settings_tab_icon = CTkImage(
            light_image=Image.open(resource_path(r"icons\settings.png")),
            size=(20, 20)
        )
        buttons = self.tabview._segmented_button._buttons_dict
        buttons["Protected Folders"].configure(
            image=self.protect_tab_icon,
            compound="left",
        )
        buttons["Settings"].configure(
            image=self.settings_tab_icon,
            compound="left",
        )
        self.setup_protect_tab()
        self.setup_settings_tab()

    def setup_protect_tab(self):
        main_frame = ctk.CTkFrame(self.protect_tab)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        bg_color = ctk.ThemeManager.theme["CTkFrame"]["fg_color"][1]
        text_color = ctk.ThemeManager.theme["CTkLabel"]["text_color"][1]
        selected_color = ctk.ThemeManager.theme["CTkButton"]["fg_color"][1]
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background=bg_color,
                        foreground=text_color,
                        fieldbackground=bg_color,
                        borderwidth=0,
                        rowheight=30,
                        font=("Segoe UI", 11))
        style.map("Treeview",
                  background=[('selected', selected_color)],
                  foreground=[('selected', 'white')])
        style.configure("Treeview.Heading",
                        background="#1f538d",
                        foreground="white",
                        font=("Segoe UI", 12, "bold"),
                        borderwidth=0)
        style.map("Treeview.Heading",
                  background=[('active', '#2a6bb0')])
        tree_frame = ctk.CTkFrame(main_frame)
        tree_frame.pack(fill="both", expand=True, pady=5)
        columns = ("Folder Path",)
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                 selectmode="extended", height=15)
        self.tree.heading("Folder Path", text="Protected Folder Path")
        self.tree.column("Folder Path", width=550, anchor="w")
        tree_scrollbar = ctk.CTkScrollbar(tree_frame, command=self.tree.yview,
                                          corner_radius=8,
                                          button_color="#3a3a3a",
                                          button_hover_color="#555555")
        self.tree.configure(yscrollcommand=tree_scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scrollbar.pack(side="right", fill="y")
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=10)
        self.add_icon = CTkImage(light_image= Image.open(resource_path(r"icons\add.png")), size=(27, 27))
        self.remove_icon = CTkImage(light_image= Image.open(resource_path(r"icons\remove.png")), size=(27, 27))
        self.refresh_icon = CTkImage(light_image= Image.open(resource_path(r"icons\refresh.png")), size=(27, 27))
        ctk.CTkButton(btn_frame, image=self.add_icon, text=" Add Folder",
                    command=self.add_folder, width=140).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame,image=self.remove_icon, text=" Remove Selected", command=self.remove_folder, width=160).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame,image=self.refresh_icon, text=" Refresh", command=self.refresh_treeview, width=120).pack(side="left", padx=5)
        self.status_label = ctk.CTkLabel(main_frame, text="", font=("Segoe UI", 12))
        self.status_label.pack(pady=5)

    def setup_settings_tab(self):
        pw_frame = ctk.CTkFrame(self.settings_tab)
        pw_frame.pack(fill="x", padx=20, pady=20)
        ctk.CTkLabel(pw_frame, text="MASTER PASSWORD", font=("Segoe UI", 20, "bold")).pack(pady=(10,20))
        ctk.CTkLabel(pw_frame, text="Current Password (Default Password 1234):", anchor="w").pack(fill="x", padx=10)
        self.old_pass_entry = ctk.CTkEntry(pw_frame, show="•", width=300)
        self.old_pass_entry.pack(padx=10, pady=5, fill="x")
        ctk.CTkLabel(pw_frame, text="New Password:", anchor="w").pack(fill="x", padx=10)
        self.new_pass_entry = ctk.CTkEntry(pw_frame, show="•", width=300)
        self.new_pass_entry.pack(padx=10, pady=5, fill="x")
        ctk.CTkLabel(pw_frame, text="Confirm New Password:", anchor="w").pack(fill="x", padx=10)
        self.confirm_pass_entry = ctk.CTkEntry(pw_frame, show="•", width=300)
        self.confirm_pass_entry.pack(padx=10, pady=5, fill="x")
        self.show_pass_var = ctk.BooleanVar(value=False)

        def toggle_password_visibility():

            if self.show_pass_var.get():
                self.old_pass_entry.configure(show="")
                self.new_pass_entry.configure(show="")
                self.confirm_pass_entry.configure(show="")
            else:
                self.old_pass_entry.configure(show="•")
                self.new_pass_entry.configure(show="•")
                self.confirm_pass_entry.configure(show="•")
        show_cb = ctk.CTkCheckBox(
            pw_frame,
            text="Show passwords",
            variable=self.show_pass_var,
            command=toggle_password_visibility
        )
        show_cb.pack(pady=(5, 10), padx=10, anchor="w")
        ctk.CTkButton(pw_frame, text="Change Password", command=self.change_password, width=200).pack(pady=(10,5))
        info_frame = ctk.CTkFrame(self.settings_tab)
        info_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(info_frame, text="Application Information", font=("Segoe UI", 16, "bold")).pack(pady=(10,5))
        info_text = (
            "• SecureLock Pro v1.0\n\n"
            "• Gmail:   souravbhattacharya8159@gmail.com.\n"
            "• If you forgot your current password, then contact with me.\n"
            "• MOB:   +91 8159058135.\n"
        )
        ctk.CTkLabel(info_frame, text=info_text, justify="left", font=("Segoe UI", 11)).pack(padx=10, pady=5)

    def refresh_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for path in sorted(self.protected_paths):
            self.tree.insert("", "end", values=(path,))
        self.status_label.configure(text=f"Total protected folders: {len(self.protected_paths)}")

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select Folder to Protect")

        if not folder:
            return
        folder = os.path.abspath(folder)

        if folder in self.protected_paths:
            messagebox.showinfo("Already Protected", f"Folder is already in the protected list:\n{folder}")
            return
        self.protected_paths.add(folder)
        self.save_and_refresh()
        self.status_label.configure(text=f"Added: {folder}")

    def verify_master_password(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Password Required")
        dialog.geometry("450x200")
        dialog.resizable(False, False)
        dialog.attributes('-topmost', True)
        dialog.update()
        dialog.attributes('-topmost', False)
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        main_frame = ctk.CTkFrame(dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(main_frame, text="Enter Master Password", font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))
        pwd_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        pwd_frame.pack(pady=10)
        entry = ctk.CTkEntry(pwd_frame, show="*", width=250, font=("Arial", 12))
        entry.pack(side="left", padx=(0,5))
        entry.focus_set()
        show_img = CTkImage(light_image=Image.open(resource_path(r"icons\show_password.png")), size=(20,20))
        hide_img = CTkImage(light_image=Image.open(resource_path(r"icons\hide_password.png")), size=(20,20))
        show_var = False

        def toggle_show():
            nonlocal show_var

            if show_var:
                entry.configure(show="*")
                show_var = False
                eye_btn.configure(image=hide_img)
            else:
                entry.configure(show="")
                show_var = True
                eye_btn.configure(image=show_img)
        eye_btn = ctk.CTkButton(pwd_frame, text="", image=hide_img, width=40, command=toggle_show)
        eye_btn.pack(side="right")
        result = False

        def check():
            nonlocal result

            if entry.get() == self.password:
                result = True
                dialog.destroy()
            else:
                messagebox.showerror("Wrong Password", "Incorrect master password.")

        def on_cancel():
            nonlocal result
            result = False
            dialog.destroy()
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="OK", command=check, width=100).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=on_cancel, width=100).pack(side="left", padx=5)
        dialog.bind("<Return>", lambda event: check())
        dialog.bind("<Escape>", lambda event: on_cancel())
        dialog.grab_set()
        dialog.wait_window()
        return result

    def remove_folder(self):
        selected = self.tree.selection()

        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one folder to remove.")
            return

        if not self.verify_master_password():
            return
        folders_to_remove = [self.tree.item(item, "values")[0] for item in selected]

        if len(folders_to_remove) == 1:
            msg = f"Remove protection from:\n{folders_to_remove[0]}?"
        else:
            msg = f"Remove protection from {len(folders_to_remove)} folders?"

        if not messagebox.askyesno("Confirm Removal", msg):
            return
        for folder in folders_to_remove:
            self.protected_paths.discard(folder)
        self.save_and_refresh()
        self.status_label.configure(text=f"Removed {len(folders_to_remove)} folder(s).")

    def change_password(self):
        old = self.old_pass_entry.get()
        new = self.new_pass_entry.get()
        confirm = self.confirm_pass_entry.get()

        if not old:
            messagebox.showerror("Error", "Please enter current password.")
            return

        if old != self.password:
            messagebox.showerror("Error", "Current password is incorrect.")
            return

        if not new:
            messagebox.showerror("Error", "New password cannot be empty.")
            return

        if new != confirm:
            messagebox.showerror("Error", "New password and confirmation do not match.")
            return

        if len(new) < 4:
            messagebox.showerror("Error", "Password must be at least 4 characters.")
            return
        self.password = new
        self.save_and_refresh()
        self.old_pass_entry.delete(0, 'end')
        self.new_pass_entry.delete(0, 'end')
        self.confirm_pass_entry.delete(0, 'end')
        self.status_label.configure(text="Password changed successfully")
        messagebox.showinfo("Success", "Master password updated.")

    def save_and_refresh(self):
        self.config["password"] = self.password
        self.config["protected_paths"] = list(self.protected_paths)
        save_config(self.config)
        self.refresh_treeview()

def is_monitor_running():

    try:
        mutex = ctypes.windll.kernel32.OpenMutexW(0x100000, False, MUTEX_NAME)

        if mutex:
            ctypes.windll.kernel32.CloseHandle(mutex)
            return True

    except:
        pass
    return False

if __name__ == "__main__":
    multiprocessing.freeze_support() 

    if sys.platform == "win32":

        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(2)

        except:

            try:
                windll.user32.SetProcessDPIAware()

            except:
                pass
    icon_path = resource_path(r"icons\icon.ico")
    icon_setter = WindowIconSetter("Folder Locker", icon_path)

    def apply_icon_async():
        time.sleep(1)
        icon_setter.set_icon()
    threading.Thread(target=apply_icon_async, daemon=True).start()

    if not is_monitor_running():
        monitor_process = multiprocessing.Process(target=run_monitor, daemon=False)
        monitor_process.start()
    app = FolderProtectorGUI()
    app.mainloop()
