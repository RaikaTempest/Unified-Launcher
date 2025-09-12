import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk, ImageOps, ImageDraw
import os
import shutil
import tempfile
import threading
import sys
import subprocess
import json
import base64
from io import BytesIO
import pandas as pd
import html  # for escaping notes/text in HTML export
from queue import Queue  # NEW: async handoff queue
# --- Hide console on Windows (when launched by double-click or from a shortcut) ---
# --- Hide console on Windows safely (do NOT detach) ---
def _hide_console_on_windows():
    try:
        import sys
        if sys.platform.startswith("win"):
            import ctypes
            # If we're running inside an interactive console (cmd/PowerShell), keep it attached.
            # Only hide if there *is* a console window and we're likely double-clicked/launched by shortcut.
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass
# --- Make print safe when no console is attached (e.g., pythonw or hidden console) ---
def _install_safe_print():
    import builtins, os, tempfile
    if getattr(builtins, "_orig_print", None):
        return  # already installed

    builtins._orig_print = builtins.print

    def _safe_print(*args, **kwargs):
        try:
            builtins._orig_print(*args, **kwargs)
        except Exception:
            # Fallback to a logfile in %TEMP% rather than crashing with WinError 6
            try:
                msg = " ".join(str(a) for a in args)
                with open(os.path.join(tempfile.gettempdir(), "upr_log.txt"), "a", encoding="utf-8") as fh:
                    fh.write(msg + "\n")
            except Exception:
                pass

    builtins.print = _safe_print

_install_safe_print()



# --- HEIC/HEIF Image Support ---
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_SUPPORT = True
except ImportError:
    HEIC_SUPPORT = False

# --- Data Processing Logic ---
def process_data(main_file_path, pn_file_path, dr_file_path):
    """
    Reads, processes, and merges data from the main barcode sheet and the PN/DR files.
    This version is optimized to handle multiple matches per barcode.
    """
    try:
        print("--- Reading Input Files ---")
        # Ensure Barcode columns are read as strings to preserve leading zeros
        dtype_spec = {'Barcode ID': str, 5: str, 0: str}
        
        print(f"--> Loading main file: {os.path.basename(main_file_path)}...")
        main_df = pd.read_csv(main_file_path, header=0, dtype=dtype_spec) if main_file_path.lower().endswith('.csv') else pd.read_excel(main_file_path, header=0, dtype=dtype_spec)
        print("    ...Done.")

        print(f"--> Loading PN file: {os.path.basename(pn_file_path)}...")
        pn_df = pd.read_excel(pn_file_path, header=None, dtype=dtype_spec)
        print("    ...Done.")

        print(f"--> Loading DR file: {os.path.basename(dr_file_path)}...")
        dr_df = pd.read_excel(dr_file_path, header=None, dtype=dtype_spec)
        print("    ...Done.")
        
        request_col_name = next((col for col in main_df.columns if 'request' in str(col).lower()), None)
        
        required_cols = ['Barcode ID', 'Latitude', 'Longitude']
        new_col_names = ['Barcode', 'Latitude', 'Longitude']
        if request_col_name:
            required_cols.append(request_col_name)
            new_col_names.append('Request')

        main_data = main_df[required_cols].copy()
        main_data.columns = new_col_names
        main_data.dropna(subset=['Barcode'], inplace=True)
        main_data.drop_duplicates(subset=['Barcode'], keep='first', inplace=True)

        pn_data = pn_df[[5, 0, 2, 4]].copy()
        pn_data.columns = ['Barcode', 'ID', 'Info', 'Location']
        pn_data.dropna(subset=['Barcode'], inplace=True)
        pn_data['Type'] = 'PN'
        pn_data['Requirement'] = 'Required'

        dr_data = dr_df[[0, 5, 3, 4, 8, 12]].copy()
        dr_data.columns = ['Barcode', 'ID', 'Info_Part1', 'Info_Part2', 'Location', 'Requirement']
        dr_data['Info'] = dr_data['Info_Part1'].fillna('').astype(str) + ' ' + dr_data['Info_Part2'].fillna('').astype(str)
        dr_data.drop(columns=['Info_Part1', 'Info_Part2'], inplace=True)
        dr_data['Info'] = dr_data['Info'].str.strip()
        dr_data.dropna(subset=['Barcode'], inplace=True)
        dr_data['Type'] = 'DR'
        
        combined_data = pd.concat([pn_data, dr_data], ignore_index=True)

        print("\n--- Merging Barcode Data ---")
        merged_df = pd.merge(main_data, combined_data, on='Barcode', how='left')
        merged_df['Type'].fillna('N/A', inplace=True)
        for col in ['ID', 'Info', 'Location', 'Requirement']:
             merged_df[col].fillna('', inplace=True)

        print("\nData processing complete.")
        return merged_df

    except FileNotFoundError as e:
        messagebox.showerror("File Not Found", f"Error: A required file was not found.\n{e}")
        return None
    except KeyError as e:
        messagebox.showerror("Column Not Found", f"Could not find required column: {e}.\nPlease ensure your main file has 'Barcode ID', 'Latitude', and 'Longitude'.")
        return None
    except Exception as e:
        messagebox.showerror("Processing Error", f"An unexpected error occurred during data processing:\n{e}")
        import traceback
        traceback.print_exc()
        return None

class PoleReviewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unified Pole Photo Reviewer")
        self.root.geometry("1800x1000")
        self.root.minsize(1200, 700)

        # --- Dark Mode Color Palette ---
        self.COLOR_BG = "#2e2e2e"
        self.COLOR_BG_LIGHT = "#3c3c3c"
        self.COLOR_BG_DARK = "#252525"
        self.COLOR_FG = "#dcdcdc"
        self.COLOR_ACCENT = "#007acc"
        self.COLOR_BORDER = "#4a4a4a"

        self._setup_styles()

        # --- Async image loading state & caches ---
        self.thumb_cache = {}      # path -> PIL.Image (thumbnail-sized)
        self.large_cache = {}      # (path, canvas_w, canvas_h) -> (PIL.Image resized, new_w, new_h)
        self.load_queue = Queue()  # background -> main-thread handoff
        self.nav_token = 0         # bump on pole/photo change to ignore stale results

        # Skeleton placeholder for thumbnails (120x120 grey)
        self.placeholder_thumb = ImageTk.PhotoImage(Image.new('RGB', (120, 120), (60, 60, 60)))

        # --- Data Storage ---
        self.poles_data = {}
        self.lookup_data = None
        self.current_pole_id = None
        self.current_photo_index = 0
        self.photo_references = []
        self.original_parent_folder = None
        self.temp_dir = tempfile.mkdtemp(prefix="pole_reviewer_")
        self.lookup_sources = None  # {'main': <path>, 'pn': <path>, 'dr': <path>}

        
        # --- Markup State ---
        self.markup_mode = tk.BooleanVar(value=False)
        self.start_x = None
        self.start_y = None
        self.current_oval = None
        self.displayed_image_info = {} # To store scale and offset of the image on canvas

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._create_menu()
        self._create_main_layout()
        self._setup_key_bindings()

        # Start the queue pump
        self.root.after(30, self._process_load_queue)

        if not HEIC_SUPPORT:
            messagebox.showwarning(
                "HEIC Support Missing",
                "The 'pillow-heif' library is not found. HEIC files will be ignored.\nTo add support, run: pip install pillow-heif"
            )

    # ---------- Unified progress dialog (status + bar) ----------
    def _open_progress(self, title="Working..."):
        # Create once and reuse
        if getattr(self, "_progress_win", None) and self._progress_win.winfo_exists():
            try:
                self._progress_win.title(title)
            except Exception:
                pass
            return

        self._progress_win = tk.Toplevel(self.root)
        self._progress_win.title(title)
        self._progress_win.geometry("420x140")
        self._progress_win.transient(self.root)
        self._progress_win.grab_set()
        try:
            self._progress_win.configure(bg=self.COLOR_BG)
        except Exception:
            pass

        frm = ttk.Frame(self._progress_win, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)
        self._progress_label = ttk.Label(frm, text="Starting…")
        self._progress_label.pack(pady=(4, 10), anchor="w")
        self._progress_bar = ttk.Progressbar(frm, orient="horizontal", length=360, mode="determinate")
        self._progress_bar.pack(pady=(0, 6))
        self._progress_note = ttk.Label(frm, text="", foreground="#9fb1c1")
        self._progress_note.pack(anchor="w")

    def _set_progress_status(self, text, note=""):
        if getattr(self, "_progress_win", None) and self._progress_win.winfo_exists():
            self._progress_label.config(text=text)
            self._progress_note.config(text=note)
            self._progress_win.update_idletasks()

    def _set_progress_max(self, maximum):
        if getattr(self, "_progress_win", None) and self._progress_win.winfo_exists():
            self._progress_bar.config(mode="determinate")
            self._progress_bar["maximum"] = max(1, int(maximum))
            self._progress_bar["value"] = 0
            self._progress_win.update_idletasks()

    def _step_progress(self, value):
        if getattr(self, "_progress_win", None) and self._progress_win.winfo_exists():
            self._progress_bar["value"] = int(value)
            self._progress_win.update_idletasks()

    def _close_progress(self):
        if getattr(self, "_progress_win", None) and self._progress_win.winfo_exists():
            try:
                self._progress_win.destroy()
            except Exception:
                pass


    def _setup_styles(self):
        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')
        self.style.configure('.', background=self.COLOR_BG, foreground=self.COLOR_FG, fieldbackground=self.COLOR_BG_LIGHT, bordercolor=self.COLOR_BORDER)
        self.style.map('.', background=[('active', self.COLOR_BG_LIGHT)])
        self.root.configure(bg=self.COLOR_BG)

        self.style.configure("TButton", padding=6, relief="flat", font=('Helvetica', 10), background=self.COLOR_BG_LIGHT)
        self.style.map("TButton", background=[('active', self.COLOR_ACCENT)])
        self.style.configure("TLabel", font=('Helvetica', 10), background=self.COLOR_BG, foreground=self.COLOR_FG)
        self.style.configure("TFrame", background=self.COLOR_BG)
        self.style.configure("Header.TLabel", font=('Helvetica', 16, "bold"), background=self.COLOR_BG, foreground=self.COLOR_FG)
        self.style.configure("TPanedwindow", background=self.COLOR_BG)
        self.style.configure("TLabelFrame", background=self.COLOR_BG, bordercolor=self.COLOR_BORDER)
        self.style.configure("TLabelFrame.Label", font=('Helvetica', 11, 'bold'), background=self.COLOR_BG, foreground=self.COLOR_FG)
        self.style.configure("TCheckbutton", font=('Helvetica', 10), background=self.COLOR_BG, foreground=self.COLOR_FG)
        self.style.map("TCheckbutton", background=[('active', self.COLOR_BG)], indicatorcolor=[('selected', self.COLOR_ACCENT), ('!selected', self.COLOR_BG_LIGHT)])
        self.style.configure("Treeview", rowheight=25, fieldbackground=self.COLOR_BG_LIGHT, background=self.COLOR_BG_LIGHT, foreground=self.COLOR_FG)
        self.style.map("Treeview", background=[('selected', self.COLOR_ACCENT)])
        self.style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold'), background=self.COLOR_BG_DARK, foreground=self.COLOR_FG)
        self.style.configure("Vertical.TScrollbar", background=self.COLOR_BG_DARK, troughcolor=self.COLOR_BG, bordercolor=self.COLOR_BG, arrowcolor=self.COLOR_FG)
        self.style.map("Vertical.TScrollbar", background=[('active', self.COLOR_BG_LIGHT)])

    def _create_menu(self):
        menu_bar = tk.Menu(self.root, bg=self.COLOR_BG_DARK, fg=self.COLOR_FG, activebackground=self.COLOR_ACCENT, activeforeground=self.COLOR_FG, relief=tk.FLAT)
        self.root.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0, bg=self.COLOR_BG_LIGHT, fg=self.COLOR_FG, activebackground=self.COLOR_ACCENT)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Review Session...", command=self.start_new_review_flow)
        file_menu.add_command(label="Load Review...", command=self.load_review)
        file_menu.add_command(label="Save Review...", command=self.save_review)
        file_menu.add_separator()
        file_menu.add_command(label="Export HTML Report", command=self.export_to_html)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing)

    def _create_main_layout(self):
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.paned_window = ttk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        self.left_frame = ttk.Frame(self.paned_window, width=350)
        self.left_frame.pack_propagate(False)
        self.paned_window.add(self.left_frame, weight=1)

        self.right_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.right_frame, weight=4)

        self._create_left_widgets()
        self._create_right_widgets()

    def _create_left_widgets(self):
        controls_frame = ttk.Frame(self.left_frame, padding="10")
        controls_frame.pack(fill=tk.X)

        nav_frame = ttk.Frame(controls_frame)
        nav_frame.pack(fill=tk.X, pady=5)
        ttk.Button(nav_frame, text="◀ Previous Pole", command=self.prev_pole).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        ttk.Button(nav_frame, text="Next Pole ▶", command=self.next_pole).pack(side=tk.LEFT, expand=True, fill=tk.X)

        ttk.Button(controls_frame, text="Export HTML Report", command=self.export_to_html).pack(fill=tk.X, pady=(10,0))

        list_frame = ttk.LabelFrame(self.left_frame, text="Pole Locations", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.pole_list = DraggableCheckboxListbox(list_frame, self.COLOR_BG_DARK, self.COLOR_ACCENT, command=self.on_pole_select)
        self.pole_list.pack(fill=tk.BOTH, expand=True)

    def _create_right_widgets(self):
        self.right_paned_window = ttk.PanedWindow(self.right_frame, orient=tk.VERTICAL)
        self.right_paned_window.pack(fill=tk.BOTH, expand=True)

        data_pane = ttk.Frame(self.right_paned_window)
        self.right_paned_window.add(data_pane, weight=2)

        viewer_pane = ttk.Frame(self.right_paned_window)
        self.right_paned_window.add(viewer_pane, weight=3)

        self._create_data_pane_widgets(data_pane)
        self._create_viewer_pane_widgets(viewer_pane)

    def _create_data_pane_widgets(self, parent):
        # Horizontal paned window so the user can resize Checklist | Notes | Lookup
        data_panes = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        data_panes.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # ==== Pane 1: Checklist ===================================================
        checklist_pane = ttk.Frame(data_panes)
        data_panes.add(checklist_pane, weight=3)  # give it some flexible weight

        checklist_container = ttk.LabelFrame(checklist_pane, text="Assessment Checklist", padding="10")
        checklist_container.pack(fill=tk.BOTH, expand=True)

        # Scrollable checklist (canvas + inner frame)
        self.checklist_canvas = tk.Canvas(checklist_container, highlightthickness=0, bg=self.COLOR_BG)
        checklist_scrollbar = ttk.Scrollbar(checklist_container, orient="vertical", command=self.checklist_canvas.yview)
        self.checklist_canvas.configure(yscrollcommand=checklist_scrollbar.set)
        self.checklist_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        checklist_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.checklist_frame = ttk.Frame(self.checklist_canvas, style="TFrame")
        self.checklist_canvas.create_window((0, 0), window=self.checklist_frame, anchor="nw")
        self.checklist_frame.bind("<Configure>", lambda e: self.checklist_canvas.configure(scrollregion=self.checklist_canvas.bbox("all")))

        # Mouse wheel (if you added the helper earlier)
        if hasattr(self, "_enable_mousewheel"):
            self._enable_mousewheel(self.checklist_canvas)
            self._enable_mousewheel(self.checklist_frame, target=self.checklist_canvas)

        # ==== Pane 2: Notes =======================================================
        notes_pane = ttk.Frame(data_panes)
        data_panes.add(notes_pane, weight=2)

        notes_container = ttk.LabelFrame(notes_pane, text="Additional Notes", padding="10")
        notes_container.pack(fill=tk.BOTH, expand=True)

        self.notes_text = tk.Text(
            notes_container, height=4, wrap="word", relief=tk.FLAT, font=('Helvetica', 10),
            bg=self.COLOR_BG_LIGHT, fg=self.COLOR_FG, insertbackground=self.COLOR_FG, borderwidth=0
        )
        notes_scrollbar = ttk.Scrollbar(notes_container, orient="vertical", command=self.notes_text.yview)
        self.notes_text.configure(yscrollcommand=notes_scrollbar.set)
        self.notes_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        notes_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.notes_text.bind("<<Modified>>", self.on_notes_modified)

        if hasattr(self, "_enable_mousewheel"):
            self._enable_mousewheel(self.notes_text)

        # ==== Pane 3: Lookup ======================================================
        lookup_pane = ttk.Frame(data_panes)
        data_panes.add(lookup_pane, weight=3)

        lookup_container = ttk.LabelFrame(lookup_pane, text="Pole Lookup Data", padding="10")
        lookup_container.pack(fill=tk.BOTH, expand=True)

        cols = ('Type', 'ID', 'Info', 'Location', 'Requirement')
        self.lookup_tree = ttk.Treeview(lookup_container, columns=cols, show='headings', style="Treeview")
        for col in cols:
            self.lookup_tree.heading(col, text=col)
            self.lookup_tree.column(col, width=100, anchor='w')

        tree_scrollbar = ttk.Scrollbar(lookup_container, orient="vertical", command=self.lookup_tree.yview)
        self.lookup_tree.configure(yscrollcommand=tree_scrollbar.set)
        self.lookup_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        if hasattr(self, "_enable_mousewheel"):
            self._enable_mousewheel(self.lookup_tree)

        # Keep your checklist item definitions + dict init (unchanged)
        self.checklist_items_defs = [
            ('pole_condition', 'Pole Condition (Rot, Damage)'),
            ('pole_clearance', 'Pole Clearance'),
            ('pole_loading', 'Pole Loading'),
            ('crossarm_condition', 'Crossarm Condition'),
            ('insulator_condition', 'Insulator Condition'),
            ('transformer_leaks', 'Transformer Leaks'),
            ('vegetation_clearance', 'Vegetation Clearance'),
            ('ground_wire', 'Ground Wire Secure'),
            ('equipment_attached', 'Foreign Equipment Attached'),
            ('accessibility', 'Site Accessibility'),
            ('down_guy_condition', 'Down Guy Condition'),
            ('down_guy_guard', 'Down Guy Guard'),
        ]
        self.checklist_vars = {}


    def _create_viewer_pane_widgets(self, parent):
        photo_paned_window = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        photo_paned_window.pack(fill=tk.BOTH, expand=True)

        # Left: large photo + controls
        large_photo_area = ttk.Frame(photo_paned_window)
        photo_paned_window.add(large_photo_area, weight=4)

        self.pole_name_label = ttk.Label(large_photo_area, text="Select a pole to begin", style="Header.TLabel")
        self.pole_name_label.pack(pady=10, padx=10, anchor="w")

        self.photo_canvas = tk.Canvas(large_photo_area, background=self.COLOR_BG_DARK, highlightthickness=0)
        self.photo_canvas.pack(fill=tk.BOTH, expand=True, padx=10)
        self.photo_canvas.bind("<ButtonPress-1>", self.on_canvas_press)
        self.photo_canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.photo_canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.photo_canvas.bind("<Double-Button-1>", self.open_current_photo_externally)

        markup_controls_frame = ttk.Frame(large_photo_area)
        markup_controls_frame.pack(fill=tk.X, pady=5, padx=10)
        ttk.Checkbutton(markup_controls_frame, text="Toggle Markup", variable=self.markup_mode, style="TCheckbutton").pack(side=tk.LEFT, padx=(0, 20))
        ttk.Button(markup_controls_frame, text="Save Markups", command=self.save_markups).pack(side=tk.LEFT, padx=5)
        ttk.Button(markup_controls_frame, text="Clear Markups", command=self.clear_temporary_markups).pack(side=tk.LEFT, padx=5)

        photo_nav_frame = ttk.Frame(large_photo_area)
        photo_nav_frame.pack(fill=tk.X, pady=5, padx=10)
        ttk.Button(photo_nav_frame, text="◀ Previous Photo", command=self.prev_photo).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,5))
        ttk.Button(photo_nav_frame, text="Next Photo ▶", command=self.next_photo).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5,0))

        # Right: thumbnails
        thumbnail_area = ttk.Frame(photo_paned_window, width=180)
        thumbnail_area.pack_propagate(False)
        photo_paned_window.add(thumbnail_area, weight=1)
        thumbnail_container = ttk.LabelFrame(thumbnail_area, text="Photos", padding="5")
        thumbnail_container.pack(fill=tk.BOTH, expand=True, padx=(5,10), pady=(10,10))
        self.thumbnail_canvas = tk.Canvas(thumbnail_container, bg=self.COLOR_BG, highlightthickness=0)
        thumb_scrollbar = ttk.Scrollbar(thumbnail_container, orient="vertical", command=self.thumbnail_canvas.yview)
        self.thumbnail_canvas.configure(yscrollcommand=thumb_scrollbar.set)
        thumb_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.thumbnail_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.thumbnail_frame = ttk.Frame(self.thumbnail_canvas, style="TFrame")
        self.thumbnail_canvas.create_window((0, 0), window=self.thumbnail_frame, anchor="nw")
        self.thumbnail_frame.bind("<Configure>", lambda e: self.thumbnail_canvas.configure(scrollregion=self.thumbnail_canvas.bbox("all")))
        self._enable_mousewheel(self.thumbnail_canvas)
        self._enable_mousewheel(self.thumbnail_frame, target=self.thumbnail_canvas)  # wheel over child frame too


        self.style.configure("SelectedThumbnail.TFrame", bordercolor=self.COLOR_ACCENT, background=self.COLOR_BG)
        self.style.configure("Thumbnail.TFrame", padding=4, relief="solid", borderwidth=2, bordercolor="transparent", background=self.COLOR_BG)

    def _setup_key_bindings(self):
        self.root.bind('<KeyPress>', self._on_key_press)
    
    def _enable_mousewheel(self, widget, *, target=None):
        """
        Attach OS-friendly mouse-wheel scrolling to a widget.
        If target is provided, we call target.yview_scroll; else try widget.yview_scroll.
        """
        tv = target or widget
        def on_wheel(e):
            # Windows/macOS normalize to 1 unit per notch
            delta = -1 if getattr(e, "delta", 0) > 0 else 1
            try:
                tv.yview_scroll(delta, "units")
            except Exception:
                pass
            return "break"

        # Windows/macOS
        widget.bind("<MouseWheel>", on_wheel)
        # Linux
        widget.bind("<Button-4>", lambda e: (tv.yview_scroll(-1, "units"), "break"))
        widget.bind("<Button-5>", lambda e: (tv.yview_scroll( 1, "units"), "break"))


    def _on_key_press(self, event):
        focused_widget = self.root.focus_get()
        if event.keysym == 'Left':
            self.prev_photo()
        elif event.keysym == 'Right':
            self.next_photo()
        elif event.keysym == 'Up' and focused_widget != self.notes_text:
            self.prev_pole()
        elif event.keysym == 'Down' and focused_widget != self.notes_text:
            self.next_pole()

    # --- Background results pump ---
    # In the PoleReviewerApp class

    # In the PoleReviewerApp class

    def _process_load_queue(self):
        try:
            while True:
                kind, token, payload = self.load_queue.get_nowait()
                # Ignore results from old, slow threads
                if token != self.nav_token:
                    continue

                if kind == 'thumb':
                    # CORRECT: Unpack the payload into the image and the direct widget reference
                    pil_img, label_widget = payload
                    try:
                        # Check if the widget still exists before trying to update it
                        if label_widget.winfo_exists():
                            photo_img = ImageTk.PhotoImage(pil_img)
                            label_widget.configure(image=photo_img)
                            label_widget.image = photo_img # Prevent garbage collection
                            label_widget.update_idletasks()
                    except Exception:
                        # This can happen if the user navigates away very quickly.
                        # It's safe to ignore as the widget is now gone.
                        pass

                elif kind == 'large':
                    path, pil_img = payload
                    try:
                        tk_img = ImageTk.PhotoImage(pil_img)
                        self.displayed_image_info['tk_photo'] = tk_img
                        self.photo_canvas.delete("all")
                        self.photo_canvas.create_image(
                            self.displayed_image_info.get('offset_x', 0),
                            self.displayed_image_info.get('offset_y', 0),
                            anchor='nw',
                            image=tk_img,
                            tags="photo"
                        )
                        self.draw_temporary_markups()
                    except Exception as e:
                        print(f"Large UI update error for {path}: {e}")
        except Exception:
            # Queue is empty, which is normal
            pass
        finally:
            self.root.after(30, self._process_load_queue)

    # --- File/session flows ---
    def start_new_review_flow(self):
        parent_folder = filedialog.askdirectory(title="Select Parent Folder of Pole Photos")
        if not parent_folder:
            return

        main_data_file = filedialog.askopenfilename(
            title="Select the Main Barcode Sheet (Excel/CSV)",
            filetypes=(("Supported Files", "*.xlsx *.xls *.csv"), ("All files", "*.*"))
        )
        if not main_data_file:
            return

        self.original_parent_folder = parent_folder
        script_dir = os.path.dirname(os.path.realpath(__file__))
        pn_file = os.path.join(script_dir, 'PN.xlsx')
        dr_file = os.path.join(script_dir, 'DR.xlsx')
        self.lookup_sources = {'main': main_data_file, 'pn': pn_file, 'dr': dr_file}

        # Progress UI for data load + file copy
        self._open_progress("Preparing review session")
        self._set_progress_status("Loading lookup data (PN & DR)…", note=os.path.basename(main_data_file))

        # Build lookup (show friendly sub-steps even if process_data doesn't expose them)
        try:
            # If your process_data can accept a callback, you could pass one here to update finer-grained steps.
            # For now we just show coarse status.
            self.lookup_data = process_data(main_data_file, pn_file, dr_file)
        except Exception as e:
            self._close_progress()
            messagebox.showerror("Lookup Error", f"Failed to load lookup data:\n{e}")
            return

        if self.lookup_data is None:
            self._close_progress()
            messagebox.showerror("Lookup Error", "Lookup data could not be created.")
            return

        # Hand off to photo copying using the SAME window
        self._set_progress_status("Copying images locally…")
        self.start_photo_review(parent_folder)



    def start_photo_review(self, folder_path, saved_data=None):
        self.poles_data.clear()
        self.pole_list.clear()
        self.clear_right_panel()
        copy_thread = threading.Thread(target=self.copy_files_with_progress, args=(folder_path, saved_data))
        copy_thread.start()

    def copy_files_with_progress(self, source_folder, saved_data=None):
        # Reuse existing progress window; create if needed
        self._open_progress("Copying Files")

        valid_extensions = ('.png', '.jpg', '.jpeg') + (('.heic',) if HEIC_SUPPORT else ())
        files_to_copy = [
            os.path.join(r, f)
            for r, _, fs in os.walk(source_folder)
            for f in fs if f.lower().endswith(valid_extensions)
        ]
        if not files_to_copy:
            self._close_progress()
            self.root.after(0, lambda: messagebox.showwarning("No Images", "No supported image files found."))
            return

        self._set_progress_status("Copying images locally…")
        self._set_progress_max(len(files_to_copy))

        for i, src_path in enumerate(files_to_copy):
            relative_path = os.path.relpath(src_path, source_folder)
            dest_path = os.path.join(self.temp_dir, relative_path)
            os.makedirs(os.path.dirname(dest_path), exist_ok=True)
            try:
                shutil.copy(src_path, dest_path)
            except Exception as e:
                print(f"Copy failed for {src_path}: {e}")
            self._step_progress(i + 1)

        # Close progress and continue loading into UI
        self._close_progress()
        self.root.after(0, self.load_data_from_temp, saved_data)


    def load_data_from_temp(self, saved_data=None):
        valid_extensions = ('.png', '.jpg', '.jpeg') + (('.heic',) if HEIC_SUPPORT else ())
        for item in os.listdir(self.temp_dir):
            item_path = os.path.join(self.temp_dir, item)
            if os.path.isdir(item_path):
                pole_id = item
                linked_data_list = []
                if self.lookup_data is not None:
                    pole_lookup_info = self.lookup_data[self.lookup_data['Barcode'] == pole_id]
                    if not pole_lookup_info.empty:
                        linked_data_list = pole_lookup_info.to_dict('records')
                    else:
                        print(f"Warning: Pole ID '{pole_id}' not found in the lookup data sheet.")

                self.poles_data[pole_id] = {
                    "path": item_path,
                    "photos": [],
                    "checklist": {key: tk.BooleanVar() for key, _ in self.checklist_items_defs},
                    "notes": "",
                    "reviewed": tk.BooleanVar(),
                    "lookup_info": linked_data_list
                }

                # --- MODIFICATION START ---
                # Intelligently pair original and marked-up photos
                all_photo_files = [f for f in os.listdir(item_path) if f.lower().endswith(valid_extensions)]
                
                # Separate marked files from original files
                marked_files = {f for f in all_photo_files if f.startswith('marked_')}
                original_files = {f for f in all_photo_files if not f.startswith('marked_')}
                
                # Create a lookup for marked files based on their original name
                marked_lookup = {f.replace('marked_', ''): f for f in marked_files}

                for original_name in sorted(list(original_files)):
                    photo_entry = {
                        "original": os.path.join(item_path, original_name),
                        "marked_up": None,
                        "markups": [] # This holds temporary, unsaved on-canvas ovals
                    }
                    
                    # Check if a corresponding marked-up version exists
                    if original_name in marked_lookup:
                        marked_name = marked_lookup[original_name]
                        photo_entry["marked_up"] = os.path.join(item_path, marked_name)
                    
                    self.poles_data[pole_id]["photos"].append(photo_entry)
                # --- MODIFICATION END ---
        
        pole_items = [(pole_id, data['reviewed']) for pole_id, data in sorted(self.poles_data.items())]
        
        if saved_data:
            self.apply_saved_review_data(saved_data)
            saved_order = saved_data.get('pole_order')
            if saved_order:
                pole_items.sort(key=lambda x: saved_order.index(x[0]) if x[0] in saved_order else float('inf'))

        self.pole_list.populate(pole_items)
        if self.poles_data:
            messagebox.showinfo("Success", f"Finished loading {len(self.poles_data)} pole folders.")
            if self.pole_list.get_item_count() > 0:
                self.pole_list.select_item(0)
        else:
            messagebox.showwarning("No Folders Found", "No sub-folders with images were found.")

    def apply_saved_review_data(self, saved_data):
        for pole_id, state in saved_data.get('poles', {}).items():
            if pole_id in self.poles_data:
                self.poles_data[pole_id]['reviewed'].set(state.get('reviewed', False))
                self.poles_data[pole_id]['notes'] = state.get('notes', '')
                for key, value in state.get('checklist', {}).items():
                    if key in self.poles_data[pole_id]['checklist']:
                        self.poles_data[pole_id]['checklist'][key].set(value)
                for i, photo_entry in enumerate(self.poles_data[pole_id]['photos']):
                    saved_photo_state = state.get('photos', [])
                    if i < len(saved_photo_state):
                        saved_photo = saved_photo_state[i]
                        if saved_photo.get('marked_up'):
                            base_name = os.path.basename(saved_photo['marked_up'])
                            photo_entry['marked_up'] = os.path.join(self.poles_data[pole_id]['path'], base_name)
                        photo_entry['markups'] = saved_photo.get('markups', [])

    # --- Navigation ---
    def on_pole_select(self, pole_id):
        self.current_pole_id = pole_id
        self.display_pole_data()
    
    def navigate_pole(self, direction):
        if not self.current_pole_id:
            return
        current_index = self.pole_list.get_selected_index()
        if current_index is None:
            return
        self.poles_data[self.pole_list.items[current_index][0]]['reviewed'].set(True)
        new_index = (current_index + direction) % self.pole_list.get_item_count()
        self.pole_list.select_item(new_index)

    def next_pole(self): self.navigate_pole(1)
    def prev_pole(self): self.navigate_pole(-1)

    def display_pole_data(self):
        if not self.current_pole_id:
            return
        # Invalidate any in-flight image loads for prior selection
        self.nav_token += 1

        self.pole_name_label.config(text=self.current_pole_id)
        pole_data = self.poles_data[self.current_pole_id]
        self.display_thumbnails(pole_data["photos"])
        self.display_checklist(pole_data["checklist"])
        self.display_notes(pole_data["notes"])
        self.display_lookup_data(pole_data.get("lookup_info", []))

        if pole_data["photos"]:
            self.set_photo(0)
        else:
            self.clear_photo_canvas()

    def display_lookup_data(self, lookup_list):
        for i in self.lookup_tree.get_children():
            self.lookup_tree.delete(i)
        for record in lookup_list:
            values = (
                record.get('Type', 'N/A'),
                record.get('ID', ''),
                record.get('Info', ''),
                record.get('Location', ''),
                record.get('Requirement', '')
            )
            self.lookup_tree.insert('', 'end', values=values)

    # --- Thumbnails (async) ---
    # In the PoleReviewerApp class

    # In the PoleReviewerApp class

    def display_thumbnails(self, photo_entries):
        for widget in self.thumbnail_frame.winfo_children():
            widget.destroy()
        self.photo_references.clear()

        thumb_size = (120, 120)
        # This token invalidates any slower threads from a previous selection
        self.nav_token += 1
        local_token = self.nav_token

        for i, photo_entry in enumerate(photo_entries):
            path = photo_entry.get('marked_up') or photo_entry['original']
            thumb_frame = ttk.Frame(self.thumbnail_frame, style="Thumbnail.TFrame")
            thumb_frame.grid(row=i, column=0, padx=5, pady=5)

            img_label = ttk.Label(
                thumb_frame,
                image=self.placeholder_thumb,
                cursor="hand2",
                background=self.COLOR_BG_DARK
            )
            img_label.pack()
            img_label.image = self.placeholder_thumb
            img_label.bind("<Button-1>", lambda e, index=i: self.set_photo(index))

            if path in self.thumb_cache:
                try:
                    pil_img = self.thumb_cache[path]
                    photo_img = ImageTk.PhotoImage(pil_img)
                    img_label.configure(image=photo_img)
                    img_label.image = photo_img
                except Exception as e:
                    print(f"Thumb cache display error for {path}: {e}")
            else:
                # CORRECT WORKER DEFINITION: It only needs the label (lbl) and token.
                def worker(p=path, lbl=img_label, token=local_token):
                    try:
                        with Image.open(p) as img:
                            img = ImageOps.exif_transpose(img)
                            img.thumbnail(thumb_size, Image.Resampling.LANCZOS)
                            pil_img = img.copy()
                        self.thumb_cache[p] = pil_img
                        # Put the direct widget reference on the queue
                        self.load_queue.put(('thumb', token, (pil_img, lbl)))
                    except Exception as e:
                        print(f"Could not create thumbnail for {p}: {e}")
                threading.Thread(target=worker, daemon=True).start()
    # ...existing code...

    def display_checklist(self, checklist_data):
        for widget in self.checklist_frame.winfo_children():
            widget.destroy()
        self.checklist_vars.clear()
        for i, (key, label) in enumerate(self.checklist_items_defs):
            var = checklist_data.get(key, tk.BooleanVar())
            self.checklist_vars[key] = var
            cb = ttk.Checkbutton(self.checklist_frame, text=label, variable=var, style="TCheckbutton")
            cb.grid(row=i, column=0, sticky='w', padx=5, pady=2)

    def display_notes(self, notes_text):
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", notes_text)
        self.notes_text.edit_modified(False)

    def on_notes_modified(self, event=None):
        if not self.current_pole_id:
            return
        if self.notes_text.edit_modified():
            self.poles_data[self.current_pole_id]['notes'] = self.notes_text.get("1.0", "end-1c")
            self.notes_text.edit_modified(False)

    # --- Photo navigation ---
    def set_photo(self, index):
        
        self.current_photo_index = index
        pole_data = self.poles_data.get(self.current_pole_id)
        if not pole_data or not pole_data["photos"]:
            return
        path = pole_data["photos"][self.current_photo_index].get('marked_up') or pole_data["photos"][self.current_photo_index]['original']
        self.display_large_photo(path)
        self.update_thumbnail_selection()

    def navigate_photo(self, direction):
        if not self.current_pole_id:
            return
        photo_count = len(self.poles_data[self.current_pole_id]["photos"])
        if photo_count == 0:
            return
        self.set_photo((self.current_photo_index + direction) % photo_count)

    def next_photo(self): self.navigate_photo(1)
    def prev_photo(self): self.navigate_photo(-1)

    # --- Large image (async with skeleton) ---
    def display_large_photo(self, path):
        self.clear_photo_canvas()
        canvas_w = self.photo_canvas.winfo_width()
        canvas_h = self.photo_canvas.winfo_height()
        if canvas_w <= 1 or canvas_h <= 1:
            return

        # Skeleton
        self.photo_canvas.create_rectangle(
            10, 10, canvas_w-10, canvas_h-10,
            fill=self.COLOR_BG, outline=self.COLOR_BORDER,
            width=2, tags="skeleton"
        )
        self.photo_canvas.create_text(
            canvas_w//2, canvas_h//2,
            text="Loading photo…",
            fill=self.COLOR_FG, tags="skeleton"
        )

        self.displayed_image_info['offset_x'] = 0
        self.displayed_image_info['offset_y'] = 0
        self.displayed_image_info['scale'] = 1

        local_token = self.nav_token
        cache_key = (path, canvas_w, canvas_h)

        if cache_key in self.large_cache:
            pil_img, new_w, new_h = self.large_cache[cache_key]
            try:
                self.displayed_image_info['offset_x'] = (canvas_w - new_w) // 2
                self.displayed_image_info['offset_y'] = (canvas_h - new_h) // 2
                self.displayed_image_info['scale'] = new_w / pil_img.width if pil_img.width else 1
                self.photo_canvas.delete("skeleton")
                self.load_queue.put(('large', local_token, (path, pil_img)))
            except Exception as e:
                import traceback
                traceback.print_exc()
                print(f"Large cache display error for {path}: {e}")
            return

        def worker(p=path, cw=canvas_w, ch=canvas_h, token=local_token, key=cache_key):
            try:
                with Image.open(p) as img:
                    img = ImageOps.exif_transpose(img)
                    iw, ih = img.size
                    if iw <= 0 or ih <= 0:
                        return
                    ratio = min(cw / iw, ch / ih)
                    new_w, new_h = int(iw * ratio), int(ih * ratio)
                    resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS).copy()
                    self.large_cache[key] = (resized, new_w, new_h)

                def post():
                    if token != self.nav_token:
                        return
                    self.photo_canvas.delete("skeleton")
                    self.displayed_image_info['offset_x'] = (cw - new_w) // 2
                    self.displayed_image_info['offset_y'] = (ch - new_h) // 2
                    self.displayed_image_info['scale'] = ratio
                    self.load_queue.put(('large', token, (p, resized)))
                self.root.after(0, post)
            except Exception as e:
                import traceback
                traceback.print_exc()
                print(f"Could not display large image {p}: {e}")
                self.root.after(0, self.clear_photo_canvas)

        threading.Thread(target=worker, daemon=True).start()

    # --- Markup handlers ---
    def on_canvas_press(self, event):
        if not self.markup_mode.get():
            return
        self.start_x = event.x
        self.start_y = event.y
        self.current_oval = self.photo_canvas.create_oval(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=3, tags="markup")

    def on_canvas_drag(self, event):
        if not self.markup_mode.get() or self.current_oval is None:
            return
        self.photo_canvas.coords(self.current_oval, self.start_x, self.start_y, event.x, event.y)

    def on_canvas_release(self, event):
        if not self.markup_mode.get() or self.current_oval is None:
            return
        offset_x = self.displayed_image_info.get('offset_x', 0)
        offset_y = self.displayed_image_info.get('offset_y', 0)
        x1, y1 = self.start_x - offset_x, self.start_y - offset_y
        x2, y2 = event.x - offset_x, event.y - offset_y
        photo_entry = self.poles_data[self.current_pole_id]['photos'][self.current_photo_index]
        photo_entry['markups'].append((x1, y1, x2, y2))
        self.current_oval = None

    def draw_temporary_markups(self):
        self.photo_canvas.delete("markup")
        if not self.current_pole_id:
            return
        photo_entry = self.poles_data[self.current_pole_id]['photos'][self.current_photo_index]
        markups = photo_entry.get('markups', [])
        offset_x = self.displayed_image_info.get('offset_x', 0)
        offset_y = self.displayed_image_info.get('offset_y', 0)
        for x1, y1, x2, y2 in markups:
            canvas_x1, canvas_y1 = x1 + offset_x, y1 + offset_y
            canvas_x2, canvas_y2 = x2 + offset_x, y2 + offset_y
            self.photo_canvas.create_oval(canvas_x1, canvas_y1, canvas_x2, canvas_y2, outline="red", width=3, tags="markup")

    def clear_temporary_markups(self):
        self.photo_canvas.delete("markup")
        if not self.current_pole_id:
            return
        photo_entry = self.poles_data[self.current_pole_id]['photos'][self.current_photo_index]
        photo_entry['markups'] = []

    def save_markups(self):
        if not self.current_pole_id or not self.original_parent_folder:
            messagebox.showwarning("Cannot Save", "No active pole or original folder is set.")
            return

        photo_entry = self.poles_data[self.current_pole_id]['photos'][self.current_photo_index]
        temp_photo_path = photo_entry['original'] # This is the path in the temp directory

        try:
            # 1. Determine the save path in the ORIGINAL source directory
            relative_path = os.path.relpath(temp_photo_path, self.temp_dir)
            source_image_path = os.path.join(self.original_parent_folder, relative_path)
            marked_up_save_path = os.path.join(
                os.path.dirname(source_image_path),
                f"marked_{os.path.basename(source_image_path)}"
            )

            # 2. Open the image from the temp dir, draw markups, and save to the source dir
            with Image.open(temp_photo_path) as img:
                img = ImageOps.exif_transpose(img)
                draw = ImageDraw.Draw(img)
                markups = photo_entry.get('markups', [])
                if not markups:
                    messagebox.showinfo("No Markups", "There are no markups to save.")
                    return

                # --- CORRECTION START ---
                # Get the scaling factor used to fit the image on the canvas
                scale = self.displayed_image_info.get('scale', 1.0)
                if scale == 0: scale = 1.0 # Avoid division by zero

                for x1, y1, x2, y2 in markups:
                    # Scale the coordinates from the displayed size back to the original image size
                    scaled_coords = [c / scale for c in [x1, y1, x2, y2]]
                    
                    # Also scale the line width to maintain its visual weight on the full-res image
                    scaled_width = max(1, int(5 / scale))

                    draw.ellipse(
                        scaled_coords,
                        outline="red",
                        width=scaled_width
                    )
                # --- CORRECTION END ---

                img.save(marked_up_save_path)

            # 3. Copy the newly saved file BACK to the temp directory for the current session
            temp_marked_up_path = os.path.join(
                os.path.dirname(temp_photo_path),
                os.path.basename(marked_up_save_path)
            )
            shutil.copy(marked_up_save_path, temp_marked_up_path)

            # 4. Update the application's internal state
            photo_entry['marked_up'] = temp_marked_up_path
            
            # Invalidate caches to force a reload of the new marked-up image
            self.thumb_cache.pop(temp_marked_up_path, None)
            self.large_cache.clear()

            # 5. Refresh the view to show the new marked-up photo
            self.display_thumbnails(self.poles_data[self.current_pole_id]["photos"])
            self.set_photo(self.current_photo_index)
            
            messagebox.showinfo("Saved", f"Markup saved permanently to the original source folder.")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Save Error", f"Failed to save markups:\n{e}")

    # --- Right-side helpers ---
    def clear_photo_canvas(self):
        self.photo_canvas.delete("all")
        self.displayed_image_info.clear()

    def update_thumbnail_selection(self):
        for idx, frame in enumerate(self.thumbnail_frame.winfo_children()):
            if idx == self.current_photo_index:
                frame.configure(style="SelectedThumbnail.TFrame")
            else:
                frame.configure(style="Thumbnail.TFrame")

    def clear_right_panel(self):
        self.pole_name_label.config(text="")
        self.clear_photo_canvas()
        for widget in self.thumbnail_frame.winfo_children():
            widget.destroy()
        for widget in self.checklist_frame.winfo_children():
            widget.destroy()
        self.notes_text.delete("1.0", tk.END)
        for i in self.lookup_tree.get_children():
            self.lookup_tree.delete(i)

    # --- Save/Load/Export ---
    def save_review(self):
        if not self.original_parent_folder:
            messagebox.showerror("No Review Loaded", "No active review session to save.")
            return
        save_path = filedialog.asksaveasfilename(
            title="Save Review State",
            defaultextension=".uprreview",
            filetypes=(("Review Files", "*.uprreview"),)
        )
        if not save_path:
            return

        data_to_save = {
            "original_parent_folder": self.original_parent_folder,
            "lookup_sources": self.lookup_sources,  # <-- new
            "poles": {},
            "pole_order": [item[0] for item in self.pole_list.items]
        }
        for pole_id, data in self.poles_data.items():
            pole_state = {
                "reviewed": data['reviewed'].get(),
                "notes": data['notes'],
                "checklist": {k: v.get() for k, v in data['checklist'].items()},
                "photos": []
            }
            for photo_entry in data['photos']:
                photo_state = {
                    "marked_up": photo_entry['marked_up'],
                    "markups": photo_entry.get('markups', [])
                }
                pole_state['photos'].append(photo_state)
            data_to_save['poles'][pole_id] = pole_state

        try:
            with open(save_path, 'w') as f:
                json.dump(data_to_save, f)
            messagebox.showinfo("Saved", f"Review saved to {os.path.basename(save_path)}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Save Error", f"Failed to save review:\n{e}")


    def load_review(self):
        review_file = filedialog.askopenfilename(
            title="Load Review File",
            filetypes=(("Review Files", "*.uprreview"),)
        )
        if not review_file:
            return
        try:
            with open(review_file, 'r') as f:
                saved_data = json.load(f)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load review file:\n{e}")
            return

        # Determine parent folder (prefer saved path)
        parent_folder = saved_data.get("original_parent_folder")
        if not parent_folder or not os.path.isdir(parent_folder):
            parent_folder = filedialog.askdirectory(title="Select the Parent Folder of Pole Photos")
        if not parent_folder:
            return
        self.original_parent_folder = parent_folder

        # Resolve lookup sources (prefer saved; else prompt main; PN/DR next to script)
        sources = saved_data.get("lookup_sources") or {}
        main_path = sources.get('main')
        pn_path = sources.get('pn') or os.path.join(os.path.dirname(os.path.realpath(__file__)), 'PN.xlsx')
        dr_path = sources.get('dr') or os.path.join(os.path.dirname(os.path.realpath(__file__)), 'DR.xlsx')
        if not (main_path and os.path.isfile(main_path)):
            main_path = filedialog.askopenfilename(
                title="Select the Main Barcode Sheet (Excel/CSV)",
                filetypes=(("Supported Files", "*.xlsx *.xls *.csv"), ("All files", "*.*"))
            )

        # Progress UI
        self._open_progress("Preparing review session")
        if main_path and os.path.isfile(main_path) and os.path.isfile(pn_path) and os.path.isfile(dr_path):
            self._set_progress_status("Loading lookup data (PN & DR)…", note=os.path.basename(main_path))
            try:
                self.lookup_data = process_data(main_path, pn_path, dr_path)
                self.lookup_sources = {'main': main_path, 'pn': pn_path, 'dr': dr_path}
            except Exception as e:
                self.lookup_data = None
                messagebox.showwarning("Lookup Data", f"Could not rebuild lookup data:\n{e}\n\nContinuing without lookup.")
        else:
            self.lookup_data = None
            messagebox.showwarning("Lookup Data", "Lookup sources not available; continuing without lookup data.")

        # Continue into photo copy using the SAME window
        self._set_progress_status("Copying images locally…")
        self.start_photo_review(parent_folder, saved_data)



    def export_to_html(self):
        if not self.poles_data:
            messagebox.showwarning("No Data", "No poles loaded.")
            return

        # Toggle prompt
        only_markups = messagebox.askyesno(
            "Export Options",
            "Export only poles with markups?\n\n"
            "Markups include any checked checklist item, any on-photo marks, "
            "any Additional Notes text, or any Pole Lookup matches."
        )

        export_path = filedialog.asksaveasfilename(
            title="Export to HTML",
            defaultextension=".html",
            filetypes=(("HTML Files", "*.html"),)
        )
        if not export_path:
            return

        try:
            import html
            from datetime import datetime

            # Helper: summarize a pole's “signals”
            def summarize_pole(pole_data):
                # Failures = items the user checked (you only check failures)
                failed = []
                for key, label in self.checklist_items_defs:
                    var = pole_data['checklist'].get(key)
                    if var and bool(var.get()):
                        failed.append(label)

                # Any canvas markups or saved marked_up images?
                has_marks = any(
                    (p.get('marked_up') or (p.get('markups') and len(p['markups']) > 0))
                    for p in pole_data.get('photos', [])
                )

                # Notes
                notes = (pole_data.get('notes') or "").strip()
                has_notes = len(notes) > 0

                # Lookup
                lookup = pole_data.get('lookup_info') or []
                has_lookup = len(lookup) > 0

                return failed, has_marks, has_notes, has_lookup, notes, lookup

            # Helper: return (thumb_b64, large_b64) honoring EXIF orientation
            def encode_pair(img_path, prefer_jpeg=True):
                with Image.open(img_path) as im:
                    im = ImageOps.exif_transpose(im)
                    # Large first (for modal)
                    large = im.copy()
                    large.thumbnail((1400, 1400), Image.Resampling.LANCZOS)
                    lb = BytesIO()
                    fmt = "JPEG" if prefer_jpeg else "PNG"
                    save_kwargs = {"quality": 85} if fmt == "JPEG" else {}
                    large.save(lb, format=fmt, **save_kwargs)
                    large_b64 = base64.b64encode(lb.getvalue()).decode("utf-8")

                    # Thumb (grid)
                    thumb = im.copy()
                    thumb.thumbnail((600, 600), Image.Resampling.LANCZOS)
                    tb = BytesIO()
                    thumb.save(tb, format=fmt, **save_kwargs)
                    thumb_b64 = base64.b64encode(tb.getvalue()).decode("utf-8")
                    mime = "image/jpeg" if fmt == "JPEG" else "image/png"
                    return f"data:{mime};base64,{thumb_b64}", f"data:{mime};base64,{large_b64}"

            css = """
            <style>
            :root{
                --bg:#111418; --panel:#1a1f24; --card:#1f262c; --muted:#9fb1c1; --fg:#e6eef5; --accent:#49a6ff; --bad:#ff5a7a; --chip:#2c353d; --line:#2a3942;
            }
            *{box-sizing:border-box}
            body{margin:0; background:var(--bg); color:var(--fg); font:15px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,"Apple Color Emoji","Segoe UI Emoji";}
            header{padding:32px 24px; border-bottom:1px solid #223038; background:linear-gradient(180deg,#141a1f,transparent);}
            h1{margin:0 0 6px; font-size:26px;}
            .meta{color:var(--muted)}
            .wrap{padding:24px; max-width:1400px; margin:0 auto;}
            .grid{display:grid; gap:18px; grid-template-columns:repeat(auto-fill,minmax(320px,1fr));}
            .card{background:var(--card); border:1px solid #26323a; border-radius:14px; overflow:hidden; display:flex; flex-direction:column;}
            .card header{padding:14px 16px; border-bottom:1px solid #26323a; background:linear-gradient(180deg,#20272e,#1f262c);}
            .card h2{margin:0; font-size:18px;}
            .section{padding:12px 16px;}
            .chips{display:flex; flex-wrap:wrap; gap:8px; margin:8px 0 0}
            .chip{background:var(--chip); color:var(--muted); padding:4px 8px; border-radius:999px; font-size:12px; border:1px solid #2b3942;}
            .chip.bad{color:var(--bad); border-color:#3a2b33}
            .notes{background:var(--panel); border:1px dashed var(--line); border-radius:10px; padding:10px; color:var(--fg); white-space:pre-wrap}
            .photos{display:grid; grid-template-columns:repeat(auto-fill,minmax(140px,1fr)); gap:8px; margin-top:8px}
            .photos img{width:100%; height:120px; object-fit:cover; border-radius:8px; border:1px solid #2a3840; display:block; cursor:pointer}
            .empty{color:var(--muted); font-size:13px}
            table.lookup{width:100%; border-collapse:collapse; margin-top:8px; font-size:13px}
            .lookup th, .lookup td{border:1px solid var(--line); padding:6px 8px; text-align:left; vertical-align:top}
            .lookup th{background:#20272e; color:var(--fg); font-weight:600}
            .lookup td{color:var(--fg)}
            
            .tablewrap{overflow:auto}              /* allows horizontal scroll if needed */
            table.lookup{width:max-content}        /* let table be as wide as needed */
            .lookup th,.lookup td{
            word-break:break-word;               /* wrap long tokens */
            overflow-wrap:anywhere;              /* wrap long IDs/strings */
            white-space:normal;
            }

            /* Modal */
            .lightbox{position:fixed; inset:0; background:rgba(0,0,0,.8); display:none; align-items:center; justify-content:center; z-index:9999; padding:24px}
            .lightbox.open{display:flex}
            .lightbox img{max-width:95vw; max-height:95vh; display:block; border-radius:10px; border:1px solid #2a3840}
            .lightbox .hint{position:fixed; bottom:16px; color:#c9d6e2; font-size:12px}
            </style>
            """

            js = """
            <script>
            (function(){
                const lb = document.getElementById('lb');
                const lbImg = document.getElementById('lb-img');
                function open(src){ lbImg.src = src; lb.classList.add('open'); }
                function close(){ lb.classList.remove('open'); lbImg.src=''; }
                document.addEventListener('click', (e)=>{
                const t = e.target;
                if(t.matches('img[data-large]')){ open(t.getAttribute('data-large')); }
                else if(t.closest('#lb') && !t.closest('#lb-img')){ close(); }
                });
                document.addEventListener('keydown', (e)=>{ if(e.key==='Escape') close(); });
            })();
            </script>
            """

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

            # Determine order from the left panel if available
            ordered_ids = [item[0] for item in self.pole_list.items] or list(self.poles_data.keys())

            # Build filtered list based on toggle
            poles_to_export = []
            for pole_id in ordered_ids:
                data = self.poles_data.get(pole_id)
                if not data:
                    continue
                failed, has_marks, has_notes, has_lookup, notes, lookup_rows = summarize_pole(data)
                include = True
                if only_markups:
                    include = bool(failed or has_marks or has_notes or has_lookup)
                if include:
                    poles_to_export.append((pole_id, data, failed, has_marks, has_notes, has_lookup, notes, lookup_rows))

            if not poles_to_export:
                messagebox.showinfo("Export", "No poles matched the selected export option.")
                return

            with open(export_path, 'w', encoding='utf-8') as f:
                f.write("<!doctype html><html><head><meta charset='utf-8'>")
                f.write("<meta name='viewport' content='width=device-width, initial-scale=1'>")
                f.write("<title>Pole Review Report</title>")
                f.write(css)
                f.write("</head><body>")
                f.write("<header><h1>Pole Review Report</h1>")
                f.write(f"<div class='meta'>Generated {html.escape(now_str)}"
                        f"{' — Only poles with markups' if only_markups else ' — All poles'}</div></header>")
                # Modal root (once per page)
                f.write("<div id='lb' class='lightbox'><img id='lb-img' alt=''><div class='hint'>Click anywhere or press Esc to close</div></div>")
                f.write("<div class='wrap'><div class='grid'>")

                for pole_id, data, failed, has_marks, has_notes, has_lookup, notes, lookup_rows in poles_to_export:
                    f.write("<article class='card'>")
                    f.write("<header>")
                    f.write(f"<h2>{html.escape(str(pole_id))}</h2>")
                    f.write("</header>")

                    # Failures (only checked items)
                    f.write("<div class='section'><div><strong>Failures</strong></div>")
                    if failed:
                        f.write("<div class='chips'>")
                        for label in failed:
                            f.write(f"<span class='chip bad'>✗ {html.escape(label)}</span>")
                        f.write("</div>")
                    else:
                        f.write("<div class='empty'>No checklist failures recorded.</div>")
                    f.write("</div>")

                    # Notes
                    if has_notes:
                        f.write("<div class='section'><div><strong>Additional Notes</strong></div>")
                        f.write(f"<div class='notes'>{html.escape(notes)}</div></div>")

                    # Lookup Matches
                    if has_lookup:
                        f.write("<div class='section'><div><strong>Lookup Matches</strong></div>")
                        f.write("<div class='tablewrap'><table class='lookup'><thead><tr>")
                        f.write("<th>Type</th><th>ID</th><th>Info</th><th>Location</th><th>Requirement</th>")
                        f.write("</tr></thead><tbody>")
                        for row in lookup_rows:
                            t = html.escape(str(row.get('Type', '')))
                            i = html.escape(str(row.get('ID', '')))
                            info = html.escape(str(row.get('Info', '')))
                            loc = html.escape(str(row.get('Location', '')))
                            required = html.escape(str(row.get('Requirement', '')))
                            f.write(f"<tr><td>{t}</td><td>{i}</td><td>{info}</td><td>{loc}</td><td>{required}</td></tr>")
                        f.write("</tbody></table></div></div>")  # close table + wrapper

                    # Photos (prefer marked-up versions), with EXIF orientation and modal
                    f.write("<div class='section'><div><strong>Photos</strong></div><div class='photos'>")
                    any_photo = False
                    for photo_entry in data.get('photos', []):
                        path = photo_entry.get('marked_up') or photo_entry.get('original')
                        if not path:
                            continue
                        try:
                            thumb_src, large_src = encode_pair(path, prefer_jpeg=True)
                            alt = html.escape(os.path.basename(path))
                            # Use data-large attribute for the modal
                            f.write(f"<img src='{thumb_src}' data-large='{large_src}' alt='{alt}'>")
                            any_photo = True
                        except Exception as e:
                            import traceback
                            traceback.print_exc()
                            # Fallback to original if marked_up missing/unreadable
                            if photo_entry.get('marked_up') and photo_entry.get('original'):
                                try:
                                    thumb_src, large_src = encode_pair(photo_entry['original'], prefer_jpeg=True)
                                    alt = html.escape(os.path.basename(photo_entry['original']))
                                    f.write(f"<img src='{thumb_src}' data-large='{large_src}' alt='{alt}'>")
                                    any_photo = True
                                except Exception as e2:
                                    f.write(f"<div class='small'>Could not load image: {html.escape(str(e2))}</div>")
                            else:
                                f.write(f"<div class='small'>Could not load image: {html.escape(str(e))}</div>")
                    if not any_photo:
                        f.write("<div class='empty'>No photos available.</div>")
                    f.write("</div></div>")  # photos

                    f.write("</article>")

                f.write("</div></div>")  # grid+wrap
                f.write("<footer><div class='small'>Unified Pole Photo Reviewer — HTML export</div></footer>")
                f.write(js)
                f.write("</body></html>")

            messagebox.showinfo("Export Complete", f"Report saved to {export_path}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Export Error", f"Failed to export HTML:\n{e}")




    # --- External open & close ---
    def open_current_photo_externally(self, event=None):
        if not self.current_pole_id:
            return
        photo_entry = self.poles_data[self.current_pole_id]['photos'][self.current_photo_index]
        path = photo_entry.get('marked_up') or photo_entry['original']
        try:
            if sys.platform.startswith('darwin'):
                subprocess.call(('open', path))
            elif os.name == 'nt':
                os.startfile(path)
            elif os.name == 'posix':
                subprocess.call(('xdg-open', path))
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Open Error", f"Could not open image externally:\n{e}")

    def on_closing(self):
        try:
            shutil.rmtree(self.temp_dir)
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"Failed to delete temp dir: {e}")
        self.root.destroy()

class DraggableCheckboxListbox(tk.Frame):
    """
    Checkbox list with click-to-select, toggle, and drag-to-reorder.
    Selection stays in sync after reorders; dragging works from anywhere on the row.
    Mouse wheel scrolling is enabled on Windows/macOS/Linux.
    """
    def __init__(self, master, bg_color, accent_color, command=None, **kwargs):
        super().__init__(master, **kwargs)
        self.bg_color = bg_color
        self.accent_color = accent_color
        self.command = command

        self.items = []     # [(text, tk.BooleanVar), ...]
        self._rows = []     # [row_frame, ...]
        self.selected_index = None

        # drag state
        self._dragging = False
        self._drag_index = None
        self._y_start = 0

        self.canvas = tk.Canvas(self, background=self.bg_color, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, background=self.bg_color)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # --- Mouse wheel scrolling (Windows/macOS) ---
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)            # Windows + most macOS Tk builds
        self.scrollable_frame.bind("<MouseWheel>", self._on_mousewheel)  # wheel over child rows

        # --- Mouse wheel scrolling (Linux/X11) ---
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll( 1, "units"))
        self.scrollable_frame.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.scrollable_frame.bind("<Button-5>", lambda e: self.canvas.yview_scroll( 1, "units"))

    # ---------- helpers ----------
    def _row_index_for_widget(self, widget):
        """Resolve current row index by walking up to the row frame and finding it in self._rows."""
        w = widget
        while w is not None and w not in self._rows:
            w = w.master
        if w in self._rows:
            return self._rows.index(w)
        return None

    def _apply_row_bindings(self, row, *children):
        targets = (row, *children)
        for t in targets:
            t.bind("<Button-1>", self._on_click_select)
            t.bind("<ButtonPress-1>", self._on_drag_start)
            t.bind("<B1-Motion>", self._on_drag_motion)
            t.bind("<ButtonRelease-1>", self._on_drag_end)

    def _sel_bg(self, i):
        return self.bg_color if i != self.selected_index else "#29323a"

    # ---------- scrolling ----------
    def _on_mousewheel(self, event):
        # On Windows, event.delta is +/-120 per notch. On macOS it can be smaller; normalize to 1 step.
        delta = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(delta, "units")

    # ---------- public API ----------
    def clear(self):
        for w in self.scrollable_frame.winfo_children():
            w.destroy()
        self.items.clear()
        self._rows.clear()
        self.selected_index = None

    def populate(self, items):
        self.clear()
        for text, var in items:
            row = tk.Frame(self.scrollable_frame, background=self.bg_color)
            row.pack(fill="x", pady=1)

            cb = tk.Checkbutton(
                row, text=text, variable=var, background=self.bg_color,
                activebackground=self.bg_color, fg="white",
                selectcolor=self.bg_color, anchor="w"
            )
            cb.pack(side="left", fill="x", expand=True, padx=(6, 6))

            # bind to both row and checkbox (so dragging works anywhere on the row)
            self._apply_row_bindings(row, cb)

            self.items.append((text, var))
            self._rows.append(row)

        if self.items:
            self.on_select(0)

    def get_selected_index(self):
        return self.selected_index

    def select_item(self, index):
        self.on_select(index)

    def get_item_count(self):
        return len(self.items)

    # ---------- selection & checking ----------
    def on_select(self, index):
        if index is None or index < 0 or index >= len(self._rows):
            return
        self.selected_index = index
        for i, row in enumerate(self._rows):
            row.configure(background=self._sel_bg(i))
        if self.command:
            self.command(self.items[index][0])

    def _on_click_select(self, event):
        idx = self._row_index_for_widget(event.widget)
        if idx is not None:
            self.on_select(idx)

    def _toggle_check_at(self, index):
        if index is None: return
        text, var = self.items[index]
        var.set(not var.get())
        if self.command:
            self.command(text)

    # ---------- drag & reorder ----------
    def _on_drag_start(self, event):
        idx = self._row_index_for_widget(event.widget)
        if idx is None: return
        self._dragging = True
        self._drag_index = idx
        self._y_start = event.y_root

    def _on_drag_motion(self, event):
        if not self._dragging:
            return
        y_local = event.widget.winfo_pointery() - self.scrollable_frame.winfo_rooty()
        target_index = self._index_from_y(y_local)
        if target_index is None or target_index == self._drag_index:
            return
        self._swap_rows(self._drag_index, target_index)
        self._drag_index = target_index

    def _on_drag_end(self, event):
        if not self._dragging:
            return
        self._dragging = False
        if self._drag_index is not None:
            # keep selection on the dragged row’s new position
            self.on_select(self._drag_index)
        self._drag_index = None

    def _index_from_y(self, y):
        centers = []
        for i, row in enumerate(self._rows):
            row.update_idletasks()
            y1 = row.winfo_y()
            h = row.winfo_height() or 1
            centers.append((abs((y1 + h/2) - y), i))
        if not centers:
            return None
        centers.sort()
        return centers[0][1]

    def _swap_rows(self, i, j):
        if i == j: return
        # swap backing data
        self.items[i], self.items[j] = self.items[j], self.items[i]
        self._rows[i], self._rows[j] = self._rows[j], self._rows[i]
        # re-pack visually in new order
        for row in self._rows:
            row.pack_forget()
        for row in self._rows:
            row.pack(fill="x", pady=1)


    # ----- External helpers -----
    def get_selected_index(self):
        return self.selected_index

    def select_item(self, index):
        self.on_select(index)

    def get_item_count(self):
        return len(self.items)

if __name__ == "__main__":
    _hide_console_on_windows() #comment this out if you want to see console output
    root = tk.Tk()
    app = PoleReviewerApp(root)
    root.mainloop()

