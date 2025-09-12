#!/usr/bin/env python3
"""
Unified Launcher (Bootstrap Edition) — Corrected
- Pretty UI using ttkbootstrap
- Mouse wheel scrolling
- Strict JSON loader (no regex munging); reads tools.json as-is
- Types: html | url | python | exe | file
- Open containing folder
- Ctrl+F to focus search, Enter to launch first visible
- Double-click safe (uses script dir for tools.json, logs crash + shows dialog)
"""

import json, os, sys, subprocess, webbrowser, pathlib, platform, shlex, traceback

# --- Portable path helpers ---
def _app_base_dir():
    # If frozen by PyInstaller
    if getattr(sys, 'frozen', False) and hasattr(sys, 'executable'):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))

import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame

APP_TITLE = "Unified Launcher"
CONFIG_BASENAME = "tools.json"
DEFAULT_THEME = "flatly"   # try "darkly" for dark mode

# ------------------------ Utilities ------------------------
def is_windows():
    return platform.system().lower().startswith("win")

def expand_path(p, config_dir=None):
    if not p:
        return p
    # If it's a URL or other URI scheme, leave it alone
    pl = str(p).lower()
    if pl.startswith(('http://', 'https://', 'file://', 'mailto:', 'ftp://')):
        return p
    base_dir = _app_base_dir()
    cfg_dir = config_dir or base_dir
    # Expand env vars and ~
    p = os.path.expandvars(os.path.expanduser(p))
    # Placeholders
    p = p.replace('{BASE}', base_dir).replace('{CONFIG}', cfg_dir)
    # Relative paths resolve from config dir (hands-off sharing)
    if not os.path.isabs(p):
        p = os.path.normpath(os.path.join(cfg_dir, p))
    return p

def open_path(p):
    p = expand_path(p)
    if not p:
        return
    if os.path.isdir(p):
        if is_windows():
            os.startfile(p)  # type: ignore
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    else:
        lower = str(p).lower()
        if lower.startswith("http://") or lower.startswith("https://"):
            webbrowser.open(p)
        elif lower.endswith(".html") or lower.endswith(".htm"):
            if os.path.exists(p):
                webbrowser.open(pathlib.Path(p).as_uri())
            else:
                webbrowser.open(p)
        else:
            try:
                if is_windows():
                    os.startfile(p)  # type: ignore
                elif platform.system() == "Darwin":
                    subprocess.Popen(["open", p])
                else:
                    subprocess.Popen(["xdg-open", p])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open:\n{p}\n\n{e}")

def tolerant_json_load(path):
    # Strict JSON load (no regex transforms)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def tolerant_json_save(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def launch_tool(tool, settings, config_dir):
    ttype = (tool.get("type") or "").lower()
    raw_path = tool.get("path", "")
    # Use raw for URL type to avoid accidental path joining
    path = raw_path if ttype == "url" else expand_path(raw_path, config_dir)
    if ttype in ("html", "url"):
        if ttype == "url" or (path.lower().startswith("http://") or path.lower().startswith("https://")):
            webbrowser.open(path)
        else:
            if not os.path.exists(path):
                messagebox.showerror("Missing File", f"HTML file not found:\n{path}")
                return
            webbrowser.open(pathlib.Path(path).as_uri())
        return

    if ttype == "file":
        if not path:
            messagebox.showerror("Error", "No path provided for file.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Missing File", f"File not found:\n{path}")
            return
        open_path(path)
        return

    if ttype == "exe":
        if not path:
            messagebox.showerror("Error", "No path provided for executable/file.")
            return
        # If this looks like a normal data file, open with system default
        ext = os.path.splitext(path)[1].lower()
        non_exec_exts = {".xlsx",".xlsm",".xls",".csv",".txt",".pdf",".docx",".pptx",".png",".jpg",".jpeg",".gif",".html",".htm"}
        if ext in non_exec_exts:
            open_path(path)
            return
        args = tool.get("arguments", [])
        try:
            subprocess.Popen([path] + args, cwd=expand_path(tool.get("working_dir",""), config_dir) or None)
        except Exception as e:
            messagebox.showerror("Launch Failed", f"Failed to launch:\n{path}\n\n{e}")
        return

    if ttype == "python":
        if not path:
            messagebox.showerror("Error", "No path provided for Python script.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Missing File", f"Python script not found:\n{path}")
            return

        interpreter = expand_path(tool.get("interpreter") or settings.get("default_python","python"), config_dir)
        args = tool.get("arguments", [])
        env = os.environ.copy()
        env.update({k:str(v) for k,v in tool.get("env", {}).items()})
        workdir = expand_path(tool.get("working_dir",""), config_dir) or os.path.dirname(path) or None

        cmd = [interpreter, path] + args
        try:
            subprocess.Popen(cmd, cwd=workdir, env=env, shell=False)
        except FileNotFoundError:
            messagebox.showerror("Interpreter Not Found", f"Could not find Python interpreter:\n{interpreter}\n\nSet 'settings.default_python' or tool.interpreter to a full path.")
        except Exception as e:
            messagebox.showerror("Launch Failed", f"Failed to launch:\n{' '.join(shlex.quote(c) for c in cmd)}\n\n{e}")
        return

    messagebox.showerror("Unsupported Type", f"Unsupported tool type: {ttype}")

# ------------------------ App ------------------------
class LauncherApp(tb.Window):
    def __init__(self, config_path):
        super().__init__(themename=DEFAULT_THEME)
        self.title(APP_TITLE)

        self.config_path = config_path
        self.config_dir = os.path.dirname(os.path.abspath(self.config_path))
        try:
            self.data = tolerant_json_load(config_path)
        except Exception as e:
            messagebox.showerror("Config Error", str(e))
            raise

        # Window size from settings
        s = self.data.get("settings", {})
        w = int(s.get("window_width", 1024))
        h = int(s.get("window_height", 700))
        self.geometry(f"{w}x{h}")

        # Apply theme from settings if present (case-insensitive)
        theme = (s.get("theme") or DEFAULT_THEME).lower()
        try:
            self.style.theme_use(theme)
        except Exception:
            pass  # ignore invalid theme names

        self._build_ui()
        self._populate()

        # Show the loaded config path (useful when sharing internally)
        try:
            self.status.set(f"Loaded config: {os.path.abspath(self.config_path)}")
        except Exception:
            pass

    # ---------- UI ----------
    def _build_ui(self):
        # Top bar
        top = tb.Frame(self, padding=(12,8))
        top.pack(side=tk.TOP, fill=tk.X)

        left = tb.Frame(top)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        right = tb.Frame(top)
        right.pack(side=tk.RIGHT)

        tb.Label(left, text="🧰 Unified Launcher", bootstyle="inverse", font=("-size", 14, "-weight", "bold")).pack(side=tk.LEFT, padx=(0,10))

        # Search
        search_wrap = tb.Frame(left)
        search_wrap.pack(side=tk.RIGHT, padx=(8,0))
        tb.Label(search_wrap, text="🔎").pack(side=tk.LEFT, padx=(0,6))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._populate())
        self.search_entry = tb.Entry(search_wrap, textvariable=self.search_var, width=36)
        self.search_entry.pack(side=tk.LEFT)
        self.search_entry.bind("<Control-f>", lambda e: (self.search_entry.focus_set(), "break"))
        self.search_entry.bind("<Return>", self._launch_first_visible)

        # Category
        tb.Label(right, text="Category:", bootstyle="secondary").pack(side=tk.LEFT, padx=(0,6))
        self.cat_var = tk.StringVar(value="All")
        self.cat_combo = tb.Combobox(right, textvariable=self.cat_var, width=18, state="readonly")
        self.cat_combo.pack(side=tk.LEFT)
        self.cat_combo.bind("<<ComboboxSelected>>", lambda e: self._populate())

        tb.Button(right, text="Reload", command=self.reload, bootstyle=PRIMARY).pack(side=tk.LEFT, padx=6)
        tb.Button(right, text="Open Config", command=self.open_config, bootstyle=SECONDARY).pack(side=tk.LEFT)
        tb.Button(right, text="Theme", command=self.cycle_theme, bootstyle=INFO).pack(side=tk.LEFT, padx=(6,0))

        # Content area with scrolling
        self.scrollframe = ScrolledFrame(self, autohide=True, padding=(6,6))
        self.scrollframe.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.card_parent = tb.Frame(self.scrollframe)
        self.card_parent.pack(fill=tk.BOTH, expand=True)

        # Status bar
        self.status = tk.StringVar(value="Ready.")
        stat = tb.Label(self, textvariable=self.status, anchor="w", bootstyle="secondary")
        stat.pack(side=tk.BOTTOM, fill=tk.X, padx=6, pady=4)

        # ---- Custom styles (no filled backgrounds) ----
        self.style.configure("Title.TLabel", font=("-size", 12, "-weight", "bold"))
        # Use a muted foreground color derived from theme
        try:
            muted = self.style.colors.secondary
        except Exception:
            muted = "#777777"
        self.style.configure("Muted.TLabel", foreground=muted)
        self.style.configure("Path.TLabel", font=("-size", 9))
        self.style.configure("Card.TFrame", padding=14, borderwidth=1, relief="groove")

        # Bind mousewheel globally (ScrolledFrame already handles most cases, but ensure)
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.bind_all("<Button-4>", self._on_mousewheel, add="+")  # Linux up
        self.bind_all("<Button-5>", self._on_mousewheel, add="+")  # Linux down

    def _on_mousewheel(self, event):
        widget = self.scrollframe
        if event.num == 4:
            widget.yview_scroll(-1, "units")
        elif event.num == 5:
            widget.yview_scroll(1, "units")
        else:
            delta = -1 if event.delta > 0 else 1
            widget.yview_scroll(delta, "units")

    def open_config(self):
        p = os.path.abspath(self.config_path)
        open_path(p)

    def reload(self):
        try:
            self.data = tolerant_json_load(self.config_path)
            self._populate()
            self.status.set("Config reloaded.")
        except Exception as e:
            messagebox.showerror("Reload Failed", str(e))

    def _gather_categories(self):
        cats = {"All"}
        for t in self.data.get("tools", []):
            cats.add(t.get("category","Uncategorized") or "Uncategorized")
        return sorted(cats)

    def _clear_cards(self):
        for w in list(self.card_parent.children.values()):
            w.destroy()

    def _matches(self, tool, q):
        hay = " ".join([
            tool.get("name",""),
            tool.get("type",""),
            tool.get("path",""),
            tool.get("category",""),
            tool.get("description","")
        ]).lower()
        return all(part in hay for part in q.lower().split())

    def _populate(self):
        self._clear_cards()
        cats = self._gather_categories()
        current = self.cat_var.get()
        if current not in cats:
            self.cat_var.set("All")
            current = "All"
        self.cat_combo["values"] = cats

        q = self.search_var.get().strip()
        shown = 0

        for tool in self.data.get("tools", []):
            if q and not self._matches(tool, q):
                continue
            cat = tool.get("category","Uncategorized") or "Uncategorized"
            if current != "All" and cat != current:
                continue

            shown += 1
            card = tb.Frame(self.card_parent, style="Card.TFrame")
            card.pack(side=tk.TOP, fill=tk.X, padx=12, pady=8)
            # subtle hover relief change
            card.bind("<Enter>", lambda e, w=card: w.configure(relief="ridge"))
            card.bind("<Leave>", lambda e, w=card: w.configure(relief="groove"))

            # Title row
            top = tb.Frame(card)
            top.pack(side=tk.TOP, fill=tk.X)
            tb.Label(top, text=tool.get("name","(unnamed)"), style="Title.TLabel").pack(side=tk.LEFT)

            tb.Label(top, text=f"[{cat} — {tool.get('type','?').upper()}]", style="Muted.TLabel").pack(side=tk.RIGHT)

            # Description (plain label, no filled background)
            desc = tool.get("description","").strip()
            if desc:
                tb.Label(card, text=desc, wraplength=820, justify="left").pack(side=tk.TOP, anchor="w", pady=(6,2))

            path = tool.get("path","").strip()
            tb.Label(card, text=path, style="Path.TLabel").pack(side=tk.TOP, anchor="w")

            # subtle separator
            tb.Separator(card).pack(side=tk.TOP, fill=tk.X, pady=(8,6))

            # Buttons
            btns = tb.Frame(card)
            btns.pack(side=tk.TOP, fill=tk.X, pady=(4,0))

            tb.Button(btns, text="🚀 Launch", command=lambda t=tool: self.launch_and_status(t), bootstyle=SUCCESS).pack(side=tk.LEFT)
            tb.Button(btns, text="📁 Open Folder", command=lambda p=path: self.open_containing(p), bootstyle=SECONDARY).pack(side=tk.LEFT, padx=8)

        if shown == 0:
            tb.Label(self.card_parent, text="No tools match your filters.", bootstyle="secondary").pack(side=tk.TOP, pady=20)

    def open_containing(self, p):
        p = expand_path(p or "", self.config_dir)
        if not p:
            return
        folder = p if os.path.isdir(p) else os.path.dirname(p)
        if not folder:
            return
        try:
            open_path(folder)
        except Exception as e:
            messagebox.showerror("Open Folder Failed", str(e))

    def launch_and_status(self, tool):
        try:
            launch_tool(tool, self.data.get("settings", {}), self.config_dir)
            self.status.set(f"Launched: {tool.get('name','(unnamed)')}")
        except Exception as e:
            messagebox.showerror("Launch Failed", str(e))

    def _launch_first_visible(self, *_):
        # Launch the first card's Launch button
        for child in self.card_parent.winfo_children():
            for sub in child.winfo_children():
                if isinstance(sub, tb.Frame):
                    for w in sub.winfo_children():
                        try:
                            if isinstance(w, tb.Button) and "Launch" in w.cget("text"):
                                w.invoke()
                                return
                        except Exception:
                            pass

    def cycle_theme(self):
        # Use the themes available in this install
        themes = list(self.style.theme_names())
        try:
            cur = self.style.theme.name
            idx = themes.index(cur)
        except Exception:
            idx = -1
        new = themes[(idx + 1) % len(themes)]

        self.style.theme_use(new)

        # persist to settings
        s = self.data.get("settings", {})
        s["theme"] = new
        self.data["settings"] = s
        try:
            tolerant_json_save(self.config_path, self.data)
        except Exception:
            pass

# ------------------------ Entrypoint ------------------------
def main():
    # Use tools.json next to the script by default (so double-click works)
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    cfg = os.path.join(script_dir, CONFIG_BASENAME)
    if len(sys.argv) > 1:
        cfg = sys.argv[1]
    app = LauncherApp(cfg)
    app.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        # Write a log next to the script
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        log_path = os.path.join(script_dir, "launcher_error.log")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        try:
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("Launcher crashed", f"An error occurred and was saved to:\n{log_path}\n\n{err}")
        finally:
            sys.exit(1)
