#!/usr/bin/env python3
"""
GFH Legacy Excel Converter  (one-time batch tool)
=================================================
Pick a folder → converts every legacy .xls (and .xlsm/.xlt/.xlsb) file inside it
to modern .xlsx, using real Excel so formatting, formulas and data are preserved.

  • Browse to a folder (optionally include subfolders)
  • Each  Name.xls  →  Name.xlsx  in the same place
  • Originals are KEPT by default (tick a box to delete them after success)
  • Already-converted files are skipped unless you tick "overwrite"

Requires Microsoft Excel installed (uses Excel COM for faithful conversion).
"""

# ── Auto-installer (version-aware) ─────────────────────────────────────────────
import sys, subprocess
def _pkg_version(dist):
    try:
        import importlib.metadata as _md
        return _md.version(dist)
    except Exception:
        return None
def _ensure(pip_name, imp_name):
    if _pkg_version(pip_name) is not None: return
    try: __import__(imp_name.split(".")[0])
    except ImportError:
        try:
            print(f"Installing {pip_name}…")
            subprocess.check_call([sys.executable,"-m","pip","install","--upgrade",pip_name,"-q"])
        except Exception as e:
            print(f"  [WARN] could not install {pip_name}: {e}")
for _p,_i in [("pywin32","win32com")]:
    _ensure(_p,_i)

import os, time, threading, queue, traceback
from datetime import datetime
import tkinter as tk
from theme_manager import ThemeManager, apply_theme_to_window, get_copyright_year
from header_manager import FixedHeaderManager
from logo_handler import LogoHandler
from tkinter import ttk, scrolledtext, messagebox, filedialog
import win32com.client

# Excel file-format codes
XLSX = 51        # xlOpenXMLWorkbook
LEGACY_EXTS = (".xls", ".xlsm", ".xlt", ".xlsb", ".xlc")   # convert these → .xlsx
RESTART_EVERY = 60   # restart Excel periodically to avoid memory bloat on big batches

_CANCEL = threading.Event()


def _find_files(folder, recurse):
    out=[]
    if recurse:
        for root,_,files in os.walk(folder):
            for f in files: out.append(os.path.join(root,f))
    else:
        out=[os.path.join(folder,f) for f in os.listdir(folder)]
    res=[]
    for p in out:
        base=os.path.basename(p)
        if base.startswith("~$"): continue                    # Excel lock files
        if os.path.splitext(base)[1].lower() in LEGACY_EXTS:
            res.append(p)
    return sorted(res)


class Converter:
    def __init__(self, log):
        self.log=log; self.xl=None; self._opened=0

    def _start_excel(self):
        self.xl=win32com.client.DispatchEx("Excel.Application")
        self.xl.Visible=False
        self.xl.DisplayAlerts=False
        try: self.xl.AutomationSecurity=3   # block macros from prompting
        except Exception: pass
        try: self.xl.AskToUpdateLinks=False
        except Exception: pass

    def _stop_excel(self):
        if self.xl is not None:
            try: self.xl.Quit()
            except Exception: pass
        self.xl=None

    def _recycle_if_needed(self):
        self._opened+=1
        if self._opened % RESTART_EVERY == 0:
            self._stop_excel(); time.sleep(1); self._start_excel()

    def convert_one(self, path, overwrite, delete_original):
        out=os.path.splitext(path)[0]+".xlsx"
        if os.path.exists(out) and not overwrite:
            return "skip"
        wb=None
        try:
            try:
                wb=self.xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0, ReadOnly=True)
            except Exception:
                # corrupt/odd file → try Excel's repair-open
                wb=self.xl.Workbooks.Open(os.path.abspath(path), UpdateLinks=0,
                                          ReadOnly=True, CorruptLoad=1)
            wb.SaveAs(os.path.abspath(out), FileFormat=XLSX)
            wb.Close(False); wb=None
            self._recycle_if_needed()
            if delete_original:
                try: os.remove(path)
                except Exception as e: self.log(f"      (kept original — delete failed: {e})","warning")
            return "ok"
        except Exception as e:
            if wb is not None:
                try: wb.Close(False)
                except Exception: pass
            # a bad file can wedge the instance — recycle it
            try: self._stop_excel(); self._start_excel()
            except Exception: pass
            return f"error: {e}"

    def run(self, files, overwrite, delete_original):
        self._start_excel()
        ok=skip=err=0
        try:
            for i,p in enumerate(files,1):
                if _CANCEL.is_set():
                    self.log("  ⏹ Cancelled by user.","warning"); break
                name=os.path.relpath(p, os.path.commonpath(files)) if len(files)>1 else os.path.basename(p)
                r=self.convert_one(p, overwrite, delete_original)
                if r=="ok":
                    ok+=1;  self.log(f"  [{i}/{len(files)}] ✅ {name}")
                elif r=="skip":
                    skip+=1; self.log(f"  [{i}/{len(files)}] ↷ {name} (xlsx exists)")
                else:
                    err+=1; self.log(f"  [{i}/{len(files)}] ❌ {name} — {r}","error")
        finally:
            self._stop_excel()
        return ok, skip, err


# ── GUI ────────────────────────────────────────────────────────────────────────
# Brand palette kept in sync with GFH_Inventory_Aging_Processor.pyw
NAVY  = "#090d26"
EMBEDDED_LOGO_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_logo_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_logo_b64.txt"), "r").read().strip()
EMBEDDED_ICON_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_icon_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_icon_b64.txt"), "r").read().strip()

RED   = "#e8212a"
WHITE = "#ffffff"
LIGHT = "#f6f7fb"
LOG_BG   = "#10182e"
LOG_FG   = "#a8d8ff"

ICON_ICO_NAME = "gfh_icon.ico"
LOGO_PNG_NAME = "GFH_Telecom_Logo.png"
COPYRIGHT_TEXT = f"Developed by Abad Umair Channa | Copyright © {date.today().year} | All rights reserved."
ICON_ICO_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon_ico_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "icon_ico_b64.txt"), "r").read().strip()


def _script_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _resource_path(name):
    """Resolve a bundled resource (logo PNG) from source or from a
    PyInstaller one-file EXE (extra files extract to _MEIPASS)."""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", _script_dir())
        return os.path.join(base, name)
    return os.path.join(_script_dir(), name)




def _extract_embedded_icon(b64, filename):
    """Decode an embedded base64 icon to a temp file; return path or None."""
    try:
        if not b64:
            return None
        import base64 as _b64, tempfile, os
        target = os.path.join(tempfile.gettempdir(), filename)
        with open(target, "wb") as fh:
            fh.write(_b64.b64decode(b64))
        return target if os.path.isfile(target) else None
    except Exception:
        return None

def _set_window_icon(root):
    """Set taskbar + titlebar icon from embedded base64 ICO."""
    import base64, tempfile, atexit, os, sys

    # 1. Try sys._MEIPASS (PyInstaller onefile extraction dir)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        ico_path = os.path.join(meipass, "gfh_icon.ico")
        if os.path.exists(ico_path):
            try:
                root.iconbitmap(ico_path)
                root.iconbitmap(ico_path)
                return
            except Exception:
                pass

    # 2. Try next to the exe/script
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, "gfh_icon.ico")
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
            root.iconbitmap(ico_path)
            return
        except Exception:
            pass

    # 3. Decode EMBEDDED_ICON_B64 to %TEMP% (no spaces, always writable)
    try:
        data = base64.b64decode(EMBEDDED_ICON_B64.strip())
        tmp_dir = os.environ.get("TEMP", tempfile.gettempdir())
        ico_path = os.path.join(tmp_dir, "gfh_app_icon.ico")
        with open(ico_path, "wb") as f:
            f.write(data)
        root.iconbitmap(ico_path)
        root.iconbitmap(ico_path)
        return
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root=root; self._q=queue.Queue(); self._busy=False
        root.title("GFH Telecom - Legacy Excel Converter")
        # Dynamic screen resolution support: size to 90% of the screen and
        # center it (DPI-aware), then stay a normal resizable top-level so
        # Windows Snap (50% left/right, corners, Win+arrow) keeps working.
        self._apply_dynamic_geometry()
        root.configure(bg=LIGHT)
        _set_window_icon(root)

        self._logo_img=None
        self.theme_manager = ThemeManager("GFH Legacy Excel Converter", app_name="gfh-xls-to-xlsx")
        self._styles(); self._header(); self._body(); self._copyright_bar(); self._poll()
        apply_theme_to_window(self.root, self.theme_manager, self._apply_theme)

    def _apply_dynamic_geometry(self) -> None:
        """Size the window to 90% of the screen and center it.

        Works on any laptop/monitor/PC (1080p, 1440p, 2K, 4K) and respects
        Windows DPI scaling (run after _enable_dpi_awareness()). The window
        stays resizable so Windows Snap gestures keep working — it centers
        on launch, then snaps normally to 50% left/right, corners or via
        Win+arrow shortcuts.
        """
        try:
            root = self.root
            root.update_idletasks()
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            w = max(640, min(int(sw * 0.90), sw - 20))
            h = max(480, min(int(sh * 0.90), sh - 40))
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            root.geometry(f"{w}x{h}+{x}+{y}")
            # minsize <= half the screen so 50% / corner snap is never blocked
            root.minsize(min(660, max(480, sw // 2)),
                         min(540, max(400, sh // 2)))
            root.resizable(True, True)
        except Exception:
            pass

    def _styles(self):
        s=ttk.Style(); s.theme_use("clam")
        s.configure("Run.TButton",background=RED,foreground=WHITE,
                    font=("Calibri",11,"bold"),padding=(16,9),borderwidth=0)
        s.map("Run.TButton",background=[("active","#c01820"),("disabled","#aaa")])
        s.configure("Browse.TButton",background=NAVY,foreground=WHITE,
                    font=("Calibri",10),padding=(10,6),borderwidth=0)
        s.map("Browse.TButton",background=[("active","#1a2550")])
        s.configure("Cancel.TButton",background="#1a2550",foreground=WHITE,
                    font=("Calibri",10),padding=(10,6),borderwidth=0)
        s.map("Cancel.TButton",background=[("active","#2a3560")])
        s.configure("Accent.Horizontal.TProgressbar",
                    troughcolor="#dde6f0",background=RED,borderwidth=0)


    def _extract_embedded(self, b64, filename):
        """Decode an embedded base64 asset into a temp file; return path or None."""
        try:
            if not b64:
                return None
            import base64 as _b64, tempfile, os
            target = os.path.join(tempfile.gettempdir(), filename)
            with open(target, "wb") as fh:
                fh.write(_b64.b64decode(b64))
            return target if os.path.isfile(target) else None
        except Exception:
            return None


    def _lock_header_colors(self, widget, navy):
        """Recursively bind <Enter>/<Leave> on all header widgets to force navy."""
        try:
            widget.bind("<Enter>", lambda e, w=widget, c=navy: w.configure(bg=c) if not isinstance(w, type(None)) else None)
            widget.bind("<Leave>", lambda e, w=widget, c=navy: w.configure(bg=c) if not isinstance(w, type(None)) else None)
        except Exception:
            pass
        try:
            for child in widget.winfo_children():
                self._lock_header_colors(child, navy)
        except Exception:
            pass
    def _header(self):
        """Header using FixedHeaderManager."""
        self.header_mgr = FixedHeaderManager(self.root, title="GFH Legacy Excel Converter")


    def _apply_theme(self, colors=None):
        """Re-apply the current theme across the window."""
        apply_theme_to_window(self.root, self.theme_manager)

    def _body(self):
        body=tk.Frame(self.root,bg=LIGHT)
        body.pack(fill="both",expand=True,padx=24,pady=18)

        # folder row
        fr=tk.Frame(body,bg=LIGHT); fr.pack(fill="x",pady=(0,14))
        fr.columnconfigure(0,weight=1)
        self.folder=tk.StringVar()
        tk.Entry(fr,textvariable=self.folder,font=("Calibri",9),
                 relief="flat",bg="#e8eff8",fg=NAVY,
                 readonlybackground="#e8eff8",
                 highlightbackground="#b0c4de",highlightthickness=1
                 ).grid(row=0,column=0,sticky="ew",ipady=5,padx=(0,8))
        ttk.Button(fr,text="Browse",style="Browse.TButton",
                   command=self._browse).grid(row=0,column=1)

        # options
        opt=tk.Frame(body,bg=LIGHT); opt.pack(fill="x",pady=(0,14))
        self.recurse=tk.BooleanVar(value=True)
        self.overwrite=tk.BooleanVar(value=True)
        self.delete=tk.BooleanVar(value=True)
        for txt,var in [("Include subfolders",self.recurse),
                        ("Overwrite existing .xlsx",self.overwrite),
                        ("Delete original after converting",self.delete)]:
            tk.Checkbutton(opt,text=txt,variable=var,font=("Calibri",10),
                           fg=NAVY,bg=LIGHT,selectcolor=WHITE,
                           activebackground=LIGHT,activeforeground=NAVY
                           ).pack(side="left",padx=(0,16))

        # action buttons
        act=tk.Frame(body,bg=LIGHT); act.pack(fill="x",pady=(0,12))
        self.run_btn=ttk.Button(act,text="▶  Convert",style="Run.TButton",
                                command=self._start)
        self.run_btn.pack(side="left")
        self.cancel_btn=ttk.Button(act,text="⏹  Cancel",style="Cancel.TButton",
                                   command=lambda:_CANCEL.set(),state="disabled")
        self.cancel_btn.pack(side="left",padx=8)
        self.pv=ttk.Progressbar(act,mode="determinate",
                                style="Accent.Horizontal.TProgressbar")
        self.pv.pack(side="left",fill="x",expand=True,padx=8)

        # log
        tk.Label(body,text="Activity Log",font=("Calibri",9,"bold"),
                 fg=NAVY,bg=LIGHT).pack(anchor="w")
        self.log_w=scrolledtext.ScrolledText(body,font=("Consolas",8),
                    bg=LOG_BG,fg=LOG_FG,relief="flat",wrap="word")
        self.log_w.pack(fill="both",expand=True)
        for tag,clr in [("info","#90CDF4"),("success","#68D391"),
                        ("error","#FC8181"),("warning","#F6E05E")]:
            self.log_w.tag_config(tag,foreground=clr)

    def _copyright_bar(self):
        bar=tk.Frame(self.root,bg=NAVY,height=26)
        bar.pack(fill="x",side="bottom"); bar.pack_propagate(False)
        tk.Label(bar,text=COPYRIGHT_TEXT,bg=NAVY,fg="#9d9db8",
                 font=("Calibri",8)).pack(pady=4)

    def _browse(self):
        d=filedialog.askdirectory(title="Select folder with legacy Excel files")
        if d: self.folder.set(d)

    def _log(self,m,tag=""): self._q.put(("log",m,tag))
    def _poll(self):
        try:
            while True:
                it=self._q.get_nowait()
                if it[0]=="log":
                    self.log_w.insert(tk.END,f"[{datetime.now():%H:%M:%S}]  {it[1]}\n",it[2] or ())
                    self.log_w.see(tk.END)
                elif it[0]=="prog":
                    self.pv["maximum"]=it[1]; self.pv["value"]=it[2]
                elif it[0]=="done":
                    self._busy=False; self.cancel_btn.config(state="disabled")
        except queue.Empty: pass
        self.root.after(80,self._poll)

    def _start(self):
        if self._busy: return
        folder=self.folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("No folder","Please browse to a valid folder first."); return
        _CANCEL.clear(); self._busy=True; self.cancel_btn.config(state="normal")
        threading.Thread(target=self._run,args=(folder,),daemon=True).start()

    def _run(self,folder):
        log=self._log
        log("="*54); log(f"Converter started {datetime.now():%Y-%m-%d %H:%M:%S}")
        log(f"Folder: {folder}"); log("="*54)
        try:
            files=_find_files(folder, self.recurse.get())
            log(f"  Legacy Excel files found: {len(files)}","info")
            if not files:
                log("  Nothing to convert.","warning"); self._q.put(("done",)); return
            self._q.put(("prog",len(files),0))
            conv=Converter(log)
            # progress wrapper
            t0=time.time()
            def prog(i): self._q.put(("prog",len(files),i))
            ok=skip=err=0
            conv._start_excel()
            try:
                base=os.path.commonpath(files) if len(files)>1 else folder
                for i,p in enumerate(files,1):
                    if _CANCEL.is_set(): log("  ⏹ Cancelled.","warning"); break
                    name=os.path.relpath(p,base)
                    r=conv.convert_one(p, self.overwrite.get(), self.delete.get())
                    if r=="ok": ok+=1; log(f"  [{i}/{len(files)}] ✅ {name}")
                    elif r=="skip": skip+=1; log(f"  [{i}/{len(files)}] ↷ {name} (xlsx exists)")
                    else: err+=1; log(f"  [{i}/{len(files)}] ❌ {name} — {r}","error")
                    prog(i)
            finally:
                conv._stop_excel()
            log("\n"+"="*54)
            log(f"  Done in {time.time()-t0:,.0f}s — converted {ok}, skipped {skip}, failed {err}.",
                "success" if err==0 else "warning")
            log("="*54)
        except Exception as e:
            log(f"[FATAL] {e}","error"); log(traceback.format_exc(),"error")
        self._q.put(("done",))


def _enable_dpi_awareness() -> None:
    """Make Windows report physical pixels so winfo_screen* is accurate on
    high-DPI displays (1080p, 1440p, 2K, 4K, DPI-scaled laptops)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    root=tk.Tk(); App(root); root.mainloop()

if __name__=="__main__":
    main()
