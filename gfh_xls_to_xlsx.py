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
from theme_manager import ThemeManager, apply_theme_to_window, get_copyright_year, create_theme_toggle_button
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
ICON_ICO_B64 = "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"


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
        self.theme_manager = ThemeManager("GFH Legacy Excel Converter")
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
        """Header: NAVY bar 108px tall with logo on the left, a truly centered
        title/subtitle block, and a theme toggle on the right.

        The title block is centered relative to the full header width (relx=0.5)
        so the heading sits in the visual middle of the bar instead of being
        pushed right by the logo.
        """
        hdr=tk.Frame(self.root,bg=NAVY,height=108)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        hdr._tag = "header"

        # Load logo from GFH_Telecom_Logo.png next to this script, composite on
        # NAVY, thumbnail to 260x82 (same recipe as Aging Processor).
        logo_path=_resource_path(LOGO_PNG_NAME)
        if os.path.exists(logo_path):
            try:
                from PIL import Image as _PI,ImageTk as _PIT
                img=_PI.open(logo_path).convert("RGBA")
                bg2=_PI.new("RGBA",img.size,(9,13,38,255))
                bg2.paste(img,mask=img.split()[3])
                img=bg2.convert("RGB"); img.thumbnail((260,82),_PI.Resampling.LANCZOS)
                self._logo_img=_PIT.PhotoImage(img)
            except Exception:
                self._logo_img=None

        lf=tk.Frame(hdr,bg=NAVY); lf.place(relx=0,rely=0.5,anchor="w",x=24)
        lf._tag = "header"
        if self._logo_img:
            _ll = tk.Label(lf,image=self._logo_img,bg=NAVY)
            _ll._tag = "logo"
            _ll.pack()
        else:
            _ll = tk.Label(lf,text="GFH TELECOM",font=("Calibri",16,"bold"),
                     fg=RED,bg=NAVY)
            _ll._tag = "logo"
            _ll.pack()

        # Centered title block — relx=0.5 anchors it to the true horizontal
        # middle of the header regardless of logo width or window size.
        tf=tk.Frame(hdr,bg=NAVY); tf.place(relx=0.5,rely=0.5,anchor="center")
        tf._tag = "header"
        _t1 = tk.Label(tf,text="LEGACY EXCEL CONVERTER",
                 font=("Calibri",18,"bold"),fg=WHITE,bg=NAVY)
        _t1._tag = "logo"
        _t1.pack()
        _t2 = tk.Label(tf,text="Convert .xls / .xlsm / .xlt / .xlsb → .xlsx using real Excel",
                 font=("Calibri",9),fg=WHITE,bg=NAVY)
        _t2._tag = "logo"
        _t2.pack()

        # One-click light/dark theme toggle in the header's right side
        theme_btn = create_theme_toggle_button(hdr, self.theme_manager, on_toggle=self._apply_theme)
        theme_btn.place(relx=0.98, rely=0.5, anchor="e")

        self._lock_header_colors(hdr, NAVY)

        self._lock_header_colors(hdr, NAVY)

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
