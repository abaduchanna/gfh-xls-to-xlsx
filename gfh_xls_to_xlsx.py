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
from tkinter import ttk, scrolledtext, messagebox, filedialog
import win32com.client

# ── GFH brand assets/constants (icon + top-right logo header) ─────────────────
COLOR_NAVY = "#161632"
COLOR_RED = "#E91B2F"
COLOR_WHITE = "#ffffff"
COPYRIGHT_TEXT = "Created by Abad Umair Channa  |  Copyright © 2026  |  All rights reserved."
ICON_ICO_NAME = "gfh_icon.ico"
ICON_PNG_NAME = "gfh_icon.png"
WORDMARK_PNG_NAME = "gfh_wordmark.png"

def get_script_dir():
    return (os.path.dirname(sys.executable)
            if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__)))

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
class App:
    def __init__(self, root):
        self.root=root; self._q=queue.Queue(); self._busy=False
        root.title("GFH Legacy Excel Converter"); root.geometry("920x660")
        root.configure(bg="#10142B")

        self._wordmark_img = None
        self._set_window_icon(root)
        self._build_brand_header(root)

        top=tk.Frame(root,bg="#10142B"); top.pack(fill="x",padx=16,pady=(12,6))
        tk.Label(top,text="Legacy Excel Converter  (.xls → .xlsx)",
                 font=("Segoe UI",18,"bold"),fg="#FFFFFF",bg="#10142B").pack(anchor="w")
        tk.Label(top,text="Pick a folder. Every legacy Excel file inside is converted to .xlsx using real Excel.",
                 font=("Segoe UI",9),fg="#A0AEC0",bg="#10142B").pack(anchor="w")

        # folder row
        fr=tk.Frame(root,bg="#10142B"); fr.pack(fill="x",padx=16,pady=4)
        self.folder=tk.StringVar()
        tk.Entry(fr,textvariable=self.folder,font=("Segoe UI",10),width=70).pack(side="left",fill="x",expand=True)
        s=ttk.Style(); s.theme_use("clam")
        s.configure("A.TButton",background="#F02428",foreground="#fff",
                    font=("Segoe UI",10,"bold"),borderwidth=0,padding=(14,7))
        s.map("A.TButton",background=[("active","#C81D21")])
        s.configure("D.TButton",background="#1A1F3A",foreground="#fff",
                    font=("Segoe UI",10,"bold"),borderwidth=0,padding=(14,7))
        ttk.Button(fr,text="Browse…",style="D.TButton",command=self._browse).pack(side="left",padx=(8,0))

        # options
        opt=tk.Frame(root,bg="#10142B"); opt.pack(fill="x",padx=16,pady=4)
        self.recurse=tk.BooleanVar(value=True)
        self.overwrite=tk.BooleanVar(value=False)
        self.delete=tk.BooleanVar(value=False)
        for txt,var in [("Include subfolders",self.recurse),
                        ("Overwrite existing .xlsx",self.overwrite),
                        ("Delete original .xls after converting",self.delete)]:
            tk.Checkbutton(opt,text=txt,variable=var,font=("Segoe UI",9),
                           fg="#E2E8F0",bg="#10142B",selectcolor="#1A1F3A",
                           activebackground="#10142B",activeforeground="#fff").pack(side="left",padx=(0,16))

        # action buttons
        act=tk.Frame(root,bg="#10142B"); act.pack(fill="x",padx=16,pady=6)
        ttk.Button(act,text="▶  Convert",style="A.TButton",command=self._start).pack(side="left")
        self.cancel_btn=ttk.Button(act,text="⏹  Cancel",style="D.TButton",
                                   command=lambda:_CANCEL.set(),state="disabled")
        self.cancel_btn.pack(side="left",padx=8)
        self.pv=ttk.Progressbar(act,mode="determinate"); self.pv.pack(side="left",fill="x",expand=True,padx=8)

        self.log_w=scrolledtext.ScrolledText(root,font=("Consolas",9),wrap=tk.WORD,
                    bg="#0B0E20",fg="#E2E8F0",relief="flat")
        self.log_w.pack(fill="both",expand=True,padx=16,pady=12)
        for tag,clr in [("info","#90CDF4"),("success","#68D391"),
                        ("error","#FC8181"),("warning","#F6E05E")]:
            self.log_w.tag_config(tag,foreground=clr)

        self._build_copyright_bar(root)
        self._poll()

    # ---- GFH branding: window icon (titlebar + taskbar) --------------------
    def _set_window_icon(self, root):
        icon_dir = get_script_dir()
        ico_path = os.path.join(icon_dir, ICON_ICO_NAME)
        png_path = os.path.join(icon_dir, ICON_PNG_NAME)
        try:
            if os.path.exists(ico_path):
                root.iconbitmap(ico_path)
        except Exception:
            pass
        try:
            if os.path.exists(png_path):
                self._icon_img = tk.PhotoImage(file=png_path)
                root.iconphoto(True, self._icon_img)
        except Exception:
            pass

    # ---- GFH branding: top navy strip with logo pinned top-right -----------
    def _build_brand_header(self, root):
        header = tk.Frame(root, bg=COLOR_NAVY, height=54)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        wordmark_path = os.path.join(get_script_dir(), WORDMARK_PNG_NAME)
        logo_holder = tk.Frame(header, bg=COLOR_NAVY)
        logo_holder.pack(side="right", padx=(0, 18), pady=8)

        if os.path.exists(wordmark_path):
            try:
                self._wordmark_img = tk.PhotoImage(file=wordmark_path)
                tk.Label(logo_holder, image=self._wordmark_img, bg=COLOR_NAVY).pack()
            except Exception:
                self._wordmark_img = None

        if self._wordmark_img is None:
            badge = tk.Label(
                logo_holder, text="G", bg=COLOR_NAVY, fg=COLOR_RED,
                font=("Segoe UI", 16, "bold"), width=2, highlightthickness=2,
                highlightbackground=COLOR_RED, highlightcolor=COLOR_RED,
            )
            badge.pack(side="left", padx=(0, 6))
            word_frame = tk.Frame(logo_holder, bg=COLOR_NAVY)
            word_frame.pack(side="left")
            tk.Label(word_frame, text="GFH", bg=COLOR_NAVY, fg=COLOR_RED,
                     font=("Segoe UI", 11, "bold")).pack(anchor="w")
            tk.Label(word_frame, text="TELECOM", bg=COLOR_NAVY, fg=COLOR_WHITE,
                     font=("Segoe UI", 7, "bold")).pack(anchor="w")

    def _build_copyright_bar(self, root):
        bar = tk.Frame(root, bg=COLOR_NAVY, height=24)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        tk.Label(bar, text=COPYRIGHT_TEXT, bg=COLOR_NAVY, fg="#9d9db8",
                  font=("Segoe UI", 8)).pack(pady=3)

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


def main():
    root=tk.Tk(); App(root); root.mainloop()

if __name__=="__main__":
    main()
