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
EMBEDDED_LOGO_B64 = "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O3rtmg3FUu2Sj13u+8Gbbz79yMPPnn7m2Z+47NM/+OHV73/fBxYsXjh3zqyZf1tw+ac/k05nq9Xek6cfd9NNN2zYsDWO65MnT7z+hz/9znXXXnD+R3/ykx+sWLGC2U2cMPErX7n6plt+/NCDT0w7/NBjjj3lzbmv/eB73zvttBn77LPfzJmv7DFi5JFHnrxs5bzvXvu9Y4457AMfOOmN2a+mM5nVq9Y3N+evuurbL7380gP33zNy1PAlSxYfeOCBTzzx9Isv/O3CCy9oa2vLZjMbNmx88MHHa1H129/+6quvvtHW1vzoY0/8+pZbX3nluQ0bNpdK5XETRt11131fvvJazw95W3Zuar/6xQ1HHvm+RYsWH3f8Md+46juT99133333Hjd29MaNmyuV6heuuHLZ0qWrVy9evnz5oYceefxx0++66zenn3HeX5566OEHH0mFLaecei55Cojieu+0ww56+ukH7r//kQs+/EnrzPeu/c4JJxxz+LQTGcy5Hzz72mu/Om7C5Ft/84uTTznxzTfmDeoYcuut9/zXj34cpryn//LggQfuN3nKESuWLd97n73/+uxjGzatLRVLk/bZ5557HvzSlVcJBszgxP5j2jTeIdAXBw4dOJ16ujs+JKcPzTSnap1BcePhzcOX17KPVBMIUizSE+buKHRXJTW9dVA+6c3ZClU3j/XDYW1NU9pbl9fMothsMraccCwoqABFo2QJ9vCCvUM1HuMOqahyve6SovIkCCtWCpZ6vLZ1jlbWojUuMaBOzqYPwpqf9JR8rzvXurJo54mfiKfjpL49ydqlanTqgVOOPe7Iww8/5o03Xrn88itqEZ17/r8NHzbu+ece/tznvv7iK3+zpnLN1Vdt2rRln0nj99v34F/e9Ivzzjvr3z97yTPPvtDe1vGNq7+lPe3Y5bLB9df/x1NPP3vhhZ9nNlde8e+XXnrhnx+495pvf+nue+7+9Ke/ACjf/+73vv/9/3j0sYcKhc58U/DD668+/rhjo3ohSWqkoVDo3n//va/9zpc/dN55xhSieiVJqs7Zq6665re/+xWpPLvKT37yozFjRx191MlLl8899JCj7r33d08++fzhRx1/zTe+cvChB08/+Wx28TXf/NZLL7563PEfECAiGjd+ggh/8lNfevW1V089+aR77r55/psrf3vn70mHjf6RYUOGn3nmKR+5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QWeQvX3zxtbBg4/IDy5b1i6ewPaiIYPyhcrD1XLZCz0g4+VKPsw0tUWFaFyJx2l/ZBCO9lJ5MEoJaREQB+IErXgFQ2uT+krhFa62jlVVpSEgQCYgsuV2W3tfS3hYgC21rTEZdmHk52bG9HgSd6UylDh4tyo0UKRLvYWzzznzpJOPXrly1fvff9jXv34tkaeUn07nBLijffgHzzr9k5+68u57fu978sILL1522WWvvPpUW2tHPt8srIDBitGeVyoXv/DFr//m1zdO3f+Q3kJh6gH7/Of1P1m48PWrr77uxht/eNTRh4LQPpMmf+XKqzdsWNna0rF+w2ZU4Te+8d1jjz2qtbUDRTfnWzyVRvK/dfX10086rq21XZGPoK699ppLL724paXlG9/87jXf/o8jph3ywl8fW7R40X77TVkwf+GDDz3AzqD2giCFAETaOjnyyMOffebZ1taWv/zl+Ztu/k0Y+jffdEN3T+d++069+eY77rzz7scff+z11+d86cqvh6mWQm/hmGOOPefcM19/ffYRR0674b9uATRIXi7T5PuhVv6ll158ww23fP+H3yPy/vSHOy679GP33X9XJtOUzWQRtXVWexRFlTFjJp5yysmXXnLZfX++N/TTb7w586Mf/XAcxW2trZnUUEZubm7P5Zp8P8uOPn7pJQcecNDQoUO+d91PTjt1+mszZ5986ukC+OUrrvz3f7/soYced1Z+8YsbCr3dkyfve+cd999x1x9SqWa3fd/s967s2WmNUgg0IwKUNK0ud2cy2Y50G1XroXEZW9+zNdOuMSmWbCLFlA8InnUag6rVy2L7WhzPrXpvVPzX6vRmpF+P9MyafrXm/a3GL9fqrzhZIGorehZUM3spUYZsJilPpNoH2lOT05KOyppNLF5PpnmmU49Walv9FAqC8LuvIintdXVtue++B4U5qsU/uuGmP9//aBDmTWJXrVg7a/Yb2g+WL1v54MNPaZ21lufPW9HTXV28eE1Pb8+bb8xbtWa99NdLiPTihYseffRxZ6izs/Oab33vrrvvyzcNm/na7KeefF6RWr589bevvv6Bhx8O0oPWrdn85psL1m/YWqlFM1+d/frsBUuXrdm4fsvM1+Z0dld6CoWZf5v9xux5K1asWb9u8/x5ixYtWjLnzQVvzFm6bt3m++57tKerJ47dXXfc+61rvl+rMaqwVCrOm7dgxap1SuneQu/CBQsWzF+8cMHKuW8uXbp05aoVG+fPX7Bs6aof//jWX//6DmPrhx16uAg99vhjYSpXqVXv/dMD9ShmwJt++bvbfnuXH2ad5Y3rO2e+9jqDv2b1uj/88X7AwDEuXbJs06auRUtWlgrFhQuWLFm6atvQ5Xy+aeni5U888Qxh2rJatmTZli3dc+YsfH32/KXL1whAHPG8uUsWLl7c01OaM2fu/PnzZ8+ev3jJqpUr1zz8yJObtvQEQXbF8pUb1m9dsXLd3DkL3nxz9py5i37x89t++vNbADzPCxoNL7u6v+95T0/sVwN6gAmxDe3YQvXfm0e+L8/1ZKvHkU44Tg2a7+f/0lt+s1orAlWVZiREBSDsrJBq6ARJGJEZkAVBGtv4ADUEJQgAGFg72vDkrLd/mgcnpbStO2ZRQVW1/E3kwVqtx29Srq/OJPxup8lhQw0cR8KRgCDqwG9q6FHYRUQBAlhXzYStjX1C6/UIgFNhLqqXANj3s/3CGkIRRVhPKiJGWEiFQZgDBkW6GhURDVsD6GWzLdbaOK4hoR9kUaQeV0FsEGQTUyEIgyAlyLVqGQlyqWy5VhKpI5Kw87z2dCofm2oc9wIwoAq9ZiB0wDaJ2NWDsBmRjKmzqyBqJC2iQi8f24KIYSdK61TQ7Ez8vvcfvHDhwt7ekta+UjqJ64krISoA8nRGKQ8A6vUCUqC1NqZMlNXa00TVqEQkSmdsXBFwfpAHQAYmRGsMuyiTanHCgpjEJRQPUETY89ICwGJtYsIgFSdVgKRv1oRKCTtAFQYZECGicrk3k87XoghVjEhEgaIAsCGlQLdrffA/vjN4Y6e+RpltVGzPb9X7D82mSpva4zomGAXN5TA/F+J5pdq8sqx0fs3zkISgIflm6JtSCiSAQALSvzkEaeGss+1IYwN8f0Z1uLglKgWmEhGU035vkH++NmROqdQVkhXo877SmA+C724VAWDYsWDXr158y1USYSKNIkBEiNaZxoCf3Y3XIOLGKEvso6KVUs65hrBLkxKExpC1RgtTY4ojETnnCBvze8A5q7XXKB07YeEGwcKKqG9fCGu39V8gNAa7SKNazn3CRWKWRv2ZCK11iIQo7Jw0Ngvr22jsLZQjEW3b//Zt/Sxak9vp+Bhs1H69RrfetmkVhKi0buzkQESKyFirFCFSo7mQCBGpb4ilgNKkSFlrSSlrrVJKtsv1/45I45+ztz0idNRLRzRlTmtrHR0VQ1OtUhxpySW5CDOrdXpeInNr5fU2KkmcYBij51AYSUBIgAA8Bl84RaqF3Z5aj9NqtMJmqHtYgiQBZuuH3UF2Per5PfEbLlMKlGGzbXj0wCTof1n75yC40W2Yrdf3BzxxcPv4ME67Tu3KuVrKGTShV/d0UQWdTL2JW2tcL0vsTCwWlGo4xCbRg1XQrFWrhqyLAqgorviJyddVRVE517RBZxcl3pyi7QRd1WKxfwxIv0x9AMEDCP7vAZi0ELE1TSZ5X847MqvGaafFcFINbU2xsaCNysYq67PSzImwQTYogsgIVjlLJqhzYCywTdDWyALoENq7NK0EmBu5JcYrUgoUiY2hMaVX+ufxDyB4AMH/PX0zoOqfHA2UT8xwm4zLpw5K48gQM64aJFUlaFnH7HlSJ0mYxQlbFCa0CEIGxHhWKeOxhNbPReRvIW+RCtbWyiuTao8OYhWIAAKxc/3BqwwAdwDB/xwfjArACYJKoWaEmK2w3cPW9vD16FQ4PhWOQMonkhUSqLBEyCCOGcCiWGQlpBxWEUtKd6twA+vNkawxbjG5BNEpJQgILIIDoB2w/x0+GH3wDRoBaeyRDU4UKgcg4ELr8gJDlAz1oT3AFgUZAB+UD8QiVtiAq0jQ7ajX1Xo56WYpWqqJF2vPEQIDCfY1JvR1HG7f5WUAzgMI/ucgWIOy4AT7mhdBhIAazO72FkKwyMYn9JE0EjXaDAEAJAGqQWPLlMYGE4TbR4L0d6Rv428Ggt4B25aA/VM+RUBMY2jXDu22DraLM/sJPS2kY4B4FyON+zukdiLsfeuOKQM2YP9UBL8LiL817dvZMyD/qruzD9j/Awh+G5p36X8HbMDeo9HAJRiwAQQP2IANIHjABmwAwQM2gOABG7ABBA/YgA0geMAGbADBA/YvYf8fMD1ywfapnOUAAAAASUVORK5CYII="
EMBEDDED_ICON_B64 = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAA0GRv/NBkb/zQZG/80GRv/NBkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zQZG/81GRv/NBgb/zUZG/81GRv/NBkb/zQZG/80GBv/NRkb/zQYG/80GRv/NBkb/zQZG/80GRv/NBkb/zQZG/81GRv/NRkb/zQZG/80GRv/NBgb/zUZG/80GRv/NBkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GBv/NRgb/zUYG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zYZHP81GRv/NRkb/zUYG/81GRv/NRkb/zUZG/81GRv/NRkc/zUZG/81GRz/Nhkc/zYZG/82GRz/Nhkc/zYZHP82GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUYG/8/JCf/TDM2/zYaHP81GRv/NRgb/zUZG/81GRv/NRkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkb/zYZHP81GRv/NRkb/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkb/zYZG/82GRz/NRkb/zoeIP8/JCf/Nhkb/zUZG/81GRv/Nhkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zUYG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZG/82GRv/Nhkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZG/82GRr/NRkh/zMZKv8zGS3/Mxkn/zUZHv82GRr/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Mxoy/yseZv8lIIn/IyCS/yIfkf8iHoz/JB16/ywbUP80GSX/Nhka/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhke/y8eWf8mI53/JyGL/y0dW/8wG0H/MBo7/y4aRf8pHGT/IR6O/yMdg/8wGjr/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHf8uIGj/JyWn/y8eWf80Giv/MRxH/ywfaP8rH3D/LR1e/zIaOf8zGS3/Jxxr/yEekv8wGj//Nhka/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRr/MR9S/ykorP8xHk3/Mxw4/yojjf8nJKD/KSGD/ysgd/8oIIj/JSKb/yoecP8vG0f/KB1t/yMfjv8zGi7/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zcZHP83GRz/Nxkb/zUaKP8sKKL/MCJu/zQbM/8qJqD/LCOE/zQaMv83GRv/Nxkb/zYZHf8uHVn/IyOp/yQhnP8sHWX/Jh+G/yodav82GRv/Nxkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zcZHP82GRz/Nxkc/zcZHP83GRz/Nxkc/zcZHP83GRn/NCBT/y4rsP81GzL/LyWD/y0mkv82GiP/NRsu/y8hcP8sI4T/MB5Z/zAdSv8mI6H/JiOc/ywfbP8vHEz/JiGW/zQaKv83GRv/Nxkc/zcZHP83GRz/Nxkc/zcZHP82GRz/Nxkc/zYZHP83GRz/Nxkc/zYZHP83GRz/Nxkc/zcZGv8yJX3/MSeR/zUbNf8uKrL/NB5K/zYaJv8tJpb/LCaY/y4hdv8pJqT/LiFy/y4eX/8sH3D/NBou/zQaK/8mIpn/MRxC/zcZGv83GRz/Nxkc/zcZHP83GRz/Nxkc/zcZHP82GBv/Nhgb/zYYG/82GRz/Nhkc/zYZHP82GRz/Nxkb/zIplP8zJnr/NB5I/y8ss/81Giz/Mx9P/y0qrv81GzD/NxgX/zEeUf8pKLD/LSN//y0iev8tInn/LCF9/yclq/8wHU3/NxgZ/zcZHP83GRz/Nxkc/zcZHP83GRz/Nhkc/zYYG/82GBv/Nhgc/zYZHP82GBz/Nhgb/zYYG/82GR//Miyi/zIrnf80IVv/MS62/zUbLP8zH0//Lyyy/zUbMf82GBf/Mh9U/ysqtP8uI4D/LiJ7/y0iev8tIXn/Lh9r/zQaK/82GBv/Nhgb/zYYHP82GRz/Nhkc/zYYHP82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/NhgZ/zUfSf8yM8X/MjHC/zIrmv8xMLv/NCBO/zUZJv8xKp//MCul/zElgP8uK7H/MSNz/zYYGf82GBf/NhgY/zUaK/8zHUL/Nhgc/zYYG/82GBv/Nhgc/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBn/NSBL/zM0yf8yM8X/MyiB/zMqkf8yLaP/Nhok/zUbL/8yJXr/MSiQ/zIiY/81GSH/NRst/zMfUv80HDX/Mxw+/zEiaP82GSD/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GR//NSJb/zQpiv81I1z/NRw5/zIwuf8zLJr/NRw3/zYYGv80H0n/MimS/zMmfP8wK6j/Lyy4/y8qrv80HT7/Nhke/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBb/NSFS/zMzyP81Ilf/NR09/zIuqP8yMsT/Myh//zEtpf8xLrP/MDLP/zInh/80HDn/MCyx/zMiZv82GBj/Nhgb/zYYG/82GBv/Nxgb/zYYG/82GBv/Nxgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GR7/NSh//zQ0zf81JWr/NRsu/zQhVf80I2L/MyeE/zExxf8yLar/NB1D/zInh/8wL7v/NB1A/zYYGf82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYGv82GR//NSZv/zQ0yv80MLL/NSZy/zUgT/81IVP/MyiI/zIsnf8xMLz/MS6v/zQfSv82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBr/NR07/zUogv80MLX/NDLE/zQyxf8zMb3/Myyj/zQjZ/81Gin/NhgZ/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBn/NhgZ/zUZIv82Gy//NRw0/zUaK/82GB7/NhgY/zYYGv82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/83GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYGv82GBr/Nhga/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/NRgb/zUYG/81GBv/NRgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zUYG/81GBv/NRgb/zYYG/81GBv/Nhgb/zYYG/81GBv/NRgb/zUYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/81GBv/NRgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/82GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zYZG/81GBv/NRgb/zUYG/81GBv/NRgb/zUZG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="

RED   = "#e8212a"
WHITE = "#ffffff"
LIGHT = "#f0f4fa"
LOG_BG   = "#10182e"
LOG_FG   = "#a8d8ff"

ICON_ICO_NAME = "GFH_Telecom_TBLogo.ico"
LOGO_PNG_NAME = "GFH_Telecom_Logo.png"
COPYRIGHT_TEXT = f"Developed by Abad Umair Channa  |  Copyright © {get_copyright_year()}  |  All rights reserved."
ICON_ICO_B64 = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAA0GRv/NBkb/zQZG/80GRv/NBkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zQZG/81GRv/NBgb/zUZG/81GRv/NBkb/zQZG/80GBv/NRkb/zQYG/80GRv/NBkb/zQZG/80GRv/NBkb/zQZG/81GRv/NRkb/zQZG/80GRv/NBgb/zUZG/80GRv/NBkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GBv/NRgb/zUYG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zYZHP81GRv/NRkb/zUYG/81GRv/NRkb/zUZG/81GRv/NRkc/zUZG/81GRz/Nhkc/zYZG/82GRz/Nhkc/zYZHP82GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUYG/8/JCf/TDM2/zYaHP81GRv/NRgb/zUZG/81GRv/NRkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkb/zYZHP81GRv/NRkb/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkb/zYZG/82GRz/NRkb/zoeIP8/JCf/Nhkb/zUZG/81GRv/Nhkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zUYG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZG/82GRv/Nhkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZG/82GRr/NRkh/zMZKv8zGS3/Mxkn/zUZHv82GRr/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Mxoy/yseZv8lIIn/IyCS/yIfkf8iHoz/JB16/ywbUP80GSX/Nhka/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhke/y8eWf8mI53/JyGL/y0dW/8wG0H/MBo7/y4aRf8pHGT/IR6O/yMdg/8wGjr/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHf8uIGj/JyWn/y8eWf80Giv/MRxH/ywfaP8rH3D/LR1e/zIaOf8zGS3/Jxxr/yEekv8wGj//Nhka/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRr/MR9S/ykorP8xHk3/Mxw4/yojjf8nJKD/KSGD/ysgd/8oIIj/JSKb/yoecP8vG0f/KB1t/yMfjv8zGi7/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zcZHP83GRz/Nxkb/zUaKP8sKKL/MCJu/zQbM/8qJqD/LCOE/zQaMv83GRv/Nxkb/zYZHf8uHVn/IyOp/yQhnP8sHWX/Jh+G/yodav82GRv/Nxkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zcZHP82GRz/Nxkc/zcZHP83GRz/Nxkc/zcZHP83GRn/NCBT/y4rsP81GzL/LyWD/y0mkv82GiP/NRsu/y8hcP8sI4T/MB5Z/zAdSv8mI6H/JiOc/ywfbP8vHEz/JiGW/zQaKv83GRv/Nxkc/zcZHP83GRz/Nxkc/zcZHP82GRz/Nxkc/zYZHP83GRz/Nxkc/zYZHP83GRz/Nxkc/zcZGv8yJX3/MSeR/zUbNf8uKrL/NB5K/zYaJv8tJpb/LCaY/y4hdv8pJqT/LiFy/y4eX/8sH3D/NBou/zQaK/8mIpn/MRxC/zcZGv83GRz/Nxkc/zcZHP83GRz/Nxkc/zcZHP82GBv/Nhgb/zYYG/82GRz/Nhkc/zYZHP82GRz/Nxkb/zIplP8zJnr/NB5I/y8ss/81Giz/Mx9P/y0qrv81GzD/NxgX/zEeUf8pKLD/LSN//y0iev8tInn/LCF9/yclq/8wHU3/NxgZ/zcZHP83GRz/Nxkc/zcZHP83GRz/Nhkc/zYYG/82GBv/Nhgc/zYZHP82GBz/Nhgb/zYYG/82GR//Miyi/zIrnf80IVv/MS62/zUbLP8zH0//Lyyy/zUbMf82GBf/Mh9U/ysqtP8uI4D/LiJ7/y0iev8tIXn/Lh9r/zQaK/82GBv/Nhgb/zYYHP82GRz/Nhkc/zYYHP82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/NhgZ/zUfSf8yM8X/MjHC/zIrmv8xMLv/NCBO/zUZJv8xKp//MCul/zElgP8uK7H/MSNz/zYYGf82GBf/NhgY/zUaK/8zHUL/Nhgc/zYYG/82GBv/Nhgc/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBn/NSBL/zM0yf8yM8X/MyiB/zMqkf8yLaP/Nhok/zUbL/8yJXr/MSiQ/zIiY/81GSH/NRst/zMfUv80HDX/Mxw+/zEiaP82GSD/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GR//NSJb/zQpiv81I1z/NRw5/zIwuf8zLJr/NRw3/zYYGv80H0n/MimS/zMmfP8wK6j/Lyy4/y8qrv80HT7/Nhke/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBb/NSFS/zMzyP81Ilf/NR09/zIuqP8yMsT/Myh//zEtpf8xLrP/MDLP/zInh/80HDn/MCyx/zMiZv82GBj/Nhgb/zYYG/82GBv/Nxgb/zYYG/82GBv/Nxgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GR7/NSh//zQ0zf81JWr/NRsu/zQhVf80I2L/MyeE/zExxf8yLar/NB1D/zInh/8wL7v/NB1A/zYYGf82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYGv82GR//NSZv/zQ0yv80MLL/NSZy/zUgT/81IVP/MyiI/zIsnf8xMLz/MS6v/zQfSv82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBr/NR07/zUogv80MLX/NDLE/zQyxf8zMb3/Myyj/zQjZ/81Gin/NhgZ/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBn/NhgZ/zUZIv82Gy//NRw0/zUaK/82GB7/NhgY/zYYGv82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/83GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYGv82GBr/Nhga/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/NRgb/zUYG/81GBv/NRgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zUYG/81GBv/NRgb/zYYG/81GBv/Nhgb/zYYG/81GBv/NRgb/zUYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/81GBv/NRgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/82GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zYZG/81GBv/NRgb/zUYG/81GBv/NRgb/zUZG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="


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


def _set_window_icon(root):
    """Set taskbar + titlebar icon from the embedded GFH_Telecom_TBLogo.ico."""
    try:
        import base64, tempfile, atexit
        data = base64.b64decode(ICON_ICO_B64.strip())
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".ico")
        tmp.write(data); tmp.close()
        atexit.register(lambda p=tmp.name: os.path.exists(p) and os.unlink(p))
        root.iconbitmap(default=False, bitmap=tmp.name)
        root.iconbitmap(tmp.name)
        return
    except Exception:
        pass
    # Fallback: use the brand PNG as the window icon
    png_path = _resource_path(LOGO_PNG_NAME)
    try:
        if os.path.exists(png_path):
            from PIL import Image as _PI, ImageTk as _PIT
            root.iconphoto(True, _PIT.PhotoImage(_PI.open(png_path)))
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root=root; self._q=queue.Queue(); self._busy=False
        root.title("GFH Legacy Excel Converter")
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
