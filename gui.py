"""Simple window: pick PDF + Excel -> preview extracted values -> write row."""
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import extractor as ex


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Trax CDC Export Report - PDF to Excel")
        self.geometry("760x620")
        self.pdf, self.xlsx = tk.StringVar(), tk.StringVar()
        self.extras = tk.BooleanVar()
        self.res = None
        for i, (lbl, var, cmd) in enumerate((("Shipment PDF", self.pdf, self.pick_pdf),
                                             ("CDC Excel", self.xlsx, self.pick_xlsx))):
            ttk.Label(self, text=lbl).grid(row=i, column=0, padx=8, pady=6, sticky="w")
            ttk.Entry(self, textvariable=var, width=70).grid(row=i, column=1)
            ttk.Button(self, text="Browse", command=cmd).grid(row=i, column=2, padx=6)
        ttk.Checkbutton(self, text="Also fill unmarked columns E (flight date) and J (procedure)",
                        variable=self.extras).grid(row=2, column=1, sticky="w")
        self.gem = tk.BooleanVar()
        ttk.Checkbutton(self, text="Use Gemini when OCR fails (sends pages to Google)", variable=self.gem,
                        command=self.confirm_gemini).grid(row=2, column=2, sticky="w")
        bar = ttk.Frame(self); bar.grid(row=3, column=0, columnspan=3, pady=6)
        self.b1 = ttk.Button(bar, text="1. Extract", command=self.extract); self.b1.pack(side="left", padx=6)
        self.b2 = ttk.Button(bar, text="2. Write to Excel", command=self.write, state="disabled")
        self.b2.pack(side="left", padx=6)
        self.tree = ttk.Treeview(self, columns=("col", "val", "src"), show="headings", height=16)
        for c, w in (("col", 60), ("val", 300), ("src", 200)):
            self.tree.heading(c, text=c.upper()); self.tree.column(c, width=w)
        self.tree.bind("<Double-1>", self.edit_value)
        self.docs = []
        self.tree.grid(row=4, column=0, columnspan=3, padx=8, sticky="nsew")
        self.log = tk.Text(self, height=9, wrap="word"); self.log.grid(row=5, column=0, columnspan=3, padx=8, pady=6)
        self.status = ttk.Label(self, text="Ready"); self.status.grid(row=6, column=0, columnspan=3, sticky="w", padx=8)
        self.timeinfo = ttk.Label(self, text="Start: --      Finish: --      Duration: --", font=("Segoe UI", 10, "bold"))
        self.timeinfo.grid(row=8, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))
        self.bar = ttk.Progressbar(self, maximum=100); self.bar.grid(row=7, column=0, columnspan=3, sticky="ew", padx=8, pady=4)

    def pick_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        self.pdf.set(p or self.pdf.get())

    def pick_xlsx(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        self.xlsx.set(p or self.xlsx.get())

    def confirm_gemini(self):
        if self.gem.get() and not messagebox.askyesno(
                "Gemini", "Only the PDF pages OCR could not read are sent to Google's Gemini API.\n"
                "Do company rules allow this?"):
            self.gem.set(False)

    def extract(self):
        if not self.pdf.get():
            return messagebox.showwarning("PDF", "Choose a PDF first")
        self.b1.config(state="disabled"); self.b2.config(state="disabled")
        self.log.delete("1.0", "end"); self.bar["value"] = 0
        self.t0 = time.time()
        self.t_start = datetime.now()
        self.timeinfo.config(text=f"Start: {self.t_start:%H:%M:%S}      Finish: running...      Duration: --")
        self._log(f"STARTED  {self.t_start:%d/%m/%Y %H:%M:%S}  -  {Path(self.pdf.get()).name}")
        self.status.config(text="Working... please wait")
        threading.Thread(target=self._work, daemon=True).start()

    def _work(self):
        try:
            res, docs, issues = ex.analyze(
                self.pdf.get(), self.gem.get(),
                log=lambda m: self.after(0, self._log, m),
                progress=lambda i, n: self.after(0, self._prog, i, n))
            self.docs = docs
            self.after(0, self._show, res, issues)
        except Exception as e:
            msg = ex.friendly_error(e)
            if "Tesseract OCR not found" not in msg:
                self.after(0, lambda m=msg: ex.record_skip(self.pdf.get(), m, log=self._log))
            self.after(0, self._fail, msg)

    def _log(self, msg):
        self.log.insert("end", msg + "\n"); self.log.see("end")

    def _prog(self, i, n):
        self.bar["value"] = (i / n * 90) if n else 0
        self.status.config(text=f"Reading PDF (OCR)... page {i} of {n}")
        self._log(f"  OCR page {i}/{n} done")

    def _fail(self, msg):
        self.b1.config(state="normal")
        end = datetime.now()
        self.timeinfo.config(text=f"Start: {getattr(self, 't_start', end):%H:%M:%S}      Finish: {end:%H:%M:%S} (FAILED)")
        self.status.config(text="Failed - see message")
        messagebox.showerror("Error", msg)

    def _show(self, res, issues, record=True):
        self.res = res
        if self.xlsx.get() and Path(self.xlsx.get()).is_file() and res["inv_no"]["value"]:
            dup = ex.precheck(self.xlsx.get(), res)          # duplicate invoice / E number check
            if dup:
                issues = issues + [("ERROR", dup)]
        self.tree.delete(*self.tree.get_children())
        cols = {**ex.COLUMNS, **(ex.EXTRA_COLUMNS if self.extras.get() else {})}
        for f, c in cols.items():
            v = res[f]["value"]
            self.tree.insert("", "end", iid=f, values=(c, v.strftime("%d/%b/%Y") if isinstance(v, datetime) else v,
                                                f"{f}  [{', '.join(res[f]['sources'])}]"))
        self._log("  Checking values...")
        self._log("\n".join(f"[{l}] {m}" for l, m in issues) or "All checks passed")
        bad = any(l == "ERROR" for l, _ in issues)
        errs = [m for l, m in issues if l == "ERROR"]
        if errs and record:
            ex.record_skip(self.pdf.get(), ex.error_reason(errs), log=self._log)
        self.bar["value"] = 100
        end, secs = datetime.now(), time.time() - getattr(self, "t0", time.time())
        self.timeinfo.config(text=f"Start: {self.t_start:%H:%M:%S}      Finish: {end:%H:%M:%S}      Duration: {secs:.0f} sec")
        self._log(f"FINISHED {end:%d/%m/%Y %H:%M:%S}  (took {secs:.0f} sec)"
                  + ("  -  fix the [ERROR] (double-click a value), then write" if bad else "  -  review, then click Write to Excel"))
        self.b1.config(state="normal"); self.b2.config(state="disabled" if bad else "normal")
        self.status.config(text="Finished reading - review the values" if not bad else "Finished - [ERROR] must be fixed")

    def edit_value(self, _event):
        """Double-click a row to type/correct its value, then re-check."""
        f = self.tree.focus()
        if not f or not self.res:
            return
        cur = self.res[f]["value"]
        cur = cur.strftime("%d/%m/%Y") if isinstance(cur, datetime) else ("" if cur is None else str(cur))
        text = simpledialog.askstring("Edit value", f"{f}  (dates as DD/MM/YYYY)", initialvalue=cur, parent=self)
        if text is None:
            return
        try:
            ex.set_manual(self.res, f, text)
        except ValueError as e:
            return messagebox.showerror("Invalid", str(e))
        self._show(self.res, ex.validate(self.res, self.docs), record=False)

    def write(self):
        if not self.xlsx.get():
            return messagebox.showwarning("Excel", "Choose the CDC Excel file")
        w0 = datetime.now()
        self._log(f"WRITING to Excel  start {w0:%d/%m/%Y %H:%M:%S} ...")
        self.status.config(text="Writing to Excel...")
        self.update_idletasks()
        try:
            ex.backup(self.xlsx.get())
            row = ex.fill_workbook(self.xlsx.get(), self.res, self.xlsx.get(), self.extras.get())
            self.status.config(text=f"FINISHED - written to row {row}")
            w1 = datetime.now()
            self._log(f"FINISHED writing {w1:%d/%m/%Y %H:%M:%S}: row {row} added (backup saved in /backup)")
            self.timeinfo.config(text=f"Excel write  Start: {w0:%H:%M:%S}      Finish: {w1:%H:%M:%S}      Row: {row}")
            self.b2.config(state="disabled")
            messagebox.showinfo("Finished", f"Written to Excel row {row}.\nBackup saved in the backup folder.")
        except PermissionError:
            messagebox.showerror("Excel", "Close the Excel file first, then try again.")
        except RuntimeError as e:
            ex.record_skip(self.pdf.get(), str(e), log=self._log)
            messagebox.showerror("Not written", str(e))


if __name__ == "__main__":
    App().mainloop()
