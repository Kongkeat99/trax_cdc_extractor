"""Trax CDC Export Report - PDF to Excel
Drag & drop MANY shipment PDFs (or a folder) + the CDC Excel file -> extract all -> review -> write all.
Drag & drop needs:  pip install tkinterdnd2   (without it the Browse buttons still work)."""
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import extractor as ex

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _Base, HAS_DND = TkinterDnD.Tk, True
except Exception:                                   # package missing -> window still works, no drag & drop
    _Base, HAS_DND = tk.Tk, False


class Item:
    """One PDF in the list."""
    def __init__(self, path):
        self.path = Path(path)
        self.res, self.docs, self.issues = None, [], []
        self.status, self.note, self.written = "QUEUED", "", False


def status_of(issues):
    if any(l == "ERROR" for l, _ in issues):
        return "ERROR"
    if any(l == "WARN" and not m.startswith("document not found") for l, m in issues):
        return "CHECK"
    return "OK"


class App(_Base):
    def __init__(self):
        super().__init__()
        self.title("Trax CDC Export Report - PDF to Excel")
        self.geometry("980x800")
        self.minsize(820, 640)
        self.xlsx, self.extras, self.gem = tk.StringVar(), tk.BooleanVar(), tk.BooleanVar()
        self.items, self.byiid = [], {}
        self.cur, self.rec_idx, self.busy = None, 0, False
        self.cur_name, self.t_start, self.t0 = "", datetime.now(), time.time()
        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=3)
        self.rowconfigure(6, weight=1)

        # ---- row 0: PDF list (drop zone) ----
        ttk.Label(self, text="Shipment PDFs").grid(row=0, column=0, padx=8, pady=6, sticky="nw")
        ff = ttk.Frame(self)
        ff.grid(row=0, column=1, sticky="nsew", pady=6)
        ff.columnconfigure(0, weight=1)
        self.files = ttk.Treeview(ff, columns=("file", "status", "note"), show="headings", height=7,
                                  selectmode="extended")
        for c, t, w in (("file", "PDF FILE", 300), ("status", "STATUS", 130), ("note", "NOTE", 380)):
            self.files.heading(c, text=t)
            self.files.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(ff, orient="vertical", command=self.files.yview)
        self.files.configure(yscrollcommand=sb.set)
        self.files.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        for tag, color in (("ERROR", "#b00020"), ("FAILED", "#b00020"), ("NOT", "#b00020"), ("CHECK", "#b36b00"),
                           ("OK", "#1a7f37"), ("WRITTEN", "#0b57d0")):
            self.files.tag_configure(tag, foreground=color)
        self.files.bind("<<TreeviewSelect>>", self.on_select)
        self.files.bind("<Delete>", lambda e: self.remove_selected())
        self.hint = ttk.Label(ff, foreground="gray", font=("Segoe UI", 11), justify="center",
                              text=("Drag & drop PDF files or a folder here\n(or click Add PDFs)" if HAS_DND else
                                    "Click Add PDFs\n(for drag & drop run:  pip install tkinterdnd2)"))
        self.hint.place(relx=0.5, rely=0.5, anchor="center")
        side = ttk.Frame(self)
        side.grid(row=0, column=2, padx=6, sticky="n", pady=6)
        ttk.Button(side, text="Add PDFs", width=14, command=self.pick_pdfs).pack(pady=2)
        ttk.Button(side, text="Remove selected", width=14, command=self.remove_selected).pack(pady=2)
        ttk.Button(side, text="Clear list", width=14, command=self.clear_list).pack(pady=2)

        # ---- row 1: Excel ----
        ttk.Label(self, text="CDC Excel").grid(row=1, column=0, padx=8, pady=6, sticky="w")
        self.xl_entry = ttk.Entry(self, textvariable=self.xlsx)
        self.xl_entry.grid(row=1, column=1, sticky="ew")
        ttk.Button(self, text="Browse", command=self.pick_xlsx).grid(row=1, column=2, padx=6)

        # ---- row 2: options ----
        opt = ttk.Frame(self)
        opt.grid(row=2, column=1, columnspan=2, sticky="w")
        ttk.Checkbutton(opt, text="Also fill unmarked columns E (flight date) and J (procedure)",
                        variable=self.extras, command=self._render).pack(side="left")
        ttk.Checkbutton(opt, text="Use Gemini when OCR fails (sends pages to Google)", variable=self.gem,
                        command=self.confirm_gemini).pack(side="left", padx=20)

        # ---- row 3: buttons ----
        bar = ttk.Frame(self)
        bar.grid(row=3, column=0, columnspan=3, pady=6)
        self.b1 = ttk.Button(bar, text="1. Extract all", command=self.extract)
        self.b1.pack(side="left", padx=6)
        self.b2 = ttk.Button(bar, text="2. Write to Excel", command=self.write, state="disabled")
        self.b2.pack(side="left", padx=6)
        self.b3 = ttk.Button(bar, text="Cut PDF pages", command=self.trim)
        self.b3.pack(side="left", padx=6)

        # ---- row 4/5: invoice selector + value table ----
        sel = ttk.Frame(self)
        sel.grid(row=4, column=0, columnspan=3, sticky="w", padx=8)
        ttk.Label(sel, text="Invoice:").pack(side="left")
        self.rec_box = ttk.Combobox(sel, state="readonly", width=50)
        self.rec_box.pack(side="left", padx=6)
        self.rec_box.bind("<<ComboboxSelected>>", self.on_rec)
        ttk.Label(sel, text="( ! = a required value is missing; double-click a value to fix it )",
                  foreground="gray").pack(side="left")
        self.tree = ttk.Treeview(self, columns=("col", "val", "src"), show="headings", height=14)
        for c, w in (("col", 60), ("val", 420), ("src", 260)):
            self.tree.heading(c, text=c.upper())
            self.tree.column(c, width=w)
        self.tree.bind("<Double-1>", self.edit_value)
        self.tree.grid(row=5, column=0, columnspan=3, padx=8, sticky="nsew")

        # ---- row 6..9: log, status, progress, time ----
        self.log = tk.Text(self, height=9, wrap="word")
        self.log.grid(row=6, column=0, columnspan=3, padx=8, pady=6, sticky="nsew")
        self.status = ttk.Label(self, text="Ready")
        self.status.grid(row=7, column=0, columnspan=3, sticky="w", padx=8)
        self.pbar = ttk.Progressbar(self, maximum=100)
        self.pbar.grid(row=8, column=0, columnspan=3, sticky="ew", padx=8, pady=4)
        self.timeinfo = ttk.Label(self, text="Start: --      Finish: --      Duration: --",
                                  font=("Segoe UI", 10, "bold"))
        self.timeinfo.grid(row=9, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))

        if HAS_DND:
            for w in (self, self.files, self.hint, self.xl_entry, self.tree, self.log):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._drop)
        else:
            self.status.config(text="Ready  (drag & drop is off: pip install tkinterdnd2)")

    # ------------------------------------------------------------------ list handling
    def _drop(self, event):
        """Any PDF / folder -> PDF list; first .xlsx -> Excel box (whatever widget it was dropped on)."""
        pdfs, xl = [], None
        for p in map(Path, self.tk.splitlist(event.data)):
            if p.is_dir():
                pdfs += sorted(q for q in p.glob("*.pdf") if not q.stem.endswith("_trimmed"))
            elif p.suffix.lower() == ".pdf":
                pdfs.append(p)
            elif p.suffix.lower() in (".xlsx", ".xlsm") and xl is None:
                xl = p
        if xl:
            self.xlsx.set(str(xl))
            self._log(f"Excel: {xl}")
        if pdfs:
            self.add_pdfs(pdfs)
        if not pdfs and not xl:
            self.status.config(text="Drop PDF files, a folder of PDFs, or an .xlsx file")
        return event.action

    def pick_pdfs(self):
        ps = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if ps:
            self.add_pdfs(map(Path, ps))

    def pick_xlsx(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        self.xlsx.set(p or self.xlsx.get())

    def add_pdfs(self, paths):
        if self.busy:
            return self.status.config(text="Busy - add files when the current job has finished")
        have = {it.path.resolve() for it in self.items}
        added = 0
        for p in paths:
            if p.resolve() in have:
                continue
            it = Item(p)
            self.items.append(it)
            self.byiid[str(id(it))] = it
            self.files.insert("", "end", iid=str(id(it)), values=(p.name, it.status, ""))
            self._row(it)
            have.add(p.resolve())
            added += 1
        self._hint()
        self.status.config(text=f"{added} PDF added - {len(self.items)} in the list. Press '1. Extract all'")

    def remove_selected(self):
        if self.busy:
            return
        for iid in self.files.selection():
            it = self.byiid.pop(iid, None)
            if it:
                self.items.remove(it)
                if self.cur is it:
                    self.cur = None
            self.files.delete(iid)
        self._render()
        self._hint()
        self._update_b2()

    def clear_list(self):
        if self.busy:
            return
        self.files.delete(*self.files.get_children())
        self.items.clear(); self.byiid.clear(); self.cur = None
        self._render(); self._hint(); self._update_b2()

    def _hint(self):
        if self.items:
            self.hint.place_forget()
        else:
            self.hint.place(relx=0.5, rely=0.5, anchor="center")

    def _row(self, it):
        tag = it.status.split()[0]
        self.files.item(str(id(it)), values=(it.path.name, it.status, it.note), tags=(tag,))

    def _set_busy(self, b):
        self.busy = b
        st = "disabled" if b else "normal"
        self.b1.config(state=st); self.b3.config(state=st)
        self._update_b2()

    def _update_b2(self):
        ok = (not self.busy) and any(it.status in ("OK", "CHECK") and not it.written for it in self.items)
        self.b2.config(state="normal" if ok else "disabled")

    def confirm_gemini(self):
        if self.gem.get() and not messagebox.askyesno(
                "Gemini", "Only the PDF pages OCR could not read are sent to Google's Gemini API.\n"
                "Do company rules allow this?"):
            self.gem.set(False)

    def _log(self, msg):
        self.log.insert("end", msg + "\n"); self.log.see("end")

    # ------------------------------------------------------------------ extract (all files)
    def extract(self):
        todo = [it for it in self.items if not it.written]
        if not todo:
            return messagebox.showwarning("PDF", "Add PDF files first (drag & drop, or click Add PDFs)")
        self._set_busy(True)
        self.log.delete("1.0", "end"); self.pbar["value"] = 0
        for it in todo:
            it.res, it.issues, it.note, it.status = None, [], "", "QUEUED"
            self._row(it)
        self.t0, self.t_start = time.time(), datetime.now()
        self.timeinfo.config(text=f"Start: {self.t_start:%H:%M:%S}      Finish: running...      Duration: --")
        self._log(f"STARTED  {self.t_start:%d/%m/%Y %H:%M:%S}  -  {len(todo)} PDF file(s)")
        threading.Thread(target=self._work, args=(todo, self.gem.get(), self.xlsx.get()), daemon=True).start()

    def _work(self, todo, gem, xlsx):
        n = len(todo)
        for k, it in enumerate(todo, 1):
            self.after(0, self._start_item, it, k, n)
            try:
                res, docs, issues = ex.analyze(
                    str(it.path), gem,
                    log=lambda m: self.after(0, self._log, m),
                    progress=lambda i, p, k=k: self.after(0, self._prog, k, n, i, p))
                if xlsx and Path(xlsx).is_file() and res["inv_no"]["value"]:
                    try:
                        dup = ex.precheck(xlsx, res)            # duplicate invoice / E number already in Excel
                    except Exception:
                        dup = None
                    if dup:
                        issues = issues + [("ERROR", dup)]
            except Exception as e:
                msg = ex.friendly_error(e)
                if "Tesseract OCR not found" in msg:
                    return self.after(0, self._stop, msg)
                self.after(0, self._item_failed, it, msg)
                continue
            self.after(0, self._item_done, it, res, docs, issues)
        self.after(0, self._all_done)

    def _start_item(self, it, k, n):
        self.cur_name = it.path.name
        it.status = "READING"
        self._row(it)
        self._log(f"\n=== [{k}/{n}] {it.path.name}")

    def _prog(self, k, n, i, p):
        self.pbar["value"] = ((k - 1) + (i / p * 0.9 if p else 0)) / n * 100
        self.status.config(text=f"[{k}/{n}] {self.cur_name} - reading (OCR) page {i} of {p}")

    def _item_done(self, it, res, docs, issues):
        it.res, it.docs, it.issues = res, docs, issues
        self._set_status(it)
        for l, m in issues:
            if l != "INFO":
                self._log(f"[{l}] {m}")
        if it.status == "ERROR":
            errs = [m for l, m in issues if l == "ERROR"]
            ex.record_skip(str(it.path), ex.error_reason(errs), log=self._log, res=res, issues=issues)
        self._row(it)

    def _item_failed(self, it, msg):
        it.status, it.note = "FAILED", msg
        ex.record_skip(str(it.path), msg, log=self._log)
        self._row(it)

    def _set_status(self, it):
        it.status = status_of(it.issues)
        errs = [m for l, m in it.issues if l == "ERROR"]
        warns = [m for l, m in it.issues if l == "WARN" and not m.startswith("document not found")]
        it.note = ex.error_reason(errs) if errs else ("; ".join(warns[:2]) if warns else "")
        if it.res and len(ex.group_of(it.res)) > 1:
            it.note = f"{len(ex.group_of(it.res))} invoices. " + it.note

    def _stop(self, msg):
        self._set_busy(False)
        self.status.config(text="Stopped - OCR engine missing")
        messagebox.showerror("Error", msg)

    def _all_done(self):
        end, secs = datetime.now(), time.time() - self.t0
        cnt = {s: sum(1 for it in self.items if it.status == s) for s in ("OK", "CHECK", "ERROR", "FAILED")}
        self.pbar["value"] = 100
        self.timeinfo.config(text=f"Start: {self.t_start:%H:%M:%S}      Finish: {end:%H:%M:%S}      Duration: {secs:.0f} sec")
        summary = (f"{cnt['OK']} OK, {cnt['CHECK']} to check, {cnt['ERROR']} with [ERROR], {cnt['FAILED']} failed")
        self._log(f"\nFINISHED {end:%d/%m/%Y %H:%M:%S}  (took {secs:.0f} sec)  -  {summary}")
        self.status.config(text=("Finished - " + summary + (" - fix the ERROR files (select one, double-click a value), "
                                 "then Write" if cnt["ERROR"] else " - review, then click Write to Excel")))
        self._set_busy(False)
        first = next((it for it in self.items if it.status in ("ERROR", "FAILED")), None) or \
            (self.items[0] if self.items else None)
        if first:
            self.files.selection_set(str(id(first)))
            self.files.see(str(id(first)))
            self.cur = None                       # force a fresh display of the re-read result
            self.on_select()

    # ------------------------------------------------------------------ show / edit one file
    def on_select(self, _event=None):
        sel = self.files.selection()
        it = self.byiid.get(sel[0]) if sel else None
        if it is self.cur and it is not None:
            return
        self.cur, self.rec_idx = it, 0
        if it:
            self._log(f"\n--- {it.path.name}: {it.status}" + (f" - {it.note}" if it.note else ""))
            for l, m in it.issues:
                if l != "INFO":
                    self._log(f"[{l}] {m}")
        self._render()

    def on_rec(self, _event=None):
        self.rec_idx = self.rec_box.current()
        self._render(keep_box=True)

    def _render(self, keep_box=False):
        self.tree.delete(*self.tree.get_children())
        it = self.cur
        if not it or not it.res:
            self.rec_box.config(values=()); self.rec_box.set("")
            return
        recs = ex.group_of(it.res)
        self.rec_idx = min(self.rec_idx, len(recs) - 1)
        if not keep_box:
            names = []
            for n, r in enumerate(recs, 1):
                bad = any(r[f]["value"] in (None, "") for f in ex.COLUMNS)
                names.append(f"{n}. {r['inv_no']['value'] or '?'}  ({r['customs_no']['value'] or '-'})" + ("  !" if bad else ""))
            self.rec_box.config(values=names)
        self.rec_box.current(self.rec_idx)
        rec = recs[self.rec_idx]
        cols = {**ex.COLUMNS, **(ex.EXTRA_COLUMNS if self.extras.get() else {})}
        for f, c in cols.items():
            v = rec[f]["value"]
            self.tree.insert("", "end", iid=f, values=(c, v.strftime("%d/%b/%Y") if isinstance(v, datetime) else v,
                                                       f"{f}  [{', '.join(rec[f]['sources'])}]"))

    def edit_value(self, _event):
        """Double-click a value of the selected invoice to type/correct it, then re-check the file."""
        f, it = self.tree.focus(), self.cur
        if not f or not it or not it.res:
            return
        rec = ex.group_of(it.res)[self.rec_idx]
        cur = rec[f]["value"]
        cur = cur.strftime("%d/%m/%Y") if isinstance(cur, datetime) else ("" if cur is None else str(cur))
        text = simpledialog.askstring("Edit value", f"{f}  (dates as DD/MM/YYYY)", initialvalue=cur, parent=self)
        if text is None:
            return
        try:
            ex.set_manual(rec, f, text)
        except ValueError as e:
            return messagebox.showerror("Invalid", str(e))
        except Exception as e:
            return messagebox.showerror("Edit", f"{type(e).__name__}: {e}")
        issues = ex.validate(it.res, it.docs)
        xl = self.xlsx.get()
        if xl and Path(xl).is_file() and it.res["inv_no"]["value"]:
            try:
                dup = ex.precheck(xl, it.res)
            except Exception:
                dup = None
            if dup:
                issues = issues + [("ERROR", dup)]
        it.issues = issues
        self._set_status(it)
        self._log(f"\n--- {it.path.name} edited -> {it.status}")
        for l, m in issues:
            if l != "INFO":
                self._log(f"[{l}] {m}")
        self._row(it)
        self._render()
        self._update_b2()

    # ------------------------------------------------------------------ write all
    def write(self):
        xl = self.xlsx.get()
        if not xl or not Path(xl).is_file():
            return messagebox.showwarning("Excel", "Choose the CDC Excel file (drop it on the window or click Browse)")
        todo = [it for it in self.items if it.status in ("OK", "CHECK") and not it.written]
        if not todo:
            return
        w0 = datetime.now()
        self._log(f"\nWRITING {len(todo)} file(s) to Excel  start {w0:%d/%m/%Y %H:%M:%S} ...")
        wrote = bad = 0
        try:
            ex.backup(xl)
            for k, it in enumerate(todo, 1):
                self.status.config(text=f"Writing {k}/{len(todo)}: {it.path.name}")
                self.pbar["value"] = k / len(todo) * 100
                self.update_idletasks()
                try:
                    row = ex.fill_workbook(xl, it.res, xl, self.extras.get())
                    rows = it.res.get("_rows") or [row]
                    it.status, it.written, it.note = f"WRITTEN row {rows[0]}" + (f"-{rows[-1]}" if len(rows) > 1 else ""), True, ""
                    wrote += 1
                    self._log(f"  {it.path.name}: {it.status}")
                except PermissionError:
                    raise
                except Exception as e:
                    msg = ex.friendly_error(e)
                    it.status, it.note = "NOT WRITTEN", msg
                    ex.record_skip(str(it.path), msg, log=self._log, res=it.res, issues=it.issues)
                    bad += 1
                self._row(it)
        except PermissionError:
            messagebox.showerror("Excel", "Close the Excel file first, then try again.")
        w1 = datetime.now()
        self._log(f"FINISHED writing {w1:%d/%m/%Y %H:%M:%S}: {wrote} file(s) written, {bad} not written (backup saved in /backup)")
        self.timeinfo.config(text=f"Excel write  Start: {w0:%H:%M:%S}      Finish: {w1:%H:%M:%S}      Files: {wrote}")
        self.status.config(text=f"FINISHED - {wrote} written, {bad} not written")
        self._update_b2()
        skipped = sum(1 for it in self.items if it.status in ("ERROR", "FAILED"))
        messagebox.showinfo("Finished", f"{wrote} PDF written to Excel, {bad} not written (duplicate / locked).\n"
                            + (f"{skipped} file(s) still have [ERROR] - fix them and press Write again.\n" if skipped else "")
                            + "Backup saved in the backup folder.")

    # ------------------------------------------------------------------ cut PDF pages
    def trim(self):
        """Cut selected PDFs (or all if none selected): finds every E number with its SAD page + COMMERCIAL
        INVOICE pages and shows the result here (log + NOTE column). No file is written."""
        sel = [self.byiid[i] for i in self.files.selection() if i in self.byiid]
        todo = [it for it in (sel or self.items) if not it.written]
        if not todo:
            return messagebox.showwarning("PDF", "Add PDF files first (drag & drop, or click Add PDFs)")
        self._set_busy(True)
        self.pbar["value"] = 0
        self.status.config(text="Cutting PDF pages (OCR)... please wait")
        threading.Thread(target=self._trim_work, args=(todo,), daemon=True).start()

    def _trim_work(self, todo):
        out = []
        for k, it in enumerate(todo, 1):
            self.after(0, self._log, f"\nCUT PDF [{k}/{len(todo)}] {it.path.name}")
            try:
                f, kept = ex.trim_pdf(str(it.path), log=lambda m: self.after(0, self._log, m),
                                      progress=lambda i, p, k=k: self.after(0, self._prog, k, len(todo), i, p))
                out.append((it, Path(f), kept, None, list(getattr(ex.trim_pdf, "last_results", []))))
            except Exception as e:
                out.append((it, None, None, ex.friendly_error(e), []))
        self.after(0, self._trim_done, out)

    def _trim_done(self, out):
        ok = 0
        for it, f, kept, err, results in out:
            if err:
                it.status, it.note = "FAILED", f"cut failed: {err}"
                ex.record_skip(str(it.path), it.note, log=self._log)
            else:
                ok += 1
                it.res, it.issues = None, []
                if f.resolve() != it.path.resolve():
                    it.path = f                                # list now holds the trimmed PDF
                it.status = "QUEUED"
                it.note = ("cut: " + ", ".join(f"{r['e']} {r['status']}" for r in results)) if results else f"cut: kept pages {kept}"
            self._row(it)
        self.pbar["value"] = 100
        self.status.config(text=f"Cut finished: {ok} ok, {len(out) - ok} failed - press '1. Extract all'")
        self._set_busy(False)
        self._render()


if __name__ == "__main__":
    App().mainloop()
