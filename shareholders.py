# -*- coding: utf-8 -*-
"""برنامه محاسبه سهامداری - ورودی: نام، تعداد سهام، تعداد وثیقه. بقیه محاسبه می‌شود."""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json, os, csv, re, tempfile, webbrowser, html, datetime

DATA = os.path.join(os.path.expanduser("~"), "shareholders_data.json")
FA = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
MODES = ["start", "end", "both"]
MODE_TXT = ["وثیقه در ابتدای بازه هر نفر", "وثیقه در انتهای بازه هر نفر", "هر دو (گزارش چاپی دو جدول)"]
TITLES = {"start": "حالت اول: سهام وثیقه در ابتدای بازه", "end": "حالت دوم: سهام وثیقه در انتهای بازه"}
COLS = [f"c{i}" for i in range(13)]


def num(s):
    s = str(s).translate(FA).replace(",", "").replace("٬", "").strip()
    return int(s) if s else 0


def fmt(n):
    return f"{int(n):,}"


def pct_s(p):
    return f"{p:.5f}".rstrip("0").rstrip(".").replace(".", "/")


def compute(people, start0, mode):
    tot = sum(p["total"] for p in people)
    rows, cur = [], start0
    for p in people:
        t, pl = p["total"], p["pledged"]
        fr = t - pl
        s, e = cur, cur + t - 1
        cur = e + 1
        if mode == "start":
            pr = (s, s + pl - 1) if pl else None
            fv = (s + pl, e) if fr else None
        else:
            fv = (s, s + fr - 1) if fr else None
            pr = (s + fr, e) if pl else None
        rows.append(dict(name=p["name"], total=t, pl=pl, fr=fr, pct=t / tot * 100 if tot else 0,
                         s=s, e=e, P=pr, F=fv))
    return rows, tot


def build(people, start0, mode):
    rows, tot = compute(people, start0, mode)
    order = "PF" if mode == "start" else "FP"
    H = {"P": ["ابتدای وثیقه", "انتهای وثیقه", "تعداد وثیقه"], "F": ["ابتدای آزاد", "انتهای آزاد", "تعداد آزاد"]}
    heads = ["سهامدار", "تعداد سهام", "تعداد سهام وثیقه", "تعداد سهام آزاد", "%", "ابتدای سهام", "انتهای سهام"]
    for k in order:
        heads += H[k]
    body = []
    for r in rows:
        line = [r["name"], fmt(r["total"]), fmt(r["pl"]), fmt(r["fr"]), pct_s(r["pct"]), fmt(r["s"]), fmt(r["e"])]
        for k in order:
            rg = r[k]
            line += [fmt(rg[0]) if rg else "—", fmt(rg[1]) if rg else "—", fmt(r["pl"] if k == "P" else r["fr"])]
        body.append(line)
    spl, sfr = sum(r["pl"] for r in rows), sum(r["fr"] for r in rows)
    total = ["جمع", fmt(tot), fmt(spl), fmt(sfr), "100" if tot else "0", "—", "—"]
    for k in order:
        total += ["—", "—", fmt(spl if k == "P" else sfr)]
    return heads, body, total


def write_xlsx(path, people, start0, modes):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook()
    wb.remove(wb.active)
    side = Side(style="thin")
    box = Border(left=side, right=side, top=side, bottom=side)
    gray = PatternFill("solid", fgColor="DDDDDD")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    names = {"start": "وثیقه در ابتدا", "end": "وثیقه در انتها"}
    for m in modes:
        ws = wb.create_sheet(names[m])
        ws.sheet_view.rightToLeft = True
        heads, rows, total = build(people, start0, m)
        for ri, line in enumerate([heads] + rows + [total], 1):
            for ci, x in enumerate(line, 1):
                if ci > 1 and re.fullmatch(r"[\d,]+", x):
                    v, nf = int(x.replace(",", "")), "#,##0"
                elif ci > 1 and re.fullmatch(r"\d+/\d+", x):
                    v, nf = float(x.replace("/", ".")), "0.00000"
                else:
                    v, nf = x, "General"
                c = ws.cell(row=ri, column=ci, value=v)
                c.number_format, c.alignment, c.border = nf, center, box
                if ri == 1 or ri == len(rows) + 2:
                    c.font = Font(bold=True)
                    c.fill = gray
        for ci in range(1, 14):
            ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = 22 if ci == 1 else 15
        ws.freeze_panes = "B2"
    wb.save(path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("برنامه محاسبه سهامداری")
        self.geometry("1350x640")
        self.db = {"company": "", "start0": 1, "mode": "start", "people": []}
        if os.path.exists(DATA):
            try:
                self.db.update(json.load(open(DATA, encoding="utf-8")))
            except Exception:
                pass
        self.edit_idx = None
        self.build_ui()
        self.refresh()

    def build_ui(self):
        top = ttk.LabelFrame(self, text="تنظیمات")
        top.pack(fill="x", padx=8, pady=4)
        self.company = tk.StringVar(value=self.db["company"])
        self.start0 = tk.StringVar(value=str(self.db["start0"]))
        ttk.Label(top, text="نام شرکت:").pack(side="right", padx=4)
        ttk.Entry(top, textvariable=self.company, width=30, justify="right").pack(side="right")
        ttk.Label(top, text="سریال شروع:").pack(side="right", padx=8)
        ttk.Entry(top, textvariable=self.start0, width=10, justify="right").pack(side="right")
        ttk.Label(top, text="جایگاه وثیقه:").pack(side="right", padx=8)
        self.modebox = ttk.Combobox(top, values=MODE_TXT, state="readonly", width=30, justify="right")
        self.modebox.current(MODES.index(self.db["mode"]))
        self.modebox.pack(side="right")
        ttk.Button(top, text="اعمال", command=self.refresh).pack(side="right", padx=8)

        form = ttk.LabelFrame(self, text="ورود اطلاعات اولیه (بقیه خودکار محاسبه می‌شود)")
        form.pack(fill="x", padx=8, pady=4)
        self.vars = {k: tk.StringVar() for k in ("name", "total", "pledged")}
        for k, label in (("pledged", "تعداد سهام وثیقه"), ("total", "تعداد سهام"), ("name", "نام سهامدار")):
            ttk.Label(form, text=label).pack(side="right", padx=6)
            ttk.Entry(form, textvariable=self.vars[k], justify="right", width=22).pack(side="right")
        bt = ttk.Frame(self)
        bt.pack(fill="x", padx=8)
        for t, cmd in [("ذخیره", self.save), ("جدید", self.clear), ("حذف", self.delete),
                       ("▲ بالا", lambda: self.move(-1)), ("▼ پایین", lambda: self.move(1)),
                       ("چاپ", self.print_report), ("خروجی Excel", self.export_xlsx),
                       ("خروجی CSV", self.export)]:
            ttk.Button(bt, text=t, command=cmd).pack(side="right", padx=3, pady=4)

        self.tree = ttk.Treeview(self, columns=COLS, show="headings", height=12,
                                 displaycolumns=COLS[::-1])
        for c in COLS:
            self.tree.column(c, width=100, anchor="center")
        self.tree.column("c0", width=140)
        self.tree.pack(fill="both", expand=True, padx=8, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        ttk.Label(self, text="ترتیب ردیف‌ها تعیین‌کننده سریال سهام است؛ با ▲ ▼ جابه‌جا کنید.").pack(pady=2)

    def clear(self):
        self.edit_idx = None
        for v in self.vars.values():
            v.set("")
        self.tree.selection_remove(self.tree.selection())

    def on_select(self, _=None):
        sel = self.tree.selection()
        if not sel or sel[0] == "tot":
            return
        self.edit_idx = int(sel[0])
        p = self.db["people"][self.edit_idx]
        for k in self.vars:
            self.vars[k].set(str(p[k]))

    def save(self):
        try:
            p = {"name": self.vars["name"].get().strip(), "total": num(self.vars["total"].get()),
                 "pledged": num(self.vars["pledged"].get())}
        except ValueError:
            return messagebox.showerror("خطا", "مقادیر عددی نامعتبر است.")
        if not p["name"] or p["total"] <= 0:
            return messagebox.showerror("خطا", "نام و تعداد سهام را وارد کنید.")
        if p["pledged"] > p["total"]:
            return messagebox.showerror("خطا", "سهام وثیقه نمی‌تواند از کل سهام بیشتر باشد.")
        if self.edit_idx is None:
            self.db["people"].append(p)
        else:
            self.db["people"][self.edit_idx] = p
        self.clear()
        self.refresh()

    def delete(self):
        if self.edit_idx is None:
            return messagebox.showinfo("توجه", "ابتدا یک ردیف را انتخاب کنید.")
        if messagebox.askyesno("حذف", "ردیف انتخاب‌شده حذف شود؟"):
            del self.db["people"][self.edit_idx]
            self.clear()
            self.refresh()

    def move(self, d):
        i, n = self.edit_idx, len(self.db["people"])
        if i is None or not 0 <= i + d < n:
            return
        L = self.db["people"]
        L[i], L[i + d] = L[i + d], L[i]
        self.refresh()
        self.tree.selection_set(str(i + d))

    def cur_mode(self):
        m = self.db["mode"]
        return "start" if m == "both" else m

    def refresh(self, *_):
        try:
            self.db["start0"] = max(1, num(self.start0.get()))
        except ValueError:
            self.db["start0"] = 1
        self.db["company"] = self.company.get()
        self.db["mode"] = MODES[self.modebox.current()]
        heads, rows, total = build(self.db["people"], self.db["start0"], self.cur_mode())
        for c, h in zip(COLS, heads):
            self.tree.heading(c, text=h)
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(rows):
            self.tree.insert("", "end", iid=str(i), values=r)
        self.tree.insert("", "end", iid="tot", values=total, tags=("t",))
        self.tree.tag_configure("t", background="#e6e6e6")
        json.dump(self.db, open(DATA, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    def export(self):
        f = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not f:
            return
        heads, rows, total = build(self.db["people"], self.db["start0"], self.cur_mode())
        with open(f, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(heads)
            w.writerows(rows + [total])

    def export_xlsx(self):
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            return messagebox.showerror("خطا", "کتابخانه openpyxl نصب نیست.\nدستور: pip install openpyxl")
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not f:
            return
        modes = ["start", "end"] if self.db["mode"] == "both" else [self.db["mode"]]
        try:
            write_xlsx(f, self.db["people"], self.db["start0"], modes)
        except PermissionError:
            return messagebox.showerror("خطا", "فایل باز است یا امکان نوشتن ندارد. آن را ببندید و دوباره تلاش کنید.")
        if messagebox.askyesno("انجام شد", "فایل ذخیره شد. باز شود؟"):
            os.startfile(f) if hasattr(os, "startfile") else webbrowser.open("file:///" + f)

    def print_report(self):
        e = html.escape
        modes = ["start", "end"] if self.db["mode"] == "both" else [self.db["mode"]]
        parts = []
        for m in modes:
            heads, rows, total = build(self.db["people"], self.db["start0"], m)
            tr = lambda r, tag="td", cls="": f"<tr{cls}>" + "".join(f"<{tag}>{e(x)}</{tag}>" for x in r) + "</tr>"
            parts.append(f"<div class=pg><h3>{TITLES[m]}</h3><table>{tr(heads, 'th')}"
                         + "".join(tr(r) for r in rows) + tr(total, "td", " class=t") + "</table></div>")
        doc = f"""<!DOCTYPE html><html dir="rtl" lang="fa"><meta charset="utf-8">
<style>body{{font-family:Tahoma,sans-serif;margin:15px}}h2,h3,p{{text-align:center}}
table{{width:100%;border-collapse:collapse;font-size:11px}}
td,th{{border:1px solid #444;padding:5px;text-align:center}}th{{background:#ddd}}
.t{{font-weight:bold;background:#eee}}.pg{{page-break-after:always}}.pg:last-child{{page-break-after:auto}}
@page{{size:A4 landscape;margin:10mm}}</style>
<h2>{e(self.company.get() or 'گزارش سهامداری')}</h2>
<p>تاریخ چاپ: {datetime.date.today()}</p>{''.join(parts)}
<script>window.onload=()=>window.print()</script></html>"""
        f = os.path.join(tempfile.gettempdir(), "shareholders_report.html")
        open(f, "w", encoding="utf-8").write(doc)
        webbrowser.open("file:///" + f.replace("\\", "/"))


if __name__ == "__main__":
    App().mainloop()
