# -*- coding: utf-8 -*-
"""برنامه محاسبه سهامداری - Tkinter (ویندوز)"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json, os, csv, tempfile, webbrowser, html, datetime

DATA = os.path.join(os.path.expanduser("~"), "shareholders_data.json")
FA = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def num(s):
    s = str(s).translate(FA).replace(",", "").replace("٬", "").strip()
    return int(s) if s else 0


def fmt(n):
    return f"{int(n):,}"


class App(tk.Tk):
    FIELDS = [("name", "نام سهامدار"), ("nid", "کد ملی"), ("total", "تعداد کل سهام"),
              ("pledged", "سهام وثیقه"), ("start", "شروع سهم (سریال)"),
              ("end", "پایان سهم (سریال)"), ("note", "توضیحات")]

    def __init__(self):
        super().__init__()
        self.title("برنامه محاسبه سهامداری")
        self.geometry("1150x650")
        self.db = {"company": "", "nominal": 1000, "people": []}
        if os.path.exists(DATA):
            try:
                self.db.update(json.load(open(DATA, encoding="utf-8")))
            except Exception:
                pass
        self.edit_idx = None
        self.build()
        self.refresh()

    # ---------- UI ----------
    def build(self):
        top = ttk.LabelFrame(self, text="مشخصات شرکت")
        top.pack(fill="x", padx=8, pady=4)
        self.company = tk.StringVar(value=self.db["company"])
        self.nominal = tk.StringVar(value=str(self.db["nominal"]))
        ttk.Label(top, text="نام شرکت:").pack(side="right", padx=4)
        ttk.Entry(top, textvariable=self.company, width=35, justify="right").pack(side="right")
        ttk.Label(top, text="ارزش اسمی هر سهم (ریال):").pack(side="right", padx=8)
        ttk.Entry(top, textvariable=self.nominal, width=12, justify="right").pack(side="right")
        ttk.Button(top, text="اعمال", command=self.refresh).pack(side="right", padx=8)

        form = ttk.LabelFrame(self, text="ثبت / ویرایش سهامدار")
        form.pack(fill="x", padx=8, pady=4)
        self.vars = {}
        for i, (k, label) in enumerate(self.FIELDS):
            v = tk.StringVar()
            self.vars[k] = v
            r, c = divmod(i, 4)
            ttk.Label(form, text=label).grid(row=r * 2, column=3 - c, sticky="e", padx=6)
            ttk.Entry(form, textvariable=v, justify="right", width=24).grid(
                row=r * 2 + 1, column=3 - c, padx=6, pady=2)
        bt = ttk.Frame(form)
        bt.grid(row=4, column=0, columnspan=4, pady=6)
        for t, cmd in [("ذخیره", self.save), ("جدید / پاک کردن فرم", self.clear),
                       ("حذف انتخاب‌شده", self.delete), ("چاپ", self.print_report),
                       ("خروجی Excel (CSV)", self.export)]:
            ttk.Button(bt, text=t, command=cmd).pack(side="right", padx=4)

        cols = ("row", "name", "nid", "total", "pledged", "main", "start", "end", "pct", "value", "note")
        heads = ["ردیف", "نام سهامدار", "کد ملی", "کل سهام", "سهام وثیقه", "سهام اصلی (آزاد)",
                 "شروع سهم", "پایان سهم", "درصد", "ارزش اسمی (ریال)", "توضیحات"]
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        for c, h in zip(cols, heads):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=60 if c == "row" else 105, anchor="center")
        self.tree.column("name", width=160)
        self.tree.pack(fill="both", expand=True, padx=8, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.summary = ttk.Label(self, font=("Tahoma", 10, "bold"))
        self.summary.pack(pady=4)

    # ---------- data ----------
    def nominal_val(self):
        try:
            return num(self.nominal.get())
        except ValueError:
            return 0

    def clear(self):
        self.edit_idx = None
        for v in self.vars.values():
            v.set("")
        self.tree.selection_remove(self.tree.selection())

    def on_select(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        self.edit_idx = int(sel[0])
        p = self.db["people"][self.edit_idx]
        for k, v in self.vars.items():
            v.set(str(p.get(k, "")))

    def save(self):
        try:
            p = {k: self.vars[k].get().strip() for k in self.vars}
            for k in ("total", "pledged", "start", "end"):
                p[k] = num(p[k])
        except ValueError:
            return messagebox.showerror("خطا", "مقادیر عددی نامعتبر است.")
        if not p["name"]:
            return messagebox.showerror("خطا", "نام سهامدار را وارد کنید.")
        if p["pledged"] > p["total"]:
            return messagebox.showerror("خطا", "سهام وثیقه نمی‌تواند از کل سهام بیشتر باشد.")
        if p["start"] and p["end"]:
            if p["end"] < p["start"]:
                return messagebox.showerror("خطا", "پایان سهم کوچک‌تر از شروع است.")
            for i, o in enumerate(self.db["people"]):
                if i != self.edit_idx and o["start"] and o["end"] and \
                        not (p["end"] < o["start"] or p["start"] > o["end"]):
                    return messagebox.showerror(
                        "خطا", f"بازه سریال با «{o['name']}» همپوشانی دارد.")
            if p["end"] - p["start"] + 1 != p["total"]:
                if not messagebox.askyesno(
                        "هشدار", f"تعداد سهام بازه ({p['end']-p['start']+1}) با کل سهام ({p['total']}) برابر نیست. ادامه می‌دهید؟"):
                    return
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

    def rows(self):
        ppl = self.db["people"]
        tot = sum(p["total"] for p in ppl)
        nv = self.nominal_val()
        out = []
        for i, p in enumerate(ppl, 1):
            main = p["total"] - p["pledged"]
            pct = p["total"] / tot * 100 if tot else 0
            out.append([i, p["name"], p["nid"], p["total"], p["pledged"], main,
                        p["start"] or "", p["end"] or "", f"{pct:.2f}", p["total"] * nv, p["note"]])
        return out, tot, nv

    def refresh(self):
        self.db["company"] = self.company.get()
        self.db["nominal"] = self.nominal_val()
        rows, tot, nv = self.rows()
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(rows):
            vals = [fmt(x) if isinstance(x, int) and j in (3, 4, 5, 9) else x
                    for j, x in enumerate(r)]
            self.tree.insert("", "end", iid=str(i), values=vals)
        pl = sum(p["pledged"] for p in self.db["people"])
        self.summary.config(
            text=f"تعداد سهامداران: {len(rows)}   |   کل سهام: {fmt(tot)}   |   "
                 f"سهام وثیقه: {fmt(pl)}   |   سهام اصلی (آزاد): {fmt(tot-pl)}   |   "
                 f"سرمایه: {fmt(tot*nv)} ریال")
        json.dump(self.db, open(DATA, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    # ---------- output ----------
    def export(self):
        f = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not f:
            return
        rows, _, _ = self.rows()
        with open(f, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["ردیف", "نام", "کد ملی", "کل سهام", "وثیقه", "سهام اصلی",
                        "شروع", "پایان", "درصد", "ارزش اسمی", "توضیحات"])
            w.writerows(rows)

    def print_report(self):
        rows, tot, nv = self.rows()
        pl = sum(p["pledged"] for p in self.db["people"])
        e = html.escape
        head = ["ردیف", "نام سهامدار", "کد ملی", "کل سهام", "سهام وثیقه", "سهام اصلی (آزاد)",
                "شروع سهم", "پایان سهم", "درصد", "ارزش اسمی (ریال)", "توضیحات"]
        body = "".join(
            "<tr>" + "".join(
                f"<td>{e(fmt(x) if isinstance(x, int) and j in (3,4,5,9) else str(x))}</td>"
                for j, x in enumerate(r)) + "</tr>" for r in rows)
        foot = (f"<tr class=t><td colspan=3>جمع</td><td>{fmt(tot)}</td><td>{fmt(pl)}</td>"
                f"<td>{fmt(tot-pl)}</td><td></td><td></td><td>100.00</td><td>{fmt(tot*nv)}</td><td></td></tr>")
        doc = f"""<!DOCTYPE html><html dir="rtl" lang="fa"><meta charset="utf-8">
<style>body{{font-family:Tahoma,sans-serif;margin:20px}}h2,p{{text-align:center}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
td,th{{border:1px solid #444;padding:5px;text-align:center}}th{{background:#ddd}}
.t{{font-weight:bold;background:#eee}}@page{{size:A4 landscape;margin:10mm}}</style>
<h2>{e(self.company.get() or 'گزارش سهامداری')} - فهرست سهامداران</h2>
<p>تاریخ چاپ: {datetime.date.today()} &nbsp; | &nbsp; ارزش اسمی هر سهم: {fmt(nv)} ریال</p>
<table><tr>{''.join(f'<th>{h}</th>' for h in head)}</tr>{body}{foot}</table>
<script>window.onload=()=>window.print()</script></html>"""
        f = os.path.join(tempfile.gettempdir(), "shareholders_report.html")
        open(f, "w", encoding="utf-8").write(doc)
        webbrowser.open("file:///" + f.replace("\\", "/"))


if __name__ == "__main__":
    App().mainloop()
