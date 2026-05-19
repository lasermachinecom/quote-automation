"""売上・在庫管理 - Tkinter GUI

タブ:
  1. 入荷登録   2. 出荷登録   3. 一覧/検索   4. 在庫
  5. 詳細/画像  6. ツール
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from config import APP_TITLE, APP_VERSION, ATTACHMENTS_DIR, DB_PATH, EXPORT_DIR
from db import connect, find_units_by_serial, init_db, stock_summary
from export_csv import export_all_units, export_stock


UNIT_STATUSES = ["在庫", "予約済", "出荷済", "保留", "返品"]
SALE_METHODS = ["ヤフオク", "会社直販", "直メール", "Y直メ", "代理店", "1円スタート", "1円即決", "レンタル", "その他"]
PAYMENT_STATUSES = ["未入金", "入金済", "一部入金", "保留"]
ORDER_STATUSES = ["受注", "入金済", "検品済", "出荷済", "キャンセル"]
INSPECTION_STATUSES = ["未検品", "検品済"]


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def today() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def open_path(path: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        messagebox.showerror("エラー", f"ファイルを開けませんでした: {e}")


def parse_int(s: str) -> int | None:
    s = (s or "").strip().replace(",", "").replace("円", "")
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


class LabeledEntry(ttk.Frame):
    def __init__(self, master, label: str, width: int = 28, **kw):
        super().__init__(master, **kw)
        ttk.Label(self, text=label, width=14, anchor="e").pack(side="left", padx=(0, 6))
        self.var = tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self.var, width=width)
        self.entry.pack(side="left", fill="x", expand=True)

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, v: str):
        self.var.set(v or "")


class LabeledCombo(ttk.Frame):
    def __init__(self, master, label: str, values, width: int = 26, **kw):
        super().__init__(master, **kw)
        ttk.Label(self, text=label, width=14, anchor="e").pack(side="left", padx=(0, 6))
        self.var = tk.StringVar()
        self.combo = ttk.Combobox(self, textvariable=self.var, values=values, width=width)
        self.combo.pack(side="left", fill="x", expand=True)

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, v: str):
        self.var.set(v or "")


# ------------------------------------------------------------------
# Tab 1: 入荷登録
# ------------------------------------------------------------------

class IncomingTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self._build()

    def _build(self):
        ttk.Label(self, text="入荷登録（仕入れ）", font=("", 14, "bold")).pack(anchor="w", pady=(0, 8))

        form = ttk.LabelFrame(self, text="機械情報")
        form.pack(fill="x", pady=4)
        self.f_model = LabeledEntry(form, "機種名 *")
        self.f_serial = LabeledEntry(form, "シリアル番号")
        self.f_mfg = LabeledEntry(form, "製造年月日")
        self.f_unit_memo = LabeledEntry(form, "機体メモ", width=50)
        for w in (self.f_model, self.f_serial, self.f_mfg, self.f_unit_memo):
            w.pack(fill="x", pady=3, padx=6)
        self.f_serial.entry.bind("<FocusOut>", self._check_duplicate)

        self.dup_label = ttk.Label(self, text="", foreground="#c0392b")
        self.dup_label.pack(anchor="w", padx=12)

        form2 = ttk.LabelFrame(self, text="仕入情報")
        form2.pack(fill="x", pady=4)
        self.f_pdate = LabeledEntry(form2, "入荷日")
        self.f_pdate.set(today())
        self.f_vendor = LabeledEntry(form2, "仕入先")
        self.f_vcompany = LabeledEntry(form2, "仕入先会社")
        self.f_amount = LabeledEntry(form2, "仕入金額")
        self.f_invoice = LabeledEntry(form2, "請求書番号")
        self.f_pmemo = LabeledEntry(form2, "仕入メモ", width=50)
        for w in (self.f_pdate, self.f_vendor, self.f_vcompany, self.f_amount, self.f_invoice, self.f_pmemo):
            w.pack(fill="x", pady=3, padx=6)

        btn = ttk.Frame(self)
        btn.pack(fill="x", pady=10)
        ttk.Button(btn, text="登録", command=self._save).pack(side="left", padx=4)
        ttk.Button(btn, text="クリア", command=self._clear).pack(side="left", padx=4)

    def _check_duplicate(self, _evt=None):
        serial = self.f_serial.get()
        if not serial:
            self.dup_label.config(text="")
            return
        hits = find_units_by_serial(serial)
        if not hits:
            self.dup_label.config(text="", foreground="#27ae60")
            return
        active = [h for h in hits if h["status"] != "出荷済"]
        if active:
            self.dup_label.config(
                text=f"⚠ 同シリアルが在庫中に {len(active)} 件あります。重複入荷の可能性！",
                foreground="#c0392b",
            )
        else:
            self.dup_label.config(
                text=f"ℹ 同シリアルの過去出荷履歴あり ({len(hits)} 件)。中古再販なら問題なし。",
                foreground="#d35400",
            )

    def _save(self):
        model = self.f_model.get()
        if not model:
            messagebox.showwarning("入力エラー", "機種名は必須です")
            return
        amount = parse_int(self.f_amount.get())
        with connect() as conn:
            cur = conn.execute(
                "INSERT INTO units(serial_no, model, mfg_date, status, memo) VALUES (?,?,?,?,?)",
                (self.f_serial.get() or None, model, self.f_mfg.get() or None,
                 "在庫", self.f_unit_memo.get() or None),
            )
            unit_id = cur.lastrowid
            conn.execute(
                """INSERT INTO purchases(unit_id, purchase_date, vendor_name, vendor_company,
                                         amount, invoice_no, memo)
                   VALUES (?,?,?,?,?,?,?)""",
                (unit_id, self.f_pdate.get() or None, self.f_vendor.get() or None,
                 self.f_vcompany.get() or None, amount, self.f_invoice.get() or None,
                 self.f_pmemo.get() or None),
            )
        messagebox.showinfo("登録完了", f"入荷を登録しました（unit_id={unit_id}）")
        self._clear()
        self.app.refresh_all()

    def _clear(self):
        for w in (self.f_model, self.f_serial, self.f_mfg, self.f_unit_memo,
                  self.f_vendor, self.f_vcompany, self.f_amount, self.f_invoice, self.f_pmemo):
            w.set("")
        self.f_pdate.set(today())
        self.dup_label.config(text="")


# ------------------------------------------------------------------
# Tab 2: 出荷登録
# ------------------------------------------------------------------

class OutgoingTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self.selected_unit_id: int | None = None
        self._build()
        self.refresh()

    def _build(self):
        ttk.Label(self, text="出荷登録（販売）", font=("", 14, "bold")).pack(anchor="w", pady=(0, 8))

        top = ttk.Frame(self)
        top.pack(fill="both", expand=False)

        # Left: stock list
        left = ttk.LabelFrame(top, text="在庫から個体を選択")
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))

        filter_row = ttk.Frame(left)
        filter_row.pack(fill="x", padx=4, pady=4)
        ttk.Label(filter_row, text="絞り込み:").pack(side="left")
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(filter_row, textvariable=self.filter_var).pack(side="left", fill="x", expand=True, padx=4)

        self.tree = ttk.Treeview(left, columns=("model", "serial", "mfg", "status"), show="headings", height=10)
        for col, lbl, w in (("model", "機種", 180), ("serial", "シリアル", 100),
                            ("mfg", "製造日", 90), ("status", "状態", 70)):
            self.tree.heading(col, text=lbl)
            self.tree.column(col, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=4, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Right: customer form
        right = ttk.LabelFrame(top, text="お客様 / 出荷情報")
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))

        self.f_unit_label = ttk.Label(right, text="(個体未選択)", foreground="#888")
        self.f_unit_label.pack(anchor="w", padx=6, pady=4)

        self.f_sdate = LabeledEntry(right, "売上日"); self.f_sdate.set(today())
        self.f_ddate = LabeledEntry(right, "納品日")
        self.f_method = LabeledCombo(right, "販売方法", SALE_METHODS)
        self.f_cust = LabeledEntry(right, "お客様名")
        self.f_company = LabeledEntry(right, "会社名/発注番号")
        self.f_postal = LabeledEntry(right, "郵便番号")
        self.f_addr = LabeledEntry(right, "住所", width=50)
        self.f_phone = LabeledEntry(right, "電話番号")
        self.f_email = LabeledEntry(right, "メール")
        self.f_yid = LabeledEntry(right, "Yahoo ID")
        self.f_inv = LabeledEntry(right, "請求書番号")
        self.f_smonth = LabeledEntry(right, "売上計上月")
        self.f_freight = LabeledEntry(right, "送料")
        self.f_total = LabeledEntry(right, "合計金額 *")
        self.f_pay = LabeledCombo(right, "入金状態", PAYMENT_STATUSES)
        self.f_paydate = LabeledEntry(right, "入金日")
        self.f_smemo = LabeledEntry(right, "メモ", width=50)

        for w in (self.f_sdate, self.f_ddate, self.f_method, self.f_cust, self.f_company,
                  self.f_postal, self.f_addr, self.f_phone, self.f_email, self.f_yid,
                  self.f_inv, self.f_smonth, self.f_freight, self.f_total,
                  self.f_pay, self.f_paydate, self.f_smemo):
            w.pack(fill="x", pady=2, padx=6)

        btn = ttk.Frame(self)
        btn.pack(fill="x", pady=10)
        ttk.Button(btn, text="出荷を登録", command=self._save).pack(side="left", padx=4)
        ttk.Button(btn, text="クリア", command=self._clear).pack(side="left", padx=4)
        ttk.Label(btn, text="※ 納品先未定の場合は空欄のまま保存可。後で更新できます。",
                  foreground="#888").pack(side="left", padx=10)

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        keyword = self.filter_var.get().strip().lower() if hasattr(self, "filter_var") else ""
        with connect() as conn:
            rows = conn.execute(
                """SELECT u.id, u.model, u.serial_no, u.mfg_date, u.status
                   FROM units u
                   LEFT JOIN sales s ON s.unit_id = u.id
                   WHERE s.id IS NULL AND u.status != '出荷済'
                   ORDER BY u.id DESC"""
            ).fetchall()
        for r in rows:
            text = f"{r['model']} {r['serial_no'] or ''}".lower()
            if keyword and keyword not in text:
                continue
            self.tree.insert("", "end", iid=str(r["id"]),
                             values=(r["model"], r["serial_no"] or "", r["mfg_date"] or "", r["status"]))

    def _on_select(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        self.selected_unit_id = int(sel[0])
        with connect() as conn:
            u = conn.execute("SELECT * FROM units WHERE id=?", (self.selected_unit_id,)).fetchone()
        self.f_unit_label.config(
            text=f"選択中: #{u['id']}  {u['model']}  シリアル: {u['serial_no'] or '(なし)'}",
            foreground="#2c3e50",
        )

    def _save(self):
        if not self.selected_unit_id:
            messagebox.showwarning("選択エラー", "在庫から個体を選択してください")
            return
        total = parse_int(self.f_total.get())
        if total is None:
            messagebox.showwarning("入力エラー", "合計金額は必須です（数値）")
            return
        freight = parse_int(self.f_freight.get())
        try:
            with connect() as conn:
                conn.execute(
                    """INSERT INTO sales(unit_id, sale_date, delivery_date,
                        customer_name, customer_company, postal, address, phone, email, yahoo_id,
                        sale_method, invoice_no, sale_month, freight, total_amount,
                        payment_status, payment_date, memo)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (self.selected_unit_id, self.f_sdate.get() or None, self.f_ddate.get() or None,
                     self.f_cust.get() or None, self.f_company.get() or None,
                     self.f_postal.get() or None, self.f_addr.get() or None,
                     self.f_phone.get() or None, self.f_email.get() or None, self.f_yid.get() or None,
                     self.f_method.get() or None, self.f_inv.get() or None, self.f_smonth.get() or None,
                     freight, total, self.f_pay.get() or None,
                     self.f_paydate.get() or None, self.f_smemo.get() or None),
                )
                conn.execute("UPDATE units SET status='出荷済', updated_at=datetime('now','localtime') WHERE id=?",
                             (self.selected_unit_id,))
        except Exception as e:
            messagebox.showerror("エラー", f"登録に失敗しました: {e}")
            return
        messagebox.showinfo("登録完了", "出荷を登録しました")
        self._clear()
        self.app.refresh_all()

    def _clear(self):
        self.selected_unit_id = None
        self.f_unit_label.config(text="(個体未選択)", foreground="#888")
        for w in (self.f_ddate, self.f_method, self.f_cust, self.f_company,
                  self.f_postal, self.f_addr, self.f_phone, self.f_email, self.f_yid,
                  self.f_inv, self.f_smonth, self.f_freight, self.f_total,
                  self.f_pay, self.f_paydate, self.f_smemo):
            w.set("")
        self.f_sdate.set(today())


# ------------------------------------------------------------------
# Tab 3: 一覧 / 検索
# ------------------------------------------------------------------

class ListTab(ttk.Frame):
    COLUMNS = [
        ("unit_id", "ID", 50),
        ("sale_date", "売上日", 90),
        ("model", "機種", 200),
        ("serial_no", "シリアル", 100),
        ("customer_name", "お客様", 180),
        ("address", "住所", 220),
        ("sale_method", "販売方法", 100),
        ("total_amount", "金額", 90),
        ("payment_status", "入金", 70),
    ]

    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self._build()
        self.refresh()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="検索:").pack(side="left")
        self.q = tk.StringVar()
        self.q.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(top, textvariable=self.q, width=30).pack(side="left", padx=4)

        ttk.Label(top, text="機種:").pack(side="left", padx=(10, 2))
        self.model_q = tk.StringVar()
        self.model_q.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(top, textvariable=self.model_q, width=20).pack(side="left")

        ttk.Label(top, text="年:").pack(side="left", padx=(10, 2))
        self.year_q = tk.StringVar()
        self.year_q.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(top, textvariable=self.year_q, width=8).pack(side="left")

        ttk.Button(top, text="CSV書き出し", command=self._export).pack(side="right")

        self.tree = ttk.Treeview(self, columns=[c[0] for c in self.COLUMNS], show="headings", height=20)
        for col, lbl, w in self.COLUMNS:
            self.tree.heading(col, text=lbl, command=lambda c=col: self._sort(c))
            self.tree.column(col, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_double)

        ysb = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        ysb.pack(side="right", fill="y")

        self.count_label = ttk.Label(self, text="0 件")
        self.count_label.pack(anchor="w", pady=4)

        self._sort_col = "sale_date"
        self._sort_asc = False

    def _query(self):
        q = self.q.get().strip()
        m = self.model_q.get().strip()
        y = self.year_q.get().strip()
        where = ["1=1"]
        params: list = []
        if q:
            where.append("(s.customer_name LIKE ? OR s.address LIKE ? OR u.serial_no LIKE ?)")
            params += [f"%{q}%"] * 3
        if m:
            where.append("u.model LIKE ?")
            params.append(f"%{m}%")
        if y:
            where.append("s.sale_date LIKE ?")
            params.append(f"%{y}%")
        sql = f"""
            SELECT u.id AS unit_id, s.sale_date, u.model, u.serial_no,
                   s.customer_name, s.address, s.sale_method,
                   s.total_amount, s.payment_status
            FROM units u
            LEFT JOIN sales s ON s.unit_id = u.id
            WHERE {' AND '.join(where)}
            ORDER BY s.sale_date DESC, u.id DESC
        """
        with connect() as conn:
            return list(conn.execute(sql, params))

    def refresh(self):
        rows = self._query()
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", iid=str(r["unit_id"]),
                             values=tuple(r[c[0]] if r[c[0]] is not None else "" for c in self.COLUMNS))
        self.count_label.config(text=f"{len(rows)} 件")

    def _sort(self, col):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        rows = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        try:
            rows.sort(key=lambda x: (float(str(x[0]).replace(",", "") or 0)), reverse=not self._sort_asc)
        except ValueError:
            rows.sort(reverse=not self._sort_asc)
        for i, (_, k) in enumerate(rows):
            self.tree.move(k, "", i)

    def _on_double(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        self.app.show_detail(int(sel[0]))

    def _export(self):
        rows = self._query()
        from export_csv import export_sales
        dest = export_sales(rows)
        messagebox.showinfo("CSV書き出し", f"書き出しました:\n{dest}")
        open_path(dest.parent)


# ------------------------------------------------------------------
# Tab 4: 在庫
# ------------------------------------------------------------------

class StockTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self._build()
        self.refresh()

    def _build(self):
        ttk.Label(self, text="在庫一覧（未出荷の個体）", font=("", 14, "bold")).pack(anchor="w", pady=(0, 8))

        btns = ttk.Frame(self)
        btns.pack(fill="x", pady=4)
        ttk.Button(btns, text="更新", command=self.refresh).pack(side="left")
        ttk.Button(btns, text="在庫CSV書き出し", command=self._export).pack(side="left", padx=6)
        ttk.Label(btns, text="    機種で絞り込み:").pack(side="left", padx=(20, 4))
        self.filter_var = tk.StringVar()
        ent = ttk.Entry(btns, textvariable=self.filter_var, width=24)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda e: self.refresh())

        cols = ("serial", "model", "mfg_date", "purchase_date", "vendor", "amount", "memo")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=22)
        for k, t, w, a in [
            ("serial", "シリアルNo.", 110, "w"),
            ("model", "機種", 220, "w"),
            ("mfg_date", "製造年月日", 100, "w"),
            ("purchase_date", "入荷日", 100, "w"),
            ("vendor", "仕入先", 140, "w"),
            ("amount", "入荷金額", 100, "e"),
            ("memo", "メモ", 320, "w"),
        ]:
            self.tree.heading(k, text=t)
            self.tree.column(k, width=w, anchor=a)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._open_detail)

        self.total_label = ttk.Label(self, text="合計: 0 台", font=("", 11, "bold"))
        self.total_label.pack(anchor="e", pady=6)

    def refresh(self):
        kw = (self.filter_var.get() if hasattr(self, "filter_var") else "").strip()
        sql = """
            SELECT u.id, u.serial_no, u.model, u.mfg_date, u.memo AS unit_memo,
                   p.purchase_date, p.vendor_name, p.amount
            FROM units u
            LEFT JOIN sales s ON s.unit_id = u.id
            LEFT JOIN purchases p ON p.unit_id = u.id
            WHERE s.id IS NULL AND u.status != '出荷済'
        """
        params: list = []
        if kw:
            sql += " AND (u.model LIKE ? OR u.serial_no LIKE ?)"
            like = f"%{kw}%"
            params += [like, like]
        sql += " ORDER BY u.model, u.id"
        with connect() as conn:
            rows = list(conn.execute(sql, params))

        self.tree.delete(*self.tree.get_children())
        total_amount = 0
        for r in rows:
            memo = (r["unit_memo"] or "").replace("\n", " / ")
            amt = r["amount"]
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["serial_no"] or "",
                r["model"] or "",
                r["mfg_date"] or "",
                r["purchase_date"] or "",
                r["vendor_name"] or "",
                f"{amt:,}" if amt is not None else "",
                memo,
            ))
            if amt is not None:
                total_amount += amt
        self.total_label.config(text=f"合計: {len(rows)} 台 / 入荷金額計: {total_amount:,} 円")

    def _open_detail(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        unit_id = int(sel[0])
        self.app.detail.load(unit_id)
        self.app.nb.select(self.app.detail)

    def _export(self):
        dest = export_stock()
        messagebox.showinfo("CSV書き出し", f"書き出しました:\n{dest}")
        open_path(dest.parent)


# ------------------------------------------------------------------
# Tab 5: 詳細 / 画像
# ------------------------------------------------------------------

class DetailTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self.unit_id: int | None = None
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text="個体ID:").pack(side="left")
        self.id_var = tk.StringVar()
        e = ttk.Entry(top, textvariable=self.id_var, width=10)
        e.pack(side="left", padx=4)
        e.bind("<Return>", lambda _e: self.load(parse_int(self.id_var.get())))
        ttk.Button(top, text="表示", command=lambda: self.load(parse_int(self.id_var.get()))).pack(side="left")
        ttk.Label(top, text="（一覧タブの行をダブルクリックでも開きます）", foreground="#888").pack(side="left", padx=10)

        self.info = tk.Text(self, height=18, wrap="word")
        self.info.pack(fill="both", expand=True, pady=8)
        self.info.config(state="disabled")

        ttk.Label(self, text="添付ファイル", font=("", 11, "bold")).pack(anchor="w", pady=(8, 2))
        att_row = ttk.Frame(self)
        att_row.pack(fill="x")
        ttk.Button(att_row, text="画像/伝票を追加", command=self._add_attachment).pack(side="left")
        ttk.Button(att_row, text="フォルダを開く", command=self._open_folder).pack(side="left", padx=6)
        ttk.Button(att_row, text="メモ編集", command=self._edit_memo).pack(side="left", padx=6)
        ttk.Button(att_row, text="このレコードを削除", command=self._delete).pack(side="right")

        self.att_list = tk.Listbox(self, height=6)
        self.att_list.pack(fill="x", pady=4)
        self.att_list.bind("<Double-1>", self._open_attachment)

    def load(self, unit_id: int | None):
        if not unit_id:
            return
        self.unit_id = unit_id
        self.id_var.set(str(unit_id))
        with connect() as conn:
            u = conn.execute("SELECT * FROM units WHERE id=?", (unit_id,)).fetchone()
            if not u:
                messagebox.showerror("エラー", "該当個体が見つかりません")
                return
            p = conn.execute("SELECT * FROM purchases WHERE unit_id=?", (unit_id,)).fetchone()
            s = conn.execute("SELECT * FROM sales WHERE unit_id=?", (unit_id,)).fetchone()
            atts = conn.execute(
                "SELECT * FROM attachments WHERE unit_id=? ORDER BY created_at DESC",
                (unit_id,),
            ).fetchall()

        lines = []
        lines.append(f"=== 個体 #{u['id']} ===")
        lines.append(f"機種:       {u['model']}")
        lines.append(f"シリアル:   {u['serial_no'] or '(なし)'}")
        lines.append(f"製造年月日: {u['mfg_date'] or ''}")
        lines.append(f"状態:       {u['status']}")
        lines.append(f"機体メモ:   {u['memo'] or ''}")
        lines.append("")
        lines.append("--- 入荷 ---")
        if p:
            lines.append(f"入荷日:     {p['purchase_date'] or ''}")
            lines.append(f"仕入先:     {p['vendor_name'] or ''}  {p['vendor_company'] or ''}")
            lines.append(f"仕入金額:   {p['amount'] or ''}")
            lines.append(f"請求書番号: {p['invoice_no'] or ''}")
            lines.append(f"仕入メモ:   {p['memo'] or ''}")
        else:
            lines.append("(入荷情報なし — 既存Excelからのインポート分は未入力です)")
        lines.append("")
        lines.append("--- 出荷 ---")
        if s:
            lines.append(f"売上日:     {s['sale_date'] or ''}")
            lines.append(f"納品日:     {s['delivery_date'] or ''}")
            lines.append(f"お客様:     {s['customer_name'] or ''}")
            lines.append(f"会社/番号:  {s['customer_company'] or ''}")
            lines.append(f"郵便番号:   {s['postal'] or ''}")
            lines.append(f"住所:       {s['address'] or ''}")
            lines.append(f"電話:       {s['phone'] or ''}")
            lines.append(f"メール:     {s['email'] or ''}")
            lines.append(f"Yahoo ID:   {s['yahoo_id'] or ''}")
            lines.append(f"販売方法:   {s['sale_method'] or ''}")
            lines.append(f"請求書:     {s['invoice_no'] or ''}")
            lines.append(f"計上月:     {s['sale_month'] or ''}")
            lines.append(f"送料:       {s['freight'] or ''}")
            lines.append(f"合計金額:   {s['total_amount'] or ''}")
            lines.append(f"入金:       {s['payment_status'] or ''}  入金日: {s['payment_date'] or ''}")
            lines.append(f"出荷メモ:   {s['memo'] or ''}")
        else:
            lines.append("(未出荷)")

        self.info.config(state="normal")
        self.info.delete("1.0", "end")
        self.info.insert("1.0", "\n".join(lines))
        self.info.config(state="disabled")

        self.att_list.delete(0, "end")
        self.attachments = list(atts)
        for a in self.attachments:
            self.att_list.insert("end", f"[{a['kind'] or 'その他'}] {Path(a['file_path']).name}  {a['caption'] or ''}")

    def _add_attachment(self):
        if not self.unit_id:
            messagebox.showwarning("選択エラー", "先に個体を表示してください")
            return
        path = filedialog.askopenfilename(title="添付するファイルを選択",
                                         filetypes=[("画像/PDF", "*.jpg *.jpeg *.png *.gif *.bmp *.pdf"),
                                                    ("すべて", "*.*")])
        if not path:
            return
        src = Path(path)
        kind = self._ask_kind()
        if kind is None:
            return
        dest_dir = ATTACHMENTS_DIR / f"unit_{self.unit_id}"
        dest_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = dest_dir / f"{ts}_{src.name}"
        shutil.copy2(src, dest)
        with connect() as conn:
            conn.execute(
                "INSERT INTO attachments(unit_id, file_path, kind) VALUES (?,?,?)",
                (self.unit_id, str(dest.relative_to(ATTACHMENTS_DIR.parent)), kind),
            )
        self.load(self.unit_id)

    def _ask_kind(self) -> str | None:
        dlg = tk.Toplevel(self)
        dlg.title("種別")
        dlg.transient(self.winfo_toplevel())
        result = {"value": None}
        ttk.Label(dlg, text="ファイルの種別を選んでください").pack(padx=12, pady=8)
        for k in ("検品写真", "伝票写真", "納品書", "保証書", "その他"):
            ttk.Button(dlg, text=k, width=18,
                       command=lambda v=k: (result.update(value=v), dlg.destroy())).pack(pady=2)
        ttk.Button(dlg, text="キャンセル", command=dlg.destroy).pack(pady=8)
        dlg.grab_set()
        self.wait_window(dlg)
        return result["value"]

    def _open_attachment(self, _evt=None):
        sel = self.att_list.curselection()
        if not sel:
            return
        a = self.attachments[sel[0]]
        path = ATTACHMENTS_DIR.parent / a["file_path"]
        open_path(path)

    def _open_folder(self):
        if not self.unit_id:
            return
        folder = ATTACHMENTS_DIR / f"unit_{self.unit_id}"
        folder.mkdir(parents=True, exist_ok=True)
        open_path(folder)

    def _edit_memo(self):
        if not self.unit_id:
            return
        with connect() as conn:
            u = conn.execute("SELECT memo FROM units WHERE id=?", (self.unit_id,)).fetchone()
            s = conn.execute("SELECT memo FROM sales WHERE unit_id=?", (self.unit_id,)).fetchone()
        dlg = tk.Toplevel(self)
        dlg.title(f"メモ編集 unit #{self.unit_id}")
        dlg.transient(self.winfo_toplevel())
        ttk.Label(dlg, text="機体メモ").pack(anchor="w", padx=8, pady=(8, 2))
        t1 = tk.Text(dlg, height=4, width=60)
        t1.pack(padx=8)
        t1.insert("1.0", u["memo"] or "")
        ttk.Label(dlg, text="出荷メモ").pack(anchor="w", padx=8, pady=(8, 2))
        t2 = tk.Text(dlg, height=4, width=60)
        t2.pack(padx=8)
        t2.insert("1.0", (s["memo"] if s else "") or "")

        def save():
            with connect() as conn:
                conn.execute("UPDATE units SET memo=?, updated_at=datetime('now','localtime') WHERE id=?",
                             (t1.get("1.0", "end").strip() or None, self.unit_id))
                if s:
                    conn.execute("UPDATE sales SET memo=? WHERE unit_id=?",
                                 (t2.get("1.0", "end").strip() or None, self.unit_id))
            dlg.destroy()
            self.load(self.unit_id)

        ttk.Button(dlg, text="保存", command=save).pack(pady=8)
        dlg.grab_set()

    def _delete(self):
        if not self.unit_id:
            return
        if not messagebox.askyesno("削除確認", f"個体 #{self.unit_id} とそれに紐づく入荷/出荷/添付を削除します。よろしいですか？"):
            return
        with connect() as conn:
            conn.execute("DELETE FROM units WHERE id=?", (self.unit_id,))
        self.unit_id = None
        self.id_var.set("")
        self.info.config(state="normal")
        self.info.delete("1.0", "end")
        self.info.config(state="disabled")
        self.att_list.delete(0, "end")
        self.app.refresh_all()


# ------------------------------------------------------------------
# Tab 7: 受注
# ------------------------------------------------------------------

class OrderTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self._build()
        self.refresh()

    def _build(self):
        ttk.Label(self, text="受注管理", font=("", 14, "bold")).pack(anchor="w", pady=(0, 8))

        top = ttk.Frame(self)
        top.pack(fill="x", pady=4)
        ttk.Label(top, text="状態:").pack(side="left")
        self.status_filter = tk.StringVar(value="(全部)")
        cb = ttk.Combobox(top, textvariable=self.status_filter,
                          values=["(全部)"] + ORDER_STATUSES, width=10, state="readonly")
        cb.pack(side="left", padx=4)
        self.status_filter.trace_add("write", lambda *_: self.refresh())

        ttk.Label(top, text="絞り込み:").pack(side="left", padx=(10, 4))
        self.q_var = tk.StringVar()
        self.q_var.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(top, textvariable=self.q_var, width=24).pack(side="left")

        ttk.Button(top, text="＋ 新規受注", command=self._new).pack(side="left", padx=(20, 4))
        ttk.Button(top, text="編集", command=self._edit).pack(side="left", padx=2)
        ttk.Button(top, text="削除", command=self._delete).pack(side="left", padx=2)
        ttk.Button(top, text="📦 出荷確定", command=self._ship).pack(side="left", padx=(20, 4))

        cols = ("status", "order_date", "customer", "model", "total",
                "payment", "invoice", "desired", "assigned")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=22)
        for k, t, w, a in [
            ("status", "状態", 70, "w"),
            ("order_date", "受注日", 90, "w"),
            ("customer", "お客様", 180, "w"),
            ("model", "受注機種", 200, "w"),
            ("total", "合計金額", 100, "e"),
            ("payment", "入金", 70, "w"),
            ("invoice", "請求書#", 100, "w"),
            ("desired", "希望出荷日", 90, "w"),
            ("assigned", "割当シリアル", 110, "w"),
        ]:
            self.tree.heading(k, text=t)
            self.tree.column(k, width=w, anchor=a)
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<Double-1>", lambda _e: self._edit())

        self.count_label = ttk.Label(self, text="0 件")
        self.count_label.pack(anchor="w", pady=2)

    def refresh(self):
        sql = """
            SELECT o.*, u.serial_no AS assigned_serial
            FROM orders o
            LEFT JOIN units u ON u.id = o.assigned_unit_id
            WHERE 1=1
        """
        params: list = []
        st = self.status_filter.get()
        if st and st != "(全部)":
            sql += " AND o.status = ?"
            params.append(st)
        kw = self.q_var.get().strip()
        if kw:
            sql += (" AND (o.customer_name LIKE ? OR o.customer_company LIKE ?"
                    " OR o.model_requested LIKE ? OR o.invoice_no LIKE ?)")
            like = f"%{kw}%"
            params += [like, like, like, like]
        sql += " ORDER BY o.order_date DESC, o.id DESC"

        with connect() as conn:
            rows = list(conn.execute(sql, params))

        self.tree.delete(*self.tree.get_children())
        for r in rows:
            total = r["total_amount"]
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["status"] or "",
                r["order_date"] or "",
                (r["customer_name"] or "") + (f" / {r['customer_company']}" if r["customer_company"] else ""),
                r["model_requested"] or "",
                f"{total:,}" if total is not None else "",
                r["payment_status"] or "",
                r["invoice_no"] or "",
                r["desired_ship_date"] or "",
                r["assigned_serial"] or "",
            ))
        self.count_label.config(text=f"{len(rows)} 件")

    def _selected_id(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def _new(self):
        OrderDialog(self, self.app, on_save=self.refresh)

    def _edit(self):
        oid = self._selected_id()
        if not oid:
            messagebox.showinfo("情報", "編集する受注を選択してください")
            return
        OrderDialog(self, self.app, order_id=oid, on_save=self.refresh)

    def _delete(self):
        oid = self._selected_id()
        if not oid:
            return
        if not messagebox.askyesno("確認", f"受注 #{oid} を削除しますか？"):
            return
        with connect() as conn:
            conn.execute("DELETE FROM orders WHERE id=?", (oid,))
        self.refresh()

    def _ship(self):
        oid = self._selected_id()
        if not oid:
            messagebox.showinfo("情報", "出荷確定する受注を選択してください")
            return
        with connect() as conn:
            o = conn.execute("SELECT * FROM orders WHERE id=?", (oid,)).fetchone()
        if not o:
            return
        if o["status"] == "出荷済":
            messagebox.showinfo("情報", "この受注は既に出荷済みです")
            return
        ShipDialog(self, self.app, order=dict(o), on_done=self._on_shipped)

    def _on_shipped(self):
        self.refresh()
        self.app.refresh_all()


class OrderDialog(tk.Toplevel):
    """新規 / 編集 受注フォーム（ポップアップ）"""

    def __init__(self, parent, app, order_id=None, on_save=None):
        super().__init__(parent)
        self.app = app
        self.order_id = order_id
        self.on_save = on_save
        self.title(f"受注 編集 #{order_id}" if order_id else "受注 新規登録")
        self.transient(parent)
        self.grab_set()
        self.geometry("680x640")
        self._build()
        if order_id:
            self._load(order_id)
        else:
            self.f_odate.set(today())

    def _build(self):
        wrap = ttk.Frame(self, padding=10)
        wrap.pack(fill="both", expand=True)

        f1 = ttk.LabelFrame(wrap, text="受注基本")
        f1.pack(fill="x", pady=4)
        self.f_odate = LabeledEntry(f1, "受注日")
        self.f_status = LabeledCombo(f1, "状態", ORDER_STATUSES)
        self.f_status.set("受注")
        self.f_method = LabeledCombo(f1, "販売方法", SALE_METHODS)
        self.f_model = LabeledEntry(f1, "受注機種")
        for w in (self.f_odate, self.f_status, self.f_method, self.f_model):
            w.pack(fill="x", pady=2, padx=6)

        f2 = ttk.LabelFrame(wrap, text="お客様")
        f2.pack(fill="x", pady=4)
        self.f_cust = LabeledEntry(f2, "お客様名")
        self.f_company = LabeledEntry(f2, "会社名/発注番号")
        self.f_postal = LabeledEntry(f2, "郵便番号")
        self.f_addr = LabeledEntry(f2, "住所", width=46)
        self.f_phone = LabeledEntry(f2, "電話番号")
        self.f_email = LabeledEntry(f2, "メール")
        self.f_yid = LabeledEntry(f2, "Yahoo ID")
        for w in (self.f_cust, self.f_company, self.f_postal, self.f_addr,
                  self.f_phone, self.f_email, self.f_yid):
            w.pack(fill="x", pady=2, padx=6)

        f3 = ttk.LabelFrame(wrap, text="金額・入金・出荷予定")
        f3.pack(fill="x", pady=4)
        self.f_invoice = LabeledEntry(f3, "請求書番号")
        self.f_total = LabeledEntry(f3, "合計金額")
        self.f_freight = LabeledEntry(f3, "送料")
        self.f_pay = LabeledCombo(f3, "入金状態", PAYMENT_STATUSES)
        self.f_pay.set("未入金")
        self.f_paydate = LabeledEntry(f3, "入金日")
        self.f_dship = LabeledEntry(f3, "希望出荷日")
        self.f_insp = LabeledCombo(f3, "検品状態", INSPECTION_STATUSES)
        self.f_insp.set("未検品")
        for w in (self.f_invoice, self.f_total, self.f_freight, self.f_pay,
                  self.f_paydate, self.f_dship, self.f_insp):
            w.pack(fill="x", pady=2, padx=6)

        self.f_memo = LabeledEntry(wrap, "メモ", width=60)
        self.f_memo.pack(fill="x", pady=6)

        btn = ttk.Frame(wrap)
        btn.pack(fill="x", pady=10)
        ttk.Button(btn, text="保存", command=self._save).pack(side="right", padx=4)
        ttk.Button(btn, text="キャンセル", command=self.destroy).pack(side="right", padx=4)

    def _load(self, oid):
        with connect() as conn:
            o = conn.execute("SELECT * FROM orders WHERE id=?", (oid,)).fetchone()
        if not o:
            messagebox.showerror("エラー", "受注が見つかりません")
            self.destroy()
            return
        self.f_odate.set(o["order_date"] or "")
        self.f_status.set(o["status"] or "受注")
        self.f_method.set(o["sale_method"] or "")
        self.f_model.set(o["model_requested"] or "")
        self.f_cust.set(o["customer_name"] or "")
        self.f_company.set(o["customer_company"] or "")
        self.f_postal.set(o["postal"] or "")
        self.f_addr.set(o["address"] or "")
        self.f_phone.set(o["phone"] or "")
        self.f_email.set(o["email"] or "")
        self.f_yid.set(o["yahoo_id"] or "")
        self.f_invoice.set(o["invoice_no"] or "")
        self.f_total.set(str(o["total_amount"]) if o["total_amount"] is not None else "")
        self.f_freight.set(str(o["freight"]) if o["freight"] is not None else "")
        self.f_pay.set(o["payment_status"] or "未入金")
        self.f_paydate.set(o["payment_date"] or "")
        self.f_dship.set(o["desired_ship_date"] or "")
        self.f_insp.set(o["inspection_status"] or "未検品")
        self.f_memo.set(o["memo"] or "")

    def _save(self):
        cust = self.f_cust.get()
        model = self.f_model.get()
        if not cust and not model:
            messagebox.showwarning("入力エラー", "お客様名 または 受注機種 は必須です")
            return

        values = {
            "order_date": self.f_odate.get() or None,
            "customer_name": cust or None,
            "customer_company": self.f_company.get() or None,
            "postal": self.f_postal.get() or None,
            "address": self.f_addr.get() or None,
            "phone": self.f_phone.get() or None,
            "email": self.f_email.get() or None,
            "yahoo_id": self.f_yid.get() or None,
            "sale_method": self.f_method.get() or None,
            "model_requested": model or None,
            "invoice_no": self.f_invoice.get() or None,
            "total_amount": parse_int(self.f_total.get()),
            "freight": parse_int(self.f_freight.get()),
            "payment_status": self.f_pay.get() or "未入金",
            "payment_date": self.f_paydate.get() or None,
            "desired_ship_date": self.f_dship.get() or None,
            "inspection_status": self.f_insp.get() or "未検品",
            "status": self.f_status.get() or "受注",
            "memo": self.f_memo.get() or None,
        }
        with connect() as conn:
            if self.order_id:
                sets = ", ".join(f"{k}=?" for k in values)
                conn.execute(
                    f"UPDATE orders SET {sets}, updated_at=datetime('now','localtime') WHERE id=?",
                    list(values.values()) + [self.order_id],
                )
            else:
                cols = ", ".join(values.keys())
                ph = ", ".join(["?"] * len(values))
                conn.execute(f"INSERT INTO orders({cols}) VALUES ({ph})", list(values.values()))
        if self.on_save:
            self.on_save()
        self.destroy()


class ShipDialog(tk.Toplevel):
    """受注を出荷確定するダイアログ：在庫から個体を選んで売上に転記"""

    def __init__(self, parent, app, order: dict, on_done=None):
        super().__init__(parent)
        self.app = app
        self.order = order
        self.on_done = on_done
        self.selected_unit_id = None
        self.title(f"出荷確定（受注 #{order['id']}）")
        self.transient(parent)
        self.grab_set()
        self.geometry("760x520")
        self._build()
        self._refresh_units()

    def _build(self):
        wrap = ttk.Frame(self, padding=10)
        wrap.pack(fill="both", expand=True)

        info = ttk.LabelFrame(wrap, text="受注内容")
        info.pack(fill="x", pady=4)
        ttk.Label(info, text=(
            f"お客様: {self.order.get('customer_name') or ''}    "
            f"会社: {self.order.get('customer_company') or ''}\n"
            f"受注機種: {self.order.get('model_requested') or ''}    "
            f"合計: {self.order.get('total_amount') or ''}    "
            f"入金: {self.order.get('payment_status') or ''}"
        ), justify="left").pack(anchor="w", padx=6, pady=4)

        sel = ttk.LabelFrame(wrap, text="在庫から個体を選択（シリアル割当）")
        sel.pack(fill="both", expand=True, pady=4)

        filt = ttk.Frame(sel)
        filt.pack(fill="x", padx=4, pady=4)
        ttk.Label(filt, text="絞り込み:").pack(side="left")
        self.q_var = tk.StringVar(value=self.order.get("model_requested") or "")
        self.q_var.trace_add("write", lambda *_: self._refresh_units())
        ttk.Entry(filt, textvariable=self.q_var, width=30).pack(side="left", padx=4)

        self.tree = ttk.Treeview(sel, columns=("serial", "model", "mfg", "purchase"),
                                 show="headings", height=12)
        for k, t, w, a in [("serial", "シリアル", 120, "w"),
                            ("model", "機種", 240, "w"),
                            ("mfg", "製造日", 100, "w"),
                            ("purchase", "入荷日", 100, "w")]:
            self.tree.heading(k, text=t)
            self.tree.column(k, width=w, anchor=a)
        self.tree.pack(fill="both", expand=True, padx=4, pady=4)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        bottom = ttk.Frame(wrap)
        bottom.pack(fill="x", pady=8)
        self.f_ddate = LabeledEntry(bottom, "実出荷日")
        self.f_ddate.set(today())
        self.f_ddate.pack(side="left")

        ttk.Button(bottom, text="✓ 出荷確定", command=self._confirm).pack(side="right", padx=4)
        ttk.Button(bottom, text="キャンセル", command=self.destroy).pack(side="right", padx=4)

    def _refresh_units(self):
        kw = self.q_var.get().strip().lower()
        with connect() as conn:
            rows = conn.execute("""
                SELECT u.id, u.serial_no, u.model, u.mfg_date, p.purchase_date
                FROM units u
                LEFT JOIN sales s ON s.unit_id = u.id
                LEFT JOIN purchases p ON p.unit_id = u.id
                WHERE s.id IS NULL AND u.status != '出荷済'
                ORDER BY u.model, u.id
            """).fetchall()
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            txt = f"{r['model'] or ''} {r['serial_no'] or ''}".lower()
            if kw and kw not in txt:
                continue
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["serial_no"] or "",
                r["model"] or "",
                r["mfg_date"] or "",
                r["purchase_date"] or "",
            ))

    def _on_select(self, _evt=None):
        sel = self.tree.selection()
        self.selected_unit_id = int(sel[0]) if sel else None

    def _confirm(self):
        if not self.selected_unit_id:
            messagebox.showwarning("選択エラー", "在庫から個体を選択してください")
            return
        ddate = self.f_ddate.get() or today()
        o = self.order
        try:
            with connect() as conn:
                cur = conn.execute(
                    """INSERT INTO sales(unit_id, sale_date, delivery_date,
                        customer_name, customer_company, postal, address, phone, email, yahoo_id,
                        sale_method, invoice_no, freight, total_amount,
                        payment_status, payment_date, memo)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (self.selected_unit_id, o.get("order_date"), ddate,
                     o.get("customer_name"), o.get("customer_company"),
                     o.get("postal"), o.get("address"), o.get("phone"),
                     o.get("email"), o.get("yahoo_id"),
                     o.get("sale_method"), o.get("invoice_no"),
                     o.get("freight"), o.get("total_amount"),
                     o.get("payment_status"), o.get("payment_date"),
                     o.get("memo")),
                )
                sale_id = cur.lastrowid
                conn.execute(
                    "UPDATE units SET status='出荷済', updated_at=datetime('now','localtime') WHERE id=?",
                    (self.selected_unit_id,),
                )
                conn.execute(
                    """UPDATE orders
                       SET status='出荷済', assigned_unit_id=?, sale_id=?,
                           updated_at=datetime('now','localtime')
                       WHERE id=?""",
                    (self.selected_unit_id, sale_id, o["id"]),
                )
        except Exception as e:
            messagebox.showerror("エラー", f"出荷確定に失敗しました: {e}")
            return
        messagebox.showinfo("完了", "出荷確定しました。売上に計上されました。")
        if self.on_done:
            self.on_done()
        self.destroy()


# ------------------------------------------------------------------
# Tab 6: ツール
# ------------------------------------------------------------------

class ToolsTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self._build()

    def _build(self):
        ttk.Label(self, text="ツール", font=("", 14, "bold")).pack(anchor="w", pady=(0, 8))

        f1 = ttk.LabelFrame(self, text="Excel取込（シリアル照合表）")
        f1.pack(fill="x", pady=4)
        ttk.Label(f1, text="「全機種　シリアル照合表」シートを読み込んでDBに取り込みます。\n"
                          "出荷日が空の行→在庫、出荷日に日付がある行→出荷済として登録します。\n"
                          "※重複登録を避けるため、再取込前にDBバックアップ推奨。",
                  foreground="#555", justify="left").pack(anchor="w", padx=8, pady=4)
        ttk.Button(f1, text="Excelファイルを選択してインポート", command=self._import).pack(anchor="w", padx=8, pady=4)

        f2 = ttk.LabelFrame(self, text="CSV書き出し")
        f2.pack(fill="x", pady=10)
        row = ttk.Frame(f2)
        row.pack(fill="x", padx=8, pady=6)
        ttk.Button(row, text="全データ書き出し (入荷+出荷)", command=self._export_all).pack(side="left", padx=2)
        ttk.Button(row, text="在庫のみ", command=self._export_stock).pack(side="left", padx=2)
        ttk.Button(row, text="書き出しフォルダを開く", command=lambda: open_path(EXPORT_DIR)).pack(side="left", padx=10)

        f3 = ttk.LabelFrame(self, text="DBバックアップ")
        f3.pack(fill="x", pady=10)
        ttk.Label(f3, text=f"DB: {DB_PATH}", foreground="#555").pack(anchor="w", padx=8)
        ttk.Button(f3, text="DBファイルをコピー保存", command=self._backup).pack(anchor="w", padx=8, pady=4)

        ttk.Label(self, text=f"バージョン: {APP_VERSION}", foreground="#888").pack(anchor="e", pady=8)

    def _import(self):
        path = filedialog.askopenfilename(title="売上Excelを選択",
                                         filetypes=[("Excel", "*.xlsx"), ("すべて", "*.*")])
        if not path:
            return
        if not messagebox.askyesno("確認", "Excelの全行を取り込みます。既存DBに追加されます。実行しますか？"):
            return
        try:
            from import_excel import import_excel
            imp, skp = import_excel(Path(path))
        except Exception as e:
            messagebox.showerror("エラー", f"取込失敗: {e}")
            return
        messagebox.showinfo("完了", f"取込: {imp} 行 / スキップ: {skp} 行")
        self.app.refresh_all()

    def _export_all(self):
        dest = export_all_units()
        messagebox.showinfo("CSV書き出し", f"書き出しました:\n{dest}")
        open_path(dest.parent)

    def _export_stock(self):
        dest = export_stock()
        messagebox.showinfo("CSV書き出し", f"書き出しました:\n{dest}")
        open_path(dest.parent)

    def _backup(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = filedialog.asksaveasfilename(initialfile=f"sales_backup_{ts}.db",
                                            defaultextension=".db")
        if not dest:
            return
        shutil.copy2(DB_PATH, dest)
        messagebox.showinfo("バックアップ", f"保存しました:\n{dest}")


# ------------------------------------------------------------------
# Main app
# ------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("1180x780")
        self._setup_style()

        init_db()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.incoming = IncomingTab(nb, self)
        self.orders = OrderTab(nb, self)
        self.outgoing = OutgoingTab(nb, self)
        self.list_tab = ListTab(nb, self)
        self.stock_tab = StockTab(nb, self)
        self.detail = DetailTab(nb, self)
        self.tools = ToolsTab(nb, self)

        nb.add(self.incoming, text="📥 入荷登録")
        nb.add(self.orders, text="📝 受注")
        nb.add(self.outgoing, text="📤 出荷登録")
        nb.add(self.list_tab, text="📋 一覧/検索")
        nb.add(self.stock_tab, text="📦 在庫")
        nb.add(self.detail, text="🖼 詳細/画像")
        nb.add(self.tools, text="⚙ ツール")

        self.nb = nb

    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista" if sys.platform.startswith("win") else "clam")
        except tk.TclError:
            pass
        style.configure("TLabel", font=("Yu Gothic UI", 10))
        style.configure("TButton", font=("Yu Gothic UI", 10))
        style.configure("TEntry", font=("Yu Gothic UI", 10))
        style.configure("Treeview", font=("Yu Gothic UI", 10), rowheight=24)
        style.configure("Treeview.Heading", font=("Yu Gothic UI", 10, "bold"))

    def refresh_all(self):
        self.outgoing.refresh()
        self.list_tab.refresh()
        self.stock_tab.refresh()
        self.orders.refresh()

    def show_detail(self, unit_id: int):
        self.nb.select(self.detail)
        self.detail.load(unit_id)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
