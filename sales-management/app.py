"""売上・在庫管理 - Tkinter GUI

タブ:
  1. 入荷登録   2. 受注   3. 一覧   4. ツール
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


UNIT_STATUSES = ["在庫", "資産", "部品取り", "予約済", "出荷済", "保留", "返品"]
STATUS_COLORS = {
    "在庫":   "#fff7b0",   # yellow
    "資産":   "#c8efc1",   # green
    "部品取り": "#cfe6f5",   # light blue
    "出荷済": "#eeeeee",   # gray
    "予約済": "#ffe2c2",
    "保留":   "#f3d4ff",
    "返品":   "#ffd6d6",
}
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
# Tab 2: 受注
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
        super().__init__(app)
        self.app = app
        self.order_id = order_id
        self.on_save = on_save
        self.title(f"受注 編集 #{order_id}" if order_id else "受注 新規登録")
        try:
            px = app.winfo_rootx()
            py = app.winfo_rooty()
        except Exception:
            px = py = 100
        self.geometry(f"680x640+{px + 60}+{py + 60}")
        self._build()
        if order_id:
            self._load(order_id)
        else:
            self.f_odate.set(today())
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)
        self.after(200, lambda: self.attributes("-topmost", False))

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
        super().__init__(app)
        self.app = app
        self.order = order
        self.on_done = on_done
        self.selected_unit_id = None
        self.title(f"出荷確定（受注 #{order['id']}）")
        try:
            px = app.winfo_rootx()
            py = app.winfo_rooty()
        except Exception:
            px = py = 100
        self.geometry(f"760x520+{px + 60}+{py + 60}")
        self._build()
        self._refresh_units()
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)
        self.after(200, lambda: self.attributes("-topmost", False))

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
                WHERE s.id IS NULL AND u.status = '在庫'
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
# Tab 3: 一覧（Master + 詳細編集ポップアップ）
# ------------------------------------------------------------------

class MasterTab(ttk.Frame):
    """全シリアルのマスター一覧。状態色付け、フィルタ、検索、ソート、
    ダブルクリックで詳細ポップアップ編集。"""

    def __init__(self, master, app):
        super().__init__(master, padding=12)
        self.app = app
        self._build()
        self.refresh()

    def _build(self):
        ttk.Label(self, text="個体一覧（全シリアル・状態色分け）",
                  font=("", 14, "bold")).pack(anchor="w", pady=(0, 8))

        top = ttk.Frame(self)
        top.pack(fill="x", pady=4)
        ttk.Button(top, text="更新", command=self.refresh).pack(side="left")

        ttk.Label(top, text="    状態:").pack(side="left", padx=(20, 4))
        self.status_var = tk.StringVar(value="(全部)")
        ttk.Combobox(top, textvariable=self.status_var,
                     values=["(全部)", "在庫", "資産", "部品取り", "出荷済",
                             "予約済", "保留", "返品"],
                     width=12, state="readonly").pack(side="left")
        self.status_var.trace_add("write", lambda *_: self.refresh())

        ttk.Label(top, text="    機種/シリアル/購入者:").pack(side="left", padx=(20, 4))
        self.q_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.q_var, width=30)
        ent.pack(side="left")
        ent.bind("<KeyRelease>", lambda e: self.refresh())

        ttk.Button(top, text="CSV書き出し", command=self._export).pack(side="right", padx=4)

        cols = ("status", "serial", "model", "mfg_date",
                "purchase_date", "vendor", "amount",
                "sale_date", "customer", "address", "total", "payment", "memo")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=22)
        for k, t, w, a in [
            ("status", "状態", 70, "w"),
            ("serial", "シリアル", 110, "w"),
            ("model", "機種", 180, "w"),
            ("mfg_date", "製造日", 90, "w"),
            ("purchase_date", "入荷日", 90, "w"),
            ("vendor", "仕入先", 160, "w"),
            ("amount", "入荷金額", 90, "e"),
            ("sale_date", "出荷日", 90, "w"),
            ("customer", "購入者", 140, "w"),
            ("address", "住所", 180, "w"),
            ("total", "売上金額", 90, "e"),
            ("payment", "入金", 60, "w"),
            ("memo", "メモ", 180, "w"),
        ]:
            self.tree.heading(k, text=t, command=lambda c=k: self._sort(c))
            self.tree.column(k, width=w, anchor=a)
        for st, color in STATUS_COLORS.items():
            self.tree.tag_configure(st, background=color)
        self.tree.pack(fill="both", expand=True, pady=4)
        self.tree.bind("<Double-1>", self._open_detail)

        sb = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", pady=4)
        self.summary_label = ttk.Label(bottom, text="", font=("", 11, "bold"))
        self.summary_label.pack(side="left")

        self._sort_col: str | None = None
        self._sort_asc = True

    def refresh(self):
        st = self.status_var.get() if hasattr(self, "status_var") else "(全部)"
        kw = self.q_var.get().strip() if hasattr(self, "q_var") else ""

        sql = """
            SELECT u.id, u.serial_no, u.model, u.mfg_date, u.status,
                   u.memo AS unit_memo,
                   p.purchase_date, p.vendor_name, p.amount,
                   s.sale_date, s.customer_name, s.address,
                   s.total_amount, s.payment_status
            FROM units u
            LEFT JOIN sales s ON s.unit_id = u.id
            LEFT JOIN purchases p ON p.unit_id = u.id
            WHERE 1=1
        """
        params: list = []
        if st != "(全部)":
            sql += " AND u.status = ?"
            params.append(st)
        if kw:
            sql += (" AND (u.model LIKE ? OR u.serial_no LIKE ?"
                    " OR s.customer_name LIKE ? OR s.address LIKE ?)")
            like = f"%{kw}%"
            params += [like, like, like, like]
        sql += " ORDER BY u.model, u.id"

        with connect() as conn:
            rows = list(conn.execute(sql, params))

        self.tree.delete(*self.tree.get_children())
        from collections import Counter
        counter: Counter = Counter()
        in_stock_amount = 0
        for r in rows:
            counter[r["status"]] += 1
            amt = r["amount"]
            tot = r["total_amount"]
            tag = (r["status"],) if r["status"] in STATUS_COLORS else ()
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["status"] or "",
                r["serial_no"] or "",
                r["model"] or "",
                r["mfg_date"] or "",
                r["purchase_date"] or "",
                r["vendor_name"] or "",
                f"{amt:,}" if amt is not None else "",
                r["sale_date"] or "",
                r["customer_name"] or "",
                r["address"] or "",
                f"{tot:,}" if tot is not None else "",
                r["payment_status"] or "",
                (r["unit_memo"] or "").replace("\n", " / "),
            ), tags=tag)
            if r["status"] == "在庫" and amt is not None:
                in_stock_amount += amt

        parts = [f"表示: {len(rows)} 台"]
        for s in ("在庫", "資産", "部品取り", "出荷済", "予約済", "保留", "返品"):
            if counter[s]:
                parts.append(f"{s}: {counter[s]}")
        parts.append(f"在庫の入荷金額計: {in_stock_amount:,} 円")
        self.summary_label.config(text="  /  ".join(parts))

    def _open_detail(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        DetailDialog(self, self.app, unit_id=int(sel[0]), on_save=self._after_save)

    def _after_save(self):
        self.refresh()
        # refresh sibling tabs that also depend on units/sales
        self.app.orders.refresh()

    def _sort(self, col):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]

        def keyfn(v):
            s = str(v).replace(",", "").strip()
            if not s:
                return (1, "")
            try:
                return (0, float(s))
            except ValueError:
                return (0, s)

        items.sort(key=lambda x: keyfn(x[0]), reverse=not self._sort_asc)
        for i, (_, k) in enumerate(items):
            self.tree.move(k, "", i)

    def _export(self):
        from export_csv import export_sales as _export_rows
        st = self.status_var.get()
        kw = self.q_var.get().strip()
        sql = """
            SELECT u.id AS unit_id, u.serial_no, u.model, u.mfg_date, u.status,
                   u.memo AS unit_memo,
                   p.purchase_date, p.vendor_name, p.vendor_company,
                   p.amount AS purchase_amount, p.invoice_no AS purchase_invoice,
                   s.sale_date, s.delivery_date, s.customer_name, s.customer_company,
                   s.postal, s.address, s.phone, s.email, s.yahoo_id,
                   s.sale_method, s.invoice_no, s.freight, s.total_amount,
                   s.payment_status, s.payment_date, s.memo AS sale_memo
            FROM units u
            LEFT JOIN sales s ON s.unit_id = u.id
            LEFT JOIN purchases p ON p.unit_id = u.id
            WHERE 1=1
        """
        params: list = []
        if st != "(全部)":
            sql += " AND u.status = ?"
            params.append(st)
        if kw:
            sql += (" AND (u.model LIKE ? OR u.serial_no LIKE ?"
                    " OR s.customer_name LIKE ? OR s.address LIKE ?)")
            like = f"%{kw}%"
            params += [like, like, like, like]
        sql += " ORDER BY u.model, u.id"
        with connect() as conn:
            rows = list(conn.execute(sql, params))
        dest = _export_rows(rows)
        messagebox.showinfo("CSV書き出し", f"書き出しました:\n{dest}")
        open_path(dest.parent)


class DetailDialog(tk.Toplevel):
    """個体の詳細編集ポップアップ。機体・入荷・出荷・添付の4セクション。"""

    def __init__(self, parent, app, unit_id: int, on_save=None):
        super().__init__(app)
        self.app = app
        self.unit_id = unit_id
        self.on_save = on_save
        self.title(f"個体詳細 #{unit_id}")
        try:
            px = app.winfo_rootx()
            py = app.winfo_rooty()
        except Exception:
            px = py = 100
        self.geometry(f"780x720+{px + 60}+{py + 60}")
        self._build()
        self._load()
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)
        self.after(200, lambda: self.attributes("-topmost", False))

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # 機体
        t1 = ttk.Frame(nb, padding=10)
        nb.add(t1, text="機体")
        self.f_model = LabeledEntry(t1, "機種 *")
        self.f_serial = LabeledEntry(t1, "シリアル")
        self.f_mfg = LabeledEntry(t1, "製造年月日")
        self.f_status = LabeledCombo(t1, "状態", UNIT_STATUSES)
        for w in (self.f_model, self.f_serial, self.f_mfg, self.f_status):
            w.pack(fill="x", pady=3)
        ttk.Label(t1, text="機体メモ:").pack(anchor="w", pady=(8, 2))
        self.f_memo = tk.Text(t1, height=8, wrap="word")
        self.f_memo.pack(fill="both", expand=True)

        # 入荷
        t2 = ttk.Frame(nb, padding=10)
        nb.add(t2, text="入荷")
        self.f_pdate = LabeledEntry(t2, "入荷日")
        self.f_vendor = LabeledEntry(t2, "仕入先", width=42)
        self.f_vcompany = LabeledEntry(t2, "仕入先会社")
        self.f_amount = LabeledEntry(t2, "仕入金額")
        self.f_pinvoice = LabeledEntry(t2, "請求書番号")
        for w in (self.f_pdate, self.f_vendor, self.f_vcompany, self.f_amount, self.f_pinvoice):
            w.pack(fill="x", pady=3)
        ttk.Label(t2, text="仕入メモ:").pack(anchor="w", pady=(8, 2))
        self.f_pmemo = tk.Text(t2, height=4, wrap="word")
        self.f_pmemo.pack(fill="x")

        # 出荷
        t3 = ttk.Frame(nb, padding=10)
        nb.add(t3, text="出荷")
        self.f_sdate = LabeledEntry(t3, "売上日")
        self.f_ddate = LabeledEntry(t3, "納品日")
        self.f_method = LabeledCombo(t3, "販売方法", SALE_METHODS)
        self.f_cust = LabeledEntry(t3, "お客様")
        self.f_company = LabeledEntry(t3, "会社/発注番号")
        self.f_postal = LabeledEntry(t3, "郵便番号")
        self.f_addr = LabeledEntry(t3, "住所", width=50)
        self.f_phone = LabeledEntry(t3, "電話")
        self.f_email = LabeledEntry(t3, "メール")
        self.f_sinvoice = LabeledEntry(t3, "請求書番号")
        self.f_freight = LabeledEntry(t3, "送料")
        self.f_total = LabeledEntry(t3, "売上金額")
        self.f_pay = LabeledCombo(t3, "入金状態", PAYMENT_STATUSES)
        self.f_paydate = LabeledEntry(t3, "入金日")
        for w in (self.f_sdate, self.f_ddate, self.f_method, self.f_cust,
                  self.f_company, self.f_postal, self.f_addr, self.f_phone,
                  self.f_email, self.f_sinvoice, self.f_freight,
                  self.f_total, self.f_pay, self.f_paydate):
            w.pack(fill="x", pady=2)
        ttk.Label(t3, text="売上メモ:").pack(anchor="w", pady=(6, 2))
        self.f_smemo = tk.Text(t3, height=3, wrap="word")
        self.f_smemo.pack(fill="x")

        # 添付
        t4 = ttk.Frame(nb, padding=10)
        nb.add(t4, text="添付ファイル")
        bar = ttk.Frame(t4); bar.pack(fill="x")
        ttk.Button(bar, text="画像/伝票を追加", command=self._add_attachment).pack(side="left")
        ttk.Button(bar, text="フォルダを開く", command=self._open_folder).pack(side="left", padx=6)
        self.att_list = tk.Listbox(t4, height=12)
        self.att_list.pack(fill="both", expand=True, pady=8)
        self.att_list.bind("<Double-1>", self._open_attachment)

        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=8, pady=6)
        ttk.Button(bottom, text="個体を削除", command=self._delete).pack(side="left")
        ttk.Button(bottom, text="閉じる", command=self.destroy).pack(side="right")
        ttk.Button(bottom, text="保存", command=self._save).pack(side="right", padx=6)

    def _set_text(self, widget, value):
        widget.delete("1.0", "end")
        if value:
            widget.insert("1.0", value)

    def _load(self):
        with connect() as conn:
            u = conn.execute("SELECT * FROM units WHERE id=?", (self.unit_id,)).fetchone()
            if not u:
                messagebox.showerror("エラー", "個体が見つかりません")
                self.destroy()
                return
            p = conn.execute("SELECT * FROM purchases WHERE unit_id=?", (self.unit_id,)).fetchone()
            s = conn.execute("SELECT * FROM sales WHERE unit_id=?", (self.unit_id,)).fetchone()
            atts = conn.execute(
                "SELECT * FROM attachments WHERE unit_id=? ORDER BY created_at DESC",
                (self.unit_id,)).fetchall()

        self.f_model.set(u["model"] or "")
        self.f_serial.set(u["serial_no"] or "")
        self.f_mfg.set(u["mfg_date"] or "")
        self.f_status.set(u["status"] or "在庫")
        self._set_text(self.f_memo, u["memo"] or "")

        if p:
            self.f_pdate.set(p["purchase_date"] or "")
            self.f_vendor.set(p["vendor_name"] or "")
            self.f_vcompany.set(p["vendor_company"] or "")
            self.f_amount.set(str(p["amount"]) if p["amount"] is not None else "")
            self.f_pinvoice.set(p["invoice_no"] or "")
            self._set_text(self.f_pmemo, p["memo"] or "")

        if s:
            self.f_sdate.set(s["sale_date"] or "")
            self.f_ddate.set(s["delivery_date"] or "")
            self.f_method.set(s["sale_method"] or "")
            self.f_cust.set(s["customer_name"] or "")
            self.f_company.set(s["customer_company"] or "")
            self.f_postal.set(s["postal"] or "")
            self.f_addr.set(s["address"] or "")
            self.f_phone.set(s["phone"] or "")
            self.f_email.set(s["email"] or "")
            self.f_sinvoice.set(s["invoice_no"] or "")
            self.f_freight.set(str(s["freight"]) if s["freight"] is not None else "")
            self.f_total.set(str(s["total_amount"]) if s["total_amount"] is not None else "")
            self.f_pay.set(s["payment_status"] or "")
            self.f_paydate.set(s["payment_date"] or "")
            self._set_text(self.f_smemo, s["memo"] or "")

        self.att_list.delete(0, "end")
        self._att_paths: list[str] = []
        for a in atts:
            self.att_list.insert("end", a["file_path"])
            self._att_paths.append(a["file_path"])

    def _save(self):
        model = self.f_model.get()
        if not model:
            messagebox.showwarning("入力エラー", "機種は必須です")
            return
        try:
            with connect() as conn:
                conn.execute(
                    """UPDATE units SET serial_no=?, model=?, mfg_date=?,
                       status=?, memo=?, updated_at=datetime('now','localtime')
                       WHERE id=?""",
                    (self.f_serial.get() or None, model, self.f_mfg.get() or None,
                     self.f_status.get() or "在庫",
                     self.f_memo.get("1.0", "end").strip() or None,
                     self.unit_id),
                )

                p_values = (
                    self.f_pdate.get() or None,
                    self.f_vendor.get() or None,
                    self.f_vcompany.get() or None,
                    parse_int(self.f_amount.get()),
                    self.f_pinvoice.get() or None,
                    self.f_pmemo.get("1.0", "end").strip() or None,
                )
                if any(v is not None for v in p_values):
                    p_exists = conn.execute(
                        "SELECT id FROM purchases WHERE unit_id=?", (self.unit_id,)
                    ).fetchone()
                    if p_exists:
                        conn.execute(
                            """UPDATE purchases SET purchase_date=?, vendor_name=?,
                               vendor_company=?, amount=?, invoice_no=?, memo=?
                               WHERE unit_id=?""",
                            p_values + (self.unit_id,),
                        )
                    else:
                        conn.execute(
                            """INSERT INTO purchases(unit_id, purchase_date, vendor_name,
                               vendor_company, amount, invoice_no, memo)
                               VALUES (?,?,?,?,?,?,?)""",
                            (self.unit_id,) + p_values,
                        )

                s_values = (
                    self.f_sdate.get() or None,
                    self.f_ddate.get() or None,
                    self.f_method.get() or None,
                    self.f_cust.get() or None,
                    self.f_company.get() or None,
                    self.f_postal.get() or None,
                    self.f_addr.get() or None,
                    self.f_phone.get() or None,
                    self.f_email.get() or None,
                    self.f_sinvoice.get() or None,
                    parse_int(self.f_freight.get()),
                    parse_int(self.f_total.get()),
                    self.f_pay.get() or None,
                    self.f_paydate.get() or None,
                    self.f_smemo.get("1.0", "end").strip() or None,
                )
                if any(v is not None for v in s_values):
                    s_exists = conn.execute(
                        "SELECT id FROM sales WHERE unit_id=?", (self.unit_id,)
                    ).fetchone()
                    if s_exists:
                        conn.execute(
                            """UPDATE sales SET sale_date=?, delivery_date=?, sale_method=?,
                               customer_name=?, customer_company=?, postal=?, address=?,
                               phone=?, email=?, invoice_no=?, freight=?, total_amount=?,
                               payment_status=?, payment_date=?, memo=?
                               WHERE unit_id=?""",
                            s_values + (self.unit_id,),
                        )
                    else:
                        conn.execute(
                            """INSERT INTO sales(unit_id, sale_date, delivery_date,
                               sale_method, customer_name, customer_company,
                               postal, address, phone, email, invoice_no,
                               freight, total_amount, payment_status, payment_date, memo)
                               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (self.unit_id,) + s_values,
                        )
        except Exception as e:
            messagebox.showerror("エラー", f"保存失敗: {e}")
            return
        if self.on_save:
            self.on_save()
        self.destroy()

    def _delete(self):
        if not messagebox.askyesno("確認", f"個体 #{self.unit_id} を削除しますか？"
                                            "（関連する入荷・出荷情報も消えます）"):
            return
        with connect() as conn:
            conn.execute("DELETE FROM units WHERE id=?", (self.unit_id,))
        if self.on_save:
            self.on_save()
        self.destroy()

    def _add_attachment(self):
        path = filedialog.askopenfilename(title="添付ファイル選択")
        if not path:
            return
        src = Path(path)
        dest_dir = ATTACHMENTS_DIR / str(self.unit_id)
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / src.name
        try:
            shutil.copy2(src, dest)
            with connect() as conn:
                conn.execute(
                    "INSERT INTO attachments(unit_id, file_path) VALUES (?,?)",
                    (self.unit_id, str(dest)),
                )
        except Exception as e:
            messagebox.showerror("エラー", f"添付に失敗: {e}")
            return
        self._load()

    def _open_attachment(self, _evt=None):
        sel = self.att_list.curselection()
        if not sel:
            return
        open_path(Path(self._att_paths[sel[0]]))

    def _open_folder(self):
        d = ATTACHMENTS_DIR / str(self.unit_id)
        d.mkdir(parents=True, exist_ok=True)
        open_path(d)


# ------------------------------------------------------------------
# Tab 4: ツール
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

        # Count existing units to detect re-import situation
        with connect() as conn:
            n_units = conn.execute("SELECT COUNT(*) FROM units").fetchone()[0]
            n_orders = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]

        clear_db = False
        if n_units > 0:
            ans = messagebox.askyesnocancel(
                "既存データの取扱い",
                f"既に {n_units} 件の個体データがあります（受注 {n_orders} 件）。\n\n"
                "【はい】 = 既存データを全削除してから取り込む（重複なし、推奨）\n"
                "【いいえ】 = 既存データに追加（重複の可能性あり）\n"
                "【キャンセル】 = 取込中止"
            )
            if ans is None:
                return
            clear_db = bool(ans)
        try:
            if clear_db:
                with connect() as conn:
                    # Order matters: child tables first, then units (FK cascades clean up)
                    conn.execute("DELETE FROM attachments")
                    conn.execute("DELETE FROM orders")
                    conn.execute("DELETE FROM sales")
                    conn.execute("DELETE FROM purchases")
                    conn.execute("DELETE FROM units")
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
        self.master = MasterTab(nb, self)
        self.tools = ToolsTab(nb, self)

        nb.add(self.incoming, text="📥 入荷登録")
        nb.add(self.orders, text="📝 受注")
        nb.add(self.master, text="📋 一覧")
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
        self.master.refresh()
        self.orders.refresh()

    def show_detail(self, unit_id: int):
        DetailDialog(self, self, unit_id=unit_id, on_save=self.refresh_all)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
