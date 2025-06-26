import json
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import mysql.connector

# Columns to hide in the UI table
EXCLUDED_COLUMNS = [
    "product sales tax",
    "shipping credits",
    "shipping credits tax",
    "gift wrap credits",
    "giftwrap credits tax",
    "Regulatory Fee",
    "Tax On Regulatory Fee",
    "promotional rebates",
    "promotional rebates tax",
    "marketplace withheld tax",
    "selling fees",
    "fba fees",
    "other transaction fees",
    "other",
    "marketplace",
    "account type",
    "fulfilment order city",
    "order state",
    "order postal",
    "tax collection model",
]

class DataManager:
    """Load Excel files and compute monthly profit margin."""

    def __init__(self):
        self.campingcar_df = None
        self.grill_df = None
        self.combined_df = pd.DataFrame()
        self.exchange_rate = 1.0
        self.config = self._load_config()
        self.conn = self._connect_db()

    def _load_config(self):
        path = os.path.join(os.path.dirname(__file__), 'config.json')
        if not os.path.exists(path):
            return {}
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _connect_db(self):
        cfg = self.config.get('db_internal') or {}
        try:
            return mysql.connector.connect(
                host=cfg.get('host'),
                port=cfg.get('port'),
                user=cfg.get('user'),
                password=cfg.get('password'),
                database=cfg.get('database'),
            )
        except Exception:
            return None

    def _fetch_product_info(self, sku):
        if not self.conn or not sku:
            return None
        prefix = str(sku)[:12]
        try:
            cur = self.conn.cursor(dictionary=True)
            query = (
                "SELECT category, brand, `공급가(W)` AS supply, "
                "`배송비(W)` AS shipping FROM products "
                "WHERE `upc/ean`=%s LIMIT 1"
            )
            cur.execute(query, (prefix,))
            row = cur.fetchone()
            cur.close()
            return row
        except Exception:
            return None

    def load_excel(self, file_type, header_row=7):
        file_path = filedialog.askopenfilename(
            title=f"{file_type} 파일 선택",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            df_raw = pd.read_excel(file_path, header=None)
            if len(df_raw) <= header_row:
                messagebox.showerror("오류", f"파일에 {header_row + 1}행 이상 데이터가 없습니다.")
                return
            df_raw.columns = df_raw.iloc[header_row]
            df = df_raw[header_row + 1:].reset_index(drop=True)

            if file_type == "캠핑카":
                self.campingcar_df = self._add_profit(df)
            else:
                self.grill_df = self._add_profit(df)

            self._combine()
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _add_profit(self, df):
        required = {"날짜", "매출", "비용"}
        if required.issubset(df.columns):
            df["월"] = pd.to_datetime(df["날짜"]).dt.to_period("M")
            monthly = df.groupby("월").agg({"매출": "sum", "비용": "sum"})
            monthly["이익률"] = (monthly["매출"] - monthly["비용"]) / monthly["매출"] * 100
            df = df.merge(monthly["이익률"], on="월", how="left")
            df["이익금"] = (df["매출"] - df["비용"]) * self.exchange_rate
        else:
            if "이익금" not in df.columns:
                df["이익금"] = ""
        for col in ["Category", "Brand", "공급가(W)", "배송비(W)"]:
            if col not in df.columns:
                df[col] = ""

        sku_col = None
        for c in ["sku", "SKU", "Sku", "SKU#"]:
            if c in df.columns:
                sku_col = c
                break
        if sku_col:
            def enrich(row):
                if not row.get(sku_col):
                    return row
                info = self._fetch_product_info(row[sku_col])
                if info:
                    row["Category"] = info.get("category", row.get("Category", ""))
                    row["Brand"] = info.get("brand", row.get("Brand", ""))
                    row["공급가(W)"] = info.get("supply", row.get("공급가(W)", ""))
                    row["배송비(W)"] = info.get("shipping", row.get("배송비(W)", ""))
                return row
            df = df.apply(enrich, axis=1)
        if "total" in df.columns:
            idx = df.columns.get_loc("total") + 1
            for col in ["Category", "Brand", "공급가(W)", "배송비(W)", "이익금"][::-1]:
                series = df.pop(col)
                df.insert(idx, col, series)
        return df

    def set_exchange_rate(self, value):
        try:
            self.exchange_rate = float(value)
        except ValueError:
            messagebox.showerror("오류", "숫자를 입력하세요")
            return
        for df in [self.campingcar_df, self.grill_df]:
            if df is not None and {"매출", "비용"}.issubset(df.columns):
                df["이익금"] = (df["매출"] - df["비용"]) * self.exchange_rate
        self._combine()

    def _combine(self):
        frames = [df for df in [self.campingcar_df, self.grill_df] if df is not None]
        if frames:
            self.combined_df = pd.concat(frames, ignore_index=True)
        else:
            self.combined_df = pd.DataFrame()

    def run_actions(self):
        """Apply business rules based on the 'type' column."""
        if self.combined_df.empty:
            messagebox.showinfo("알림", "실행할 데이터가 없습니다.")
            return

        # Determine column names
        type_col = None
        for c in ["type", "Type", "TYPE"]:
            if c in self.combined_df.columns:
                type_col = c
                break
        if not type_col:
            messagebox.showerror("오류", "type 열이 없습니다.")
            return

        total_col = None
        for c in ["Total", "total"]:
            if c in self.combined_df.columns:
                total_col = c
                break
        if not total_col:
            messagebox.showerror("오류", "Total 열이 없습니다.")
            return

        df = self.combined_df.copy()
        # Drop rows where type is Transfer
        df = df[df[type_col] != "Transfer"]

        if "이익금" not in df.columns:
            df["이익금"] = ""

        mask = ~df[type_col].isin(["Order", "Refund"])
        totals = pd.to_numeric(df.loc[mask, total_col], errors="coerce").fillna(0)
        df.loc[mask, "이익금"] = totals * self.exchange_rate

        self.combined_df = df.reset_index(drop=True)

    def export(self):
        if self.combined_df.empty:
            messagebox.showinfo("알림", "저장할 데이터가 없습니다.")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if not file_path:
            return
        try:
            if file_path.endswith(".csv"):
                self.combined_df.to_csv(file_path, index=False)
            else:
                self.combined_df.to_excel(file_path, index=False)
            messagebox.showinfo("완료", "파일 저장 성공")
        except Exception as e:
            messagebox.showerror("오류", str(e))


def treeview_sort_column(tv, col, reverse):
    data = [(tv.set(k, col), k) for k in tv.get_children("")]
    try:
        data.sort(key=lambda t: float(str(t[0]).replace(",", "")), reverse=reverse)
    except ValueError:
        data.sort(key=lambda t: t[0], reverse=reverse)
    for index, (_, k) in enumerate(data):
        tv.move(k, "", index)
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))


def update_tree(tree, df):
    for item in tree.get_children():
        tree.delete(item)
    if df.empty:
        return
    # Drop columns that should not be shown in the UI
    display_df = df.drop(columns=[c for c in EXCLUDED_COLUMNS if c in df.columns])
    tree['columns'] = list(display_df.columns)
    tree['show'] = 'headings'
    for col in display_df.columns:
        tree.heading(col, text=col, command=lambda c=col: treeview_sort_column(tree, c, False))
        tree.column(col, width=100, anchor='center')
    for _, row in display_df.iterrows():
        tree.insert('', tk.END, values=list(row))


def main():
    manager = DataManager()
    root = tk.Tk()
    root.title('월별 이익률 계산 프로그램')
    root.geometry('1200x600')

    title = tk.Label(root, text='월별 이익률 계산 프로그램', font=('Arial', 14))
    title.pack(pady=5)

    control_frame = tk.Frame(root)
    control_frame.pack(pady=5)

    btn_camping = tk.Button(control_frame, text='캠핑카 파일', command=lambda: [manager.load_excel('캠핑카'), update_tree(tree, manager.combined_df)])
    btn_camping.pack(side=tk.LEFT, padx=5)

    btn_grill = tk.Button(control_frame, text='고기판 파일', command=lambda: [manager.load_excel('고기판'), update_tree(tree, manager.combined_df)])
    btn_grill.pack(side=tk.LEFT, padx=5)

    tk.Label(control_frame, text='환율').pack(side=tk.LEFT)
    rate_var = tk.StringVar(value='1.0')
    rate_entry = tk.Entry(control_frame, textvariable=rate_var, width=10)
    rate_entry.pack(side=tk.LEFT, padx=5)
    tk.Button(control_frame, text='적용', command=lambda: [manager.set_exchange_rate(rate_var.get()), update_tree(tree, manager.combined_df)]).pack(side=tk.LEFT, padx=5)

    run_btn = tk.Button(control_frame, text='실행', command=lambda: [manager.run_actions(), update_tree(tree, manager.combined_df)])
    run_btn.pack(side=tk.LEFT, padx=5)

    tree_frame = tk.Frame(root)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    x_scroll = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
    x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
    y_scroll = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
    y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    tree = ttk.Treeview(tree_frame, xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)
    tree.pack(fill=tk.BOTH, expand=True)

    x_scroll.config(command=tree.xview)
    y_scroll.config(command=tree.yview)

    export_btn = tk.Button(root, text='결과 저장', command=lambda: manager.export())
    export_btn.pack(pady=5)

    root.mainloop()


if __name__ == '__main__':
    main()
