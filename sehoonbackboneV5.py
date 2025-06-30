import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import socket
try:
    import pymysql
except ImportError as e:
    raise SystemExit(
        "pymysql \ud328\ud0a4\uc9c0\uac00 \uc124\uce58\ub418\uc5b4 \uc788\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4. 'pip install -r requirements.txt'\ub85c \uc124\uce58\ud558\uc138\uc694."
    ) from e
try:
    import pandas as pd
except ImportError as e:
    raise SystemExit(
        "pandas \ud328\ud0a4\uc9c0\uac00 \uc124\uce58\ub418\uc5b4 \uc788\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4. 'pip install -r requirements.txt'\ub85c \uc124\uce58\ud558\uc138\uc694."
    ) from e

try:
    import requests
except ImportError as e:
    raise SystemExit(
        "requests \ud328\ud0a4\uc9c0\uac00 \uc124\uce58\ub418\uc5b4 \uc788\uc9c0 \uc54a\uc2b5\ub2c8\ub2e4. 'pip install -r requirements.txt'\ub85c \uc124\uce58\ud558\uc138\uc694."
    ) from e

CONFIG_FILE = 'config.json'

API_URL = "https://api.exchangerate.host/latest?base=USD&symbols=KRW"

# Load configuration and business rate
if os.path.exists(CONFIG_FILE):
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
else:
    cfg = {}

INT_HOST = cfg.get('db_internal', {}).get('host', '127.0.0.1')
INT_PORT = cfg.get('db_internal', {}).get('port', 3306)
EXT_HOST = cfg.get('db_external', {}).get('host', '127.0.0.1')
EXT_PORT = cfg.get('db_external', {}).get('port', 3306)
business_rate = cfg.get('business_rate', 0.0)

def save_business_rate(value):
    cfg['business_rate'] = value
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False)

# Determine accessible DB host/port
def can_connect(host, port, timeout=0.5):
    s = socket.socket()
    s.settimeout(timeout)
    try:
        s.connect((host, port))
        return True
    except Exception:
        return False
    finally:
        s.close()

if can_connect(INT_HOST, INT_PORT):
    DB_HOST, DB_PORT = INT_HOST, INT_PORT
else:
    DB_HOST, DB_PORT = EXT_HOST, EXT_PORT

# Setup MariaDB connection
conn = pymysql.connect(
    host=DB_HOST,
    port=DB_PORT,
    user='sehoonyoloapi',
    password='ZEEMyPassStron25@',
    database='myappdb',
    charset='utf8mb4',
    cursorclass=pymysql.cursors.DictCursor,
)
cur = conn.cursor()
pcur = conn.cursor()

# Setup shipping table
cur.execute(
    """
    CREATE TABLE IF NOT EXISTS shipping (
        weight INT PRIMARY KEY,
        cost INT
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
)
conn.commit()

# default shipping data if empty
cur.execute('SELECT COUNT(*) AS cnt FROM shipping')
if cur.fetchone()['cnt'] == 0:
    data = [
        (0, 5981),
        (100, 5981),
        (200, 7528),
        (300, 9088),
        (400, 10665),
        (500, 12258),
        (600, 13791),
        (700, 15285),
        (800, 16771),
        (900, 18365),
        (1000, 19792),
        (1100, 20722),
        (1200, 22222),
        (1300, 23715),
        (1400, 25200),
        (1500, 26707),
        (1600, 28138),
        (1700, 29576),
        (1800, 31008),
        (1900, 32446),
        (2000, 33862),
    ]
    cur.executemany('INSERT INTO shipping(weight, cost) VALUES (%s, %s)', data)
    conn.commit()

# Setup products table
pcur.execute(
    """
    CREATE TABLE IF NOT EXISTS products (
        id INT AUTO_INCREMENT PRIMARY KEY,
        type VARCHAR(100),
        upc VARCHAR(50),
        en_name TEXT,
        ko_name TEXT,
        item TEXT,
        hs_code VARCHAR(50),
        supply INT,
        weight INT,
        ship FLOAT,
        finance FLOAT,
        cost_sum INT,
        sell_price FLOAT,
        input_val FLOAT,
        profit INT,
        margin FLOAT,
        category VARCHAR(100),
        brand VARCHAR(100)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
)
conn.commit()

def get_shipping_cost(weight):
    cur = conn.cursor()
    cur.execute('SELECT cost FROM shipping WHERE weight <= %s ORDER BY weight DESC LIMIT 1', (weight,))
    row = cur.fetchone()
    return row['cost'] if row else 0

# HS table
HS_TABLE = {
    'Blanket': '9404900000',
    'Car Accessories': '8708290000',
    'CDs': '8523491020',
    'Cleansing Foam': '3307109000',
    'Foundation': '3304992000',
    'Kitchen Stainless': '7323940000',
    'Laver': '2008995010',
    'Lipstick': '3304101000',
    'Makeup Tool': '3304999000',
    'Mask Pack': '3307909000',
    'Mixed Coffee': '2101111000',
    'Moisture Cream': '3304991000',
    'Other': '8215999000',
    'Other-METAL': '7323990000',
    'Painting Palette': '3926109000',
    'Paper Filter': '4823200000',
    'Pastel': '9609901000',
    'Pencil sharpener': '8472904000',
    'Photo Cards': '4902109000',
    'Plastic Product': '3926909000',
    'Roasted Tea': '2101301000',
    'Roller Pen': '9608100000',
    'Rubber Gloves': '4015190000',
    'Stationary': '4820200000',
}
CATEGORIES = sorted([
    'Stationery',
    'Camping',
    'General',
    'Cosmetics',
    'Automobile',
    'Apparel',
    'Coffee & Tea',
    'Food',
    'Music',
    'Painting',
    'Cards',
    'Toy',
    'Stainless Steel',
    'laver',
    'Electronics',
])

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('\uae30\uc900\uc815\ubcf4 \uad00\ub9ac \uc2dc\uc2a4\ud15c')
        self.configure(bg='#e6ffe6')
        self.geometry('1200x700')
        # remember sort direction for each column (False means ascending)
        self.sort_states = {}
        self.create_widgets()
        self.refresh_rate()
        self.load_data()

    def create_widgets(self):
        title_frame = tk.Frame(self, bg='#e6ffe6')
        title_frame.pack(fill='x')
        try:
            img = tk.PhotoImage(file='logo.png')
            w, h = img.width(), img.height()
            max_size = 40
            factor = max(w // max_size, h // max_size, 1)
            if factor > 1:
                img = img.subsample(factor, factor)
            self.logo_img = img
            tk.Label(
                title_frame,
                image=self.logo_img,
                bg="#e6ffe6",
            ).pack(side="left", padx=5)
        except Exception:
            self.logo_img = None
        tk.Label(
            title_frame,
            text="\uae30\uc900\uc815\ubcf4 \uad00\ub9ac \uc2dc\uc2a4\ud15c",
            bg="#e6ffe6",
            font=("Arial", 16, "bold"),
        ).pack(side="left")

        top_frame = tk.Frame(self, bg='#e6ffe6')
        top_frame.pack(fill='x')

        btn_font = ('Arial', 10)
        label_font = ('Arial', 10)

        tk.Button(top_frame, text='\uc5d1\uc140 \uc785\ub825', command=self.import_excel, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(top_frame, text='\ub2e4\uc6b4\ub85c\ub4dc', command=self.export_excel, font=btn_font).pack(side='left', padx=5, pady=5)

        tk.Label(top_frame, text='\ud604\uc7ac\ud658\uc728(USD\u2192KRW):', bg='#e6ffe6', font=label_font).pack(side='left')
        self.rate_var = tk.StringVar(value='0')
        tk.Label(top_frame, textvariable=self.rate_var, bg='#e6ffe6', font=label_font).pack(side='left')

        tk.Label(top_frame, text='\uacbd\uc601\ud658\uc728:', bg='#e6ffe6', font=label_font).pack(side='left', padx=(10,0))
        self.brate_var = tk.DoubleVar(value=business_rate)
        tk.Entry(top_frame, textvariable=self.brate_var, width=10, font=label_font).pack(side='left')
        tk.Button(top_frame, text='\uc800\uc7a5', command=self.save_rate, font=btn_font).pack(side='left', padx=5)

        search_frame = tk.Frame(self, bg='#e6ffe6')
        search_frame.pack(fill='x')
        tk.Label(search_frame, text='\uac80\uc0c9:', bg='#e6ffe6', font=label_font).pack(side='left')
        self.search_var = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.search_var, font=label_font, width=40).pack(side='left')
        tk.Button(search_frame, text='\uac80\uc0c9', command=self.search, font=btn_font).pack(side='left', padx=5)

        btn_frame = tk.Frame(self, bg='#e6ffe6')
        btn_frame.pack(fill='x')
        tk.Button(btn_frame, text='\uc804\uccb4 \uc120\ud0dd', command=self.select_all, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(btn_frame, text='\uc804\uccb4 \ud574\uc81c', command=self.deselect_all, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(btn_frame, text='\ud589 \ucd94\uac00', command=self.add_row, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(btn_frame, text='\uc0ad\uc81c', command=self.delete_rows, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(btn_frame, text='\uce74\ud14c\uace0\ub9ac/\ube0c\ub79c\ub4dc \uc218\uc815', command=self.bulk_edit_brand_category, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(btn_frame, text='\ud310\uac00\uc218\uc815', command=self.bulk_edit_sell_price, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(btn_frame, text='\uc0c8\ub85c\uace0\uce68', command=self.refresh_from_db, font=btn_font).pack(side='left', padx=5, pady=5)
        tk.Button(btn_frame, text='\ub9e8 \uc704', command=self.go_first, font=btn_font).pack(side='right', padx=5, pady=5)
        tk.Button(btn_frame, text='\ub9e8 \uc544\ub798', command=self.go_last, font=btn_font).pack(side='right', padx=5, pady=5)


        columns = [
            'check','type','upc','en_name','ko_name','item','hs_code','supply','weight','ship','finance','cost_sum','sell_price','input_val','profit','margin','category','brand'
        ]
        tree_frame = tk.Frame(self, bg='#e6ffe6')
        tree_frame.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', selectmode='extended')
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        headings = [
            '\uc120\ud0dd','\uad6c\ubd84','UPC/EAN','\uc601\uc5b4 \uc81c\ud488\uba85','\ud55c\uae00 \uc81c\ud488\uba85','\ud488\ubaa9\uba85','HS CODE','\uacf5\uae09\uac00(W)','\ubb34\uac8c(g)','\ubc30\uc1a1\ube44(W)','\uae08\uc735($)','\uc6d0\uac00\ud569(W)','\ud310\uac00($)','\uc785\ub825\uac12($)','\uc774\uc775\uae08(W)','\uc774\uc775\ub960(%)','Category','Brand'
        ]
        width_map = {
            'check': 40,
            'type': 64,
            'upc': 96,
            'en_name': 250,
            'ko_name': 250,
            'weight': 64,
            'ship': 64,
            'finance': 64,
            'sell_price': 64,
            'margin': 64,
        }
        anchor_map = {
            'en_name': 'w',
            'ko_name': 'w',
        }
        for col, hd in zip(columns, headings):
            self.tree.heading(col, text=hd,
                             command=lambda c=col: self.sort_by_column(c))
            w = width_map.get(col, 80)
            anchor = anchor_map.get(col, 'center')
            self.tree.column(col, width=w, anchor=anchor)

        style = ttk.Style(self.tree)
        style.configure('Treeview', rowheight=30)
        # alternate row colors for readability
        self.tree.tag_configure('even', background='#ffffff')
        self.tree.tag_configure('odd', background='#f5f5f5')
        style.map('Treeview')
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        self.tree.bind('<Double-1>', self.on_double_click)
        self.tree.bind('<Button-1>', self.on_check_click)
        self.tree.bind('<Button-3>', self.copy_upc)
        self.refresh_tags()


    def refresh_rate(self):
        try:
            resp = requests.get(API_URL, timeout=5)
            data = resp.json()
            rate = data.get('rates', {}).get('KRW')
            self.rate_var.set(f'{rate:.2f}')
        except Exception:
            self.rate_var.set('0')

    def save_rate(self):
        save_business_rate(self.brate_var.get())
        self.recompute_all()
        messagebox.showinfo('\uc800\uc7a5', '\uacbd\uc601\ud658\uc728\uc774 \uc800\uc7a5\ub418\uc5c8\uc2b5\ub2c8\ub2e4.')

    def get_checked_indices(self):
        return [iid for iid in self.tree.get_children() if self.tree.set(iid, 'check') == '\u2714']

    def delete_rows(self):
        for iid in self.get_checked_indices():
            self.tree.delete(iid)
            pcur.execute('DELETE FROM products WHERE id=%s', (int(iid),))
        conn.commit()
        self.refresh_tags()

    def select_all(self):
        for iid in self.tree.get_children():
            self.tree.set(iid, 'check', '\u2714')

    def deselect_all(self):
        for iid in self.tree.get_children():
            self.tree.set(iid, 'check', '\u25a1')

    def _sort(self, col, descending=False):
        def key_func(iid):
            val = self.tree.set(iid, col)
            try:
                return float(str(val).replace(',', ''))
            except ValueError:
                return val

        items = list(self.tree.get_children())
        items.sort(key=key_func, reverse=descending)
        for idx, iid in enumerate(items):
            self.tree.move(iid, '', idx)

    def sort_by_column(self, col):
        descending = self.sort_states.get(col, False)
        self._sort(col, descending)
        self.sort_states[col] = not descending

    def go_first(self):
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.see(children[0])

    def go_last(self):
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[-1])
            self.tree.see(children[-1])

    def toggle_type_value(self, current):
        """Return the toggled FBA/FBM value based on the current string."""
        if current.startswith('FBA-'):
            suffix = current[4:]
            if suffix == '\uceac':
                return '\uceac\ud551\uce74'
            elif suffix == '\uace0':
                return '\uace0\uae30\ud310'
            else:
                return suffix or 'FBM'
        if current == 'FBA':
            return 'FBM'
        if current == 'FBM':
            return 'FBA'
        if current == '\uceac\ud551\uce74':
            return 'FBA-\uceac'
        if current == '\uace0\uae30\ud310':
            return 'FBA-\uace0'
        return 'FBA'

    def fba_toggle(self):
        """Toggle checked rows between FBA and FBM states."""
        checked = self.get_checked_indices()
        if not checked:
            return
        for iid in checked:
            current = self.tree.set(iid, 'type')
            new_type = self.toggle_type_value(current)
            self.tree.set(iid, 'type', new_type)
            pcur.execute('UPDATE products SET type=%s WHERE id=%s', (new_type, int(iid)))
        conn.commit()
        self.refresh_tags()

    def bulk_edit_brand_category(self):
        """Update Category and Brand for all checked rows."""
        checked = self.get_checked_indices()
        if not checked:
            messagebox.showinfo('\uc54c\ub9bc', '\uc218\uc815\ud560 \ud56d\ubaa9\uc744 \uc120\ud0dd\ud558\uc138\uc694.')
            return

        w = tk.Toplevel(self)
        w.title('\uc77c\uacb0 \uc218\uc815')
        w.geometry('300x150')
        w.configure(bg='#e6ffe6')

        tk.Label(w, text='Category', bg='#e6ffe6').grid(row=0, column=0, padx=5, pady=5, sticky='e')
        cat_combo = ttk.Combobox(w, values=CATEGORIES)
        cat_combo.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(w, text='Brand', bg='#e6ffe6').grid(row=1, column=0, padx=5, pady=5, sticky='e')
        brand_entry = tk.Entry(w)
        brand_entry.grid(row=1, column=1, padx=5, pady=5)

        def apply_changes():
            category = cat_combo.get()
            brand = brand_entry.get()
            for iid in checked:
                current_category = category or self.tree.set(iid, 'category')
                current_brand = brand or self.tree.set(iid, 'brand')
                self.tree.set(iid, 'category', current_category)
                self.tree.set(iid, 'brand', current_brand[:10])
                pcur.execute(
                    'UPDATE products SET category=%s, brand=%s WHERE id=%s',
                    (current_category, current_brand, int(iid))
                )
            conn.commit()
            self.refresh_tags()
            w.destroy()

        tk.Button(w, text='\uc801\uc6a9', command=apply_changes).grid(row=2, column=0, padx=5, pady=10)
        tk.Button(w, text='\ucde8\uc18c', command=w.destroy).grid(row=2, column=1, padx=5, pady=10)

    def bulk_edit_sell_price(self):
        """Update sell price for all checked rows."""
        checked = self.get_checked_indices()
        if not checked:
            messagebox.showinfo('\uc54c\ub9bc', '\uc218\uc815\ud560 \ud56d\ubaa9\uc744 \uc120\ud0dd\ud558\uc138\uc694.')
            return

        w = tk.Toplevel(self)
        w.title('\ud310\uac00\uc218\uc815')
        w.geometry('250x120')
        w.configure(bg='#e6ffe6')

        tk.Label(w, text='\ud310\uac00($)', bg='#e6ffe6').grid(row=0, column=0, padx=5, pady=10, sticky='e')
        price_entry = tk.Entry(w)
        price_entry.grid(row=0, column=1, padx=5, pady=10)

        def apply_price():
            try:
                price = float(price_entry.get())
            except ValueError:
                messagebox.showerror('\uc624\ub958', '\uc62c\ubc14\ub978 \uac12\uc744 \uc785\ub825\ud558\uc138\uc694.')
                return
            for iid in checked:
                vals = self.tree.item(iid, 'values')
                supply = int(float(str(vals[7]).replace(',', '')))
                weight = int(float(str(vals[8]).replace(',', '')))
                values = self.compute_values(
                    vals[1], vals[2], vals[3], vals[4], vals[5],
                    supply, weight, price, vals[16], vals[17]
                )
                values[0] = vals[0]
                self.tree.item(iid, values=values)
                hs, ship, finance, cost_sum, input_val, profit, margin = self.compute_numbers(
                    vals[5], supply, weight, price
                )
                pcur.execute(
                    'UPDATE products SET sell_price=%s, ship=%s, finance=%s, cost_sum=%s, input_val=%s, profit=%s, margin=%s WHERE id=%s',
                    (price, ship, finance, cost_sum, input_val, profit, margin, int(iid))
                )
            conn.commit()
            self.refresh_tags()
            w.destroy()

        tk.Button(w, text='\uc801\uc6a9', command=apply_price).grid(row=1, column=0, padx=5, pady=10)
        tk.Button(w, text='\ucde8\uc18c', command=w.destroy).grid(row=1, column=1, padx=5, pady=10)

    def fba_toggle_popup(self, entries):
        """Toggle the type field within the row editor window."""
        cur_val = entries['\uad6c\ubd84'].get()
        new_val = self.toggle_type_value(cur_val)
        entries['\uad6c\ubd84'].delete(0, tk.END)
        entries['\uad6c\ubd84'].insert(0, new_val)

    def add_row(self):
        self.open_row_window()

    def on_double_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        vals = self.tree.item(iid, 'values')
        data = {
            '\uad6c\ubd84': vals[1],
            'UPC/EAN': vals[2],
            '\uc601\uc5b4 \uc81c\ud488\uba85': vals[3],
            '\ud55c\uae00 \uc81c\ud488\uba85': vals[4],
            '\ud488\ubaa9\uba85': vals[5],
            '\uacf5\uae09\uac00(W)': vals[7],
            '\ubb34\uac8c(g)': vals[8],
            '\ud310\uac00($)': vals[12],
            '\uc785\ub825\uac12($)': vals[13],
            '\uc774\uc775\uae08(W)': vals[14],
            'Category': vals[16],
            'Brand': vals[17],
        }
        self.open_row_window(iid, data)

    def on_check_click(self, event):
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if col != '#1' or not iid:
            return
        val = self.tree.set(iid, 'check')
        new_val = '\u2714' if val == '\u25a1' else '\u25a1'
        self.tree.set(iid, 'check', new_val)

    def copy_upc(self, event):
        """Copy the UPC/EAN value of the clicked row to the clipboard."""
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        upc = self.tree.set(iid, 'upc')
        self.clipboard_clear()
        self.clipboard_append(upc)
        messagebox.showinfo('\ubcf5\uc0ac', 'UPC/EAN\uc774 \ubcf5\uc0ac\ub418\uc5c8\uc2b5\ub2c8\ub2e4.')


    def open_row_window(self, iid=None, values=None):
        w = tk.Toplevel(self)
        w.title('\ud589 \uc785\ub825')
        w.geometry('560x500')
        w.configure(bg='#e6ffe6')

        # allow the entry column to expand
        w.grid_columnconfigure(1, weight=1)

        entries = {}
        int_vcmd = (w.register(lambda P: P.isdigit() or P == ''), '%P')
        lbl_font = ('Arial', 10)
        btn_font = ('Arial', 10)
        labels = [
            '\uad6c\ubd84','UPC/EAN','\uc601\uc5b4 \uc81c\ud488\uba85','\ud55c\uae00 \uc81c\ud488\uba85','\ud488\ubaa9\uba85',
            '\uacf5\uae09\uac00(W)','\ubb34\uac8c(g)','\ud310\uac00($)','\uc785\ub825\uac12($)','\uae08\uc735($)','\uc774\uc775\uae08(W)','\uc774\uc775\ub960(%)','Category','Brand'
        ]

        for i, lbl in enumerate(labels):
            tk.Label(w, text=lbl, font=lbl_font, bg='#e6ffe6').grid(
                row=i, column=0, sticky='e', padx=5, pady=2
            )

            if lbl == '\ud488\ubaa9\uba85':
                widget = ttk.Combobox(w, values=list(HS_TABLE.keys()), font=lbl_font)
            elif lbl == 'Category':
                widget = ttk.Combobox(w, values=CATEGORIES, font=lbl_font)
            else:
                if lbl in ('\uc601\uc5b4 \uc81c\ud488\uba85', '\ud55c\uae00 \uc81c\ud488\uba85'):
                    widget = tk.Entry(w, font=lbl_font, width=60)
                elif lbl in ('\uc785\ub825\uac12($)', '\uc774\uc775\uae08(W)', '\uc774\uc775\ub960(%)'):
                    widget = tk.Entry(w, font=lbl_font, state='readonly')
                elif lbl == '\ubb34\uac8c(g)':
                    widget = tk.Entry(w, font=lbl_font, width=16, validate='key', validatecommand=int_vcmd)
                elif lbl == '\uacf5\uae09\uac00(W)':
                    widget = tk.Entry(w, font=lbl_font, width=20, validate='key', validatecommand=int_vcmd)
                elif lbl == '\ud310\uac00($)':
                    widget = tk.Entry(w, font=lbl_font, width=16)
                else:
                    widget = tk.Entry(w, font=lbl_font, width=20)

            widget.grid(row=i, column=1, sticky='ew', padx=5, pady=2)
            entries[lbl] = widget
        btn_row = len(labels)
        btn_frame = tk.Frame(w, bg='#e6ffe6')
        btn_frame.grid(row=btn_row, column=0, columnspan=3, pady=5)
        tk.Button(btn_frame, text='\uc800\uc7a5', command=lambda:self.save_row(iid, w, entries), font=btn_font, bg='#a0d8ef').pack(side='left', padx=5)
        if iid:
            tk.Button(btn_frame, text='\uc0ad\uc81c', command=lambda:self.delete_row_in_window(iid, w), font=btn_font, bg='#ffbbbb').pack(side='left', padx=5)
        tk.Button(btn_frame, text='\ucde8\uc18c', command=w.destroy, font=btn_font, bg='#cccccc').pack(side='left', padx=5)

        if values:
            for k, v in values.items():
                if k in entries and k not in ('\uc785\ub825\uac12($)', '\uc774\uc775\uae08(W)', '\uc774\uc775\ub960(%)'):
                    if k in ('\uacf5\uae09\uac00(W)', '\ubc30\uc1a1\ube44(W)', '\uc6d0\uac00\ud569(W)', '\uc774\uc775\uae08(W)'):
                        try:
                            v = str(int(float(str(v).replace(',', ''))))
                        except ValueError:
                            v = ''
                    entries[k].insert(0, v)
        self._update_input_field(entries)

        entries['\ud310\uac00($)'].bind('<KeyRelease>', lambda e: self._update_input_field(entries))
        entries['\ubb34\uac8c(g)'].bind('<KeyRelease>', lambda e: self._update_input_field(entries))

    def save_row(self, iid, window, entries):
        """Save or update a row and persist it in MariaDB."""
        try:
            type_ = entries['\uad6c\ubd84'].get()
            upc = entries['UPC/EAN'].get()
            en = entries['\uc601\uc5b4 \uc81c\ud488\uba85'].get()
            ko = entries['\ud55c\uae00 \uc81c\ud488\uba85'].get()
            item = entries['\ud488\ubaa9\uba85'].get()
            supply = int(entries['\uacf5\uae09\uac00(W)'].get().replace(',', '') or 0)
            weight = int(entries['\ubb34\uac8c(g)'].get() or 0)
            sell = float(entries['\ud310\uac00($)'].get() or 0)
            category = entries['Category'].get()
            brand = entries['Brand'].get()

            values = self.compute_values(type_, upc, en, ko, item, supply, weight, sell, category, brand)
            hs, ship, finance, cost_sum, input_val, profit, margin = self.compute_numbers(item, supply, weight, sell)

            if iid:
                row_id = int(iid)
                check = self.tree.set(iid, 'check')
                values[0] = check
                self.tree.item(iid, values=values)
                pcur.execute(
                    'UPDATE products SET type=%s, upc=%s, en_name=%s, ko_name=%s, item=%s, hs_code=%s, '
                    'supply=%s, weight=%s, ship=%s, finance=%s, cost_sum=%s, sell_price=%s, input_val=%s, '
                    'profit=%s, margin=%s, category=%s, brand=%s WHERE id=%s',
                    (
                        type_, upc, en, ko, item, hs,
                        supply, weight, ship, finance, cost_sum, sell,
                        input_val, profit, margin, category, brand, row_id
                    )
                )
            else:
                pcur.execute(
                    'INSERT INTO products (type, upc, en_name, ko_name, item, hs_code, supply, weight, ship, finance, cost_sum, sell_price, input_val, profit, margin, category, brand) '
                    'VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                    (
                        type_, upc, en, ko, item, hs, supply, weight, ship, finance,
                        cost_sum, sell, input_val, profit, margin, category, brand
                    )
                )
                row_id = pcur.lastrowid
                self.insert_tree_row(values, iid=row_id)

            conn.commit()
            self.refresh_tags()
            window.destroy()
        except Exception as e:
            messagebox.showerror('\uc800\uc7a5 \uc624\ub958', str(e))

    def delete_row_in_window(self, iid, window):
        self.tree.delete(iid)
        pcur.execute('DELETE FROM products WHERE id=%s', (int(iid),))
        conn.commit()
        self.refresh_tags()
        window.destroy()

    def _update_input_field(self, entries):
        def to_int(v):
            try:
                return int(float(str(v).replace(',', '')))
            except (ValueError, TypeError):
                return 0

        def to_float(v):
            try:
                return float(str(v).replace(',', ''))
            except (ValueError, TypeError):
                return 0.0

        weight = to_int(entries['\ubb34\uac8c(g)'].get())
        sell = to_float(entries['\ud310\uac00($)'].get())
        supply = to_int(entries['\uacf5\uae09\uac00(W)'].get())
        ship = get_shipping_cost(weight)
        brate = self.brate_var.get()
        input_val = sell - (ship / brate if brate else 0)
        ent = entries['\uc785\ub825\uac12($)']
        ent.config(state='normal')
        ent.delete(0, tk.END)
        ent.insert(0, f'{input_val:.2f}')
        ent.config(state='readonly')
        if '\uae08\uc735($)' in entries:
            finance = sell * 0.15 + 0.4
            cost_sum_val = supply + ship + finance * brate
            profit_val = sell * brate - cost_sum_val
            p_ent = entries['\uae08\uc735($)']
            p_ent.config(state='normal')
            p_ent.delete(0, tk.END)
            p_ent.insert(0, f"{int(round(profit_val)):,}")
            p_ent.config(state='readonly')
            if '\uc774\uc775\ub960(%)' in entries:
                margin = (profit_val / (sell * brate) * 100) if sell > 0 and brate else 0
                m_ent = entries['\uc774\uc775\ub960(%)']
                m_ent.config(state='normal')
                m_ent.delete(0, tk.END)
                m_ent.insert(0, f"{margin:.0f}")
                m_ent.config(state='readonly')

    def compute_numbers(self, item, supply, weight, sell):
        def to_int(v):
            try:
                return int(float(str(v).replace(',', '')))
            except (ValueError, TypeError):
                return 0

        supply = to_int(supply)
        weight = to_int(weight)
        hs = HS_TABLE.get(item, '')
        ship = get_shipping_cost(weight)
        finance = sell * 0.15 + 0.4
        brate = self.brate_var.get()
        cost_sum_val = supply + ship + finance * brate
        input_val = sell - (ship / brate if brate else 0)
        profit_val = sell * brate - cost_sum_val
        cost_sum = int(round(cost_sum_val))
        profit = int(round(profit_val))
        margin = (profit_val / (sell * brate) * 100) if sell > 0 and brate else 0
        return hs, ship, finance, cost_sum, input_val, profit, margin

    def compute_values(self, type_, upc, en, ko, item, supply, weight, sell, category, brand):
        def to_int(v):
            try:
                return int(float(str(v).replace(',', '')))
            except (ValueError, TypeError):
                return 0

        supply_i = to_int(supply)
        weight_i = to_int(weight)

        # handle None values gracefully
        en = en or ""
        ko = ko or ""
        brand = brand or ""
        category = category or ""

        hs, ship, finance, cost_sum, input_val, profit, margin = self.compute_numbers(
            item, supply_i, weight_i, sell
        )
        return [
            '\u25a1',
            type_,
            upc,
            en[:80],
            ko[:80],
            item,
            hs,
            f'{supply_i:,}',
            weight_i,
            f'{ship:,}',
            f'{finance:.2f}',
            f'{cost_sum:,}',
            f'{sell:.2f}',
            f'{input_val:.2f}',
            f'{profit:,}',
            f'{margin:.0f}',
            category,
            brand[:10],
        ]

    def insert_tree_row(self, values, iid=None):
        index = len(self.tree.get_children())
        tag = 'even' if index % 2 == 0 else 'odd'
        if iid is not None:
            self.tree.insert('', 'end', iid=str(iid), values=values, tags=(tag,))
        else:
            self.tree.insert('', 'end', values=values, tags=(tag,))

    def refresh_tags(self):
        for idx, iid in enumerate(self.tree.get_children()):
            tag = 'even' if idx % 2 == 0 else 'odd'
            self.tree.item(iid, tags=(tag,))

    def recompute_all(self):
        def to_int(v):
            try:
                return int(float(str(v).replace(',', '')))
            except (ValueError, TypeError):
                return 0

        def to_float(v):
            try:
                return float(str(v).replace(',', ''))
            except (ValueError, TypeError):
                return 0.0

        for iid in self.tree.get_children():
            vals = self.tree.item(iid, 'values')
            supply = to_int(vals[7])
            weight = to_int(vals[8])
            sell = to_float(vals[12])
            values = self.compute_values(
                vals[1], vals[2], vals[3], vals[4], vals[5],
                supply, weight, sell,
                vals[16], vals[17]
            )
            values[0] = vals[0]
            self.tree.item(iid, values=values)
            hs, ship, finance, cost_sum, input_val, profit, margin = self.compute_numbers(
                vals[5], supply, weight, sell
            )
            pcur.execute(
                'UPDATE products SET type=%s, upc=%s, en_name=%s, ko_name=%s, item=%s, hs_code=%s, '
                'supply=%s, weight=%s, ship=%s, finance=%s, cost_sum=%s, sell_price=%s, input_val=%s, '
                'profit=%s, margin=%s, category=%s, brand=%s WHERE id=%s',
                (
                    vals[1], vals[2], vals[3], vals[4], vals[5], hs,
                    supply, weight, ship, finance, cost_sum, sell,
                    input_val, profit, margin, vals[16], vals[17], int(iid)
                )
            )
        conn.commit()

    def refresh_from_db(self):
        """Reload all rows from the database."""
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        if hasattr(self, 'hidden_items'):
            self.hidden_items.clear()
        self.load_data()

    def load_data(self):
        pcur.execute(
            'SELECT id, type, upc, en_name, ko_name, item, supply, weight, sell_price, category, brand FROM products'
        )
        for row in pcur.fetchall():
            pid = row['id']
            type_ = row['type']
            upc = row['upc']
            en = row['en_name']
            ko = row['ko_name']
            item = row['item']
            supply = int(float(row['supply'])) if row['supply'] is not None else 0
            weight = int(float(row['weight'])) if row['weight'] is not None else 0
            sell_price = float(row['sell_price']) if row['sell_price'] is not None else 0.0
            category = row['category']
            brand = row['brand']

            values = self.compute_values(
                type_,
                upc,
                en,
                ko,
                item,
                supply,
                weight,
                sell_price,
                category,
                brand,
            )
            self.insert_tree_row(values, iid=pid)
        self.refresh_tags()

    def import_excel(self):
        file = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx *.xls')])
        if not file:
            return
        try:
            # Read all rows as strings to avoid truncation or type issues
            df = pd.read_excel(file, dtype=str)
        except Exception as e:
            messagebox.showerror('\uc624\ub958', str(e))
            return

        df.fillna('', inplace=True)

        col_map = {
            '\uad6c\ubd84': 'type',
            'UPC/EAN': 'upc',
            'UPC': 'upc',
            '\uc601\uc5b4 \uc81c\ud488\uba85': 'en_name',
            '\ud55c\uae00 \uc81c\ud488\uba85': 'ko_name',
            '\ud488\ubaa9\uba85': 'item',
            '\uacf5\uae09\uac00(W)': 'supply',
            '\ubb34\uac8c(g)': 'weight',
            '\ud310\uac00($)': 'sell_price',
            'Category': 'category',
            'Brand': 'brand',
        }

        cols = {}
        for k, v in col_map.items():
            if k in df.columns:
                cols[v] = k

        if 'upc' not in cols:
            messagebox.showerror('\uc624\ub958', 'UPC \ucee4\ub7fc\uc774 \uc5c6\uc2b5\ub2c8\ub2e4')
            return

        def to_int(v):
            try:
                return int(float(str(v).replace(',', '')))
            except (ValueError, TypeError):
                return 0

        def to_float(v):
            try:
                return float(str(v).replace(',', ''))
            except (ValueError, TypeError):
                return 0.0

        for _, row in df.iterrows():
            upc = str(row[cols['upc']]) if pd.notna(row[cols['upc']]) else ''
            if not upc:
                continue
            # skip duplicates already present in the table
            if any(self.tree.set(iid, 'upc') == upc for iid in self.tree.get_children()):
                continue
            data = {
                'type_': row[cols['type']] if 'type' in cols and pd.notna(row[cols['type']]) else '',
                'upc': upc,
                'en': row[cols['en_name']] if 'en_name' in cols and pd.notna(row[cols['en_name']]) else '',
                'ko': row[cols['ko_name']] if 'ko_name' in cols and pd.notna(row[cols['ko_name']]) else '',
                'item': row[cols['item']] if 'item' in cols and pd.notna(row[cols['item']]) else '',
                'supply': to_int(row[cols['supply']]) if 'supply' in cols else 0,
                'weight': to_int(row[cols['weight']]) if 'weight' in cols else 0,
                'sell_price': to_float(row[cols['sell_price']]) if 'sell_price' in cols else 0.0,
                'category': row[cols['category']] if 'category' in cols and pd.notna(row[cols['category']]) else '',
                'brand': row[cols['brand']] if 'brand' in cols and pd.notna(row[cols['brand']]) else '',
            }

            values = self.compute_values(data['type_'], data['upc'], data['en'], data['ko'],
                                         data['item'], data['supply'], data['weight'],
                                         data['sell_price'], data['category'], data['brand'])
            hs, ship, finance, cost_sum, input_val, profit, margin = self.compute_numbers(
                data['item'], data['supply'], data['weight'], data['sell_price']
            )
            pcur.execute(
                'INSERT INTO products (type, upc, en_name, ko_name, item, hs_code, supply, weight, ship, finance, cost_sum, sell_price, input_val, profit, margin, category, brand) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',
                (
                    data['type_'],
                    data['upc'],
                    data['en'],
                    data['ko'],
                    data['item'],
                    hs,
                    data['supply'],
                    data['weight'],
                    ship,
                    finance,
                    cost_sum,
                    data['sell_price'],
                    input_val,
                    profit,
                    margin,
                    data['category'],
                    data['brand'],
                )
            )
            conn.commit()
            new_id = pcur.lastrowid
            self.insert_tree_row(values, iid=new_id)
        self.refresh_tags()

    def search(self):
        """Filter rows by UPC/EAN or name."""
        keyword = self.search_var.get().strip().lower()

        # restore previously hidden items
        if hasattr(self, 'hidden_items'):
            for iid in self.hidden_items:
                self.tree.reattach(iid, '', 'end')
            self.hidden_items.clear()
        else:
            self.hidden_items = []

        if not keyword:
            return

        matches = []
        for item in list(self.tree.get_children()):
            vals = self.tree.item(item, 'values')
            upc = str(vals[2]).lower().replace('-', '')
            en = str(vals[3]).lower()
            ko = str(vals[4]).lower()
            if keyword in upc or keyword in en or keyword in ko:
                matches.append(item)
            else:
                self.tree.detach(item)
                self.hidden_items.append(item)

        if matches:
            self.tree.see(matches[0])
        messagebox.showinfo('\uac80\uc0c9 \uacb0\uacfc', f'{len(matches)}\uac1c \ud56d\ubaa9\uc774 \uac80\uc0c9\ub418\uc5c8\uc2b5\ub2c8\ub2e4.')

    def export_excel(self):
        file = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if not file:
            return
        cols = [self.tree.heading(c)['text'] for c in self.tree['columns']]
        data = [self.tree.item(i,'values') for i in self.tree.get_children()]
        df = pd.DataFrame(data, columns=cols)
        df.to_excel(file, index=False)
        messagebox.showinfo('\uc800\uc7a5', '\uc800\uc7a5\ub418\uc5c8\uc2b5\ub2c8\ub2e4')

if __name__ == '__main__':
    app = App()
    app.mainloop()