import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


# -------------------------
# Helpers: header detection
# -------------------------
def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def find_column(df: pd.DataFrame, candidates):
    """
    Find a column by checking if any candidate substring appears in the header.
    Returns column name or raises ValueError.
    """
    cols = list(df.columns)
    norm_cols = {c: normalize(c) for c in cols}
    cand_norm = [normalize(x) for x in candidates]

    for c in cols:
        h = norm_cols[c]
        if any(cn in h for cn in cand_norm):
            return c
    raise ValueError(f"Не намерих колона за: {candidates}. Налични колони: {cols}")


def detect_tier_columns(df_prices: pd.DataFrame):
    """
    Detect tier columns like 1000, 2000, 3000... from header names.
    Returns list of tuples: (tier_int, colname), sorted ascending.
    """
    tiers = []
    for c in df_prices.columns:
        s = normalize(c)
        digits = re.findall(r"\d+", s)
        if not digits:
            continue
        # if header is like "1000" or "за 1000 бр"
        tier = int(digits[0])
        if tier > 0:
            tiers.append((tier, c))
    tiers.sort(key=lambda x: x[0])
    return tiers


def to_int(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int,)):
        return int(x)
    try:
        s = str(x).strip().replace(",", ".")
        v = float(s)
        return int(round(v))
    except Exception:
        return None


def to_float(x):
    if pd.isna(x):
        return None
    if isinstance(x, (float, int)):
        return float(x)
    try:
        s = str(x).strip().replace(",", ".")
        return float(s)
    except Exception:
        return None


def resolve_unit_price(qty: int, tiers_map: dict):
    """
    tiers_map: {tier_int: price_float}
    pick the smallest tier >= qty; else pick the largest tier.
    """
    if not tiers_map:
        return None
    keys = sorted(tiers_map.keys())
    for k in keys:
        if k >= qty and tiers_map.get(k) is not None:
            return tiers_map[k]
    # fallback: last available
    for k in reversed(keys):
        if tiers_map.get(k) is not None:
            return tiers_map[k]
    return None


# -------------------------
# Excel reading
# -------------------------
def read_excel_any(path: str) -> pd.DataFrame:
    """
    Read .xls/.xlsx with pandas.
    .xls requires xlrd==2.0.1 installed.
    """
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".xls":
            return pd.read_excel(path, engine="xlrd")
        else:
            return pd.read_excel(path, engine="openpyxl")
    except ImportError as e:
        raise ImportError(
            "Липсва библиотека за четене. Инсталирай:\n"
            "pip install pandas openpyxl xlrd==2.0.1\n\n"
            f"Оригинална грешка: {e}"
        )
    except Exception as e:
        raise RuntimeError(f"Не успях да прочета файла: {path}\nГрешка: {e}")


def merge_order_and_prices(order_path: str, prices_path: str) -> pd.DataFrame:
    df_order = read_excel_any(order_path)
    df_prices = read_excel_any(prices_path)

    # Find order columns
    col_order_no = find_column(df_order, ["номер на поръчка", "поръчка", "Purchase Order"])
    col_item = find_column(df_order, ["име на артикул", "артикул", "продукт", "Item Number"])
    col_qty = find_column(df_order, ["заявени бройки", "бройки", "количество", "Quantity Ordered"])
    col_date = find_column(df_order, ["дата на доставка", "доставка", "delivery", "Due Date"])

    # Find prices columns
    p_item = find_column(df_prices, ["код АЛ филтър", "артикул", "item"])
    p_tl = None
    p_size = None
    p_mat = None
    try:
        p_tl = find_column(df_prices, ["технологичен лист", "ТЛ", "tech"])
    except Exception:
        p_tl = None
    try:
        p_size = find_column(df_prices, ["размер", "шир./вис."])
    except Exception:
        p_size = None
    try:
        p_mat = find_column(df_prices, ["материал", "material"])
    except Exception:
        p_mat = None

    tier_cols = detect_tier_columns(df_prices)  # [(1000,"1000"), (2000,"2000")...]

    # Build quick lookup for prices by item
    prices_lookup = {}
    for _, r in df_prices.iterrows():
        name = r.get(p_item)
        if pd.isna(name):
            continue
        name = str(name).strip()
        tiers_map = {}
        for tier, col in tier_cols:
            tiers_map[tier] = to_float(r.get(col))
        prices_lookup[name] = {
            "Технологичен лист": None if p_tl is None else (None if pd.isna(r.get(p_tl)) else str(r.get(p_tl)).strip()),
            "Размер": None if p_size is None else (None if pd.isna(r.get(p_size)) else str(r.get(p_size)).strip()),
            "Материал": None if p_mat is None else (None if pd.isna(r.get(p_mat)) else str(r.get(p_mat)).strip()),
            "tiers": tiers_map,
        }

    # Create merged rows
    out_rows = []
    line_no = 0

    for _, r in df_order.iterrows():
        order_no = r.get(col_order_no)
        item = r.get(col_item)
        qty = r.get(col_qty)
        ddate = r.get(col_date)

        if pd.isna(order_no) or pd.isna(item):
            continue

        order_no = str(order_no).strip()
        item = str(item).strip()
        qty_i = to_int(qty)
        if qty_i is None:
            continue

        line_no += 1
        order_ref = f"{order_no}-{line_no}"

        price_info = prices_lookup.get(item)
        unit_price = None
        size = None
        tl = None
        mat = None

        if price_info:
            unit_price = resolve_unit_price(qty_i, price_info.get("tiers") or {})
            size = price_info.get("Размер")
            tl = price_info.get("Технологичен лист")
            mat = price_info.get("Материал")

        total = unit_price * qty_i if unit_price is not None else None

        out_rows.append({
            "Артикул": item,
            "Размер": size or "",
            "Бройки": qty_i,
            "Ед. Цена": "" if unit_price is None else unit_price,
            "Сума": "" if total is None else total,
            "Номер на поръчка и ред": order_ref,
            "Дата на доставка": ddate,
            "Технологичен лист": tl or "",
            "Материал": mat or "",
        })

    out = pd.DataFrame(out_rows)
    return out
def excel_cell_to_string(x) -> str:
    """
    Връща стойността ТОЧНО като текст от Excel.
    Никакви дати, никакви формати.
    """
    if pd.isna(x):
        return ""
    return str(x).strip()

# -------------------------
# Tkinter UI
# -------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Сливане на поръчка + цени (Excel)")
        self.geometry("1200x650")

        self.order_path = tk.StringVar(value="")
        self.prices_path = tk.StringVar(value="")

        self.df_merged = None

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(side=tk.TOP, fill=tk.X)

        btn_order = ttk.Button(top, text="Качи Поръчка (.xls/.xlsx)", command=self.pick_order)
        btn_prices = ttk.Button(top, text="Качи Цени (.xls/.xlsx)", command=self.pick_prices)
        btn_merge = ttk.Button(top, text="Слей", command=self.do_merge)
        btn_save = ttk.Button(top, text="Запази .xlsx", command=self.save_xlsx)

        btn_order.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        btn_prices.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        btn_merge.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        btn_save.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        ttk.Label(top, text="Поръчка:").grid(row=1, column=0, sticky="w")
        ttk.Label(top, textvariable=self.order_path).grid(row=1, column=1, columnspan=3, sticky="w")

        ttk.Label(top, text="Цени:").grid(row=2, column=0, sticky="w")
        ttk.Label(top, textvariable=self.prices_path).grid(row=2, column=1, columnspan=3, sticky="w")

        # Table
        mid = ttk.Frame(self, padding=(10, 0, 10, 10))
        mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(mid, show="headings")
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(mid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Status
        self.status = tk.StringVar(value="Избери двата файла и натисни 'Слей'.")
        ttk.Label(self, textvariable=self.status, padding=10).pack(side=tk.BOTTOM, fill=tk.X)

    def pick_order(self):
        path = filedialog.askopenfilename(
            title="Избери файл Поръчка",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.order_path.set(path)

    def pick_prices(self):
        path = filedialog.askopenfilename(
            title="Избери файл Цени",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.prices_path.set(path)

    def do_merge(self):
        op = self.order_path.get().strip()
        pp = self.prices_path.get().strip()
        if not op or not pp:
            messagebox.showwarning("Липсват файлове", "Моля избери и двата файла (Поръчка и Цени).")
            return

        try:
            self.df_merged = merge_order_and_prices(op, pp)
            self._load_table(self.df_merged)
            self.status.set(f"Готово: {len(self.df_merged)} реда слети.")
        except Exception as e:
            messagebox.showerror("Грешка", str(e))
            self.status.set("Грешка при сливане.")

    def save_xlsx(self):
        if self.df_merged is None or self.df_merged.empty:
            messagebox.showinfo("Няма данни", "Първо натисни 'Слей'.")
            return

        # default name: Porachka_<orderNo>.xlsx (ако можем да го извлечем)
        default_name = "Porachka.xlsx"
        try:
            # order ref is like B26306-1
            first_ref = str(self.df_merged.iloc[0]["Номер на поръчка и ред"])
            order_no = first_ref.split("-")[0]
            default_name = f"Porachka_{order_no}.xlsx"
        except Exception:
            pass

        out_path = filedialog.asksaveasfilename(
            title="Запази като",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out_path:
            return

        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                self.df_merged.to_excel(writer, index=False, sheet_name="Porachka")
            messagebox.showinfo("Записано", f"Файлът е записан:\n{out_path}")
            self.status.set(f"Записано: {out_path}")
        except Exception as e:
            messagebox.showerror("Грешка при запис", str(e))
            self.status.set("Грешка при запис.")

    def _load_table(self, df: pd.DataFrame):
        # clear old
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())

        cols = list(df.columns)
        self.tree["columns"] = cols

        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140, anchor="w")

        # insert rows (limit render if huge)
        max_rows = 2000
        for i, row in df.head(max_rows).iterrows():
            values = [row.get(c, "") for c in cols]
            self.tree.insert("", "end", values=values)

        if len(df) > max_rows:
            self.status.set(f"Показвам първите {max_rows} реда от {len(df)} (всички се записват при export).")


if __name__ == "__main__":
    # nicer ttk theme on some systems
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App()
    app.mainloop()
