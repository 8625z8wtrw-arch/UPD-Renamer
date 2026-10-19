# -*- coding: utf-8 -*-
"""
UPD/СФ Renamer — программа для массового переименования файлов УПД и СФ.
Автоопределение типа, умный парсинг ИНН, растягиваемые таблицы и открытие файлов.
"""

import os
import sys
import csv
import re
import subprocess
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Drag-and-drop ---
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# --- Парсинг ---
try:
    from lxml import etree
except ImportError:
    etree = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import openpyxl
except ImportError:
    openpyxl = None


# ============== УТИЛИТЫ ==============

BAD_CHARS = r'[\\/:*?"<>|]'

def safe_filename(name: str, max_len: int = 120) -> str:
    name = re.sub(BAD_CHARS, "_", name)
    name = name.strip().strip(".")
    if len(name) > max_len:
        name = name[:max_len - 1] + "…"
    return name

def shorten_counterparty(name: str, limit: int = 40) -> str:
    name = name.strip()
    if len(name) > limit:
        return name[:limit - 1] + "…"
    return name

def normalize_number(num: str) -> str:
    return re.sub(r'\s+', '', str(num)).strip()

def format_date_ddmmyyyy(raw: str) -> str:
    raw = raw.strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y%m%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return raw

def to_float(s) -> float:
    if s is None: return 0.0
    if isinstance(s, (int, float)): return float(s)
    s = str(s).replace(" ", "").replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0

def open_file(path: Path):
    """Открывает файл в системном просмотрщике (Mac или Windows)."""
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)])
        elif sys.platform == "win32":
            os.startfile(str(path))
        else:
            subprocess.run(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")


# ============== ПАРСЕРЫ ==============

def extract_inns(text: str) -> list:
    """Надежно ищет ИНН в тексте, игнорируя КПП и лишние символы."""
    # Ищем "ИНН" и забираем 10 или 12 цифр, игнорируя пробелы и слеши
    matches = re.findall(r'ИНН[^\d]{0,15}(\d{10,12})', text, re.IGNORECASE)
    # Если не нашли, ищем просто 10 или 12 цифр подряд (запасной вариант)
    if not matches:
        matches = re.findall(r'\b(\d{10}|\d{12})\b', text)
    return matches

def parse_xml_upd(path: Path) -> dict:
    if etree is None: return {"_error": "lxml не установлен"}
    data = {}
    try:
        tree = etree.parse(str(path))
        root = tree.getroot()
    except Exception as e:
        return {"_error": f"XML не прочитан: {e}"}

    def find_text(xpath_list):
        for xp in xpath_list:
            try:
                res = root.xpath(xp)
                if res:
                    val = res[0]
                    if hasattr(val, "text"): val = val.text
                    if val: return str(val).strip()
            except Exception: continue
        return ""

    xml_type = ""
    try:
        types = root.xpath("//*[local-name()='Файл']/@*[local-name()='Тип']")
        if types: xml_type = str(types[0]).upper()
    except: pass
    
    data["type"] = "УПД" if ("УПД" in xml_type or "УПД" in path.name.upper()) else "СФ"

    data["sign"] = find_text([
        "//*[local-name()='СвСчФакт']/*[local-name()='ИнфПолФХЖ1']/*[local-name()='Признак']",
        "//*[local-name()='Признак']",
    ]) or "1"

    data["number"] = find_text(["//*[local-name()='СчетФактура']/*[local-name()='НомерДок']", "//*[local-name()='НомерДок']"])
    raw_date = find_text(["//*[local-name()='СчетФактура']/*[local-name()='ДатаДок']", "//*[local-name()='ДатаДок']"])
    data["date"] = format_date_ddmmyyyy(raw_date)

    data["seller_name"] = find_text(["//*[local-name()='СвПрод']/*[local-name()='НаимОрг']", "//*[local-name()='Продавец']/*[local-name()='НаимОрг']"])
    data["seller_inn"] = find_text(["//*[local-name()='СвПрод']/*[local-name()='ИННЮЛ']", "//*[local-name()='Продавец']/*[local-name()='ИННЮЛ']"])
    
    data["buyer_name"] = find_text(["//*[local-name()='СвПокуп']/*[local-name()='НаимОрг']", "//*[local-name()='Покупатель']/*[local-name()='НаимОрг']"])
    data["buyer_inn"] = find_text(["//*[local-name()='СвПокуп']/*[local-name()='ИННЮЛ']", "//*[local-name()='Покупатель']/*[local-name()='ИННЮЛ']"])

    data["sum_vat"] = find_text(["//*[local-name()='ВсегоОпл']/*[local-name()='СтТовУчНалВсего']", "//*[local-name()='СтТовУчНалВсего']"])
    data["sum_no_vat"] = find_text(["//*[local-name()='ВсегоОпл']/*[local-name()='СтТовБезНДСВсего']", "//*[local-name()='СтТовБезНДСВсего']"])
    data["vat"] = find_text(["//*[local-name()='ВсегоОпл']/*[local-name()='СумНалВсего']/*[local-name()='СумНал']", "//*[local-name()='СумНал']"])
    return data


def parse_pdf(path: Path) -> dict:
    if pdfplumber is None: return {"_error": "pdfplumber не установлен"}
    text = ""
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
    except Exception as e:
        return {"_error": f"PDF не прочитан: {e}"}

    data = {}
    text_upper = text.upper()
    
    data["type"] = "УПД" if ("УНИВЕРСАЛЬНЫЙ ПЕРЕДАТОЧНЫЙ ДОКУМЕНТ" in text_upper or "УПД" in path.name.upper()) else "СФ"

    m = re.search(r'[Пп]ризнак\s*[:\-]?\s*([12])', text)
    data["sign"] = m.group(1) if m else "1"

    m = re.search(r'(?:№|Номер)\s*[:\-]?\s*([\w\-/]+)\s+от\s+(\d{2}[.\-/]\d{2}[.\-/]\d{4})', text)
    if m:
        data["number"] = m.group(1)
        data["date"] = m.group(2)
    else:
        data["number"] = ""
        data["date"] = ""

    # Умный поиск ИНН
    inns = extract_inns(text)
    data["seller_inn"] = inns[0] if inns else ""
    data["buyer_inn"] = inns[1] if len(inns) > 1 else ""

    m = re.search(r'[Вв]сего\s+к\s+оплате.*?([\d\s]+[.,]\d{2})', text)
    data["sum_vat"] = m.group(1).strip() if m else ""
    m = re.search(r'[Сс]тоимость\s+без\s+НДС.*?([\d\s]+[.,]\d{2})', text)
    data["sum_no_vat"] = m.group(1).strip() if m else ""
    m = re.search(r'НДС[^\d]*([\d\s]+[.,]\d{2})', text)
    data["vat"] = m.group(1).strip() if m else ""

    m = re.search(r'[Пп]родавец[:\s]+([«"\w][^,\n]{3,80})', text)
    data["seller_name"] = m.group(1).strip() if m else ""
    m = re.search(r'[Пп]окупатель[:\s]+([«"\w][^,\n]{3,80})', text)
    data["buyer_name"] = m.group(1).strip() if m else ""

    return data


def parse_excel(path: Path) -> dict:
    if openpyxl is None: return {"_error": "openpyxl не установлен"}
    text = ""
    try:
        wb = openpyxl.load_workbook(str(path), data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is not None: text += str(cell) + " "
    except Exception as e:
        return {"_error": f"Excel не прочитан: {e}"}

    data = {}
    data["type"] = "УПД" if ("УПД" in text.upper() or "УПД" in path.name.upper()) else "СФ"

    m = re.search(r'[Пп]ризнак[:\s]*([12])', text)
    data["sign"] = m.group(1) if m else "1"
    m = re.search(r'№\s*([\w\-/]+)\s*от\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})', text)
    if m:
        data["number"] = m.group(1)
        data["date"] = m.group(2)
    else:
        data["number"] = data["date"] = ""
    
    inns = extract_inns(text)
    data["seller_inn"] = inns[0] if inns else ""
    data["buyer_inn"] = inns[1] if len(inns) > 1 else ""
    data["seller_name"] = data["buyer_name"] = ""
    data["sum_vat"] = data["sum_no_vat"] = data["vat"] = ""
    return data


def parse_file(path: Path) -> dict:
    ext = path.suffix.lower()
    if ext == ".xml": return parse_xml_upd(path)
    if ext == ".pdf": return parse_pdf(path)
    if ext in (".xlsx", ".xls"): return parse_excel(path)
    return {"_error": f"Формат {ext} не поддерживается"}


# ============== ИМЯ ==============

def build_new_name(data: dict, source: str, doc_type: str) -> str:
    counterparty = data.get("buyer_name" if source == "buyer" else "seller_name", "")
    inn = data.get("buyer_inn" if source == "buyer" else "seller_inn", "")
    counterparty = shorten_counterparty(counterparty, 40)

    number = normalize_number(data.get("number", ""))
    date = data.get("date", "")
    sum_vat = data.get("sum_vat", "")
    try:
        sum_vat = f"{to_float(sum_vat):,.2f}".replace(",", " ").replace(".", ",")
    except Exception: pass

    if doc_type == "УПД":
        name = f"УПД_({data.get('sign', '1')})_№{number}_от_{date}_{sum_vat} руб._{counterparty}_{inn}"
    else:
        name = f"СФ_№{number}_от_{date}_{sum_vat} руб._{counterparty}_{inn}"

    return safe_filename(name)


def unique_path(folder: Path, name: str, ext: str) -> Path:
    candidate = folder / f"{name}{ext}"
    i = 1
    while candidate.exists():
        candidate = folder / f"{name}_{i}{ext}"
        i += 1
    return candidate


# ============== GUI ==============

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Переименование УПД / СФ")
        self.root.geometry("1300x700") # Сделал окно чуть шире

        self.folder = tk.StringVar()
        self.source = tk.StringVar(value="buyer")
        self.files_data = []

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Папка:").pack(side="left")
        ttk.Entry(top, textvariable=self.folder, width=60).pack(side="left", padx=5)
        ttk.Button(top, text="Выбрать…", command=self.choose_folder).pack(side="left")

        opts = ttk.Frame(self.root, padding=(10, 0))
        opts.pack(fill="x")

        ttk.Label(opts, text="Источник данных для имени:").pack(side="left")
        ttk.Radiobutton(opts, text="Покупатель", variable=self.source, value="buyer").pack(side="left", padx=5)
        ttk.Radiobutton(opts, text="Продавец", variable=self.source, value="seller").pack(side="left")

        ttk.Button(opts, text="Предпросмотр", command=self.preview).pack(side="left", padx=15)
        # Кнопка автоподгонки ширины колонок
        ttk.Button(opts, text="Авто-ширина", command=self.auto_fit_columns).pack(side="left")

        # Таблица
        columns = ("old", "type", "sign", "number", "date", "sum_vat", "sum_no_vat",
                   "vat", "counterparty", "inn", "new")
        headers = ["Старое имя (двойной клик - открыть)", "Тип", "Пр.", "№", "Дата", "С НДС", "Без НДС",
                   "НДС", "Контрагент", "ИНН", "Новое имя"]

        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # stretch=True позволяет мышкой растягивать колонки
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=20)
        for col, head in zip(columns, headers):
            self.tree.heading(col, text=head)
            # Ставим более широкие колонки по умолчанию
            width = 250 if col == "new" else (200 if col == "old" else (120 if col == "counterparty" else (60 if col in ("type", "sign") else 90)))
            self.tree.column(col, width=width, minwidth=50, stretch=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_double_click)

        # Кнопки
        btns = ttk.Frame(self.root, padding=10)
        btns.pack(fill="x")
        ttk.Button(btns, text="Переименовать всё", command=self.rename_all).pack(side="left")
        ttk.Button(btns, text="Откатить по логу", command=self.rollback).pack(side="left", padx=10)
        ttk.Button(btns, text="Очистить", command=self.clear).pack(side="left")

        self.stats = ttk.Label(self.root, text="Файлов: 0  |  С НДС: 0  |  Без НДС: 0  |  НДС: 0", padding=10)
        self.stats.pack(fill="x")

        if DND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
            ttk.Label(self.root, text="Можно перетащить файлы прямо в окно", foreground="gray").pack(pady=(0, 5))

    def auto_fit_columns(self):
        """Автоматически подгоняет ширину колонок под текст."""
        for col in self.tree["columns"]:
            max_len = len(self.tree.heading(col)["text"])
            for item in self.tree.get_children():
                val = str(self.tree.set(item, col))
                max_len = max(max_len, len(val))
            # Примерная ширина символа в пикселях (зависит от шрифта)
            self.tree.column(col, width=max(80, max_len * 8 + 20))

    def _on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            self.folder.set(str(Path(files[0]).parent))
            self.preview()

    def choose_folder(self):
        p = filedialog.askdirectory()
        if p: self.folder.set(p)

    def clear(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        self.files_data = []
        self._update_stats()

    def preview(self):
        folder = Path(self.folder.get())
        if not folder.exists():
            messagebox.showerror("Ошибка", "Папка не найдена")
            return

        self.clear()
        exts = {".xml", ".pdf", ".xlsx", ".xls"}
        files = sorted([f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in exts])

        for f in files:
            data = parse_file(f)
            row = self._make_row(f, data)
            self.files_data.append(row)
            self.tree.insert("", "end", values=row["values"])

        self._update_stats()
        self.auto_fit_columns() # Подгоняем ширину после загрузки

    def _make_row(self, path: Path, data: dict) -> dict:
        source = self.source.get()
        counterparty = data.get("buyer_name" if source == "buyer" else "seller_name", "")
        inn = data.get("buyer_inn" if source == "buyer" else "seller_inn", "")
        doc_type = data.get("type", "УПД")
        
        new_name = build_new_name(data, source, doc_type)

        return {
            "path": path, "type": doc_type, "sign": data.get("sign", ""),
            "number": data.get("number", ""), "date": data.get("date", ""),
            "sum_vat": data.get("sum_vat", ""), "sum_no_vat": data.get("sum_no_vat", ""),
            "vat": data.get("vat", ""), "counterparty": counterparty, "inn": inn,
            "new_name": new_name,
            "values": (path.name, doc_type, data.get("sign", ""), data.get("number", ""),
                       data.get("date", ""), data.get("sum_vat", ""), data.get("sum_no_vat", ""),
                       data.get("vat", ""), counterparty, inn, new_name),
        }

    def _on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col: return
        
        col_idx = int(col.replace("#", "")) - 1
        idx = self.tree.index(item)
        
        # Если двойной клик по "Старому имени" - открываем файл
        if col_idx == 0:
            if idx < len(self.files_data):
                open_file(self.files_data[idx]["path"])
            return
            
        if col_idx in (0, 10): return # Старое и новое имя не редактируем напрямую
        
        x, y, w, h = self.tree.bbox(item, col)
        value = self.tree.set(item, self.tree["columns"][col_idx])

        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.focus()

        def save(_):
            new_val = entry.get()
            self.tree.set(item, self.tree["columns"][col_idx], new_val)
            entry.destroy()
            self._recalc_new_name(item)

        entry.bind("<Return>", save)
        entry.bind("<FocusOut>", save)

    def _recalc_new_name(self, item):
        idx = self.tree.index(item)
        if idx >= len(self.files_data): return
        row = self.files_data[idx]
        vals = self.tree.item(item, "values")
        
        row["type"] = vals[1] 
        row["sign"] = vals[2]
        row["number"] = vals[3]
        row["date"] = vals[4]
        row["sum_vat"] = vals[5]
        row["sum_no_vat"] = vals[6]
        row["vat"] = vals[7]
        row["counterparty"] = vals[8]
        row["inn"] = vals[9]

        data = {
            "sign": row["sign"], "number": row["number"], "date": row["date"],
            "sum_vat": row["sum_vat"], "buyer_name": row["counterparty"],
            "buyer_inn": row["inn"], "seller_name": row["counterparty"], "seller_inn": row["inn"],
        }
        row["new_name"] = build_new_name(data, self.source.get(), row["type"])
        self.tree.set(item, "new", row["new_name"])

    def rename_all(self):
        if not self.files_data:
            messagebox.showinfo("Инфо", "Сначала нажми «Предпросмотр»")
            return

        folder = Path(self.folder.get())
        log_path = folder / f"rename_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
        done = 0

        with open(log_path, "w", newline="", encoding="utf-8-sig") as fp:
            w = csv.writer(fp, delimiter=";")
            w.writerow(["old_name", "new_name", "date", "time"])

            for i, item in enumerate(self.tree.get_children()):
                row = self.files_data[i]
                old_path = row["path"]
                if not old_path.exists(): continue
                new_name = safe_filename(self.tree.set(item, "new"))
                target = unique_path(folder, new_name, old_path.suffix)
                try:
                    old_path.rename(target)
                    w.writerow([old_path.name, target.name, datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S")])
                    done += 1
                except Exception as e:
                    messagebox.showerror("Ошибка", f"{old_path.name}: {e}")

        messagebox.showinfo("Готово", f"Переименовано файлов: {done}\nЛог сохранен в папке.")
        self.preview()

    def rollback(self):
        folder = Path(self.folder.get())
        if not folder.exists(): return
        logs = sorted(folder.glob("rename_log_*.csv"), reverse=True)
        if not logs:
            messagebox.showinfo("Инфо", "Логов не найдено")
            return
        log_path = logs[0]

        with open(log_path, "r", encoding="utf-8-sig") as fp:
            rows = list(csv.reader(fp, delimiter=";"))

        restored = 0
        for old_name, new_name, *_ in rows[1:]:
            new_file = folder / new_name
            old_file = folder / old_name
            if new_file.exists() and not old_file.exists():
                try:
                    new_file.rename(old_file)
                    restored += 1
                except Exception: pass
        messagebox.showinfo("Откат", f"Восстановлено: {restored}")
        self.preview()

    def _update_stats(self):
        n = len(self.files_data)
        sum_vat = sum(to_float(r["sum_vat"]) for r in self.files_data)
        sum_no_vat = sum(to_float(r["sum_no_vat"]) for r in self.files_data)
        vat = sum(to_float(r["vat"]) for r in self.files_data)
        self.stats.config(text=f"Файлов: {n}  |  С НДС: {sum_vat:,.2f}  |  Без НДС: {sum_no_vat:,.2f}  |  НДС: {vat:,.2f}".replace(",", " "))


def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()