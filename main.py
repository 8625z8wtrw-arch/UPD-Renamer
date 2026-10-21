# -*- coding: utf-8 -*-
"""
UPD/СФ Renamer — программа для массового переименования файлов УПД и СФ.
Автоопределение типа, сокращение ОПФ, экспорт в Excel и умный поиск сумм.
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

def abbreviate_org(name: str) -> str:
    """Заменяет длинные названия ОПФ на аббревиатуры."""
    if not name: return name
    replacements = {
        r'Общество\s+с\s+ограниченной\s+ответственностью': 'ООО',
        r'Акционерное\s+общество': 'АО',
        r'Публичное\s+акционерное\s+общество': 'ПАО',
        r'Индивидуальный\s+предприниматель': 'ИП',
    }
    for pattern, short in replacements.items():
        name = re.sub(pattern, short, name, flags=re.IGNORECASE)
    return name.strip()

def shorten_counterparty(name: str, limit: int = 100) -> str:
    """Обрезает слишком длинное название с добавлением многоточия."""
    name = abbreviate_org(name)
    if len(name) > limit:
        return name[:limit - 1].strip() + "…"
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
    try:
        if sys.platform == "darwin": subprocess.run(["open", str(path)])
        elif sys.platform == "win32": os.startfile(str(path))
        else: subprocess.run(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть файл: {e}")


# ============== ПАРСЕРЫ ==============

def extract_inns(text: str) -> list:
    matches = re.findall(r'ИНН[^\d]{0,15}(\d{10,12})', text, re.IGNORECASE)
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

    inns = extract_inns(text)
    data["seller_inn"] = inns[0] if inns else ""
    data["buyer_inn"] = inns[1] if len(inns) > 1 else ""

    # Ищем суммы по разным шаблонам (учитывая продавцов и покупателей)
    m = re.search(r'(?:Всего\s+к\s+оплате|Итого\s+с\s+НДС|Сумма\s+с\s+НДС)[^\d]*([\d\s]+[.,]\d{2})', text, re.IGNORECASE)
    data["sum_vat"] = m.group(1).strip() if m else ""
    
    m = re.search(r'(?:Стоимость\s+без\s+НДС|Всего\s+без\s+НДС)[^\d]*([\d\s]+[.,]\d{2})', text, re.IGNORECASE)
    data["sum_no_vat"] = m.group(1).strip() if m else ""
    
    m = re.search(r'(?:Итого\s+)?НДС[^\d]*(?:в\s+том\s+числе\s+)?([\d\s]+[.,]\d{2})', text, re.IGNORECASE)
    if not m: m = re.search(r'Сумма\s+НДС[^\d]*([\d\s]+[.,]\d{2})', text, re.IGNORECASE)
    data["vat"] = m.group(1).strip() if m else ""

    m = re.search(r'[Пп]родавец[:\s]+([«"\w][^,\n]{3,100})', text)
    data["seller_name"] = m.group(1).strip() if m else ""
    m = re.search(r'[Пп]окупатель[:\s]+([«"\w][^,\n]{3,100})', text)
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

def build_new_name(row: dict, source: str, doc_type: str) -> str:
    if source == "buyer":
        raw_counterparty = row.get("buyer_name", "")
        inn = row.get("buyer_inn", "")
    else:
        raw_counterparty = row.get("seller_name", "")
        inn = row.get("seller_inn", "")
        
    counterparty = shorten_counterparty(raw_counterparty, 100)

    number = normalize_number(row.get("number", ""))
    date = row.get("date", "")
    sum_vat = row.get("sum_vat", "")
    try:
        sum_vat = f"{to_float(sum_vat):,.2f}".replace(",", " ").replace(".", ",")
    except Exception: pass

    if doc_type == "УПД":
        name = f"УПД_({row.get('sign', '1')})_№{number}_от_{date}_{sum_vat} руб._{counterparty}_{inn}"
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
        self.root.geometry("1350x700")

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
        # Добавил command=self.update_source для моментального пересчета при клике
        ttk.Radiobutton(opts, text="Покупатель", variable=self.source, value="buyer", command=self.update_source).pack(side="left", padx=5)
        ttk.Radiobutton(opts, text="Продавец", variable=self.source, value="seller", command=self.update_source).pack(side="left")

        ttk.Button(opts, text="Предпросмотр", command=self.preview).pack(side="left", padx=15)
        ttk.Button(opts, text="Авто-ширина", command=self.auto_fit_columns).pack(side="left")

        columns = ("old", "type", "sign", "number", "date", "sum_vat", "sum_no_vat",
                   "vat", "counterparty", "inn", "new")
        headers = ["Старое имя (двойной клик - открыть)", "Тип", "Пр.", "№", "Дата", "С НДС", "Без НДС",
                   "НДС", "Контрагент", "ИНН", "Новое имя"]

        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=20)
        for col, head in zip(columns, headers):
            self.tree.heading(col, text=head)
            width = 250 if col == "new" else (200 if col == "old" else (150 if col == "counterparty" else (60 if col in ("type", "sign") else 90)))
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

        btns = ttk.Frame(self.root, padding=10)
        btns.pack(fill="x")
        ttk.Button(btns, text="Переименовать всё", command=self.rename_all).pack(side="left")
        ttk.Button(btns, text="Экспорт в Excel", command=self.export_to_excel).pack(side="left", padx=10)
        ttk.Button(btns, text="Откатить по логу", command=self.rollback).pack(side="left", padx=10)
        ttk.Button(btns, text="Очистить", command=self.clear).pack(side="left")

        self.stats = ttk.Label(self.root, text="Файлов: 0  |  С НДС: 0  |  Без НДС: 0  |  НДС: 0", padding=10)
        self.stats.pack(fill="x")

        if DND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
            ttk.Label(self.root, text="Можно перетащить файлы прямо в окно", foreground="gray").pack(pady=(0, 5))

    def update_source(self):
        """Пересчитывает таблицу при переключении Покупатель/Продавец."""
        if not self.files_data: return
        for i, item in enumerate(self.tree.get_children()):
            row = self.files_data[i]
            source = self.source.get()
            
            if source == "buyer":
                raw_name = row.get("buyer_name", "")
                inn = row.get("buyer_inn", "")
            else:
                raw_name = row.get("seller_name", "")
                inn = row.get("seller_inn", "")
                
            counterparty = shorten_counterparty(raw_name, 100)
            new_name = build_new_name(row, source, row["type"])
            
            vals = list(self.tree.item(item, "values"))
            vals[8] = counterparty
            vals[9] = inn
            vals[10] = new_name
            self.tree.item(item, values=vals)
            row["new_name"] = new_name

    def auto_fit_columns(self):
        for col in self.tree["columns"]:
            max_len = len(self.tree.heading(col)["text"])
            for item in self.tree.get_children():
                val = str(self.tree.set(item, col))
                max_len = max(max_len, len(val))
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
        self.auto_fit_columns()

    def _make_row(self, path: Path, data: dict) -> dict:
        doc_type = data.get("type", "УПД")
        row = {
            "path": path, "type": doc_type, "sign": data.get("sign", ""),
            "number": data.get("number", ""), "date": data.get("date", ""),
            "sum_vat": data.get("sum_vat", ""), "sum_no_vat": data.get("sum_no_vat", ""),
            "vat": data.get("vat", ""), 
            "buyer_name": data.get("buyer_name", ""),
            "seller_name": data.get("seller_name", ""),
            "buyer_inn": data.get("buyer_inn", ""),
            "seller_inn": data.get("seller_inn", ""),
        }
        
        source = self.source.get()
        counterparty = shorten_counterparty(row["buyer_name" if source == "buyer" else "seller_name"], 100)
        inn = row["buyer_inn" if source == "buyer" else "seller_inn"]
        new_name = build_new_name(row, source, doc_type)
        row["new_name"] = new_name
        
        row["values"] = (
            path.name, doc_type, row["sign"], row["number"], row["date"], 
            row["sum_vat"], row["sum_no_vat"], row["vat"], counterparty, inn, new_name
        )
        return row

    def _on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col: return
        
        col_idx = int(col.replace("#", "")) - 1
        idx = self.tree.index(item)
        
        if col_idx == 0:
            if idx < len(self.files_data): open_file(self.files_data[idx]["path"])
            return
            
        if col_idx in (0, 10): return
        
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
        # Если пользователь правит контрагента вручную, пишем это в нужное поле
        if self.source.get() == "buyer":
            row["buyer_name"] = vals[8]
            row["buyer_inn"] = vals[9]
        else:
            row["seller_name"] = vals[8]
            row["seller_inn"] = vals[9]

        row["new_name"] = build_new_name(row, self.source.get(), row["type"])
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

        messagebox.showinfo("Готово", f"Переименовано файлов: {done}\nЛог сохранен.")
        self.preview()

    def export_to_excel(self):
        """Сохраняет таблицу в Excel."""
        if not self.files_data:
            messagebox.showinfo("Инфо", "Нет данных. Сначала нажми «Предпросмотр»")
            return
        try:
            import pandas as pd
            folder = Path(self.folder.get())
            if not folder.exists(): folder = Path.home()
            
            filepath = filedialog.asksaveasfilename(
                initialdir=str(folder), title="Сохранить Excel",
                defaultextension=".xlsx", filetypes=[("Excel файлы", "*.xlsx")]
            )
            if not filepath: return
            
            export_data = []
            for item in self.tree.get_children():
                vals = self.tree.item(item, "values")
                export_data.append({
                    "Старое имя": vals[0], "Тип": vals[1], "Признак": vals[2],
                    "Номер": vals[3], "Дата": vals[4], "С НДС": vals[5],
                    "Без НДС": vals[6], "НДС": vals[7], "Контрагент": vals[8],
                    "ИНН": vals[9], "Новое имя": vals[10]
                })
            
            df = pd.DataFrame(export_data)
            df.to_excel(filepath, index=False)
            messagebox.showinfo("Готово", f"Данные экспортированы:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")

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
