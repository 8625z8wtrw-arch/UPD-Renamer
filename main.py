# -*- coding: utf-8 -*-
"""
UPD/СФ Renamer — v4.1 (Точный поиск ИНН продавца и покупателя)
"""

import os
import sys
import csv
import re
import subprocess
import threading
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu, scrolledtext

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
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import pandas as pd
except ImportError:
    pd = None

# ============== УТИЛИТЫ ==============

BAD_CHARS = r'[\\/:*?"<>|]'

def safe_filename(name: str, max_len: int = 120) -> str:
    name = re.sub(BAD_CHARS, "_", name)
    name = name.strip().strip(".")
    if len(name) > max_len:
        name = name[:max_len - 1] + "…"
    return name

def abbreviate_org(name: str) -> str:
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
    s = re.sub(r'[^\d,.]', '', str(s))
    s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0

def open_file(path: Path):
    try:
        if sys.platform == "darwin": subprocess.run(["open", str(path)])
        elif sys.platform == "win32": os.startfile(str(path))
        else: subprocess.run(["xdg-open", str(path)])
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось открыть: {e}")

def open_folder(path: Path):
    try:
        if sys.platform == "darwin": subprocess.run(["open", str(path)])
        elif sys.platform == "win32": os.startfile(str(path))
        else: subprocess.run(["xdg-open", str(path)])
    except Exception: pass

def get_pdf_debug_text(path: Path) -> str:
    text = ""
    if FITZ_AVAILABLE:
        try:
            doc = fitz.open(str(path))
            for page in doc:
                text += page.get_text("text") + "\n--- Разрыв страницы ---\n"
            if text.strip(): return text
        except Exception as e: pass
    
    if pdfplumber:
        try:
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages:
                    text += (page.extract_text() or "") + "\n--- Разрыв страницы ---\n"
            if text.strip(): return text
        except: pass
    return "⚠️ Не удалось извлечь текст."


# ============== ПАРСЕРЫ ==============

def parse_xml_upd(path: Path) -> dict:
    if etree is None: return {"_error": "lxml не установлен"}
    data = {"type": "УПД", "sign": "1", "number": "", "date": "", "sum_vat": "", "sum_no_vat": "", "vat": "", "vat_rate": "", "seller_name": "", "seller_inn": "", "buyer_name": "", "buyer_inn": ""}
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
    
    if "УПД" not in xml_type and "УПД" not in path.name.upper(): data["type"] = "СФ"

    data["sign"] = find_text(["//*[local-name()='Признак']"]) or "1"
    data["number"] = find_text(["//*[local-name()='НомерДок']"])
    data["date"] = format_date_ddmmyyyy(find_text(["//*[local-name()='ДатаДок']"]))
    data["seller_name"] = find_text(["//*[local-name()='СвПрод']/*[local-name()='НаимОрг']", "//*[local-name()='Продавец']/*[local-name()='НаимОрг']"])
    data["seller_inn"] = find_text(["//*[local-name()='СвПрод']/*[local-name()='ИННЮЛ']", "//*[local-name()='Продавец']/*[local-name()='ИННЮЛ']"])
    data["buyer_name"] = find_text(["//*[local-name()='СвПокуп']/*[local-name()='НаимОрг']", "//*[local-name()='Покупатель']/*[local-name()='НаимОрг']"])
    data["buyer_inn"] = find_text(["//*[local-name()='СвПокуп']/*[local-name()='ИННЮЛ']", "//*[local-name()='Покупатель']/*[local-name()='ИННЮЛ']"])
    data["sum_vat"] = find_text(["//*[local-name()='СтТовУчНалВсего']"])
    data["sum_no_vat"] = find_text(["//*[local-name()='СтТовБезНДСВсего']"])
    data["vat"] = find_text(["//*[local-name()='СумНал']"])
    data["vat_rate"] = find_text(["//*[local-name()='СтНал']", "//*[local-name()='НалСт']"])
    
    if data["vat_rate"] and data["vat_rate"] not in ["Без НДС", "без НДС"] and "%" not in data["vat_rate"]:
        data["vat_rate"] += "%"
    return data


def parse_pdf(path: Path) -> dict:
    if not FITZ_AVAILABLE and not pdfplumber: 
        return {"_error": "Не установлены библиотеки для чтения PDF"}
    
    text = ""
    if FITZ_AVAILABLE:
        try:
            doc = fitz.open(str(path))
            for page in doc:
                text += page.get_text("text") + " "
        except Exception: pass
    
    if not text.strip() and pdfplumber:
        try:
            with pdfplumber.open(str(path)) as pdf:
                for page in pdf.pages:
                    text += (page.extract_text() or "") + " "
        except Exception as e:
            return {"_error": f"PDF не прочитан: {e}"}

    text = re.sub(r'\s+', ' ', text)

    data = {"type": "УПД", "sign": "1", "number": "", "date": "", "sum_vat": "", "sum_no_vat": "", "vat": "", "vat_rate": "", "seller_name": "", "seller_inn": "", "buyer_name": "", "buyer_inn": ""}
    text_upper = text.upper()
    
    if "УНИВЕРСАЛЬНЫЙ ПЕРЕДАТОЧНЫЙ ДОКУМЕНТ" not in text_upper and "УПД" not in path.name.upper():
        data["type"] = "СФ"

    m = re.search(r'[Пп]ризнак\s*[:\-]?\s*([12])', text)
    if m: data["sign"] = m.group(1)

    m = re.search(r'(?:№|Номер)\s*[:\-]?\s*([\w\-/]+)\s*(?:от|ОТ)\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})', text)
    if m:
        data["number"] = m.group(1)
        data["date"] = m.group(2)
    else:
        m_d = re.search(r'\b(\d{2}[.\-/]\d{2}[.\-/]\d{4})\b', text)
        if m_d: data["date"] = m_d.group(1)

    # --- УМНЫЙ ПОИСК ИНН ---
    # 1. Ищем конкретно продавца (как на твоем скриншоте)
    m_seller = re.search(r'ИНН[/\s]*КПП\s*продавца[^\d]{0,20}(\d{10,12})', text, re.IGNORECASE)
    if m_seller: data["seller_inn"] = m_seller.group(1)
    
    # 2. Ищем конкретно покупателя
    m_buyer = re.search(r'ИНН[/\s]*КПП\s*покупателя[^\d]{0,20}(\d{10,12})', text, re.IGNORECASE)
    if m_buyer: data["buyer_inn"] = m_buyer.group(1)

    # 3. Если по ключевым словам не нашли, берем просто все 10-12 значные числа
    if not data["seller_inn"] or not data["buyer_inn"]:
        inns = re.findall(r'ИНН[^\d]{0,30}(\d{10,12})', text, re.IGNORECASE)
        if not inns:
            inns = re.findall(r'\b(\d{10}|\d{12})\b', text)
        
        if len(inns) > 0 and not data["seller_inn"]: data["seller_inn"] = inns[0]
        if len(inns) > 1 and not data["buyer_inn"]: data["buyer_inn"] = inns[1]

    # --- МАТЕМАТИЧЕСКИЙ ПОИСК СУММ ---
    numbers = re.findall(r'(\d{1,3}(?:\s\d{3})*[.,]\d{2})', text)
    floats = [to_float(n) for n in numbers]
    
    found_math = False
    # Перебираем комбинации А + Б = В
    for i in range(len(floats)):
        for j in range(i+1, len(floats)):
            for k in range(len(floats)):
                if i == j or i == k or j == k: continue
                a, b, c = floats[i], floats[j], floats[k]
                
                if a > 0 and b > 0 and c > 0 and abs((a + b) - c) < 0.05:
                    rate = round((b / a) * 100)
                    if rate in (5, 7, 10, 20, 22):
                        data["sum_no_vat"] = numbers[i]
                        data["vat"] = numbers[j]
                        data["sum_vat"] = numbers[k]
                        data["vat_rate"] = f"{rate}%"
                        found_math = True
                        break
            if found_math: break
        if found_math: break

    # Если связка не найдена, ищем "Без НДС"
    if not found_math:
        if re.search(r'Без\s+НДС', text, re.IGNORECASE):
            data["vat_rate"] = "Без НДС"
            data["vat"] = "0.00"
            m = re.search(r'(?:Всего\s*к\s*оплате|Итого)[^\d]{0,50}?([\d\s]{3,20}[.,]\d{2})', text, re.IGNORECASE)
            if m:
                data["sum_vat"] = re.sub(r'[^\d,.]', '', m.group(1))
                data["sum_no_vat"] = data["sum_vat"]

    m = re.search(r'[Пп]родавец[:\s]+([«"\w][^,\n]{3,100})', text)
    if m: data["seller_name"] = m.group(1).strip()
    m = re.search(r'[Пп]окупатель[:\s]+([«"\w][^,\n]{3,100})', text)
    if m: data["buyer_name"] = m.group(1).strip()

    return data


def parse_file(path: Path) -> dict:
    ext = path.suffix.lower()
    if ext == ".xml": return parse_xml_upd(path)
    if ext == ".pdf": return parse_pdf(path)
    return {"_error": f"Формат {ext} не поддерживается"}


# ============== ИМЯ ==============

def build_new_name(row: dict, source: str, doc_type: str) -> str:
    raw_counterparty = row.get("buyer_name" if source == "buyer" else "seller_name", "")
    inn = row.get("buyer_inn" if source == "buyer" else "seller_inn", "")
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
        self.root.title("Переименование УПД / СФ (v4.1)")
        self.root.geometry("1250x800")
        self.root.minsize(900, 600)

        self.folder = tk.StringVar()
        self.source = tk.StringVar(value="buyer")
        self.search_var = tk.StringVar()
        
        self.files_data = []
        self.filtered_data = []
        self.is_processing = False
        self.auto_width_on = True

        self.all_columns = {
            "num": "№", "old": "Старое имя", 
            "type": "Тип", "sign": "Пр.", "number": "№ док", "date": "Дата", 
            "sum_vat": "Сумма с НДС", "sum_no_vat": "Без НДС", 
            "vat": "Сумма НДС", "vat_rate": "Ставка НДС", 
            "counterparty": "Контрагент", "inn": "ИНН", "new": "Новое имя"
        }
        self.visible_columns = list(self.all_columns.keys())

        self._build_ui()
        self._setup_tags()

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=5)
        top.pack(fill="x")

        ttk.Label(top, text="Папка:").pack(side="left")
        ttk.Entry(top, textvariable=self.folder).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(top, text="Выбрать…", command=self.choose_folder).pack(side="left")
        ttk.Button(top, text="📂", command=self.open_current_folder, width=3).pack(side="left", padx=2)
        ttk.Button(top, text="🛠 Текст PDF", command=self.debug_pdf).pack(side="left", padx=5)

        opts = ttk.Frame(self.root, padding=(5, 5))
        opts.pack(fill="x")

        ttk.Label(opts, text="В имя:").pack(side="left")
        ttk.Radiobutton(opts, text="Покупатель", variable=self.source, value="buyer", command=self.update_source).pack(side="left")
        ttk.Radiobutton(opts, text="Продавец", variable=self.source, value="seller", command=self.update_source).pack(side="left")
        
        ttk.Button(opts, text="🔍 Предпросмотр", command=self.preview).pack(side="left", padx=10)
        
        ttk.Label(opts, text="|  Поиск:").pack(side="left", padx=(10, 0))
        search_entry = ttk.Entry(opts, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.search_var.trace_add("write", self._on_search)

        ttk.Button(opts, text="↔ Авто-ширина", command=self.toggle_auto_width).pack(side="right")
        ttk.Button(opts, text="Сброс", command=self.reset_width).pack(side="right", padx=2)

        frame = ttk.Frame(self.root)
        frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.tree = ttk.Treeview(frame, columns=tuple(self.visible_columns), show="headings", height=15, selectmode="extended")
        self._setup_columns()

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._show_header_menu)
        self.tree.bind("<Control-Button-1>", self._show_header_menu)
        self.tree.bind("<ButtonRelease-1>", self._on_column_drag)

        btns = ttk.Frame(self.root, padding=5)
        btns.pack(fill="x")
        
        ttk.Button(btns, text="✅ Переименовать всё", command=lambda: self.rename(False)).pack(side="left")
        ttk.Button(btns, text="Переименовать выделенные", command=lambda: self.rename(True)).pack(side="left", padx=5)
        ttk.Button(btns, text="📊 Экспорт в Excel", command=self.export_to_excel).pack(side="left", padx=10)
        ttk.Button(btns, text="⏪ Откат по логу", command=self.rollback).pack(side="left")
        ttk.Button(btns, text="Очистить", command=self.clear).pack(side="right")

        status_frame = ttk.Frame(self.root, padding=(10, 5))
        status_frame.pack(fill="x", side="bottom")
        
        self.progress = ttk.Progressbar(status_frame, mode="determinate", length=200)
        self.progress.pack(side="left", padx=(0, 10))
        
        self.stats = ttk.Label(status_frame, text="Файлов: 0  |  Готово к работе", font=("Arial", 11, "bold"))
        self.stats.pack(side="left")

        legend_frame = ttk.Frame(status_frame)
        legend_frame.pack(side="right")
        
        ttk.Label(legend_frame, text="🟢 Все ок", foreground="green", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        ttk.Label(legend_frame, text="🟡 Мелкая помарка", foreground="#b8860b", font=("Arial", 10, "bold")).pack(side="left", padx=5)
        ttk.Label(legend_frame, text="🔴 Ошибка", foreground="red", font=("Arial", 10, "bold")).pack(side="left", padx=5)

        if DND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)
            ttk.Label(self.root, text="Перетащите папку ИЛИ отдельные файлы прямо в окно", foreground="gray").pack(pady=(0, 5))

    def debug_pdf(self):
        try:
            selection = self.tree.selection()
            if not selection:
                messagebox.showwarning("Внимание", "Сначала нажмите «Предпросмотр».\n\nЗатем кликните мышкой на нужную строку в таблице, и только потом нажимайте эту кнопку.")
                return
            
            num_id = selection[0]
            row = next((r for r in self.files_data if str(r["num"]) == num_id), None)
            if not row: return
            
            self.root.config(cursor="watch")
            self.root.update()
            
            text = get_pdf_debug_text(row["path"])
            
            win = tk.Toplevel(self.root)
            win.title(f"Текст PDF: {row['old']}")
            win.geometry("800x600")
            win.attributes('-topmost', True)
            win.focus_force()
            
            st = scrolledtext.ScrolledText(win, wrap=tk.WORD, font=("Courier", 12))
            st.pack(fill="both", expand=True)
            st.insert(tk.END, text)
            st.config(state="disabled")
            
            self.root.config(cursor="")
        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("Ошибка", f"Не удалось открыть текст PDF:\n{e}")

    def _setup_tags(self):
        self.tree.tag_configure('error', background='#ffcccc', foreground='black')
        self.tree.tag_configure('warning', background='#fff3cd', foreground='black')
        self.tree.tag_configure('ok', background='#d4edda', foreground='black')
        self.tree.tag_configure('even', background='#f9f9f9', foreground='black')
        self.tree.tag_configure('odd', background='white', foreground='black')

    def _setup_columns(self):
        self.tree["displaycolumns"] = tuple(self.visible_columns)
        for col in self.visible_columns:
            self.tree.heading(col, text=self.all_columns[col])
            width = 220 if col == "new" else (160 if col == "old" else (130 if col == "counterparty" else (50 if col in ("type", "sign", "num") else 90)))
            self.tree.column(col, width=width, minwidth=40, stretch=True)

    def toggle_auto_width(self):
        self.auto_width_on = not self.auto_width_on
        if self.auto_width_on: self.auto_fit_columns()
        else: self.reset_width()

    def reset_width(self):
        self._setup_columns()

    def _on_column_drag(self, event):
        if self.tree.identify_region(event.x, event.y) == "separator":
            self.auto_width_on = False

    def _show_header_menu(self, event):
        if self.tree.identify_region(event.x, event.y) != "heading": return
        menu = Menu(self.root, tearoff=0)
        for col, title in self.all_columns.items():
            var = tk.BooleanVar(value=(col in self.visible_columns))
            menu.add_checkbutton(label=title, variable=var, command=lambda c=col, v=var: self._toggle_column(c, v.get()))
        menu.post(event.x_root, event.y_root)

    def _toggle_column(self, col, show):
        if show and col not in self.visible_columns: self.visible_columns.append(col)
        elif not show and col in self.visible_columns: self.visible_columns.remove(col)
        self.visible_columns = [c for c in self.all_columns if c in self.visible_columns]
        self.tree["displaycolumns"] = tuple(self.visible_columns)
        self._refresh_tree()

    def _on_search(self, *args):
        query = self.search_var.get().lower()
        self.filtered_data = []
        for row in self.files_data:
            if (query in row.get("counterparty", "").lower() or 
                query in row.get("inn", "").lower() or 
                query in row.get("number", "").lower() or
                query in row.get("old", "").lower()):
                self.filtered_data.append(row)
        self._refresh_tree()
        self._update_stats()

    def _refresh_tree(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for i, row in enumerate(self.filtered_data):
            self._insert_row_to_ui(row, i)

    def _insert_row_to_ui(self, row, index):
        vals = tuple(row.get(c, "") for c in self.visible_columns)
        
        if not row["number"] or not row["date"]:
            tag = "error"
        elif not row["sum_vat"] or not row["inn"] or not row["counterparty"]:
            tag = "warning"
        else:
            tag = "ok"

        self.tree.insert("", "end", values=vals, tags=(tag,), iid=str(row["num"]))

    def update_source(self):
        if not self.files_data: return
        source = self.source.get()
        for row in self.files_data:
            row["counterparty"] = shorten_counterparty(row["buyer_name" if source == "buyer" else "seller_name"], 100)
            row["inn"] = row["buyer_inn" if source == "buyer" else "seller_inn"]
            row["new"] = build_new_name(row, source, row["type"])
        self._on_search()

    def open_current_folder(self):
        p = Path(self.folder.get())
        if p.exists(): open_folder(p)

    def choose_folder(self):
        p = filedialog.askdirectory()
        if p: self.folder.set(p)

    def clear(self):
        self.files_data = []
        self.filtered_data = []
        self._refresh_tree()
        self._update_stats()

    def auto_fit_columns(self):
        if not self.auto_width_on: return
        for col in self.tree["displaycolumns"]:
            max_len = len(self.tree.heading(col)["text"])
            for item in self.tree.get_children():
                max_len = max(max_len, len(str(self.tree.set(item, col))))
            self.tree.column(col, width=max(60, max_len * 8 + 15))

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        if not items: return
        
        paths = [Path(i) for i in items]
        dirs = [p for p in paths if p.is_dir()]
        files = [p for p in paths if p.is_file() and p.suffix.lower() in (".pdf", ".xml", ".xlsx", ".xls")]
        
        if dirs:
            self.folder.set(str(dirs[0]))
            self.preview()
        elif files:
            self.folder.set(str(files[0].parent))
            self.preview(specific_files=files)

    def preview(self, specific_files=None):
        if self.is_processing: return
        folder = Path(self.folder.get())
        if not folder.exists():
            messagebox.showerror("Ошибка", "Папка не найдена")
            return

        self.clear()
        
        if specific_files:
            files = sorted(specific_files)
        else:
            exts = {".xml", ".pdf", ".xlsx", ".xls"}
            files = sorted([f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in exts])
        
        if not files:
            messagebox.showinfo("Инфо", "Нет подходящих файлов.")
            return

        self.is_processing = True
        self.progress["maximum"] = len(files)
        self.progress["value"] = 0
        self.progress.pack(side="left", padx=(0, 10))
        
        threading.Thread(target=self._parse_worker, args=(files,), daemon=True).start()

    def _parse_worker(self, files):
        for i, f in enumerate(files, 1):
            data = parse_file(f)
            row = self._make_row(i, f, data)
            self.files_data.append(row)
            self.root.after(0, self.progress.configure, {"value": i})
        
        self.root.after(0, self._finish_parsing)

    def _finish_parsing(self):
        self.is_processing = False
        self.progress.pack_forget()
        self._on_search()
        if self.auto_width_on:
            self.auto_fit_columns()

    def _make_row(self, num: int, path: Path, data: dict) -> dict:
        doc_type = data.get("type", "УПД")
        row = {
            "num": num, "path": path, "type": doc_type, "sign": data.get("sign", ""),
            "number": data.get("number", ""), "date": data.get("date", ""),
            "sum_vat": data.get("sum_vat", ""), "sum_no_vat": data.get("sum_no_vat", ""),
            "vat": data.get("vat", ""), "vat_rate": data.get("vat_rate", ""),
            "buyer_name": data.get("buyer_name", ""), "seller_name": data.get("seller_name", ""),
            "buyer_inn": data.get("buyer_inn", ""), "seller_inn": data.get("seller_inn", ""),
            "old": path.name
        }
        source = self.source.get()
        row["counterparty"] = shorten_counterparty(row["buyer_name" if source == "buyer" else "seller_name"], 100)
        row["inn"] = row["buyer_inn" if source == "buyer" else "seller_inn"]
        row["new"] = build_new_name(row, source, doc_type)
        return row

    def _on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or not col: return
        
        col_name = self.tree["displaycolumns"][int(col.replace("#", "")) - 1]
        
        if col_name == "old":
            num_id = item
            for r in self.files_data:
                if str(r["num"]) == num_id:
                    open_file(r["path"])
                    break
            return
            
        if col_name in ("num", "new"): return
        
        x, y, w, h = self.tree.bbox(item, col)
        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, self.tree.set(item, col_name))
        entry.focus()

        def save(_):
            new_val = entry.get()
            self.tree.set(item, col_name, new_val)
            entry.destroy()
            self._recalc_row(item, col_name, new_val)

        entry.bind("<Return>", save)
        entry.bind("<FocusOut>", save)

    def _recalc_row(self, item, col_name, value):
        num_id = item
        row = next((r for r in self.files_data if str(r["num"]) == num_id), None)
        if not row: return
        
        row[col_name] = value
        if col_name == "counterparty":
            if self.source.get() == "buyer": row["buyer_name"] = value
            else: row["seller_name"] = value
        elif col_name == "inn":
            if self.source.get() == "buyer": row["buyer_inn"] = value
            else: row["seller_inn"] = value

        row["new"] = build_new_name(row, self.source.get(), row["type"])
        
        vals = tuple(row.get(c, "") for c in self.visible_columns)
        if not row["number"] or not row["date"]:
            tag = "error"
        elif not row["sum_vat"] or not row["inn"] or not row["counterparty"]:
            tag = "warning"
        else:
            tag = "ok"
            
        self.tree.item(item, values=vals, tags=(tag,))
        self._update_stats()

    def rename(self, only_selected=False):
        if not self.files_data:
            messagebox.showinfo("Инфо", "Сначала нажми «Предпросмотр»")
            return

        targets = []
        if only_selected:
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showinfo("Инфо", "Выдели строки в таблице.")
                return
            selected_nums = list(selected_items)
            targets = [r for r in self.files_data if str(r["num"]) in selected_nums]
        else:
            targets = self.files_data

        folder = Path(self.folder.get())
        log_path = folder / f"rename_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
        done, errors = 0, 0

        with open(log_path, "w", newline="", encoding="utf-8-sig") as fp:
            w = csv.writer(fp, delimiter=";")
            w.writerow(["old_name", "new_name", "date", "time"])

            for row in targets:
                old_path = row["path"]
                if not old_path.exists(): continue
                new_name = safe_filename(row["new"])
                target = unique_path(folder, new_name, old_path.suffix)
                try:
                    old_path.rename(target)
                    w.writerow([old_path.name, target.name, datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%H:%M:%S")])
                    done += 1
                except Exception:
                    errors += 1

        messagebox.showinfo("Готово", f"Успешно: {done}\nОшибок: {errors}")
        self.preview()

    def export_to_excel(self):
        if not self.filtered_data:
            messagebox.showinfo("Инфо", "Нет данных.")
            return
        if pd is None:
            messagebox.showerror("Ошибка", "Установите pandas.")
            return
        try:
            filepath = filedialog.asksaveasfilename(
                initialdir=self.folder.get(), title="Сохранить Excel",
                defaultextension=".xlsx", filetypes=[("Excel файлы", "*.xlsx")]
            )
            if not filepath: return
            
            data = []
            for item in self.tree.get_children():
                vals = self.tree.item(item, "values")
                row_dict = {self.all_columns[col]: vals[i] for i, col in enumerate(self.tree["displaycolumns"])}
                data.append(row_dict)
            
            pd.DataFrame(data).to_excel(filepath, index=False)
            messagebox.showinfo("Готово", f"Экспортировано:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось: {e}")

    def rollback(self):
        folder = Path(self.folder.get())
        logs = sorted(folder.glob("rename_log_*.csv"), reverse=True)
        if not logs:
            messagebox.showinfo("Инфо", "Логов не найдено")
            return
        
        with open(logs[0], "r", encoding="utf-8-sig") as fp:
            rows = list(csv.reader(fp, delimiter=";"))

        restored = 0
        for old_name, new_name, *_ in rows[1:]:
            new_file, old_file = folder / new_name, folder / old_name
            if new_file.exists() and not old_file.exists():
                try:
                    new_file.rename(old_file)
                    restored += 1
                except Exception: pass
        messagebox.showinfo("Откат", f"Восстановлено: {restored}")
        self.preview()

    def _update_stats(self):
        sum_vat = sum(to_float(r["sum_vat"]) for r in self.filtered_data)
        vat = sum(to_float(r["vat"]) for r in self.filtered_data)
        
        sum_vat_str = f"{sum_vat:,.2f}".replace(",", " ").replace(".", ",")
        vat_str = f"{vat:,.2f}".replace(",", " ").replace(".", ",")
        
        self.stats.config(text=f"Показано файлов: {len(self.filtered_data)}  |  Сумма с НДС: {sum_vat_str} руб.  |  НДС: {vat_str} руб.")

def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
