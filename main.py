# -*- coding: utf-8 -*-
"""
Avito Price Analyzer — настольное приложение (Windows/macOS/Linux). Обновлённая версия.

Главное:
- Поиск по десктопной версии Avito (www.avito.ru).
- Укреплён парсер (больше селекторов и фолбэков).
- Мягче фильтрация (порог совпадения 0.3).
- Запрос собирается из ключевых токенов (до 6), меньше "шума".

Входной Excel без шапки:
A — бренд
B — модель/характеристики (например: "iPhone 12, 128 ГБ, белый")
C — закупочная цена в ₽
Выход:
D — средняя цена по Avito (₽)
E — наценка, % = ((D − C)/C) * 100
F — ссылка на самое дешёвое из учтённых объявлений
Подсветка: 5–10% жёлтый; ≥10% зелёный.
"""

import threading
import queue
import time
import random
import re
import sys
import os
import math
from dataclasses import dataclass
from typing import List, Optional, Tuple

import requests
from bs4 import BeautifulSoup

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Avito Price Analyzer"
RESULT_SUFFIX = "_analyzed"
MAX_RESULTS_PER_QUERY = 20
REQUEST_TIMEOUT = 20
REQUEST_DELAY_RANGE = (1.0, 2.0)

MOBILE_UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.3 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
]

STOPWORDS_RU = {"и","или","а","для","с","на","в","из","к","по","без","до","от","что",
    "цвет","новый","б/у","бу","есть","нет","про","про-","встроенный","версия"}

@dataclass
class Listing:
    title: str
    url: str
    price_rub: Optional[int]

def log_print(widget: tk.Text, text: str):
    widget.insert(tk.END, text + "\n")
    widget.see(tk.END)
    widget.update_idletasks()

def normalize_text(s: str) -> str:
    s = s.lower().replace("чёр", "чер")
    s = re.sub(r"(\d+)\s*(гб|gb)", lambda m: f"{m.group(1)}gb", s)
    s = re.sub(r"(\d+)\s*(тб|tb)", lambda m: f"{m.group(1)}tb", s)
    s = re.sub(r"[^a-z0-9а-яё\s\-]+", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def extract_tokens(brand: str, name: str) -> List[str]:
    full = normalize_text(f"{brand} {name}")
    raw = re.findall(r"[a-z0-9а-яё\-]+", full, flags=re.IGNORECASE)
    toks, seen = [], set()
    for t in raw:
        if len(t) <= 1 or t in STOPWORDS_RU:
            continue
        if t not in seen:
            seen.add(t); toks.append(t)
    return toks

def build_search_url(query: str) -> str:
    # Десктопная версия Avito по всей России
    from urllib.parse import quote_plus
    return f"https://www.avito.ru/rossiya?q={quote_plus(query)}"

def fetch_html(url: str, session: requests.Session) -> Optional[str]:
    headers = {
        "User-Agent": random.choice(MOBILE_UAS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.7,en;q=0.6",
        "Connection": "keep-alive",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "DNT": "1",
        "Upgrade-Insecure-Requests": "1",
    }
    try:
        resp = session.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        if resp.status_code == 200 and "<html" in resp.text.lower():
            return resp.text
    except requests.RequestException:
        pass
    return None

def parse_listings_from_html(html: str, base_url: str) -> List[Listing]:
    soup = BeautifulSoup(html, "lxml")
    full_text = soup.get_text(" ", strip=True).lower()
    if "не робот" in full_text or "captcha" in full_text or "подтвердите" in full_text:
        return []

    candidates = []
    candidates.extend(soup.select('div[data-marker="item"]'))
    candidates.extend(soup.select('div[class*="iva-item"]'))
    candidates.extend(soup.select('article'))
    # dedup
    uniq, seen = [], set()
    for c in candidates:
        key = c.get("data-item-id") or str(hash(c.get_text(" ", strip=True)))[:16]
        if key not in seen:
            seen.add(key); uniq.append(c)
    candidates = uniq

    def extract_title_url(node) -> Tuple[Optional[str], Optional[str]]:
        a = node.select_one('a[data-marker="item-title"]') or node.select_one('a[class*="link-link"]') or node.find('a', href=True)
        title, url = None, None
        if a:
            title = a.get_text(" ", strip=True) or None
            href = a.get("href")
            if href:
                if href.startswith("http"):
                    url = href
                else:
                    if href.startswith("/"):
                        url = f"https://www.avito.ru{href}"
                    else:
                        url = base_url.rstrip("/") + "/" + href
        return title, url

    def extract_price(node) -> Optional[int]:
        mp = node.select_one('meta[itemprop="price"]')
        if mp and mp.get("content") and mp["content"].isdigit():
            return int(mp["content"])
        p = node.select_one('[data-marker="item-price"]') or node.select_one('span[itemprop="price"]') or node.select_one('strong[class*="price"]')
        if p and p.get_text(strip=True):
            digits = re.findall(r"\d+", p.get_text(" ", strip=True))
            if digits: return int(" ".join(digits).replace(" ", ""))
        txt = node.get_text(" ", strip=True)
        if "₽" in txt:
            digits = re.findall(r"\d+", txt)
            if digits: return int("".join(digits))
        return None

    listings = []
    for node in candidates:
        title, url = extract_title_url(node)
        if not title or not url: continue
        price = extract_price(node)
        listings.append(Listing(title=title, url=url, price_rub=price))

    return listings[:80]

def score_listing_match(listing_title: str, required_tokens: List[str]) -> float:
    if not required_tokens: return 0.0
    lt = normalize_text(listing_title)
    hits = sum(1 for t in required_tokens if t in lt)
    return hits / len(required_tokens)

def choose_relevant(listings: List[Listing], required_tokens: List[str]) -> List[Listing]:
    scored = [(score_listing_match(li.title, required_tokens), li) for li in listings if li.price_rub]
    scored.sort(key=lambda x: x[0], reverse=True)
    filtered = [li for s, li in scored if s >= 0.3]
    return filtered[:20]

def make_output_path(input_path: str) -> str:
    base, ext = os.path.splitext(input_path)
    if ext.lower() not in [".xlsx", ".xls"]:
        ext = ".xlsx"
    return f"{base}{RESULT_SUFFIX}{ext}"

def apply_row_fill(ws, row_idx: int, fill: PatternFill, max_col: int):
    for col in range(1, max_col + 1):
        ws.cell(row=row_idx, column=col).fill = fill

def save_with_formatting(df: pd.DataFrame, out_path: str):
    tmp_path = out_path
    df.to_excel(tmp_path, header=False, index=False)
    from openpyxl import load_workbook
    wb = load_workbook(tmp_path)
    ws = wb.active

    col_widths = [20, 50, 14, 16, 12, 50]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    max_row = ws.max_row
    for row in range(1, max_row + 1):
        c_cell = ws.cell(row=row, column=3)
        d_cell = ws.cell(row=row, column=4)
        e_cell = ws.cell(row=row, column=5)
        c_cell.number_format = '#,##0" ₽"'
        d_cell.number_format = '#,##0" ₽"'
        e_cell.number_format = '0.00" %"'

        e_val = e_cell.value
        if isinstance(e_val, (int, float)):
            if 5.0 <= e_val < 10.0:
                fill = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
                apply_row_fill(ws, row, fill, max_col=6)
            elif e_val >= 10.0:
                fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
                apply_row_fill(ws, row, fill, max_col=6)

    wb.save(tmp_path)

def process_excel(input_path: str, log_widget: tk.Text, progress: ttk.Progressbar) -> Optional[str]:
    try:
        df = pd.read_excel(input_path, header=None)
    except Exception as e:
        log_print(log_widget, f"Ошибка чтения файла: {e}")
        return None

    if df.shape[1] < 3:
        log_print(log_widget, "Ошибка: в файле должно быть минимум 3 колонки (A,B,C).")
        return None

    def to_float(x):
        try:
            if pd.isna(x): return math.nan
            s = str(x).replace(" ", "").replace("\xa0", "").replace(",", ".")
            s = re.sub(r"[^0-9.\-]", "", s)
            return float(s) if s else math.nan
        except Exception:
            return math.nan

    session = requests.Session()
    total_rows = len(df)

    for idx, row in df.iterrows():
        brand = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else ""
        name  = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
        cost_c = to_float(row.iloc[2])

        progress["value"] = int((idx / max(1, total_rows)) * 100); progress.update_idletasks()

        if not brand.strip() and not name.strip():
            log_print(log_widget, f"[{idx+1}] Пустая строка — пропуск.")
            df.loc[idx, 3] = None; df.loc[idx, 4] = None; df.loc[idx, 5] = None
            continue

        tokens = extract_tokens(brand, name)
        query = " ".join(tokens[:6]) if tokens else f"{brand} {name}".strip()
        log_print(log_widget, f"[{idx+1}] Ищу: {query} | ключи: {', '.join(tokens) if tokens else '—'}")

        url = build_search_url(query)
        html = fetch_html(url, session)

        if not html:
            log_print(log_widget, "   ⚠️ Не удалось получить результаты (возможно, защита сайта/сеть).")
            df.loc[idx, 3] = None; df.loc[idx, 4] = None; df.loc[idx, 5] = None
            time.sleep(random.uniform(*REQUEST_DELAY_RANGE))
            continue

        listings_all = parse_listings_from_html(html, base_url="https://www.avito.ru")
        log_print(log_widget, f"   Найдено сырых карточек: {len(listings_all)}")

        if not listings_all:
            log_print(log_widget, "   ⚠️ Ничего не найдено по запросу.")
            df.loc[idx, 3] = None; df.loc[idx, 4] = None; df.loc[idx, 5] = None
            time.sleep(random.uniform(*REQUEST_DELAY_RANGE)); continue

        listings = choose_relevant(listings_all, tokens)
        log_print(log_widget, f"   Соответствует фильтру: {len(listings)}")

        if not listings:
            log_print(log_widget, "   ⚠️ Нашёл объявления, но ни одно не прошло фильтр по конфигурации.")
            df.loc[idx, 3] = None; df.loc[idx, 4] = None; df.loc[idx, 5] = None
            time.sleep(random.uniform(*REQUEST_DELAY_RANGE)); continue

        prices = [li.price_rub for li in listings if li.price_rub is not None]
        avg_price = round(sum(prices) / len(prices), 2) if prices else None
        cheapest = min(listings, key=lambda li: li.price_rub or 10**12)

        df.loc[idx, 3] = avg_price
        if avg_price is not None and not (cost_c is None or (isinstance(cost_c, float) and math.isnan(cost_c))):
            markup_percent = (avg_price - cost_c) / cost_c * 100.0 if cost_c != 0 else None
            df.loc[idx, 4] = markup_percent
        else:
            df.loc[idx, 4] = None
        df.loc[idx, 5] = cheapest.url if cheapest and cheapest.url else None

        log_print(log_widget, f"   ✓ Учтено объявлений: {len(listings)}; средняя: {avg_price if avg_price else '—'} ₽; мин: {cheapest.price_rub if cheapest.price_rub else '—'} ₽")
        time.sleep(random.uniform(*REQUEST_DELAY_RANGE))

    out_path = make_output_path(input_path)
    try:
        save_with_formatting(df, out_path)
    except Exception as e:
        log_print(log_widget, f"⚠️ Не удалось применить форматирование, сохраню без него: {e}")
        df.to_excel(out_path, header=False, index=False)

    progress["value"] = 100; progress.update_idletasks()
    return out_path

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x680")
        self.resizable(True, True)

        self.input_path = None
        self.output_path = None
        self.worker_thread = None
        self.task_queue = queue.Queue()

        self.create_widgets()
        self.after(100, self.process_queue)

    def create_widgets(self):
        pad = 10
        frm_top = ttk.Frame(self, padding=pad); frm_top.pack(fill="x")
        self.path_var = tk.StringVar()
        ttk.Label(frm_top, text="Файл Excel (A=бренд, B=модель/характеристики, C=закупка ₽):").pack(anchor="w")
        row = ttk.Frame(frm_top); row.pack(fill="x", pady=(6,0))
        self.path_entry = ttk.Entry(row, textvariable=self.path_var); self.path_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Выбрать файл…", command=self.choose_file).pack(side="left", padx=(6,0))

        frm_btns = ttk.Frame(self, padding=pad); frm_btns.pack(fill="x")
        self.btn_start = ttk.Button(frm_btns, text="Старт", command=self.on_start, state="disabled"); self.btn_start.pack(side="left")
        ttk.Button(frm_btns, text="Закрыть", command=self.on_close).pack(side="right")

        frm_prog = ttk.Frame(self, padding=pad); frm_prog.pack(fill="x")
        ttk.Label(frm_prog, text="Ход обработки:").pack(anchor="w")
        self.progress = ttk.Progressbar(frm_prog, orient="horizontal", mode="determinate"); self.progress.pack(fill="x", pady=(6,0))

        frm_log = ttk.Frame(self, padding=pad); frm_log.pack(fill="both", expand=True)
        ttk.Label(frm_log, text="Лог:").pack(anchor="w")
        self.log = tk.Text(frm_log, height=20, wrap="word"); self.log.pack(fill="both", expand=True)

        frm_out = ttk.Frame(self, padding=pad); frm_out.pack(fill="x")
        ttk.Label(frm_out, text="Итоговый файл:").pack(anchor="w")
        out_row = ttk.Frame(frm_out); out_row.pack(fill="x", pady=(6,0))
        self.out_var = tk.StringVar(value="—")
        ttk.Entry(out_row, textvariable=self.out_var).pack(side="left", fill="x", expand=True)
        ttk.Button(out_row, text="Открыть папку", command=self.open_output_folder).pack(side="left", padx=(6,0))

        self.update_start_state()

    def choose_file(self):
        path = filedialog.askopenfilename(title="Выберите Excel-файл", filetypes=[("Excel files","*.xlsx *.xls")])
        if path:
            self.input_path = path; self.path_var.set(path); self.update_start_state()

    def update_start_state(self):
        self.btn_start["state"] = "normal" if self.input_path else "disabled"

    def on_start(self):
        if not self.input_path:
            messagebox.showwarning("Внимание", "Сначала выберите Excel-файл."); return
        self.progress["value"] = 0; self.out_var.set("—"); self.output_path = None; self.log.delete("1.0", tk.END); self.btn_start["state"] = "disabled"

        def worker():
            try:
                out = process_excel(self.input_path, self.log, self.progress)
                def ok():
                    if out:
                        self.output_path = out; self.out_var.set(out); messagebox.showinfo("Готово", "Обработка завершена.")
                    else:
                        messagebox.showwarning("Готово", "Обработка завершена с предупреждениями. См. лог.")
                self.task_queue.put(ok)
            except Exception as e:
                def err():
                    messagebox.showerror("Ошибка", f"Непредвиденная ошибка: {e}"); self.btn_start["state"] = "normal"
                self.task_queue.put(err)
            finally:
                def fin():
                    self.btn_start["state"] = "normal"
                self.task_queue.put(fin)

        self.worker_thread = threading.Thread(target=worker, daemon=True); self.worker_thread.start()

    def process_queue(self):
        try:
            while True:
                cb = self.task_queue.get_nowait(); cb()
        except queue.Empty:
            pass
        self.after(100, self.process_queue)

    def open_output_folder(self):
        if not self.output_path or not os.path.exists(self.output_path):
            messagebox.showinfo("Инфо", "Итоговый файл пока не создан."); return
        folder = os.path.dirname(self.output_path) or "."
        if sys.platform.startswith("win"): os.startfile(folder)
        elif sys.platform == "darwin": os.system(f'open "{folder}"')
        else: os.system(f'xdg-open "{folder}"')

    def on_close(self): self.destroy()

def main():
    app = App(); app.mainloop()

if __name__ == "__main__":
    main()
