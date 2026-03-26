"""
Извлечение 4 форм МСФО из PDF-отчёта российского банка с помощью MinerU.

Формы:
  1. Отчёт о финансовом положении (баланс)
  2. Отчёт о прибылях и убытках и прочем совокупном доходе
  3. Отчёт о движении денежных средств
  4. Отчёт об изменениях в собственном капитале

Использование:
  python extract_ifrs.py sber.pdf
  python extract_ifrs.py sber.pdf --output output_dir
"""

import sys
import os
import re
import json
import csv
import argparse
import shutil
from pathlib import Path
from html.parser import HTMLParser

# Windows: переключаем stdout на UTF-8, чтобы корректно выводить кириллицу
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ─────────────────────────── MinerU imports ────────────────────────────

from mineru.cli.common import read_fn, do_parse
from mineru.utils.enum_class import MakeMode

# ──────────────────────── Ключевые фразы форм ───────────────────────────

FORM_KEYWORDS: dict[str, list[str]] = {
    "balance_sheet": [
        r"финансово.{0,5}\s*положени",  # финансовом положении / финансового положения
        r"отчет.{0,20}финансово",
        "statement of financial position",
        "balance sheet",
    ],
    "income_statement": [
        r"прибыл.{0,5}\s*и\s*убытк",   # прибылях и убытках / прибыли и убытках
        r"совокупн.{0,5}\s*доход",
        "прочем совокупном",
        "profit or loss",
        "income statement",
        "comprehensive income",
    ],
    "cash_flow": [
        r"движени.{0,10}денежн",        # движении денежных средств
        "cash flow",
        "денежные потоки",
    ],
    "equity_changes": [
        r"изменени.{0,20}собственн",    # изменениях в составе собственных
        r"изменени.{0,10}капитал",
        "changes in equity",
        "собственных средств",
    ],
}

FORM_NAMES_RU: dict[str, str] = {
    "balance_sheet":    "Отчёт о финансовом положении",
    "income_statement": "Отчёт о прибылях и убытках",
    "cash_flow":        "Отчёт о движении денежных средств",
    "equity_changes":   "Отчёт об изменениях в собственном капитале",
}

# ──────────────────────────── HTML → таблица ────────────────────────────

class _TableParser(HTMLParser):
    """
    Парсер HTML-таблиц → список строк (list[list[str]]).
    Обрабатывает colspan/rowspan, <br> внутри ячеек.
    """

    def __init__(self):
        super().__init__()
        self.rows: list[list[str]] = []
        self._cur_row: list[str] = []
        self._cur_cell: list[str] = []
        self._in_cell = False

    @staticmethod
    def _attr(attrs: list, name: str, default: int = 1) -> int:
        for k, v in attrs:
            if k == name:
                try:
                    return max(1, int(v))
                except (TypeError, ValueError):
                    pass
        return default

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self._cur_row = []
        elif tag in ("td", "th"):
            self._cur_cell = []
            self._in_cell = True
            # colspan → add extra empty cells after this one later
            self._cur_colspan = self._attr(attrs, "colspan")
        elif tag == "br" and self._in_cell:
            self._cur_cell.append(" ")

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            text = "".join(self._cur_cell).strip()
            self._cur_row.append(text)
            # Fill extra columns for colspan
            for _ in range(self._cur_colspan - 1):
                self._cur_row.append("")
            self._in_cell = False
            self._cur_colspan = 1
        elif tag == "tr":
            if self._cur_row:
                self.rows.append(self._cur_row)

    def handle_data(self, data):
        if self._in_cell:
            self._cur_cell.append(data)


def html_to_rows(html: str) -> list[list[str]]:
    parser = _TableParser()
    parser._cur_colspan = 1
    parser.feed(html)
    return parser.rows


# ───────────────────────── Нормализация числа ───────────────────────────

# Типичные артефакты OCR в числах:
#   «О» вместо «0», «l» вместо «1», слитые тысячные разделители
_DIGIT_FIXES = [
    (re.compile(r'(?<=\d)O(?=\d)'), '0'),   # О → 0 между цифрами
    (re.compile(r'(?<=\d)l(?=\d)'), '1'),   # l → 1 между цифрами
    (re.compile(r'(?<=\d)I(?=\d)'), '1'),   # I → 1 между цифрами
]

# Пробел / тонкая пробельная черта как тысячный разделитель
_THOUSANDS_SEP = re.compile(r'(\d)\s{1,2}(\d{3})(?!\d)')
# Ситуация, когда знак минус — тире/дефис
_MINUS_NORM = re.compile(r'^[–—−‒-]\s*')


def normalize_number(text: str) -> str:
    """
    Нормализует строку с числом:
    - убирает тысячные пробелы (но сохраняет значение)
    - нормализует минус
    - исправляет OCR-артефакты
    Если строка не является числом — возвращает как есть.
    """
    t = text.strip()
    if not t:
        return t

    # Исправление OCR-артефактов
    for pat, repl in _DIGIT_FIXES:
        t = pat.sub(repl, t)

    # Нормализация минуса
    t = _MINUS_NORM.sub('-', t)

    # Убираем тысячные разделители (пробел между группами цифр)
    # Применяем многократно, пока есть совпадения
    prev = None
    while prev != t:
        prev = t
        t = _THOUSANDS_SEP.sub(r'\1\2', t)

    # Проверяем, что это действительно число
    clean = t.replace(',', '.').replace(' ', '')
    try:
        float(clean)
        return t          # числовая строка — возвращаем нормализованную
    except ValueError:
        return text.strip()  # не число — оригинал


def clean_cell(text: str) -> str:
    """Очистка ячейки: убирает лишние переносы, нормализует пробелы."""
    # Убираем мягкий перенос и другие невидимые символы
    text = text.replace('\u00ad', '').replace('\u200b', '').replace('\xa0', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    return normalize_number(text)


# ──────────────────────── Определение формы ────────────────────────────

def _text_matches(text: str, form_key: str) -> bool:
    t = text.lower()
    for kw in FORM_KEYWORDS[form_key]:
        if re.search(kw, t, re.IGNORECASE):
            return True
    return False


def detect_form_from_caption(caption: str) -> str | None:
    """
    Определяем форму только по caption таблицы (более надёжно).
    Приоритет: caption проверяется первым.
    """
    if not caption:
        return None
    for key in ("cash_flow", "equity_changes", "income_statement", "balance_sheet"):
        if _text_matches(caption, key):
            return key
    return None


def detect_form(context_texts: list[str]) -> str | None:
    """
    По списку текстовых блоков рядом с таблицей определяем,
    к какой форме она относится.
    Возвращает ключ формы или None.
    """
    combined = " ".join(context_texts)
    # Приоритет: сначала более специфичные
    for key in ("cash_flow", "equity_changes", "income_statement", "balance_sheet"):
        if _text_matches(combined, key):
            return key
    return None


# ───────────────────── Парсинг content_list.json ───────────────────────

def extract_tables_from_content_list(
    content_list_path: str,
) -> dict[str, list[list[list[str]]]]:
    """
    Читает content_list.json и собирает таблицы по 4 формам.

    Возвращает словарь: form_key → список таблиц (каждая таблица = list[list[str]]).
    """
    with open(content_list_path, encoding="utf-8") as f:
        blocks = json.load(f)

    result: dict[str, list[list[list[str]]]] = {k: [] for k in FORM_KEYWORDS}

    # Скользящее окно контекста — последние N текстовых блоков
    CONTEXT_WINDOW = 8
    text_window: list[str] = []

    for block in blocks:
        btype = block.get("type", "")

        if btype == "text":
            txt = block.get("text", "").strip()
            if txt:
                text_window.append(txt)
                if len(text_window) > CONTEXT_WINDOW:
                    text_window.pop(0)

        elif btype == "table":
            html = block.get("table_body") or block.get("html") or ""
            if not html:
                continue

            rows = html_to_rows(html)
            if not rows:
                continue

            # Дополнительно проверяем caption / заголовок внутри блока
            caption = block.get("table_caption", [])
            caption_text = " ".join(caption) if isinstance(caption, list) else str(caption)

            # Только caption: единственный надёжный признак для
            # основных форм отчётности. Контекстный поиск по тексту
            # захватывает таблицы из примечаний — не используем.
            form_key = detect_form_from_caption(caption_text)
            if form_key:
                result[form_key].append(rows)

    return result


# ─────────────────────── Склейка таблиц (pagination) ──────────────────

def merge_table_pages(tables: list[list[list[str]]]) -> list[list[str]]:
    """
    Склеивает список таблиц одной формы (разбитых по страницам).
    Дублирующийся заголовок (первая строка) при склейке пропускается.
    """
    if not tables:
        return []
    merged: list[list[str]] = list(tables[0])
    if len(tables) == 1:
        return merged

    header = tables[0][0] if tables[0] else []

    for tbl in tables[1:]:
        if not tbl:
            continue
        # Если первая строка совпадает с заголовком — пропускаем
        start = 1 if (tbl[0] == header or _rows_similar(tbl[0], header)) else 0
        merged.extend(tbl[start:])

    return merged


def _rows_similar(a: list[str], b: list[str]) -> bool:
    """Проверяет, похожи ли две строки (заголовки) по большинству ячеек."""
    if not a or not b:
        return False
    matches = sum(1 for x, y in zip(a, b) if x.strip().lower() == y.strip().lower())
    return matches / max(len(a), len(b)) > 0.6


# ──────────────────────── Исправление разрывов строк ─────────────────────

def fix_line_breaks_in_table(rows: list[list[str]]) -> list[list[str]]:
    """
    Исправляет разрывы строк внутри ячеек таблиц.
    
    Правила:
    1. Если строка начинается со строчной буквы (в первой непустой ячейке),
       она считается продолжением предыдущей строки.
    2. Текст из первой ячейки строки-продолжения добавляется к последней
       непустой ячейке предыдущей строки.
    3. Остальные ячейки строки-продолжения (данные) переносятся в 
       соответствующие колонки предыдущей строки, если там пусто.
    4. Строка-продолжение удаляется из результата.
    
    Пример:
      Было:
        ["Процентные доходы, рассчитанные по методу эффективной", "", "", ""]
        ["ставки", "463,3", "269,0", "-20,9%", "24,0%"]
      Стало:
        ["Процентные доходы, рассчитанные по методу эффективной ставки", "463,3", "269,0", "-20,9%", "24,0%"]
    """
    if not rows:
        return []
    
    result: list[list[str]] = []
    
    for row in rows:
        if not row:
            continue
        
        # Проверяем, является ли текущая строка продолжением предыдущей
        is_continuation = False
        if result:
            first_non_empty = ""
            for cell in row:
                stripped = cell.strip()
                if stripped:
                    first_non_empty = stripped
                    break
            
            # Если первая непустая ячейка начинается со строчной буквы
            if first_non_empty and first_non_empty[0].islower():
                is_continuation = True
        
        if is_continuation and result:
            # Склейка с предыдущей строкой
            prev_row = result[-1]
            
            # Расширяем предыдущую строку до размера текущей, если нужно
            while len(prev_row) < len(row):
                prev_row.append("")
            
            # Находим последнюю непустую ячейку в предыдущей строке
            # и добавляем к ней текст из первой ячейки текущей строки
            last_non_empty_idx = -1
            for i in range(len(prev_row) - 1, -1, -1):
                if prev_row[i].strip():
                    last_non_empty_idx = i
                    break
            
            if last_non_empty_idx >= 0:
                # Добавляем текст продолжения к последней ячейке
                continuation_text = row[0].strip() if row else ""
                if continuation_text:
                    prev_row[last_non_empty_idx] = prev_row[last_non_empty_idx].rstrip() + " " + continuation_text
            elif row and row[0].strip():
                # Если в предыдущей строке всё пусто, просто добавляем текст
                if len(prev_row) > 0:
                    prev_row[0] = row[0].strip()
            
            # Переносим данные из остальных ячеек строки-продолжения
            for i in range(1, len(row)):
                if i < len(prev_row) and row[i].strip():
                    # Если в предыдущей строке ячейка пуста, заполняем её
                    if not prev_row[i].strip():
                        prev_row[i] = row[i].strip()
                    # Иначе оставляем как есть (предполагаем, что данные уже там)
        else:
            # Новая строка — добавляем как есть
            result.append(list(row))
    
    return result


# ──────────────────────────── Запись CSV ────────────────────────────────

def write_csv(rows: list[list[str]], path: str) -> None:
    """Записывает таблицу в CSV (UTF-8 с BOM для совместимости с Excel)."""
    max_cols = max((len(r) for r in rows), default=0)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        for row in rows:
            # Выравниваем строки по максимальной ширине
            padded = row + [""] * (max_cols - len(row))
            cleaned = [clean_cell(c) for c in padded]
            writer.writerow(cleaned)


# ────────────────────────────── Main ────────────────────────────────────

def run(pdf_path: str, output_dir: str) -> None:
    pdf_path = str(Path(pdf_path).resolve())
    pdf_name = Path(pdf_path).stem
    output_dir = str(Path(output_dir).resolve())

    mineru_out = os.path.join(output_dir, "_mineru_raw")
    os.makedirs(mineru_out, exist_ok=True)

    # Кэш: если content_list.json уже создан — пропускаем долгий OCR
    expected_cl = os.path.join(mineru_out, pdf_name, "txt", f"{pdf_name}_content_list.json")
    cached_matches = list(Path(mineru_out).rglob("*_content_list.json"))
    content_list_path = str(cached_matches[0]) if cached_matches else None

    if content_list_path and os.path.exists(content_list_path):
        print(f"[1/3] Кэш найден, MinerU пропускается: {content_list_path}")
    else:
        print(f"[1/3] Запуск MinerU на '{pdf_path}' (может занять несколько минут) ...")
        pdf_bytes = read_fn(pdf_path)

        do_parse(
            output_dir=mineru_out,
            pdf_file_names=[pdf_name],
            pdf_bytes_list=[pdf_bytes],
            p_lang_list=["east_slavic"],   # русский / кириллица
            backend="pipeline",
            parse_method="txt",            # текстовый PDF (не OCR)
            formula_enable=False,          # формулы нам не нужны
            table_enable=True,
            f_draw_layout_bbox=False,
            f_draw_span_bbox=False,
            f_dump_md=False,
            f_dump_middle_json=False,
            f_dump_model_output=False,
            f_dump_orig_pdf=False,
            f_dump_content_list=True,      # нужен content_list.json
            f_make_md_mode=MakeMode.MM_MD,
        )

        # Находим content_list.json
        content_list_path = expected_cl
        if not os.path.exists(content_list_path):
            found = list(Path(mineru_out).rglob("*_content_list.json"))
            if not found:
                raise FileNotFoundError(
                    f"content_list.json не найден в {mineru_out}. "
                    "Возможно, MinerU не смог обработать PDF."
                )
            content_list_path = str(found[0])

    print(f"[2/3] Разбор content_list.json: {content_list_path}")
    form_tables = extract_tables_from_content_list(content_list_path)

    print(f"[3/3] Запись CSV в '{output_dir}' ...")
    os.makedirs(output_dir, exist_ok=True)

    found_forms: list[str] = []
    missing_forms: list[str] = []

    for form_key, tables in form_tables.items():
        name_ru = FORM_NAMES_RU[form_key]
        if not tables:
            print(f"  [!] НЕ НАЙДЕНО: {name_ru}")
            missing_forms.append(form_key)
            continue

        merged = merge_table_pages(tables)
        merged = fix_line_breaks_in_table(merged)
        csv_path = os.path.join(output_dir, f"{form_key}.csv")
        write_csv(merged, csv_path)
        print(f"  [+] {name_ru} -> {csv_path}  ({len(merged)} строк, {len(tables)} частей)")
        found_forms.append(form_key)

    print()
    if missing_forms:
        print(
            "Не найдены формы:", ", ".join(FORM_NAMES_RU[k] for k in missing_forms)
        )
        print("Совет: если PDF на основе OCR, повторите с parse_method='ocr'.")
    if found_forms:
        print(f"Готово! Найдено {len(found_forms)}/4 форм. CSV сохранены в: {output_dir}")

    # Удаляем временные файлы MinerU (images и т.п.) — оставляем только CSV
    # Раскомментируйте, если не нужны промежуточные файлы:
    # shutil.rmtree(mineru_out, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(description="Извлечение 4 форм МСФО из PDF банковского отчёта")
    parser.add_argument("pdf", help="Путь к PDF-файлу (например, sber.pdf)")
    parser.add_argument(
        "--output", "-o",
        default="ifrs_output",
        help="Папка для CSV-файлов (по умолчанию: ifrs_output)",
    )
    args = parser.parse_args()

    run(args.pdf, args.output)


if __name__ == "__main__":
    main()