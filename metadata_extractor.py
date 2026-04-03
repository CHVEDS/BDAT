"""
Извлечение названия банка и года отчётности из content_list.json (MinerU).

Алгоритм:
  1. Сканируем первые MAX_SCAN_BLOCKS текстовых блоков (обычно 3–5 страниц).
  2. Год — по специфичным паттернам с датой («31 декабря 2024 года»),
     затем по общим («за 2024 год»). Выбираем наиболее часто встречающийся год.
  3. Название банка:
     а) сначала проверяем список известных банков (точное совпадение);
     б) затем ищем юридическое название (АО/ПАО «…»);
     в) крайний случай — используем имя PDF-файла.
"""

import re
import json
from collections import Counter
from pathlib import Path


# ─────────────────── Константы ──────────────────────────────────────────

# Сколько блоков в начале документа просматривать
MAX_SCAN_BLOCKS = 60
# Сколько символов из блока использовать максимально
MAX_BLOCK_LEN   = 400


# ─────────────────── Паттерны для года ──────────────────────────────────

# Приоритетные: год рядом с конкретной датой
_YEAR_PRIORITY = [
    re.compile(r'31\s+декабря\s+(20\d{2})', re.IGNORECASE),
    re.compile(r'за\s+(?:год|период)[^.]{0,40}?31\s+декабря\s+(20\d{2})', re.IGNORECASE | re.DOTALL),
    re.compile(r'(?:по состоянию на|на дату)\s+31\s+декабря\s+(20\d{2})', re.IGNORECASE),
    re.compile(r'as\s+(?:of|at)\s+(?:31\s+december|december\s+31),?\s+(20\d{2})', re.IGNORECASE),
    re.compile(r'for\s+the\s+year\s+ended\s+(?:31\s+december|december\s+31),?\s+(20\d{2})', re.IGNORECASE),
]

# Общие: просто «за 2024 год» или «2024 года»
_YEAR_GENERAL = [
    re.compile(r'за\s+(20\d{2})\s+год', re.IGNORECASE),
    re.compile(r'(20\d{2})\s+год[ау]?\b', re.IGNORECASE),
]


# ─────────────────── Известные банки ────────────────────────────────────

# Каждая запись: (compiled_regex, каноническое_название)
# Порядок важен: более специфичные — выше
KNOWN_BANKS: list[tuple[re.Pattern, str]] = [
    (re.compile(r'газпромбанк|gazprombank',      re.IGNORECASE), 'Газпромбанк'),
    (re.compile(r'сбербанк\s+россий|sberbank',   re.IGNORECASE), 'Сбербанк'),
    (re.compile(r'сбербанк',                      re.IGNORECASE), 'Сбербанк'),
    (re.compile(r'\bвтб\b|bank\s+vtb|\bvtb\b',  re.IGNORECASE), 'ВТБ'),
    (re.compile(r'альфа.{0,3}банк|alfa.{0,3}bank', re.IGNORECASE), 'Альфа-Банк'),
    (re.compile(r'россельхозбанк|rosselkhozbank', re.IGNORECASE), 'Россельхозбанк'),
    (re.compile(r'тинькофф|tinkoff',             re.IGNORECASE), 'Тинькофф (Т-Банк)'),
    (re.compile(r'\bт-банк\b',                   re.IGNORECASE), 'Тинькофф (Т-Банк)'),
    (re.compile(r'банк\s+открытие|otkritie',     re.IGNORECASE), 'Банк Открытие'),
    (re.compile(r'промсвязьбанк|psb\b',          re.IGNORECASE), 'Промсвязьбанк'),
    (re.compile(r'совкомбанк|sovcombank',        re.IGNORECASE), 'Совкомбанк'),
    (re.compile(r'московский\s+кредитный\s+банк|\bмкб\b', re.IGNORECASE), 'МКБ'),
    (re.compile(r'росбанк|rosbank',              re.IGNORECASE), 'Росбанк'),
    (re.compile(r'райффайзен|raiffeisen',        re.IGNORECASE), 'Райффайзенбанк'),
    (re.compile(r'уралсиб|uralsib',              re.IGNORECASE), 'Уралсиб'),
    (re.compile(r'банк\s+санкт.петербург',       re.IGNORECASE), 'Банк Санкт-Петербург'),
    (re.compile(r'ренессанс.{0,5}кредит',        re.IGNORECASE), 'Ренессанс Кредит'),
    (re.compile(r'\bак\s*барс\b',               re.IGNORECASE), 'АК БАРС'),
    (re.compile(r'абсолют.{0,3}банк',            re.IGNORECASE), 'Абсолют Банк'),
    (re.compile(r'хоум\s+кредит|home\s+credit', re.IGNORECASE), 'Хоум Кредит'),
    (re.compile(r'экспобанк|expobank',           re.IGNORECASE), 'Экспобанк'),
    (re.compile(r'зенит',                        re.IGNORECASE), 'Банк Зенит'),
    (re.compile(r'кредит\s+европа\s+банк',       re.IGNORECASE), 'Кредит Европа Банк'),
    (re.compile(r'мтс.{0,5}банк',               re.IGNORECASE), 'МТС-Банк'),
    (re.compile(r'почта\s+банк',                 re.IGNORECASE), 'Почта Банк'),
    (re.compile(r'кредит\s+банк\s+москвы|кбм\b', re.IGNORECASE), 'Кредит Банк Москвы'),
    (re.compile(r'ситибанк|citibank',            re.IGNORECASE), 'Ситибанк'),
    (re.compile(r'юникредит|unicredit',          re.IGNORECASE), 'ЮниКредит Банк'),
    (re.compile(r'банк\s+россия\b',              re.IGNORECASE), 'Банк Россия'),
]

# ─────────────────── Паттерн юридического названия ───────────────────

# «АО / ПАО / ЗАО «ИМЯ БАНКА»» или «Банк ИМЯ (ПАО)»
_LEGAL_FULL_RE = re.compile(
    r'(?:П?АО|ЗАО|ООО|НКО|АКБ)\s*[«"\'"]?\s*([А-ЯЁ][А-ЯЁа-яёA-Za-z0-9\s\-]{2,40}?)\s*[»"\'"]',
    re.UNICODE,
)
# «Банк ИМЯ (публичное/закрытое акционерное...)»
_LEGAL_BANK_RE = re.compile(
    r'(?:Банк|Bank)\s+([А-ЯЁ][А-ЯЁа-яёA-Za-z0-9\s\-]{1,30}?)\s*(?:\(|,|\.|$)',
    re.UNICODE,
)


# ─────────────────── Публичные функции ───────────────────────────────

def extract_metadata(content_list_path: str, pdf_stem: str = "") -> dict:
    """
    Читает content_list.json и возвращает:
        {
            "bank_name": str,   # каноническое название банка
            "year":      str,   # год отчётности (строка, напр. "2024")
            "source":    str,   # откуда взяли название: "known" / "legal" / "filename"
        }
    """
    try:
        with open(content_list_path, encoding="utf-8") as f:
            blocks = json.load(f)
    except Exception:
        return _fallback(pdf_stem)

    texts = _collect_texts(blocks)
    combined = " ".join(texts)

    year      = _detect_year(texts)
    bank_name, source = _detect_bank(texts, combined, pdf_stem)

    return {"bank_name": bank_name, "year": year, "source": source}


# ─────────────────── Внутренние функции ─────────────────────────────

def _collect_texts(blocks: list) -> list[str]:
    """Возвращает первые MAX_SCAN_BLOCKS текстовых блоков (обрезанных)."""
    texts = []
    for block in blocks:
        if block.get("type") != "text":
            continue
        txt = block.get("text", "").strip()
        if txt:
            texts.append(txt[:MAX_BLOCK_LEN])
        if len(texts) >= MAX_SCAN_BLOCKS:
            break
    return texts


def _detect_year(texts: list[str]) -> str:
    """
    Возвращает наиболее вероятный год отчётности.
    Приоритет: паттерны с «31 декабря», затем «за YYYY год».
    """
    candidates: list[str] = []

    for txt in texts:
        for pat in _YEAR_PRIORITY:
            m = pat.search(txt)
            if m:
                candidates.append(m.group(1))

    # Если по приоритетным паттернам нашли — берём самый частый
    if candidates:
        return Counter(candidates).most_common(1)[0][0]

    # Общие паттерны
    for txt in texts:
        for pat in _YEAR_GENERAL:
            for m in pat.finditer(txt):
                candidates.append(m.group(1))

    if candidates:
        return Counter(candidates).most_common(1)[0][0]

    return ""


def _detect_bank(texts: list[str], combined: str, pdf_stem: str) -> tuple[str, str]:
    """
    Возвращает (название_банка, источник).
    Источник: 'known' | 'legal' | 'filename'
    """
    # 1. Известные банки — по всем текстам первых страниц
    counts: dict[str, int] = Counter()
    for txt in texts:
        for pattern, canon_name in KNOWN_BANKS:
            if pattern.search(txt):
                counts[canon_name] += 1

    # Берём самое часто встречающееся известное название
    if counts:
        best = counts.most_common(1)[0][0]
        return best, "known"

    # 2. Юридическое название: АО/ПАО «…»
    legal_candidates: list[str] = []
    for txt in texts[:20]:  # только начало документа
        for m in _LEGAL_FULL_RE.finditer(txt):
            name = m.group(1).strip()
            # Отфильтровываем слишком короткие и нерелевантные
            if len(name) >= 3 and not re.search(r'отчёт|отчет|финансов|консолид', name, re.IGNORECASE):
                legal_candidates.append(_normalize_name(name))

        for m in _LEGAL_BANK_RE.finditer(txt):
            name = m.group(1).strip()
            if len(name) >= 3:
                legal_candidates.append(_normalize_name(name))

    if legal_candidates:
        best = Counter(legal_candidates).most_common(1)[0][0]
        # Убираем слишком длинные (скорее всего захватили лишнее)
        if len(best) <= 50:
            return best, "legal"

    # 3. Крайний случай — из имени файла
    if pdf_stem:
        return _normalize_name(pdf_stem), "filename"

    return "Неизвестный банк", "filename"


def _normalize_name(name: str) -> str:
    """Нормализует название: убирает лишние пробелы, исправляет капитализацию."""
    name = re.sub(r'\s+', ' ', name).strip()
    # Если всё в верхнем регистре — приводим к Title Case
    if name == name.upper() and len(name) > 3:
        name = name.title()
    return name


def _fallback(pdf_stem: str) -> dict:
    return {
        "bank_name": _normalize_name(pdf_stem) if pdf_stem else "Неизвестный банк",
        "year": "",
        "source": "filename",
    }
