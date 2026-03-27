"""
МСФО-экстрактор — Streamlit-интерфейс
Корпоративный стиль (синяя гамма)

Запуск:
    streamlit run app_ifrs.py
"""

import io
import os
import sys
import contextlib
import tempfile
import zipfile
from pathlib import Path
from collections import defaultdict

import streamlit as st
import pandas as pd

import extract_ifrs
import metadata_extractor as meta_ext

# ─────────────────────────── Стили ──────────────────────────────────────

C_BLUE      = "#002B6E"   # основной синий (тёмный) — белый текст поверх
C_BLUE2     = "#0047AB"   # средний синий
C_BLUE_LT   = "#C8DAEA"   # светло-голубой фон (только для акцентов)
C_BORDER    = "#9DBDD6"   # граница элементов
C_BG        = "#B0E0E6"   # фон страницы (заметно отличается от карточек)
C_WHITE     = "#4682B4"   # фон карточек
C_TEXT      = "#FFFFFF"   # основной текст — максимальный контраст на белом
C_TEXT_2    = "#2D3748"   # вторичный текст — тёмно-серый (8:1 на белом)
C_TEXT_MUTE = "#4A5568"   # приглушённый текст (7:1 на белом)
C_GREEN     = "#145A32"   # зелёный текст (11:1 на зелёном фоне)
C_GREEN_BG  = "#C3E6CB"   # зелёный фон (для зелёного текста)
C_RED       = "#7B1E1E"   # красный текст (10:1 на красном фоне)
C_RED_BG    = "#F5C6CB"   # красный фон
C_YELLOW_BG = "#FFEEBA"   # жёлтый фон
C_YELLOW    = "#6B4A00"   # тёмно-жёлтый текст (9:1 на жёлтом фоне)

CSS = f"""
<style>
/* ── Base: явные цвета для всех текстовых элементов ── */
.stApp {{
    background-color: {C_BG};
}}
/* Streamlit текст — явный тёмный цвет */
.stApp p,
.stApp span,
.stApp label,
.stApp li,
.stApp div,
.stApp small,
.stMarkdown p,
.stMarkdown li,
.stMarkdown span,
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stText"] {{
    color: {C_TEXT} !important;
    font-family: "Arial", "Helvetica Neue", sans-serif;
}}
/* Заголовки */
.stApp h1, .stApp h2, .stApp h3, .stApp h4 {{
    color: {C_BLUE} !important;
    font-family: "Arial", "Helvetica Neue", sans-serif;
}}
/* Подписи к виджетам */
.stApp .stSelectbox label,
.stApp .stRadio label,
.stApp .stMultiSelect label,
.stApp .stFileUploader label {{
    color: {C_TEXT} !important;
    font-weight: 600;
}}
/* Текст внутри radio/checkbox/select */
.stApp .stRadio div[role="radiogroup"] label,
.stApp .stCheckbox label {{
    color: {C_TEXT} !important;
}}
/* Экспандер */
.stApp .streamlit-expanderHeader {{
    color: {C_TEXT} !important;
    font-weight: 600;
    background: {C_WHITE};
    border: 1px solid {C_BORDER};
    border-radius: 6px;
}}

/* ── Header ── */
.app-header {{
    background: {C_BLUE};
    padding: 20px 32px 16px;
    border-radius: 0 0 10px 10px;
    margin-bottom: 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.18);
}}
.app-header h1 {{
    color: #FFFFFF !important;
    font-size: 20px;
    font-weight: 700;
    margin: 0;
    line-height: 1.3;
    font-family: "Arial", sans-serif;
}}
.app-header p {{
    color: #B8CDE8 !important;
    font-size: 13px;
    margin: 4px 0 0;
    font-family: "Arial", sans-serif;
}}

/* ── Карточки ── */
.card {{
    background: {C_WHITE};
    border: 1px solid {C_BORDER};
    border-radius: 8px;
    padding: 18px 20px;
    margin-bottom: 12px;
    box-shadow: 0 1px 4px rgba(0,43,110,0.08);
}}
.card-header {{
    font-size: 13px;
    font-weight: 700;
    color: {C_BLUE} !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 10px;
    padding-bottom: 6px;
    border-bottom: 2px solid {C_BLUE_LT};
    font-family: "Arial", sans-serif;
}}

/* ── Стат-плитки ── */
.stat-tile {{
    background: {C_WHITE};
    border: 1px solid {C_BORDER};
    border-top: 3px solid {C_BLUE2};
    border-radius: 8px;
    padding: 14px 16px;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,43,110,0.07);
}}
.stat-tile .val {{
    font-size: 28px;
    font-weight: 700;
    color: {C_BLUE} !important;
    line-height: 1.1;
    font-family: "Arial", sans-serif;
}}
.stat-tile .lbl {{
    font-size: 12px;
    font-weight: 600;
    color: {C_TEXT_2} !important;
    margin-top: 3px;
    font-family: "Arial", sans-serif;
}}
.stat-tile .sub {{
    font-size: 11px;
    margin-top: 6px;
}}

/* ── Бейджи: тёмный текст на цветном фоне ── */
.badge-ok  {{
    background: {C_GREEN_BG};
    color: {C_GREEN} !important;
    border-radius: 10px;
    padding: 2px 9px;
    font-size: 11px;
    font-weight: 700;
    white-space: nowrap;
    font-family: "Arial", sans-serif;
}}
.badge-err {{
    background: {C_RED_BG};
    color: {C_RED} !important;
    border-radius: 10px;
    padding: 2px 9px;
    font-size: 11px;
    font-weight: 700;
    white-space: nowrap;
    font-family: "Arial", sans-serif;
}}
.badge-warn {{
    background: {C_YELLOW_BG};
    color: {C_YELLOW} !important;
    border-radius: 10px;
    padding: 2px 9px;
    font-size: 11px;
    font-weight: 700;
    white-space: nowrap;
    font-family: "Arial", sans-serif;
}}

/* ── Панель заголовка формы ── */
.form-bar {{
    background: {C_BLUE};
    color: #FFFFFF !important;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 14px;
    border-radius: 6px 6px 0 0;
    margin-bottom: 0;
    font-family: "Arial", sans-serif;
}}

/* ── Блок информации о банке ── */
.bank-info {{
    background: {C_WHITE};
    border: 1px solid {C_BORDER};
    border-left: 5px solid {C_BLUE2};
    border-radius: 0 6px 6px 0;
    padding: 12px 18px;
    margin-bottom: 16px;
    box-shadow: 0 1px 4px rgba(0,43,110,0.08);
}}
.bank-info .name {{
    font-size: 18px;
    font-weight: 700;
    color: {C_BLUE} !important;
    font-family: "Arial", sans-serif;
}}
.bank-info .year {{
    font-size: 13px;
    color: {C_TEXT_2} !important;
    margin-top: 4px;
    font-family: "Arial", sans-serif;
}}

/* ── Sidebar: белый фон, тёмный текст ── */
section[data-testid="stSidebar"] {{
    background: {C_WHITE};
    border-right: 2px solid {C_BLUE_LT};
}}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] div,
section[data-testid="stSidebar"] small {{
    color: {C_TEXT} !important;
}}
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {{
    color: {C_BLUE} !important;
}}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
    border-bottom: 2px solid {C_BLUE_LT};
    background: transparent;
}}
.stTabs [data-baseweb="tab"] {{
    font-size: 13px;
    font-weight: 600;
    color: {C_TEXT_2} !important;
    background: #EBF2F9;
    border: 1px solid {C_BORDER};
    border-bottom: none;
    border-radius: 6px 6px 0 0;
    padding: 6px 14px;
}}
.stTabs [aria-selected="true"] {{
    color: {C_BLUE} !important;
    border-color: {C_BLUE} !important;
    background: {C_WHITE} !important;
}}
.stTabs [data-baseweb="tab-highlight"] {{
    background-color: {C_BLUE} !important;
}}
/* Содержимое вкладки — белый фон */
.stTabs [data-baseweb="tab-panel"] {{
    background: {C_WHITE};
    border: 1px solid {C_BLUE_LT};
    border-top: none;
    border-radius: 0 0 8px 8px;
    padding: 14px;
}}

/* ── Кнопки ── */
div.stButton > button {{
    border-radius: 6px;
    font-weight: 600;
    font-size: 13px;
    padding: 8px 20px;
    transition: background-color 0.2s;
}}
div.stButton > button[kind="primary"] {{
    background: {C_BLUE} !important;
    color: #FFFFFF !important;
    border: none;
}}
div.stButton > button[kind="primary"]:hover {{
    background: {C_BLUE2} !important;
}}
div.stButton > button[kind="secondary"] {{
    background: {C_WHITE} !important;
    color: {C_BLUE} !important;
    border: 2px solid {C_BLUE};
}}

/* ── Кнопки скачивания ── */
div[data-testid="stDownloadButton"] > button {{
    background: {C_BLUE} !important;
    color: #FFFFFF !important;
    border: none;
    border-radius: 6px;
    font-weight: 600;
    font-size: 13px;
    padding: 8px 18px;
}}
div[data-testid="stDownloadButton"] > button:hover {{
    background: {C_BLUE2} !important;
}}

/* ── Таблицы ── */
.stDataFrame {{
    border: 1px solid {C_BORDER};
    border-radius: 6px;
    overflow: hidden;
}}
/* Текст внутри таблиц — тёмный */
.stDataFrame td, .stDataFrame th {{
    color: {C_TEXT} !important;
}}

/* ── Сообщения (info / warning / success / error) ── */
div[data-testid="stAlert"] {{
    border-radius: 6px;
}}

/* ── Expander ── */
details summary {{
    color: {C_TEXT} !important;
    font-weight: 600;
}}

/* ── Разделитель ── */
hr {{ border-top: 2px solid {C_BLUE_LT}; margin: 20px 0; }}

/* ── Заголовок таблицы сравнения ── */
.cmp-header {{
    background: {C_BLUE};
    color: #FFFFFF !important;
    font-size: 14px;
    font-weight: 700;
    padding: 10px 16px;
    border-radius: 8px 8px 0 0;
    margin-bottom: 0;
    font-family: "Arial", sans-serif;
}}
</style>
"""

# ─────────────────────────── Константы ──────────────────────────────────

FORM_ORDER = ["balance_sheet", "income_statement", "cash_flow", "equity_changes"]
FORM_NAMES: dict[str, str] = {
    "balance_sheet":    "Отчёт о финансовом положении",
    "income_statement": "Отчёт о прибылях и убытках",
    "cash_flow":        "Отчёт о движении денежных средств",
    "equity_changes":   "Отчёт об изменениях в собственном капитале",
}
FORM_SHORT: dict[str, str] = {
    "balance_sheet":    "Баланс",
    "income_statement": "П&У",
    "cash_flow":        "ДДС",
    "equity_changes":   "Капитал",
}
FORM_ICONS: dict[str, str] = {
    "balance_sheet":    "📊",
    "income_statement": "📈",
    "cash_flow":        "💵",
    "equity_changes":   "🏛",
}


# ─────────────────────────── Утилиты ────────────────────────────────────

@contextlib.contextmanager
def _capture_stdout():
    buf = io.StringIO()
    old = sys.stdout
    sys.stdout = buf
    try:
        yield buf
    finally:
        sys.stdout = old


def _read_df(csv_path: str) -> pd.DataFrame:
    try:
        return pd.read_csv(csv_path, encoding="utf-8-sig", header=None, dtype=str).fillna("")
    except Exception:
        return pd.DataFrame()


def _find_content_list(output_dir: str) -> str | None:
    found = list(Path(output_dir).rglob("*_content_list.json"))
    return str(found[0]) if found else None


def _build_zip(bank_data: dict) -> bytes:
    """Собирает ZIP со всеми результатами всех банков."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for bank_key, data in bank_data.items():
            folder = data["pdf_stem"]
            # XLSX
            if data.get("xlsx_bytes"):
                zf.writestr(f"{folder}/{folder}_ifrs.xlsx", data["xlsx_bytes"])
            # CSV
            for fk in FORM_ORDER:
                df: pd.DataFrame = data["results"].get(fk, {}).get("df", pd.DataFrame())
                if not df.empty:
                    csv_str = df.to_csv(index=False, header=False, encoding="utf-8-sig")
                    zf.writestr(f"{folder}/{folder}_{fk}.csv", csv_str.encode("utf-8-sig"))
    buf.seek(0)
    return buf.read()


# ─────────────────────────── Обработка PDF ──────────────────────────────

def process_pdf(uploaded_file, output_base: str, status_container) -> dict | None:
    """
    Обрабатывает один PDF. Возвращает dict с результатами или None при ошибке.
    status_container — st.empty() для вывода шагов.
    """
    pdf_stem = Path(uploaded_file.name).stem

    # Создаём папку для этого PDF
    out_dir = os.path.join(output_base, pdf_stem)
    os.makedirs(out_dir, exist_ok=True)

    # Сохраняем PDF
    pdf_path = os.path.join(out_dir, uploaded_file.name)
    with open(pdf_path, "wb") as f:
        f.write(uploaded_file.getvalue())

    def _step(text: str):
        status_container.markdown(text)

    _step(f"**⏳ [1/3]** Запуск MinerU — разбор PDF `{uploaded_file.name}`...")

    log_text = ""
    try:
        log_buf = io.StringIO()
        with contextlib.redirect_stdout(log_buf):
            extract_ifrs.run(pdf_path, out_dir)
        log_text = log_buf.getvalue()
    except Exception as exc:
        _step(f"❌ **Ошибка обработки:** {exc}")
        return None

    _step("**⏳ [2/3]** Определение названия банка и года...")

    # Ищем content_list.json для метаданных
    cl_path = _find_content_list(out_dir)
    if cl_path:
        meta = meta_ext.extract_metadata(cl_path, pdf_stem)
    else:
        meta = {"bank_name": pdf_stem, "year": "", "source": "filename"}

    _step("**⏳ [3/3]** Загрузка результатов...")

    # Читаем результаты
    results: dict[str, dict] = {}
    for fk in FORM_ORDER:
        csv_path = os.path.join(out_dir, f"{fk}.csv")
        df = _read_df(csv_path) if os.path.exists(csv_path) else pd.DataFrame()
        results[fk] = {"rows": len(df), "csv": csv_path if df.shape[0] else None, "df": df}

    # Читаем XLSX
    xlsx_path = os.path.join(out_dir, f"{pdf_stem}_ifrs.xlsx")
    xlsx_bytes = open(xlsx_path, "rb").read() if os.path.exists(xlsx_path) else None

    _step(f"✅ **Готово:** `{uploaded_file.name}`")

    return {
        "pdf_stem":  pdf_stem,
        "filename":  uploaded_file.name,
        "bank_name": meta["bank_name"],
        "year":      meta["year"],
        "meta_src":  meta["source"],
        "results":   results,
        "xlsx_bytes": xlsx_bytes,
        "log":       log_text,
    }


# ─────────────────────────── Компоненты UI ──────────────────────────────

def render_header():
    st.markdown("""
    <div class="app-header">
        <h1>🏦 МСФО Экстрактор</h1>
        <p>Автоматическое извлечение финансовых форм из PDF-отчётов банков</p>
    </div>
    """, unsafe_allow_html=True)


def render_bank_info(data: dict):
    src_badge = {
        "known":    '<span class="badge-ok">✓ Определён точно</span>',
        "legal":    '<span class="badge-warn">~ Из юр. названия</span>',
        "filename": '<span class="badge-err">⚠ Из имени файла</span>',
    }.get(data.get("meta_src", ""), "")

    year_str = f"Год отчётности: <b>{data['year']}</b>" if data.get("year") else \
               '<span style="color:#999">Год не определён</span>'

    st.markdown(f"""
    <div class="bank-info">
        <div class="name">{data['bank_name']} &nbsp;{src_badge}</div>
        <div class="year">{year_str} &nbsp;·&nbsp; Файл: {data['filename']}</div>
    </div>
    """, unsafe_allow_html=True)


def render_stats(results: dict):
    cols = st.columns(4)
    for i, fk in enumerate(FORM_ORDER):
        r = results.get(fk, {})
        rows = r.get("rows", 0)
        found = rows > 0
        badge = '<span class="badge-ok">Найдено</span>' if found \
                else '<span class="badge-err">Не найдено</span>'
        with cols[i]:
            st.markdown(f"""
            <div class="stat-tile">
                <div style="font-size:24px">{FORM_ICONS[fk]}</div>
                <div class="val">{rows if found else "—"}</div>
                <div class="lbl">{FORM_SHORT[fk]}</div>
                <div class="sub">{badge}</div>
            </div>
            """, unsafe_allow_html=True)


def render_forms_tabs(results: dict, key_prefix: str = ""):
    tabs = st.tabs([f"{FORM_ICONS[fk]} {FORM_SHORT[fk]}" for fk in FORM_ORDER])
    for tab, fk in zip(tabs, FORM_ORDER):
        with tab:
            r = results.get(fk, {})
            df: pd.DataFrame = r.get("df", pd.DataFrame())

            st.markdown(f"""
            <div class="form-bar">{FORM_ICONS[fk]} {FORM_NAMES[fk]}
            {"&nbsp;·&nbsp; " + str(len(df)) + " строк" if not df.empty else ""}
            </div>
            """, unsafe_allow_html=True)

            if df.empty:
                st.info("Форма не найдена в документе.")
                continue

            display = df.copy()
            display.columns = [f"{'Показатель' if j == 0 else f'Кол. {j}'}"
                                for j in range(len(display.columns))]
            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True,
                height=min(600, 36 * len(display) + 40),
                key=f"tbl_{key_prefix}_{fk}",
            )


def render_download(data: dict, key_sfx: str = ""):
    st.markdown("#### ⬇ Скачать")
    col1, col2 = st.columns([1, 1])

    with col1:
        if data.get("xlsx_bytes"):
            st.download_button(
                label="📊 Excel (все формы)",
                data=data["xlsx_bytes"],
                file_name=f"{data['pdf_stem']}_ifrs.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_xl_{key_sfx}",
            )
        else:
            st.warning("Excel не создан")

    with col2:
        for fk in FORM_ORDER:
            r = data["results"].get(fk, {})
            df: pd.DataFrame = r.get("df", pd.DataFrame())
            if not df.empty:
                csv_bytes = df.to_csv(index=False, header=False).encode("utf-8-sig")
                st.download_button(
                    label=f"📄 CSV · {FORM_SHORT[fk]}",
                    data=csv_bytes,
                    file_name=f"{data['pdf_stem']}_{fk}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key=f"dl_csv_{key_sfx}_{fk}",
                )


def render_comparison(bank_data: dict):
    """Сводная таблица сравнения нескольких банков."""
    st.markdown('<div class="cmp-header">📊 Сравнение отчётов</div>', unsafe_allow_html=True)

    rows = []
    for bank_key, data in bank_data.items():
        row = {
            "Банк":        data["bank_name"],
            "Год":         data["year"] or "—",
            "Файл":        data["filename"],
        }
        for fk in FORM_ORDER:
            r = data["results"].get(fk, {})
            n = r.get("rows", 0)
            row[FORM_SHORT[fk]] = f"✓ {n} стр." if n > 0 else "✗"
        rows.append(row)

    df_cmp = pd.DataFrame(rows)
    st.dataframe(df_cmp, use_container_width=True, hide_index=True)


def render_export_section(bank_data: dict):
    """Секция экспорта — аналог reference app."""
    st.markdown("---")
    st.markdown("### 📥 Экспорт результатов")

    with st.expander("⚙ Настройки экспорта", expanded=True):
        col_left, col_right = st.columns(2)

        with col_left:
            report_type = st.radio(
                "Тип отчёта:",
                [
                    "📋 Индивидуальный по банку",
                    "📊 Сводная таблица (все банки)",
                    "📦 Полный архив (ZIP)",
                ],
                index=0,
            )

        with col_right:
            bank_keys = list(bank_data.keys())
            if report_type == "📋 Индивидуальный по банку" and len(bank_keys) > 1:
                selected_keys = st.multiselect(
                    "Банки:",
                    options=bank_keys,
                    format_func=lambda k: bank_data[k]["bank_name"],
                    default=bank_keys,
                )
            else:
                selected_keys = bank_keys

            format_opts = st.multiselect(
                "Форматы:",
                [".xlsx", ".csv"],
                default=[".xlsx"],
                disabled=(report_type == "📦 Полный архив (ZIP)"),
            )

    if st.button("📊 Сгенерировать", type="primary"):
        if report_type == "📦 Полный архив (ZIP)":
            zip_bytes = _build_zip({k: bank_data[k] for k in selected_keys})
            st.download_button(
                "⬇ Скачать ZIP-архив",
                data=zip_bytes,
                file_name="ifrs_export.zip",
                mime="application/zip",
                key="dl_zip_final",
            )

        elif report_type == "📊 Сводная таблица (все банки)":
            rows = []
            for k in selected_keys:
                d = bank_data[k]
                row = {"Банк": d["bank_name"], "Год": d["year"] or "—"}
                for fk in FORM_ORDER:
                    n = d["results"].get(fk, {}).get("rows", 0)
                    row[FORM_NAMES[fk]] = n if n > 0 else 0
                rows.append(row)
            df_sum = pd.DataFrame(rows)

            if ".xlsx" in format_opts:
                buf = io.BytesIO()
                df_sum.to_excel(buf, index=False, engine="openpyxl")
                st.download_button("⬇ Сводная таблица (XLSX)", buf.getvalue(),
                                   "summary.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_sum_xl")
            if ".csv" in format_opts:
                st.download_button("⬇ Сводная таблица (CSV)", df_sum.to_csv(index=False).encode(),
                                   "summary.csv", "text/csv", key="dl_sum_csv")

        else:  # индивидуальный
            for k in selected_keys:
                d = bank_data[k]
                st.markdown(f"**{d['bank_name']}** ({d['year'] or '—'})")
                _render_individual_exports(d, k, format_opts)


def _render_individual_exports(data: dict, key_sfx: str, format_opts: list):
    cols = st.columns(len(FORM_ORDER) + 1)
    with cols[0]:
        if ".xlsx" in format_opts and data.get("xlsx_bytes"):
            st.download_button(
                "📊 XLSX",
                data["xlsx_bytes"],
                f"{data['pdf_stem']}_ifrs.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"exp_xl_{key_sfx}",
                use_container_width=True,
            )
    if ".csv" in format_opts:
        for i, fk in enumerate(FORM_ORDER):
            df: pd.DataFrame = data["results"].get(fk, {}).get("df", pd.DataFrame())
            with cols[i + 1]:
                if not df.empty:
                    st.download_button(
                        f"{FORM_SHORT[fk]}",
                        df.to_csv(index=False, header=False).encode("utf-8-sig"),
                        f"{data['pdf_stem']}_{fk}.csv",
                        "text/csv",
                        key=f"exp_csv_{key_sfx}_{fk}",
                        use_container_width=True,
                    )
                else:
                    st.button(f"✗ {FORM_SHORT[fk]}", disabled=True,
                              key=f"exp_na_{key_sfx}_{fk}", use_container_width=True)


# ─────────────────────────── Сайдбар ────────────────────────────────────

def render_sidebar() -> str:
    with st.sidebar:
        st.markdown(f"""
        <div style="background:{C_BLUE};color:white;padding:14px 16px;
                    border-radius:8px;margin-bottom:16px">
            <b style="font-size:15px">⚙ Параметры</b>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("**Метод разбора PDF**")
        parse_method = st.selectbox(
            "Метод:",
            ["txt — текстовый PDF (быстро)", "ocr — сканированный PDF (медленно)"],
            label_visibility="collapsed",
        )

        st.markdown("---")
        st.markdown("**Формы МСФО**")
        for fk in FORM_ORDER:
            st.markdown(f"<small>{FORM_ICONS[fk]} {FORM_NAMES[fk]}</small>",
                        unsafe_allow_html=True)

        st.markdown("---")
        if st.button("🗑 Очистить все данные", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

        st.markdown("---")
        st.caption("Использует MinerU · openpyxl")

    return "ocr" if "ocr" in parse_method else "txt"


# ─────────────────────────────── Main ───────────────────────────────────

def main():
    st.set_page_config(
        page_title="МСФО Экстрактор",
        page_icon="🏦",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(CSS, unsafe_allow_html=True)

    render_header()
    render_sidebar()

    # ── Постоянная папка для сессии ──────────────────────────────────
    if "session_dir" not in st.session_state:
        st.session_state.session_dir = tempfile.mkdtemp(prefix="ifrs_")
    if "bank_data" not in st.session_state:
        st.session_state.bank_data = {}  # key=pdf_stem → данные

    session_dir = st.session_state.session_dir
    bank_data: dict = st.session_state.bank_data

    # ── Загрузчик файлов ─────────────────────────────────────────────
    st.markdown("### 📂 Загрузка файлов")
    uploaded_files = st.file_uploader(
        "Перетащите PDF-отчёты или нажмите «Browse files»",
        type=["pdf"],
        accept_multiple_files=True,
        help="Поддерживается загрузка нескольких банковских отчётов",
        label_visibility="collapsed",
    )

    if not uploaded_files and not bank_data:
        st.markdown("""
        <div class="card" style="text-align:center;padding:36px;color:#5A6478">
            <div style="font-size:48px">📄</div>
            <div style="font-size:16px;font-weight:600;margin-top:8px">
                Загрузите PDF-отчёты для анализа
            </div>
            <div style="font-size:13px;margin-top:6px">
                Поддерживаются МСФО-отчёты российских банков в формате PDF
            </div>
        </div>
        """, unsafe_allow_html=True)
        return

    # ── Кнопка запуска ──────────────────────────────────────────────
    col_run, col_space = st.columns([2, 3])
    with col_run:
        run_clicked = st.button("🚀 Извлечь данные", type="primary",
                                use_container_width=True,
                                disabled=not uploaded_files)

    # ── Обработка ───────────────────────────────────────────────────
    if run_clicked and uploaded_files:
        progress = st.progress(0)
        total = len(uploaded_files)

        for idx, uf in enumerate(uploaded_files):
            stem = Path(uf.name).stem

            with st.container():
                st.markdown(f"""
                <div class="card">
                <div class="card-header">Обработка файла {idx+1}/{total}: {uf.name}</div>
                """, unsafe_allow_html=True)
                status_el = st.empty()

                result = process_pdf(uf, session_dir, status_el)

                if result:
                    bank_data[stem] = result
                    status_el.markdown(
                        f'✅ Завершено — найдено форм: '
                        f'**{sum(1 for r in result["results"].values() if r["rows"] > 0)}/4**'
                    )
                else:
                    status_el.error("Обработка не выполнена")

                st.markdown("</div>", unsafe_allow_html=True)

            progress.progress((idx + 1) / total)

        progress.progress(1.0)
        st.success(f"✅ Обработано файлов: {total}")

    # ── Результаты ──────────────────────────────────────────────────
    if not bank_data:
        return

    st.markdown("---")
    st.markdown("### 📋 Результаты")

    # Сравнение если несколько банков
    if len(bank_data) > 1:
        render_comparison(bank_data)
        st.markdown("---")

    # Детали по каждому банку
    if len(bank_data) > 1:
        bank_tabs = st.tabs([
            f"{FORM_ICONS['balance_sheet']} {d['bank_name']} {d['year'] or ''}"
            for d in bank_data.values()
        ])
    else:
        bank_tabs = [st.container()]

    for tab, (stem, data) in zip(bank_tabs, bank_data.items()):
        with tab:
            render_bank_info(data)
            render_stats(data["results"])
            render_forms_tabs(data["results"], key_prefix=stem)
            st.markdown("---")
            render_download(data, key_sfx=stem)

            with st.expander("📋 Журнал обработки"):
                log = data.get("log", "")
                st.code(log if log else "(нет данных)", language=None)

    # Экспорт
    render_export_section(bank_data)

    # Сброс сессии
    st.markdown("---")
    col_fin, _ = st.columns([1, 3])
    with col_fin:
        if st.button("🔄 Начать новый анализ", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()


if __name__ == "__main__":
    main()
