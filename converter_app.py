# converter_app.py — Конвертер банковской выписки

import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

from ui_common import section_header, apply_global_css

# Дата/время релиза конвертера (обновлять при выкладке)
CONVERTER_RELEASE = "06.03.2025 10:00"


# ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====

def extract_bank_account(text: str) -> str:
    """Извлечь 20-значный счёт из строки."""
    match = re.search(r"\b\d{20}\b", str(text))
    return match.group(0) if match else ""


def extract_is_from_bailiff(text: str) -> str:
    """
    Определить, пришёл ли платёж от ФССП / УФК и т.п.
    Возвращает "Y" / "N".
    """
    txt = str(text).lower().replace("\n", " ")
    keywords = [
        "уфк", "росп", "осп", "уфссп", "гуссп", "гуфссп",
        "фссп", "фссп россии", "государственная служба судебных приставов",
    ]
    return "Y" if any(kw in txt for kw in keywords) else "N"


def determine_payment_provider(is_from_bailiff: str, payment_purpose: str) -> str:
    """
    Определить провайдера платежа:
    - если платёж от УФК (IsFromBailiff == "Y") -> "FSSP";
    - если назначение начинается с "{} 54ПБ Взыскание по ИД" -> "SB IP";
    - иначе -> "OTHER".
    """
    if str(is_from_bailiff).upper() == "Y":
        return "FSSP"

    purpose = str(payment_purpose).lstrip()
    if purpose.lower().startswith("{} 54пб взыскание по ид"):
        return "SB IP"

    return "OTHER"


def extract_court_order_number(text: str) -> str:
    """Извлечь номер судебного приказа / ИД (текущая версия с приоритетами и исключениями)."""
    text_l = str(text).lower()

    # Приоритет: ВС/ФС + 9 цифр
    priority_match = re.search(r"\b(вс|фс)\s?(\d{9})\b", text_l)
    if priority_match:
        return f"{priority_match.group(1).upper()} {priority_match.group(2)}"

    # Прямой шаблон ИД с 'ид'
    match_id_direct = re.search(r"\bид\s+([\d\-]+/\d{4}(?:-\d{1,3})?)\b", text_l)
    if match_id_direct:
        return match_id_direct.group(1)

    # Приоритет: «Судебный приказ 2-515/2025/5м», «2-4411/2024-5-8» и похожие —
    # не отбрасывать из‑за ИП в предыдущем тексте, захватывать все сегменты до пробела
    sp_match = re.search(
        r"(?:судебный приказ|суд\.? приказ|с/пр)\s*(?:№|:)?\s*"
        r"([\d\-/]+(?:[а-яёa-z][\d\-/]*)?)",
        text_l,
    )
    if sp_match:
        value = sp_match.group(1)
        if len(value.strip()) >= 5 and not re.search(r"-ип$", value):
            return value

    patterns = [
        r"№[а-яa-z]+[\d\-]*-([\d\-]+/\d{4}(?:-\d{1,3})?)",
        r"(?:судебный приказ|суд\.? приказ|с/пр)[^\d]{0,3}([\d]{1,2}-\d{1,4}-\d{1,5}/\d{4})",
        r"(?:судебный приказ|суд\.? приказ|с/пр)\s*(?:№|:)?\s*([\d\-/]+)",
        r"взыскание по ид от \d{2}\.\d{2}\.\d{4} ?№([\d\-/]+)",
        r"по и/д\s*№?\s*([\d\-/]+)",
        r"\bи/д\s*№?\s*([\d\-/]+)",
        r"(?:по\s+)?и/л\s*(?:№|n)?\s*([\d\-/]+)",
        r"\b(?:ид n|ид|n)\s*(?:№|n)?\s*([\d\-]+/\d{4}(?:-\d{1,3})?)\b",
        r"№\s*([\d\-]+/\d{4}(?:-\d{1,3})?)",
        r"суд\.пр\s*([\d\-]+/[\d\-]+)",
        r"исполнительный лист\s*([\d\-]+/\d{4})",
        r"\bил\s+([\d\-]+/\d{4})",
        r"и/л\s*(?:№|n)?\s*([\w\-]+/\d{4})",
        r"по документу\s+([\d\-]+/\d{4})",
        r"с/п\s*([\d\-]+/\d{4})",
    ]

    for pattern in patterns:
        match = re.search(pattern, text_l)
        if not match:
            continue

        value = match.group(1)
        if len(value.strip()) < 5:
            continue

        # Проверяем, что это не номер ИП
        before = text_l[: match.start()]
        if re.search(r"(ип|\bисп\w*)\s*$", before.strip()[-20:]):
            continue
        if re.search(r"-ип$", value):
            continue

        return value

    return ""


def extract_court_order_date(text: str, court_number: str) -> str:
    """Дата судебного приказа вблизи номера приказа."""
    txt = str(text).lower()
    cn = court_number.strip().lower()
    if not cn or len(cn) < 5:
        return ""

    txt_clean = re.sub(r"[()\[\]]", " ", txt)
    pos = txt_clean.find(cn)
    if pos == -1:
        return ""

    context = txt_clean[max(0, pos - 50): pos + 50]
    date_patterns = [
        r"от\s*(\d{2}\.\d{2}\.\d{4})",
        r"от\s*(\d{4}-\d{2}-\d{2})",
        r"(\d{2}\.\d{2}\.\d{4})",
        r"(\d{4}-\d{2}-\d{2})",
    ]
    for pattern in date_patterns:
        m = re.search(pattern, context)
        if m:
            return m.group(1)
    return ""


def extract_ip_number(text: str) -> str:
    """Извлечь номер ИП (исполнительного производства)."""
    t = str(text).lower()

    m1 = re.search(r"(?:и/п|ип)?[ №:]*([0-9]{4,8}/[0-9]{2}/[0-9]{4,8}-ип)\b", t)
    if m1:
        return m1.group(1)

    m2 = re.search(r"(?:и/п|ип)?[ №:]*([0-9]{4,8}/[0-9]{2}/[0-9]{4,8})\b", t)
    if m2:
        before = t[: m2.start()]
        if "ид" not in before[-20:]:
            return m2.group(1)

    m3 = re.search(r"\(ип\s+([\w\-\/]+)", t)
    if m3:
        return m3.group(1)

    return ""


def extract_fio(text: str) -> str:
    """Попытка вытащить ФИО из текста назначения."""
    txt = str(text)
    patterns = [
        r"(?:взыскано\s+)?с\s+должника\s+([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"\bс\s+([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"\bдолг[а]?:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"\bдолжника:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"долга взыскателю\s*:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
        r"\bс:\s*([А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+ [А-ЯЁ][а-яё]+)",
    ]
    for pattern in patterns:
        m = re.search(pattern, txt, flags=re.IGNORECASE)
        if m:
            return m.group(1).title().strip()

    # Доп. проверка: три подряд идущих слова, третье оканчивается на -ович/-евич/-овна/-евна
    fallback = re.search(
        r"\b([А-ЯЁа-яё]+)\s+([А-ЯЁа-яё]+)\s+([А-ЯЁа-яё]+(?:ович|евич|овна|евна))\b",
        txt,
    )
    if fallback:
        return f"{fallback.group(1)} {fallback.group(2)} {fallback.group(3)}".title().strip()
    return ""


def _patronymic_ok(word: str) -> bool:
    """Проверка окончания отчества: -ович/-евич/-овна/-евна."""
    w = (word or "").lower()
    return (
        w.endswith("ович")
        or w.endswith("евич")
        or w.endswith("овна")
        or w.endswith("евна")
    )


def extract_fio_from_debet_54pb(text: str) -> str:
    """
    ФИО для формата 54ПБ. В ячейке данные могут быть разделены переносами строки (как в Excel).
    Сначала ищем строку, содержащую «//» — берём из неё текст до «//» (перед «//» может быть пробел);
    если там ровно 3 слова и третье с окончанием -ович/-евич/-овна/-евна — это ФИО.
    Запасной вариант: весь текст до первого «//», слова 3–5 с проверкой окончания 5-го.
    """
    raw = str(text).replace("\r\n", "\n").replace("\r", "\n")

    # Приоритет: строка с «//» (когда в ячейке разделение по переносам строки)
    for line in raw.splitlines():
        line = line.strip()
        if "//" not in line:
            continue
        part = re.split(r"\s*//", line, maxsplit=1)[0].strip()
        parts_words = part.split()
        if len(parts_words) == 3 and _patronymic_ok(parts_words[2]):
            return " ".join(parts_words).title()

    # Запасной вариант: весь текст до «//», слова 3–5
    before_slash = re.split(r"\s*//", raw, maxsplit=1)[0].strip()
    words = before_slash.split()
    if len(words) >= 5 and _patronymic_ok(words[4]):
        return f"{words[2]} {words[3]} {words[4]}".title()

    return ""


def process_bank_statement(df: pd.DataFrame) -> pd.DataFrame:
    """Основная обработка выписки -> итоговая таблица."""
    # Берём только кредитовые операции
    df = df[pd.to_numeric(df["Сумма по кредиту"], errors="coerce") > 0].copy()

    res = pd.DataFrame()
    res["CaseID"] = ""
    res["TransactionType"] = "Оплата"
    res["Sum"] = df["Сумма по кредиту"]

    res["PaymentDate"] = (
        pd.to_datetime(df["Дата проводки"], errors="coerce").dt.date
    )
    res["BookingDate"] = datetime.now().date()

    res["BankAccount"] = df["Кредит"].apply(extract_bank_account)
    res["InvoiceNum"] = ""
    res["InvoiceID"] = ""

    # В части выписок колонки «Счет» нет — тогда для УФК/ФССП используем «Дебет»
    bailiff_series = df["Счет"] if "Счет" in df.columns else df["Дебет"]

    # Порядок колонок: PaymentProvider, затем IsFromBailiff. Для PaymentProvider нужен уже вычисленный IsFromBailiff ("Y"/"N")
    is_from_bailiff_series = bailiff_series.apply(extract_is_from_bailiff)
    res["PaymentProvider"] = df.apply(
        lambda row: determine_payment_provider(
            is_from_bailiff_series[row.name], row["Назначение платежа"]
        ),
        axis=1,
    )
    res["IsFromBailiff"] = is_from_bailiff_series

    res["CourtOrderNumber"] = df["Назначение платежа"].apply(
        extract_court_order_number
    )
    res["Дата приказа"] = df.apply(
        lambda row: extract_court_order_date(
            row["Назначение платежа"],
            extract_court_order_number(row["Назначение платежа"]),
        ),
        axis=1,
    )

    res["Номер ИП"] = df["Назначение платежа"].apply(extract_ip_number)

    def _get_fio(row):
        purpose = str(row["Назначение платежа"]).lstrip()
        if purpose.lower().startswith("{} 54пб взыскание по ид"):
            debet_val = row["Дебет"]
            if pd.isna(debet_val) or str(debet_val).strip().lower() in ("", "nan"):
                if "Счет" in df.columns and pd.notna(row.get("Счет")) and str(row["Счет"]).strip().lower() not in ("", "nan"):
                    return extract_fio_from_debet_54pb(row["Счет"])
                return extract_fio(row["Назначение платежа"])
            return extract_fio_from_debet_54pb(debet_val)
        return extract_fio(row["Назначение платежа"])

    res["ФИО"] = df.apply(_get_fio, axis=1)
    res["Назначение платежа"] = df["Назначение платежа"]

    return res


# ===== ОСНОВНАЯ ФУНКЦИЯ МОДУЛЯ =====

def run():
    # ← единый CSS, как на остальных вкладках!
    apply_global_css()

    section_header(
        "Конвертер банковской выписки",
        "Загрузите файл выписки. Я выделю только нужные операции и соберу таблицу...",
        release=f"Релиз: {CONVERTER_RELEASE}",
    )

    uploaded_file = st.file_uploader(
        "Загрузите файл выписки (Excel)", type=["xlsx", "xls"]
    )

    if not uploaded_file:
        return

    try:
        # Читаем выписку — как и раньше, пропуская первые 2 строки
        df_raw = pd.read_excel(uploaded_file, skiprows=2)

        # Переименовываем нужные столбцы: если в файле есть колонка с нужным именем — используем её
        def _col_by_name_or_index(df, name, fallback_idx):
            for i, c in enumerate(df.columns):
                if str(c).strip().lower() == name.lower():
                    if str(c).strip() != name:
                        df.rename(columns={df.columns[i]: name}, inplace=True)
                    return
            # Не перезаписывать колонку, если там уже «Дебет» или «Счет» (в части файлов Дебет на индексе 4)
            existing = str(df.columns[fallback_idx]).strip() if fallback_idx < len(df.columns) else ""
            if name == "Счет" and existing == "Дебет":
                return
            if name == "Дебет" and existing == "Счет":
                return
            df.columns.values[fallback_idx] = name

        _col_by_name_or_index(df_raw, "Дата проводки", 1)
        _col_by_name_or_index(df_raw, "Счет", 4)
        _col_by_name_or_index(df_raw, "Дебет", 6)  # в файле колонка может быть под именем "Дебет"
        _col_by_name_or_index(df_raw, "Кредит", 8)
        _col_by_name_or_index(df_raw, "Сумма по кредиту", 13)
        _col_by_name_or_index(df_raw, "№ документа", 14)
        _col_by_name_or_index(df_raw, "Назначение платежа", 20)

        df = df_raw.copy()

        st.success("Файл успешно загружен и распознан.")

        df_result = process_bank_statement(df)

        col1, col2 = st.columns([3, 1])
        with col1:
            st.subheader("Результат обработки")
        with col2:
            output = BytesIO()
            df_result.to_excel(output, index=False, engine="openpyxl")
            st.download_button(
                "📥 Скачать результат (Excel)",
                data=output.getvalue(),
                file_name="результат_выписки.xlsx",
            )

        st.dataframe(df_result)

    except Exception as e:
        st.error(f"Ошибка при обработке: {e}")
