
# percent395_app.py — Расчет процентов по ст. 395 ГК РФ

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import List, Tuple, Dict, Any

import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet


# ----------------------------
# Helpers
# ----------------------------

def _register_cyrillic_font():
    """Register a font that supports Cyrillic (best effort)."""
    try:
        pdfmetrics.registerFont(TTFont("DejaVuSans", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
        return "DejaVuSans"
    except Exception:
        return "Helvetica"


def _to_date(x) -> date:
    if isinstance(x, date) and not isinstance(x, datetime):
        return x
    if isinstance(x, datetime):
        return x.date()
    return pd.to_datetime(x).date()


def _days_inclusive(d1: date, d2: date) -> int:
    return (d2 - d1).days + 1


def _fmt_money_ru(x: float) -> str:
    # 15000.00 -> "15 000,00"
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return s


def _fmt_percent_ru(rate_decimal: float) -> str:
    # 0.21 -> "21,00%"
    return f"{rate_decimal*100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", " ")


def _safe_filename(s: str) -> str:
    s = str(s).strip()
    s = re.sub(r'[\\/:*?"<>|]+', "_", s)
    return s or "file"


@dataclass
class RatePeriod:
    start: date
    end: date
    days_in_year: int
    rate: float  # decimal, e.g. 0.21


def _load_rate_periods(df_rate: pd.DataFrame) -> List[RatePeriod]:
    # Expect columns: 'С', 'По', 'дней в году', 'ставка, %' (maybe with NBSP)
    col_map = {c: c.strip().replace("\xa0", " ") for c in df_rate.columns}
    df_rate = df_rate.rename(columns=col_map)

    # tolerant matching
    def pick(name_variants):
        for v in name_variants:
            if v in df_rate.columns:
                return v
        # contains match
        for c in df_rate.columns:
            for v in name_variants:
                if v.lower() in c.lower():
                    return c
        raise KeyError(f"Не найдена колонка: {name_variants}")

    c_from = pick(["С", "C", "с"])
    c_to = pick(["По", "по"])
    c_days_year = pick(["дней в году", "дней_в_году"])
    c_rate = pick(["ставка, %", "ставка,%", "ставка", "ставка %"])

    periods: List[RatePeriod] = []
    for _, r in df_rate.iterrows():
        if pd.isna(r[c_from]) or pd.isna(r[c_to]) or pd.isna(r[c_days_year]) or pd.isna(r[c_rate]):
            continue
        periods.append(
            RatePeriod(
                start=_to_date(r[c_from]),
                end=_to_date(r[c_to]),
                days_in_year=int(r[c_days_year]),
                rate=float(r[c_rate]),
            )
        )
    periods.sort(key=lambda p: p.start)
    return periods


def _calc_395_for_sum(
    principal: float,
    date_from: date,
    date_to: date,
    periods: List[RatePeriod],
) -> Tuple[float, List[Dict[str, Any]]]:
    """
    Return: (total_interest, rows_for_pdf)
    """
    rows: List[Dict[str, Any]] = []
    total = 0.0

    for p in periods:
        seg_start = max(date_from, p.start)
        seg_end = min(date_to, p.end)
        if seg_start > seg_end:
            continue

        days = _days_inclusive(seg_start, seg_end)
        interest = principal * days / p.days_in_year * p.rate
        total += interest

        rows.append(
            {
                "start": seg_start,
                "end": seg_end,
                "days": days,
                "rate": p.rate,
                "days_in_year": p.days_in_year,
                "principal": principal,
                "interest": interest,
            }
        )

    return round(total + 1e-9, 2), rows


def _build_pdf_bytes(
    contract_no: str,
    contract_date: date | None,
    fio: str | None,
    as_of: date,
    calc_rows: List[Dict[str, Any]],
    total_interest: float,
) -> bytes:
    font = _register_cyrillic_font()
    styles = getSampleStyleSheet()
    base = ParagraphStyle(
        "base",
        parent=styles["Normal"],
        fontName=font,
        fontSize=10,
        leading=12,
    )
    title = ParagraphStyle(
        "title",
        parent=base,
        fontSize=11,
        leading=14,
        spaceAfter=6,
    )

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=15*mm,
        rightMargin=15*mm,
        topMargin=15*mm,
        bottomMargin=15*mm,
        title=f"395_{contract_no}",
    )

    cd = f"{contract_date:%d.%m.%Y}" if isinstance(contract_date, date) else ""
    fio_txt = fio or ""

    header_text = (
        "Расчет процентов за неправомерное пользование чужими денежными средствами "
        "по периодам действия ключевой ставки ЦБ РФ по номеру договора "
        f"№{contract_no}"
        + (f" от {cd}" if cd else "")
        + (f" {fio_txt}" if fio_txt else "")
        + f" на {as_of:%d.%m.%Y} г."
    )

    elements = [
        Paragraph(header_text, title),
        Spacer(1, 4*mm),
    ]

    # Table header
    data = [[
        "№",
        "Период\nпросрочки c",
        "Период\nпросрочки по",
        "Коли\nчеств\nо\nдней",
        "Ставка в\n%",
        "Сумма платежа",
        "Дата платежа",
        "Основной долг",
        "Формула",
        "Сумма\nпроцентов",
    ]]

    for i, r in enumerate(calc_rows, 1):
        principal = float(r["principal"])
        days = int(r["days"])
        rate = float(r["rate"])
        diy = int(r["days_in_year"])
        interest = float(r["interest"])

        formula = f"{principal:.2f}*{days}*1/{diy}*{_fmt_percent_ru(rate)}"
        data.append([
            str(i),
            r["start"].strftime("%d.%m.%Y"),
            r["end"].strftime("%d.%m.%Y"),
            str(days),
            _fmt_percent_ru(rate),
            "-",   # Сумма платежа
            "-",   # Дата платежа
            _fmt_money_ru(principal),
            formula,
            _fmt_money_ru(round(interest + 1e-9, 2)),
        ])

    data.append(["", "", "", "", "", "", "", "", "Итого:", _fmt_money_ru(total_interest)])

    tbl = Table(
        data,
        colWidths=[10*mm, 22*mm, 22*mm, 14*mm, 18*mm, 22*mm, 20*mm, 22*mm, 48*mm, 22*mm],
        repeatRows=1,
    )

    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -2), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 1), (0, -2), "CENTER"),
        ("ALIGN", (3, 1), (4, -2), "CENTER"),
        ("ALIGN", (5, 1), (6, -2), "CENTER"),
        ("ALIGN", (7, 1), (7, -2), "RIGHT"),
        ("ALIGN", (9, 1), (9, -2), "RIGHT"),
        ("SPAN", (0, -1), (8, -1)),
        ("ALIGN", (8, -1), (8, -1), "RIGHT"),
        ("ALIGN", (9, -1), (9, -1), "RIGHT"),
        ("LINEABOVE", (0, -1), (-1, -1), 1.0, colors.black),
    ]))

    elements.append(tbl)
    doc.build(elements)
    return buf.getvalue()


# ----------------------------
# Streamlit UI
# ----------------------------

def run():
    st.title("Расчет % по 395 статье")

    st.write("Загрузите Excel с листами «Ставка» и «Список».")

    uploaded = st.file_uploader("Файл Excel", type=["xlsx", "xls"])
    c1, c2 = st.columns(2)
    with c1:
        date_from = st.date_input("Дата от (дата начала расчета)", value=date.today())
    with c2:
        date_to = st.date_input("Дата до (дата до которой производится расчет)", value=date.today())

    if uploaded is None:
        return

    if date_from > date_to:
        st.error("Дата от не может быть больше Даты до.")
        return

    try:
        xls = pd.ExcelFile(uploaded)
        if "Ставка" not in xls.sheet_names or "Список" not in xls.sheet_names:
            st.error("В файле должны быть листы «Ставка» и «Список».")
            return

        df_rate = pd.read_excel(xls, sheet_name="Ставка")
        df_list = pd.read_excel(xls, sheet_name="Список")
        periods = _load_rate_periods(df_rate)

        # Normalize list columns
        list_col_map = {c: c.strip().replace("\xa0", " ") for c in df_list.columns}
        df_list = df_list.rename(columns=list_col_map)

        required = ["Номер договора", "Сумма ОД"]
        missing = [c for c in required if c not in df_list.columns]
        if missing:
            st.error(f"На листе «Список» нет колонок: {', '.join(missing)}")
            return

        if "ФИО" not in df_list.columns:
            df_list["ФИО"] = ""
        if "Дата договора" not in df_list.columns:
            df_list["Дата договора"] = pd.NaT

        out_rows = []
        pdf_files: List[Tuple[str, bytes]] = []

        for _, row in df_list.iterrows():
            contract_no = str(row.get("Номер договора", "")).strip()
            fio = str(row.get("ФИО", "")).strip() if not pd.isna(row.get("ФИО", "")) else ""
            cdate = row.get("Дата договора", pd.NaT)
            contract_date = None if pd.isna(cdate) else _to_date(cdate)

            principal = float(row.get("Сумма ОД", 0) or 0)
            total_interest, calc_rows = _calc_395_for_sum(principal, date_from, date_to, periods)

            out_row = dict(row)
            out_row["Сума по 395"] = total_interest
            out_rows.append(out_row)

            pdf_bytes = _build_pdf_bytes(
                contract_no=contract_no,
                contract_date=contract_date,
                fio=fio,
                as_of=date_to,
                calc_rows=calc_rows,
                total_interest=total_interest,
            )
            pdf_files.append((f"{_safe_filename(contract_no)}.pdf", pdf_bytes))

        df_out = pd.DataFrame(out_rows)

        st.success(f"Готово. Договоров: {len(df_out)}")

        # -------- Excel output (both sheets)
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            df_rate.to_excel(writer, sheet_name="Ставка", index=False)
            df_out.to_excel(writer, sheet_name="Список", index=False)
        excel_buf.seek(0)

        st.download_button(
            "Скачать Excel с «Сума по 395»",
            data=excel_buf,
            file_name="395_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # -------- ZIP of PDFs
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for fname, b in pdf_files:
                zf.writestr(fname, b)
        zip_buf.seek(0)

        st.download_button(
            "Скачать PDF-расчеты (ZIP)",
            data=zip_buf,
            file_name="395_pdfs.zip",
            mime="application/zip",
        )

        st.caption("PDF формируется по примеру: один файл на договор, имя файла — номер договора.")

    except Exception as e:
        st.exception(e)
