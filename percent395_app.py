# percent395_app.py

import io
import os
import zipfile
import math
import datetime as dt
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional

import pandas as pd
import openpyxl
import streamlit as st

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# =========================
# Data
# =========================

@dataclass
class RateRow:
    start: dt.date
    end: dt.date
    days_in_year: int
    rate: float


# =========================
# Entry point for app.py
# =========================

def run():
    percent395_app()


# =========================
# Fonts
# =========================

def _register_cyrillic_font():
    candidates = [
        ("LiberationSerif", "fonts/LiberationSerif-Regular.ttf"),
        ("DejaVuSans", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ("DejaVuSans", "/usr/share/fonts/dejavu/DejaVuSans.ttf"),
    ]

    for name, path in candidates:
        if os.path.exists(path):
            if name not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(name, path))
            return name, name

    return "Helvetica", "Helvetica-Bold"


# =========================
# UI
# =========================

def percent395_app():
    st.title("Начисление процентов по ст. 395 ГК РФ")

    uploaded = st.file_uploader("Загрузка файла (Excel)", type=["xlsx"])

    col1, col2 = st.columns(2)
    with col1:
        date_from = st.date_input("Дата от")
    with col2:
        date_to = st.date_input("Дата до")

    calc = st.button("Рассчитать", type="primary", disabled=(uploaded is None))

    sig = None
    if uploaded:
        sig = (uploaded.name, str(date_from), str(date_to))

    if sig and st.session_state.get("p395_sig") != sig:
        st.session_state["p395_sig"] = sig
        st.session_state.pop("p395_zip", None)
        st.session_state.pop("p395_xlsx", None)

    if uploaded and date_from > date_to:
        st.error("Дата от не может быть больше даты до.")
        return

    if calc and uploaded:
        try:
            zip_bytes, xlsx_bytes = run_calculation(
                uploaded.getvalue(), date_from, date_to
            )
            st.session_state["p395_zip"] = zip_bytes
            st.session_state["p395_xlsx"] = xlsx_bytes
            st.success("Готово. Скачайте результат.")
        except Exception as e:
            st.exception(e)

    if "p395_zip" in st.session_state and "p395_xlsx" in st.session_state:
        st.info("Результаты готовы")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Скачать ZIP (только PDF)",
                data=st.session_state["p395_zip"],
                file_name="percent395_pdfs.zip",
                mime="application/zip",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Скачать Excel",
                data=st.session_state["p395_xlsx"],
                file_name="percent395_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# =========================
# Helpers
# =========================

def _to_date(x) -> Optional[dt.date]:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    return pd.to_datetime(x, dayfirst=True).date()


def _sheet_to_df(wb, sheet_name):
    ws = wb[sheet_name]
    data = list(ws.values)
    header = [str(v) for v in data[0]]
    return pd.DataFrame(data[1:], columns=header).dropna(how="all")


def _find_sheet_name(wb, names):
    for n in names:
        if n in wb.sheetnames:
            return n
    return None


def _parse_rates(df):
    out = []
    for _, r in df.iterrows():
        out.append(
            RateRow(
                start=_to_date(r["С"]),
                end=_to_date(r["По"]),
                days_in_year=int(r["Дней в году"]),
                rate=float(r["Ставка"]) / 100,
            )
        )
    return out


def _rate_for_date(rates, d):
    for r in rates:
        if r.start <= d <= r.end:
            return r
    raise ValueError("Нет ставки")


def _fmt(d):
    return d.strftime("%d.%m.%Y")


def _fmt_money(x):
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


# =========================
# PDF
# =========================

def _build_pdf(contract, rows, total):
    font, _ = _register_cyrillic_font()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("T", fontName=font, fontSize=12, alignment=1))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )

    data = [["№", "Период с", "Период по", "Дней", "Ставка %", "ОД", "Проценты"]]
    for i, r in enumerate(rows, 1):
        data.append([
            i, _fmt(r["from"]), _fmt(r["to"]),
            r["days"], r["rate"] * 100,
            _fmt_money(r["principal"]),
            _fmt_money(r["interest"]),
        ])

    data.append(["", "", "", "", "ИТОГО", "", _fmt_money(total)])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))

    doc.build([
        Paragraph(f"Расчет процентов по договору №{contract}", styles["T"]),
        Spacer(1, 10),
        table,
    ])

    return buf.getvalue()


# =========================
# Calculation
# =========================

def run_calculation(excel_bytes, date_from, date_to):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)

    df_list = _sheet_to_df(wb, "Список")
    df_rate = _sheet_to_df(wb, "Ставка")

    rates = _parse_rates(df_rate)

    totals = {}
    pdfs = {}

    for _, r in df_list.iterrows():
        od = float(r["Сумма ОД"])
        cur = od
        total = 0
        rows = []

        d = date_from
        while d <= date_to:
            rate = _rate_for_date(rates, d)
            total += cur * rate.rate / rate.days_in_year
            rows.append({
                "from": d,
                "to": d,
                "days": 1,
                "rate": rate.rate,
                "principal": cur,
                "interest": cur * rate.rate / rate.days_in_year,
            })
            d += dt.timedelta(days=1)

        pdfs[str(r["Номер договора"])] = _build_pdf(
            r["Номер договора"], rows, total
        )
        totals[str(r["Номер договора"])] = round(total, 2)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for k, v in pdfs.items():
            z.writestr(f"{k}.pdf", v)

    xlsx = io.BytesIO()
    wb.save(xlsx)

    return out.getvalue(), xlsx.getvalue()
