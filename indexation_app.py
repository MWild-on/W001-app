import io
import calendar
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP
import zipfile

import pandas as pd
import streamlit as st
from ui_common import apply_global_css, section_header
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# Регистрируем обычный и жирный шрифт
FONT_DIR = BASE_DIR / "fonts"
pdfmetrics.registerFont(TTFont("DejaVuSans", str(FONT_DIR / "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", str(FONT_DIR / "DejaVuSans-Bold.ttf")))


# Привязываем семейство, чтобы <b> работало
pdfmetrics.registerFontFamily(
    "DejaVuSans",
    normal="DejaVuSans",
    bold="DejaVuSans-Bold",
)

# Константы для стилей
FONT_NAME = "DejaVuSans"          # базовое семейство
FONT_NAME_BOLD = "DejaVuSans-Bold"



# ---------- Форматирование ----------

def fmt_money(value: Decimal | float | int) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    # обнуляем микроскопические остатки типа -1.3E-12
    if value.copy_abs() < Decimal("0.005"):
        value = Decimal("0.00")
    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    # формат с пробелами и запятой: 12345.67 -> "12 345,67"
    s = f"{value:,.2f}"          # "12,345.67"
    s = s.replace(",", " ").replace(".", ",")
    return f"{s} руб."



def fmt_date(d: dt.date | None) -> str:
    if d is None:
        return ""
    return d.strftime("%d.%m.%Y")

def fmt_plain(value: Decimal | float | int) -> str:
    """
    Форматирование чисел для формул:
    12345.67 → 12 345,67
    (без слова 'руб.')
    """
    if not isinstance(value, Decimal):
        value = Decimal(str(value))

    # нули и мусор округляем
    if value.copy_abs() < Decimal("0.005"):
        value = Decimal("0.00")

    value = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    s = f"{value:.2f}"
    integer, frac = s.split(".")

    integer = integer.replace(",", "")
    integer = "{:,}".format(int(integer)).replace(",", " ")

    return f"{integer},{frac}"

# =========================
#  Вспомогательные функции
# =========================

def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


def next_month(year: int, month: int):
    if month == 12:
        return year + 1, 1
    return year, month + 1


def load_cpi_from_sheet(xls: pd.ExcelFile, sheet_name: str = "Индекс"):
    """
    Лист 'Индекс':
      Год | Месяц | Индексы потребительских цен
    """
    df = pd.read_excel(xls, sheet_name=sheet_name)
    cpi_map = {}
    for _, row in df.iterrows():
        y = int(row["Год"])
        m = int(row["Месяц"])
        idx_val = Decimal(str(row["Индексы потребительских цен"]))
        cpi_map[(y, m)] = idx_val
    return cpi_map, df


def compute_indexation_for_period(
    amount: Decimal,
    start_date: dt.date,
    end_date: dt.date,
    cpi_map,
) -> Decimal:
    """
    Индексация суммы `amount` за период [start_date, end_date] включительно.
    - первый и последний месяц пропорционально дням;
    - месяцы между ними берутся целиком.
    """
    if amount <= 0 or start_date > end_date:
        return Decimal("0.00")

    product = Decimal("1.0")
    s = start_date
    e = end_date

    y, m = s.year, s.month

    while True:
        key = (y, m)
        if key not in cpi_map:
            raise ValueError(
                f"Нет ИПЦ для {y}-{m:02d} — добавь этот месяц на лист 'Индекс'."
            )

        cpi = cpi_map[key]  # например 100.84
        delta = cpi - Decimal("100")
        dim = days_in_month(y, m)

        # Период целиком в одном месяце
        if y == s.year and m == s.month and y == e.year and m == e.month:
            days = e.day - s.day + 1
            prop = Decimal(days) / Decimal(dim)
            factor = (delta * prop + Decimal("100")) / Decimal("100")
            product *= factor
            break

        # Первый месяц (не последний)
        elif y == s.year and m == s.month:
            days = dim - s.day + 1
            prop = Decimal(days) / Decimal(dim)
            factor = (delta * prop + Decimal("100")) / Decimal("100")
            product *= factor

        # Последний месяц (не первый)
        elif y == e.year and m == e.month:
            days = e.day
            prop = Decimal(days) / Decimal(dim)
            factor = (delta * prop + Decimal("100")) / Decimal("100")
            product *= factor
            break

        # Полный месяц внутри периода
        else:
            factor = cpi / Decimal("100")
            product *= factor

        y, m = next_month(y, m)

    ind = (amount * product - amount).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return ind


# =========================
#  Расчёт по одному долгу
# =========================

def calculate_indexation_for_debt(
    reg_num: int,
    main_row: pd.Series,
    payments_df: pd.DataFrame,
    cpi_map,
    cutoff_date: dt.date,
):
    """
    По одному долгу:
    - индексация считается только до cutoff_date (включительно);
    - по каждому платежу отдельный период;
    - если после последнего платежа до cutoff_date есть дни,
      считаем индексацию на остаток долга до cutoff_date.
    """

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    initial_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))

    pays = payments_df[payments_df["Рег. номер"] == reg_num].copy()
    if pays.empty:
        # только хвост до cutoff_date, если приказ раньше
        if order_date <= cutoff_date:
            ind = compute_indexation_for_period(
                initial_debt, order_date, cutoff_date, cpi_map
            )
            return ind, [{
                "period_start": order_date,
                "period_end": cutoff_date,
                "debt_before": initial_debt,
                "indexation": ind,
                "payment_date": None,
                "payment_amount": Decimal("0.00"),
                "debt_after_payment": initial_debt,
            }]
        return Decimal("0.00"), []

    # даты платежей
    pays["date"] = pd.to_datetime(pays["Дата платежа"]).dt.date
    # учитываем только платежи не позже крайней даты
    pays = pays[pays["date"] <= cutoff_date]
    pays = pays[pays["Сумма платежа"] > 0]

    if pays.empty:
        # нет платежей в периоде расчёта – одна сплошная индексация до cutoff_date
        if order_date <= cutoff_date:
            ind = compute_indexation_for_period(
                initial_debt, order_date, cutoff_date, cpi_map
            )
            return ind, [{
                "period_start": order_date,
                "period_end": cutoff_date,
                "debt_before": initial_debt,
                "indexation": ind,
                "payment_date": None,
                "payment_amount": Decimal("0.00"),
                "debt_after_payment": initial_debt,
            }]
        return Decimal("0.00"), []

    grouped = (
        pays.groupby("date")["Сумма платежа"]
        .sum()
        .reset_index()
        .sort_values("date")
    )

    total_indexation = Decimal("0.00")
    remaining_debt = initial_debt
    current_start = order_date
    periods = []

    # периоды к платежам
    for _, row in grouped.iterrows():
        pay_date = row["date"]
        pay_amount = Decimal(str(row["Сумма платежа"]))

        if remaining_debt <= 0:
            break
        if current_start > cutoff_date:
            break

        period_end = min(pay_date, cutoff_date)

        ind = compute_indexation_for_period(
            remaining_debt, current_start, period_end, cpi_map
        )
        total_indexation += ind

        remaining_after_payment = remaining_debt - pay_amount

        periods.append(
            {
                "period_start": current_start,
                "period_end": period_end,
                "debt_before": remaining_debt,
                "indexation": ind,
                "payment_date": pay_date,
                "payment_amount": pay_amount,
                "debt_after_payment": remaining_after_payment,
            }
        )

        remaining_debt = remaining_after_payment
        current_start = pay_date + dt.timedelta(days=1)

    # хвост до крайней даты, если долг ещё есть
    if remaining_debt > 0 and current_start <= cutoff_date:
        ind = compute_indexation_for_period(
            remaining_debt, current_start, cutoff_date, cpi_map
        )
        total_indexation += ind
        periods.append(
            {
                "period_start": current_start,
                "period_end": cutoff_date,
                "debt_before": remaining_debt,
                "indexation": ind,
                "payment_date": None,
                "payment_amount": Decimal("0.00"),
                "debt_after_payment": remaining_debt,
            }
        )

    return total_indexation, periods


# =========================
#  PDF по долгу (в память)
# =========================

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import mm

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import mm

def generate_pdf_bytes_for_debt(
    reg_num: int,
    main_row: pd.Series,
    total_indexation: Decimal,
    periods,
    cutoff_date: dt.date,
) -> bytes:

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
        title=f"Расчёт индексации {reg_num}",
    )

    story = []

    # -----------------------------
    # Стили
    # -----------------------------
    styles = getSampleStyleSheet()

    style_title = ParagraphStyle(
        "Title",
        parent=styles["Normal"],
        fontName=FONT_NAME_BOLD,
        fontSize=14,
        leading=18,
        spaceAfter=12,
        alignment=TA_CENTER,
    )

    style_subtitle = ParagraphStyle(
        "Subtitle",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=11,
        leading=14,
        spaceAfter=12,
        alignment=TA_CENTER,
    )
   
    style_h2 = ParagraphStyle(
        "Heading2",
        parent=styles["Normal"],
        fontName=FONT_NAME_BOLD,
        fontSize=12,
        leading=16,
        spaceAfter=10,
        alignment=TA_LEFT,
    )

    style_text = ParagraphStyle(
        "NormalText",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=11,
        leading=15,
        spaceAfter=6,
        alignment=TA_LEFT,
    )

    style_formula = ParagraphStyle(
        "Formula",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=11,
        leading=15,
        spaceAfter=12,
        alignment=TA_LEFT,
    )

    # -----------------------------
    # Заголовок
    # -----------------------------
    story.append(Paragraph("Расчёт индексации присуждённых сумм", style_title))

    # Если в исходном Excel есть колонка «Номер приказа» и она заполнена,
    # добавляем строку "по делу № ..."
    order_number = ""
    if "Номер приказа" in main_row.index:
        raw_num = main_row.get("Номер приказа")
        if pd.notna(raw_num):
            order_number = str(raw_num).strip()

    if order_number and order_number.lower() not in ("nan", "none"):
        story.append(Paragraph(f"по делу № {order_number}", style_subtitle))

    
    story.append(Spacer(1, 12))

    order_date = pd.to_datetime(main_row["Дата вынесения приказа"]).date()
    base_debt = Decimal(str(main_row["Сумма платежей с декабря 2024"]))

    # --- НОВОЕ: конец периода индексации = максимальная дата платежа ---
    payment_dates = [
        p["payment_date"]
        for p in periods
        if p.get("payment_date") is not None and p.get("payment_amount", Decimal("0.00")) != Decimal("0.00")
    ]
    if payment_dates:
        header_end_date = max(payment_dates)
    else:
        # если платежей нет – как раньше, до cutoff_date
        header_end_date = cutoff_date

    total_days = (header_end_date - order_date).days + 1

    # Взысканная сумма (сумма на дату приказа)
    story.append(Paragraph(
        f"<b>Взысканная сумма на дату начала периода индексации ({fmt_date(order_date)}):</b><br/>{fmt_money(base_debt)}",
        style_text,
    ))

    # Период индексации в шапке
    story.append(Paragraph(
        f"<b>Период индексации: </b>{fmt_date(order_date)} – {fmt_date(header_end_date)} ({total_days} дней)",
        style_text,
    ))

    # Регион
    story.append(Paragraph("<b>Регион: </b>Российская Федерация", style_text))

    # Сумма индексации
    story.append(Paragraph(
        f"<b>Сумма индексации: </b>{fmt_money(total_indexation)}",
        style_text,
    ))

    story.append(Spacer(1, 12))

    # -----------------------------
    # Порядок расчёта
    # -----------------------------
    story.append(Paragraph("Порядок расчёта:", style_h2))

    story.append(Paragraph(
        "Сумма долга × ИПЦ1 × пропорция первого месяца × ИПЦ2 × ИПЦ3 × ... × ИПЦn<br/>"
        "× пропорция последнего месяца – сумма долга = И",
        style_formula
    ))

    story.append(Spacer(1, 12))

    # -----------------------------
    # Первый период (до первого платежа)
    # -----------------------------
    if periods:
        first = periods[0]

        story.append(Paragraph(
            f"Индексация за период: {fmt_date(first['period_start'])} – {fmt_date(first['period_end'])}",
            style_h2
        ))

        story.append(Paragraph(
            f"Индексируемая сумма на начало периода: {fmt_money(first['debt_before'])}",
            style_text
        ))

        story.append(Paragraph(
            f"Индексация за период: {fmt_money(first['indexation'])}",
            style_text
        ))

        story.append(Spacer(1, 12))

    # -----------------------------
    # Частичные оплаты
    # -----------------------------
    # Количество периодов с платежами
    pay_period_count = sum(
        1 for p in periods
        if p.get("payment_date") is not None and p.get("payment_amount", Decimal("0.00")) != Decimal("0.00")
    )

    for i in range(pay_period_count):
        pay_period = periods[i]
        # период после этого платежа (если есть)
        next_period = periods[i + 1] if i + 1 < len(periods) else None

        story.append(Paragraph(f"Частичная оплата долга #{i + 1}", style_h2))

        # платёж
        story.append(Paragraph(
            f"Платёж #{i + 1}: дата {fmt_date(pay_period['payment_date'])}, "
            f"сумма {fmt_money(pay_period['payment_amount'])}",
            style_text
        ))

        # период индексации ПОСЛЕ данного платежа
        if next_period is not None:
            story.append(Paragraph(
                f"Период индексации: {fmt_date(next_period['period_start'])} – "
                f"{fmt_date(next_period['period_end'])}",
                style_text
            ))
        else:
            story.append(Paragraph(
                "Период индексации после данного платежа отсутствует",
                style_text
            ))

        # формула остатка
        line = (
            f"Остаток долга на начало периода: "
            f"{fmt_plain(pay_period['debt_before'])} - {fmt_plain(pay_period['payment_amount'])} "
            f"= {fmt_plain(pay_period['debt_after_payment'])} руб."
        )
        story.append(Paragraph(line, style_text))

        story.append(Paragraph(
            f"Остаток долга после периода: {fmt_money(pay_period['debt_after_payment'])}",
            style_text
        ))

        # индексация за период после платежа
        if next_period is not None:
            story.append(Paragraph(
                f"Индексация за период: {fmt_money(next_period['indexation'])}",
                style_text
            ))
        else:
            story.append(Paragraph(
                "Индексация за период: 0,00 руб.",
                style_text
            ))

        story.append(Spacer(1, 8))

    # -----------------------------
    # Итог
    # -----------------------------
    story.append(Paragraph(
        f"Итоговая индексация = {fmt_money(total_indexation)}",
        style_h2
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()




# =========================
#  Основная обработка файла
# =========================

def process_workbook(uploaded_file, cutoff_date: dt.date):
    """
    Принимает загруженный Excel-файл и крайнюю дату расчёта.
    Возвращает:
      - bytes итогового Excel
      - bytes ZIP с PDF по каждому долгу
      - DataFrame с основным листом
    """
    xls = pd.ExcelFile(uploaded_file)

    main_df = pd.read_excel(xls, sheet_name="Основной")
    payments_df = pd.read_excel(xls, sheet_name="Платежи")
    cpi_map, cpi_df = load_cpi_from_sheet(xls, sheet_name="Индекс")

    # ограничим крайнюю дату последним месяцем, который есть в индексации
    max_year = int(cpi_df["Год"].max())
    max_month = int(cpi_df["Месяц"].max())
    last_cpi_date = dt.date(max_year, max_month, days_in_month(max_year, max_month))

    effective_cutoff = min(cutoff_date, last_cpi_date)

    new_index_col = []
    pdf_files = []  # (file_name, bytes)

    for _, row in main_df.iterrows():
        reg_num = int(row["Рег номер"])

        total_indexation, periods = calculate_indexation_for_debt(
            reg_num=reg_num,
            main_row=row,
            payments_df=payments_df,
            cpi_map=cpi_map,
            cutoff_date=effective_cutoff,
        )

        new_index_col.append(float(total_indexation))

        if periods:
            pdf_bytes = generate_pdf_bytes_for_debt(
                reg_num=reg_num,
                main_row=row,
                total_indexation=total_indexation,
                periods=periods,
                cutoff_date=effective_cutoff,
            )
            pdf_files.append((f"{reg_num}.pdf", pdf_bytes))

    main_df["Сумма индексации (расчёт)"] = new_index_col

 # ----- Excel в память -----
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        main_df.to_excel(writer, sheet_name="Основной", index=False)
        payments_df.to_excel(writer, sheet_name="Платежи", index=False)
        cpi_df.to_excel(writer, sheet_name="Индекс", index=False)
    excel_buffer.seek(0)
    excel_bytes = excel_buffer.getvalue()

    # ----- ЕДИНЫЙ ZIP: Excel + все PDF -----
    combined_zip_buffer = io.BytesIO()
    with zipfile.ZipFile(combined_zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Excel-файл
        zf.writestr("Файл для расчета_с_индексацией.xlsx", excel_bytes)

        # PDF-отчёты (в папке pdf/)
        for fname, fbytes in pdf_files:
            zf.writestr(f"pdf/{fname}", fbytes)

    combined_zip_buffer.seek(0)
    combined_zip_bytes = combined_zip_buffer.getvalue()

    # теперь возвращаем единый архив
    return combined_zip_bytes, main_df, effective_cutoff

# =========================
#  Streamlit entry point для вкладки
# =========================

def run():

    apply_global_css()

    section_header(
        "Расчёт индексации присуждённых сумм",
        "Загрузите Excel-файл со страницами «Основной», "
        "«Платежи» и «Индекс»."
    )

    today = dt.date.today()
    first_day_this_month = today.replace(day=1)
    default_cutoff = first_day_this_month - dt.timedelta(days=1)

    uploaded_file = st.file_uploader(
        "Загрузите файл формата .xlsx",
        type=["xlsx"],
    )
    
    cutoff_date = st.date_input(
        "Крайняя дата для расчёта",
        value=default_cutoff,
        format="DD.MM.YYYY",
    )

    if uploaded_file is not None:
        st.success("Файл загружен. Нажмите кнопку ниже для расчёта.")

        if st.button("Рассчитать индексацию"):
            try:
                zip_bytes, main_df, effective_cutoff = process_workbook(
                    uploaded_file, cutoff_date
                )

                st.info(f"Фактическая крайняя дата расчёта: {effective_cutoff}")

                st.download_button(
                    label="Скачать архив (Excel + PDF-отчёты)",
                    data=zip_bytes,
                    file_name="indexation_results.zip",
                    mime="application/zip",
                )


            except Exception as e:
                st.error(f"Ошибка при обработке файла: {e}")
