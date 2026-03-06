# ui_common.py
import streamlit as st

def apply_global_css():
    """Единое оформление для всех разделов."""
    st.markdown(
        """
        <style>
        /* ширина и отступы основной колонки */
        .block-container {
            max-width: 1100px;
            padding-top: 1.5rem;
            padding-bottom: 3rem;
        }

        /* кнопки */
        .stButton>button {
            border-radius: 999px;
            padding: 0.55rem 1.4rem;
            font-weight: 600;
        }

        /* file uploader — чуть плотнее */
        .uploadedFile,
        .stFileUploader label {
            font-size: 0.95rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def section_header(title: str, subtitle: str, release: str | None = None):
    """Единый заголовок раздела (как большая шапка страницы). release — дата/время релиза мелким шрифтом под названием."""
    st.markdown(
        f"""
        <h1 style="font-size: 40px; font-weight: 700; margin-bottom: 0.3rem;">
            {title}
        </h1>
        """
        + (
            f'<p style="font-size: 12px; color: #6b7280; margin-top: 0; margin-bottom: 0.4rem;">{release}</p>'
            if release else ""
        )
        + f"""
        <p style="font-size: 16px; color: #4b5563; margin-bottom: 1.8rem;">
            {subtitle}
        </p>
        """,
        unsafe_allow_html=True,
    )
