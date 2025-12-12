# app.py — ГЛАВНЫЙ ФАЙЛ

import streamlit as st
import converter_app
import indexation_app
import statement_app
import percent395_app


from ui_common import apply_global_css  # общий стиль для всех страниц

# Сначала настраиваем страницу (до любых других st.*)
st.set_page_config(
    page_title="Bank tools",
    layout="wide",
)

# ===== Простая авторизация =====

def check_auth():
    """Авторизация по логину и паролю."""
    users = st.secrets["users"]

    # инициализация
    if "auth_ok" not in st.session_state:
        st.session_state["auth_ok"] = False
        st.session_state["current_user"] = None

    # уже авторизован
    if st.session_state["auth_ok"]:
        return True

    # три колонки: центрируем всю форму
    col_left, col_center, col_right = st.columns([1, 2, 1])

    with col_center:
        # небольшой ограничитель ширины, чтобы поля не были слишком широкими
        st.markdown(
            "<div style='max-width: 480px; margin: 0 auto;'>",
            unsafe_allow_html=True,
        )

        st.markdown("### Добро пожаловать в W001-app")

        login = st.text_input("Логин:", key="login")
        password = st.text_input("Пароль:", type="password", key="password")

        if st.button("Войти", use_container_width=False):
            if login in users and password == users[login]:
                st.session_state["auth_ok"] = True
                st.session_state["current_user"] = login
                st.rerun()
            else:
                st.error("Неверный логин или пароль")

        st.markdown("</div>", unsafe_allow_html=True)

    return False



    return True

# Блок защиты
if not check_auth():
    st.stop()



def main():
    # применяем единый CSS для всех разделов
    apply_global_css()
    user = st.session_state.get("current_user", "—")
    st.sidebar.caption(f"Пользователь: {user}")

    
    st.sidebar.title("Навигация")
    page = st.sidebar.radio(
        "",
        ("Конвертер", "Индексация", "Создание выписки", "Расчет % по 395 статье"),
    )

    if page == "Конвертер":
        converter_app.run()
    elif page == "Индексация":
        indexation_app.run()
    elif page == "Создание выписки":
        statement_app.run()
    elif page == "Расчет % по 395 статье":
        percent395_app.run()


if __name__ == "__main__":
    main()
