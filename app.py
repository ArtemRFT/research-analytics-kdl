import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Аналитика направлений", page_icon="📊", layout="wide")

st.title("📊 Аналитика по направлениям")
st.caption("Загрузите Excel-файл — данные обрабатываются только в памяти и не сохраняются.")

uploaded = st.file_uploader("Выберите файл .xlsx", type=["xlsx", "xls"])

# ---------- Парсинг столбца "Статус" ----------
STATUS_RE = re.compile(r"(\d+)\s*/\s*(\d+)")

def parse_status(s: str):
    """Возвращает (заказано, выполнено). Если не распарсили — (None, None)."""
    if not isinstance(s, str):
        return None, None
    m = STATUS_RE.search(s)
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))

# ---------- Парсинг столбца "Исследования" ----------
def split_research(s: str):
    if not isinstance(s, str):
        return []
    # разделитель ';', чистим пробелы и пустые
    parts = [p.strip() for p in s.split(";")]
    return [p for p in parts if p]

@st.cache_data(show_spinner=False)
def load_and_prepare(file_bytes: bytes):
    df = pd.read_excel(pd.io.common.BytesIO(file_bytes))

    # Находим строку заголовка: ищем ячейку "Направление" в столбце A
    header_row = None
    for i in range(min(20, len(df))):
        if str(df.iloc[i, 0]).strip() == "Направление":
            header_row = i
            break
    if header_row is None:
        raise ValueError("Не найден заголовок 'Направление' в первом столбце")

    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    # Оставляем нужные колонки (если есть)
    need = ["Направление", "Отделение", "Врач", "Исследования", "Статус"]
    for col in need:
        if col not in df.columns:
            raise ValueError(f"В файле нет колонки '{col}'")

    df = df[need].copy()
    df = df[df["Направление"].notna()]

    # Статус -> заказано/выполнено
    parsed = df["Статус"].apply(parse_status)
    df["Заказано"] = [p[0] for p in parsed]
    df["Выполнено"] = [p[1] for p in parsed]

    # Разворачиваем исследования: одна строка = одно исследование в одном направлении
    rows = []
    for _, r in df.iterrows():
        for res in split_research(r["Исследования"]):
            rows.append({
                "Направление": r["Направление"],
                "Отделение": r["Отделение"],
                "Врач": r["Врач"],
                "Исследование": res,
                "Заказано": r["Заказано"],
                "Выполнено": r["Выполнено"],
            })
    long_df = pd.DataFrame(rows)
    return df, long_df

if uploaded is not None:
    try:
        raw_df, long_df = load_and_prepare(uploaded.getvalue())
    except Exception as e:
        st.error(f"Ошибка чтения файла: {e}")
        st.stop()

    # ---------- Общая статистика ----------
    st.header("1. Общая статистика")
    total_ordered = int(raw_df["Заказано"].fillna(0).sum())
    total_done = int(raw_df["Выполнено"].fillna(0).sum())
    pct = (total_done / total_ordered * 100) if total_ordered else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Направлений", len(raw_df))
    c2.metric("Заказано исследований", total_ordered)
    c3.metric("Выполнено исследований", total_done)
    c4.metric("Выполнение", f"{pct:.1f}%")

    # ---------- Аналитика по исследованиям ----------
    st.header("2. По исследованиям")

    # сколько раз заказано каждое исследование (по направлениям)
    by_research = (
        long_df.groupby("Исследование", dropna=False)
        .agg(
            Заказано=("Заказано", "sum"),
            Выполнено=("Выполнено", "sum"),
            Направлений=("Направление", "nunique"),
        )
        .reset_index()
    )
    by_research["% выполнения"] = (
        by_research["Выполнено"] / by_research["Заказано"].replace(0, pd.NA) * 100
    ).round(1)
    by_research = by_research.sort_values("Заказано", ascending=False)

    st.dataframe(by_research, use_container_width=True, hide_index=True)

    # ---------- Отделения × исследования ----------
    st.header("3. Какие отделения сколько заказывают")

    pivot = (
        long_df.pivot_table(
            index="Отделение",
            columns="Исследование",
            values="Заказано",
            aggfunc="sum",
            fill_value=0,
        )
    )
    st.dataframe(pivot, use_container_width=True)

    # Топ-исследование по каждому отделению
    st.subheader("Топ-3 исследования по каждому отделению")
    top3 = (
        long_df.groupby(["Отделение", "Исследование"], dropna=False)["Заказано"]
        .sum()
        .reset_index()
        .sort_values(["Отделение", "Заказано"], ascending=[True, False])
        .groupby("Отделение")
        .head(3)
    )
    st.dataframe(top3, use_container_width=True, hide_index=True)

    # ---------- Сводка по отделениям ----------
    st.header("4. Сводка по отделениям")
    by_dept = (
        long_df.groupby("Отделение", dropna=False)
        .agg(
            Направлений=("Направление", "nunique"),
            Заказано=("Заказано", "sum"),
            Выполнено=("Выполнено", "sum"),
            Видов_исследований=("Исследование", "nunique"),
        )
        .reset_index()
        .sort_values("Заказано", ascending=False)
    )
    by_dept["% выполнения"] = (
        by_dept["Выполнено"] / by_dept["Заказано"].replace(0, pd.NA) * 100
    ).round(1)
    st.dataframe(by_dept, use_container_width=True, hide_index=True)

    # ---------- Фильтр по отделению ----------
    st.header("5. Детализация по отделению")
    dept = st.selectbox("Отделение", sorted(long_df["Отделение"].dropna().unique()))
    if dept:
        sub = long_df[long_df["Отделение"] == dept]
        detail = (
            sub.groupby("Исследование")
            .agg(Заказано=("Заказано", "sum"), Выполнено=("Выполнено", "sum"))
            .reset_index()
            .sort_values("Заказано", ascending=False)
        )
        st.dataframe(detail, use_container_width=True, hide_index=True)

    # ---------- Скачать результат ----------
    csv = by_research.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Скачать аналитику по исследованиям (CSV)", csv, "research_stats.csv")
else:
    st.info("Файл не загружен. Загрузите .xlsx, чтобы увидеть аналитику.")