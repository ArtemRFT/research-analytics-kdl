import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Аналитика направлений", page_icon="📊", layout="wide")
st.title("📊 Аналитика по направлениям")
st.caption("Загрузите Excel-файл — данные обрабатываются только в памяти и не сохраняются.")

# ------------------------------------------------------------------
# Вспомогательные функции (все — ДО load_and_prepare)
# ------------------------------------------------------------------
STATUS_RE = re.compile(r"(\d+)\s*/\s*(\d+)")


def parse_status(s):
    """Возвращает (заказано, выполнено). Если не распарсили — (None, None)."""
    if not isinstance(s, str):
        return None, None
    m = STATUS_RE.search(s)
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))


def split_research(s):
    """Делит строку исследований по ';', чистит пробелы и пустые."""
    if not isinstance(s, str):
        return []
    parts = [p.strip() for p in s.split(";")]
    return [p for p in parts if p]


def _norm(s) -> str:
    """Нормализация заголовка: BOM, NBSP, юникод, регистр, пробелы."""
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\ufeff", "").replace("\xa0", " ")
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def _find_header(df_raw: pd.DataFrame):
    """Ищет строку и колонку, где встречается 'направление'."""
    target = "направление"
    for r in range(len(df_raw)):
        for c in range(df_raw.shape[1]):
            if _norm(df_raw.iat[r, c]) == target:
                return r, c
    return None, None


@st.cache_data(show_spinner=False)
def load_and_prepare(file_bytes: bytes):
    xls = pd.ExcelFile(pd.io.common.BytesIO(file_bytes))
    chosen = None
    for sheet in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        hr, hc = _find_header(df_raw)
        if hr is not None:
            chosen = (sheet, df_raw, hr, hc)
            break

    if chosen is None:
        sample = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, nrows=10)
        preview = sample.head(10).to_string(max_colwidth=40)
        raise ValueError(
            "Не найден заголовок 'Направление' ни на одном листе.\n\n"
            f"Листы: {xls.sheet_names}\n\n"
            f"Первые строки листа '{xls.sheet_names[0]}':\n{preview}"
        )

    sheet, df_raw, hr, hc = chosen
    df = df_raw.iloc[hr:, hc:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

    need = ["Направление", "Отделение", "Врач", "Исследования", "Статус"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(
            f"После заголовка не найдены колонки: {missing}.\n"
            f"Найденные колонки: {list(df.columns)}"
        )

    df = df[need].copy()
    df = df[df["Направление"].notna()]

    parsed = df["Статус"].apply(parse_status)
    df["Заказано"] = [p[0] for p in parsed]
    df["Выполнено"] = [p[1] for p in parsed]

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


# ------------------------------------------------------------------
# UI
# ------------------------------------------------------------------
uploaded = st.file_uploader("Выберите файл .xlsx", type=["xlsx", "xls"])

if uploaded is None:
    st.info("Файл не загружен. Загрузите .xlsx, чтобы увидеть аналитику.")
    st.stop()

try:
    raw_df, long_df = load_and_prepare(uploaded.getvalue())
except Exception as e:
    st.error(f"Ошибка чтения файла: {e}")
    st.exception(e)
    st.stop()

# ---------- 1. Общая статистика ----------
st.header("1. Общая статистика")
total_ordered = int(raw_df["Заказано"].fillna(0).sum())
total_done = int(raw_df["Выполнено"].fillna(0).sum())
pct = (total_done / total_ordered * 100) if total_ordered else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Направлений", len(raw_df))
c2.metric("Заказано исследований", total_ordered)
c3.metric("Выполнено исследований", total_done)
c4.metric("Выполнение", f"{pct:.1f}%")

# ---------- 2. По исследованиям ----------
st.header("2. По исследованиям")
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

# ---------- 3. Отделения × исследования ----------
st.header("3. Какие отделения сколько заказывают")
pivot = long_df.pivot_table(
    index="Отделение",
    columns="Исследование",
    values="Заказано",
    aggfunc="sum",
    fill_value=0,
)
st.dataframe(pivot, use_container_width=True)

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

# ---------- 4. Сводка по отделениям ----------
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

# ---------- 5. Детализация по отделению ----------
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

# ---------- Скачать ----------
csv = by_research.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "Скачать аналитику по исследованиям (CSV)",
    csv,
    "research_stats.csv",
    mime="text/csv",
)
