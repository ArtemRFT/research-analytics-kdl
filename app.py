import re
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Аналитика направлений", page_icon="📊", layout="wide")
st.title("📊 Аналитика по исследованиям")
st.caption("Загрузите Excel-файл — данные обрабатываются только в памяти и не сохраняются.")

# ------------------------------------------------------------------
# Вспомогательные функции
# ------------------------------------------------------------------
STATUS_RE = re.compile(r"(\d+)\s*/\s*(\d+)")


def parse_status(s):
    if not isinstance(s, str):
        return None, None
    m = STATUS_RE.search(s)
    if not m:
        return None, None
    return int(m.group(1)), int(m.group(2))


def split_research(s):
    if not isinstance(s, str):
        return []
    parts = [p.strip() for p in s.split(";")]
    return [p for p in parts if p]


def _norm(s) -> str:
    if s is None:
        return ""
    s = str(s).replace("\ufeff", "").replace("\xa0", " ")
    s = unicodedata.normalize("NFKC", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def _find_header(df_raw: pd.DataFrame):
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
        raise ValueError(
            "Не найден заголовок 'Направление' ни на одном листе.\n\n"
            f"Листы: {xls.sheet_names}\n\n"
            f"Первые строки листа '{xls.sheet_names[0]}':\n{sample.head(10).to_string(max_colwidth=40)}"
        )

    sheet, df_raw, hr, hc = chosen
    df = df_raw.iloc[hr:, hc:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)

    need = ["Направление", "Отделение", "Исследования", "Статус"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(
            f"Не найдены колонки: {missing}.\nНайденные: {list(df.columns)}"
        )

    df = df[need].copy()
    df = df[df["Направление"].notna()].reset_index(drop=True)

    parsed = df["Статус"].apply(parse_status)
    df["Статус_Заказано"] = [p[0] for p in parsed]
    df["Статус_Выполнено"] = [p[1] for p in parsed]

    rows = []
    for _, r in df.iterrows():
        for res in split_research(r["Исследования"]):
            rows.append({
                "Направление": r["Направление"],
                "Отделение": r["Отделение"],
                "Исследование": res,
            })
    long_df = pd.DataFrame(rows)

    return df, long_df


# ------------------------------------------------------------------
# UI
# ------------------------------------------------------------------
uploaded = st.file_uploader("Выберите файл .xlsx", type=["xlsx", "xls"])

if uploaded is None:
    st.info("Файл не загружен. Загрузите .xlsx.")
    st.stop()

try:
    raw_df, long_df = load_and_prepare(uploaded.getvalue())
except Exception as e:
    st.error(f"Ошибка чтения файла: {e}")
    st.exception(e)
    st.stop()

if long_df.empty:
    st.warning("Не удалось извлечь ни одного исследования.")
    st.stop()

# ============================================================
# 1. Сводка + топ-исследования графиком
# ============================================================
summary = (
    long_df.groupby("Исследование", dropna=False)
    .size()
    .reset_index(name="Заказано")
    .sort_values("Заказано", ascending=False)
    .reset_index(drop=True)
)

total_orders = int(summary["Заказано"].sum())
total_research = len(summary)
total_depts = long_df["Отделение"].nunique()

c1, c2, c3 = st.columns(3)
c1.metric("Всего заказов исследований", total_orders)
c2.metric("Уникальных исследований", total_research)
c3.metric("Отделений", total_depts)

st.divider()

# --- Горизонтальный бар-чарт топ-N ---
st.subheader("🔝 Топ исследований по числу заказов")

top_n = st.slider("Сколько показать в графике", 5, min(50, len(summary)), min(15, len(summary)))
chart_df = summary.head(top_n).set_index("Исследование")["Заказано"]
st.bar_chart(chart_df, horizontal=True, use_container_width=True)

st.divider()

# ============================================================
# 2. Раскрывающиеся таблицы с графиком по отделениям
# ============================================================
st.subheader("📂 Детализация по исследованиям")
st.caption("Нажмите на исследование, чтобы увидеть распределение по отделениям.")

search = st.text_input("🔍 Фильтр по названию исследования", "").strip().lower()
filtered = summary[summary["Исследование"].str.lower().str.contains(search)] if search else summary

if filtered.empty:
    st.info("Ничего не найдено по фильтру.")
else:
    for _, row in filtered.iterrows():
        research = row["Исследование"]
        total = int(row["Заказано"])

        with st.expander(f"{research} — {total} заказ(ов)"):
            dept_stats = (
                long_df[long_df["Исследование"] == research]
                .groupby("Отделение", dropna=False)
                .size()
                .reset_index(name="Заказано раз")
                .sort_values("Заказано раз", ascending=False)
                .reset_index(drop=True)
            )

            col_chart, col_table = st.columns([2, 3])

            with col_chart:
                st.caption("График по отделениям")
                st.bar_chart(
                    dept_stats.set_index("Отделение")["Заказано раз"],
                    horizontal=True,
                    use_container_width=True,
                    height=max(200, 35 * len(dept_stats)),
                )

            with col_table:
                st.caption("Таблица")
                styled = dept_stats.style.background_gradient(
                    subset=["Заказано раз"], cmap="Blues"
                ).format({"Заказано раз": "{:d}"})
                st.dataframe(styled, use_container_width=True, hide_index=True, height=max(200, 35 * len(dept_stats)))

# ============================================================
# 3. Статусы — визуально внизу
# ============================================================
st.divider()
st.header("📋 Статусы по направлениям")

status_ordered = int(raw_df["Статус_Заказано"].fillna(0).sum())
status_done = int(raw_df["Статус_Выполнено"].fillna(0).sum())
status_pct = (status_done / status_ordered * 100) if status_ordered else 0

s1, s2, s3 = st.columns(3)
s1.metric("Заказано (из статуса)", status_ordered)
s2.metric("Выполнено (из статуса)", status_done)
s3.metric("Выполнение", f"{status_pct:.1f}%")

st.progress(min(status_pct / 100, 1.0))
st.caption(f"Выполнено {status_done} из {status_ordered} позиций по столбцу «Статус».")
