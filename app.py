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
    df = df[df["Направление"].notna()]

    parsed = df["Статус"].apply(parse_status)
    df["Заказано"] = [p[0] for p in parsed]

    rows = []
    for _, r in df.iterrows():
        for res in split_research(r["Исследования"]):
            rows.append({
                "Направление": r["Направление"],
                "Отделение": r["Отделение"],
                "Исследование": res,
                "Заказано": r["Заказано"],
            })
    return pd.DataFrame(rows)


# ------------------------------------------------------------------
# UI
# ------------------------------------------------------------------
uploaded = st.file_uploader("Выберите файл .xlsx", type=["xlsx", "xls"])

if uploaded is None:
    st.info("Файл не загружен. Загрузите .xlsx.")
    st.stop()

try:
    long_df = load_and_prepare(uploaded.getvalue())
except Exception as e:
    st.error(f"Ошибка чтения файла: {e}")
    st.exception(e)
    st.stop()

if long_df.empty:
    st.warning("Не удалось извлечь ни одного исследования.")
    st.stop()

# ---------- Раскрывающиеся таблицы ----------
# Сводка: исследование -> сколько всего заказано
summary = (
    long_df.groupby("Исследование", dropna=False)["Заказано"]
    .sum()
    .reset_index()
    .sort_values("Заказано", ascending=False)
)

st.subheader("Исследования (нажмите на название, чтобы раскрыть отделения)")

for _, row in summary.iterrows():
    research = row["Исследование"]
    total = int(row["Заказано"]) if pd.notna(row["Заказано"]) else 0

    with st.expander(f"{research} — заказано: {total}"):
        dept_stats = (
            long_df[long_df["Исследование"] == research]
            .groupby("Отделение", dropna=False)["Заказано"]
            .sum()
            .reset_index()
            .rename(columns={"Отделение": "Отделение", "Заказано": "Заказано раз"})
            .sort_values("Заказано раз", ascending=False)
        )
        st.dataframe(dept_stats, use_container_width=True, hide_index=True)
