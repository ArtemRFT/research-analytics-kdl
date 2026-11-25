import re
import unicodedata
import pandas as pd
import streamlit as st

def _norm(s) -> str:
    if s is None:
        return ""
    s = str(s)
    # убираем BOM и неразрывные пробелы
    s = s.replace("\ufeff", "").replace("\xa0", " ")
    # нормализуем юникод (ё/е, разные тире и т.п.)
    s = unicodedata.normalize("NFKC", s)
    # схлопываем пробелы и приводим к нижнему регистру
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
    # читаем все листы, берём первый, где нашёлся заголовок
    xls = pd.ExcelFile(pd.io.common.BytesIO(file_bytes))
    chosen = None
    for sheet in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
        hr, hc = _find_header(df_raw)
        if hr is not None:
            chosen = (sheet, df_raw, hr, hc)
            break

    if chosen is None:
        # диагностика: покажем, что реально в файле
        sample = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, nrows=10)
        preview = sample.head(10).to_string(max_colwidth=40)
        raise ValueError(
            "Не найден заголовок 'Направление' ни на одном листе.\n\n"
            f"Листы: {xls.sheet_names}\n\n"
            f"Первые строки листа '{xls.sheet_names[0]}':\n{preview}"
        )

    sheet, df_raw, hr, hc = chosen
    # берём всё, что ниже и правее найденного заголовка
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
