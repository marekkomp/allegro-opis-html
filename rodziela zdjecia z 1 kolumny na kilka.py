import streamlit as st
import pandas as pd
import io
import re

st.title("Rozdzielanie kolumny 'Zdjęcia' na wiele kolumn")

uploaded = st.file_uploader("Wybierz plik Excel", type=["xlsx", "xlsm"])

# Ustawienia
deduplicate = st.checkbox("Usuń duplikaty linków w komórce", value=True)
strip_params = st.checkbox("Usuń parametry zapytań (?...) z linków", value=False)

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
    )
    return df

def split_photos_cell(x: object) -> list:
    if pd.isna(x):
        return []
    s = str(x).strip()
    if not s:
        return []
    parts = [p.strip() for p in s.split("|") if p.strip()]
    if deduplicate:
        # zachowaj kolejność
        seen, out = set(), []
        for p in parts:
            if p not in seen:
                seen.add(p); out.append(p)
        parts = out
    if strip_params:
        parts = [re.sub(r"\?.*$", "", p) for p in parts]
    return parts

if uploaded is not None:
    # Wczytaj pierwszy arkusz
    sheets = pd.read_excel(uploaded, sheet_name=None, header=0)
    first_name = list(sheets.keys())[0]
    df = sheets[first_name].copy()
    df = normalize_cols(df)

    # Szukaj kolumny 'Zdjęcia' case-insensitive
    name_map = {c.casefold(): c for c in df.columns}
    key = "zdjęcia".casefold()
    if key not in name_map:
        st.error("Brak kolumny 'Zdjęcia' w pierwszym arkuszu (sprawdź nagłówek w wierszu 1).")
        st.stop()

    photos_col = name_map[key]

    st.caption("Podgląd przed:")
    st.dataframe(df[[photos_col]].head(5))

    # Rozbij komórki na listy
    lists = df[photos_col].apply(split_photos_cell)

    # Ustal maksymalną liczbę zdjęć
    max_n = lists.map(len).max() if len(lists) else 0

    if max_n == 0:
        st.warning("Nie znaleziono żadnych linków do zdjęć w kolumnie 'Zdjęcia'.")
        st.stop()

    # Zbuduj nowe kolumny Zdjęcie 1..N
    for i in range(max_n):
        df[f"Zdjęcie {i+1}"] = lists.apply(lambda L: L[i] if i < len(L) else "")

    st.caption(f"Dodano {max_n} kolumn: Zdjęcie 1 … Zdjęcie {max_n}")
    st.dataframe(df[[photos_col] + [f"Zdjęcie {i+1}" for i in range(max_n)]].head(5))

    # Zapis do Excela
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=first_name)
    output.seek(0)

    st.download_button(
        "Pobierz plik z rozdzielonymi zdjęciami",
        data=output,
        file_name="oferty_zdjecia_rozdzielone.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
