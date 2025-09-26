import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import io
import json
import re
import html

# --- Helpers ---
def html_to_text(raw_html: str) -> str:
    """Zamienia HTML na czysty tekst z sensownymi nowymi liniami."""
    if not isinstance(raw_html, str):
        return ""
    # Parsowanie
    soup = BeautifulSoup(raw_html, "html.parser")

    # Usuwamy skrypty/style i komentarze
    for tag in soup(["script", "style"]):
        tag.decompose()

    # Wymuszamy nowe linie w typowych blokach
    for br in soup.find_all(["br"]):
        br.replace_with("\n")
    for block in soup.find_all(["p", "div", "section", "article", "h1","h2","h3","h4","h5","h6","li"]):
        if block.name == "li":
            # punktor dla list
            prefix = "• "
            block.insert(0, prefix)
        # dodajemy newline po bloku (jeśli jeszcze nie ma)
        if not (block.text.endswith("\n")):
            block.append("\n")

    text = soup.get_text(separator=" ")
    text = html.unescape(text)

    # Sprzątanie białych znaków
    text = re.sub(r"[ \t]+\n", "\n", text)      # spacje przed \n
    text = re.sub(r"\n{3,}", "\n\n", text)      # maks 2 nowe linie pod rząd
    text = re.sub(r"[ \t]{2,}", " ", text)      # wielokrotne spacje -> jedna
    return text.strip()

def extract_content_from_json(maybe_json: str) -> str:
    """Jeśli to JSON sekcji, zwróć zlepione 'content' z itemów TEXT; wpp. zwróć oryginał."""
    if not isinstance(maybe_json, str):
        return ""
    try:
        data = json.loads(maybe_json)
        clean = []
        for section in data.get("sections", []):
            for item in section.get("items", []):
                if item.get("type") == "TEXT" and "content" in item:
                    clean.append(item["content"])
        return "".join(clean) if clean else maybe_json
    except json.JSONDecodeError:
        return maybe_json

def clean_html_to_plain(content: str) -> str:
    """Pipeline: JSON -> HTML -> TEXT"""
    html_content = extract_content_from_json(content)
    return html_to_text(html_content)

# --- Streamlit UI ---
st.title("Przetwarzanie Excel → czysty tekst w 'Opis oferty'")

uploaded_file = st.file_uploader("Wybierz plik Excel", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    # Wczytanie wszystkich arkuszy, pierwszy jako aktywny
    sheets = pd.read_excel(uploaded_file, sheet_name=None, header=0)

    first_sheet_name = list(sheets.keys())[0]
    df = sheets[first_sheet_name]

    # Normalizacja nagłówków (spacje, BOM, wielkość liter)
    df.columns = (
        df.columns.astype(str)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
    )
    # Szukamy kolumny „Opis oferty” case-insensitive
    name_map = {c.casefold(): c for c in df.columns}
    opis_key = "opis oferty".casefold()
    if opis_key not in name_map:
        st.error("Brak kolumny 'Opis oferty' w pierwszym arkuszu (sprawdź pisownię/nagłówek w wierszu 1).")
    else:
        opis_col = name_map[opis_key]

        st.caption("Podgląd przed konwersją:")
        st.dataframe(df[[opis_col]].head(5))

        # Konwersja do czystego tekstu
        df[opis_col] = df[opis_col].apply(clean_html_to_plain)

        st.caption("Podgląd po konwersji do czystego tekstu:")
        st.dataframe(df[[opis_col]].head(5))

        # Zapis do Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=first_sheet_name)
        output.seek(0)

        st.download_button(
            label="Pobierz plik z czystym tekstem",
            data=output,
            file_name="oferty_czysty_tekst.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
