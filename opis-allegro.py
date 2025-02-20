import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import io
import json

# Funkcja czyszcząca HTML i usuwająca niepotrzebne fragmenty JSON
def clean_html(content):
    # Usuwamy fragmenty JSON-like, zachowując tylko dane zawarte w "content"
    try:
        data = json.loads(content)  # Próbujemy zinterpretować zawartość jako JSON
        # Sprawdzamy, czy to jest lista "items" i wyciągamy tylko zawartość typu "TEXT"
        clean_content = ""
        for section in data.get("sections", []):
            for item in section.get("items", []):
                if item.get("type") == "TEXT" and "content" in item:
                    clean_content += item["content"]
        return clean_content
    except json.JSONDecodeError:
        # Jeśli nie jest to poprawny JSON, po prostu zwracamy oryginalny tekst
        return content

# Streamlit UI
st.title("Przetwarzanie plików Excel i czyszczenie HTML w opisach ofert")

# Wgrywanie pliku Excel
uploaded_file = st.file_uploader("Wybierz plik Excel", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    # Wczytanie pliku Excel do DataFrame, pomijając pierwsze 3 wiersze i traktując czwarty wiersz jako nagłówek
    df = pd.read_excel(uploaded_file, sheet_name=None, header=3)  # Wczytujemy z pominięciem pierwszych trzech wierszy
    
    # Sprawdzenie, czy kolumna "Opis oferty" istnieje w pierwszym arkuszu
    sheet_names = df.keys()
    first_sheet_name = list(sheet_names)[0]
    df = df[first_sheet_name]  # Wybór pierwszego arkusza, jeśli jest ich więcej
    
    if "Opis oferty" in df.columns:
        st.write("Oryginalne dane:")
        st.write(df[["Opis oferty"]].head())

        # Przetwarzanie kolumny "Opis oferty"
        df['Opis oferty'] = df['Opis oferty'].apply(clean_html)

        # Pokazanie przetworzonych danych
        st.write("Przetworzone dane:")
        st.write(df[["Opis oferty"]].head())

        # Przygotowanie pliku do pobrania
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=first_sheet_name)
        output.seek(0)

        # Opcja pobrania poprawionego pliku
        st.download_button(
            label="Pobierz poprawiony plik Excel",
            data=output,
            file_name="poprawiony_oferty.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Brak kolumny 'Opis oferty' w pliku Excel.")
