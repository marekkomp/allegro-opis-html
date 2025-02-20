import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import io
import re

# Funkcja czyszcząca HTML i usuwająca niepożądane fragmenty JSON
def clean_html(content):
    # Usuwanie tylko niepotrzebnych fragmentów JSON-like (zostawiając tekst)
    content = re.sub(r'({"sections":.*?})', '', content)  # Usuwa JSON-like struktury
    content = re.sub(r'\[.*?\]', '', content)  # Usuwa inne tablice JSON, jeśli są

    # Parsowanie HTML
    soup = BeautifulSoup(content, 'html.parser')
    
    # Usuwamy wszystkie niepożądane tagi
    for tag in soup.find_all(True):  # True oznacza wszystkie tagi
        if tag.name not in ['h1', 'h2', 'p', 'b']:  # Pozostawiamy tylko h1, h2, p, b
            tag.unwrap()  # Usuwa tag, pozostawiając jego zawartość

    # Zwracamy oczyszczony HTML
    return str(soup)

# Streamlit UI
st.title("Przetwarzanie plików Excel i czyszczenie HTML w opisach ofert")

# Wgrywanie pliku Excel
uploaded_file = st.file_uploader("Wybaj plik Excel", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    # Wczytanie pliku Excel do DataFrame, obsługujemy zarówno .xlsm, jak i .xlsx
    df = pd.read_excel(uploaded_file, sheet_name=None)  # Wczytanie wszystkich arkuszy
    
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
