import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import io

# Funkcja czyszcząca HTML
def clean_html(content):
    # Parsujemy HTML
    soup = BeautifulSoup(content, 'html.parser')
    
    # Usuwamy wszystkie niepożądane tagi
    for tag in soup.find_all(True):  # True oznacza wszystkie tagi
        if tag.name not in ['h1', 'h2', 'p', 'b']:  # Pozostawiamy tylko h1, h2, p, b
            tag.unwrap()  # Usuwa tag, pozostawiając jego zawartość

    # Zwracamy oczyszczony HTML
    return str(soup)

# Funkcja konwertująca plik .xlsm do .xlsx
def convert_xlsm_to_xlsx(file):
    # Wczytaj plik XLSM
    df = pd.read_excel(file, sheet_name=None)  # Wczytaj wszystkie arkusze
    
    # Zapisz go jako plik XLSX
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, data in df.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

# Streamlit UI
st.title("Przetwarzanie plików Excel i czyszczenie HTML w opisach ofert")

# Wgrywanie pliku Excel
uploaded_file = st.file_uploader("Wybierz plik Excel", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    # Obsługa pliku XLSM - konwersja do XLSX
    if uploaded_file.name.endswith('.xlsm'):
        st.write("Plik XLSM wykryty. Trwa konwersja do XLSX...")
        converted_file = convert_xlsm_to_xlsx(uploaded_file)
        st.download_button(
            label="Pobierz plik XLSX",
            data=converted_file,
            file_name="converted_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        # Wczytanie pliku Excel do DataFrame
        df = pd.read_excel(uploaded_file)

        # Sprawdzenie, czy kolumna "Opis oferty" istnieje
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
                df.to_excel(writer, index=False, sheet_name='Przetworzone oferty')
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
