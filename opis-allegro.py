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

# Streamlit UI
st.title("Przetwarzanie opisów ofert z pliku Excel")

# Wgrywanie pliku Excel
uploaded_file = st.file_uploader("Wybierz plik Excel", type=["xlsx"])

if uploaded_file is not None:
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
