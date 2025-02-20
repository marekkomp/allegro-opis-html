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

# Lista kolumn do zachowania w pliku wynikowym
columns_to_keep = [
    'Status', 'Rezultat', 'ID oferty', 'Link do oferty', 'Akcja', 'Status oferty', 
    'Kategoria główna', 'Podkategoria', 'Sygnatura/SKU Sprzedającego', 'Liczba sztuk', 
    'Cena PL', 'Tytuł oferty', 'Zdjęcia', 'Opis oferty', 'Informacje o gwarancjach (opcjonalne)', 
    'Stan', 'Model', 'Marka', 'Rodzaj', 'Typ baterii', 'Przekątna ["]', 'Rozdzielczość (px)', 
    'Powłoka matrycy', 'Rodzaj podświetlenia', 'Typ matrycy', 'Rodzaj podświetlania', 
    'Przekątna ekranu (cale) ["]', 'Rozdzielczość natywna [px]', 'Złącza', 'Typ napędu', 
    'Komunikacja', 'Rodzaj karty graficznej', 'System operacyjny', 'Seria', 'Taktowanie bazowe procesora [GHz]', 
    'Liczba rdzeni procesora', 'Typ pamięci RAM', 'Wielkość pamięci RAM', 'Typ dysku twardego', 
    'Pojemność dysku [GB]', 'Przekątna ekranu ["', 'Seria procesora', 'Ekran dotykowy', 
    'Model procesora', 'Typ obudowy', 'Model procesora1'
]

# Streamlit UI
st.title("Przetwarzanie plików Excel i czyszczenie HTML w opisach ofert")

# Wgrywanie pliku Excel
uploaded_file = st.file_uploader("Wybierz plik Excel", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    # Wczytanie pliku Excel do DataFrame, traktując pierwszy wiersz jako nagłówek
    df = pd.read_excel(uploaded_file, sheet_name=None, header=0)  # Wczytujemy z pierwszym wierszem jako nagłówek
    
    # Sprawdzenie, jakie kolumny zostały wczytane
    st.write("Kolumny w wczytanym pliku:")
    st.write(df.keys())  # Wypisujemy nazwy arkuszy
    
    # Wybieramy pierwszy arkusz z pliku
    sheet_names = df.keys()
    first_sheet_name = list(sheet_names)[0]
    df = df[first_sheet_name]  # Wybór pierwszego arkusza
    
    # Wyświetlenie wczytanych kolumn
    st.write("Wczytane kolumny:")
    st.write(df.columns)

    # Sprawdzenie, czy kolumna "Opis oferty" istnieje
    if "Opis oferty" in df.columns:
        st.write("Oryginalne dane:")
        st.write(df[["Opis oferty"]].head())

        # Przetwarzanie kolumny "Opis oferty"
        df['Opis oferty'] = df['Opis oferty'].apply(clean_html)

        # Selekcja tylko wybranych kolumn
        df_selected = df[columns_to_keep]

        # Pokazanie przetwor
