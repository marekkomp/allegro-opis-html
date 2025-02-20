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

# Funkcja do zmiany nazw powtarzających się kolumn
def rename_duplicate_columns(df):
    # Słownik do przechowywania liczby wystąpień kolumn
    columns_count = {}
    
    # Zmieniamy nazwę kolumn, które się powtarzają
    new_columns = []
    for column in df.columns:
        if column in columns_count:
            columns_count[column] += 1
            new_column_name = f"{column}_{columns_count[column]}"  # Dodajemy licznik do powtarzającej się kolumny
        else:
            columns_count[column] = 0
            new_column_name = column
        new_columns.append(new_column_name)

    # Zmieniamy nazwy kolumn
    df.columns = new_columns
    return df

# Streamlit UI
st.title("Przetwarzanie plików Excel i czyszczenie HTML w opisach ofert")

# Wgrywanie pliku Excel
uploaded_file = st.file_uploader("Wybierz plik Excel", type=["xlsm", "xlsx"])

if uploaded_file is not None:
    # Wczytanie pliku Excel do DataFrame, traktując pierwszy wiersz jako nagłówek
    df = pd.read_excel(uploaded_file, sheet_name=None, header=0)  # Wczytujemy z pierwszym wierszem jako nagłówek
    
    # Sprawdzenie, czy kolumna "Opis oferty" istnieje w pierwszym arkuszu
    sheet_names = df.keys()
    first_sheet_name = list(sheet_names)[0]
    df = df[first_sheet_name]  # Wybór pierwszego arkusza, jeśli jest ich więcej
    
    # Zmiana nazw powtarzających się kolumn
    df = rename_duplicate_columns(df)

    # Lista kolumn do zachowania
    columns_to_keep = [
        "Status", "Rezultat", "ID oferty", "Link do oferty", "Akcja", "Status oferty",
        "Kategoria główna", "Podkategoria", "Sygnatura/SKU Sprzedającego", "Liczba sztuk", 
        "Cena PL", "Tytuł oferty", "Zdjęcia", "Opis oferty", "Informacje o gwarancjach (opcjonalne)", 
        "Stan", "Model", "Marka", "Rodzaj", "Typ baterii", "Przekątna [\"]", 
        "Rozdzielczość (px)", "Powłoka matrycy", "Rodzaj podświetlenia", "Typ matrycy", 
        "Rodzaj podświetlania", "Przekątna ekranu (cale) [\"]", "Rozdzielczość natywna [px]", 
        "Złącza", "Typ napędu", "Komunikacja", "Rodzaj karty graficznej", "System operacyjny", 
        "Seria", "Taktowanie bazowe procesora [GHz]", "Liczba rdzeni procesora", "Typ pamięci RAM", 
        "Wielkość pamięci RAM", "Typ dysku twardego", "Pojemność dysku [GB]", "Przekątna ekranu [\"]", 
        "Seria procesora", "Ekran dotykowy", "Model procesora", "Typ obudowy", "Model procesora"
    ]
    
    # Zachowanie tylko wybranych kolumn
    df = df[columns_to_keep]

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
