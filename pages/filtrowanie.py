import streamlit as st
import pandas as pd
import io
import os

# Tytuł aplikacji
st.title('Filtrowanie arkusza Excel')

st.html('Skrypt filtruje arkusz excela zgodnie z wybraną przez nas kolumną oraz podanymi adresami url. <br /> W pole można również wkleić zwykłe frazy. Finalny plik zawiera wiersze zgodne z filtrem, nawet jeżeli występują w danej kolumnie wielokrotnie.')

# Wczytaj plik Excel
uploaded_file = st.file_uploader("Wybierz plik Excel", type="xlsx")

if uploaded_file is not None:
    # Wczytaj dane z przesłanego pliku Excel
    @st.cache_data
    def load_data(file):
        return pd.read_excel(file)

    df = load_data(uploaded_file)

    # Wyświetl dostępne kolumny i daj możliwość wyboru kolumny
    selected_column = st.selectbox("Wybierz kolumnę do filtrowania", df.columns)

    # Okno do wklejenia listy URL
    url_input = st.text_area("Wklej listę URL (każdy URL w nowej linii)")

    # Dodaj wybór, co zrobić z listą URL-i
    action = st.radio("Wybierz akcję:", ('Zostaw tylko adresy z listy', 'Usuń z pliku adresy z listy'))

    # Konsola i pasek postępu
    progress_bar = st.progress(0)
    console = st.empty()

    # Przyciski do przetwarzania
    if st.button('Działaj!'):
        if url_input:
            # Pobierz listę URL-i z formularza (każdy URL w nowej linii)
            url_list = [url.strip() for url in url_input.splitlines() if url.strip()]

            console.write("Rozpoczynam filtrowanie danych...")

            # Filtrowanie w zależności od wybranej akcji
            if action == 'Zostaw tylko adresy z listy':
                filtered_df = df[df[selected_column].isin(url_list)]
                console.write("Zachowano tylko adresy z listy.")
            else:  # Usuń adresy z listy
                filtered_df = df[~df[selected_column].isin(url_list)]
                console.write("Usunięto adresy z listy.")

            progress_bar.progress(50)

            # Przygotowanie nazwy pliku wyjściowego
            original_file_name = os.path.splitext(uploaded_file.name)[0]
            output_file_name = f"{original_file_name}_przefiltrowany.xlsx"

            # Zapisz przefiltrowane dane do pliku Excel w pamięci
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False)

            progress_bar.progress(100)

            # Pobierz przefiltrowany plik
            st.success("Przefiltrowano dane! Kliknij poniżej, aby pobrać plik.")
            console.write("Przetwarzanie zakończone.")

            # Przygotuj plik do pobrania
            st.download_button(
                label="Pobierz przefiltrowany plik Excel",
                data=output.getvalue(),
                file_name=output_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Musisz wprowadzić listę URL.")
            console.write("Błąd: lista URL jest pusta.")
