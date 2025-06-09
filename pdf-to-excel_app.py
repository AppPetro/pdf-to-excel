import streamlit as st
import pandas as pd
import re
import io
import PyPDF2
import pdfplumber

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Wyciąga tekst przez PyPDF2 lub pdfplumber.
    2. Parser:
       - Skupia się na wykrywaniu kolejnych EAN-ów (13 cyfr) i ilości (\"<liczba> szt.\").
       - Gdy znajdzie EAN i później ilość, tworzy rekord {Lp, Symbol: EAN, Ilość}.
    3. Wyświetla tabelę i umożliwia pobranie pliku Excel.
    """
)

def extract_text(pdf_bytes: bytes) -> list[str]:
    """Wyciąga linie tekstu z PDF -> listę niepustych stringów."""
    # PyPDF2
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        lines = []
        for page in reader.pages:
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                ln = ln.strip()
                if ln:
                    lines.append(ln)
        if lines:
            return lines
    except Exception:
        pass
    # pdfplumber fallback
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            lines = []
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    ln = ln.strip()
                    if ln:
                        lines.append(ln)
            return lines
    except Exception:
        return []


def parse_records(lines: list[str]) -> pd.DataFrame:
    """Parsuje EAN i odpowiadające im ilości w kolejności występowania."""
    records = []
    lp_counter = 1
    last_ean = None
    # skanuj linia po linii
    for ln in lines:
        # wykryj EAN (13 cyfr)
        m_ean = re.search(r"\b(\d{13})\b", ln)
        if m_ean:
            last_ean = m_ean.group(1)
        # wykryj ilość przed 'szt'
        m_qty = re.search(r"(\d{1,4})\s*szt", ln, flags=re.IGNORECASE)
        if m_qty and last_ean:
            qty = int(m_qty.group(1))
            records.append({"Lp": lp_counter, "Symbol": last_ean, "Ilość": qty})
            lp_counter += 1
            last_ean = None  # reset po użyciu
    return pd.DataFrame(records)


# Główna logika
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded_file:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()
lines = extract_text(pdf_bytes)
if not lines:
    st.error("Nie udało się wyciągnąć tekstu z tego PDF-a. Spróbuj użyć OCR.")
    st.stop()

df = parse_records(lines)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# Przygotuj Excel
def to_excel(df_in: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamówienie")
    return out.getvalue()

excel_data = to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=excel_data,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
