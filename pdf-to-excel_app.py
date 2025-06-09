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
    1. Próbuje wyciągnąć tekst przez PyPDF2, a gdy to się nie uda – przez pdfplumber.
    2. Próbuje znaleźć pełne linie z EAN-em i ilością (layout D):
       linia zaczyna się od 13 cyfr i zawiera ",<xx>,xx szt.".
    3. Jeśli znajdzie takie linie, zwraca je jako pozycje z automatycznie numerowanym Lp.
    4. W przeciwnym razie przechodzi do trybu uniwersalnego:
       - Aktualizuje bieżący EAN (każdy 13-cyfrowy ciąg).
       - Wykrywa numer pozycji (Lp) ze czystej linii z cyframi.
       - Wykrywa ilość przy słowie "szt." w tej samej linii.
       - Tworzy rekordy tylko, gdy ma wszystkie trzy elementy: Lp, EAN, ilość.
    5. Wynik wyświetla się w tabeli i można go pobrać jako Excel.
    """
)


def extract_text(pdf_bytes):
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
    # fallback
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


def parse_records(lines):
    # Layout D: pełne linie zaczynające się od EAN i zawierające ilość
    pattern_full = re.compile(r"^(\d{13}).*?(\d{1,4}),\d{2}\s*szt", flags=re.IGNORECASE)
    records = []
    lp_counter = 1
    for ln in lines:
        m = pattern_full.match(ln)
        if m:
            records.append({
                "Lp": lp_counter,
                "Symbol": m.group(1),
                "Ilość": int(m.group(2))
            })
            lp_counter += 1
    if records:
        return pd.DataFrame(records)

    # Tryb uniwersalny
    records = []
    current_ean = None
    current_lp = None
    for ln in lines:
        # Aktualizuj EAN
        m_ean = re.search(r"\b(\d{13})\b", ln)
        if m_ean:
            current_ean = m_ean.group(1)
            continue
        # Aktualizuj numer pozycji (Lp)
        m_lp = re.fullmatch(r"(\d{1,3})", ln)
        if m_lp:
            current_lp = int(m_lp.group(1))
            continue
        # Wykryj ilość w tej linii
        m_qty = re.search(r"(\d{1,4})\s*szt\.?", ln, flags=re.IGNORECASE)
        if m_qty and current_lp is not None and current_ean:
            qty = int(m_qty.group(1))
            records.append({
                "Lp": current_lp,
                "Symbol": current_ean,
                "Ilość": qty
            })
            current_lp = None  # reset po użyciu
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
def to_excel(df_in):
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
