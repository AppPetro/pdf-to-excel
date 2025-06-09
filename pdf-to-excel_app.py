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


def extract_text(pdf_bytes: bytes) -> list[str]:
    """Wyciąga linie tekstu z PDF-a za pomocą PyPDF2, fallback pdfplumber."""
    # PyPDF2
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        lines: list[str] = []
        for page in reader.pages:
            text = page.extract_text() or ""
            lines += [ln.strip() for ln in text.split("\n") if ln.strip()]
        if lines:
            return lines
    except Exception:
        pass
    # fallback pdfplumber
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            lines = []
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines += [ln.strip() for ln in text.split("\n") if ln.strip()]
            return lines
    except Exception:
        return []


def parse_records(lines: list[str]) -> pd.DataFrame:
    """Parsuje linie i zwraca DataFrame z kolumnami Lp, Symbol (EAN), Ilość."""
    # 1) Layout D – pełne linie: EAN + ",<qty>,xx szt."
    pattern_full = re.compile(r"^(\d{13}).*?(\d{1,4}),\d{2}\s*szt", flags=re.IGNORECASE)
    records: list[dict] = []
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

    # 2) Tryb uniwersalny
    records = []
    current_ean: str | None = None
    current_lp: int | None = None

    for ln in lines:
        # aktualizuj obecny EAN, jeśli w linii jest 13-cyfrowy ciąg
        m_ean = re.search(r"\b(\d{13})\b", ln)
        if m_ean:
            current_ean = m_ean.group(1)

        # wykryj Lp jako osobną linię zawierającą tylko cyfry
        m_lp = re.fullmatch(r"(\d{1,3})", ln)
        if m_lp:
            current_lp = int(m_lp.group(1))

        # wykryj ilość w tej samej linii
        m_qty = re.search(r"(\d{1,4})\s*szt\.?",
