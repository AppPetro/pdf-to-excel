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
       linia zaczyna się od 13 cyfr i zawiera ",<xx> szt.".
    3. Jeśli znajdzie takie linie, zwraca je jako pozycje z automatycznie numerowanym Lp.
    4. W przeciwnym razie przechodzi do trybu uniwersalnego:
       - Tworzy mapę wszystkich 13-cyfrowych EAN-ów w tekście.
       - Przetwarza linie: aktualizuje bieżący EAN, rozpoznaje Lp i ilość przy słowie "szt.".
       - Dla każdej wykrytej ilości zapisuje wiersz {Lp, Symbol: aktualny EAN, Ilość}.
    5. Wynik wyświetla się w tabeli i można go pobrać jako Excel.
    """
)


def extract_text(pdf_bytes: bytes) -> list[str]:
    # PyPDF2
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        lines = []
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
    # 1) Layout D – linie zaczynające się od 13 cyfr i zawierające ',<qty> szt.'
    pattern_full = re.compile(r"^(\d{13}).*?(\d{1,4}),\d{2}\s*szt", re.IGNORECASE)
    records = []
    lp_counter = 1
    for ln in lines:
        m = pattern_full.match(ln)
        if m:
            ean = m.group(1)
            qty = int(m.group(2))
            records.append({"Lp": lp_counter, "Symbol": ean, "Ilość": qty})
            lp_counter += 1
    if records:
        return pd.DataFrame(records)

    # 2) Tryb uniwersalny
    records = []
    current_ean = None
    current_lp = None

    for ln in lines:
        # aktualizuj EAN, jeśli w linii jest 13-cyfrowy ciąg
        m_ean = re.search(r"\b(\d{13})\b", ln)
        if m_ean:
            current_ean = m_ean.group(1)

        # wykryj Lp jako osobną linię z kilkoma cyframi
        m_lp = re.fullmatch(r"(\d{1,3})", ln)
        if m_lp:
            current_lp = int(m_lp.group(1))

        # wykryj ilość przy słowie 'szt.'
        m_qty = re.search(r"(\d{1,4})\s*szt\.?",
