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
    2. Najpierw sprawdza, czy mamy "pełne" linie z EAN-em i ilością (layout D):
       linia zaczyna się od 13 cyfr i zawiera ",<xx> szt.".
    3. Jeśli znajdzie takie linie, parsuje je jako oddzielne pozycje z autoinkrementowanym Lp.
    4. W przeciwnym razie przechodzi do trybu uniwersalnego:
       - Buduje się mapa wszystkich 13-cyfrowych EAN-ów w tekście.
       - Przechodzi przez kolejne linie, śledząc najnowszy EAN oraz numer Lp (czysta linia z cyframi).
       - Gdy natrafi na ilość (liczba + "szt."), zapisuje wpis {Lp, Symbol: aktualny EAN, Ilość}.
    5. Wynik wyświetla się w tabeli i można pobrać go jako plik Excel.
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
    # 1) Spróbuj layout D (pełna linia: EAN + ",<qty> szt.")
    pattern_full = re.compile(r"^(\d{13}).*?(\d{1,3}),\d{2}\s*szt", re.IGNORECASE)
    out = []
    lp_counter = 1
    for ln in lines:
        m = pattern_full.match(ln)
        if m:
            out.append({"Lp": lp_counter, "Symbol": m.group(1), "Ilość": int(m.group(2))})
            lp_counter += 1
    if out:
        return pd.DataFrame(out)

    # 2) Tryb uniwersalny
    # mapa ean -> jego ostatnie wystąpienie
    ean_idx = {}
    for i, ln in enumerate(lines):
        for ean in re.findall(r"\b(\d{13})\b", ln):
            ean_idx[i] = ean

    records = []
    current_ean = None
    current_lp = None
    for ln in lines:
        # aktualizuj ean, jeśli jest w linii
        m_ean = re.search(r"\b(\d{13})\b", ln)
        if m_ean:
            current_ean = m_ean.group(1)

        # wykryj Lp jako osobną linię z cyframi
        m_lp = re.fullmatch(r"(\d{1,3})", ln)
        if m_lp:
            current_lp = int(m_lp.group(1))

        # wykryj ilość przed 'szt.'
        m_qty = re.search(r"(\d{1,4})\s*szt\.?",
