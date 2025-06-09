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
    1. Próbuje wyciągnąć tekst przez PyPDF2.
    2. Jeśli nie wykryje pełnych linii z EAN+ilość (“layout D”), 
       używa uniwersalnego parsera, który:
       - Buduje mapę wszystkich 13-cyfrowych EAN-ów w tekście.
       - Szuka wierszy zaczynających się od Lp i zawierających “…<ilość> szt.”
       - Łączy każdą pozycję z najbliższym wcześniejszym EAN-em.
    3. Wyświetla wynik w tabeli **Lp | Symbol | Ilość** i pozwala pobrać Excel.
    """
)

def extract_text(pdf_bytes: bytes) -> list[str]:
    # najpierw PyPDF2
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
    # fallback na pdfplumber
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            lines = []
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines += [ln.strip() for ln in text.split("\n") if ln.strip()]
            return lines
    except Exception:
        return []

def parse_layout_d(lines: list[str]) -> pd.DataFrame:
    """
    Layout D: linia zaczyna się od 13 cyfr, potem gdzieś ',<xx> szt'
    """
    pattern = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
    out, lp = [], 1
    for ln in lines:
        m = pattern.match(ln)
        if m:
            out.append({"Lp": lp, "Symbol": m.group(1), "Ilość": int(m.group(2))})
            lp += 1
    return pd.DataFrame(out)

def parse_generic(lines: list[str]) -> pd.DataFrame:
    """
    Uniwersalny parser B/CE:
    - Buduje mapę wszystkich wystąpień 13-cyfrowych EAN-ów (dowolnie w liniach).
    - Szuka wierszy typu:
         ^(\d{1,3})\D+.*?(\d{1,4})\s*szt\.?
      gdzie:
        1) grupa1 = numer Lp,
        2) grupa2 = ilość.
    - Każdy taki wiersz łączy z najbliższym wcześniejszym EAN-em z mapy.
    """
    # 1) Mapa EAN-ów
    ean_idx = {}
    for i, ln in enumerate(lines):
        for ean in re.findall(r"\b(\d{13})\b", ln):
            ean_idx[i] = ean

    # 2) Szukamy wierszy z Lp i ilością
    pattern = re.compile(r"^(\d{1,3})\D+.*?(\d{1,4})\s*szt\.?", re.IGNORECASE)
    out = []
    for i, ln in enumerate(lines):
        m = pattern.match(ln)
        if not m:
            continue
        lp  = int(m.group(1))
        qty = int(m.group(2))
        # znajdź najbliższy wcześniejszy EAN
        prev = [idx for idx in ean_idx if idx < i]
        if not prev:
            continue
        symbol = ean_idx[max(prev)]
        out.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})

    # posortuj po Lp
    df = pd.DataFrame(out)
    return df.sort_values("Lp").reset_index(drop=True)

# ────── G Ł Ó W N A   L O G I K A ──────

uploaded = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

txt = extract_text(uploaded.read())
if not txt:
    st.error("Nie udało się wyciągnąć tekstu z tego PDF-a. Spróbuj OCR-em.")
    st.stop()

# 1) Spróbuj layout D
df = parse_layout_d(txt)

# 2) Jeśli pusto, albo zbyt mało wierszy, użyj parsera generic
if df.empty or len(df) < 5:
    df = parse_generic(txt)

# Odfiltruj puste i uporządkuj
df = df.dropna(subset=["Ilość"]).reset_index(drop=True)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

def to_excel(df_in: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df_in.to_excel(w, index=False, sheet_name="Zamówienie")
    return out.getvalue()

data_xl = to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=data_xl,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
