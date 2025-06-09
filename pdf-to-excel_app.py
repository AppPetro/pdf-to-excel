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
    1. Wyciąga każdą linię tekstu z PDF za pomocą PyPDF2 lub pdfplumber.
    2. Używa uniwersalnego parsera:
       - Śledzi każdy EAN (13 cyfr) z prefiksem „Kod kres.” lub dowolnie w tekście.
       - Gdy linia zaczyna się od numeru pozycji i zawiera „szt.”, wyciąga ilość i łączy z ostatnim EAN.
    3. Wyświetla tabelę z kolumnami Lp, Symbol (EAN), Ilość i pozwala pobrać plik Excel.
    """
)


def extract_text(pdf_bytes: bytes) -> list[str]:
    # próba PyPDF2
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
            lines: list[str] = []
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines += [ln.strip() for ln in text.split("\n") if ln.strip()]
            return lines
    except Exception:
        return []


def parse_generic(lines: list[str]) -> pd.DataFrame:
    products = []
    last_ean: str | None = None
    for ln in lines:
        # aktualizuj EAN prefiksem "Kod kres." lub dowolny 13-cyfrowy ciąg
        m_pref = re.search(r"Kod\s+kres.*?(\d{13})", ln)
        if m_pref:
            last_ean = m_pref.group(1)
            continue
        m_any = re.search(r"\b(\d{13})\b", ln)
        if m_any:
            last_ean = m_any.group(1)
        # wykryj linię produktu: zaczyna się od numeru i zawiera 'szt.'
        m_lp = re.match(r"^(\d{1,2})", ln)
        if m_lp and 'szt' in ln.lower() and last_ean:
            lp = int(m_lp.group(1))
            m_qty = re.search(r"(\d{1,4})\s*szt", ln, flags=re.IGNORECASE)
            if m_qty:
                qty = int(m_qty.group(1))
                products.append({"Lp": lp, "Symbol": last_ean, "Ilość": qty})
                last_ean = None  # kolejny produkt wymaga nowego EAN
    return pd.DataFrame(products)


# Główna logika
uploaded = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded.read()
lines = extract_text(pdf_bytes)
if not lines:
    st.error("Nie udało się wyciągnąć tekstu z PDF. Spróbuj OCR-em.")
    st.stop()

df = parse_generic(lines)
# usuwamy puste i resetujemy indeks
if 'Ilość' in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)
else:
    st.error("Nie znaleziono kolumny 'Ilość'. Parsowanie nie powiodło się.")
    st.stop()

if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# eksport do Excela
def to_excel(df_input: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_input.to_excel(writer, index=False, sheet_name="Zamówienie")
    return buffer.getvalue()

st.download_button(
    label="Pobierz wynik jako Excel",
    data=to_excel(df),
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
