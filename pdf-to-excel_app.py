import streamlit as st
import pandas as pd
import re
import io
import PyPDF2
import pdfplumber

# Konfiguracja aplikacji
st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Wydobywa tekst przez PyPDF2, a jeśli nie zadziała – przez pdfplumber.
    2. Scala wiersze rozbite między stronami (kontynuacje bez numeru LP).
    3. Wykrywa układ i parsuje tabelę **Lp | Symbol | Ilość**.
    4. Sprawdza, czy liczba pozycji (max Lp) zgadza się z liczbą unikalnych EAN-ów.
    5. Umożliwia pobranie wyników jako Excel.
    """
)

# Regex wykrywające początek nowego wiersza (numer LP) i EAN
NEW_ROW_REGEX = re.compile(r"^\s*\d+\s+")
EAN_REGEX = re.compile(r"(\d{8,13})")


def merge_continued_lines(lines: list[str]) -> list[str]:
    """
    Scala fragmenty tekstu: jeśli linia nie zaczyna się od LP, dokleja ją
    do poprzedniej jako kontynuację.
    """
    merged = []
    for ln in lines:
        text = ln.strip()
        if not text:
            continue
        if NEW_ROW_REGEX.match(text):
            merged.append(text)
        else:
            # kontynuacja poprzedniego wiersza
            if merged:
                merged[-1] += ' ' + text
            else:
                merged.append(text)
    return merged


def extract_text_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        return []
    lines = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            ln = ln.strip()
            if ln:
                lines.append(ln)
    return lines


def extract_text_with_pdfplumber(pdf_bytes: bytes) -> list[str]:
    lines = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    ln = ln.strip()
                    if ln:
                        lines.append(ln)
    except Exception:
        return []
    return lines


def parse_layout_default(all_lines: list[str]) -> pd.DataFrame:
    """
    Prosty parser domyślny – znajduje LP, ilość i EAN w jednej linii lub w kontynuacji.
    Linia musi zawierać LP i ilość, EAN musi być w fragmencie "Kod kres".
    """
    pat_item = re.compile(r"^(\d+)\s+.*?(\d{1,3})\s+szt", re.IGNORECASE)
    pat_ean  = re.compile(r"kod kres\.?\s*:\s*(\d{8,13})", re.IGNORECASE)

    products = []
    for ln in all_lines:
        m_item = pat_item.search(ln)
        if not m_item:
            continue
        lp = int(m_item.group(1))
        qty = int(m_item.group(2))
        # EAN w tej samej linii?
        m_ean = pat_ean.search(ln)
        symbol = m_ean.group(1) if m_ean else ''
        # jeśli nie, spróbuj znaleźć w fragmencie po spacji
        if not symbol and 'kod kres' in ln.lower():
            e = EAN_REGEX.search(ln)
            symbol = e.group(1) if e else ''
        products.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})

    return pd.DataFrame(products)


# Główna logika
uploaded = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()
pdf_bytes = uploaded.read()

# 1) Ekstrakcja tekstu
lines = extract_text_with_pypdf2(pdf_bytes)
if not lines:
    lines = extract_text_with_pdfplumber(pdf_bytes)
    if not lines:
        st.error("Nie udało się wyciągnąć tekstu z PDF-a.")
        st.stop()

# 2) Scalanie kontynuacji
merged = merge_continued_lines(lines)

# 3) Parsowanie (domyślny uniwersalny parser)
df = parse_layout_default(merged)

# Walidacja
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji.")
    st.stop()
max_lp     = int(df['Lp'].max())
unique_ean = df['Symbol'].nunique()
if max_lp != unique_ean:
    st.warning(
        f"Znaleziono {max_lp} pozycji, ale tylko {unique_ean} unikalnych kodów EAN – sprawdź parsowanie."
    )

# Wyświetlenie i pobranie
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)


def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamówienie")
    return buf.getvalue()

st.download_button(
    label="Pobierz jako Excel",
    data=convert_df_to_excel(df),
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
