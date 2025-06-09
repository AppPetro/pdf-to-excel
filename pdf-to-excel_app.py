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
    2. W zależności od układu (D/E/B/C/A) parsuje pozycje.
    3. Wyświetla wynik i umożliwia pobranie Excela.
    """
)


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


def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    pattern = re.compile(r"^(\d{13}).*?(\d{1,3}),\d{2}\s*szt", flags=re.IGNORECASE)
    products = []
    lp = 1
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            products.append({"Lp": lp, "Symbol": m.group(1), "Ilość": int(m.group(2))})
            lp += 1
    return pd.DataFrame(products)


def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern_item = re.compile(r"^(\d+)\s+.*?(\d{1,3})\s+szt\.", flags=re.IGNORECASE)
    i = 0
    while i < len(all_lines):
        ln = all_lines[i]
        m = pattern_item.match(ln)
        if m:
            lp = int(m.group(1))
            qty = int(m.group(2))
            # znajdź EAN w kolejnych 3 wierszach
            ean = None
            for j in range(i+1, min(i+4, len(all_lines))):
                found = re.search(r"(\d{13})", all_lines[j])
                if found:
                    ean = found.group(1)
                    break
            products.append({"Lp": lp, "Symbol": ean or "", "Ilość": qty})
            i += 1
        else:
            i += 1
    return pd.DataFrame(products)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern = re.compile(r"^(\d+)\s+(\d{13})\s+.*?(\d{1,3}),\d{2}\s*szt", flags=re.IGNORECASE)
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            products.append({"Lp": int(m.group(1)), "Symbol": m.group(2), "Ilość": int(m.group(3))})
    return pd.DataFrame(products)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    # Zbieramy wszystkie indeksy linii z EAN
    ean_map = {}
    for idx, ln in enumerate(all_lines):
        m = re.search(r"\b(\d{13})\b", ln)
        if m:
            ean_map[idx] = m.group(1)
    products = []
    # Szukamy linii z Lp, a potem ilości
    for idx, ln in enumerate(all_lines):
        # wykryj Lp
        if re.fullmatch(r"\d+", ln):
            lp = int(ln)
            # znajdź najbliższy wcześniejszy EAN
            prev_idxs = [i for i in ean_map if i < idx]
            if not prev_idxs:
                continue
            symbol = ean_map[max(prev_idxs)]
            # szukamy ilości w następnych 3 liniach
            qty = None
            for j in range(idx+1, min(idx+4, len(all_lines))):
                m2 = re.search(r"(\d{1,4})\s*szt", all_lines[j], flags=re.IGNORECASE)
                if m2:
                    qty = int(m2.group(1))
                    break
            if qty is not None:
                products.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})
    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    # fallback - analogicznie do C, ale z 'Kod kres'
    products = []
    ean_idx = {i: re.search(r"(\d{13})", ln).group(1)
               for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres") and re.search(r"\d{13}", ln)}
    for idx, ln in enumerate(all_lines):
        if re.fullmatch(r"\d+", ln):
            lp = int(ln)
            prevs = [i for i in ean_idx if i < idx]
            if not prevs:
                continue
            symbol = ean_idx[max(prevs)]
            # ilość w 2 liniach
            qty = None
            if idx+2 < len(all_lines) and all_lines[idx+2].lower().startswith("szt"):
                if re.fullmatch(r"\d+", all_lines[idx+1]):
                    qty = int(all_lines[idx+1])
            if qty:
                products.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})
    return pd.DataFrame(products)

# Główna logika
uploaded = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded.read()
# najpierw PyPDF2
lines_py = extract_text_with_pypdf2(pdf_bytes)
lines = lines_py if lines_py else extract_text_with_pdfplumber(pdf_bytes)

# detekcja layoutu
df = pd.DataFrame()
if any(re.match(r"^\d{13}.*\d{1,3},\d{2}\s*szt", ln) for ln in lines):
    df = parse_layout_d(lines)
elif any(re.match(r"^\d+.*\d{1,3}\s+szt", ln, flags=re.IGNORECASE) for ln in lines):
    df = parse_layout_e(lines)
elif any(re.match(r"^\d+\s+\d{13}", ln) for ln in lines):
    df = parse_layout_b(lines)
elif any(re.fullmatch(r"\d{13}", ln) for ln in lines):
    df = parse_layout_c(lines)
else:
    df = parse_layout_a(lines)

# finalizacja
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

st.download_button(
    label="Pobierz wynik jako Excel",
    data=to_excel(df),
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
