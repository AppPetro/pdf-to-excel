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
    1. Próbuje wyciągnąć tekst przez PyPDF2 (stare „trudniejsze” PDF-y).
    2. Jeśli w wyciągniętym przez PyPDF2 tekście nie występują układy D ani E, 
       używa starych parserów (układ B, C lub A).
    3. W przeciwnym razie (lub gdy PyPDF2 nie wyciągnie w ogóle linii) 
       wyciąga tekst przez pdfplumber i próbuje wykryć układy:
       - **Układ D**: linie zawierające EAN i ilość w jednym wierszu.
       - **Układ E**: Lp i ilość, poniżej linia “Kod kres.: <EAN>”.
       - **Układ B**: jedna linia: Lp, EAN, nazwa, ilość.
       - **Układ C**: EAN w osobnej linii, potem Lp, nazwa, “szt.”, ilość.
       - **Układ A**: “Kod kres.: <EAN>” w osobnej linii, Lp w osobnej linii, fragmenty nazwy przed i po liczbie.
    4. Wyświetla wynik w tabeli **Lp | Symbol | Ilość** i pozwala pobrać go jako Excel.
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
            stripped = ln.strip()
            if stripped:
                lines.append(stripped)
    return lines

def extract_text_with_pdfplumber(pdf_bytes: bytes) -> list[str]:
    lines = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    stripped = ln.strip()
                    if stripped:
                        lines.append(stripped)
    except Exception:
        return []
    return lines

def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
    lp = 1
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            ean = m.group(1)
            qty = int(m.group(2))
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
            lp += 1
    return pd.DataFrame(products)

def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    products = []
    # znajdź wszystkie linie "Kod kres.: <EAN>"
    ean_idx = {i: re.search(r"(\d{13})", ln).group(1)
               for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")}
    # znajdź pure-digit Lp
    for i, ln in enumerate(all_lines):
        if re.fullmatch(r"\d+", ln):
            lp = int(ln)
            # szukaj ilości i powiązanej nazwy w kolejnych liniach
            qty = None
            for j in range(i+1, min(i+6, len(all_lines))):
                # łap "120 szt." w tej samej linii
                m = re.search(r"(\d{1,3})\s*szt\.?", all_lines[j], flags=re.IGNORECASE)
                if m:
                    qty = int(m.group(1))
                    break
                # lub linię "szt." i poprzedzającą liczbę
                if all_lines[j].lower().strip() == "szt." and j-1 > i:
                    prev = all_lines[j-1].strip()
                    if re.fullmatch(r"\d+", prev):
                        qty = int(prev)
                        break
            # znajdź najbliższy wcześniejszy EAN
            prev_eans = [idx for idx in ean_idx if idx < i]
            if qty is not None and prev_eans:
                ean = ean_idx[max(prev_eans)]
                products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)

def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern = re.compile(r"^(\d+)\s+(\d{13})\s+.+?\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            products.append({
                "Lp": int(m.group(1)),
                "Symbol": m.group(2),
                "Ilość": int(m.group(3))
            })
    return pd.DataFrame(products)

def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ C – każdy 13-cyfrowy ciąg w linii traktujemy jako EAN,
    potem Lp i qty mogą być osobno (numer, nazwa, qty + 'szt.').
    """
    # 1) mapuj wszystkie EAN-y
    ean_idx = {}
    for i, ln in enumerate(all_lines):
        m = re.search(r"\b(\d{13})\b", ln)
        if m:
            ean_idx[i] = m.group(1)
    # 2) znajdź indeksy Lp
    idx_lp = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d+", ln)]
    products = []
    for lp_idx in idx_lp:
        lp = int(all_lines[lp_idx])
        # znajdź najbliższy wcześniejszy EAN
        prev = [i for i in ean_idx if i < lp_idx]
        if not prev:
            continue
        ean = ean_idx[max(prev)]
        # znajdź qty w okolicy Lp
        qty = None
        for j in range(lp_idx+1, min(lp_idx+8, len(all_lines))):
            # '120 szt.' w jednej linii
            m = re.search(r"(\d{1,4})\s*szt\.?", all_lines[j], flags=re.IGNORECASE)
            if m:
                qty = int(m.group(1))
                break
            # lub oddzielnie: numer, potem 'szt.'
            if all_lines[j].lower().strip() == "szt." and j-1 > lp_idx:
                prev_ln = all_lines[j-1].strip()
                if re.fullmatch(r"\d+", prev_ln):
                    qty = int(prev_ln)
                    break
        if qty is not None:
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)

# ───────────── MAIN ──────────────────────────

uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded_file:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# próba PyPDF2
lines_py = extract_text_with_pypdf2(pdf_bytes)
pattern_d = re.compile(r"^\d{13}.*\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
pattern_e = re.compile(r"^\d+\s+.*\d{1,3}\s+szt\.?", flags=re.IGNORECASE)
has_kod = any(ln.lower().startswith("kod kres") for ln in lines_py)

df = pd.DataFrame()
if lines_py:
    if any(pattern_d.match(ln) for ln in lines_py):
        df = parse_layout_d(lines_py)
    elif any(pattern_e.match(ln) for ln in lines_py) and has_kod:
        df = parse_layout_e(lines_py)
    else:
        # B / C / A
        if any(re.match(r"^\d+\s+\d{13}", ln) for ln in lines_py):
            df = parse_layout_b(lines_py)
        elif any(re.fullmatch(r"\d{13}", ln) for ln in lines_py):
            df = parse_layout_c(lines_py)
        else:
            df = parse_layout_e(lines_py)

# fallback pdfplumber
if df.empty:
    lines_pl = extract_text_with_pdfplumber(pdf_bytes)
    if not lines_pl:
        st.error("Nie udało się wyciągnąć tekstu z PDF-a. Wykonaj OCR i spróbuj ponownie.")
        st.stop()
    if any(pattern_d.match(ln) for ln in lines_pl):
        df = parse_layout_d(lines_pl)
    elif any(pattern_e.match(ln) for ln in lines_pl) and any(ln.lower().startswith("kod kres") for ln in lines_pl):
        df = parse_layout_e(lines_pl)
    else:
        if any(re.match(r"^\d+\s+\d{13}", ln) for ln in lines_pl):
            df = parse_layout_b(lines_pl)
        elif any(re.fullmatch(r"\d{13}", ln) for ln in lines_pl):
            df = parse_layout_c(lines_pl)
        else:
            df = parse_layout_e(lines_pl)

df = df.dropna(subset=["Ilość"]).reset_index(drop=True)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamówienie")
    return out.getvalue()

excel_data = convert_df_to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=excel_data,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
