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
    1. Próbuje wyciągnąć tekst przez PyPDF2 (stare PDF-y) lub pdfplumber.
    2. Parsuje w układach D, E, B, C lub A (w tej kolejności).
    3. Wyświetla wynik w tabeli **Lp | Symbol | Ilość** i pozwala pobrać Excel.
    """
)

# ───────────── Ekstrakcja tekstu ─────────────
def extract_text_py2(pdf_bytes: bytes) -> list[str]:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        lines: list[str] = []
        for page in reader.pages:
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                ln = ln.strip()
                if ln:
                    lines.append(ln)
        return lines
    except Exception:
        return []

def extract_text_plumber(pdf_bytes: bytes) -> list[str]:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            lines: list[str] = []
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    ln = ln.strip()
                    if ln:
                        lines.append(ln)
            return lines
    except Exception:
        return []

# ───────────── Parsery layoutów ─────────────
def parse_layout_d(lines: list[str]) -> pd.DataFrame:
    patt = re.compile(r"^(\d{13}).*?(\d{1,3}),\d{2}\s*szt", flags=re.IGNORECASE)
    out, lp = [], 1
    for ln in lines:
        m = patt.match(ln)
        if m:
            out.append({"Lp": lp, "Symbol": m.group(1), "Ilość": int(m.group(2))})
            lp += 1
    return pd.DataFrame(out)

def parse_layout_e(lines: list[str]) -> pd.DataFrame:
    out = []
    patt = re.compile(r"^(\d+)\s+.+?(\d{1,3})\s+szt\.", flags=re.IGNORECASE)
    for i in range(len(lines)):
        m = patt.match(lines[i])
        if m:
            lp, qty = int(m.group(1)), int(m.group(2))
            ean = None
            for j in range(i+1, min(i+4, len(lines))):
                f = re.search(r"(\d{13})", lines[j])
                if f:
                    ean = f.group(1)
                    break
            out.append({"Lp": lp, "Symbol": ean or "", "Ilość": qty})
    return pd.DataFrame(out)

def parse_layout_b(lines: list[str]) -> pd.DataFrame:
    out = []
    patt = re.compile(r"^(\d+)\s+(\d{13})\s+.+?(\d{1,3}),\d{2}\s*szt", flags=re.IGNORECASE)
    for ln in lines:
        m = patt.match(ln)
        if m:
            out.append({"Lp": int(m.group(1)), "Symbol": m.group(2), "Ilość": int(m.group(3))})
    return pd.DataFrame(out)

def parse_layout_c(lines: list[str]) -> pd.DataFrame:
    """
    Układ C – EAN w osobnej linii jako czysty 13-cyfrowy kod lub linia zaczynająca się od 'Kod kres.'.
    Następnie numer pozycji (Lp) i ilość w formacie '... szt.' w kolejnych liniach.
    """
    # 1) Zbuduj mapę indeks->EAN tylko dla czystych linii 13 cyfr lub z prefiksem 'Kod kres.'
    ean_map: dict[int, str] = {}
    for idx, ln in enumerate(lines):
        if re.fullmatch(r"\d{13}", ln):
            ean_map[idx] = ln
        elif ln.lower().startswith("kod kres"):
            parts = ln.split(":", 1)
            if len(parts) == 2 and (m := re.search(r"(\d{13})", parts[1])):
                ean_map[idx] = m.group(1)

    out: list[dict] = []
    # 2) Dla każdej linii z numerem pozycji (Lp) dobierz ostatni wcześniejszy EAN
    for idx, ln in enumerate(lines):
        if re.fullmatch(r"\d+", ln):
            lp = int(ln)
            prev_idxs = [i for i in ean_map if i < idx]
            if not prev_idxs:
                continue
            symbol = ean_map[max(prev_idxs)]
            # 3) poszukaj ilości w kilku kolejnych liniach
            qty = None
            for j in range(idx+1, min(idx+5, len(lines))):
                if m2 := re.search(r"(\d{1,4})\s*szt", lines[j], flags=re.IGNORECASE):
                    qty = int(m2.group(1))
                    break
            if qty is not None:
                out.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})
    return pd.DataFrame(out)

def parse_layout_a(lines: list[str]) -> pd.DataFrame:
    out = []
    ean_idx = {i: re.search(r"(\d{13})", ln).group(1)
               for i, ln in enumerate(lines)
               if ln.lower().startswith("kod kres") and re.search(r"\d{13}", ln)}
    for i, ln in enumerate(lines):
        if re.fullmatch(r"\d+", ln):
            lp = int(ln)
            prev = [k for k in ean_idx if k < i]
            if not prev:
                continue
            symbol = ean_idx[max(prev)]
            qty = None
            if i+2 < len(lines) and lines[i+2].lower().startswith("szt") and re.fullmatch(r"\d+", lines[i+1]):
                qty = int(lines[i+1])
            if qty is not None:
                out.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})
    return pd.DataFrame(out)

# ───────────── Main ─────────────
uploaded = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded.read()
lines = extract_text_py2(pdf_bytes) or extract_text_plumber(pdf_bytes)

# próbuj kolejno układów D, E, B, C, A
df = parse_layout_d(lines)
if df.empty:
    df = parse_layout_e(lines)
if df.empty:
    df = parse_layout_b(lines)
if df.empty:
    df = parse_layout_c(lines)
if df.empty:
    df = parse_layout_a(lines)

# finalizacja
df = df.dropna(subset=["Ilość"]).reset_index(drop=True)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# eksport do Excela
def to_excel(df_in: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_in.to_excel(w, index=False, sheet_name="Zamówienie")
    return buf.getvalue()

st.download_button(
    label="Pobierz wynik jako Excel",
    data=to_excel(df),
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet"
)
