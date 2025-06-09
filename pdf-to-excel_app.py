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
    1. Próbuje wyciągnąć tekst przez PyPDF2 (stare PDF-y).
    2. Jeśli PyPDF2 nie zwróci pełnych pozycji lub nie znajdzie EAN, używa pdfplumber.
    3. Obsługuje różne layouty (D, E, B, C, A) z kolejnością prób:
       - **Układ D**: linia zaczyna się od 13-cyfrowego EAN-u i zawiera ",<qty>,xx szt."
       - **Układ E**: linia z Lp i ilością „szt.”, poniżej linia z EAN (dowolnie w treści)
       - **Układ B**: jedna linia: Lp, EAN, nazwa, ilość z przecinkami
       - **Układ C**: EAN w osobnej linii (dowolnie w treści), obok Lp i ilość w kolejnych wierszach
       - **Układ A**: prefiks “Kod kres.:", Lp, nazwa fragmenty i ilość w oddzielnych liniach
    4. Wyświetla wynik w tabeli **Lp | Symbol | Ilość** i pozwala pobrać Excel.
    """
)

# ───────────── Ekstrakcja tekstu ─────────────
def extract_py2(pdf_bytes: bytes) -> list[str]:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        lines = []
        for page in reader.pages:
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                ln = ln.strip()
                if ln:
                    lines.append(ln)
        return lines
    except:
        return []


def extract_plumber(pdf_bytes: bytes) -> list[str]:
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
    except:
        return []

# ───────────── Parsery layoutów ─────────────
def parse_d(lines: list[str]) -> pd.DataFrame:
    patt = re.compile(r"^(\d{13}).*?(\d{1,3}),\d{2}\s*szt", flags=re.IGNORECASE)
    out, lp = [], 1
    for ln in lines:
        m = patt.match(ln)
        if m:
            out.append({"Lp": lp, "Symbol": m.group(1), "Ilość": int(m.group(2))})
            lp += 1
    return pd.DataFrame(out)


def parse_e(lines: list[str]) -> pd.DataFrame:
    out = []
    patt = re.compile(r"^(\d+)\s+.+?(\d{1,3})\s+szt\.", flags=re.IGNORECASE)
    i = 0
    while i < len(lines):
        m = patt.match(lines[i])
        if m:
            lp = int(m.group(1))
            qty = int(m.group(2))
            ean = None
            for j in range(i+1, min(i+4, len(lines))):
                found = re.search(r"(\d{13})", lines[j])
                if found:
                    ean = found.group(1)
                    break
            out.append({"Lp": lp, "Symbol": ean or "", "Ilość": qty})
        i += 1
    return pd.DataFrame(out)


def parse_b(lines: list[str]) -> pd.DataFrame:
    out = []
    patt = re.compile(r"^(\d+)\s+(\d{13})\s+.+?(\d{1,3}),\d{2}\s*szt", flags=re.IGNORECASE)
    for ln in lines:
        m = patt.match(ln)
        if m:
            out.append({"Lp": int(m.group(1)), "Symbol": m.group(2), "Ilość": int(m.group(3))})
    return pd.DataFrame(out)


def parse_c(lines: list[str]) -> pd.DataFrame:
    # zbierz mapę linii->EAN (dowolne 13-cyfrowe)
    ean_map = {i: m.group(1) for i, ln in enumerate(lines)
               if (m := re.search(r"\b(\d{13})\b", ln))}
    out = []
    for i, ln in enumerate(lines):
        if re.fullmatch(r"\d+", ln):
            lp = int(ln)
            # najbliższy wcześniejszy EAN
            prev = [k for k in ean_map if k < i]
            if not prev:
                continue
            symbol = ean_map[max(prev)]
            qty = None
            for j in range(i+1, min(i+4, len(lines))):
                if m2 := re.search(r"(\d{1,4})\s*szt", lines[j], flags=re.IGNORECASE):
                    qty = int(m2.group(1))
                    break
            if qty is not None:
                out.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})
    return pd.DataFrame(out)


def parse_a(lines: list[str]) -> pd.DataFrame:
    out = []
    ean_idx = {i: re.search(r"(\d{13})", ln).group(1)
               for i, ln in enumerate(lines) if ln.lower().startswith("kod kres") and re.search(r"\d{13}", ln)}
    for i, ln in enumerate(lines):
        if re.fullmatch(r"\d+", ln):
            lp = int(ln)
            prev = [k for k in ean_idx if k < i]
            if not prev:
                continue
            symbol = ean_idx[max(prev)]
            qty = None
            if i+2 < len(lines) and lines[i+2].lower().startswith("szt"):
                if re.fullmatch(r"\d+", lines[i+1]):
                    qty = int(lines[i+1])
            if qty:
                out.append({"Lp": lp, "Symbol": symbol, "Ilość": qty})
    return pd.DataFrame(out)

# ───────────── Main ─────────────
uploaded = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded.read()
lines = extract_py2(pdf_bytes)
if not lines:
    lines = extract_plumber(pdf_bytes)

# Wypróbuj kolejno wszystkie parsery w kolejności D, E, B, C, A
parsers = [parse_d, parse_e, parse_b, parse_c, parse_a]
df = pd.DataFrame()
for parser in parsers:
    tmp = parser(lines)
    if "Ilość" in tmp.columns and not tmp.empty:
        df = tmp
        break

# Jeśli dalej pusto, komunikat błędu
df = df.reset_index(drop=True)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

# Wyświetl i eksportuj
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
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

    label="Pobierz wynik jako Excel",
    data=to_excel(df),
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
