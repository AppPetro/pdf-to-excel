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
    2. Jeśli nie wykryje układów D/E, używa parserów B/C/A.
    3. W przeciwnym razie (lub gdy PyPDF2 nie da linii) 
       wyciąga tekst przez pdfplumber i próbuje wykryć układy D/E/B/C/A.
    4. Parsuje i wyświetla tabelę **Lp | Symbol | Ilość**.
    5. Sprawdza, czy liczba pozycji (max Lp) zgadza się z liczbą unikalnych EAN-ów.
       Jeśli nie – pokazuje ostrzeżenie.
    6. Umożliwia pobranie wyników jako Excel.
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
    products, lp = [], 1
    pat = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
    for ln in all_lines:
        m = pat.match(ln)
        if m:
            products.append({"Lp": lp, "Symbol": m.group(1), "Ilość": int(m.group(2))})
            lp += 1
    return pd.DataFrame(products)


def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ E – każdemu Lp przypisujemy pierwszy następujący po nim EAN,
    ale usuwamy go z dalszej puli, żeby nie pojawił się ponownie.
    """
    # 1) Zbierz pozycje (index, Lp, Ilość)
    pat_item = re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", re.IGNORECASE)
    items = []
    for i, ln in enumerate(all_lines):
        m = pat_item.match(ln)
        if m:
            items.append({"idx": i, "Lp": int(m.group(1)), "Ilość": int(m.group(2))})

    # 2) Zbierz wszystkie EANy (index, ean)
    pat_ean = re.compile(r"^kod kres\.\s*:\s*(\d{13})", re.IGNORECASE)
    eans = []
    for i, ln in enumerate(all_lines):
        m = pat_ean.match(ln)
        if m:
            eans.append({"idx": i, "ean": m.group(1)})

    # 3) Przypisuj po kolei, usuwając wykorzystane eany
    products = []
    remaining = eans.copy()
    for it in items:
        cand = [e for e in remaining if e["idx"] > it["idx"]]
        if cand:
            chosen = min(cand, key=lambda e: e["idx"])
            symbol = chosen["ean"]
            remaining.remove(chosen)
        else:
            symbol = ""
        products.append({
            "Lp": it["Lp"],
            "Symbol": symbol,
            "Ilość": it["Ilość"]
        })

    return pd.DataFrame(products)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pat = re.compile(r"^(\d+)\s+(\d{13})\s+.+?\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
    for ln in all_lines:
        m = pat.match(ln)
        if m:
            products.append({"Lp": int(m.group(1)), "Symbol": m.group(2), "Ilość": int(m.group(3))})
    return pd.DataFrame(products)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    idx_lp = [
        i for i in range(len(all_lines)-1)
        if re.fullmatch(r"\d+", all_lines[i])
        and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", all_lines[i+1])
    ]
    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]
    products = []
    for lp_i in idx_lp:
        before = [e for e in idx_ean if e < lp_i]
        ean = all_lines[max(before)] if before else ""
        qty = None
        for j in range(lp_i+1, len(all_lines)-2):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"\d+", all_lines[j+2]):
                qty = int(all_lines[j+2])
                break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_i]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    idx_lp = []
    for i in range(len(all_lines)-1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i+1]
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                and nxt.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt)
                and not nxt.lower().startswith("kod kres")
            ):
                idx_lp.append(i)
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products = []
    for k, lp_i in enumerate(idx_lp):
        prev_i = idx_lp[k-1] if k > 0 else -1
        next_i = idx_lp[k+1] if k+1 < len(idx_lp) else len(all_lines)
        val = [e for e in idx_ean if prev_i < e < next_i]
        ean = all_lines[max(val)].split(":",1)[1].strip() if val else ""
        qty = None
        for j in range(lp_i+1, next_i):
            if re.fullmatch(r"\d+", all_lines[j]) and all_lines[j+1].lower() == "szt.":
                qty = int(all_lines[j])
                break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_i]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────

uploaded = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()
pdf_bytes = uploaded.read()

# próba PyPDF2
lines_py = extract_text_with_pypdf2(pdf_bytes)
pat_d = re.compile(r"^\d{13}.*\d{1,3},\d{2}\s+szt", re.IGNORECASE)
pat_e = re.compile(r"^\d+\s+.+?\s+\d{1,3}\s+szt\.", re.IGNORECASE)
has_kod_py = any(ln.lower().startswith("kod kres") for ln in lines_py)
is_d_py = any(pat_d.match(ln) for ln in lines_py)
is_e_py = any(pat_e.match(ln) for ln in lines_py) and has_kod_py

df = pd.DataFrame()
if lines_py and not (is_d_py or is_e_py):
    is_b_py = any(re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", re.IGNORECASE).match(ln) for ln in lines_py)
    is_c_py = any(re.fullmatch(r"\d{13}", ln) for ln in lines_py) and not is_b_py
    if is_b_py:
        df = parse_layout_b(lines_py)
    elif is_c_py:
        df = parse_layout_c(lines_py)
    else:
        df = parse_layout_a(lines_py)

# jeśli nic, pdfplumber
if df.empty:
    lines_new = extract_text_with_pdfplumber(pdf_bytes)
    if not lines_new:
        st.error("Nie udało się wyciągnąć tekstu z tego PDF-a."); st.stop()
    is_d = any(pat_d.match(ln) for ln in lines_new)
    is_e = any(pat_e.match(ln) for ln in lines_new) and any(ln.lower().startswith("kod kres") for ln in lines_new)
    is_b = any(re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", re.IGNORECASE).match(ln) for ln in lines_new)
    is_c = any(re.fullmatch(r"\d{13}", ln) for ln in lines_new) and not is_b

    if is_d:
        df = parse_layout_d(lines_new)
    elif is_e:
        df = parse_layout_e(lines_new)
    elif is_b:
        df = parse_layout_b(lines_new)
    elif is_c:
        df = parse_layout_c(lines_new)
    else:
        df = parse_layout_a(lines_new)

# porządkowanie
if "Ilość" in df:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)

if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji."); st.stop()

# WALIDACJA: ile pozycji vs ile unikalnych EAN
max_lp = int(df["Lp"].max())
unique_ean = df["Symbol"].nunique()
if max_lp != unique_ean:
    st.warning(
        f"Uwaga! Znalazłem {max_lp} pozycji, ale tylko {unique_ean} unikalnych kodów EAN – sprawdź, czy parsowanie się nie pogubiło."
    )

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_in.to_excel(w, index=False, sheet_name="Zamówienie")
    return buf.getvalue()

st.download_button(
    label="Pobierz jako Excel",
    data=convert_df_to_excel(df),
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
