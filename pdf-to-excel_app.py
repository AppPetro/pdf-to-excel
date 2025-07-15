import streamlit as st
import pandas as pd
import re
import io
import pdfplumber

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj PDF ze zleceniem wydania / WZ-em lub fakturą:
    1. Pobieramy wszystkie linie przez pdfplumber (łączy strony).  
    2. Usuwamy stopki/numerację stron.  
    3. Wstawiamy brakującą spację między numerem a nazwą.  
    4. Wykrywamy format:
       - Tabela WZ (Kolumny: Lp, Kod produktu, Nazwa, Termin ważności, Ilość, Waga brutto) :contentReference[oaicite:0]{index=0}  
       - WZ/Subiekt GT (z `szt.` i EAN)  
       - Faktura D, E, B, C lub A.  
    5. Pokazujemy tabelę, statystyki (z kontrolą brakujących i duplikatów) i eksport do Excela.
    """
)

def extract_text(pdf_bytes: bytes) -> list[str]:
    try:
        lines: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for ln in (page.extract_text() or "").split("\n"):
                    stripped = ln.strip()
                    if stripped:
                        lines.append(stripped)
        return lines
    except Exception:
        return []

def parse_layout_wz_table(all_lines: list[str]) -> pd.DataFrame:
    """
    Parsuje WZ w formie tabelarycznej:
    Lp, Kod produktu, Nazwa, Termin ważności, Ilość, Waga brutto
    """
    products = []
    # znajdź nagłówek "Kod produktu"
    header_idx = next((i for i, ln in enumerate(all_lines)
                       if ln.startswith("Kod produktu")), None)
    if header_idx is None:
        return pd.DataFrame(products)
    # każda kolejna linia to wiersz danych
    pat = re.compile(
        r"^(\d+)\s+"            # Lp
        r"(\d{13})\s+"          # Kod produktu (EAN)
        r"(.+?)\s+"             # Nazwa
        r"(\d{4}-\d{2}-\d{2})\s+"# Termin ważności
        r"([\d\s,]+)\s+"        # Ilość (z separatorami)
        r"([\d\s,]+)$"          # Waga brutto
    )
    for ln in all_lines[header_idx+1:]:
        if m := pat.match(ln):
            qty = int(float(m.group(5).replace(" ", "").replace(",", ".")))
            products.append({
                "Lp": int(m.group(1)),
                "Symbol": m.group(2),
                "Ilość": qty
            })
    return pd.DataFrame(products)

def parse_layout_wz(all_lines: list[str]) -> pd.DataFrame:
    products = []
    lp = 1

    # 1) '<ilość> szt. <EAN>'
    pat1 = re.compile(r"([\d\s,]+)\s+szt\.\s+(\d{13}\.?)")
    for ln in all_lines:
        if m := pat1.search(ln):
            qty = int(float(m.group(1).replace(" ", "").replace(",", ".")))
            ean = m.group(2).rstrip(".")
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
            lp += 1
    if products:
        return pd.DataFrame(products)

    # 2) kolumnowy: '^Lp  <EAN>  ...  <ilość> szt.'
    pat2 = re.compile(
        r"^(\d+)\s+"           # Lp
        r"(\d{13})\s+"         # Symbol=EAN
        r".+?\s+"              # Nazwa i inne
        r"([\d\s,]+)\s+szt\."  # Ilość (z separatorem tys.)
    )
    for ln in all_lines:
        if m := pat2.match(ln):
            qty2 = int(float(m.group(3).replace(" ", "").replace(",", ".")))
            products.append({
                "Lp": int(m.group(1)),
                "Symbol": m.group(2),
                "Ilość": qty2
            })
    if products:
        return pd.DataFrame(products)

    # 3) EAN na końcu linii (non-greedy, bez spacji przed EAN)
    pat3 = re.compile(r"^(\d+)\s+.+\s+([\d\s,]+)\s+szt\.\s+.*?(\d{13})$")
    for ln in all_lines:
        if m := pat3.match(ln):
            qty3 = int(float(m.group(2).replace(" ", "").replace(",", ".")))
            products.append({
                "Lp": int(m.group(1)),
                "Symbol": m.group(3),
                "Ilość": qty3
            })
    return pd.DataFrame(products)

# — pozostałe parsery fakturowe D, E, B, C, A — (bez zmian) —
def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    products = []; lp = 1
    pat = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
    for ln in all_lines:
        if m := pat.match(ln):
            products.append({"Lp": lp, "Symbol": m.group(1), "Ilość": int(m.group(2).replace(" ", ""))})
            lp += 1
    return pd.DataFrame(products)

def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    products = []; i = 0
    pat_item = re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", re.IGNORECASE)
    while i < len(all_lines):
        if m := pat_item.match(all_lines[i]):
            lp, q = int(m.group(1)), int(m.group(2))
            ean = ""; j = i + 1
            while j < len(all_lines) and not all_lines[j].lower().startswith("kod kres"):
                j += 1
            if j < len(all_lines):
                parts = all_lines[j].split(":", 1)
                if len(parts) == 2:
                    ean = parts[1].strip()
            products.append({"Lp": lp, "Symbol": ean, "Ilość": q})
            i = j + 1
        else:
            i += 1
    return pd.DataFrame(products)

def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pat = re.compile(r"^(\d+)\s+(\d{13})\s+.+?\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
    for ln in all_lines:
        if m := pat.match(ln):
            products.append({
                "Lp": int(m.group(1)),
                "Symbol": m.group(2),
                "Ilość": int(m.group(3).replace(" ", ""))
            })
    return pd.DataFrame(products)

def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    idx_lp = [
        i for i in range(len(all_lines)-1)
        if re.fullmatch(r"\d+", all_lines[i])
           and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", all_lines[i+1])
           and not all_lines[i+1].lower().startswith("kod kres")
    ]
    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]
    products = []
    for lp_idx in idx_lp:
        prev_lp = max((e for e in idx_lp if e < lp_idx), default=-1)
        next_lp = min((e for e in idx_lp if e > lp_idx), default=len(all_lines))
        valid = [e for e in idx_ean if prev_lp < e < next_lp]
        ean = all_lines[max(valid)] if valid else ""
        qty = None
        for j in range(lp_idx+1, next_lp):
            if re.fullmatch(r"\d+", all_lines[j]) and j+1<next_lp and all_lines[j+1].lower()=="szt.":
                qty = int(all_lines[j])
                break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_idx]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)

def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    idx_lp = [
        i for i in range(len(all_lines)-1)
        if re.fullmatch(r"\d+", all_lines[i])
           and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", all_lines[i+1])
           and not all_lines[i+1].lower().startswith("kod kres")
    ]
    idx_kod = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products = []
    for k, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[k-1] if k>0 else -1
        next_lp = idx_lp[k+1] if k+1<len(idx_lp) else len(all_lines)
        valid = [e for e in idx_kod if prev_lp<e<next_lp]
        ean = ""
        if valid:
            parts = all_lines[max(valid)].split(":",1)
            if len(parts)==2:
                ean = parts[1].strip()
        qty = None
        for j in range(lp_idx+1, next_lp):
            if re.fullmatch(r"\d+", all_lines[j]) and j+1<next_lp and all_lines[j+1].lower()=="szt.":
                qty = int(all_lines[j])
                break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_idx]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)

# ────────────────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader("Wybierz plik PDF", type=["pdf"])
if not uploaded_file:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()
lines = extract_text(pdf_bytes)

# 1) usuń stopki/numerację
lines = [ln for ln in lines if not ln.startswith("/") and "Strona" not in ln]

# 2) wstaw spację Lp→Nazwa
lines = [re.sub(r"^(\d+)(?=[A-Za-z])", r"\1 ", ln) for ln in lines]

# 3) detekcje
is_wz_table = any(ln.startswith("Kod produktu") for ln in lines)           # tabela WZ :contentReference[oaicite:1]{index=1}
is_wz       = (
    any(re.search(r"[\d\s,]+\s+szt\.\s+\d{13}\.?", ln) for ln in lines)
    or any(re.match(r"^\d+\s+\d{13}\s+.+?\s+[\d\s,]+\s+szt\.", ln) for ln in lines)
    or any(re.match(r"^\d+\s+.+\s+[\d\s,]+\,\d+\s+szt\.\s+.*\d{13}$", ln) for ln in lines)
)
is_d     = any(re.match(r"^(\d{13})", ln) for ln in lines)
has_kres = any(ln.lower().startswith("kod kres") for ln in lines)
is_e     = any(re.match(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", ln) for ln in lines) and has_kres
is_b     = any(re.match(r"^\d+\s+\d{13}", ln) for ln in lines)
has_plain= any(re.fullmatch(r"\d{13}", ln) for ln in lines)
is_c     = has_plain and not is_b

# 4) wybór parsera
if   is_wz_table: df = parse_layout_wz_table(lines)
elif is_wz:       df = parse_layout_wz(lines)
elif is_d:        df = parse_layout_d(lines)
elif is_e:        df = parse_layout_e(lines)
elif is_b:        df = parse_layout_b(lines)
elif is_c:        df = parse_layout_c(lines)
else:             df = parse_layout_a(lines)

# 5) filtr pustych ilości
if "Ilość" in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji.")
    st.stop()

# 6) statystyki i walidacja spójności
total    = df.shape[0]
unique_e = df["Symbol"].nunique()
sum_qty  = int(df["Ilość"].sum())

if unique_e != total:
    st.error(f"Liczba pozycji ({total}) różni się od liczby unikalnych EAN-ów ({unique_e}).")

st.markdown(
    f"**Znaleziono w sumie:** {total} pozycji  \n"
    f"**Unikalnych EAN-ów:** {unique_e}  \n"
    f"**Łączna ilość:** {sum_qty}"
)

# 7) wynik i eksport
st.subheader("Wyekstrahowane pozycje")
st.dataframe(df, use_container_width=True)

def to_excel(df_in: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df_in.to_excel(w, index=False, sheet_name="Zamówienie")
    return out.getvalue()

st.download_button(
    label="Pobierz jako Excel",
    data=to_excel(df),
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
)
