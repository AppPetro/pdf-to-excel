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
    3. Łączymy kontynuacje nazw produktów w jedną linię.  # ✂️ zmiana  
    4. Wstawiamy brakującą spację między numerem a nazwą.  
    5. Wykrywamy format “WZ/Subiekt GT” i parsujemy EAN:
       - Patrzymy na linie zaczynające się od Lp + tekst + ‘szt.’  
       - Z każdego takiego wyciągamy Ilość i **EAN z końca linii**.  # ✂️ zmiana  
    6. Albo – jeśli to faktura – D, E, B, C lub A.  
    7. Pokazujemy tabelę, statystyki i eksport do Excela.
    """
)

def extract_text(pdf_bytes: bytes) -> list[str]:
    try:
        lines: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for ln in (page.extract_text() or "").split("\n"):
                    if (stripped := ln.strip()):
                        lines.append(stripped)
        return lines
    except Exception:
        return []

def parse_layout_wz(all_lines: list[str]) -> pd.DataFrame:
    """
    Parsuje tylko wiersze produktów:
     - zaczynają się od numeru pozycji i zawierają 'szt.',
     - wyciąga Lp, Ilość oraz EAN jako 13 cyfr na końcu linii.
    """
    products = []
    for ln in all_lines:
        # tylko linie z Lp i szt.
        if not re.match(r"^\d+\s+[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln) or "szt." not in ln:
            continue

        # numer pozycji
        lp = int(re.match(r"^(\d+)", ln).group(1))

        # ilość
        m_qty = re.search(r"([\d,]+)\s+szt\.", ln)
        qty = int(float(m_qty.group(1).replace(",", "."))) if m_qty else None

        # EAN na końcu linii: 13 cyfr + opcjonalna kropka
        m_ean = re.search(r"(\d{13})\.?$", ln)
        ean = m_ean.group(1) if m_ean else ""

        products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})

    return pd.DataFrame(products)

# — pozostałe parsery fakturowe D, E, B, C, A — (bez zmian!) —
def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    products=[]; lp=1
    pat=re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
    for ln in all_lines:
        if m:=pat.match(ln):
            products.append({"Lp":lp,"Symbol":m.group(1),"Ilość":int(m.group(2).replace(" ",""))})
            lp+=1
    return pd.DataFrame(products)

def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    products=[]; i=0
    pat_item=re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", re.IGNORECASE)
    while i<len(all_lines):
        if m:=pat_item.match(all_lines[i]):
            lp,q=int(m.group(1)),int(m.group(2))
            ean=""; j=i+1
            while j<len(all_lines) and not all_lines[j].lower().startswith("kod kres"): j+=1
            if j<len(all_lines):
                parts=all_lines[j].split(":",1)
                if len(parts)==2: ean=parts[1].strip()
            products.append({"Lp":lp,"Symbol":ean,"Ilość":q})
            i=j+1
        else: i+=1
    return pd.DataFrame(products)

def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    products=[]; 
    pat=re.compile(r"^(\d+)\s+(\d{13})\s+.+?\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
    for ln in all_lines:
        if m:=pat.match(ln):
            products.append({"Lp":int(m.group(1)),"Symbol":m.group(2),"Ilość":int(m.group(3).replace(" ",""))})
    return pd.DataFrame(products)

def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    idx_lp=[i for i in range(len(all_lines)-1)
            if re.fullmatch(r"\d+",all_lines[i])
               and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]",all_lines[i+1])
               and not all_lines[i+1].lower().startswith("kod kres")]
    idx_ean=[i for i,ln in enumerate(all_lines) if re.fullmatch(r"\d{13}",ln)]
    products=[]
    for lp_idx in idx_lp:
        prev_lp=max((e for e in idx_lp if e<lp_idx),default=-1)
        next_lp=min((e for e in idx_lp if e>lp_idx),default=len(all_lines))
        valid=[e for e in idx_ean if prev_lp<e<next_lp]
        ean=all_lines[max(valid)] if valid else ""
        qty=None
        for j in range(lp_idx+1,next_lp):
            if re.fullmatch(r"\d+",all_lines[j]) and j+1<next_lp and all_lines[j+1].lower()=="szt.":
                qty=int(all_lines[j]);break
        if qty is not None:
            products.append({"Lp":int(all_lines[lp_idx]),"Symbol":ean,"Ilość":qty})
    return pd.DataFrame(products)

def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    idx_lp=[i for i in range(len(all_lines)-1)
            if re.fullmatch(r"\d+",all_lines[i])
               and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]",all_lines[i+1])
               and not all_lines[i+1].lower().startswith("kod kres")]
    idx_kod=[i for i,ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products=[]
    for k,lp_idx in enumerate(idx_lp):
        prev_lp=idx_lp[k-1] if k>0 else -1
        next_lp=idx_lp[k+1] if k+1<len(idx_lp) else len(all_lines)
        valid=[e for e in idx_kod if prev_lp<e<next_lp]
        ean=""
        if valid:
            parts=all_lines[max(valid)].split(":",1)
            if len(parts)==2: ean=parts[1].strip()
        qty=None
        for j in range(lp_idx+1,next_lp):
            if re.fullmatch(r"\d+",all_lines[j]) and j+1<next_lp and all_lines[j+1].lower()=="szt.":
                qty=int(all_lines[j]);break
        if qty is not None:
            products.append({"Lp":int(all_lines[lp_idx]),"Symbol":ean,"Ilość":qty})
    return pd.DataFrame(products)

# ────────────────────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader("Wybierz plik PDF", type=["pdf"])
if not uploaded_file:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# 1) pobieramy linie
lines = extract_text(pdf_bytes)

# 2) usuwamy stopki/numerację
lines = [ln for ln in lines if not ln.startswith("/") and "Strona" not in ln]

# 3) ✂️ zmiana: łączymy linie kontynuacji w opisie
merged = []
for ln in lines:
    if re.match(r"^\d+\s+", ln):
        merged.append(ln)
    else:
        if merged:
            merged[-1] += " " + ln
        else:
            merged.append(ln)
lines = merged

# 4) wstawiamy spację między numerem a nazwą
lines = [re.sub(r"^(\d+)(?=[A-Za-z])", r"\1 ", ln) for ln in lines]

# 5) detekcja WZ/Subiekt GT
is_wz = any(re.search(r"[\d,]+\s+szt\..*?\d{13}\.?", ln) for ln in lines) \
      or any(re.match(r"^\d+\s+\d{13}\s+.+?\s+[\d,]+\s+szt\.", ln) for ln in lines)

# pozostałe detekcje faktury
is_d     = any(re.match(r"^(\d{13})", ln) for ln in lines)
has_kres = any(ln.lower().startswith("kod kres") for ln in lines)
is_e     = any(re.match(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", ln) for ln in lines) and has_kres
is_b     = any(re.match(r"^\d+\s+\d{13}", ln) for ln in lines)
has_plain= any(re.fullmatch(r"\d{13}", ln) for ln in lines)
is_c     = has_plain and not is_b

# 6) wybieramy parser
if is_wz:
    df = parse_layout_wz(lines)
elif is_d:
    df = parse_layout_d(lines)
elif is_e:
    df = parse_layout_e(lines)
elif is_b:
    df = parse_layout_b(lines)
elif is_c:
    df = parse_layout_c(lines)
else:
    df = parse_layout_a(lines)

# 7) filtrujemy puste ilości
if "Ilość" in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji.")
    st.stop()

# 8) statystyki
total   = df.shape[0]
unique  = df["Symbol"].nunique()
sum_qty = int(df["Ilość"].sum())

# 9) komunikaty o błędach/prawidłowości
missing = df["Symbol"].eq("").sum()
if missing:
    st.error(f"Brakuje EAN w {missing} pozycjach!")
elif total != unique:
    st.error(
        f"Znaleziono w sumie: {total} pozycji  \n"
        f"Unikalnych EAN-ów: {unique}"
    )
else:
    st.markdown(
        f"**Znaleziono w sumie:** {total} pozycji  \n"
        f"**Unikalnych EAN-ów:** {unique}  \n"
        f"**Łączna ilość:** {sum_qty}"
    )

# 10) wyświetlenie i eksport
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
