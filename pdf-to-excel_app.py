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
    4. Wykrywamy format “WZ/Subiekt GT” na podstawie wzorca linii i parsujemy go.  
    5. Albo – jeśli to faktura – D, E, B, C lub A.  
    6. Pokazujemy tabelę, statystyki EAN-ów i umożliwiamy eksport do Excela.
    """
)

def extract_text(pdf_bytes: bytes) -> list[str]:
    try:
        lines: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for ln in (page.extract_text() or "").split("\n"):
                    if stripped := ln.strip():
                        lines.append(stripped)
        return lines
    except Exception:
        return []

def parse_layout_wz(all_lines: list[str]) -> pd.DataFrame:
    """
    Parsuje linię:
    Lp  Nazwa ...  Ilość szt  EAN(13)  Masa
    bez nagłówków – sam pattern.
    """
    products = []
    wz_pat = re.compile(
        r"^(\d+)\s+"         # grupa 1: Lp
        r"(.+?)\s+"          # grupa 2: Nazwa (najkrócej jak się da)
        r"([\d,]+)\s+szt\s+" # grupa 3: Ilość
        r"(\d{13})\s+"       # grupa 4: EAN
        r"([\d,]+)"          # grupa 5: Masa
    )
    for ln in all_lines:
        if m := wz_pat.match(ln):
            lp   = int(m.group(1))
            name = m.group(2)
            qty  = int(float(m.group(3).replace(",", ".")))
            ean  = m.group(4)
            # masę ignorujemy w danych wyjściowych lub można dodać jako kolumnę
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)

# — dotychczasowe parsery fakturowe —
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
            products.append({
                "Lp":int(m.group(1)),
                "Symbol":m.group(2),
                "Ilość":int(m.group(3).replace(" ",""))
            })
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

# ───────────── GŁÓWNA LOGIKA ────────────────────────────────────────────

uploaded_file = st.file_uploader("Wybierz plik PDF", type=["pdf"])
if not uploaded_file:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# 1) wszystkie linie przez pdfplumber
lines = extract_text(pdf_bytes)

# 2) filtrujemy stopki i numerację stron
lines = [ln for ln in lines if not ln.startswith("/") and "Strona" not in ln]

# 3) wstawiamy spację między numerem a nazwą
lines = [re.sub(r"^(\d+)(?=[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż])", r"\1 ", ln)
         for ln in lines]

# 4) detekcja WZ/Subiekt GT po wzorcu linii
wz_pat = re.compile(
    r"^\d+\s+.+?\s+[\d,]+\s+szt\s+\d{13}\s+[\d,]+"
)
is_wz = any(wz_pat.match(ln) for ln in lines)  # :contentReference[oaicite:0]{index=0}

# pozostałe detekcje faktur D/E/B/C/A
pat_d = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", re.IGNORECASE)
pat_e = re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", re.IGNORECASE)
is_d = any(pat_d.match(ln) for ln in lines)
has_kres = any(ln.lower().startswith("kod kres") for ln in lines)
is_e = any(pat_e.match(ln) for ln in lines) and has_kres
is_b = any(re.compile(r"^\d+\s+\d{13}", re.IGNORECASE).match(ln) for ln in lines)
has_plain = any(re.fullmatch(r"\d{13}", ln) for ln in lines)
is_c = has_plain and not is_b

# 5) wybór parsera
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

# 6) filtrowanie pustych ilości
if "Ilość" in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji.")
    st.stop()

# 7) statystyki EAN-ów
total_eans  = df.shape[0]
unique_eans = df["Symbol"].nunique()
total_qty   = int(df["Ilość"].sum())
st.markdown(
    f"**Znaleziono w sumie:** {total_eans} pozycji  \n"
    f"**Unikalnych EAN-ów:** {unique_eans}  \n"
    f"**Sumaryczna ilość:** {total_qty}"
)

# 8) tabela i eksport do Excela
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
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
