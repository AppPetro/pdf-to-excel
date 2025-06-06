import streamlit as st
import pandas as pd
import re
import io
import PyPDF2

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Wyciąga tekst przez PyPDF2.
    2. Na podstawie wykrytego układu wybiera odpowiedni parser:
       - **Układ D**: linie zawierające tylko EAN (13 cyfr) i ilość (`<ilość>,<xx> szt.`).  
       - **Układ E**: linie zaczynające się od Lp i nazwy, potem ilość, a poniżej „Kod kres.: <EAN>”.  
         (Przykłady: pliki typu `Gussto wola park.pdf`, `Zamówienie nr ZD 0175_05_25.pdf`.)  
       - **Układ B**: cała pozycja w jednej linii: `<Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt.`  
       - **Układ C**: czysty 13-cyfrowy EAN w osobnej linii, potem Lp, potem nazwa, „szt.” i ilość.  
       - **Układ A**: „Kod kres.: <EAN>” w osobnej linii, Lp w osobnej linii, fragmenty nazwy przed i po liczbie.
    3. Wywołuje odpowiedni parser i wyświetla wynik w formie tabeli (`Lp`, `Symbol`, `Quantity`, `Kod EAN`).
    4. Umożliwia pobranie danych jako plik Excel.
    """
)


# ──────────────────────────────────────────────────────────────────────────────
# 1) POMOCNICZE FUNKCJE DO WYCIĄGANIA TEKSTU

def extract_text_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    """
    Wyciąga wszystkie niepuste linie tekstu przez PyPDF2.
    Jeśli nic nie znajdzie lub wystąpi błąd, zwraca pustą listę.
    """
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        return []
    lines: list[str] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                lines.append(stripped)
    return lines


# ──────────────────────────────────────────────────────────────────────────────
# 2) PARSERY UKŁADÓW D, E, B, C, A – wszystkie pracują na linii z PyPDF2

def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu D – linie zawierające EAN (13 cyfr) i ilość („<ilość>,<xx> szt.”).
    Przykład:
      5029040012366 Nazwa Produktu 96,00 szt.
      5029040012403 96,00 szt.
    - Lp automatycznie rośnie od 1.
    - Symbol (kolumna) pozostaje pusty, bo nazwa nie zawsze jest w tej samej linii.
    """
    products = []
    pattern = re.compile(
        r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt",
        flags=re.IGNORECASE
    )
    lp_counter = 1
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            barcode_val = m.group(1)
            qty_val = int(m.group(2).replace(" ", ""))
            products.append({
                "Lp": lp_counter,
                "Symbol": "",
                "Quantity": qty_val,
                "Kod EAN": barcode_val
            })
            lp_counter += 1
    return pd.DataFrame(products)


def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu E – linie zaczynające się od Lp i nazwy oraz ilości w tej samej linii,
    a poniżej (ewentualnie po liniach typu "ARA...") znajduje się linia "Kod kres.: <EAN>".
    """
    products = []
    i = 0
    pattern_item = re.compile(r"^(\d+)\s+(.+?)\s+(\d{1,3})\s+szt\.", flags=re.IGNORECASE)

    while i < len(all_lines):
        ln = all_lines[i]
        m = pattern_item.match(ln)
        if m:
            lp_val = int(m.group(1))
            initial_name = m.group(2).strip()
            qty_val = int(m.group(3))
            name_parts = [initial_name]
            barcode_val = None

            j = i + 1
            while j < len(all_lines):
                next_ln = all_lines[j]

                if next_ln.lower().startswith("kod kres"):
                    parts = next_ln.split(":", 1)
                    if len(parts) == 2:
                        barcode_val = parts[1].strip()
                    j += 1
                    break

                if re.fullmatch(r"[A-Za-z0-9]+", next_ln):
                    # linia katalogu (ARA...), pomijamy
                    j += 1
                    continue

                # w przeciwnym razie traktujemy to jako fragment nazwy
                name_parts.append(next_ln.strip())
                j += 1

            full_name = " ".join(name_parts).strip()
            products.append({
                "Lp": lp_val,
                "Symbol": full_name,
                "Quantity": qty_val,
                "Kod EAN": barcode_val
            })

            i = j
        else:
            i += 1

    return pd.DataFrame(products)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu B – cała pozycja w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt. …
    """
    products = []
    pattern = re.compile(
        r"^(\d+)\s+(\d{13})\s+(.+?)\s+(\d{1,3}),\d{2}\s+szt",
        flags=re.IGNORECASE
    )
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            lp_val = int(m.group(1))
            barcode_val = m.group(2)
            name_val = m.group(3).strip()
            qty_val = int(m.group(4).replace(" ", ""))
            products.append({
                "Lp": lp_val,
                "Symbol": name_val,
                "Quantity": qty_val,
                "Kod EAN": barcode_val
            })
    return pd.DataFrame(products)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu C – czysty 13-cyfrowy EAN w osobnej linii, potem Lp, potem nazwa,
    potem "szt." i ilość w kolejnych wierszach.
    """
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt):
                idx_lp.append(i)

    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]
    products = []
    for lp_idx in idx_lp:
        eans_before = [e for e in idx_ean if e < lp_idx]
        barcode_val = all_lines[max(eans_before)] if eans_before else None

        name_val = all_lines[lp_idx + 1] if lp_idx + 1 < len(all_lines) else None

        qty_val = None
        for j in range(lp_idx + 1, len(all_lines) - 2):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"\d+", all_lines[j + 2]):
                qty_val = int(all_lines[j + 2])
                break

        if name_val and qty_val is not None:
            products.append({
                "Lp": int(all_lines[lp_idx]),
                "Symbol": name_val.strip(),
                "Quantity": qty_val,
                "Kod EAN": barcode_val
            })

    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu A – "Kod kres.: <EAN>" w osobnej linii, Lp w osobnej linii,
    fragmenty nazwy przed i po kolumnie cen/ilości.
    """
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                and nxt.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt)
                and not nxt.startswith("Kod kres")
            ):
                idx_lp.append(i)

    idx_ean = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        barcode_val = None
        if valid_eans:
            parts = all_lines[max(valid_eans)].split(":", 1)
            if len(parts) == 2:
                barcode_val = parts[1].strip()

        name_parts: list[str] = []
        qty_val = None
        qty_idx = None

        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                qty_val = int(ln)
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln)
                and not ln.startswith("VAT")
                and ln != "/"
                and not ln.startswith("ARA")
                and not ln.startswith("KAT")
            ):
                name_parts.append(ln)

        if qty_idx is None:
            continue

        for k in range(qty_idx + 1, next_lp):
            ln2 = all_lines[k]
            if ln2.lower().startswith("kod kres"):
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln2)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln2)
                and not ln2.startswith("VAT")
                and ln2 != "/"
                and not ln2.startswith("ARA")
                and not ln2.startswith("KAT")
            ):
                name_parts.append(ln2)

        full_name = " ".join(name_parts).strip()
        products.append({
            "Lp": int(all_lines[lp_idx]),
            "Symbol": full_name,
            "Quantity": qty_val,
            "Kod EAN": barcode_val
        })

    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────
# 3) GŁÓWNA LOGIKA: WCZYTANIE PLIKÓW I WYBÓR PARSERA

# 3.1) Wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# 3.2) Wyciągnięcie tekstu przez PyPDF2
all_lines = extract_text_with_pypdf2(pdf_bytes)

# 3.3) Sprawdź, czy w tekście są układy D/E, aby wiedzieć, czy natychmiast parsować D/E
pattern_d = re.compile(r"^\d{13}(?:\s+.*?)*\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
is_layout_d = any(pattern_d.match(ln) for ln in all_lines)

pattern_e = re.compile(r"^\d+\s+.+?\s+\d{1,3}\s+szt\.", flags=re.IGNORECASE)
has_kod_kres = any(ln.lower().startswith("kod kres") for ln in all_lines)
is_layout_e = any(pattern_e.match(ln) for ln in all_lines) and has_kod_kres

# 3.4) Jeśli nie znaleziono linii albo wykryto D/E → i tak użyjemy PyPDF2 do wszystkich układów
#       (po prostu nie ma dwóch etapów – używamy tylko PyPDF2),
#       ale jeśli wykryto D lub E, wiemy, że chcemy parse_layout_d lub parse_layout_e.
df = pd.DataFrame()
if all_lines:
    if is_layout_d:
        df = parse_layout_d(all_lines)
    elif is_layout_e:
        df = parse_layout_e(all_lines)
    else:
        # Spróbuj układ B
        pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
        is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

        # Spróbuj układ C
        has_pure_ean = any(re.fullmatch(r"\d{13}", ln) for ln in all_lines)
        is_layout_c = has_pure_ean and not is_layout_b

        if is_layout_b:
            df = parse_layout_b(all_lines)
        elif is_layout_c:
            df = parse_layout_c(all_lines)
        else:
            df = parse_layout_a(all_lines)

# 3.5) Usuń wiersze bez wartości „Quantity” (jeśli kolumna istnieje)
if "Quantity" in df.columns:
    df = df.dropna(subset=["Quantity"]).reset_index(drop=True)

# 3.6) Jeśli nic nie znaleziono, wyświetl błąd
if df.empty:
    st.error(
        "Po parsowaniu nie znaleziono pozycji zamówienia. "
        "Upewnij się, że PDF zawiera kody EAN oraz ilości w formacie rozpoznawalnym przez parser."
    )
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# 4) WYŚWIETLENIE WYNIKÓW I EKSPORT DO EXCEL

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)


def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamówienie")
    return output.getvalue()


excel_data = convert_df_to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=excel_data,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet",
)
