"""
Menu Digitalizer - foodora

Selgeren laster opp en meny (PDF, Word, Excel eller bilde).
Verktoyet ekstraherer rettene, normaliserer tekst, utleder allergener,
og viser alt i et redigerbart grid. Selgeren retter ved behov og laster
ned en Excel-fil i MDS-formatet, navngitt <Vendor>_<GRID>.xlsx.
"""

import io
import re

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from extraction import extract_menu_from_files
from rules import to_title_case, to_sentence_case

# --- MDS-kolonner. Speiler malen fra MDS sitt CSV-output. -------------------
# Menneskevennlig eksport-format - agenten leser dette manuelt inn i MDS.
# Rekkefolge og navn er valgt for lesbarhet, ikke MDS-kolonnesamsvar.
EXPORT_COLUMNS = [
    "Kategori",
    "Tittel",
    "Variant",
    "Beskrivelse",
    "Pris (NOK)",
    "Allergener",
    "Status",
    "Tilleggsvalg",
    "Growth+",
]

# Statusverdier for allergen-bekreftelse per rad.
STATUS_UBEKREFTET = "Ubekreftet"
STATUS_BEKREFTET = "Bekreftet"
STATUS_IKKE_AKTUELT = "Ikke aktuelt"
STATUS_OPTIONS = [STATUS_UBEKREFTET, STATUS_BEKREFTET, STATUS_IKKE_AKTUELT]

# --- Pandora designsystem (Digital Sales, Commercial) -----------------------
# Fargepalett basert paa teamets dashboard-spec.
PANDORA_PINK = "#FF2B85"
PANDORA_PURPLE = "#882D80"
PANDORA_AMBER = "#F7A21B"
PANDORA_GREEN = "#10B981"
PANDORA_RED = "#EF4444"

# Marked -> sprakkode. Kun NO aktivt naa; lett aa utvide senere.
MARKETS = {
    "Norge": "NO",
}

st.set_page_config(page_title="Menu Digitalizer - Pandora",
                   page_icon="\U0001F37D", layout="wide")

# --- Stil: Pandora-designet, lyst dashboard ---------------------------------
st.markdown("""
<style>
  :root {
      --pink: #FF2B85;
      --purple: #882D80;
      --amber: #F7A21B;
      --green: #10B981;
      --red: #EF4444;
      --grad: linear-gradient(135deg, #FF2B85 0%, #882D80 100%);
      --bg: #f4f2f5;
      --surface: #ffffff;
      --ink: #2d3748;
      --muted: #718096;
      --faint: #a0aec0;
      --border: #e2e8f0;
      --line: #edf2f7;
      --pink-50: #fff1f7;
      --pink-100: #ffe0ee;
      --pink-200: #ffc2da;
      --purple-50: #f7eef6;
      --radius: 16px;
      --shadow: 0 1px 2px rgba(45,55,72,.05), 0 6px 20px rgba(136,45,128,.06);
  }

  /* Grunnlag - lys lilla-graa sidebakgrunn */
  .stApp { background-color: var(--bg); color: var(--ink); }
  .stApp, .stApp p, .stApp label, .stApp span, .stApp li,
  .stApp div[data-testid="stMarkdownContainer"],
  .stApp div[data-testid="stMarkdownContainer"] * {
      color: var(--ink);
  }
  .stApp h1, .stApp h2, .stApp h3,
  section[data-testid="stSidebar"] h1,
  section[data-testid="stSidebar"] h2,
  section[data-testid="stSidebar"] h3 {
      color: var(--ink) !important;
      font-weight: 700;
      letter-spacing: -0.01em;
  }
  .stApp small, .stApp [data-testid="stCaptionContainer"] {
      color: var(--muted) !important;
  }

  /* Sidebar - hvit overflate */
  section[data-testid="stSidebar"] {
      background-color: var(--surface);
      border-right: 1px solid var(--border);
  }
  section[data-testid="stSidebar"] * { color: var(--ink); }

  /* Primaerknapper - pink */
  .stButton > button[kind="primary"],
  .stDownloadButton > button,
  .stFormSubmitButton > button {
      background: var(--pink);
      color: #FFFFFF !important;
      border: none;
      border-radius: 10px;
      font-weight: 600;
      box-shadow: 0 1px 2px rgba(255,43,133,.2);
      transition: transform .05s ease, box-shadow .15s ease;
  }
  .stButton > button[kind="primary"]:hover,
  .stDownloadButton > button:hover,
  .stFormSubmitButton > button:hover {
      background: var(--purple);
      color: #FFFFFF !important;
      box-shadow: 0 2px 6px rgba(136,45,128,.25);
  }
  /* Sekundaerknapper */
  .stButton > button:not([kind="primary"]) {
      background: var(--surface);
      color: var(--ink);
      border: 1px solid var(--border);
      border-radius: 10px;
  }
  .stButton > button:not([kind="primary"]):hover {
      border-color: var(--pink);
      color: var(--pink);
  }

  /* Input, selectbox, radio, multiselect */
  .stTextInput input, .stNumberInput input, .stTextArea textarea {
      background: var(--surface) !important;
      color: var(--ink) !important;
      border: 1px solid var(--border) !important;
      border-radius: 10px !important;
  }
  .stTextInput input:focus, .stNumberInput input:focus,
  .stTextArea textarea:focus {
      border-color: var(--pink) !important;
      box-shadow: 0 0 0 3px var(--pink-50) !important;
  }
  div[data-baseweb="select"] > div,
  div[data-baseweb="input"] > div,
  div[data-baseweb="textarea"] > div {
      background: var(--surface) !important;
      color: var(--ink) !important;
      border-color: var(--border) !important;
      border-radius: 10px !important;
  }
  div[data-baseweb="select"] * { color: var(--ink) !important; }
  div[data-baseweb="popover"] li,
  div[data-baseweb="popover"] div {
      background: var(--surface) !important;
      color: var(--ink) !important;
  }
  div[data-baseweb="popover"] li:hover {
      background: var(--pink-50) !important;
  }
  .stRadio label, .stRadio span { color: var(--ink) !important; }

  /* Filopplaster */
  section[data-testid="stFileUploaderDropzone"] {
      background: var(--purple-50) !important;
      border: 1px dashed var(--purple) !important;
      border-radius: var(--radius) !important;
  }
  section[data-testid="stFileUploaderDropzone"] * {
      color: var(--ink) !important;
  }

  /* Metrikk-kort - hvite kort med lilla skygge */
  div[data-testid="stMetric"] {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 16px 20px;
      box-shadow: var(--shadow);
  }
  div[data-testid="stMetric"] [data-testid="stMetricValue"] {
      color: var(--pink) !important;
      font-weight: 700;
      font-size: 28px;
  }
  div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
      color: var(--muted) !important;
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.04em;
  }

  /* Tabeller */
  .stDataFrame, [data-testid="stDataFrame"] {
      background: var(--surface) !important;
      border-radius: var(--radius);
      border: 1px solid var(--border);
      box-shadow: var(--shadow);
      overflow: hidden;
  }

  /* Alert-bokser - subtile overflater */
  div[data-testid="stAlertContainer"] {
      border-radius: 12px !important;
      border: 1px solid var(--border) !important;
  }

  /* Skille-linjer */
  hr {
      border-color: var(--line) !important;
      margin: 20px 0 !important;
  }

  /* Containere med ramme (st.container(border=True)) */
  div[data-testid="stVerticalBlockBorderWrapper"] {
      border-radius: var(--radius) !important;
      border: 1px solid var(--border) !important;
      box-shadow: var(--shadow);
      background: var(--surface);
  }

  /* Expander - saa den matcher kortstilen */
  details[data-testid="stExpander"] {
      border-radius: var(--radius) !important;
      border: 1px solid var(--border) !important;
      background: var(--surface) !important;
      box-shadow: var(--shadow);
  }

  /* Topp-banner: Pandora-gradient med dekorative sirkler */
  .md-banner {
      position: relative;
      background: var(--grad);
      color: #FFFFFF;
      padding: 26px 30px;
      border-radius: var(--radius);
      margin-bottom: 20px;
      display: flex;
      align-items: center;
      gap: 22px;
      overflow: hidden;
      box-shadow: 0 6px 24px rgba(136,45,128,.18);
  }
  .md-banner::before {
      content: "";
      position: absolute;
      top: -40px; right: -30px;
      width: 180px; height: 180px;
      border-radius: 50%;
      background: rgba(255,255,255,.06);
  }
  .md-banner::after {
      content: "";
      position: absolute;
      bottom: -60px; right: 80px;
      width: 120px; height: 120px;
      border-radius: 50%;
      background: rgba(255,255,255,.06);
  }
  .md-banner h1 { color: #FFFFFF !important; margin: 0; font-size: 30px;
                  font-weight: 700; letter-spacing: -0.02em; position: relative; }
  .md-banner p  { color: rgba(255,255,255,.88) !important; margin: 6px 0 0 0;
                  font-size: 15px; position: relative; }
  .md-logo {
      background: rgba(255,255,255,.18);
      border-radius: 14px;
      width: 46px; height: 46px;
      flex-shrink: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      position: relative;
      backdrop-filter: blur(4px);
  }
  .md-logo img { height: 28px; width: auto; display: block; }

  /* Status-badges - tekst paa tint */
  .badge { display: inline-block; padding: 3px 10px; border-radius: 999px;
           font-size: 12px; font-weight: 600; }
  .badge-success { background: #D1FAE5; color: #047857; }
  .badge-warning { background: #FEF3C7; color: #B45309; }
  .badge-danger  { background: #FEE2E2; color: #B91C1C; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Hjelpefunksjoner
# ---------------------------------------------------------------------------

def items_to_dataframe(items):
    """Gjor raa ekstraksjon om til redigerbart grid med normalisering."""

    def _strip_leading_number(s):
        """Fjern ledende listenummer: '95. Seafood' -> 'Seafood'.
        Lar tall som er del av navnet staa: '7UP', 'Pizza 4 Stagioni'."""
        s = str(s or "").strip()
        # [23] Sushi -> Sushi
        m = re.match(r"^\[\s*\d{1,3}\s*\]\s*(.+)$", s)
        if m:
            return m.group(1).strip()
        # 95. Seafood / 12) Pizza / 7 - Pad Thai / 95.Seafood
        m = re.match(r"^\d{1,3}\s*[.):\-]\s*(.+)$", s)
        if m:
            return m.group(1).strip()
        return s

    rows = []
    for it in items:
        raw_title = _strip_leading_number(str(it.get("title", "")).strip())
        title = to_title_case(raw_title)
        desc = to_sentence_case(str(it.get("description", "")).strip())
        price = it.get("price")

        # To separate allergen-felt:
        # - menu_allergens: kun det restauranten selv har oppgitt
        # - suggested_allergens: modellens typiske forslag (vises som
        #   forslag under gridet, ikke i Allergener-kolonnen)
        menu_al = str(it.get("menu_allergens", "")).strip()
        sug_al = str(it.get("suggested_allergens", "")).strip()

        # Filtrer bort "Ingen allergener"-tekst hvis modellen skulle
        # skrive den - vi vil ha ekte tomme felt slik at forslagspanelet
        # utloses og selgeren tar en aktiv beslutning.
        if menu_al.lower() in ("ingen allergener", "ingen", "none",
                                "n/a", "-"):
            menu_al = ""
        if sug_al.lower() in ("ingen allergener", "ingen", "none",
                               "n/a", "-"):
            sug_al = ""

        add_ons = str(it.get("add_ons", "")).strip()

        # Hvis modellen fant ingen allergener OG heller ingen forslag,
        # sett "Ingen allergener" med en gang - da slipper selgeren aa
        # skrive det inn manuelt.
        allergener_display = menu_al
        if not menu_al and not sug_al:
            allergener_display = "Ingen allergener"

        rows.append({
            "Kategori": str(it.get("category", "")).strip(),
            "Tittel": title,
            "Variant": str(it.get("variation", "")).strip(),
            "Beskrivelse": desc,
            "Pris (NOK)": price if price is not None else 0.0,
            "Allergener": allergener_display,
            "Status": STATUS_UBEKREFTET,
            "Tilleggsvalg": add_ons,
            "Growth+": "",
            "_Forslag": sug_al,      # skjult kolonne - brukes i forslagspanel
        })
    # Skjulte kolonner starter med underscore og filtreres bort ved eksport.
    return pd.DataFrame(rows, columns=EXPORT_COLUMNS + ["_Forslag"])


def safe_filename_part(text):
    """Gjor en tekst trygg som del av et filnavn."""
    cleaned = re.sub(r"[^\w\-]+", "_", text.strip())
    return cleaned.strip("_")


def build_export_filename(vendor, grid):
    """Bygg filnavn etter MDS-konvensjonen <Vendor>_<GRID>.xlsx."""
    v = safe_filename_part(vendor) or "Vendor"
    g = safe_filename_part(grid) or "GRID"
    return f"{v}_{g}.xlsx"


def build_export_excel(df):
    """
    Bygg Excel-fila i den menneskevennlige eksportstrukturen.

    Fila leses manuelt av en agent som legger menyen inn i MDS,
    saa strukturen er valgt for lesbarhet, ikke MDS-kolonnesamsvar.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Meny"

    ws.append(EXPORT_COLUMNS)
    header_fill = PatternFill("solid", start_color=PANDORA_PINK.lstrip("#"))
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", name="Arial")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    price_col_idx = EXPORT_COLUMNS.index("Pris (NOK)") + 1
    status_col_idx = EXPORT_COLUMNS.index("Status") + 1
    red_fill = PatternFill("solid", start_color="FFD6D6")
    green_fill = PatternFill("solid", start_color="D4EDDA")
    yellow_fill = PatternFill("solid", start_color="FFF2CC")
    gray_fill = PatternFill("solid", start_color="E9ECEF")

    for _, r in df.iterrows():
        ws.append([r.get(c, "") for c in EXPORT_COLUMNS])
        row_num = ws.max_row

        # Mangler pris -> rod.
        if not r.get("Pris (NOK)") or r.get("Pris (NOK)") == 0:
            ws.cell(row=row_num, column=price_col_idx).fill = red_fill

        # Status-fargekoding.
        status = str(r.get("Status", "")).strip()
        if status == STATUS_BEKREFTET:
            ws.cell(row=row_num, column=status_col_idx).fill = green_fill
        elif status == STATUS_UBEKREFTET:
            ws.cell(row=row_num, column=status_col_idx).fill = yellow_fill
        elif status == STATUS_IKKE_AKTUELT:
            ws.cell(row=row_num, column=status_col_idx).fill = gray_fill

    # Kolonnebredder tilpasset innhold.
    widths = {
        "Kategori": 18, "Tittel": 26, "Variant": 14,
        "Beskrivelse": 55, "Pris (NOK)": 12, "Allergener": 30,
        "Status": 14, "Tilleggsvalg": 45, "Growth+": 10,
    }
    for col_idx, name in enumerate(EXPORT_COLUMNS, start=1):
        letter = ws.cell(row=1, column=col_idx).column_letter
        ws.column_dimensions[letter].width = widths.get(name, 18)

    # Tekstbryting paa lange kolonner slik at agenten ser hele teksten.
    wrap = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = wrap

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def _logo_data_uri():
    """Les Delivery Hero-logoen som base64 slik at den kan ligge i banneret."""
    import base64
    import os
    path = os.path.join(os.path.dirname(__file__), "dh_logo.png")
    try:
        with open(path, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode("utf-8")
        return f"data:image/png;base64,{b64}"
    except FileNotFoundError:
        return None


_logo = _logo_data_uri()
_logo_html = (
    f'<div class="md-logo"><img src="{_logo}" alt="Delivery Hero"/></div>'
    if _logo else ""
)

st.markdown(f"""
<div class="md-banner">
  {_logo_html}
  <div>
    <h1>Menu Digitalizer</h1>
    <p>Last opp en meny &rarr; verkt&oslash;yet strukturerer den &rarr;
       rediger og last ned klar for agent.</p>
  </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Innstillinger")

    market_name = st.selectbox(
        "Marked",
        options=list(MARKETS.keys()),
        help="Bestemmer hvilken Title-/Description-kolonne i MDS-malen "
             "som fylles ut. Flere markeder kan legges til senere.",
    )
    market = MARKETS[market_name]

    st.divider()
    st.subheader("Vendor")
    vendor_name = st.text_input(
        "Vendornavn",
        placeholder="f.eks. Randis Gatekjokken",
        help="Brukes i filnavnet paa den nedlastede Excel-fila.",
    )
    grid_id = st.text_input(
        "GRID-id",
        placeholder="f.eks. a4b2",
        help="Vendor-id paa DH. Brukes i filnavnet etter "
             "MDS-konvensjonen <Vendor>_<GRID>.xlsx.",
    )

    st.divider()
    st.subheader("Prisjustering")
    adjust_mode = st.radio(
        "Type",
        options=["Ingen", "Prosent (%)", "Kroner (NOK)"],
        horizontal=False,
        help="Velg om alle priser skal \u00f8kes med en prosent eller "
             "et fast kronebel\u00f8p.",
    )
    price_pct = 0.0
    price_kr = 0.0
    if adjust_mode == "Prosent (%)":
        price_pct = st.number_input(
            "\u00d8k alle priser med (%)",
            min_value=0.0, max_value=100.0, value=0.0, step=1.0,
        )
    elif adjust_mode == "Kroner (NOK)":
        price_kr = st.number_input(
            "\u00d8k alle priser med (NOK)",
            min_value=0.0, max_value=500.0, value=0.0, step=1.0,
        )

# API-nokkelen leses fra Streamlit Secrets - aldri vist i UI.
try:
    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
except Exception:
    api_key = ""

if "menu_df" not in st.session_state:
    st.session_state.menu_df = None

uploaded = st.file_uploader(
    "Last opp meny",
    type=["pdf", "docx", "xlsx", "xls", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
    help="PDF, Word, Excel eller bilde av menyen. Du kan laste opp "
         "flere bilder av samme meny \u2013 de analyseres samlet.",
)
analyze_file = st.button("Analyser meny", type="primary",
                         disabled=not uploaded)

st.caption(
    "\U0001F4A1 Meny p\u00e5 nett? \u00c5pne siden i nettleseren, scroll "
    "helt til bunnen sl\u00e5 hele menyen er lastet, og lagre siden som "
    "PDF (Cmd+P \u2192 Lagre som PDF). Last s\u00e5 opp PDF-en her."
)

# --- Filopplasting -----------------------------------------------------------
if analyze_file and uploaded:
    with st.spinner(f"Analyserer {len(uploaded)} fil(er) \u2026"):
        try:
            files = [(f.getvalue(), f.name) for f in uploaded]
            items = extract_menu_from_files(files, api_key or None)
            if not items:
                st.warning("Fant ingen menyelementer i filen(e).")
            else:
                st.session_state.menu_df = items_to_dataframe(items)
                st.success(f"Hentet ut {len(items)} element(er) fra "
                           f"{len(uploaded)} fil(er). Rediger ved behov "
                           "under.")
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Noe gikk galt under analysen: {e}")

# ---------------------------------------------------------------------------
# Redigerbart grid
# ---------------------------------------------------------------------------

if st.session_state.menu_df is not None:
    st.subheader("Rediger menyen")
    st.caption("Klikk i cellene for \u00e5 rette. Endringer lagres f\u00f8rst "
               "n\u00e5r du trykker **Lagre endringer** \u2013 da slipper du "
               "at siden oppdaterer seg etter hver lille rettelse.")

    with st.form("editor_form", clear_on_submit=False):
        edited_form = st.data_editor(
            st.session_state.menu_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Kategori": st.column_config.TextColumn(width="small"),
                "Tittel": st.column_config.TextColumn(width="medium"),
                "Variant": st.column_config.TextColumn(width="small"),
                "Beskrivelse": st.column_config.TextColumn(width="large"),
                "Pris (NOK)": st.column_config.NumberColumn(
                    format="%.0f", min_value=0),
                "Allergener": st.column_config.TextColumn(width="medium"),
                "Status": st.column_config.SelectboxColumn(
                    width="small",
                    options=STATUS_OPTIONS,
                    required=True,
                    help="Bekreft allergener mot vendor for hver rett.",
                ),
                "Tilleggsvalg": st.column_config.TextColumn(width="large"),
                "Growth+": st.column_config.TextColumn(width="small"),
                "_Forslag": None,  # skjul: brukes bare i forslagspanel
            },
            key="editor",
        )
        saved = st.form_submit_button("\U0001F4BE Lagre endringer",
                                      type="primary")
        if saved:
            st.session_state.menu_df = edited_form

    # Det vi viser videre er den sist lagrede versjonen.
    edited = st.session_state.menu_df

    # --- Forslagspanel: allergener modellen mener er typiske ----------------
    def _merge_allergens(current, added):
        """Slaa sammen kommaseparerte allergenlister uten duplikater."""
        cur = [a.strip() for a in (current or "").split(",") if a.strip()]
        add = [a.strip() for a in (added or "").split(",") if a.strip()]
        seen = {a.lower(): a for a in cur}
        for a in add:
            if a.lower() not in seen:
                seen[a.lower()] = a
                cur.append(a)
        return ", ".join(cur)

    open_suggestions = edited[
        edited["_Forslag"].fillna("").str.strip() != ""
    ]

    if len(open_suggestions) > 0:
        st.subheader(
            f"\U0001F4A1 Forslag til ekstra allergener "
            f"({len(open_suggestions)})"
        )
        st.caption(
            "Modellen mener disse allergenene er typiske for retten, "
            "men de sto ikke paa menyen. Vurder hver enkelt: **Godta** "
            "legger dem til i Allergener-kolonnen, **Avvis** fjerner "
            "forslaget. Redigering gjor du direkte i Allergener-cellen "
            "i tabellen over."
        )

        for idx, row in open_suggestions.iterrows():
            title = str(row["Tittel"]).strip()
            variant = str(row["Variant"]).strip()
            label = f"{title} \u2013 {variant}" if variant else title
            current = str(row["Allergener"]).strip() or "*(ingen)*"
            suggested = str(row["_Forslag"]).strip()

            with st.container(border=True):
                cc1, cc2, cc3, cc4 = st.columns([3, 3, 1, 1])
                with cc1:
                    st.markdown(f"**{label}**")
                    st.caption(f"Fra meny: {current}")
                with cc2:
                    st.markdown(f"Forslag: **{suggested}**")
                with cc3:
                    if st.button("\u2705 Godta", key=f"acc_{idx}",
                                 use_container_width=True):
                        new_df = st.session_state.menu_df.copy()
                        new_df.at[idx, "Allergener"] = _merge_allergens(
                            new_df.at[idx, "Allergener"], suggested)
                        new_df.at[idx, "_Forslag"] = ""
                        st.session_state.menu_df = new_df
                        st.rerun()
                with cc4:
                    if st.button("\u274c Avvis", key=f"rej_{idx}",
                                 use_container_width=True):
                        new_df = st.session_state.menu_df.copy()
                        new_df.at[idx, "_Forslag"] = ""
                        # Hvis Allergener-cella ble tom etter avvisning,
                        # sett "Ingen allergener" saa selgeren slipper
                        # aa skrive det inn manuelt.
                        current = str(new_df.at[idx, "Allergener"]).strip()
                        if not current:
                            new_df.at[idx, "Allergener"] = "Ingen allergener"
                        st.session_state.menu_df = new_df
                        st.rerun()
    else:
        st.success("\u2705 Ingen ubehandlede allergen-forslag.")

    # --- Sok og erstatt ------------------------------------------------------
    with st.expander("\U0001F50D S\u00f8k og erstatt"):
        st.caption("Erstatt tekst p\u00e5 tvers av flere rader p\u00e5 en gang. "
                   "Case-sensitiv. Trykk **Erstatt** for \u00e5 gj\u00f8re "
                   "endringen.")
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            sr_find = st.text_input("S\u00f8k etter", key="sr_find")
        with c2:
            sr_repl = st.text_input("Erstatt med", key="sr_repl")
        with c3:
            sr_cols = st.multiselect(
                "I kolonner",
                options=["Tittel", "Beskrivelse", "Variant", "Kategori",
                         "Allergener", "Tilleggsvalg"],
                default=["Tittel", "Beskrivelse"],
                key="sr_cols",
            )
        # Forhaandsvis antall treff for tilbakemelding.
        if sr_find and sr_cols:
            hits = 0
            for c in sr_cols:
                hits += edited[c].fillna("").astype(str).str.count(
                    re.escape(sr_find)).sum()
            st.caption(f"Fant **{int(hits)}** forekomst(er) av "
                       f"«{sr_find}» i valgte kolonner.")
        if st.button("Erstatt", key="sr_do",
                     disabled=not (sr_find and sr_cols)):
            new_df = edited.copy()
            for c in sr_cols:
                new_df[c] = new_df[c].fillna("").astype(str).str.replace(
                    sr_find, sr_repl, regex=False)
            st.session_state.menu_df = new_df
            st.success(f"Erstattet «{sr_find}» med «{sr_repl}» i "
                       f"{len(sr_cols)} kolonne(r).")
            st.rerun()

    missing_mask = edited["Pris (NOK)"] == 0
    missing_price = int(missing_mask.sum())
    total_rows = len(edited)
    confirmed = int((edited["Status"] == STATUS_BEKREFTET).sum())
    unconfirmed = int((edited["Status"] == STATUS_UBEKREFTET).sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Retter", total_rows)
    c2.metric("Mangler pris", missing_price)
    c3.metric("Allergener bekreftet", f"{confirmed}/{total_rows}")

    if missing_price:
        manglende = edited.loc[missing_mask, "Tittel"].tolist()
        liste = ", ".join(str(t) for t in manglende if str(t).strip())
        st.error(f"\u26a0\ufe0f {missing_price} rett(er) mangler pris og er "
                 f"merket r\u00f8dt i eksportfila: {liste}. "
                 "Fyll inn riktig pris f\u00f8r du laster opp menyen.")
    if unconfirmed:
        st.warning(
            f"\U0001F514 {unconfirmed} rett(er) har fortsatt status "
            "**Ubekreftet**. Gaa gjennom allergenene med vendoren og "
            "sett status til *Bekreftet* pr rad f\u00f8r du laster ned."
        )

    # --- Forhaandsvisning av prisjustering -----------------------------------
    def _apply_adjustment(prices):
        """Returner justert prisserie basert paa valgt modus."""
        if price_pct > 0:
            return (prices * (1.0 + price_pct / 100.0)).round(0)
        if price_kr > 0:
            # Bare priser stoerre enn 0 justeres - vi vil ikke gi en
            # manglende-pris-rad et tilfeldig tall.
            return prices.where(prices == 0, prices + price_kr).round(0)
        return prices

    adjustment_active = price_pct > 0 or price_kr > 0
    if adjustment_active:
        adj_label = (f"+{price_pct:.0f}%" if price_pct > 0
                     else f"+{price_kr:.0f} kr")
        preview = edited[["Tittel", "Variant", "Pris (NOK)"]].copy()
        preview = preview.rename(columns={"Pris (NOK)": "Pris f\u00f8r"})
        preview[f"Pris etter {adj_label}"] = _apply_adjustment(
            edited["Pris (NOK)"])
        st.markdown(f"**Prisjustering: {adj_label}** \u2013 "
                    "slik blir prisene i eksportfila:")
        st.dataframe(preview, use_container_width=True, hide_index=True)

    st.divider()

    # --- Growth+ -------------------------------------------------------------
    st.subheader("Growth+")
    sold_growth = st.radio(
        "Er Growth+ solgt til denne vendoren?",
        options=["Nei", "Ja"],
        horizontal=True,
    )

    growth_indices = []
    if sold_growth == "Ja":
        # Bygg valg: "Tittel - Variant" saa identiske titler skilles.
        def _label(idx, row):
            t = str(row["Tittel"]).strip() or "(uten tittel)"
            v = str(row["Variant"]).strip()
            return f"{t} - {v}" if v else t

        option_map = {
            _label(idx, row): idx for idx, row in edited.iterrows()
        }
        picked = st.multiselect(
            "Velg de 2 rettene som inng\u00e5r i Growth+-kampanjen",
            options=list(option_map.keys()),
            max_selections=2,
        )
        growth_indices = [option_map[p] for p in picked]
        if len(picked) != 2:
            st.info(f"Velg n\u00f8yaktig 2 retter \u2013 valgt n\u00e5: "
                    f"{len(picked)}.")
        else:
            st.success("2 retter valgt \u2013 merkes med TRUE i "
                       "Growth+-kolonnen.")

    st.divider()

    # Anvend prisjustering og Growth+-merking paa eksport-dataframe.
    export_df = edited.copy()
    if adjustment_active:
        export_df["Pris (NOK)"] = _apply_adjustment(export_df["Pris (NOK)"])
    if growth_indices:
        export_df.loc[growth_indices, "Growth+"] = "TRUE"

    # Fallback: tomme Allergener-celler fylles med "Ingen allergener"
    # slik at eksportfila aldri har en tom allergen-verdi.
    empty_mask = export_df["Allergener"].fillna("").str.strip() == ""
    export_df.loc[empty_mask, "Allergener"] = "Ingen allergener"

    export_name = build_export_filename(vendor_name, grid_id)
    if not vendor_name or not grid_id:
        st.info(f"Fyll inn vendornavn og GRID-id i sidepanelet for "
                f"riktig filnavn. N\u00e5v\u00e6rende: **{export_name}**")
    else:
        st.caption(f"Fila lastes ned som: **{export_name}**")

    excel_buf = build_export_excel(export_df)
    st.download_button(
        "\u2b07\ufe0f Last ned Excel",
        data=excel_buf,
        file_name=export_name,
        mime="application/vnd.openxmlformats-officedocument."
             "spreadsheetml.sheet",
        type="primary",
    )
