"""
Ekstraksjon av menyelementer fra opplastede filer.

Stotter: PDF, Word (.docx), Excel (.xlsx), bilder (.jpg/.jpeg/.png).
Alle formater konverteres til bilde(r) og sendes til Claude sin
vision-modell, som returnerer strukturert JSON.

Hvorfor vision for alt: en meny i PDF/Word kan se ut som ren tekst,
men layouten (kolonner, prislister, varianter) er det vanskelige.
Vision-modellen ser menyen slik en selger ville og kobler
tittel/beskrivelse/pris/variant korrekt i ett steg.

Excel-filer som ALLEREDE er strukturerte leses direkte med pandas -
da trengs ingen modell.
"""

import base64
import io
import json
import os

import anthropic

MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 8000

# Systemprompt som styrer ekstraksjonen. Reglene speiler MDS-skjemaet
# og testrapportens funn (sentence case beskrivelser osv.).
EXTRACTION_PROMPT = """Du er et ekstraksjonsverktoy for restaurantmenyer som skal pa foodora sin plattform.

Du faar et bilde (eller flere) av en meny. Hent ut HVER rett som et eget element.

Returner KUN gyldig JSON - ingen forklaring, ingen markdown-fences. Format:
{
  "items": [
    {
      "title": "Produktnavn slik det staar paa menyen",
      "description": "Beskrivelse / ingredienser hvis oppgitt, ellers tom streng",
      "price": 189.0,
      "variation": "Variantnavn hvis retten har varianter, ellers tom streng",
      "category": "Kategori-overskrift fra menyen hvis synlig, ellers tom streng"
    }
  ]
}

Viktige regler:
- IKKE oversett. Behold originalspraaket fra menyen.
- Pris skal vaere et tall (decimal). Fjern valutasymboler. Finner du ingen pris, sett null.
- Hvis en rett har flere storrelser/varianter med ulik pris, lag ETT element per variant
  og bruk "variation"-feltet (f.eks. "Liten", "Stor", "0,5L"). IKKE putt storrelsen i tittelen.
- Ta med ALLE retter, ogsaa drikke og barnemeny.
- Hvis samme rett staar to ganger paa menyen, ta den med to ganger (selgeren rydder).
- Beskrivelser: skriv av det som staar, ikke dikt opp ingredienser.
"""


def _b64_image(image_bytes, media_type):
    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": media_type,
            "data": base64.standard_b64encode(image_bytes).decode("utf-8"),
        },
    }


def _pdf_to_images(pdf_bytes):
    """Render hver PDF-side til PNG-bytes."""
    import fitz  # PyMuPDF

    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(dpi=180)
        images.append(pix.tobytes("png"))
    doc.close()
    return images


def _docx_to_images(docx_bytes):
    """Konverter Word til PDF via LibreOffice, deretter til bilder."""
    import subprocess
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "input.docx")
        with open(src, "wb") as f:
            f.write(docx_bytes)
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             "--outdir", tmp, src],
            check=True, capture_output=True, timeout=120,
        )
        pdf_path = os.path.join(tmp, "input.pdf")
        with open(pdf_path, "rb") as f:
            return _pdf_to_images(f.read())


def extract_structured_excel(xlsx_bytes):
    """
    Excel som allerede er strukturert leses direkte - ingen modell.
    Returnerer None hvis filen ikke ser ut som en ferdig menytabell.
    """
    import pandas as pd

    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    cols = {c.lower().strip(): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None

    title_col = pick("title", "tittel", "navn", "name", "title_en_no", "produkt")
    price_col = pick("price", "pris", "bel\u00f8p")
    desc_col = pick("description", "beskrivelse", "description_en_no")

    if not title_col or not price_col:
        return None  # ikke en ferdig tabell - send til vision i stedet

    items = []
    for _, row in df.iterrows():
        title = row.get(title_col)
        if pd.isna(title) or not str(title).strip():
            continue
        price = row.get(price_col)
        items.append({
            "title": str(title).strip(),
            "description": "" if not desc_col or pd.isna(row.get(desc_col))
                           else str(row.get(desc_col)).strip(),
            "price": None if pd.isna(price) else float(price),
            "variation": "",
            "category": "",
        })
    return items


def _images_for_file(file_bytes, filename):
    """Konverter en hvilken som helst stottet fil til en liste med (bytes, media_type)."""
    ext = filename.lower().rsplit(".", 1)[-1]

    if ext == "pdf":
        return [(img, "image/png") for img in _pdf_to_images(file_bytes)]
    if ext in ("docx", "doc"):
        return [(img, "image/png") for img in _docx_to_images(file_bytes)]
    if ext in ("jpg", "jpeg"):
        return [(file_bytes, "image/jpeg")]
    if ext == "png":
        return [(file_bytes, "image/png")]
    if ext in ("xlsx", "xls"):
        # ustrukturert excel: konverter til pdf og behandle som bilde
        import subprocess
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, f"input.{ext}")
            with open(src, "wb") as f:
                f.write(file_bytes)
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf",
                 "--outdir", tmp, src],
                check=True, capture_output=True, timeout=120,
            )
            with open(os.path.join(tmp, "input.pdf"), "rb") as f:
                return [(img, "image/png") for img in _pdf_to_images(f.read())]
    raise ValueError(f"Ustottet filtype: .{ext}")


def extract_menu(file_bytes, filename, api_key=None):
    """
    Hovedfunksjon. Tar filinnhold + filnavn, returnerer liste med dicts:
    {title, description, price, variation, category}.

    Kaster ValueError ved ustottet filtype eller manglende API-nokkel.
    """
    ext = filename.lower().rsplit(".", 1)[-1]

    # Snarvei: strukturert Excel trenger ingen modell.
    if ext in ("xlsx", "xls"):
        structured = extract_structured_excel(file_bytes)
        if structured is not None:
            return structured

    images = _images_for_file(file_bytes, filename)
    if not images:
        return []

    return extract_menu_from_images(images, api_key)


def extract_menu_from_files(files, api_key=None):
    """
    Tar flere opplastede filer og analyserer dem som EN meny.

    `files` er en liste med (file_bytes, filename). Alle bildene slaas
    sammen og sendes i ETT modellkall, slik at modellen ser hele menyen
    under ett - viktig naar en fysisk meny er fotografert i flere bilder,
    saa en rett splittet over to bilder ikke telles dobbelt.

    Strukturert Excel handteres fortsatt direkte (uten modell); lastes
    flere strukturerte Excel-filer opp, slaas radene sammen.
    """
    if not files:
        return []

    # Ett enkelt strukturert Excel-ark -> direkte lesing, ingen modell.
    structured_rows = []
    image_files = []
    for file_bytes, filename in files:
        ext = filename.lower().rsplit(".", 1)[-1]
        if ext in ("xlsx", "xls"):
            rows = extract_structured_excel(file_bytes)
            if rows is not None:
                structured_rows.extend(rows)
                continue
        image_files.append((file_bytes, filename))

    # Samle alle bilder fra alle bilde-/PDF-/Word-filer.
    all_images = []
    for file_bytes, filename in image_files:
        all_images.extend(_images_for_file(file_bytes, filename))

    items = []
    if all_images:
        items = extract_menu_from_images(all_images, api_key)

    return structured_rows + items


def extract_menu_from_images(images, api_key=None):
    """
    Kjorer vision-ekstraksjon paa en liste med (bytes, media_type)-bilder.

    Brukes baade av extract_menu (filopplasting) og av URL-modulen, slik
    at det finnes EN ekstraksjonslogikk uansett hvor bildene kommer fra.
    """
    key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        raise ValueError(
            "Mangler ANTHROPIC_API_KEY. Legg den i Streamlit Secrets "
            "eller som milj\u00f8variabel."
        )

    if not images:
        return []

    client = anthropic.Anthropic(api_key=key)

    content = []
    for img_bytes, media_type in images[:20]:  # tak paa 20 bilder
        content.append(_b64_image(img_bytes, media_type))
    content.append({
        "type": "text",
        "text": "Hent ut alle menyelementer fra bildet/bildene over som JSON.",
    })

    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=EXTRACTION_PROMPT,
        messages=[{"role": "user", "content": content}],
    )

    raw = "".join(b.text for b in response.content if b.type == "text").strip()
    # fjern eventuelle markdown-fences modellen kan ha lagt paa
    if raw.startswith("```"):
        raw = raw.split("```", 2)[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(f"Klarte ikke tolke modellsvaret som JSON: {e}")

    return data.get("items", [])
