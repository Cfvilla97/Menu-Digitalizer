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
MAX_TOKENS = 16000

# Systemprompt som styrer ekstraksjonen. Reglene speiler MDS-skjemaet
# og testrapportens funn (sentence case beskrivelser osv.).
EXTRACTION_PROMPT = """Du er et ekstraksjonsverktoy for restaurantmenyer som skal pa foodora sin plattform.

Du faar et bilde (eller flere) av en meny. Hent ut HVER rett som et eget element.

Returner KUN gyldig JSON - ingen forklaring, ingen markdown-fences. Format:
{
  "items": [
    {
      "title": "Produktnavn slik det staar paa menyen",
      "description": "Beskrivelse av retten. Avsluttes med punktum.",
      "price": 189.0,
      "variation": "Variantnavn hvis retten har varianter, ellers tom streng",
      "category": "Kategori-overskrift fra menyen hvis synlig, ellers tom streng",
      "allergens": "Gluten, Melk, Egg"
    }
  ]
}

Viktige regler:
- IKKE oversett. Behold originalspraaket fra menyen.
- Pris skal vaere et tall (decimal). Fjern valutasymboler. Finner du ingen pris, sett null.
- Ta med ALLE retter, ogsaa drikke og barnemeny.
- Hvis samme rett staar to ganger paa menyen, ta den med to ganger (selgeren rydder).

VARIANTER - dette er viktig, MDS bommer ofte her:
- Hvis en rett eller drikke tilbys i flere STORRELSER med ulik pris
  (Liten/Stor, 0,33L/0,5L/1,5L), lag ETT element per storrelse. Bruk
  "variation"-feltet. IKKE i tittelen. Dette gjelder ogsaa brus:
  "Coca-Cola 0,5L" og "Coca-Cola 1,5L" blir to elementer med samme
  tittel "Coca-Cola" og variation "0,5L" / "1,5L".
- For DRIKKE skal "variation" KUN vaere storrelsen (f.eks. "0,5L").
  IKKE ta med emballasje. Ord som "flaske", "boks", "PET", "i boks",
  "paa flaske" skal IKKE staa i variation og IKKE bli egne elementer.
  Eksempel: "Appelsinsmakende Fanta, 0,5L flaske eller boks" ->
  ETT element, variation "0,5L".
- Hvis en rett tilbys med ulikt INNHOLD/PROTEIN (f.eks. "ris med kylling
  eller kjott eller scampi"), lag ETT element per valg. Da blir
  "variation" f.eks. "Kylling", "Kjott", "Scampi" - hver med sin pris.
- Kort sagt: hver kombinasjon kunden faktisk kan bestille og betale for
  skal vaere sin egen rad.

BESKRIVELSE - dette feltet skal ALLTID fylles ut, aldri tom:
- Hvis menyen har en beskrivelse: bruk den, men gjor den utfyllende og
  innholdsrik.
- Hvis menyen mangler beskrivelse: skriv en utfyllende, innholdsrik
  beskrivelse basert paa hva retten ER (rettnavnet og kjente kjennetegn).
  Eksempel: "Sprobunnspizza med tomatsaus, mozzarella, skinke og
  champignon." eller "Wokrett med fritert kylling, gronnsaker og
  sotsur saus, servert med ris."
- Beskriv bunntype, hovedingredienser og tilbehor der det er kjent.
- IKKE dikt opp spesifikke ingredienser som retten ikke har.
- Skriv i sentence case (stor forbokstav forst, ikke title case).
- Avslutt ALLTID med punktum.

ALLERGENS - list forventede allergener for HVER rett:
- Bruk de 14 EU-allergenene: Gluten, Skalldyr, Egg, Fisk, Peanotter, Soya,
  Melk, Notter, Selleri, Sennep, Sesam, Sulfitter, Lupin, Blotdyr.
- Ta med bade allergener nevnt i menyteksten OG allergener som er typiske
  for retten (f.eks. Pad Thai -> Peanotter, Egg, Fisk; pizza -> Gluten, Melk).
- Skriv KUN allergennavnene, kommaseparert: "Gluten, Melk, Egg".
- IKKE skriv "antatt", "bekreft", "typisk" eller liknende - bare navnene.
- For RENE produkter uten noen av de 14 allergenene (f.eks. brus, vann,
  juice, sort kaffe, fersk frukt): skriv "Ingen allergener".
- Klarer du virkelig ikke vurdere produktet, la feltet vaere tom streng.
"""


def _compress_image(image_bytes, max_b64_bytes=4_900_000):
    """
    Krymp et bilde slik at den BASE64-KODEDE versjonen holder seg under
    API-grensa paa 5 MB.

    VIKTIG: API-grensa gjelder den base64-kodede strengen, ikke raa-fila.
    Base64 gjor data ca. 33 % storre, saa vi maaler base64-storrelsen
    direkte i stedet for raa bytes.

    Telefonbilder er ofte 6-12 MB. Menytekst er fullt lesbar i lavere
    opplosning, saa nedskalering paavirker ikke ekstraksjonen merkbart.

    Returnerer (nye_bytes, media_type).
    """
    import base64 as _b64
    from io import BytesIO
    from PIL import Image

    def _b64_len(data):
        # storrelsen paa base64-strengen uten aa faktisk kode alt i minnet
        return (len(data) + 2) // 3 * 4

    # Allerede liten nok? Behold som den er.
    if _b64_len(image_bytes) <= max_b64_bytes:
        return image_bytes, None

    img = Image.open(BytesIO(image_bytes))
    if img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")

    # Skaler ned i pikseldimensjon, og senk JPEG-kvaliteten trinnvis.
    # Vi proever flere kombinasjoner til base64-storrelsen er trygg.
    for max_dim in (2200, 1800, 1500, 1200, 1000):
        work = img
        if max(work.size) > max_dim:
            ratio = max_dim / max(work.size)
            work = work.resize(
                (max(1, int(work.size[0] * ratio)),
                 max(1, int(work.size[1] * ratio))),
                Image.LANCZOS,
            )
        for quality in (85, 70, 55, 40):
            buf = BytesIO()
            work.save(buf, format="JPEG", quality=quality)
            data = buf.getvalue()
            if _b64_len(data) <= max_b64_bytes:
                return data, "image/jpeg"

    # Siste utvei: kraftig nedskalering.
    work = img.resize(
        (max(1, img.size[0] // 4), max(1, img.size[1] // 4)),
        Image.LANCZOS,
    )
    buf = BytesIO()
    work.save(buf, format="JPEG", quality=50)
    return buf.getvalue(), "image/jpeg"


def _b64_image(image_bytes, media_type):
    compressed, new_type = _compress_image(image_bytes)
    if new_type:
        media_type = new_type
    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": media_type,
            "data": base64.standard_b64encode(compressed).decode("utf-8"),
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
        return data.get("items", [])
    except json.JSONDecodeError:
        # Svaret kan ha blitt avkuttet (naadd token-taket). Forsok aa
        # redde de komplette elementene i stedet for aa feile helt.
        salvaged = _salvage_truncated_items(raw)
        if salvaged:
            return salvaged
        raise ValueError(
            "Modellsvaret kunne ikke tolkes som JSON. Menyen kan vaere "
            "for stor - prov aa dele den opp i faerre sider per opplasting."
        )


def _salvage_truncated_items(raw):
    """
    Redder komplette menyelementer fra et JSON-svar som ble avkuttet.

    Finner alle hele {...}-objekter inne i "items"-lista og parser dem
    enkeltvis. Et halvt siste objekt forkastes.
    """
    import re

    start = raw.find('"items"')
    if start == -1:
        return []
    bracket = raw.find("[", start)
    if bracket == -1:
        return []

    items = []
    depth = 0
    obj_start = None
    for i in range(bracket + 1, len(raw)):
        ch = raw[i]
        if ch == "{":
            if depth == 0:
                obj_start = i
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and obj_start is not None:
                chunk = raw[obj_start:i + 1]
                try:
                    items.append(json.loads(chunk))
                except json.JSONDecodeError:
                    pass
                obj_start = None
    return items
