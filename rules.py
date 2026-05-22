"""
Allergen-utledning og tekstformatering for Menu Digitalizer.

Allergener bygger på EUs 14 obligatoriske allergener (EU 1169/2011).
Utledning skjer KUN fra ingredienstekst i beskrivelsen. Resultatet er
alltid en ANTAKELSE som selgeren maa bekrefte mot vendoren.
"""

import re

# ---------------------------------------------------------------------------
# EUs 14 allergener. Hver allergen har et sett norske nokkelord som,
# hvis de finnes i beskrivelsen, indikerer at allergenet kan vaere til stede.
# Listene er bevisst konservative - de fanger vanlige ingredienser, ikke
# alle tenkelige tilfeller. Selgeren bekrefter alltid.
# ---------------------------------------------------------------------------

ALLERGEN_KEYWORDS = {
    "Gluten": [
        "hvete", "mel", "brod", "bun", "pasta", "nudler", "noodle", "spaghetti",
        "soyasaus", "soya saus", "panko", "tempura", "fritert", "panert",
        "byggryn", "bygg", "rug", "havre", "couscous", "bulgur", "wrap",
        "tortilla", "pita", "baguette", "kjeks", "vareruller", "varrull",
        "spring roll", "dumpling", "wonton", "bao",
    ],
    "Skalldyr": [
        "scampi", "reke", "reker", "krabbe", "hummer", "kreps", "languster",
        "shrimp", "prawn", "crab", "lobster",
    ],
    "Egg": [
        "egg", "majones", "mayo", "aioli", "bearnaise", "hollandaise",
        "eggnudler", "pasta", "kake", "pannekake", "omelett",
    ],
    "Fisk": [
        "fisk", "laks", "torsk", "tunfisk", "tuna", "ansjos", "sardin",
        "fish", "salmon", "fiskesaus", "fish sauce", "nam pla",
    ],
    "Peanotter": [
        "peanott", "peanot", "peanut", "satay", "sate", "jordnott",
    ],
    "Soya": [
        "soya", "soyasaus", "soy sauce", "tofu", "edamame", "miso",
        "teriyaki", "soyabonner",
    ],
    "Melk": [
        "melk", "flote", "ost", "smor", "kremost", "yoghurt", "rommе",
        "romme", "creme fraiche", "parmesan", "mozzarella", "cheddar",
        "iskrem", "is ", "butter", "cream", "cheese", "ghee", "paneer",
    ],
    "Notter": [
        "mandel", "cashew", "valnott", "hasselnott", "pistasj", "pekan",
        "paranott", "macadamia", "almond", "walnut", "hazelnut", "cashew",
    ],
    "Selleri": [
        "selleri", "sellerirot", "celery",
    ],
    "Sennep": [
        "sennep", "mustard", "dijon",
    ],
    "Sesam": [
        "sesam", "sesame", "tahini", "sesamolje", "sesamfro",
    ],
    "Sulfitter": [
        "vin", "eddik", "torket frukt", "rosin", "wine", "balsamico",
    ],
    "Lupin": [
        "lupin",
    ],
    "Blotdyr": [
        "blekksprut", "akkar", "calamari", "musling", "blaaskjell",
        "blaskjell", "ostron", "osters", "snegle", "oystersaus",
        "oyster sauce", "squid", "octopus", "mussel", "clam", "scallop",
    ],
}

# Beskrivelser kortere enn dette regnes som for tynne til aa utlede fra.
MIN_DESC_LENGTH_FOR_ALLERGENS = 12


def _normalize(text):
    """Senk til lowercase og fjern aksenter for robust nokkelord-matching."""
    if not text:
        return ""
    t = text.lower()
    repl = {"\u00e6": "ae", "\u00f8": "o", "\u00e5": "a",
            "\u00e9": "e", "\u00e8": "e", "\u00ea": "e"}
    for k, v in repl.items():
        t = t.replace(k, v)
    return t


def detect_allergens(description):
    """
    Returnerer en streng med utledede allergener fra en beskrivelse.

    - Tom/for kort beskrivelse  -> "Sjekk med vendor"
    - Ingen treff               -> "Ingen funnet (bekreft)"
    - Ett eller flere treff     -> "Egg, Peanotter (antatt - bekreft)"
    """
    if not description or len(description.strip()) < MIN_DESC_LENGTH_FOR_ALLERGENS:
        return "Sjekk med vendor"

    norm = _normalize(description)
    found = []
    for allergen, keywords in ALLERGEN_KEYWORDS.items():
        for kw in keywords:
            # ordgrense slik at "is " ikke matcher "vis", "egg" ikke "legge"
            if re.search(r"\b" + re.escape(kw.strip()) + r"", norm):
                found.append(allergen)
                break

    if not found:
        return "Ingen funnet (bekreft)"
    return ", ".join(found) + " (antatt - bekreft)"


# ---------------------------------------------------------------------------
# Tekstformatering. Reglene kommer direkte fra MDS-testrapporten:
#   - Titler skal vaere Title Case
#   - Beskrivelser skal vaere sentence case (ikke title case)
# ---------------------------------------------------------------------------

# Smaa ord som ikke skal stor forbokstav i titler (med mindre forst/sist).
_TITLE_LOWERCASE = {
    "og", "i", "med", "til", "for", "av", "pa", "uten", "eller",
    "and", "or", "with", "the", "of", "in",
}


def to_title_case(text):
    """Title Case for produkttitler, men behold smaa ord smaa midt i."""
    if not text:
        return text
    words = text.strip().split()
    out = []
    for idx, w in enumerate(words):
        low = w.lower()
        if idx != 0 and idx != len(words) - 1 and low in _TITLE_LOWERCASE:
            out.append(low)
        else:
            # bevar tall/maleenheter som 0,5L
            if any(c.isdigit() for c in w):
                out.append(w)
            else:
                out.append(w[0].upper() + w[1:].lower() if w else w)
    return " ".join(out)


def to_sentence_case(text):
    """
    Sentence case for beskrivelser: stor forbokstav i hver setning,
    resten smaa. Loser MDS-klagen om title case i beskrivelser.
    """
    if not text:
        return text
    text = text.strip().lower()
    # stor forbokstav etter . ! ? og helt i starten
    def _cap(match):
        return match.group(1) + match.group(2).upper()
    text = re.sub(r"(^|[.!?]\s+)([a-zaeoa])", _cap, text)
    # sikre at aller forste tegn er stort selv om regex bommet
    if text:
        text = text[0].upper() + text[1:]
    return text
