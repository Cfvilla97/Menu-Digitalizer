# Menu Digitalizer — foodora

Selger laster opp en meny (PDF, Word, Excel eller bilde) → verktøyet
ekstraherer rettene, normaliserer tekst, utleder allergener, og viser
alt i et redigerbart grid. Selger retter ved behov og laster ned en
Excel-fil i MDS-formatet, navngitt `<Vendor>_<GRID>.xlsx`.

## Filer

| Fil | Ansvar |
|-----|--------|
| `app.py` | Streamlit-UI: opplasting, redigerbart grid, eksport |
| `extraction.py` | Filhåndtering + vision-modell-kall → strukturert JSON |
| `rules.py` | Allergenutledning (EUs 14) + tekstnormalisering |
| `requirements.txt` | Python-avhengigheter |
| `packages.txt` | Systempakker (LibreOffice) |
| `dh_logo.png` | Delivery Hero-logo, vises i banneret |

## Slik kjører du lokalt

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# rediger secrets.toml og legg inn din ANTHROPIC_API_KEY
streamlit run app.py
```

## Streamlit Cloud

Push til GitHub, koble repoet i Streamlit Cloud, og legg
`ANTHROPIC_API_KEY` inn under **Settings → Secrets**. `packages.txt`
installerer LibreOffice for Word/Excel-konvertering.

## Hvordan det fungerer

1. **Ekstraksjon** — filer konverteres til bilde(r) og sendes til
   Claude sin vision-modell, som leser menyens layout og returnerer
   strukturert JSON. Strukturert Excel leses direkte med pandas.
2. **Normalisering** — titler → Title Case, beskrivelser → sentence
   case. Løser to konkrete klager fra MDS-testrapporten.
3. **Allergener** — utledes fra ingrediensene i beskrivelsen, matchet
   mot EUs 14 allergener. Alltid merket *antatt – bekreft*.
4. **Eksport** — redigert grid skrives til Excel i MDS-kolonneformatet
   pluss `Allergens`-kolonnen, filnavn `<Vendor>_<GRID>.xlsx`.

## Meny på nett

For en meny som ligger på en nettside (f.eks. en bestillingsplattform):
åpne siden i nettleseren, scroll helt til bunnen så hele menyen er
lastet, og lagre siden som PDF (Cmd+P → Lagre som PDF). Last så opp
PDF-en i verktøyet som en vanlig fil.

## Viktig om allergener

Verktøyet gjetter aldri blindt — det utleder kun fra ingredienser som
faktisk står i teksten, og flagger alt for manuell bekreftelse. Selger
er ansvarlig for å verifisere mot vendoren før menyen publiseres.

## Kjente begrensninger

- Vision-ekstraksjon er ikke 100 % reproduserbar — gridet er derfor
  redigerbart.
- Bildekvalitet betyr mye: uskarpe bilder gir dårligere treff.
- Choice groups / toppings utledes ikke automatisk. Varianter med ulik
  pris splittes til egne rader.
