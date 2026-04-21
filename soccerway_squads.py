"""
Soccerway Dutch Leagues Squad Scraper
======================================
Scrapt spelersstatistieken van alle clubs in de Eredivisie en
Keuken Kampioen Divisie en slaat ze op als Excel-bestand (één tabblad per club).

Gebruik:
    python soccerway_squads.py --api-key JOUW_SCRAPERAPI_KEY
    python soccerway_squads.py --api-key JOUW_KEY --output mijn_bestand.xlsx
    python soccerway_squads.py --api-key JOUW_KEY --competities eredivisie

Via GitHub Actions: stel SCRAPERAPI_KEY in als repository secret.

Installeer:
    pip install requests beautifulsoup4 lxml openpyxl
"""

import argparse
import os
import re
import sys
import time
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuratie
# ---------------------------------------------------------------------------

BASE_URL     = "https://www.soccerway.com"
SCRAPERAPI   = "https://api.scraperapi.com"

COMPETITIES = {
    "eredivisie": {
        "naam": "Eredivisie",
        "paginas": [
            f"{BASE_URL}/netherlands/eredivisie/",
            f"{BASE_URL}/netherlands/eredivisie/results/",
            f"{BASE_URL}/netherlands/eredivisie/fixtures/",
        ],
        "kleur":  "1E3A5F",
        "accent": "FF6B00",
    },
    "kkd": {
        "naam": "Keuken Kampioen Divisie",
        "paginas": [
            f"{BASE_URL}/netherlands/eerste-divisie/",
            f"{BASE_URL}/netherlands/eerste-divisie/results/",
            f"{BASE_URL}/netherlands/eerste-divisie/fixtures/",
        ],
        "kleur":  "2D5016",
        "accent": "FFD700",
    },
}

KOLOMMEN = ["#", "Naam", "Positie", "Leeftijd", "Wedstrijden", "MIN",
            "Goals", "Assists", "Gele kaarten", "Rode kaarten"]

# Bekende niet-Nederlandse clubs die in de sidebar staan — uitsluiten
UITGESLOTEN = {
    "inter-miami", "al-nassr", "al-ahli", "al-ahli-sc",
    "messi-lionel", "ronaldo-cristiano",
    "arsenal", "manchester-city", "manchester-united", "chelsea", "liverpool",
    "barcelona", "real-madrid", "psg", "bayern-munich", "juventus",
    "milan", "inter", "borussia-dortmund", "atalanta", "lazio", "napoli",
    "cremonese", "como", "freiburg", "mainz", "bayer-leverkusen", "dortmund",
    "brighton", "burnley", "crystal-palace",
    "betis", "real-salt-lake", "new-england-revolution",
    "celta-vigo", "alaves", "elche", "angers", "nantes",
    "atl-madrid",
}

# Nederlandse clubs die we WILLEN (whitelist voor filtering)
NL_EREDIVISIE = {
    "ajax", "psv", "feyenoord", "az-alkmaar", "fc-twente", "twente",
    "fc-utrecht", "utrecht", "sc-heerenveen", "heerenveen",
    "nec-nijmegen", "nijmegen", "go-ahead-eagles", "g-a-eagles",
    "sparta-rotterdam", "pec-zwolle", "zwolle",
    "fortuna-sittard", "sittard", "fc-groningen", "groningen",
    "nac-breda", "heracles-almelo", "heracles",
    "excelsior", "fc-volendam", "telstar",
    "almere-city", "rkc-waalwijk",
}

NL_KKD = {
    "willem-ii", "fc-den-bosch", "den-bosch", "mvv-maastricht", "maastricht",
    "roda-jc", "roda", "fc-eindhoven", "eindhoven-fc", "de-graafschap",
    "jong-ajax", "jong-az-alkmaar", "jong-az", "jong-psv",
    "jong-fc-utrecht", "jong-utrecht",
    "telstar", "top-oss", "vvv-venlo", "venlo",
    "fc-dordrecht", "dordrecht", "excelsior",
    "ado-den-haag", "den-haag", "sc-cambuur", "cambuur",
    "fc-volendam", "almere-city", "helmond-sport",
    "fc-emmen", "vitesse", "rkc-waalwijk", "waalwijk",
}

NL_CLUBS = {"eredivisie": NL_EREDIVISIE, "kkd": NL_KKD}

POSITIE_MAP = {
    "Goalkeepers": "Keeper",
    "Defenders":   "Verdediger",
    "Midfielders": "Middenvelder",
    "Forwards":    "Aanvaller",
    "Coach":       "Coach",
}

# Wedstrijdlink patroon: /match/slug1-HASH1/slug2-HASH2/
MATCH_RE = re.compile(
    r"/match/([a-z][a-z0-9-]*)-([A-Za-z0-9]{6,12})/([a-z][a-z0-9-]*)-([A-Za-z0-9]{6,12})/"
)

REQUEST_DELAY = 3   # seconden — ScraperAPI is sneller dan directe requests

# ---------------------------------------------------------------------------
# ScraperAPI fetch
# ---------------------------------------------------------------------------

def scraper_fetch(url: str, api_key: str, retries: int = 3,
                  render: bool = False) -> BeautifulSoup | None:
    """Haalt een URL op via ScraperAPI en retourneert BeautifulSoup."""
    params = {
        "api_key":      api_key,
        "url":          url,
        "country_code": "nl",
        "follow_redirect": "true",
    }
    if render:
        params["render"] = "true"

    for poging in range(1, retries + 1):
        try:
            r = requests.get(SCRAPERAPI, params=params, timeout=70)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, "lxml")
                has_table = bool(soup.select_one("table.statistics"))
                has_gk    = "Goalkeepers" in r.text
                if not has_table and not has_gk:
                    # Log eerste 300 tekens zodat we zien wat er terugkomt
                    preview = r.text[:300].replace("\n"," ")
                    print(f"\n    [debug] geen data. Preview: {preview[:200]}")
                return soup
            elif r.status_code == 500:
                print(f"    [retry {poging}/{retries}] status 500...", end=" ", flush=True)
                time.sleep(5)
            else:
                print(f"  [WAARSCHUWING] HTTP {r.status_code} voor {url}", file=sys.stderr)
                return None
        except requests.RequestException as e:
            print(f"  [WAARSCHUWING] {url}: {e}", file=sys.stderr)
            time.sleep(5)
    return None


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


# ---------------------------------------------------------------------------
# Club-hashes extraheren uit wedstrijdlinks
# ---------------------------------------------------------------------------

def haal_clubs_op(comp_key: str, api_key: str) -> list:
    cfg = COMPETITIES[comp_key]
    whitelist = NL_CLUBS[comp_key]
    gevonden = {}   # hash → slug

    for pagina_url in cfg["paginas"]:
        print(f"  Scannen: {pagina_url}")
        soup = scraper_fetch(pagina_url, api_key, render=True)
        if not soup:
            continue

        for a in soup.find_all("a", href=True):
            m = MATCH_RE.search(a["href"])
            if m:
                slug1, hash1, slug2, hash2 = m.groups()
                for slug, hsh in [(slug1, hash1), (slug2, hash2)]:
                    if hsh not in gevonden and slug in whitelist:
                        gevonden[hsh] = slug

        time.sleep(REQUEST_DELAY)

    clubs = []
    for hsh, slug in gevonden.items():
        squad_url = f"{BASE_URL}/team/{slug}/{hsh}/squad/"
        naam = slug.replace("-", " ").title()
        clubs.append((naam, squad_url))

    clubs.sort(key=lambda x: x[0])
    print(f"  {len(clubs)} club(s) gevonden: {', '.join(n for n, _ in clubs)}")
    return clubs


# ---------------------------------------------------------------------------
# Spelersdata scrapen
# ---------------------------------------------------------------------------

def haal_clubnaam_op(soup: BeautifulSoup) -> str:
    for sel in ["div.page-title h1", "h1"]:
        tag = soup.select_one(sel)
        if tag:
            naam = clean(re.sub(r"\s*[-–]\s*Squad.*", "", tag.get_text(), flags=re.I))
            if naam and len(naam) > 1:
                return naam
    title = soup.find("title")
    if title:
        parts = re.split(r"\s*[-–|]\s*", clean(title.get_text()))
        if parts:
            return parts[0].strip()
    return ""


def scrape_squad(url: str, api_key: str) -> tuple:
    # Probeer eerst zonder render (goedkoper, 1 credit)
    soup = scraper_fetch(url, api_key, render=False)

    # Als er geen tabel is, probeer met JS-rendering (2 credits maar werkt bij JS-sites)
    if soup and not soup.select_one("table.statistics") and "Goalkeepers" not in soup.get_text():
        print(f"    [render=true]", end=" ", flush=True)
        soup = scraper_fetch(url, api_key, render=True)
    if not soup:
        return "", []

    clubnaam = haal_clubnaam_op(soup)
    spelers = []
    current_positie = "Onbekend"

    for row in soup.select("table.statistics tr, table tr"):
        th = row.find("th")
        if th:
            tekst = clean(th.get_text())
            if tekst in POSITIE_MAP:
                current_positie = POSITIE_MAP[tekst]
            continue

        if current_positie == "Coach":
            continue

        tds = row.find_all("td")
        if len(tds) < 3:
            continue

        def td(i):
            return clean(tds[i].get_text()) if i < len(tds) else "-"

        naam_td = tds[1]
        a_tag = naam_td.find("a")
        naam = clean(a_tag.get_text() if a_tag else naam_td.get_text())

        if not naam or naam in ("#", "Name", "") or naam.isdigit():
            continue

        try:
            spelers.append({
                "#":            td(0),
                "Naam":         naam,
                "Positie":      current_positie,
                "Leeftijd":     td(2),
                "Wedstrijden":  td(3),
                "MIN":          td(4),
                "Goals":        td(5),
                "Assists":      td(6),
                "Gele kaarten": td(7),
                "Rode kaarten": td(8),
            })
        except (IndexError, AttributeError):
            continue

    return clubnaam, spelers


# ---------------------------------------------------------------------------
# Excel opmaak
# ---------------------------------------------------------------------------

def maak_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def schrijf_tabblad(wb, clubnaam, comp_cfg, spelers):
    tab = re.sub(r"[/\\?*:\[\]]", "", clubnaam)[:31]
    if tab in [ws.title for ws in wb.worksheets]:
        tab = tab[:28] + " (2)"

    ws = wb.create_sheet(title=tab)
    ws.sheet_properties.tabColor = comp_cfg["kleur"]
    border = maak_border()

    ws.merge_cells("A1:J1")
    ws["A1"] = clubnaam
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", start_color=comp_cfg["kleur"])
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:J2")
    ws["A2"] = comp_cfg["naam"]
    ws["A2"].font = Font(name="Arial", italic=True, size=10, color="FFFFFF")
    ws["A2"].fill = PatternFill("solid", start_color=comp_cfg["accent"])
    ws["A2"].alignment = Alignment(horizontal="center")

    for col, h in enumerate(KOLOMMEN, 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", start_color="2C3E50")
        c.alignment = Alignment(horizontal="center")
        c.border = border

    pk = {"Keeper":"E8F4FD","Verdediger":"E8F8E8","Middenvelder":"FFF8E1","Aanvaller":"FDE8E8","Onbekend":"F5F5F5"}
    for ri, sp in enumerate(spelers, 4):
        vul = PatternFill("solid", start_color=pk.get(sp.get("Positie","Onbekend"),"FFFFFF"))
        for col, h in enumerate(KOLOMMEN, 1):
            c = ws.cell(row=ri, column=col, value=sp.get(h,"-"))
            c.font = Font(name="Arial", size=10)
            c.fill = vul
            c.border = border
            c.alignment = Alignment(horizontal="center" if col != 2 else "left")

    for i, b in enumerate([6,30,16,10,14,8,8,9,14,13], 1):
        ws.column_dimensions[get_column_letter(i)].width = b
    ws.freeze_panes = "A4"


def maak_overzicht(wb, data):
    ws = wb.active
    ws.title = "Overzicht"
    ws.sheet_properties.tabColor = "1A1A2E"

    ws.merge_cells("A1:D1")
    ws["A1"] = "Nederlandse Voetbal — Spelersdata Overzicht"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", start_color="1A1A2E")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28

    for col, h in enumerate(["Club","Competitie","Spelers","Tabblad"], 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", start_color="2C3E50")
        c.alignment = Alignment(horizontal="center")

    for ri, d in enumerate(data, 3):
        ws.cell(ri,1,d["club"]).font   = Font(name="Arial",size=10)
        ws.cell(ri,2,d["comp"]).font   = Font(name="Arial",size=10)
        c = ws.cell(ri,3,d["aantal"]); c.font=Font(name="Arial",size=10); c.alignment=Alignment(horizontal="center")
        ws.cell(ri,4,d["tab"]).font    = Font(name="Arial",size=10)
        if ri % 2 == 0:
            for col in range(1,5):
                ws.cell(ri,col).fill = PatternFill("solid",start_color="F0F4FF")

    for col, w in zip("ABCD",[28,28,10,28]):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Hoofdprogramma
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Soccerway squad scraper → Excel via ScraperAPI")
    parser.add_argument("--api-key",  "-k", default=os.getenv("SCRAPERAPI_KEY"),
                        help="ScraperAPI sleutel (of stel SCRAPERAPI_KEY omgevingsvariabele in)")
    parser.add_argument("--output",   "-o", default="nederland_voetbal_squads.xlsx")
    parser.add_argument("--competities","-c", nargs="+",
                        choices=["eredivisie","kkd"], default=["eredivisie","kkd"])
    args = parser.parse_args()

    if not args.api_key:
        print("FOUT: geen ScraperAPI key opgegeven.\n"
              "Gebruik --api-key JOUW_KEY of stel de SCRAPERAPI_KEY omgevingsvariabele in.",
              file=sys.stderr)
        sys.exit(1)

    wb = Workbook()
    overzicht = []

    print(f"\nSoccerway Squad Scraper (via ScraperAPI)")
    print("=" * 50)
    print(f"Output: {args.output}\n")

    for comp_key in args.competities:
        cfg = COMPETITIES[comp_key]
        print(f"\n▶ {cfg['naam']}")

        clubs = haal_clubs_op(comp_key, args.api_key)
        if not clubs:
            print("  [OVERGESLAGEN] Geen clubs gevonden")
            continue

        print(f"\n  Squads scrapen ({len(clubs)} clubs):")
        for i, (naam_slug, url) in enumerate(clubs, 1):
            print(f"  [{i}/{len(clubs)}] {naam_slug} ...", end=" ", flush=True)
            echte_naam, spelers = scrape_squad(url, args.api_key)
            clubnaam = echte_naam or naam_slug

            if spelers:
                schrijf_tabblad(wb, clubnaam, cfg, spelers)
                print(f"✓ {len(spelers)} spelers")
            else:
                print("⚠ Geen data")

            overzicht.append({
                "club":   clubnaam,
                "comp":   cfg["naam"],
                "aantal": len(spelers),
                "tab":    re.sub(r"[/\\?*:\[\]]","",clubnaam)[:31],
            })

            if i < len(clubs):
                time.sleep(REQUEST_DELAY)

    maak_overzicht(wb, overzicht)
    wb.save(args.output)
    totaal = sum(d["aantal"] for d in overzicht)
    print(f"\n{'='*50}")
    print(f"✓ Opgeslagen: {args.output}")
    print(f"  {len(wb.worksheets)} tabbladen  |  {totaal} spelers totaal")


if __name__ == "__main__":
    main()
