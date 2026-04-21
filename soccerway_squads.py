"""
Soccerway Dutch Leagues Squad Scraper
======================================
Scrapt spelersstatistieken van alle clubs in de Eredivisie en
Keuken Kampioen Divisie en slaat ze op als Excel-bestand (één tabblad per club).

Gebruik:
    python soccerway_squads.py
    python soccerway_squads.py --output mijn_bestand.xlsx
    python soccerway_squads.py --competities eredivisie kkd

Installeer benodigdheden:
    pip install requests beautifulsoup4 lxml openpyxl
"""

import argparse
import re
import sys
import time

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuratie
# ---------------------------------------------------------------------------

BASE_URL = "https://www.soccerway.com"

COMPETITIES = {
    "eredivisie": {
        "naam": "Eredivisie",
        "url": f"{BASE_URL}/netherlands/eredivisie/",
        "kleur": "1E3A5F",       # donkerblauw
        "accent": "FF6B00",      # oranje
    },
    "kkd": {
        "naam": "Keuken Kampioen Divisie",
        "url": f"{BASE_URL}/netherlands/eerste-divisie/",
        "kleur": "2D5016",       # donkergroen
        "accent": "FFD700",      # goud
    },
}

# Handmatige clublijsten (soccerway laadt clubs dynamisch via JS)
CLUBS = {
    "eredivisie": [
        ("AFC Ajax",            f"{BASE_URL}/team/ajax/iMt2JyI8/squad/"),
        ("AZ",                  f"{BASE_URL}/team/az/qmCZMHn0/squad/"),
        ("FC Twente",           f"{BASE_URL}/team/fc-twente/2qVFxFJZ/squad/"),
        ("PSV",                 f"{BASE_URL}/team/psv/zBz87jTv/squad/"),
        ("Feyenoord",           f"{BASE_URL}/team/feyenoord/KWjGU5Bk/squad/"),
        ("FC Utrecht",          f"{BASE_URL}/team/fc-utrecht/kT5JeZpX/squad/"),
        ("Vitesse",             f"{BASE_URL}/team/vitesse/L3kPqiLW/squad/"),
        ("SC Heerenveen",       f"{BASE_URL}/team/sc-heerenveen/lrwLHfS6/squad/"),
        ("Feyenoord",           f"{BASE_URL}/team/feyenoord/KWjGU5Bk/squad/"),
        ("NEC Nijmegen",        f"{BASE_URL}/team/nec-nijmegen/GBjhSCmB/squad/"),
        ("Go Ahead Eagles",     f"{BASE_URL}/team/go-ahead-eagles/pS1c9zEH/squad/"),
        ("Sparta Rotterdam",    f"{BASE_URL}/team/sparta-rotterdam/C6IyQz9i/squad/"),
        ("PEC Zwolle",          f"{BASE_URL}/team/pec-zwolle/T6yYlJo0/squad/"),
        ("RKC Waalwijk",        f"{BASE_URL}/team/rkc-waalwijk/7i27wXGE/squad/"),
        ("Almere City FC",      f"{BASE_URL}/team/almere-city-fc/OVDMU52H/squad/"),
        ("Fortuna Sittard",     f"{BASE_URL}/team/fortuna-sittard/F6nJ0WyM/squad/"),
        ("FC Groningen",        f"{BASE_URL}/team/fc-groningen/zO0qhNH9/squad/"),
        ("NAC Breda",           f"{BASE_URL}/team/nac-breda/gHlqCHQM/squad/"),
        ("Heracles Almelo",     f"{BASE_URL}/team/heracles-almelo/wJxQU6IS/squad/"),
    ],
    "kkd": [
        ("Willem II",           f"{BASE_URL}/team/willem-ii/6u3qag0G/squad/"),
        ("FC Den Bosch",        f"{BASE_URL}/team/fc-den-bosch/fzJovhPi/squad/"),
        ("MVV Maastricht",      f"{BASE_URL}/team/mvv-maastricht/3OcMRbzN/squad/"),
        ("Roda JC",             f"{BASE_URL}/team/roda-jc/yNH3bCeU/squad/"),
        ("FC Eindhoven",        f"{BASE_URL}/team/fc-eindhoven/9pZvJo01/squad/"),
        ("De Graafschap",       f"{BASE_URL}/team/de-graafschap/hBkSs3M6/squad/"),
        ("Jong Ajax",           f"{BASE_URL}/team/jong-ajax/e4EBSz4l/squad/"),
        ("Jong AZ",             f"{BASE_URL}/team/jong-az/Xbq4FcnC/squad/"),
        ("Jong PSV",            f"{BASE_URL}/team/jong-psv/zZlhJf1t/squad/"),
        ("Jong FC Utrecht",     f"{BASE_URL}/team/jong-fc-utrecht/OhaMRMaS/squad/"),
        ("Telstar",             f"{BASE_URL}/team/telstar/R0Y2Liyy/squad/"),
        ("TOP Oss",             f"{BASE_URL}/team/top-oss/aPl93Aex/squad/"),
        ("VVV-Venlo",           f"{BASE_URL}/team/vvv-venlo/WxI7YYBY/squad/"),
        ("FC Dordrecht",        f"{BASE_URL}/team/fc-dordrecht/Cq2a6P83/squad/"),
        ("Excelsior",           f"{BASE_URL}/team/excelsior/gOt7aOVi/squad/"),
        ("ADO Den Haag",        f"{BASE_URL}/team/ado-den-haag/8lmGIBzD/squad/"),
        ("Cambuur",             f"{BASE_URL}/team/cambuur/cHb5NHWP/squad/"),
        ("FC Volendam",         f"{BASE_URL}/team/fc-volendam/TnVgniAS/squad/"),
        ("Almere City FC",      f"{BASE_URL}/team/almere-city-fc/OVDMU52H/squad/"),
        ("Helmond Sport",       f"{BASE_URL}/team/helmond-sport/qbXiNzKJ/squad/"),
    ],
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "nl-NL,nl;q=0.9,en;q=0.8",
    "Referer": BASE_URL + "/",
}

KOLOMMEN = ["#", "Naam", "Positie", "Leeftijd", "Wedstrijden", "MIN", "Goals", "Assists", "Gele kaarten", "Rode kaarten"]
REQUEST_DELAY = 2


# ---------------------------------------------------------------------------
# Scraping
# ---------------------------------------------------------------------------

def fetch(url: str, session: requests.Session) -> BeautifulSoup:
    try:
        r = session.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        return BeautifulSoup(r.text, "lxml")
    except requests.RequestException as e:
        print(f"  [WAARSCHUWING] Fout bij ophalen {url}: {e}", file=sys.stderr)
        return None


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "")).strip()


def scrape_squad(url: str, session: requests.Session) -> tuple[str, list[dict]]:
    """Haal spelersdata op. Geeft (stadiumnaam, lijst van spelers) terug."""
    soup = fetch(url, session)
    if not soup:
        return "", []

    spelers = []
    positie = "Onbekend"

    # Detecteer positiekoppen (h2/th met 'Goalkeepers', 'Defenders', etc.)
    positie_map = {
        "Goalkeepers": "Keeper",
        "Defenders": "Verdediger",
        "Midfielders": "Middenvelder",
        "Forwards": "Aanvaller",
        "Coach": "Coach",
    }

    table = soup.select_one("table.statistics")
    if not table:
        # Sommige pagina's hebben de data in een andere tabel
        tables = soup.find_all("table")
        for t in tables:
            if t.find("th", string=re.compile("Name|MIN")):
                table = t
                break

    if not table:
        print("  [WAARSCHUWING] Geen spelerstabel gevonden.", file=sys.stderr)
        return "", []

    # Doorloop alle rijen inclusief sectie-headers
    current_positie = "Onbekend"
    for row in soup.select("table.statistics tr, table tr"):
        # Sectie-header rij
        header_cell = row.find("th", class_=re.compile("player-position|position-header"))
        if not header_cell:
            # Probeer het via tekst
            th = row.find("th")
            if th and clean(th.get_text()) in positie_map:
                current_positie = positie_map[clean(th.get_text())]
                continue

        tds = row.find_all("td")
        if not tds or len(tds) < 3:
            continue

        # Sla header-achtige rijen over
        if row.find("th"):
            tekst = clean(row.find("th").get_text())
            if tekst in positie_map:
                current_positie = positie_map[tekst]
            continue

        # Probeer spelerdata te parsen
        try:
            def td(i, default="-"):
                if i < len(tds):
                    return clean(tds[i].get_text()) or default
                return default

            rugnummer = td(0)
            # Naam staat soms met link
            naam_td = tds[1] if len(tds) > 1 else None
            naam = "-"
            if naam_td:
                a = naam_td.find("a")
                naam = clean(a.get_text() if a else naam_td.get_text())

            if not naam or naam in ("#", "Name", ""):
                continue

            leeftijd = td(2)
            wedstrijden = td(3)
            minuten = td(4)
            goals = td(5)
            assists = td(6)
            geel = td(7)
            rood = td(8)

            # Sla rijen over waar naam puur een getal is (foutief geparsed)
            if naam.isdigit():
                continue

            if current_positie == "Coach":
                continue

            spelers.append({
                "#": rugnummer,
                "Naam": naam,
                "Positie": current_positie,
                "Leeftijd": leeftijd,
                "Wedstrijden": wedstrijden,
                "MIN": minuten,
                "Goals": goals,
                "Assists": assists,
                "Gele kaarten": geel,
                "Rode kaarten": rood,
            })
        except (IndexError, AttributeError):
            continue

    return current_positie, spelers


# ---------------------------------------------------------------------------
# Excel opmaak
# ---------------------------------------------------------------------------

def maak_stijlen():
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    return border


def schrijf_tabblad(wb: Workbook, clubnaam: str, competitie_cfg: dict,
                    spelers: list[dict], tab_kleur: str):
    # Tabbladnaam max 31 tekens, geen verboden tekens
    tab_naam = re.sub(r"[/\\?*:\[\]]", "", clubnaam)[:31]

    # Voorkom dubbele namen
    bestaand = [ws.title for ws in wb.worksheets]
    if tab_naam in bestaand:
        tab_naam = tab_naam[:28] + " (2)"

    ws = wb.create_sheet(title=tab_naam)
    ws.sheet_properties.tabColor = tab_kleur

    accent_hex = competitie_cfg["accent"]
    header_hex = competitie_cfg["kleur"]

    # Rij 1: clubnaam als titel
    ws.merge_cells("A1:J1")
    ws["A1"] = clubnaam
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", start_color=header_hex)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Rij 2: competitienaam
    ws.merge_cells("A2:J2")
    ws["A2"] = competitie_cfg["naam"]
    ws["A2"].font = Font(name="Arial", italic=True, size=10, color="FFFFFF")
    ws["A2"].fill = PatternFill("solid", start_color=accent_hex)
    ws["A2"].alignment = Alignment(horizontal="center")

    # Rij 3: kolomkoppen
    header_fill = PatternFill("solid", start_color="2C3E50")
    border = maak_stijlen()

    for col, kol_naam in enumerate(KOLOMMEN, start=1):
        cel = ws.cell(row=3, column=col, value=kol_naam)
        cel.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cel.fill = header_fill
        cel.alignment = Alignment(horizontal="center")
        cel.border = border

    # Data
    positie_kleuren = {
        "Keeper":       "E8F4FD",
        "Verdediger":   "E8F8E8",
        "Middenvelder": "FFF8E1",
        "Aanvaller":    "FDE8E8",
        "Onbekend":     "F5F5F5",
    }

    for rij_idx, speler in enumerate(spelers, start=4):
        pos = speler.get("Positie", "Onbekend")
        vulkleur = positie_kleuren.get(pos, "FFFFFF")
        vul = PatternFill("solid", start_color=vulkleur)

        for col, kol_naam in enumerate(KOLOMMEN, start=1):
            waarde = speler.get(kol_naam, "-")
            cel = ws.cell(row=rij_idx, column=col, value=waarde)
            cel.font = Font(name="Arial", size=10)
            cel.fill = vul
            cel.border = border
            cel.alignment = Alignment(horizontal="center" if col != 2 else "left")

    # Kolombreedtes
    breedtes = [6, 30, 16, 10, 14, 8, 8, 9, 14, 13]
    for i, breedte in enumerate(breedtes, start=1):
        ws.column_dimensions[get_column_letter(i)].width = breedte

    # Freeze titelbalk + headers
    ws.freeze_panes = "A4"


def maak_overzicht_tab(wb: Workbook, club_data: list[dict]):
    ws = wb.active
    ws.title = "📋 Overzicht"
    ws.sheet_properties.tabColor = "1A1A2E"

    ws.merge_cells("A1:E1")
    ws["A1"] = "Nederlandse Voetbal — Spelersdata Overzicht"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", start_color="1A1A2E")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28

    headers = ["Club", "Competitie", "Aantal spelers", "Tabblad"]
    header_fill = PatternFill("solid", start_color="2C3E50")
    for col, h in enumerate(headers, start=1):
        cel = ws.cell(row=2, column=col, value=h)
        cel.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cel.fill = header_fill
        cel.alignment = Alignment(horizontal="center")

    for rij, info in enumerate(club_data, start=3):
        ws.cell(row=rij, column=1, value=info["club"]).font = Font(name="Arial", size=10)
        ws.cell(row=rij, column=2, value=info["competitie"]).font = Font(name="Arial", size=10)
        ws.cell(row=rij, column=3, value=info["aantal"]).font = Font(name="Arial", size=10)
        ws.cell(row=rij, column=3).alignment = Alignment(horizontal="center")
        ws.cell(row=rij, column=4, value=info["tab"]).font = Font(name="Arial", size=10)

        if rij % 2 == 0:
            vul = PatternFill("solid", start_color="F0F4FF")
            for col in range(1, 5):
                ws.cell(row=rij, column=col).fill = vul

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 28


# ---------------------------------------------------------------------------
# Hoofdprogramma
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Soccerway squad scraper → Excel")
    parser.add_argument("--output", "-o", default="nederland_voetbal_squads.xlsx",
                        help="Naam outputbestand (standaard: nederland_voetbal_squads.xlsx)")
    parser.add_argument("--competities", "-c", nargs="+",
                        choices=["eredivisie", "kkd"], default=["eredivisie", "kkd"],
                        help="Welke competities scrapen (standaard: beide)")
    parser.add_argument("--vertraging", "-v", type=float, default=REQUEST_DELAY,
                        help=f"Seconden vertraging tussen verzoeken (standaard: {REQUEST_DELAY})")
    args = parser.parse_args()

    session = requests.Session()
    wb = Workbook()
    club_data_overzicht = []

    totaal_clubs = sum(len(CLUBS[c]) for c in args.competities)
    print(f"\nSoccerway Squad Scraper")
    print(f"{'='*50}")
    print(f"Competities : {', '.join(args.competities)}")
    print(f"Clubs       : {totaal_clubs}")
    print(f"Output      : {args.output}")
    print(f"{'='*50}\n")

    teller = 0
    for comp_key in args.competities:
        cfg = COMPETITIES[comp_key]
        clubs = CLUBS[comp_key]
        print(f"\n▶ {cfg['naam']} ({len(clubs)} clubs)")

        for clubnaam, url in clubs:
            teller += 1
            print(f"  [{teller}/{totaal_clubs}] {clubnaam} ...", end=" ", flush=True)

            _, spelers = scrape_squad(url, session)

            if spelers:
                tab_kleur = cfg["kleur"]
                schrijf_tabblad(wb, clubnaam, cfg, spelers, tab_kleur)
                print(f"✓ {len(spelers)} spelers")
            else:
                print("⚠ Geen data gevonden")
                spelers = []

            club_data_overzicht.append({
                "club": clubnaam,
                "competitie": cfg["naam"],
                "aantal": len(spelers),
                "tab": re.sub(r"[/\\?*:\[\]]", "", clubnaam)[:31],
            })

            if teller < totaal_clubs:
                time.sleep(args.vertraging)

    # Overzichtstabblad als eerste
    maak_overzicht_tab(wb, club_data_overzicht)

    wb.save(args.output)
    print(f"\n{'='*50}")
    print(f"✓ Opgeslagen: {args.output}")
    print(f"  {len(wb.worksheets)} tabbladen  |  {sum(d['aantal'] for d in club_data_overzicht)} spelers totaal")


if __name__ == "__main__":
    main()
