# ⚽ Soccerway Squad Scraper

Scrapt automatisch alle spelersdata van de **Eredivisie** en **Keuken Kampioen Divisie** van [soccerway.com](https://www.soccerway.com) en slaat alles op in één Excel-bestand — één tabblad per club.

---

## 📋 Wat krijg je?

Een `.xlsx` bestand met:
- Een **overzichtstabblad** met alle clubs en het aantal spelers
- Per club een **eigen tabblad** met: rugnummer, naam, positie, leeftijd, wedstrijden, minuten, goals, assists, gele kaarten en rode kaarten

---

## 🖥️ Installatie — stap voor stap

### Stap 1 — Controleer of Python is geïnstalleerd

Open een terminal (Mac/Linux) of Command Prompt (Windows) en typ:

```
python --version
```

Zie je iets als `Python 3.10.0`? Dan ben je klaar voor stap 2.

Zie je een foutmelding? Download dan Python via [python.org/downloads](https://www.python.org/downloads/) en installeer het. **Vink tijdens de installatie "Add Python to PATH" aan.**

---

### Stap 2 — Download dit project

Klik rechtsboven op de groene knop **Code → Download ZIP**, pak het zipbestand uit en onthoud waar je het neergezet hebt.

Of als je Git hebt:
```
git clone https://github.com/jouw-gebruikersnaam/soccerway-scraper.git
cd soccerway-scraper
```

---

### Stap 3 — Ga naar de projectmap

**Windows:**
```
cd C:\Users\JouwNaam\Downloads\soccerway-scraper
```

**Mac/Linux:**
```
cd ~/Downloads/soccerway-scraper
```

Tip: typ `cd ` en sleep dan de map naar het terminalvenster — dan wordt het pad automatisch ingevuld.

---

### Stap 4 — Installeer de benodigde packages

```
pip install -r requirements.txt
```

Dit installeert alle benodigde bibliotheken. Je hoeft dit maar één keer te doen.

---

### Stap 5 — Klaar! Voer het script uit

```
python soccerway_squads.py
```

Na een paar minuten verschijnt het bestand `nederland_voetbal_squads.xlsx` in de projectmap.

---

## ⚙️ Opties

| Optie | Wat het doet | Voorbeeld |
|---|---|---|
| `--output` | Andere bestandsnaam kiezen | `--output mijn_bestand.xlsx` |
| `--competities` | Alleen één competitie scrapen | `--competities eredivisie` of `--competities kkd` |
| `--vertraging` | Meer tijd tussen verzoeken (bij blokkering) | `--vertraging 5` |

Voorbeelden:

```
# Alleen Eredivisie
python soccerway_squads.py --competities eredivisie

# Beide competities, eigen bestandsnaam
python soccerway_squads.py --output squads_2526.xlsx

# Als de site je blokkeert: verhoog de vertraging
python soccerway_squads.py --vertraging 5
```

---

## ❗ Problemen?

**`pip` werkt niet**
Probeer `pip3` in plaats van `pip`.

**`python` werkt niet**
Probeer `python3` in plaats van `python`.

**De site blokkeert het script**
Verhoog de vertraging: `--vertraging 5` of hoger.

**Een club heeft geen data**
De URL van die club op soccerway.com is mogelijk gewijzigd. Pas de URL bovenin `soccerway_squads.py` aan bij de juiste club.

---

## 📁 Bestandsoverzicht

```
soccerway-scraper/
├── soccerway_squads.py   ← het script
├── requirements.txt      ← benodigde packages
├── .gitignore            ← bestanden die Git negeert
└── README.md             ← deze pagina
```

---

## ⚠️ Let op

Dit script is bedoeld voor persoonlijk gebruik. Respecteer de gebruiksvoorwaarden van soccerway.com en wees zuinig met verzoeken (gebruik de vertraging-optie).
