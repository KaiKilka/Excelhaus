import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ── Farben ──────────────────────────────────────────────────────────────────
DUNKELBLAU   = "1F3864"
WEISS        = "FFFFFF"
HELLGRAU     = "F2F2F2"
GRUEN        = "E2EFDA"
BLAU_HELL    = "DDEEFF"
ORANGE_HELL  = "FFF2CC"

def titel_style(cell, bg=DUNKELBLAU, fg=WEISS):
    cell.font = Font(bold=True, color=fg, size=11)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def zeilen_fill(ws, row, col_max, grau=True):
    farbe = HELLGRAU if grau else WEISS
    fill = PatternFill("solid", fgColor=farbe)
    for c in range(1, col_max + 1):
        ws.cell(row=row, column=c).fill = fill

def border_thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def apply_border(ws, row, col_max):
    b = border_thin()
    for c in range(1, col_max + 1):
        ws.cell(row=row, column=c).border = b

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def set_row_height(ws, row, h):
    ws.row_dimensions[row].height = h

wb = openpyxl.Workbook()

# ════════════════════════════════════════════════════════════════════════════
# BLATT 1 – PSP
# ════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "1_PSP"
ws1.sheet_properties.tabColor = "1F3864"

headers1 = ["PSP-Code", "Ebene", "Gewerk / Arbeitspaket", "Beschreibung", "Typ (F/E)", "Anmerkung"]
set_col_widths(ws1, [12, 8, 34, 52, 12, 38])

# Haupttitel
ws1.merge_cells("A1:F1")
c = ws1["A1"]
c.value = "PROJEKTSTRUKTURPLAN – Kernsanierung EFH Baujahr 1969"
titel_style(c)
set_row_height(ws1, 1, 28)

# Header
for col, h in enumerate(headers1, 1):
    cell = ws1.cell(row=2, column=col, value=h)
    titel_style(cell)
set_row_height(ws1, 2, 22)

psp_data = [
    # PSP-Code, Ebene, Gewerk/AP, Beschreibung, Typ, Anmerkung
    ("1",       1, "PLANUNG & VORBEREITUNG",           "Gesamte Planungsphase vor Baubeginn",                           "",    ""),
    ("1.1",     2, "Bestandsaufnahme",                 "Bausubstanz, Schadstoffe (Asbest!), Energieausweis",            "F",   "Sachverständiger für Asbest-Erstinspektion Pflicht"),
    ("1.2",     2, "Energieberatung / iSFP",           "Individueller Sanierungsfahrplan (Pflicht für BEG-Förderung)",  "F",   "Zugelassener Energieberater (BAFA-Liste)"),
    ("1.3",     2, "Architekten- / Fachplanung",       "Leistungsphase 1–4 nach HOAI",                                 "F",   "Architekt für Genehmigungsplanung, Statiker für Dach/PV"),
    ("1.4",     2, "Fördermittelbeantragung",          "KfW BEG EH, BEG Einzelmaßnahmen, BAFA, NRW.Bank etc.",         "F/E", "Antrag VOR Maßnahmenbeginn stellen!"),
    ("1.5",     2, "Ausschreibung & Vergabe",          "Leistungsverzeichnisse erstellen, mind. 3 Angebote einholen",  "F/E", "Eigenleistung: Angebote einholen & vergleichen"),
    ("1.6",     2, "Baugenehmigung",                   "Je nach Bundesland für Dach/Fassade ggf. Baugenehmigung nötig","F",   "Klären: Denkmalschutz? Gestaltungssatzung?"),
    ("2",       1, "SICHERUNG & VORBEREITUNG OBJEKT",  "Baustelle einrichten, Schadstoffsanierung",                    "",    ""),
    ("2.1",     2, "Baustelleneinrichtung",            "Container, WC, Stromanschluss Baustelle, Absperrung",          "F/E", ""),
    ("2.2",     2, "Schadstoffsanierung",              "Asbest (Dach, Boden, Rohre), Bleifarbe, PCB entfernen",        "F",   "Zertifizierter Entsorgungsfachbetrieb Pflicht (TRGS 519)"),
    ("2.3",     2, "Entkernung / Rückbau",             "Alte Böden, Putz, Einbauten, Fenster demontieren",             "F/E", "Eigenleistung möglich, Vorsicht bei tragenden Wänden"),
    ("3",       1, "ROHBAU & DACH",                    "Dach, tragende Konstruktion, Dachdämmung",                     "",    ""),
    ("3.1",     2, "Dachstuhl-Inspektion & Reparatur", "Statik prüfen, schadhafte Sparren ersetzen",                   "F",   "Zimmerer + Statiker; Grundlage für PV-Statik"),
    ("3.2",     2, "Dachdämmung (Aufsparren)",         "WLG 032/035, mind. 200 mm, GEG-konform",                       "F",   "Dachdecker/Zimmerer – luftdichte Ebene kritisch"),
    ("3.3",     2, "Neue Ziegeleindeckung",            "Inkl. Lattung, Konterlattung, Ortgang, First, Traufe",         "F",   "Dachdecker; Grundlage für PV-Montagesystem"),
    ("3.4",     2, "Dachentwässerung",                 "Neue Rinnen, Fallrohre, Anschlüsse Kanalisation",              "F",   ""),
    ("3.5",     2, "Dachflächenfenster / Gauben",      "Einbau inkl. Dämmkragen und Abdichtung",                       "F",   ""),
    ("4",       1, "PV-ANLAGE",                        "Photovoltaik nach Dachfertigstellung",                         "",    ""),
    ("4.1",     2, "PV-Planung & Dimensionierung",     "Auslegung kWp, Speicher, Ausrichtung, Verschattungsanalyse",   "F",   "PV-Fachbetrieb oder zertifizierter Energieberater"),
    ("4.2",     2, "Statik-Nachweis PV",               "Zusatzlast auf Dachstuhl nachweisen",                          "F",   "Statiker; Pflicht für Netzanschluss"),
    ("4.3",     2, "Montage Unterkonstruktion",        "Schienenmontagesy­stem auf neuem Dach befestigen",              "F",   "PV-Fachbetrieb; Dachdeckerkoordination nötig"),
    ("4.4",     2, "Modulmontage & Verkabelung DC",    "Module, Wechselrichter, DC-Verkabelung",                       "F",   "Elektrofachbetrieb/PV-Fachbetrieb"),
    ("4.5",     2, "Netzanschluss & Anmeldung",        "Netzanschlussantrag beim Netzbetreiber, Zählertausch",         "F",   "Nur zugelassener Elektriker; VDE-AR-N 4105"),
    ("4.6",     2, "Marktstammdatenregister",          "Anlage beim MaStR der BNetzA registrieren",                    "F/E", "Eigenleistung möglich (Online-Portal), Frist: 1 Monat"),
    ("4.7",     2, "Inbetriebnahme & Übergabe",        "Protokoll, Einweisung Wechselrichter/App, Messung",            "F",   ""),
    ("5",       1, "HEIZUNG",                          "Neue Heizungsanlage (GEG-konform ab 2024)",                    "",    ""),
    ("5.1",     2, "Heizlastberechnung (DIN 12831)",   "Grundlage für Wärmepumpen-Dimensionierung",                    "F",   "Heizungsbauer/Energieberater; Pflicht für Förderung"),
    ("5.2",     2, "Rückbau alte Heizung",             "Demontage Öl-/Gasheizung, Tank-Entsorgung",                    "F",   "Konzessionspflicht; Ölbehälter: zugelassener Fachbetrieb"),
    ("5.3",     2, "Installation Wärmepumpe",          "Luft-Wasser oder Sole-Wasser inkl. Fundament, Kältemittel",    "F",   "Kältemittelschein Pflicht (EU 517/2014)"),
    ("5.4",     2, "Flächenheizung / Heizkörper",      "Fußbodenheizung (vor Estrich!) oder neue Heizkörper",          "F",   "SHK-Fachbetrieb"),
    ("5.5",     2, "Hydraulischer Abgleich",           "Pflicht für Förderung (BEG), Protokoll erforderlich",          "F",   ""),
    ("5.6",     2, "Warmwasserbereitung",               "Warmwasserspeicher, Trinkwasserzone, Legionellenschutz",        "F",   "TRWI DIN 1988 beachten"),
    ("6",       1, "ELEKTROINSTALLATION",               "Komplette Erneuerung nach VDE",                                "",    ""),
    ("6.1",     2, "Planung Elektro (Stromlaufplan)",  "Lichtkreise, Steckdosen, Verbraucher, Smart Home",             "F",   "Elektrofachbetrieb erstellt Stromlaufplan"),
    ("6.2",     2, "Rückbau alte Installation",        "Entfernen alter Leitungen, Sicherungskasten",                  "F",   "Nur Elektrofachbetrieb; VDE 0100"),
    ("6.3",     2, "Leerrohre verlegen (Schlitze)",    "Kabelkanäle stemmen, Leerrohre einlegen",                      "F/E", "Stemmen = Eigenleistung möglich; Rohre = Fachfirma legt"),
    ("6.4",     2, "Unterputzinstallation",            "NYM-Leitungen, Dosen, Verteiler",                              "F",   "Konzessionspflicht Elektrik; VDE 0100-520"),
    ("6.5",     2, "Unterverteiler / Zähleranlage",    "HAK, Zählerplatz, FI/LS-Kombis nach aktueller Norm",          "F",   "Abnahme durch Netzbetreiber erforderlich"),
    ("6.6",     2, "Endmontage Elektro",               "Schalter, Steckdosen, Leuchten einbauen",                      "F/E", "Leuchten = Eigenleistung; Anschlüsse = Fachfirma"),
    ("6.7",     2, "Prüfprotokoll VDE 0100-600",       "Messung, DGUV-Protokoll, Übergabe",                            "F",   "Pflicht vor Inbetriebnahme"),
    ("7",       1, "SANITÄR",                          "Wasser, Abwasser, Bad",                                        "",    ""),
    ("7.1",     2, "Planung Sanitär",                  "Grundriss Bad/WC, Leitungsführung",                            "F",   ""),
    ("7.2",     2, "Rohinstallation Wasser/Abwasser",  "Steig- und Verteilerleitungen, Abfluss-Gefälle",               "F",   "SHK-Fachbetrieb; TRWI DIN 1988; vor Estrich!"),
    ("7.3",     2, "Abdichtung Nassräume",             "Verbundabdichtung nach DIN 18534",                             "F",   "Haftungstechnisch kritisch; Fliesenleger oder SHK"),
    ("7.4",     2, "Fliesen Bad/WC",                   "Wandfliesen, Bodenfliesen, Fugen",                             "F/E", "Eigenleistung möglich bei handwerklichem Geschick"),
    ("7.5",     2, "Endmontage Sanitär",               "WC, Waschbecken, Dusche/Wanne, Armaturen",                     "F",   "SHK-Fachbetrieb für Druckprüfung"),
    ("8",       1, "FASSADE & AUSSENBEREICH",          "WDVS oder Vorhangfassade, Sockel",                             "",    ""),
    ("8.1",     2, "Fassadenplanung & U-Wert",         "Dämmstärke nach GEG, Befestigungsstatik",                      "F",   ""),
    ("8.2",     2, "Gerüst stellen",                   "Fassadengerüst inkl. Seitengeländer und Einhausung",           "F",   "Gerüstbauer; Arbeitsschutz ArbSchG"),
    ("8.3",     2, "WDVS anbringen",                   "Dämmplatten kleben + dübeln, Armierungsgewebe, Putzträger",   "F",   "Fachbetrieb; sonst Gewährleistung verfällt"),
    ("8.4",     2, "Außenputz / Anstrich",             "Edelputz, Silikonharzfarbe, Sockelabdichtung",                 "F/E", "Eigenleistung Anstrich möglich"),
    ("8.5",     2, "Gerüst abbauen",                   "Nach Fertigstellung Fassade und Dachrandarbeiten",             "F",   ""),
    ("8.6",     2, "Außenanlagen",                     "Terrasse, Wege, Einfriedung",                                  "E",   "Eigenleistung gut möglich"),
    ("9",       1, "FENSTER & TÜREN",                  "Erneuerung aller Öffnungen",                                   "",    ""),
    ("9.1",     2, "Maßaufnahme & Bestellung",         "Rohbaumaße, U-Wert-Anforderung (GEG), Rahmenprofile",          "F",   "Fensterbauer nimmt Maß"),
    ("9.2",     2, "Rückbau alte Fenster/Türen",       "Vorsichtig demontieren, Öffnungen sichern",                    "F/E", "Schnell ausführen wegen Witterung"),
    ("9.3",     2, "Einbau neue Fenster",              "Montage nach RAL-Leitfaden, Anschlussfugen abdichten",         "F",   "RAL-Montage Pflicht für Luftdichtheit/Förderung"),
    ("9.4",     2, "Haustür & Nebentüren",             "Sicherheitsklasse RC2 empfohlen, Thermotrennung",              "F",   ""),
    ("9.5",     2, "Innenfensterbänke & Leibungen",    "Leibungsputz, Fensterbänke innen einbauen",                    "F/E", ""),
    ("10",      1, "INNENPUTZ & TROCKENBAU",           "Wandoberflächen, Ständerwände, Decken",                        "",    ""),
    ("10.1",    2, "Innenputz (Nassputz)",             "Kalk-Zement-Putz, mind. 2-lagig auf Mauerwerk",               "F",   "Maschineller Putz schneller/günstiger"),
    ("10.2",    2, "Trockenbau Ständerwände",          "GK-Wände, Vorsatzschalen für Dämmung/Installation",           "F/E", "Eigenleistung gut möglich (YouTube-Lernkurve)"),
    ("10.3",    2, "Trockenbau Decken",                "Abhangdecken, Dämmung Decke EG/OG",                            "F/E", ""),
    ("10.4",    2, "Spachteln & Schleifen",            "Q2/Q3-Spachtelung je nach Anforderung",                        "F/E", "Q3/Q4 = Fachfirma empfohlen"),
    ("11",      1, "ESTRICH & BÖDEN",                  "Unterböden und Bodenbeläge",                                   "",    ""),
    ("11.1",    2, "Trittschalldämmung verlegen",      "PE-Folie + Trittschalldämmmatte vor Estrich",                  "F/E", "Einfache Eigenleistung möglich"),
    ("11.2",    2, "Heizestrich einbringen",           "Fließestrich (CA) oder Zementestrich (CT), Nivellierung",      "F",   "Estrichleger; Estrich muss heizbar sein (FBH)"),
    ("11.3",    2, "Estrichbelegreife abwarten",       "Trocknungszeit: CT ~28 Tage; CA ~7–14 Tage (beheizt)",         "",    "Kritische Wartezeit! Boden erst nach CM-Messung"),
    ("11.4",    2, "Fliesen Wohnbereich",              "Großformat-Fliesen kleben, fugfrei oder Fuge",                 "F/E", ""),
    ("11.5",    2, "Parkett / Laminat / LVT",         "Klicksystem (E) oder geklebt (F)",                             "F/E", "Eigenleistung sehr gut möglich"),
    ("11.6",    2, "Treppensanierung",                 "Stufen schleifen oder verkleiden",                             "F/E", ""),
    ("12",      1, "MALERARBEITEN",                    "Innen- und ggf. Außenanstrich",                                "",    ""),
    ("12.1",    2, "Grundierung & Vorarbeiten",        "Tiefengrund, Tapete entfernen, Risse schließen",               "F/E", "Eigenleistung gut möglich"),
    ("12.2",    2, "Wandfarbe innen",                  "2× Dispersionsfarbe oder Kalkfarbe",                           "E",   "Sehr gute Eigenleistungsmöglichkeit"),
    ("12.3",    2, "Deckenanstrich",                   "Weiß, ggf. Strukturfarbe",                                     "E",   ""),
    ("12.4",    2, "Lackierarbeiten",                  "Türen, Fensterrahmen, Heizkörper",                             "F/E", "Hochglanz = Fachfirma empfohlen"),
    ("13",      1, "ABNAHME & DOKUMENTATION",          "Abschlussphase",                                               "",    ""),
    ("13.1",    2, "Bauabnahmen je Gewerk",            "Förmliche Abnahme mit Mängelprotokoll",                        "F",   "Sachverständiger empfohlen"),
    ("13.2",    2, "Energieausweis (neu)",             "Bedarfsausweis nach GEG nach Sanierung",                       "F",   "Energieberater; für Verkauf/Vermietung Pflicht"),
    ("13.3",    2, "Fördermittel-Verwendungsnachweis", "Rechnungen sammeln, Nachweise einreichen (KfW/BAFA)",          "F/E", "Frist beachten! Meist 24 Monate"),
    ("13.4",    2, "Dokumentationsmappe anlegen",      "Alle Pläne, Protokolle, Garantien, Revisionsunterlagen",       "E",   "Ordnerstruktur analog + digital"),
    ("13.5",    2, "Einzug & Übergabe",                "Übergabeprotokoll, Zählerstände, Einweisung Technik",          "E",   ""),
]

row = 3
for i, (code, ebene, gewerk, beschr, typ, anm) in enumerate(psp_data):
    grau = (i % 2 == 0)
    zeilen_fill(ws1, row, 6, grau)

    indent = "  " * (ebene - 1)
    ws1.cell(row=row, column=1, value=code)
    ws1.cell(row=row, column=2, value=ebene)
    c_gew = ws1.cell(row=row, column=3, value=indent + gewerk)
    if ebene == 1:
        c_gew.font = Font(bold=True, size=11)
        ws1.cell(row=row, column=1).font = Font(bold=True)
    ws1.cell(row=row, column=4, value=beschr)

    c_typ = ws1.cell(row=row, column=5, value=typ)
    if typ == "F":
        c_typ.fill = PatternFill("solid", fgColor=BLAU_HELL)
    elif typ == "E":
        c_typ.fill = PatternFill("solid", fgColor=GRUEN)
    elif typ == "F/E":
        c_typ.fill = PatternFill("solid", fgColor=ORANGE_HELL)
    c_typ.alignment = Alignment(horizontal="center")

    ws1.cell(row=row, column=6, value=anm)
    apply_border(ws1, row, 6)
    set_row_height(ws1, row, 18)
    row += 1

# Legende
row += 1
ws1.cell(row=row, column=1, value="Legende:").font = Font(bold=True)
ws1.cell(row=row+1, column=1, value="F = Fachfirma Pflicht")
ws1.cell(row=row+1, column=2).fill = PatternFill("solid", fgColor=BLAU_HELL)
ws1.cell(row=row+2, column=1, value="E = Eigenleistung empfohlen")
ws1.cell(row=row+2, column=2).fill = PatternFill("solid", fgColor=GRUEN)
ws1.cell(row=row+3, column=1, value="F/E = Teils Fachfirma, teils Eigenleistung")
ws1.cell(row=row+3, column=2).fill = PatternFill("solid", fgColor=ORANGE_HELL)

ws1.auto_filter.ref = f"A2:{get_column_letter(6)}{row-2}"
ws1.freeze_panes = "A3"


# ════════════════════════════════════════════════════════════════════════════
# BLATT 2 – PROJEKTPLANUNG
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("2_Projektplanung")
ws2.sheet_properties.tabColor = "2E75B6"

headers2 = [
    "Nr.", "Gewerk / Arbeitspaket", "Abhängigkeit von", "Typ",
    "Dauer (Wo)", "Beginn\n(Wo ab Start)", "Ende\n(Wo)", "Begründung Fachfirma-Pflicht", "Risiko"
]
set_col_widths(ws2, [6, 36, 28, 8, 10, 14, 10, 52, 30])

ws2.merge_cells("A1:I1")
c = ws2["A1"]
c.value = "PROJEKTPLANUNG – Kernsanierung EFH Baujahr 1969 (Logische Reihenfolge mit Abhängigkeiten)"
titel_style(c)
set_row_height(ws2, 1, 28)

for col, h in enumerate(headers2, 1):
    cell = ws2.cell(row=2, column=col, value=h)
    titel_style(cell)
set_row_height(ws2, 2, 30)

# Nr, Gewerk, Abhängigkeit, Typ, Dauer(Wo), Beginn(Wo), Ende(Wo), Begründung, Risiko
proj_data = [
    # ── PHASE 0: PLANUNG ──────────────────────────────────────────────────
    ("0",   "── PHASE 0: PLANUNG & VORBEREITUNG ──",          "",                          "",    "",   "",   "", "", ""),
    ("0.1", "Bestandsaufnahme & Asbestgutachten",              "Projektstart",              "F",    2,    1,    2,
     "TRGS 519: Asbestarbeiten nur durch zertifizierte Fachbetriebe; Strafbar bei Eigenleistung",
     "Asbest-Fund verlängert Zeitplan um 2–8 Wo"),
    ("0.2", "Energieberatung & iSFP erstellen",               "0.1",                       "F",    3,    2,    4,
     "BEG-Förderung (KfW/BAFA) setzt Energieberater auf BAFA-Expertenliste voraus",
     "Ohne iSFP kein max. Fördersatz (20% Bonus)"),
    ("0.3", "Statiker beauftragen (Dach + PV)",               "0.1",                       "F",    2,    2,    4,
     "Standsicherheitsnachweis nach LBO; ohne Statik keine Baugenehmigung",
     "Statik-Problem kann Dachplanung komplett ändern"),
    ("0.4", "Architektenleistung LP 1–4",                     "0.2, 0.3",                  "F",    4,    3,    6,
     "HOAI LP 1–4 für Baugenehmigung; Architekt koordiniert Fachplaner",
     "Planungsfehler teuerste Fehler im Projekt"),
    ("0.5", "Fördermittelantrag stellen (KfW/BAFA)",          "0.2",                       "F/E",  1,    4,    5,
     "Antrag MUSS vor Auftragserteilung gestellt sein – keine Rückwirkung!",
     "Vergessener Antrag = Förderverlust 10.000–30.000 €"),
    ("0.6", "Ausschreibung & Angebote einholen",              "0.4, 0.5",                  "F/E",  4,    5,    8,
     "Für öffentliche Förderung mind. 3 Vergleichsangebote pro Gewerk",
     ""),
    ("0.7", "Vergabe & Vertragsschluss",                      "0.6",                       "F/E",  2,    8,   10,
     "",
     "VOB-Verträge empfehlenswert; Zahlungsplan aushandeln"),
    ("0.8", "Baugenehmigung einholen",                        "0.4",                       "F",    8,    4,   12,
     "Je nach Bundesland und Maßnahme erforderlich; Bearbeitungszeit 4–12 Wo",
     "Behörden-Verzögerung kann Projektstart verschieben"),

    # ── PHASE 1: SCHADSTOFF & RÜCKBAU ────────────────────────────────────
    ("1",   "── PHASE 1: SCHADSTOFFSANIERUNG & RÜCKBAU ──",  "",                           "",    "",   "",   "", "", ""),
    ("1.1", "Baustelleneinrichtung",                          "0.7",                        "F/E",  1,   10,   10,
     "",
     ""),
    ("1.2", "Schadstoffsanierung (Asbest, Blei)",            "0.1, 1.1",                   "F",    2,   10,   11,
     "TRGS 519 § 7: Nur zugelassene Betriebe; Entsorgungsnachweis erforderlich",
     "Schwebender Staub, Kontamination des Gebäudes; Bewohner ausquartieren"),
    ("1.3", "Entkernung & Rückbau innen",                    "1.2",                        "F/E",  2,   11,   12,
     "Tragende Wände NUR mit Statiker-Freigabe entfernen",
     "Eigenleistung spart 3.000–8.000 €; Verletzungsgefahr"),
    ("1.4", "Alte Fenster & Türen demontieren",              "1.3",                        "F/E",  1,   12,   12,
     "",
     "Öffnungen sofort sichern (Folie/Holz), Einbruch- und Witterungsschutz"),

    # ── PHASE 2: DACH ────────────────────────────────────────────────────
    ("2",   "── PHASE 2: DACH (Basis für PV!) ──",           "",                           "",    "",   "",   "", "", ""),
    ("2.1", "Dachstuhl prüfen & reparieren",                 "1.3, 0.3",                   "F",    1,   12,   12,
     "Zimmermeister beurteilt Sparrenzustand; Statiker genehmigt PV-Zusatzlast",
     "Morsche Sparren → Mehrkosten 5.000–20.000 €"),
    ("2.2", "Aufsparrendämmung montieren",                   "2.1",                        "F",    2,   13,   14,
     "Luftdichte Ebene + Dampfbremse erfordert Fachkenntnis; Pfusch = Schimmel",
     "Fehler bei Dampfbremse → Schimmel im Dachstuhl, Totalschaden möglich"),
    ("2.3", "Neue Ziegeleindeckung",                         "2.2",                        "F",    2,   14,   15,
     "Dachdecker-Konzession; Gewährleistungspflicht bei Fachfirma",
     "ACHTUNG: PV erst nach vollständiger Eindeckung möglich"),
    ("2.4", "Dachentwässerung & Details",                    "2.3",                        "F",    1,   15,   15,
     "",
     ""),
    ("2.5", "Dachflächenfenster einbauen",                   "2.2",                        "F",    1,   14,   14,
     "Velux/Roto erfordert zertifizierte Montagebetriebe",
     ""),

    # ── PHASE 3: PV-ANLAGE ───────────────────────────────────────────────
    ("3",   "── PHASE 3: PV-ANLAGE (nach Dachfertigstellung!) ──", "",                     "",    "",   "",   "", "", ""),
    ("3.1", "PV-Feinplanung & Modulauswahl",                 "2.3 (Dach fertig), 4.2",     "F",    1,   16,   16,
     "Nur nach fertiger Eindeckung: exakte Maße, Verschattungsanalyse",
     ""),
    ("3.2", "Statik-Nachweis PV-Anlage",                     "2.1, 3.1",                   "F",    1,   16,   16,
     "Netzbetreiber verlangt Statiknachweis; sonst kein Netzanschluss",
     ""),
    ("3.3", "Montage Unterkonstruktion + Module",            "2.3, 3.2",                   "F",    1,   17,   17,
     "PV-Fachbetrieb; Dachdurchdringung muss mit Dachdecker abgestimmt sein",
     "Koordination Dachdecker + PV-Monteur KRITISCH"),
    ("3.4", "DC-Verkabelung & Wechselrichter",               "3.3, 6.3",                   "F",    1,   17,   17,
     "VDE 0100-712; DC-Seite Gleichstrom-gefährlich",
     ""),
    ("3.5", "Netzanschlussantrag & Zählertausch",            "3.4",                        "F",    4,   18,   22,
     "VDE-AR-N 4105; nur zugelassener Netzbetreiber tauscht Zähler",
     "Wartezeit Netzbetreiber 4–12 Wo; früh beantragen!"),
    ("3.6", "Marktstammdatenregister eintragen",             "3.5",                        "F/E",  0,   22,   22,
     "",
     "Frist 1 Monat nach Inbetriebnahme; Bußgeld bis 50.000 €"),

    # ── PHASE 4: HEIZUNG ─────────────────────────────────────────────────
    ("4",   "── PHASE 4: HEIZUNG ──",                        "",                           "",    "",   "",   "", "", ""),
    ("4.1", "Heizlastberechnung DIN 12831",                  "0.2 (Energieberatung)",       "F",    1,   10,   10,
     "Grundlage für korrekte WP-Dimensionierung; Pflicht für BEG-Förderung",
     "Falsche Dimensionierung → Anlage läuft ineffizient"),
    ("4.2", "Rückbau alte Heizanlage / Öltank",             "1.2",                        "F",    1,   11,   11,
     "Öltanksanierung nach AwSV; nur zugelassener Fachbetrieb",
     "Bodenverunreinigung durch Öl → Haftung Bauherr"),
    ("4.3", "Fußbodenheizung (Rohre verlegen)",              "7.1 (Rohbau-E fertig)",       "F",    2,   13,   14,
     "SHK-Fachbetrieb; VOR Estrich verlegen – sonst teurer Nachrüstung",
     "KRITISCH: FBH muss VOR Estrich liegen!"),
    ("4.4", "Installation Wärmepumpe (Außeneinheit)",        "4.3",                        "F",    1,   15,   15,
     "F-Gase-Verordnung: Kältemittelschein Pflicht; Lärm-Abstandsregelung beachten",
     "Aufstellort prüfen: Abstand Nachbar mind. 3 m"),
    ("4.5", "Warmwasserspeicher & Rohrnetz",                 "4.4",                        "F",    1,   15,   15,
     "DVGW/TRWI; Legionellenschutz ab 60°C",
     ""),
    ("4.6", "Hydraulischer Abgleich + Inbetriebnahme",       "4.5, 11.2 (Estrich fertig)", "F",    1,   21,   21,
     "Pflicht für BEG-Förderung; Protokoll erforderlich",
     "Ohne Abgleich ca. 15–20% höherer Energieverbrauch"),

    # ── PHASE 5: ELEKTRO ─────────────────────────────────────────────────
    ("5",   "── PHASE 5: ELEKTROINSTALLATION ──",            "",                           "",    "",   "",   "", "", ""),
    ("5.1", "Elektroplanung & Stromlaufplan",                "0.4 (Architektenplan)",       "F",    1,   10,   10,
     "Stromlaufplan Grundlage für Schlitz- und Leerrohre",
     ""),
    ("5.2", "Schlitze stemmen & Leerrohre",                  "1.3, 5.1",                   "F/E",  2,   12,   13,
     "Stemmen Eigenleistung möglich; Leerrohre fachgerecht nach Plan",
     "Tragende Wände nicht übermäßig schlitzen (Statik!)"),
    ("5.3", "Unterputzinstallation (Rohinstallation)",       "5.2",                        "F",    3,   13,   15,
     "VDE 0100-520; Konzessionspflicht; nur Elektrofachbetrieb",
     "Pfusch = Brandgefahr, Versicherung zahlt nicht"),
    ("5.4", "Unterverteiler / Zähleranlage",                 "5.3",                        "F",    1,   15,   15,
     "TAB Netzbetreiber; Abnahme durch Netzbetreiber Pflicht",
     ""),
    ("5.5", "Endmontage Schalter, Steckdosen, Leuchten",     "10.2 (Putz/GK fertig)",      "F/E",  2,   20,   21,
     "Anschlussklemmen = Fachfirma; Leuchtenmontage = Eigenleistung",
     ""),
    ("5.6", "Prüfprotokoll VDE 0100-600",                    "5.5",                        "F",    1,   21,   21,
     "Messpflicht vor Inbetriebnahme; Ergebnisdokumentation",
     "Ohne Protokoll keine Abnahme durch Versicherung"),

    # ── PHASE 6: FENSTER & TÜREN ─────────────────────────────────────────
    ("6",   "── PHASE 6: FENSTER & TÜREN ──",               "",                            "",    "",   "",   "", "", ""),
    ("6.1", "Maßaufnahme & Bestellung Fenster",              "1.4 (Demontage),  0.5",       "F",    1,   12,   12,
     "Fensterbauer nimmt Maß nach Rückbau; Lieferzeit 4–8 Wochen beachten",
     "Lieferzeit einplanen!"),
    ("6.2", "Einbau neue Fenster",                           "6.1 (+Lieferzeit), 2.x",     "F",    2,   17,   18,
     "RAL-Leitfaden Fenster: dreischichtige Abdichtung Pflicht für Luftdichtheit",
     "Fehler = Kondensat, Schimmel an Laibung"),
    ("6.3", "Haustür & Nebentüren einbauen",                 "6.1",                        "F",    1,   18,   18,
     "",
     ""),
    ("6.4", "Leibungen verputzen / Innenfensterbänke",       "6.2, 10.1",                  "F/E",  1,   19,   19,
     "",
     ""),

    # ── PHASE 7: SANITÄR ─────────────────────────────────────────────────
    ("7",   "── PHASE 7: SANITÄR (Rohinstallation vor Estrich!) ──", "",                   "",    "",   "",   "", "", ""),
    ("7.1", "Sanitär Rohinstallation Wasser/Abwasser",       "1.3",                        "F",    2,   12,   13,
     "TRWI DIN 1988-200; Druckprüfung vor Estrich Pflicht",
     "KRITISCH: Vor Estrich fertig sein; sonst Aufbruch"),
    ("7.2", "Abdichtung Nassräume (DIN 18534)",              "7.1, 10.1",                  "F",    1,   19,   19,
     "Verbundabdichtung; Fliesenleger/SHK; bei Eigenleistung Gewährleistungslücke",
     "Mangelnde Abdichtung = Wasserschaden, Totalschaden möglich"),
    ("7.3", "Fliesen Bad & WC",                              "7.2",                        "F/E",  2,   19,   20,
     "",
     ""),
    ("7.4", "Endmontage Sanitär (WC, Wascht., Dusche)",      "7.3, 11.2",                  "F",    1,   22,   22,
     "Druckprüfung, Einregulierung; SHK-Fachbetrieb",
     ""),

    # ── PHASE 8: FASSADE ─────────────────────────────────────────────────
    ("8",   "── PHASE 8: FASSADE ──",                        "",                           "",    "",   "",   "", "", ""),
    ("8.1", "Gerüst stellen",                                "2.4 (Dach fertig)",           "F",    1,   16,   16,
     "Gerüstbauer; Arbeitsschutzregeln DGUV V38",
     "Gerüst = teuer; Fassade + Dachrand gleichzeitig nutzen!"),
    ("8.2", "WDVS aufbringen",                               "8.1, 6.2 (Fenster eingebaut)","F",   4,   18,   21,
     "Systemzulassung (AbP); keine Mischsysteme; Fachbetrieb für Gewährleistung",
     "Systembruch = Zulassung erlischt"),
    ("8.3", "Außenputz / Anstrich",                          "8.2",                        "F/E",  2,   21,   22,
     "Unterputz = Fachfirma; Anstrich = Eigenleistung möglich",
     ""),
    ("8.4", "Gerüst abbauen",                                "8.3, 2.4",                   "F",    1,   23,   23,
     "",
     ""),

    # ── PHASE 9: INNENAUSBAU ─────────────────────────────────────────────
    ("9",   "── PHASE 9: INNENAUSBAU ──",                    "",                           "",    "",   "",   "", "", ""),
    ("9.1", "Innenputz (nach Rohinstallationen!)",           "5.3, 7.1, 4.3",              "F",    3,   15,   17,
     "Maschinenputz schneller; nach allen Rohinstallationen!",
     "ACHTUNG: Kein Putz vor allen Leerrohren/Schlitzen"),
    ("9.2", "Trockenbau Ständerwände",                       "9.1",                        "F/E",  2,   17,   18,
     "Tragende GK-Wände: Fachfirma; Trennwände: Eigenleistung gut möglich",
     ""),
    ("9.3", "Trockenbau Decken / Abhangdecken",              "9.2",                        "F/E",  2,   18,   19,
     "",
     ""),
    ("9.4", "Spachteln Q2/Q3",                               "9.3",                        "F/E",  2,   19,   20,
     "Q3/Q4 = Fachfirma empfohlen für gleichmäßige Oberfläche",
     ""),

    # ── PHASE 10: ESTRICH & BÖDEN ─────────────────────────────────────────
    ("10",  "── PHASE 10: ESTRICH & BÖDEN ──",               "",                           "",    "",   "",   "", "", ""),
    ("10.1","Trittschalldämmung verlegen",                    "4.3 (FBH), 7.1 (Sanitär)",  "F/E",  1,   15,   15,
     "",
     "KRITISCH: FBH muss VOR Dämmung und Estrich liegen!"),
    ("10.2","Estrich einbringen",                             "10.1, 5.3, 7.1, 4.3",       "F",    1,   15,   15,
     "Estrichleger; Nivellierlaser; Estrich auf FBH = Heizestrich",
     "WARTEZEIT: CT mind. 28 Tage; CA 7–14 Tage – Bauzeitpuffer einplanen!"),
    ("10.3","Estrich-Trocknungszeit",                         "10.2",                       "",     4,   16,   19,
     "CM-Messung vor Belegreife-Freigabe; keine Aktivitäten auf frischem Estrich",
     "Wichtigste Wartezeit im Projekt; Parallele Arbeiten planen"),
    ("10.4","Bodenbeläge verlegen (Fliesen/Parkett/LVT)",    "10.3 (belegreif!), 9.4",     "F/E",  3,   20,   22,
     "Fliesen Großformat = Fachfirma empfohlen; Klick-LVT/-Parkett = Eigenleistung",
     "CM-Wert prüfen! Sonst Gewährleistung Bodenleger erlischt"),

    # ── PHASE 11: MALER ───────────────────────────────────────────────────
    ("11",  "── PHASE 11: MALERARBEITEN ──",                 "",                           "",    "",   "",   "", "", ""),
    ("11.1","Grundierung & Vorarbeiten",                      "9.4 (Spachtel fertig)",      "F/E",  1,   20,   20,
     "",
     ""),
    ("11.2","Wandfarbe & Deckenanstrich innen",               "11.1, 5.5 (Endmontage E.)",  "E",    3,   21,   23,
     "Sehr gute Eigenleistungsmöglichkeit; spart 3.000–8.000 €",
     ""),
    ("11.3","Lackierarbeiten (Türen, Zargen)",                "11.2",                       "F/E",  1,   22,   22,
     "Hochglanz = Fachfirma; Seidenmatt = Eigenleistung möglich",
     ""),

    # ── PHASE 12: ABNAHME ────────────────────────────────────────────────
    ("12",  "── PHASE 12: ABNAHME & DOKUMENTATION ──",       "",                           "",    "",   "",   "", "", ""),
    ("12.1","Bauabnahmen je Gewerk mit Protokoll",            "alle Gewerke",               "F",    2,   23,   24,
     "Sachverständiger empfohlen; Mängelprotokoll = Rechtsgrundlage",
     "Ohne schriftliche Abnahme läuft Gewährleistung nicht"),
    ("12.2","Fördermittel Verwendungsnachweis",               "12.1",                       "F/E",  1,   24,   24,
     "",
     "Frist meist 24 Monate nach Förderzusage; alle Rechnungen sammeln!"),
    ("12.3","Energieausweis (neu) nach Sanierung",            "12.1",                       "F",    1,   24,   24,
     "Pflicht nach GEG § 80 für Verkauf/Vermietung",
     ""),
    ("12.4","Dokumentationsmappe anlegen",                    "laufend",                    "E",    1,   24,   24,
     "",
     "Garantieurkunden, Wartungsverträge, Pläne digital sichern"),
    ("12.5","Einzug",                                         "12.1",                       "E",    "",  24,   24,
     "",
     ""),
]

row = 3
for i, entry in enumerate(proj_data):
    nr, gewerk, abh, typ, dauer, beginn, ende, begruendung, risiko = entry
    grau = (i % 2 == 0)
    zeilen_fill(ws2, row, 9, grau)

    ws2.cell(row=row, column=1, value=nr)

    c_gew = ws2.cell(row=row, column=2, value=gewerk)
    if gewerk.startswith("──"):
        c_gew.font = Font(bold=True, size=11)
        for c in range(1, 10):
            ws2.cell(row=row, column=c).fill = PatternFill("solid", fgColor="D6E4F0")
        ws2.cell(row=row, column=1).font = Font(bold=True)
    else:
        c_gew.alignment = Alignment(indent=1, wrap_text=True)

    ws2.cell(row=row, column=3, value=abh).alignment = Alignment(wrap_text=True)

    c_typ = ws2.cell(row=row, column=4, value=typ)
    if typ == "F":
        c_typ.fill = PatternFill("solid", fgColor=BLAU_HELL)
    elif typ == "E":
        c_typ.fill = PatternFill("solid", fgColor=GRUEN)
    elif typ == "F/E":
        c_typ.fill = PatternFill("solid", fgColor=ORANGE_HELL)
    c_typ.alignment = Alignment(horizontal="center")

    ws2.cell(row=row, column=5, value=dauer).alignment = Alignment(horizontal="center")
    ws2.cell(row=row, column=6, value=beginn).alignment = Alignment(horizontal="center")
    ws2.cell(row=row, column=7, value=ende).alignment = Alignment(horizontal="center")
    ws2.cell(row=row, column=8, value=begruendung).alignment = Alignment(wrap_text=True)
    ws2.cell(row=row, column=9, value=risiko).alignment = Alignment(wrap_text=True)

    apply_border(ws2, row, 9)
    set_row_height(ws2, row, 32)
    row += 1

ws2.auto_filter.ref = f"A2:{get_column_letter(9)}{row-1}"
ws2.freeze_panes = "A3"

# ── Speichern ────────────────────────────────────────────────────────────────
wb.save("/home/user/Excelhaus/sanierungsplan.xlsx")
print("Datei gespeichert: /home/user/Excelhaus/sanierungsplan.xlsx")
