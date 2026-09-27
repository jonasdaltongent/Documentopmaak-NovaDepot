#!/usr/bin/env python3
"""
maak_werkdocumenten.py
Genereert het werkdocument voor les 04 "Klaar om te versturen" (Tekstverwerking 3:
documentopmaak en pdf, NovaDepot, ORLO).

  TV3_Documentopmaak.docx   het werkdocument dat de leerling INLEVERT (samen met een pdf)
                            pagina 1-2: de onthaalbrochure van NovaDepot
                            pagina 3:   het vragenblad

Gebruik:  python3 maak_werkdocumenten.py
Vereist:  python-docx

LET OP bij aanpassen: de fouten in de paginaopmaak zijn OPZETTELIJK. Ze zijn het lesmateriaal.
 - De pagina staat LIGGEND met marges van 1 cm. De leerling zet ze staand, A4, marges 2 cm
   (stap 2, de demo van de leraar; rij 1 van het vragenblad is daarom al ingevuld).
 - Bovenaan pagina 1 staat de koptekst als GEWONE TEKST (grijze regel, links uitgelijnd).
   De leerling knipt hem en plakt hem in de echte koptekst (stap 3), rechts uitgelijnd.
 - Er is GEEN kop- of voettekst en er zijn GEEN paginanummers (stap 3 en 4).
 - Tussen deel 1 en deel 2 staan LEGE_REGELS lege alinea's (Enters). Daarmee "duwde" iemand
   deel 2 naar een nieuwe pagina. De leerling wist ze en voegt een pagina-einde in (stap 5).
   Liggend vullen ze de rest van pagina 1, zodat deel 2 ongeveer bovenaan pagina 2 begint. Na de
   nieuwe pagina-instelling (staand, 2 cm) belandt deel 2 daardoor op een verkeerde plek: precies
   het probleem dat de leerling met een pagina-einde oplost.
   De lege alinea's hebben dezelfde alinea-uitlijning en -afstand als de titel van deel 2
   (gecentreerd, 12 pt erna): zo houdt die titel zijn opmaak, hoe de leerling ze ook wist.
 - De brochure zelf volgt de huisstijlkaart van les 03 (Verdana; titel 20 pt vet gecentreerd;
   tussentitel 14 pt vet; tekst 12 pt; lijsten met opsommingstekens; laatste regel rechts).
   Het vragenblad staat bewust in Arial: alles in Verdana hoort bij de brochure.
 - Het document wordt in Drive omgezet naar een Google-document (Drive-instelling "Uploads
   converteren", uploaden in Drive zelf; Classroom maakt daar een kopie per leerling van).
   Controleer na het omzetten of het nog liggend staat, met smalle marges en met de lege regels.
   Zie README.md, paragraaf 2.
 - NovaDepot, alle namen en toestelnummers zijn fictief.
"""
import os

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "..", "assets", "novadepot-logo.png")
UIT = os.path.join(HERE, "TV3_Documentopmaak.docx")

NAVY = RGBColor(0x2A, 0x39, 0x73)
GRIJS = RGBColor(0x80, 0x80, 0x80)
GREY = RGBColor(0x4D, 0x55, 0x73)
LEGE_REGELS = 9           # de "Enters" tussen deel 1 en deel 2 (elk ± 1 cm hoog)
KOPREGEL = "NovaDepot · Onthaalbrochure jobstudenten · oktober 2026"


# ---------- hulpfuncties ----------
def lettertype(run, naam):
    run.font.name = naam
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), naam)


def basis_document():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Verdana"
    st.font.size = Pt(12)
    rpr = st.element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:eastAsia"), "Verdana")
    taal = OxmlElement("w:lang")
    taal.set(qn("w:val"), "nl-BE")
    rpr.append(taal)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.15
    for naam in ("List Bullet",):
        doc.styles[naam].font.name = "Verdana"
        doc.styles[naam].font.size = Pt(12)

    # FOUT 1 (opzettelijk): liggend, marges van 1 cm
    s = doc.sections[0]
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = Cm(29.7), Cm(21.0)
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Cm(1)
    s.header_distance = s.footer_distance = Cm(0.6)

    doc.core_properties.title = "TV3 Documentopmaak - onthaalbrochure NovaDepot"
    doc.core_properties.author = "Toegepaste Informatica"
    return doc


def alinea(doc, tekst="", grootte=12, vet=False, uitlijning=None, na=6, kleur=None, stijl=None):
    p = doc.add_paragraph(style=stijl) if stijl else doc.add_paragraph()
    p.paragraph_format.space_after = Pt(na)
    if uitlijning is not None:
        p.alignment = uitlijning
    if tekst:
        r = p.add_run(tekst)
        r.bold = vet
        r.font.size = Pt(grootte)
        if kleur is not None:
            r.font.color.rgb = kleur
    return p


def titel(doc, tekst):
    return alinea(doc, tekst, grootte=20, vet=True, uitlijning=WD_ALIGN_PARAGRAPH.CENTER, na=12)


def tussentitel(doc, tekst):
    p = alinea(doc, tekst, grootte=14, vet=True, na=4)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.keep_with_next = True
    return p


def opsomming(doc, regels):
    for regel in regels:
        alinea(doc, regel, stijl="List Bullet", na=2)


def lege_regel_als_titel(doc):
    """Een lege alinea met de uitlijning en afstand van de titel van deel 2 (gecentreerd, 12 pt erna)."""
    return alinea(doc, "", uitlijning=WD_ALIGN_PARAGRAPH.CENTER, na=12)


def vb_tekst(doc, s, vet=False, klein=False, na=6, grootte=11):
    """Tekst op het vragenblad (Arial)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(na)
    p.paragraph_format.line_spacing = 1.1
    r = p.add_run(s)
    lettertype(r, "Arial")
    r.bold = vet
    r.font.size = Pt(9.5 if klein else grootte)
    if klein:
        r.font.color.rgb = GREY
    return p


def vb_kop(doc, s, grootte=13, voor=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(voor)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(s)
    lettertype(r, "Arial")
    r.bold = True
    r.font.size = Pt(grootte)
    r.font.color.rgb = NAVY
    return p


# ---------- tabellen met vaste kolommen en randen (schemavolgorde gerespecteerd) ----------
def _voeg_in_voor(ouder, el, opvolgers):
    """Voeg el in vóór het eerste bestaande element uit opvolgers (OOXML-volgorde), anders achteraan."""
    for tag in opvolgers:
        ref = ouder.find(qn(tag))
        if ref is not None:
            ref.addprevious(el)
            return el
    ouder.append(el)
    return el


def vaste_tabel(t, breedtes_cm, randkleur="8C8C8C"):
    """Vaste kolombreedtes en dunne randen, zodat Google Documenten de tabel toont zoals bedoeld."""
    tbl = t._tbl
    tblpr = tbl.tblPr
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = _voeg_in_voor(tblpr, OxmlElement("w:tblW"),
                             ("w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
                              "w:tblLayout", "w:tblCellMar", "w:tblLook"))
    tblw.set(qn("w:w"), str(int(round(sum(breedtes_cm) * 567))))
    tblw.set(qn("w:type"), "dxa")
    randen = OxmlElement("w:tblBorders")
    for kant in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + kant)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), randkleur)
        randen.append(el)
    _voeg_in_voor(tblpr, randen, ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook"))
    indeling = OxmlElement("w:tblLayout")
    indeling.set(qn("w:type"), "fixed")
    _voeg_in_voor(tblpr, indeling, ("w:tblCellMar", "w:tblLook"))
    for kolom, b in zip(tbl.tblGrid.findall(qn("w:gridCol")), breedtes_cm):
        kolom.set(qn("w:w"), str(int(round(b * 567))))
    for rij in t.rows:
        for cel, b in zip(rij.cells, breedtes_cm):
            cel.width = Cm(b)
    return t


def schaduw(cel, kleur="E4E8F6"):
    tcpr = cel._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), kleur)
    tcpr.append(shd)


def cel_tekst(cel, s, vet=False, grootte=10, cursief=False):
    cel.text = ""
    p = cel.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(s)
    lettertype(r, "Arial")
    r.bold = vet
    r.italic = cursief
    r.font.size = Pt(grootte)


def antwoordlijnen(doc, aantal=2):
    for _ in range(aantal):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("_" * 78)
        lettertype(r, "Arial")
        r.font.size = Pt(11)
        r.font.color.rgb = GRIJS


# ---------- de brochure (pagina 1 en 2) ----------
def brochure(doc):
    # FOUT 2 (opzettelijk): de koptekst als gewone tekst bovenaan pagina 1
    alinea(doc, KOPREGEL, grootte=9, kleur=GRIJS, na=12)

    # ---- deel 1 ----
    titel(doc, "Welkom bij NovaDepot!")
    alinea(doc, "Fijn dat je deze eindejaarsperiode bij ons komt werken. In deze brochure lees je "
                "wat je moet weten voor je eerste werkdag. Lees ze rustig door voor je begint.")
    tussentitel(doc, "Wie zijn we?")
    alinea(doc, "NovaDepot is een groothandel en een logistiek bedrijf. In ons magazijn liggen "
                "duizenden producten klaar voor winkels in de hele regio. Elke dag komen er "
                "vrachtwagens aan bij onze laadpoorten, en elke dag vertrekken er bestellingen "
                "naar onze klanten.")
    tussentitel(doc, "Je eerste werkdag")
    opsomming(doc, [
        "Kom om 7.30 uur naar het onthaal.",
        "Breng je identiteitskaart mee.",
        "Om 8.00 uur is er een infosessie in de kantine.",
        "Daarna krijg je veiligheidsschoenen, een fluohesje en een rondleiding door het magazijn.",
    ])
    tussentitel(doc, "Werkuren")
    alinea(doc, "Je werkt van maandag tot en met vrijdag, van 7.30 uur tot 16.00 uur. Om 12.00 uur "
                "heb je een halfuur middagpauze in de kantine.")

    # FOUT 3 (opzettelijk): Enters om deel 2 naar een nieuwe pagina te duwen
    for _ in range(LEGE_REGELS):
        lege_regel_als_titel(doc)

    # ---- deel 2 ----
    titel(doc, "Afspraken en contact")
    tussentitel(doc, "Veilig werken")
    opsomming(doc, [
        "Draag altijd je veiligheidsschoenen en je fluohesje.",
        "Loop alleen op de groene wandelpaden.",
        "Blijf uit de buurt van een rijdende heftruck.",
        "Zie je iets gevaarlijks? Meld het meteen aan je verantwoordelijke.",
    ])
    tussentitel(doc, "Ziek of te laat?")
    alinea(doc, "Verwittig het onthaal vóór 7.30 uur. Zo weet je ploeg op tijd dat je er niet bent.")
    tussentitel(doc, "Wie kan je bellen?")
    opsomming(doc, [
        "Onthaal – toestel 200",
        "Tom Peeters, magazijn – toestel 214",
        "Sara Claes, preventiedienst – toestel 220",
        "Lotte Maes, personeelsdienst – toestel 230",
    ])
    alinea(doc, "Lotte Maes, personeelsdienst – 1 oktober 2026",
           uitlijning=WD_ALIGN_PARAGRAPH.RIGHT, na=0).paragraph_format.space_before = Pt(14)


# ---------- het vragenblad (pagina 3) ----------
def vragenblad(doc):
    doc.add_page_break()

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("VRAGENBLAD — voor je leraar")
    lettertype(r, "Arial")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = NAVY
    vb_tekst(doc, "Les 04 — Klaar om te versturen  ·  Tekstverwerking 3  ·  Toegepaste Informatica  ·  "
                  "NovaDepot", klein=True, na=8)

    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    vaste_tabel(t, [7.0, 4.5, 5.5])
    for i, label in enumerate(["Naam:", "Klas:", "Datum:"]):
        cel_tekst(t.rows[0].cells[i], label + " ", vet=True, grootte=11)

    vb_kop(doc, "Zo werk je", 12, voor=10)
    for s in [
        "1.  Op de lespagina lees je wat je moet doen, stap voor stap.",
        "2.  Pagina 1 en 2 zijn de brochure van NovaDepot. Die maak jij klaar om te versturen.",
        "3.  Op dit vragenblad schrijf je wat je deed en waarom.",
        "4.  Je levert dit document én je pdf in.",
    ]:
        vb_tekst(doc, s, na=1)

    # ---- deel 1: de tabel ----
    vb_kop(doc, "Deel 1 — Wat was er mis?")
    vb_tekst(doc, "Vul rij 2, 3 en 4 in terwijl je werkt. Rij 1 deed je leraar voor.", klein=True)
    t = doc.add_table(rows=5, cols=3)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    vaste_tabel(t, [5.5, 5.5, 6.0])
    for i, kop in enumerate(["Wat was er mis?", "Wat heb jij gedaan?", "Waarom is dat beter?"]):
        cel_tekst(t.rows[0].cells[i], kop, vet=True)
        schaduw(t.rows[0].cells[i])
    rijen = [
        ("1. De pagina stond liggend en de tekst liep tot tegen de rand.",
         "Bestand › Pagina-instelling: staand, A4, marges 2 cm.",
         "Een brochure lees je staand. Met een witte rand leest de tekst rustiger."),
        ("2. De koptekst stond als gewone tekst bovenaan pagina 1.", "", ""),
        ("3. Er stonden geen paginanummers.", "", ""),
        ("4. Met veel Enters werd deel 2 naar een nieuwe pagina geduwd.", "", ""),
    ]
    for i, rij in enumerate(rijen, start=1):
        for j, waarde in enumerate(rij):
            cel_tekst(t.rows[i].cells[j], waarde, cursief=(i == 1 and j > 0))
        t.rows[i].height = Cm(1.6)

    # ---- deel 2: vragen ----
    vb_kop(doc, "Deel 2 — Vragen")
    vb_tekst(doc, "Vraag 1. Kijk naar pagina 2. Staat je koptekst daar ook? Waarom moest je hem maar "
                  "één keer typen?", vet=True)
    antwoordlijnen(doc, 2)
    vb_tekst(doc, "Vraag 2. Lotte schrijft later drie regels extra bij deel 1. Wat gebeurt er dan met "
                  "deel 2 als je Enters had gebruikt? En nu, met een pagina-einde?", vet=True)
    antwoordlijnen(doc, 3)
    vb_tekst(doc, "Vraag 3. Waarom stuur je de brochure als pdf naar de jobstudenten, en niet als "
                  "Google-document? Geef één reden.", vet=True)
    antwoordlijnen(doc, 2)

    vb_kop(doc, "Deel 3 — Tot slot")
    vb_tekst(doc, "Vraag 4. Welke stap vond je vandaag het moeilijkst? Waarom?", vet=True)
    antwoordlijnen(doc, 2)

    # ---- zelfcontrole ----
    vb_kop(doc, "Zelfcontrole — aankruisen vóór je inlevert")
    for s in [
        "Mijn brochure staat staand, op A4, met marges van 2 cm.",
        "De grijze regel staat in de koptekst, rechts, en niet meer bovenaan pagina 1.",
        "Onderaan elke pagina staat Pagina en het juiste nummer.",
        "Tussen deel 1 en deel 2 staan geen lege regels meer.",
        "Deel 2 begint bovenaan pagina 2, met een pagina-einde.",
        "Mijn pdf heeft 3 pagina's en ik heb hem nagekeken.",
        "Mijn pdf staat bij Jouw werk in de opdracht, naast dit document.",
        "Rij 2, 3 en 4 en de vier vragen zijn ingevuld.",
    ]:
        vb_tekst(doc, "☐  " + s, na=1)

    # ---- extra ----
    vb_kop(doc, "Extra — niet verplicht")
    vb_tekst(doc, "Alleen als je al ingeleverd hebt. Klik in de opdracht op Inleveren ongedaan maken. "
                  "Lever daarna opnieuw in. Je pdf hoef je niet opnieuw te maken.", klein=True)
    vb_tekst(doc, "a) Maak van Pagina 1 → Pagina 1 van 3. Kijk op de lespagina bij Extra.", vet=True, na=2)
    vb_tekst(doc, "b) Zet dit logo in je koptekst. Kopieer het, plak het in de koptekst en maak het "
                  "kleiner.", vet=True, na=4)
    if os.path.exists(LOGO):
        doc.add_picture(LOGO, width=Cm(5))


if __name__ == "__main__":
    doc = basis_document()
    brochure(doc)
    vragenblad(doc)
    doc.save(UIT)
    print("Klaar: " + os.path.basename(UIT))
    print("  pagina 1-2: brochure met opzettelijke opmaakfouten (liggend, marges 1 cm, getypte "
          "koptekst, {} lege regels)".format(LEGE_REGELS))
    print("  pagina 3:   vragenblad (rij 1 al ingevuld = de demo)")
