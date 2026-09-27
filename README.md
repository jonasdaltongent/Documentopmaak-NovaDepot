# Handleiding voor de leraar — Les 04: Klaar om te versturen

**Vak:** Toegepaste Informatica
**Doelgroep:** de ORLO-klassen — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten + 30 minuten keuzewerktijd (formatief)
**Context:** NovaDepot, fictief logistiek bedrijf en groothandel (vervolg op les 02–03)
**Toestel:** Windows 10-pc met AZERTY-klavier, Google Chrome en Google Workspace
**Kernleerplandoelen:** `BK2_02.05` / `BK2_02.05.02` (tekstverwerking, minimale inhoud *documentopmaak*) en `BV2_04.02` (digitale inhouden creëren) — toepassen
**Deadline:** vrijdag 2 oktober 2026, 20.00 uur

De leerling maakt een onthaalbrochure van NovaDepot klaar om te versturen: pagina-instelling,
koptekst, paginanummers, een pagina-einde in plaats van lege regels, en een pdf. Hij levert
**twee** bestanden in: zijn werkdocument en zijn pdf.

---

## 1. Inhoud van het pakket

```text
W05 - Les 04 - ORLO - Tekstverwerking - documentopmaak/
├── index.html                   # de lespagina: route · één stap · checklist (zie §1b)
├── presentatie.html             # 8 klassikale dia's voor de lesstart en "Ik doe"
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/
│   ├── novadepot-logo.svg/.png, novadepot-icon.svg   # logo NovaDepot (eigen werk)
│   ├── dalton-gent-logo.png
│   ├── fonts/                   # Atkinson Hyperlegible + Montserrat (OFL, zelf gehost)
│   └── screenshots/             # knop-uitlijnen.png (jouw schermafbeelding uit les 03) + LEESMIJ.md
├── werkdocument/
│   ├── TV3_Documentopmaak.docx  # het werkdocument dat de leerling INLEVERT (met de pdf)
│   └── maak_werkdocumenten.py   # maakt het werkdocument opnieuw (python-docx)
├── lesvoorbereiding.md          # volgens §9.2 van de AI-lesplanner v2
├── dalton-lesfiche.md           # lestijd + keuzewerktijd
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

Er zijn **geen bronbestanden** en dus geen zip-bestand: de brochure zit in het werkdocument zelf,
net als bij les 02–03.

### 1b. Hoe de lespagina werkt

Dezelfde opbouw als les 2 van 3MWWE (AI-lesplanner v2):

- **Links de route** — *Start*, *Theoriekaart*, zeven stappen in twee groepen (*De pagina's
  klaarmaken* · *Versturen*) en *Extra*. Een afgewerkte stap krijgt een groen vinkje.
- **Midden één stap** — vijf vaste blokken: één zin uitleg · *Wat moet je doen?* · hoogstens één
  tip · *Hulp nodig?* (dichtgeklapt) · *Klaar als*.
- **Rechts de checklist** — 22 concrete taken. Op een half scherm staat alles onder elkaar en zie je
  alleen de taken van de huidige stap; in stap 7 staat de hele lijst open.
- **Theoriekaart** — tien kaartjes, met een schets van een pagina (koptekst, marge, tekst,
  voettekst), een kaartje *Uit les 02 en 03*, *Woorden* en *Veelgemaakte fouten*.
- **Geen toestelkeuze**: de klas werkt op Windows. Alles gebeurt in de browser, behalve de pdf
  openen.

`localStorage` bewaart alleen de vinkjes en de laatste stap (voorvoegsel `novadepot_tv3_v1_`), met
een wisknop.

---

## 2. Klaarzetten (± 15 minuten)

### Stap 1 — Publiceren via GitHub Pages ✅ *gebeurd op 27-09-2026*

Op jouw vraag gepubliceerd, vóór je inhoudelijke controle, zodat je het op je gsm kan nakijken.
Pages staat op branch `main`, map `/ (root)`:

- **Repository:** <https://github.com/jonasdaltongent/Documentopmaak-NovaDepot>
- **Lespagina voor de leerlingen:** <https://jonasdaltongent.github.io/Documentopmaak-NovaDepot/>
- **Dia's voor het bord:** <https://jonasdaltongent.github.io/Documentopmaak-NovaDepot/presentatie.html>

Dat adres staat al op dia 7 en in `lesdoelen.json` (veld `bron`). Deel met de leerlingen altijd
het **Pages-adres**, niet de repositorylink. Wil je na je controle iets veranderen, zeg het: ik pas
het aan en push opnieuw. Een wijziging staat 1 à 2 minuten na de push online.

### Stap 2 — Het werkdocument omzetten en nakijken

1. Upload `werkdocument/TV3_Documentopmaak.docx` naar Drive.
2. **Zet het om naar een Google-document.** Classroom maakt de kopie per leerling alleen in
   Google-formaat ([Classroom-help](https://support.google.com/edu/classroom/answer/6020265?hl=nl)).
   Een gedocumenteerde manier: in Drive-instellingen (`drive.google.com/drive/settings`) het vakje
   **Uploads converteren naar de indeling van een Editor van Google Documenten** aanvinken en
   daarna uploaden ([Drive-help 2424368](https://support.google.com/drive/answer/2424368?hl=nl)).
   Let op: die instelling geldt dan voor al je volgende uploads. Het menupad om één bestand om te
   zetten staat niet in de helppagina's die ik kon lezen.
3. **Kijk het omgezette document na.** De fouten zijn het lesmateriaal, ze moeten er nog in zitten:
   - [ ] *Bestand › Pagina-instelling* toont **liggend** en marges van **1 cm**;
   - [ ] bovenaan pagina 1 staat de grijze regel *NovaDepot · Onthaalbrochure jobstudenten · oktober 2026*;
   - [ ] er is **geen** koptekst, voettekst of paginanummer;
   - [ ] met *Bekijken › Opmaakmarkeringen tonen* zie je **negen lege regels** tussen deel 1 en
     *Afspraken en contact*;
   - [ ] pagina 3 is het vragenblad, rij 1 van de tabel is ingevuld, en onderaan staat het logo;
   - [ ] de tekst staat niet vol rode golflijnen (zo wel: zet de taal van het document op Nederlands).
4. Verwijder daarna het `.docx`-bestand uit Drive, zodat je niet per ongeluk het verkeerde bestand
   aan de opdracht hangt.

### Stap 3 — Eén opdracht in Google Classroom

Onderwerp **Tekstverwerking – de basis** (zoals les 02–03):

| | Opdracht: **Tekstverwerking 3 — Klaar om te versturen** |
|---|---|
| **Bijlage 1** | de link naar de lespagina |
| **Bijlage 2** | `TV3_Documentopmaak` (Google-document) — **Een kopie maken voor elke leerling** |
| **Punten** | zonder cijfer (formatief) |
| **Deadline** | vrijdag 2 oktober 2026, 20.00 uur |

> [!NOTE]
> *Een kopie maken voor elke leerling* kan je alleen kiezen **vóór** je de opdracht post. Pas je het
> origineel daarna aan, dan komt dat niet meer in de kopieën van de leerlingen.

Instructietekst (kopieer):

```text
1. Open de lespagina (link). Zet ze links op je scherm.
2. Open je werkdocument TV3_Documentopmaak. Zet het rechts.
3. Volg de stappen op de lespagina. In stap 6 maak je een pdf.
4. Lever je werkdocument én je pdf in. Niet klaar? Lever toch in en schrijf onder Privéreacties tot welke stap je kwam.
```

Je e-mailadres is deze les **niet** nodig: de leerlingen delen niets.

### Afvinklijst vóór de les

- [x] De lespagina is gepubliceerd en het adres op dia 7 klopt. (Getest op 27-09-2026.)
- [ ] Het werkdocument is een **Google-document** en bevat nog alle fouten (stap 2 hierboven).
- [ ] Een testleerling krijgt een eigen kopie met de eigen naam in de titel.
- [ ] Met die testleerling: het menu **Invoegen** toont **Pagina-elementen** en **Eindemarkering**
  (zie §8). Zo niet: zeg het tijdens de demo en pas de lespagina aan.
- [ ] Met die testleerling: *Bestand › Downloaden* geeft een pdf, en *Jouw werk › Toevoegen of
  maken › Bestand* laat die pdf toevoegen.
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden (26-09-2026)

Alle knopnamen op de lespagina, de dia's en het vragenblad komen uit de Nederlandse helppagina's
(ruwe tekst, niet een samenvatting). De helppagina's melden zelf dat ze deels met AI vertaald zijn
en zijn niet altijd eenvormig (dezelfde pagina schrijft *Bekijken* en *Weergeven*). Waar jouw scherm
anders zegt, geldt jouw scherm.

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Opmaakmarkeringen | **Bekijken** › **Opmaakmarkeringen tonen** (ook door jou bevestigd in les 02) | [Docs 6367684](https://support.google.com/docs/answer/6367684?hl=nl) |
| Pagina-instelling | **Bestand** › **Pagina-instelling**, bovenaan **Pagina's**; **Afdrukstand** · **Papierformaat** · **Marges**; **OK** | [Docs 10296604](https://support.google.com/docs/answer/10296604?hl=nl) |
| Met of zonder paginering | kop- en voetteksten en paginanummers bestaan alleen **met paginering** | [Docs 11528737](https://support.google.com/docs/answer/11528737?hl=nl) |
| Koptekst / voettekst | **Invoegen** › **Pagina-elementen** › **Koptekst** of **Voettekst** | [Docs 86629](https://support.google.com/docs/answer/86629?hl=nl) |
| Paginanummer | **Invoegen** › **Pagina-elementen** › **Paginanummer** › **Paginanummer** (plaats kiezen) of **Aantal pagina's** (op de plek van de cursor) | idem |
| Pagina-einde | **Invoegen** › **Eindemarkering** › **Pagina-einde** · **Ctrl + Enter** · wegdoen: klik eronder, **Backspace** | [Docs 11526892](https://support.google.com/docs/answer/11526892?hl=nl) · [Sneltoetsen](https://support.google.com/docs/answer/179738?hl=nl) |
| Rechts uitlijnen | **Ctrl + Shift + R**; werkbalk **Uitlijnen** | [Sneltoetsen](https://support.google.com/docs/answer/179738?hl=nl) · [Docs 1663349](https://support.google.com/docs/answer/1663349?hl=nl) |
| Pdf maken | **Bestand** › **Downloaden**, de pdf-indeling kiezen in de lijst | [Drive 2423534](https://support.google.com/drive/answer/2423534?hl=nl) |
| Download openen | de **downloadlade** rechts naast de adresbalk, op het bestand klikken; of **Meer** › **Downloads** | [Chrome 95759](https://support.google.com/chrome/answer/95759?hl=nl) |
| Bestand toevoegen | **Jouw werk** › **Toevoegen of maken** › **Bestand**, bijlage kiezen, **Toevoegen**; **Verwijderen** naast een bijlage | [Classroom 6020285](https://support.google.com/edu/classroom/answer/6020285?hl=nl) |
| Inleveren | **Inleveren** · **Inleveren ongedaan maken** · **Privéreacties** › **Posten** | idem |

**Niet in de helppagina's, dus beschreven in plaats van benoemd:** de namen van de keuzes in
*Pagina-instelling* (staand, A4, de vier margevakken), de voorbeelden in het menu *Paginanummer*
(*"het voorbeeld met het nummer rechtsonder"*), de exacte naam van de pdf in de lijst van
*Downloaden* (*"het pdf-document (.pdf)"*), en de knoppen in het venster om in Classroom een
bestand van je computer te kiezen.

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| Lesstart | 3' | Dia 1–2: retrieval van les 02–03 (¶ en centreren) |
| **Ik doe** | **5'** | Dia 3–6: lesdoel, voor en na, en twee dingen voordoen: de pagina-instelling (en tonen dat deel 2 verschuift) en de koptekst |
| Jullie doen | 34' | Dia 7 blijft staan; de leerlingen werken stap 1 tot 6 af |
| Controle en indiening | 6' | Stap 7: zelftest, pdf toevoegen, checklist, vraag 4, Inleveren |
| Afsluiting | 2' | Dia 8: waarom een pdf? Vooruitblik op les 05 (tabellen) |

Volledige uitwerking: `lesvoorbereiding.md` §15–17. De hardop-denk-tekst voor de demo staat in de
notities (`N` in `presentatie.html`).

**Eerste rondgang, kijk alleen naar twee dingen.** Ze blokkeren alles wat erna komt:

1. Werkt iedereen in de **eigen kopie** van het werkdocument (met de eigen naam in de titel)?
2. Staan de **opmaakmarkeringen** aan? Zonder ¶ zien ze de lege regels in stap 5 niet.

**Tweede rondgang, rond stap 3:** wie in de koptekst blijft typen, zit er nog in. Eén zin volstaat:
*"Klik in je gewone tekst."*

---

## 4. Verbetersleutel

De leerling levert het werkdocument en de pdf in. De pdf toont de opmaak; de antwoorden lees je in
het werkdocument (vraag 3 en 4 kunnen na het maken van de pdf ingevuld zijn).

### Het werkdocument en de pdf

- Staand, A4, marges 2 cm; 3 pagina's (brochure 2 + vragenblad 1).
- De grijze regel staat in de **koptekst**, rechts, op elke pagina — en niet meer in de tekst.
- Onderaan elke pagina **Pagina 1**, **Pagina 2**, **Pagina 3** (automatisch nummer, *Pagina*
  ervoor getypt).
- Geen lege regels tussen deel 1 en deel 2; een **pagina-einde** vóór *Afspraken en contact*,
  zodat deel 2 bovenaan pagina 2 begint.
- De titel *Afspraken en contact* is nog 20 pt, vet en gecentreerd.

### Vragenblad, deel 1 (rij 1 was het voorbeeld)

| Rij | Wat heb jij gedaan? | Waarom is dat beter? |
|---|---|---|
| 2 | De grijze regel geknipt en in de koptekst geplakt (*Invoegen › Pagina-elementen › Koptekst*), rechts uitgelijnd. | Hij staat nu op elke pagina en ik moest hem maar één keer typen. |
| 3 | *Invoegen › Pagina-elementen › Paginanummer*, rechtsonder, met *Pagina* ervoor. | Het programma telt zelf; komt er een pagina bij, dan klopt het nummer nog. |
| 4 | De lege regels gewist en een pagina-einde ingevoegd vóór *Afspraken en contact*. | Deel 2 blijft bovenaan pagina 2, ook als er tekst bijkomt of de marges veranderen. |

### Vragen

| Vraag | Waar het om gaat |
|---|---|
| 1 | Ja, op elke pagina: een koptekst herhaalt zich vanzelf. Eén keer typen volstaat. |
| 2 | Met Enters schuift deel 2 mee naar beneden en staat het niet meer bovenaan een pagina. Met een pagina-einde blijft deel 2 bovenaan pagina 2 beginnen. |
| 3 | Eén van: ziet er op elk toestel hetzelfde uit · niemand verandert hem per ongeluk · je kan hem openen zonder Google-account · je kan hem zo afdrukken of mailen. |
| 4 | Vrij. De vaakst genoemde stap wordt de lesstart van les 05. |

### Essentiële fouten — geef hier altijd feedback op

- Deel 2 staat nog steeds met lege regels op zijn plaats (of begint niet bovenaan pagina 2).
- Paginanummers zelf getypt (overal hetzelfde nummer).
- De koptekst staat nog als gewone tekst bovenaan pagina 1.
- Geen pdf toegevoegd, of een pdf van een ander document.
- Het document staat op *Zonder paginering* (dan zijn kop- en voettekst onzichtbaar).

**Feedback:** één top en één tip als privéreactie in Classroom. Verbeteren en opnieuw inleveren
mag: de leerling klikt op *Inleveren ongedaan maken*.

---

## 5. Schermafbeeldingen (optioneel)

`index.html` heeft vijf plaatsen voor schermafbeeldingen. `knop-uitlijnen.png` staat er al (jouw
schermafbeelding uit les 03). De andere vier zijn niet verplicht: ontbreekt er een, dan laat de
pagina die plaats weg. Open `index.html?leraar` om te zien waar ze komen. De lijst staat in
`assets/screenshots/LEESMIJ.md`. De nuttigste is die van stap 7 (het venster om een bestand te
kiezen in Classroom).

> Zolang ze ontbreken, zie je vier 404-meldingen in de console van de browser. Dat is normaal.

---

## 6. Het werkdocument opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Vereist `python-docx`. Lees eerst de waarschuwing bovenaan het script: de fouten in de
paginaopmaak (liggend, 1 cm, de grijze regel, de negen lege regels) zijn **opzettelijk**.

---

## 7. Leerplandoelen in je jaaroverzicht

De repository heeft een `pre-push` hook (in `.git/hooks/`, zoals bij les 02–03): bij elke push
roept hij `_tools/update_leerdoelen.py` aan met `lesdoelen.json`. Bij de eerste push op 27-09-2026
zijn de 7 doelen van deze les op het blad *Registratie* van `Leerplandoelen 2026-2027.xlsx`
gezet, en ze staan allemaal op het blad **ORLO**, dus ze worden geteld. Een tweede push voegt niets
dubbel toe.

Met de hand, als dat ooit nodig is:

```bash
python3 "../../_tools/update_leerdoelen.py" lesdoelen.json
```

> [!NOTE]
> De datum in `lesdoelen.json` is een plaatshouder (maandag 28 september 2026). Het script schrijft
> een regel maar één keer: pas je de datum later aan, verbeter hem dan ook op het blad
> *Registratie*. Een hook wordt niet mee gekloond: haal je de repository opnieuw binnen, dan moet
> hij opnieuw geïnstalleerd worden.

## 8. Wat nog moet blijken in de klas

Dingen die ik niet vooraf kon testen. Noteer na de les wat er gebeurde:

1. **De menunamen *Pagina-elementen* en *Eindemarkering*.** Zo staan ze in de huidige Nederlandse
   helppagina's, maar die zijn deels met AI vertaald. Heet het op de schoolcomputers anders
   (bijvoorbeeld een oudere indeling), pas dan stap 3, 4 en 5 en de theoriekaart aan. Voor het
   pagina-einde werkt **Ctrl + Enter** in elk geval (staat in *Hulp nodig?*).
2. **Het omzetten van .docx naar Google-document.** Ik kon niet controleren of de liggende pagina,
   de marges van 1 cm en de lege regels de omzetting overleven, en waar deel 2 daarna precies
   staat. Kijk het na volgens §2, stap 3.
3. **Paginanummer + "Pagina" ervoor typen.** De lespagina laat eerst het automatische nummer
   invoegen en dan *Pagina* ervoor typen. Of het nummer daarbij in een aparte regel komt, kon ik
   niet testen.
4. **Het venster om een bestand toe te voegen in Classroom.** De knoppen daarin staan niet in de
   helppagina's; de lespagina beschrijft ze alleen in woorden. Een schermafbeelding in stap 7 helpt.
5. **De pdf openen.** Op de schoolcomputers opent een pdf in het standaardprogramma (Chrome,
   Microsoft Edge of een ander). Dat maakt niet uit, maar het ziet er misschien anders uit dan de
   lespagina doet vermoeden.
6. **Haalbaarheid.** Zeven stappen met een pdf erbij, in 50 + 30 minuten. Te krap? Gebruik de
   minimumroute uit `lesvoorbereiding.md` §20.
7. **Klasnaam.** `lesdoelen.json` gaat uit van `3ORLO` ("de ORLO-klassen"). Pas het veld `klasnaam`
   aan als het om andere klassen gaat.
