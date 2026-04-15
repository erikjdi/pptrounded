# Hoekradius — PowerPoint Add-in

Stel de hoekradius van afgeronde rechthoeken in PowerPoint in met een exacte waarde.
Werkt op Mac en Windows, compatibel met AutoSave en SharePoint.

---

## Bestanden

```
pptx-addin/
├── taskpane.html     ← De UI van het taakvenster
├── taskpane.js       ← Logica (Office JS API)
├── manifest.xml      ← Add-in manifest (aanpassen met jouw URL)
└── assets/
    ├── icon-16.png   ← Zelf aan te maken (16×16 px)
    ├── icon-32.png   ← Zelf aan te maken (32×32 px)
    └── icon-80.png   ← Zelf aan te maken (80×80 px)
```

---

## Installatie (3 opties)

### Optie A — GitHub Pages (aanbevolen, gratis)

1. Maak een GitHub repository aan (mag privé zijn)
2. Zet `taskpane.html` en `taskpane.js` in de root
3. Activeer GitHub Pages via Settings → Pages → Branch: main
4. Vervang in `manifest.xml` alle `https://JOUW-URL` door jouw GitHub Pages URL
   - Voorbeeld: `https://jandevries.github.io/pptx-hoekradius`
5. Sla het manifest op

### Optie B — SharePoint / OneDrive

1. Upload `taskpane.html` en `taskpane.js` naar een SharePoint-bibliotheek
2. Zorg dat de bestanden publiek toegankelijk zijn (of via organisatie-intranet)
3. Vervang `https://JOUW-URL` in manifest.xml met de directe URL naar taskpane.html
4. Sla het manifest op

### Optie C — Lokaal testen (alleen Windows)

1. Zet de map op een netwerkshare of gebruik localhost met een tool als `http-server`
2. Gebruik het manifest direct via "Sideloading" (zie hieronder)

---

## Add-in laden in PowerPoint

### Windows
1. Ga naar **Invoegen → Add-ins → Mijn add-ins**
2. Klik op **Een aangepaste add-in uploaden** (of "Sideload")
3. Kies het bestand `manifest.xml`
4. De knop "Hoekradius" verschijnt in het tabblad **Start**

### Mac
1. Ga naar **Invoegen → Add-ins**
2. Klik **Mijn invoegtoepassingen**
3. Klik het ⚙️-tandwiel → **Een aangepaste invoegtoepassing uploaden**
4. Kies `manifest.xml`

### Organisatie-breed (via IT/admin)
Upload `manifest.xml` naar het Microsoft 365 Admin Center onder
**Instellingen → Geïntegreerde apps → Apps uploaden**.
Dan is de add-in beschikbaar voor alle gebruikers in de organisatie.

---

## Gebruik

1. Open PowerPoint
2. Klik op **Hoekradius** in het tabblad Start
3. Kies het bereik:
   - **Selectie** — alleen geselecteerde vormen
   - **Huidige slide** — alle afgeronde rechthoeken op de actieve slide
   - **Alle slides** — hele presentatie
4. Stel de radius in met de slider of kies een preset
5. Klik **Toepassen**

De knop **Huidige radius uitlezen** leest de waarde van de geselecteerde vorm uit
en stelt de slider automatisch in op die waarde.

---

## Radiuswaarden

| Waarde | Beschrijving |
|--------|-------------|
| 0.00   | Volledig rechthoekig (geen afronding) |
| 0.10   | Licht afgerond |
| 0.25   | Normaal afgerond |
| 0.50   | Maximaal afgerond (halve cirkel aan de korte zijde) |

---

## Technische details

- Gebruikt de **Office JS API** (`PowerPoint.run`) — geen VBA, geen macro's
- Het `.pptx`-bestand zelf bevat geen code → AutoSave en SharePoint werken normaal
- Vereist Microsoft 365 (desktop of web); werkt niet met eeuwigdurende licenties
  van Office 2016/2019 zonder internet
- De `adjustmentHandles` API is beschikbaar vanaf Office JS API versie 1.5

---

## Problemen oplossen

**"adjustmentHandles wordt niet ondersteund"**
→ Update PowerPoint naar de nieuwste versie via Microsoft 365.

**Knop verschijnt niet in het lint**
→ Controleer of de URL in manifest.xml bereikbaar is (HTTPS vereist, geen HTTP).

**Mac: add-in laadt niet**
→ Zorg dat de URL in manifest.xml een geldig SSL-certificaat heeft.
→ GitHub Pages heeft altijd HTTPS — aanbevolen voor Mac-compatibiliteit.
