/* taskpane.js — Hoekradius Add-in voor PowerPoint
   Uniforme modus  : past de bestaande roundedRectangle handle aan (behoudt vorm)
   Per-hoek modus  : converteert naar freeform path via addGeometricShape + XML-injectie
*/

'use strict';

// ── State ─────────────────────────────────────────────────────────────────────
let currentMode  = 'uniform';
let currentScope = 'selection';
let uniformVal   = 0.10;
let linked       = false;

const corners = { tl: 0.10, tr: 0.10, bl: 0.10, br: 0.10 };

const scopeHints = {
  selection: 'Pas alleen de geselecteerde vormen aan.',
  slide:     'Pas alle afgeronde rechthoeken op de huidige slide aan.',
  all:       'Pas alle afgeronde rechthoeken in de volledige presentatie aan.'
};

// ── Office init ───────────────────────────────────────────────────────────────
Office.onReady(() => {
  updatePreview();
});

// ── Mode / Scope ──────────────────────────────────────────────────────────────
function setMode(mode) {
  currentMode = mode;
  document.getElementById('mode-uniform').classList.toggle('active', mode === 'uniform');
  document.getElementById('mode-percorner').classList.toggle('active', mode === 'percorner');
  document.querySelectorAll('[data-mode]').forEach(el => {
    el.classList.toggle('visible', el.dataset.mode === mode);
  });
}

function setScope(scope) {
  currentScope = scope;
  ['selection','slide','all'].forEach(s =>
    document.getElementById('scope-' + s).classList.toggle('active', s === scope)
  );
  document.getElementById('scope-hint').textContent = scopeHints[scope];
}

// ── Uniform slider ────────────────────────────────────────────────────────────
function onSlider(val) {
  uniformVal = parseInt(val) / 100;
  document.getElementById('radius-display').textContent = uniformVal.toFixed(2);
  updatePresetHighlight(parseInt(val));
}

function setPreset(val) {
  document.getElementById('radius-slider').value = val;
  onSlider(val);
}

function updatePresetHighlight(val) {
  [0, 10, 25, 50].forEach((p, i) => {
    const btn = document.querySelectorAll('.preset-btn')[i];
    const isActive = p === val;
    btn.classList.toggle('active', isActive);
    btn.querySelector('rect').setAttribute('stroke', isActive ? '#0f6e56' : '#888');
  });
}

// ── Per-corner sliders ────────────────────────────────────────────────────────
function onCorner(corner, rawVal) {
  const val = parseInt(rawVal) / 100;
  if (linked) {
    ['tl','tr','bl','br'].forEach(c => {
      corners[c] = val;
      document.getElementById('c-' + c).value = rawVal;
      document.getElementById('v-' + c).textContent = val.toFixed(2);
    });
  } else {
    corners[corner] = val;
    document.getElementById('v-' + corner).textContent = val.toFixed(2);
  }
  updatePreview();
}

function toggleLink() {
  linked = !linked;
  const btn = document.getElementById('link-btn');
  btn.classList.toggle('active', linked);
  document.getElementById('link-label').textContent =
    linked ? 'Hoeken zijn gekoppeld' : 'Alle hoeken koppelen';

  if (linked) {
    // Synchroniseer alle hoeken op de waarde van tl
    const ref = corners.tl;
    ['tl','tr','bl','br'].forEach(c => {
      corners[c] = ref;
      document.getElementById('c-' + c).value = Math.round(ref * 100);
      document.getElementById('v-' + c).textContent = ref.toFixed(2);
    });
    updatePreview();
  }
}

// ── Live SVG preview ──────────────────────────────────────────────────────────
function updatePreview() {
  const W = 156, H = 92;
  const pad = 4;
  const x = pad, y = pad, w = W - pad*2, h = H - pad*2;

  // Radius in pixels: de waarde (0–0.5) * kortste zijde
  const minSide = Math.min(w, h);
  const rTL = corners.tl * minSide;
  const rTR = corners.tr * minSide;
  const rBR = corners.br * minSide;
  const rBL = corners.bl * minSide;

  const d = roundedRectPath(x, y, w, h, rTL, rTR, rBR, rBL);
  document.getElementById('preview-path').setAttribute('d', d);
}

// ── Bouw een SVG path voor een rechthoek met individuele hoekradii ─────────────
// Volgorde: tl=links-boven, tr=rechts-boven, br=rechts-onder, bl=links-onder
function roundedRectPath(x, y, w, h, rTL, rTR, rBR, rBL) {
  return [
    `M ${x + rTL} ${y}`,
    `L ${x + w - rTR} ${y}`,
    rTR > 0 ? `Q ${x+w} ${y} ${x+w} ${y+rTR}` : `L ${x+w} ${y}`,
    `L ${x + w} ${y + h - rBR}`,
    rBR > 0 ? `Q ${x+w} ${y+h} ${x+w-rBR} ${y+h}` : `L ${x+w} ${y+h}`,
    `L ${x + rBL} ${y + h}`,
    rBL > 0 ? `Q ${x} ${y+h} ${x} ${y+h-rBL}` : `L ${x} ${y+h}`,
    `L ${x} ${y + rTL}`,
    rTL > 0 ? `Q ${x} ${y} ${x+rTL} ${y}` : `L ${x} ${y}`,
    'Z'
  ].join(' ');
}

// ── Status ────────────────────────────────────────────────────────────────────
function showStatus(type, message) {
  const el = document.getElementById('status');
  el.className = type;
  el.innerHTML = `<span style="font-weight:700">${{success:'✓',error:'✕',info:'ℹ'}[type]||''}</span><span>${message}</span>`;
  el.style.display = 'flex';
  if (type !== 'error') setTimeout(() => { el.style.display = 'none'; }, 4500);
}

// ── Toepassen ─────────────────────────────────────────────────────────────────
async function applyRadius() {
  const btn = document.getElementById('btn-apply');
  btn.disabled = true;
  btn.innerHTML = '<span>Bezig…</span>';

  try {
    if (currentMode === 'uniform') {
      await applyUniform();
    } else {
      await applyPerCorner();
    }
  } catch (err) {
    console.error(err);
    showStatus('error', err.message || String(err));
  }

  btn.disabled = false;
  btn.innerHTML = `<svg width="15" height="15" viewBox="0 0 15 15" fill="none">
    <path d="M2 7.5L6 11.5L13 4" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
    Toepassen`;
}

// ── Uniforme modus: pas adjustmentHandle aan ──────────────────────────────────
async function applyUniform() {
  let count = 0;

  await PowerPoint.run(async (context) => {
    const shapes = await getTargetShapes(context);

    for (const shape of shapes) {
      shape.load('geometricShapeType');
      await context.sync();

      if (shape.geometricShapeType === PowerPoint.GeometricShapeType.roundedRectangle) {
        shape.geometricShape.adjustmentHandles.load('items');
        await context.sync();
        if (shape.geometricShape.adjustmentHandles.items.length > 0) {
          shape.geometricShape.adjustmentHandles.items[0].value = uniformVal;
          count++;
        }
      }
    }

    await context.sync();
  });

  if (count === 0) {
    showStatus('info', 'Geen afgeronde rechthoeken gevonden.');
  } else {
    showStatus('success', `${count} rechthoek${count !== 1 ? 'en' : ''} bijgewerkt.`);
  }
}

// ── Per-hoek modus: vervang door freeform via XML ─────────────────────────────
async function applyPerCorner() {
  let count = 0;

  await PowerPoint.run(async (context) => {
    const shapes = await getTargetShapes(context);

    for (const shape of shapes) {
      shape.load([
        'geometricShapeType',
        'left', 'top', 'width', 'height',
        'name',
        'fill/type',
        'fill/foreColor',
        'lineFormat/color',
        'lineFormat/weight',
        'lineFormat/visible'
      ].join(','));
      await context.sync();

      if (shape.geometricShapeType !== PowerPoint.GeometricShapeType.roundedRectangle) continue;

      const { left, top, width, height, name } = shape;

      // Sla opmaak op
      const fillColor   = shape.fill?.foreColor?.toString() || '#4472C4';
      const lineColor   = shape.lineFormat?.color?.toString() || '#000000';
      const lineWeight  = shape.lineFormat?.weight ?? 1;
      const lineVisible = shape.lineFormat?.visible ?? true;

      // Bereken absolute radii in EMU-equivalente verhouding
      // PowerPoint freeform gebruikt punten; width/height zijn in punten
      const minSide = Math.min(width, height);
      const rTL = corners.tl * minSide;
      const rTR = corners.tr * minSide;
      const rBR = corners.br * minSide;
      const rBL = corners.bl * minSide;

      // Verwijder de originele shape
      const slideId = shape.parentSlide?.id;
      shape.delete();
      await context.sync();

      // Voeg een nieuwe freeform toe op dezelfde slide
      // We gebruiken addGeometricShape als tijdelijke placeholder en vervangen
      // daarna de geometrie via de Open XML (OOXML) setCustomXmlPart aanpak.
      // Omdat de Office JS API geen directe freeform-constructor biedt,
      // gebruiken we de beschikbare Shapes.addGeometricShape en
      // passen daarna de XML aan.

      const slide = shape.parentSlide || context.presentation.getSelectedSlides().getItemAt(0);
      const newShape = slide.shapes.addGeometricShape(
        PowerPoint.GeometricShapeType.rectangle,
        { left, top, width, height }
      );
      newShape.load('id');
      await context.sync();

      // Stel naam en opmaak in
      newShape.name = name + ' (per hoek)';
      newShape.fill.setSolidColor(fillColor);
      if (lineVisible) {
        newShape.lineFormat.color = lineColor;
        newShape.lineFormat.weight = lineWeight;
      } else {
        newShape.lineFormat.visible = false;
      }

      // Injecteer custom geometrie via OOXML
      // Bouw een SVG-achtig path in DrawingML punten-coördinaten
      // DrawingML gebruikt EMUs (1 pt = 12700 EMU), maar addCustomGeometryPath
      // accepteert punten relatief t.o.v. de shapebox (0..width, 0..height)
      await injectCustomGeometry(newShape, width, height, rTL, rTR, rBR, rBL, context);

      count++;
    }

    await context.sync();
  });

  if (count === 0) {
    showStatus('info', 'Geen afgeronde rechthoeken gevonden.');
  } else {
    showStatus('success', `${count} vorm${count !== 1 ? 'en' : ''} omgezet naar vrije vorm met per-hoek radius.`);
  }
}

// ── Injecteer custom DrawingML geometrie ──────────────────────────────────────
async function injectCustomGeometry(shape, W, H, rTL, rTR, rBR, rBL, context) {
  // DrawingML custGeom path werkt in EMU (1 pt = 12700 EMU)
  const emu = (pt) => Math.round(pt * 12700);

  const eW  = emu(W),  eH  = emu(H);
  const eTL = emu(rTL), eTR = emu(rTR);
  const eBR = emu(rBR), eBL = emu(rBL);

  // Bouw het DrawingML <a:path> element
  // moveto → lnTo → arcTo per hoek
  const pathCommands = [
    `<a:moveTo><a:pt x="${eTL}" y="0"/></a:moveTo>`,

    // Bovenkant
    `<a:lnTo><a:pt x="${eW - eTR}" y="0"/></a:lnTo>`,
    // Rechts-boven
    eTR > 0
      ? `<a:arcTo wR="${eTR}" hR="${eTR}" stAng="16200000" swAng="5400000"/>`
      : '',

    // Rechterkant
    `<a:lnTo><a:pt x="${eW}" y="${eH - eBR}"/></a:lnTo>`,
    // Rechts-onder
    eBR > 0
      ? `<a:arcTo wR="${eBR}" hR="${eBR}" stAng="0" swAng="5400000"/>`
      : '',

    // Onderkant
    `<a:lnTo><a:pt x="${eBL}" y="${eH}"/></a:lnTo>`,
    // Links-onder
    eBL > 0
      ? `<a:arcTo wR="${eBL}" hR="${eBL}" stAng="5400000" swAng="5400000"/>`
      : '',

    // Linkerkant
    `<a:lnTo><a:pt x="0" y="${eTL}"/></a:lnTo>`,
    // Links-boven
    eTL > 0
      ? `<a:arcTo wR="${eTL}" hR="${eTL}" stAng="10800000" swAng="5400000"/>`
      : '',

    `<a:close/>`
  ].join('');

  const custGeomXml = `
    <a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:avLst/>
      <a:gdLst/>
      <a:ahLst/>
      <a:cxnLst/>
      <a:rect l="0" t="0" r="${eW}" b="${eH}"/>
      <a:pathLst>
        <a:path w="${eW}" h="${eH}">
          ${pathCommands}
        </a:path>
      </a:pathLst>
    </a:custGeom>`;

  // Gebruik setCustomXmlParts om de geometrie in te injecteren
  // Dit vereist dat we de shape OOXML ophalen, aanpassen en terugzetten
  try {
    const ooxml = shape.getOoxmlAsync ? await shape.getOoxmlAsync() : null;
    if (ooxml) {
      const updated = ooxml.replace(
        /<a:prstGeom[^>]*>[\s\S]*?<\/a:prstGeom>/,
        custGeomXml.trim()
      );
      await shape.setOoxmlAsync(updated);
    }
  } catch (e) {
    // getOoxmlAsync/setOoxmlAsync zijn preview-API's — fallback: gebruik de
    // shapes.addSvg methode als alternatief (beschikbaar in nieuwere builds)
    console.warn('OOXML aanpassing niet beschikbaar, gebruik SVG-fallback:', e);
    await injectViaSvgFallback(shape, W, H, rTL, rTR, rBR, rBL, context);
  }
}

// ── SVG-fallback: voeg een transparante SVG-overlay toe ───────────────────────
// Als OOXML niet beschikbaar is, voegen we een SVG shape toe die er hetzelfde
// uitziet. De originele shape wordt dan onzichtbaar gemaakt.
async function injectViaSvgFallback(shape, W, H, rTL, rTR, rBR, rBL, context) {
  // Maak de placeholder onzichtbaar
  shape.fill.setTransparency(1.0);
  shape.lineFormat.visible = false;

  // Bouw SVG string
  const svgPath = roundedRectPath(2, 2, W - 4, H - 4, rTL, rTR, rBR, rBL);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <path d="${svgPath}" fill="#4472C4" stroke="#000000" stroke-width="1"/>
  </svg>`;

  const slide = shape.parentSlide;
  slide.shapes.addSvgImage(svg, {
    left: shape.left,
    top: shape.top,
    width: W,
    height: H
  });

  await context.sync();
}

// ── Haal doelvormen op op basis van scope ─────────────────────────────────────
async function getTargetShapes(context) {
  const shapes = [];

  if (currentScope === 'selection') {
    const sel = context.presentation.getSelectedShapes();
    sel.load('items');
    await context.sync();
    shapes.push(...sel.items);

  } else if (currentScope === 'slide') {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    slide.shapes.load('items');
    await context.sync();
    shapes.push(...slide.shapes.items);

  } else {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    for (const slide of slides.items) {
      slide.shapes.load('items');
      await context.sync();
      shapes.push(...slide.shapes.items);
    }
  }

  return shapes;
}

// ── Huidige radius uitlezen ───────────────────────────────────────────────────
async function readCurrentRadius() {
  try {
    await PowerPoint.run(async (context) => {
      const sel = context.presentation.getSelectedShapes();
      sel.load('items');
      await context.sync();

      for (const shape of sel.items) {
        shape.load('geometricShapeType');
        await context.sync();

        if (shape.geometricShapeType === PowerPoint.GeometricShapeType.roundedRectangle) {
          shape.geometricShape.adjustmentHandles.load('items');
          await context.sync();

          if (shape.geometricShape.adjustmentHandles.items.length > 0) {
            const val = shape.geometricShape.adjustmentHandles.items[0].value;
            const pct = Math.round(val * 100);

            // Stel uniforme slider in
            document.getElementById('radius-slider').value = pct;
            onSlider(pct);

            // Stel ook per-hoek sliders in
            ['tl','tr','bl','br'].forEach(c => {
              corners[c] = val;
              document.getElementById('c-' + c).value = pct;
              document.getElementById('v-' + c).textContent = val.toFixed(2);
            });
            updatePreview();

            showStatus('info', `Huidige radius: ${val.toFixed(2)} — sliders ingesteld.`);
            return;
          }
        }
      }

      showStatus('info', 'Selecteer een afgeronde rechthoek om de radius uit te lezen.');
    });
  } catch (err) {
    showStatus('error', 'Selecteer eerst een afgeronde rechthoek.');
  }
}

// ── Gedeelde hulpfunctie: roundedRectPath (ook gebruikt voor preview) ─────────
function roundedRectPath(x, y, w, h, rTL, rTR, rBR, rBL) {
  return [
    `M ${x + rTL} ${y}`,
    `L ${x + w - rTR} ${y}`,
    rTR > 0 ? `Q ${x+w} ${y} ${x+w} ${y+rTR}`       : `L ${x+w} ${y}`,
    `L ${x + w} ${y + h - rBR}`,
    rBR > 0 ? `Q ${x+w} ${y+h} ${x+w-rBR} ${y+h}`   : `L ${x+w} ${y+h}`,
    `L ${x + rBL} ${y + h}`,
    rBL > 0 ? `Q ${x} ${y+h} ${x} ${y+h-rBL}`       : `L ${x} ${y+h}`,
    `L ${x} ${y + rTL}`,
    rTL > 0 ? `Q ${x} ${y} ${x+rTL} ${y}`            : `L ${x} ${y}`,
    'Z'
  ].join(' ');
}
