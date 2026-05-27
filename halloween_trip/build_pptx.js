'use strict';
const PptxGenJS = require('pptxgenjs');
const path = require('path');

const IMG = __dirname;
const OUT = path.join(__dirname, '..', 'halloween_trip_2025.pptx');

// ── Palette ──────────────────────────────────────────────────────────────────
const C = {
  bg:      '1A1A2E', bgCard:  '16213E', bgDeep:  '0F3460',
  orange:  'FF8C00', amber:   'FFC107', coral:   'F96167',
  white:   'FFFFFF', offWhite:'E8E8E8', gray:    '9CA3AF',
  dimGray: '6B7280', green:   '22C55E', gold:    'FFD700',
  black:   '000000',
};

// ── Helpers ───────────────────────────────────────────────────────────────────
function rect(slide, x, y, w, h, fill, opts = {}) {
  slide.addShape(PptxGenJS.ShapeType ? PptxGenJS.ShapeType.RECT : 'rect', {
    x, y, w, h, fill: { color: fill }, line: { color: fill }, ...opts,
  });
}

function sectionHeader(slide, title, accent = C.orange) {
  // Top orange bar
  slide.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: accent }, line: { color: accent } });
  // Header bg
  slide.addShape('rect', { x: 0, y: 0.08, w: '100%', h: 0.72, fill: { color: C.bgCard }, line: { color: C.bgCard } });
  // Title text
  slide.addText(title, {
    x: 0.4, y: 0.08, w: 12.5, h: 0.72,
    fontSize: 22, fontFace: 'Trebuchet MS', bold: true, color: C.white, valign: 'middle',
  });
  // Bottom accent line
  slide.addShape('rect', { x: 0, y: 0.8, w: '100%', h: 0.03, fill: { color: accent }, line: { color: accent } });
}

function bottomBar(slide) {
  slide.addShape('rect', { x: 0, y: 7.42, w: '100%', h: 0.08, fill: { color: C.orange }, line: { color: C.orange } });
}

function statBox(slide, x, y, w, h, label, value, accent = C.orange) {
  slide.addShape('rect', { x, y, w, h, fill: { color: C.bgDeep }, line: { color: accent, pt: 1.5 }, rounding: 0.04 });
  slide.addText(label, { x, y: y + 0.05, w, h: 0.35, align: 'center', fontSize: 10, fontFace: 'Calibri', color: C.gray });
  slide.addText(value, { x, y: y + 0.38, w, h: h - 0.45, align: 'center', fontSize: 16, fontFace: 'Trebuchet MS', bold: true, color: accent });
}

function dayCard(slide, x, y, w, h, day, date, items, accent = C.orange) {
  slide.addShape('rect', { x, y, w, h, fill: { color: C.bgCard }, line: { color: accent, pt: 1 }, rounding: 0.04 });
  // Day header
  slide.addShape('rect', { x, y, w, h: 0.42, fill: { color: C.bgDeep }, line: { color: C.bgDeep }, rounding: 0.04 });
  slide.addText(day, { x: x + 0.08, y: y + 0.02, w: w - 0.16, h: 0.20, fontSize: 11, fontFace: 'Trebuchet MS', bold: true, color: accent });
  slide.addText(date, { x: x + 0.08, y: y + 0.22, w: w - 0.16, h: 0.18, fontSize: 9, fontFace: 'Calibri', color: C.gray });
  // Items
  let ty = y + 0.48;
  items.forEach(item => {
    slide.addText(item, {
      x: x + 0.1, y: ty, w: w - 0.2, h: 0.28,
      fontSize: 9, fontFace: 'Calibri', color: C.offWhite, valign: 'top',
      wrap: true,
    });
    ty += 0.29;
  });
}

// ── Slide 1: Title ────────────────────────────────────────────────────────────
function slide1(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  s.addImage({ path: path.join(IMG, 'title_bg.png'), x: 0, y: 0, w: '100%', h: '100%' });
  s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 45 }, line: { color: '000000' } });

  // Top bar
  s.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.1, fill: { color: C.orange }, line: { color: C.orange } });

  // Pumpkin
  s.addText('🎃', { x: 0, y: 0.7, w: '100%', h: 1.0, align: 'center', fontSize: 72, fontFace: 'Segoe UI Emoji' });

  // Main title
  s.addText('HALLOWEEN TRIP 2025', {
    x: 0.5, y: 1.55, w: 12.33, h: 1.2,
    align: 'center', fontSize: 62, fontFace: 'Trebuchet MS', bold: true, color: C.orange,
    glow: { size: 12, opacity: 0.5, color: C.orange },
  });

  // Tagline
  s.addText('Oct 30 – Nov 2  ·  4 Guys  ·  European Capital  ·  Halloween Mayhem', {
    x: 0.5, y: 2.85, w: 12.33, h: 0.55,
    align: 'center', fontSize: 20, fontFace: 'Trebuchet MS', color: C.offWhite,
  });

  // Past trips box
  s.addShape('rect', { x: 2.0, y: 3.7, w: 9.33, h: 0.6, fill: { color: C.bgCard, transparency: 20 }, line: { color: C.orange, pt: 1 }, rounding: 0.05 });
  s.addText('Already conquered:   ✓ Riga   ✓ Athens   ✓ Tirana   ✓ Belgrade', {
    x: 2.0, y: 3.7, w: 9.33, h: 0.6,
    align: 'center', fontSize: 15, fontFace: 'Calibri', color: C.gray, valign: 'middle',
  });

  // Next stop teaser
  s.addText('TWO NEW DESTINATIONS — LET\'S VOTE', {
    x: 0.5, y: 4.6, w: 12.33, h: 0.65,
    align: 'center', fontSize: 26, fontFace: 'Trebuchet MS', bold: true, color: C.coral,
  });

  // Destination pill teasers
  s.addShape('rect', { x: 1.5, y: 5.45, w: 4.3, h: 0.7, fill: { color: C.bgDeep }, line: { color: C.orange, pt: 2 }, rounding: 0.06 });
  s.addText('🇵🇹  LISBON', { x: 1.5, y: 5.45, w: 4.3, h: 0.7, align: 'center', fontSize: 22, fontFace: 'Trebuchet MS', bold: true, color: C.white, valign: 'middle' });

  s.addShape('rect', { x: 7.53, y: 5.45, w: 4.3, h: 0.7, fill: { color: C.bgDeep }, line: { color: C.amber, pt: 2 }, rounding: 0.06 });
  s.addText('🇭🇺  BUDAPEST', { x: 7.53, y: 5.45, w: 4.3, h: 0.7, align: 'center', fontSize: 22, fontFace: 'Trebuchet MS', bold: true, color: C.white, valign: 'middle' });

  s.addText('VS', { x: 5.8, y: 5.45, w: 1.73, h: 0.7, align: 'center', fontSize: 20, fontFace: 'Trebuchet MS', bold: true, color: C.coral, valign: 'middle' });

  // Subtitle bottom
  s.addText('Flights from DUS · FMO · DTM', {
    x: 0, y: 6.5, w: '100%', h: 0.4,
    align: 'center', fontSize: 13, fontFace: 'Calibri', color: C.dimGray, italic: true,
  });

  bottomBar(s);
}

// ── Slide 2: Lisbon Overview ───────────────────────────────────────────────────
function slide2(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };

  sectionHeader(s, '🇵🇹  DESTINATION 1 — LISBON, PORTUGAL', C.orange);

  // Left: hero image
  s.addImage({ path: path.join(IMG, 'lisbon_hero.png'), x: 0, y: 0.83, w: 7.0, h: 5.4 });
  // Gradient fade right edge
  s.addShape('rect', { x: 5.8, y: 0.83, w: 1.2, h: 5.4, fill: { color: C.bg }, line: { color: C.bg }, transparency: 20 });

  // Right side content
  const rx = 7.2;

  // City tagline
  s.addText('Nightlife Capital of the Atlantic', {
    x: rx, y: 0.95, w: 5.9, h: 0.45,
    fontSize: 16, fontFace: 'Trebuchet MS', bold: true, color: C.orange,
  });
  s.addText('Bairro Alto  ·  Pink Street  ·  Alfama Fado  ·  Castelo de São Jorge', {
    x: rx, y: 1.38, w: 5.9, h: 0.32,
    fontSize: 10, fontFace: 'Calibri', color: C.gray, italic: true,
  });

  // Stats grid: 2x3
  const stats = [
    ['✈ Flight', '3h 10m\nfrom DUS'],
    ['🍺 Beer', '€2 – €3\navg per pint'],
    ['🌡 Weather', '17–22°C\nlate October'],
    ['💰 Trip est.', '~€620\nper person'],
    ['🏙 Nightlife', '⭐⭐⭐⭐⭐\nWorld-class'],
    ['📍 Fly from', 'DUS only\n(Eurowings)'],
  ];

  let si = 0;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 3; col++) {
      const bx = rx + col * 1.98;
      const by = 1.85 + row * 1.35;
      statBox(s, bx, by, 1.85, 1.2, stats[si][0], stats[si][1], C.orange);
      si++;
    }
  }

  // Description
  s.addShape('rect', { x: rx, y: 4.65, w: 5.9, h: 1.55, fill: { color: C.bgCard }, line: { color: C.bgDeep, pt: 1 }, rounding: 0.04 });
  s.addText(
    'Lisbon delivers warm Atlantic charm with some of Europe\'s cheapest, most vibrant nightlife. ' +
    'Bairro Alto bursts with 200+ bars from midnight. Pink Street clubs rage until dawn. ' +
    'The city fully embraces Halloween — pub crawls, costumed crowds, and massive club events. ' +
    'Late-October weather means terrace drinking in a t-shirt.',
    {
      x: rx + 0.15, y: 4.75, w: 5.6, h: 1.35,
      fontSize: 10.5, fontFace: 'Calibri', color: C.offWhite, wrap: true, valign: 'top',
    }
  );

  bottomBar(s);
}

// ── Slide 3: Lisbon Flights ───────────────────────────────────────────────────
function slide3(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇵🇹  LISBON — FLIGHTS', C.orange);

  // Subsection label
  s.addText('OUTBOUND: Germany → Lisbon (LIS)', {
    x: 0.4, y: 1.0, w: 12.5, h: 0.35,
    fontSize: 13, fontFace: 'Trebuchet MS', bold: true, color: C.orange,
  });

  // Outbound flights table
  const headerOpts = { bold: true, color: C.orange, fontSize: 11, fontFace: 'Trebuchet MS', fill: { color: C.bgDeep }, align: 'center' };
  const rowOpts    = { color: C.offWhite, fontSize: 10, fontFace: 'Calibri', fill: { color: C.bgCard }, align: 'center' };
  const altRow     = { ...rowOpts, fill: { color: C.bg } };

  s.addTable([
    [
      { text: 'Option', options: headerOpts },
      { text: 'Flight', options: headerOpts },
      { text: 'Route & Date', options: headerOpts },
      { text: 'Departure', options: headerOpts },
      { text: 'Arrival', options: headerOpts },
      { text: 'Duration', options: headerOpts },
      { text: 'Price/Person', options: headerOpts },
      { text: 'Total (×4)', options: { ...headerOpts, color: C.amber } },
    ],
    [
      { text: '⭐ Best', options: { ...rowOpts, color: C.green, bold: true } },
      { text: 'EW9602\nEurowings', options: rowOpts },
      { text: 'DUS → LIS\nOct 30 morning', options: rowOpts },
      { text: '10:15', options: rowOpts },
      { text: '12:20', options: rowOpts },
      { text: '3h 05m', options: rowOpts },
      { text: '~€100', options: rowOpts },
      { text: '~€400', options: { ...rowOpts, color: C.amber, bold: true } },
    ],
    [
      { text: 'Early Bird', options: altRow },
      { text: 'EW9602\nEurowings', options: altRow },
      { text: 'DUS → LIS\nOct 30', options: altRow },
      { text: '06:20', options: altRow },
      { text: '08:25', options: altRow },
      { text: '3h 05m', options: altRow },
      { text: '~€80', options: altRow },
      { text: '~€320', options: { ...altRow, color: C.green, bold: true } },
    ],
    [
      { text: 'Eve. before', options: rowOpts },
      { text: 'EW9604\nEurowings', options: rowOpts },
      { text: 'DUS → LIS\nOct 29 evening', options: rowOpts },
      { text: '~18:30', options: rowOpts },
      { text: '~20:30', options: rowOpts },
      { text: '3h 00m', options: rowOpts },
      { text: '~€90', options: rowOpts },
      { text: '~€360', options: { ...rowOpts, color: C.amber, bold: true } },
    ],
    [
      { text: 'Premium', options: altRow },
      { text: 'TP541\nTAP Portugal', options: altRow },
      { text: 'DUS → LIS\nOct 30', options: altRow },
      { text: 'Various', options: altRow },
      { text: 'Various', options: altRow },
      { text: '3h 10m', options: altRow },
      { text: '€150–180', options: altRow },
      { text: '€600–720', options: altRow },
    ],
  ], {
    x: 0.4, y: 1.38, w: 12.5,
    colW: [1.2, 1.7, 2.1, 1.2, 1.2, 1.2, 1.5, 1.4],
    rowH: [0.38, 0.52, 0.52, 0.52, 0.52],
    border: { type: 'solid', pt: 1, color: C.bgDeep },
  });

  // Return flights
  s.addText('RETURN: Lisbon → Germany (Nov 2)', {
    x: 0.4, y: 4.25, w: 12.5, h: 0.35,
    fontSize: 13, fontFace: 'Trebuchet MS', bold: true, color: C.orange,
  });

  s.addTable([
    [
      { text: 'Flight', options: headerOpts },
      { text: 'Route & Date', options: headerOpts },
      { text: 'Departure', options: headerOpts },
      { text: 'Arrival', options: headerOpts },
      { text: 'Duration', options: headerOpts },
      { text: 'Price/Person', options: headerOpts },
      { text: 'Total (×4)', options: { ...headerOpts, color: C.amber } },
    ],
    [
      { text: 'EW9603 Eurowings', options: { ...rowOpts, color: C.green } },
      { text: 'LIS → DUS  /  Nov 2', options: rowOpts },
      { text: '18:00', options: rowOpts },
      { text: '21:50', options: rowOpts },
      { text: '3h 50m', options: rowOpts },
      { text: '~€90', options: rowOpts },
      { text: '~€360', options: { ...rowOpts, color: C.amber, bold: true } },
    ],
  ], {
    x: 0.4, y: 4.62, w: 12.5,
    colW: [2.0, 2.5, 1.5, 1.5, 1.5, 1.8, 1.7],
    rowH: [0.38, 0.52],
    border: { type: 'solid', pt: 1, color: C.bgDeep },
  });

  // Notes
  s.addShape('rect', { x: 0.4, y: 5.38, w: 12.5, h: 0.72, fill: { color: C.bgCard }, line: { color: C.bgDeep }, rounding: 0.04 });
  s.addText(
    '⚠  DUS (Düsseldorf) is the ONLY airport with direct Lisbon service — Eurowings & TAP Air Portugal operate daily non-stop routes.  ' +
    'DTM & FMO require connections (8h+ total).  •  Book via eurowings.com or kayak.com  •  Round-trip average: ~€170–190pp',
    {
      x: 0.55, y: 5.43, w: 12.2, h: 0.62,
      fontSize: 10, fontFace: 'Calibri', color: C.gray, wrap: true, valign: 'middle',
    }
  );

  // Price summary pill
  s.addShape('rect', { x: 9.0, y: 6.2, w: 3.9, h: 0.9, fill: { color: C.bgDeep }, line: { color: C.orange, pt: 2 }, rounding: 0.06 });
  s.addText('Round Trip · 4 Persons', { x: 9.0, y: 6.22, w: 3.9, h: 0.32, align: 'center', fontSize: 10, fontFace: 'Calibri', color: C.gray });
  s.addText('€640 – €760 total', { x: 9.0, y: 6.52, w: 3.9, h: 0.52, align: 'center', fontSize: 20, fontFace: 'Trebuchet MS', bold: true, color: C.amber });

  bottomBar(s);
}

// ── Slide 4: Lisbon Airbnb ────────────────────────────────────────────────────
function slide4(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇵🇹  LISBON — ACCOMMODATION', C.orange);

  // Left image
  s.addImage({ path: path.join(IMG, 'lisbon_airbnb.png'), x: 0, y: 0.83, w: 5.5, h: 3.4 });
  s.addImage({ path: path.join(IMG, 'lisbon_map.png'), x: 0, y: 4.28, w: 5.5, h: 2.3 });

  // Right details
  const rx = 5.8;

  // Listing title bar
  s.addShape('rect', { x: rx, y: 0.9, w: 7.3, h: 0.55, fill: { color: C.bgDeep }, line: { color: C.orange, pt: 1.5 }, rounding: 0.04 });
  s.addText('Bairro Alto / São Roque  —  2 Bedroom Apartment', {
    x: rx + 0.15, y: 0.9, w: 7.0, h: 0.55,
    fontSize: 14, fontFace: 'Trebuchet MS', bold: true, color: C.white, valign: 'middle',
  });

  // Spec badges
  const specs = [
    ['🛏  2 Bedrooms', C.bgDeep, C.orange],
    ['🚿  2 Bathrooms', C.bgDeep, C.orange],
    ['🏗  Balcony  ✓', C.bgDeep, C.green],
    ['👥  Sleeps 4', C.bgDeep, C.orange],
    ['📍  Central', C.bgDeep, C.amber],
    ['❄  A/C  ✓', C.bgDeep, C.orange],
  ];
  specs.forEach((spec, i) => {
    const bx = rx + (i % 3) * 2.45;
    const by = 1.56 + Math.floor(i / 3) * 0.55;
    s.addShape('rect', { x: bx, y: by, w: 2.3, h: 0.42, fill: { color: spec[1] }, line: { color: spec[2], pt: 1 }, rounding: 0.04 });
    s.addText(spec[0], { x: bx, y: by, w: 2.3, h: 0.42, align: 'center', fontSize: 11, fontFace: 'Calibri', bold: true, color: spec[2], valign: 'middle' });
  });

  // Neighbourhood box
  s.addShape('rect', { x: rx, y: 2.74, w: 7.3, h: 0.55, fill: { color: C.bgCard }, line: { color: C.bgDeep }, rounding: 0.04 });
  s.addText('📌  Bairro Alto — steps from 200+ bars, Cervejaria Ramiro, Pink Street (10 min walk)', {
    x: rx + 0.1, y: 2.74, w: 7.1, h: 0.55,
    fontSize: 10.5, fontFace: 'Calibri', color: C.gray, valign: 'middle',
  });

  // Price section
  s.addShape('rect', { x: rx, y: 3.4, w: 7.3, h: 1.5, fill: { color: C.bgCard }, line: { color: C.orange, pt: 1.5 }, rounding: 0.04 });
  s.addText('PRICING  (Oct 30 – Nov 2, 3 nights)', {
    x: rx + 0.2, y: 3.46, w: 7.0, h: 0.32,
    fontSize: 11, fontFace: 'Trebuchet MS', bold: true, color: C.orange,
  });

  const priceRows = [
    ['Per night', '~€130'],
    ['3 nights subtotal', '~€390'],
    ['Cleaning + fees', '~€60–80'],
    ['Total for 4 persons', '~€450–480'],
    ['Per person', '~€113–120'],
  ];
  priceRows.forEach((row, i) => {
    const py = 3.82 + i * 0.215;
    s.addText(row[0], { x: rx + 0.2, y: py, w: 4.0, h: 0.21, fontSize: 10, fontFace: 'Calibri', color: C.gray });
    s.addText(row[1], { x: rx + 4.5, y: py, w: 2.6, h: 0.21, align: 'right', fontSize: 10, fontFace: 'Trebuchet MS', bold: i === 3 || i === 4, color: i === 3 ? C.amber : i === 4 ? C.green : C.offWhite });
  });

  // URL box
  s.addShape('rect', { x: rx, y: 5.03, w: 7.3, h: 0.52, fill: { color: C.bgDeep }, line: { color: C.bgDeep }, rounding: 0.04 });
  s.addText('airbnb.com/rooms/26081783  (Bairro Alto / São Roque — verify availability)', {
    x: rx + 0.15, y: 5.03, w: 7.0, h: 0.52,
    fontSize: 9.5, fontFace: 'Calibri', color: C.amber, italic: true, valign: 'middle',
  });

  // Alternative option
  s.addText('Alternative: "Garrett 48" in Chiado — terrace, 2BR, ~€100–130/night → airbnb.com/rooms/1016061', {
    x: rx + 0.1, y: 5.65, w: 7.1, h: 0.32,
    fontSize: 9.5, fontFace: 'Calibri', color: C.dimGray, italic: true,
  });

  s.addText('🔑  Book immediately — Halloween weekend sells out weeks in advance!', {
    x: rx + 0.1, y: 6.05, w: 7.1, h: 0.32,
    fontSize: 9.5, fontFace: 'Calibri', bold: true, color: C.coral,
  });

  bottomBar(s);
}

// ── Slide 5: Lisbon Activities ────────────────────────────────────────────────
function slide5(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇵🇹  LISBON — DAY-BY-DAY ITINERARY', C.orange);

  const days = [
    {
      day: 'DAY 1 — THU Oct 30', date: 'Arrival Day',
      items: [
        '✈  EW9602 lands LIS 12:20 · Check into Bairro Alto',
        '🚶  SANDEMANs Free Tour: 3:30 PM @ Praça do Comércio',
        '   (2.5h, pay €10–15 tip · covers Baixa, Alfama, miradouros)',
        '🍽  Dinner: Petiscos at Tasca do Chico ~8 PM',
        '🍺  Bairro Alto bar crawl: 200+ bars, €2–3 beers',
        '   Start @ Park Bar (rooftop) → Pavilhão Chinês',
      ],
    },
    {
      day: 'DAY 2 — FRI Oct 31 🎃', date: 'HALLOWEEN NIGHT',
      items: [
        '🍳  Cooking Lisbon Market Tour 9:30 AM (4h, €110pp)',
        '   Mercado 31 de Janeiro · 3-course meal + wine included',
        '   Book: cookinglisbon.com/market-tour',
        '🕙  Halloween Haunts Pub Crawl 10 PM',
        '   Meet: Praça Luís de Camões, Chiado · €28pp',
        '   3 bars · 1h open bar · shots · VIP club entry',
        '🎶  After 1 AM: Brunch Electronik Halloween Special',
        '   Pavilhão Carlos Lopes · Chris Liebing B2B Luke Slater',
      ],
    },
    {
      day: 'DAY 3 — SAT Nov 1', date: 'Recovery & Explore',
      items: [
        '🌅  Late start — hangover pastel de nata @ Pasteis de Belém',
        '🚋  Tram 28: scenic ride through Alfama to Castelo',
        '🍺  Cerveja Musa craft brewery, Mouraria (local IPA)',
        '🌇  Miradouro da Graça for sunset views',
        '🎵  Fado dinner: Tasca do Chico or Zé da Viola',
        '🌃  Pink Street night out: Pensão Amor + Music Box',
        '   Lux Frágil riverside club (open till 6 AM)',
      ],
    },
    {
      day: 'DAY 4 — SUN Nov 2', date: 'Departure Day',
      items: [
        '☕  Morning coffee @ A Brasileira café, Chiado',
        '🏰  Quick visit: Castelo de São Jorge (optional)',
        '🛍  Souvenirs: ginjinha liqueur, tinned fish',
        '🚌  Transfer to DUS airport (allow 45 min to airport)',
        '✈  EW9603  18:00 LIS → 21:50 DUS',
      ],
    },
  ];

  // 4 cards in a 2×2 grid
  days.forEach((d, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const cx = 0.3 + col * 6.5;
    const cy = 0.98 + row * 3.15;
    dayCard(s, cx, cy, 6.2, 3.0, d.day, d.date, d.items, C.orange);
  });

  // Footer note
  s.addShape('rect', { x: 0.3, y: 7.12, w: 12.7, h: 0.22, fill: { color: C.bgCard }, line: { color: C.bgDeep } });
  s.addText(
    'Halloween Pub Crawl ticket: eventbrite.co.uk/e/1981093800888  ·  Cooking class: cookinglisbon.com/booking  ·  Free tour: neweuropetours.eu',
    { x: 0.4, y: 7.12, w: 12.5, h: 0.22, fontSize: 8, fontFace: 'Calibri', color: C.dimGray, valign: 'middle' }
  );

  bottomBar(s);
}

// ── Slide 6: Lisbon Cost Summary ──────────────────────────────────────────────
function slide6(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇵🇹  LISBON — COST SUMMARY', C.orange);

  // Table header + rows
  const hOpt  = { bold: true, color: C.orange, fontSize: 12, fontFace: 'Trebuchet MS', fill: { color: C.bgDeep }, align: 'center', valign: 'middle' };
  const rOpt  = { color: C.offWhite, fontSize: 11, fontFace: 'Calibri', fill: { color: C.bgCard }, align: 'left', valign: 'middle' };
  const rAlt  = { ...rOpt, fill: { color: C.bg } };
  const rNum  = { ...rOpt, align: 'right' };
  const rNumA = { ...rAlt, align: 'right' };
  const total = { bold: true, color: C.amber, fontSize: 13, fontFace: 'Trebuchet MS', fill: { color: C.bgDeep }, align: 'right', valign: 'middle' };
  const totalL = { ...total, align: 'left', color: C.white };

  s.addTable([
    [
      { text: 'Item', options: hOpt },
      { text: 'Details', options: hOpt },
      { text: 'Per Person', options: hOpt },
      { text: 'Total × 4', options: { ...hOpt, color: C.amber } },
    ],
    [
      { text: 'Flights (return)', options: rOpt },
      { text: 'Eurowings EW9602/9603  DUS ↔ LIS', options: rOpt },
      { text: '~€170', options: rNum },
      { text: '~€680', options: rNum },
    ],
    [
      { text: 'Accommodation', options: rAlt },
      { text: 'Airbnb Bairro Alto — 3 nights (inc. fees)', options: rAlt },
      { text: '~€115', options: rNumA },
      { text: '~€460', options: rNumA },
    ],
    [
      { text: 'Halloween Pub Crawl', options: rOpt },
      { text: 'Halloween Haunts — Chiado (Oct 31)', options: rOpt },
      { text: '€28', options: rNum },
      { text: '€112', options: rNum },
    ],
    [
      { text: 'Cooking Class', options: rAlt },
      { text: 'Cooking Lisbon — Market Tour + 3-course meal', options: rAlt },
      { text: '€110', options: rNumA },
      { text: '€440', options: rNumA },
    ],
    [
      { text: 'Free Walking Tour', options: rOpt },
      { text: 'SANDEMANs — Praça do Comércio (tip)', options: rOpt },
      { text: '~€12', options: rNum },
      { text: '~€48', options: rNum },
    ],
    [
      { text: 'Food (3 days)', options: rAlt },
      { text: 'Meals, petiscos, pastéis — avg €15–25/meal', options: rAlt },
      { text: '~€120', options: rNumA },
      { text: '~€480', options: rNumA },
    ],
    [
      { text: 'Drinks & Nightlife', options: rOpt },
      { text: 'Bairro Alto + Pink Street + clubs (3 nights)', options: rOpt },
      { text: '~€150', options: rNum },
      { text: '~€600', options: rNum },
    ],
    [
      { text: 'Misc (transport, tips)', options: rAlt },
      { text: 'Bolt rides, Uber, souvenirs', options: rAlt },
      { text: '~€30', options: rNumA },
      { text: '~€120', options: rNumA },
    ],
    [
      { text: '🏆  TOTAL ESTIMATE', options: totalL },
      { text: 'Lisbon, Oct 30 – Nov 2, 2025', options: { ...total, align: 'left', color: C.gray } },
      { text: '~€735', options: { ...total, color: C.amber } },
      { text: '~€2,940', options: { ...total, color: C.orange } },
    ],
  ], {
    x: 0.5, y: 1.02, w: 12.3,
    colW: [2.5, 5.2, 2.1, 2.1],
    rowH: [0.42, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.55],
    border: { type: 'solid', pt: 1, color: C.bgDeep },
  });

  // Budget range note
  s.addShape('rect', { x: 0.5, y: 6.52, w: 12.3, h: 0.68, fill: { color: C.bgCard }, line: { color: C.orange, pt: 1 }, rounding: 0.04 });
  s.addText(
    '💡 Budget range: €600 – €900pp depending on party intensity.  ' +
    'Beer at €2–3 and supermarket runs keep costs low.  ' +
    'Clubs (Lux Frágil, Brunch Electronik) add €20–30 each.  ' +
    'Tap water is drinkable and free in most restaurants.',
    {
      x: 0.65, y: 6.54, w: 12.0, h: 0.64,
      fontSize: 10.5, fontFace: 'Calibri', color: C.gray, wrap: true, valign: 'middle',
    }
  );

  bottomBar(s);
}

// ── MADRID ACCENT COLOR ───────────────────────────────────────────────────────
const MAD = 'E6321E'; // Spanish red

// ── Slide 7: Madrid Overview ──────────────────────────────────────────────────
function slide7(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇪🇸  DESTINATION 2 — MADRID, SPAIN', MAD);

  s.addImage({ path: path.join(IMG, 'madrid_hero.png'), x: 0, y: 0.83, w: 7.0, h: 5.4 });
  s.addShape('rect', { x: 5.8, y: 0.83, w: 1.2, h: 5.4, fill: { color: C.bg }, line: { color: C.bg }, transparency: 20 });

  const rx = 7.2;
  s.addText('The City That Never Sleeps — Literally', {
    x: rx, y: 0.95, w: 5.9, h: 0.45,
    fontSize: 16, fontFace: 'Trebuchet MS', bold: true, color: MAD,
  });
  s.addText('Malasaña  ·  Chueca  ·  Gran Vía  ·  La Latina tapas  ·  Clubs till 7 AM', {
    x: rx, y: 1.38, w: 5.9, h: 0.32,
    fontSize: 10, fontFace: 'Calibri', color: C.gray, italic: true,
  });

  const stats = [
    ['✈ Flight', '2h 30m\nfrom DUS'],
    ['🍺 Beer', '€2 – €4\ncaña or tubo'],
    ['🌡 Weather', '18–20°C ☀\nlate October'],
    ['💰 Trip est.', '~€712\nper person'],
    ['🏙 Nightlife', '⭐⭐⭐⭐⭐\nClubs till 7 AM'],
    ['📍 Fly from', 'DUS only\n(Iberia/Eurowings)'],
  ];
  let si = 0;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 3; col++) {
      statBox(s, rx + col * 1.98, 1.85 + row * 1.35, 1.85, 1.2, stats[si][0], stats[si][1], MAD);
      si++;
    }
  }

  s.addShape('rect', { x: rx, y: 4.65, w: 5.9, h: 1.55, fill: { color: C.bgCard }, line: { color: C.bgDeep, pt: 1 }, rounding: 0.04 });
  s.addText(
    'Madrid is the warmest Halloween in Europe. Oct 31 is a Friday AND Nov 1 is a Spanish public holiday — ' +
    'meaning absolutely everyone is out and clubs stay open until 7 AM. ' +
    'Malasaña\'s bohemian bar scene, free tapas at El Tigre, vermouth on tap at 150-year-old bodegas, ' +
    'and Teatro Kapital\'s 7-floor nightclub. T-shirt weather guaranteed.',
    {
      x: rx + 0.15, y: 4.75, w: 5.6, h: 1.35,
      fontSize: 10.5, fontFace: 'Calibri', color: C.offWhite, wrap: true, valign: 'top',
    }
  );

  bottomBar(s);
}

// ── Slide 8: Madrid Flights ───────────────────────────────────────────────────
function slide8(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇪🇸  MADRID — FLIGHTS', MAD);

  const hO = { bold: true, color: MAD, fontSize: 11, fontFace: 'Trebuchet MS', fill: { color: C.bgDeep }, align: 'center' };
  const rO  = { color: C.offWhite, fontSize: 10, fontFace: 'Calibri', fill: { color: C.bgCard }, align: 'center' };
  const rA  = { ...rO, fill: { color: C.bg } };

  s.addText('OUTBOUND: Germany → Madrid (MAD / Barajas)', {
    x: 0.4, y: 1.0, w: 12.5, h: 0.35,
    fontSize: 13, fontFace: 'Trebuchet MS', bold: true, color: MAD,
  });

  s.addTable([
    [
      { text: 'Option', options: hO },
      { text: 'Airline', options: hO },
      { text: 'Route & Date', options: hO },
      { text: 'Departure', options: hO },
      { text: 'Arrival', options: hO },
      { text: 'Duration', options: hO },
      { text: 'Price/Person', options: hO },
      { text: 'Total (×4)', options: { ...hO, color: C.orange } },
    ],
    [
      { text: '⭐ Best\nEve. Oct 29', options: { ...rO, color: C.green, bold: true } },
      { text: 'Iberia\n(IB3166)', options: rO },
      { text: 'DUS → MAD\nOct 29 evening', options: rO },
      { text: '~19:55', options: rO },
      { text: '~22:35', options: rO },
      { text: '2h 30m', options: rO },
      { text: '~€95', options: rO },
      { text: '~€380', options: { ...rO, color: C.orange, bold: true } },
    ],
    [
      { text: 'Early Oct 30', options: rA },
      { text: 'Iberia\n(IB3162)', options: rA },
      { text: 'DUS → MAD\nOct 30 morning', options: rA },
      { text: '~06:20', options: rA },
      { text: '~08:50', options: rA },
      { text: '2h 30m', options: rA },
      { text: '~€75', options: rA },
      { text: '~€300', options: { ...rA, color: C.green, bold: true } },
    ],
    [
      { text: 'Midday Oct 30', options: rO },
      { text: 'Eurowings\n(EW)', options: rO },
      { text: 'DUS → MAD\nOct 30 midday', options: rO },
      { text: '~13:xx', options: rO },
      { text: '~15:30', options: rO },
      { text: '2h 30m', options: rO },
      { text: '~€90', options: rO },
      { text: '~€360', options: { ...rO, color: C.amber, bold: true } },
    ],
    [
      { text: 'Premium', options: rA },
      { text: 'Vueling\n(VY)', options: rA },
      { text: 'DUS → MAD\nOct 30', options: rA },
      { text: 'Various', options: rA },
      { text: 'Various', options: rA },
      { text: '2h 30m', options: rA },
      { text: '€100–130', options: rA },
      { text: '€400–520', options: rA },
    ],
  ], {
    x: 0.4, y: 1.38, w: 12.5,
    colW: [1.4, 1.6, 2.0, 1.2, 1.2, 1.2, 1.6, 1.5],
    rowH: [0.38, 0.55, 0.52, 0.52, 0.52],
    border: { type: 'solid', pt: 1, color: C.bgDeep },
  });

  s.addText('RETURN: Madrid → Germany (Nov 2)', {
    x: 0.4, y: 4.35, w: 12.5, h: 0.35,
    fontSize: 13, fontFace: 'Trebuchet MS', bold: true, color: MAD,
  });

  s.addTable([
    [
      { text: 'Flight', options: hO },
      { text: 'Route & Date', options: hO },
      { text: 'Departure', options: hO },
      { text: 'Arrival', options: hO },
      { text: 'Duration', options: hO },
      { text: 'Price/Person', options: hO },
      { text: 'Total (×4)', options: { ...hO, color: C.orange } },
    ],
    [
      { text: 'Iberia / Eurowings  (MAD→DUS)', options: { ...rO, color: C.green } },
      { text: 'MAD → DUS  /  Nov 2', options: rO },
      { text: 'Various', options: rO },
      { text: 'Various', options: rO },
      { text: '2h 40m', options: rO },
      { text: '~€70–90', options: rO },
      { text: '~€280–360', options: { ...rO, color: C.amber, bold: true } },
    ],
  ], {
    x: 0.4, y: 4.72, w: 12.5,
    colW: [2.5, 2.0, 1.3, 1.3, 1.3, 1.8, 1.6],
    rowH: [0.38, 0.52],
    border: { type: 'solid', pt: 1, color: C.bgDeep },
  });

  s.addShape('rect', { x: 0.4, y: 5.44, w: 12.5, h: 0.65, fill: { color: C.bgCard }, line: { color: C.bgDeep }, rounding: 0.04 });
  s.addText(
    'ℹ  DUS (Düsseldorf) is the recommended airport — direct flights with Iberia (3×/day), Eurowings & Vueling.  ' +
    'DTM (Dortmund) has no direct Madrid service.  FMO requires connections.  ' +
    'Book at iberia.com · eurowings.com · kayak.com  ·  Airport to city centre: Metro Line 8 (€5) or taxi (~€35 fixed)',
    {
      x: 0.55, y: 5.48, w: 12.2, h: 0.6,
      fontSize: 10, fontFace: 'Calibri', color: C.gray, wrap: true, valign: 'middle',
    }
  );

  s.addShape('rect', { x: 8.5, y: 6.2, w: 4.4, h: 0.9, fill: { color: C.bgDeep }, line: { color: MAD, pt: 2 }, rounding: 0.06 });
  s.addText('Round Trip · 4 Persons', { x: 8.5, y: 6.22, w: 4.4, h: 0.32, align: 'center', fontSize: 10, fontFace: 'Calibri', color: C.gray });
  s.addText('€580 – €740 total', { x: 8.5, y: 6.52, w: 4.4, h: 0.52, align: 'center', fontSize: 20, fontFace: 'Trebuchet MS', bold: true, color: MAD });

  bottomBar(s);
}

// ── Slide 9: Madrid Airbnb ────────────────────────────────────────────────────
function slide9(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇪🇸  MADRID — ACCOMMODATION', MAD);

  s.addImage({ path: path.join(IMG, 'madrid_airbnb.png'), x: 0, y: 0.83, w: 5.5, h: 3.4 });
  s.addImage({ path: path.join(IMG, 'madrid_map.png'), x: 0, y: 4.28, w: 5.5, h: 2.3 });

  const rx = 5.8;

  s.addShape('rect', { x: rx, y: 0.9, w: 7.3, h: 0.55, fill: { color: C.bgDeep }, line: { color: MAD, pt: 1.5 }, rounding: 0.04 });
  s.addText('Sonder Malasaña — 2 Bedroom · Balcony · Nightlife on the Doorstep', {
    x: rx + 0.15, y: 0.9, w: 7.0, h: 0.55,
    fontSize: 14, fontFace: 'Trebuchet MS', bold: true, color: C.white, valign: 'middle',
  });

  const specs = [
    ['🛏  2 Bedrooms', C.bgDeep, MAD],
    ['🚿  2 Bathrooms', C.bgDeep, MAD],
    ['🏗  Balcony  ✓', C.bgDeep, C.green],
    ['👥  Sleeps 4', C.bgDeep, MAD],
    ['❄  A/C  ✓', C.bgDeep, MAD],
    ['🎸  Malasaña bars', C.bgDeep, C.orange],
  ];
  specs.forEach((spec, i) => {
    const bx = rx + (i % 3) * 2.45;
    const by = 1.56 + Math.floor(i / 3) * 0.55;
    s.addShape('rect', { x: bx, y: by, w: 2.3, h: 0.42, fill: { color: spec[1] }, line: { color: spec[2], pt: 1 }, rounding: 0.04 });
    s.addText(spec[0], { x: bx, y: by, w: 2.3, h: 0.42, align: 'center', fontSize: 11, fontFace: 'Calibri', bold: true, color: spec[2], valign: 'middle' });
  });

  s.addShape('rect', { x: rx, y: 2.74, w: 7.3, h: 0.55, fill: { color: C.bgCard }, line: { color: C.bgDeep }, rounding: 0.04 });
  s.addText('📌  Malasaña — 1 min to Calle de la Palma bars, 5 min to Gran Vía, 10 min to Puerta del Sol', {
    x: rx + 0.1, y: 2.74, w: 7.1, h: 0.55,
    fontSize: 10.5, fontFace: 'Calibri', color: C.gray, valign: 'middle',
  });

  s.addShape('rect', { x: rx, y: 3.4, w: 7.3, h: 1.5, fill: { color: C.bgCard }, line: { color: MAD, pt: 1.5 }, rounding: 0.04 });
  s.addText('PRICING  (Oct 30 – Nov 2, 3 nights)', {
    x: rx + 0.2, y: 3.46, w: 7.0, h: 0.32,
    fontSize: 11, fontFace: 'Trebuchet MS', bold: true, color: MAD,
  });

  const priceRows = [
    ['Per night (Sonder managed)', '~€150'],
    ['3 nights subtotal', '~€450'],
    ['Cleaning + fees', '~€60–80'],
    ['Total for 4 persons', '~€510–530'],
    ['Per person', '~€128–133'],
  ];
  priceRows.forEach((row, i) => {
    const py = 3.82 + i * 0.215;
    s.addText(row[0], { x: rx + 0.2, y: py, w: 4.0, h: 0.21, fontSize: 10, fontFace: 'Calibri', color: C.gray });
    s.addText(row[1], { x: rx + 4.5, y: py, w: 2.6, h: 0.21, align: 'right', fontSize: 10, fontFace: 'Trebuchet MS', bold: i === 3 || i === 4, color: i === 3 ? MAD : i === 4 ? C.green : C.offWhite });
  });

  s.addShape('rect', { x: rx, y: 5.03, w: 7.3, h: 0.52, fill: { color: C.bgDeep }, line: { color: C.bgDeep }, rounding: 0.04 });
  s.addText('airbnb.com/rooms/47308216  (Sonder Malasaña — professional managed, 24/7 support)', {
    x: rx + 0.15, y: 5.03, w: 7.0, h: 0.52,
    fontSize: 9.5, fontFace: 'Calibri', color: MAD, italic: true, valign: 'middle',
  });

  s.addText('Budget alt: 2BR central Madrid ~€80–110/night → ~€310 total → ~€78pp total', {
    x: rx + 0.1, y: 5.65, w: 7.1, h: 0.32,
    fontSize: 9.5, fontFace: 'Calibri', color: C.dimGray, italic: true,
  });

  s.addText('🔑  Book immediately — Halloween weekend Madrid fills up completely!', {
    x: rx + 0.1, y: 6.05, w: 7.1, h: 0.32,
    fontSize: 9.5, fontFace: 'Calibri', bold: true, color: C.coral,
  });

  bottomBar(s);
}

// ── Slide 10: Madrid Activities ───────────────────────────────────────────────
function slide10(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇪🇸  MADRID — DAY-BY-DAY ITINERARY', MAD);

  const days = [
    {
      day: 'DAY 1 — WED Oct 29', date: 'Arrival (Evening)',
      items: [
        '✈  Iberia ~22:35 lands MAD from DUS',
        '🚇  Metro Line 8 to centre (~€5, 20 min)',
        '🏠  Check in Malasaña apartment',
        '🍷  Bodega de la Ardosa: vermouth on tap since 1892',
        '🍺  Malasaña bar crawl: Calle de la Palma, Tupperware',
        '   Beer: Mahou 5 Estrellas ~€3 · La Vía Láctea till 3 AM',
      ],
    },
    {
      day: 'DAY 2 — THU Oct 30', date: 'Walking Tour + Cooking',
      items: [
        '🚶  SANDEMANs Free Tour 11:30 AM @ Plaza Mayor',
        '   2.5h · Royal Palace, Sol, La Latina, historic Madrid',
        '🥘  Lunch tapas on Calle de la Cava Baja (La Latina)',
        '🍳  Cooking Point Tapas Class 17:30–21:30 (€85pp)',
        '   7 tapas recipes + sangría · Calle de Moratín, 11',
        '   Book: cookingpoint.es  · Group max 12',
        '🌃  After: Chueca bar scene + El Tigre free tapas',
      ],
    },
    {
      day: 'DAY 3 — FRI Oct 31 🎃', date: 'HALLOWEEN NIGHT',
      items: [
        '🌞  Sleep in — Madrid nights go till 7 AM',
        '🏛  Afternoon: Prado Museum or Retiro Park stroll',
        '🎃  Halloween Pub Crawl 22:00 · €20pp',
        '   Meet: Mulberry\'s Bar, Calle de Núñez de Arce 9',
        '   4 bars · tequila shots · €100 costume contest',
        '🎶  After 1 AM: Teatro Kapital — 7 floors, till 6 AM',
        '   Calle Atocha 125 · €21pp entry · most central club',
        '   OR Fabrik Halloween party (20 min Uber, bigger)',
      ],
    },
    {
      day: 'DAY 4 — SAT Nov 1 (hol) + Nov 2', date: 'Public Holiday + Depart',
      items: [
        '🌞  Nov 1 = ALL SAINTS DAY (public holiday in Spain)',
        '☕  Brunch: Churros con Chocolate at Chocolatería San Ginés',
        '🏺  Mercado de San Miguel: pintxos + local wines',
        '🍺  Afternoon vermut: La Hora del Vermut (~noon–3 PM)',
        '🌃  Final night: Malasaña or Lavapiés late bars',
        '',
        '✈  NOV 2: DUS return flight (check Iberia/Eurowings)',
        '   Metro Line 8 to airport (~€5, allow 45 min)',
      ],
    },
  ];

  days.forEach((d, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    dayCard(s, 0.3 + col * 6.5, 0.98 + row * 3.15, 6.2, 3.0, d.day, d.date, d.items, MAD);
  });

  s.addShape('rect', { x: 0.3, y: 7.12, w: 12.7, h: 0.22, fill: { color: C.bgCard }, line: { color: C.bgDeep } });
  s.addText(
    'Halloween Pub Crawl: eventbrite.es/e/1037812363717  ·  Cooking: cookingpoint.es  ·  Free tour: neweuropetours.eu/madrid  ·  Club: teatrokapital.com',
    { x: 0.4, y: 7.12, w: 12.5, h: 0.22, fontSize: 8, fontFace: 'Calibri', color: C.dimGray, valign: 'middle' }
  );

  bottomBar(s);
}

// ── Slide 11: Madrid Cost Summary ─────────────────────────────────────────────
function slide11(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🇪🇸  MADRID — COST SUMMARY', MAD);

  const hO  = { bold: true, color: MAD, fontSize: 12, fontFace: 'Trebuchet MS', fill: { color: C.bgDeep }, align: 'center', valign: 'middle' };
  const rO  = { color: C.offWhite, fontSize: 11, fontFace: 'Calibri', fill: { color: C.bgCard }, align: 'left', valign: 'middle' };
  const rA  = { ...rO, fill: { color: C.bg } };
  const rNm = { ...rO, align: 'right' };
  const rNA = { ...rA, align: 'right' };
  const tot = { bold: true, color: MAD, fontSize: 13, fontFace: 'Trebuchet MS', fill: { color: C.bgDeep }, align: 'right', valign: 'middle' };
  const totL = { ...tot, align: 'left', color: C.white };

  s.addTable([
    [
      { text: 'Item', options: hO },
      { text: 'Details', options: hO },
      { text: 'Per Person', options: hO },
      { text: 'Total × 4', options: { ...hO, color: C.orange } },
    ],
    [
      { text: 'Flights (return)', options: rO },
      { text: 'Iberia / Eurowings  DUS ↔ MAD', options: rO },
      { text: '~€175', options: rNm },
      { text: '~€700', options: rNm },
    ],
    [
      { text: 'Accommodation', options: rA },
      { text: 'Airbnb Malasaña — 3 nights (inc. fees)', options: rA },
      { text: '~€130', options: rNA },
      { text: '~€520', options: rNA },
    ],
    [
      { text: 'Halloween Pub Crawl', options: rO },
      { text: 'Bar Crawl Madrid — Puerta del Sol (Oct 31)', options: rO },
      { text: '€20', options: rNm },
      { text: '€80', options: rNm },
    ],
    [
      { text: 'Cooking Class', options: rA },
      { text: 'Cooking Point — 7 tapas + sangría (4h)', options: rA },
      { text: '€85', options: rNA },
      { text: '€340', options: rNA },
    ],
    [
      { text: 'Free Walking Tour', options: rO },
      { text: 'SANDEMANs — Plaza Mayor (tip)', options: rO },
      { text: '~€12', options: rNm },
      { text: '~€48', options: rNm },
    ],
    [
      { text: 'Food (3 days)', options: rA },
      { text: 'Tapas, churros, market food — avg €15–20/meal', options: rA },
      { text: '~€115', options: rNA },
      { text: '~€460', options: rNA },
    ],
    [
      { text: 'Drinks & Nightlife', options: rO },
      { text: 'Malasaña + Chueca + Halloween club (3 nights)', options: rO },
      { text: '~€140', options: rNm },
      { text: '~€560', options: rNm },
    ],
    [
      { text: 'Misc (transport, tips)', options: rA },
      { text: 'Metro, Uber, souvenirs, airport taxi', options: rA },
      { text: '~€35', options: rNA },
      { text: '~€140', options: rNA },
    ],
    [
      { text: '🏆  TOTAL ESTIMATE', options: totL },
      { text: 'Madrid, Oct 29/30 – Nov 2, 2025', options: { ...tot, align: 'left', color: C.gray } },
      { text: '~€712', options: { ...tot, color: MAD } },
      { text: '~€2,848', options: { ...tot, color: C.orange } },
    ],
  ], {
    x: 0.5, y: 1.02, w: 12.3,
    colW: [2.5, 5.2, 2.1, 2.1],
    rowH: [0.42, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.48, 0.55],
    border: { type: 'solid', pt: 1, color: C.bgDeep },
  });

  s.addShape('rect', { x: 0.5, y: 6.52, w: 12.3, h: 0.68, fill: { color: C.bgCard }, line: { color: MAD, pt: 1 }, rounding: 0.04 });
  s.addText(
    '💡 Budget range: €600–900pp depending on accommodation choice and club nights.  ' +
    'El Tigre bar gives MASSIVE free tapas plates with every €3 beer — take advantage.  ' +
    'Oct 31 is a Friday + Nov 1 is a public holiday = the best possible Halloween combo in Europe.',
    {
      x: 0.65, y: 6.54, w: 12.0, h: 0.64,
      fontSize: 10.5, fontFace: 'Calibri', color: C.gray, wrap: true, valign: 'middle',
    }
  );

  bottomBar(s);
}

// ── Slide 12: Side-by-Side Comparison ─────────────────────────────────────────
function slide12(pres) {
  const s = pres.addSlide();
  s.addImage({ path: path.join(IMG, 'comparison_bg.png'), x: 0, y: 0, w: '100%', h: '100%' });
  s.addShape('rect', { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '000000', transparency: 55 }, line: { color: '000000' } });

  s.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.72, fill: { color: C.bgCard, transparency: 20 }, line: { color: C.bgCard } });
  s.addShape('rect', { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText('SIDE-BY-SIDE COMPARISON  —  WHERE ARE WE GOING?', {
    x: 0.4, y: 0.08, w: 12.5, h: 0.64,
    fontSize: 22, fontFace: 'Trebuchet MS', bold: true, color: C.white, valign: 'middle',
  });

  s.addShape('rect', { x: 3.9, y: 0.82, w: 4.1, h: 0.58, fill: { color: C.bgDeep, transparency: 10 }, line: { color: C.orange, pt: 2 }, rounding: 0.04 });
  s.addText('🇵🇹  LISBON', { x: 3.9, y: 0.82, w: 4.1, h: 0.58, align: 'center', fontSize: 20, fontFace: 'Trebuchet MS', bold: true, color: C.orange, valign: 'middle' });

  s.addShape('rect', { x: 8.15, y: 0.82, w: 4.1, h: 0.58, fill: { color: C.bgDeep, transparency: 10 }, line: { color: MAD, pt: 2 }, rounding: 0.04 });
  s.addText('🇪🇸  MADRID', { x: 8.15, y: 0.82, w: 4.1, h: 0.58, align: 'center', fontSize: 20, fontFace: 'Trebuchet MS', bold: true, color: MAD, valign: 'middle' });

  const rows = [
    ['💰 Trip Cost (pp)',      '~€735',                                '~€712',                         false],
    ['✈ Flight',              '3h 10m · DUS only',                    '2h 30m · DUS only',              false],
    ['🛫 Airport Choice',     'DUS only (Eurowings / TAP)',           'DUS only (Iberia / Eurowings)',   false],
    ['🍺 Beer Price',         '€2 – €3 per pint',                    '€2 – €4 per caña / tubo',        false],
    ['🌡 Late Oct Weather',   '17–22°C ☀ T-shirt nights',            '18–20°C ☀ Warmer at night',      false],
    ['🎃 Halloween Vibe',     '⭐⭐⭐⭐  Fri pub crawl + clubs',      '⭐⭐⭐⭐⭐ Fri + PUBLIC HOLIDAY', false],
    ['🏙 Unique Feature',     'Pink Street · Fado · Atlantic',        'Malasaña · Tapas · Vermouth',     false],
    ['🎭 Nightlife Hours',    'Bars till 2 AM · clubs till 6 AM',    'Clubs till 6–7 AM · free tapas',  false],
    ['🍴 Food Scene',         'Petiscos · Bacalhau · €15–25/meal',   'Tapas · Pintxos · €15–20/meal',   false],
    ['📍 Pub Crawl',          'Halloween Haunts · €28pp',             'Bar Crawl Madrid · €20pp',        false],
    ['🍳 Cooking Class',      'Cooking Lisbon · €110pp',              'Cooking Point · €85pp',           false],
    ['💡 Our Verdict',        'Atlantic charm, Fado, warmer Airbnb', 'BEST Halloween in Europe 2025',   true],
  ];

  rows.forEach((row, i) => {
    const ry = 1.5 + i * 0.455;
    const isLast = row[3];
    const rowFill = isLast ? C.bgDeep : (i % 2 === 0 ? C.bgCard : C.bg);
    const textColor = isLast ? C.white : C.offWhite;

    s.addShape('rect', { x: 0.3, y: ry, w: 3.5, h: 0.42, fill: { color: rowFill, transparency: isLast ? 0 : 20 }, line: { color: C.bgDeep } });
    s.addText(row[0], { x: 0.35, y: ry, w: 3.4, h: 0.42, fontSize: isLast ? 12 : 10.5, fontFace: 'Calibri', bold: isLast, color: isLast ? C.amber : C.gray, valign: 'middle' });

    s.addShape('rect', { x: 3.9, y: ry, w: 4.1, h: 0.42, fill: { color: rowFill, transparency: isLast ? 0 : 20 }, line: { color: isLast ? C.orange : C.bgDeep, pt: isLast ? 1.5 : 0.5 } });
    s.addText(row[1], { x: 3.95, y: ry, w: 4.0, h: 0.42, align: 'center', fontSize: isLast ? 11.5 : 10, fontFace: isLast ? 'Trebuchet MS' : 'Calibri', bold: isLast, color: isLast ? C.orange : textColor, valign: 'middle' });

    s.addShape('rect', { x: 8.15, y: ry, w: 4.1, h: 0.42, fill: { color: rowFill, transparency: isLast ? 0 : 20 }, line: { color: isLast ? MAD : C.bgDeep, pt: isLast ? 1.5 : 0.5 } });
    s.addText(row[2], { x: 8.2, y: ry, w: 4.0, h: 0.42, align: 'center', fontSize: isLast ? 11.5 : 10, fontFace: isLast ? 'Trebuchet MS' : 'Calibri', bold: isLast, color: isLast ? MAD : textColor, valign: 'middle' });
  });

  s.addShape('rect', { x: 0, y: 7.42, w: '100%', h: 0.08, fill: { color: C.orange }, line: { color: C.orange } });
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Halloween Trip 2025';
  pres.title = 'Halloween Trip 2025 – Lisbon vs Madrid';

  console.log('Building slides...');
  slide1(pres);  console.log('  ✓ Slide 1: Title');
  slide2(pres);  console.log('  ✓ Slide 2: Lisbon Overview');
  slide3(pres);  console.log('  ✓ Slide 3: Lisbon Flights');
  slide4(pres);  console.log('  ✓ Slide 4: Lisbon Airbnb');
  slide5(pres);  console.log('  ✓ Slide 5: Lisbon Activities');
  slide6(pres);  console.log('  ✓ Slide 6: Lisbon Cost Summary');
  slide7(pres);  console.log('  ✓ Slide 7: Madrid Overview');
  slide8(pres);  console.log('  ✓ Slide 8: Madrid Flights');
  slide9(pres);  console.log('  ✓ Slide 9: Madrid Airbnb');
  slide10(pres); console.log('  ✓ Slide 10: Madrid Activities');
  slide11(pres); console.log('  ✓ Slide 11: Madrid Cost Summary');
  slide12(pres); console.log('  ✓ Slide 12: Comparison');

  await pres.writeFile({ fileName: OUT });
  console.log(`\n✅ Saved: ${OUT}`);
}

main().catch(err => { console.error(err); process.exit(1); });
