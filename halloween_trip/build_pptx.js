'use strict';
const PptxGenJS = require('pptxgenjs');
const path = require('path');

const IMG = __dirname;
const OUT = path.join(__dirname, '..', 'halloween_trip_2025.pptx');

// ── Color palette — clean light theme ────────────────────────────────────────
const C = {
  bg:       'F7F8FC',   // page background (off-white)
  bgCard:   'FFFFFF',   // pure white cards
  bgSubtle: 'EEF2F7',   // subtle fill: table headers, alternating rows
  bgAlt:    'F3F4F8',   // alternating table row
  text:     '111827',   // primary text
  textMid:  '374151',   // body text
  textSoft: '6B7280',   // secondary labels
  border:   'E2E8F0',   // card / table borders
  green:    '059669',   // positive highlights
  white:    'FFFFFF',
  amber:    'D97706',   // totals, highlights
  coral:    'DC2626',   // warnings
  black:    '000000',
};
const LIS_C = 'F07800';  // Orange  — Lissabon
const SOF_C = '7B3FC5';  // Violet  — Sofia

// ── Helpers ───────────────────────────────────────────────────────────────────
function sectionHeader(s, title, accent) {
  s.addShape('rect',{x:0,y:0,w:'100%',h:0.78,fill:{color:C.bgCard},line:{color:C.border,pt:0.5}});
  s.addShape('rect',{x:0,y:0,w:0.08,h:0.78,fill:{color:accent},line:{color:accent}});
  s.addText(title,{x:0.22,y:0,w:13.0,h:0.78,fontSize:20,fontFace:'Calibri',bold:true,color:C.text,valign:'middle'});
  s.addShape('rect',{x:0,y:0.78,w:'100%',h:0.025,fill:{color:accent},line:{color:accent}});
}

function bottomBar(s, accent) {
  const col = accent || C.border;
  s.addShape('rect',{x:0,y:7.42,w:'100%',h:0.08,fill:{color:C.bgSubtle},line:{color:C.bgSubtle}});
  s.addShape('rect',{x:0,y:7.42,w:'100%',h:0.02,fill:{color:col},line:{color:col}});
}

function statBox(s,x,y,w,h,label,val,accent) {
  s.addShape('rect',{x,y,w,h,fill:{color:C.bgCard},line:{color:C.border,pt:0.75},rounding:0.05});
  s.addShape('rect',{x,y,w,h:0.06,fill:{color:accent},line:{color:accent},rounding:0.05});
  s.addText(label,{x,y:y+0.1,w,h:0.3,align:'center',fontSize:9,fontFace:'Calibri',color:C.textSoft});
  s.addText(val,{x,y:y+0.38,w,h:h-0.45,align:'center',fontSize:15,fontFace:'Calibri',bold:true,color:accent});
}

function dayCard(s,x,y,w,h,day,date,items,accent) {
  s.addShape('rect',{x,y,w,h,fill:{color:C.bgCard},line:{color:C.border,pt:0.75},rounding:0.05});
  s.addShape('rect',{x,y,w,h:0.44,fill:{color:accent},line:{color:accent},rounding:0.05});
  s.addText(day,{x:x+0.1,y:y+0.02,w:w-0.2,h:0.24,fontSize:11,fontFace:'Calibri',bold:true,color:C.white});
  s.addText(date,{x:x+0.1,y:y+0.26,w:w-0.2,h:0.16,fontSize:8.5,fontFace:'Calibri',color:'FFE0B2',italic:true});
  let ty = y+0.5;
  items.forEach(item => {
    s.addText(item,{x:x+0.1,y:ty,w:w-0.2,h:0.28,fontSize:9,fontFace:'Calibri',color:C.textMid,valign:'top',wrap:true});
    ty += 0.29;
  });
}

// ── Generic city slide builders ───────────────────────────────────────────────
function buildOverview(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  Reiseziel ${cfg.num} — ${cfg.name}, ${cfg.country}`, cfg.accent);

  // Hero image panel
  s.addImage({path:path.join(IMG,cfg.heroImg),x:0,y:0.83,w:7.0,h:5.6});
  // Fade from image to white panel
  s.addShape('rect',{x:5.5,y:0.83,w:1.8,h:5.6,fill:{color:C.bgCard},line:{color:C.bgCard},transparency:15});

  // Right content panel
  const rx = 7.15;
  s.addShape('rect',{x:rx,y:0.88,w:6.0,h:5.55,fill:{color:C.bgCard},line:{color:C.border,pt:0.5},rounding:0.05});

  s.addText(cfg.tagline,{x:rx+0.18,y:0.98,w:5.65,h:0.42,fontSize:16,fontFace:'Calibri',bold:true,color:cfg.accent});
  s.addText(cfg.subtitle,{x:rx+0.18,y:1.38,w:5.65,h:0.28,fontSize:9.5,fontFace:'Calibri',color:C.textSoft,italic:true});

  let si = 0;
  for(let row=0;row<2;row++) for(let col=0;col<3;col++){
    statBox(s,rx+0.18+col*1.92,1.8+row*1.32,1.78,1.18,cfg.stats[si][0],cfg.stats[si][1],cfg.accent);
    si++;
  }

  s.addShape('rect',{x:rx+0.12,y:4.54,w:5.76,h:0.025,fill:{color:C.border},line:{color:C.border}});
  s.addText(cfg.description,{x:rx+0.18,y:4.6,w:5.65,h:1.66,fontSize:10,fontFace:'Calibri',color:C.textMid,wrap:true,valign:'top'});

  bottomBar(s, cfg.accent);
}

function buildFlights(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name} — Flüge`, cfg.accent);

  const hO = {bold:true,color:C.white,fontSize:11,fontFace:'Calibri',fill:{color:cfg.accent},align:'center',valign:'middle'};
  const rO = {color:C.textMid,fontSize:10,fontFace:'Calibri',fill:{color:C.bgCard},align:'center',valign:'middle'};
  const rA = {...rO,fill:{color:C.bgAlt}};

  s.addText(`Hinflug: Deutschland → ${cfg.iataLabel}`,{x:0.4,y:1.02,w:12.5,h:0.34,fontSize:13,fontFace:'Calibri',bold:true,color:C.text});
  const hdr = ['Option','Flug','Strecke & Datum','Abflug','Ankunft','Dauer','Preis p.P.','Gesamt (×4)'];
  const outRows = cfg.flightsOut.map((r,i) => {
    const st = i%2===0?rO:rA;
    return [
      {text:r[0],options:i===0?{...st,color:C.green,bold:true}:st},
      {text:r[1],options:st},{text:r[2],options:st},
      {text:r[3],options:st},{text:r[4],options:st},{text:r[5],options:st},
      {text:r[6],options:st},
      {text:r[7],options:{...st,color:i===0?cfg.accent:C.amber,bold:i===0}},
    ];
  });
  s.addTable([[...hdr.map(h=>({text:h,options:hO}))],...outRows],{
    x:0.4,y:1.38,w:12.5,colW:[1.4,1.65,2.0,1.2,1.2,1.2,1.6,1.5],
    rowH:[0.38,...Array(outRows.length).fill(0.52)],
    border:{type:'solid',pt:0.5,color:C.border},
  });

  const ry = 1.38+0.38+outRows.length*0.52+0.22;
  s.addText(`Rückflug: ${cfg.iataLabel} → Deutschland (2. Nov.)`,{x:0.4,y:ry,w:12.5,h:0.34,fontSize:13,fontFace:'Calibri',bold:true,color:C.text});
  const retHdr = ['Flug','Strecke & Datum','Abflug','Ankunft','Dauer','Preis p.P.','Gesamt (×4)'];
  const retRow = cfg.flightRet;
  s.addTable([
    [...retHdr.map(h=>({text:h,options:hO}))],
    [{text:retRow[0],options:{...rO,color:C.green}},{text:retRow[1],options:rO},{text:retRow[2],options:rO},{text:retRow[3],options:rO},{text:retRow[4],options:rO},{text:retRow[5],options:rO},{text:retRow[6],options:{...rO,color:C.amber,bold:true}}],
  ],{x:0.4,y:ry+0.37,w:12.5,colW:[2.5,2.0,1.3,1.3,1.3,1.8,1.6],rowH:[0.38,0.52],border:{type:'solid',pt:0.5,color:C.border}});

  const ny = ry+0.37+0.38+0.52+0.15;
  s.addShape('rect',{x:0.4,y:ny,w:12.5,h:0.72,fill:{color:C.bgSubtle},line:{color:C.border,pt:0.5},rounding:0.04});
  s.addText(cfg.flightNote,{x:0.55,y:ny+0.05,w:12.2,h:0.62,fontSize:9.5,fontFace:'Calibri',color:C.textSoft,wrap:true,valign:'middle'});

  const py = ny+0.85;
  s.addShape('rect',{x:8.8,y:py,w:4.1,h:0.85,fill:{color:C.bgCard},line:{color:cfg.accent,pt:2},rounding:0.06});
  s.addText('Hin & Zurück · 4 Personen',{x:8.8,y:py+0.02,w:4.1,h:0.28,align:'center',fontSize:9.5,fontFace:'Calibri',color:C.textSoft});
  s.addText(cfg.flightTotal,{x:8.8,y:py+0.32,w:4.1,h:0.48,align:'center',fontSize:20,fontFace:'Calibri',bold:true,color:cfg.accent});

  bottomBar(s, cfg.accent);
}

function buildAirbnb(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name} — Unterkunft`, cfg.accent);

  s.addImage({path:path.join(IMG,cfg.airbnbImg),x:0,y:0.83,w:5.5,h:3.4});
  s.addImage({path:path.join(IMG,cfg.mapImg),x:0,y:4.28,w:5.5,h:2.3});

  const rx = 5.8;
  s.addShape('rect',{x:rx,y:0.9,w:7.3,h:0.55,fill:{color:C.bgSubtle},line:{color:C.border,pt:0.5},rounding:0.04});
  s.addText(cfg.airbnbTitle,{x:rx+0.15,y:0.9,w:7.0,h:0.55,fontSize:14,fontFace:'Calibri',bold:true,color:C.text,valign:'middle'});

  cfg.airbnbSpecs.forEach((spec,i)=>{
    const bx = rx+(i%3)*2.45; const by = 1.56+Math.floor(i/3)*0.55;
    const col = spec[2]||cfg.accent;
    s.addShape('rect',{x:bx,y:by,w:2.3,h:0.42,fill:{color:C.bgCard},line:{color:col,pt:1.5},rounding:0.04});
    s.addText(spec[0],{x:bx,y:by,w:2.3,h:0.42,align:'center',fontSize:11,fontFace:'Calibri',bold:true,color:col,valign:'middle'});
  });

  s.addShape('rect',{x:rx,y:2.74,w:7.3,h:0.52,fill:{color:C.bgSubtle},line:{color:C.border,pt:0.5},rounding:0.04});
  s.addText(cfg.airbnbLocation,{x:rx+0.1,y:2.74,w:7.1,h:0.52,fontSize:10,fontFace:'Calibri',color:C.textMid,valign:'middle'});

  s.addShape('rect',{x:rx,y:3.38,w:7.3,h:1.52,fill:{color:C.bgCard},line:{color:cfg.accent,pt:1.5},rounding:0.04});
  s.addText('Preise  (30. Okt. – 2. Nov., 3 Nächte)',{x:rx+0.2,y:3.44,w:7.0,h:0.3,fontSize:11,fontFace:'Calibri',bold:true,color:cfg.accent});
  cfg.airbnbPricing.forEach((row,i)=>{
    const py = 3.8+i*0.215;
    const isTotal = i===3; const isPP = i===4;
    s.addText(row[0],{x:rx+0.2,y:py,w:4.0,h:0.21,fontSize:10,fontFace:'Calibri',color:C.textSoft});
    s.addText(row[1],{x:rx+4.5,y:py,w:2.6,h:0.21,align:'right',fontSize:10,fontFace:'Calibri',bold:isTotal||isPP,color:isTotal?cfg.accent:isPP?C.green:C.textMid});
  });

  s.addShape('rect',{x:rx,y:5.0,w:7.3,h:0.66,fill:{color:cfg.accent},line:{color:cfg.accent},rounding:0.06});
  s.addText('🔍  Airbnb: ' + cfg.name + ' durchsuchen  →',{x:rx,y:5.0,w:7.3,h:0.66,align:'center',fontSize:15,fontFace:'Calibri',bold:true,color:C.white,valign:'middle',hyperlink:{url:cfg.airbnbSearchUrl,tooltip:'Airbnb öffnen — 30. Okt.–2. Nov., 4 Gäste, 2 SZ'}});
  s.addText(cfg.airbnbAlt,{x:rx+0.1,y:5.74,w:7.1,h:0.28,fontSize:9,fontFace:'Calibri',color:C.textSoft,italic:true,hyperlink:{url:cfg.airbnbAltUrl,tooltip:'Alternative Suche öffnen'}});
  s.addText('⚠  Halloween-Wochenende — frühzeitig buchen!',{x:rx+0.1,y:6.08,w:7.1,h:0.28,fontSize:9.5,fontFace:'Calibri',bold:true,color:C.coral});

  bottomBar(s, cfg.accent);
}

function buildActivities(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name} — Tagesprogramm`, cfg.accent);
  cfg.days.forEach((d,i)=>{
    const col = i%2; const row = Math.floor(i/2);
    dayCard(s,0.3+col*6.5,0.98+row*3.15,6.2,3.0,d.day,d.date,d.items,cfg.accent);
  });
  s.addShape('rect',{x:0.3,y:7.1,w:12.7,h:0.24,fill:{color:C.bgSubtle},line:{color:C.border,pt:0.5},rounding:0.03});
  s.addText(cfg.activityLinks,{x:0.4,y:7.1,w:12.5,h:0.24,fontSize:8,fontFace:'Calibri',color:C.textSoft,valign:'middle'});
  bottomBar(s, cfg.accent);
}

function buildCosts(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name} — Kostenübersicht`, cfg.accent);

  const hO  = {bold:true,color:C.white,fontSize:12,fontFace:'Calibri',fill:{color:cfg.accent},align:'center',valign:'middle'};
  const rO  = {color:C.textMid,fontSize:11,fontFace:'Calibri',fill:{color:C.bgCard},align:'left',valign:'middle'};
  const rA  = {...rO,fill:{color:C.bgAlt}};
  const rN  = {...rO,align:'right'};
  const rNA = {...rA,align:'right'};
  const tot = {bold:true,color:C.white,fontSize:13,fontFace:'Calibri',fill:{color:cfg.accent},align:'right',valign:'middle'};
  const totL = {...tot,align:'left'};

  const rows = [
    [{text:'Posten',options:hO},{text:'Details',options:hO},{text:'p.P.',options:{...hO,color:C.white}},{text:'Gesamt ×4',options:{...hO,color:'FFE0B2'}}],
    ...cfg.costRows.map((r,i) => {
      const isLast = i===cfg.costRows.length-1;
      if(isLast) return [{text:r[0],options:totL},{text:r[1],options:{...tot,align:'left',color:'E0E0FF',fill:{color:cfg.accent}}},{text:r[2],options:{...tot,color:C.white}},{text:r[3],options:{...tot,color:'FFE0B2'}}];
      const st = i%2===0?rO:rA; const stn = i%2===0?rN:rNA;
      return [{text:r[0],options:st},{text:r[1],options:st},{text:r[2],options:stn},{text:r[3],options:stn}];
    }),
  ];

  const nData = cfg.costRows.length;
  s.addTable(rows,{x:0.5,y:1.02,w:12.3,colW:[2.5,5.2,2.1,2.1],rowH:[0.42,...Array(nData).fill(0.47)],border:{type:'solid',pt:0.5,color:C.border}});

  const ty = 1.02+0.42+nData*0.47+0.08;
  s.addShape('rect',{x:0.5,y:ty,w:12.3,h:0.65,fill:{color:C.bgSubtle},line:{color:C.border,pt:0.5},rounding:0.04});
  s.addText(cfg.costNote,{x:0.65,y:ty+0.05,w:12.0,h:0.55,fontSize:10,fontFace:'Calibri',color:C.textSoft,wrap:true,valign:'middle'});

  bottomBar(s, cfg.accent);
}

// ── CITY DATA ─────────────────────────────────────────────────────────────────
const LISSABON = {
  num:1, name:'Lissabon', country:'Portugal', flag:'🇵🇹', accent:LIS_C,
  heroImg:'lisbon_hero.png', mapImg:'lisbon_map.png', airbnbImg:'lisbon_airbnb.png',
  iataLabel:'LIS (Lissabon)',
  tagline:'Nachtleben-Hauptstadt des Atlantiks',
  subtitle:'Bairro Alto · Pink Street · Alfama Fado · Castelo de São Jorge',
  stats:[
    ['✈ Flugdauer','3h 10m\nab DUS'],['🍺 Bier','€2 – €3\npro Pint'],
    ['🌡 Wetter','17–22°C ☀\nEnde Oktober'],['💰 Kosten ca.','~€735\npro Person'],
    ['🏙 Nightlife','⭐⭐⭐⭐⭐\nWeltklasse'],['📍 Ab Flughafen','Nur DUS\n(Eurowings/TAP)'],
  ],
  description:'Lissabon punktet mit warmem Atlantik-Charme und einer der günstigsten, lebendigsten Nachtlebensszenen Europas. Bairro Alto explodiert ab Mitternacht mit 200+ Bars. Pink-Street-Clubs laufen bis in den Morgen. Die Stadt feiert Halloween ausgiebig — Kneipentour, Kostümparties und Mega-Club-Events. Ende Oktober: T-Shirt-Wetter garantiert.',
  flightsOut:[
    ['⭐ Empfohlen','EW9602\nEurowings','DUS → LIS\n30. Okt. morgens','10:15','12:20','3h 05m','~€100','~€400'],
    ['Frühabflug','EW9602\nEurowings','DUS → LIS\n30. Okt.','06:20','08:25','3h 05m','~€80','~€320'],
    ['Vorabend','EW9604\nEurowings','DUS → LIS\n29. Okt. abends','~18:30','~20:30','3h 00m','~€90','~€360'],
    ['Premium','TP541\nTAP Portugal','DUS → LIS\n30. Okt.','Verschieden','Verschieden','3h 10m','€150–180','€600–720'],
  ],
  flightRet:['EW9603 Eurowings','LIS → DUS  /  2. Nov.','18:00','21:50','3h 50m','~€90','~€360'],
  flightNote:'⚠  DUS (Düsseldorf) ist der EINZIGE Flughafen mit Direktverbindung nach Lissabon — Eurowings & TAP Air Portugal fliegen täglich non-stop.  DTM & FMO erfordern Umsteigen (8h+ Gesamtreisezeit).  Buchen: eurowings.com · kayak.com · Hin & Zurück Ø ~€170–190 p.P.',
  flightTotal:'€640 – €760 gesamt',
  airbnbTitle:'Bairro Alto / São Roque — 2-Zimmer-Wohnung',
  airbnbSpecs:[
    ['🛏 2 Schlafzimmer'],['🚿 2 Bäder'],['🏗 Balkon ✓',[],C.green],
    ['👥 4 Gäste'],['📍 Zentrallage'],['❄ Klimaanlage ✓'],
  ],
  airbnbLocation:'📌 Bairro Alto — direkt im Herz der 200+ Bars, Pink Street 10 Fußminuten entfernt',
  airbnbPricing:[['Pro Nacht','~€130'],['3 Nächte','~€390'],['Reinigung & Gebühren','~€60–80'],['Gesamt für 4 Pers.','~€450–480'],['Pro Person','~€113–120']],
  airbnbSearchUrl:'https://www.airbnb.com/s/Bairro-Alto--Lisbon--Portugal/homes?checkin=2025-10-30&checkout=2025-11-02&adults=4&min_bedrooms=2&room_types%5B%5D=Entire+home%2Fapt',
  airbnbAlt:'📍 Alternativ in Chiado suchen: airbnb.com/s/Chiado--Lisbon (2 SZ, 4 Gäste, gleiche Daten)',
  airbnbAltUrl:'https://www.airbnb.com/s/Chiado--Lisbon--Portugal/homes?checkin=2025-10-30&checkout=2025-11-02&adults=4&min_bedrooms=2&room_types%5B%5D=Entire+home%2Fapt',
  days:[
    {day:'TAG 1 — Do., 30. Okt.',date:'Ankunftstag',items:[
      '✈  EW9602 landet LIS 12:20 · Einchecken in Bairro Alto',
      '🚶  SANDEMANs Stadtführung 15:30 @ Praça do Comércio',
      '   (2,5h, Trinkgeld €10–15 · Baixa, Alfama, Aussichtspunkte)',
      '🍽  Abendessen: Petiscos bei Tasca do Chico ab 20 Uhr',
      '🍺  Bairro Alto Kneipenbummel: 200+ Bars, €2–3 Bier',
      '   Start @ Park Bar (Dachterrasse) → Pavilhão Chinês',
    ]},
    {day:'TAG 2 — Fr., 31. Okt. · Halloween-Nacht',date:'Highlights des Trips',items:[
      '🍳  Kochkurs Cooking Lisbon ab 9:30 Uhr (4h, €110 p.P.)',
      '   Mercado 31 de Janeiro · 3-Gang-Menü + Wein inklusive',
      '   Buchung: cookinglisbon.com/market-tour',
      '🕙  Halloween-Kneipentour 22 Uhr · €28 p.P.',
      '   Treffpkt.: Praça Luís de Camões · 3 Bars · 1h offene Bar',
      '   Shots · Kostüm-Wettbewerb · VIP-Clubeintritt',
      '🎶  Ab 1 Uhr: Brunch Electronik Halloween Special',
      '   Pavilhão Carlos Lopes · Chris Liebing B2B Luke Slater',
    ]},
    {day:'TAG 3 — Sa., 1. Nov.',date:'Erkunden & Entspannen',items:[
      '🌅  Ausschlafen — Pastéis de Nata bei Pasteis de Belém',
      '🚋  Straßenbahn 28: malerische Fahrt durch Alfama',
      '🍺  Cerveja Musa Craft-Brauerei, Mouraria (lokale IPAs)',
      '🌇  Miradouro da Graça: bester Sonnenuntergangs-Aussichtsp.',
      '🎵  Fado-Abend: Tasca do Chico oder Zé da Viola',
      '🌃  Pink Street Nacht: Pensão Amor + Music Box · Lux Frágil',
    ]},
    {day:'TAG 4 — So., 2. Nov.',date:'Abreisetag',items:[
      '☕  Morgen-Kaffee @ A Brasileira, Chiado',
      '🏰  Kurzer Besuch Castelo de São Jorge (optional)',
      '🛍  Souvenirs: Ginjinha-Likör, Sardinenkonserven',
      '🚌  Transfer zum Flughafen (45 Min. einplanen)',
      '✈  EW9603  18:00 LIS → 21:50 DUS',
    ]},
  ],
  activityLinks:'Kneipentour: eventbrite.co.uk/e/1981093800888  ·  Kochkurs: cookinglisbon.com/booking  ·  Stadtführung: neweuropetours.eu/lisbon  ·  Lux Frágil: luxfragil.com',
  costRows:[
    ['Hin- und Rückflug','Eurowings EW9602/9603  DUS ↔ LIS','~€170','~€680'],
    ['Unterkunft','Airbnb Bairro Alto — 3 Nächte (inkl. Gebühren)','~€115','~€460'],
    ['Halloween-Kneipentour','Halloween Haunts — Chiado (31. Okt.)','€28','€112'],
    ['Kochkurs','Cooking Lisbon — Markt + 3-Gang-Menü + Wein','€110','€440'],
    ['Stadtführung','SANDEMANs — Praça do Comércio (Trinkgeld)','~€12','~€48'],
    ['Essen (3 Tage)','Mahlzeiten, Petiscos, Pastéis — Ø €15–25/Mahlzeit','~€120','~€480'],
    ['Getränke & Ausgehen','Bairro Alto + Pink Street + Clubs (3 Nächte)','~€150','~€600'],
    ['Sonstiges','Bolt/Uber, Trinkgelder, Souvenirs','~€30','~€120'],
    ['🏆  Gesamt','Lissabon, 30. Okt. – 2. Nov. 2025','~€735','~€2.940'],
  ],
  costNote:'💡 Budgetrahmen: €600–900 p.P. je nach Partylust.  Bier ab €2–3, Supermarkt-Runs helfen sparen.  Clubs (Lux Frágil, Brunch Electronik) kosten je ~€20–30 extra.  Leitungswasser in Restaurants gratis.',
};

const SOFIA = {
  num:2, name:'Sofia', country:'Bulgarien', flag:'🇧🇬', accent:SOF_C,
  heroImg:'sofia_hero.png', mapImg:'sofia_map.png', airbnbImg:'sofia_airbnb.png',
  iataLabel:'SOF (Sofia)',
  tagline:'Günstigstes Ziel — €1 Bier & Vitosha-Bergblick',
  subtitle:'Vitosha Blvd · Yalta Club · Alexander-Nevsky-Kathedrale · Studentski Grad',
  stats:[
    ['✈ Flugdauer','~2h 30m\nab DUS'],['🍺 Bier','€1 – €2\n(günstigstes Eu.!)'],
    ['🌡 Wetter','10–14°C 🧥\nEnde Oktober'],['💰 Kosten ca.','~€520\npro Person'],
    ['🏙 Nightlife','⭐⭐⭐⭐\nYalta · Sin City'],['📍 Ab Flughafen','DUS direkt\n(Bulgaria Air/Ryanair)'],
  ],
  description:'Sofia ist mit Abstand das günstigste Ziel — ein Bier kostet €1–2, Abendessen in einem guten Restaurant ~€8–12. Die bulgarische Hauptstadt überrascht mit einer lebhaften Ausgehszene rund um den Vitosha Boulevard und den Studentenvierteln. Im Hintergrund: das Vitosha-Gebirge, direkt an der Stadtgrenze. Die Alexander-Nevsky-Kathedrale leuchtet nachts golden — perfekte Halloween-Kulisse.',
  flightsOut:[
    ['⭐ Empfohlen','Bulgaria Air\n(FB)','DUS → SOF\n30. Okt.','Morgens','~11:00','~2h 30m','~€80–120','~€320–480'],
    ['Option 2','Ryanair\n(FR)','DUS → SOF\n30. Okt.','Verschieden','Verschieden','~2h 30m','~€50–90','~€200–360'],
    ['Vorabend','Bulgaria Air\n(FB)','DUS → SOF\n29. Okt.','Abends','Abends','~2h 30m','~€90–130','~€360–520'],
    ['Budget+','Wizz Air\n(W6)','DUS → SOF\n30. Okt.','Verschieden','Verschieden','~2h 30m','~€40–80','~€160–320'],
  ],
  flightRet:['Bulgaria Air / Ryanair  (SOF→DUS)','SOF → DUS  /  2. Nov.','Verschieden','Verschieden','~2h 30m','~€60–100','~€240–400'],
  flightNote:'✅  DUS (Düsseldorf) → SOF (Sofia): Direktflüge mit Bulgaria Air und Ryanair.  Flughafen Sofia liegt 10 km östlich vom Zentrum — Metro Linie M1 (€0,80, 20 Min.) direkt ins Zentrum.  Buchen: bulgariaair.com · ryanair.com · wizzair.com · skyscanner.de',
  flightTotal:'€320 – €560 gesamt',
  airbnbTitle:'Vitosha Blvd / Zentrum — 2-Zi.-Wohnung',
  airbnbSpecs:[
    ['🛏 2 Schlafzimmer'],['🚿 2 Bäder'],['🏗 Balkon ✓',[],C.green],
    ['👥 4 Gäste'],['❄ Klimaanlage ✓'],['🏔 Vitosha-Blick',[],SOF_C],
  ],
  airbnbLocation:'📌 Vitosha Blvd / Stadtzentrum — direkt an der besten Bar-Meile, 10 Min. zu Fuß zur Kathedrale',
  airbnbPricing:[['Pro Nacht','~€60–80'],['3 Nächte','~€180–240'],['Reinigung & Gebühren','~€30–50'],['Gesamt für 4 Pers.','~€210–290'],['Pro Person','~€53–73']],
  airbnbSearchUrl:'https://www.airbnb.com/s/Sofia--Bulgaria/homes?checkin=2025-10-30&checkout=2025-11-02&adults=4&min_bedrooms=2&room_types%5B%5D=Entire+home%2Fapt',
  airbnbAlt:'📍 Alternativ im Studentenviertel suchen: airbnb.com/s/Studentski-Grad--Sofia (lebhafter, günstiger)',
  airbnbAltUrl:'https://www.airbnb.com/s/Studentski-Grad--Sofia--Bulgaria/homes?checkin=2025-10-30&checkout=2025-11-02&adults=4&min_bedrooms=2&room_types%5B%5D=Entire+home%2Fapt',
  days:[
    {day:'TAG 1 — Do., 30. Okt.',date:'Ankunftstag',items:[
      '✈  Landet SOF ~11:00 · Metro M1 ins Zentrum (€0,80!)',
      '🏠  Einchecken in der Vitosha-Nähe',
      '🚶  SANDEMANs Free Tour 14:30 @ Alexander-Nevsky-Platz',
      '   Kathedrale, Präsidentenpalast, Rotunde Sveti Georgi',
      '🍺  Erste Biere am Vitosha Blvd: €1–2 pro Bier (kein Tippfehler!)',
      '🌃  Erste Nacht: Vitosha Blvd Bars & Lounge-Cafés',
    ]},
    {day:'TAG 2 — Fr., 31. Okt. · Halloween-Nacht',date:'Highlights des Trips',items:[
      '🏔  Morgens: Vitosha-Berg Wanderung (Bus/Gondel ab Stadtrand)',
      '   Bergblick über ganz Sofia — kostenloser Nationalpark',
      '🍳  Bulgarischer Kochkurs 15:00 Uhr (3h, ~€45 p.P.)',
      '   Shopska-Salat + Moussaka + Rakija-Verkostung',
      '🎃  Halloween Kneipentour ab 22 Uhr · ~€15 p.P.',
      '   Studentski Grad: Sin City · Yalta Club · Mixtape 5',
      '   Clubs bis 5 Uhr morgens · €1 Bier die ganze Nacht',
    ]},
    {day:'TAG 3 — Sa., 1. Nov.',date:'Erkunden & Ausgehen',items:[
      '🕌  Alexander-Nevsky-Kathedrale innen (kostenlos, atemberaubend)',
      '🏛  Nationales Historisches Museum (optional)',
      '🛒  Weiblicher Markt (Ženski Pazar): lokales Essen & Gewürze',
      '🍽  Abendessen: Mehana-Restaurant — Shopska + Bier €8 gesamt',
      '🌃  Letzte Nacht: Lozenets-Viertel Bars oder Yalta-Club-Afterparty',
    ]},
    {day:'TAG 4 — So., 2. Nov.',date:'Abreisetag',items:[
      '☕  Banitsa (bulgarisches Blätterteig-Gebäck) zum Frühstück (€0,50)',
      '🛍  Souvenirs: Rosen-Öl, Rakija, Mamarnitsa (Glücksbringer)',
      '🚇  Metro M1 zum Flughafen (€0,80, 20 Min.)',
      '✈  Bulgaria Air / Ryanair SOF → DUS (Nachmittag)',
    ]},
  ],
  activityLinks:'Kneipentour: eventbrite.com → "Sofia Halloween Bar Crawl"  ·  Kochkurs: cookly.app → "Sofia Cooking Class"  ·  Stadtführung: neweuropetours.eu/sofia  ·  Yalta Club: yaltaclub.com',
  costRows:[
    ['Hin- und Rückflug','Bulgaria Air / Ryanair  DUS ↔ SOF','~€140','~€560'],
    ['Unterkunft','Airbnb Zentrum — 3 Nächte (inkl. Gebühren)','~€65','~€260'],
    ['Halloween Kneipentour','Bar Crawl Sofia (31. Okt.)','~€15','~€60'],
    ['Kochkurs','Bulgarische Küche — Moussaka + Rakija','~€45','~€180'],
    ['Stadtführung','SANDEMANs — Alexander-Nevsky (Trinkgeld)','~€10','~€40'],
    ['Essen (3 Tage)','Mehana, Banitsa, Markt — Ø €10–15/Mahlzeit','~€90','~€360'],
    ['Getränke & Ausgehen','Vitosha Blvd + Clubs (€1 Bier, 3 Nächte)','~€120','~€480'],
    ['Sonstiges','Metro, Uber, Souvenirs','~€25','~€100'],
    ['🏆  Gesamt','Sofia, 30. Okt. – 2. Nov. 2025','~€510','~€2.040'],
  ],
  costNote:'💡 Günstigstes Ziel: €1–2 Bier, €8 Abendessen, €0,80 Metro!  Budgetrahmen: €400–650 p.P.  Ideal wenn Ihr nach dem letzten Trip die Kasse schonen wollt — ohne auf Spaß zu verzichten.',
};

// ── Slide 1: Titelfolie ───────────────────────────────────────────────────────
function slide1(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bgCard };

  // Top accent bar — gradient effect with two city colors
  s.addShape('rect',{x:0,y:0,w:6.67,h:0.07,fill:{color:LIS_C},line:{color:LIS_C}});
  s.addShape('rect',{x:6.67,y:0,w:6.67,h:0.07,fill:{color:SOF_C},line:{color:SOF_C}});

  // Title block
  s.addText('TRIP-PLANUNG 2025',{x:0.8,y:0.55,w:11.73,h:1.1,align:'center',fontSize:54,fontFace:'Calibri',bold:true,color:C.text});
  s.addText('30. Oktober – 2. November  ·  4 Personen  ·  Europäische Hauptstadt',{x:0.8,y:1.62,w:11.73,h:0.48,align:'center',fontSize:18,fontFace:'Calibri',color:C.textSoft});

  // Divider
  s.addShape('rect',{x:2.5,y:2.24,w:8.33,h:0.015,fill:{color:C.border},line:{color:C.border}});

  // Previously visited chip
  s.addShape('rect',{x:2.5,y:2.3,w:8.33,h:0.52,fill:{color:C.bgSubtle},line:{color:C.border,pt:0.5},rounding:0.05});
  s.addText('Bereits bereist:   ✓ Riga 2021   ✓ Athen 2022   ✓ Tirana 2023   ✓ Belgrad 2024',{x:2.5,y:2.3,w:8.33,h:0.52,align:'center',fontSize:13,fontFace:'Calibri',color:C.textSoft,valign:'middle'});

  // "Die Kandidaten" heading
  s.addText('Die Kandidaten für Trip #5:',{x:0.8,y:3.06,w:11.73,h:0.44,align:'center',fontSize:18,fontFace:'Calibri',bold:true,color:C.text});

  // City cards
  const cards = [
    { x:1.0,  label:'🇵🇹  LISSABON', col:LIS_C, stats:['✈ 3h 10m  ·  DUS direkt','🌡 17–22°C  Ende Oktober','💰 ~€735 pro Person','🍺 €2–3 Bier'] },
    { x:7.17, label:'🇧🇬  SOFIA',    col:SOF_C, stats:['✈ ~2h 30m  ·  DUS direkt','🌡 10–14°C  Ende Oktober','💰 ~€510 pro Person','🍺 €1–2 Bier'] },
  ];
  cards.forEach(c => {
    s.addShape('rect',{x:c.x,y:3.6,w:5.17,h:3.55,fill:{color:C.bgCard},line:{color:c.col,pt:2},rounding:0.08});
    s.addShape('rect',{x:c.x,y:3.6,w:5.17,h:0.72,fill:{color:c.col},line:{color:c.col},rounding:0.08});
    s.addText(c.label,{x:c.x,y:3.6,w:5.17,h:0.72,align:'center',fontSize:24,fontFace:'Calibri',bold:true,color:C.white,valign:'middle'});
    c.stats.forEach((line,i) => {
      s.addText(line,{x:c.x+0.25,y:4.44+i*0.62,w:4.67,h:0.56,fontSize:13.5,fontFace:'Calibri',color:C.textMid,valign:'middle'});
    });
  });

  // Footer note
  s.addText('Beide Ziele: Direktflug ab Düsseldorf (DUS)',{x:0,y:7.2,w:'100%',h:0.32,align:'center',fontSize:12,fontFace:'Calibri',color:C.textSoft,italic:true});

  // Bottom accent bar
  s.addShape('rect',{x:0,y:7.43,w:6.67,h:0.07,fill:{color:LIS_C},line:{color:LIS_C}});
  s.addShape('rect',{x:6.67,y:7.43,w:6.67,h:0.07,fill:{color:SOF_C},line:{color:SOF_C}});
}

// ── Slide 12: Direktvergleich ─────────────────────────────────────────────────
function buildComparison(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '⚖️  Direktvergleich — Lissabon vs. Sofia', C.text);

  const catW=3.4, cityW=4.55;
  const cols=[
    {label:'🇵🇹 Lissabon', color:LIS_C, x:catW+0.35},
    {label:'🇧🇬 Sofia',    color:SOF_C, x:catW+0.35+cityW+0.1},
  ];

  cols.forEach(col=>{
    s.addShape('rect',{x:col.x,y:0.85,w:cityW,h:0.54,fill:{color:col.color},line:{color:col.color},rounding:0.05});
    s.addText(col.label,{x:col.x,y:0.85,w:cityW,h:0.54,align:'center',fontSize:18,fontFace:'Calibri',bold:true,color:C.white,valign:'middle'});
  });

  const rows=[
    ['💰 Kosten p.P.',      '~€735',                          '~€510  💸 Günstigstes Ziel!'],
    ['✈ Flugdauer',         '3h 10m  ab DUS',                 '~2h 30m  ab DUS'],
    ['🛫 Flughafen',        'DUS — Eurowings / TAP',          'DUS — Bulgaria Air / Ryanair'],
    ['🍺 Bierpreis',        '€2–3  (Super Bock / Sagres)',    '€1–2  (günstigstes Eu. 🏆)'],
    ['🌡 Wetter Ende Okt.', '17–22°C  ☀',                    '10–14°C  🧥'],
    ['🎃 Halloween',        '⭐⭐⭐⭐  Clubs bis 6 Uhr',       '⭐⭐⭐⭐  Yalta · Sin City bis 5h'],
    ['🏙 Viertel',          'Bairro Alto · Pink Street',      'Vitosha Blvd · Studentski Grad'],
    ['🌟 Highlight',        'Fado · Atlantik-Charme',         'Alexander-Nevsky · Vitosha-Berg'],
    ['🍳 Kochkurs',         '€110 p.P.',                      '~€45 p.P.'],
    ['🏠 Unterkunft p.P.',  '~€113–120',                      '~€53–73  (3× günstiger!)'],
    ['💡 Stärke',           'Nightlife-Weltklasse + Wärme',   'Extremes Budget + Überraschung'],
  ];

  const rowH=0.5;
  rows.forEach((row,i)=>{
    const ry=1.43+i*rowH;
    const isLast=i===rows.length-1;
    const rowFill=isLast?C.bgSubtle:(i%2===0?C.bgCard:C.bgAlt);

    s.addShape('rect',{x:0.3,y:ry,w:catW,h:rowH,fill:{color:rowFill},line:{color:C.border,pt:0.5}});
    s.addText(row[0],{x:0.4,y:ry,w:catW-0.15,h:rowH,fontSize:isLast?11:10,fontFace:'Calibri',bold:isLast,color:isLast?C.text:C.textSoft,valign:'middle'});

    cols.forEach((col,ci)=>{
      s.addShape('rect',{x:col.x,y:ry,w:cityW,h:rowH,fill:{color:rowFill},line:{color:isLast?col.color:C.border,pt:isLast?1.5:0.5}});
      s.addText(row[ci+1],{x:col.x+0.1,y:ry,w:cityW-0.2,h:rowH,align:'center',fontSize:isLast?11:10,fontFace:'Calibri',bold:isLast,color:isLast?col.color:C.textMid,valign:'middle'});
    });
  });

  bottomBar(s);
}

// ── Slide 13: Hall of Fame ────────────────────────────────────────────────────
function buildHallOfFame(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🏆  Trip Hall of Fame — Die Legende wächst', C.amber);

  const past = [
    { year:'2021', city:'Riga',    flag:'🇱🇻', color:'C9973A',
      lines:['🍺 Bier: €0,80 — ewiger Rekord', '🌃 Clubs bis 7 Uhr morgens', '🏆 Günstigster Trip aller Zeiten'] },
    { year:'2022', city:'Athen',   flag:'🇬🇷', color:'3A72C9',
      lines:['🌅 Dachbar mit Akropolis-Blick', '🐟 Ouzo-Entscheidung: bereut', '🏆 Wärmster Trip bisher'] },
    { year:'2023', city:'Tirana',  flag:'🇦🇱', color:'C93A3A',
      lines:['🥘 Bestes Essen: unerwartet gut', '🎵 Barkeeper wurde Freund', '🏆 Geheimtipp des Jahres'] },
    { year:'2024', city:'Belgrad', flag:'🇷🇸', color:'7A3AC9',
      lines:['⛵ Splavovi-Clubs am Fluss', '🌄 Heimkehr beim Sonnenaufgang', '🏆 Wildeste Nacht aller Zeiten'] },
  ];

  const cW=3.6, cH=2.82, gX=0.2, gY=0.18;
  past.forEach((t,i)=>{
    const col=i%2, row=Math.floor(i/2);
    const x=0.25+col*(cW+gX);
    const y=1.02+row*(cH+gY);
    s.addShape('rect',{x,y,w:cW,h:cH,fill:{color:C.bgCard},line:{color:C.border,pt:0.75},rounding:0.06});
    s.addShape('rect',{x,y,w:cW,h:0.46,fill:{color:t.color},line:{color:t.color},rounding:0.06});
    s.addText(`${t.flag}  ${t.year}`,{x,y,w:cW,h:0.46,align:'center',fontSize:17,fontFace:'Calibri',bold:true,color:C.white,valign:'middle'});
    s.addText(t.city.toUpperCase(),{x,y:y+0.5,w:cW,h:0.38,align:'center',fontSize:14,fontFace:'Calibri',bold:true,color:C.text});
    s.addShape('rect',{x:x+0.15,y:y+0.94,w:cW-0.3,h:0.02,fill:{color:C.border},line:{color:C.border}});
    t.lines.forEach((line,li)=>{
      const isLast=li===t.lines.length-1;
      s.addText(line,{x:x+0.1,y:y+1.0+li*0.56,w:cW-0.2,h:0.52,fontSize:9.5,fontFace:'Calibri',
        color:isLast?t.color:C.textMid,bold:isLast,wrap:true,valign:'top'});
    });
  });

  // Right: 2025 feature panel
  const nx=7.9, nw=5.15, ny=1.0, nh=6.28;
  s.addShape('rect',{x:nx,y:ny,w:nw,h:nh,fill:{color:C.bgCard},line:{color:C.border,pt:0.75},rounding:0.08});
  s.addShape('rect',{x:nx,y:ny,w:nw,h:0.62,fill:{color:C.amber},line:{color:C.amber},rounding:0.08});
  s.addText('2025  ·  Next Stop',{x:nx,y:ny,w:nw,h:0.62,align:'center',fontSize:16,fontFace:'Calibri',bold:true,color:C.white,valign:'middle'});

  s.addText('✈',{x:nx,y:ny+0.7,w:nw,h:0.7,align:'center',fontSize:44,fontFace:'Segoe UI Emoji'});
  s.addText('TRIP #5 — DIE WAHL IST GEFALLEN',{x:nx+0.15,y:ny+1.48,w:nw-0.3,h:0.42,align:'center',fontSize:12,fontFace:'Calibri',bold:true,color:C.textMid});

  const opts=[
    ['🇵🇹  LISSABON',LIS_C,'Nightlife-Hauptstadt des Atlantiks\n17–22°C · Direktflug ab DUS'],
    ['🇧🇬  SOFIA',   SOF_C,'Günstigstes Ziel: €1 Bier\nVitosha · Clubs · Kathedrale'],
  ];
  opts.forEach(([lbl,col,sub],oi)=>{
    const py=ny+2.04+oi*1.38;
    s.addShape('rect',{x:nx+0.2,y:py,w:nw-0.4,h:0.56,fill:{color:col},line:{color:col},rounding:0.06});
    s.addText(lbl,{x:nx+0.2,y:py,w:nw-0.4,h:0.56,align:'center',fontSize:18,fontFace:'Calibri',bold:true,color:C.white,valign:'middle'});
    s.addText(sub,{x:nx+0.2,y:py+0.6,w:nw-0.4,h:0.64,align:'center',fontSize:9.5,fontFace:'Calibri',color:C.textSoft,wrap:true,valign:'top'});
  });

  s.addShape('rect',{x:nx+0.15,y:ny+4.96,w:nw-0.3,h:0.02,fill:{color:C.border},line:{color:C.border}});
  const facts=['📅  31. Oktober 2025 = Freitag','Jeder Trip unvergessen — Trip 5 wird der Beste'];
  facts.forEach((f,fi)=>{
    s.addText(f,{x:nx+0.1,y:ny+5.02+fi*0.5,w:nw-0.2,h:0.46,align:'center',fontSize:9.5,fontFace:'Calibri',
      color:fi===0?C.amber:C.textSoft,bold:fi===0,italic:fi===1,wrap:true});
  });

  bottomBar(s, C.amber);
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Trip-Planung 2025';
  pres.title  = 'Trip-Planung 2025 – Lissabon vs. Sofia';

  console.log('Erstelle Folien...');
  slide1(pres);                        console.log('  ✓ Folie  1: Titelfolie');

  const cities = [LISSABON, SOFIA];
  cities.forEach((cfg,ci) => {
    const base = (ci*5)+2;
    buildOverview(pres,cfg);    console.log(`  ✓ Folie ${String(base).padStart(2)}: ${cfg.name} Übersicht`);
    buildFlights(pres,cfg);     console.log(`  ✓ Folie ${String(base+1).padStart(2)}: ${cfg.name} Flüge`);
    buildAirbnb(pres,cfg);      console.log(`  ✓ Folie ${String(base+2).padStart(2)}: ${cfg.name} Unterkunft`);
    buildActivities(pres,cfg);  console.log(`  ✓ Folie ${String(base+3).padStart(2)}: ${cfg.name} Programm`);
    buildCosts(pres,cfg);       console.log(`  ✓ Folie ${String(base+4).padStart(2)}: ${cfg.name} Kosten`);
  });

  buildComparison(pres);  console.log('  ✓ Folie 12: Direktvergleich');
  buildHallOfFame(pres);  console.log('  ✓ Folie 13: Hall of Fame');

  await pres.writeFile({ fileName: OUT });
  console.log(`\n✅ Gespeichert: ${OUT}`);
}

main().catch(err => { console.error(err); process.exit(1); });
