'use strict';
const PptxGenJS = require('pptxgenjs');
const path = require('path');

const IMG = __dirname;
const OUT = path.join(__dirname, '..', 'halloween_trip_2025.pptx');

const C = {
  bg:'1A1A2E', bgCard:'16213E', bgDeep:'0F3460',
  orange:'FF8C00', amber:'FFC107', coral:'F96167',
  white:'FFFFFF', offWhite:'E8E8E8', gray:'9CA3AF',
  dimGray:'6B7280', green:'22C55E', gold:'FFD700', black:'000000',
};
const LIS_C = 'FF8C00'; // Orange (Lissabon)
const MIL_C = 'C8003C'; // Rosso Milano (Mailand)

// ── Helpers ───────────────────────────────────────────────────────────────────
function sectionHeader(s, title, accent) {
  s.addShape('rect',{x:0,y:0,w:'100%',h:0.08,fill:{color:accent},line:{color:accent}});
  s.addShape('rect',{x:0,y:0.08,w:'100%',h:0.72,fill:{color:C.bgCard},line:{color:C.bgCard}});
  s.addText(title,{x:0.4,y:0.08,w:12.5,h:0.72,fontSize:22,fontFace:'Trebuchet MS',bold:true,color:C.white,valign:'middle'});
  s.addShape('rect',{x:0,y:0.8,w:'100%',h:0.03,fill:{color:accent},line:{color:accent}});
}
function bottomBar(s) {
  s.addShape('rect',{x:0,y:7.42,w:'100%',h:0.08,fill:{color:C.orange},line:{color:C.orange}});
}
function statBox(s,x,y,w,h,label,val,accent) {
  s.addShape('rect',{x,y,w,h,fill:{color:C.bgDeep},line:{color:accent,pt:1.5},rounding:0.04});
  s.addText(label,{x,y:y+0.05,w,h:0.35,align:'center',fontSize:10,fontFace:'Calibri',color:C.gray});
  s.addText(val,{x,y:y+0.38,w,h:h-0.45,align:'center',fontSize:16,fontFace:'Trebuchet MS',bold:true,color:accent});
}
function dayCard(s,x,y,w,h,day,date,items,accent) {
  s.addShape('rect',{x,y,w,h,fill:{color:C.bgCard},line:{color:accent,pt:1},rounding:0.04});
  s.addShape('rect',{x,y,w,h:0.42,fill:{color:C.bgDeep},line:{color:C.bgDeep},rounding:0.04});
  s.addText(day,{x:x+0.08,y:y+0.02,w:w-0.16,h:0.20,fontSize:11,fontFace:'Trebuchet MS',bold:true,color:accent});
  s.addText(date,{x:x+0.08,y:y+0.22,w:w-0.16,h:0.18,fontSize:9,fontFace:'Calibri',color:C.gray});
  let ty=y+0.48;
  items.forEach(item=>{
    s.addText(item,{x:x+0.1,y:ty,w:w-0.2,h:0.28,fontSize:9,fontFace:'Calibri',color:C.offWhite,valign:'top',wrap:true});
    ty+=0.29;
  });
}

// ── Generic city slide builders ───────────────────────────────────────────────
function buildOverview(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  REISEZIEL ${cfg.num} — ${cfg.name.toUpperCase()}, ${cfg.country.toUpperCase()}`, cfg.accent);
  s.addImage({path:path.join(IMG,cfg.heroImg),x:0,y:0.83,w:7.0,h:5.4});
  s.addShape('rect',{x:5.8,y:0.83,w:1.2,h:5.4,fill:{color:C.bg},line:{color:C.bg},transparency:20});
  const rx=7.2;
  s.addText(cfg.tagline,{x:rx,y:0.95,w:5.9,h:0.45,fontSize:16,fontFace:'Trebuchet MS',bold:true,color:cfg.accent});
  s.addText(cfg.subtitle,{x:rx,y:1.38,w:5.9,h:0.32,fontSize:10,fontFace:'Calibri',color:C.gray,italic:true});
  let si=0;
  for(let row=0;row<2;row++) for(let col=0;col<3;col++){
    statBox(s,rx+col*1.98,1.85+row*1.35,1.85,1.2,cfg.stats[si][0],cfg.stats[si][1],cfg.accent);
    si++;
  }
  s.addShape('rect',{x:rx,y:4.65,w:5.9,h:1.55,fill:{color:C.bgCard},line:{color:C.bgDeep,pt:1},rounding:0.04});
  s.addText(cfg.description,{x:rx+0.15,y:4.75,w:5.6,h:1.35,fontSize:10.5,fontFace:'Calibri',color:C.offWhite,wrap:true,valign:'top'});
  bottomBar(s);
}

function buildFlights(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name.toUpperCase()} — FLÜGE`, cfg.accent);
  const hO={bold:true,color:cfg.accent,fontSize:11,fontFace:'Trebuchet MS',fill:{color:C.bgDeep},align:'center'};
  const rO={color:C.offWhite,fontSize:10,fontFace:'Calibri',fill:{color:C.bgCard},align:'center'};
  const rA={...rO,fill:{color:C.bg}};

  s.addText(`HINFLUG: Deutschland → ${cfg.iataLabel}`,{x:0.4,y:1.0,w:12.5,h:0.35,fontSize:13,fontFace:'Trebuchet MS',bold:true,color:cfg.accent});
  const hdr=['Option','Flug','Strecke & Datum','Abflug','Ankunft','Dauer','Preis p.P.','Gesamt (×4)'];
  const outRows=cfg.flightsOut.map((r,i)=>{
    const style=i%2===0?rO:rA;
    return [
      {text:r[0],options:i===0?{...style,color:C.green,bold:true}:style},
      {text:r[1],options:style},{text:r[2],options:style},
      {text:r[3],options:style},{text:r[4],options:style},{text:r[5],options:style},
      {text:r[6],options:style},{text:r[7],options:{...style,color:i===0?C.orange:C.amber,bold:i===0}},
    ];
  });
  s.addTable([[...hdr.map(h=>({text:h,options:hO}))],...outRows],{
    x:0.4,y:1.38,w:12.5,colW:[1.4,1.65,2.0,1.2,1.2,1.2,1.6,1.5],
    rowH:[0.38,...Array(outRows.length).fill(0.52)],
    border:{type:'solid',pt:1,color:C.bgDeep},
  });

  const ry=1.38+0.38+outRows.length*0.52+0.22;
  s.addText(`RÜCKFLUG: ${cfg.iataLabel} → Deutschland (2. Nov.)`,{x:0.4,y:ry,w:12.5,h:0.35,fontSize:13,fontFace:'Trebuchet MS',bold:true,color:cfg.accent});
  const retHdr=['Flug','Strecke & Datum','Abflug','Ankunft','Dauer','Preis p.P.','Gesamt (×4)'];
  const retRow=cfg.flightRet;
  s.addTable([
    [...retHdr.map(h=>({text:h,options:hO}))],
    [{text:retRow[0],options:{...rO,color:C.green}},{text:retRow[1],options:rO},{text:retRow[2],options:rO},{text:retRow[3],options:rO},{text:retRow[4],options:rO},{text:retRow[5],options:rO},{text:retRow[6],options:{...rO,color:C.amber,bold:true}}],
  ],{x:0.4,y:ry+0.37,w:12.5,colW:[2.5,2.0,1.3,1.3,1.3,1.8,1.6],rowH:[0.38,0.52],border:{type:'solid',pt:1,color:C.bgDeep}});

  const ny=ry+0.37+0.38+0.52+0.15;
  s.addShape('rect',{x:0.4,y:ny,w:12.5,h:0.72,fill:{color:C.bgCard},line:{color:C.bgDeep},rounding:0.04});
  s.addText(cfg.flightNote,{x:0.55,y:ny+0.05,w:12.2,h:0.62,fontSize:10,fontFace:'Calibri',color:C.gray,wrap:true,valign:'middle'});

  const py=ny+0.85;
  s.addShape('rect',{x:8.8,y:py,w:4.1,h:0.85,fill:{color:C.bgDeep},line:{color:cfg.accent,pt:2},rounding:0.06});
  s.addText('Hin & Zurück · 4 Personen',{x:8.8,y:py+0.02,w:4.1,h:0.3,align:'center',fontSize:10,fontFace:'Calibri',color:C.gray});
  s.addText(cfg.flightTotal,{x:8.8,y:py+0.32,w:4.1,h:0.48,align:'center',fontSize:20,fontFace:'Trebuchet MS',bold:true,color:cfg.accent});
  bottomBar(s);
}

function buildAirbnb(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name.toUpperCase()} — UNTERKUNFT`, cfg.accent);
  s.addImage({path:path.join(IMG,cfg.airbnbImg),x:0,y:0.83,w:5.5,h:3.4});
  s.addImage({path:path.join(IMG,cfg.mapImg),x:0,y:4.28,w:5.5,h:2.3});
  const rx=5.8;
  s.addShape('rect',{x:rx,y:0.9,w:7.3,h:0.55,fill:{color:C.bgDeep},line:{color:cfg.accent,pt:1.5},rounding:0.04});
  s.addText(cfg.airbnbTitle,{x:rx+0.15,y:0.9,w:7.0,h:0.55,fontSize:14,fontFace:'Trebuchet MS',bold:true,color:C.white,valign:'middle'});
  cfg.airbnbSpecs.forEach((spec,i)=>{
    const bx=rx+(i%3)*2.45; const by=1.56+Math.floor(i/3)*0.55;
    const col=spec[2]||cfg.accent;
    s.addShape('rect',{x:bx,y:by,w:2.3,h:0.42,fill:{color:C.bgDeep},line:{color:col,pt:1},rounding:0.04});
    s.addText(spec[0],{x:bx,y:by,w:2.3,h:0.42,align:'center',fontSize:11,fontFace:'Calibri',bold:true,color:col,valign:'middle'});
  });
  s.addShape('rect',{x:rx,y:2.74,w:7.3,h:0.55,fill:{color:C.bgCard},line:{color:C.bgDeep},rounding:0.04});
  s.addText(cfg.airbnbLocation,{x:rx+0.1,y:2.74,w:7.1,h:0.55,fontSize:10.5,fontFace:'Calibri',color:C.gray,valign:'middle'});
  s.addShape('rect',{x:rx,y:3.4,w:7.3,h:1.5,fill:{color:C.bgCard},line:{color:cfg.accent,pt:1.5},rounding:0.04});
  s.addText('PREISE  (30. Okt. – 2. Nov., 3 Nächte)',{x:rx+0.2,y:3.46,w:7.0,h:0.32,fontSize:11,fontFace:'Trebuchet MS',bold:true,color:cfg.accent});
  cfg.airbnbPricing.forEach((row,i)=>{
    const py=3.82+i*0.215;
    s.addText(row[0],{x:rx+0.2,y:py,w:4.0,h:0.21,fontSize:10,fontFace:'Calibri',color:C.gray});
    const isTotal=i===3; const isPP=i===4;
    s.addText(row[1],{x:rx+4.5,y:py,w:2.6,h:0.21,align:'right',fontSize:10,fontFace:'Trebuchet MS',bold:isTotal||isPP,color:isTotal?cfg.accent:isPP?C.green:C.offWhite});
  });
  // Primary search button
  s.addShape('rect',{x:rx,y:5.03,w:7.3,h:0.66,fill:{color:cfg.accent},line:{color:cfg.accent},rounding:0.06});
  s.addText('🔍  Airbnb: ' + cfg.name + ' durchsuchen  →',{x:rx,y:5.03,w:7.3,h:0.66,align:'center',fontSize:15,fontFace:'Trebuchet MS',bold:true,color:C.white,valign:'middle',hyperlink:{url:cfg.airbnbSearchUrl,tooltip:'Airbnb öffnen — 30. Okt.–2. Nov., 4 Gäste, 2 SZ'}});
  // Alt hint
  s.addText(cfg.airbnbAlt,{x:rx+0.1,y:5.77,w:7.1,h:0.28,fontSize:9,fontFace:'Calibri',color:C.dimGray,italic:true,hyperlink:{url:cfg.airbnbAltUrl,tooltip:'Alternative Suche öffnen'}});
  s.addText('⚠️  Halloween-Wochenende ist schnell ausgebucht — jetzt klicken & buchen!',{x:rx+0.1,y:6.1,w:7.1,h:0.3,fontSize:9.5,fontFace:'Calibri',bold:true,color:C.coral});
  bottomBar(s);
}

function buildActivities(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name.toUpperCase()} — TAGESPROGRAMM`, cfg.accent);
  cfg.days.forEach((d,i)=>{
    const col=i%2; const row=Math.floor(i/2);
    dayCard(s,0.3+col*6.5,0.98+row*3.15,6.2,3.0,d.day,d.date,d.items,cfg.accent);
  });
  s.addShape('rect',{x:0.3,y:7.12,w:12.7,h:0.22,fill:{color:C.bgCard},line:{color:C.bgDeep}});
  s.addText(cfg.activityLinks,{x:0.4,y:7.12,w:12.5,h:0.22,fontSize:8,fontFace:'Calibri',color:C.dimGray,valign:'middle'});
  bottomBar(s);
}

function buildCosts(pres, cfg) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, `${cfg.flag}  ${cfg.name.toUpperCase()} — KOSTENÜBERSICHT`, cfg.accent);
  const hO={bold:true,color:cfg.accent,fontSize:12,fontFace:'Trebuchet MS',fill:{color:C.bgDeep},align:'center',valign:'middle'};
  const rO={color:C.offWhite,fontSize:11,fontFace:'Calibri',fill:{color:C.bgCard},align:'left',valign:'middle'};
  const rA={...rO,fill:{color:C.bg}};
  const rN={...rO,align:'right'}; const rNA={...rA,align:'right'};
  const tot={bold:true,color:cfg.accent,fontSize:13,fontFace:'Trebuchet MS',fill:{color:C.bgDeep},align:'right',valign:'middle'};
  const totL={...tot,align:'left',color:C.white};
  const rows=[
    [{text:'Posten',options:hO},{text:'Details',options:hO},{text:'p.P.',options:hO},{text:'Gesamt ×4',options:{...hO,color:C.orange}}],
    ...cfg.costRows.map((r,i)=>{
      const isLast=i===cfg.costRows.length-1;
      if(isLast) return [{text:r[0],options:totL},{text:r[1],options:{...tot,align:'left',color:C.gray}},{text:r[2],options:{...tot,color:cfg.accent}},{text:r[3],options:{...tot,color:C.orange}}];
      const st=i%2===0?rO:rA; const stn=i%2===0?rN:rNA;
      return [{text:r[0],options:st},{text:r[1],options:st},{text:r[2],options:stn},{text:r[3],options:stn}];
    }),
  ];
  const nData=cfg.costRows.length;
  s.addTable(rows,{x:0.5,y:1.02,w:12.3,colW:[2.5,5.2,2.1,2.1],rowH:[0.42,...Array(nData).fill(0.47)],border:{type:'solid',pt:1,color:C.bgDeep}});
  const ty=1.02+0.42+nData*0.47+0.08;
  s.addShape('rect',{x:0.5,y:ty,w:12.3,h:0.65,fill:{color:C.bgCard},line:{color:cfg.accent,pt:1},rounding:0.04});
  s.addText(cfg.costNote,{x:0.65,y:ty+0.05,w:12.0,h:0.55,fontSize:10.5,fontFace:'Calibri',color:C.gray,wrap:true,valign:'middle'});
  bottomBar(s);
}

// ── CITY DATA ────────────────────────────────────────────────────────────────

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
    {day:'TAG 1 — DO., 30. OKT.',date:'Ankunftstag',items:[
      '✈  EW9602 landet LIS 12:20 · Einchecken in Bairro Alto',
      '🚶  SANDEMANs Stadtführung 15:30 @ Praça do Comércio',
      '   (2,5h, Trinkgeld €10–15 · Baixa, Alfama, Aussichtspunkte)',
      '🍽  Abendessen: Petiscos bei Tasca do Chico ab 20 Uhr',
      '🍺  Bairro Alto Kneipenbummel: 200+ Bars, €2–3 Bier',
      '   Start @ Park Bar (Dachterrasse) → Pavilhão Chinês',
    ]},
    {day:'TAG 2 — FR., 31. OKT. 🎃',date:'HALLOWEEN-NACHT',items:[
      '🍳  Kochkurs Cooking Lisbon ab 9:30 Uhr (4h, €110 p.P.)',
      '   Mercado 31 de Janeiro · 3-Gang-Menü + Wein inklusive',
      '   Buchung: cookinglisbon.com/market-tour',
      '🕙  Halloween-Kneipentour 22 Uhr · €28 p.P.',
      '   Treffpkt.: Praça Luís de Camões · 3 Bars · 1h offene Bar',
      '   Shots · Kostüm-Wettbewerb · VIP-Clubeintritt',
      '🎶  Ab 1 Uhr: Brunch Electronik Halloween Special',
      '   Pavilhão Carlos Lopes · Chris Liebing B2B Luke Slater',
    ]},
    {day:'TAG 3 — SA., 1. NOV.',date:'Erkunden & Entspannen',items:[
      '🌅  Ausschlafen — Pastéis de Nata bei Pasteis de Belém',
      '🚋  Straßenbahn 28: malerische Fahrt durch Alfama',
      '🍺  Cerveja Musa Craft-Brauerei, Mouraria (lokale IPAs)',
      '🌇  Miradouro da Graça: bester Sonnenuntergangs-Aussichtsp.',
      '🎵  Fado-Abend: Tasca do Chico oder Zé da Viola',
      '🌃  Pink Street Nacht: Pensão Amor + Music Box · Lux Frágil',
    ]},
    {day:'TAG 4 — SO., 2. NOV.',date:'Abreisetag',items:[
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
    ['🏆  GESAMTSCHÄTZUNG','Lissabon, 30. Okt. – 2. Nov. 2025','~€735','~€2.940'],
  ],
  costNote:'💡 Budgetrahmen: €600–900 p.P. je nach Partylust.  Bier ab €2–3, Supermarkt-Runs helfen sparen.  Clubs (Lux Frágil, Brunch Electronik) kosten je ~€20–30 extra.  Leitungswasser in Restaurants gratis.',
};

const MAILAND = {
  num:2, name:'Mailand', country:'Italien', flag:'🇮🇹', accent:MIL_C,
  heroImg:'milan_hero.png', mapImg:'milan_map.png', airbnbImg:'milan_airbnb.png',
  iataLabel:'MXP (Mailand-Malpensa)',
  tagline:'Aperitivo, Mode & Nachtleben an den Navigli',
  subtitle:'Navigli · Corso Como · Duomo · Brera · Clubs bis 5 Uhr',
  stats:[
    ['✈ Flugdauer','1h 40m\nab DUS'],['🍺 Bier','€5–7\n(Nastro Azzurro)'],
    ['🌡 Wetter','12–16°C\nEnde Oktober'],['💰 Kosten ca.','~€700\npro Person'],
    ['🏙 Nightlife','⭐⭐⭐⭐\nNavigli & Corso Como'],['📍 Ab Flughafen','DUS direkt!\n(Eurowings/ITA)'],
  ],
  description:'Mailand ist mit nur 1h 40m Flugzeit das nächste Ziel — Direktflug ab DUS mit Eurowings. Das Navigli-Viertel mit seinen Kanälen ist die Aperitivo-Hochburg Italiens: ab 18 Uhr Drinks + gratis Buffet für ~€10. Corso Como und die Isola-Szene bieten Clubs bis 5 Uhr. Das Duomo di Milano leuchtet nachts orange — perfekt für Halloween-Fotos. Und die Mailänder feiern Halloween ausgiebig.',
  flightsOut:[
    ['⭐ Empfohlen','EW\nEurowings','DUS → MXP\n30. Okt. früh','~07:00','~08:40','1h 40m','~€80','~€320'],
    ['Vorabend','EW\nEurowings','DUS → MXP\n29. Okt. abends','~18:xx','~19:40','1h 40m','~€90','~€360'],
    ['Mittag','ITA Airways\n(ex-Alitalia)','DUS → MXP\n30. Okt.','~12:xx','~13:50','1h 40m','~€100–130','~€400–520'],
    ['Budget','FR\nRyanair','DUS → BGY\n30. Okt.','Verschieden','Verschieden','1h 45m','~€50–80','~€200–320'],
  ],
  flightRet:['Eurowings / Ryanair  (MXP/BGY→DUS)','MXP → DUS  /  2. Nov.','Verschieden','Verschieden','1h 40m','~€70–100','~€280–400'],
  flightNote:'✅  DUS (Düsseldorf) → MXP (Mailand-Malpensa): Direktflug mit Eurowings täglich, ~1h 40m.  Malpensa liegt 45 km nordwestlich vom Zentrum — Malpensa Express Zug ~52 Min., €13 ins Stadtzentrum.  Budget-Option: DUS → BGY (Bergamo/Orio al Serio) mit Ryanair — Bus nach Mailand ~1h, €5–6.  Buchen: eurowings.com · ryanair.com',
  flightTotal:'€480 – €640 gesamt',
  airbnbTitle:'Navigli / Porta Ticinese — 2-Zi.-Wohnung am Kanal',
  airbnbSpecs:[
    ['🛏 2 Schlafzimmer'],['🚿 2 Bäder'],['🏗 Balkon/Terrasse',[],C.green],
    ['👥 4 Gäste'],['❄ Klimaanlage ✓'],['🍹 Navigli Bars 1 Min',[],MIL_C],
  ],
  airbnbLocation:'📌 Navigli — direkt an der Aperitivo-Meile, 15 Min. per Metro/Tram zum Duomo & Corso Como',
  airbnbPricing:[['Pro Nacht','~€130'],['3 Nächte','~€390'],['Reinigung & Gebühren','~€60–80'],['Gesamt für 4 Pers.','~€450–470'],['Pro Person','~€113–118']],
  airbnbSearchUrl:'https://www.airbnb.com/s/Navigli--Milan--Italy/homes?checkin=2025-10-30&checkout=2025-11-02&adults=4&min_bedrooms=2&room_types%5B%5D=Entire+home%2Fapt',
  airbnbAlt:'📍 Alternativ in Brera/Isola suchen: airbnb.com/s/Brera--Milan (Designviertel, gleiche Daten)',
  airbnbAltUrl:'https://www.airbnb.com/s/Brera--Milan--Italy/homes?checkin=2025-10-30&checkout=2025-11-02&adults=4&min_bedrooms=2&room_types%5B%5D=Entire+home%2Fapt',
  days:[
    {day:'TAG 1 — DO., 30. OKT.',date:'Ankunftstag',items:[
      '✈  Eurowings landet MXP ~08:40 · Malpensa Express ins Zentrum',
      '🏠  Einchecken in Navigli (Check-in ab 15 Uhr)',
      '🚶  SANDEMANs Stadtführung 14:30 @ Duomo (2,5h, Trinkgeld)',
      '   Duomo, Galleria Vittorio Emanuele, Castello Sforzesco',
      '🍹  Aperitivo ab 18 Uhr am Naviglio Grande',
      '   €8–10 Drink + gratis Buffet — Mailänder Tradition!',
      '🌃  Erste Nacht: Navigli-Bars & Darsena-Platz',
    ]},
    {day:'TAG 2 — FR., 31. OKT. 🎃',date:'HALLOWEEN-NACHT',items:[
      '🍳  Pasta-Kochkurs 10:00 Uhr (3h, €70 p.P., Navigli)',
      '   Risotto alla Milanese + Ossobuco + Tiramisu',
      '🛍  Nachmittags: Brera Designviertel oder Modeboulevard',
      '🎃  Halloween Kneipentour Navigli ab 21 Uhr · ~€20 p.P.',
      '   Treffpunkt Darsena · 4 Bars · Shots · Kostüm-Contest',
      '🎶  Ab 1 Uhr: Alcatraz oder Volt Club — Halloween Special',
      '   Corso Como Party bis 5 Uhr morgens',
    ]},
    {day:'TAG 3 — SA., 1. NOV.',date:'Erkunden & Ausgehen',items:[
      '🌅  Ausschlafen — Mailands Nächte enden spät',
      '☕  Frühstück: Cappuccino + Cornetto an der Bar (€2)',
      '🏛  Duomo Dachterrasse (Blick über ganz Mailand, ~€14)',
      '🛒  Mercato Metropolitano oder Fiera di Sinigaglia',
      '🍝  Abendessen: Risotto Milanese in der Navigli-Trattoria',
      '🌃  Letzte Nacht: Isola-Viertel Bars + Corso Como',
    ]},
    {day:'TAG 4 — SO., 2. NOV.',date:'Abreisetag',items:[
      '☕  Letzter Cappuccino an der Bar — nie im Sitzen!',
      '🛍  Souvenirs: Panettone, Amaro, Design-Mitbringsel',
      '🚂  Malpensa Express zum Flughafen (52 Min.)',
      '✈  Eurowings MXP → DUS (Nachmittag)',
    ]},
  ],
  activityLinks:'Halloween Kneipentour: eventbrite.it → "Milan Halloween Bar Crawl"  ·  Kochkurs: cookly.app → "Milan Pasta Class"  ·  Stadtführung: neweuropetours.eu/milan  ·  Alcatraz Club: alcatrazmilano.com',
  costRows:[
    ['Hin- und Rückflug','Eurowings EW  DUS ↔ MXP (Direktflug)','~€170','~€680'],
    ['Unterkunft','Airbnb Navigli — 3 Nächte (inkl. Gebühren)','~€115','~€460'],
    ['Halloween Kneipentour','Bar Crawl Navigli / Darsena (31. Okt.)','€20','€80'],
    ['Kochkurs','Pasta-Kochkurs Navigli — Risotto + Ossobuco','€70','€280'],
    ['Stadtführung','SANDEMANs — Duomo (Trinkgeld)','~€12','~€48'],
    ['Essen (3 Tage)','Trattoria, Aperitivo-Buffet, Bar-Frühstück','~€130','~€520'],
    ['Getränke & Ausgehen','Navigli + Corso Como + Clubs (3 Nächte)','~€160','~€640'],
    ['Sonstiges','Malpensa Express, Tram, Souvenirs','~€30','~€120'],
    ['🏆  GESAMTSCHÄTZUNG','Mailand, 30. Okt. – 2. Nov. 2025','~€707','~€2.828'],
  ],
  costNote:'💡 Kürzester Flug dieser Auswahl (1h 40m ab DUS)!  Aperitivo-Zeit (18–21 Uhr): Drink + gratis Buffet ~€10 — spart Abendessenkosten.  Cappuccino an der Bar stehend: €1,50.  Budgetrahmen: €600–850 p.P.',
};

// ── Slide 1: Titelfolie ───────────────────────────────────────────────────────
function slide1(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  s.addImage({path:path.join(IMG,'title_bg.png'),x:0,y:0,w:'100%',h:'100%'});
  s.addShape('rect',{x:0,y:0,w:'100%',h:'100%',fill:{color:'000000',transparency:42},line:{color:'000000'}});
  s.addShape('rect',{x:0,y:0,w:'100%',h:0.1,fill:{color:C.orange},line:{color:C.orange}});
  s.addText('🎃',{x:0,y:0.65,w:'100%',h:1.0,align:'center',fontSize:72,fontFace:'Segoe UI Emoji'});
  s.addText('HALLOWEEN-TRIP 2025',{x:0.5,y:1.5,w:12.33,h:1.2,align:'center',fontSize:60,fontFace:'Trebuchet MS',bold:true,color:C.orange});
  s.addText('30. Okt. – 2. Nov.  ·  4 Jungs  ·  Europäische Hauptstadt  ·  Halloween-Nacht',{x:0.5,y:2.78,w:12.33,h:0.55,align:'center',fontSize:19,fontFace:'Trebuchet MS',color:C.offWhite});
  s.addShape('rect',{x:2.0,y:3.55,w:9.33,h:0.6,fill:{color:C.bgCard,transparency:20},line:{color:C.orange,pt:1},rounding:0.05});
  s.addText('Bereits bereist:   ✓ Riga   ✓ Athen   ✓ Tirana   ✓ Belgrad',{x:2.0,y:3.55,w:9.33,h:0.6,align:'center',fontSize:15,fontFace:'Calibri',color:C.gray,valign:'middle'});
  s.addText('DIE ZIELE STEHEN FEST — ES GEHT LOS!',{x:0.5,y:4.35,w:12.33,h:0.65,align:'center',fontSize:26,fontFace:'Trebuchet MS',bold:true,color:C.coral});
  const pills=[['🇵🇹 LISSABON',LIS_C],['🇮🇹 MAILAND',MIL_C]];
  pills.forEach(([label,col],i)=>{
    const px=2.17+i*4.5;
    s.addShape('rect',{x:px,y:5.15,w:4.2,h:0.75,fill:{color:C.bgDeep},line:{color:col,pt:2.5},rounding:0.08});
    s.addText(label,{x:px,y:5.15,w:4.2,h:0.75,align:'center',fontSize:24,fontFace:'Trebuchet MS',bold:true,color:C.white,valign:'middle'});
  });
  s.addText('Beide Ziele: Direktflug ab DUS  ·  Lissabon 17–22°C · Mailand 12–16°C',{x:0,y:6.25,w:'100%',h:0.4,align:'center',fontSize:13,fontFace:'Calibri',color:C.dimGray,italic:true});
  s.addShape('rect',{x:0,y:7.42,w:'100%',h:0.08,fill:{color:C.orange},line:{color:C.orange}});
}

// ── Slide 12: Direktvergleich (2 Städte) ─────────────────────────────────────
function buildComparison(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '⚔️  DIREKTVERGLEICH — LISSABON vs. MAILAND', C.orange);

  const catW=3.4, cityW=4.55;
  const cols=[
    {label:'🇵🇹 LISSABON', color:LIS_C, x:catW+0.35},
    {label:'🇮🇹 MAILAND',  color:MIL_C, x:catW+0.35+cityW+0.1},
  ];

  cols.forEach(col=>{
    s.addShape('rect',{x:col.x,y:0.85,w:cityW,h:0.56,fill:{color:C.bgDeep},line:{color:col.color,pt:2},rounding:0.05});
    s.addText(col.label,{x:col.x,y:0.85,w:cityW,h:0.56,align:'center',fontSize:18,fontFace:'Trebuchet MS',bold:true,color:col.color,valign:'middle'});
  });

  const rows=[
    ['💰 Kosten p.P.',      '~€735',                          '~€707'],
    ['✈ Flugdauer',         '3h 10m  ab DUS',                 '1h 40m  ab DUS  ✈ schnellste!'],
    ['🛫 Flughafen',        'DUS — Eurowings / TAP',          'DUS — Eurowings / ITA / Ryanair'],
    ['🍺 Bierpreis',        '€2–3  (Super Bock / Sagres)',    '€5–7  (Nastro Azzurro)'],
    ['🌡 Wetter Ende Okt.', '17–22°C  ☀',                    '12–16°C  🧥'],
    ['🎃 Halloween',        '⭐⭐⭐⭐  Clubs bis 6 Uhr',       '⭐⭐⭐⭐  Corso Como bis 5 Uhr'],
    ['🏙 Viertel',          'Bairro Alto · Pink Street',      'Navigli · Corso Como · Isola'],
    ['🌟 Highlight',        'Fado · Atlantik-Charme',         'Aperitivo + gratis Buffet ab 18h'],
    ['🍳 Kochkurs',         '€110 p.P.',                      '€70 p.P.'],
    ['🏠 Unterkunft p.P.',  '~€113–120',                      '~€113–118'],
    ['💡 Stärke',           'Nightlife-Weltklasse + Wärme',   'Kürzester Flug + Fashion + Pasta'],
  ];

  const rowH=0.5;
  rows.forEach((row,i)=>{
    const ry=1.45+i*rowH;
    const isLast=i===rows.length-1;
    const rowFill=isLast?C.bgDeep:(i%2===0?C.bgCard:C.bg);

    s.addShape('rect',{x:0.3,y:ry,w:catW,h:rowH,fill:{color:rowFill},line:{color:C.bgDeep}});
    s.addText(row[0],{x:0.4,y:ry,w:catW-0.15,h:rowH,fontSize:isLast?11:10,fontFace:'Calibri',bold:isLast,color:isLast?C.amber:C.gray,valign:'middle'});

    cols.forEach((col,ci)=>{
      s.addShape('rect',{x:col.x,y:ry,w:cityW,h:rowH,fill:{color:rowFill},line:{color:isLast?col.color:C.bgDeep,pt:isLast?1.5:0.5}});
      s.addText(row[ci+1],{x:col.x+0.1,y:ry,w:cityW-0.2,h:rowH,align:'center',fontSize:isLast?11:10,fontFace:isLast?'Trebuchet MS':'Calibri',bold:isLast,color:isLast?col.color:C.offWhite,valign:'middle'});
    });
  });
  bottomBar(s);
}

// ── Slide 13: Hall of Fame ────────────────────────────────────────────────────
function buildHallOfFame(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  sectionHeader(s, '🏆  TRIP HALL OF FAME — DIE LEGENDE WÄCHST', C.orange);

  // Left 2×2 grid: past trips
  const past = [
    { year:'2021', city:'Riga',    flag:'🇱🇻', color:'E8C85A',
      lines:['🍺 Bier: €0,80 — ewiger Rekord', '🌃 Clubs bis 7 Uhr morgens', '🏆 Günstigster Trip aller Zeiten'] },
    { year:'2022', city:'Athen',   flag:'🇬🇷', color:'5A8FE8',
      lines:['🌅 Dachbar mit Akropolis-Blick', '🐟 Ouzo-Entscheidung: bereut', '🏆 Wärmster Trip bisher'] },
    { year:'2023', city:'Tirana',  flag:'🇦🇱', color:'E85A5A',
      lines:['🥘 Bestes Essen: unerwartet gut', '🎵 Barkeeper wurde Freund', '🏆 Geheimtipp des Jahres'] },
    { year:'2024', city:'Belgrad', flag:'🇷🇸', color:'A55AE8',
      lines:['⛵ Splavovi-Clubs am Fluss', '🌄 Heimkehr beim Sonnenaufgang', '🏆 Wildeste Nacht aller Zeiten'] },
  ];

  const cW=3.6, cH=2.82, gX=0.2, gY=0.18;
  past.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.25 + col*(cW+gX);
    const y = 1.02 + row*(cH+gY);

    s.addShape('rect',{x,y,w:cW,h:cH,fill:{color:C.bgCard},line:{color:t.color,pt:1.5},rounding:0.06});
    s.addShape('rect',{x,y,w:cW,h:0.46,fill:{color:C.bgDeep},line:{color:C.bgDeep},rounding:0.06});
    s.addText(`${t.flag}  ${t.year}`,{x,y,w:cW,h:0.46,align:'center',fontSize:17,fontFace:'Trebuchet MS',bold:true,color:t.color,valign:'middle'});
    s.addText(t.city.toUpperCase(),{x,y:y+0.5,w:cW,h:0.4,align:'center',fontSize:15,fontFace:'Trebuchet MS',bold:true,color:C.white});
    s.addShape('rect',{x:x+0.15,y:y+0.96,w:cW-0.3,h:0.025,fill:{color:t.color},line:{color:t.color}});
    t.lines.forEach((line, li) => {
      const isLast = li === t.lines.length - 1;
      s.addText(line,{x:x+0.1,y:y+1.0+li*0.56,w:cW-0.2,h:0.52,fontSize:9.5,fontFace:'Calibri',
        color:isLast?t.color:C.offWhite,bold:isLast,wrap:true,valign:'top'});
    });
  });

  // Right side: 2025 Next Stop (x=7.9, fits within 13.333")
  const nx=7.9, nw=5.15, ny=1.0, nh=6.28;
  s.addShape('rect',{x:nx,y:ny,w:nw,h:nh,fill:{color:C.bgDeep},line:{color:C.orange,pt:2.5},rounding:0.08});
  s.addShape('rect',{x:nx,y:ny,w:nw,h:0.6,fill:{color:C.orange},line:{color:C.orange},rounding:0.08});
  s.addText('2025  ·  NEXT STOP',{x:nx,y:ny,w:nw,h:0.6,align:'center',fontSize:16,fontFace:'Trebuchet MS',bold:true,color:C.bg,valign:'middle'});

  s.addText('🎃',{x:nx,y:ny+0.68,w:nw,h:0.7,align:'center',fontSize:44,fontFace:'Segoe UI Emoji'});
  s.addText('TRIP #5 — DIE WAHL IST GEFALLEN',{x:nx+0.15,y:ny+1.46,w:nw-0.3,h:0.44,align:'center',fontSize:12,fontFace:'Trebuchet MS',bold:true,color:C.amber});

  const opts=[['🇵🇹  LISSABON',LIS_C,'Nightlife-Hauptstadt\ndes Atlantiks · 17–22°C'],
              ['🇮🇹  MAILAND', MIL_C,'Kürzester Flug (1h40)\nNavigli · Aperitivo · Mode']];
  opts.forEach(([lbl,col,sub],oi)=>{
    const py = ny+2.02+oi*1.38;
    s.addShape('rect',{x:nx+0.2,y:py,w:nw-0.4,h:0.58,fill:{color:C.bgCard},line:{color:col,pt:2},rounding:0.06});
    s.addText(lbl,{x:nx+0.2,y:py,w:nw-0.4,h:0.58,align:'center',fontSize:18,fontFace:'Trebuchet MS',bold:true,color:col,valign:'middle'});
    s.addText(sub,{x:nx+0.2,y:py+0.62,w:nw-0.4,h:0.62,align:'center',fontSize:9.5,fontFace:'Calibri',color:C.gray,wrap:true,valign:'top'});
  });

  s.addShape('rect',{x:nx+0.15,y:ny+4.94,w:nw-0.3,h:0.025,fill:{color:C.orange},line:{color:C.orange}});
  const facts=['🎃  31. Oktober 2025 = FREITAG','Jeder Trip unvergessen — Trip 5 wird der Beste'];
  facts.forEach((f,fi)=>{
    s.addText(f,{x:nx+0.1,y:ny+5.0+fi*0.52,w:nw-0.2,h:0.48,align:'center',fontSize:9.5,fontFace:'Calibri',
      color:fi===0?C.orange:C.gray,bold:fi===0,italic:fi===1,wrap:true});
  });

  bottomBar(s);
}

// ── MAIN ─────────────────────────────────────────────────────────────────────
async function main() {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Halloween-Trip 2025';
  pres.title = 'Halloween-Trip 2025 – Lissabon vs. Mailand';

  console.log('Erstelle Folien...');
  slide1(pres); console.log('  ✓ Folie 1: Titelfolie');

  const cities=[LISSABON,MAILAND];
  cities.forEach((cfg,ci)=>{
    const base=(ci*5)+2;
    buildOverview(pres,cfg);   console.log(`  ✓ Folie ${base}: ${cfg.name} Übersicht`);
    buildFlights(pres,cfg);    console.log(`  ✓ Folie ${base+1}: ${cfg.name} Flüge`);
    buildAirbnb(pres,cfg);     console.log(`  ✓ Folie ${base+2}: ${cfg.name} Unterkunft`);
    buildActivities(pres,cfg); console.log(`  ✓ Folie ${base+3}: ${cfg.name} Programm`);
    buildCosts(pres,cfg);      console.log(`  ✓ Folie ${base+4}: ${cfg.name} Kosten`);
  });

  buildComparison(pres);   console.log('  ✓ Folie 12: Direktvergleich');
  buildHallOfFame(pres);   console.log('  ✓ Folie 13: Hall of Fame');

  await pres.writeFile({ fileName: OUT });
  console.log(`\n✅ Gespeichert: ${OUT}`);
}

main().catch(err => { console.error(err); process.exit(1); });
