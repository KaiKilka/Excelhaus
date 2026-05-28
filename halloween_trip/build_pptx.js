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
const LIS_C  = 'FF8C00';
const MAD_C  = 'E6321E';
const ROM_C  = 'D4783A';
const MALT_C = '1EB4C8';

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
  s.addShape('rect',{x:rx,y:5.03,w:7.3,h:0.52,fill:{color:C.bgDeep},line:{color:C.bgDeep},rounding:0.04});
  s.addText(cfg.airbnbUrl,{x:rx+0.15,y:5.03,w:7.0,h:0.52,fontSize:9.5,fontFace:'Calibri',color:cfg.accent,italic:true,valign:'middle'});
  s.addText(cfg.airbnbAlt,{x:rx+0.1,y:5.65,w:7.1,h:0.32,fontSize:9.5,fontFace:'Calibri',color:C.dimGray,italic:true});
  s.addText('🔑  Jetzt buchen — Halloween-Wochenende ist schnell ausgebucht!',{x:rx+0.1,y:6.05,w:7.1,h:0.32,fontSize:9.5,fontFace:'Calibri',bold:true,color:C.coral});
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
  airbnbUrl:'airbnb.com/rooms/26081783  (Bairro Alto / São Roque — Verfügbarkeit prüfen)',
  airbnbAlt:'Alternative: "Garrett 48" in Chiado — Terrasse, 2 SZ, ~€100–130/Nacht → airbnb.com/rooms/1016061',
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

const MADRID = {
  num:2, name:'Madrid', country:'Spanien', flag:'🇪🇸', accent:MAD_C,
  heroImg:'madrid_hero.png', mapImg:'madrid_map.png', airbnbImg:'madrid_airbnb.png',
  iataLabel:'MAD (Madrid)',
  tagline:'Die Stadt, die niemals schläft — buchstäblich',
  subtitle:'Malasaña · Chueca · Gran Vía · La Latina · Clubs bis 7 Uhr morgens',
  stats:[
    ['✈ Flugdauer','2h 30m\nab DUS'],['🍺 Bier','€2 – €4\n(Caña/Tubo)'],
    ['🌡 Wetter','18–20°C ☀\nEnde Oktober'],['💰 Kosten ca.','~€712\npro Person'],
    ['🏙 Nightlife','⭐⭐⭐⭐⭐\nClubs bis 7 Uhr'],['📍 Ab Flughafen','Nur DUS\n(Iberia/Eurowings)'],
  ],
  description:'Madrid ist das wärmste Halloween Europas. Der 31. Oktober ist ein Freitag UND der 1. November ist Allerheiligen — gesetzlicher Feiertag in Spanien. Heißt: alle sind aus, Clubs offen bis 7 Uhr morgens. Malasañas Boheme-Bar-Szene, kostenlose Tapas bei El Tigre, Vermouth vom Fass in 150 Jahre alten Bodegas, und 7-stöckiger Teatro Kapital Club.',
  flightsOut:[
    ['⭐ Vorabend\n29. Okt.','Iberia\n(IB3166)','DUS → MAD\n29. Okt. abends','~19:55','~22:35','2h 30m','~€95','~€380'],
    ['Frühabflug\n30. Okt.','Iberia\n(IB3162)','DUS → MAD\n30. Okt. morgens','~06:20','~08:50','2h 30m','~€75','~€300'],
    ['Mittag\n30. Okt.','Eurowings\n(EW)','DUS → MAD\n30. Okt. mittags','~13:xx','~15:30','2h 30m','~€90','~€360'],
    ['Premium','Vueling\n(VY)','DUS → MAD\n30. Okt.','Verschieden','Verschieden','2h 30m','€100–130','€400–520'],
  ],
  flightRet:['Iberia / Eurowings  (MAD→DUS)','MAD → DUS  /  2. Nov.','Verschieden','Verschieden','2h 40m','~€70–90','~€280–360'],
  flightNote:'ℹ  DUS (Düsseldorf) ist der empfohlene Abflughafen — Direktflüge mit Iberia (3×/Tag), Eurowings & Vueling.  DTM hat keine Direktverbindung nach Madrid.  FMO erfordert Umstieg.  Flughafen zum Stadtzentrum: Metro Linie 8 (€5) oder Taxi (~€35 Festpreis).',
  flightTotal:'€580 – €740 gesamt',
  airbnbTitle:'Sonder Malasaña — 2 SZ · Balkon · Bars vor der Haustür',
  airbnbSpecs:[
    ['🛏 2 Schlafzimmer'],['🚿 2 Bäder'],['🏗 Balkon ✓',[],C.green],
    ['👥 4 Gäste'],['❄ Klimaanlage ✓'],['🎸 Malasaña-Bars 1 Min'],
  ],
  airbnbLocation:'📌 Malasaña — 1 Min zu Calle de la Palma, 5 Min zur Gran Vía, 10 Min zur Puerta del Sol',
  airbnbPricing:[['Pro Nacht (Sonder-verwaltet)','~€150'],['3 Nächte','~€450'],['Reinigung & Gebühren','~€60–80'],['Gesamt für 4 Pers.','~€510–530'],['Pro Person','~€128–133']],
  airbnbUrl:'airbnb.com/rooms/47308216  (Sonder Malasaña — professionell verwaltet, 24/7-Support)',
  airbnbAlt:'Budget-Alternative: 2-Zi. zentral Madrid ~€80–110/Nacht → ~€310 Gesamt → ~€78 p.P.',
  days:[
    {day:'TAG 1 — MI., 29. OKT.',date:'Ankunft (Abends)',items:[
      '✈  Iberia landet MAD ~22:35 ab DUS',
      '🚇  Metro Linie 8 ins Zentrum (~€5, 20 Min.)',
      '🏠  Einchecken in Malasaña',
      '🍷  Bodega de la Ardosa: Vermouth vom Fass seit 1892',
      '🍺  Malasaña Kneipenbummel: Calle de la Palma, Tupperware',
      '   Bier: Mahou 5 Estrellas ~€3 · La Vía Láctea bis 3 Uhr',
    ]},
    {day:'TAG 2 — DO., 30. OKT.',date:'Stadtführung & Kochkurs',items:[
      '🚶  SANDEMANs Stadtführung 11:30 Uhr @ Plaza Mayor (2,5h)',
      '   Königspalast, Sol, La Latina, historisches Madrid',
      '🥘  Mittagessen Tapas in der Calle de la Cava Baja',
      '🍳  Kochkurs Cooking Point 17:30–21:30 Uhr (€85 p.P.)',
      '   7 Tapas-Rezepte + Sangría · Calle de Moratín, 11',
      '   Buchung: cookingpoint.es',
      '🌃  Danach: Chueca-Bar-Szene + El Tigre (kostenlose Tapas!)',
    ]},
    {day:'TAG 3 — FR., 31. OKT. 🎃',date:'HALLOWEEN-NACHT',items:[
      '🌞  Ausschlafen — Madrider Nächte gehen bis 7 Uhr',
      '🏛  Nachmittags: Prado-Museum oder Retiro-Park',
      '🎃  Halloween-Kneipentour 22 Uhr · €20 p.P.',
      '   Treffpkt.: Mulberry\'s Bar, Calle de Núñez de Arce 9',
      '   4 Bars · Tequila-Shots · €100 Kostüm-Wettbewerb',
      '🎶  Ab 1 Uhr: Teatro Kapital — 7 Etagen, offen bis 6 Uhr',
      '   Calle Atocha 125 · ~€21 Eintritt · zentralster Club',
    ]},
    {day:'TAG 4 — SA., 1. NOV. 🎉 + SO., 2. NOV.',date:'Feiertag + Abreise',items:[
      '🌞  1. Nov. = ALLERHEILIGEN (gesetzl. Feiertag in Spanien)',
      '☕  Brunch: Churros con Chocolate bei Chocolatería San Ginés',
      '🏺  Mercado de San Miguel: Pintxos + lokale Weine',
      '🍺  La Hora del Vermut (~12–15 Uhr) in Malasaña',
      '🌃  Letzte Nacht: Malasaña oder Lavapiés Spät-Bars',
      '',
      '✈  2. NOV.: Rückflug DUS (Iberia/Eurowings)',
      '   Metro Linie 8 zum Flughafen (~€5, 45 Min. einplanen)',
    ]},
  ],
  activityLinks:'Halloween-Kneipentour: eventbrite.es/e/1037812363717  ·  Kochkurs: cookingpoint.es  ·  Stadtführung: neweuropetours.eu/madrid  ·  Teatro Kapital: teatrokapital.com',
  costRows:[
    ['Hin- und Rückflug','Iberia / Eurowings  DUS ↔ MAD','~€175','~€700'],
    ['Unterkunft','Airbnb Malasaña — 3 Nächte (inkl. Gebühren)','~€130','~€520'],
    ['Halloween-Kneipentour','Bar Crawl Madrid — Puerta del Sol (31. Okt.)','€20','€80'],
    ['Kochkurs','Cooking Point — 7 Tapas + Sangría (4h)','€85','€340'],
    ['Stadtführung','SANDEMANs — Plaza Mayor (Trinkgeld)','~€12','~€48'],
    ['Essen (3 Tage)','Tapas, Churros, Markt — Ø €15–20/Mahlzeit','~€115','~€460'],
    ['Getränke & Ausgehen','Malasaña + Chueca + Halloween-Club (3 Nächte)','~€140','~€560'],
    ['Sonstiges','Metro, Uber, Souvenirs, Flughafen-Taxi','~€35','~€140'],
    ['🏆  GESAMTSCHÄTZUNG','Madrid, 29./30. Okt. – 2. Nov. 2025','~€712','~€2.848'],
  ],
  costNote:'💡 Budgetrahmen: €600–900 p.P.  El Tigre Bar: RIESIGE kostenlose Tapas-Teller zum €3-Bier — unbedingt nutzen!  31. Okt. ist Freitag + 1. Nov. Nationalfeiertag = bestes Halloween Europas 2025.',
};

const ROM = {
  num:3, name:'Rom', country:'Italien', flag:'🇮🇹', accent:ROM_C,
  heroImg:'rome_hero.png', mapImg:'rome_map.png', airbnbImg:'rome_airbnb.png',
  iataLabel:'CIA/FCO (Rom)',
  tagline:'Die Ewige Stadt — auf die partyfähige Art',
  subtitle:'Trastevere · Testaccio · Pigneto · Campo de\'Fiori · Clubs bis 5 Uhr',
  stats:[
    ['✈ Flugdauer','2h 15m\nab DTM'],['🍺 Bier','€4 – €6\n(Peroni/Moretti)'],
    ['🌡 Wetter','16–19°C\nEnde Oktober'],['💰 Kosten ca.','~€597\npro Person'],
    ['🏙 Nightlife','⭐⭐⭐⭐\nTestaccio & Pigneto'],['📍 Ab Flughafen','DTM (Ryanair)\noder DUS (Eurowings)'],
  ],
  description:'Rom ist das günstigste Ziel — dank Ryanair direkt ab DTM (Dortmund) zum Flughafen Ciampino. Trastevere ist eines der schönsten Ausgehviertel Europas: enges Kopfsteinpflaster, authentische Ristoranti, Bars bis 2 Uhr. Clubs in Testaccio und Pigneto gehen bis 5 Uhr. Kolosseum und Forum Romanum sind von außen kostenlos. Aperitivo-Zeit (ab 18 Uhr): kostenlose Snacks zum €5-Drink.',
  flightsOut:[
    ['⭐ Empfohlen','FR (Ryanair)','DTM → CIA\n30. Okt. morgens','~08:30','~10:45','2h 15m','~€55','~€220'],
    ['Vorabend 29.10.','FR (Ryanair)','DTM → CIA\n29. Okt. abends','~18:xx','~20:xx','2h 15m','~€65','~€260'],
    ['Ab DUS','EW\nEurowings','DUS → FCO\n30. Okt.','Verschieden','Verschieden','2h 20m','~€75–90','~€300–360'],
    ['Premium','ITA Airways','DUS → FCO\n30. Okt.','Verschieden','Verschieden','2h 20m','€90–130','€360–520'],
  ],
  flightRet:['FR (Ryanair)  CIA → DTM','CIA → DTM  /  2. Nov.','Verschieden','Verschieden','2h 15m','~€45–65','~€180–260'],
  flightNote:'✅  DTM (Dortmund) → CIA (Ciampino): günstigste Option via Ryanair! Ciampino liegt 12 km südlich von Rom (Bus/Metro ~40 Min., ~€7).  DUS → Fiumicino (FCO) mit Eurowings — 30 Min. per Expresszug ins Zentrum (~€14).  Buchen: ryanair.com · skyscanner.de',
  flightTotal:'€400 – €520 gesamt',
  airbnbTitle:'Trastevere / Centro Storico — 2-Zi.-Wohnung mit Terrasse',
  airbnbSpecs:[
    ['🛏 2 Schlafzimmer'],['🚿 2 Bäder'],['🏗 Terrasse/Dachterrasse ✓',[],C.green],
    ['👥 4 Gäste'],['❄ Klimaanlage ✓'],['🍕 Trastevere-Lage'],
  ],
  airbnbLocation:'📌 Trastevere — Roms schönstes Ausgehviertel, 5 Min. zum Tiber, 10 Min. zum Campo de\'Fiori',
  airbnbPricing:[['Pro Nacht','~€120'],['3 Nächte','~€360'],['Reinigung & Gebühren','~€60–80'],['Gesamt für 4 Pers.','~€420–440'],['Pro Person','~€105–110']],
  airbnbUrl:'airbnb.com → Suche: Trastevere, 4 Gäste, 2 SZ, Terrasse/Balkon, 30. Okt.–2. Nov.',
  airbnbAlt:'Alternativ: Prati (nahe Vatikan), Pigneto (angesagt) oder Testaccio (Clubs vor der Tür)',
  days:[
    {day:'TAG 1 — DO., 30. OKT.',date:'Ankunftstag',items:[
      '✈  Ryanair landet CIA ~10:45 · Bus/Metro ins Zentrum (~40 Min.)',
      '🏠  Einchecken in Trastevere',
      '🏛  Kolosseum-Außenansicht & Forum Romanum (kostenlos!)',
      '🍷  Aperol Spritz am Campo de\'Fiori zum Sonnenuntergang',
      '🍝  Abendessen Trastevere: Pasta Cacio e Pepe ~€15',
      '🍺  Erste Bar-Nacht: Trastevere-Gassen, Peroni €4–5',
    ]},
    {day:'TAG 2 — FR., 31. OKT. 🎃',date:'HALLOWEEN-NACHT',items:[
      '🍳  Pasta-Kochkurs 10:00 Uhr (3h, €65 p.P., Testaccio)',
      '   Frische Pasta + Cacio e Pepe + Tiramisu + Wein',
      '🚶  SANDEMANs Free Tour 15:30 @ Piazza Venezia (2,5h)',
      '🕙  Halloween-Kneipentour 21 Uhr · ~€20 p.P.',
      '   Treffpunkt: Nähe Trevi-Brunnen · 4 Bars · Shots',
      '   Kostüm-Wettbewerb · VIP-Clubeintritt',
      '🎉  Danach: Testaccio Clubs bis 5 Uhr (Rashomon, Goa Club)',
    ]},
    {day:'TAG 3 — SA., 1. NOV.',date:'Sightseeing & Ausgehen',items:[
      '🏛  Vatikan & Piazza Navona (früh, vor den Massen)',
      '🍕  Mittag: Original Pizza al Taglio €3–5 pro Stück',
      '🥂  Aperitivo-Zeit ab 18 Uhr am Campo de\'Fiori',
      '   Spritz + kostenlose Snacks in vielen Bars',
      '🌃  Testaccio Clubs: Rashomon, Goa Club (ab 1 Uhr)',
      '   Oder Pigneto-Bars: angesagtes Künstler-Viertel',
    ]},
    {day:'TAG 4 — SO., 2. NOV.',date:'Abreisetag',items:[
      '☕  Espresso & Cornetto (€2 an der Bar stehend — echt so!)',
      '🛍  Souvenirs: Campo de\'Fiori Markt',
      '🚌  Bus CIA (Ryanair) oder Zug FCO (Eurowings)',
      '✈  Ryanair CIA → DTM (Mittag/Abend)',
    ]},
  ],
  activityLinks:'Kneipentour: viator.com → "Rome Halloween Pub Crawl"  ·  Kochkurs: cookly.app → "Rome Pasta Class"  ·  Stadtführung: neweuropetours.eu/rome  ·  Ryanair: ryanair.com',
  costRows:[
    ['Hin- und Rückflug','Ryanair FR  DTM ↔ CIA (Ciampino)','~€120','~€480'],
    ['Unterkunft','Airbnb Trastevere — 3 Nächte (inkl. Gebühren)','~€110','~€440'],
    ['Halloween-Kneipentour','Halloween Pub Crawl Rom (31. Okt.)','€20','€80'],
    ['Kochkurs','Pasta-Kochkurs Testaccio — Pasta + Tiramisu + Wein','€65','€260'],
    ['Stadtführung','SANDEMANs — Piazza Venezia (Trinkgeld)','~€12','~€48'],
    ['Essen (3 Tage)','Pasta, Pizza al Taglio, Aperitivo — Ø €15–20','~€110','~€440'],
    ['Getränke & Ausgehen','Trastevere + Testaccio + Pigneto (3 Nächte)','~€130','~€520'],
    ['Sonstiges','Bus CIA, Metro, Souvenirs','~€30','~€120'],
    ['🏆  GESAMTSCHÄTZUNG','Rom, 30. Okt. – 2. Nov. 2025','~€597','~€2.388'],
  ],
  costNote:'💡 Günstigstes Ziel dank Ryanair ab DTM!  Aperitivo-Zeit (18–20 Uhr): kostenlose Snacks zum €5-Drink in vielen Bars.  Kolosseum + Forum Romanum Außenansicht = kostenlos.  Budgetrahmen: €500–700 p.P.',
};

const MALTA = {
  num:4, name:'Malta', country:'Valletta', flag:'🇲🇹', accent:MALT_C,
  heroImg:'malta_hero.png', mapImg:'malta_map.png', airbnbImg:'malta_airbnb.png',
  iataLabel:'MLA (Malta)',
  tagline:'Wärmstes Ziel — Mittelmeer im Oktober',
  subtitle:'Paceville · Sliema · Valletta · Blaue Lagune · Cisk Lager',
  stats:[
    ['✈ Flugdauer','~3h 00m\nab DUS'],['🍺 Bier','€2 – €3\n(Cisk Lager)'],
    ['🌡 Wetter','22–24°C ☀☀\nWärmstes Ziel!'],['💰 Kosten ca.','~€580\npro Person'],
    ['🏙 Nightlife','⭐⭐⭐⭐\nPaceville Clubs'],['📍 Ab Flughafen','DUS\n(Eurowings/Ryanair)'],
  ],
  description:'Malta ist mit 22–24°C das WÄRMSTE europäische Ziel dieser Auswahl — und Ende Oktober noch Badesaison! Die Insel bietet UNESCO-Welterbe Valletta, das mittelalterliche Mdina und das Nachtleben von Paceville (St. Julian\'s). Halloween in den Clubs von Paceville ist eine eigene Kategorie. Und das Cisk Lager gehört zu den günstigsten Bieren Europas.',
  flightsOut:[
    ['⭐ Empfohlen','EW\nEurowings','DUS → MLA\n30. Okt. früh','~07:xx','~10:00','2h 55m','~€90','~€360'],
    ['Vorabend 29.10.','Eurowings\n/ Ryanair','DUS → MLA\n29. Okt. abends','Abends','Abends','~3h','~€100','~€400'],
    ['Budget','FR\nRyanair','DUS → MLA\n30. Okt.','Verschieden','Verschieden','3h 00m','~€70–90','~€280–360'],
  ],
  flightRet:['Eurowings/Ryanair  (MLA→DUS)','MLA → DUS  /  2. Nov.','Verschieden','Verschieden','~3h','~€80–100','~€320–400'],
  flightNote:'🌞  Malta ist mit 22–24°C Anfang November das WÄRMSTE Reiseziel dieser Auswahl!  Flughafen Malta (MLA) liegt 8 km südlich von Valletta — Taxi/Bus ins Zentrum ~20 Min.  Direktflüge ab DUS mit Eurowings & Ryanair.  Buchen: eurowings.com · ryanair.com · skyscanner.de',
  flightTotal:'€600 – €800 gesamt',
  airbnbTitle:'St. Julian\'s / Sliema — 2-Zi.-Wohnung mit Terrasse & Meerblick',
  airbnbSpecs:[
    ['🛏 2 Schlafzimmer'],['🚿 2 Bäder'],['🏗 Terrasse ✓',[],C.green],
    ['👥 4 Gäste'],['❄ Klimaanlage ✓'],['🌊 Meerblick mögl.',[],MALT_C],
  ],
  airbnbLocation:'📌 St. Julian\'s / Paceville — direkt am Nachtleben, 2 Min. zu Hugo\'s Lounge & Beach Clubs',
  airbnbPricing:[['Pro Nacht','~€95'],['3 Nächte','~€285'],['Reinigung & Gebühren','~€50–70'],['Gesamt für 4 Pers.','~€335–355'],['Pro Person','~€84–89']],
  airbnbUrl:'airbnb.com → Suche: St. Julian\'s Malta, 4 Gäste, 2 SZ, Terrasse, 30. Okt.–2. Nov.',
  airbnbAlt:'Alternative: Sliema (ruhiger, Wasserfront) oder Valletta (historisch, UNESCO) ebenfalls empfohlen',
  days:[
    {day:'TAG 1 — DO., 30. OKT.',date:'Ankunftstag',items:[
      '✈  Eurowings landet MLA ~10:00 · Taxi nach St. Julian\'s',
      '🏠  Einchecken · erstes kühles Cisk-Bier auf der Terrasse',
      '🏛  Valletta Altstadt-Spaziergang (UNESCO-Welterbe)',
      '🌊  Sliema Waterfront zum Sonnenuntergang',
      '🍺  Erste Bar-Nacht in Paceville: Hugo\'s Lounge',
      '   Cisk 0,5l: ~€2,50 · T-Shirt-Wetter um Mitternacht!',
    ]},
    {day:'TAG 2 — FR., 31. OKT. 🎃',date:'HALLOWEEN-NACHT',items:[
      '⛵  Tagesausflug Blaue Lagune (Bootsfahrt, ~€35 p.P.)',
      '   Kristallklares Mittelmeer · letztes Bad im Oktober!',
      '🍳  Maltesischer Kochkurs nachmittags (€65 p.P.)',
      '   Kaninchen-Eintopf (Fenek) + Pastizzi + lokaler Wein',
      '🎃  Halloween in Paceville ab 22 Uhr',
      '   Clubmeile: Sky Club · Footloose · Level 22',
      '   Bis 4 Uhr morgens · alle mit Halloween-Deko',
    ]},
    {day:'TAG 3 — SA., 1. NOV.',date:'Malta Erkunden',items:[
      '🏰  Mdina: "Die Stille Stadt" — mittelalterl. Festungsstadt',
      '🐟  Marsaxlokk Fischmarkt: frischer Fang, bunte Boote',
      '🍺  Mittagspause: Pasta + Cisk am Hafen (~€12)',
      '🌊  Nachmittag: Baden an den Felsenbädern (noch warm!)',
      '🌃  Letzte Nacht: Paceville + Strand-Bars',
    ]},
    {day:'TAG 4 — SO., 2. NOV.',date:'Abreisetag',items:[
      '☀  Frühstück auf der Terrasse bei 20°C',
      '🥐  Pastizzi zum Frühstück (Ricotta-Blätterteig, €0,50!)',
      '🏖  Letzte Stunde am Meer (badefertig!)',
      '🛍  Souvenirs: Maltesisches Kreuz, Ftira-Brot, Cisk-Sixpack',
      '✈  Eurowings/Ryanair Rückflug nach DUS',
    ]},
  ],
  activityLinks:'Bootsausflug Blaue Lagune: katarinaLine.com oder lokal buchen  ·  Kochkurs: experiencemalta.com  ·  Stadtführung: visitvalletta.mt  ·  Paceville Clubs: hugo.com.mt',
  costRows:[
    ['Hin- und Rückflug','Eurowings/Ryanair  DUS ↔ MLA','~€185','~€740'],
    ['Unterkunft','Airbnb St. Julian\'s — 3 Nächte (inkl. Gebühren)','~€89','~€355'],
    ['Halloween-Party','Paceville Clubeintritt (31. Okt.) — meist günstiger','~€15','~€60'],
    ['Kochkurs','Maltesische Küche — Fenek + Pastizzi + Wein','€65','€260'],
    ['Bootsausflug','Blaue Lagune (31. Okt. tagsüber)','€35','€140'],
    ['Stadtführung','Valletta Free Tour (Trinkgeld)','~€12','~€48'],
    ['Essen (3 Tage)','Fisch, Pasta, Pastizzi, Hafenrestaurants','~€95','~€380'],
    ['Getränke & Ausgehen','Paceville Clubs + Strandbar (3 Nächte)','~€110','~€440'],
    ['Sonstiges','Taxi Flughafen, Bus, Souvenirs','~€25','~€100'],
    ['🏆  GESAMTSCHÄTZUNG','Malta, 30. Okt. – 2. Nov. 2025','~€631','~€2.523'],
  ],
  costNote:'💡 Günstigste Unterkunft dieser Auswahl!  Cisk-Bier ab €2,50.  Pastizzi (maltesisches Gebäck) für €0,50.  Baden im Mittelmeer noch möglich!  Blaue Lagune ist eines der schönsten Gewässer Europas.',
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
  s.addText('VIER NEUE ZIELE — JETZT ABSTIMMEN!',{x:0.5,y:4.35,w:12.33,h:0.65,align:'center',fontSize:26,fontFace:'Trebuchet MS',bold:true,color:C.coral});
  // 4 destination pills
  const pills=[['🇵🇹 LISSABON',LIS_C],['🇪🇸 MADRID',MAD_C],['🇮🇹 ROM',ROM_C],['🇲🇹 MALTA',MALT_C]];
  pills.forEach(([label,col],i)=>{
    const px=0.5+i*3.2;
    s.addShape('rect',{x:px,y:5.15,w:3.0,h:0.65,fill:{color:C.bgDeep},line:{color:col,pt:2},rounding:0.06});
    s.addText(label,{x:px,y:5.15,w:3.0,h:0.65,align:'center',fontSize:18,fontFace:'Trebuchet MS',bold:true,color:C.white,valign:'middle'});
  });
  s.addText('Flüge ab DUS · FMO · DTM  ·  Alle Ziele: Ende Oktober 16–24°C',{x:0,y:6.25,w:'100%',h:0.4,align:'center',fontSize:13,fontFace:'Calibri',color:C.dimGray,italic:true});
  s.addShape('rect',{x:0,y:7.42,w:'100%',h:0.08,fill:{color:C.orange},line:{color:C.orange}});
}

// ── Slide 22: Direktvergleich ─────────────────────────────────────────────────
function buildComparison(pres) {
  const s = pres.addSlide();
  s.addImage({path:path.join(IMG,'comparison4_bg.png'),x:0,y:0,w:'100%',h:'100%'});
  s.addShape('rect',{x:0,y:0,w:'100%',h:'100%',fill:{color:'000000',transparency:50},line:{color:'000000'}});
  s.addShape('rect',{x:0,y:0,w:'100%',h:0.72,fill:{color:C.bgCard,transparency:15},line:{color:C.bgCard}});
  s.addShape('rect',{x:0,y:0,w:'100%',h:0.08,fill:{color:C.orange},line:{color:C.orange}});
  s.addText('DIREKTVERGLEICH — WO FAHREN WIR HIN?',{x:0.3,y:0.08,w:12.7,h:0.64,fontSize:22,fontFace:'Trebuchet MS',bold:true,color:C.white,valign:'middle'});

  const cols=[
    {label:'🇵🇹 LISSABON', color:LIS_C, x:2.8},
    {label:'🇪🇸 MADRID',   color:MAD_C, x:5.4},
    {label:'🇮🇹 ROM',      color:ROM_C, x:8.0},
    {label:'🇲🇹 MALTA',   color:MALT_C,x:10.6},
  ];
  // Column headers
  cols.forEach(col=>{
    s.addShape('rect',{x:col.x,y:0.82,w:2.45,h:0.52,fill:{color:C.bgDeep,transparency:10},line:{color:col.color,pt:1.5},rounding:0.04});
    s.addText(col.label,{x:col.x,y:0.82,w:2.45,h:0.52,align:'center',fontSize:15,fontFace:'Trebuchet MS',bold:true,color:col.color,valign:'middle'});
  });

  const rows=[
    ['💰 Kosten p.P.',       '~€735',                  '~€712',                 '~€597',                '~€631'],
    ['✈ Flugdauer',          '3h 10m · DUS',           '2h 30m · DUS',          '2h 15m · DTM',         '~3h · DUS'],
    ['🛫 Flughafen',         'DUS (Eurowings/TAP)',     'DUS (Iberia)',           'DTM (Ryanair)',         'DUS (Eurowings)'],
    ['🍺 Bierpreis (Ø)',     '€2–3 / Pint',            '€2–4 / Caña',           '€4–6 / Flasche',       '€2–3 (Cisk)'],
    ['🌡 Wetter Okt.',       '17–22°C ☀',              '18–20°C ☀',             '16–19°C',              '22–24°C ☀☀'],
    ['🎃 Halloween',         '⭐⭐⭐⭐ Fr.+Clubs',      '⭐⭐⭐⭐⭐ Fr.+Feiert.',  '⭐⭐⭐⭐ Testaccio',     '⭐⭐⭐⭐ Paceville'],
    ['🏙 Besonderheit',      'Pink St.·Fado·Atlantik', 'Malasaña·Tapas·Vermouth','Kolosseum·Pasta·Aperol','Blaue Lagune·Baden'],
    ['🎭 Clubs bis',         '6 Uhr',                  '7 Uhr',                 '5 Uhr',                '4 Uhr'],
    ['🍴 Essen',             'Petiscos €15–25',        'Tapas €15–20',          'Pizza/Pasta €15–20',   'Fisch €12–20'],
    ['🍳 Kochkurs',          '€110 p.P.',              '€85 p.P.',              '€65 p.P.',             '€65 p.P.'],
    ['💡 Unser Tipp',        'Atlantik-Charme+Fado',   'Fr.+Feiertag = Top!',   'Günstig ab DTM!',      'Wärmstes+Meer!'],
  ];

  rows.forEach((row,i)=>{
    const ry=1.44+i*0.54;
    const isLast=i===rows.length-1;
    const rowFill=isLast?C.bgDeep:(i%2===0?C.bgCard:C.bg);
    const tr=isLast?0:25;

    // Category
    s.addShape('rect',{x:0.25,y:ry,w:2.45,h:0.5,fill:{color:rowFill,transparency:tr},line:{color:C.bgDeep}});
    s.addText(row[0],{x:0.3,y:ry,w:2.4,h:0.5,fontSize:isLast?11:9.5,fontFace:'Calibri',bold:isLast,color:isLast?C.amber:C.gray,valign:'middle'});

    cols.forEach((col,ci)=>{
      s.addShape('rect',{x:col.x,y:ry,w:2.45,h:0.5,fill:{color:rowFill,transparency:tr},line:{color:isLast?col.color:C.bgDeep,pt:isLast?1.5:0.5}});
      s.addText(row[ci+1],{x:col.x+0.05,y:ry,w:2.35,h:0.5,align:'center',fontSize:isLast?10.5:9,fontFace:isLast?'Trebuchet MS':'Calibri',bold:isLast,color:isLast?col.color:C.offWhite,valign:'middle'});
    });
  });
  s.addShape('rect',{x:0,y:7.42,w:'100%',h:0.08,fill:{color:C.orange},line:{color:C.orange}});
}

// ── Slide 23: Männerreise-Statistiken ────────────────────────────────────────
function buildMaennerStats(pres) {
  const s = pres.addSlide();
  s.background = { color: C.bg };
  s.addShape('rect',{x:0,y:0,w:'100%',h:0.08,fill:{color:C.orange},line:{color:C.orange}});
  s.addShape('rect',{x:0,y:0.08,w:'100%',h:0.72,fill:{color:C.bgCard},line:{color:C.bgCard}});
  s.addText('📊  WISSENSCHAFT DER MÄNNERREISE — BELEGBARE FAKTEN',{x:0.4,y:0.08,w:12.5,h:0.72,fontSize:22,fontFace:'Trebuchet MS',bold:true,color:C.white,valign:'middle'});
  s.addShape('rect',{x:0,y:0.8,w:'100%',h:0.03,fill:{color:C.orange},line:{color:C.orange}});

  // Big stat boxes: 3 per row, 3 rows
  const stats = [
    { icon:'🍺', val:'+340%',   label:'Bierkonsum vs. Zuhause',       sub:'gemessen in Litern, nicht Reue',   color:C.orange },
    { icon:'⌛', val:'4h 12m',  label:'Schlafdauer pro Nacht',         sub:'geplant waren: 8h',                color:C.amber  },
    { icon:'📱', val:'-83%',    label:'Nachrichten nach Hause',        sub:'pro Tag nach Ankunft',             color:C.coral  },
    { icon:'🔁', val:'7×',      label:'"Nur noch eine Bar" pro Nacht', sub:'davon eingehalten: 0×',            color:LIS_C   },
    { icon:'🔑', val:'1,4',     label:'Verlorene Schlüssel pro Trip',  sub:'einer davon taucht nie wieder auf',color:MALT_C  },
    { icon:'💸', val:'+47%',    label:'Budget-Überschreitung',         sub:'Ø über alle Männertouren weltweit',color:MAD_C   },
    { icon:'📸', val:'847 / 3', label:'Fotos gemacht vs. geteilt',     sub:'die anderen 844 bleiben unter uns',color:ROM_C   },
    { icon:'🕙', val:'22:47',   label:'Tatsächliche Abfahrtszeit',     sub:'geplant: 21:00 Uhr',              color:C.orange },
    { icon:'🌅', val:'04:35',   label:'Heimkehrzeit (versprochen: 01)', sub:'Frühstück direkt inbegriffen',    color:C.amber  },
  ];

  const colW = 4.1, rowH = 1.78;
  const startX = 0.25, startY = 0.97;

  stats.forEach((st, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = startX + col * (colW + 0.1);
    const y = startY + row * (rowH + 0.08);

    s.addShape('rect',{x,y,w:colW,h:rowH,fill:{color:C.bgCard},line:{color:st.color,pt:1.5},rounding:0.05});
    s.addShape('rect',{x,y,w:colW,h:0.44,fill:{color:C.bgDeep},line:{color:C.bgDeep},rounding:0.05});

    // Emoji icon
    s.addText(st.icon, {x, y:y+0.02, w:colW, h:0.44, align:'center', fontSize:20, fontFace:'Segoe UI Emoji', valign:'middle'});
    // Big number
    s.addText(st.val, {x, y:y+0.46, w:colW, h:0.68, align:'center', fontSize:32, fontFace:'Trebuchet MS', bold:true, color:st.color, valign:'middle'});
    // Label
    s.addText(st.label, {x:x+0.08, y:y+1.16, w:colW-0.16, h:0.3, align:'center', fontSize:10, fontFace:'Trebuchet MS', bold:true, color:C.offWhite, valign:'middle'});
    // Subtext
    s.addText(st.sub, {x:x+0.08, y:y+1.46, w:colW-0.16, h:0.26, align:'center', fontSize:8.5, fontFace:'Calibri', color:C.dimGray, italic:true, valign:'top', wrap:true});
  });

  s.addShape('rect',{x:0,y:7.3,w:'100%',h:0.12,fill:{color:C.bgCard},line:{color:C.bgDeep}});
  s.addText('Quelle: Erfahrungswerte, n = unzählige Männerreisen seit 1987  ·  Statistiken ohne Gewähr, aber mit vollem Überzeugungsgrad  ·  🏆 In Memoriam: Budgets, die nie überlebt haben',
    {x:0.3,y:7.3,w:12.7,h:0.12,fontSize:7.5,fontFace:'Calibri',color:C.dimGray,italic:true,valign:'middle'});
  bottomBar(s);
}

// ── MAIN ─────────────────────────────────────────────────────────────────────
async function main() {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Halloween-Trip 2025';
  pres.title = 'Halloween-Trip 2025 – Lissabon vs Madrid vs Rom vs Malta';

  console.log('Erstelle Folien...');
  slide1(pres); console.log('  ✓ Folie 1: Titelfolie');

  const cities=[LISSABON,MADRID,ROM,MALTA];
  cities.forEach((cfg,ci)=>{
    const base=(ci*5)+2;
    buildOverview(pres,cfg);   console.log(`  ✓ Folie ${base}: ${cfg.name} Übersicht`);
    buildFlights(pres,cfg);    console.log(`  ✓ Folie ${base+1}: ${cfg.name} Flüge`);
    buildAirbnb(pres,cfg);     console.log(`  ✓ Folie ${base+2}: ${cfg.name} Unterkunft`);
    buildActivities(pres,cfg); console.log(`  ✓ Folie ${base+3}: ${cfg.name} Programm`);
    buildCosts(pres,cfg);      console.log(`  ✓ Folie ${base+4}: ${cfg.name} Kosten`);
  });

  buildComparison(pres);    console.log('  ✓ Folie 22: Direktvergleich');
  buildMaennerStats(pres);  console.log('  ✓ Folie 23: Männerreise-Statistiken');

  await pres.writeFile({ fileName: OUT });
  console.log(`\n✅ Gespeichert: ${OUT}`);
}

main().catch(err => { console.error(err); process.exit(1); });
