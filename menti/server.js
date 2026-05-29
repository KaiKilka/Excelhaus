const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const path = require('path');

const app = express();
const server = http.createServer(app);
const io = new Server(server);

app.use(express.static(path.join(__dirname, 'public')));

const CRITERIA = [
  { id: 'wetter',          label: '🌡 Wetter',                   desc: 'Wie wichtig ist dir das Wetter Ende Oktober?' },
  { id: 'flug',            label: '✈ Flug & Erreichbarkeit',     desc: 'Flugdauer und Direktflug ab DUS?' },
  { id: 'halloween',       label: '🎃 Halloween-Atmosphäre',      desc: 'Wie stark ist die Halloween-Stimmung?' },
  { id: 'nightlife',       label: '🍺 Nightlife & Bars',          desc: 'Clubbing, Bars, Ausgehen bis in die Nacht?' },
  { id: 'preis',           label: '💰 Preis-Leistung',            desc: 'Budget: Flug + Unterkunft + Essen + Bier?' },
  { id: 'sehensw',         label: '🏙 Sehenswürdigkeiten',        desc: 'Kultur, Architektur und Stadtflair?' },
  { id: 'essen',           label: '🍳 Essen & Trinken',           desc: 'Lokale Küche, Restaurants, Streetfood?' },
  { id: 'unterkunft',      label: '🏠 Unterkunft',                desc: 'Airbnb-Verfügbarkeit, Lage und Preis?' },
];

const CITIES = ['lissabon', 'sofia'];

// votes[socketId] = { gender, ratings: { criterionId: { lissabon: 1-5, sofia: 1-5 } } }
const votes = {};

function computeResults() {
  const participants = Object.values(votes);
  const total = participants.length;
  const byGender = { maennlich: [], weiblich: [] };
  participants.forEach(p => {
    if (p.gender === 'maennlich') byGender.maennlich.push(p);
    else if (p.gender === 'weiblich') byGender.weiblich.push(p);
  });

  function avgRatings(group) {
    const out = {};
    CRITERIA.forEach(c => {
      out[c.id] = {};
      CITIES.forEach(city => {
        const vals = group.map(p => p.ratings?.[c.id]?.[city]).filter(v => v > 0);
        out[c.id][city] = vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length) : 0;
      });
    });
    return out;
  }

  function overallAvg(group) {
    const avgByC = avgRatings(group);
    const totals = {};
    CITIES.forEach(city => {
      const vals = CRITERIA.map(c => avgByC[c.id][city]).filter(v => v > 0);
      totals[city] = vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length) : 0;
    });
    return totals;
  }

  return {
    total,
    genderCount: { maennlich: byGender.maennlich.length, weiblich: byGender.weiblich.length },
    all: avgRatings(participants),
    byGender: {
      maennlich: avgRatings(byGender.maennlich),
      weiblich: avgRatings(byGender.weiblich),
    },
    overall: overallAvg(participants),
    overallByGender: {
      maennlich: overallAvg(byGender.maennlich),
      weiblich: overallAvg(byGender.weiblich),
    },
    criteria: CRITERIA,
  };
}

io.on('connection', (socket) => {
  votes[socket.id] = { gender: null, ratings: {} };

  socket.emit('init', { criteria: CRITERIA, cities: CITIES });
  socket.emit('results', computeResults());

  socket.on('setGender', (gender) => {
    if (!['maennlich', 'weiblich'].includes(gender)) return;
    votes[socket.id].gender = gender;
    io.emit('results', computeResults());
  });

  socket.on('setRating', ({ criterionId, city, value }) => {
    if (!CRITERIA.find(c => c.id === criterionId)) return;
    if (!CITIES.includes(city)) return;
    const v = parseInt(value, 10);
    if (v < 1 || v > 5) return;
    if (!votes[socket.id].ratings[criterionId]) votes[socket.id].ratings[criterionId] = {};
    votes[socket.id].ratings[criterionId][city] = v;
    io.emit('results', computeResults());
  });

  socket.on('disconnect', () => {
    delete votes[socket.id];
    io.emit('results', computeResults());
  });
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Menti app running on http://localhost:${PORT}`);
  console.log(`Results dashboard: http://localhost:${PORT}/results.html`);
});
