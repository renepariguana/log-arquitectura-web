/**
 * update-cost.js
 * Scrapea el costo m² de construcción desde arquitecturayconstrucciondigital.com
 * Uso: node update-cost.js
 * Cron diario (ej. 8am): 0 8 * * * cd /ruta && node update-cost.js
 */

const https = require('https');
const fs    = require('fs');
const path  = require('path');

const URL    = 'https://arquitecturayconstrucciondigital.com/';
const OUTPUT = path.join(__dirname, 'cost-data.js');

function fetch(url, cb) {
  https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, (res) => {
    if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
      return fetch(res.headers.location, cb);
    }
    let data = '';
    res.on('data', c => data += c);
    res.on('end', () => cb(null, data));
  }).on('error', cb);
}

fetch(URL, (err, html) => {
  if (err) { console.error('Error al obtener la página:', err.message); return; }

  // Busca patrón: "Costo m² | $2.009.439"
  const m = html.match(/Costo\s+m[²2]\s*[\|l]\s*\$([\d\.,]+)/i);
  if (!m) { console.log('⚠ No se encontró el costo en la página.'); return; }

  const cost = '$' + m[1];
  const ts   = new Date().toLocaleDateString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });
  const out  = `// Auto-actualizado por update-cost.js — ${ts}\nwindow.CONSTRUCTION_COST = { value: "${cost}", date: "${ts}" };\n`;

  fs.writeFileSync(OUTPUT, out, 'utf8');
  console.log(`✓ Costo m² actualizado: ${cost} (${ts})`);
});
