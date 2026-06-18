/**
 * update-projects.js
 * Escanea proyectos/[ubicacion]/[proyecto]/ y genera projects-data.js
 * Uso: node update-projects.js
 * Para actualización automática: node update-projects.js --watch
 */

const fs   = require('fs');
const path = require('path');

const ROOT        = __dirname;
const PROYECTOS   = path.join(ROOT, 'proyectos');
const OUTPUT      = path.join(ROOT, 'projects-data.js');
const IMAGE_EXTS  = new Set(['.jpg', '.jpeg', '.png', '.webp', '.avif', '.gif']);

function isImage(f) { return IMAGE_EXTS.has(path.extname(f).toLowerCase()); }

function toTitle(name) {
  // Capitaliza primera letra de cada palabra; \b falla con letras acentuadas
  return name.replace(/[-_]/g, ' ').trim()
    .split(' ')
    .map(w => w.length ? w[0].toUpperCase() + w.slice(1) : w)
    .join(' ');
}

function extractYear(str) {
  const m = str.match(/\b(20\d{2})\b/);
  return m ? m[1] : null;
}

function stripYear(str) {
  return str.replace(/\s*\b20\d{2}\b\s*/g, '').trim();
}

function scan() {
  const projects = [];

  if (!fs.existsSync(PROYECTOS)) {
    fs.mkdirSync(PROYECTOS, { recursive: true });
    console.log('✓ Carpeta "proyectos/" creada. Agregá subcarpetas con tus proyectos.');
    return projects;
  }

  const locations = fs.readdirSync(PROYECTOS)
    .filter(d => fs.statSync(path.join(PROYECTOS, d)).isDirectory())
    .sort();

  for (const loc of locations) {
    const locPath  = path.join(PROYECTOS, loc);
    const projDirs = fs.readdirSync(locPath)
      .filter(d => fs.statSync(path.join(locPath, d)).isDirectory())
      .sort();

    for (const proj of projDirs) {
      const projPath = path.join(locPath, proj);
      const files    = fs.readdirSync(projPath)
        .filter(isImage)
        .sort();

      if (files.length === 0) continue;

      const rawName    = toTitle(proj);
      const rawLoc     = toTitle(loc);
      const year       = extractYear(rawName) || extractYear(rawLoc) || String(new Date().getFullYear());
      const name       = stripYear(rawName);
      const location   = stripYear(rawLoc);

      // Rutas relativas al index.html (mismo directorio raíz)
      const images = files.map(f => `proyectos/${loc}/${proj}/${f}`);

      projects.push({ name, location, year, cover: images[0], images });
    }
  }

  return projects;
}

function write(projects) {
  const ts  = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });
  const out = `// Auto-generado por update-projects.js — ${ts}
// NO editar manualmente. Para actualizar: node update-projects.js
window.PROJECTS = ${JSON.stringify(projects, null, 2)};
`;
  fs.writeFileSync(OUTPUT, out, 'utf8');
  console.log(`[${ts}] ✓ ${projects.length} proyecto(s) → projects-data.js`);
  projects.forEach(p => console.log(`   · ${p.name} | ${p.location} | ${p.year} | ${p.images.length} imagen(es)`));
}

// --- main ---
const projects = scan();
write(projects);

// Modo watch: node update-projects.js --watch
if (process.argv.includes('--watch')) {
  console.log('\nModo watch activo. Escuchando cambios en proyectos/...\n');
  let debounce;
  fs.watch(PROYECTOS, { recursive: true }, () => {
    clearTimeout(debounce);
    debounce = setTimeout(() => { write(scan()); }, 600);
  });
}
