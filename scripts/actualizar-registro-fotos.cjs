// Ejecutar después de cambiar los catálogos de fotos embebidos en los dashboards.
const fs = require('node:fs');
const path = require('node:path');
const root = path.join(__dirname, '..');
const img = Object.create(null);
for (const file of ['Dashboard sell out.html', 'Dashboard Steve Madden.html']) {
  const html = fs.readFileSync(path.join(root, file), 'utf8');
  const match = html.match(/<script[^>]*id="IMG"[^>]*>([\s\S]*?)<\/script>/)
    || html.match(/const IMG = (?:FotosManuales\.combinar\()?(\{[^\r\n]+\})(?:,await FotosManuales\.leer\(\)\))?;/);
  if (!match) throw new Error('Catálogo no encontrado: ' + file);
  for (const [key, value] of Object.entries(JSON.parse(match[1]))) if (value) img[key] = true;
}
fs.writeFileSync(path.join(root, 'data/fotos-registradas.json'), JSON.stringify({img}));
console.log(Object.keys(img).length + ' modelos registrados');
