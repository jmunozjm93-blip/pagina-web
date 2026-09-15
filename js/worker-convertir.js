/* Web Worker: lee el Excel y lo convierte sin congelar la pantalla */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'convertir.js');

self.onmessage = function (e) {
  const { id, tipo, semana, buffer } = e.data;
  const t0 = Date.now();
  try {
    self.postMessage({ id, etapa: 'leyendo' });
    const wb = XLSX.read(buffer, Convertir.opcionesLectura(tipo));
    const tLectura = Date.now() - t0;
    self.postMessage({ id, etapa: 'convirtiendo', hojas: wb.SheetNames });
    const json = Convertir.convertir(tipo, wb, semana);
    const semanaArchivo = Convertir.semanaDelLibro(tipo, wb);
    const resumen = Convertir.resumen(tipo, json, semana);
    delete json.filasLeidas;
    const texto = JSON.stringify(json);
    self.postMessage({ id, etapa: 'listo', texto, resumen, semanaArchivo, ms: { lectura: tLectura, total: Date.now() - t0 } });
  } catch (err) {
    self.postMessage({ id, etapa: 'error', error: err.message || String(err) });
  }
};
