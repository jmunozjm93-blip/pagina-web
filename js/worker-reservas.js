importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'reservas.js?v=1');
self.onmessage = function (event) {
  const {buffer, semana} = event.data;
  try {
    const wb = XLSX.read(buffer, {type:'array'});
    const data = Reservas.convertir(wb, semana, XLSX);
    self.postMessage({data});
  } catch (error) {
    self.postMessage({error: error.message || String(error)});
  }
};
