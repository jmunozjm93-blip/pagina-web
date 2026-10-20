/* Usa la misma conexión y cola de cargas de cargar.html. */
function procesarReservas(file) {
  const semana = el('resSemana').value;
  const card = document.createElement('div');
  card.className = 'card';
  const title = document.createElement('div');
  title.className = 'name';
  title.textContent = `${file.name} · Reservas${semana ? ' · S' + semana : ''}`;
  const status = document.createElement('div');
  status.className = 'st';
  const body = document.createElement('div'), actions = document.createElement('div');
  actions.className = 'acts';
  card.append(title, status, body, actions);
  el('resCards').prepend(card);
  const message = (text, cls = '') => { status.textContent = text; status.className = 'st ' + cls; };
  if (!semana) { message('Primero elige una semana con datos de sell out.', 'bad'); return; }
  if (!/\.xlsx$/i.test(file.name)) { message('Selecciona un archivo .xlsx.', 'bad'); return; }
  message('Leyendo y validando las tres hojas…');
  const worker = new Worker('js/worker-reservas.js?v=1');
  const fail = error => { worker.terminate(); message(error, 'bad'); };
  worker.onerror = () => fail('No se pudo leer el Excel. Revisa la conexión e inténtalo nuevamente.');
  worker.onmessage = event => {
    worker.terminate();
    if (event.data.error) { message(event.data.error, 'bad'); return; }
    const data = event.data.data;
    const table = document.createElement('table');
    table.innerHTML = '<thead><tr><th>Reserva</th><th class="n">Modelos</th><th class="n">Unidades a cargar</th></tr></thead>';
    const tbody = document.createElement('tbody');
    for (const item of data.resumen) {
      const tr = document.createElement('tr');
      for (const [i, value] of [item.hoja, fmt(item.modelos), fmt(item.unidades)].entries()) {
        const td = document.createElement('td'); td.textContent = value;
        if (i) td.className = 'n';
        tr.append(td);
      }
      tbody.append(tr);
    }
    table.append(tbody); body.append(table);
    const omitted = data.omitidas;
    const warning = omitted.length
      ? ` Se ${omitted.length === 1 ? 'omite 1 fila' : 'omiten ' + fmt(omitted.length) + ' filas'} sin Modelo (${fmt(omitted.reduce((s,r)=>s+r.unidades,0))} unidades). Corrige el origen para incluirlas.`
      : '';
    message(`Validado para S${semana}. Se reemplazarán solo BO02, KAM y REGIONALES.${warning}`, omitted.length ? 'warn' : 'ok');
    const upload = document.createElement('button');
    upload.textContent = `Actualizar Reservas S${semana}`;
    actions.append(upload);
    upload.onclick = () => {
      if (!token()) { message('Primero guarda el token de GitHub en el paso 1.', 'bad'); return; }
      upload.disabled = true;
      message('En cola para actualizar Reservas…');
      enCola(async () => {
        try {
          const manifest = await leerManifiesto();
          if (!['calzado','ropa','acc'].some(t => manifest.semanas?.[semana]?.[t])) throw new Error('La semana elegida ya no tiene datos de sell out. Actualiza la página.');
          message(`Actualizando Reservas de S${semana}…`);
          data.actualizado = new Date().toISOString();
          const commit = await subir(`data/w${semana}/reservas.json`, JSON.stringify(data),
            `Semana ${semana}: actualiza BO02, KAM y REGIONALES por Modelo`);
          await registrar(semana, 'reservas');
          message(`Reservas de S${semana} actualizadas. GitHub Pages lo publica en 1–2 minutos.${warning}`, omitted.length ? 'warn' : 'ok');
          actions.replaceChildren();
          const link = document.createElement('a');
          link.textContent = 'Ver actualización'; link.href = commit.html_url;
          link.target = '_blank'; link.rel = 'noopener';
          actions.append(link);
          cargarEstado();
        } catch (error) {
          message(error.message, 'bad');
          upload.disabled = false;
        }
      });
    };
  };
  file.arrayBuffer().then(buffer => worker.postMessage({buffer, semana}, [buffer]))
    .catch(() => fail('No se pudo abrir el archivo seleccionado.'));
}
el('resFile').addEventListener('change', event => {
  if (event.target.files[0]) procesarReservas(event.target.files[0]);
  event.target.value = '';
});
