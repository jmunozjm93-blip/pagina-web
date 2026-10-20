(function() {
  'use strict';
  const el=id=>document.getElementById(id),fmt=n=>n.toLocaleString('es-CL');
  let defaults=null,stockFile=null,detailFile=null,worker=null,revision=0,busy=false;
  const status=(text,cls='')=>{el('status').textContent=text;el('status').className='status '+cls;};
  function showRules(rules,custom=false) {
    el('rulesStatus').textContent=custom?'Usando el detalle adjunto: '+detailFile.name:'Reglas guardadas de detalle reservas.xlsx.';
    el('rulesStatus').className='status ok';el('rules').replaceChildren();
    for(const [name,codes] of Object.entries(rules.grupos)) {
      const row=document.createElement('div');row.className='rule';
      const label=document.createElement('b');label.textContent=name;
      const text=document.createElement('span');text.className='codes';text.textContent=codes.join(', ');
      row.append(label,text);el('rules').append(row);
    }
    const all=Object.values(rules.grupos),shared=[...new Set(all.flat())].filter(c=>all.filter(list=>list.includes(c)).length>1);
    el('shared').textContent=shared.length?`${shared.join(', ')} se incluye en más de una hoja, según el detalle. No sumes las hojas como un total único.`:'Cada almacén está asignado a una sola hoja.';
  }
  function invalidate() {
    revision++;if(worker)worker.terminate();worker=null;busy=false;
    el('review').hidden=true;el('downloadStatus').textContent='';
    el('download').disabled=false;
    el('separate').replaceChildren();
  }
  async function process() {
    invalidate();const current=revision;
    if(!stockFile){status('Adjunta el stock para generar la revisión automáticamente.');return;}
    if(!defaults&&!detailFile){status('Espera las reglas guardadas o adjunta un detalle de reservas.','warn');return;}
    status('Leyendo el stock y aplicando los filtros…');
    try {
      const [stock,detalle]=await Promise.all([stockFile.arrayBuffer(),detailFile?detailFile.arrayBuffer():null]);
      if(current!==revision)return;
      worker=new Worker('js/worker-armar-reservas.js?v=1');
      worker.onerror=()=>{invalidate();status('No se pudo iniciar el lector de Excel. Revisa la conexión y vuelve a elegir el archivo.','bad');};
      worker.onmessage=event=>{
        if(current!==revision)return;
        const msg=event.data;
        if(msg.tipo==='error') {
          busy=false;el('download').disabled=false;
          for(const button of el('separate').querySelectorAll('button'))button.disabled=false;
          status(msg.error,'bad');return;
        }
        if(msg.tipo==='listo')showResult(msg.resultado);
        if(msg.tipo==='excel') {
          const filename=msg.solo?`Reservas_${msg.solo}.xlsx`:'Reservas_consolidado.xlsx';
          const url=URL.createObjectURL(new Blob([msg.buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
          const a=document.createElement('a');a.href=url;a.download=filename;a.click();
          setTimeout(()=>URL.revokeObjectURL(url),30000);
          busy=false;el('download').disabled=false;
          for(const button of el('separate').querySelectorAll('button'))button.disabled=false;
          el('downloadStatus').textContent=`Descargado: ${filename}.`;
        }
      };
      worker.postMessage({tipo:'armar',stock,detalle,reglas:defaults},detalle?[stock,detalle]:[stock]);
    }catch(error){if(current===revision)status('No pude leer los archivos: '+error.message,'bad');}
  }
  function showResult(result) {
    showRules(result.reglas,!!detailFile);
    status('Reservas calculadas. Revisa los totales y descarga el Excel.','ok');
    el('source').textContent=`${stockFile.name} · ${fmt(result.filas)} filas de stock · Hoja ${result.hojaOrigen}`;
    el('summary').replaceChildren();
    for(const r of result.resumen) {
      const tr=document.createElement('tr');
      [r.hoja,fmt(r.modelos),fmt(r.unidades),fmt(r.sinModelo)].forEach((text,i)=>{
        const td=document.createElement('td');td.textContent=text;if(i)td.className='n';tr.append(td);
      });
      el('summary').append(tr);
    }
    el('warnings').replaceChildren();
    if(result.sinModelo.length) {
      const p=document.createElement('p');
      p.textContent=`Hay ${result.sinModelo.length===1?'1 fila':fmt(result.sinModelo.length)+' filas'} sin Modelo (${fmt(result.sinModelo.reduce((s,r)=>s+r.unidades,0))} unidades). Se conservan como “(en blanco)” en el Excel, pero no se subirán al sell out. Completa el Modelo en el stock para incluirlas.`;
      el('warnings').append(p);
    }
    el('diagnostics').replaceChildren();
    for(const r of result.resumen) {
      const p=document.createElement('p');p.textContent=`${r.hoja} · Almacenes definidos sin registros: ${r.ausentes.join(', ')||'ninguno'}.`;el('diagnostics').append(p);
    }
    const p=document.createElement('p');p.textContent='Excluidos por no estar en el detalle: '+(result.excluidos.map(([code,qty])=>`${code} (${fmt(qty)})`).join(', ')||'ninguno')+'.';el('diagnostics').append(p);
    el('separate').replaceChildren();
    for(const name of ['BO02','KAM','REGIONALES']) {
      const button=document.createElement('button');button.type='button';button.className='sec';button.textContent=name;
      button.onclick=()=>download(name);el('separate').append(button);
    }
    el('review').hidden=false;
  }
  function download(solo) {
    if(!worker||busy)return;
    busy=true;el('download').disabled=true;
    for(const button of el('separate').querySelectorAll('button'))button.disabled=true;
    el('downloadStatus').textContent='Generando Excel…';
    worker.postMessage({tipo:'excel',solo});
  }
  function setStock(file) {
    if(!file)return;
    if(!/\.xlsx$/i.test(file.name)){invalidate();stockFile=null;el('stockName').textContent='Selecciona un Excel .xlsx.';status('El archivo debe ser .xlsx.','bad');return;}
    stockFile=file;el('stockName').textContent=file.name;process();
  }
  el('stock').onchange=event=>{setStock(event.target.files[0]);event.target.value='';};
  el('detail').onchange=event=>{
    const file=event.target.files[0];if(!file)return;
    detailFile=file;el('resetRules').hidden=false;
    el('rulesStatus').textContent='Detalle adjunto: '+file.name+'. Se validará junto al stock.';
    el('rules').replaceChildren();el('shared').textContent='';process();
  };
  el('resetRules').onclick=()=>{detailFile=null;el('detail').value='';el('resetRules').hidden=true;if(defaults)showRules(defaults);process();};
  el('download').onclick=()=>download();
  for(const name of ['dragenter','dragover'])el('drop').addEventListener(name,e=>{e.preventDefault();el('drop').classList.add('over');});
  for(const name of ['dragleave','drop'])el('drop').addEventListener(name,e=>{e.preventDefault();el('drop').classList.remove('over');});
  el('drop').addEventListener('drop',event=>{
    const files=[...event.dataTransfer.files];
    if(files.length!==1){invalidate();status('Arrastra un solo archivo de stock.','bad');return;}
    setStock(files[0]);
  });
  fetch('data/reglas-reservas.json?v=1',{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error('HTTP '+r.status);return r.json();}).then(r=>{
    defaults=r;if(!detailFile)showRules(r);if(stockFile&&!detailFile)process();
  }).catch(()=>{
    if(!detailFile){el('rulesStatus').textContent='No se pudieron descargar las reglas. Adjunta detalle reservas.xlsx para continuar.';el('rulesStatus').className='status warn';}
  });
})();
