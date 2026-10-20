/* Contexto de sucursal para la descarga; no modifica las filas del dashboard. */
(function(global){
 const key=r=>JSON.stringify([r._w,r.m,r.dep,r.b,r.g,r.st,r.du,r.te]);
 function preparar(source,seleccion){
  const index=new Map();
  for(const r of source){
   if(!r.suc||(seleccion.size&&!seleccion.has(r.suc)))continue;
   const k=key(r);let retailers=index.get(k);if(!retailers){retailers=new Map();index.set(k,retailers);}
   for(const ret of Object.keys(r.r)){if(!retailers.has(ret))retailers.set(ret,new Set());retailers.get(ret).add(r.suc);}
  }
  return (row,rets)=>{
   const candidates=rets.filter(ret=>row.r[ret]&&((row.r[ret][0]||0)>0||(row.r[ret][3]||0)>0));
   const best=candidates.reduce((a,b)=>!a||(row.r[b][0]||0)>(row.r[a][0]||0)?b:a,'');
   const names=index.get(key(row))?.get(best);
   return [best||'–',names?.size?[...names].sort().join(' | '):seleccion.size?'Selección: '+[...seleccion].sort().join(' | '):'Sin detalle de sucursal'];
  };
 }
 global.SucursalExport={preparar};
})(typeof window!=='undefined'?window:globalThis);
