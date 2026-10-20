/* Exportación por modelo, retailer y sucursal sin sumar tiendas diferentes. */
(function(global){
 const key=r=>JSON.stringify([r._w,r.m,r.dep,r.b,r.g,r.st,r.du,r.te]);
 function desglosar(source,rows,seleccion,rets){
  const wanted=new Set(rows.map(key)), found=new Set(), result=[];
  for(const r of source){
   const k=key(r);
   if(!wanted.has(k)||(seleccion.size&&!seleccion.has(r.suc)))continue;
   if(!r.suc)continue;
   found.add(k);
   for(const ret of rets){
    if(!r.r[ret])continue;
    result.push({...r,r:{[ret]:[...r.r[ret]]},_exportRetailer:ret});
   }
  }
  if([...wanted].some(k=>!found.has(k)))throw new Error('Esta hoja no contiene el detalle de sucursales necesario. Vuelve a cargar el Excel con Nombre Sucursal para descargarlo desglosado.');
  return result;
 }
 global.SucursalExport={desglosar};
})(typeof window!=='undefined'?window:globalThis);
