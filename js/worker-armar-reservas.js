importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js','armar-reservas.js?v=1');
let resultado=null;
self.onmessage=function(event) {
  try {
    const {tipo,stock,detalle,reglas,solo}=event.data;
    if(tipo==='armar') {
      resultado=null;
      const rules=detalle?ArmarReservas.leerDetalle(XLSX.read(detalle,{type:'array',dense:true}),XLSX):reglas;
      const wb=XLSX.read(stock,{type:'array',dense:true});
      resultado=ArmarReservas.armar(wb,rules,XLSX);
      self.postMessage({tipo:'listo',resultado});
    } else if(tipo==='excel') {
      if(!resultado)throw new Error('Primero adjunta el stock y genera la revisión.');
      const buffer=XLSX.write(ArmarReservas.libro(resultado,XLSX,solo),{type:'array',bookType:'xlsx',compression:true});
      self.postMessage({tipo:'excel',buffer,solo},[buffer]);
    } else throw new Error('Operación no válida.');
  }catch(error){self.postMessage({tipo:'error',error:error.message||String(error)});}
};
