const {test}=require('node:test');
const assert=require('node:assert/strict');
require('../js/convertir.js');
test('Agru y detalle tienen destinos distintos de la misma semana',()=>{
 assert.deepEqual(Convertir.identificar('Reporte de zapatilla SellOut w42 Agru.xlsx'),{tipo:'calzado_resumen',semana:'42'});
 assert.deepEqual(Convertir.identificar('Reporte de zapatilla SellOut w42..xlsx'),{tipo:'calzado',semana:'42'});
});
test('Rechaza totales, hojas o indicadores diferentes, conservando vacíos',()=>{
 const a={Hoja:{gt:{RIPLEY:[92,.027,.38,501,36.7,3980304],_TOTAL:[92,.027,.38,501,36.7,3980304]}}};
 assert(Convertir.parCompatible(a,structuredClone(a)));
 const b=structuredClone(a);b.Hoja.gt.RIPLEY[0]++;assert(!Convertir.parCompatible(a,b));
 assert(!Convertir.parCompatible(a,{}));
 b.Hoja.gt.RIPLEY[0]=92;b.Hoja.gt.RIPLEY[1]=null;assert(!Convertir.parCompatible(a,b));
});
