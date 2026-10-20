// Ejecutar: node --test tests/reservas.test.cjs
const test = require('node:test');
const assert = require('node:assert/strict');
require('../js/reservas.js');
const XLSX = {utils:{sheet_to_json:sheet=>structuredClone(sheet)}};
const base = () => ({
  BO02:[['Modelo','Available'],['001-A',5],['(en blanco)',13]],
  KAM:[['Modelo','Available'],['001-A',0],['OTRO',-2]],
  REGIONALES:[['Modelo','Available'],['001-A',9]],
});
const parse = sheets => Reservas.convertir({SheetNames:Object.keys(sheets),Sheets:sheets},'42',XLSX);
test('une por Modelo y conserva ceros y negativos; informa filas sin modelo',()=>{
  const data=parse(base());
  assert.deepEqual(data.modelos['001-A'],[5,0,9]);
  assert.deepEqual(data.modelos.OTRO,[0,-2,0]);
  assert.equal(data.omitidas[0].unidades,13);
  assert.ok(!data.modelos['(EN BLANCO)']);
  assert.equal(Reservas.validar(data,'42'),data);
  assert.throws(()=>Reservas.validar(data,'41'));
});
test('modifica exclusivamente las tres columnas y reinicia modelos ausentes',()=>{
  const data=parse(base());
  const row={m:'001-a',bo:99,kam:99,reg:99,fp:9990,r:{retailer:[1,2,3,4,5,6]},d:'Producto'};
  const before=structuredClone(row);
  Reservas.aplicar(row,data);
  assert.deepEqual(row,{...before,bo:5,kam:0,reg:9});
  const missing={m:'AUSENTE',bo:99,kam:99,reg:99,fp:42};
  Reservas.aplicar(missing,data);
  assert.deepEqual(missing,{m:'AUSENTE',bo:0,kam:0,reg:0,fp:42});
  const historical=structuredClone(before);
  Reservas.aplicar(historical,null);
  assert.deepEqual(historical,before);
});
test('rechaza filas duplicadas, totales, cantidades inválidas y hojas incompletas',()=>{
  for(const row of [['001-a',5],['Total general',7],['X','7'],['X',Infinity]]) {
    const sheets=base();sheets.BO02.push(row);
    assert.throws(()=>parse(sheets));
  }
  const missing=base();delete missing.REGIONALES;
  assert.throws(()=>parse(missing));
  const extra=base();extra.REGIONALES[0].push('Extra');
  assert.throws(()=>parse(extra));
});
