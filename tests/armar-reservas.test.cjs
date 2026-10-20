const test=require('node:test'),assert=require('node:assert/strict');
require('../js/armar-reservas.js');
const XLSX={utils:{
  decode_range:()=>({e:{c:2}}),
  sheet_to_json:(sheet,options)=>structuredClone(options.range?sheet.rows.slice(0,1):sheet.rows),
}};
const book=rows=>({SheetNames:['Stock'],Sheets:{Stock:{rows,'!ref':'A1:C10'}}});
const rules={version:1,grupos:{BO02:['BO02'],KAM:['C_DEPORT','9000'],REGIONALES:['C_DEPORT','C_ALMAND']}};
const head=['Código de almacén','Modelo','Available'];
test('suma por modelo, aplica cada filtro y conserva el almacén compartido',()=>{
  const data=ArmarReservas.armar(book([head,['BO02','001-A',2],['BO02','001-A',3],
    ['C_DEPORT','001-A',7],['9000','001-A',4],['C_ALMAND','NEG',-2],
    ['BO02',null,13],['OTRO','EXCLUIDO',99],['BO02','CERO',null]]),rules,XLSX);
  assert.deepEqual(data.tablas.BO02,[['Modelo','Available'],['001-A',5],['CERO',0],['(en blanco)',13]]);
  assert.deepEqual(data.tablas.KAM,[['Modelo','Available'],['001-A',11]]);
  assert.deepEqual(data.tablas.REGIONALES,[['Modelo','Available'],['001-A',7],['NEG',-2]]);
  assert.deepEqual(data.compartidos,['C_DEPORT']);
  assert.deepEqual(data.excluidos,[['OTRO',99]]);
  assert.equal(data.sinModelo[0].unidades,13);
});
test('acepta detalle sin encabezado de Regiones y normaliza nombres',()=>{
  const wb={SheetNames:['KAMS','BO02','Regiones'],Sheets:{
    KAMS:{rows:[['código almacen','Nombre','Tipo'],['C_DEPORT','Gonzalo','KAM']]},
    BO02:{rows:[['código almacen','Nombre','Tipo'],['BO02','BO02','BO02']]},
    Regiones:{rows:[['C_DEPORT','gonzalo y vicente nuevo','REGIONALES']]},
  }};
  const result=ArmarReservas.leerDetalle(wb,XLSX);
  assert.deepEqual(result.grupos.REGIONALES,['C_DEPORT']);
  assert.deepEqual(result.grupos.KAM,['C_DEPORT']);
  wb.Sheets.KAMS.rows.push(['C_DEPORT','Repetido','KAM']);
  assert.throws(()=>ArmarReservas.leerDetalle(wb,XLSX),/duplicados/);
});
test('rechaza datos inválidos y reglas incompletas sin producir un resultado anterior',()=>{
  assert.throws(()=>ArmarReservas.armar(book([head,['BO02','X','12']]),rules,XLSX),/numérico/);
  assert.throws(()=>ArmarReservas.armar(book([head,[null,'X',12]]),rules,XLSX),/almacén/);
  assert.throws(()=>ArmarReservas.armar(book([head,['OTRO','X',12]]),rules,XLSX),/Ningún/);
  assert.throws(()=>ArmarReservas.armar(book([['Modelo','Available']]),rules,XLSX),/No encuentro/);
  assert.throws(()=>ArmarReservas.validarReglas({version:1,grupos:{BO02:['BO02']}}),/Faltan/);
});
test('una nueva ejecución reemplaza los valores en vez de acumular reservas anteriores',()=>{
  const first=ArmarReservas.armar(book([head,['BO02','001-A',5]]),rules,XLSX);
  const second=ArmarReservas.armar(book([head,['BO02','NUEVO',9]]),rules,XLSX);
  assert.equal(first.tablas.BO02[1][1],5);
  assert.deepEqual(second.tablas.BO02,[head.slice(1),['NUEVO',9]]);
});
