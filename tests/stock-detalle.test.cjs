const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
require('../js/convertir.js');
test('W43: stock de detalle e indicadores de Agru, sin mutar fuentes',()=>{
 const read=f=>JSON.parse(fs.readFileSync(path.join(__dirname,'../data/w43/'+f+'.json')));
 const a=read('calzado_resumen'),b=read('calzado'),before=JSON.stringify(a),out=Convertir.resumenConStockDetalle(a,b);
 for(const n of Object.keys(a)){
  for(const ret of Object.keys(a[n].gt))for(let i=0;i<6;i++)assert.equal(out[n].gt[ret][i],i===3?(b[n].gt[ret]?.[i]??null):a[n].gt[ret][i]);
  a[n].rows.forEach((r,j)=>{for(const ret in r.r)for(const i of [0,1,2,4,5])assert.equal(out[n].rows[j].r[ret][i],r.r[ret][i]);});
 }
 const gt=out[Object.keys(out)[0]].gt._TOTAL;assert.equal(gt[0],9050);assert.equal(gt[3],304239);assert.equal(gt[5],354639841);assert.equal(JSON.stringify(a),before);
});
test('Vacíos, ceros, negativos, modelos exclusivos y duplicados',()=>{
 const row=(m,v)=>({m,dep:'D',r:{R:v}}),sheet=rows=>({H:{rows,retailers:['R'],gt:{R:[1,.1,.2,99,10,100]}}});
 const a=sheet([row('A',[1,.1,.2,9,10,100]),row('B',[2,.2,.3,8,5,200]),row('B',[3,.3,.4,7,4,300]),row('C',[0,0,null,4,null,0])]);
 const b=sheet([row('A',[2,.2,.3,-2,5,200]),row('A',[1,.1,.2,2,10,100]),row('B',[1,.1,.2,10,10,100]),row('D',[4,.2,.3,7,5,500])]);
 const out=Convertir.resumenConStockDetalle(a,b).H.rows;
 assert.equal(out[0].r.R[3],0);assert.equal(out[1].r.R[3],null);assert.equal(out[2].r.R[3],null);assert.equal(out[3].r.R[3],null);
 assert.deepEqual(out[4].r.R,[null,null,null,7,null,null]);
});
