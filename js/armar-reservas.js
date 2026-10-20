(function(global) {
  'use strict';
  const HOJAS=['BO02','KAM','REGIONALES'];
  const clave=value=>String(value==null?'':value).trim().toUpperCase();
  const norm=value=>clave(value).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');
  function validarReglas(reglas) {
    if(!reglas || reglas.version!==1 || !reglas.grupos)throw new Error('No se pudieron leer las reglas de reservas.');
    const grupos={};
    for(const hoja of HOJAS) {
      const source=reglas.grupos[hoja];
      if(!Array.isArray(source)||!source.length)throw new Error(`Faltan almacenes para ${hoja}.`);
      const codes=source.map(clave);
      if(codes.some(c=>!c)||new Set(codes).size!==codes.length)throw new Error(`${hoja}: hay códigos de almacén vacíos o duplicados.`);
      grupos[hoja]=codes;
    }
    return {version:1,fuente:reglas.fuente||'Detalle adjunto',grupos};
  }
  function leerDetalle(wb,XLSX) {
    const grupos={},aliases={BO02:'BO02',B002:'BO02',KAM:'KAM',KAMS:'KAM',REGIONES:'REGIONALES',REGIONALES:'REGIONALES'};
    for(const name of wb.SheetNames) {
      const hoja=aliases[norm(name)];
      if(!hoja)continue;
      if(grupos[hoja])throw new Error(`Hay más de una hoja de reglas para ${hoja}.`);
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:true,defval:null});
      grupos[hoja]=rows.filter(r=>r.some(v=>v!==null&&v!=='')).filter(r=>!norm(r[0]).startsWith('CODIGOALMACEN')).map(r=>{
        if(!clave(r[0]))throw new Error(`${name}: una regla no tiene código de almacén.`);
        return clave(r[0]);
      });
    }
    return validarReglas({version:1,grupos});
  }
  function armar(wb,reglas,XLSX) {
    reglas=validarReglas(reglas);
    let source=null,headers=null,sourceName='';
    for(const name of wb.SheetNames) {
      const sheet=wb.Sheets[name];
      // Read only the first row while locating the raw stock table.
      const extent=XLSX.utils.decode_range(sheet['!ref']||'A1');
      const head=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:null,range:{s:{r:0,c:0},e:{r:0,c:extent.e.c}}})[0]||[];
      const indexes=['CODIGODEALMACEN','MODELO','AVAILABLE'].map(key=>head.map(norm).indexOf(key));
      if(indexes.some(i=>i<0))continue;
      if(source)throw new Error('Hay más de una hoja con Código de almacén, Modelo y Available. Deja una sola tabla de stock.');
      source=sheet;headers=indexes;sourceName=name;
    }
    if(!source)throw new Error('No encuentro Código de almacén, Modelo y Available en la primera fila. Adjunta el archivo de stock original.');
    const rows=XLSX.utils.sheet_to_json(source,{header:1,raw:true,defval:null});
    const sets=Object.fromEntries(HOJAS.map(h=>[h,new Set(reglas.grupos[h])]));
    const sums=Object.fromEntries(HOJAS.map(h=>[h,new Map()]));
    const present=new Set(), excluded=new Map(), blanks=[];
    let count=0,total=0,matched=0;
    rows.slice(1).forEach((row,i)=>{
      if(row.every(v=>v===null||v===''))return;
      const code=clave(row[headers[0]]), model=clave(row[headers[1]]), value=row[headers[2]];
      if(!code)throw new Error(`Fila ${i+2}: falta Código de almacén.`);
      if(value!==null&&value!==''&&(typeof value!=='number'||!Number.isFinite(value)))throw new Error(`Fila ${i+2}: Available debe ser numérico.`);
      const qty=value==null||value===''?0:value;
      count++;total+=qty;present.add(code);
      const groups=HOJAS.filter(h=>sets[h].has(code));
      if(!groups.length){excluded.set(code,(excluded.get(code)||0)+qty);return;}
      matched++;
      if(!model)blanks.push({fila:i+2,almacen:code,unidades:qty,hojas:groups});
      for(const h of groups) sums[h].set(model||'(en blanco)',(sums[h].get(model||'(en blanco)')||0)+qty);
    });
    if(!count)throw new Error('La tabla de stock está vacía.');
    if(!matched)throw new Error('Ningún almacén del stock coincide con las reglas. Revisa el archivo de stock o el detalle.');
    const tablas={},resumen=[];
    for(const h of HOJAS) {
      const entries=[...sums[h]].sort((a,b)=>a[0]==='(en blanco)'?1:b[0]==='(en blanco)'?-1:a[0].localeCompare(b[0],'es'));
      tablas[h]=[['Modelo','Available'],...entries];
      resumen.push({hoja:h,modelos:entries.filter(r=>r[0]!=='(en blanco)').length,
        unidades:entries.reduce((s,r)=>s+r[1],0),sinModelo:sums[h].get('(en blanco)')||0,
        ausentes:reglas.grupos[h].filter(c=>!present.has(c))});
    }
    const compartidos=[...new Set(Object.values(reglas.grupos).flat())].filter(c=>HOJAS.filter(h=>sets[h].has(c)).length>1);
    return {tablas,resumen,sinModelo:blanks,excluidos:[...excluded].sort((a,b)=>a[0].localeCompare(b[0])),
      compartidos,filas:count,unidadesOrigen:total,hojaOrigen:sourceName,reglas};
  }
  function libro(resultado,XLSX,solo) {
    const wb=XLSX.utils.book_new();
    for(const name of solo?[solo]:HOJAS) {
      if(!HOJAS.includes(name))throw new Error('Hoja no válida.');
      const ws=XLSX.utils.aoa_to_sheet(resultado.tablas[name]);
      ws['!cols']=[{wch:25},{wch:18}];
      ws['!autofilter']={ref:ws['!ref']};
      for(let i=1;i<resultado.tablas[name].length;i++)if(ws['B'+(i+1)])ws['B'+(i+1)].z='#,##0;[Red]-#,##0;0';
      XLSX.utils.book_append_sheet(wb,ws,name);
    }
    return wb;
  }
  global.ArmarReservas={HOJAS,validarReglas,leerDetalle,armar,libro};
})(typeof self!=='undefined'?self:globalThis);
