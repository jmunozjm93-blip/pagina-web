/* Conversión de los Excel semanales de sell out al JSON que consume el dashboard.
 * Corre en el navegador sobre libros abiertos con SheetJS (XLSX.read).
 * Regla de oro: los números son los del Excel, tal cual. No se estima nada.
 */
(function (global) {
  'use strict';

  // ---------- nombres de archivo ----------
  const TIPOS = [
    { tipo: 'calzado', re: /^Reporte de zapatilla SellOut\s*w\s*(\d{1,2})\.*\.xlsx$/i, etiqueta: 'Reporte de zapatilla SellOut wNN.xlsx' },
    { tipo: 'ropa',    re: /^Reporte de Ropa SellOut\s*w\s*(\d{1,2})\.*\.xlsx$/i,      etiqueta: 'Reporte de Ropa SellOut wNN.xlsx' },
    { tipo: 'acc',     re: /^Reporte de Accesorios SellOut\s*w\s*(\d{1,2})\.*\.xlsx$/i, etiqueta: 'Reporte de Accesorios SellOut WNN.xlsx' },
    { tipo: 'tiendas', re: /^Reporte por tienda\s*w\s*(\d{1,2})\.*\.xlsx$/i,           etiqueta: 'Reporte por tienda wNN.xlsx' },
    { tipo: 'topsuc',  re: /^reporte por SKU por tienda\s*w\s*(\d{1,2})\.*\.xlsx$/i,   etiqueta: 'reporte por SKU por tienda wNN.xlsx' },
  ];
  function identificar(nombre) {
    for (const t of TIPOS) { const m = nombre.match(t.re); if (m) return { tipo: t.tipo, semana: String(+m[1]) }; }
    return null;
  }

  // ---------- utilitarios ----------
  // redondeo a 4 decimales igual que Python: los empates exactos (.5) van al par
  function r4(x) {
    if (typeof x !== 'number' || !isFinite(x)) return (x == null || x === '') ? null : x;
    if (Number.isInteger(x)) return x;
    // toFixed usa la expansión decimal exacta del double; solo los empates exactos
    // (p. ej. 0.03125) necesitan la regla del par
    const s = x.toFixed(30);
    const resto = s.slice(s.indexOf('.') + 5);
    if (/^50*$/.test(resto)) {
      const f = Math.floor(Math.abs(x) * 1e4), e = (f % 2 === 0 ? f : f + 1) / 1e4;
      return x < 0 ? -e : e;
    }
    return Number(x.toFixed(4));
  }
  const vacio = v => v == null || v === '';
  const txt = v => vacio(v) ? '' : String(v).trim();
  const num = v => (typeof v === 'number') ? v : (vacio(v) ? null : (isNaN(+v) ? v : +v));

  // hoja -> arreglo de filas (arreglo de valores crudos), sin depender de modo denso o disperso
  function filasDe(ws) {
    if (ws['!data']) {
      return ws['!data'].map(row => row ? row.map(c => (c ? c.v : null)) : []);
    }
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }
  const norm = s => txt(s).toLowerCase().replace(/\s+/g, '');

  // ---------- líneas de producto: calzado / ropa / acc ----------
  const ATTR = { 'modelo': 'm', 'descripcion': 'd', 'preciopubblanco_3': 'fp', 'status': 'st', 'duracion': 'du',
                 'temporada': 'te', 'grupogenero': 'g', 'departamentointerno': 'dep', 'marca': 'b' };
  // posición dentro del arreglo de 6: [vtaUN, rotación, margen, stockActual, semanasStock, vtaNeta]
  const MET = { 'vtaun': 0, 'rotacionun': 1, 'margen': 2, 'stockactualun': 3, 'stkfinalun': 3, 'semanasdestock': 4, 'vtaneta': 5 };
  const HOJAS_RE = [/^w\s*x?\d*\(tienda\+\.com\)$/, /^tiendaw\s*x?\d*$/, /^w\s*x?\d*\.com$/,
                    /^acumulado\(tienda\+\.com\)$/, /^acumuladotienda$/, /^acumulado\.com$/];
  const nombresHojas = w => [`W${w}(tienda + .com)`, `Tienda W${w}`, `W${w}.com`,
                             'Acumulado (tienda + .com)', 'Acumulado tienda', 'Acumulado.com'];

  function validarLinea(wb) {
    const errs = [];
    const names = wb.SheetNames;
    if (names.length !== 6) errs.push(`Se esperaban 6 hojas y el archivo tiene ${names.length}: ${names.join(', ')}`);
    names.slice(0, 6).forEach((n, i) => {
      if (HOJAS_RE[i] && !HOJAS_RE[i].test(norm(n))) errs.push(`Hoja ${i + 1} debería ser "${nombresHojas('NN')[i]}" y es "${n}"`);
    });
    return errs;
  }

  function parseHojaLinea(ws, nombre) {
    const rows = filasDe(ws);
    const h = rows.findIndex(r => r && norm(r[0]) === 'modelo');
    if (h < 0) throw new Error(`Hoja "${nombre}": no encontré la fila de encabezados (columna A = "Modelo")`);
    const hdr = rows[h], retRow = rows[h - 1] || [];
    // columnas de atributos
    const attrCols = [];
    let c = 0;
    for (; c < hdr.length; c++) {
      const k = ATTR[norm(hdr[c])];
      if (!k) break;
      attrCols.push([c, k]);
    }
    if (!attrCols.some(a => a[1] === 'm')) throw new Error(`Hoja "${nombre}": falta la columna Modelo`);
    // bloques por cliente
    const blocks = [];   // {ret, cols:[[col, idx]]}
    let cur = null;
    for (; c < hdr.length; c++) {
      const hn = norm(hdr[c]);
      if (!hn) continue;
      if (hn.startsWith('total')) break;
      const rn = txt(retRow[c]);
      if (rn) {
        if (/^total/i.test(rn)) break;
        cur = { ret: rn, cols: [] }; blocks.push(cur);
      }
      if (!cur) throw new Error(`Hoja "${nombre}": la columna ${c + 1} (${hdr[c]}) no tiene cliente encima`);
      const mi = MET[hn];
      if (mi === undefined) throw new Error(`Hoja "${nombre}": métrica desconocida "${hdr[c]}"`);
      cur.cols.push([c, mi]);
    }
    if (!blocks.length) throw new Error(`Hoja "${nombre}": no encontré bloques por cliente`);
    const retailers = blocks.map(b => b.ret);
    const leer = (row, b) => {
      const a = [null, null, null, null, null, null];
      let any = false;
      for (const [col, mi] of b.cols) { const v = row[col]; if (!vacio(v)) { any = true; a[mi] = r4(num(v)); } }
      return any ? a : null;
    };
    const out = [];
    let gt = null, prevAttrs = null;
    for (let i = h + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      const m = txt(row[0]);
      if (norm(m) === 'totalgeneral') {
        gt = {};
        // en el total solo entran los clientes con algún valor distinto de cero
        for (const b of blocks) { const a = leer(row, b); if (a && a.some(x => x)) gt[b.ret] = a; }
        gt._TOTAL = totalPonderado(gt);
        break;
      }
      if (row.every(vacio)) continue;
      const o = {};
      if (m) {
        for (const [col, k] of attrCols) {
          const v = row[col];
          if (k === 'm') o.m = m;
          else if (k === 'fp') { const n = num(v); if (n != null && n !== 0) o.fp = n; }   // precio 0 = sin precio
          else if (!vacio(v)) o[k] = txt(v);
        }
        prevAttrs = Object.assign({}, o); delete prevAttrs.dep;
      } else {
        // dinámica anidada Modelo > Departamento: el modelo y sus atributos vienen solo en la
        // primera fila del grupo; las siguientes traen únicamente el departamento
        if (!prevAttrs) continue;
        Object.assign(o, prevAttrs);
        const depCol = attrCols.find(a => a[1] === 'dep');
        if (depCol && !vacio(row[depCol[0]])) o.dep = txt(row[depCol[0]]);
      }
      o.r = {};
      for (const b of blocks) { const a = leer(row, b); if (a) o.r[b.ret] = a; }
      out.push(o);
    }
    if (!gt) throw new Error(`Hoja "${nombre}": no encontré la fila "Total general"`);
    return { retailers, rows: out, gt };
  }

  // _TOTAL: suma de UN, $ y stock; rotación ponderada por venta UN, margen por venta neta,
  // semanas de stock ponderadas por stock (o stock/venta si ningún cliente trae semanas)
  function totalPonderado(gt) {
    let v = 0, stk = 0, vn = 0, rotW = 0, mgW = 0, wosW = 0, wosS = 0;
    for (const k in gt) {
      const a = gt[k];
      v += a[0] || 0; stk += a[3] || 0; vn += a[5] || 0;
      if (a[1] != null) rotW += a[1] * (a[0] || 0);
      if (a[2] != null) mgW += a[2] * (a[5] || 0);
      if (a[4] != null) { wosW += a[4] * (a[3] || 0); wosS += a[3] || 0; }
    }
    // rotación y margen se dividen por el total completo (un cliente sin indicador pesa 0);
    // semanas de stock solo por el stock de los clientes que sí traen el indicador
    return [v, v ? r4(rotW / v) : null, vn ? r4(mgW / vn) : null, stk,
            wosS ? r4(wosW / wosS) : (v ? r4(stk / v) : null), r4(vn)];
  }

  function parseLinea(wb, semana) {
    const errs = validarLinea(wb);
    if (errs.length) throw new Error(errs.join('\n'));
    const names = nombresHojas(semana);
    const out = {};
    wb.SheetNames.slice(0, 6).forEach((sn, i) => { out[names[i]] = parseHojaLinea(wb.Sheets[sn], sn); });
    rellenarAtributos(out);
    return out;
  }

  // Una hoja puede traer la fila sin atributos y otra hoja del mismo archivo sí:
  // se completa con lo que el propio modelo tiene en las demás hojas.
  const ATTR_KEYS = ['d', 'st', 'du', 'te', 'g', 'b', 'fp', 'dep'];
  function rellenarAtributos(hojas) {
    const cat = {};
    for (const s in hojas) for (const r of hojas[s].rows) {
      const c = (cat[r.m] = cat[r.m] || {});
      for (const k of ATTR_KEYS) if (r[k] != null && r[k] !== '' && c[k] == null) c[k] = r[k];
    }
    let n = 0;
    for (const s in hojas) for (const r of hojas[s].rows) {
      const c = cat[r.m]; if (!c) continue;
      for (const k of ATTR_KEYS) if ((r[k] == null || r[k] === '') && c[k] != null) { r[k] = c[k]; n++; }
    }
    return n;
  }

  // ---------- relleno jerárquico de dinámicas ----------
  // Excel escribe cada nivel solo cuando cambia. El nivel más a la izquierda con valor es el que
  // cambió: los niveles a su izquierda se heredan, los de su derecha se toman de la fila (aunque vengan vacíos).
  function rellenar(row, cols, prev) {
    const out = new Array(cols.length);
    let k = cols.findIndex(c => !vacio(row[c]));
    if (k < 0) k = cols.length;   // fila sin dimensiones: hereda todo
    for (let i = 0; i < cols.length; i++) out[i] = i < k ? (prev ? prev[i] : '') : txt(row[cols[i]]);
    return out;
  }

  // ---------- TIENDAS (hoja "tiendas" del Reporte por tienda) ----------
  const T_DIM = { 'tiposucursal': 't', 'linea': 'l', 'temporada': 'te', 'marca': 'm', 'nombrecliente': 'cli', 'duracion': 'du',
                  'status': 'st', 'genero': 'g', 'departamentointerno': 'd', 'nombresucursal': 's', 'nrosemana': 'w' };
  const T_DIMS = ['t', 'l', 'cli', 'm', 'te', 'g', 'du', 'st', 's', 'w', 'd'];
  const T_METS = ['Vta UN', 'Vta UN LY', 'Vta Neta', 'Vta Neta LY', 'Contribucion', 'Contribucion LY', 'Margen', 'Margen LY',
                  'Rotacion UN', 'Rotacion UN LY', 'Stk Inicial UN', 'Stk Inicial UN LY', 'Stk Inicial Neta', 'Stk Inicial Neta LY',
                  'Stock Actual UN', 'Semanas de Stock'];

  function parseTiendas(wb, semana) {
    const sn = wb.SheetNames.find(n => norm(n) === 'tiendas');
    if (!sn) throw new Error(`No encontré la hoja "tiendas". Hojas: ${wb.SheetNames.join(', ')}`);
    const rows = filasDe(wb.Sheets[sn]);
    const h = rows.findIndex(r => r && norm(r[0]) === 'tiposucursal');
    if (h < 0) throw new Error('Hoja "tiendas": no encontré la fila de encabezados (columna A = "Tipo Sucursal")');
    const hdr = rows[h];
    const dimCol = {}, metCol = {};
    hdr.forEach((hv, c) => { const k = T_DIM[norm(hv)]; if (k) dimCol[k] = c; });
    T_METS.forEach(mn => { const c = hdr.findIndex(hv => norm(hv) === norm(mn)); if (c >= 0) metCol[mn] = c; });
    const faltan = Object.keys(T_DIM).map(k => T_DIM[k]).filter(k => dimCol[k] === undefined);
    if (faltan.length) throw new Error('Hoja "tiendas": faltan columnas ' + faltan.join(', '));
    // orden de anidación = orden de las columnas en el Excel
    const nest = Object.entries(dimCol).sort((a, b) => a[1] - b[1]);
    const nestCols = nest.map(e => e[1]), nestKeys = nest.map(e => e[0]);
    const vals = {}, idx = {};
    T_DIMS.forEach(k => { vals[k] = []; idx[k] = new Map(); });
    const code = (k, v) => { let i = idx[k].get(v); if (i === undefined) { i = vals[k].length; vals[k].push(v); idx[k].set(v, i); } return i; };
    const out = [];
    let prev = null, total = 0;
    const wiCol = nestKeys.indexOf('w');
    for (let i = h + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      if (norm(row[0]) === 'totalgeneral') break;
      const d = rellenar(row, nestCols, prev); prev = d;
      total++;
      if (d[wiCol] !== String(semana)) continue;
      const rec = new Array(T_DIMS.length + T_METS.length);
      nestKeys.forEach((k, j) => { rec[T_DIMS.indexOf(k)] = code(k, d[j]); });
      T_METS.forEach((mn, j) => { const c = metCol[mn]; rec[T_DIMS.length + j] = c === undefined ? null : r4(num(row[c])); });
      out.push(rec);
    }
    if (!out.length) throw new Error(`Hoja "tiendas": no hay filas con Nro Semana = ${semana}`);
    return { dim: T_DIMS, vals, mets: T_METS, rows: out, gt: null, filasLeidas: total };
  }

  // ---------- Top-5 sucursales (reporte por SKU por tienda) ----------
  const CLI_KEY = { 'DEPOR': 'D', 'FALABELLA': 'F', 'RIPLEY': 'R', 'PARIS': 'P', 'LA POLAR': 'L', 'HITES': 'H', 'GE2': 'G' };
  function top5(list) {
    // orden: venta UN descendente y, a igual venta, nombre de sucursal ascendente
    return list.sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0)).slice(0, 5);
  }
  function parseTopSuc(wb, semana) {
    const sn = wb.SheetNames[0];
    const rows = filasDe(wb.Sheets[sn]);
    const h = rows.findIndex(r => r && norm(r[0]) === 'nombrecliente');
    if (h < 0) throw new Error('No encontré la fila de encabezados (columna A = "Nombre Cliente")');
    const hdr = rows[h].map(norm);
    const col = n => { const c = hdr.indexOf(norm(n)); if (c < 0) throw new Error(`Falta la columna "${n}"`); return c; };
    const cCli = col('Nombre Cliente'), cSuc = col('Nro Sucursal Nombre'), cSem = col('Nro Semana'), cMod = col('Modelo'),
          cUN = col('Vta UN'), cNeta = col('Vta Neta');
    const nestCols = [cCli, cSuc, cSem];
    const W = {}, ACC = {}, weeks = new Set();
    let prev = null;
    for (let i = h + 1; i < rows.length; i++) {
      const row = rows[i]; if (!row) continue;
      if (norm(row[0]) === 'totalgeneral') break;
      const d = rellenar(row, nestCols, prev); prev = d;
      const key = CLI_KEY[d[0].toUpperCase()];
      const w = d[2], mod = txt(row[cMod]).toUpperCase();
      if (!key || !w || !mod) continue;
      const un = num(row[cUN]) || 0, neta = r4(num(row[cNeta]) || 0);
      if (!un) continue;   // solo filas con unidades (las devoluciones, negativas, sí cuentan)
      weeks.add(w);
      if (w === String(semana)) {
        const a = (W[mod] = W[mod] || {}); (a[key] = a[key] || []).push([d[1], un, r4(neta)]);
      }
      const b = (ACC[mod] = ACC[mod] || {}); const m = (b[key] = b[key] || new Map());
      const e = m.get(d[1]) || [d[1], 0, 0]; e[1] += un; e[2] += neta; m.set(d[1], e);
    }
    for (const mod in W) for (const k in W[mod]) W[mod][k] = top5(W[mod][k]);
    const ACC2 = {};
    for (const mod in ACC) { ACC2[mod] = {}; for (const k in ACC[mod]) ACC2[mod][k] = top5([...ACC[mod][k].values()].map(e => [e[0], e[1], r4(e[2])])); }
    if (!Object.keys(W).length) throw new Error(`No hay ventas con Nro Semana = ${semana}`);
    const accWeeks = [...weeks].sort((a, b) => a - b);
    return { W: { [String(semana)]: W }, ACC: ACC2, accWeeks };
  }

  // ---------- semana declarada dentro del archivo (para cotejar con el nombre) ----------
  function semanaDelLibro(tipo, wb) {
    if (tipo === 'tiendas' || tipo === 'topsuc') return null;   // traen todas las semanas
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = filasDe(ws).slice(0, 12);
    for (const r of rows) {
      if (!r) continue;
      const k = norm(r[0]);
      if (k === 'nrosemana' && !vacio(r[1])) return String(+r[1]);
      if (k === 'periodosemana(nombre)') { const m = String(r[1] || '').match(/W(\d{1,2})/i); if (m) return String(+m[1]); }
    }
    return null;
  }

  // ---------- resumen para mostrar en pantalla ----------
  function resumen(tipo, json, semana) {
    if (tipo === 'tiendas') {
      const si = json.dim.indexOf('s'), ci = json.dim.indexOf('cli');
      const sucs = new Set(json.rows.map(r => r[si])), clis = new Set(json.rows.map(r => r[ci]));
      const iv = json.dim.length + json.mets.indexOf('Vta UN'), inn = json.dim.length + json.mets.indexOf('Vta Neta');
      const vta = json.rows.reduce((a, r) => a + (r[iv] || 0), 0), neta = json.rows.reduce((a, r) => a + (r[inn] || 0), 0);
      return { tipo, filas: json.rows.length, sucursales: sucs.size, clientes: clis.size, vtaUN: vta, vtaNeta: neta, filasLeidas: json.filasLeidas };
    }
    if (tipo === 'topsuc') {
      const W = json.W[String(semana)] || {};
      return { tipo, modelos: Object.keys(W).length, modelosAcc: Object.keys(json.ACC).length, accWeeks: json.accWeeks };
    }
    const hojas = Object.keys(json).map(h => {
      const s = json[h], t = s.gt._TOTAL || [];
      return { hoja: h, filas: s.rows.length, clientes: s.retailers.length, vtaUN: t[0], vtaNeta: t[5], stock: t[3] };
    });
    return { tipo, hojas };
  }

  // ---------- punto de entrada ----------
  function convertir(tipo, wb, semana) {
    switch (tipo) {
      case 'calzado': case 'ropa': case 'acc': return parseLinea(wb, semana);
      case 'tiendas': return parseTiendas(wb, semana);
      case 'topsuc': return parseTopSuc(wb, semana);
      default: throw new Error('Tipo desconocido: ' + tipo);
    }
  }
  // opciones de lectura para XLSX.read según el tipo (las hojas grandes solo se leen si hacen falta)
  function opcionesLectura(tipo) {
    const o = { dense: true, cellText: false, cellHTML: false, cellStyles: false };
    if (tipo === 'tiendas') o.sheets = ['tiendas'];
    return o;
  }

  global.Convertir = { TIPOS, identificar, convertir, opcionesLectura, nombresHojas, validarLinea, totalPonderado, rellenar, semanaDelLibro, resumen };
})(typeof window !== 'undefined' ? window : globalThis);
