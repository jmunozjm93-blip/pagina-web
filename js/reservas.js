/* Reservas: instantánea semanal. Solo modifica bo, kam y reg por Modelo. */
(function (global) {
  'use strict';
  const HOJAS = ['BO02', 'KAM', 'REGIONALES'];
  const norm = value => String(value == null ? '' : value).trim().toUpperCase();
  const propio = (object, key) => Object.prototype.hasOwnProperty.call(object, key);

  function convertir(wb, semana, XLSX) {
    if (!/^(?:[1-9]|[1-4]\d|5[0-3])$/.test(String(semana))) throw new Error('Elige una semana entre 1 y 53.');
    const nombres = {};
    for (const name of wb.SheetNames) {
      const key = norm(name) === 'B002' ? 'BO02' : norm(name);
      if (!HOJAS.includes(key) || propio(nombres, key)) throw new Error('El Excel debe tener exactamente las hojas BO02, KAM y REGIONALES.');
      nombres[key] = name;
    }
    if (HOJAS.some(key => !propio(nombres, key))) throw new Error('Faltan hojas. Sube el consolidado con BO02, KAM y REGIONALES.');
    const modelos = Object.create(null), resumen = [], omitidas = [];
    HOJAS.forEach((hoja, index) => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[nombres[hoja]], {header: 1, raw: true, defval: null});
      const headers = (rows.shift() || []).map(norm);
      if (headers[0] !== 'MODELO' || headers[1] !== 'AVAILABLE' || headers.slice(2).some(Boolean)) {
        throw new Error(`${hoja}: se esperan solo las columnas Modelo y Available.`);
      }
      const vistos = new Set();
      let total = 0, omitido = 0;
      rows.forEach((row, i) => {
        if (row.every(value => value === null || value === '')) return;
        const model = norm(row[0]), value = row[1];
        if (row.slice(2).some(value => value !== null && value !== '')) throw new Error(`${hoja}, fila ${i + 2}: hay datos fuera de las dos columnas.`);
        if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error(`${hoja}, fila ${i + 2}: Available debe ser un número.`);
        if (!model || model === '(EN BLANCO)') {
          omitidas.push({hoja, fila: i + 2, unidades: value});
          omitido += value;
          return;
        }
        if (model === 'TOTAL GENERAL') throw new Error(`${hoja}: quita la fila Total general.`);
        if (vistos.has(model)) throw new Error(`${hoja}: el Modelo ${model} está repetido.`);
        vistos.add(model);
        if (!propio(modelos, model)) modelos[model] = [0, 0, 0];
        modelos[model][index] = value;
        total += value;
      });
      resumen.push({hoja, modelos: vistos.size, unidades: total, unidadesOmitidas: omitido});
    });
    if (!Object.keys(modelos).length) throw new Error('El archivo no contiene modelos válidos.');
    return {version: 1, semana: String(semana), actualizado: new Date().toISOString(), modelos, resumen, omitidas};
  }

  function validar(data, semana) {
    if (!data || data.version !== 1 || data.semana !== String(semana) ||
        !data.modelos || typeof data.modelos !== 'object' || Array.isArray(data.modelos) ||
        !Object.keys(data.modelos).length) throw new Error('Archivo de Reservas inválido o de otra semana.');
    for (const [model, values] of Object.entries(data.modelos)) {
      if (!model || norm(model) !== model || !Array.isArray(values) || values.length !== 3 ||
          values.some(value => typeof value !== 'number' || !Number.isFinite(value))) {
        throw new Error('Reservas contiene un Modelo o una cantidad inválidos.');
      }
    }
    return data;
  }

  function aplicar(row, data) {
    if (!data) return row;
    const key = norm(row.m);
    const values = propio(data.modelos, key) ? data.modelos[key] : [0, 0, 0];
    row.bo = values[0];
    row.kam = values[1];
    row.reg = values[2];
    return row;
  }
  global.Reservas = {HOJAS, convertir, validar, aplicar};
})(typeof self !== 'undefined' ? self : globalThis);
