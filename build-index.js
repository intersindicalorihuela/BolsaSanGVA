#!/usr/bin/env node
// Uso: node scripts/build-index.js <ruta-excel> <ruta-index.html>
//
// Lee TODAS las hojas del Excel. Cada hoja con al menos una fila válida se
// convierte en una categoría/bolsa de la app; las hojas vacías o sin
// columnas reconocibles se ignoran automáticamente.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ROOT = path.resolve(__dirname, '..');
const EXCEL_PATH = process.argv[2] || path.join(ROOT, 'data', 'bolsa.xls');
const HTML_PATH = process.argv[3] || path.join(ROOT, 'index.html');
const RULES_PATH = path.join(__dirname, 'reglas.json');

function normalizeHeader(str) {
  return String(str || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function findColumnIndexes(headerRow) {
  const norm = headerRow.map(normalizeHeader);
  const find = (needle) => norm.findIndex(h => h.includes(needle));
  return {
    puesto: find('PUESTO'),
    nombre: find('NOMBRE'),
    situacion: find('SITUACION'),
    categoria: find('CATEGORIA'),
    departamento: find('DEPARTAMENTO')
  };
}

function parseSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (rows.length < 2) return [];
  const idx = findColumnIndexes(rows[0]);
  if (idx.puesto === -1 || idx.nombre === -1) return []; // no es una hoja de datos reconocible

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length === 0) continue;
    const puesto = parseInt(r[idx.puesto], 10);
    if (isNaN(puesto)) continue;
    const nombre = String(r[idx.nombre] || '').trim();
    const situacion = idx.situacion > -1 ? String(r[idx.situacion] || '').trim() : '';
    let categoria = idx.categoria > -1 ? String(r[idx.categoria] || '').trim() : '';
    let departamento = idx.departamento > -1 ? String(r[idx.departamento] || '').trim() : '';
    if (/^categor/i.test(categoria)) categoria = '';
    if (/^departamento/i.test(departamento)) departamento = '';
    out.push([puesto, nombre, situacion, categoria, departamento]);
  }
  return out;
}

function resolveCategory(sheetName, aliasHojas) {
  const alias = aliasHojas[sheetName];
  if (alias) {
    return {
      id: alias.id || sheetName,
      label: alias.label || sheetName,
      label_va: alias.label_va || alias.label || sheetName
    };
  }
  return { id: sheetName, label: sheetName, label_va: sheetName };
}

function buildSituaciones(allRows, base) {
  const set = new Set(base);
  allRows.forEach(r => { if (r[2]) set.add(r[2]); });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'es'));
}

// Extrae el objeto JSON balanceado que empieza en el primer '{' a partir de startIdx
function extractJsonSpan(str, startIdx) {
  let i = startIdx;
  while (i < str.length && str[i] !== '{') i++;
  if (i >= str.length) throw new Error("No se encontró '{' de BOLSA_DATA");
  let depth = 0, inStr = false, esc = false;
  const start = i;
  for (; i < str.length; i++) {
    const c = str[i];
    if (inStr) {
      if (esc) esc = false;
      else if (c === '\\') esc = true;
      else if (c === '"') inStr = false;
    } else {
      if (c === '"') inStr = true;
      else if (c === '{') depth++;
      else if (c === '}') { depth--; if (depth === 0) return { start, end: i + 1 }; }
    }
  }
  throw new Error('BOLSA_DATA no está bien formado (llaves sin cerrar)');
}

function injectData(html, newData) {
  const marker = 'const BOLSA_DATA = ';
  const markerIdx = html.indexOf(marker);
  if (markerIdx === -1) throw new Error('No se encontró "const BOLSA_DATA = " en index.html');
  const span = extractJsonSpan(html, markerIdx + marker.length);
  return html.slice(0, span.start) + JSON.stringify(newData) + html.slice(span.end);
}

function main() {
  if (!fs.existsSync(EXCEL_PATH)) {
    console.error('No existe el Excel en', EXCEL_PATH);
    process.exit(1);
  }
  const rules = JSON.parse(fs.readFileSync(RULES_PATH, 'utf8'));
  const wb = XLSX.read(fs.readFileSync(EXCEL_PATH), { type: 'buffer' });

  const categories = [];
  const candidates = {};
  const allRows = [];
  const ignoradas = [];

  wb.SheetNames.forEach(sheetName => {
    const rows = parseSheet(wb.Sheets[sheetName]);
    if (rows.length === 0) { ignoradas.push(sheetName); return; }
    const cat = resolveCategory(sheetName, rules.aliasHojas || {});
    categories.push(cat);
    candidates[cat.id] = rows;
    allRows.push(...rows);
  });

  if (categories.length === 0) {
    console.error('Ninguna hoja del Excel tiene columnas reconocibles (Puesto/Nombre). Revisa el archivo.');
    process.exit(1);
  }

  const situaciones = buildSituaciones(allRows, rules.situacionesBase || []);
  const newData = {
    generated: new Date().toISOString().slice(0, 10),
    situaciones,
    categories,
    candidates
  };

  const html = fs.readFileSync(HTML_PATH, 'utf8');
  fs.writeFileSync(HTML_PATH, injectData(html, newData), 'utf8');

  console.log('index.html actualizado. Bolsas encontradas con datos:');
  categories.forEach(c => console.log(' -', c.label, '(' + c.id + '):', candidates[c.id].length, 'candidatos'));
  if (ignoradas.length) {
    console.log('Hojas ignoradas (sin datos o sin columnas reconocibles):', ignoradas.join(', '));
  }
}

main();
