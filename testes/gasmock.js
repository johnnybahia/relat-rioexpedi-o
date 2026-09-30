// Simulador mínimo do Apps Script para rodar Código.gs em Node: planilha em memória
// (SpreadsheetApp), Utilities, PropertiesService, CacheService, LockService, ScriptApp.
// Não reproduz: tipos de célula fora do que o CSV traz, limites de tempo/cota, gatilhos reais.
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

const ARQ_CODIGO = process.env.CODIGO_GS || path.join(__dirname, '..', 'Código.gs'); // CODIGO_GS: teste de mutação
const TZ_OFFSET_H = -3; // America/Fortaleza (sem horário de verão)

function pad(n, w = 2) { return String(n).padStart(w, '0'); }
function partesFortaleza(d) {
  const t = new Date(d.getTime() + TZ_OFFSET_H * 3600 * 1000);
  return { y: t.getUTCFullYear(), M: t.getUTCMonth() + 1, d: t.getUTCDate(), H: t.getUTCHours(), m: t.getUTCMinutes(), s: t.getUTCSeconds(), dow: t.getUTCDay() };
}
function formatDate(d, tz, pat) {
  const p = partesFortaleza(d);
  return pat.replace(/yyyy|MM|dd|HH|mm|ss|H|u/g, tok => ({
    yyyy: p.y, MM: pad(p.M), dd: pad(p.d), HH: pad(p.H), mm: pad(p.m), ss: pad(p.s), H: String(p.H), u: String(p.dow === 0 ? 7 : p.dow)
  })[tok]);
}
function dataFortaleza(y, M, d, H = 0, m = 0, s = 0) { return new Date(Date.UTC(y, M - 1, d, H - TZ_OFFSET_H, m, s)); }

function colunaParaNumero(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
function lerA1(a1) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a1);
  if (!m) throw new Error('A1 inválido: ' + a1);
  const c1 = colunaParaNumero(m[1]), r1 = Number(m[2]);
  const c2 = m[3] ? colunaParaNumero(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}
function exibir(v) {
  if (v instanceof Date) return formatDate(v, '', 'dd/MM/yyyy');
  if (v === null || v === undefined) return '';
  return String(v);
}

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = this.sh.data[this.r - 1 + i] || [];
      const o = [];
      for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; o.push(v === undefined || v === null ? '' : v); }
      out.push(o);
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map(r => r.map(exibir)); }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return exibir(this.getValue()); }
  setValues(vals) {
    if (vals.length !== this.nr || vals.some(r => r.length !== this.nc)) {
      throw new Error(`setValues: dimensões ${vals.length}x${vals[0] && vals[0].length} ≠ ${this.nr}x${this.nc} (${this.sh.name})`);
    }
    for (let i = 0; i < this.nr; i++) {
      const ri = this.r - 1 + i;
      while (this.sh.data.length <= ri) this.sh.data.push([]);
      for (let j = 0; j < this.nc; j++) this.sh.data[ri][this.c - 1 + j] = vals[i][j];
    }
    this.sh.maxCols = Math.max(this.sh.maxCols, this.c + this.nc - 1);
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  clearContent() {
    for (let i = 0; i < this.nr; i++) {
      const row = this.sh.data[this.r - 1 + i];
      if (!row) continue;
      for (let j = 0; j < this.nc; j++) row[this.c - 1 + j] = '';
    }
    return this;
  }
  clear() { return this.clearContent(); }
  clearContents() { return this.clearContent(); }
  setNumberFormat() { return this; }
  setFontWeight() { return this; } setFontStyle() { return this; } setFontColor() { return this; }
  setBackground() { return this; } setFontSize() { return this; } setWrap() { return this; }
  setHorizontalAlignment() { return this; } setBorder() { return this; } setNote() { return this; }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
}

class Sheet {
  constructor(name, data, maxCols) { this.name = name; this.data = data || []; this.maxCols = Math.max(maxCols || 0, ...this.data.map(r => r.length), 1); }
  getName() { return this.name; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const [r, cc, nr, nc] = lerA1(a); return new Range(this, r, cc, nr, nc); }
    return new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d);
  }
  getLastRow() {
    for (let i = this.data.length - 1; i >= 0; i--) {
      const row = this.data[i];
      if (row && row.some(v => v !== '' && v !== null && v !== undefined)) return i + 1;
    }
    return 0;
  }
  getLastColumn() {
    let m = 0;
    this.data.forEach(row => {
      if (!row) return;
      for (let j = row.length - 1; j >= 0; j--) {
        const v = row[j];
        if (v !== '' && v !== null && v !== undefined) { m = Math.max(m, j + 1); break; }
      }
    });
    return m;
  }
  getMaxColumns() { return Math.max(this.maxCols, this.getLastColumn()); }
  getMaxRows() { return Math.max(this.data.length, 1000); }
  getDataRange() { return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  deleteRows(start, n) { this.data.splice(start - 1, n); }
  deleteRow(r) { this.data.splice(r - 1, 1); }
  insertColumnsAfter(c, n) { this.maxCols += n; }
  appendRow(vals) { this.data.push([...vals]); }
  clearContents() { this.data = this.data.map(r => r.map(() => '')); return this; }
  clear() { return this.clearContents(); }
  setFrozenRows() {} setColumnWidth() {} setTabColor() {} hideSheet() {} activate() {}
  getSheetId() { return 1; }
}

class Spreadsheet {
  constructor() { this.sheets = new Map(); }
  getSheetByName(n) { return this.sheets.get(n) || null; }
  insertSheet(n) { const s = new Sheet(n, []); this.sheets.set(n, s); return s; }
  addSheet(n, data, maxCols) { const s = new Sheet(n, data, maxCols); this.sheets.set(n, s); return s; }
  getSheets() { return [...this.sheets.values()]; }
  getId() { return 'simulador'; }
  toast() {}
}

/**
 * Carrega Código.gs num contexto isolado com os serviços simulados.
 * Funções declaradas no .gs viram propriedades de `ctx`; constantes (const/let) só por vm.runInContext.
 * Janelas de confirmação (ui.alert) respondem "SIM".
 */
function criarContexto(opts = {}) {
  const ss = new Spreadsheet();
  const props = new Map();
  const logs = [];
  const ui = {
    alert: () => 'YES',
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
    Button: { YES: 'YES', NO: 'NO', OK: 'OK', CANCEL: 'CANCEL' },
    createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addSubMenu() { return this; }, addToUi() {} })
  };
  const ctx = {
    console,
    Logger: { log: (m) => { if (opts.verbose) console.log(String(m)); logs.push(String(m)); }, clear() {} },
    Utilities: {
      formatDate, getUuid: () => crypto.randomUUID(),
      DigestAlgorithm: { MD5: 'md5' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, txt) => [...crypto.createHash(alg).update(txt, 'utf8').digest()].map(b => b > 127 ? b - 256 : b),
      base64Encode: (bytes) => Buffer.from(bytes.map(b => (b + 256) % 256)).toString('base64'),
      sleep() {}
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (props.has(k) ? props.get(k) : null),
        setProperty: (k, v) => { props.set(k, String(v)); },
        deleteProperty: k => { props.delete(k); },
        getProperties: () => Object.fromEntries(props)
      })
    },
    CacheService: { getScriptCache: () => ({ get: () => null, put() {}, remove() {}, removeAll() {}, getAll: () => ({}), putAll() {} }) },
    LockService: {
      getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true }),
      getDocumentLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true })
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, openById: () => ss, flush() {}, getUi: () => ui },
    ScriptApp: {
      getProjectTriggers: () => [], deleteTrigger() {},
      newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create() {} }), everyHours: () => ({ create() {} }), after: () => ({ create() {} }) }) })
    },
    Session: { getActiveUser: () => ({ getEmail: () => '' }), getEffectiveUser: () => ({ getEmail: () => '' }) },
    HtmlService: {},
    Date, Math, JSON, Number, String, Array, Object, Set, Map, RegExp, Error, isNaN, isFinite, parseInt, parseFloat, Infinity, NaN, Symbol, Promise, Buffer
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(ARQ_CODIGO, 'utf8'), ctx, { filename: 'Código.gs' });
  return { ctx, ss, props, logs };
}

/** Célula de CSV → valor de getValues(): Date (dd/mm/aaaa), número ou texto (texto=true mantém string). */
function celula(v, texto) {
  if (v === '') return '';
  if (texto) return v;
  const m = /^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}):(\d{2}))?$/.exec(v);
  if (m) return dataFortaleza(+m[3], +m[2], +m[1], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(v)) return Number(v);
  return v;
}

function lerCsv(arq) {
  const txt = fs.readFileSync(arq, 'utf8');
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < txt.length; i++) {
    const ch = txt[i];
    if (q) { if (ch === '"') { if (txt[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === ',') { row.push(cur); cur = ''; }
    else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
    else if (ch !== '\r') cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function escreverCsv(arq, rows) {
  const q = v => { const s = String(v === null || v === undefined ? '' : v); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  fs.writeFileSync(arq, rows.map(r => r.map(q).join(',')).join('\n') + '\n', 'utf8');
}

module.exports = { criarContexto, lerCsv, escreverCsv, celula, dataFortaleza, formatDate };
