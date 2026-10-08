/**
 * Tablero Negocios Financieros – El Puerto de Liverpool
 * Backend de Apps Script.
 */

const CONFIG = {
  SPREADSHEET_ID: '',                        // vacío = usa el Sheet al que está ligado el proyecto
  HOJA_TOTAL: ['E_Total'],
  HOJA_CTAS: ['E_CtasNvas', 'E_CtaNvas'],    // acepta cualquiera de los dos nombres
  ORGS: [
    { key: 'visa',     org: '200' },
    { key: 'dept',     org: '210' },
    { key: 'mini',     org: '230' },
    { key: 'suburbia', org: 'TOTAL' }        // SECCIÓN SUBURBIA AL FINAL (Suma/Consolidado de 200 + 210 + 230)
  ],
  ORGS_A_SUMAR: ['200', '210', '230'],       // Organizaciones que componen el Total
  FORZAR_PERIODO: ''                         // solo para pruebas, ej. '2026-10'. Vacío = mes en curso
};

const MESES_ES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

// Cada métrica se define con columnas: [numerador] o [numerador, denominador] (para razones).
const TABS = {
  colocacion: {
    title: 'Cuentas Nuevas', metricName: 'Cuentas', src: 'ctas',
    actual: ['CUENTAS'], ppto: ['PPTO_REV'], me: ['ME'],
    vs: 'ppto', fmt: 'int', diff: 'abs', showScorecards: true
  },
  aprobacion: {
    title: 'Aprobación', metricName: 'Aprobación', src: 'total',
    actual: ['APROBADAS', 'SOLICITUDES'], ppto: ['CTA_PPTO_REV', 'SOL_PPTO_REV'], me: null,
    vs: 'ppto', fmt: 'pct', diff: 'rel', showScorecards: false
  },
  solicitudes: {
    title: 'Solicitudes', metricName: 'Solicitudes', src: 'total',
    actual: ['SOLICITUDES'], ppto: ['SOL_PPTO_REV'], me: ['SOL_ME'],
    vs: 'ppto', fmt: 'int', diff: 'abs', showScorecards: true
  },
  lineas: {
    title: 'Líneas', metricName: null, src: 'ctas',
    actual: ['LINEAS', 'CUENTAS'], ppto: null, me: null,   // línea promedio = LINEAS / CUENTAS, contra el año anterior
    vs: 'aa', fmt: 'int', diff: 'rel', showScorecards: false
  }
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Tablero Negocios Financieros')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Llamada desde el front con google.script.run. Lee las hojas en cada ejecución. */
function getDashboardData() {
  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const maps = {
    total: loadMap_(ss, CONFIG.HOJA_TOTAL, colsOf_('total'), tz),
    ctas: loadMap_(ss, CONFIG.HOJA_CTAS, colsOf_('ctas'), tz)
  };

  const p = (CONFIG.FORZAR_PERIODO || Utilities.formatDate(new Date(), tz, 'yyyy-M')).split('-').map(Number);
  const y = p[0], m = p[1];

  const out = { periodo: MESES_ES[m - 1] + ' ' + y, year: y, month: m, tabs: {} };
  Object.keys(TABS).forEach(k => {
    out.tabs[k] = buildTab_(TABS[k], maps[TABS[k].src], y, m);
  });
  return out;
}

// ---------- Construcción de cada pestaña ----------

function buildTab_(cfg, map, y, m) {
  const aa = cfg.vs === 'aa';
  const refSpec = aa ? cfg.actual : cfg.ppto;
  const refY = aa ? y - 1 : y;
  const pct = cfg.fmt === 'pct';
  const months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

  const show = v => v === null ? '–' : (pct ? fmtPct_(v) : fmtInt_(v));
  const chartV = v => v === null ? null : (pct ? Math.round(v * 10000) / 100 : Math.round(v));

  const kpi = (orgKey, m1, m2) => {
    const act = measure_(map, orgKey, y, m1, m2, cfg.actual);
    const ref = measure_(map, orgKey, refY, m1, m2, refSpec);
    const d = { ppto: show(ref), actual: show(act), diffNum: '–', diffPct: cfg.diff === 'abs' ? '–' : '', isPositive: true, neutral: true };
    if (act !== null && ref) {
      const diff = act - ref;
      d.neutral = false;
      d.isPositive = diff >= 0;
      if (cfg.diff === 'abs') { d.diffNum = fmtInt_(diff); d.diffPct = fmtPct_(diff / ref); }
      else { d.diffNum = fmtPct_((act / ref) - 1); d.diffPct = ''; }
    }
    return d;
  };

  const rows = [];
  CONFIG.ORGS.forEach(o => {
    const orgKey = String(o.org).trim();
    if (!orgKey) return;

    const actualData = months.map(mm => mm <= m ? chartV(measure_(map, orgKey, y, mm, mm, cfg.actual)) : null);
    const pptoData = months.map(mm => chartV(measure_(map, orgKey, refY, mm, mm, refSpec)));
    const meData = cfg.me ? months.map(mm => chartV(measure_(map, orgKey, y, mm, mm, cfg.me))) : null;

    // Totales para scorecards
    const meTotal = cfg.showScorecards && meData ? sumArr_(meData) : null;
    const pptoTotal = cfg.showScorecards && pptoData ? sumArr_(pptoData) : null;

    const chart = {
      yearLabel: String(y),
      isPct: pct,
      actual: actualData,
      pptoRev: pptoData,
      meTotal: meTotal !== null ? fmtInt_(meTotal) : null,
      pptoTotal: pptoTotal !== null ? fmtInt_(pptoTotal) : null
    };

    if (aa) {
      chart.lineLabel = String(y - 1);
    } else {
      chart.prevLabel = String(y - 1);
      chart.prev = months.map(mm => chartV(measure_(map, orgKey, y - 1, mm, mm, cfg.actual)));
    }
    if (cfg.me) chart.me = meData;

    rows.push({ key: o.key, name: o.key === 'suburbia' ? 'Suburbia' : o.key, acumulado: kpi(orgKey, 1, m), mes: kpi(orgKey, m, m), chartData: chart });
  });

  return {
    title: cfg.title,
    metricName: aa ? String(y) : cfg.metricName,
    showSeparateDiffHeaders: cfg.diff === 'abs',
    showScorecards: cfg.showScorecards,
    kpiLabels: aa
      ? { main: String(y), ppto: String(y - 1), diff: 'Diferencia vs AA' }
      : { main: cfg.metricName, ppto: 'Ppto Rev', diff: 'Diferencia vs Ppto Rev' },
    rows: rows
  };
}

// ---------- Cálculo ----------

function sumArr_(arr) {
  if (!arr) return null;
  let total = 0, count = 0;
  arr.forEach(v => { if (v !== null && !isNaN(v)) { total += v; count++; } });
  return count > 0 ? total : null;
}

function sumRange_(map, orgKey, y, m1, m2, col) {
  let t = null;
  const targetOrgs = (orgKey === 'TOTAL') ? CONFIG.ORGS_A_SUMAR : [orgKey];

  for (let mm = m1; mm <= m2; mm++) {
    targetOrgs.forEach(o => {
      const r = map[o + '|' + y + '|' + mm];
      if (r && r[col] !== undefined) t = (t || 0) + r[col];
    });
  }
  return t;
}

// spec = [num] o [num, den]. Para razones suma numerador y denominador del periodo y luego divide.
function measure_(map, orgKey, y, m1, m2, spec) {
  const n = sumRange_(map, orgKey, y, m1, m2, spec[0]);
  if (spec.length === 1) return n;
  const d = sumRange_(map, orgKey, y, m1, m2, spec[1]);
  return (n === null || !d) ? null : n / d;
}

// ---------- Lectura de hojas ----------

function colsOf_(src) {
  const s = {};
  Object.keys(TABS).map(k => TABS[k]).filter(t => t.src === src).forEach(t => {
    [t.actual, t.ppto, t.me].forEach(sp => { if (sp) sp.forEach(c => { s[c] = 1; }); });
  });
  return Object.keys(s);
}

// Devuelve { 'ORG|año|mes': { COLUMNA: suma } } leyendo la hoja en este momento.
function loadMap_(ss, nombres, cols, tz) {
  let sh = null;
  for (let i = 0; i < nombres.length && !sh; i++) sh = ss.getSheetByName(nombres[i]);
  if (!sh) throw new Error('No encontré la hoja "' + nombres[0] + '".');

  const vals = sh.getDataRange().getValues();
  const idx = {};
  vals[0].forEach((h, i) => { const k = String(h).trim().toUpperCase(); if (k && !(k in idx)) idx[k] = i; });
  ['MES', 'ORG'].concat(cols).forEach(c => {
    if (!(c in idx)) throw new Error('Falta la columna "' + c + '" en la hoja "' + sh.getName() + '".');
  });

  const map = {};
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const ym = ym_(row[idx.MES], tz);
    if (!ym) continue;
    const key = String(row[idx.ORG]).trim() + '|' + ym[0] + '|' + ym[1];
    const o = map[key] || (map[key] = {});
    cols.forEach(c => {
      const n = num_(row[idx[c]]);
      if (n !== null) o[c] = (o[c] || 0) + n;
    });
  }
  return map;
}

// Devuelve [año, mes] de la columna MES (Date o texto m/d/aaaa).
function ym_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-M').split('-').map(Number);
  if (typeof v === 'string' && v.trim()) {
    const p = v.trim().split(/[\/\-\s]/).map(Number);
    if (p.length >= 3 && p.every(isFinite)) return p[0] > 31 ? [p[0], p[1]] : [p[2], p[0]];
  }
  return null;
}

function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v.replace(/[,$%\s]/g, ''));
    return isFinite(n) ? n : null;
  }
  return null;
}

function fmtInt_(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function fmtPct_(f) { return (f * 100).toFixed(1) + '%'; }
