const xlsx = require('xlsx');
const XLSX = xlsx;
const mongoose = require('mongoose');
const puppeteer = require('puppeteer');

// Usa o mesmo modelo UploadedSheet registrado em server.js, mas só o resolve na hora da chamada,
// depois que o schema já foi definido.
function getUploadedSheetModel() {
  if (mongoose.models && mongoose.models.UploadedSheet) {
    return mongoose.models.UploadedSheet;
  }
  return mongoose.model('UploadedSheet');
}

// Helpers básicos que espelham a lógica do frontend (versões JS simplificadas)

function matchesSetorFilter(setorVal, filtro) {
  if (!filtro) return true;
  if (Array.isArray(filtro)) {
    if (!setorVal) return false;
    return filtro.includes(setorVal);
  }
  if (!setorVal) return false;
  return setorVal === filtro;
}

function formatSetorFiltroLabel(filtro) {
  if (!filtro) return 'Todos os setores';
  if (Array.isArray(filtro)) return filtro.join(', ');
  return filtro;
}

function parseSortableDate(s) {
  if (!s) return null;
  const trimmed = String(s).trim();
  const dmYTime = /^([0-3]?\d)\/(0?[1-9]|1[0-2])\/(\d{4})(?:\s+([0-1]?\d|2[0-3]):([0-5]?\d)(?::([0-5]?\d))?)?$/;
  const m = trimmed.match(dmYTime);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]) - 1;
    const y = Number(m[3]);
    const hh = m[4] ? Number(m[4]) : 0;
    const mm = m[5] ? Number(m[5]) : 0;
    const ss = m[6] ? Number(m[6]) : 0;
    return new Date(y, mo, d, hh, mm, ss).getTime();
  }
  const parsed = Date.parse(trimmed);
  if (!isNaN(parsed)) return parsed;
  const cleanedNum = Number(trimmed.replace(/[^0-9.-]/g, ''));
  if (!isNaN(cleanedNum)) {
    try {
      const ssf = xlsx.SSF;
      if (ssf && typeof ssf.parse_date_code === 'function') {
        const pc = ssf.parse_date_code(cleanedNum);
        if (pc && pc.y && pc.m && pc.d) {
          return new Date(pc.y, pc.m - 1, pc.d).getTime();
        }
      }
    } catch {
      // ignore
    }
  }
  return null;
}

// Abaixo, portamos a lógica de parse e cálculo de indicadores do frontend para o backend,
// para que o relatório consolidado tenha exatamente a mesma forma e números.

function parsePreventivaRows(rows, setorFiltroLocal = null, tipo, selectedMonthForMeta, preditivaOverride) {
  const out = [];
  const daysInMonth = (() => {
    if (selectedMonthForMeta) {
      const [yStr, mStr] = selectedMonthForMeta.split('-');
      const y = Number(yStr);
      const m = Number(mStr);
      if (!isNaN(y) && !isNaN(m) && m >= 1 && m <= 12) {
        return new Date(y, m, 0).getDate();
      }
    }
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  })();

  rows.forEach((row) => {
    if (!Array.isArray(row)) return;
    const colunaAF = row[31];
    const colunaF = row[5];
    const searchKeyword = tipo === 'preditivas' ? 'PREDITIVA' : 'PREVENTIVA';
    const temPreventiva = colunaAF !== undefined && colunaAF !== null &&
      String(colunaAF).toUpperCase().includes(searchKeyword);
    let temMaiorQue31 = false;
    if (colunaF !== undefined && colunaF !== null) {
      if (typeof colunaF === 'number') {
        temMaiorQue31 = colunaF > 31;
      } else {
        const cleaned = String(colunaF).replace(/[^0-9.-]/g, '').trim();
        const num = Number(cleaned);
        if (!isNaN(num)) {
          temMaiorQue31 = num > 31;
        } else {
          temMaiorQue31 = String(colunaF).toLowerCase().includes('maior que 31');
        }
      }
    }

    const isNivel1Match = temPreventiva && temMaiorQue31;
    const isNivel2Match = temPreventiva && !temMaiorQue31;

    let shouldInclude;
    if (typeof tipo === 'undefined') {
      shouldInclude = temPreventiva;
    } else if (tipo === 'preditivas') {
      shouldInclude = temPreventiva;
    } else if (tipo === 'preventiva_nivel_1') {
      shouldInclude = isNivel1Match;
    } else if (tipo === 'preventiva_nivel_2') {
      shouldInclude = isNivel2Match;
    } else {
      shouldInclude = temPreventiva;
    }

    if (!shouldInclude) return;

    try {
      const colunaAG = row[32];
      const effectivePreds = (typeof preditivaOverride !== 'undefined') ? preditivaOverride : [];
      if (tipo === 'preditivas' && effectivePreds && effectivePreds.length > 0) {
        const agRaw = colunaAG === undefined || colunaAG === null ? '' : String(colunaAG).trim();
        const agDigits = agRaw.replace(/[^0-9]/g, '');
        const match = effectivePreds.some(code => agDigits === String(code).replace(/[^0-9]/g, ''));
        if (!match) return;
      }
    } catch {
      // ignore
    }

    const parseExcelDate = (val) => {
      if (val === undefined || val === null) return null;
      const sffAccessor = XLSX && XLSX.SSF;
      if (typeof val === 'number') {
        try {
          const parsed = sffAccessor && sffAccessor.parse_date_code
            ? sffAccessor.parse_date_code(val)
            : null;
          if (parsed && parsed.y && parsed.m && parsed.d) {
            const d = String(parsed.d).padStart(2, '0');
            const m = String(parsed.m).padStart(2, '0');
            const y = parsed.y;
            return `${d}/${m}/${y}`;
          }
        } catch {
          // ignore
        }
        return String(val);
      }
      const cleaned = String(val).replace(/[^0-9.-]/g, '').trim();
      const num = Number(cleaned);
      if (!isNaN(num) && cleaned.length > 0) {
        try {
          const parsed = sffAccessor && sffAccessor.parse_date_code
            ? sffAccessor.parse_date_code(num)
            : null;
          if (parsed && parsed.y && parsed.m && parsed.d) {
            const d = String(parsed.d).padStart(2, '0');
            const m = String(parsed.m).padStart(2, '0');
            const y = parsed.y;
            return `${d}/${m}/${y}`;
          }
        } catch {
          // ignore
        }
      }
      return String(val);
    };

    const get = (idx, asDate = false) => {
      const v = row[idx];
      if (v === undefined || v === null) return null;
      return asDate ? parseExcelDate(v) : String(v);
    };

    const setorMap = {
      '5063': 'TRIO','5065': 'TRIO','5068': 'SEMI','5070': 'SEMI','5071': 'AREA EXTERNA','5072': 'SEMI','5073': 'SEMI','5074': 'SEMI','5077': 'SEMI','5080': 'TRIO','5082': 'TRIO','5083': 'SEMI','5089': 'AREA EXTERNA','5091': 'AREA EXTERNA','5092': 'AREA EXTERNA','5133': 'AREA EXTERNA','5135': 'AREA EXTERNA','7537': 'AREA EXTERNA','9022': 'SEMI','9316': 'AREA EXTERNA','9317': 'AREA EXTERNA','9318': 'AREA EXTERNA','9325': 'AREA EXTERNA','9326': 'AREA EXTERNA','9340': 'TRIO','9341': 'AREA EXTERNA','9342': 'TRIO','9343': 'TRIO','9344': 'TRIO','9346': 'TRIO','9352': 'AREA EXTERNA','9354': 'SEMI','9357': 'SEMI','9360': 'SEMI','9361': 'SEMI','9362': 'SEMI','9363': 'SEMI','9364': 'SEMI','9365': 'SEMI','9366': 'SEMI','9449': 'AREA EXTERNA','9596': 'AREA EXTERNA','11691': 'AREA EXTERNA','12367': 'AREA EXTERNA','13011': 'TRIO','13846': 'TRIO','13847': 'TRIO','13848': 'SEMI','13849': 'SEMI','13852': 'TRIO','14701': 'SEMI','14702': 'SEMI','14703': 'SEMI','14704': 'SEMI','14705': 'SEMI','14706': 'SEMI','14707': 'SEMI','14708': 'SEMI','14709': 'TRIO','14710': 'TRIO','14711': 'TRIO','14716': 'TRIO','14717': 'SEMI','14718': 'SEMI','14721': 'AREA EXTERNA','14725': 'TRIO','14967': 'SEMI','16977': 'AREA EXTERNA','6858': 'SEMI','5085': 'SEMI','9348': 'TRIO','6124': 'AREA EXTERNA'
    };

    const resolveSetor = (val) => {
      if (val === undefined || val === null) return null;
      const key = String(val).trim();
      const numericKey = key.replace(/[^0-9]/g, '');
      if (numericKey && setorMap[numericKey]) return setorMap[numericKey];
      if (setorMap[key]) return setorMap[key];
      return null;
    };

    const setorVal = resolveSetor(row[38]);
    if (!matchesSetorFilter(setorVal, setorFiltroLocal)) {
      return;
    }

    let metaIndividual = null;
    const freqVal = colunaF;
    if (freqVal !== undefined && freqVal !== null) {
      let freqNum = null;
      if (typeof freqVal === 'number') freqNum = freqVal;
      else {
        const cleanedF = String(freqVal).replace(/[^0-9.-]/g, '').trim();
        const numF = Number(cleanedF);
        if (!isNaN(numF) && cleanedF.length > 0) freqNum = numF;
      }
      if (freqNum && freqNum > 0) {
        metaIndividual = Math.round(daysInMonth / freqNum);
        if (metaIndividual < 1) metaIndividual = 1;
      }
    }

    out.push({
      ordemCodigo: get(9),
      codigoEquip: get(2),
      nomeEquip: get(3),
      c: get(2),
      d: get(3),
      f: get(5),
      j: get(9),
      reprogramada: get(10),
      n: get(13, true),
      p: get(15),
      r: get(17),
      ad: get(29),
      af: get(31),
      al: get(37),
      am: get(38),
      setor: setorVal,
      raw: row,
      meta: metaIndividual,
    });
  });
  return out;
}

function parseHigienizacaoRows(rows, setorFiltroLocal = null, selectedMonthForMeta) {
  const out = [];
  const daysInMonth = (() => {
    if (selectedMonthForMeta) {
      const [yStr, mStr] = selectedMonthForMeta.split('-');
      const y = Number(yStr);
      const m = Number(mStr);
      if (!isNaN(y) && !isNaN(m) && m >= 1 && m <= 12) {
        return new Date(y, m, 0).getDate();
      }
    }
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  })();

  rows.forEach((row) => {
    if (!Array.isArray(row)) return;
    const colunaAF = row[31];
    const colunaF = row[5];

    const temHigienizacao = colunaAF !== undefined && colunaAF !== null &&
      String(colunaAF).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().includes('HIGIENIZACAO');

    if (!temHigienizacao) return;

    const parseExcelDate = (val) => {
      if (val === undefined || val === null) return null;
      const sffAccessor = XLSX && XLSX.SSF;
      if (typeof val === 'number') {
        try {
          const parsed = sffAccessor && sffAccessor.parse_date_code
            ? sffAccessor.parse_date_code(val)
            : null;
          if (parsed && parsed.y && parsed.m && parsed.d) {
            const d = String(parsed.d).padStart(2, '0');
            const m = String(parsed.m).padStart(2, '0');
            const y = parsed.y;
            return `${d}/${m}/${y}`;
          }
        } catch {
          // ignore
        }
        return String(val);
      }
      const cleaned = String(val).replace(/[^0-9.-]/g, '').trim();
      const num = Number(cleaned);
      if (!isNaN(num) && cleaned.length > 0) {
        try {
          const parsed = sffAccessor && sffAccessor.parse_date_code
            ? sffAccessor.parse_date_code(num)
            : null;
          if (parsed && parsed.y && parsed.m && parsed.d) {
            const d = String(parsed.d).padStart(2, '0');
            const m = String(parsed.m).padStart(2, '0');
            const y = parsed.y;
            return `${d}/${m}/${y}`;
          }
        } catch {
          // ignore
        }
      }
      return String(val);
    };

    const get = (idx, asDate = false) => {
      const v = row[idx];
      if (v === undefined || v === null) return null;
      return asDate ? parseExcelDate(v) : String(v);
    };

    const setorMap = {
      '5063': 'TRIO','5065': 'TRIO','5068': 'SEMI','5070': 'SEMI','5071': 'AREA EXTERNA','5072': 'SEMI','5073': 'SEMI','5074': 'SEMI','5077': 'SEMI','5080': 'TRIO','5082': 'TRIO','5083': 'SEMI','5089': 'AREA EXTERNA','5091': 'AREA EXTERNA','5092': 'AREA EXTERNA','5133': 'AREA EXTERNA','5135': 'AREA EXTERNA','7537': 'AREA EXTERNA','9022': 'SEMI','9316': 'AREA EXTERNA','9317': 'AREA EXTERNA','9318': 'AREA EXTERNA','9325': 'AREA EXTERNA','9326': 'AREA EXTERNA','9340': 'TRIO','9341': 'AREA EXTERNA','9342': 'TRIO','9343': 'TRIO','9344': 'TRIO','9346': 'TRIO','9352': 'AREA EXTERNA','9354': 'SEMI','9357': 'SEMI','9360': 'SEMI','9361': 'SEMI','9362': 'SEMI','9363': 'SEMI','9364': 'SEMI','9365': 'SEMI','9366': 'SEMI','9449': 'AREA EXTERNA','9596': 'AREA EXTERNA','11691': 'AREA EXTERNA','12367': 'AREA EXTERNA','13011': 'TRIO','13846': 'TRIO','13847': 'TRIO','13848': 'SEMI','13849': 'SEMI','13852': 'TRIO','14701': 'SEMI','14702': 'SEMI','14703': 'SEMI','14704': 'SEMI','14705': 'SEMI','14706': 'SEMI','14707': 'SEMI','14708': 'SEMI','14709': 'TRIO','14710': 'TRIO','14711': 'TRIO','14716': 'TRIO','14717': 'SEMI','14718': 'SEMI','14721': 'AREA EXTERNA','14725': 'TRIO','14967': 'SEMI','16977': 'AREA EXTERNA','6858': 'SEMI','5085': 'SEMI','9348': 'TRIO','6124': 'AREA EXTERNA'
    };

    const resolveSetor = (val) => {
      if (val === undefined || val === null) return null;
      const key = String(val).trim();
      const numericKey = key.replace(/[^0-9]/g, '');
      if (numericKey && setorMap[numericKey]) return setorMap[numericKey];
      if (setorMap[key]) return setorMap[key];
      return null;
    };

    const setorVal = resolveSetor(row[38]);
    if (!matchesSetorFilter(setorVal, setorFiltroLocal)) {
      return;
    }

    let metaIndividual = null;
    const freqVal = colunaF;
    if (freqVal !== undefined && freqVal !== null) {
      let freqNum = null;
      if (typeof freqVal === 'number') freqNum = freqVal;
      else {
        const cleanedF = String(freqVal).replace(/[^0-9.-]/g, '').trim();
        const numF = Number(cleanedF);
        if (!isNaN(numF) && cleanedF.length > 0) freqNum = numF;
      }
      if (freqNum && freqNum > 0) {
        metaIndividual = Math.round(daysInMonth / freqNum);
        if (metaIndividual < 1) metaIndividual = 1;
      }
    }

    out.push({
      ordemCodigo: get(9),
      codigoEquip: get(2),
      nomeEquip: get(3),
      c: get(2),
      d: get(3),
      f: get(5),
      j: get(9),
      reprogramada: get(10),
      n: get(13, true),
      p: get(15),
      r: get(17),
      ad: get(29),
      af: get(31),
      al: get(37),
      am: get(38),
      setor: setorVal,
      raw: row,
      meta: metaIndividual,
    });
  });
  return out;
}

function parseLubrificacaoRows(rows, setorFiltroLocal = null, selectedMonthForMeta) {
  const out = [];
  const daysInMonth = (() => {
    if (selectedMonthForMeta) {
      const [yStr, mStr] = selectedMonthForMeta.split('-');
      const y = Number(yStr);
      const m = Number(mStr);
      if (!isNaN(y) && !isNaN(m) && m >= 1 && m <= 12) {
        return new Date(y, m, 0).getDate();
      }
    }
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  })();

  rows.forEach((row) => {
    if (!Array.isArray(row)) return;
    const colunaAF = row[31];
    const colunaF = row[5];

    const temLubrificacao = colunaAF !== undefined && colunaAF !== null &&
      String(colunaAF).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().includes('LUBRIFICACAO');

    if (!temLubrificacao) return;

    const parseExcelDate = (val) => {
      if (val === undefined || val === null) return null;
      const sffAccessor = XLSX && XLSX.SSF;
      if (typeof val === 'number') {
        try {
          const parsed = sffAccessor && sffAccessor.parse_date_code
            ? sffAccessor.parse_date_code(val)
            : null;
          if (parsed && parsed.y && parsed.m && parsed.d) {
            const d = String(parsed.d).padStart(2, '0');
            const m = String(parsed.m).padStart(2, '0');
            const y = parsed.y;
            return `${d}/${m}/${y}`;
          }
        } catch {
          // ignore
        }
        return String(val);
      }
      const cleaned = String(val).replace(/[^0-9.-]/g, '').trim();
      const num = Number(cleaned);
      if (!isNaN(num) && cleaned.length > 0) {
        try {
          const parsed = sffAccessor && sffAccessor.parse_date_code
            ? sffAccessor.parse_date_code(num)
            : null;
          if (parsed && parsed.y && parsed.m && parsed.d) {
            const d = String(parsed.d).padStart(2, '0');
            const m = String(parsed.m).padStart(2, '0');
            const y = parsed.y;
            return `${d}/${m}/${y}`;
          }
        } catch {
          // ignore
        }
      }
      return String(val);
    };

    const get = (idx, asDate = false) => {
      const v = row[idx];
      if (v === undefined || v === null) return null;
      return asDate ? parseExcelDate(v) : String(v);
    };

    const setorMap = {
      '5063': 'TRIO','5065': 'TRIO','5068': 'SEMI','5070': 'SEMI','5071': 'AREA EXTERNA','5072': 'SEMI','5073': 'SEMI','5074': 'SEMI','5077': 'SEMI','5080': 'TRIO','5082': 'TRIO','5083': 'SEMI','5089': 'AREA EXTERNA','5091': 'AREA EXTERNA','5092': 'AREA EXTERNA','5133': 'AREA EXTERNA','5135': 'AREA EXTERNA','7537': 'AREA EXTERNA','9022': 'SEMI','9316': 'AREA EXTERNA','9317': 'AREA EXTERNA','9318': 'AREA EXTERNA','9325': 'AREA EXTERNA','9326': 'AREA EXTERNA','9340': 'TRIO','9341': 'AREA EXTERNA','9342': 'TRIO','9343': 'TRIO','9344': 'TRIO','9346': 'TRIO','9352': 'AREA EXTERNA','9354': 'SEMI','9357': 'SEMI','9360': 'SEMI','9361': 'SEMI','9362': 'SEMI','9363': 'SEMI','9364': 'SEMI','9365': 'SEMI','9366': 'SEMI','9449': 'AREA EXTERNA','9596': 'AREA EXTERNA','11691': 'AREA EXTERNA','12367': 'AREA EXTERNA','13011': 'TRIO','13846': 'TRIO','13847': 'TRIO','13848': 'SEMI','13849': 'SEMI','13852': 'TRIO','14701': 'SEMI','14702': 'SEMI','14703': 'SEMI','14704': 'SEMI','14705': 'SEMI','14706': 'SEMI','14707': 'SEMI','14708': 'SEMI','14709': 'TRIO','14710': 'TRIO','14711': 'TRIO','14716': 'TRIO','14717': 'SEMI','14718': 'SEMI','14721': 'AREA EXTERNA','14725': 'TRIO','14967': 'SEMI','16977': 'AREA EXTERNA','6858': 'SEMI','5085': 'SEMI','9348': 'TRIO','6124': 'AREA EXTERNA'
    };

    const resolveSetor = (val) => {
      if (val === undefined || val === null) return null;
      const key = String(val).trim();
      const numericKey = key.replace(/[^0-9]/g, '');
      if (numericKey && setorMap[numericKey]) return setorMap[numericKey];
      if (setorMap[key]) return setorMap[key];
      return null;
    };

    const setorVal = resolveSetor(row[38]);
    if (!matchesSetorFilter(setorVal, setorFiltroLocal)) {
      return;
    }

    let metaIndividual = null;
    const freqVal = colunaF;
    if (freqVal !== undefined && freqVal !== null) {
      let freqNum = null;
      if (typeof freqVal === 'number') freqNum = freqVal;
      else {
        const cleanedF = String(freqVal).replace(/[^0-9.-]/g, '').trim();
        const numF = Number(cleanedF);
        if (!isNaN(numF) && cleanedF.length > 0) freqNum = numF;
      }
      if (freqNum && freqNum > 0) {
        metaIndividual = Math.round(daysInMonth / freqNum);
        if (metaIndividual < 1) metaIndividual = 1;
      }
    }

    out.push({
      ordemCodigo: get(9),
      codigoEquip: get(2),
      nomeEquip: get(3),
      c: get(2),
      d: get(3),
      f: get(5),
      j: get(9),
      reprogramada: get(10),
      n: get(13, true),
      p: get(15),
      r: get(17),
      ad: get(29),
      af: get(31),
      al: get(37),
      am: get(38),
      setor: setorVal,
      raw: row,
      meta: metaIndividual,
    });
  });
  return out;
}

function parseSolicitacoesRows(rows, setorFiltroLocal = null, selectedMonthForMeta) {
  const out = [];
  void selectedMonthForMeta;

  const parseExcelDate = (val) => {
    if (val === undefined || val === null) return null;
    const sffAccessor = XLSX && XLSX.SSF;
    if (typeof val === 'number') {
      try {
        const parsed = sffAccessor && sffAccessor.parse_date_code ? sffAccessor.parse_date_code(val) : null;
        if (parsed && parsed.y && parsed.m && parsed.d) {
          const d = String(parsed.d).padStart(2, '0');
          const m = String(parsed.m).padStart(2, '0');
          const y = parsed.y;
          const pd = parsed;
          const hh = pd['H'] ?? pd['h'] ?? 0;
          const mmn = pd['M'] ?? pd['m'] ?? 0;
          const ss = pd['S'] ?? pd['s'] ?? 0;
          const timePart = (hh || mmn || ss) ? ` ${String(hh).padStart(2,'0')}:${String(mmn).padStart(2,'0')}:${String(ss).padStart(2,'0')}` : '';
          return `${d}/${m}/${y}${timePart}`;
        }
      } catch {
        // ignore
      }
      try {
        const serial = Number(val);
        const jsTime = (serial - 25569) * 86400 * 1000;
        const jsDate = new Date(jsTime);
        if (!isNaN(jsDate.getTime())) {
          const dd = String(jsDate.getDate()).padStart(2, '0');
          const mm = String(jsDate.getMonth() + 1).padStart(2, '0');
          const yyyy = jsDate.getFullYear();
          const hh = String(jsDate.getHours()).padStart(2, '0');
          const mn = String(jsDate.getMinutes()).padStart(2, '0');
          const ssn = String(jsDate.getSeconds()).padStart(2, '0');
          return `${dd}/${mm}/${yyyy} ${hh}:${mn}:${ssn}`;
        }
      } catch {
        // ignore
      }
      return String(val);
    }
    const strVal = String(val).trim();
    const dmYwithTime = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
    const match = strVal.match(dmYwithTime);
    if (match) {
      const dd = String(match[1]).padStart(2, '0');
      const mm = String(match[2]).padStart(2, '0');
      const yy = match[3];
      const hh = match[4] ? String(match[4]).padStart(2, '0') : null;
      const mn = match[5] ? String(match[5]).padStart(2, '0') : null;
      const ss = match[6] ? String(match[6]).padStart(2, '0') : null;
      const timePart = hh ? ` ${hh}:${mn ?? '00'}:${ss ?? '00'}` : '';
      return `${dd}/${mm}/${yy}${timePart}`;
    }
    const cleaned = String(val).replace(/[^0-9.-]/g, '').trim();
    const num = Number(cleaned);
    if (!isNaN(num) && cleaned.length > 0) {
      try {
        const parsed = sffAccessor && sffAccessor.parse_date_code ? sffAccessor.parse_date_code(num) : null;
        if (parsed && parsed.y && parsed.m && parsed.d) {
          const d = String(parsed.d).padStart(2, '0');
          const m = String(parsed.m).padStart(2, '0');
          const y = parsed.y;
          return `${d}/${m}/${y}`;
        }
      } catch {
        // ignore
      }
    }
    return String(val);
  };

  const setorMap = {
    '5063': 'TRIO','5065': 'TRIO','5068': 'SEMI','5070': 'SEMI','5071': 'AREA EXTERNA','5072': 'SEMI','5073': 'SEMI','5074': 'SEMI','5077': 'SEMI','5080': 'TRIO','5082': 'TRIO','5083': 'SEMI','5089': 'AREA EXTERNA','5091': 'AREA EXTERNA','5092': 'AREA EXTERNA','5133': 'AREA EXTERNA','5135': 'AREA EXTERNA','7537': 'AREA EXTERNA','9022': 'SEMI','9316': 'AREA EXTERNA','9317': 'AREA EXTERNA','9318': 'AREA EXTERNA','9325': 'AREA EXTERNA','9326': 'AREA EXTERNA','9340': 'TRIO','9341': 'AREA EXTERNA','9342': 'TRIO','9343': 'TRIO','9344': 'TRIO','9346': 'TRIO','9352': 'AREA EXTERNA','9354': 'SEMI','9357': 'SEMI','9360': 'SEMI','9361': 'SEMI','9362': 'SEMI','9363': 'SEMI','9364': 'SEMI','9365': 'SEMI','9366': 'SEMI','9449': 'AREA EXTERNA','9596': 'AREA EXTERNA','11691': 'AREA EXTERNA','12367': 'AREA EXTERNA','13011': 'TRIO','13846': 'TRIO','13847': 'TRIO','13848': 'SEMI','13849': 'SEMI','13852': 'TRIO','14701': 'SEMI','14702': 'SEMI','14703': 'SEMI','14704': 'SEMI','14705': 'SEMI','14706': 'SEMI','14707': 'SEMI','14708': 'SEMI','14709': 'TRIO','14710': 'TRIO','14711': 'TRIO','14716': 'TRIO','14717': 'SEMI','14718': 'SEMI','14721': 'AREA EXTERNA','14725': 'TRIO','14967': 'SEMI','16977': 'AREA EXTERNA','6858': 'SEMI','5085': 'SEMI','9348': 'TRIO','6124': 'AREA EXTERNA'
  };

  const resolveSetor = (val) => {
    if (val === undefined || val === null) return null;
    const key = String(val).trim();
    const numericKey = key.replace(/[^0-9]/g, '');
    if (numericKey && setorMap[numericKey]) return setorMap[numericKey];
    if (setorMap[key]) return setorMap[key];
    return null;
  };

  rows.forEach((row) => {
    if (!Array.isArray(row)) return;

    const get = (idx, asDate = false) => {
      const v = row[idx];
      if (v === undefined || v === null) return null;
      return asDate ? parseExcelDate(v) : String(v);
    };

    const cleanCell = (v) => {
      if (v === null || v === undefined) return v;
      return String(v)
        .replace(/\u00A0/g, ' ')
        .replace(/\u200B/g, '')
        .replace(/\uFEFF/g, '')
        .replace(/[\t\r\n]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    };

    const setorVal = resolveSetor(row[26]);
    if (!matchesSetorFilter(setorVal, setorFiltroLocal)) return;

    const rawStatus = get(2);
    const raw16 = get(11);
    const solicitacaoVal = cleanCell(get(0));
    const statusVal = cleanCell(rawStatus) || '';
    const dez16Val = cleanCell(raw16) || '';
    const servicoVal = cleanCell(get(4)) || '';

    const servicoNormalized = servicoVal
      ? servicoVal.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim()
      : '';

    let hiddenForTable = false;
    if (
      servicoNormalized === 'COU - PROBLEMA ELETRICO' ||
      servicoNormalized === 'COU - PROBLEMA MECANICO'
    ) {
      hiddenForTable = true;
    }

    const removeNumericPrefixLocal = (x) => String(x).replace(/^\s*\d+\s*-?\s*/, '').trim();
    const statusForCheck = removeNumericPrefixLocal(statusVal.toLowerCase());
    const isPending = statusForCheck.includes('pend');

    const ordemDeServicoVal = (() => {
      const existing = cleanCell(get(10));
      if (isPending) {
        return existing && existing.length > 0 ? existing : 'GERAR O.S';
      }
      return existing;
    })();

    out.push({
      solicitacao: solicitacaoVal,
      statusSolicitacao: statusVal,
      prioridade: cleanCell(get(3)),
      servicoSolicitacao: cleanCell(get(4)),
      equipamentoSolicitacao: cleanCell(get(9)),
      ordemDeServico: ordemDeServicoVal,
      dezesseis15: dez16Val,
      dataServico: get(12, true),
      horasApropriadas: get(14, true),
      usuarioSolicitante: cleanCell(get(20)),
      setorSolicitacao: setorVal,
      ad: cleanCell(get(29)),
      af: cleanCell(get(31)),
      aa: cleanCell(get(26)),
      raw: row,
      statusExtra: null,
      highlightOrdemServico: isPending,
      hiddenFromTable: hiddenForTable,
    });
  });

  return out;
}

function computeComparisonResultsForTipo(tipo, ordensForTipo, completedOrdensAll, selectedMonthValue, setorFiltro, parseSortableDateFn) {
  if (!selectedMonthValue || !/^\d{4}-\d{2}$/.test(selectedMonthValue)) return null;
  const [yStr, mStr] = selectedMonthValue.split('-');
  const selY = Number(yStr);
  const selM = Number(mStr);
  const matched = [];

  const keyForOrder = (o) => {
    const key = (o.ordemCodigo || o.j || '').toString().trim();
    return key.replace(/\s+/g, ' ');
  };

  const completedInMonth = completedOrdensAll.filter(c => {
    const dateField = tipo === 'solicitacoes' ? (c.dataServico ?? '') : (c.n ?? '');
    const t = parseSortableDateFn(dateField);
    if (!t) return false;
    const d = new Date(t);
    const inMonth = d.getFullYear() === selY && (d.getMonth() + 1) === selM;
    if (!inMonth) return false;

    if (tipo === 'solicitacoes') {
      const rawLVal = c.dezesseis15 ?? (c.raw && Array.isArray(c.raw) ? c.raw[11] ?? '' : '');
      const normalizeL = (s) => String(s ?? '')
        .replace(/\u00A0/g, ' ')
        .replace(/\u200B/g, '')
        .replace(/\uFEFF/g, '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .toUpperCase().replace(/\s+/g, ' ').trim();
      const lNorm = normalizeL(rawLVal);
      const reFinalizada = /\b2\s*-?\s*FINALIZADA\b/;
      if (!reFinalizada.test(lNorm)) return false;
    }

    const isFreqGreaterThan31ForOrder = (ord) => {
      const rawVal = ord.raw && Array.isArray(ord.raw) ? ord.raw[5] : ord.f;
      const val = rawVal === undefined || rawVal === null ? ord.f : rawVal;
      if (val === undefined || val === null) return false;
      if (typeof val === 'number') return val > 31;
      const cleaned = String(val).replace(/[^0-9.-]/g, '').trim();
      const num = Number(cleaned);
      if (!isNaN(num) && cleaned.length > 0) return num > 31;
      const s = String(val).toLowerCase();
      return s.includes('maior que 31') || s.includes('> 31');
    };

    if (tipo === 'preditivas') {
      const afVal = (c.af !== undefined && c.af !== null) ? c.af : (c.raw && Array.isArray(c.raw) ? c.raw[31] : null);
      if (!afVal) return false;
      if (!String(afVal).toUpperCase().includes('PREDITIVA')) return false;
    } else if (tipo === 'preventiva_nivel_1') {
      if (!isFreqGreaterThan31ForOrder(c)) return false;
    } else if (tipo === 'preventiva_nivel_2') {
      if (isFreqGreaterThan31ForOrder(c)) return false;
    }

    if (setorFiltro) {
      const setorVal = (c.setor || '').toString() || null;
      return matchesSetorFilter(setorVal, setorFiltro);
    }
    return true;
  });

  const completedKeys = new Set(completedInMonth.map(c => keyForOrder(c)).filter(Boolean));

  ordensForTipo.forEach(p => {
    const keyP = keyForOrder(p);
    if (!keyP) return;
    if (completedKeys.has(keyP)) {
      const found = completedInMonth.find(c => keyForOrder(c) === keyP);
      if (found) matched.push({ pending: p, completed: found });
    }
  });

  let totalPendentes = ordensForTipo.length;

  if (tipo === 'solicitacoes') {
    const isPendingRow = (o) => {
      const raw = o.raw && Array.isArray(o.raw) ? o.raw : null;
      const statusField = String(o.statusSolicitacao ?? (raw ? raw[2] ?? '' : ''));
      const colLField = String(o.dezesseis15 ?? (raw ? raw[11] ?? '' : ''));

      const normalize = (s) => String(s ?? '')
        .replace(/\u00A0/g, ' ')
        .replace(/\u200B/g, '')
        .replace(/\uFEFF/g, '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .toUpperCase().replace(/\s+/g, ' ').trim();

      const status = normalize(statusField);
      const colL = normalize(colLField);

      const rePending = /^1\s*-?\s*PENDENTE\b/;
      const reCancelled = /\b3\s*-?\s*CANCELADA\b/;
      return rePending.test(status) || rePending.test(colL) || reCancelled.test(colL);
    };
    totalPendentes = ordensForTipo.filter(isPendingRow).length;
  }

  if (tipo === 'preventiva_nivel_1' || tipo === 'preventiva_nivel_2' || tipo === 'lubrificacao' || tipo === 'higienizacao' || tipo === 'preditivas' || tipo === 'corretiva' || tipo === 'predial' || tipo === 'melhoria' || tipo === 'outros') {
    totalPendentes = ordensForTipo.filter(o => {
      const dateField = o.n ?? '';
      const t = parseSortableDateFn(dateField);
      if (!t) return false;
      const d = new Date(t);
      return d.getFullYear() === selY && (d.getMonth() + 1) === selM;
    }).length;
  }

  let completedCount = completedInMonth.length;

  if (tipo === 'solicitacoes') {
    const reFinalizada = /\b2\s*-?\s*FINALIZADA\b/;
    completedCount = ordensForTipo.filter(o => {
      const rawL = o.dezesseis15 ?? (o.raw && Array.isArray(o.raw) ? o.raw[11] ?? '' : '');
      const lNorm = String(rawL ?? '')
        .replace(/\u00A0/g, ' ')
        .replace(/\u200B/g, '')
        .replace(/\uFEFF/g, '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .toUpperCase().replace(/\s+/g, ' ').trim();
      return reFinalizada.test(lNorm);
    }).length;
  }

  const matchedCount = matched.length;
  const pendentesCount = totalPendentes;

  let ordensPendentesParaGerarOS = 0;
  if (tipo === 'solicitacoes') {
    const rePendingOnly = /^1\s*-?\s*PENDENTE\b/;
    ordensPendentesParaGerarOS = ordensForTipo.filter(o => {
      const raw = o.raw && Array.isArray(o.raw) ? o.raw : null;
      const statusField = String(o.statusSolicitacao ?? (raw ? raw[2] ?? '' : ''));
      const statusNorm = String(statusField ?? '')
        .replace(/\u00A0/g, ' ')
        .replace(/\u200B/g, '')
        .replace(/\uFEFF/g, '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .toUpperCase().replace(/\s+/g, ' ').trim();
      return rePendingOnly.test(statusNorm);
    }).length;
  }

  let solicitacoesReprovadas = 0;
  if (tipo === 'solicitacoes') {
    const reReprovado = /\b3\s*-?\s*REPROVADO\b/;
    solicitacoesReprovadas = ordensForTipo.filter(o => {
      const raw = o.raw && Array.isArray(o.raw) ? o.raw : null;
      const statusField = String(o.statusSolicitacao ?? (raw ? raw[2] ?? '' : ''));
      const statusNorm = String(statusField ?? '')
        .replace(/\u00A0/g, ' ')
        .replace(/\u200B/g, '')
        .replace(/\uFEFF/g, '')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .toUpperCase().replace(/\s+/g, ' ').trim();
      return reReprovado.test(statusNorm);
    }).length;
  }

  let metaValue;
  if (tipo === 'preventiva_nivel_1' || tipo === 'preditivas') {
    metaValue = Math.max(0, Math.round((totalPendentes + completedCount) * 1.0));
  } else if (tipo === 'preventiva_nivel_2') {
    metaValue = Math.max(0, Math.round((totalPendentes + completedCount) * 0.85));
  } else if (tipo === 'solicitacoes') {
    metaValue = Math.max(0, Math.round((totalPendentes + completedCount) * 0.95));
  } else {
    metaValue = Math.max(0, Math.round((totalPendentes + completedCount) * 0.95));
  }

  const percentCompleted = metaValue > 0 ? (completedCount / metaValue) * 100 : null;
  const percentMatchedOfPendentes = totalPendentes > 0 ? (matchedCount / totalPendentes) * 100 : null;

  let reprogramadasCount = 0;
  if (tipo === 'corretiva' || tipo === 'predial' || tipo === 'melhoria' || tipo === 'outros') {
    reprogramadasCount = ordensForTipo.filter(o => {
      const raw = o.raw && Array.isArray(o.raw) ? o.raw : null;
      const repVal = String(o.reprogramada ?? (raw ? raw[10] ?? '' : '')).replace(/\s+/g, '').replace(',', '.');
      const num = Number(repVal.replace(/[^0-9.-]/g, ''));
      return !isNaN(num) && num >= 1;
    }).length;
  }

  return {
    totalPendentes,
    completedInMonth: completedCount,
    ordensPendentesParaGerarOS,
    solicitacoesReprovadas,
    reprogramadasCount,
    pendentes: pendentesCount,
    matched,
    percentCompleted,
    percentMatchedOfPendentes,
    meta: metaValue,
  };
}

function computeAdherenceMetricsForTipo(tipo, comparison) {
  if (!comparison) return null;

  const now = new Date();
  const year = now.getFullYear();
  const monthIndex = now.getMonth();

  const meta = comparison.meta ?? 0;
  const actualCompleted = comparison.completedInMonth ?? 0;

  if (tipo === 'solicitacoes') {
    const prevMonthIndex = (monthIndex + 11) % 12;
    const prevMonthYear = monthIndex === 0 ? year - 1 : year;

    const periodStart = new Date(prevMonthYear, prevMonthIndex, 16);
    const periodEnd = new Date(year, monthIndex, 15, 23, 59, 59, 999);

    const msInDay = 24 * 60 * 60 * 1000;
    const totalDays = Math.max(1, Math.floor((periodEnd.getTime() - periodStart.getTime()) / msInDay) + 1);

    let daysElapsed = 0;
    if (now >= periodStart) {
      if (now > periodEnd) {
        daysElapsed = totalDays;
      } else {
        daysElapsed = Math.floor((now.getTime() - periodStart.getTime()) / msInDay) + 1;
      }
    }

    const expectedToDate = totalDays > 0 ? (meta / totalDays) * daysElapsed : 0;
    const adherencePercent = expectedToDate > 0 ? (actualCompleted / expectedToDate) * 100 : null;
    return { daysInMonth: totalDays, today: daysElapsed, expectedToDate, actualCompleted, adherencePercent };
  }

  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const today = now.getDate();
  const expectedToDate = daysInMonth > 0 ? (meta / daysInMonth) * today : 0;
  const adherencePercent = expectedToDate > 0 ? (actualCompleted / expectedToDate) * 100 : null;
  return { daysInMonth, today, expectedToDate, actualCompleted, adherencePercent };
}

async function buildConsolidatedHtmlFromDb(params = {}) {
  const { selectedMonth = null, setorFiltro = null } = params;

  const UploadedSheet = getUploadedSheetModel();
  const pendentesDoc = await UploadedSheet.findOne({ type: 'pendentes' }).sort({ uploadedAt: -1 }).lean();
  if (!pendentesDoc) return null;
  const completedDoc = await UploadedSheet.findOne({ type: 'completed' }).sort({ uploadedAt: -1 }).lean();
  const solicitacoesDoc = await UploadedSheet.findOne({ type: 'solicitacoes' }).sort({ uploadedAt: -1 }).lean();

  const dadosOriginais = (pendentesDoc.parsedRows || []).slice(1); // remove cabeçalho
  const completedRows = completedDoc ? (completedDoc.parsedRows || []).slice(1) : [];
  const solicitacoesRows = solicitacoesDoc ? (solicitacoesDoc.parsedRows || []).slice(1) : [];

  const completedPreventiva = parsePreventivaRows(completedRows);
  const completedPreditiva = parsePreventivaRows(completedRows, undefined, 'preditivas');
  const completedLub = parseLubrificacaoRows(completedRows);
  const completedHig = parseHigienizacaoRows(completedRows);
  const completedSolic = parseSolicitacoesRows(completedRows);
  const completedOrdensAll = [...completedPreventiva, ...completedPreditiva, ...completedLub, ...completedHig, ...completedSolic];

  const now = new Date();
  const dataStr = now.toLocaleString('pt-BR');
  const mesStr = selectedMonth
    ? `${selectedMonth.slice(5)}/${selectedMonth.slice(0, 4)}`
    : 'Todos os meses';
  const setorStr = formatSetorFiltroLabel(setorFiltro || undefined);

  const escapeHtml = (value) => {
    const s = String(value ?? '');
    return s
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  };

  const tiposRelatorio = [
    'preventiva_nivel_1',
    'preventiva_nivel_2',
    'lubrificacao',
    'preditivas',
    'solicitacoes',
  ];

  const tiposOrdemLabels = {
    preventiva_nivel_1: 'Preventiva Nível 1',
    preventiva_nivel_2: 'Preventiva Nível 2',
    lubrificacao: 'Lubrificação',
    higienizacao: 'Higienização',
    preditivas: 'Preditivas',
    solicitacoes: 'Solicitações',
    corretiva: 'Corretiva',
    predial: 'Predial',
    melhoria: 'Melhoria',
    outros: 'Outros',
  };

  const secoes = [];

  const getOrdensForTipoForMetrics = (tipo) => {
    if (tipo === 'solicitacoes') {
      const source = (solicitacoesRows && solicitacoesRows.length > 0)
        ? solicitacoesRows
        : dadosOriginais;
      if (!source || source.length === 0) return [];
      return parseSolicitacoesRows(source, setorFiltro, selectedMonth);
    }

    if (!dadosOriginais || dadosOriginais.length === 0) return [];

    switch (tipo) {
      case 'lubrificacao':
        return parseLubrificacaoRows(dadosOriginais, setorFiltro, selectedMonth);
      case 'higienizacao':
        return parseHigienizacaoRows(dadosOriginais, setorFiltro, selectedMonth);
      case 'corretiva':
        // não usado no relatório consolidado atual
        return [];
      case 'predial':
        return [];
      case 'melhoria':
        return [];
      case 'outros':
        return [];
      case 'preventiva_nivel_1':
      case 'preventiva_nivel_2':
      case 'preditivas':
      default:
        return parsePreventivaRows(
          dadosOriginais,
          setorFiltro,
          tipo,
          selectedMonth,
          tipo === 'preditivas' ? [] : undefined,
        );
    }
  };

  for (const tipo of tiposRelatorio) {
    const ordensTipo = getOrdensForTipoForMetrics(tipo);
    if (!ordensTipo || ordensTipo.length === 0) continue;

    const comp = computeComparisonResultsForTipo(
      tipo,
      ordensTipo,
      completedOrdensAll,
      selectedMonth,
      setorFiltro,
      parseSortableDate,
    );
    if (!comp) continue;

    const adh = computeAdherenceMetricsForTipo(tipo, comp);
    const label = tiposOrdemLabels[tipo] || tipo;
    secoes.push({ tipo, label, comparison: comp, adherence: adh });
  }

  if (secoes.length === 0) return null;

  let secoesHtml = '';
  let solicitacoesExtrasHtml = '';

  for (const sec of secoes) {
    const { tipo, label, comparison, adherence } = sec;
    const pendentes = comparison.pendentes ?? 0;
    const realizadasMes = comparison.completedInMonth ?? 0;
    const meta = comparison.meta ?? 0;
    const percentualMes = comparison.percentCompleted !== null && comparison.percentCompleted !== undefined
      ? comparison.percentCompleted.toFixed(1)
      : '';
    const esperadoAteHoje = adherence
      ? Math.round(adherence.expectedToDate)
      : '';
    const aderencia = adherence && adherence.adherencePercent !== null
      ? adherence.adherencePercent.toFixed(1)
      : '';

    const ordensIndicador = tipo === 'solicitacoes'
      ? (comparison.meta ?? 0) - (comparison.completedInMonth ?? 0)
      : '';
    const pendentesGerarOS = tipo === 'solicitacoes'
      ? (comparison.ordensPendentesParaGerarOS ?? 0)
      : '';
    const reprovadas = tipo === 'solicitacoes'
      ? (comparison.solicitacoesReprovadas ?? 0)
      : '';

    secoesHtml += `
      <h2>${escapeHtml(label)}</h2>
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-title">Pendentes</div>
          <div class="kpi-value">${pendentes}</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Realizadas no mês</div>
          <div class="kpi-value">${realizadasMes}</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Meta</div>
          <div class="kpi-value">${meta}</div>
          <div class="kpi-sub">${percentualMes ? escapeHtml(percentualMes + '% do mês') : ''}</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-title">Aderência ao plano</div>
          <div class="kpi-value">${aderencia ? escapeHtml(aderencia + '%') : '--'}</div>
          <div class="kpi-sub">${esperadoAteHoje !== '' ? 'Esperado até hoje: ' + escapeHtml(String(esperadoAteHoje)) : ''}</div>
        </div>
      </div>
    `;

    if (tipo === 'solicitacoes') {
      solicitacoesExtrasHtml = `
      <h3>Indicadores específicos de Solicitações</h3>
      <table>
        <thead>
          <tr>
            <th>Ordens para indicador</th>
            <th>Pendentes para gerar O.S</th>
            <th>Solicitações reprovadas</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>${escapeHtml(ordensIndicador === '' ? '-' : String(ordensIndicador))}</td>
            <td>${escapeHtml(pendentesGerarOS === '' ? '-' : String(pendentesGerarOS))}</td>
            <td>${escapeHtml(reprovadas === '' ? '-' : String(reprovadas))}</td>
          </tr>
        </tbody>
      </table>
      `;
    }
  }

  const html = `<!DOCTYPE html>
  <html lang="pt-BR">
    <head>
      <meta charSet="utf-8" />
      <title>Relatório de Indicadores - Processa Plano</title>
      <style>
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; margin: 24px; color: #111827; }
        h1 { font-size: 22px; margin-bottom: 4px; }
        h2 { font-size: 16px; margin-top: 20px; margin-bottom: 8px; }
        h3 { font-size: 14px; margin-top: 14px; margin-bottom: 6px; }
        p { margin: 2px 0; font-size: 12px; }
        .meta { font-size: 12px; color: #4b5563; margin-bottom: 16px; }
        .kpi-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 8px; margin-top: 8px; }
        .kpi-card { border-radius: 8px; padding: 8px 10px; border: 1px solid #e5e7eb; }
        .kpi-title { font-size: 11px; text-transform: uppercase; letter-spacing: .04em; color: #6b7280; margin-bottom: 4px; }
        .kpi-value { font-size: 16px; font-weight: 600; }
        .kpi-sub { font-size: 11px; color: #6b7280; margin-top: 2px; }
        table { border-collapse: collapse; width: 100%; margin-top: 8px; font-size: 11px; }
        th, td { border: 1px solid #d1d5db; padding: 4px 6px; text-align: left; }
        th { background: #f3f4f6; font-weight: 600; }
        ul { margin: 4px 0 0 18px; padding: 0; font-size: 11px; }
        li { margin: 2px 0; }
        @media print {
          body { margin: 10mm; }
          .no-print { display: none; }
        }
      </style>
    </head>
    <body>
      <h1>Relatório de Indicadores - Processa Plano (Consolidado)</h1>
      <div class="meta">
        <p><strong>Mês de referência:</strong> ${escapeHtml(mesStr)}</p>
        <p><strong>Setor:</strong> ${escapeHtml(setorStr)}</p>
        <p><strong>Gerado em:</strong> ${escapeHtml(dataStr)}</p>
        <p><strong>Guias incluídas:</strong> ${escapeHtml(secoes.map(s => s.label).join(', '))}</p>
      </div>
      ${secoesHtml}
      ${solicitacoesExtrasHtml}

      <h2>Notas para análise</h2>
      <ul>
        <li>Use este relatório como base para apresentações e reuniões de acompanhamento.</li>
        <li>Os números refletem a combinação entre pendentes e ordens concluídas de acordo com o filtro de mês e setor.</li>
        <li>Para detalhes linha a linha, utilize a própria tela do Processa Plano ou exporte via copiar/colar para o Excel.</li>
      </ul>

      <p class="no-print" style="margin-top:16px; font-size:11px; color:#6b7280;">Dica: utilize a opção "Imprimir" do navegador e selecione "Salvar como PDF" para gerar o arquivo.</p>
    </body>
  </html>`;

  return html;
}

async function buildConsolidatedPdfBufferFromDb(params = {}) {
  const html = await buildConsolidatedHtmlFromDb(params);
  if (!html) return null;

  // Usa Chromium headless (puppeteer) para renderizar o mesmo HTML do relatório e gerar um PDF
  // com o mesmo padrão visual do navegador.

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '10mm', right: '10mm', bottom: '10mm', left: '10mm' },
    });
    return pdfBuffer;
  } finally {
    await browser.close();
  }
}

module.exports = {
  buildConsolidatedHtmlFromDb,
  buildConsolidatedPdfBufferFromDb,
};
