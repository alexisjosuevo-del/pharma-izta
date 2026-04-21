
const STORAGE_KEY = "mezclas_homologadas_uploads_v1";
const STORAGE_PACKAGING = "mezclas_packaging_catalog_v1";
const STORAGE_PRICES = "mezclas_precios_v1";

const monthOrder = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];

const state = {
  baseData: [],
  uploadedData: [],
  packagingCatalog: [],
  priceList: [],
  filteredData: [],
  monthlyChart: null,
  labChart: null,
  hospChart: null,
  marketShareChart: null,
  forecastChart: null,
  forecastHistoryChart: null,
  forecastModelChart: null,
  selectedForecastGroup: null,
};

const CHART_OPTIONS = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: { labels: { color: '#94a3b8', font: { family: 'Outfit' } } }
  },
  scales: {
    x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', font: { family: 'Outfit' } } },
    y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#94a3b8', font: { family: 'Outfit' } } }
  }
};

const _cache_normalizeText = new Map();
function normalizeText(value){
  if(value === null || value === undefined) return "";
  const s = String(value);
  if(_cache_normalizeText.has(s)) return _cache_normalizeText.get(s);
  
  const res = s
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[.,;:]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  _cache_normalizeText.set(s, res);
  return res;
}

const _cache_normalizeTextKP = new Map();
function normalizeTextKeepPunct(value){
  if(value === null || value === undefined) return "";
  const s = String(value);
  if(_cache_normalizeTextKP.has(s)) return _cache_normalizeTextKP.get(s);

  const res = s
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  _cache_normalizeTextKP.set(s, res);
  return res;
}

const _cache_normalizeHospital = new Map();
function normalizeHospital(value){
  const s = String(value || "");
  if(_cache_normalizeHospital.has(s)) return _cache_normalizeHospital.get(s);
  const res = normalizeText(value);
  _cache_normalizeHospital.set(s, res);
  return res;
}

const _cache_normalizeFabricante = new Map();
function normalizeFabricante(value){
  const s = String(value || "");
  if(_cache_normalizeFabricante.has(s)) return _cache_normalizeFabricante.get(s);

  let txt = normalizeText(value)
    .replace(/[\-_]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  txt = txt
    .replace(/^LABORATORIOS?\s+/g, "")
    .replace(/^LABS?\s+/g, "")
    .replace(/^S\s*A\s*DE\s*C\s*V\b/g, "")
    .replace(/^SA\s+DE\s+CV\b/g, "")
    .replace(/^S\s*A\b/g, "")
    .replace(/^DE\s+C\s*V\b/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if(txt === "ZURICH PHARMA" || txt === "ZURICH") txt = "ZURICH";
  if(txt.includes("FRENESUS") || txt.includes("FRESENIUS")) txt = "FRESENIUS KABI";
  if(txt.includes("ACCORD")) txt = "ACCORD";
  if(txt.includes("JANSSEN")) txt = "JANSSEN";

  _cache_normalizeFabricante.set(s, txt);
  return txt;
}



function normalizeMedicamentoOnco(value){
  let txt = normalizeText(value)
    .replace(/\b\d+(?:[.,]\d+)?\s*(MG\/ML|MG|MCG|G|GRAMOS?|GR|ML|UI|IU|MEQ|MMOL|%|MUI|U|UNIDADES?)\b/g, " ")
    .replace(/\b\d+(?:[.,]\d+)?(?:\/\d+(?:[.,]\d+)?)?\b/g, " ")
    .replace(/\bSOLUCION\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if(/^TRASTUZUMAB EMTANSINA\b/.test(txt)) return "TRASTUZUMAB";
  if(/^TRASTUZUMAB DERUXTECAN?\b/.test(txt)) return "TRASTUZUMAB";
  if(/^DOXORUBICINA LIPOSOMAL PEGILADA\b/.test(txt)) return "DOXORUBICINA";
  if(/^BRENTUXIMAB(?:\s+FCO\s+AMP)?\b/.test(txt)) return "BRENTUXIMAB";
  if(/^MIRVETUXIMAB SORAVTANSINE\b/.test(txt)) return "MIRVETUXIMAB";
  if(/^ENFORTUMAB VEDOTINA\b/.test(txt)) return "ENFORTUMAB";
  return txt;
}

function normalizeMedicamentoNutri(value){
  return normalizeTextKeepPunct(value);
}

function parseNumber(value){
  if(value === null || value === undefined || value === "") return null;
  if(typeof value === "number") return Number.isFinite(value) ? value : null;
  const normalized = String(value).replace(/,/g, ".").match(/-?\d+(?:\.\d+)?/);
  return normalized ? Number(normalized[0]) : null;
}

function extractFirstNumber(value){
  return parseNumber(value);
}

function extractPresentationValue(value){
  if(value === null || value === undefined || value === "") return null;
  if(typeof value === "number") return Number.isFinite(value) ? value : null;
  const txt = String(value).trim();
  if(!txt) return null;

  const normalized = txt.replace(/,/g, ".");
  let match = normalized.match(/\/\s*(-?\d+(?:\.\d+)?)\s*(?:ml|mL|ML|l|L|meq|MEQ|ui|UI|iu|IU|u|U|mg|MG|g|G)(?:\s|$)/);
  if(match) return Number(match[1]);

  match = normalized.match(/(-?\d+(?:\.\d+)?)\s*(?:ml|mL|ML|l|L)(?:\s|$)/g);
  if(match && match.length){
    const last = match[match.length - 1].match(/-?\d+(?:\.\d+)?/);
    if(last) return Number(last[0]);
  }

  const nums = [...normalized.matchAll(/-?\d+(?:\.\d+)?/g)].map(m => Number(m[0])).filter(Number.isFinite);
  return nums.length ? nums[nums.length - 1] : null;
}

function computeFrascosFromRecord(record){
  if(record && record.frascos !== null && record.frascos !== undefined && record.frascos !== "") {
    const explicit = Number(record.frascos);
    if(Number.isFinite(explicit)) return explicit;
  }
  const dosis = parseNumber(record?.dosis);
  const presentacion = parseNumber(record?.presentacion_valor ?? record?.presentacion_original);
  if(dosis !== null && presentacion !== null && Number.isFinite(dosis) && Number.isFinite(presentacion) && presentacion > 0){
    return dosis / presentacion;
  }
  return null;
}

function computePiezas(record, piezasPorEmpaque = 1){
  const frascos = computeFrascosFromRecord(record);
  const empaque = Number(piezasPorEmpaque) || 1;
  if(frascos === null || !Number.isFinite(frascos) || empaque <= 0) return null;
  return Number((frascos / empaque).toFixed(4));
}

function formatLabelNumber(value){
  const num = parseNumber(value);
  if(num === null || !Number.isFinite(num)) return "";
  // Round to 2 decimals if it has decimals, else integer
  if(Math.abs(num - Math.round(num)) < 1e-9) return String(Math.round(num));
  return String(Number(num.toFixed(2)));
}

function buildMedicamentoHomologadoPresentacion(baseName, presentacionOriginal, presentacionValor){
  const num = parseNumber(presentacionValor ?? presentacionOriginal);
  if(num === null || !Number.isFinite(num)) return baseName || "";
  return [baseName || "", formatLabelNumber(num)].filter(Boolean).join(" ").trim();
}

function excelDateToISO(serial){
  if(typeof serial !== "number" || !Number.isFinite(serial)) return null;
  // Corrección para fechas de Excel que a veces vienen con decimales (hora)
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  // Ajuste de zona horaria para evitar desfases de un día al usar ISO
  const year = dateInfo.getUTCFullYear();
  const month = String(dateInfo.getUTCMonth() + 1).padStart(2, "0");
  const day = String(dateInfo.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseDate(value){
  if(value === null || value === undefined || value === "") return null;
  if(value instanceof Date && !isNaN(value)) return value.toISOString().slice(0,10);
  if(typeof value === "number") return excelDateToISO(value);
  const txt = String(value).trim();
  if(!txt) return null;

  let match = txt.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if(match){
    return `${match[1]}-${String(match[2]).padStart(2,"0")}-${String(match[3]).padStart(2,"0")}`;
  }

  match = txt.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if(match){
    const year = Number(match[3]) < 100 ? `20${String(match[3]).padStart(2,"0")}` : String(match[3]);
    return `${year}-${String(match[2]).padStart(2,"0")}-${String(match[1]).padStart(2,"0")}`;
  }

  const parsed = new Date(txt);
  if(!isNaN(parsed)) return parsed.toISOString().slice(0,10);

  match = txt.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[1]}-${match[2]}-${match[3]}` : null;
}

function inferMonthFromSourceName(sourceName, fullPath = ""){
  const txt = normalizeText((fullPath || sourceName) || "");
  const yearMatch = txt.match(/(20\d{2})/);
  const year = yearMatch ? yearMatch[1] : "";
  const monthAliases = [
    ["ENERO","01"],["FEBRERO","02"],["MARZO","03"],["ABRIL","04"],["MAYO","05"],["JUNIO","06"],
    ["JULIO","07"],["AGOSTO","08"],["SEPTIEMBRE","09"],["SETIEMBRE","09"],["OCTUBRE","10"],["NOVIEMBRE","11"],["DICIEMBRE","12"],
    ["ENE","01"],["FEB","02"],["MAR","03"],["ABR","04"],["MAY","05"],["JUN","06"],["JUL","07"],["AGO","08"],["SEP","09"],["OCT","10"],["NOV","11"],["DIC","12"]
  ];
  const found = monthAliases.find(([label]) => txt.includes(label));
  if(found && year) return `${year}-${found[1]}`;
  return "";
}

function getFechaEntregaFromRow(row, linea){
  // Forzamos búsqueda exacta para evitar confusiones entre columnas con y sin asterisco
  const rowKeys = Object.keys(row);
  if(linea === "NUTRICIONAL") {
     const target = rowKeys.find(k => k.trim().toUpperCase() === "FECHA DE ENTREGA*");
     if(target) return row[target];
  } else {
     const target = rowKeys.find(k => k.trim().toUpperCase() === "FECHA DE ENTREGA");
     if(target) return row[target];
  }
  
  const preferred = linea === "NUTRICIONAL"
    ? ["FECHA DE ENTREGA*", "FECHA ENTREGA*", "FECHA DE ENTREGA", "FECHA ENTREGA"]
    : ["FECHA DE ENTREGA", "FECHA ENTREGA", "FECHA DE ENTREGA*", "FECHA ENTREGA*"];
  return coalesce(row, preferred);
}

function getMonthNameFromDate(iso){
  if(!iso) return "";
  const d = new Date(iso + "T00:00:00");
  return monthOrder[d.getMonth()] || "";
}

function getYearMonth(iso){
  if(!iso) return "";
  return iso.slice(0,7);
}

function coalesceNormalized(norm, keys){
  for(const key of keys){
    const target = normalizeText(key);
    // Buscamos directamente en el objeto ya normalizado
    if(norm[target] !== undefined && norm[target] !== null && String(norm[target]).trim() !== "") {
      return norm[target];
    }
  }
  // Búsqueda por inclusión en las llaves del objeto normalizado
  const rowKeys = Object.keys(norm);
  for(const key of keys){
    const target = normalizeText(key);
    const foundKey = rowKeys.find(rk => rk.includes(target));
    if(foundKey) return norm[foundKey];
  }
  return null;
}

function getFechaEntregaFromNormalized(norm, linea){
  const target = linea === "NUTRICIONAL" ? "FECHA DE ENTREGA*" : "FECHA DE ENTREGA";
  const val = norm[normalizeText(target)];
  if(val !== undefined && val !== null && String(val).trim() !== "") return val;
  
  const preferred = linea === "NUTRICIONAL"
    ? ["FECHA DE ENTREGA*", "FECHA ENTREGA*", "FECHA DE ENTREGA", "FECHA ENTREGA"]
    : ["FECHA DE ENTREGA", "FECHA ENTREGA", "FECHA DE ENTREGA*", "FECHA ENTREGA*"];
  return coalesceNormalized(norm, preferred);
}

function detectSheet(workbook){
  const names = workbook.SheetNames;
  if(names.includes("NUTRICIÓN")) return {sheet: "NUTRICIÓN", linea: "NUTRICIONAL"};
  if(names.includes("ONCOLÓGIA")) return {sheet: "ONCOLÓGIA", linea: "ONCOLOGICO"};
  const normalized = names.map(n => normalizeText(n));
  const nutriIndex = normalized.findIndex(n => n.includes("NUTRICION"));
  if(nutriIndex >= 0) return {sheet: names[nutriIndex], linea: "NUTRICIONAL"};
  const oncoIndex = normalized.findIndex(n => n.includes("ONCOLOGIA"));
  if(oncoIndex >= 0) return {sheet: names[oncoIndex], linea: "ONCOLOGICO"};
  return null;
}

function detectHeaderRow(sheet){
  // IMPORTANTE: NO usar blankrows:false aquí. Necesitamos el índice real de fila en el sheet
  // (no el índice en el array filtrado), porque ese mismo número se pasa al range: de sheet_to_json.
  // Con blankrows:false, la fila 2 aparece en idx=0 y sheet_to_json(range:0) usaría la fila 0
  // (vacía) como cabecera → sin columnas → 0 registros.
  const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:null, raw:true, range:0});
  for(let idx = 0; idx < Math.min(rows.length, 20); idx++){
    const row = (rows[idx] || []).map(value => normalizeText(value || ""));
    const hasHospital = row.some(cell => cell.includes("HOSPITAL") || cell.includes("UNIDAD") || cell.includes("PUNTO DE VENTA"));
    const hasMedicamento = row.some(cell => cell.includes("MEDICAMENTO") || cell.includes("DESCRIPCION"));
    if(hasHospital && hasMedicamento) return idx;
  }
  return 0;
}

function sheetToJsonRows(workbook, sheetName){
  const sheet = workbook.Sheets[sheetName];
  const headerRow = detectHeaderRow(sheet);
  return XLSX.utils.sheet_to_json(sheet, {defval: null, raw: true, range: headerRow});
}

function buildRecord(norm, linea, sourceName){
  const hospitalOriginal = coalesceNormalized(norm, ["HOSPITAL*", "HOSPITAL", "UNIDAD", "PUNTO DE VENTA"]);
  const medicamentoOriginal = coalesceNormalized(norm, ["MEDICAMENTO", "DESCRIPCION", "CONCEPTO", "&&"]);
  const marca = coalesceNormalized(norm, ["MARCA *", "MARCA"]);
  const fabricanteOriginal = coalesceNormalized(norm, ["FABRICANTE *", "FABRICANTE", "LABORATORIO"]);
  const dosis = coalesceNormalized(norm, ["DOSIS CON SOBRELLENADO (mL) *", "DOSIS (mL)", "DOSIS"]);
  const presentacionOriginal = coalesceNormalized(norm, ["PRESENTACIÓN", "PRESENTACION *", "PRESENTACION"]);
  const tipoMedicamento = coalesceNormalized(norm, ["TIPO MEDICAMENTO", "TIPO MEDICAMENTO 1", "TIPO"]);
  const montoTotal = coalesceNormalized(norm, ["TOTAL MEZCLAS", "TOTAL PRECIO MEZCLAS", "IMPORTE TOTAL", "TOTAL"]);
  const fechaEntrega = getFechaEntregaFromNormalized(norm, linea);
  let nombreCliente = coalesceNormalized(norm, ["NOMBRE CLIENTE", "CLIENTE"]);
  const noCliente = coalesceNormalized(norm, ["NO. CLIENTE", "NO CLIENTE", "NUM CLIENTE"]);
  const contrato = coalesceNormalized(norm, ["CONTRATO"]);
  const estatus = coalesceNormalized(norm, ["ESTATUS"]);
  const fechaISO = parseDate(fechaEntrega);
  const inferredYearMonthFromSource = inferMonthFromSourceName(sourceName, norm._fullPath || "");
  const finalYearMonth = getYearMonth(fechaISO) || inferredYearMonthFromSource;
  
  const presentacionValor = extractPresentationValue(presentacionOriginal);
  const frascosFuente = linea === "NUTRICIONAL"
    ? coalesceNormalized(norm, ["FRASCOS CONSUMIDOS CON SOBRELLENADO", "FRASCOS CONSUMIDOS", "FRASCOS"])
    : coalesceNormalized(norm, ["FRASCOS", "FRASCO", "NO FRASCOS"]);
  const frascos = parseNumber(frascosFuente);

  if(String(noCliente || "").trim() === "SON-0028" && (!nombreCliente || ["#N/A","N/A","NAN"].includes(String(nombreCliente).trim().toUpperCase()))){
    nombreCliente = "INSTITUTO DE SEGURIDAD Y SERVICIOS SOCIALES DE LOS TRABAJADORES AL SERVICIO DE LOS PODERES DEL ESTADO DE PUEBLA";
  }

  if(!hospitalOriginal && !medicamentoOriginal) return null;

  const medicamentoHomologado = linea === "ONCOLOGICO"
    ? normalizeMedicamentoOnco(medicamentoOriginal || "")
    : normalizeMedicamentoNutri(medicamentoOriginal || "");

  return {
    archivo_origen: sourceName,
    linea,
    hospital_original: hospitalOriginal || "",
    hospital_homologado: normalizeHospital(hospitalOriginal || ""),
    medicamento_original: medicamentoOriginal || "",
    medicamento_homologado: medicamentoHomologado,
    medicamento_homologado_presentacion: buildMedicamentoHomologadoPresentacion(medicamentoHomologado, presentacionOriginal, presentacionValor),
    marca: normalizeText(marca || ""),
    fabricante_original: fabricanteOriginal || "",
    laboratorio_homologado: normalizeFabricante(fabricanteOriginal || ""),
    dosis: parseNumber(dosis),
    presentacion_original: presentacionOriginal || "",
    presentacion_valor: presentacionValor,
    tipo_medicamento: normalizeText(tipoMedicamento || ""),
    mes: (getMonthNameFromDate(fechaISO) || (coalesceNormalized(norm, ["MES"]) || "").toString().trim().replace(/\s+/g," ")),
    monto_total: parseNumber(montoTotal),
    no_cliente: noCliente || "",
    nombre_cliente: (nombreCliente || "").toString().trim(),
    fecha_entrega: fechaISO,
    dia: fechaISO ? Number(fechaISO.slice(8,10)) : null,
    estatus: (estatus || "").toString().trim(),
    contrato: (contrato || "").toString().trim(),
    consigna: normalizeText(tipoMedicamento || "").includes("CONSIGNA"),
    frascos: frascos !== null && frascos !== undefined ? frascos : (parseNumber(dosis) && presentacionValor ? (parseNumber(dosis) / presentacionValor) : null),
    piezas_por_empaque: 1,
    piezas: null,
    anio_mes: finalYearMonth,
    producto_cliente_laboratorio: "",
    empaque_ambiguo: false,
  };
}

function finalizeRecord(record){
  if(record._finalized) return record;
  // Simplificamos: no creamos la llave pesada producto_cliente_laboratorio aquí para ahorrar memoria en el arranque masivo.
  // Se calculará solo cuando sea necesario para deduplicación o exportación.
  const fixedCliente = String(record.no_cliente || "").trim() === "SON-0028"
    && (!record.nombre_cliente || ["#N/A","N/A","NAN"].includes(String(record.nombre_cliente).trim().toUpperCase()))
      ? "INSTITUTO DE SEGURIDAD Y SERVICIOS SOCIALES DE LOS TRABAJADORES AL SERVICIO DE LOS PODERES DEL ESTADO DE PUEBLA"
      : (record.nombre_cliente || "");
  record.nombre_cliente = String(fixedCliente || "").trim();
  record.laboratorio_homologado = normalizeFabricante(record.laboratorio_homologado || record.fabricante_original || "");
  
  const v = record.presentacion_valor;
  record.presentacion_valor = (v !== null && v !== undefined && v !== "") ? Number(v) : null;
  if((record.presentacion_valor === null || !Number.isFinite(record.presentacion_valor) || record.presentacion_valor <= 0)){
    record.presentacion_valor = extractPresentationValue(record.presentacion_original || "");
  }

  record.frascos = (record.frascos !== null && record.frascos !== undefined && record.frascos !== "") ? Number(record.frascos) : computeFrascosFromRecord(record);
  record.piezas_por_empaque = record.piezas_por_empaque ? Math.round(Number(record.piezas_por_empaque)) : 1;
  record.piezas = computePiezas(record, record.piezas_por_empaque);
  record.medicamento_homologado_presentacion = buildMedicamentoHomologadoPresentacion(record.medicamento_homologado || "", record.presentacion_original || "", record.presentacion_valor);
  
  record._finalized = true;
  return record;
}

function getFullRecordKey(r){
  return `${r.linea || ""} | ${r.nombre_cliente || ""} | ${r.hospital_homologado || ""} | ${r.medicamento_homologado_presentacion || ""} | ${r.laboratorio_homologado || ""}`;
}

function detectPackagingHeaders(row){
  const mapped = {};
  for(const key of Object.keys(row)){
    const n = normalizeText(key);
    if(!mapped.medicamento && n.includes("MEDICAMENTO")) mapped.medicamento = key;
    if(!mapped.tipo && n.includes("TIPO")) mapped.tipo = key;
    if(!mapped.piezas && (n.includes("PIEZA X CAJA") || n.includes("PIEZAS POR EMPAQUE") || (n.includes("PIEZAS") && n.includes("EMPAQUE")))) mapped.piezas = key;
  }
  return mapped;
}

function buildPackagingCatalog(rows){
  if(!rows.length) return [];
  const headers = detectPackagingHeaders(rows[0]);
  if(!headers.medicamento || !headers.tipo || !headers.piezas){
    throw new Error("El archivo de piezas por empaque debe incluir al menos: MEDICAMENTO, Pieza x Caja y TIPO.");
  }
  return rows.map(row => {
    const tipo = normalizeText(row[headers.tipo] || "");
    const medicamento = row[headers.medicamento] || "";
    const piezas = parseNumber(row[headers.piezas]);
    return {
      linea: tipo,
      medicamento_homologado: tipo === "ONCOLOGICO" ? normalizeMedicamentoOnco(medicamento) : normalizeMedicamentoNutri(medicamento),
      piezas_por_empaque: piezas,
    };
  }).filter(item => item.medicamento_homologado && item.linea && item.piezas_por_empaque && item.piezas_por_empaque !== 1);
}

function applyPackagingToData(data){
  if(!state.packagingCatalog.length){
    // Si no hay catálogo, evitamos mapear si no es estrictamente necesario, 
    // pero necesitamos asegurar que los campos existan. 
    // Si ya vienen de finalizeRecord, podemos saltarnos esto si no hay cambios.
    return data;
  }

  const grouped = new Map();
  state.packagingCatalog.forEach(item => {
    const key = `${item.linea}|${item.medicamento_homologado}`;
    if(!grouped.has(key)) grouped.set(key, new Set());
    grouped.get(key).add(Math.round(Number(item.piezas_por_empaque)));
  });

  for(let i=0; i<data.length; i++){
    const record = data[i];
    const key = `${record.linea}|${record.medicamento_homologado}`;
    const valuesSet = grouped.get(key);
    if(!valuesSet) continue;

    const uniqueValues = [...valuesSet].filter(v => Number.isFinite(v) && v !== 1);
    const piezasPorEmpaque = uniqueValues.length === 1 ? Math.round(uniqueValues[0]) : 1;
    
    if(record.piezas_por_empaque === piezasPorEmpaque) continue;

    const frascos = record.frascos !== null && record.frascos !== undefined ? record.frascos : computeFrascosFromRecord(record);
    const piezas = computePiezas({ ...record, frascos }, piezasPorEmpaque);
    
    record.piezas_por_empaque = piezasPorEmpaque;
    record.piezas = (piezas !== null && Number.isFinite(piezas)) ? piezas : null;
    record.empaque_ambiguo = uniqueValues.length > 1;
    record.empaque_enlazado = piezasPorEmpaque > 1 ? "SI" : "NO";
  }
  return data;
}

function getAllData(){
  // Optimizacion radical: si no hay cargas nuevas ni cambios en empaque, devolvemos baseData directamente.
  if(!state.uploadedData.length && !state.packagingCatalog.length) return state.baseData;

  // Evitamos spreads masivos [...state.baseData, ...state.uploadedData]
  let merged = state.baseData.concat(state.uploadedData);
  
  // Si hay catálogo de empaque, lo aplicamos por mutación (in-place)
  if(state.packagingCatalog.length){
    merged = applyPackagingToData(merged);
  }

  // Deduplicación rápida si hay riesgo de solape entre base e importación
  if(state.uploadedData.length > 0){
    const seen = new Set();
    const result = [];
    for(let i=0; i<merged.length; i++){
      const r = merged[i];
      // Clave ligera para reducir impacto en memoria
      const k = `${r.linea[0]}|${r.hospital_homologado.slice(0,5)}|${r.medicamento_homologado.slice(0,10)}|${r.fecha_entrega}|${r.no_cliente}|${r.monto_total}`;
      if(!seen.has(k)){
        seen.add(k);
        result.push(r);
      }
    }
    return result;
  }

  return merged;
}

function formatNumber(value, decimals = 0){
  if(value === null || value === undefined || value === "" || Number.isNaN(value)) return "-";
  return new Intl.NumberFormat("es-MX", {maximumFractionDigits: decimals, minimumFractionDigits: decimals}).format(value);
}

function formatCurrency(value){
  if(value === null || value === undefined || Number.isNaN(value)) return "-";
  return new Intl.NumberFormat("es-MX", {style:"currency", currency:"MXN", maximumFractionDigits:2}).format(value);
}

function setStatus(message){
  const el = document.getElementById("statusBox");
  if(!message) {
    el.innerHTML = "";
    return;
  }
  
  if(message.includes("\n")) {
    const lines = message.split("\n").filter(l => l.trim());
    el.innerHTML = `<div style="display:flex; flex-direction:column; gap:12px;">
      ${lines.map(line => {
        let icon = "◈";
        if(line.includes("agregaron") || line.includes("precargada")) icon = "✓";
        if(line.includes("Iniciando") || line.includes("Procesando")) icon = "🔄";
        if(line.includes("Error") || line.includes("problema")) icon = "⚠";
        return `<div style="display:flex; align-items:start; gap:10px;">
          <span style="opacity:0.6; font-size:16px;">${icon}</span>
          <span>${line}</span>
        </div>`;
      }).join("")}
    </div>`;
  } else {
    el.innerHTML = `<div style="display:flex; align-items:center; gap:10px;"><span style="opacity:0.6; font-size:16px;">◈</span><span>${message}</span></div>`;
  }
}

function showToast(title, message, type = 'success'){
  const container = document.getElementById('toastWrapper') || createToastContainer();
  const toast = document.createElement('div');
  toast.className = `status-box toast-item toast-${type}`;
  toast.style.marginBottom = '10px';
  toast.style.boxShadow = '0 10px 25px rgba(0,0,0,0.3)';
  toast.style.borderLeft = `5px solid ${type === 'success' ? 'var(--success)' : 'var(--danger)'}`;
  toast.style.animation = 'slideInRight 0.4s ease forwards';
  
  toast.innerHTML = `
    <div style="font-weight:800; margin-bottom:4px; font-size:14px; color:#fff;">${title}</div>
    <div style="font-size:12px; opacity:0.9;">${message}</div>
  `;
  
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'fadeOut 0.5s ease forwards';
    setTimeout(() => toast.remove(), 500);
  }, 5000);
}

function createToastContainer(){
  const div = document.createElement('div');
  div.id = 'toastWrapper';
  div.style.position = 'fixed';
  div.style.top = '30px';
  div.style.right = '30px';
  div.style.zIndex = '10000';
  div.style.display = 'flex';
  div.style.flexDirection = 'column';
  div.style.width = '320px';
  document.body.appendChild(div);
  
  const style = document.createElement('style');
  style.innerHTML = `
    @keyframes slideInRight {
      from { transform: translateX(100%); opacity: 0; }
      to { transform: translateX(0); opacity: 1; }
    }
    @keyframes fadeOut {
      from { opacity: 1; transform: scale(1); }
      to { opacity: 0; transform: scale(0.95); }
    }
  `;
  document.head.appendChild(style);
  return div;
}

function populateSelect(id, values, includeAll = true){
  const el = document.getElementById(id);
  const current = el.value;
  el.innerHTML = "";
  if(includeAll){
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "Todos";
    el.appendChild(opt);
  }
  values.forEach(value => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value || "(Vacío)";
    el.appendChild(opt);
  });
  if([...el.options].some(o => o.value === current)) el.value = current;
}

function renderSummary(data){
  const wrap = document.getElementById("summaryCards");
  
  const pricedCount = data.filter(x => state.priceList.some(p => p.medicamento === normalizeText(x.medicamento_homologado_presentacion))).length;
  const packagedCount = data.filter(x => Number(x.piezas_por_empaque) > 1).length;
  const healthPrice = data.length ? (pricedCount / data.length * 100).toFixed(1) : 0;
  const healthPack = data.length ? (packagedCount / data.length * 100).toFixed(1) : 0;

  const totalMonto = data.reduce((a,b)=>a+(Number(b.monto_total)||0),0);
  const totalPiezas = data.reduce((a,b)=>a+(Number(b.piezas)||0),0);
  const totalFrascos = data.reduce((a,b)=>a+(Number(b.frascos)||0),0);

  const cards = [
    {label:"Registros", value:formatNumber(data.length), sub:"Histórico acumulado"},
    {label:"Monto Total", value:formatCurrency(totalMonto), sub:"Inversión histórica"},
    {label:"Volumen (Piezas)", value:formatNumber(totalPiezas,0) + " pzas", sub:"Unidades finales (Empaque)"},
    {label:"Unidades (Frascos)", value:formatNumber(totalFrascos,0) + " unid", sub:"Viales/unidades individuales"},
    {label:"Hospitales", value:formatNumber(new Set(data.map(x=>x.hospital_homologado).filter(Boolean)).size), sub:"Unidades homologadas"},
  ];
  wrap.innerHTML = cards.map(card => `
    <div class="stat-card">
      <div class="label">${card.label}</div>
      <div class="value">${card.value}</div>
      <div class="sub">${card.sub}</div>
    </div>
  `).join("");
}

function renderFilters(data){
  // Optimizacion: Un solo paso sobre los datos para recolectar todos los valores unicos
  // Esto evita crear 7 arreglos temporales de 70k+ elementos
  const sets = {
    linea: new Set(),
    mes: new Set(),
    hospital: new Set(),
    medicamento: new Set(),
    laboratorio: new Set(),
    cliente: new Set(),
    marca: new Set(),
    tipo: new Set()
  };

  for(let i=0; i<data.length; i++){
    const x = data[i];
    if(x.linea) sets.linea.add(x.linea);
    if(x.anio_mes) sets.mes.add(x.anio_mes);
    if(x.hospital_homologado) sets.hospital.add(x.hospital_homologado);
    if(x.medicamento_homologado) sets.medicamento.add(x.medicamento_homologado);
    if(x.laboratorio_homologado) sets.laboratorio.add(x.laboratorio_homologado);
    if(x.nombre_cliente) sets.cliente.add(x.nombre_cliente);
    if(x.marca) sets.marca.add(x.marca);
    if(x.tipo_medicamento) sets.tipo.add(x.tipo_medicamento);
  }

  const formatMonth = (ym) => {
    const [y, m] = ym.split("-");
    return `${monthOrder[Number(m)-1] || m} ${y}`;
  };

  populateSelect("filterLinea", [...sets.linea].sort());
  
  const mesSelect = document.getElementById("filterMes");
  const currentMes = mesSelect.value;
  mesSelect.innerHTML = '<option value="">Todos</option>';
  [...sets.mes].sort().reverse().forEach(ym => {
    const opt = document.createElement("option");
    opt.value = ym;
    opt.textContent = formatMonth(ym);
    mesSelect.appendChild(opt);
  });
  mesSelect.value = currentMes;

  populateSelect("filterHospital", [...sets.hospital].sort());
  populateSelect("filterMedicamento", [...sets.medicamento].sort());
  populateSelect("filterLaboratorio", [...sets.laboratorio].sort());
  populateSelect("filterCliente", [...sets.cliente].sort());
  populateSelect("filterConsigna", ["Sí", "No"]);
  populateSelect("filterMarca", [...sets.marca].sort());
  populateSelect("filterTipo", [...sets.tipo].sort());
}

function getFilteredData(){
  const data = getAllData();
  const linea = document.getElementById("filterLinea").value;
  const mes = document.getElementById("filterMes").value;
  const hospital = document.getElementById("filterHospital").value;
  const medicamento = document.getElementById("filterMedicamento").value;
  const laboratorio = document.getElementById("filterLaboratorio").value;
  const cliente = document.getElementById("filterCliente").value;
  const consigna = document.getElementById("filterConsigna").value;
  const marca = document.getElementById("filterMarca").value;
  const tipo = document.getElementById("filterTipo").value;
  const search = normalizeText(document.getElementById("filterSearch").value);

  return data.filter(row => {
    if(linea && row.linea !== linea) return false;
    if(mes && row.anio_mes !== mes) return false;
    if(hospital && row.hospital_homologado !== hospital) return false;
    if(medicamento && row.medicamento_homologado !== medicamento) return false;
    if(laboratorio && row.laboratorio_homologado !== laboratorio) return false;
    if(cliente && row.nombre_cliente !== cliente) return false;
    if(consigna === "Sí" && !row.consigna) return false;
    if(consigna === "No" && row.consigna) return false;
    if(marca && row.marca !== marca) return false;
    if(tipo && row.tipo_medicamento !== tipo) return false;
    if(search){
      const haystack = normalizeText([
        row.hospital_homologado,row.medicamento_homologado,row.nombre_cliente,row.laboratorio_homologado,row.archivo_origen
      ].join(" "));
      if(!haystack.includes(search)) return false;
    }
    return true;
  });
}

function renderTable(data){
  const head = document.querySelector("#dataTable thead");
  const body = document.querySelector("#dataTable tbody");
  const columns = [
    "linea","mes","fecha_entrega","hospital_homologado","medicamento_homologado","medicamento_homologado_presentacion","marca","laboratorio_homologado",
    "dosis","presentacion_valor","frascos","piezas_por_empaque","piezas","tipo_medicamento","monto_total","no_cliente","nombre_cliente","archivo_origen"
  ];
  head.innerHTML = `<tr>${columns.map(col => `<th>${col}</th>`).join("")}</tr>`;
    body.innerHTML = data.slice(0, 500).map(row => `
    <tr>
      ${columns.map(col => {
        let value = row[col];
        if(col === "monto_total") return `<td>${formatCurrency(Number(value))}</td>`;
        if(["dosis","presentacion_valor","frascos","piezas_por_empaque","piezas"].includes(col)) return `<td>${formatNumber(Number(value), col === "piezas" ? 2 : 4)}</td>`;
        if(col === "medicamento_homologado" && row.empaque_ambiguo) {
          return `<td><span title="Diferentes piezas por empaque detectadas para este producto" style="cursor:help">⚠ ${value ?? ""}</span></td>`;
        }
        return `<td>${value ?? ""}</td>`;
      }).join("")}
    </tr>
  `).join("");
  document.getElementById("tableCount").textContent = `Mostrando ${Math.min(data.length,500)} de ${formatNumber(data.length)} registros filtrados`;
}

function aggregateBy(items, keyFn, metric){
  const map = new Map();
  items.forEach(item => {
    const key = keyFn(item);
    const value = Number(item[metric]) || 0;
    if(!map.has(key)) map.set(key, 0);
    map.set(key, map.get(key) + value);
  });
  return map;
}

function renderCharts(data){
  const monthKeys = [...new Set(data.map(x => x.anio_mes).filter(Boolean))].sort();
  const lineas = [...new Set(data.map(x => x.linea))];
  const colors = ['#38bdf8', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'];
  
  const mainMetric = document.getElementById("mainMetric")?.value || "monto_total";
  const isPiezas = mainMetric === "piezas";
  const labelPrefix = isPiezas ? "Volumen" : "Monto";

  const lineDatasets = lineas.map((linea, idx) => {
    const color = colors[idx % colors.length];
    return {
      label: linea,
      backgroundColor: color,
      borderColor: color,
      borderWidth: 2,
      borderRadius: 6,
      data: monthKeys.map(month => data.filter(x => x.linea === linea && x.anio_mes === month).reduce((a,b)=>a+(Number(b[mainMetric])||0),0)),
    };
  });

  if(state.monthlyChart) state.monthlyChart.destroy();
  state.monthlyChart = new Chart(document.getElementById("monthlyChart"), {
    type: "bar",
    data: {labels: monthKeys, datasets: lineDatasets},
    options: {
      ...CHART_OPTIONS,
      scales: {
        ...CHART_OPTIONS.scales,
        y: { 
          ...CHART_OPTIONS.scales.y, 
          beginAtZero: true, 
          ticks: { 
            ...CHART_OPTIONS.scales.y.ticks, 
            callback: v => isPiezas ? formatNumber(v, 0) : formatCurrency(v) 
          } 
        }
      }
    }
  });

  // Top Labs
  const labMap = aggregateBy(data, x => x.laboratorio_homologado || "(Sin laboratorio)", mainMetric);
  const topLabs = [...labMap.entries()].sort((a,b)=>b[1]-a[1]).slice(0,10);
  if(state.labChart) state.labChart.destroy();
  state.labChart = new Chart(document.getElementById("labChart"), {
    type: "bar",
    data: {
      labels: topLabs.map(x=>x[0]), 
      datasets:[{
        label: labelPrefix, 
        data: topLabs.map(x=>x[1]), 
        backgroundColor: '#38bdf8',
        borderRadius: 8,
        borderSkipped: false
      }]
    },
    options: {
      ...CHART_OPTIONS, 
      indexAxis:"y",
      scales: {
        x: { 
          ...CHART_OPTIONS.scales.x, 
          beginAtZero: true, 
          ticks: { ...CHART_OPTIONS.scales.x.ticks, callback: v => isPiezas ? formatNumber(v, 0) : formatCurrency(v) } 
        },
        y: { ...CHART_OPTIONS.scales.y, grid: { display: false } }
      }
    }
  });

  // Top Hospitales
  const hospMap = aggregateBy(data, x => x.hospital_homologado || "(Sin hospital)", mainMetric);
  const topHosps = [...hospMap.entries()].sort((a,b)=>b[1]-a[1]).slice(0,10);
  if(state.hospChart) state.hospChart.destroy();
  state.hospChart = new Chart(document.getElementById("hospChart"), {
    type: "bar",
    data: {
      labels: topHosps.map(x=>x[0]), 
      datasets:[{
        label: labelPrefix, 
        data: topHosps.map(x=>x[1]), 
        backgroundColor: '#10b981',
        borderRadius: 8,
        borderSkipped: false
      }]
    },
    options: {
      ...CHART_OPTIONS, 
      indexAxis:"y",
      scales: {
        x: { 
          ...CHART_OPTIONS.scales.x, 
          beginAtZero: true, 
          ticks: { ...CHART_OPTIONS.scales.x.ticks, callback: v => isPiezas ? formatNumber(v, 0) : formatCurrency(v) } 
        },
        y: { ...CHART_OPTIONS.scales.y, grid: { display: false } }
      }
    }
  });

  // Market Share (Lab Distribution)
  const labShareMap = aggregateBy(data, x => x.laboratorio_homologado || "(Otro)", mainMetric);
  const sortedLabs = [...labShareMap.entries()].sort((a,b)=>b[1]-a[1]);
  const mainLabs = sortedLabs.slice(0, 5);
  const otherSum = sortedLabs.slice(5).reduce((acc, curr) => acc + curr[1], 0);
  if(otherSum > 0) mainLabs.push(["Resto", otherSum]);

  if(state.marketShareChart) state.marketShareChart.destroy();
  state.marketShareChart = new Chart(document.getElementById("marketShareChart"), {
    type: "doughnut",
    data: {
      labels: mainLabs.map(x=>x[0]),
      datasets: [{
        data: mainLabs.map(x=>x[1]),
        backgroundColor: ['#38bdf8', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#94a3b8'],
        borderWidth: 0,
        hoverOffset: 20
      }]
    },
    options: {
      ...CHART_OPTIONS,
      cutout: '72%',
      plugins: {
        ...CHART_OPTIONS.plugins,
        legend: { position: 'right', labels: { color: '#94a3b8', boxWidth: 12, padding: 20, font: { family: 'Outfit', size: 12, weight: '600' } } },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const label = ctx.label || '';
              const value = ctx.parsed || 0;
              const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
              const percentage = ((value/total)*100).toFixed(1);
              return ` ${label}: ${isPiezas ? formatNumber(value, 2) : formatCurrency(value)} (${percentage}%)`;
            }
          }
        }
      }
    }
  });

  // Gráfica: Línea, medicamento, presentacion, mes -> piezas (Top 10 agregados por medicamento+presentacion)
  const medPresMap = aggregateBy(data, x => `${x.linea} | ${x.medicamento_homologado} | ${x.presentacion_original || x.presentacion_valor || ""}`, isPiezas ? "piezas" : mainMetric);
  const topMedPres = [...medPresMap.entries()].sort((a,b)=>b[1]-a[1]).slice(0, 5);
  const medPresLabels = topMedPres.map(x=>x[0]);
  
  if(state.medPresChart) state.medPresChart.destroy();
  state.medPresChart = new Chart(document.getElementById("medPresChart"), {
    type: "line",
    data: {
      labels: monthKeys,
      datasets: medPresLabels.map((lbl, idx) => ({
        label: lbl,
        data: monthKeys.map(m => data.filter(x => x.anio_mes === m && `${x.linea} | ${x.medicamento_homologado} | ${x.presentacion_original || x.presentacion_valor || ""}` === lbl).reduce((a,b)=>a+(Number(b[isPiezas ? "piezas" : mainMetric])||0),0)),
        borderColor: colors[idx % colors.length],
        tension: 0.3
      }))
    },
    options: {
      ...CHART_OPTIONS,
      scales: {
        x: { ...CHART_OPTIONS.scales.x },
        y: { ...CHART_OPTIONS.scales.y, beginAtZero: true }
      }
    }
  });

  // Gráfica: Hospital, Línea, medicamento, presentacion, mes -> piezas
  const hospMedPresMap = aggregateBy(data, x => `${x.hospital_homologado} | ${x.linea} | ${x.medicamento_homologado} | ${x.presentacion_original || x.presentacion_valor || ""}`, isPiezas ? "piezas" : mainMetric);
  const topHospMedPres = [...hospMedPresMap.entries()].sort((a,b)=>b[1]-a[1]).slice(0, 5);
  const hospMedPresLabels = topHospMedPres.map(x=>x[0].substring(0, 60)); // truncate long
  
  if(state.hospMedPresChart) state.hospMedPresChart.destroy();
  state.hospMedPresChart = new Chart(document.getElementById("hospMedPresChart"), {
    type: "bar",
    data: {
      labels: monthKeys,
      datasets: topHospMedPres.map((item, idx) => ({
        label: hospMedPresLabels[idx],
        data: monthKeys.map(m => data.filter(x => x.anio_mes === m && `${x.hospital_homologado} | ${x.linea} | ${x.medicamento_homologado} | ${x.presentacion_original || x.presentacion_valor || ""}` === item[0]).reduce((a,b)=>a+(Number(b[isPiezas ? "piezas" : mainMetric])||0),0)),
        backgroundColor: colors[idx % colors.length]
      }))
    },
    options: {
      ...CHART_OPTIONS,
      scales: {
        x: { ...CHART_OPTIONS.scales.x, stacked: true },
        y: { ...CHART_OPTIONS.scales.y, stacked: true, beginAtZero: true }
      }
    }
  });

  // Interacción Drill-down: Click en el gráfico de Market Share para filtrar por Lab
  document.getElementById("marketShareChart").onclick = (evt) => {
    const points = state.marketShareChart.getElementsAtEventForMode(evt, 'nearest', { intersect: true }, true);
    if (points.length) {
      const firstPoint = points[0];
      const label = state.marketShareChart.data.labels[firstPoint.index];
      if (label && label !== "Resto") {
        const sel = document.getElementById("filterLaboratorio");
        sel.value = label;
        refresh();
      }
    }
  };
}

function nextMonths(lastYearMonth, count){
  if(!lastYearMonth) return [];
  const [year,month] = lastYearMonth.split("-").map(Number);
  const out = [];
  let y = year, m = month;
  for(let i=0;i<count;i++){
    m += 1;
    if(m > 12){ m = 1; y += 1; }
    out.push(`${y}-${String(m).padStart(2,"0")}`);
  }
  return out;
}

function weightedMovingAverage(series){
  const vals = series.slice(-3);
  if(!vals.length) return 0;
  const weights = vals.length === 1 ? [1] : vals.length === 2 ? [0.4,0.6] : [0.2,0.3,0.5];
  return vals.reduce((sum, value, idx) => sum + value * weights[idx], 0);
}

function simpleMovingAverage(series){
  const vals = series.slice(-3);
  if(!vals.length) return 0;
  return vals.reduce((sum, value) => sum + value, 0) / vals.length;
}

function trendAverageForecast(series, periods, isPieces){
  const diffs = series.slice(1).map((v, i) => v - series[i]);
  let avgDiff = 0;
  if(diffs.length){
    // Mayor peso en la tendencia reciente
    const weights = diffs.map((_, i) => i + 1);
    const sumW = weights.reduce((a, b) => a + b, 0);
    avgDiff = diffs.reduce((a, b, i) => a + b * (weights[i] / sumW), 0);
  }
  let current = series.length ? series[series.length - 1] : 0;
  const out = [];
  for(let i = 0; i < periods; i++){
    current = Math.max(0, current + avgDiff);
    out.push(isPieces ? Number(Math.max(0, current).toFixed(4)) : Number(current.toFixed(2)));
  }
  return out;
}

function flatForecast(series, periods, isPieces){
  const base = series.length ? Math.max(0, Number(series[series.length - 1]) || 0) : 0;
  return Array.from({length: periods}, () => isPieces ? Number(base.toFixed(4)) : Number(base.toFixed(2)));
}

function movingAverageForecast(series, periods, isPieces, weighted){
  const history = series.slice();
  const out = [];
  for(let i = 0; i < periods; i++){
    const value = weighted ? weightedMovingAverage(history) : simpleMovingAverage(history);
    const clean = Math.max(0, Number(value) || 0);
    history.push(clean);
    out.push(isPieces ? Number(clean.toFixed(4)) : Number(clean.toFixed(2)));
  }
  return out;
}

function recentMape(series, forecastFn){
  if(series.length < 4) return null;
  const actuals = series.slice(-3);
  const train = series.slice(0, -3);
  if(!train.length) return null;
  const predicted = forecastFn(train, 3, false);
  const errors = [];
  for(let i = 0; i < actuals.length; i++){
    const actual = Number(actuals[i]) || 0;
    const pred = Number(predicted[i]) || 0;
    if(actual <= 0) continue;
    errors.push(Math.abs((actual - pred) / actual));
  }
  if(!errors.length) return null;
  return Number(((errors.reduce((a, b) => a + b, 0) / errors.length) * 100).toFixed(2));
}

function getForecastModels(){
  return [
    { key: "trend", label: "Tendencia promedio", forecast: (series, periods, isPieces) => trendAverageForecast(series, periods, isPieces) },
    { key: "wma", label: "Promedio móvil ponderado", forecast: (series, periods, isPieces) => movingAverageForecast(series, periods, isPieces, true) },
    { key: "sma", label: "Promedio móvil simple", forecast: (series, periods, isPieces) => movingAverageForecast(series, periods, isPieces, false) },
    { key: "flat", label: "Último valor", forecast: (series, periods, isPieces) => flatForecast(series, periods, isPieces) },
  ];
}

function pickBestForecastModel(metrics){
  const models = getForecastModels();
  const evaluations = models.map(model => ({
    ...model,
    mape: recentMape(metrics.monto_total, model.forecast)
  }));
  const withMape = evaluations.filter(item => item.mape !== null && Number.isFinite(item.mape));
  let selected = models.find(m => m.key === "wma") || models[0];
  if(withMape.length > 0){
    withMape.forEach(item => {
      if(item.key === "trend") item.mape = item.mape * 0.7; // Mayor peso a la tendencia
    });
    selected = withMape.sort((a,b) => a.mape - b.mape)[0];
  }
  return { selected, evaluations };
}

function buildDetailedForecastRows(data){
  if(!data || !data.length) return {rows: [], nextMonths: []};
  
  // Optimizacion: Calculamos meses una sola vez
  const monthsSet = new Set();
  const valid = [];
  for(let i=0; i<data.length; i++) {
    const r = data[i];
    if(r.anio_mes) {
      valid.push(r);
      monthsSet.add(r.anio_mes);
    }
  }
  const allMonths = [...monthsSet].sort();
  if(!allMonths.length) return {rows: [], nextMonths: []};

  const grouped = new Map();
  for(let i=0; i<valid.length; i++) {
    const row = valid[i];
    const key = `${row.linea}|${row.nombre_cliente}|${row.hospital_homologado}|${row.medicamento_homologado_presentacion}|${row.laboratorio_homologado}`;
    
    let g = grouped.get(key);
    if(!g){
      g = {
        linea: row.linea || "",
        nombre_cliente: row.nombre_cliente || "",
        hospital_homologado: row.hospital_homologado || "",
        medicamento_homologado_presentacion: row.medicamento_homologado_presentacion || "",
        laboratorio_homologado: row.laboratorio_homologado || "",
        months: new Map(),
      };
      grouped.set(key, g);
    }
    
    const ym = row.anio_mes;
    if(!g.months.has(ym)) g.months.set(ym, {monto_total:0, piezas:0, monto_sin_consigna:0, piezas_sin_consigna:0});
    const bucket = g.months.get(ym);
    bucket.monto_total += Number(row.monto_total) || 0;
    bucket.piezas += Number(row.piezas) || 0;
    if(!row.consigna){
      bucket.monto_sin_consigna += Number(row.monto_total) || 0;
      bucket.piezas_sin_consigna += Number(row.piezas) || 0;
    }
  }

  const next = nextMonths(allMonths[allMonths.length - 1], 3);
  const rows = [];
  const models = getForecastModels();

  for(const g of grouped.values()){
    const metrics = {
      monto_total: allMonths.map(m => g.months.get(m)?.monto_total || 0),
      piezas: allMonths.map(m => g.months.get(m)?.piezas || 0),
      monto_sin_consigna: allMonths.map(m => g.months.get(m)?.monto_sin_consigna || 0),
      piezas_sin_consigna: allMonths.map(m => g.months.get(m)?.piezas_sin_consigna || 0),
    };

    const modelChoice = pickBestForecastModel(metrics);
    const chosen = modelChoice.selected;
    const unitPrice = findPrice(g.medicamento_homologado_presentacion);
    
    const row = {
      linea: g.linea,
      nombre_cliente: g.nombre_cliente,
      hospital_homologado: g.hospital_homologado,
      medicamento_homologado_presentacion: g.medicamento_homologado_presentacion,
      laboratorio_homologado: g.laboratorio_homologado,
      metodo: chosen.label,
      __sort: 0
    };

    const selectedPiezas = chosen.forecast(metrics.piezas, 3, true);
    const selectedPiezasSin = chosen.forecast(metrics.piezas_sin_consigna, 3, true);
    let selectedMonto, selectedMontoSin;
    
    if(unitPrice !== null){
      selectedMonto = selectedPiezas.map(p => Number((p * unitPrice).toFixed(2)));
      selectedMontoSin = selectedPiezasSin.map(p => Number((p * unitPrice).toFixed(2)));
    } else {
      selectedMonto = chosen.forecast(metrics.monto_total, 3, false);
      selectedMontoSin = chosen.forecast(metrics.monto_sin_consigna, 3, false);
    }

    next.forEach((m, idx) => {
      row[`piezas_${m}`] = selectedPiezas[idx];
      row[`piezas_sin_consigna_${m}`] = selectedPiezasSin[idx];
      row[`monto_total_${m}`] = selectedMonto[idx];
      row[`monto_sin_consigna_${m}`] = selectedMontoSin[idx];
      row.__sort += row[`monto_total_${m}`];
    });

    rows.push(row);
  }

  rows.sort((a,b)=>b.__sort - a.__sort);
  return {rows, nextMonths: next};
}

function withMapeLabel(evaluations){
  const withMape = evaluations.filter(item => item.mape !== null && Number.isFinite(item.mape));
  if(!withMape.length) return "Sin histórico suficiente: se usa Tendencia promedio";
  const best = [...withMape].sort((a, b) => a.mape - b.mape)[0];
  return `Menor error reciente (%MAPE): ${best.label}`;
}


function buildForecastSummaryRows(detailRows, next){
  return next.map(month => ({
    Mes: month,
    "Monto total": Number(detailRows.reduce((acc, row)=>acc + (Number(row[`monto_total_${month}`]) || 0), 0).toFixed(2)),
    "Piezas total": Number(detailRows.reduce((acc, row)=>acc + (Number(row[`piezas_${month}`]) || 0), 0).toFixed(4)),
    "Monto sin consigna": Number(detailRows.reduce((acc, row)=>acc + (Number(row[`monto_sin_consigna_${month}`]) || 0), 0).toFixed(2)),
    "Piezas sin consigna": Number(detailRows.reduce((acc, row)=>acc + (Number(row[`piezas_sin_consigna_${month}`]) || 0), 0).toFixed(4)),
  }));
}


function buildAggregateForecastSummaryRows(data, next){
  const valid = data.filter(row => row.anio_mes);
  const months = [...new Set(valid.map(r => r.anio_mes).filter(Boolean))].sort();
  if(!months.length) return next.map(month => ({Mes: month, "Monto total": 0, "Piezas total": 0, "Monto sin consigna": 0, "Piezas sin consigna": 0, metodo_importe: "N/D", metodo_piezas: "N/D"}));

  const monthly = new Map();
  valid.forEach(row => {
    const ym = row.anio_mes;
    if(!monthly.has(ym)) monthly.set(ym, {monto_total:0, piezas:0, monto_sin_consigna:0, piezas_sin_consigna:0});
    const bucket = monthly.get(ym);
    bucket.monto_total += Number(row.monto_total) || 0;
    bucket.piezas += Number(row.piezas) || 0;
    if(!row.consigna){
      bucket.monto_sin_consigna += Number(row.monto_total) || 0;
      bucket.piezas_sin_consigna += Number(row.piezas) || 0;
    }
  });

  const metrics = {
    monto_total: months.map(m => monthly.get(m)?.monto_total || 0),
    piezas: months.map(m => monthly.get(m)?.piezas || 0),
    monto_sin_consigna: months.map(m => monthly.get(m)?.monto_sin_consigna || 0),
    piezas_sin_consigna: months.map(m => monthly.get(m)?.piezas_sin_consigna || 0),
  };

  const amountChoice = pickBestForecastModel({monto_total: metrics.monto_total});
  const piecesChoice = pickBestForecastModel({monto_total: metrics.piezas});
  const amountNoChoice = pickBestForecastModel({monto_total: metrics.monto_sin_consigna});
  const piecesNoChoice = pickBestForecastModel({monto_total: metrics.piezas_sin_consigna});

  const montoFc = amountChoice.selected.forecast(metrics.monto_total, 3, false);
  const piezasFc = piecesChoice.selected.forecast(metrics.piezas, 3, true);
  const montoNoFc = amountNoChoice.selected.forecast(metrics.monto_sin_consigna, 3, false);
  const piezasNoFc = piecesNoChoice.selected.forecast(metrics.piezas_sin_consigna, 3, true);

  return next.map((month, idx) => ({
    Mes: month,
    "Monto total": Number((montoFc[idx] || 0).toFixed(2)),
    "Piezas total": Number((piezasFc[idx] || 0).toFixed(4)),
    "Monto sin consigna": Number((montoNoFc[idx] || 0).toFixed(2)),
    "Piezas sin consigna": Number((piezasNoFc[idx] || 0).toFixed(4)),
    metodo_importe: amountChoice.selected.label,
    metodo_piezas: piecesChoice.selected.label,
  }));
}

function buildForecastRows(data, metric, level){
  const valid = data.filter(row => row.anio_mes && row[metric] !== null && row[metric] !== undefined && row[metric] !== "");
  const groupKey = level === "overall"
    ? () => "General"
    : level === "detalle_demanda"
      ? (row) => [row.linea || "", row.nombre_cliente || "", row.hospital_homologado || "", row.medicamento_homologado_presentacion || "", row.laboratorio_homologado || ""].join(" | ")
      : (row) => row.producto_cliente_laboratorio || "(Sin llave)";
  const byGroup = new Map();
  const historyMonths = [...new Set(valid.map(r => r.anio_mes))].sort();
  const next = nextMonths(historyMonths[historyMonths.length - 1], 3);

  valid.forEach(row => {
    const key = groupKey(row);
    if(!byGroup.has(key)) byGroup.set(key, new Map());
    const monthMap = byGroup.get(key);
    const month = row.anio_mes;
    monthMap.set(month, (monthMap.get(month) || 0) + (Number(row[metric]) || 0));
  });

  const rows = [];
  for(const [group, monthMap] of byGroup.entries()){
    const values = historyMonths.map(m => monthMap.get(m) || 0);
    const diffs = values.slice(1).map((v,i) => v - values[i]);
    const avgDiff = diffs.length ? diffs.reduce((a,b)=>a+b,0) / diffs.length : 0;
    let current = values.length ? values[values.length - 1] : 0;
    const fc = [];
    for(let i = 0; i < 3; i++){
      current = Math.max(0, current + avgDiff);
      fc.push(metric === "piezas" ? Number(Math.max(0, current).toFixed(4)) : Number(current.toFixed(2)));
    }
    rows.push({
      grupo: group,
      historial_meses: historyMonths.join(", "),
      metodo: "Tendencia promedio",
      [next[0] || "Mes+1"]: fc[0],
      [next[1] || "Mes+2"]: fc[1],
      [next[2] || "Mes+3"]: fc[2],
      valor_base: values.reduce((a,b)=>a+b,0),
    });
  }
  return rows.sort((a,b)=>b.valor_base-a.valor_base).slice(0, level === "overall" ? 1 : 200);
}

function renderForecastChart(summaryRows){
  const canvas = document.getElementById("forecastChart");
  if(state.forecastChart) state.forecastChart.destroy();
  if(!summaryRows.length){
    state.forecastChart = new Chart(canvas, {
      type: "bar",
      data: {labels:["Sin datos"], datasets:[{label:"Pronóstico", data:[0]}]},
      options: {responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}}
    });
    return;
  }

  const metric = document.getElementById("forecastMetric").value;
  const isPiezas = metric === "piezas";
  const labelPrefix = isPiezas ? "Piezas" : "Monto";

  state.forecastChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: summaryRows.map(r => r.Mes),
      datasets: [
        {
          label: `${labelPrefix} total`, 
          data: summaryRows.map(r => Number(r[`${labelPrefix} total`]) || 0), 
          backgroundColor: '#38bdf8', 
          borderRadius: 6
        },
        {
          label: `${labelPrefix} sin consigna`, 
          data: summaryRows.map(r => Number(r[`${labelPrefix} sin consigna`]) || 0), 
          backgroundColor: '#8b5cf6', 
          borderRadius: 6
        }
      ]
    },
    options: {
      ...CHART_OPTIONS,
      scales: {
        x: { ...CHART_OPTIONS.scales.x, grid: { display: false } },
        y: { 
          ...CHART_OPTIONS.scales.y, 
          beginAtZero: true, 
          ticks: { 
            ...CHART_OPTIONS.scales.y.ticks, 
            callback: v => isPiezas ? formatNumber(v, 2) : formatCurrency(v) 
          } 
        }
      }
    }
  });
}

function renderForecastHistoryChart(data, summaryRows){
  const canvas = document.getElementById("forecastHistoryChart");
  if(state.forecastHistoryChart) state.forecastHistoryChart.destroy();

  const metric = document.getElementById("forecastMetric").value;
  const isPiezas = metric === "piezas";
  const labelPrefix = isPiezas ? "piezas" : "monto";
  const unitSuffix = isPiezas ? " pzas" : "";

  // Si hay un grupo seleccionado, filtramos la data histórica para ese grupo
  let filteredHistory = data.filter(row => row.anio_mes);
  let chartTitle = "General (Top-down)";

  if(state.selectedForecastGroup){
    filteredHistory = filteredHistory.filter(row => {
      // Intentamos matchear por la llave compuesta del grupo
      const key = `${row.linea || ""}|${row.nombre_cliente || ""}|${row.hospital_homologado || ""}|${row.medicamento_homologado_presentacion || ""}|${row.laboratorio_homologado || ""}`;
      return key === state.selectedForecastGroup;
    });
    chartTitle = "Detalle de Grupo (Bottom-up)";
    
    // Si el grupo no tiene data (por filtros), volvemos al general
    if(!filteredHistory.length) {
      filteredHistory = data.filter(row => row.anio_mes);
      chartTitle = "General (Sin data para selección)";
    }
  }

  const monthlyMap = new Map();
  const monthlyNo = new Map();
  const metricKey = isPiezas ? "piezas" : "monto_total";

  filteredHistory.forEach(row => {
    const month = row.anio_mes;
    const val = Number(row[metricKey]) || 0;
    monthlyMap.set(month, (monthlyMap.get(month) || 0) + val);
    if(!row.consigna){
      monthlyNo.set(month, (monthlyNo.get(month) || 0) + val);
    }
  });

  // Los summaryRows del pronóstico ya vienen del build correspondiente (consolidado o específico)
  // Pero si tenemos un grupo seleccionado, necesitamos los summaryRows de ESE grupo específicamente
  let fcSource = summaryRows;
  if(state.selectedForecastGroup){
     const detailedResult = buildDetailedForecastRows(data);
     const targetRow = detailedResult.rows.find(r => `${r.linea || ""}|${r.nombre_cliente || ""}|${r.hospital_homologado || ""}|${r.medicamento_homologado_presentacion || ""}|${r.laboratorio_homologado || ""}` === state.selectedForecastGroup);
     if(targetRow){
        const next = detailedResult.nextMonths;
        fcSource = next.map(m => ({
          Mes: m,
          [`${isPiezas ? "Piezas" : "Monto"} total`]: targetRow[isPiezas ? `piezas_${m}` : `monto_total_${m}`],
          [`${isPiezas ? "Piezas" : "Monto"} sin consigna`]: targetRow[isPiezas ? `piezas_sin_consigna_${m}` : `monto_sin_consigna_${m}`]
        }));
     }
  }

  const historyMonths = [...new Set(data.map(r => r.anio_mes).filter(Boolean))].sort();
  const labels = [...historyMonths, ...fcSource.map(r => r.Mes)];
  
  const mLabel = isPiezas ? "Piezas" : "Monto";
  const histTotal = [...historyMonths.map(m => monthlyMap.get(m) ?? null), ...fcSource.map(()=>null)];
  const histNo = [...historyMonths.map(m => monthlyNo.get(m) ?? null), ...fcSource.map(()=>null)];
  const fcTotal = [...historyMonths.map(()=>null), ...fcSource.map(r => Number(r[`${mLabel} total`]) || 0)];
  const fcNo = [...historyMonths.map(()=>null), ...fcSource.map(r => Number(r[`${mLabel} sin consigna`]) || 0)];

  state.forecastHistoryChart = new Chart(canvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {label:`Histórico ${labelPrefix} total`, data: histTotal, tension:0.3, borderColor: '#38bdf8', backgroundColor: 'rgba(56, 189, 248, 0.1)', fill: true, pointRadius: 4},
        {label:`Histórico ${labelPrefix} sin consigna`, data: histNo, tension:0.3, borderColor: '#10b981', borderDash: [5, 5], pointRadius: 4},
        {label:`Pronóstico ${labelPrefix} total`, data: fcTotal, tension:0.3, borderColor: '#f59e0b', backgroundColor: 'rgba(245, 158, 11, 0.1)', fill: true, borderDash: [2, 2], pointRadius: 6, pointStyle: 'rectRot'},
        {label:`Pronóstico ${labelPrefix} sin consigna`, data: fcNo, tension:0.3, borderColor: '#ef4444', borderDash: [2, 2], pointRadius: 6, pointStyle: 'rectRot'}
      ]
    },
    options: {
      ...CHART_OPTIONS,
      plugins: {
        ...CHART_OPTIONS.plugins,
        title: { display: true, text: chartTitle, color: '#94a3b8', font: { size: 12 } },
        tooltip: {
          callbacks: {
            label: (ctx) => ` ${ctx.dataset.label}: ${isPiezas ? formatNumber(ctx.parsed.y, 2) : formatCurrency(ctx.parsed.y)}${unitSuffix}`
          }
        }
      },
      scales: {
        x: { ...CHART_OPTIONS.scales.x, grid: { display: false } },
        y: { 
          ...CHART_OPTIONS.scales.y, 
          beginAtZero: true, 
          title: { display: true, text: mLabel, color: '#94a3b8' },
          ticks: { ...CHART_OPTIONS.scales.y.ticks, callback: v => isPiezas ? formatNumber(v, 0) : formatCurrency(v) }
        }
      }
    }
  });
}

function renderForecastModelChart(modelRows){
  const canvas = document.getElementById("forecastModelChart");
  if(!canvas) return;
  if(state.forecastModelChart) state.forecastModelChart.destroy();
  const usable = (modelRows || []).filter(r => r.Modelo !== "Sin método");
  if(!usable.length){
    state.forecastModelChart = new Chart(canvas, {
      type: "bar",
      data: {labels:["Sin datos"], datasets:[{label:"Participación %", data:[0]}]},
      options: {responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}}
    });
    return;
  }
  state.forecastModelChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: usable.map(r => r.Modelo),
      datasets: [{label:"Participación %", data: usable.map(r => Number(r.Participacion) || 0)}]
    },
    options: {responsive:true, maintainAspectRatio:false}
  });
}

function renderForecastHighlights(cards){
  const wrap = document.getElementById("forecastHighlights");
  wrap.innerHTML = cards.map(card => `
    <div class="stat-card compact">
      <div class="label">${card.label}</div>
      <div class="value">${card.value}</div>
      <div class="sub">${card.sub}</div>
    </div>
  `).join("");
}

function renderSimpleTable(tableId, rows){
  const head = document.querySelector(`#${tableId} thead`);
  const body = document.querySelector(`#${tableId} tbody`);
  if(!rows.length){
    head.innerHTML = "";
    body.innerHTML = `<tr><td>Sin datos</td></tr>`;
    return;
  }
  const columns = Object.keys(rows[0]);
  head.innerHTML = `<tr>${columns.map(col => `<th>${col}</th>`).join("")}</tr>`;
  body.innerHTML = rows.map(row => `
    <tr>${columns.map(col => {
      const val = row[col];
      const colLow = col.toLowerCase();

      if(colLow.includes("metodo") || col === "Modelo") {
        let badgeClass = "badge-accent";
        if(String(val).includes("Tendencia")) badgeClass = "badge-success";
        if(String(val).includes("ponderado")) badgeClass = "badge-warning";
        if(String(val).includes("Último")) badgeClass = "badge-warning";
        return `<td><span class="badge ${badgeClass}">${val ?? ""}</span></td>`;
      }

      if(typeof val === "number"){
        if(/error|mape|participacion/i.test(col)) return `<td><span class="badge badge-accent">${formatNumber(val,2)}%</span></td>`;
        if(/monto|importe/i.test(colLow)) return `<td>${formatCurrency(val)}</td>`;
        if(/piezas/i.test(colLow)) return `<td>${formatNumber(val, 2)}</td>`;
        return `<td>${formatNumber(val, 2)}</td>`;
      }
      return `<td>${val ?? ""}</td>`;
    }).join("")}</tr>
  `).join("");
}

function renderForecast(data){
  const metric = document.getElementById("forecastMetric").value;
  const level = document.getElementById("forecastLevel").value;
  const result = buildDetailedForecastRows(data);
  const rows = result.rows;

  const aggSummaryRows = buildAggregateForecastSummaryRows(data, result.nextMonths);

  const methodCounts = rows.reduce((acc, row) => {
    const key = row.metodo || "Sin método";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  const modelRows = Object.entries(methodCounts)
    .sort((a,b) => b[1] - a[1])
    .map(([modelo, grupos]) => ({Modelo:modelo, Grupos:grupos, Participacion: rows.length ? Number(((grupos / rows.length) * 100).toFixed(2)) : 0}));

  document.getElementById("forecastInfo").textContent = `Pronóstico visible a 3 meses. Se evalúan 4 modelos por grupo: Tendencia promedio, Promedio móvil ponderado, Promedio móvil simple y Último valor. El resumen mensual y los gráficos generales usan la serie total mensual consolidada (Top-down) para evitar distorsiones. El detalle de la tabla usa el modelo óptimo calculado individualmente para cada grupo (Bottom-up).`;

  renderForecastHighlights([
    {label:"Mes 1", value: aggSummaryRows[0] ? formatCurrency(aggSummaryRows[0]["Monto total"]) + " / " + formatNumber(aggSummaryRows[0]["Piezas total"], 0) + " pzas" : "-", sub: result.nextMonths[0] || "Sin mes"},
    {label:"Mes 2", value: aggSummaryRows[1] ? formatCurrency(aggSummaryRows[1]["Monto total"]) + " / " + formatNumber(aggSummaryRows[1]["Piezas total"], 0) + " pzas" : "-", sub: result.nextMonths[1] || "Sin mes"},
    {label:"Mes 3", value: aggSummaryRows[2] ? formatCurrency(aggSummaryRows[2]["Monto total"]) + " / " + formatNumber(aggSummaryRows[2]["Piezas total"], 0) + " pzas" : "-", sub: result.nextMonths[2] || "Sin mes"},
    {label:"Grupos detallados", value: formatNumber(rows.length), sub: `${formatNumber(rows.length)} combinaciones con modelo individual`},
  ]);

  renderSimpleTable("forecastSummaryTable", aggSummaryRows);
  renderSimpleTable("forecastModelsTable", modelRows);
  renderForecastModelChart(modelRows);

  const head = document.querySelector("#forecastTable thead");
  const body = document.querySelector("#forecastTable tbody");
  if(!rows.length){
    head.innerHTML = "";
    body.innerHTML = `<tr><td>No hay datos suficientes para el pronóstico con los filtros actuales.</td></tr>`;
    renderForecastChart([]);
    renderForecastHistoryChart([], []);
    renderForecastModelChart([]);
    return;
  }

  // Preparamos columnas para la tabla de la UI
  const next = result.nextMonths;
  let finalTableRows = [];

  if(level === "overall"){
    finalTableRows = [{
      Grupo: "Total consolidado",
      Metodo: metric === "piezas" ? aggSummaryRows[0].metodo_piezas : aggSummaryRows[0].metodo_importe,
      [next[0]]: metric === "piezas" ? aggSummaryRows[0]["Piezas total"] : aggSummaryRows[0]["Monto total"],
      [next[1]]: metric === "piezas" ? aggSummaryRows[1]["Piezas total"] : aggSummaryRows[1]["Monto total"],
      [next[2]]: metric === "piezas" ? aggSummaryRows[2]["Piezas total"] : aggSummaryRows[2]["Monto total"],
    }];
  } else {
    // Para niveles detallados, usamos los rows de buildDetailedForecastRows
    finalTableRows = rows.map(r => {
      const row = {
        Linea: r.linea,
        Cliente: r.nombre_cliente,
        Hospital: r.hospital_homologado,
        Medicamento: r.medicamento_homologado_presentacion,
        Metodo: r.metodo
      };
      next.forEach(m => {
        row[m] = metric === "piezas" ? r[`piezas_${m}`] : r[`monto_total_${m}`];
      });
      return row;
    });
  }

  const columns = Object.keys(finalTableRows[0]);
  const visibleRows = finalTableRows.slice(0,300);
  head.innerHTML = `<tr>${columns.map(col => `<th>${col}</th>`).join("")}</tr>`;
  body.innerHTML = visibleRows.map((row, idx) => {
    const detailSource = rows[idx]; // Corresponde al row original de buildDetailedForecastRows
    const groupKey = detailSource ? `${detailSource.linea || ""}|${detailSource.nombre_cliente || ""}|${detailSource.hospital_homologado || ""}|${detailSource.medicamento_homologado_presentacion || ""}|${detailSource.laboratorio_homologado || ""}` : "";
    const isSelected = state.selectedForecastGroup === groupKey;

    return `
      <tr class="${isSelected ? 'selected-row' : ''}" data-group='${groupKey}' style="${isSelected ? 'background:var(--accent-soft); border-left:4px solid var(--accent)' : ''}">
        ${columns.map(col => {
          const val = row[col];
          if(col === "Metodo") {
             let badgeClass = "badge-accent";
             if(val.includes("Tendencia")) badgeClass = "badge-success";
             if(val.includes("ponderado")) badgeClass = "badge-warning";
             return `<td><span class="badge ${badgeClass}">${val}</span></td>`;
          }
          if(/^\d{4}-\d{2}$/.test(col) || col === "Monto total") {
             return metric === "piezas" ? `<td>${formatNumber(Number(val), 2)}</td>` : `<td>${formatCurrency(Number(val))}</td>`;
          }
          if(typeof val === "number") return `<td>${formatNumber(val, 2)}</td>`;
          return `<td>${val ?? ""}</td>`;
        }).join("")}
      </tr>
    `;
  }).join("");

  // Bind click para drill-down
  body.querySelectorAll("tr").forEach(tr => {
    tr.addEventListener("click", () => {
      const g = tr.getAttribute("data-group");
      if(state.selectedForecastGroup === g) state.selectedForecastGroup = null;
      else state.selectedForecastGroup = g;
      renderForecast(data); // Re-render local
    });
  });

  renderForecastChart(aggSummaryRows);
  renderForecastHistoryChart(data, aggSummaryRows);
  renderGrowthInsights(data);
}

function renderGrowthInsights(data){
  const panel = document.getElementById("growthInsightsPanel");
  const wrap = document.getElementById("growthCards");
  const metric = document.getElementById("forecastMetric").value;
  const isPiezas = metric === "piezas";

  const result = buildDetailedForecastRows(data);
  const nextMonth = result.nextMonths[0];
  if(!nextMonth || !result.rows.length){
    panel.style.display = "none";
    return;
  }

  // OPTIMIZACIÓN CRÍTICA: Pre-agrupamos el histórico por llave para evitar filter() dentro de map()
  // Esto reduce la complejidad de O(N*M) a O(N+M), vital para 70k+ registros.
  const historyMap = new Map();
  data.forEach(d => {
    if(!d.anio_mes) return;
    const key = `${d.linea || ""}|${d.nombre_cliente || ""}|${d.hospital_homologado || ""}|${d.medicamento_homologado_presentacion || ""}|${d.laboratorio_homologado || ""}`;
    if(!historyMap.has(key)) historyMap.set(key, []);
    historyMap.get(key).push(d);
  });

  const growth = result.rows.map(row => {
    const key = `${row.linea || ""}|${row.nombre_cliente || ""}|${row.hospital_homologado || ""}|${row.medicamento_homologado_presentacion || ""}|${row.laboratorio_homologado || ""}`;
    const fcVal = Number(row[isPiezas ? `piezas_${nextMonth}` : `monto_total_${nextMonth}`]) || 0;
    
    const history = historyMap.get(key) || [];
    if(!history.length) return null;
    
    // Promedio de los últimos 3 meses como base
    const months = [...new Set(history.map(h => h.anio_mes))].sort().slice(-3);
    const sumBase = history.filter(h => months.includes(h.anio_mes)).reduce((s, h) => s + (Number(h[isPiezas ? "piezas" : "monto_total"])||0), 0);
    const avgBase = sumBase / (months.length || 1);
    
    const pct = avgBase > 0 ? ((fcVal - avgBase) / avgBase) * 100 : 0;
    return {
      name: row.medicamento_homologado_presentacion || row.hospital_homologado,
      sub: row.nombre_cliente,
      value: fcVal,
      pct: pct
    };
  })
  .filter(g => g && g.value > 0 && g.pct > 5)
  .sort((a,b) => b.pct - a.pct)
  .slice(0, 5);

  if(!growth.length){
    panel.style.display = "none";
    return;
  }

  panel.style.display = "block";
  wrap.innerHTML = growth.map(g => `
    <div class="stat-card insight">
      <div class="label">${g.name}<br><small style="opacity:0.7">${g.sub}</small></div>
      <div class="value">${isPiezas ? formatNumber(g.value, 0) + ' pzas' : formatCurrency(g.value)}</div>
      <div class="trend-up">↑ ${formatNumber(g.pct, 1)}% vs promedio</div>
    </div>
  `).join("");
}

function updateHeroText(data){
  const months = [...new Set(data.map(x => x.anio_mes).filter(Boolean))].sort();
  if(!months.length) return;
  const start = months[0];
  const end = months[months.length - 1];
  
  const formatDate = (ym) => {
    const [y, m] = ym.split("-");
    const mName = monthOrder[Number(m) - 1] || m;
    return `${mName} ${y}`;
  };

  const heroH2 = document.querySelector(".hero h2");
  if(heroH2) {
    if(start === end) heroH2.textContent = `Analizando datos de ${formatDate(start)}`;
    else heroH2.textContent = `Histórico consolidado: ${formatDate(start)} - ${formatDate(end)}`;
  }
}

function refresh(){
  const loader = document.getElementById("loader");
  const isLarge = (state.baseData.length + state.uploadedData.length) > 30000;
  
  if(isLarge && loader) {
    loader.style.display = "flex";
    loader.classList.remove("fade-out");
  }

  // PASO 1: Renderizamos Dashboard Principal (Instantáneo)
  setTimeout(() => {
    const allData = getAllData();
    renderSummary(allData);
    renderFilters(allData);
    const filtered = getFilteredData();
    state.filteredData = filtered;
    renderTable(filtered);
    renderCharts(filtered);
    updateHeroText(allData);

    // Ocultamos loader principal apenas el dashboard sea interactivo
    if(loader) {
      loader.classList.add("fade-out");
      setTimeout(() => { loader.style.display = "none"; }, 500);
    }
    
    // PASO 2: Cálculos pesados diferidos (Pronósticos e Insights)
    // Mostramos un mini-loader en la sección de pronóstico
    const fInfo = document.getElementById("forecastInfo");
    if(fInfo) fInfo.innerHTML = `<div style="display:flex; align-items:center; gap:10px;"><div class="spinner" style="width:20px; height:20px; border-width:2px;"></div><span>Calculando pronósticos y analizando tendencias...</span></div>`;

    setTimeout(() => {
      renderForecast(filtered);
      document.getElementById("piecesNote").textContent = (state.packagingCatalog.length ? "Catálogo empaque ✔" : "Sin empaque") + " | " + (state.priceList.length ? "Lista de precios ✔" : "Sin precios");
    }, 100);
  }, isLarge ? 30 : 0);
}

function saveLocal(){
  try{
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.uploadedData));
    localStorage.setItem(STORAGE_PACKAGING, JSON.stringify(state.packagingCatalog));
    localStorage.setItem(STORAGE_PRICES, JSON.stringify(state.priceList));
  }catch(e){
    if(e.name === "QuotaExceededError"){
      console.warn("Límite de almacenamiento excedido. Las cargas actuales son demasiado grandes para guardarse en el navegador.");
      showToast("Límite de memoria", "Los archivos son muy grandes para guardarse permanentemente en el navegador. Procesa y descarga tu Excel antes de cerrar la pestaña.", "warning");
    } else {
      console.error("Error al guardar en localStorage:", e);
    }
  }
}

function loadLocal(){
  try{
    state.uploadedData = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]").map(item => finalizeRecord({...item}));
    state.packagingCatalog = JSON.parse(localStorage.getItem(STORAGE_PACKAGING) || "[]");
    state.priceList = JSON.parse(localStorage.getItem(STORAGE_PRICES) || "[]");
  }catch(e){
    state.uploadedData = [];
    state.packagingCatalog = [];
    state.priceList = [];
  }
}

async function readWorkbook(file){
  const data = await file.arrayBuffer();
  return XLSX.read(data, {type:"array"});
}

async function processMonthlyFiles(files){
  const added = [];
  let totalRowsProcessed = 0;
  setStatus(`Iniciando lectura de ${files.length} archivos...`);
  
  for(const file of files){
    const workbook = await readWorkbook(file);
    const detected = detectSheet(workbook);
    if(!detected) {
      console.warn(`No identifiqué hoja válida en ${file.name}.`);
      continue;
    }
    
    setStatus(`Procesando ${file.name} (${detected.linea})...`);
    const rows = sheetToJsonRows(workbook, detected.sheet);
    const CHUNK_SIZE = 1000;
    for(let i=0; i<rows.length; i+=CHUNK_SIZE){
      const chunk = rows.slice(i, i + CHUNK_SIZE);
      for(const row of chunk){
        // Normalización única por fila: O(N) en lugar de O(N * C)
        const norm = { _fullPath: file.webkitRelativePath || "" };
        Object.keys(row).forEach(k => { norm[normalizeText(k)] = row[k]; });
        
        const record = buildRecord(norm, detected.linea, file.name);
        if(record) added.push(finalizeRecord(record));
      }
      totalRowsProcessed += chunk.length;
      setStatus(`Procesando ${file.name}: ${totalRowsProcessed.toLocaleString()} registros totales...`);
      await new Promise(resolve => setTimeout(resolve, 10));
    }
  }

  setStatus("Finalizando deduplicación de registros...");
  await new Promise(resolve => setTimeout(resolve, 50));

  // Deduplicación de alto rendimiento: usamos Set con llaves más cortas
  const existingKeys = new Set();
  const getLightKey = (r) => `${r.linea[0]}|${r.hospital_homologado.slice(0,5)}|${r.medicamento_homologado.slice(0,8)}|${r.fecha_entrega}|${r.monto_total}`;
  
  // Precargamos llaves existentes
  state.uploadedData.forEach(r => existingKeys.add(getLightKey(r)));

  const deduped = [];
  for(let i=0; i<added.length; i++){
    const row = added[i];
    const key = getLightKey(row);
    if(!existingKeys.has(key)){
      existingKeys.add(key);
      deduped.push(row);
    }
    // No saturar el hilo en bases masivas (70k+)
    if(i % 5000 === 0) await new Promise(r => setTimeout(r, 0));
  }

  state.uploadedData = [...state.uploadedData, ...deduped];
  return deduped.length;
}

async function processPriceFile(file){
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {defval:null, raw:true});
  
  state.priceList = rows.map(row => {
    let medicamento = "";
    let precio = null;
    
    for(const key of Object.keys(row)){
      const k = normalizeText(key);
      if(/DESCRIPCIO|MEDICAMENTO|CONCEPTO|PRODUCTO|ITEM/i.test(k)) medicamento = row[key];
      if(/PRECIO|COSTO|UNITARIO|VALOR/i.test(k)) precio = parseNumber(row[key]);
    }
    
    return {
      medicamento: normalizeText(medicamento),
      precio: precio
    };
  }).filter(item => item.medicamento && item.precio !== null);
  
  // Forzamos que se recalculen los pronósticos al haber nuevos precios
  saveLocal();
  refresh();
  
  return state.priceList.length;
}

function findPrice(medicamentoHomologado){
  if(!state.priceList.length) return null;
  const target = normalizeText(medicamentoHomologado);
  // Intento de match exacto
  const exact = state.priceList.find(p => p.medicamento === target);
  if(exact) return exact.precio;
  
  // Intento de match por inclusión (el de la lista de precios suele ser más largo)
  const partial = state.priceList.find(p => target.includes(p.medicamento) || p.medicamento.includes(target));
  if(partial) return partial.precio;
  
  return null;
}

async function processPackagingFile(file){
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {defval:null, raw:true});
  state.packagingCatalog = buildPackagingCatalog(rows);
  return state.packagingCatalog.length;
}

function exportCsv(rows, filename){
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8;"});
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

function sanitizeRowsForExport(rows){
  // Optimización masiva: evitamos Object.fromEntries y Object.entries en cada una de las 70k filas.
  const result = [];
  for(let i=0; i<rows.length; i++){
    const row = rows[i];
    const newRow = {};
    for(let key in row){
      let val = row[key];
      if(val === undefined || val === null) { newRow[key] = ""; continue; }
      if(typeof val === "number") {
        if(!Number.isFinite(val)) { newRow[key] = ""; continue; }
        newRow[key] = val; continue;
      }
      if(typeof val === "boolean") { newRow[key] = val ? "Sí" : "No"; continue; }
      newRow[key] = val;
    }
    result.push(newRow);
  }
  return result;
}

function downloadBlob(blob, filename){
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 3000);
}

function getExportColumns(){
  return [
    ["archivo_origen", 28],["linea", 16],["hospital_original", 30],["hospital_homologado", 30],
    ["medicamento_original", 28],["medicamento_homologado", 28],["medicamento_homologado_presentacion", 34],
    ["marca", 22],["fabricante_original", 24],["laboratorio_homologado", 24],["dosis", 14],
    ["presentacion_original", 20],["presentacion_valor", 16],["frascos", 16],["tipo_medicamento", 18],["mes", 12],
    ["monto_total", 16],["no_cliente", 14],["nombre_cliente", 30],["fecha_entrega", 14],["dia", 10],
    ["consigna", 12],["piezas_por_empaque", 16],["empaque_ambiguo", 14],["piezas", 14],
    ["producto_cliente_laboratorio", 44],["anio_mes", 12]
  ];
}

function styleHeaderRow(row){
  row.eachCell(cell => {
    cell.fill = {type:"pattern", pattern:"solid", fgColor:{argb:"A9D0F5"}};
    cell.font = {bold:true, color:{argb:"000000"}, name:"Calibri", size:11};
    cell.alignment = {vertical:"middle", horizontal:"center", wrapText:true};
    cell.border = {
      top:{style:"thin", color:{argb:"7F7F7F"}}, left:{style:"thin", color:{argb:"7F7F7F"}},
      bottom:{style:"thin", color:{argb:"7F7F7F"}}, right:{style:"thin", color:{argb:"7F7F7F"}}
    };
  });
  row.height = 26;
}

function styleDataSheet(ws){
  ws.views = [{state:"frozen", ySplit:1, showGridLines:false}];
  // Solo aplicamos bordes y estilos a las primeras N filas para evitar bloqueos de memoria en bases masivas
  const MAX_STYLED_ROWS = 2000; 
  ws.eachRow((row, rowNumber) => {
    if(rowNumber === 1) return;
    if(rowNumber > MAX_STYLED_ROWS) return; // Optimizacion crítica para 70k+ filas
    row.eachCell(cell => {
      cell.font = {name:"Calibri", size:10, color:{argb:"000000"}};
      cell.alignment = {vertical:"top", horizontal:"left", wrapText:false};
      cell.border = {
        top:{style:"thin", color:{argb:"D9D9D9"}}, left:{style:"thin", color:{argb:"D9D9D9"}},
        bottom:{style:"thin", color:{argb:"D9D9D9"}}, right:{style:"thin", color:{argb:"D9D9D9"}}
      };
    });
  });
  if(ws.rowCount > 1) {
    ws.autoFilter = {from:{row:1,column:1}, to:{row:1,column:ws.columnCount}};
  }
}

function applyColumnFormats(ws, columns){
  const integerCols = new Set(["piezas_por_empaque","dia"]);
  const currencyCols = new Set(["monto_total"]);
  const decimalCols = new Set(["dosis","presentacion_valor","frascos"]);
  columns.forEach((colName, idx) => {
    const column = ws.getColumn(idx + 1);
    if(integerCols.has(colName)) column.numFmt = '0';
    if(currencyCols.has(colName)) column.numFmt = '$#,##0.00';
    if(decimalCols.has(colName)) column.numFmt = '0.0000';
  });
}

function addDataSheet(workbook, name, rows){
  const columns = getExportColumns();
  const ws = workbook.addWorksheet(name);
  ws.columns = columns.map(([key, width]) => ({header:key, key, width}));
  
  // Procesamos en un solo paso para evitar duplicar la memoria
  const rowsToAdd = rows.map(row => {
    const rowData = {};
    for(let i=0; i<columns.length; i++){
      const key = columns[i][0];
      let val = row[key];
      // Sanitation in-place
      if(val === undefined || val === null) { rowData[key] = ""; }
      else if(typeof val === "number") { rowData[key] = Number.isFinite(val) ? val : ""; }
      else if(typeof val === "boolean") { rowData[key] = val ? "Sí" : "No"; }
      else { rowData[key] = val; }
    }
    return rowData;
  });
  
  ws.addRows(rowsToAdd);
  
  styleHeaderRow(ws.getRow(1));
  styleDataSheet(ws);
  applyColumnFormats(ws, columns.map(([key]) => key));
  return ws;
}

function addSummarySheet(workbook, rows){
  const ws = workbook.addWorksheet("Resumen");
  ws.views = [{showGridLines:false}];
  [30,16,4,14,16,14,4,18,18,16].forEach((w,idx) => ws.getColumn(idx+1).width = w);

  // Styling Title
  ws.mergeCells('A1:B1');
  const titleCell = ws.getCell('A1');
  titleCell.value = 'Tablero Maestro de Mezclas';
  titleCell.font = { name:'Outfit', family:4, size:18, bold:true, color: { argb: 'FFFFFF' } };
  titleCell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'1e293b' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  const totalMonto = rows.reduce((acc, row) => acc + (Number(row.monto_total) || 0), 0);
  const totalPiezas = rows.reduce((acc, row) => acc + (Number(row.piezas) || 0), 0);
  const linked = rows.filter(row => Number(row.piezas_por_empaque) > 1).length;

  const leftBlock = [
    ["Indicador", "Valor"],
    ["Registros totales", rows.length],
    ["Inversión histórica total", totalMonto],
    ["Frascos totales", rows.reduce((a,b)=>a+(Number(b.frascos)||0),0)],
    ["Piezas totales", totalPiezas],
    ["Hospitales únicos", new Set(rows.map(r => r.hospital_homologado).filter(Boolean)).size],
    ["Medicamentos únicos", new Set(rows.map(r => r.medicamento_homologado).filter(Boolean)).size],
    ["Laboratorios únicos", new Set(rows.map(r => r.laboratorio_homologado).filter(Boolean)).size],
    ["Enlazados con empaque", linked]
  ];

  leftBlock.forEach((r, i) => {
    const rowIdx = i + 3;
    ws.getCell(`A${rowIdx}`).value = r[0];
    ws.getCell(`B${rowIdx}`).value = r[1];
    ws.getCell(`A${rowIdx}`).font = { bold: i === 0 };
    if(i === 0) {
      ws.getCell(`A${rowIdx}`).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'A9D0F5' } };
      ws.getCell(`B${rowIdx}`).fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'A9D0F5' } };
    }
  });

  // Table Headers Monthly
  const headers = [['D','Mes'],['E','Monto total'],['F','Piezas total'],['H','Mes sin consigna'],['I','Monto sin consigna'],['J','Piezas sin consigna']];
  headers.forEach(h => {
    const cell = ws.getCell(`${h[0]}1`);
    cell.value = h[1];
    cell.font = { bold:true, color:{ argb:'FFFFFF' } };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'38bdf8' } };
    cell.alignment = { horizontal:'center' };
  });

  const grouped = new Map();
  const groupedNo = new Map();
  rows.forEach(row => {
    const key = row.anio_mes || row.mes || "";
    if(!grouped.has(key)) grouped.set(key, {monto:0, piezas:0});
    const a = grouped.get(key);
    a.monto += Number(row.monto_total) || 0;
    a.piezas += Number(row.piezas) || 0;
    if(!row.consigna){
      if(!groupedNo.has(key)) groupedNo.set(key, {monto:0, piezas:0});
      const b = groupedNo.get(key);
      b.monto += Number(row.monto_total) || 0;
      b.piezas += Number(row.piezas) || 0;
    }
  });
  const keys = [...grouped.keys()].sort();
  keys.forEach((key, i) => {
    const row = i + 2;
    ws.getCell(`D${row}`).value = key;
    ws.getCell(`E${row}`).value = grouped.get(key).monto;
    ws.getCell(`F${row}`).value = Number((grouped.get(key).piezas || 0).toFixed(4));
    ws.getCell(`H${row}`).value = key;
    const no = groupedNo.get(key) || {monto:0,piezas:0};
    ws.getCell(`I${row}`).value = no.monto;
    ws.getCell(`J${row}`).value = Number((no.piezas || 0).toFixed(4));
  });

  ws.getColumn(2).numFmt = '#,##0';
  ws.getCell('B5').numFmt = '$#,##0.00';
  ws.getColumn(5).numFmt = '$#,##0.00';
  ws.getColumn(6).numFmt = '#,##0.00';
  ws.getColumn(9).numFmt = '$#,##0.00';
  ws.getColumn(10).numFmt = '#,##0.00';

  styleDataSheet(ws);
}

function addForecastSheets(workbook, rows, precalculatedDetail){
  const detail = precalculatedDetail || buildDetailedForecastRows(rows);
  const detailRows = detail.rows;
  const summaryRows = buildAggregateForecastSummaryRows(rows, detail.nextMonths);

  const detailWs = workbook.addWorksheet("Pronostico_Detalle");
  if(!detailRows.length){
    detailWs.addRow(["No hay datos suficientes para el pronóstico con la selección actual."]);
    detailWs.getColumn(1).width = 70;
  } else {
    const columns = Object.keys(detailRows[0]);
    detailWs.columns = columns.map(col => ({header:col, key:col, width: ["nombre_cliente","hospital_homologado","medicamento_homologado_presentacion"].includes(col) ? 34 : ["historial_meses","base_importe","base_piezas"].includes(col) ? 28 : ["columna_origen_importe","columna_origen_piezas"].includes(col) ? 22 : 18}));
    detailRows.forEach(row => detailWs.addRow(row));
    styleHeaderRow(detailWs.getRow(1));
    styleDataSheet(detailWs);
    columns.forEach((col, idx) => {
      if(/^monto_/.test(col)) detailWs.getColumn(idx + 1).numFmt = '$#,##0.00';
      if(/^piezas/.test(col)) detailWs.getColumn(idx + 1).numFmt = '0.0000';
    });
  }

  const modelsWs = workbook.addWorksheet("Pronostico_Modelos");
  modelsWs.columns = [
    {header:"Modelo", key:"Modelo", width:28},
    {header:"Descripcion", key:"Descripcion", width:44},
    {header:"Usa", key:"Usa", width:42},
    {header:"Grupos", key:"Grupos", width:14},
    {header:"Participacion %", key:"Participacion", width:18},
  ];
  buildForecastModelRows(detailRows).forEach(row => modelsWs.addRow(row));
  styleHeaderRow(modelsWs.getRow(1));
  styleDataSheet(modelsWs);
  modelsWs.getColumn(4).numFmt = '0';
  modelsWs.getColumn(5).numFmt = '0.00';

  const helpWs = workbook.addWorksheet("Pronostico_Ayuda");
  helpWs.views = [{showGridLines:false}];
  helpWs.getColumn(1).width = 26;
  helpWs.getColumn(2).width = 90;
  helpWs.getCell('A1').value = 'Elemento';
  helpWs.getCell('B1').value = 'Explicación';
  styleHeaderRow(helpWs.getRow(1));
  [
    ['Base importe', 'Se pronostica con el histórico mensual de monto_total. No usa precio unitario.'],
    ['Base piezas', 'Se pronostica con el histórico mensual de piezas sin redondear (fuente: columna piezas).'],
    ['Columna origen importe', 'La columna fuente usada en la base para el pronóstico de importe es monto_total.'],
    ['Columna origen piezas', 'La columna fuente usada en la base para el pronóstico de piezas es la columna piezas.'],
    ['Tendencia promedio', 'Calcula el cambio promedio observado entre meses y lo proyecta hacia adelante.'],
    ['Promedio móvil ponderado', 'Da mayor peso a los meses más recientes para capturar cambios recientes.'],
    ['Promedio móvil simple', 'Promedia los últimos meses con el mismo peso para suavizar variaciones.'],
    ['Último valor', 'Replica el último valor observado como escenario base.'],
    ['Selección del modelo', 'Se elige el modelo con menor error reciente (%MAPE) cuando hay histórico suficiente.']
  ].forEach(item => helpWs.addRow(item));
  styleDataSheet(helpWs);

  const monthlySumWs = workbook.addWorksheet("Resumen_Mensual_Consolidado");
  monthlySumWs.columns = [
    {header:"Mes", key:"Mes", width:16},
    {header:"Monto total ($)", key:"Monto total", width:22},
    {header:"Piezas total (V)", key:"Piezas total", width:18},
    {header:"Monto sin consigna ($)", key:"Monto sin consigna", width:24},
    {header:"Piezas sin consigna (V)", key:"Piezas sin consigna", width:20},
  ];
  summaryRows.forEach(row => monthlySumWs.addRow(row));
  styleHeaderRow(monthlySumWs.getRow(1));
  styleDataSheet(monthlySumWs);
  monthlySumWs.getColumn(2).numFmt = '$#,##0.00';
  monthlySumWs.getColumn(3).numFmt = '#,##0.00';
  monthlySumWs.getColumn(4).numFmt = '$#,##0.00';
  monthlySumWs.getColumn(5).numFmt = '#,##0.00';
}


function buildForecastModelRows(detailRows){
  const counts = detailRows.reduce((acc, row) => {
    const key = row.metodo || "Sin método";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
  const descriptions = {
    "Tendencia promedio": "Proyecta con el cambio promedio observado en la serie histórica.",
    "Promedio móvil ponderado": "Da mayor peso a los meses más recientes para captar cambios recientes.",
    "Promedio móvil simple": "Promedia los últimos meses con el mismo peso.",
    "Último valor": "Replica el último valor observado como escenario base.",
    "Sin método": "No hubo histórico suficiente para seleccionar un modelo."
  };
  const uses = {
    "Tendencia promedio": "Usa la serie histórica mensual y el cambio promedio entre meses.",
    "Promedio móvil ponderado": "Usa los últimos meses y da más peso a los más recientes.",
    "Promedio móvil simple": "Usa el promedio de los últimos meses con el mismo peso.",
    "Último valor": "Usa el último valor mensual observado como base directa.",
    "Sin método": "Usa respaldo por histórico insuficiente."
  };
  return Object.entries(descriptions)
    .map(([modelo, descripcion]) => ({
      Modelo: modelo,
      Descripcion: descripcion,
      Usa: uses[modelo] || "",
      Grupos: counts[modelo] || 0,
      Participacion: detailRows.length ? Number((((counts[modelo] || 0) / detailRows.length) * 100).toFixed(2)) : 0
    }))
    .filter(row => row.Modelo !== 'Sin método' || row.Grupos > 0);
}

function addForecastChartsSheet(workbook, summaryRows){
  const ws = workbook.addWorksheet('Pronostico_Graficos');
  ws.views = [{showGridLines:false}];
  ws.getColumn(1).width = 4;
  ws.getColumn(2).width = 18;
  ws.getColumn(3).width = 18;
  ws.getColumn(4).width = 18;
  ws.getColumn(5).width = 18;
  ws.getColumn(6).width = 18;
  ws.getColumn(7).width = 18;
  ws.getCell('B2').value = 'Gráficos de pronóstico';
  ws.getCell('B2').font = {name:'Calibri', size:16, bold:true, color:{argb:'1F1F1F'}};
  ws.getCell('B3').value = 'Se incluyen los gráficos visibles de la herramienta, una tabla resumen de los próximos 3 meses y un gráfico extra de participación de modelos.';
  ws.getCell('B3').font = {name:'Calibri', size:10, color:{argb:'404040'}};

  const tableStartRow = 6;
  const headers = ['Mes','Monto total','Piezas total','Monto sin consigna','Piezas sin consigna'];
  headers.forEach((header, idx) => {
    const cell = ws.getCell(tableStartRow, idx + 2);
    cell.value = header;
    cell.fill = {type:'pattern', pattern:'solid', fgColor:{argb:'A9D0F5'}};
    cell.font = {name:'Calibri', size:11, bold:true, color:{argb:'000000'}};
    cell.alignment = {vertical:'middle', horizontal:'center'};
    cell.border = {
      top:{style:'thin', color:{argb:'7F7F7F'}}, left:{style:'thin', color:{argb:'7F7F7F'}},
      bottom:{style:'thin', color:{argb:'7F7F7F'}}, right:{style:'thin', color:{argb:'7F7F7F'}}
    };
  });
  summaryRows.forEach((row, idx) => {
    const values = [row['Mes'], row['Monto total'], row['Piezas total'], row['Monto sin consigna'], row['Piezas sin consigna']];
    values.forEach((value, j) => {
      const cell = ws.getCell(tableStartRow + idx + 1, j + 2);
      cell.value = value;
      cell.font = {name:'Calibri', size:10, color:{argb:'000000'}};
      cell.border = {
        top:{style:'thin', color:{argb:'D9D9D9'}}, left:{style:'thin', color:{argb:'D9D9D9'}},
        bottom:{style:'thin', color:{argb:'D9D9D9'}}, right:{style:'thin', color:{argb:'D9D9D9'}}
      };
    });
  });
  ws.getColumn(3).numFmt = '$#,##0.00';
  ws.getColumn(4).numFmt = '0';
  ws.getColumn(5).numFmt = '$#,##0.00';
  ws.getColumn(6).numFmt = '0.0000';

  const chartTargets = [
    {id:'forecastChart', title:'Pronóstico próximos 3 meses (importe: monto_total)', range:'B12:G30'},
    {id:'forecastHistoryChart', title:'Histórico vs pronóstico de piezas (fuente: columna piezas)', range:'B32:G50'},
    {id:'forecastModelChart', title:'Participación de modelos', range:'N12:S30'},
    {id:'monthlyChart', title:'Evolución mensual', range:'H32:M50'},
    {id:'labChart', title:'Top laboratorios', range:'N32:S50'}
  ];

  chartTargets.forEach(target => {
    const canvas = document.getElementById(target.id);
    if(!canvas || !canvas.toDataURL) return;
    try{
      const base64 = canvas.toDataURL('image/png', 1.0);
      const imageId = workbook.addImage({base64, extension:'png'});
      const startCell = target.range.split(':')[0];
      ws.getCell(startCell).value = target.title;
      ws.getCell(startCell).font = {name:'Calibri', size:12, bold:true, color:{argb:'1F1F1F'}};
      ws.addImage(imageId, target.range);
    }catch(error){
      console.warn('No se pudo agregar gráfico al Excel:', target.id, error);
    }
  });
}

async function exportStyledWorkbook(rows, filename){
  return exportXlsx(rows, filename);
}

function buildCatalogExportRows(){
  return state.packagingCatalog.map(item => ({
    archivo_origen: "CATALOGO EMPAQUE", linea: item.linea, hospital_original: "", hospital_homologado: "",
    medicamento_original: item.medicamento_homologado, medicamento_homologado: item.medicamento_homologado,
    medicamento_homologado_presentacion: item.medicamento_homologado, marca: "", fabricante_original: "",
    laboratorio_homologado: "", dosis: "", presentacion_original: "", presentacion_valor: "", frascos: "", tipo_medicamento: "",
    mes: "", monto_total: "", no_cliente: "", nombre_cliente: "", fecha_entrega: "", dia: "", consigna: "",
    piezas_por_empaque: item.piezas_por_empaque, empaque_ambiguo: "", piezas: "", producto_cliente_laboratorio: "", anio_mes: ""
  }));
}

function buildPriceExportRows(){
  return state.priceList.map(p => ({
    Medicamento: p.medicamento,
    Clave: p.clave,
    "Precio Unitario": p.precio,
    Laboratorio: p.laboratorio || ""
  }));
}

function buildSummaryExportRows(rows){
  const totalRegistros = rows.length;
  const montoTotal = rows.reduce((sum, row) => sum + (parseNumber(row.monto_total) || 0), 0);
  const piezasTotal = rows.reduce((sum, row) => sum + (parseNumber(row.piezas) || 0), 0);
  const montoSinConsigna = rows.filter(row => !row.consigna).reduce((sum, row) => sum + (parseNumber(row.monto_total) || 0), 0);
  const piezasSinConsigna = rows.filter(row => !row.consigna).reduce((sum, row) => sum + (parseNumber(row.piezas) || 0), 0);
  const clientes = new Set(rows.map(row => String(row.nombre_cliente || '').trim()).filter(Boolean)).size;
  const hospitales = new Set(rows.map(row => String(row.hospital_homologado || '').trim()).filter(Boolean)).size;
  const laboratorios = new Set(rows.map(row => String(row.laboratorio_homologado || '').trim()).filter(Boolean)).size;
  return [
    {Indicador:'Registros exportados', Valor: totalRegistros},
    {Indicador:'Monto total', Valor: montoTotal},
    {Indicador:'Piezas total', Valor: Number(piezasTotal.toFixed(4))},
    {Indicador:'Monto sin consigna', Valor: montoSinConsigna},
    {Indicador:'Piezas sin consigna', Valor: Number(piezasSinConsigna.toFixed(4))},
    {Indicador:'Clientes únicos', Valor: clientes},
    {Indicador:'Hospitales homologados únicos', Valor: hospitales},
    {Indicador:'Laboratorios homologados únicos', Valor: laboratorios},
  ];
}

function makeSheetFromRows(rows, columns){
  const ordered = sanitizeRowsForExport(rows).map(row => {
    if(!columns || !columns.length) return row;
    const obj = {};
    columns.forEach(col => { obj[col] = row[col] ?? ''; });
    return obj;
  });
  return XLSX.utils.json_to_sheet(ordered);
}

async function exportXlsx(rows, filename){
  if(!rows || !rows.length) {
    showToast("Sin datos", "No hay registros para exportar.", "warning");
    return;
  }
  
  const btnAll = document.getElementById('btnDownloadAll');
  const prevAll = btnAll ? btnAll.disabled : false;
  
  try{
    if(btnAll) btnAll.disabled = true;
    setStatus('Generando Excel Multi-Pestaña (Modo Ultra-Ligero)...');
    await new Promise(r => setTimeout(r, 100));

    const wb = XLSX.utils.book_new();

    // 1. Resumen
    setStatus('Generando hoja de Resumen...');
    const sumData = buildSummaryExportRows(rows);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sumData), "Resumen");

    // 2. Pronóstico Mensual (Histórico 2025 + Pronóstico 2026)
    setStatus('Generando Pronóstico Mensual...');
    const detail = buildDetailedForecastRows(rows);
    const aggRows = buildAggregateForecastSummaryRows(rows, detail.nextMonths);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(aggRows), "Pronostico_Mensual");

    // 3. Pronóstico Detallado (Por producto)
    setStatus('Generando Detalle de Pronóstico...');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detail.rows), "Pronostico_Detalle");

    // 4. Catálogos
    if(state.packagingCatalog.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildCatalogExportRows()), "Catalogo_Empaque");
    }
    if(state.priceList.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildPriceExportRows()), "Catalogo_Precios");
    }

    // 5. Base Principal (La pesada)
    setStatus('Procesando Base Homologada (272k registros)...');
    await new Promise(r => setTimeout(r, 50));
    // Pasamos rows directo para ahorrar RAM masivamente
    const wsBase = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, wsBase, "Base_Homologada");

    setStatus('Escribiendo archivo final...');
    XLSX.writeFile(wb, filename);
    setStatus(`Descarga completada: ${filename}`);
    
  }catch(error){
    console.error("Error en exportación XLSX:", error);
    setStatus(`Error al exportar: ${error.message}`);
    alert(`El archivo es demasiado grande para Excel estándar. Intenta descargar el botón CSV (Ligero).`);
  }finally{
    if(btnAll) btnAll.disabled = prevAll;
  }
}

function downloadPackagingTemplate(){
  const rows = [
    {MEDICAMENTO:"CICLOFOSFAMIDA", "Pieza x Caja":5, TIPO:"ONCOLOGICO"},
    {MEDICAMENTO:"CARNITINA 200 mg/ml", "Pieza x Caja":5, TIPO:"NUTRICIONAL"}
  ];
  exportStyledWorkbook(rows.map(r => ({archivo_origen:"PLANTILLA", linea:r.TIPO, hospital_original:"", hospital_homologado:"", medicamento_original:r.MEDICAMENTO, medicamento_homologado:r.MEDICAMENTO, medicamento_homologado_presentacion:r.MEDICAMENTO, marca:"", fabricante_original:"", laboratorio_homologado:"", dosis:"", presentacion_original:"", presentacion_valor:"", frascos:"", tipo_medicamento:"", mes:"", monto_total:"", no_cliente:"", nombre_cliente:"", fecha_entrega:"", dia:"", consigna:"", piezas_por_empaque:r["Pieza x Caja"], empaque_ambiguo:"", piezas:"", producto_cliente_laboratorio:"", anio_mes:""})), "plantilla_piezas_por_empaque.xlsx");
}


function bindCollapsiblePanel(){
  const panel = document.getElementById("controlPanel");
  const trigger = document.getElementById("toggleControlPanel");
  const body = document.getElementById("controlPanelBody");
  if(!panel || !trigger || !body) return;
  trigger.addEventListener("click", () => {
    const willOpen = panel.classList.contains("collapsed");
    panel.classList.toggle("collapsed", !willOpen);
    body.hidden = !willOpen;
    trigger.setAttribute("aria-expanded", willOpen ? "true" : "false");
  });
}

function bindEvents(){
  ["filterLinea","filterMes","filterHospital","filterMedicamento","filterLaboratorio","filterCliente","filterConsigna","filterMarca","filterTipo","mainMetric","forecastMetric","forecastLevel"]
    .forEach(id => document.getElementById(id).addEventListener("change", refresh));
  document.getElementById("filterSearch").addEventListener("input", refresh);

  document.getElementById("btnProcess").addEventListener("click", async () => {
    const monthlyFiles = [...document.getElementById("monthlyFiles").files];
    const packagingFile = document.getElementById("packagingFile").files[0];
    const priceFile = document.getElementById("priceFile").files[0];
    
    if(!monthlyFiles.length && !packagingFile && !priceFile) {
      showToast("Sin archivos", "Primero selecciona los archivos que deseas cargar.", "error");
      return;
    }

    try{
      let messages = [];
      let successCount = 0;

      if(monthlyFiles.length){
        const count = await processMonthlyFiles(monthlyFiles);
        if(count > 0) {
          messages.push(`Se agregaron ${count.toLocaleString()} registros nuevos de los archivos Excel mensuales.`);
          successCount++;
        } else {
          messages.push("No se encontraron registros válidos o ya existían en los archivos mensuales seleccionados.");
        }
      }
      if(packagingFile){
        const countPack = await processPackagingFile(packagingFile);
        if(countPack > 0) {
          messages.push(`Se cargaron ${countPack.toLocaleString()} relaciones de piezas por empaque.`);
          successCount++;
        }
      }
      if(priceFile){
        const countPrice = await processPriceFile(priceFile);
        if(countPrice > 0) {
          messages.push(`Se cargaron ${countPrice.toLocaleString()} precios de la lista.`);
          successCount++;
        }
      }

      saveLocal();
      refresh();
      
      const statusMsg = messages.join("\n");
      setStatus(statusMsg);
      
      if(successCount > 0) {
        showToast("¡Procesamiento exitoso!", "Los datos se han integrado correctamente a la base maestra.", "success");
      } else {
        showToast("Sin cambios nuevos", "Los archivos no contenían datos nuevos o hubo un problema con las columnas.", "error");
      }
    }catch(error){
      console.error(error);
      setStatus(`Ocurrió un problema al procesar: ${error.message}`);
      showToast("Error de procesamiento", error.message, "error");
    }
  });

  document.getElementById("btnResetUploads").addEventListener("click", () => {
    state.uploadedData = [];
    state.packagingCatalog = [];
    state.priceList = [];
    saveLocal();
    document.getElementById("monthlyFiles").value = "";
    document.getElementById("packagingFile").value = "";
    document.getElementById("priceFile").value = "";
    refresh();
    setStatus("Se limpiaron las cargas nuevas, catálogo de empaque y lista de precios.");
  });

  document.getElementById("btnDownloadTemplate").addEventListener("click", downloadPackagingTemplate);
  document.getElementById("btnDownloadAll").addEventListener("click", async () => { 
    try{ 
      await exportXlsx(getAllData(), "base_maestra_homologada_pronostico_2025_2026.xlsx"); 
    }catch(e){
      console.error(e);
    } 
  });

  document.getElementById("btnExportCsv").addEventListener("click", () => {
    try {
      setStatus("Generando descarga CSV...");
      exportCsv(getAllData(), "base_homologada_consolidada.csv");
      setStatus("CSV generado correctamente.");
    } catch(e) {
      console.error(e);
      showToast("Error", "No se pudo generar el CSV: " + e.message, "error");
    }
  });
  document.getElementById("btnExportFilteredCsv").addEventListener("click", () => exportCsv(state.filteredData, "base_homologada_filtrada.csv"));
  document.getElementById("btnExportFilteredXlsx").addEventListener("click", async () => { 
    try{ 
      await exportXlsx(state.filteredData, "base_homologada_filtrada.xlsx"); 
    }catch(e){
      console.error(e);
      alert("Error al descargar filtrado: " + e.message);
    } 
  });
  
  document.getElementById("btnResetFilters").addEventListener("click", () => {
    ["filterLinea","filterMes","filterHospital","filterMedicamento","filterLaboratorio","filterCliente","filterConsigna","filterMarca","filterTipo"]
      .forEach(id => document.getElementById(id).value = "");
    document.getElementById("filterSearch").value = "";
    refresh();
  });
}

// Nombre del archivo fuente (debe estar en la misma carpeta que index.html)
const XLSB_SOURCE_FILE = 'CONCENTRADO ONCOLOGICO Y NUTRICIONAL 2025 2026 TOTAL.2 BASE A.xlsb';

async function loadXlsbData(bar, statusTxt){
  // Descargamos el .xlsb directamente desde el servidor local
  const encodedName = XLSB_SOURCE_FILE.split('').map(c => encodeURIComponent(c)).join('').replace(/%20/g, '%20');
  if(statusTxt) statusTxt.textContent = `Descargando base de datos (archivo ~42 MB)...`;
  if(bar) bar.style.width = '10%';

  let resp;
  try {
    // Intentamos primero con el nombre crudo (el servidor lo puede manejar)
    resp = await fetch(XLSB_SOURCE_FILE);
    if(!resp.ok) throw new Error(`HTTP ${resp.status}`);
  } catch(e) {
    throw new Error(`No se pudo descargar el archivo fuente: ${e.message}. Asegúrate de estar corriendo un servidor local (npx serve .).`);
  }

  if(statusTxt) statusTxt.textContent = 'Leyendo bytes del archivo...';
  const arrayBuffer = await resp.arrayBuffer();
  if(bar) bar.style.width = '35%';

  if(statusTxt) statusTxt.textContent = 'Parseando estructura Excel (XLSB)...';
  await new Promise(r => setTimeout(r, 30));

  // Leemos con SheetJS (ya incluido en el HTML via CDN)
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), {type: 'array', dense: true});
  if(bar) bar.style.width = '50%';

  const allRecords = [];
  const totalSheets = workbook.SheetNames.length;

  for(let si = 0; si < totalSheets; si++){
    const sheetName = workbook.SheetNames[si];
    const sheet = workbook.Sheets[sheetName];
    const headerIdx = detectHeaderRow(sheet);
    const rows = XLSX.utils.sheet_to_json(sheet, {defval: null, raw: true, range: headerIdx});

    if(statusTxt) statusTxt.textContent = `Homologando hoja "${sheetName}" (${rows.length.toLocaleString()} filas)...`;
    await new Promise(r => setTimeout(r, 20));

    // Línea por defecto según nombre de hoja
    const lineaDefault = normalizeText(sheetName).includes('NUTRI') ? 'NUTRICIONAL' : 'ONCOLOGICO';

    const CHUNK = 3000;
    for(let i = 0; i < rows.length; i += CHUNK){
      const slice = rows.slice(i, i + CHUNK);

      for(const row of slice){
        // Normalizar todas las claves del row para que buildRecord las encuentre
        const norm = {};
        for(const k in row) norm[normalizeText(k)] = row[k];

        // Detectar línea desde la columna LÍNEA/LINEA del propio registro
        const lineaRaw = norm['LINEA'] || '';
        let linea = lineaDefault;
        if(lineaRaw){
          const lv = normalizeText(String(lineaRaw));
          if(lv.includes('NUTRI')) linea = 'NUTRICIONAL';
          else if(lv.includes('ONCO')) linea = 'ONCOLOGICO';
        }

        const rec = buildRecord(norm, linea, XLSB_SOURCE_FILE);
        if(rec) allRecords.push(rec);
      }

      const progress = 50 + ((si / totalSheets) + (i / rows.length / totalSheets)) * 35;
      if(bar) bar.style.width = Math.min(85, progress) + '%';
      await new Promise(r => setTimeout(r, 0)); // Yield para no congelar UI
    }
  }

  return allRecords;
}

async function init(){
  const loader = document.getElementById('loader');
  const bar = document.getElementById('progressBar');
  const statusTxt = document.getElementById('loaderStatus');

  if(bar) bar.style.width = '5%';
  if(statusTxt) statusTxt.textContent = 'Iniciando tablero...';

  // Pintar el loader antes de cualquier operación pesada
  await new Promise(r => requestAnimationFrame(r));
  await new Promise(r => setTimeout(r, 50));

  try {
    // ─── Cargar datos desde el XLSB directamente ───────────────────────────
    state.baseData = await loadXlsbData(bar, statusTxt);

    // ─── Finalizar registros (calcular piezas, presentacion_valor, etc.) ────
    if(statusTxt) statusTxt.textContent = 'Optimizando registros...';
    const total = state.baseData.length;
    const CHUNK = 5000;
    for(let i = 0; i < total; i += CHUNK){
      const limit = Math.min(i + CHUNK, total);
      for(let j = i; j < limit; j++) finalizeRecord(state.baseData[j]);
      if(bar) bar.style.width = (85 + (i / total) * 12) + '%';
      if(statusTxt && i % 25000 === 0)
        statusTxt.textContent = `Finalizando: ${i.toLocaleString()} de ${total.toLocaleString()} registros...`;
      await new Promise(r => setTimeout(r, 0));
    }

    if(statusTxt) statusTxt.textContent = 'Cargando componentes del tablero...';

    loadLocal();
    bindCollapsiblePanel();
    bindEvents();

    setStatus([
      `Base cargada: ${state.baseData.length.toLocaleString('es-MX')} registros.`,
      'La herramienta aplica piezas por empaque con default 1 si no hay cruce.',
    ].join('\n'));

    refresh();

    if(bar) bar.style.width = '100%';
    setTimeout(() => {
      if(loader) loader.classList.add('fade-out');
      setTimeout(() => { if(loader) loader.style.display = 'none'; }, 500);
    }, 300);

  } catch(err) {
    console.error('Error durante la inicialización:', err);
    if(loader) loader.style.display = 'none';
    showToast('Error de carga', `Hubo un problema al inicializar: ${err.message}`, 'error');
    // Mostrar el error en el statusBox para que sea visible
    setStatus(`⚠ Error: ${err.message}`);
  }
}

window.addEventListener('DOMContentLoaded', init);
