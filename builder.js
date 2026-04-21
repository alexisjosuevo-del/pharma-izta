const XLSX = require('xlsx');
const fs = require('fs');

console.log("Reading file...");
const wb = XLSX.readFile('CONCENTRADO ONCOLOGICO Y NUTRICIONAL 2025 2026 TOTAL.2 BASE A.xlsb', { dense: true });
console.log("File read. SheetNames:", wb.SheetNames);

let allValidRows = [];

function normalizeText(s) {
  if (s == null) return "";
  return String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[.,;:]+/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeHospital(v) { return normalizeText(v); }
function normalizeFabricante(v) { 
  let t = normalizeText(v).replace(/[\-_]+/g, " ").replace(/\s+/g, " ").trim();
  t = t.replace(/^LABORATORIOS?\s+/g, "").replace(/^LABS?\s+/g, "").replace(/^S\s*A\s*DE\s*C\s*V\b/g, "").replace(/^SA\s+DE\s+CV\b/g, "").replace(/^S\s*A\b/g, "").replace(/^DE\s+C\s*V\b/g, "").replace(/\s+/g, " ").trim();
  if(t==="ZURICH PHARMA" || t==="ZURICH") t="ZURICH";
  if(t.includes("FRENESUS") || t.includes("FRESENIUS")) t="FRESENIUS KABI";
  if(t.includes("ACCORD")) t="ACCORD";
  if(t.includes("JANSSEN")) t="JANSSEN";
  return t;
}
function normalizeMedicamentoOnco(v) {
  let t = normalizeText(v).replace(/\b\d+(?:[.,]\d+)?\s*(MG\/ML|MG|MCG|G|GRAMOS?|GR|ML|UI|IU|MEQ|MMOL|%|MUI|U|UNIDADES?)\b/g, " ").replace(/\b\d+(?:[.,]\d+)?(?:\/\d+(?:[.,]\d+)?)?\b/g, " ").replace(/\bSOLUCION\b/g, " ").replace(/\s+/g, " ").trim();
  if(/^TRASTUZUMAB EMTANSINA\b/.test(t)) return "TRASTUZUMAB";
  if(/^TRASTUZUMAB DERUXTECAN?\b/.test(t)) return "TRASTUZUMAB";
  return t;
}
function parseNumber(value){
  if(value === null || value === undefined || value === "") return null;
  if(typeof value === "number") return Number.isFinite(value) ? value : null;
  const normalized = String(value).replace(/,/g, ".").match(/-?\d+(?:\.\d+)?/);
  return normalized ? Number(normalized[0]) : null;
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

function detectHeaderRow(sheet){
  const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:null, raw:true, range:0});
  for(let idx = 0; idx < Math.min(rows.length, 20); idx++){
    const row = (rows[idx] || []).map(value => normalizeText(value || ""));
    const hasHospital = row.some(cell => cell.includes("HOSPITAL") || cell.includes("UNIDAD") || cell.includes("PUNTO DE VENTA") || cell.includes("CLIENTE"));
    const hasMedicamento = row.some(cell => cell.includes("MEDICAMENTO") || cell.includes("DESCRIPCION") || cell.includes("&&") || cell.includes("PRODUCTO"));
    if((hasHospital || row.includes("HOSPITAL") || row.includes("HOSPITAL ")) && hasMedicamento) return idx;
  }
  return 0;
}

wb.SheetNames.forEach(sheetName => {
  const sheet = wb.Sheets[sheetName];
  const headerIdx = detectHeaderRow(sheet);
  console.log("Sheet", sheetName, "header row:", headerIdx);
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, range: headerIdx });
  let linea = "ONCOLOGICO";
  if (sheetName.toUpperCase().includes("NUTRI") || normalizeText(sheetName).includes("NUTRI")) linea = "NUTRICIONAL";

  rows.forEach(row => {
    const norm = {};
    for (const k in row) norm[normalizeText(k)] = row[k];
    const coalesce = (keys) => {
      for (const k of keys) {
        if (norm[normalizeText(k)] !== undefined && norm[normalizeText(k)] !== null) return norm[normalizeText(k)];
      }
      return null;
    };

    let hospitalOriginal = coalesce(["HOSPITAL*", "HOSPITAL", "UNIDAD", "PUNTO DE VENTA"]);
    let medicamentoOriginal = coalesce(["MEDICAMENTO", "DESCRIPCION", "CONCEPTO", "&&"]);
    let marca = coalesce(["MARCA *", "MARCA"]);
    let fabricanteOriginal = coalesce(["FABRICANTE *", "FABRICANTE", "LABORATORIO"]);
    let dosis = coalesce(["DOSIS CON SOBRELLENADO (ML) *", "DOSIS (ML)", "DOSIS", "DOSIS CON SOBRELLENADO"]);
    let presentacionOriginal = coalesce(["PRESENTACIÓN", "PRESENTACION *", "PRESENTACION", "BM"]);
    let tipoMedicamento = coalesce(["TIPO MEDICAMENTO", "TIPO MEDICAMENTO 1", "TIPO", "LINEA", "LÍNEA"]);
    let montoTotal = coalesce(["TOTAL MEZCLAS", "TOTAL PRECIO MEZCLAS", "IMPORTE TOTAL", "TOTAL", "IMPORTE", "MONTO TOTAL"]);
    let nombreCliente = coalesce(["NOMBRE CLIENTE", "CLIENTE", "CLIENTE (NIVEL 1)"]);
    
    let fechaISO = null;
    let excelDate = coalesce(["FECHA DE ENTREGA*", "FECHA ENTREGA*", "FECHA DE ENTREGA", "FECHA ENTREGA", "MES", "FECHA", "AÑO MES"]);
    if (typeof excelDate === "number") {
      const utcValue = Math.floor(excelDate - 25569) * 86400;
      fechaISO = new Date(utcValue * 1000).toISOString().slice(0, 10);
    } else if (typeof excelDate === "string") {
      excelDate = excelDate.trim();
      let match = excelDate.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
      if(match) fechaISO = `${match[1]}-${String(match[2]).padStart(2,"0")}-${String(match[3]).padStart(2,"0")}`;
      else {
        match = excelDate.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
        if(match) {
           const y = Number(match[3]) < 100 ? `20${String(match[3]).padStart(2,"0")}` : match[3];
           fechaISO = `${y}-${String(match[2]).padStart(2,"0")}-${String(match[1]).padStart(2,"0")}`;
        } else {
             // maybe it's just year month?
             match = excelDate.match(/^(\d{4})[-\/](\d{1,2})$/);
             if(match) {
                 fechaISO = `${match[1]}-${String(match[2]).padStart(2,"0")}-01`;
             }
        }
      }
    }
    const anio_mes = fechaISO ? fechaISO.slice(0, 7) : "";
    const presVal = extractPresentationValue(presentacionOriginal);
    let frascos = parseNumber(coalesce(["FRASCOS CONSUMIDOS CON SOBRELLENADO", "FRASCOS CONSUMIDOS", "FRASCOS", "FRASCO", "NO FRASCOS", "CANTIDAD", "PIEZAS"]));
    
    if (!hospitalOriginal && !medicamentoOriginal) return;

    if (!linea && tipoMedicamento) {
        if(normalizeText(tipoMedicamento).includes("ONCO")) linea = "ONCOLOGICO";
        else linea = "NUTRICIONAL";
    }

    let medHomologado = linea === "ONCOLOGICO" ? normalizeMedicamentoOnco(medicamentoOriginal) : normalizeText(medicamentoOriginal);
    const piezas_por_empaque = 1;
    let piezas = null;
    if (frascos == null) {
      const d = parseNumber(dosis);
      if (d && presVal) frascos = d / presVal;
    }
    if (frascos != null) piezas = Number((frascos / piezas_por_empaque).toFixed(4));
    
    let presLbl = medHomologado;
    if (presVal) presLbl = medHomologado + " " + presVal;

    allValidRows.push({
      linea, archivo_origen: "CONCENTRADO 2025 2026 TOTAL.2 BASE A", hospital_original: hospitalOriginal || "", hospital_homologado: normalizeHospital(hospitalOriginal),
      medicamento_original: medicamentoOriginal || "", medicamento_homologado: medHomologado, medicamento_homologado_presentacion: presLbl,
      marca: normalizeText(marca), fabricante_original: fabricanteOriginal || "", laboratorio_homologado: normalizeFabricante(fabricanteOriginal),
      dosis: parseNumber(dosis), presentacion_original: presentacionOriginal || "", presentacion_valor: presVal,
      tipo_medicamento: normalizeText(tipoMedicamento), mes: "", monto_total: parseNumber(montoTotal), no_cliente: "", nombre_cliente: String(nombreCliente || ""),
      fecha_entrega: fechaISO, dia: fechaISO ? Number(fechaISO.slice(8,10)) : null, estatus: "", contrato: "", consigna: normalizeText(tipoMedicamento).includes("CONSIGNA"),
      frascos, piezas_por_empaque, piezas, anio_mes, producto_cliente_laboratorio: "", empaque_ambiguo: false
    });
  });
});

fs.writeFileSync("initial-data.js", "window.PRELOADED_DATA = " + JSON.stringify(allValidRows) + ";");
console.log("Done! Wrote format data to initial-data.js. Rows: " + allValidRows.length);
