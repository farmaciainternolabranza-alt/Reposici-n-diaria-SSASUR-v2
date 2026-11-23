(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const modeButton = $("drogueriaModeBtn");
  const modeConfig = $("drogueriaConfig");
  const modeResults = $("drogueriaResultsPanel");
  const generateButton = $("drogueriaGenerate");
  const stockInput = $("stockFile");
  const consumptionInput = $("consumoFile");
  const pharmacySelect = $("pharmacyWarehouse");
  const warehouseSelect = $("warehouseStock");
  const mainGeneratePanel = $("generatePanel");
  const mainResults = $("resultsPanel");
  let isDrogueria = false;
  let stockRecords = [];
  let stockDate = null;
  let productRows = [];

  const normalize = value => String(value ?? "").trim().toLocaleLowerCase("es-CL")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " ");
  const cleanCode = value => String(value ?? "").trim().replace(/\.0$/, "");
  const parseNumber = value => {
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;
    let text = String(value ?? "").trim().replace(/\s/g, "");
    if (!text) return 0;
    if (text.includes(",") && text.includes(".")) {
      text = text.lastIndexOf(",") > text.lastIndexOf(".")
        ? text.replace(/\./g, "").replace(",", ".") : text.replace(/,/g, "");
    } else if (text.includes(",")) text = text.replace(",", ".");
    const number = Number(text);
    return Number.isFinite(number) ? number : 0;
  };
  const parseDate = text => {
    const match = String(text ?? "").match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
    if (!match) return null;
    const date = new Date(+match[3], +match[2] - 1, +match[1]);
    return date.getFullYear() === +match[3] && date.getMonth() === +match[2] - 1 && date.getDate() === +match[1] ? date : null;
  };
  const formatDate = date => date ? date.toLocaleDateString("es-CL") : "No detectada";
  const selectedDocTypes = () => new Set([...( $("docOptions")?.querySelectorAll("input:checked") || [])].map(input => input.value));

  function csvRows(text) {
    const lines = text.split(/\r?\n/).filter(line => line.trim()).slice(0, 30);
    const delimiters = [";", ",", "\t"];
    const delimiter = delimiters.map(delim => ({ delim, score: lines.reduce((total, line) => total + line.split(delim).length - 1, 0) }))
      .sort((a, b) => b.score - a.score)[0].delim;
    const rows = [];
    let row = [], field = "", quoted = false;
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      if (char === '"') {
        if (quoted && text[i + 1] === '"') { field += '"'; i++; }
        else quoted = !quoted;
      } else if (char === delimiter && !quoted) { row.push(field); field = ""; }
      else if ((char === "\n" || char === "\r") && !quoted) {
        if (char === "\r" && text[i + 1] === "\n") i++;
        row.push(field);
        if (row.some(cell => String(cell).trim())) rows.push(row);
        row = []; field = "";
      } else field += char;
    }
    row.push(field);
    if (row.some(cell => String(cell).trim())) rows.push(row);
    return rows;
  }

  async function getRows(file, isConsumption) {
    if (!file) throw new Error("Carga los reportes de consumo y stock antes de generar el informe.");
    if (/\.xlsx?$/i.test(file.name)) {
      if (!window.XLSX) throw new Error("No está disponible el lector Excel. Recarga la página con conexión a internet.");
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false });
      return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, raw: true, defval: "" });
    }
    if (isConsumption) return csvRows((await file.text()).replace(/^\uFEFF/, ""));
    throw new Error("El archivo de stock debe estar en formato Excel (.xlsx o .xls).");
  }

  function locateStockHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const headers = rows[i].map(normalize);
      const code = headers.indexOf("codigo");
      const warehouse = headers.indexOf("bodega");
      const quantity = headers.indexOf("cantidad");
      const description = headers.findIndex(value => ["descripcion", "producto", "articulo"].includes(value));
      if (code >= 0 && warehouse >= 0 && quantity >= 0) return { row: i, code, warehouse, quantity, description };
    }
    throw new Error("No se reconocieron las columnas Código, Bodega y Cantidad del reporte de stock.");
  }

  async function loadStock() {
    const file = stockInput?.files?.[0];
    if (!file) return;
    try {
      const rows = await getRows(file, false);
      const header = locateStockHeader(rows);
      const meta = rows.slice(0, header.row).flat().join(" ");
      const dateMatch = meta.match(/obtenido\s*:\s*(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4})/i);
      stockDate = dateMatch ? parseDate(dateMatch[1]) : null;
      stockRecords = rows.slice(header.row + 1).map(row => ({
        code: cleanCode(row[header.code]),
        warehouse: String(row[header.warehouse] ?? "").trim(),
        quantity: parseNumber(row[header.quantity]),
        description: header.description >= 0 ? String(row[header.description] ?? "").trim() : ""
      })).filter(row => row.code && row.warehouse);
      updateModeAvailability();
    } catch (error) {
      stockRecords = [];
      updateModeAvailability(error.message);
    }
  }

  function updateModeAvailability(errorMessage = "") {
    const pharmacy = pharmacySelect?.value || "";
    const warehouse = warehouseSelect?.value || "";
    const hasFiles = Boolean(stockInput?.files?.length && consumptionInput?.files?.length);
    const validSelection = Boolean(pharmacy && warehouse && pharmacy !== warehouse);
    if (generateButton) generateButton.disabled = !validSelection || !stockRecords.length || !hasFiles;
    const info = $("drogueriaStockInfo");
    if (!info) return;
    if (errorMessage) {
      info.textContent = errorMessage;
      info.className = "notice error";
    } else if (!hasFiles) {
      info.textContent = "Carga los mismos archivos de consumo y stock en la sección 1.";
      info.className = "notice warning";
    } else if (!stockRecords.length) {
      info.textContent = "Leyendo el archivo de stock…";
      info.className = "notice warning";
    } else if (!validSelection) {
      info.textContent = "Verifica que estén seleccionadas dos bodegas distintas en la sección 3.";
      info.className = "notice warning";
    } else {
      info.textContent = `Stock CESFAM = ${pharmacy} + ${warehouse}.`;
      info.className = "notice success";
    }
  }

  function locateConsumptionHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 60); i++) {
      const headers = rows[i].map(normalize);
      const code = headers.indexOf("codigo");
      const quantity = headers.indexOf("cantidad");
      const type = headers.findIndex(value => ["tipo documento", "tipo de documento", "documento", "tipo doc"].includes(value));
      if (code >= 0 && quantity >= 0 && type >= 0) return { row: i, code, quantity, type };
    }
    const fallback = rows.slice(0, 60).findIndex(row => normalize(row[4]).includes("codigo") && normalize(row[7]).includes("cantidad") && /tipo|documento/.test(normalize(row[9])));
    if (fallback >= 0) return { row: fallback, code: 4, quantity: 7, type: 9 };
    throw new Error("No se encontraron las columnas Código, Cantidad y Tipo Documento en el reporte de consumo.");
  }

  async function readConsumption() {
    const rows = await getRows(consumptionInput?.files?.[0], true);
    const header = locateConsumptionHeader(rows);
    const meta = rows.slice(0, header.row + 1).flat().join(" ");
    const period = meta.match(/desde\s+(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4})\s+hasta\s+(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4})/i);
    const records = rows.slice(header.row + 1).map(row => ({
      code: cleanCode(row[header.code]),
      quantity: parseNumber(row[header.quantity]),
      type: String(row[header.type] ?? "").trim()
    })).filter(record => record.code && record.type);
    return { records, start: period ? parseDate(period[1]) : null, end: period ? parseDate(period[2]) : null };
  }

  function modeSwitch(enabled) {
    isDrogueria = enabled;
    document.body.classList.toggle("drogueria-active", enabled);
    modeButton?.setAttribute("aria-pressed", String(enabled));
    modeConfig?.classList.toggle("hidden", !enabled);
    mainGeneratePanel?.classList.toggle("hidden", enabled);
    mainResults?.classList.toggle("hidden", enabled);
    if (!enabled) modeResults?.classList.add("hidden");
    updateModeAvailability();
  }

  async function generateDrogueria() {
    try {
      const pharmacy = pharmacySelect?.value || "";
      const warehouse = warehouseSelect?.value || "";
      if (!pharmacy || !warehouse || pharmacy === warehouse) {
        throw new Error("Selecciona dos bodegas distintas en la sección 3: Stock Farmacia y Stock Bodega.");
      }
      const today = new Intl.DateTimeFormat("en-CA", {
        timeZone: "America/Santiago", year: "numeric", month: "2-digit", day: "2-digit"
      }).format(new Date());
      const stockDay = stockDate
        ? `${stockDate.getFullYear()}-${String(stockDate.getMonth() + 1).padStart(2, "0")}-${String(stockDate.getDate()).padStart(2, "0")}`
        : "";
      if (!stockDay || stockDay !== today) {
        throw new Error("El reporte de stock no fue obtenido hoy en Chile o no se pudo detectar su fecha. Actualiza el archivo de stock para continuar.");
      }

      const consumption = await readConsumption();
      const types = selectedDocTypes();
      if (!types.size) throw new Error("Selecciona al menos un tipo de documento para calcular el consumo.");

      const consumptionByCode = new Map();
      consumption.records.filter(record => types.has(record.type)).forEach(record => {
        consumptionByCode.set(record.code, (consumptionByCode.get(record.code) || 0) + record.quantity);
      });
      const stockByCode = new Map();
      const descriptions = new Map();
      stockRecords.forEach(record => {
        if (record.description && !descriptions.has(record.code)) descriptions.set(record.code, record.description);
        if (record.warehouse !== pharmacy && record.warehouse !== warehouse) return;
        stockByCode.set(record.code, (stockByCode.get(record.code) || 0) + record.quantity);
      });

      const codes = new Set([...consumptionByCode.keys(), ...stockByCode.keys()]);
      productRows = [...codes].map(code => {
        const consumptionValue = consumptionByCode.get(code) || 0;
        const cesfamStock = stockByCode.get(code) || 0;
        return {
          code,
          article: descriptions.get(code) || code,
          consumption: consumptionValue,
          stock: cesfamStock,
          availability: consumptionValue > 0 ? cesfamStock / consumptionValue * 100 : null
        };
      }).sort((a, b) => a.article.localeCompare(b.article, "es"));

      const tbody = $("drogueriaTable").tBodies[0];
      tbody.innerHTML = "";
      productRows.forEach(row => {
        const tr = document.createElement("tr");
        const values = [row.article, format(row.consumption), format(row.stock), row.availability === null ? "N/C" : `${row.availability.toFixed(1).replace(".", ",")}%`];
        values.forEach((value, index) => {
          const td = document.createElement("td");
          td.textContent = value;
          if (index === 0) td.title = row.article;
          if (index === 3 && row.availability !== null) {
            if (row.availability < 50) td.classList.add("availability-low");
            else if (row.availability > 50 && row.availability < 100) td.classList.add("availability-medium");
            else if (row.availability > 100) td.classList.add("availability-high");
          }
          tr.append(td);
        });
        tbody.append(tr);
      });

      $("drogueriaSubtitle").textContent = `Consumo: ${formatDate(consumption.start)} al ${formatDate(consumption.end)} · Stock: ${formatDate(stockDate)} · Stock CESFAM: ${pharmacy} + ${warehouse}`;
      modeResults?.classList.remove("hidden");
      modeResults?.scrollIntoView({ behavior: "smooth" });
    } catch (error) {
      alert(error.message || "No se pudo generar el reporte Droguería.");
    }
  }

  function format(value) {
    return Number(value || 0).toLocaleString("es-CL", { maximumFractionDigits: 2 });
  }

  function downloadCsv() {
    if (!productRows.length) return;
    const rows = [["Artículo", "Consumo", "Stock CESFAM", "Disponibilidad (%)"], ...productRows.map(row => [
      row.article, row.consumption, row.stock, row.availability === null ? "" : row.availability.toFixed(2)
    ])];
    const csv = "\uFEFF" + rows.map(row => row.map(value => `"${String(value).replace(/"/g, '""')}"`).join(";")).join("\r\n");
    const link = document.createElement("a");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    link.href = url;
    link.download = "reporte_drogueria.csv";
    link.click();
    URL.revokeObjectURL(url);
  }

  modeButton?.addEventListener("click", () => modeSwitch(!isDrogueria));
  stockInput?.addEventListener("change", () => { stockRecords = []; stockDate = null; updateModeAvailability(); loadStock(); });
  consumptionInput?.addEventListener("change", updateModeAvailability);
  pharmacySelect?.addEventListener("change", updateModeAvailability);
  warehouseSelect?.addEventListener("change", updateModeAvailability);
  $("docOptions")?.addEventListener("change", updateModeAvailability);
  generateButton?.addEventListener("click", generateDrogueria);
  $("drogueriaPrintBtn")?.addEventListener("click", () => window.print());
  $("drogueriaCsvBtn")?.addEventListener("click", downloadCsv);
  updateModeAvailability();
})();
