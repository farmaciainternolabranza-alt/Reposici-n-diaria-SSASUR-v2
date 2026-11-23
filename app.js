(() => {
"use strict";
const $ = id => document.getElementById(id);
let consumoData=null, stockData=null, consumoFile=null, stockFile=null, resultRows=[];
const STORAGE_KEY="reposicion_fecha_consumo_v1";
function setStatus(msg,kind=""){const el=$("status");el.textContent=msg;el.className="status "+kind}
function norm(v){return String(v??"").trim().toLocaleLowerCase("es-CL").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ")}
function cleanCode(v){return String(v??"").trim().replace(/\.0$/,"")}
function num(v){if(typeof v==="number")return Number.isFinite(v)?v:0;let s=String(v??"").trim().replace(/\s/g,"");if(!s)return 0;if(s.includes(",")&&s.includes(".")){s=s.lastIndexOf(",")>s.lastIndexOf(".")?s.replace(/\./g,"").replace(",","."):s.replace(/,/g,"")}else if(s.includes(","))s=s.replace(",",".");const n=Number(s);return Number.isFinite(n)?n:0}
function parseDate(s){const m=String(s??"").match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);if(!m)return null;const d=new Date(+m[3],+m[2]-1,+m[1]);return d.getFullYear()===+m[3]&&d.getMonth()===+m[2]-1&&d.getDate()===+m[1]?d:null}
function dateKey(d){return d?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`:""}
function fmtDate(d){return d?d.toLocaleDateString("es-CL"):"No detectada"}
function todayChile(){return new Intl.DateTimeFormat("en-CA",{timeZone:"America/Santiago",year:"numeric",month:"2-digit",day:"2-digit"}).format(new Date())}
function findDate(text){const m=String(text).match(/obtenido\s*:\s*(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{4})/i);return m?parseDate(m[1]):null}
function detectDelimiter(text){
 const lines=text.split(/\r?\n/).filter(x=>x.trim()).slice(0,40),candidates=[";",",","\t"];
 // Puntúa separadores en varias filas, no solo en la primera línea del reporte,
 // que normalmente contiene el título y no los encabezados.
 return candidates.map(delim=>{let counts=lines.map(line=>{let q=false,n=0;for(let i=0;i<line.length;i++){if(line[i]==='"'){if(q&&line[i+1]==='"'){i++;}else q=!q}else if(line[i]===delim&&!q)n++}return n}).filter(n=>n>0);let total=counts.reduce((a,b)=>a+b,0),consistent=counts.length?counts.filter(n=>n===Math.max(...counts)).length:0;return {delim,total,consistent}}).sort((a,b)=>b.total-a.total||b.consistent-a.consistent)[0].delim;
}
function parseCSV(text){
 text=text.replace(/^\uFEFF/,"");
 const delim=detectDelimiter(text);
 const rows=[];let row=[],field="",quoted=false;
 for(let i=0;i<text.length;i++){const c=text[i];if(c==='"'){if(quoted&&text[i+1]==='"'){field+='"';i++}else quoted=!quoted}else if(c===delim&&!quoted){row.push(field);field=""}else if((c==="\n"||c==="\r")&&!quoted){if(c==="\r"&&text[i+1]==="\n")i++;row.push(field);if(row.some(x=>String(x).trim()))rows.push(row);row=[];field=""}else field+=c}
 row.push(field);if(row.some(x=>String(x).trim()))rows.push(row);return rows
}
function locateHeader(rows,required){
 const aliases={"tipo documento":["tipo documento","tipo de documento","documento","tipo doc"]};
 const matches=(cells,want)=>{const variants=aliases[want]||[want];return variants.some(v=>cells.includes(norm(v)))};
 for(let i=0;i<Math.min(rows.length,60);i++){const cells=rows[i].map(norm);if(required.every(x=>matches(cells,norm(x))))return i}
 return -1
}
function readConsumptionRows(rows,metaText=""){
 const header=locateHeader(rows,["Código","Cantidad","Tipo Documento"]);
 if(header<0)throw Error("No se reconocieron los encabezados del reporte de salidas. Se buscaron Código, Cantidad y Tipo Documento entre las primeras 60 filas. Verifica que el archivo conserve sus encabezados.");
 const h=rows[header].map(norm);
 const find=(names)=>h.findIndex(x=>names.includes(x));
 const idx={code:find(["codigo"]),qty:find(["cantidad"]),type:find(["tipo documento","tipo de documento","documento","tipo doc"])};
 if(Object.values(idx).some(x=>x<0))throw Error("Se encontró una fila de encabezados, pero falta alguna columna requerida: Código, Cantidad o Tipo Documento.");
 const periodMatch=String(metaText).match(/desde\\s+(\\d{1,2}[\\/.-]\\d{1,2}[\\/.-]\\d{4})\\s+hasta\\s+(\\d{1,2}[\\/.-]\\d{1,2}[\\/.-]\\d{4})/i);
 const start=periodMatch?parseDate(periodMatch[1]):null,end=periodMatch?parseDate(periodMatch[2]):null;
 const records=rows.slice(header+1).map(r=>({code:cleanCode(r[idx.code]),qty:num(r[idx.qty]),type:String(r[idx.type]??"").trim()})).filter(r=>r.code&&r.type);
 return {records,start,end,types:[...new Set(records.map(r=>r.type))].sort((a,b)=>a.localeCompare(b,"es"))}
}
function readConsumption(text){return readConsumptionRows(parseCSV(text),text)}
function readConsumptionWorkbook(workbook){
 const sheet=workbook.Sheets[workbook.SheetNames[0]];
 const rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:""});
 const meta=rows.slice(0,60).flat().join(" ");
 return readConsumptionRows(rows,meta);
}
function readStock(workbook){
 const sheet=workbook.Sheets[workbook.SheetNames[0]],rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:""});
 let header=-1,idx={};
 for(let i=0;i<Math.min(rows.length,35);i++){const cells=rows[i].map(norm);const c=cells.indexOf("codigo"),b=cells.indexOf("bodega"),q=cells.indexOf("cantidad"),d=cells.findIndex(x=>["descripcion","producto","articulo"].includes(x));if(c>=0&&b>=0&&q>=0){header=i;idx={code:c,bodega:b,qty:q,description:d};break}}
 if(header<0)throw Error("No se encontró una fila de encabezados con Código, Bodega y Cantidad en el Excel de stock.");
 const meta=rows.slice(0,header).flat().join(" ");
 const obtained=findDate(meta);
 const records=rows.slice(header+1).map(r=>({code:cleanCode(r[idx.code]),bodega:String(r[idx.bodega]??"").trim(),qty:num(r[idx.qty]),description:idx.description>=0?String(r[idx.description]??"").trim():""})).filter(r=>r.code&&r.bodega);
 return {records,obtained,warehouses:[...new Set(records.map(r=>r.bodega))].sort((a,b)=>a.localeCompare(b,"es"))}
}
function renderDocs(){const box=$("docOptions");box.innerHTML="";consumoData.types.forEach((t,i)=>{const label=document.createElement("label");label.className="option";const cb=document.createElement("input");cb.type="checkbox";cb.value=t;cb.checked=true;cb.addEventListener("change",updateDocCount);const span=document.createElement("span");span.textContent=t;label.append(cb,span);box.append(label)});updateDocCount();$("consumoConfig").classList.remove("hidden")}
function updateDocCount(){const n=[...$("docOptions").querySelectorAll("input:checked")].length;$("docCount").textContent=`${n} de ${consumoData?.types.length||0} seleccionados`}
function renderWarehouses(){
 for(const id of ["pharmacyWarehouse","warehouseStock"]){const s=$(id);s.innerHTML='<option value="">Seleccionar bodega…</option>';stockData.warehouses.forEach(w=>{const o=document.createElement("option");o.value=w;o.textContent=w;s.append(o)})}
 // Sugerencias automáticas: “Farmacia Activa” -> Stock Farmacia;
 // “Bodega Farmacia” -> Stock Bodega. No se asigna la misma bodega dos veces.
 const pharmacyGuess=stockData.warehouses.find(w=>{const n=norm(w);return n.includes("farmacia")&&!n.includes("bodega")});
 const warehouseGuess=stockData.warehouses.find(w=>{const n=norm(w);return n.includes("bodega")&&n.includes("farmacia")});
 if(pharmacyGuess)$("pharmacyWarehouse").value=pharmacyGuess;
 if(warehouseGuess&&warehouseGuess!==pharmacyGuess)$("warehouseStock").value=warehouseGuess;
 $("stockConfig").classList.remove("hidden")
}
function setNotice(el,msg,kind){el.textContent=msg;el.className="notice "+kind}
function validateDates(){
 if(consumoData){const prev=localStorage.getItem(STORAGE_KEY);if(consumoData.end){let msg=`Período detectado: ${fmtDate(consumoData.start)} hasta ${fmtDate(consumoData.end)}.`;if(prev&&dateKey(consumoData.end)<prev)msg+=` ALERTA: la fecha “hasta” (${dateKey(consumoData.end)}) es anterior al último reporte procesado (${prev}). Se permitirá continuar bajo tu responsabilidad.`;setNotice($("periodInfo"),msg,prev&&dateKey(consumoData.end)<prev?"warning":"success")}else setNotice($("periodInfo"),"No se pudo detectar el período “desde… hasta…” en el encabezado. Revisa el archivo antes de procesar.","warning")}
 if(stockData){const today=todayChile(),got=dateKey(stockData.obtained);if(!got)setNotice($("stockDateInfo"),"No se pudo detectar la fecha de obtención del stock. El procesamiento quedará bloqueado.","error");else if(got!==today)setNotice($("stockDateInfo"),`Stock obtenido el ${fmtDate(stockData.obtained)}. Hoy en Chile es ${today.split("-").reverse().join("/")}. No se permite procesar un stock que no sea de hoy.`,"error");else setNotice($("stockDateInfo"),`Stock obtenido hoy (${fmtDate(stockData.obtained)}). Validación correcta.`,"success")}
 $("generatePanel").classList.toggle("hidden",!(consumoData&&stockData));
}
async function loadConsumo(file){consumoFile=file;$("consumoName").textContent=file.name;try{if(/\\.xlsx?$/i.test(file.name)){if(!window.XLSX)throw Error("No se pudo cargar el lector Excel. Comprueba tu conexión a internet y recarga la página.");const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:false});consumoData=readConsumptionWorkbook(wb)}else{const text=await file.text();consumoData=readConsumption(text)}renderDocs();validateDates();setStatus(`Consumo leído: ${consumoData.records.length} registros y ${consumoData.types.length} tipos de documento.`,"success")}catch(e){consumoData=null;$("consumoConfig").classList.add("hidden");validateDates();setStatus(e.message,"error")}}
async function loadStock(file){stockFile=file;$("stockName").textContent=file.name;try{if(!window.XLSX)throw Error("No se pudo cargar el lector Excel. Comprueba tu conexión a internet y recarga la página.");const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:false});stockData=readStock(wb);renderWarehouses();validateDates();setStatus(`Stock leído: ${stockData.records.length} registros y ${stockData.warehouses.length} bodegas detectadas.`,"success")}catch(e){stockData=null;$("stockConfig").classList.add("hidden");validateDates();setStatus(e.message,"error")}}
function setupDrop(inputId,dropId,loader){const input=$(inputId),drop=$(dropId);input.addEventListener("change",()=>{if(input.files[0])loader(input.files[0])});["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("drag")}));["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("drag")}));drop.addEventListener("drop",e=>{const f=e.dataTransfer.files[0];if(f)loader(f)})}
function process(){
 if(!consumoData||!stockData)return;
 const today=todayChile();if(dateKey(stockData.obtained)!==today){alert("Procesamiento bloqueado: el reporte de stock no fue obtenido hoy en Chile.");return}
 const pharmacy=$("pharmacyWarehouse").value,warehouse=$("warehouseStock").value;if(!pharmacy||!warehouse||pharmacy===warehouse){alert("Selecciona dos bodegas distintas: Stock Farmacia y Stock Bodega.");return}
 const selected=[...$("docOptions").querySelectorAll("input:checked")].map(x=>x.value);if(!selected.length){alert("Selecciona al menos un tipo de documento.");return}
 const filtered=consumoData.records.filter(r=>selected.includes(r.type)),cons=new Map();filtered.forEach(r=>cons.set(r.code,(cons.get(r.code)||0)+r.qty));
 const stock=new Map();stockData.records.forEach(r=>{if(r.bodega!==pharmacy&&r.bodega!==warehouse)return;const key=r.code;if(!stock.has(key))stock.set(key,{pharmacy:0,warehouse:0});const o=stock.get(key);if(r.bodega===pharmacy)o.pharmacy+=r.qty;if(r.bodega===warehouse)o.warehouse+=r.qty});
 const codes=new Set([...cons.keys(),...stock.keys()]);const descriptions=new Map();stockData.records.forEach(r=>{if(r.description&&!descriptions.has(r.code))descriptions.set(r.code,r.description)});resultRows=[...codes].map(code=>{const s=stock.get(code)||{pharmacy:0,warehouse:0},c=cons.get(code)||0;return {code,description:descriptions.get(code)||"",consumption:c,pharmacy:s.pharmacy,warehouse:s.warehouse,availability:c>0?s.pharmacy/c*100:null}}).sort((a,b)=>{const ap=a.pharmacy===0&&a.warehouse>0,bp=b.pharmacy===0&&b.warehouse>0;if(ap!==bp)return ap?-1:1;if(a.availability===null&&b.availability!==null)return 1;if(a.availability!==null&&b.availability===null)return -1;if(a.availability!==null&&b.availability!==null&&a.availability!==b.availability)return a.availability-b.availability;return a.code.localeCompare(b.code,"es")});
 const tbody=$("resultsTable").querySelector("tbody");tbody.innerHTML="";resultRows.forEach((r,i)=>{const tr=document.createElement("tr");if(r.pharmacy===0&&r.warehouse>0)tr.className="priority";const vals=[i+1,r.code,r.description,format(r.consumption),format(r.pharmacy),format(r.warehouse),r.availability===null?"N/C":r.availability.toFixed(1).replace(".",",")+"%",""];vals.forEach((v,j)=>{const td=document.createElement("td");td.textContent=v;if(j===7)td.className="manual";tr.append(td)});tbody.append(tr)});
 const priority=resultRows.filter(r=>r.pharmacy===0&&r.warehouse>0).length;$("summary").innerHTML=`<div class="metric"><span>Artículos</span><b>${resultRows.length}</b></div><div class="metric"><span>Prioridad: farmacia 0 / bodega con stock</span><b>${priority}</b></div><div class="metric"><span>Registros de consumo incluidos</span><b>${filtered.length}</b></div>`;
 $("resultSubtitle").textContent=`Consumo: ${fmtDate(consumoData.start)} al ${fmtDate(consumoData.end)} · Stock: ${fmtDate(stockData.obtained)} · Farmacia: ${pharmacy} · Bodega: ${warehouse}`;
 $("resultsPanel").classList.remove("hidden");if(consumoData.end)localStorage.setItem(STORAGE_KEY,dateKey(consumoData.end));$("resultsPanel").scrollIntoView({behavior:"smooth"});
}
function format(n){return Number(n||0).toLocaleString("es-CL",{maximumFractionDigits:2})}
function downloadCSV(){if(!resultRows.length)return;const rows=[["Código","Consumo","Stock Farmacia","Stock Bodega","Disponibilidad (%)","Cantidad a reponer"],...resultRows.map(r=>[r.code,r.consumption,r.pharmacy,r.warehouse,r.availability===null?"":r.availability.toFixed(2),""])];const csv="\uFEFF"+rows.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(";")).join("\r\n");const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8"}));a.download="informe_reposicion.csv";a.click();URL.revokeObjectURL(a.href)}
setupDrop("consumoFile","consumoDrop",loadConsumo);setupDrop("stockFile","stockDrop",loadStock);
$("selectAllDocs").addEventListener("click",()=>{$("docOptions").querySelectorAll("input").forEach(x=>x.checked=true);updateDocCount()});
$("clearDocs").addEventListener("click",()=>{$("docOptions").querySelectorAll("input").forEach(x=>x.checked=false);updateDocCount()});
$("generate").addEventListener("click",process);$("printBtn").addEventListener("click",()=>window.print());$("csvBtn").addEventListener("click",downloadCSV);
})();