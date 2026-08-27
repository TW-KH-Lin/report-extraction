
import { MsgReader } from "https://cdn.jsdelivr.net/npm/@kenjiuno/msgreader-web-ng@0.2.0-alpha1/+esm";

const pdfjsLib = window.pdfjsLib;
pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";

const CATEGORY_SHEETS = ["Final Reports", "Ongoing - Email", "Not in Detail Excel"];
const FAMILY_SHEETS = ["CN95", "CN140ub", "CN140", "CN110", "CN180"];
const CATEGORY_HEADERS = [
  "Lot","Product Family","Customer Company","Rolls Implicated","Samples Received",
  "Final Roll(s)","Master Roll(s)","MR-FR Area(s)",
  "Complaint / Notification","Problem","Tests / Assays Applied","Result / Status","Material No.","Report Date"
];
const FAMILY_HEADERS = ["Source Group", ...CATEGORY_HEADERS];

let workbookBuffer = null;
let records = [];

const $ = (id) => document.getElementById(id);
const esc = (s="") => String(s)
  .replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
  .replaceAll('"',"&quot;").replaceAll("'","&#039;");

function productFamily(material="") {
  const code = material.toUpperCase().replace(/\s/g,"");
  if (code.startsWith("1UN14AR")) return "CN140ub";
  if (code.startsWith("1UN14ER")) return "CN140";
  if (code.startsWith("1UN95")) return "CN95";
  if (code.startsWith("1UN11")) return "CN110";
  if (code.startsWith("1UN18")) return "CN180";
  return "";
}

function firstMatch(text, regexes) {
  for (const re of regexes) {
    const m = text.match(re);
    if (m?.[1]) return m[1].replace(/\s+/g," ").trim().replace(/^[:\s]+|[:\s]+$/g,"");
  }
  return "";
}

function parseMrFr(text) {
  const pairs = [];
  const re = /\bMR-FR\s*:?\s*([0-9]+(?:\/[0-9]+)?)-([0-9]+)\b/gi;
  let m;
  while ((m = re.exec(text))) {
    const key = `${m[1]}-${m[2]}`;
    if (!pairs.some(p => p.key === key)) pairs.push({ key, master:m[1], final:m[2] });
  }
  const listRe = /MR-FR\s*:?\s*((?:\d+(?:\/\d+)?-\d+)(?:\s*[,;]\s*\d+(?:\/\d+)?-\d+)+)/gi;
  while ((m = listRe.exec(text))) {
    for (const part of m[1].split(/[,;]/)) {
      const p = part.trim().match(/^(\d+(?:\/\d+)?)-(\d+)$/);
      if (p) {
        const key = `${p[1]}-${p[2]}`;
        if (!pairs.some(x => x.key === key)) pairs.push({key, master:p[1], final:p[2]});
      }
    }
  }
  const masters = [...new Set(pairs.map(p=>p.master))];
  const finals = [...new Set(pairs.map(p=>p.final))].sort((a,b)=>Number(a)-Number(b));
  const areas = pairs.map(p=>`MR-FR${p.master}-${p.final}`);

  if (!masters.length) {
    const stripped = text.replace(/MR-FR\s*\d+(?:\/\d+)?-\d+/gi,"");
    const rollLists = [...stripped.matchAll(/\bMR(?:s|[\u00B4'\u2019]s)?\s*:?\s*((?:\d+\s*(?:,|and|&)\s*)+\d+)/gi)];
    for (const match of rollLists) {
      masters.push(...match[1].match(/\d+/g) || []);
    }
    if (!masters.length) {
      const retain = [...stripped.matchAll(/\bMR\s*([0-9]+)\b/gi)].map(x=>x[1]);
      masters.push(...retain);
    }
    masters.splice(0, masters.length, ...new Set(masters));
  }
  return { masters, finals, areas };
}

function selectedCheckbox(text, label, options) {
  const i = text.toLowerCase().indexOf(label.toLowerCase());
  if (i < 0) return "";
  const win = text.slice(i, i + 350);
  for (const option of options) {
    const escaped = option.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (new RegExp(`(?:☒|☑|■|\\bX\\b)\\s*${escaped}`, "i").test(win)) return option;
  }
  return "";
}

function extractAssays(text) {
  const assays = [];
  if (/Phenol\s+red\s+buffer\s+line\s+test/i.test(text)) {
    assays.push("Phenol red buffer line test");
  }
  const hasProteinBinding = /Protein\s+(?:binding\s+capacity|lines?)/i.test(text);
  const hasLineMorphology = /line\s+morphology/i.test(text);
  if (hasProteinBinding && hasLineMorphology) {
    let label = "Protein line morphology and binding capacity";
    if (/SyproRuby/i.test(text) && /image\s+analysis/i.test(text)) {
      label += " (SyproRuby staining and image analysis)";
    }
    assays.push(label);
  }
  return assays.join("; ");
}

function parseRecord(text, filename, sourceType) {
  const searchableText = `${text}\nFile name: ${filename}`;
  let complaintNo = firstMatch(searchableText, [
    /\b(Comp\s*-\s*\d{6,10})\b/i,
    /\b(13\d{8})\b/,
    /(?:Complaint|Notification)\s*(?:number|no\.?|#)\s*[:#]?\s*([A-Z0-9][A-Z0-9\-\/]{4,})/i
  ]);
  complaintNo = complaintNo.replace(/\s+/g, "").replace(/^comp/i, "Comp");
  let material = firstMatch(text, [
    /Product\s+code\s*[:#]?\s*([A-Z0-9]+(?:\s+[A-Z])?)(?=\s+Lot\s+number)/i,
    /Product\s+code\s*[:#]?\s*([A-Z0-9]+)/i,
    /\b(1UN(?:95|14|11|18)[A-Z0-9]+)\b/i
  ]);
  material = material.replace(/\s+/g, "");
  const lot = firstMatch(text, [
    /Lot\s+number\s*[:#]?\s*([0-9]{7,9})/i,
    /\blot\s+(?:number\s*)?[:#]?\s*([0-9]{7,9})\b/i
  ]);
  let reportDate = firstMatch(text, [
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{2,4})/i,
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*([0-9]{1,2}\s*[-./]\s*[A-Za-z]{3,9}\s*[-./]\s*[0-9]{2,4})/i,
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*([0-9]{1,4}\s*[-./]\s*[0-9]{1,2}\s*[-./]\s*[0-9]{1,4})/i
  ]);
  reportDate = reportDate.replace(/\s*([./-])\s*/g, "$1");
  const customerCompany = firstMatch(text, [
    /(?:Customer\s+(?:company|organization|account)|Company\s+name)\s*[:#]?\s*([^\n]{2,160})/i,
    /\n[^\n]*@[^\n]*\n([^\n]*(?:Co\.?|Ltd\.?|Inc\.?|LLC|GmbH|Corporation|Company|Biopharm)[^\n]*)/i
  ]);
  const rollsImplicated = firstMatch(text, [
    /Total\s+units\s+implicated\s*[:#]?\s*(\d+)\s*rolls?/i,
    /(?:Number|Total)\s+of\s+rolls?\s+implicated\s*[:#]?\s*(\d+)/i
  ]);
  const samplesReceived = firstMatch(text, [
    /Number\s+of\s+samples\s+received\s*[:#]?\s*(\d+)/i,
    /Samples\s+received\s*[:#]?\s*(\d+)/i
  ]);
  const problem = firstMatch(text, [
    /(?:Issue|Problem|Complaint)\s+description\s*[:#]?\s*([\s\S]{5,600}?)(?=\s+(?:(?:[A-Za-z]+\s+)?Criticality|Complaint\s+status|Date\s+Complaint|Could\s+the\s+failure|Root\s+cause|Investigation|Final\s+assessment|Conclusion|Figure|Fig\.)\b|$)/i,
    /Customer\s+(?:statement|complaint)\s*[:#]?\s*[“"]?([\s\S]{5,600}?)(?=[”"]?\s*(?:Criticality|Complaint\s+status|Figure|Fig\.|Investigation|Conclusion|$))/i,
    /(?:Issue|Problem|Complaint)\s+description\s*[:#]?\s*(.{5,500})/i,
    /Customer\s+(?:statement|complaint)\s*[:#]?\s*[“"]?(.{5,500})/i,
    /Subject\s*:\s*(.{1,180})/i
  ]).replace(/^["“]|["”]$/g,"");
  const status = sourceType === "msg"
    ? "Ongoing – email only"
    : selectedCheckbox(text, "Complaint status", ["Confirmed","Not confirmed","Not conclusive"]);
  const criticality = selectedCheckbox(text, "Criticality", ["Critical","Major","Minor","Track&Trend"]);
  const reproduced = selectedCheckbox(text, "Could the failure be reproduced", ["Yes","No"]);
  const rootCause = selectedCheckbox(text, "Root cause identified within", ["Yes","No"])
    || selectedCheckbox(text, "root cause related", ["Yes","No"]);
  const assaysApplied = extractAssays(text);
  const mrfr = parseMrFr(text);
  const sourceGroup = sourceType === "msg" ? "Ongoing - Email" : "Final Reports";

  const warnings = [];
  if (!complaintNo) warnings.push("Complaint number not extracted");
  if (!material) warnings.push("Material number not extracted");
  if (!lot) warnings.push("Lot not extracted");
  if (!problem) warnings.push("Problem not extracted");
  if (!reportDate) warnings.push("Report date not extracted");
  if (!customerCompany) warnings.push("Customer company not extracted");
  if (sourceType === "pdf" && text.replace(/\s/g, "").length < 80) {
    warnings.push("This PDF has little or no selectable text; OCR may be required");
  }
  if (sourceType === "pdf" && !status) warnings.push("Complaint status not confidently extracted");

  return {
    sourceFile: filename, sourceType, sourceGroup,
    complaintNo, reportDate, customerCompany, rollsImplicated, samplesReceived,
    materialNo: material, productFamily: productFamily(material), lot,
    problem, assaysApplied, resultStatus:status, criticality,
    masterRolls: mrfr.masters.join("; "),
    finalRolls: mrfr.finals.join("; "),
    mrfrAreas: mrfr.areas.join("; "),
    failureReproduced: reproduced,
    rootCauseRelated: rootCause,
    finalAssessment:"",
    warnings: warnings.join("; "),
    rawText:text.slice(0,30000)
  };
}

async function pdfText(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({data:arrayBuffer}).promise;
  const pages = [];
  for (let n=1; n<=pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const content = await page.getTextContent();
    const lines = [];
    let line = [];
    let lastY = null;
    for (const item of content.items) {
      const y = Number(item.transform?.[5]);
      const changedLine = lastY !== null && Number.isFinite(y) && Math.abs(y - lastY) > 2;
      if (line.length && changedLine) {
        lines.push(line.join(" ").trim());
        line = [];
      }
      if (item.str) line.push(item.str);
      if (item.hasEOL && line.length) {
        lines.push(line.join(" ").trim());
        line = [];
      }
      if (Number.isFinite(y)) lastY = y;
    }
    if (line.length) lines.push(line.join(" ").trim());
    pages.push(lines.filter(Boolean).join("\n"));
  }
  return pages.join("\n");
}

async function msgText(arrayBuffer) {
  const reader = new MsgReader(arrayBuffer);
  const info = reader.getFileData();
  const candidates = [
    `Subject: ${info.subject || ""}`,
    `From: ${info.senderName || info.senderEmail || ""}`,
    info.body || "",
    info.bodyHtml || "",
  ];
  return candidates.join("\n").replace(/<[^>]+>/g," ");
}

async function extractOne(name, buffer) {
  const ext = name.toLowerCase().split(".").pop();
  if (ext === "pdf") return parseRecord(await pdfText(buffer), name, "pdf");
  if (ext === "msg") return parseRecord(await msgText(buffer), name, "msg");
  throw new Error(`Unsupported file: ${name}`);
}

async function expandFiles(fileList) {
  const expanded = [];
  for (const file of fileList) {
    if (file.name.toLowerCase().endsWith(".zip")) {
      const zip = await JSZip.loadAsync(await file.arrayBuffer());
      for (const [name, entry] of Object.entries(zip.files)) {
        if (!entry.dir && /\.(pdf|msg)$/i.test(name)) {
          expanded.push({name:name.split("/").pop(), buffer:await entry.async("arraybuffer")});
        }
      }
    } else {
      expanded.push({name:file.name, buffer:await file.arrayBuffer()});
    }
  }
  return expanded;
}

function recordCard(r, index) {
  const groupOptions = CATEGORY_SHEETS.map(x =>
    `<option ${x===r.sourceGroup?"selected":""}>${x}</option>`).join("");
  const famOptions = ["","CN95","CN140ub","CN140","CN110","CN180"].map(x =>
    `<option ${x===r.productFamily?"selected":""}>${x}</option>`).join("");
  return `
  <div class="record" data-index="${index}">
    <div class="record-head">
      <div class="record-title">${esc(r.sourceFile)}</div>
      <button class="secondary remove-record" data-index="${index}">Remove</button>
    </div>
    <div class="grid">
      ${fieldSelect("sourceGroup","Category",groupOptions)}
      ${field("complaintNo","Complaint / Notification",r.complaintNo)}
      ${field("reportDate","Report Date",r.reportDate)}
      ${field("customerCompany","Customer Company",r.customerCompany,"wide")}
      ${field("rollsImplicated","Rolls Implicated",r.rollsImplicated)}
      ${field("samplesReceived","Samples Received",r.samplesReceived)}
      ${field("lot","Lot",r.lot)}
      ${field("materialNo","Material No.",r.materialNo)}
      ${fieldSelect("productFamily","Product Family",famOptions)}
      ${field("masterRolls","Master Roll(s)",r.masterRolls)}
      ${field("finalRolls","Final Roll(s)",r.finalRolls)}
      ${field("mrfrAreas","MR-FR Area(s)",r.mrfrAreas,"wide")}
      ${field("resultStatus","Result / Status",r.resultStatus)}
      ${field("criticality","Criticality",r.criticality)}
      ${field("failureReproduced","Failure Reproduced?",r.failureReproduced)}
      ${field("rootCauseRelated","Root Cause Related?",r.rootCauseRelated)}
      ${textarea("problem","Problem",r.problem,"wide")}
      ${textarea("assaysApplied","Tests / Assays Applied",r.assaysApplied,"wide")}
      ${textarea("finalAssessment","Final Assessment",r.finalAssessment,"wide")}
    </div>
    ${r.warnings ? `<div class="warning">${esc(r.warnings)}</div>` : ""}
    <details><summary>Raw extracted text</summary><pre>${esc(r.rawText)}</pre></details>
  </div>`;
}
function field(name,label,value="",cls="") {
  return `<div class="field ${cls}"><label>${label}</label><input data-field="${name}" value="${esc(value)}"></div>`;
}
function textarea(name,label,value="",cls="") {
  return `<div class="field ${cls}"><label>${label}</label><textarea data-field="${name}">${esc(value)}</textarea></div>`;
}
function fieldSelect(name,label,options,cls="") {
  return `<div class="field ${cls}"><label>${label}</label><select data-field="${name}">${options}</select></div>`;
}

function renderRecords() {
  $("records").innerHTML = records.length
    ? records.map(recordCard).join("")
    : `<p class="hint">No extracted complaints yet.</p>`;
  document.querySelectorAll(".remove-record").forEach(btn => {
    btn.onclick = () => {
      syncRecordsFromDom();
      records.splice(Number(btn.dataset.index),1);
      renderRecords();
    };
  });
  document.querySelectorAll('[data-field="materialNo"]').forEach(input => {
    input.addEventListener("change", e => {
      const card=e.target.closest(".record");
      const fam=productFamily(e.target.value);
      if (fam) card.querySelector('[data-field="productFamily"]').value=fam;
    });
  });
}

function syncRecordsFromDom() {
  document.querySelectorAll(".record").forEach(card => {
    const i = Number(card.dataset.index);
    for (const input of card.querySelectorAll("[data-field]")) {
      records[i][input.dataset.field] = input.value.trim();
    }
  });
}

async function loadWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

function sheetRows(ws) {
  if (!ws || ws.rowCount < 2) return [];
  const headers = ws.getRow(1).values.slice(1).map(v=>String(v||""));
  const out = [];
  for (let r=2; r<=ws.rowCount; r++) {
    const vals=ws.getRow(r).values.slice(1);
    if (!vals.some(v=>v!==null && v!==undefined && v!=="")) continue;
    const obj={};
    headers.forEach((h,i)=>obj[h]=vals[i] ?? "");
    out.push(obj);
  }
  return out;
}

function ensureSheet(wb,name) {
  return wb.getWorksheet(name) || wb.addWorksheet(name);
}

function clearSheet(ws) {
  if (ws.rowCount) ws.spliceRows(1, ws.rowCount);
  try {
    for (const name of Object.keys(ws.tables || {})) ws.removeTable(name);
  } catch (_) {}
}

function formatSheet(ws, hasSource=false) {
  ws.views=[{state:"frozen", xSplit:2, ySplit:1}];
  const widths = hasSource
    ? [20,13,15,28,16,16,18,20,36,24,46,40,26,24,14]
    : [13,15,28,16,16,18,20,36,24,46,40,26,24,14];
  ws.columns.forEach((c,i)=>c.width=widths[i]||18);
  const header=ws.getRow(1);
  header.height=25;
  header.eachCell(cell=>{
    cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF17365D"}};
    cell.font={bold:true,color:{argb:"FFFFFFFF"}};
    cell.alignment={horizontal:"center",vertical:"middle",wrapText:true};
  });
  for (let r=2;r<=ws.rowCount;r++) {
    ws.getRow(r).alignment={vertical:"top",wrapText:true};
  }
  ws.autoFilter={from:{row:1,column:1},to:{row:ws.rowCount,column:ws.columnCount}};
}

function writeSheet(ws, headers, rows, hasSource=false) {
  clearSheet(ws);
  ws.addRow(headers);
  for (const row of rows) ws.addRow(headers.map(h=>row[h] ?? ""));
  formatSheet(ws,hasSource);
}

function normalizeId(v) { return String(v||"").trim().toLowerCase(); }

function recordToCategoryRow(r) {
  const fam = productFamily(r.materialNo) || r.productFamily || "";
  return {
    "Lot":r.lot||"", "Product Family":fam,
    "Customer Company":r.customerCompany||"",
    "Rolls Implicated":r.rollsImplicated||"", "Samples Received":r.samplesReceived||"",
    "Final Roll(s)":r.finalRolls||"", "Master Roll(s)":r.masterRolls||"",
    "MR-FR Area(s)":r.mrfrAreas||"", "Complaint / Notification":r.complaintNo||"",
    "Problem":r.problem||"", "Tests / Assays Applied":r.assaysApplied||"",
    "Result / Status":r.resultStatus||"",
    "Material No.":r.materialNo||"", "Report Date":r.reportDate||""
  };
}

function sortCategory(rows) {
  return rows.sort((a,b)=>{
    const na=Number(String(a["Lot"]||"").match(/\d+/)?.[0]||1e18);
    const nb=Number(String(b["Lot"]||"").match(/\d+/)?.[0]||1e18);
    return na-nb || String(a["Complaint / Notification"]||"").localeCompare(String(b["Complaint / Notification"]||""));
  });
}
function sortFamily(rows) {
  return rows.sort((a,b)=>{
    const p=String(a["Problem"]||"").localeCompare(String(b["Problem"]||""),undefined,{sensitivity:"base"});
    if (p) return p;
    return Number(a["Lot"]||1e18)-Number(b["Lot"]||1e18);
  });
}

async function buildUpdatedWorkbook() {
  if (!workbookBuffer) throw new Error("Load the current Excel workbook first.");
  syncRecordsFromDom();
  const wb = await loadWorkbook(workbookBuffer);

  const categoryRows={};
  for (const name of CATEGORY_SHEETS) categoryRows[name]=sheetRows(ensureSheet(wb,name));

  for (const r of records) {
    const id=normalizeId(r.complaintNo);
    if (!id) continue;
    for (const name of CATEGORY_SHEETS) {
      categoryRows[name]=categoryRows[name].filter(x=>normalizeId(x["Complaint / Notification"])!==id);
    }
    const group=CATEGORY_SHEETS.includes(r.sourceGroup)?r.sourceGroup:"Final Reports";
    categoryRows[group].push(recordToCategoryRow(r));
  }

  for (const name of CATEGORY_SHEETS) {
    const ws=ensureSheet(wb,name);
    writeSheet(ws,CATEGORY_HEADERS,sortCategory(categoryRows[name]),false);
  }

  const familyRows=Object.fromEntries(FAMILY_SHEETS.map(f=>[f,[]]));
  const labels={"Final Reports":"Final Report","Ongoing - Email":"Ongoing - Email","Not in Detail Excel":"Not in Detail Excel"};
  for (const group of CATEGORY_SHEETS) {
    for (const row of categoryRows[group]) {
      const fam=productFamily(row["Material No."]) || row["Product Family"] || "";
      row["Product Family"]=fam;
      if (!familyRows[fam]) continue;
      familyRows[fam].push({"Source Group":labels[group],...row});
    }
  }
  for (const fam of FAMILY_SHEETS) {
    const ws=ensureSheet(wb,fam);
    writeSheet(ws,FAMILY_HEADERS,sortFamily(familyRows[fam]),true);
  }

  const out=await wb.xlsx.writeBuffer();
  return out;
}

async function searchHistory(lot) {
  if (!workbookBuffer) throw new Error("Load the current Excel workbook first.");
  const wb=await loadWorkbook(workbookBuffer);
  const found=[];
  for (const group of CATEGORY_SHEETS) {
    const ws=wb.getWorksheet(group);
    for (const row of sheetRows(ws)) {
      if (String(row["Lot"]||"").trim()===String(lot||"").trim()) found.push({"Source Group":group,...row});
    }
  }
  return found;
}

function historyTable(rows) {
  if (!rows.length) return `<p class="status good">No existing complaint row found for this lot.</p>`;
  const cols=["Source Group","Lot","Product Family","Customer Company","Rolls Implicated","Samples Received","Complaint / Notification","Problem","Tests / Assays Applied","Result / Status","Final Roll(s)"];
  return `<table><thead><tr>${cols.map(c=>`<th>${c}</th>`).join("")}</tr></thead><tbody>
  ${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c]??"")}</td>`).join("")}</tr>`).join("")}
  </tbody></table>`;
}

$("excelFile").addEventListener("change", async e=>{
  const file=e.target.files?.[0];
  if (!file) return;
  workbookBuffer=await file.arrayBuffer();
  try {
    const wb=await loadWorkbook(workbookBuffer);
    $("excelStatus").className="status good";
    $("excelStatus").textContent=`Loaded ${file.name} (${wb.worksheets.length} sheets).`;
  } catch(err) {
    workbookBuffer=null;
    $("excelStatus").className="status bad";
    $("excelStatus").textContent=`Could not read workbook: ${err.message}`;
  }
});

$("extractBtn").onclick=async()=>{
  const files=$("complaintFiles").files;
  if (!files?.length) return;
  $("extractStatus").className="status";
  $("extractStatus").textContent="Extracting...";
  try {
    const expanded=await expandFiles(files);
    const newRecords=[];
    for (const f of expanded) {
      try { newRecords.push(await extractOne(f.name,f.buffer)); }
      catch(err) {
        newRecords.push({
          sourceFile:f.name, sourceType:"", sourceGroup:"Final Reports",
          complaintNo:"", reportDate:"", customerCompany:"", rollsImplicated:"", samplesReceived:"",
          materialNo:"", productFamily:"",
          lot:"", problem:"", assaysApplied:"", resultStatus:"", criticality:"", masterRolls:"",
          finalRolls:"", mrfrAreas:"", failureReproduced:"", rootCauseRelated:"",
          finalAssessment:"", warnings:`Extraction error: ${err.message}`, rawText:""
        });
      }
    }
    records.push(...newRecords);
    renderRecords();
    $("extractStatus").className="status good";
    $("extractStatus").textContent=`Extracted ${newRecords.length} file(s).`;
  } catch(err) {
    $("extractStatus").className="status bad";
    $("extractStatus").textContent=err.message;
  }
};

$("clearBtn").onclick=()=>{ records=[]; renderRecords(); };

$("historyBtn").onclick=async()=>{
  const lot=$("historyLot").value.trim();
  if (!lot) return;
  $("historyResult").innerHTML=`<p class="status">Searching...</p>`;
  try { $("historyResult").innerHTML=historyTable(await searchHistory(lot)); }
  catch(err) { $("historyResult").innerHTML=`<p class="status bad">${esc(err.message)}</p>`; }
};

$("buildBtn").onclick=async()=>{
  $("buildStatus").className="status";
  $("buildStatus").textContent="Building workbook...";
  try {
    const out=await buildUpdatedWorkbook();
    const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url; a.download="Membrane_Complaints_Updated.xlsx";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),2000);
    $("buildStatus").className="status good";
    $("buildStatus").textContent="Updated workbook created and downloaded.";
  } catch(err) {
    $("buildStatus").className="status bad";
    $("buildStatus").textContent=err.message;
  }
};

renderRecords();
