
import { MsgReader } from "./vendor/msgreader.esm.js";

const pdfjsLib = window.pdfjsLib;
pdfjsLib.GlobalWorkerOptions.workerSrc =
  "./vendor/pdf.worker.min.js";

const CATEGORY_SHEETS = ["Final Reports", "Ongoing - Email", "Not in Detail Excel"];
const FAMILY_SHEETS = ["CN95", "CN140ub", "CN140", "CN110", "CN180"];
const LOT_FAMILY_ORDER = ["CN95", "CN110", "CN140", "CN140ub", "CN180"];
const CATEGORY_HEADERS = [
  "Lot","Product Family","Membrane Type","Customer Company","Rolls Implicated","Samples Received",
  "Zone(s)","Final Roll(s)","Master Roll(s)","MR-FR Area(s)","MR-FR (s)",
  "Complaint / Notification","Formal Issue Description","Problem","Customer Reported Failure","Tests / Assays Applied",
  "Standardized Symptom(s)","Problem Type","LFA Relevance",
  "Result / Status","Criticality","Failure Reproduced?","Root Cause in Process?",
  "Product Description","Coordinator","Similar Events Same Category?","Containment Necessary?",
  "Corrective / Preventive Action Necessary?","Root Cause Analysis Conclusion","Problem Description Check",
  "Final Assessment / Root Cause","Final Scope / Decision",
  "Sample Details","Data Quality / Notes","Material No.","Complaint Registered Date","Report Date","Days"
];
const FAMILY_HEADERS = ["Source Group", ...CATEGORY_HEADERS];
const EVIDENCE_SHEET = "Extracted Test Evidence";
const COMPLAINT_SUMMARY_SHEET = "Complaint Summary";
const SUMMARY_SHEET = "Lot & Symptom Summary";
const REVIEW_OVERVIEW_SHEET = "Complaint Overview";
const REVIEW_INVESTIGATION_SHEET = "Complaint Investigation";
const REVIEW_ROOT_CAUSE_SHEET = "Tests & Root Cause";
const EVIDENCE_HEADERS = [
  "Complaint / Notification","Lot","Material No.","Product Family","Units Implicated","Samples Received",
  "Problem","Customer Reported Failure","Complaint Status","Sample Source","Sample ID","Standard Test",
  "Standard Purpose","Standard Method","Result (Source)","Outcome","Within Spec?","Issue Observed?",
  "Source Page","Case-specific Conditions / Source Detail","Source File"
];
const SUMMARY_HEADERS = ["Lot","Complaint Count","Symptom","Symptom Count","Complaint Numbers"];
const COMPLAINT_SUMMARY_HEADERS = [
  "Complaint Number","Lot","Product Family","Material No.","Membrane Type","End Customer",
  "Problem","Customer Reported Failure","Standardized Symptom(s)","Problem Type","Tests Performed",
  "Final Result / Status","Rolls Implicated","Samples Received","MR-FR(s)",
  "Registered Date","Report Date","Days","Data Quality / Notes","Source Group"
];
const REVIEW_OVERVIEW_HEADERS = [
  "Source Group","Complaint Number","Lot Number","End Customer","Final Result / Status",
  "Standardized Symptom(s)","Date Registered","Report Date","Days","Membrane Type","Material No."
];
const REVIEW_INVESTIGATION_HEADERS = [
  "Complaint Number","Lot Number","End Customer","Standardized Symptom(s)","Customer Reported Failure",
  "Tests Performed","MR-FR Area(s)","Rolls Implicated","Samples Received"
];
const REVIEW_ROOT_CAUSE_HEADERS = [
  "Complaint Number","Lot Number","End Customer","Standard Test","Sample Source","Sample ID","Purpose","Method",
  "Result","Outcome","Within Spec?","Issue Observed?","Source Page","Conditions","Conclusion of Root Cause Analysis"
];
const MANAGED_SHEETS = [
  ...CATEGORY_SHEETS,EVIDENCE_SHEET,COMPLAINT_SUMMARY_SHEET,SUMMARY_SHEET,
  REVIEW_OVERVIEW_SHEET,REVIEW_INVESTIGATION_SHEET,REVIEW_ROOT_CAUSE_SHEET,...FAMILY_SHEETS
];
const DRAFT_DB_NAME = "report-extraction-private-drafts";
const DRAFT_STORE = "drafts";
const DRAFT_KEY = "current-results";
const SUMMARY_FIELDS = [
  {key:"complaintNo",label:"Complaint Number",essential:true},
  {key:"lot",label:"Lot(s)",essential:true},
  {key:"lotComplaintCount",label:"Complaints Related to Lot",essential:true},
  {key:"customer",label:"End Customer",essential:true},
  {key:"productFamily",label:"Product Family",essential:true},
  {key:"materialNo",label:"Material No."},
  {key:"problem",label:"Reason / Problem",essential:true},
  {key:"standardizedSymptoms",label:"Standardized Symptom(s)"},
  {key:"problemType",label:"Problem Type"},
  {key:"result",label:"Final Result / Status",essential:true},
  {key:"tests",label:"Tests Performed",essential:true},
  {key:"rollsImplicated",label:"Rolls Implicated"},
  {key:"samplesReceived",label:"Samples Received"},
  {key:"registeredDate",label:"Registered Date"},
  {key:"reportDate",label:"Report Date"},
  {key:"days",label:"Days"},
  {key:"sourceSheets",label:"Source Worksheet(s)"}
];

const REVIEW_FIELD_DEFINITIONS = [
  {key:"sourceGroup",label:"Source Group",aliases:["source group","category","workbook group"]},
  {key:"complaintNo",label:"Complaint / Notification",essential:true,aliases:["complaint notification","complaint number","complaint no","complaint id","complaint ids","complaint id s","notification number","notification","complaint","complaint #"]},
  {key:"complaintRegisteredDate",label:"Complaint Registered Date",aliases:["complaint registered date","registered date","registration date"]},
  {key:"reportDate",label:"Report Date",aliases:["final report date","report date","date of sending final report to customer"]},
  {key:"daysToReport",label:"Days",aliases:["days to report","elapsed days","days"]},
  {key:"customerCompany",label:"End Customer",essential:true,aliases:["customer company","end customer","final customer","customer source label","customer source","customer"]},
  {key:"rollsImplicated",label:"Rolls / Units Implicated",aliases:["rolls implicated","number of roll implicated","implicated units","units implicated"]},
  {key:"samplesReceived",label:"Samples Received",aliases:["samples received","sample received","number of sample received"]},
  {key:"sampleDetails",label:"Sample Details",aliases:["sample details","sample source","sample id"]},
  {key:"lot",label:"Lot / Batch",essential:true,aliases:["lot s","lot number","lot no","product lot if available","batch number","batch no","batch","lot"]},
  {key:"materialNo",label:"Material No.",aliases:["material no","material number","article no","article number"]},
  {key:"productDescription",label:"Product Description",aliases:["product description","material number text","material text","item text"]},
  {key:"productFamily",label:"Product Family",essential:true,aliases:["product family","product type","product"]},
  {key:"membraneType",label:"Membrane Type",aliases:["membrane type"]},
  {key:"zones",label:"Zone(s)",aliases:["zone s","zones","zone"]},
  {key:"masterRolls",label:"Master Roll(s)",aliases:["master roll s","master rolls","master roll"]},
  {key:"finalRolls",label:"Final Roll(s)",aliases:["final roll s","final rolls","final roll"]},
  {key:"mrfrAreas",label:"MR-FR Area(s)",aliases:["mr fr area s","mr fr areas"]},
  {key:"mrfrCombined",label:"MR-FR(s)",aliases:["mr fr s","mr frs","mr fr"]},
  {key:"formalProblem",label:"Formal Issue Description",essential:true,aliases:["formal issue description","issue email statement","issue description","notification text","reported symptom","failure mode"]},
  {key:"problem",label:"Problem / Reason",essential:true,aliases:["customer reported failure","reason problem","problem s","problems","problem","reason"]},
  {key:"customerReportedFailure",label:"Customer Reported Failure",aliases:["customer reported failure"]},
  {key:"standardizedSymptoms",label:"Standardized Symptom(s)",aliases:["standardized symptom s","standardized symptoms","complaint symptom","symptom"]},
  {key:"problemTypes",label:"Problem Type",aliases:["problem type","failure family","defect classification","defect l1 defect classification for failure description","defect l2 defect classification for failure description"]},
  {key:"lfaRelevance",label:"LFA Relevance",aliases:["lfa relevance"]},
  {key:"assaysApplied",label:"Tests / Assays Applied",essential:true,aliases:["tests assays applied","tests performed","test performed","test evidence","standard test","assays applied","assay","test"]},
  {key:"resultStatus",label:"Result / Status",essential:true,aliases:["result status","final result status","claim decision","complaint status","workbook status","result","status"]},
  {key:"criticality",label:"Criticality",aliases:["criticality"]},
  {key:"failureReproduced",label:"Failure Reproduced?",aliases:["failure reproduced"]},
  {key:"rootCauseRelated",label:"Root Cause in Process?",aliases:["root cause in process","root cause related"]},
  {key:"coordinator",label:"Coordinator",aliases:["coordinator","written by"]},
  {key:"similarEvents",label:"Similar Events Same Category?",aliases:["similar events same category","similar events"]},
  {key:"containmentNecessary",label:"Containment Necessary?",aliases:["containment necessary","containment"]},
  {key:"correctiveActionNecessary",label:"Corrective / Preventive Action Necessary?",aliases:["corrective preventive action necessary","corrective action necessary"]},
  {key:"rootCauseConclusion",label:"Conclusion of Root Cause Analysis",aliases:["conclusion of root cause analysis","root cause analysis conclusion"]},
  {key:"problemValidation",label:"Problem Description Check",aliases:["problem description check"]},
  {key:"finalAssessment",label:"Final Assessment",aliases:["detailed final assessment","final assessment root cause","final assessment","historical note"]},
  {key:"finalScope",label:"Final Scope / Decision",aliases:["final scope decision","final decision","scope decision"]},
  {key:"warnings",label:"Data Quality / Notes",aliases:["data quality notes","notes"]},
  {key:"sourceFile",label:"Source File",aliases:["source file","email subject identifier","source"]}
];

const DEFAULT_REVIEW_COLUMNS = REVIEW_FIELD_DEFINITIONS
  .map((item,index)=>({id:`default-${index}`,header:item.label,field:item.key,supported:true,essential:Boolean(item.essential)}));
const DEFAULT_REVIEW_PROFILE = {name:"Default report view",headerRow:-1,columns:DEFAULT_REVIEW_COLUMNS,isDefault:true};
const COMPACT_CN_REVIEW_HEADERS = [
  "Source Group","Lot","Zone(s)","MR-FR (s)","Complaint #","Problem","Result / Status",
  "Material No.","Report Date","Standardized Symptom(s)","Problem Type","LFA Relevance"
];
const BUILT_IN_REVIEW_GROUP_HEADERS = {
  "Complaint Information": [
    "Complaint Number","Lot(s)","Membrane Type","Implicated Units","Sample Received",
    "Registered Date","Report Date","Days","End Customer","Problem(s)","Result"
  ],
  "Complaint Summary": COMPLAINT_SUMMARY_HEADERS,
  "Final Reports": CATEGORY_HEADERS,
  "Ongoing - Email": CATEGORY_HEADERS,
  "Not in Detail Excel": CATEGORY_HEADERS,
  "CN95": COMPACT_CN_REVIEW_HEADERS,
  "CN110": COMPACT_CN_REVIEW_HEADERS,
  "CN140": COMPACT_CN_REVIEW_HEADERS,
  "CN140ub": COMPACT_CN_REVIEW_HEADERS,
  "CN180": COMPACT_CN_REVIEW_HEADERS,
  "Combined Overview": [
    "Final Report Date","Year","Complaint / Notification","Customer","Product Family","Material No",
    "Batch","Reported Symptom","Claim Decision","Failure Family","Defect Classification","Historical Note",
    "Detailed Report?","Complaint Status","Criticality","Root Cause in Process?","Failure Reproduced?",
    "Detailed Final Assessment","Test Evidence?","Source"
  ]
};
const BUILT_IN_REVIEW_PROFILES = Object.entries(BUILT_IN_REVIEW_GROUP_HEADERS).map(([name,headers])=>({
  ...reviewProfileFromMatrix(name,[headers]),isBuiltIn:true
}));

const GENERIC_TEST_RULES = [
  {name:"Visual inspection",re:/\bvisual\s+inspection\b/i,purpose:"Appearance / physical defect review",method:"Inspect customer, retain or reference material and available photographs."},
  {name:"Batch record review",re:/\b(?:batch|manufacturing|production)\s+(?:record|documentation)\s+review\b|\breview\s+of\s+(?:the\s+)?(?:batch|manufacturing|production)\s+(?:record|documentation)/i,purpose:"Manufacturing history review",method:"Review production, release and in-process records for deviations or relevant trends."},
  {name:"Peel strength test",re:/\b(?:peel[- ]test|peel\s+strength\s+(?:test|measurement))\b/i,purpose:"Backing adhesion / peel strength",method:"Measure peel or adhesive strength under the report-defined conditions."},
  {name:"Adhesive strength test",re:/\badhesive\s+strength\s+(?:test|measurement)s?\b/i,purpose:"Backing adhesion",method:"Measure adhesive strength and compare with the applicable internal reference."},
  {name:"Illuminated inspection table",re:/\billuminated\s+inspection\s+table\b/i,purpose:"Surface and structure inspection",method:"Inspect membrane on an illuminated inspection table."},
  {name:"Cold light inspection",re:/\bcold\s+light\s+inspection\b/i,purpose:"Surface damage inspection",method:"Inspect membrane using cold-light illumination."},
  {name:"Phenol red drop test",re:/\bphenol\s+red\s+drop\s+test\b/i,purpose:"Surface wetting / flow behavior",method:"Apply phenol-red solution and assess wetting or spreading behavior."},
  {name:"Functional hCG assay",re:/\b(?:functional\s+)?hcg\s+(?:test|assay)\b/i,purpose:"Functional lateral-flow performance",method:"Run the report-defined hCG lateral-flow assay and compare signal behavior with reference."}
];

let workbookBuffer = null;
let workbookMode = "standard";
let workbookFileName = "";
let records = [];
let lastBuiltSheetNames = [];
let lastPreservedSheetNames = [];
let summaryDataset = [];
let summarySources = [];
let selectedLotFamilies = new Set(LOT_FAMILY_ORDER);
let activeReviewTab = "overview";
let reviewProfiles = [DEFAULT_REVIEW_PROFILE,...BUILT_IN_REVIEW_PROFILES];
let activeReviewProfileName = DEFAULT_REVIEW_PROFILE.name;
const reviewSelections = new Map(reviewProfiles.map(profile=>[
  profile.name,
  new Set(profile.columns.filter(column=>column.supported && (!profile.isDefault || column.essential)).map(column=>column.id))
]));
const collapsedRecords = new Set();

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

function normalizedProductFamily(material="", statedFamily="") {
  const materialCodes=String(material||"").toUpperCase().match(/1UN(?:14AR|14ER|95|11|18)[A-Z0-9-]*/g)||[];
  const derived=[...new Set(materialCodes.map(code=>productFamily(code)).filter(Boolean))];
  if (derived.length) return derived.join("; ");
  const stated=String(statedFamily||"").match(/CN140ub|CN140|CN180|CN110|CN95/gi)||[];
  return [...new Set(stated.map(value=>value.toLowerCase()==="cn140ub"?"CN140ub":value.toUpperCase()))].join("; ");
}

function membraneType(material="", description="") {
  const family=productFamily(material);
  const familyLabel=family==="CN140ub"?"CN140":family;
  const source=`${material} ${description}`.toLowerCase();
  let backing="";
  if (/\bunbacked\b/.test(source) || /^1un14ar/i.test(material)) backing="unbacked";
  else if (/\bbacked\b/.test(source) || /^1un(?:95|11|14er|18)/i.test(material)) backing="backed";
  return [familyLabel,backing].filter(Boolean).join(" ");
}

function complaintClassification(problem="", customerFailure="") {
  const text=cleanBlock(`${problem} ${customerFailure}`).toLowerCase().replace(/[_/]+/g," ");
  const matches=[];
  const rules=[
    ["Poor absorption","Flow / Wetting",/\b(?:poor|low|reduced)?\s*absorption\b|\babsorption\s+(?:issue|problem)\b/],
    ["Abnormal wicking / flow","Flow / Wetting",/\b(?:wicking|slow running|slow flow|fast flow|flow time|fail(?:ed)? to run)\b/],
    ["Wetting issue","Flow / Wetting",/\b(?:wetting|hydrophilic|hydrophobic|dry spot|white circle)\w*\b/],
    ["Uneven line","Line / Printing",/\buneven\b[^.]{0,45}\b(?:print(?:ing|ed)?|line)s?\b|\b(?:print(?:ing|ed)?|line)s?\b[^.]{0,45}\buneven\b/],
    ["Discontinuous line","Line / Printing",/\b(?:discontinuous|interrupted|broken)\b[^.]{0,45}\b(?:print(?:ing|ed)?|line)s?\b|\b(?:print(?:ing|ed)?|line)s?\b[^.]{0,45}\b(?:discontinuous|interrupted|broken)\b/],
    ["Wide / spreading line","Line / Printing",/\b(?:wide|wider|spreading|diffusion|fringing)\b[^.]{0,45}\b(?:print(?:ing|ed)?|line)s?\b/],
    ["Ghost line","Signal / Assay",/\bghost\s+line\b/],
    ["Abnormal signal","Signal / Assay",/\b(?:abnormal|signal)\s+(?:signal|issue)\b/],
    ["Low sensitivity","Signal / Assay",/\blow\s+sensitivity\b/],
    ["False positive","Signal / Assay",/\bfalse\s+positive\b/],
    ["Color / visual issue","Appearance / Surface",/\b(?:coloring|colouring|discolou?r|color variation|colour variation|visual issue|stain|spot)\w*\b/],
    ["Surface roughness","Appearance / Surface",/\b(?:rough|roughness)\b[^.]{0,35}\b(?:membrane|surface)?\b/],
    ["Imprint","Appearance / Surface",/\bimprints?\b/],
    ["Particles","Appearance / Surface",/\b(?:particle|dust|debris)\w*\b/],
    ["Dirt","Appearance / Surface",/\b(?:dirt|contaminat|foreign material)\w*\b/],
    ["Scratch","Mechanical",/\bscratch(?:es|ed)?\b/],
    ["Crack","Mechanical",/\bcracks?\b/],
    ["Rupture","Mechanical",/\b(?:rupture|broken membrane)\w*\b/],
    ["Edge damage / telescoping","Mechanical",/\b(?:edge damage|telescop|mechanical damage)\w*\b/],
    ["Thickness out of specification","Specification / Conversion",/\bthickness\b[^.]{0,35}\b(?:spec|specification|out)\b/],
    ["Length out of specification","Specification / Conversion",/\b(?:length|short roll)\b[^.]{0,35}\b(?:short|spec|specification|out)?\b/],
    ["Width out of specification","Specification / Conversion",/\bwidth\b[^.]{0,35}\b(?:spec|specification|out)\b/],
    ["Low peel strength","Specification / Conversion",/\b(?:low peel strength|peel strength|backing adhesion|label lifting)\b/],
    ["Bag unsealed","Packaging / Storage",/\b(?:bag unsealed|unsealed bag|open bag|packaging integrity)\b/],
    ["Odor / storage concern","Packaging / Storage",/\b(?:odor|odour|storage concern)\b/],
    ["Inter-roll / intra-lot variation","Variation",/\b(?:difference between rolls|roll-to-roll|intra-lot|zone difference|variation between rolls)\b/]
  ];
  for (const [symptom,type,re] of rules) if (re.test(text)) matches.push({symptom,type});
  const symptoms=[...new Set(matches.map(x=>x.symptom))];
  const types=[...new Set(matches.map(x=>x.type))];
  const fallback=cleanBlock(problem).replace(/^(?:performance|quality|product|functionality|functional)(?:\s+\w+)?\s+(?:issue|problem)\s*:?\s*/i,"");
  const standardizedSymptoms=symptoms.length?symptoms.join("; "):(fallback||"Review required");
  const problemTypes=types.length?types.join("; "):"Review required";
  const lfaRelevance=types.some(x=>["Flow / Wetting","Line / Printing","Signal / Assay"].includes(x))
    ?"LFA performance"
    :types.some(x=>["Appearance / Surface","Mechanical","Variation"].includes(x))
      ?"Potential LFA impact"
      :types.some(x=>["Specification / Conversion","Packaging / Storage"].includes(x))
        ?"Non-LFA":"Review required";
  return {standardizedSymptoms,problemTypes,lfaRelevance};
}

function parseFlexibleDate(value="") {
  const text=String(value).trim();
  let m=text.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
  if (m) {
    const year=Number(m[3])+(m[3].length===2?2000:0);
    return new Date(Date.UTC(year,Number(m[2])-1,Number(m[1])));
  }
  m=text.match(/^(\d{1,2})[.\/-]([A-Za-z]{3,9})[.\/-](\d{2,4})$/);
  if (m) {
    const month=new Date(`${m[2]} 1, 2000`).getMonth();
    const year=Number(m[3])+(m[3].length===2?2000:0);
    if (month>=0) return new Date(Date.UTC(year,month,Number(m[1])));
  }
  m=text.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{2,4})$/);
  if (m) {
    const month=new Date(`${m[2]} 1, 2000`).getMonth();
    const year=Number(m[3])+(m[3].length===2?2000:0);
    if (month>=0) return new Date(Date.UTC(year,month,Number(m[1])));
  }
  return null;
}

function daysBetweenDates(start,end) {
  const a=parseFlexibleDate(start), b=parseFlexibleDate(end);
  return a&&b?Math.round((b-a)/86400000):"";
}

function isoDate(value="") {
  const date=value instanceof Date?value:parseFlexibleDate(value);
  if (!date || Number.isNaN(date.getTime())) return String(value||"").trim();
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth()+1).padStart(2,"0")}-${String(date.getUTCDate()).padStart(2,"0")}`;
}

function normalizeExtractedDate(value="") {
  const text=String(value).trim();
  let m=text.match(/^(\d(?:\s*\d)?)\s+([A-Za-z]{3,9})\s+(\d(?:\s*\d){1,3})$/);
  if (m) return `${m[1].replace(/\s/g,"")} ${m[2]} ${m[3].replace(/\s/g,"")}`;
  m=text.match(/^(\d(?:\s*\d)?)\s*([.\/-])\s*([A-Za-z0-9](?:\s*[A-Za-z0-9]){1,8})\s*([.\/-])\s*(\d(?:\s*\d){1,3})$/);
  if (m) return `${m[1].replace(/\s/g,"")}${m[2]}${m[3].replace(/\s/g,"")}${m[4]}${m[5].replace(/\s/g,"")}`;
  return text.replace(/\s*([.\/-])\s*/g,"$1");
}

function firstMatch(text, regexes) {
  for (const re of regexes) {
    const m = text.match(re);
    if (m?.[1]) return m[1].replace(/\s+/g," ").trim().replace(/^[:\s]+|[:\s]+$/g,"");
  }
  return "";
}

function cleanBlock(value="") {
  return String(value)
    .replace(/--- Page \d+ ---/g, " ")
    .replace(/\s+/g, " ")
    .replace(/^[:\s“”"']+|[:\s“”"']+$/g, "")
    .trim();
}

function repairCommonPdfSpacing(value="") {
  return String(value)
    .replace(/\bL\s+ength\b/gi,"Length")
    .replace(/\bC\s+racks?\b/gi,match=>match.toLowerCase().includes("s")?"Cracks":"Crack")
    .replace(/\bDu\s+ring\b/gi,"During")
    .replace(/\bo\s+bserved\b/gi,"observed")
    .replace(/\bH\s+angzhou\b/gi,"Hangzhou")
    .replace(/\bHa\s+ngzhou\b/gi,"Hangzhou")
    .replace(/\bC\s+ustomer\b/gi,"Customer")
    .replace(/\bR\s+eport\b/gi,"Report");
}

function normalizeProblemText(value="") {
  let text=repairCommonPdfSpacing(cleanBlock(value))
    .replace(/^Customer\s+statement\s*:\s*/i,"")
    .replace(/[“”"]/g,"")
    .trim();
  text=text.split(/\s+The\s+customer\s+claimed\b/i)[0].trim()||text;
  text=text.split(/\s+(?:Customer\s+CRN|Product\s+Code|SSB\s+batch|Please\s+use\s+the\s+reference|A\s+report\s+will\s+be\s+available|Due\s+date\s+complaint\s+report)\s*:?/i)[0].trim()||text;
  return text.replace(/^[„“”"']+|[„“”"']+$/g,"").replace(/\s+([,.;:])/g,"$1").replace(/\s{2,}/g," ");
}

function problemFromChineseText(value="") {
  const matches=[];
  const rules=[["Imprints",/压痕/],["Scratches",/划痕/],["Cracks",/裂痕|裂纹/],["Bag unsealed",/未封口/],["Width out of specification",/宽度变窄/],["Thickness issue",/厚度/]];
  for (const [label,re] of rules) if (re.test(value)) matches.push(label);
  return [...new Set(matches)].join(" / ");
}

function productFamilyFromText(value="") {
  const text=String(value);
  if (/CN\s*140\s*(?:ub|unbacked)/i.test(text)) return "CN140ub";
  if (/CN\s*140/i.test(text)) return "CN140";
  if (/CN\s*180/i.test(text)) return "CN180";
  if (/CN\s*110/i.test(text)) return "CN110";
  if (/CN\s*95/i.test(text)) return "CN95";
  return "";
}

function customerFromFilename(filename="") {
  const tag=String(filename).match(/\(([^()]+)\)(?=\.[^.]+$)/)?.[1]||"";
  if (!tag) return "";
  const label=(tag.split("_").pop()||tag).replace(/[-_]+/g," ").trim();
  return /^[A-Z0-9]{2,}$/.test(label)?label:label.replace(/\b\w/g,char=>char.toUpperCase());
}

function canonicalCustomerName(value="") {
  return repairCommonPdfSpacing(cleanBlock(value))
    .replace(/^the\s+customer\s+/i,"")
    .replace(/^(?:customer|company|organization|account)\s*(?:company|name)?\s*[:#-]?\s*/i,"")
    .replace(/\s*\(\s*/g," (").replace(/\s*\)\s*/g,") ")
    .replace(/\bCo\.?\s*,?\s*Ltd\.?\b/gi,"Co., Ltd.")
    .replace(/\s+([,.;])/g,"$1")
    .replace(/\.{2,}$/,".")
    .replace(/\s{2,}/g," ")
    .trim();
}

function plausibleCustomerName(value="") {
  const text=canonicalCustomerName(value);
  return Boolean(text && text.length>=2 && text.length<=140 && !/@|Biotech\s+GmbH|Complaint\s+Report|QN\s+no|thank\s+you/i.test(text));
}

function sectionMatch(text, headingPattern, nextHeadingPattern, maxLength=2400) {
  const re = new RegExp(
    `${headingPattern}\\s*([\\s\\S]{1,${maxLength}}?)(?=${nextHeadingPattern}|$)`,
    "i"
  );
  return cleanBlock(text.match(re)?.[1] || "");
}

function customerCompanyFromHeader(text, filename="", sourceType="") {
  const explicit = firstMatch(text, [
    /(?:Customer\s+(?:company|organization|account)|Company\s+name)\s*[:#]?\s*([^\n]{2,160})/i,
    /(?:new\s+complaint\s+report\s+from|please\s+forward[^\n]{0,80}?\s+to)\s*[:#]?\s*([^\n.]{2,160})/i
  ]);
  if (plausibleCustomerName(explicit)) return canonicalCustomerName(explicit);
  if (sourceType==="msg") return customerFromFilename(filename);
  const lines = text.split(/\n/).map(x=>x.trim()).filter(Boolean);
  const emailIndex = lines.findIndex(line => /@[A-Z0-9.-]+\.[A-Z]{2,}/i.test(line));
  if (emailIndex >= 0) {
    for (let i=emailIndex+1; i<Math.min(emailIndex+4, lines.length); i++) {
      const line=lines[i];
      if (/^(?:\d|Report\s+date|Complaint\s+information)/i.test(line)) continue;
      if (/@/.test(line)) continue;
      const candidate=line.replace(/\b([A-Za-z]{5,})\s+([a-z])\b/g,"$1$2").replace(/\s{2,}.*$/, "").trim();
      if (!/^(?:From:|Address:|Subject:|Dear\b|We\s+acknowledge)/i.test(candidate) && plausibleCustomerName(candidate)) return canonicalCustomerName(candidate);
    }
  }
  const fallback=firstMatch(text, [
    /\n([^\n]*(?:Co\.?|Ltd\.?|Inc\.?|LLC|GmbH|Corporation|Company|Biopharm|Biotechnology)[^\n]*)/i
  ]);
  if (sourceType!=="msg" && !/Biotech\s+GmbH/i.test(fallback)) return canonicalCustomerName(fallback);
  return customerFromFilename(filename);
}

function coordinatorFromText(text) {
  const raw=firstMatch(text, [
    /Written\s+by\s*:\s*(?:Reviewed\s+by\s*:)?\s*\n\s*([A-Z][A-Za-zÀ-ÿ'’-]+(?:\s+[A-Z][A-Za-zÀ-ÿ'’-]+){1,3})/i,
    /\n\s*([A-Z][A-Za-zÀ-ÿ'’-]+(?:\s+[A-Z][A-Za-zÀ-ÿ'’-]+){1,3})\s*\n\s*Quality\s+Professional/i
  ]);
  const parts=raw.split(/\s+/).filter(Boolean);
  return parts.length>=4 ? parts.slice(0,2).join(" ") : raw;
}

function validateProblemAgainstRootCause(problem, customerFailure, rootCauseConclusion) {
  if (!rootCauseConclusion) return "Not checked - root-cause conclusion not extracted";
  const stop = new Set([
    "about","after","against","although","been","being","complaint","conclusion","could","failure",
    "from","have","identified","issue","most","process","product","related","report","root","sample",
    "that","their","there","these","this","those","through","were","which","within","with"
  ]);
  const tokens = value => new Set((String(value).toLowerCase().match(/[a-z]{4,}/g)||[])
    .filter(x=>!stop.has(x))
    .map(x=>x.endsWith("ing")?x.slice(0,-3):x.endsWith("ed")?x.slice(0,-2):x.endsWith("s")?x.slice(0,-1):x));
  const problemTokens=tokens(`${problem} ${customerFailure}`);
  const rootTokens=tokens(rootCauseConclusion);
  const overlap=[...problemTokens].filter(x=>rootTokens.has(x));
  if (overlap.length>=2) return `Consistent with root-cause conclusion (${overlap.slice(0,4).join(", ")})`;
  if (overlap.length===1) return `Broad match - review wording (${overlap[0]})`;
  return "Potential mismatch - review problem description against root-cause conclusion";
}

function enrichProblemDescription(formalProblem, customerFailure, rootCauseConclusion) {
  const formal=cleanBlock(formalProblem);
  let detail=cleanBlock(customerFailure).replace(/[.!?]+$/g,"");
  const isBroad=/^(?:performance|quality|product|functionality|functional)(?:\s+\w+)?\s+(?:issue|problem)$/i.test(formal);
  if (!isBroad) return formal || detail;
  if (/absorption(?:-related)?\s+(?:issue|problem)/i.test(rootCauseConclusion) && !/absorption/i.test(detail)) {
    detail=detail?`Absorption issue; ${detail}`:"Absorption issue";
  }
  if (!detail && /line\s+(?:behavior|quality|printing)/i.test(rootCauseConclusion)) detail="Line quality / printing issue";
  if (!detail) return formal;
  return `${formal}: ${detail.charAt(0).toLowerCase()}${detail.slice(1)}`;
}

function parseMrFr(text) {
  const pairs = [];
  const re = /\bMR\s*-\s*FR\s*:?\s*([0-9]+(?:\s*\/\s*[0-9]+)?)\s*-\s*([0-9]+)\b/gi;
  let m;
  while ((m = re.exec(text))) {
    const key = `${m[1]}-${m[2]}`;
    if (!pairs.some(p => p.key === key)) pairs.push({ key, master:m[1], final:m[2] });
  }
  const listRe = /MR\s*-\s*FR\s*:?\s*((?:\d+(?:\s*\/\s*\d+)?\s*-\s*\d+)(?:\s*[,;]\s*\d+(?:\s*\/\s*\d+)?\s*-\s*\d+)+)/gi;
  while ((m = listRe.exec(text))) {
    for (const part of m[1].split(/[,;]/)) {
      const p = part.trim().match(/^(\d+(?:\s*\/\s*\d+)?)\s*-\s*(\d+)$/);
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
    const stripped = text.replace(/MR\s*-\s*FR\s*\d+(?:\s*\/\s*\d+)?\s*-\s*\d+/gi,"");
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
  const labelPattern=label.replace(/[.*+?^${}()|[\]\\]/g,"\\$&").replace(/\s+/g,"\\s+");
  const labelMatch=text.match(new RegExp(labelPattern,"i"));
  if (!labelMatch || labelMatch.index===undefined) return "";
  const win = text.slice(labelMatch.index, labelMatch.index + 350);
  for (const option of options) {
    const escaped = option.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (new RegExp(`(?:☒|☑|■|\\bX\\b)\\s*${escaped}`, "i").test(win)) return option;
  }
  return "";
}

function extractAssays(text) {
  const assays = [];
  const hasFlowMeasurement = /Capillary\s+Flow\s+Time\s*#|could\s+not\s+be\s+measured\s+against\s+specification/i.test(text);
  const hasInProcessReview = /Review\s+in[- ]process\s+data\s+of\s+capillary\s+flow\s+time/i.test(text);
  if (hasFlowMeasurement) assays.push("Capillary flow time");
  else if (hasInProcessReview) assays.push("In-process data review");
  if (/Phenol\s+red\s+(?:buffer\s+)?line\s+test/i.test(text)) assays.push("Phenol red line test");
  if (/Protein\s+(?:binding\s+(?:assay|capacity)|lines?)/i.test(text)) {
    assays.push("Protein binding assay");
  }
  for (const rule of GENERIC_TEST_RULES) if (rule.re.test(text)) assays.push(rule.name);
  return [...new Set(assays)].join("; ");
}

function sourcePageFor(text, phrase) {
  const i=text.toLowerCase().indexOf(phrase.toLowerCase());
  if (i<0) return "";
  const before=text.slice(0,i);
  const pages=[...before.matchAll(/--- Page (\d+) ---/g)];
  return pages.length ? `p.${pages.at(-1)[1]}` : "";
}

function sourcePageRange(text, phrases) {
  const pages=phrases.map(x=>sourcePageFor(text,x)).filter(Boolean).map(x=>Number(x.replace("p.","")));
  if (!pages.length) return "";
  const first=Math.min(...pages), last=Math.max(...pages);
  return first===last?`p.${first}`:`p.${first}-${last}`;
}

function sourceContextForRegex(text,re,maxLength=360) {
  const flags=re.flags.includes("g")?re.flags:re.flags+"g";
  const match=new RegExp(re.source,flags).exec(text);
  if (!match || match.index===undefined) return "";
  const pageEnd=text.indexOf("--- Page ",match.index+1);
  const end=Math.min(match.index+maxLength,pageEnd>match.index?pageEnd:text.length);
  return cleanBlock(text.slice(match.index,end));
}

function genericEvidenceOutcome(result="") {
  const text=String(result).toLowerCase();
  const negatedIssue=/\b(?:no|not|without)\b[^.]{0,45}\b(?:defect|damage|issue|deviation|discrepanc|irregularit|scratch|crack|imprint)/i.test(result);
  if (/within\s+(?:the\s+)?(?:internal\s+)?specification|no\s+(?:visible\s+)?(?:defect|issue|deviation|discrepanc|irregularit)/i.test(result)) return {outcome:"Pass",withinSpec:"Yes",issueObserved:"No"};
  if (!negatedIssue && /below\s+(?:the\s+)?(?:internal\s+)?(?:threshold|specification)|out\s+of\s+specification|reveals?\s+(?:low|a\s+defect)|(?:defect|damage|scratch|crack|imprint)\s+(?:was|were)\s+(?:observed|identified)/i.test(result)) return {outcome:"Issue observed",withinSpec:"No",issueObserved:"Yes"};
  if (text) return {outcome:"Review required",withinSpec:"",issueObserved:""};
  return {outcome:"Recorded",withinSpec:"",issueObserved:""};
}

function extractTestEvidence(text) {
  const tests=[];
  const rollInfo=parseMrFr(text);
  const retainIds=rollInfo.masters.map(x=>`MR${x}`).join("; ");
  const returnedIds=firstMatch(text, [
    /returned\s+\d+\s+(?:pieces|samples)[\s\S]{0,180}?\(([#\d\s,;and]+)(?:identification|customer|\))/i
  ]).replace(/\s+/g," ").replace(/\s*,\s*/g,"; ").replace(/\s+and\s+/gi,"; ");
  if (/capillary\s+flow\s+time/i.test(text)) {
    const result=firstMatch(text, [
      /(The\s+capillary\s+flow\s+time[\s\S]{5,420}?)(?=As\s+the\s+samples|Complaint\s+number|--- Page|$)/i,
      /(in-process[- ]data[\s\S]{5,260}?within\s+specification)/i
    ]);
    tests.push({
      name:/Capillary\s+Flow\s+Time\s*#|could\s+not\s+be\s+measured\s+against\s+specification/i.test(text)
        ?"Capillary flow time":"In-process data review",
      purpose:"Flow performance / specification check",
      sampleSource:"Retain sample",
      sampleId:retainIds,
      method:"Measure capillary flow time over 40 mm using the report-defined liquid/conditions and compare with specification/reference.",
      result:cleanBlock(result),
      outcome:/within\s+specification/i.test(result)?"Pass":"Review required",
      withinSpec:/within\s+specification/i.test(result)?"Yes":"",
      issueObserved:"No",
      sourcePage:sourcePageFor(text,"capillary flow time"),
      conditions:/65\s+to\s+115\s+sec\s*\/\s*40\s*mm/i.test(text)
        ?"Specification 65-115 sec/40 mm; customer samples were post-use and not measurable against specification":""
    });
  }
  if (/Phenol\s+red\s+buffer\s+line\s+test/i.test(text)) {
    const result=firstMatch(text, [
      /(The\s+printed\s+buffer\s+lines[\s\S]{5,360}?)(?=Protein\s+binding|Complaint\s+number|--- Page|$)/i,
      /(Phenol\s+red\s+buffer\s+lines[\s\S]{5,240}?(?:irregularities|disruptions|reference))/i
    ]);
    const issue=/irregularit|interrupt|disrupt/i.test(result) && !/without\s+(?:any\s+)?(?:irregularit|interrupt|disrupt)/i.test(result);
    tests.push({
      name:"Phenol red line test", purpose:"Line quality / wetting", sampleSource:"Retain sample",
      sampleId:retainIds,
      method:"Print a phenol-red buffer line on the membrane and compare continuity, width, shape and wetting with reference.",
      result:cleanBlock(result),
      outcome:issue?"Minor irregularity / review":"Pass", withinSpec:issue?"":"Yes",
      issueObserved:issue?"Partial":"No", sourcePage:sourcePageFor(text,"Phenol red buffer line test"),
      conditions:/1\s+or\s+2\s*[μµu]l\s*\/\s*cm/i.test(text)?"Phenol red buffer printed at 1 or 2 µl/cm":""
    });
  }
  if (/Protein\s+(?:binding\s+(?:assay|capacity)|lines?)/i.test(text)) {
    const normal=firstMatch(text, [
      /(The\s+protein\s+lines[\s\S]{5,260}?(?:reference\s+membrane|without\s+any\s+disruptions))/i
    ]);
    const issues=firstMatch(text, [
      /(The\s+following\s+issues\s+were\s+identified[\s\S]{5,720}?)(?=No\s+irregularities|1\.4\.|Manufacturing\s+documentation|--- Page|$)/i
    ]);
    const result=cleanBlock([normal,issues].filter(Boolean).join(" "));
    tests.push({
      name:"Protein binding assay", purpose:"Binding capacity / line morphology",
      sampleSource:returnedIds?"Customer return + retain":"Retain sample",
      sampleId:[returnedIds?`Return pieces ${returnedIds}`:"",retainIds?`Retain ${retainIds}`:""].filter(Boolean).join("; "),
      method:"Print protein line(s), stain with SyproRuby, and evaluate line morphology and binding/signal versus reference.",
      result,
      outcome:issues?"Issue reproduced / mixed":"Pass", withinSpec:issues?"Functionality retained":"Yes",
      issueObserved:issues?"Yes":"No", sourcePage:sourcePageRange(text,["Protein binding","The following issues were identified"]),
      conditions:/0\.5\s+to\s+4\s+mg\s*\/\s*ml/i.test(text)?"Protein 0.5-4 mg/ml; 1 µl/cm; SyproRuby staining and image analysis":""
    });
  }
  const existing=new Set(tests.map(test=>test.name));
  for (const rule of GENERIC_TEST_RULES) {
    if (existing.has(rule.name) || !rule.re.test(text)) continue;
    const result=sourceContextForRegex(text,rule.re);
    const assessment=genericEvidenceOutcome(result);
    const phrase=text.match(rule.re)?.[0]||rule.name;
    tests.push({
      name:rule.name,purpose:rule.purpose,sampleSource:"See report",sampleId:retainIds,
      method:rule.method,result,...assessment,sourcePage:sourcePageFor(text,phrase),conditions:""
    });
  }
  return tests;
}

function complaintIdsFromFilename(filename="") {
  return [...String(filename).matchAll(/Comp\s*-\s*(\d{4,10})/gi)]
    .map(match=>`Comp-${String(Number(match[1])).padStart(7,"0")}`);
}

function parseRecord(text, filename, sourceType) {
  const rawSourceText=text;
  text=repairCommonPdfSpacing(text);
  const searchableText = `${text}\nFile name: ${filename}`;
  let complaintNo = firstMatch(searchableText, [
    /\b(Comp\s*-\s*\d{6,10})\b/i,
    /\b(13\d{8})\b/,
    /(?:Complaint|Notification)\s*(?:number|no\.?|#)\s*[:#]?\s*([A-Z0-9][A-Z0-9\-\/]{4,})/i
  ]);
  complaintNo = complaintNo.replace(/\s+/g, "").replace(/^comp/i, "Comp");
  let material = firstMatch(text, [
    /Product\s+code\s*[:#]?\s*((?:1UN)[A-Z0-9\s]{7,35}?)(?=\s+Lot\s+number|\s+SSB\s+batch|\n)/i,
    /Product\s+code\s*[:#]?\s*([A-Z0-9]+)/i,
    /\b(1UN(?:95|14|11|18)[A-Z0-9]+)\b/i
  ]);
  material = material.replace(/\s+/g, "");
  const productDescription = firstMatch(text, [
    /Product\s+description\s*[:#]?\s*([\s\S]{3,220}?)(?=\s+Total\s+units\s+implicated|\s+Number\s+of\s+samples|\n)/i
  ]);
  const lot = firstMatch(text, sourceType==="msg" ? [
    /Subject\s*:[^\n]*?(?:batch|lot|批次)\s*([0-9]{7,9})/i,
    /Batch\s+number\s+([0-9]{7,9})\s+must\s+be\s+the\s+correct\s+one/i,
    /SSB\s+(?:batch|lot)\s*(?:No\.?|number)?\s*[:#]?\s*([0-9]{7,9})/i,
    /批次\s*([0-9]{7,9})/i,
    /Lot\s+number\s*[:#]?\s*([0-9]{7,9})/i,
    /\blot\s+(?:number\s*)?[:#]?\s*([0-9]{7,9})\b/i
  ] : [
    /Lot\s+number\s*[:#]?\s*([0-9]{7,9})/i,
    /\blot\s+(?:number\s*)?[:#]?\s*([0-9]{7,9})\b/i
  ]);
  let reportDate = firstMatch(text, [
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*(\d(?:\s*\d)?\s+[A-Za-z]{3,9}\s+\d(?:\s*\d){1,3})/i,
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*([0-9]{1,2}\s*[-./]\s*[A-Za-z]{3,9}\s*[-./]\s*[0-9]{2,4})/i,
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*([0-9]{1,4}\s*[-./]\s*[0-9]{1,2}\s*[-./]\s*[0-9]{1,4})/i
  ]);
  reportDate = normalizeExtractedDate(reportDate);
  const customerCompany = canonicalCustomerName(customerCompanyFromHeader(text,filename,sourceType));
  const rollsImplicated = firstMatch(text, [
    /Total\s+units\s+implicated\s*[:#]?\s*(\d+)\s*rolls?/i,
    /(?:Number|Total)\s+of\s+rolls?\s+implicated\s*[:#]?\s*(\d+)/i
  ]);
  const samplesReceived = firstMatch(text, [
    /Number\s+of\s+samples\s+received\s*[:#]?\s*(\d+)/i,
    /Samples\s+received\s*[:#]?\s*(\d+)/i
  ]);
  let sampleDetails = firstMatch(text, [
    /Number\s+of\s+samples\s+received\s*[:#]?\s*([\s\S]{1,100}?)(?=\s+Date\s+(?:samples|complaint)|\s+Issue\s+description|\n)/i
  ]);
  const receivedIdentifiers=firstMatch(text, [
    /returned\s+\d+\s+(?:pieces|samples)[\s\S]{0,180}?\(([#\d\s,;and]+)(?:identification|customer|\))/i
  ]).replace(/\s+/g," ");
  if (receivedIdentifiers && !sampleDetails.includes(receivedIdentifiers)) {
    sampleDetails=cleanBlock(`${sampleDetails} (${receivedIdentifiers})`);
  }
  let formalProblem = normalizeProblemText(firstMatch(text, [
    /(?:Issue|Problem|Complaint)\s+description\s*[:#]?\s*([\s\S]{5,600}?)(?=\s+(?:(?:[A-Za-z]+\s+)?Criticality|Complaint\s+status|Date\s+Complaint|Could\s+the\s+failure|Root\s+cause|Investigation|Final\s+assessment|Conclusion|Figure|Fig\.)\b|$)/i,
    /Customer\s+(?:statement|complaint)\s*[:#]?\s*[“"]?([\s\S]{5,600}?)(?=[”"]?\s*(?:Criticality|Complaint\s+status|Figure|Fig\.|Investigation|Conclusion|$))/i,
    /(?:Issue|Problem|Complaint)\s+description\s*[:#]?\s*(.{5,500})/i,
    /Customer\s+(?:statement|complaint)\s*[:#]?\s*[“"]?(.{5,500})/i,
    /Subject\s*:\s*(.{1,240})/i
  ]).replace(/^["“]|["”]$/g,""));
  if ((!formalProblem || /投诉批次|异常沟通|Final\s+Report/i.test(formalProblem)) && problemFromChineseText(text)) {
    formalProblem=problemFromChineseText(text);
  }
  const status = sourceType === "msg"
    ? "Ongoing – email only"
    : selectedCheckbox(text, "Complaint status", ["Confirmed","Not confirmed","Not conclusive"]);
  const criticality = selectedCheckbox(text, "Criticality", ["Critical","Major","Minor","Track&Trend"]);
  const reproduced = selectedCheckbox(text, "Could the failure be reproduced", ["Yes","No"]);
  const rootCause = selectedCheckbox(text, "Root cause identified within", ["Yes","No"])
    || selectedCheckbox(text, "root cause related", ["Yes","No"]);
  const assaysApplied = extractAssays(text);
  const testEvidence = extractTestEvidence(text);
  const mrfr = parseMrFr(text);
  const zones = firstMatch(text, [
    /Zone\s*\(s\)\s*[:#]?\s*([^\n]{1,120})/i,
    /(?:Membrane|Roll)\s+zones?\s*[:#]?\s*([^\n]{1,120})/i
  ]);
  const mrfrCombined = mrfr.areas.length
    ? mrfr.areas.join("; ")
    : mrfr.masters.map(x=>`MR${x}`).join("; ");
  const sourceGroup = sourceType === "msg" ? "Ongoing - Email" : "Final Reports";
  const customerReportedFailure = normalizeProblemText(firstMatch(text, [
    /received\s+a\s+complaint\s+with\s+the\s+following\s+statement\s*:\s*[“"]?([\s\S]{3,500}?)(?=\s*(?:Picture\s*s|Pictures|Figure|1\.2\.|Criticality|--- Page))/i,
    /Customer\s+statement[\s\S]{0,420}?[“"]([^”"]{3,500})[”"]/i,
    /Customer\s+statement\s*:\s*([\s\S]{3,700}?)(?=\s+(?:The\s+customer\s+claimed|Root\s+cause|(?:Company|Manufacturer)\s+Criticality|Complaint\s+status|1\.2\.|Criticality|--- Page))/i,
    /Issue\s+description\s*:\s*[“"]?([\s\S]{3,700}?)[”"]?\s*(?=Customer\s+CRN|Product\s+Code|SSB\s+batch|Due\s+date|Please\s+use)/i
  ]).replace(/\s*-\s*/g,"-").replace(/[“”"]+/g,""));
  const rootCauseConclusion = sectionMatch(
    text,
    "Conclusion\\s+of\\s+the\\s+root\\s+cause\\s+analysis\\s*",
    "(?:4\\.\\s*Correction|4\\.\\s*Conclusion|5\\.\\s*Conclusion|Corrective\\s*/\\s*Preventive|--- Page)",
    2600
  );
  const problem = enrichProblemDescription(formalProblem, customerReportedFailure, rootCauseConclusion);
  const problemValidation = validateProblemAgainstRootCause(problem, customerReportedFailure, rootCauseConclusion);
  const finalAssessment = sectionMatch(
    text,
    "(?:^|\\n)\\s*(?:4|5)\\.\\s*Conclusion\\s*",
    "(?:Best\\s+regards|Written\\s+by|Complaint\\s+number|--- Page)",
    2200
  );
  const similarEvents = selectedCheckbox(text, "Similar events reported", ["Yes","No"]);
  const containmentNecessary = selectedCheckbox(text, "Containment action necessary", ["Yes","No"]);
  const correctiveActionNecessary = selectedCheckbox(text, "Corrective / Preventive action necessary", ["Yes","No"])
    || selectedCheckbox(text, "Corrective / Preventive actions necessary", ["Yes","No"]);
  const finalScope = firstMatch(text, [
    /(The\s+scope\s+of\s+failure\s+is\s+concluded[\s\S]{3,240}?\.)(?=\s|$)/i,
    /(Complaint\s+(?:not\s+confirmed|confirmed)[\s\S]{3,180}?(?:monitoring|claimed\s+units)\.?)/i
  ]);
  const coordinator = coordinatorFromText(text);
  const reportVersion = firstMatch(text, [/Report\s+version\s*[:#]?\s*([A-Z0-9.-]+)/i]);
  let complaintRegisteredDate = firstMatch(text, [
    /Date\s+complaint\s+registered\s*[:#]?\s*(\d(?:\s*\d)?\s*[-./]\s*[A-Za-z0-9](?:\s*[A-Za-z0-9]){1,8}\s*[-./]\s*\d(?:\s*\d){1,3})/i,
    /Date\s+complaint\s+registered\s*[:#]?\s*([0-9]{1,2}\s*[-./]\s*[A-Za-z0-9]{2,9}\s*[-./]\s*[0-9]{2,4})/i,
    /Date\s+complaint\s+registered\s*[:#]?\s*([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{2,4})/i
  ]);
  complaintRegisteredDate=normalizeExtractedDate(complaintRegisteredDate);
  const samplesReceivedDate = firstMatch(text, [/Date\s+samples\s+received\s*[:#]?\s*([A-Z0-9][A-Z0-9 .\/-]{1,30})/i]);
  const classification = complaintClassification(problem,customerReportedFailure);
  const daysToReport = daysBetweenDates(complaintRegisteredDate,reportDate);

  const warnings = [];
  if (!complaintNo) warnings.push("Complaint number not extracted");
  if (!material) warnings.push("Material number not extracted");
  if (!lot) warnings.push("Lot not extracted");
  if (!problem) warnings.push("Problem not extracted");
  if (!complaintRegisteredDate) warnings.push("Complaint registered date not extracted");
  if (!reportDate) warnings.push("Report date not extracted");
  if (!customerCompany) warnings.push("Customer company not extracted");
  if (!customerReportedFailure) warnings.push("Customer-reported failure not extracted");
  if (classification.problemTypes==="Review required") warnings.push("Standardized symptom classification requires review");
  if (problemValidation.startsWith("Potential mismatch")) warnings.push(problemValidation);
  if (sourceType === "pdf" && text.replace(/\s/g, "").length < 80) {
    warnings.push("This PDF has little or no selectable text; OCR may be required");
  }
  if (sourceType === "pdf" && !status) warnings.push("Complaint status not confidently extracted");
  const filenameComplaints=complaintIdsFromFilename(filename);
  if (filenameComplaints.length===1 && complaintNo && normalizeComplaintId(filenameComplaints[0])!==normalizeComplaintId(complaintNo)) {
    warnings.push(`Filename says ${filenameComplaints[0]}, but report content says ${complaintNo}`);
  }

  return {
    sourceFile: filename, sourceType, sourceGroup,
    complaintNo, reportDate, customerCompany, rollsImplicated, samplesReceived,
    sampleDetails, materialNo: material, productDescription,
    productFamily: productFamily(material)||productFamilyFromText(`${productDescription} ${text.slice(0,1200)}`), lot,
    membraneType:membraneType(material,productDescription)||productFamilyFromText(`${productDescription} ${text.slice(0,1200)}`), zones, mrfrCombined,
    standardizedSymptoms:classification.standardizedSymptoms,
    problemTypes:classification.problemTypes, lfaRelevance:classification.lfaRelevance,
    formalProblem, problem, assaysApplied, resultStatus:status, criticality,
    masterRolls: mrfr.masters.join("; "),
    finalRolls: mrfr.finals.join("; "),
    mrfrAreas: mrfr.areas.join("; "),
    failureReproduced: reproduced,
    rootCauseRelated: rootCause,
    customerReportedFailure, coordinator, reportVersion,
    complaintRegisteredDate, samplesReceivedDate, daysToReport,
    similarEvents, containmentNecessary, correctiveActionNecessary,
    rootCauseConclusion, problemValidation, finalAssessment, finalScope, testEvidence,
    warnings: warnings.join("; "),
    rawText:rawSourceText.slice(0,30000)
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
    pages.push(`--- Page ${n} ---\n${lines.filter(Boolean).join("\n")}`);
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

function markerIndexes(marker="") {
  const output=new Set();
  for (const part of String(marker).split(/[+,;&]/).map(value=>value.trim()).filter(Boolean)) {
    const range=part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) {
      const start=Number(range[1]), end=Number(range[2]);
      for (let value=Math.min(start,end);value<=Math.max(start,end);value++) output.add(value);
    } else if (/^\d+$/.test(part)) output.add(Number(part));
  }
  return [...output];
}

function assignMarkedValue(target,marker,value) {
  for (const index of markerIndexes(marker)) if (value) target.set(index,cleanBlock(value));
}

function markedValuesFromArea(area,valuePattern) {
  const values=new Map();
  let match;
  const re=new RegExp(`${valuePattern}\\s*\\(([0-9+,&;\\s-]+)\\)`,"gi");
  while ((match=re.exec(area))) assignMarkedValue(values,match[2],match[1]);
  return values;
}

function splitNumberedPdfCases(text,base) {
  const early=repairCommonPdfSpacing(text.slice(0,12000));
  const ids=new Map();
  let match;
  const idRe=/\b(13\d{8}|Comp\s*-\s*\d{6,10})\s*\((\d+)\)/gi;
  while ((match=idRe.exec(early))) {
    const id=match[1].replace(/\s+/g,"").replace(/^comp/i,"Comp");
    if (!ids.has(Number(match[2]))) ids.set(Number(match[2]),id);
  }
  if (ids.size<2) return [base];

  const materialByCase=new Map();
  const materialRe=/\b(1UN(?:14AR|14ER|95|11|18)[A-Z0-9\s]{5,28}?)\s*\(([0-9+,&;\s-]+)\)/gi;
  while ((match=materialRe.exec(early))) {
    const code=match[1].replace(/\s+/g,"");
    if (/^1UN(?:14AR|14ER|95|11|18)[A-Z0-9]{5,}$/i.test(code)) assignMarkedValue(materialByCase,match[2],code);
  }
  const lotByCase=new Map();
  const lotRe=/\b(2[0-9]{6})\s*\(([0-9+,&;\s-]+)\)/g;
  while ((match=lotRe.exec(early))) assignMarkedValue(lotByCase,match[2],match[1]);
  const rollsByCase=new Map();
  const rollsRe=/\b([0-9]+)\s*rolls?\s*\(([0-9+,&;\s-]+)\)/gi;
  while ((match=rollsRe.exec(early))) assignMarkedValue(rollsByCase,match[2],match[1]);
  const dateByCase=new Map();
  const dateRe=/(\d{1,2}\s*[-./]\s*[A-Za-z]{3,9}\s*[-./]\s*\d{2,4})\s*\(([0-9+,&;\s-]+)\)/gi;
  while ((match=dateRe.exec(early))) assignMarkedValue(dateByCase,match[2],normalizeExtractedDate(match[1]));

  const problemArea=early.match(/Issue\s+description[\s\S]{0,120}?Customer\s+statement\s*:\s*([\s\S]{1,900}?)(?=The\s+customer\s+claimed|Root\s+cause)/i)?.[1]||"";
  const problemByCase=new Map();
  const problemRe=/(?:^|\n)\s*([^()\n]{3,180}?)\s*\(([0-9+,&;\s-]+)\)\s*(?=\n|$)/g;
  while ((match=problemRe.exec(problemArea))) {
    const value=normalizeProblemText(match[1]);
    if (value && !/^(?:issue\s+description|customer\s+statement)$/i.test(value)) assignMarkedValue(problemByCase,match[2],value);
  }

  const confirmedArea=firstMatch(early,[/Confirmed\s+for\s+lot\s+([\s\S]{1,180}?)(?=Not\s+confirmed|Not\s+conclusive|II\.|$)/i]);
  const notConfirmedArea=firstMatch(early,[/Not\s+confirmed\s+for\s+lot\s+([\s\S]{1,220}?)(?=Not\s+conclusive|II\.|$)/i]);
  const confirmedLots=new Set(confirmedArea.match(/\b\d{7,9}\b/g)||[]);
  const notConfirmedLots=new Set(notConfirmedArea.match(/\b\d{7,9}\b/g)||[]);

  return [...ids.entries()].sort((a,b)=>a[0]-b[0]).map(([index,complaintNo])=>{
    const materialNo=materialByCase.get(index)||base.materialNo;
    const lot=lotByCase.get(index)||base.lot;
    const complaintRegisteredDate=normalizeExtractedDate(dateByCase.get(index)||base.complaintRegisteredDate);
    const problem=problemByCase.get(index)||base.problem;
    let resultStatus=base.resultStatus;
    if (confirmedLots.has(lot)) resultStatus="Confirmed";
    else if (notConfirmedLots.has(lot)) resultStatus="Not confirmed";
    const classification=complaintClassification(problem,problem);
    const notice=`Multi-complaint report: case ${index} of ${ids.size}; review case-specific MR-FR and shared test evidence`;
    return {
      ...structuredClone(base),complaintNo,materialNo,lot,problem,formalProblem:problem,
      customerReportedFailure:problem,rollsImplicated:rollsByCase.get(index)||base.rollsImplicated,
      complaintRegisteredDate,daysToReport:daysBetweenDates(complaintRegisteredDate,base.reportDate),resultStatus,
      productFamily:productFamily(materialNo),membraneType:membraneType(materialNo,base.productDescription),
      standardizedSymptoms:classification.standardizedSymptoms,problemTypes:classification.problemTypes,lfaRelevance:classification.lfaRelevance,
      warnings:mergeSummaryValue(base.warnings,notice)
    };
  });
}

function splitAcknowledgementEmailCases(text,base,filename) {
  const starts=[...text.matchAll(/We\s+acknowledge\s+the\s+receipt\s+of\s+your\s+reported\s+complaint\s*\((Comp\s*-\s*\d{6,10})\)/gi)];
  if (starts.length<2) return [base];
  return starts.map((start,index)=>{
    const end=starts[index+1]?.index ?? text.search(/Please\s+use\s+the\s+reference/i);
    const block=text.slice(start.index,end>start.index?end:text.length);
    const record=parseRecord(block,filename,"msg");
    record.complaintNo=start[1].replace(/\s+/g,"").replace(/^comp/i,"Comp");
    record.customerCompany=customerFromFilename(filename)||base.customerCompany;
    record.warnings=mergeSummaryValue(record.warnings,`Multi-complaint email: extracted complaint ${index+1} of ${starts.length}`);
    return record;
  });
}

async function extractMany(name, buffer) {
  const ext = name.toLowerCase().split(".").pop();
  if (ext === "pdf") {
    const text=await pdfText(buffer);
    return splitNumberedPdfCases(text,parseRecord(text,name,"pdf"));
  }
  if (ext === "msg") {
    const text=await msgText(buffer);
    return splitAcknowledgementEmailCases(text,parseRecord(text,name,"msg"),name);
  }
  throw new Error(`Unsupported file: ${name}`);
}

async function extractOne(name, buffer) {
  return (await extractMany(name,buffer))[0];
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
  const collapsed=collapsedRecords.has(recordUiKey(r));
  const visibleFields=activeReviewLabels();
  const summaryParts=[];
  if (visibleFields.has("complaintNo") && r.complaintNo) summaryParts.push(r.complaintNo);
  if (visibleFields.has("lot") && r.lot) summaryParts.push(`Lot ${r.lot}`);
  const showSourceDetails=Boolean($("reviewSourceDetails")?.checked);
  const groupOptions = CATEGORY_SHEETS.map(x =>
    `<option ${x===r.sourceGroup?"selected":""}>${x}</option>`).join("");
  const famOptions = ["","CN95","CN140ub","CN140","CN110","CN180"].map(x =>
    `<option ${x===r.productFamily?"selected":""}>${x}</option>`).join("");
  return `
  <div class="record${collapsed?" collapsed":""}" data-index="${index}">
    <div class="record-head">
      <div>
        <div class="record-title">${esc(r.sourceFile)}</div>
        <div class="record-summary">${esc(summaryParts.join(" · "))}</div>
      </div>
      <div class="record-actions">
        <button class="secondary toggle-record" data-index="${index}" aria-expanded="${collapsed?"false":"true"}">${collapsed?"Show details":"Fold"}</button>
        <button class="secondary remove-record" data-index="${index}">Remove</button>
      </div>
    </div>
    <div class="record-body"${collapsed?" hidden":""}>
    <div class="grid">
      ${fieldSelect("sourceGroup","Category",groupOptions)}
      ${field("complaintNo","Complaint / Notification",r.complaintNo)}
      ${field("complaintRegisteredDate","Complaint Registered Date",r.complaintRegisteredDate)}
      ${field("reportDate","Report Date",r.reportDate)}
      ${field("daysToReport","Days to Report",r.daysToReport)}
      ${field("customerCompany","Customer Company",r.customerCompany,"wide")}
      ${field("rollsImplicated","Rolls Implicated",r.rollsImplicated)}
      ${field("samplesReceived","Samples Received",r.samplesReceived)}
      ${field("sampleDetails","Sample Details",r.sampleDetails,"wide")}
      ${field("lot","Lot",r.lot)}
      ${field("materialNo","Material No.",r.materialNo)}
      ${field("productDescription","Product Description",r.productDescription,"wide")}
      ${fieldSelect("productFamily","Product Family",famOptions)}
      ${field("membraneType","Membrane Type",r.membraneType)}
      ${field("zones","Zone(s)",r.zones)}
      ${field("masterRolls","Master Roll(s)",r.masterRolls)}
      ${field("finalRolls","Final Roll(s)",r.finalRolls)}
      ${field("mrfrAreas","MR-FR Area(s)",r.mrfrAreas,"wide")}
      ${field("mrfrCombined","MR-FR (s)",r.mrfrCombined,"wide")}
      ${field("resultStatus","Result / Status",r.resultStatus)}
      ${field("criticality","Criticality",r.criticality)}
      ${field("failureReproduced","Failure Reproduced?",r.failureReproduced)}
      ${field("rootCauseRelated","Root Cause Related?",r.rootCauseRelated)}
      ${field("formalProblem","Formal Issue Description",r.formalProblem,"wide")}
      ${textarea("problem","Workbook Problem (enriched)",r.problem,"wide")}
      ${textarea("customerReportedFailure","Customer Reported Failure",r.customerReportedFailure,"wide")}
      ${field("standardizedSymptoms","Standardized Symptom(s)",r.standardizedSymptoms,"wide")}
      ${field("problemTypes","Problem Type",r.problemTypes,"wide")}
      ${field("lfaRelevance","LFA Relevance",r.lfaRelevance)}
      ${textarea("assaysApplied","Tests / Assays Applied",r.assaysApplied,"wide")}
      ${field("coordinator","Coordinator / Written By",r.coordinator)}
      ${field("similarEvents","Similar Events Same Category?",r.similarEvents)}
      ${field("containmentNecessary","Containment Necessary?",r.containmentNecessary)}
      ${field("correctiveActionNecessary","Corrective / Preventive Action Necessary?",r.correctiveActionNecessary)}
      ${textarea("rootCauseConclusion","Conclusion of Root Cause Analysis",r.rootCauseConclusion,"wide")}
      ${field("problemValidation","Problem Description Check",r.problemValidation,"wide")}
      ${textarea("finalAssessment","Final Assessment",r.finalAssessment,"wide")}
      ${textarea("finalScope","Final Scope / Decision",r.finalScope,"wide")}
    </div>
    ${visibleFields.has("assaysApplied")?testEvidenceTable(r.testEvidence):""}
    ${showSourceDetails && r.warnings ? `<div class="warning">${esc(r.warnings)}</div>` : ""}
    ${showSourceDetails?`<details><summary>Raw extracted text</summary><pre>${esc(r.rawText)}</pre></details>`:""}
    </div>
  </div>`;
}
function testEvidenceTable(tests=[]) {
  if (!tests.length) return "";
  return `<div class="test-evidence"><h3>Structured test evidence</h3><div class="table-scroll"><table><thead><tr>
    <th>Standard Test</th><th>Sample Source</th><th>Sample ID</th><th>Purpose</th><th>Method</th><th>Result</th>
    <th>Outcome</th><th>Within Spec?</th><th>Issue Observed?</th><th>Source Page</th><th>Conditions</th>
  </tr></thead><tbody>${tests.map(t=>`<tr>
    <td>${esc(t.name)}</td><td>${esc(t.sampleSource)}</td><td>${esc(t.sampleId)}</td><td>${esc(t.purpose)}</td>
    <td>${esc(t.method)}</td><td>${esc(t.result)}</td><td>${esc(t.outcome)}</td><td>${esc(t.withinSpec)}</td>
    <td>${esc(t.issueObserved)}</td><td>${esc(t.sourcePage)}</td><td>${esc(t.conditions)}</td>
  </tr>`).join("")}</tbody></table></div></div>`;
}
function field(name,label,value="",cls="") {
  const presentation=reviewFieldPresentation(name,label);
  return `<div class="field ${cls}" data-review-field="${name}"${presentation.visible?"":" hidden"}><label>${esc(presentation.label)}</label><input data-field="${name}" value="${esc(value)}"></div>`;
}
function textarea(name,label,value="",cls="") {
  const presentation=reviewFieldPresentation(name,label);
  return `<div class="field ${cls}" data-review-field="${name}"${presentation.visible?"":" hidden"}><label>${esc(presentation.label)}</label><textarea data-field="${name}">${esc(value)}</textarea></div>`;
}
function fieldSelect(name,label,options,cls="") {
  const presentation=reviewFieldPresentation(name,label);
  return `<div class="field ${cls}" data-review-field="${name}"${presentation.visible?"":" hidden"}><label>${esc(presentation.label)}</label><select data-field="${name}">${options}</select></div>`;
}

function currentLotRows() {
  const lots=new Map();
  for (const record of records) {
    const lot=String(record.lot||"").trim()||"Lot not extracted";
    if (!lots.has(lot)) lots.set(lot,{lot,complaints:new Set(),families:new Set(),symptoms:new Set(),statuses:new Set()});
    const row=lots.get(lot);
    row.complaints.add(record.complaintNo||record.sourceFile||"Unnumbered complaint");
    const family=productFamily(record.materialNo)||record.productFamily||record.membraneType||"";
    if (family) row.families.add(family);
    const classification=complaintClassification(record.problem,record.customerReportedFailure);
    for (const symptom of String(record.standardizedSymptoms||classification.standardizedSymptoms||"").split(";").map(x=>x.trim()).filter(Boolean)) row.symptoms.add(symptom);
    if (record.resultStatus) row.statuses.add(record.resultStatus);
  }
  return [...lots.values()].sort((a,b)=>a.lot.localeCompare(b.lot,undefined,{numeric:true}));
}

function renderLotsTable() {
  const target=$("lotsTable");
  if (!target) return;
  const rows=currentLotRows();
  target.innerHTML=rows.length
    ? `<div class="table-scroll"><table><thead><tr><th>Lot</th><th>Complaint count</th><th>Complaint number(s)</th><th>Product family</th><th>Standardized symptom(s)</th><th>Status</th></tr></thead><tbody>${rows.map(row=>`<tr><td>${esc(row.lot)}</td><td>${row.complaints.size}</td><td>${esc([...row.complaints].join("; "))}</td><td>${esc([...row.families].join("; "))}</td><td>${esc([...row.symptoms].join("; "))}</td><td>${esc([...row.statuses].join("; "))}</td></tr>`).join("")}</tbody></table></div>`
    : `<p class="hint">No extracted lots yet.</p>`;
}

const ORGANIZED_REVIEW_TABS = {
  overview:{
    label:"Complaint overview",
    fields:[
      ["sourceGroup","Source Group","select"],["complaintNo","Complaint Number"],["lot","Lot Number"],
      ["customerCompany","End Customer","wide"],["resultStatus","Final Result / Status"],
      ["standardizedSymptoms","Standardized Symptom(s)","wide"],["complaintRegisteredDate","Date Registered","date"],
      ["reportDate","Report Date","date"],["daysToReport","Days"],["membraneType","Membrane Type"],["materialNo","Material No."]
    ]
  },
  investigation:{
    label:"Complaint investigation",
    fields:[
      ["complaintNo","Complaint Number"],["lot","Lot Number"],["customerCompany","End Customer","wide"],
      ["standardizedSymptoms","Standardized Symptom(s)","wide"],["customerReportedFailure","Customer Reported Failure","long"],
      ["assaysApplied","Tests Performed","long"],["mrfrAreas","MR-FR Area(s)","wide"],
      ["rollsImplicated","Rolls Implicated"],["samplesReceived","Samples Received"]
    ]
  },
  evidence:{
    label:"Tests & root cause",
    fields:[
      ["complaintNo","Complaint Number"],["lot","Lot Number"],["customerCompany","End Customer","wide"],
      ["rootCauseConclusion","Conclusion of Root Cause Analysis","long"]
    ]
  }
};

function structuredTestEvidenceText(record) {
  const tests=record.testEvidence||[];
  if (!tests.length) return record.assaysApplied?`Tests performed: ${record.assaysApplied}`:"No structured test evidence extracted.";
  return tests.map((test,index)=>{
    const parts=[
      test.sampleSource&&`Sample: ${test.sampleSource}${test.sampleId?` (${test.sampleId})`:""}`,
      test.purpose&&`Purpose: ${test.purpose}`,
      test.method&&`Method: ${test.method}`,
      test.result&&`Result: ${test.result}`,
      test.outcome&&`Outcome: ${test.outcome}`,
      test.withinSpec&&`Within spec: ${test.withinSpec}`,
      test.issueObserved&&`Issue observed: ${test.issueObserved}`,
      test.sourcePage&&`Source: ${test.sourcePage}`,
      test.conditions&&`Conditions: ${test.conditions}`
    ].filter(Boolean);
    return `${index+1}. ${test.name||"Test"}${parts.length?`\n${parts.join("\n")}`:""}`;
  }).join("\n\n");
}

function testMethodItems(value="") {
  const text=String(value||"").trim();
  if (!text) return [];
  return text.split(/\n+|;\s+|\.\s+(?=[A-Z])/).map(item=>item.trim().replace(/[.;]+$/,"")).filter(Boolean);
}

function testMethodBulletText(value="") {
  return testMethodItems(value).map(item=>`• ${item}`).join("\n");
}

function testMethodBulletHtml(value="") {
  const items=testMethodItems(value);
  return items.length?`<ul class="test-method-list">${items.map(item=>`<li>${esc(item)}</li>`).join("")}</ul>`:"";
}

function organizedReviewCell(record,index,definition) {
  const [key,label,type="text"]=definition;
  if (type==="evidence") return `<td class="organized-cell evidence-column"><div class="structured-evidence-text">${esc(structuredTestEvidenceText(record)).replaceAll("\n","<br>")}</div></td>`;
  let value=record[key]||"";
  if (type==="date") value=isoDate(value);
  if (type==="select") {
    const options=CATEGORY_SHEETS.map(group=>`<option value="${esc(group)}"${group===value?" selected":""}>${esc(group)}</option>`).join("");
    return `<td class="organized-cell"><select aria-label="${esc(label)} for row ${index+1}" data-field="${key}">${options}</select></td>`;
  }
  if (type==="long") return `<td class="organized-cell long-column"><textarea aria-label="${esc(label)} for row ${index+1}" data-field="${key}">${esc(value)}</textarea></td>`;
  const cls=type==="wide"?" wide-column":"";
  return `<td class="organized-cell${cls}"><input aria-label="${esc(label)} for row ${index+1}" data-field="${key}" value="${esc(value)}"></td>`;
}

function renderStructuredEvidenceReview() {
  const headers=["Complaint Number","Lot Number","End Customer","Standard Test","Sample Source","Sample ID","Purpose","Method","Result","Outcome","Within Spec?","Issue Observed?","Source Page","Conditions"];
  const rows=records.flatMap(record=>(record.testEvidence||[]).map(test=>[
    record.complaintNo||"",record.lot||"",record.customerCompany||"",test.name||"",test.sampleSource||"",test.sampleId||"",
    test.purpose||"",test.method||"",test.result||"",test.outcome||"",test.withinSpec||"",test.issueObserved||"",test.sourcePage||"",test.conditions||""
  ]));
  return `<section class="structured-evidence-section"><div class="organized-review-heading"><strong>Structured Test Evidence</strong><span>${rows.length} test row${rows.length===1?"":"s"}</span></div>
    ${rows.length?`<div class="table-scroll"><table class="structured-evidence-table"><thead><tr>${headers.map(header=>`<th>${esc(header)}</th>`).join("")}</tr></thead><tbody>
      ${rows.map(row=>`<tr>${row.map((value,index)=>`<td>${index===7?testMethodBulletHtml(value):esc(value)}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`
      :`<p class="hint">No structured test evidence was extracted for the current complaints.</p>`}</section>`;
}

function renderRecords() {
  const target=$("records");
  const tab=ORGANIZED_REVIEW_TABS[activeReviewTab]||ORGANIZED_REVIEW_TABS.overview;
  if (!records.length) {
    target.innerHTML=`<p class="hint">No extracted complaints yet.</p>`;
    renderLotsTable();
    return;
  }
  target.innerHTML=`<div class="organized-review-heading"><strong>${esc(tab.label)}</strong><span>${records.length} complaint${records.length===1?"":"s"} · one complaint per row</span></div>
    <div class="table-scroll organized-review-scroll"><table class="organized-review-table"><thead><tr>
      ${tab.fields.map(([,label])=>`<th>${esc(label)}</th>`).join("")}<th>Action</th>
    </tr></thead><tbody>${records.map((record,index)=>`<tr data-record-row data-index="${index}">
      ${tab.fields.map(definition=>organizedReviewCell(record,index,definition)).join("")}
      <td class="organized-action"><button type="button" class="secondary remove-record" data-index="${index}">Remove</button></td>
    </tr>`).join("")}</tbody></table></div>${activeReviewTab==="evidence"?renderStructuredEvidenceReview():""}`;
  target.querySelectorAll(".remove-record").forEach(btn=>{
    btn.onclick=()=>{
      syncRecordsFromDom();
      const index=Number(btn.dataset.index);
      collapsedRecords.delete(recordUiKey(records[index]));
      records.splice(index,1);
      renderRecords();
    };
  });
  target.querySelectorAll("[data-field]").forEach(input=>{
    input.addEventListener("change",()=>{
      syncRecordsFromDom();
      renderLotsTable();
    });
  });
  renderLotsTable();
}

function syncRecordsFromDom() {
  document.querySelectorAll("[data-record-row]").forEach(row=>{
    const index=Number(row.dataset.index);
    for (const input of row.querySelectorAll("[data-field]")) records[index][input.dataset.field]=input.value.trim();
    const family=productFamily(records[index].materialNo);
    if (family) records[index].productFamily=family;
    records[index].membraneType=membraneType(records[index].materialNo,records[index].productDescription)||records[index].membraneType||"";
    records[index].daysToReport=daysBetweenDates(records[index].complaintRegisteredDate,records[index].reportDate)||records[index].daysToReport||"";
  });
}

function recordUiKey(record) {
  return normalizeComplaintId(record?.complaintNo)||normalizeId(record?.sourceFile);
}

function matchingRecordIndex(record) {
  const complaint=normalizeComplaintId(record.complaintNo);
  if (complaint) {
    const byComplaint=records.findIndex(existing=>normalizeComplaintId(existing.complaintNo)===complaint);
    return byComplaint;
  }
  const source=normalizeId(record.sourceFile);
  return source ? records.findIndex(existing=>normalizeId(existing.sourceFile)===source) : -1;
}

function openDraftDatabase() {
  return new Promise((resolve,reject)=>{
    const request=indexedDB.open(DRAFT_DB_NAME,1);
    request.onupgradeneeded=()=>{
      const db=request.result;
      if (!db.objectStoreNames.contains(DRAFT_STORE)) db.createObjectStore(DRAFT_STORE,{keyPath:"id"});
    };
    request.onsuccess=()=>resolve(request.result);
    request.onerror=()=>reject(request.error||new Error("Temporary browser storage is unavailable."));
  });
}

async function saveTemporaryDraft() {
  syncRecordsFromDom();
  if (!records.length) throw new Error("Extract at least one complaint report before saving.");
  const db=await openDraftDatabase();
  await new Promise((resolve,reject)=>{
    const tx=db.transaction(DRAFT_STORE,"readwrite");
    tx.objectStore(DRAFT_STORE).put({
      id:DRAFT_KEY,
      records:structuredClone(records),
      collapsed:[...collapsedRecords],
      selectedSheets:[...selectedExportSheets()],
      savedAt:new Date().toISOString()
    });
    tx.oncomplete=resolve;
    tx.onerror=()=>reject(tx.error||new Error("Could not save temporary results."));
    tx.onabort=()=>reject(tx.error||new Error("Temporary save was interrupted."));
  });
  db.close();
}

async function readTemporaryDraft() {
  const db=await openDraftDatabase();
  const draft=await new Promise((resolve,reject)=>{
    const tx=db.transaction(DRAFT_STORE,"readonly");
    const request=tx.objectStore(DRAFT_STORE).get(DRAFT_KEY);
    request.onsuccess=()=>resolve(request.result||null);
    request.onerror=()=>reject(request.error||new Error("Could not restore temporary results."));
  });
  db.close();
  return draft;
}

async function deleteTemporaryDraft() {
  const db=await openDraftDatabase();
  await new Promise((resolve,reject)=>{
    const tx=db.transaction(DRAFT_STORE,"readwrite");
    tx.objectStore(DRAFT_STORE).delete(DRAFT_KEY);
    tx.oncomplete=resolve;
    tx.onerror=()=>reject(tx.error||new Error("Could not clear temporary results."));
  });
  db.close();
}

async function restoreTemporaryDraft() {
  try {
    const draft=await readTemporaryDraft();
    if (!draft?.records?.length || records.length) return;
    records=draft.records;
    collapsedRecords.clear();
    for (const key of draft.collapsed||[]) collapsedRecords.add(key);
    if (Array.isArray(draft.selectedSheets)) {
      const selected=new Set(draft.selectedSheets);
      document.querySelectorAll("[data-export-sheet]").forEach(input=>{input.checked=selected.has(input.value);});
    }
    renderRecords();
    const savedTime=draft.savedAt?new Date(draft.savedAt).toLocaleString():"an earlier session";
    $("draftStatus").className="status good";
    $("draftStatus").textContent=`Restored ${records.length} temporarily saved complaint(s) from ${savedTime}.`;
  } catch(err) {
    $("draftStatus").className="status bad";
    $("draftStatus").textContent=`Temporary restore unavailable: ${err.message}`;
  }
}

async function loadWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

async function isValidXlsxContainer(buffer) {
  try {
    const zip=await JSZip.loadAsync(buffer.slice(0));
    return Boolean(zip.file("xl/workbook.xml") && zip.file("[Content_Types].xml"));
  } catch (_) {
    return false;
  }
}

async function readReferenceWorkbook(buffer) {
  const zip=await JSZip.loadAsync(buffer.slice(0));
  const parse=async path=>new DOMParser().parseFromString(await zip.file(path).async("string"),"application/xml");
  const workbook=await parse("xl/workbook.xml");
  const rels=await parse("xl/_rels/workbook.xml.rels");
  const relMap={};
  for (const rel of rels.getElementsByTagNameNS("*","Relationship")) relMap[rel.getAttribute("Id")]=rel.getAttribute("Target");
  const shared=[];
  if (zip.file("xl/sharedStrings.xml")) {
    const sharedDoc=await parse("xl/sharedStrings.xml");
    for (const si of sharedDoc.getElementsByTagNameNS("*","si")) {
      shared.push([...si.getElementsByTagNameNS("*","t")].map(x=>x.textContent||"").join(""));
    }
  }
  const sheets={};
  for (const sheet of workbook.getElementsByTagNameNS("*","sheet")) {
    const name=sheet.getAttribute("name");
    const rid=sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id")
      || sheet.getAttribute("r:id");
    let target=(relMap[rid]||"").replace(/^\//,"");
    if (!target.startsWith("xl/")) target=`xl/${target}`;
    if (!zip.file(target)) continue;
    const doc=await parse(target);
    const rows=[];
    for (const row of doc.getElementsByTagNameNS("*","row")) {
      const values=[];
      for (const cell of row.getElementsByTagNameNS("*","c")) {
        const ref=cell.getAttribute("r")||"A1";
        const letters=ref.match(/[A-Z]+/)?.[0]||"A";
        let col=0; for (const ch of letters) col=col*26+ch.charCodeAt(0)-64; col-=1;
        const type=cell.getAttribute("t");
        let value="";
        if (type==="inlineStr") value=[...cell.getElementsByTagNameNS("*","t")].map(x=>x.textContent||"").join("");
        else value=cell.getElementsByTagNameNS("*","v")[0]?.textContent||"";
        if (type==="s" && value!=="") value=shared[Number(value)]??value;
        values[col]=value;
      }
      if (values.some(x=>x!==undefined && x!=="")) rows[Number(row.getAttribute("r")||rows.length+1)]=values;
    }
    sheets[name]=rows;
  }
  return sheets;
}

function objectsFromReferenceRows(rows, headerRow) {
  const headers=rows?.[headerRow]||[];
  const out=[];
  for (let r=headerRow+1;r<(rows?.length||0);r++) {
    const values=rows[r]||[];
    if (!values.some(x=>x!==undefined && x!=="")) continue;
    const obj={}; headers.forEach((h,i)=>{if(h)obj[h]=values[i]??"";}); out.push(obj);
  }
  return out;
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

function displayCellValue(value) {
  if (value===null || value===undefined) return "";
  if (value instanceof Date) return value.toLocaleDateString("en-GB");
  if (typeof value==="object") {
    if (value.text!==undefined) return String(value.text);
    if (value.result!==undefined) return displayCellValue(value.result);
    if (Array.isArray(value.richText)) return value.richText.map(part=>part.text||"").join("");
    if (value.formula!==undefined || value.sharedFormula!==undefined) return "";
  }
  return String(value).trim();
}

function normalizeHeader(value) {
  return displayCellValue(value).toLowerCase().replace(/&/g,"and").replace(/[^a-z0-9]+/g," ").trim();
}

function worksheetMatrix(ws) {
  const rows=[];
  if (!ws) return rows;
  for (let r=1;r<=ws.rowCount;r++) {
    const values=[];
    for (let c=1;c<=ws.columnCount;c++) values.push(displayCellValue(ws.getCell(r,c).value));
    rows.push(values);
  }
  return rows;
}

function detectComplaintHeaderRow(matrix) {
  let best={index:-1,score:0};
  const clues=[
    /complaint|notification/,/\blot\b|batch/,/customer/,/problem|symptom|issue/,
    /result|status|conclusion/,/test|assay/,/material/,/report date|registered date/
  ];
  matrix.slice(0,12).forEach((row,index)=>{
    const headers=row.map(normalizeHeader).filter(Boolean);
    const score=clues.reduce((sum,re)=>sum+(headers.some(header=>re.test(header))?1:0),0);
    if (score>best.score) best={index,score};
  });
  return best.score>=2?best.index:-1;
}

function reviewDefinitionForHeader(header) {
  const normalized=normalizeHeader(header);
  if (!normalized) return null;
  let best=null;
  for (const definition of REVIEW_FIELD_DEFINITIONS) {
    if (definition.key==="rootCauseRelated" && normalized.startsWith("root cause in ")) best={definition,score:90};
    for (const alias of definition.aliases) {
      const normalizedAlias=normalizeHeader(alias);
      let score=0;
      if (normalized===normalizedAlias) score=100;
      else if (normalizedAlias.length>=6 && normalized.endsWith(` ${normalizedAlias}`)) score=80;
      else if (normalizedAlias.length>=8 && normalized.startsWith(`${normalizedAlias} `)) score=70;
      if (score>(best?.score||0)) best={definition,score};
    }
  }
  return best?.definition||null;
}

function detectReviewHeaderRow(matrix) {
  let best={index:-1,mapped:0,filled:0};
  matrix.slice(0,12).forEach((row,index)=>{
    const headers=(row||[]).map(displayCellValue).filter(Boolean);
    const mapped=new Set(headers.map(header=>reviewDefinitionForHeader(header)?.key).filter(Boolean)).size;
    if (mapped>best.mapped || (mapped===best.mapped && headers.length>best.filled)) {
      best={index,mapped,filled:headers.length};
    }
  });
  return best.index;
}

function reviewProfileFromMatrix(name,matrix) {
  const headerRow=detectReviewHeaderRow(matrix);
  const headers=headerRow>=0?(matrix[headerRow]||[]):[];
  const columns=[];
  const seenColumns=new Set();
  headers.forEach((value,index)=>{
    const originalHeader=displayCellValue(value);
    if (!originalHeader) return;
    const definition=reviewDefinitionForHeader(originalHeader);
    const signature=`${normalizeHeader(originalHeader)}|${definition?.key||""}`;
    if (seenColumns.has(signature)) return;
    seenColumns.add(signature);
    const header=definition?.key==="rootCauseRelated"?definition.label:originalHeader;
    columns.push({
      id:`${name}-${headerRow}-${index}`,
      header,
      field:definition?.key||"",
      supported:Boolean(definition),
      essential:Boolean(definition?.essential)
    });
  });
  return {name,headerRow,columns,isDefault:false};
}

function reviewProfilesFromWorkbook(workbook) {
  return workbook.worksheets.map(worksheet=>reviewProfileFromMatrix(worksheet.name,worksheetMatrix(worksheet)));
}

function reviewProfilesFromReferenceSheets(sheets) {
  return Object.entries(sheets).map(([name,rows])=>{
    const matrix=[];
    for (let index=0;index<rows.length;index++) matrix.push(rows[index]||[]);
    return reviewProfileFromMatrix(name,matrix);
  });
}

function choosePreferredReviewProfile(profiles) {
  const preferred=["Complaint Information","Combined Overview","Complaint Summary","Final Reports","CN95"];
  for (const name of preferred) {
    const profile=profiles.find(item=>item.name.toLowerCase()===name.toLowerCase() && item.columns.some(column=>column.supported));
    if (profile) return profile.name;
  }
  return profiles
    .filter(profile=>!profile.isDefault)
    .sort((a,b)=>b.columns.filter(column=>column.supported).length-a.columns.filter(column=>column.supported).length)[0]?.name
    || DEFAULT_REVIEW_PROFILE.name;
}

function setReviewProfiles(profiles=[]) {
  const hasWorkbookProfiles=profiles.length>0;
  reviewProfiles=[DEFAULT_REVIEW_PROFILE,...(hasWorkbookProfiles?profiles:BUILT_IN_REVIEW_PROFILES)];
  reviewSelections.clear();
  for (const profile of reviewProfiles) {
    reviewSelections.set(profile.name,new Set(profile.columns.filter(column=>column.supported && (!profile.isDefault || column.essential)).map(column=>column.id)));
  }
  activeReviewProfileName=hasWorkbookProfiles?choosePreferredReviewProfile(reviewProfiles):DEFAULT_REVIEW_PROFILE.name;
  renderReviewControls();
  renderRecords();
}

function currentReviewProfile() {
  return reviewProfiles.find(profile=>profile.name===activeReviewProfileName)||DEFAULT_REVIEW_PROFILE;
}

function activeReviewLabels() {
  const profile=currentReviewProfile();
  const selected=reviewSelections.get(profile.name)||new Set();
  const labels=new Map();
  for (const column of profile.columns) {
    if (column.supported && selected.has(column.id) && !labels.has(column.field)) labels.set(column.field,column.header);
  }
  return labels;
}

function renderReviewControls() {
  const select=$("reviewSheetSelect");
  if (!select) return;
  select.innerHTML=reviewProfiles.map(profile=>`<option value="${esc(profile.name)}"${profile.name===activeReviewProfileName?" selected":""}>${esc(profile.name)}</option>`).join("");
  renderReviewColumnOptions();
}

function renderReviewColumnOptions() {
  const profile=currentReviewProfile();
  const selected=reviewSelections.get(profile.name)||new Set();
  const target=$("reviewColumnOptions");
  if (!target) return;
  target.innerHTML=profile.columns.length
    ? profile.columns.map(column=>column.supported
      ? `<label><input type="checkbox" data-review-column="${esc(column.id)}"${selected.has(column.id)?" checked":""} /> <span>${esc(column.header)}</span></label>`
      : `<label class="unsupported"><input type="checkbox" disabled /> <span>${esc(column.header)}<small>Not available from report</small></span></label>`
    ).join("")
    : `<p class="hint">No column header row could be identified in this worksheet.</p>`;
  const supported=profile.columns.filter(column=>column.supported).length;
  const unsupported=profile.columns.length-supported;
  $("reviewFilterStatus").className=supported?"status good":"status";
  $("reviewFilterStatus").textContent=profile.isDefault
    ? "Use this compact view, choose a built-in worksheet group, or upload Excel to load its exact worksheet names and columns."
    : `${profile.isBuiltIn?"Built-in group · ":""}${profile.name}: ${supported} worksheet column${supported===1?"":"s"} matched to report fields${unsupported?`; ${unsupported} worksheet column${unsupported===1?" is":"s are"} not available directly from the report`:""}.`;
  target.querySelectorAll("[data-review-column]").forEach(input=>{
    input.addEventListener("change",()=>{
      const next=reviewSelections.get(profile.name)||new Set();
      if (input.checked) next.add(input.dataset.reviewColumn);
      else next.delete(input.dataset.reviewColumn);
      reviewSelections.set(profile.name,next);
      syncRecordsFromDom();
      renderRecords();
    });
  });
}

function updateReviewSelection(mode) {
  const profile=currentReviewProfile();
  const next=new Set();
  for (const column of profile.columns) {
    if (!column.supported) continue;
    if (mode==="all" || (mode==="essential" && column.essential)) next.add(column.id);
  }
  reviewSelections.set(profile.name,next);
  renderReviewColumnOptions();
  syncRecordsFromDom();
  renderRecords();
}

function reviewFieldPresentation(name,defaultLabel) {
  const labels=activeReviewLabels();
  return {visible:labels.has(name),label:labels.get(name)||defaultLabel};
}

function objectsFromMatrix(matrix, headerIndex) {
  if (headerIndex<0) return [];
  const headers=(matrix[headerIndex]||[]).map(displayCellValue);
  const rows=[];
  for (let r=headerIndex+1;r<matrix.length;r++) {
    const values=matrix[r]||[];
    if (!values.some(value=>displayCellValue(value))) continue;
    const row={};
    headers.forEach((header,index)=>{if(header) row[header]=displayCellValue(values[index]);});
    rows.push(row);
  }
  return rows;
}

const SUMMARY_ALIASES = {
  complaintNo:["complaint notification","complaint number","complaint no","comp","complaint","notification","complaint #"],
  lot:["lot","lots","lot no","lot number","batch","batch no","batch number"],
  customer:["customer company","end customer","customer","company name","company"],
  productFamily:["product family","membrane type","product"],
  materialNo:["material no","material number","article no","article number"],
  problem:["problem","problems","reported symptom","complaint symptom","issue description","formal issue description","customer reported failure","reason"],
  standardizedSymptoms:["standardized symptoms","standardized symptom","standardized symptom s"],
  problemType:["problem type"],
  result:["result status","result","complaint status","claim decision","final result","final assessment root cause","final scope decision","conclusion"],
  tests:["tests assays applied","tests performed","test performed","assays applied","standard test","test","assay"],
  rollsImplicated:["rolls implicated","amount frs","implicated units","units implicated","number of roll implicated"],
  samplesReceived:["samples received","samples received qty","sample received","number of sample received"],
  registeredDate:["complaint registered date","registered date","registration date"],
  reportDate:["report date","final report date"],
  days:["days","elapsed days"]
};

function valueByAliases(row, aliases) {
  const entries=Object.entries(row).map(([header,value])=>[normalizeHeader(header),displayCellValue(value)]);
  for (const alias of aliases) {
    const exact=entries.find(([header])=>header===alias);
    if (exact?.[1]) return exact[1];
  }
  for (const alias of aliases.filter(value=>value.length>=6)) {
    const suffix=entries.find(([header])=>header.endsWith(` ${alias}`));
    if (suffix?.[1]) return suffix[1];
  }
  return "";
}

function summaryRowFromObject(row, sourceSheet) {
  const result={sourceSheets:sourceSheet};
  for (const [key,aliases] of Object.entries(SUMMARY_ALIASES)) result[key]=valueByAliases(row,aliases);
  result.productFamily=normalizedProductFamily(result.materialNo,result.productFamily);
  result.customer=canonicalCustomerName(result.customer);
  if (/^claim\s+accepted$/i.test(result.result)) result.result="Confirmed";
  if (/^claim\s+not\s+accepted$/i.test(result.result)) result.result="Not confirmed";
  for (const key of ["registeredDate","reportDate"]) {
    if (/^\d{5}(?:\.\d+)?$/.test(result[key])) {
      const date=new Date(Date.UTC(1899,11,30)+Number(result[key])*86400000);
      result[key]=date.toLocaleDateString("en-GB");
    }
  }
  return result;
}

function splitUniqueValues(value) {
  return String(value||"").split(/\s*;\s*|\n+/).map(item=>item.trim()).filter(Boolean);
}

function mergeSummaryValue(current,incoming) {
  const values=[...splitUniqueValues(current),...splitUniqueValues(incoming)];
  return [...new Set(values.map(value=>value.trim()).filter(Boolean))].join("; ");
}

function mergeSummaryRows(target,incoming) {
  for (const key of Object.keys(incoming)) {
    if (key==="lotComplaintCount") continue;
    if (key==="complaintNo" && target[key] && normalizeComplaintId(target[key])===normalizeComplaintId(incoming[key])) {
      const values=[target[key],incoming[key]].filter(Boolean);
      const compStyle=values.find(value=>/^comp/i.test(value));
      target[key]=compStyle||values.sort((a,b)=>String(a).length-String(b).length)[0];
      continue;
    }
    if (["registeredDate","reportDate","days"].includes(key) && target[key]) continue;
    target[key]=mergeSummaryValue(target[key],incoming[key]);
  }
  return target;
}

function reconcileCustomerNames(dataset) {
  const allNames=[...new Set(dataset.flatMap(row=>splitUniqueValues(row.customer)).map(canonicalCustomerName).filter(Boolean))];
  const detailed=allNames.filter(value=>value.length>=14);
  for (const row of dataset) {
    const reconciled=splitUniqueValues(row.customer).map(value=>{
      const current=canonicalCustomerName(value);
      const firstToken=normalizeHeader(current).split(" ").find(token=>token.length>=4&&!/^(?:hangzhou|shenzhen|guangzhou|nanjing|anhui|sichuan|qingdao|beijing|company|biotech|technology)$/.test(token))||"";
      if (!firstToken) return current;
      const candidates=detailed.filter(candidate=>normalizeHeader(candidate).split(" ").includes(firstToken));
      return candidates.sort((a,b)=>b.length-a.length)[0]||current;
    });
    row.customer=[...new Set(reconciled.filter(Boolean))].join("; ");
  }
}

function shouldIncludeSummaryRow(row,sourceSheet) {
  if (row.complaintNo) return true;
  if (/test|evidence|guide|simple|summary/i.test(sourceSheet)) return false;
  return Boolean(row.lot && (row.problem || row.result || row.customer));
}

function complaintKey(row,index) {
  const complaint=normalizeComplaintId(row.complaintNo);
  if (complaint) return `complaint:${complaint}`;
  const lot=normalizeId(row.lot), problem=normalizeId(row.problem), result=normalizeId(row.result);
  return lot||problem?`unidentified:${lot}|${problem}|${result}`:`row:${index}`;
}

function recordToSummaryRow(record) {
  const category=recordToCategoryRow(record);
  return {
    complaintNo:displayCellValue(category["Complaint / Notification"]),
    lot:displayCellValue(category["Lot"]),
    customer:canonicalCustomerName(displayCellValue(category["Customer Company"])),
    productFamily:displayCellValue(category["Product Family"]),
    materialNo:displayCellValue(category["Material No."]),
    problem:displayCellValue(category["Problem"]||category["Customer Reported Failure"]),
    standardizedSymptoms:displayCellValue(category["Standardized Symptom(s)"]),
    problemType:displayCellValue(category["Problem Type"]),
    result:displayCellValue(category["Result / Status"]||category["Final Assessment / Root Cause"]),
    tests:displayCellValue(category["Tests / Assays Applied"]),
    rollsImplicated:displayCellValue(category["Rolls Implicated"]),
    samplesReceived:displayCellValue(category["Samples Received"]),
    registeredDate:displayCellValue(category["Complaint Registered Date"]),
    reportDate:displayCellValue(category["Report Date"]),
    days:displayCellValue(category["Days"]),
    sourceSheets:"Current extraction"
  };
}

function lotTokens(value) {
  return String(value||"").split(/\s*;\s*|\s*,\s*|\n+/).map(item=>item.trim()).filter(Boolean);
}

async function appendWorkbookComplaintRows(raw,buffer,fileName) {
  const sourceLabel=sheetName=>fileName?`${fileName} › ${sheetName}`:sheetName;
  const appendRows=(rows,sheetName)=>{
    let carry={};
    const isEvidence=/test|evidence/i.test(sheetName);
    for (const row of rows) {
      const normalized=summaryRowFromObject(row,sourceLabel(sheetName));
      if (isEvidence) {
        const contextKeys=["complaintNo","lot","customer","productFamily","materialNo","problem","result","rollsImplicated","samplesReceived","registeredDate","reportDate"];
        if (normalized.complaintNo) carry={};
        for (const key of contextKeys) {
          if (normalized[key]) carry[key]=normalized[key];
          else if (carry[key]) normalized[key]=carry[key];
        }
      }
      if (shouldIncludeSummaryRow(normalized,sheetName)) raw.push(normalized);
    }
  };
  try {
    const wb=await loadWorkbook(buffer.slice(0));
    for (const ws of wb.worksheets) {
      if (ws.name===SUMMARY_SHEET || /syndrome.*(?:guide|simple)/i.test(ws.name)) continue;
      const matrix=worksheetMatrix(ws);
      const headerIndex=detectComplaintHeaderRow(matrix);
      appendRows(objectsFromMatrix(matrix,headerIndex),ws.name);
    }
  } catch (_) {
    const sheets=await readReferenceWorkbook(buffer.slice(0));
    for (const [sheetName,matrixWithBlank] of Object.entries(sheets)) {
      if (sheetName===SUMMARY_SHEET || /syndrome.*(?:guide|simple)/i.test(sheetName)) continue;
      const matrix=[];
      for (let i=0;i<matrixWithBlank.length;i++) matrix.push(matrixWithBlank[i]||[]);
      const headerIndex=detectComplaintHeaderRow(matrix);
      appendRows(objectsFromMatrix(matrix,headerIndex),sheetName);
    }
  }
}

async function extractSummarySourceRecords(source) {
  if (source.records) return source.records;
  const expanded=[];
  if (/\.zip$/i.test(source.name)) {
    const zip=await JSZip.loadAsync(source.buffer.slice(0));
    for (const [name,entry] of Object.entries(zip.files)) {
      if (!entry.dir && /\.(pdf|msg)$/i.test(name)) expanded.push({name:name.split("/").pop(),buffer:await entry.async("arraybuffer")});
    }
  } else expanded.push({name:source.name,buffer:source.buffer});
  source.records=[];
  source.errors=[];
  for (const item of expanded) {
    try { source.records.push(...await extractMany(item.name,item.buffer)); }
    catch(err) { source.errors.push(`${item.name}: ${err.message}`); }
  }
  return source.records;
}

async function collectComplaintDataset() {
  const raw=[];
  if (workbookBuffer) await appendWorkbookComplaintRows(raw,workbookBuffer,workbookFileName);
  for (const source of summarySources.filter(item=>item.kind==="workbook")) {
    await appendWorkbookComplaintRows(raw,source.buffer,source.name);
  }
  for (const source of summarySources.filter(item=>item.kind==="report")) {
    raw.push(...(await extractSummarySourceRecords(source)).map(record=>({...recordToSummaryRow(record),sourceSheets:source.name})));
  }
  syncRecordsFromDom();
  raw.push(...records.map(recordToSummaryRow));

  const merged=new Map();
  raw.forEach((row,index)=>{
    const key=complaintKey(row,index);
    if (!merged.has(key)) merged.set(key,{...row});
    else mergeSummaryRows(merged.get(key),row);
  });
  const dataset=[...merged.values()];
  dataset.forEach(row=>{
    row.productFamily=normalizedProductFamily(row.materialNo,row.productFamily);
    row.customer=canonicalCustomerName(row.customer);
    const calculated=daysBetweenDates(row.registeredDate,row.reportDate);
    if (!/^\d+$/.test(String(row.days||"")) && calculated!=="") row.days=String(calculated);
  });
  reconcileCustomerNames(dataset);
  const lotCounts=new Map();
  dataset.forEach((row,index)=>{
    const uniqueComplaint=normalizeComplaintId(row.complaintNo);
    if (!uniqueComplaint) return;
    for (const lot of new Set(lotTokens(row.lot))) {
      if (!lotCounts.has(lot)) lotCounts.set(lot,new Set());
      lotCounts.get(lot).add(uniqueComplaint);
    }
  });
  for (const row of dataset) {
    const counts=lotTokens(row.lot).map(lot=>({lot,count:lotCounts.get(lot)?.size||0}));
    row.lotComplaintCount=counts.length===1?String(counts[0].count):counts.map(item=>`${item.lot}: ${item.count}`).join("; ");
  }
  return dataset.sort((a,b)=>String(a.complaintNo||"").localeCompare(String(b.complaintNo||""),undefined,{numeric:true})||String(a.lot||"").localeCompare(String(b.lot||""),undefined,{numeric:true}));
}

function selectedSummaryFields() {
  const selected=new Set([...document.querySelectorAll("[data-summary-field]:checked")].map(input=>input.value));
  return SUMMARY_FIELDS.filter(field=>selected.has(field.key));
}

function renderLotOnlyTable(rows) {
  const filterControls=`<div class="lot-family-filter" role="group" aria-label="CN types to display">
    <strong>Show CN types:</strong>
    ${LOT_FAMILY_ORDER.map(family=>`<label><input type="checkbox" data-lot-family-filter value="${family}" ${selectedLotFamilies.has(family)?"checked":""} /> ${family}</label>`).join("")}
  </div>`;
  const filteredRows=rows.filter(row=>{
    if (selectedLotFamilies.size===LOT_FAMILY_ORDER.length) return true;
    const families=splitUniqueValues(row.productFamily);
    return families.some(family=>selectedLotFamilies.has(family));
  });
  const lots=new Map();
  for (const row of filteredRows) {
    for (const lot of new Set(lotTokens(row.lot))) {
      if (!lots.has(lot)) lots.set(lot,[]);
      lots.get(lot).push(row);
    }
  }
  if (!lots.size) return `${filterControls}<p class="status good">No complaint lots match the selected CN types.</p>`;
  const groups=[...lots.entries()].sort((a,b)=>String(a[0]).localeCompare(String(b[0]),undefined,{numeric:true}));
  return `${filterControls}<table class="summary-table lot-detail-table"><thead><tr>
    <th>Lot (complaint count)</th><th>Product Family</th><th>Complaint Number</th><th>End Customer</th><th>Reason / Problem</th><th>Final Result / Status</th>
  </tr></thead><tbody>${groups.map(([lot,complaints],groupIndex)=>complaints.map((row,index)=>`<tr class="lot-tone-${groupIndex%4}">
    ${index===0?`<td class="lot-group-cell" rowspan="${complaints.length}"><strong>Lot (${complaints.length})</strong><span>${esc(lot)}</span></td>`:""}
    <td class="lot-family-cell">${esc(row.productFamily||"Not specified")}</td>
    <td>${esc(row.complaintNo||"Complaint number unavailable")}</td>
    <td>${esc(row.customer||"")}</td>
    <td>${esc(row.problem||"")}</td>
    <td>${esc(row.result||"")}</td>
  </tr>`).join("")).join("")}</tbody></table>`;
}

function renderDatasetTable(rows,fields) {
  if (!rows.length) return `<p class="status good">No matching complaint information found.</p>`;
  if (!fields.length) return `<p class="status bad">Tick at least one information field to display.</p>`;
  if (fields.length===1 && fields[0].key==="lot") return renderLotOnlyTable(rows);
  return `<table class="summary-table"><thead><tr>${fields.map(field=>`<th>${esc(field.label)}</th>`).join("")}</tr></thead><tbody>${rows.map(row=>`<tr>${fields.map(field=>`<td>${esc(row[field.key]||"")}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}

function renderSummaryMetrics(dataset) {
  const lots=new Set(dataset.flatMap(row=>lotTokens(row.lot)));
  const customers=new Set(dataset.flatMap(row=>splitUniqueValues(row.customer)));
  const tested=dataset.filter(row=>row.tests).length;
  const rollTotal=dataset.reduce((sum,row)=>sum+(Number(String(row.rollsImplicated||"").match(/\d+/)?.[0])||0),0);
  const numericDays=dataset.map(row=>Number(row.days)).filter(value=>Number.isFinite(value)&&value>=0);
  const averageDays=numericDays.length?Math.round(numericDays.reduce((sum,value)=>sum+value,0)/numericDays.length):0;
  $("summaryMetrics").innerHTML=`
    <div class="metric"><strong>${dataset.length}</strong><span>unique complaint${dataset.length===1?"":"s"}</span></div>
    <div class="metric"><strong>${lots.size}</strong><span>complaint lot${lots.size===1?"":"s"}</span></div>
    <div class="metric"><strong>${customers.size}</strong><span>end customer${customers.size===1?"":"s"}</span></div>
    <div class="metric"><strong>${tested}</strong><span>complaint${tested===1?"":"s"} with tests listed</span></div>
    <div class="metric"><strong>${rollTotal}</strong><span>implicated roll${rollTotal===1?"":"s"} recorded</span></div>
    <div class="metric"><strong>${averageDays||"—"}</strong><span>average days to report${numericDays.length?` (${numericDays.length} cases)`:""}</span></div>`;
}

function topCountEntries(values,limit=8) {
  const counts=new Map();
  for (const value of values.flatMap(item=>splitUniqueValues(item))) {
    if (!value || /^review required$/i.test(value)) continue;
    counts.set(value,(counts.get(value)||0)+1);
  }
  return [...counts.entries()].sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0])).slice(0,limit);
}

function overviewList(title,entries,empty="No information recorded") {
  return `<section class="overview-card"><h3>${esc(title)}</h3>${entries.length
    ?`<table><thead><tr><th>Item</th><th>Complaints</th></tr></thead><tbody>${entries.map(([label,count])=>`<tr><td>${esc(label)}</td><td>${count}</td></tr>`).join("")}</tbody></table>`
    :`<p class="hint">${esc(empty)}</p>`}</section>`;
}

function renderSummaryOverview(dataset) {
  const target=$("summaryOverview");
  if (!target) return;
  if (!dataset.length) { target.innerHTML=""; return; }
  const lotEntries=topCountEntries(dataset.flatMap(row=>lotTokens(row.lot))).filter(([,count])=>count>1);
  const symptomEntries=topCountEntries(dataset.flatMap(row=>symptomTokensForRow(row)));
  const customerEntries=topCountEntries(dataset.map(row=>row.customer));
  const testEntries=topCountEntries(dataset.map(row=>row.tests));
  const required=["complaintNo","lot","customer","productFamily","problem","result"];
  const quality=required.map(key=>({key,label:SUMMARY_FIELDS.find(field=>field.key===key)?.label||key,count:dataset.filter(row=>row[key]).length}));
  const testCoverage=dataset.filter(row=>row.tests).length;
  const daysCoverage=dataset.filter(row=>/^\d+$/.test(String(row.days||""))).length;
  target.innerHTML=`<div class="overview-heading"><h2>Complaint Overview</h2><p class="hint">The tables below summarize recurring lots, symptoms, customers, tests and data completeness. Select columns below for case-level detail.</p></div>
    <div class="overview-grid">
      ${overviewList("Recurring lots",lotEntries,"No repeated lots found")}
      ${overviewList("Frequent symptoms",symptomEntries)}
      ${overviewList("Customers with most complaints",customerEntries)}
      ${overviewList("Most-used tests",testEntries,"No tests recorded")}
      <section class="overview-card"><h3>Data completeness</h3><table><thead><tr><th>Field</th><th>Coverage</th></tr></thead><tbody>
        ${quality.map(item=>`<tr><td>${esc(item.label)}</td><td>${item.count}/${dataset.length}</td></tr>`).join("")}
        <tr><td>Tests performed</td><td>${testCoverage}/${dataset.length}</td></tr>
        <tr><td>Numeric days to report</td><td>${daysCoverage}/${dataset.length}</td></tr>
      </tbody></table></section>
    </div>`;
}

async function refreshSummary() {
  $("summaryResult").innerHTML=`<p class="status">Reading workbook and building summary...</p>`;
  try {
    summaryDataset=await collectComplaintDataset();
    renderSummaryMetrics(summaryDataset);
    renderSummaryOverview(summaryDataset);
    $("summaryResult").innerHTML=renderDatasetTable(summaryDataset,selectedSummaryFields());
    $("summaryExcelStatus").className="status good";
    const sourceCount=summarySources.length+(workbookBuffer?1:0);
    const errorCount=summarySources.reduce((sum,source)=>sum+(source.errors?.length||0),0);
    $("summaryExcelStatus").textContent=`Summary ready: ${summaryDataset.length} unique complaint(s) from ${sourceCount||"current"} source file(s)${errorCount?`; ${errorCount} report(s) could not be read`:""}.`;
  } catch(err) {
    $("summaryExcelStatus").className="status bad";
    $("summaryExcelStatus").textContent=`Could not create summary: ${err.message}`;
    $("summaryResult").innerHTML="";
  }
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
  const widthByHeader={
    "Source Group":20,"Lot":13,"Product Family":15,"Membrane Type":20,"Customer Company":28,"Rolls Implicated":16,
    "Samples Received":16,"Final Roll(s)":18,"Master Roll(s)":18,"MR-FR Area(s)":28,
    "Zone(s)":20,"MR-FR (s)":30,
    "Complaint / Notification":24,"Formal Issue Description":28,"Problem":38,"Customer Reported Failure":42,"Tests / Assays Applied":38,
    "Standardized Symptom(s)":36,"Problem Type":30,"LFA Relevance":22,
    "Result / Status":20,"Criticality":14,"Failure Reproduced?":18,"Root Cause in Process?":20,
    "Product Description":36,"Coordinator":20,"Final Assessment / Root Cause":48,"Final Scope / Decision":38,
    "Data Quality / Notes":38,"Material No.":24,"Complaint Registered Date":20,"Report Date":16,"Days":12,"Sample Source":22,"Sample ID":28,
    "Standard Test":24,"Standard Purpose":28,"Standard Method":44,"Result (Source)":48,"Outcome":24,
    "Within Spec?":14,"Issue Observed?":16,"Source Page":12,"Case-specific Conditions / Source Detail":42,
    "Complaint Count":18,"Symptom":38,"Symptom Count":16,"Complaint Numbers":42,"Complaint Number":24,
    "End Customer":28,"Tests Performed":38,"Final Result / Status":22,"Registered Date":18,"Date Registered":18,
    "Lot Number":15,"Structured Test Evidence":72,"Conclusion of Root Cause Analysis":64,
    "Purpose":28,"Method":44,"Result":48,"Conditions":42,
    "Source File":34
  };
  const headers=ws.getRow(1).values.slice(1).map(h=>String(h||""));
  const widths=headers.map(h=>widthByHeader[h]||20);
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
  for (const dateHeader of ["Complaint Registered Date","Registered Date","Date Registered","Report Date"]) {
    const index=headers.indexOf(dateHeader)+1;
    if (index>0) for (let r=2;r<=ws.rowCount;r++) ws.getCell(r,index).numFmt=ws.name===REVIEW_OVERVIEW_SHEET?"yyyy-mm-dd":"dd.mm.yyyy";
  }
  ws.autoFilter={from:{row:1,column:1},to:{row:ws.rowCount,column:ws.columnCount}};
}

function writeSheet(ws, headers, rows, hasSource=false) {
  clearSheet(ws);
  ws.addRow(headers);
  for (const row of rows) ws.addRow(headers.map(h=>row[h] ?? ""));
  formatSheet(ws,hasSource);
  if (headers.includes("Structured Test Evidence") || headers.includes("Conclusion of Root Cause Analysis")) {
    for (let row=2;row<=ws.rowCount;row++) ws.getRow(row).height=96;
  }
}

function normalizeId(v) { return String(v||"").trim().toLowerCase(); }

function normalizeComplaintId(v) {
  const text=normalizeId(v).replace(/[–—]/g,"-");
  const match=text.match(/^(?:comp(?:laint)?[-\s]*)?0*(\d+)$/i);
  return match?String(Number(match[1])):text;
}

function hasManagedHeaders(ws, expectedHeaders) {
  if (!ws || ws.rowCount<1) return false;
  const actual=ws.getRow(1).values.slice(1).map(normalizeHeader);
  return expectedHeaders.every(header=>actual.includes(normalizeHeader(header)));
}

function selectedExportSheets() {
  return new Set([...document.querySelectorAll("[data-export-sheet]:checked")].map(input=>input.value).filter(name=>MANAGED_SHEETS.includes(name)));
}

function removeSheetIfPresent(wb,name) {
  const sheet=wb.getWorksheet(name);
  if (sheet) wb.removeWorksheet(sheet.id);
}

function recordToCategoryRow(r) {
  const fam = normalizedProductFamily(r.materialNo,r.productFamily);
  const classification=complaintClassification(r.problem,r.customerReportedFailure);
  const registeredDate=parseFlexibleDate(r.complaintRegisteredDate)||r.complaintRegisteredDate||"";
  const reportDate=parseFlexibleDate(r.reportDate)||r.reportDate||"";
  const calculatedDays=daysBetweenDates(r.complaintRegisteredDate,r.reportDate);
  return {
    "Lot":r.lot||"", "Product Family":fam,
    "Membrane Type":membraneType(r.materialNo,r.productDescription)||r.membraneType||"",
    "Customer Company":r.customerCompany||"",
    "Rolls Implicated":r.rollsImplicated||"", "Samples Received":r.samplesReceived||"",
    "Zone(s)":r.zones||"",
    "Final Roll(s)":r.finalRolls||"", "Master Roll(s)":r.masterRolls||"",
    "MR-FR Area(s)":r.mrfrAreas||"", "MR-FR (s)":r.mrfrCombined||r.mrfrAreas||"",
    "Complaint / Notification":r.complaintNo||"",
    "Formal Issue Description":r.formalProblem||"", "Problem":r.problem||"",
    "Customer Reported Failure":r.customerReportedFailure||"",
    "Standardized Symptom(s)":classification.standardizedSymptoms,
    "Problem Type":classification.problemTypes,"LFA Relevance":classification.lfaRelevance,
    "Tests / Assays Applied":r.assaysApplied||"", "Result / Status":r.resultStatus||"",
    "Criticality":r.criticality||"", "Failure Reproduced?":r.failureReproduced||"",
    "Root Cause in Process?":r.rootCauseRelated||"",
    "Product Description":r.productDescription||"", "Coordinator":r.coordinator||"",
    "Similar Events Same Category?":r.similarEvents||"",
    "Containment Necessary?":r.containmentNecessary||"",
    "Corrective / Preventive Action Necessary?":r.correctiveActionNecessary||"",
    "Root Cause Analysis Conclusion":r.rootCauseConclusion||"",
    "Problem Description Check":r.problemValidation||"",
    "Final Assessment / Root Cause":r.finalAssessment||"", "Final Scope / Decision":r.finalScope||"",
    "Sample Details":r.sampleDetails||"", "Data Quality / Notes":r.warnings||"",
    "Material No.":r.materialNo||"", "Complaint Registered Date":registeredDate,
    "Report Date":reportDate,"Days":calculatedDays===""?(r.daysToReport||""):calculatedDays
  };
}

function recordToEvidenceRows(r) {
  const fam=normalizedProductFamily(r.materialNo,r.productFamily);
  return (r.testEvidence||[]).map(t=>({
    "Complaint / Notification":r.complaintNo||"", "Lot":r.lot||"", "Material No.":r.materialNo||"",
    "Product Family":fam, "Units Implicated":r.rollsImplicated||"", "Samples Received":r.samplesReceived||"",
    "Problem":r.problem||"", "Customer Reported Failure":r.customerReportedFailure||"",
    "Complaint Status":r.resultStatus||"", "Sample Source":t.sampleSource||"", "Sample ID":t.sampleId||"",
    "Standard Test":t.name||"", "Standard Purpose":t.purpose||"", "Standard Method":testMethodBulletText(t.method),
    "Result (Source)":t.result||"", "Outcome":t.outcome||"", "Within Spec?":t.withinSpec||"",
    "Issue Observed?":t.issueObserved||"", "Source Page":t.sourcePage||"",
    "Case-specific Conditions / Source Detail":t.conditions||"", "Source File":r.sourceFile||""
  }));
}

function recordToOverviewRow(record) {
  const row=recordToCategoryRow(record);
  return {
    "Source Group":record.sourceGroup||"Final Reports",
    "Complaint Number":row["Complaint / Notification"],
    "Lot Number":row["Lot"],
    "End Customer":row["Customer Company"],
    "Final Result / Status":row["Result / Status"],
    "Standardized Symptom(s)":row["Standardized Symptom(s)"],
    "Date Registered":parseFlexibleDate(row["Complaint Registered Date"])||"",
    "Report Date":parseFlexibleDate(row["Report Date"])||"",
    "Days":row["Days"],
    "Membrane Type":row["Membrane Type"],
    "Material No.":row["Material No."]
  };
}

function recordToInvestigationRow(record) {
  const row=recordToCategoryRow(record);
  return {
    "Complaint Number":row["Complaint / Notification"],
    "Lot Number":row["Lot"],
    "End Customer":row["Customer Company"],
    "Standardized Symptom(s)":row["Standardized Symptom(s)"],
    "Customer Reported Failure":row["Customer Reported Failure"],
    "Tests Performed":row["Tests / Assays Applied"],
    "MR-FR Area(s)":row["MR-FR Area(s)"],
    "Rolls Implicated":row["Rolls Implicated"],
    "Samples Received":row["Samples Received"]
  };
}

function recordToRootCauseRows(record) {
  const row=recordToCategoryRow(record);
  const tests=record.testEvidence?.length?record.testEvidence:[{}];
  return tests.map(test=>({
    "Complaint Number":row["Complaint / Notification"],"Lot Number":row["Lot"],"End Customer":row["Customer Company"],
    "Standard Test":test.name||"","Sample Source":test.sampleSource||"","Sample ID":test.sampleId||"",
    "Purpose":test.purpose||"","Method":testMethodBulletText(test.method),"Result":test.result||"","Outcome":test.outcome||"",
    "Within Spec?":test.withinSpec||"","Issue Observed?":test.issueObserved||"","Source Page":test.sourcePage||"",
    "Conditions":test.conditions||"","Conclusion of Root Cause Analysis":row["Root Cause Analysis Conclusion"]
  }));
}

function mergeOrganizedRows(existingRows,newRows) {
  const replacementIds=new Set(newRows.map(row=>normalizeComplaintId(row["Complaint Number"])).filter(Boolean));
  const kept=existingRows.filter(row=>!replacementIds.has(normalizeComplaintId(row["Complaint Number"])));
  return [...kept,...newRows].sort((a,b)=>String(a["Complaint Number"]||"").localeCompare(String(b["Complaint Number"]||""),undefined,{numeric:true}));
}

function updateOrganizedSheet(wb,selectedSheets,name,headers,newRows) {
  if (!selectedSheets.has(name)) {
    removeSheetIfPresent(wb,name);
    return;
  }
  const existing=sheetRows(wb.getWorksheet(name));
  writeSheet(ensureSheet(wb,name),headers,mergeOrganizedRows(existing,newRows),false);
}

function categoryRowsToComplaintSummary(categoryRows) {
  const merged=new Map();
  for (const group of CATEGORY_SHEETS) {
    for (const row of categoryRows[group]) {
      const complaint=displayCellValue(row["Complaint / Notification"]);
      if (!complaint) continue;
      const key=normalizeComplaintId(complaint);
      const item={
        "Complaint Number":complaint,
        "Lot":displayCellValue(row["Lot"]),
        "Product Family":displayCellValue(row["Product Family"]),
        "Material No.":displayCellValue(row["Material No."]),
        "Membrane Type":displayCellValue(row["Membrane Type"]),
        "End Customer":displayCellValue(row["Customer Company"]),
        "Problem":displayCellValue(row["Problem"]),
        "Customer Reported Failure":displayCellValue(row["Customer Reported Failure"]),
        "Standardized Symptom(s)":displayCellValue(row["Standardized Symptom(s)"]),
        "Problem Type":displayCellValue(row["Problem Type"]),
        "Tests Performed":displayCellValue(row["Tests / Assays Applied"]),
        "Final Result / Status":displayCellValue(row["Result / Status"]),
        "Rolls Implicated":displayCellValue(row["Rolls Implicated"]),
        "Samples Received":displayCellValue(row["Samples Received"]),
        "MR-FR(s)":displayCellValue(row["MR-FR (s)"]||row["MR-FR Area(s)"]),
        "Registered Date":row["Complaint Registered Date"]||"",
        "Report Date":row["Report Date"]||"",
        "Days":displayCellValue(row["Days"]),
        "Data Quality / Notes":displayCellValue(row["Data Quality / Notes"]),
        "Source Group":group
      };
      if (!merged.has(key)) merged.set(key,item);
      else {
        const current=merged.get(key);
        for (const header of COMPLAINT_SUMMARY_HEADERS) {
          if (!displayCellValue(current[header]) && displayCellValue(item[header])) current[header]=item[header];
        }
      }
    }
  }
  return [...merged.values()].sort((a,b)=>String(a["Complaint Number"]).localeCompare(String(b["Complaint Number"]),undefined,{numeric:true}));
}

function mergeExistingCategoryRow(existing,incoming) {
  if (!existing) return incoming;
  const merged={...existing};
  for (const [header,value] of Object.entries(incoming)) {
    if (value!=="" && value!==null && value!==undefined) merged[header]=value;
  }
  const oldCustomer=displayCellValue(existing["Customer Company"]);
  const newCustomer=displayCellValue(incoming["Customer Company"]);
  if (oldCustomer && (!newCustomer || oldCustomer.length>newCustomer.length+5)) merged["Customer Company"]=existing["Customer Company"];
  return merged;
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
  syncRecordsFromDom();
  if (!records.length) throw new Error("Extract at least one complaint report first.");
  const selectedSheets=selectedExportSheets();
  lastPreservedSheetNames=[];
  if (!selectedSheets.size) throw new Error("Tick at least one sheet to export.");
  let wb;
  if (!workbookBuffer) {
    wb=new ExcelJS.Workbook();
    workbookMode="new";
  } else {
    try {
      wb=await loadWorkbook(workbookBuffer.slice(0));
      workbookMode="standard";
    } catch (_) {
      wb=new ExcelJS.Workbook();
      workbookMode="reference-readonly";
    }
  }

  const categoryRows={};
  for (const name of CATEGORY_SHEETS) categoryRows[name]=sheetRows(wb.getWorksheet(name));

  for (const r of records) {
    const id=normalizeComplaintId(r.complaintNo);
    if (!id) continue;
    let existing=null;
    for (const name of CATEGORY_SHEETS) {
      const matched=categoryRows[name].find(x=>normalizeComplaintId(x["Complaint / Notification"])===id);
      if (matched && !existing) existing=matched;
      categoryRows[name]=categoryRows[name].filter(x=>normalizeComplaintId(x["Complaint / Notification"])!==id);
    }
    const group=CATEGORY_SHEETS.includes(r.sourceGroup)?r.sourceGroup:"Final Reports";
    categoryRows[group].push(mergeExistingCategoryRow(existing,recordToCategoryRow(r)));
  }

  for (const name of CATEGORY_SHEETS) {
    if (selectedSheets.has(name)) {
      const ws=ensureSheet(wb,name);
      writeSheet(ws,CATEGORY_HEADERS,sortCategory(categoryRows[name]),false);
    } else removeSheetIfPresent(wb,name);
  }

  let evidenceRows=sheetRows(wb.getWorksheet(EVIDENCE_SHEET));
  for (const r of records) {
    const id=normalizeComplaintId(r.complaintNo);
    if (!id) continue;
    evidenceRows=evidenceRows.filter(x=>normalizeComplaintId(x["Complaint / Notification"])!==id);
    evidenceRows.push(...recordToEvidenceRows(r));
  }
  if (selectedSheets.has(EVIDENCE_SHEET)) writeSheet(ensureSheet(wb,EVIDENCE_SHEET),EVIDENCE_HEADERS,evidenceRows,false);
  else removeSheetIfPresent(wb,EVIDENCE_SHEET);
  if (selectedSheets.has(COMPLAINT_SUMMARY_SHEET)) writeSheet(ensureSheet(wb,COMPLAINT_SUMMARY_SHEET),COMPLAINT_SUMMARY_HEADERS,categoryRowsToComplaintSummary(categoryRows),false);
  else removeSheetIfPresent(wb,COMPLAINT_SUMMARY_SHEET);
  if (selectedSheets.has(SUMMARY_SHEET)) writeSheet(ensureSheet(wb,SUMMARY_SHEET),SUMMARY_HEADERS,lotSymptomSummaryRows(categoryRows),false);
  else removeSheetIfPresent(wb,SUMMARY_SHEET);

  updateOrganizedSheet(wb,selectedSheets,REVIEW_OVERVIEW_SHEET,REVIEW_OVERVIEW_HEADERS,records.map(recordToOverviewRow));
  updateOrganizedSheet(wb,selectedSheets,REVIEW_INVESTIGATION_SHEET,REVIEW_INVESTIGATION_HEADERS,records.map(recordToInvestigationRow));
  updateOrganizedSheet(wb,selectedSheets,REVIEW_ROOT_CAUSE_SHEET,REVIEW_ROOT_CAUSE_HEADERS,records.flatMap(recordToRootCauseRows));

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
    if (selectedSheets.has(fam)) {
      const existing=wb.getWorksheet(fam);
      if (existing && !hasManagedHeaders(existing,FAMILY_HEADERS)) {
        lastPreservedSheetNames.push(fam);
      } else {
        const ws=existing||ensureSheet(wb,fam);
        writeSheet(ws,FAMILY_HEADERS,sortFamily(familyRows[fam]),true);
      }
    } else removeSheetIfPresent(wb,fam);
  }

  if (!wb.worksheets.length) throw new Error("The selected export would contain no worksheets.");
  lastBuiltSheetNames=wb.worksheets.map(sheet=>sheet.name);
  const out=await wb.xlsx.writeBuffer();
  return out;
}

async function searchHistory(lot) {
  const found=[];
  if (workbookBuffer) {
    try {
      const wb=await loadWorkbook(workbookBuffer.slice(0));
      for (const group of CATEGORY_SHEETS) {
        const ws=wb.getWorksheet(group);
        for (const row of sheetRows(ws)) {
          if (String(row["Lot"]||"").trim()===String(lot||"").trim()) found.push({"Source Group":group,...row});
        }
      }
    } catch (_) {
      const sheets=await readReferenceWorkbook(workbookBuffer);
      const configs=[
        {name:"Combined Overview",header:4,lot:"Batch",complaint:"Complaint / Notification",family:"Product Family",company:"Customer",problem:"Reported Symptom",status:"Complaint Status"},
        {name:"Case Conclusions",header:4,lot:"Lot No",complaint:"Complaint No",family:"Product",company:"",problem:"Issue Description",status:"Complaint Status"},
        {name:"Test Results",header:3,lot:"Lot No",complaint:"",family:"Product",company:"",problem:"Complaint Symptom",status:"Complaint Status"}
      ];
      for (const cfg of configs) {
        for (const row of objectsFromReferenceRows(sheets[cfg.name],cfg.header)) {
          if (String(row[cfg.lot]||"").trim()!==String(lot||"").trim()) continue;
          const complaintKey=cfg.complaint||Object.keys(row).find(h=>/Complaint.*No/i.test(h))||"";
          found.push({
            "Source Group":cfg.name,"Lot":row[cfg.lot]||"","Product Family":row[cfg.family]||"",
            "Customer Company":cfg.company?row[cfg.company]||"":"","Complaint / Notification":row[complaintKey]||"",
            "Problem":row[cfg.problem]||"","Result / Status":row[cfg.status]||""
          });
        }
      }
    }
  }

  syncRecordsFromDom();
  for (const r of records) {
    if (String(r.lot||"").trim()!==String(lot||"").trim()) continue;
    found.push({"Source Group":"Current extraction",...recordToCategoryRow(r)});
  }
  return deduplicateHistoryRows(found);
}

function deduplicateHistoryRows(rows) {
  const merged=new Map();
  rows.forEach((row,index)=>{
    const id=normalizeComplaintId(row["Complaint / Notification"]);
    const key=id?`id:${id}`:`row:${index}`;
    if (!merged.has(key)) {
      merged.set(key,{...row});
      return;
    }
    const current=merged.get(key);
    for (const [field,value] of Object.entries(row)) {
      if (field==="Source Group") {
        const groups=new Set(`${current[field]||""}; ${value||""}`.split(";").map(x=>x.trim()).filter(Boolean));
        current[field]=[...groups].join("; ");
      } else if (!current[field] && value) current[field]=value;
    }
  });
  return [...merged.values()];
}

function symptomLabels(row) {
  const text=cleanBlock(`${row["Problem"]||""} ${row["Customer Reported Failure"]||""}`)
    .toLowerCase().replace(/[_/]+/g," ");
  const labels=[];
  const themes=[
    ["Absorption / wetting",/\b(?:absorption|absorbency|wetting|wicking)\b/],
    ["Uneven or interrupted printing lines",/\b(?:uneven|interrupted|broken|discontinuous)\b[^.]{0,45}\b(?:print(?:ing|ed)?|line)s?\b|\b(?:print(?:ing|ed)?|line)s?\b[^.]{0,45}\b(?:uneven|interrupted|broken|discontinuous)\b/],
    ["Capillary flow",/\b(?:capillary|flow time|slow flow|fast flow|flow rate)\b/],
    ["Protein binding",/\b(?:protein binding|binding capacity|weak binding|low binding)\b/],
    ["Particles / contamination",/\b(?:particle|contaminat|foreign material|debris|dust)\w*\b/],
    ["Discoloration / staining",/\b(?:discolou?r|stain|spot|yellowing|colour variation|color variation)\w*\b/],
    ["Physical damage",/\b(?:tear|torn|hole|crack|crease|wrinkle|damage|dent)\w*\b/],
    ["Backing / delamination",/\b(?:delaminat|separat|backing|adhesion|peel)\w*\b/],
    ["Dimension / thickness",/\b(?:width|thickness|dimension|oversize|undersize)\w*\b/]
  ];
  for (const [label,re] of themes) if (re.test(text)) labels.push(label);
  if (labels.length) return labels;
  const fallback=text
    .replace(/^(?:performance|quality|product|functionality|functional)(?:\s+\w+)?\s+(?:issue|problem)\s*:?\s*/,"")
    .replace(/[^a-z0-9]+/g," ").trim().split(/\s+/).slice(0,8).join(" ");
  return fallback?[fallback.charAt(0).toUpperCase()+fallback.slice(1)]:["Unspecified symptom"];
}

function lotSymptomSummaryRows(categoryRows) {
  const unique=deduplicateHistoryRows(CATEGORY_SHEETS.flatMap(group=>(categoryRows[group]||[]).map(row=>({"Source Group":group,...row}))));
  const lots=new Map();
  unique.forEach((row,index)=>{
    const lot=String(row["Lot"]||"").trim();
    if (!lot) return;
    if (!lots.has(lot)) lots.set(lot,new Map());
    const complaint=row["Complaint / Notification"]||`Unnumbered complaint ${index+1}`;
    lots.get(lot).set(complaint,row);
  });
  const output=[];
  for (const [lot,complaints] of lots) {
    const symptomMap=new Map();
    for (const [complaint,row] of complaints) {
      for (const label of new Set(symptomLabels(row))) {
        if (!symptomMap.has(label)) symptomMap.set(label,new Set());
        symptomMap.get(label).add(complaint);
      }
    }
    for (const [label,ids] of [...symptomMap.entries()].sort((a,b)=>b[1].size-a[1].size||a[0].localeCompare(b[0]))) {
      output.push({
        "Lot":lot,"Complaint Count":complaints.size,"Symptom":label,
        "Symptom Count":ids.size,"Complaint Numbers":[...ids].join("; ")
      });
    }
  }
  return output.sort((a,b)=>String(a["Lot"]).localeCompare(String(b["Lot"]),undefined,{numeric:true})||b["Symptom Count"]-a["Symptom Count"]);
}

function historyTable(rows) {
  if (!rows.length) return `<p class="status good">No complaint found for this lot in the workbook or current extraction.</p>`;
  const symptomMap=new Map();
  for (const row of rows) {
    const complaint=row["Complaint / Notification"]||"Unnumbered complaint";
    for (const label of new Set(symptomLabels(row))) {
      if (!symptomMap.has(label)) symptomMap.set(label,new Set());
      symptomMap.get(label).add(complaint);
    }
  }
  const symptoms=[...symptomMap.entries()]
    .map(([label,complaints])=>({label,count:complaints.size,complaints:[...complaints]}))
    .sort((a,b)=>b.count-a.count||a.label.localeCompare(b.label));
  const repeated=symptoms.filter(x=>x.count>1).length;
  const cols=["Source Group","Lot","Product Family","Customer Company","Rolls Implicated","Samples Received","Complaint / Notification","Problem","Tests / Assays Applied","Result / Status","Final Roll(s)"];
  return `<div class="history-summary">
    <div class="metric"><strong>${rows.length}</strong><span>unique complaint${rows.length===1?"":"s"} from this lot</span></div>
    <div class="metric"><strong>${symptoms.length}</strong><span>symptom theme${symptoms.length===1?"":"s"} reported</span></div>
    <div class="metric"><strong>${repeated}</strong><span>symptom theme${repeated===1?"":"s"} reported more than once</span></div>
  </div>
  <h3>Symptom frequency</h3>
  <table class="symptom-table"><thead><tr><th>Symptom</th><th>Complaint count</th><th>Complaint number(s)</th></tr></thead><tbody>
    ${symptoms.map(x=>`<tr><td>${esc(x.label)}</td><td>${x.count}</td><td>${esc(x.complaints.join("; "))}</td></tr>`).join("")}
  </tbody></table>
  <p class="hint">Each complaint is counted once per symptom theme, even if it appears on multiple workbook sheets.</p>
  <h3>Complaints from this lot</h3>
  <table><thead><tr>${cols.map(c=>`<th>${c}</th>`).join("")}</tr></thead><tbody>
  ${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c]??"")}</td>`).join("")}</tr>`).join("")}
  </tbody></table>`;
}

async function handleWorkbookSelection(file) {
  if (!file) return;
  workbookBuffer=await file.arrayBuffer();
  workbookFileName=file.name;
  summaryDataset=[];
  try {
    const wb=await loadWorkbook(workbookBuffer.slice(0));
    workbookMode="standard";
    setReviewProfiles(reviewProfilesFromWorkbook(wb));
    for (const id of ["excelStatus","summaryExcelStatus"]) {
      $(id).className="status good";
      $(id).textContent=`Loaded ${file.name} (${wb.worksheets.length} sheets).`;
    }
  } catch(err) {
    if (await isValidXlsxContainer(workbookBuffer)) {
      workbookMode="reference-readonly";
      const sheets=await readReferenceWorkbook(workbookBuffer.slice(0));
      setReviewProfiles(reviewProfilesFromReferenceSheets(sheets));
      for (const id of ["excelStatus","summaryExcelStatus"]) {
        $(id).className="status good";
        $(id).textContent=`Loaded ${file.name} as a protected reference. The app will read it locally and leave the original unchanged.`;
      }
    } else {
      workbookBuffer=null;
      workbookFileName="";
      setReviewProfiles([]);
      for (const id of ["excelStatus","summaryExcelStatus"]) {
        $(id).className="status bad";
        $(id).textContent=`Could not read workbook: ${err.message}`;
      }
    }
  }
}

$("excelFile").addEventListener("change", async e=>{
  syncRecordsFromDom();
  await handleWorkbookSelection(e.target.files?.[0]);
});

$("summarySourceFiles").addEventListener("change", async e=>{
  const files=[...(e.target.files||[])];
  if (!files.length) return;
  $("summaryExcelStatus").className="status";
  $("summaryExcelStatus").textContent="Reading selected source files locally...";
  let added=0, failed=0;
  for (const file of files) {
    const kind=/\.xlsx$/i.test(file.name)?"workbook":/\.(pdf|msg|zip)$/i.test(file.name)?"report":"";
    if (!kind) continue;
    const key=`${file.name}|${file.size}|${file.lastModified}`;
    if (summarySources.some(source=>source.key===key)) continue;
    try {
      summarySources.push({key,name:file.name,kind,buffer:await file.arrayBuffer(),records:null,errors:[]});
      added++;
    } catch (_) { failed++; }
  }
  summaryDataset=[];
  e.target.value="";
  $("summaryExcelStatus").className=failed?"status bad":"status good";
  $("summaryExcelStatus").textContent=`Selected ${summarySources.length} summary source file(s) (${added} newly added)${failed?`; ${failed} file(s) were unavailable`:""}. Choose more files or create the summary.`;
});

$("extractBtn").onclick=async()=>{
  const files=$("complaintFiles").files;
  if (!files?.length) return;
  syncRecordsFromDom();
  $("extractStatus").className="status";
  $("extractStatus").textContent="Extracting...";
  try {
    const expanded=await expandFiles(files);
    const newRecords=[];
    for (const f of expanded) {
      try { newRecords.push(...await extractMany(f.name,f.buffer)); }
      catch(err) {
        newRecords.push({
          sourceFile:f.name, sourceType:"", sourceGroup:"Final Reports",
          complaintNo:"", reportDate:"", customerCompany:"", rollsImplicated:"", samplesReceived:"",
          sampleDetails:"", materialNo:"", productDescription:"", productFamily:"",
          lot:"", membraneType:"", zones:"", mrfrCombined:"", standardizedSymptoms:"", problemTypes:"", lfaRelevance:"",
          complaintRegisteredDate:"", daysToReport:"", formalProblem:"", problem:"", customerReportedFailure:"", assaysApplied:"", resultStatus:"", criticality:"", masterRolls:"",
          finalRolls:"", mrfrAreas:"", failureReproduced:"", rootCauseRelated:"",
          coordinator:"", similarEvents:"", containmentNecessary:"", correctiveActionNecessary:"",
          rootCauseConclusion:"", problemValidation:"", finalAssessment:"", finalScope:"", testEvidence:[],
          warnings:`Extraction error: ${err.message}`, rawText:""
        });
      }
    }
    let added=0, updated=0;
    for (const record of newRecords) {
      const existingIndex=matchingRecordIndex(record);
      if (existingIndex>=0) {
        records[existingIndex]=record;
        updated++;
      } else {
        records.push(record);
        added++;
      }
    }
    renderRecords();
    $("complaintFiles").value="";
    $("extractStatus").className="status good";
    const parts=[];
    if (added) parts.push(`added ${added}`);
    if (updated) parts.push(`updated ${updated}`);
    $("extractStatus").textContent=`Extraction complete: ${parts.join(", ")}. Existing complaints were kept.`;
  } catch(err) {
    $("extractStatus").className="status bad";
    $("extractStatus").textContent=err.message;
  }
};

$("saveDraftBtn").onclick=async()=>{
  $("draftStatus").className="status";
  $("draftStatus").textContent="Saving locally in this browser...";
  try {
    await saveTemporaryDraft();
    $("draftStatus").className="status good";
    $("draftStatus").textContent=`Saved ${records.length} complaint(s) temporarily in this browser. Nothing was uploaded.`;
  } catch(err) {
    $("draftStatus").className="status bad";
    $("draftStatus").textContent=err.message;
  }
};

$("clearBtn").onclick=async()=>{
  records=[];
  collapsedRecords.clear();
  renderRecords();
  try { await deleteTemporaryDraft(); } catch (_) {}
  $("draftStatus").className="status";
  $("draftStatus").textContent="Current and temporarily saved results were cleared.";
};

$("selectAllSheetsBtn").onclick=()=>{
  document.querySelectorAll("[data-export-sheet]").forEach(input=>{input.checked=true;});
};

$("clearAllSheetsBtn").onclick=()=>{
  document.querySelectorAll("[data-export-sheet]").forEach(input=>{input.checked=false;});
};

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
    a.href=url; a.download=workbookMode==="standard"
      ?"Report_Extraction_Updated.xlsx":"Report_Extraction_New.xlsx";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),2000);
    $("buildStatus").className="status good";
    const prefix=workbookMode==="standard"
      ?"Updated workbook created and downloaded."
      :workbookMode==="reference-readonly"
        ?"New extracted workbook created. The reference workbook was not changed."
        :"New Excel workbook created and downloaded from the extracted reports.";
    const preserved=lastPreservedSheetNames.length
      ?` Legacy-format sheet${lastPreservedSheetNames.length===1?"":"s"} preserved unchanged: ${lastPreservedSheetNames.join(", ")}. New report rows are available in Complaint Summary and the category sheets.`
      :"";
    $("buildStatus").textContent=`${prefix} Workbook sheets (${lastBuiltSheetNames.length}): ${lastBuiltSheetNames.join(", ")}.${preserved}`;
  } catch(err) {
    $("buildStatus").className="status bad";
    $("buildStatus").textContent=err.message;
  }
};

const ROLL_PLAN_REFERENCE = {
  "ZM17_20mm":{machine:"ZM17",slitWidth:"20 mm",masterWidth:"1580 mm",sectionWidth:"780 mm",slots:76,zoneStarts:[0,13,25,38,51,63]},
  "ZM17_25mm":{machine:"ZM17",slitWidth:"25 mm",masterWidth:"1580 mm",sectionWidth:"780 mm",slots:60,zoneStarts:[0,10,20,30,40,50]},
  "ZM9_18mm":{machine:"ZM9",slitWidth:"18 mm",masterWidth:"1200 mm",sectionWidth:"360 mm",slots:57,zoneStarts:[0,10,19,29,38,48]},
  "ZM9_20mm":{machine:"ZM9",slitWidth:"20 mm",masterWidth:"1200 mm",sectionWidth:"360 mm",slots:51,zoneStarts:[0,9,17,26,34,43]},
  "ZM9_25mm":{machine:"ZM9",slitWidth:"25 mm",masterWidth:"1200 mm",sectionWidth:"360 mm",slots:42,zoneStarts:[0,7,14,21,28,35]},
  "ZM10_20mm":{machine:"ZM10",slitWidth:"20 mm",masterWidth:"1200 mm",sectionWidth:"390 mm",slots:57,zoneStarts:[0,10,19,29,38,48]},
  "ZM10_25mm":{machine:"ZM10",slitWidth:"25 mm",masterWidth:"1200 mm",sectionWidth:"360 mm",slots:45,zoneStarts:[0,8,15,23,30,38]}
};

function zoneForSlot(spec,slotIndex) {
  let zone=1;
  spec.zoneStarts.forEach((start,index)=>{if (slotIndex>=start) zone=index+1;});
  return zone;
}

function renderRollPlanSheet() {
  const name=$("rollPlanSheet").value;
  const spec=ROLL_PLAN_REFERENCE[name]||Object.values(ROLL_PLAN_REFERENCE)[0];
  $("rollPlanHeading").textContent=name;
  const zoneHeader=spec.zoneStarts.map((start,index)=>{
    const end=spec.zoneStarts[index+1]??spec.slots;
    return `<th class="plan-zone zone-${index+1}" colspan="${end-start}">Zone ${index+1}</th>`;
  }).join("");
  const slotHeader=Array.from({length:spec.slots},(_,index)=>`<th class="plan-slot zone-${zoneForSlot(spec,index)}">${index+1}</th>`).join("");
  const rows=Array.from({length:50},(_,masterIndex)=>`<tr><th class="plan-master">M${masterIndex+1}</th>${Array.from({length:spec.slots},(_,slotIndex)=>`<td class="plan-position zone-${zoneForSlot(spec,slotIndex)}" title="M${masterIndex+1} · final-roll position ${slotIndex+1} · Zone ${zoneForSlot(spec,slotIndex)}"></td>`).join("")}</tr>`).join("");
  $("rollPlanTable").innerHTML=`<div class="roll-plan-meta">
      <span><strong>Machine</strong>${esc(spec.machine)}</span><span><strong>Slit width</strong>${esc(spec.slitWidth)}</span>
      <span><strong>Master width</strong>${esc(spec.masterWidth)}</span><span><strong>Reference section</strong>${esc(spec.sectionWidth)}</span>
      <span><strong>Final-roll positions</strong>${spec.slots}</span><span><strong>Zones</strong>6</span>
    </div><div class="roll-plan-grid-scroll"><table class="roll-plan-table"><thead><tr><th class="plan-master plan-sticky-corner">Master roll</th>${zoneHeader}</tr><tr><th class="plan-master">Final-roll position</th>${slotHeader}</tr></thead><tbody>${rows}</tbody></table></div>`;
  $("rollPlanStatus").className="status good";
  $("rollPlanStatus").textContent=`Displaying ${name}: 50 master-roll rows × ${spec.slots} final-roll positions. Read-only reference.`;
}

function initializeRollPlanReference() {
  $("rollPlanSheet").innerHTML=Object.keys(ROLL_PLAN_REFERENCE).map(name=>`<option value="${esc(name)}">${esc(name)}</option>`).join("");
  renderRollPlanSheet();
}

$("rollPlanSheet").addEventListener("change",renderRollPlanSheet);

function showAppView(viewId) {
  document.querySelectorAll(".app-view").forEach(view=>{view.hidden=view.id!==viewId;});
  document.querySelectorAll("[data-app-tab]").forEach(button=>{
    const active=button.dataset.appTab===viewId;
    button.classList.toggle("active",active);
    button.setAttribute("aria-selected",String(active));
  });
}

document.querySelectorAll("[data-app-tab]").forEach(button=>{
  button.addEventListener("click",()=>{
    showAppView(button.dataset.appTab);
    if (button.dataset.appTab==="searchView") {
      refreshDirectSearchChoices().then(()=>refreshSearchChoices()).then(()=>refreshDecisionSymptomChoices());
    }
  });
});

function renderSummaryFieldOptions() {
  $("summaryFieldOptions").innerHTML=SUMMARY_FIELDS.map(field=>`
    <label><input type="checkbox" data-summary-field value="${esc(field.key)}" ${field.essential?"checked":""} /> ${esc(field.label)}</label>`).join("");
  document.querySelectorAll("[data-summary-field]").forEach(input=>{
    input.addEventListener("change",()=>{
      if (summaryDataset.length) $("summaryResult").innerHTML=renderDatasetTable(summaryDataset,selectedSummaryFields());
    });
  });
}

$("summaryBasicBtn").onclick=()=>{
  const essentials=new Set(SUMMARY_FIELDS.filter(field=>field.essential).map(field=>field.key));
  document.querySelectorAll("[data-summary-field]").forEach(input=>{input.checked=essentials.has(input.value);});
  if (summaryDataset.length) $("summaryResult").innerHTML=renderDatasetTable(summaryDataset,selectedSummaryFields());
};

$("summaryLotBtn").onclick=()=>{
  selectedLotFamilies=new Set(LOT_FAMILY_ORDER);
  document.querySelectorAll("[data-summary-field]").forEach(input=>{input.checked=input.value==="lot";});
  if (summaryDataset.length) $("summaryResult").innerHTML=renderDatasetTable(summaryDataset,selectedSummaryFields());
  else {
    $("summaryResult").innerHTML=`<p class="status">Choose source files and click Create summary, then Show lot details.</p>`;
  }
};

$("summaryResult").addEventListener("change",event=>{
  const input=event.target.closest("[data-lot-family-filter]");
  if (!input) return;
  if (input.checked) selectedLotFamilies.add(input.value);
  else selectedLotFamilies.delete(input.value);
  $("summaryResult").innerHTML=renderDatasetTable(summaryDataset,selectedSummaryFields());
});

$("summaryAllBtn").onclick=()=>{
  document.querySelectorAll("[data-summary-field]").forEach(input=>{input.checked=true;});
  if (summaryDataset.length) $("summaryResult").innerHTML=renderDatasetTable(summaryDataset,selectedSummaryFields());
};

$("summaryBtn").onclick=refreshSummary;

$("clearSummaryFilesBtn").onclick=()=>{
  summarySources=[];
  summaryDataset=[];
  $("summarySourceFiles").value="";
  $("summaryExcelStatus").className="status";
  $("summaryExcelStatus").textContent="Selected summary source files were cleared. The export workbook and extracted reports were not changed.";
  $("summaryMetrics").innerHTML="";
  $("summaryOverview").innerHTML="";
  $("summaryResult").innerHTML="";
};

const SEARCH_RESULT_FIELDS = SUMMARY_FIELDS.filter(field=>[
  "complaintNo","lot","lotComplaintCount","customer","productFamily","materialNo","problem",
  "result","tests","rollsImplicated","samplesReceived","registeredDate","reportDate","sourceSheets"
].includes(field.key));

function matchesSearchFamily(row,family) {
  if (!family) return true;
  return splitUniqueValues(row.productFamily).some(value=>value.toLowerCase()===family.toLowerCase());
}

function matchesSearchCustomer(row,customer) {
  if (!customer) return true;
  const wanted=canonicalCustomerName(customer).toLowerCase();
  return splitUniqueValues(row.customer).map(canonicalCustomerName).some(value=>value.toLowerCase()===wanted);
}

function symptomTokensForRow(row) {
  const listed=splitUniqueValues(row.standardizedSymptoms).filter(value=>value.toLowerCase()!=="review required");
  if (listed.length) return listed;
  const inferred=splitUniqueValues(complaintClassification(row.problem||"","").standardizedSymptoms)
    .filter(value=>value.toLowerCase()!=="review required");
  return inferred;
}

function refreshDecisionSymptomChoices() {
  const symptomSelect=$("decisionSymptom");
  const previous=symptomSelect.value;
  const family=$("decisionFamily").value;
  const symptoms=new Map();
  for (const row of summaryDataset.filter(row=>matchesSearchFamily(row,family))) {
    symptomTokensForRow(row).forEach(symptom=>{
      const key=symptom.toLowerCase();
      if (!symptoms.has(key)) symptoms.set(key,symptom);
    });
  }
  const sorted=[...symptoms.values()].filter(Boolean).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
  symptomSelect.innerHTML=`<option value="">Choose a symptom</option>${sorted.map(value=>`<option value="${esc(value)}">${esc(value)}</option>`).join("")}`;
  if (sorted.includes(previous)) symptomSelect.value=previous;
  symptomSelect.disabled=!sorted.length;
  $("decisionSearchStatus").className=sorted.length?"status good":"status bad";
  $("decisionSearchStatus").textContent=sorted.length
    ?`${sorted.length} symptom${sorted.length===1?"":"s"} available${family?` for ${family}`:""}.`
    :`No symptoms are available${family?` for ${family}`:""}. Load source files in Workbook Summary first.`;
  $("decisionSearchResult").innerHTML="";
}

function renderDecisionSearch(rows) {
  const testCounts=new Map();
  for (const row of rows) {
    for (const test of new Set(splitUniqueValues(row.tests))) {
      testCounts.set(test,(testCounts.get(test)||0)+1);
    }
  }
  const commonTests=[...testCounts.entries()]
    .sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0]))
    .slice(0,6);
  const reference=commonTests.length
    ?commonTests.map(([test,count])=>`${esc(test)} (${count} case${count===1?"":"s"})`).join(" · ")
    :"No tests were recorded for these matching complaints.";
  return `<div class="decision-reference"><strong>Historical test reference</strong>${reference}
    <small>Use this as a history lookup; the appropriate test plan still depends on the current sample and investigation.</small></div>
    <table class="summary-table"><thead><tr>
      <th>Complaint Number</th><th>Lot(s)</th><th>Product Family</th><th>Symptom(s)</th><th>Reason / Problem</th><th>Tests Performed</th><th>Final Result / Status</th>
    </tr></thead><tbody>${rows.map(row=>`<tr>
      <td>${esc(row.complaintNo||"Complaint number unavailable")}</td>
      <td>${esc(row.lot||"")}</td>
      <td>${esc(row.productFamily||"")}</td>
      <td>${esc(symptomTokensForRow(row).join("; "))}</td>
      <td>${esc(row.problem||"")}</td>
      <td>${esc(row.tests||"Not recorded in workbook")}</td>
      <td>${esc(row.result||"Not recorded in workbook")}</td>
    </tr>`).join("")}</tbody></table>`;
}

async function refreshDirectSearchChoices() {
  const valueSelect=$("directSearchValue");
  const previous=valueSelect.value;
  const searchType=$("directSearchType").value;
  const family=$("directSearchFamily").value;
  const noun=searchType==="lot"?"lot number":"complaint number";
  valueSelect.disabled=true;
  valueSelect.innerHTML=`<option value="">Loading ${esc(noun)}s...</option>`;
  $("directSearchStatus").className="status";
  $("directSearchStatus").textContent="Reading the selected workbooks and current extracted reports...";
  try {
    summaryDataset=await collectComplaintDataset();
    const choices=new Map();
    const addChoice=(value,customer)=>{
      const cleanValue=String(value||"").trim();
      if (!cleanValue) return;
      if (!choices.has(cleanValue)) choices.set(cleanValue,new Set());
      for (const name of splitUniqueValues(customer).map(canonicalCustomerName).filter(Boolean)) choices.get(cleanValue).add(name);
    };
    for (const row of summaryDataset.filter(row=>matchesSearchFamily(row,family))) {
      if (searchType==="lot") lotTokens(row.lot).forEach(value=>addChoice(value,row.customer));
      else if (row.complaintNo) addChoice(row.complaintNo,row.customer);
    }
    const sorted=[...choices.keys()].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    valueSelect.innerHTML=`<option value="">Choose a ${esc(noun)}</option>${sorted.map(value=>{
      const customers=[...choices.get(value)].sort((a,b)=>a.localeCompare(b));
      const customerLabel=customers.length?customers.join("; "):"End customer not recorded";
      return `<option value="${esc(value)}">${esc(value)} — ${esc(customerLabel)}</option>`;
    }).join("")}`;
    if (sorted.includes(previous)) valueSelect.value=previous;
    valueSelect.disabled=!sorted.length;
    $("directSearchStatus").className=sorted.length?"status good":"status bad";
    $("directSearchStatus").textContent=sorted.length
      ?`${sorted.length} ${noun}${sorted.length===1?"":"s"} available${family?` for ${family}`:""}.`
      :`No ${noun}s are available${family?` for ${family}`:""}. Load source files in Workbook Summary first.`;
    $("directSearchResult").innerHTML="";
  } catch(err) {
    valueSelect.innerHTML=`<option value="">Choose a ${esc(noun)}</option>`;
    $("directSearchStatus").className="status bad";
    $("directSearchStatus").textContent=`Could not load search choices: ${err.message}`;
  }
}

async function runDirectSearch() {
  const query=$("directSearchValue").value.trim();
  if (!query) {
    $("directSearchStatus").className="status bad";
    $("directSearchStatus").textContent=`Choose a ${$("directSearchType").value==="lot"?"lot":"complaint"} number.`;
    $("directSearchResult").innerHTML="";
    return;
  }
  $("directSearchStatus").className="status";
  $("directSearchStatus").textContent="Searching workbook and current extracted reports...";
  try {
    summaryDataset=await collectComplaintDataset();
    const searchType=$("directSearchType").value;
    const family=$("directSearchFamily").value;
    const matches=summaryDataset.filter(row=>{
      if (!matchesSearchFamily(row,family)) return false;
      if (searchType==="complaintNo") return String(row.complaintNo||"").trim()===query;
      return lotTokens(row.lot).some(lot=>lot===query);
    });
    $("directSearchStatus").className="status good";
    $("directSearchStatus").textContent=`Found ${matches.length} matching complaint${matches.length===1?"":"s"}.`;
    $("directSearchResult").innerHTML=renderDatasetTable(matches,SEARCH_RESULT_FIELDS);
  } catch(err) {
    $("directSearchStatus").className="status bad";
    $("directSearchStatus").textContent=`Search failed: ${err.message}`;
    $("directSearchResult").innerHTML="";
  }
}

async function refreshSearchChoices() {
  const customerSelect=$("searchCustomer");
  const familySelect=$("searchFamily");
  const valueSelect=$("searchValue");
  const previousCustomer=customerSelect.value;
  const previousFamily=familySelect.value;
  const previous=valueSelect.value;
  const noun="complaint number";
  customerSelect.disabled=true;
  familySelect.disabled=true;
  valueSelect.disabled=true;
  valueSelect.innerHTML=`<option value="">Loading complaint numbers...</option>`;
  $("searchStatus").className="status";
  $("searchStatus").textContent="Reading the selected workbooks and current extracted reports...";
  try {
    summaryDataset=await collectComplaintDataset();
    const customers=[...new Set(summaryDataset.flatMap(row=>splitUniqueValues(row.customer)).map(canonicalCustomerName).filter(Boolean))]
      .sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    customerSelect.innerHTML=`<option value="">Choose end customer</option>${customers.map(value=>`<option value="${esc(value)}">${esc(value)}</option>`).join("")}`;
    if (customers.includes(previousCustomer)) customerSelect.value=previousCustomer;
    customerSelect.disabled=!customers.length;
    const customer=customerSelect.value;

    const familyRows=summaryDataset.filter(row=>matchesSearchCustomer(row,customer));
    const families=[...new Set(familyRows.flatMap(row=>splitUniqueValues(row.productFamily)).filter(Boolean))]
      .sort((a,b)=>{
        const aIndex=LOT_FAMILY_ORDER.indexOf(a), bIndex=LOT_FAMILY_ORDER.indexOf(b);
        return (aIndex<0?999:aIndex)-(bIndex<0?999:bIndex)||a.localeCompare(b);
      });
    familySelect.innerHTML=`<option value="">Choose membrane type</option>${families.map(value=>`<option value="${esc(value)}">${esc(value)}</option>`).join("")}`;
    if (customer && families.includes(previousFamily)) familySelect.value=previousFamily;
    familySelect.disabled=!customer||!families.length;
    const family=familySelect.value;

    const choices=new Map();
    const addChoice=(value,customer)=>{
      const cleanValue=String(value||"").trim();
      if (!cleanValue) return;
      if (!choices.has(cleanValue)) choices.set(cleanValue,new Set());
      for (const name of splitUniqueValues(customer).map(canonicalCustomerName).filter(Boolean)) choices.get(cleanValue).add(name);
    };
    for (const row of summaryDataset.filter(row=>matchesSearchCustomer(row,customer)&&matchesSearchFamily(row,family))) {
      if (row.complaintNo) addChoice(row.complaintNo,row.customer);
    }
    const sorted=[...choices.keys()].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    valueSelect.innerHTML=`<option value="">Choose a ${esc(noun)}</option>${sorted.map(value=>{
      const customers=[...choices.get(value)].sort((a,b)=>a.localeCompare(b));
      const customerLabel=customers.length?customers.join("; "):"End customer not recorded";
      return `<option value="${esc(value)}">${esc(value)} — ${esc(customerLabel)}</option>`;
    }).join("")}`;
    if (customer && family && sorted.includes(previous)) valueSelect.value=previous;
    valueSelect.disabled=!customer||!family||!sorted.length;
    $("searchStatus").className=customers.length?"status good":"status bad";
    $("searchStatus").textContent=!customers.length
      ?"No end customers are available. Load source files in Workbook Summary first."
      :!customer
        ?`${customers.length} end customer${customers.length===1?"":"s"} available. Choose one to continue.`
        :!family
          ?`${families.length} membrane type${families.length===1?"":"s"} available for ${customer}. Choose one to continue.`
          :sorted.length
            ?`${sorted.length} ${noun}${sorted.length===1?"":"s"} available for ${customer} · ${family}.`
            :`No complaint numbers are available for ${customer} · ${family}.`;
    $("searchResult").innerHTML="";
  } catch(err) {
    valueSelect.innerHTML=`<option value="">Choose a ${esc(noun)}</option>`;
    $("searchStatus").className="status bad";
    $("searchStatus").textContent=`Could not load search choices: ${err.message}`;
  }
}

async function runQuickSearch() {
  const customer=$("searchCustomer").value.trim();
  const family=$("searchFamily").value.trim();
  const query=$("searchValue").value.trim();
  if (!customer || !family || !query) {
    $("searchStatus").className="status bad";
    $("searchStatus").textContent=!customer?"Choose an end customer.":!family?"Choose a membrane type.":"Choose a complaint number.";
    $("searchResult").innerHTML="";
    return;
  }
  $("searchStatus").className="status";
  $("searchStatus").textContent="Searching workbook and current extracted reports...";
  try {
    summaryDataset=await collectComplaintDataset();
    const matches=summaryDataset.filter(row=>{
      if (!matchesSearchCustomer(row,customer)) return false;
      if (!matchesSearchFamily(row,family)) return false;
      return String(row.complaintNo||"").trim()===query;
    });
    $("searchStatus").className="status good";
    $("searchStatus").textContent=`Found ${matches.length} matching complaint${matches.length===1?"":"s"}.`;
    $("searchResult").innerHTML=renderDatasetTable(matches,SEARCH_RESULT_FIELDS);
  } catch(err) {
    $("searchStatus").className="status bad";
    $("searchStatus").textContent=`Search failed: ${err.message}`;
    $("searchResult").innerHTML="";
  }
}

function runDecisionSearch() {
  const symptom=$("decisionSymptom").value.trim();
  const family=$("decisionFamily").value;
  if (!symptom) {
    $("decisionSearchStatus").className="status bad";
    $("decisionSearchStatus").textContent="Choose a symptom.";
    $("decisionSearchResult").innerHTML="";
    return;
  }
  const matches=summaryDataset.filter(row=>
    matchesSearchFamily(row,family) &&
    symptomTokensForRow(row).some(value=>value.toLowerCase()===symptom.toLowerCase())
  );
  $("decisionSearchStatus").className=matches.length?"status good":"status bad";
  $("decisionSearchStatus").textContent=matches.length
    ?`Found ${matches.length} similar complaint case${matches.length===1?"":"s"}${family?` for ${family}`:""}.`
    :`No complaint cases match ${symptom}${family?` for ${family}`:""}.`;
  $("decisionSearchResult").innerHTML=matches.length?renderDecisionSearch(matches):"";
}

$("searchBtn").onclick=runQuickSearch;
$("directSearchBtn").onclick=runDirectSearch;
$("directSearchType").addEventListener("change",refreshDirectSearchChoices);
$("directSearchFamily").addEventListener("change",refreshDirectSearchChoices);
$("searchCustomer").addEventListener("change",()=>{
  $("searchFamily").value="";
  $("searchValue").value="";
  refreshSearchChoices();
});
$("searchFamily").addEventListener("change",()=>{
  $("searchValue").value="";
  refreshSearchChoices();
});
$("decisionFamily").addEventListener("change",refreshDecisionSymptomChoices);
$("decisionSearchBtn").onclick=runDecisionSearch;

document.querySelectorAll("[data-review-tab]").forEach(button=>{
  button.addEventListener("click",()=>{
    syncRecordsFromDom();
    activeReviewTab=button.dataset.reviewTab;
    document.querySelectorAll("[data-review-tab]").forEach(item=>{
      const active=item.dataset.reviewTab===activeReviewTab;
      item.classList.toggle("active",active);
      item.setAttribute("aria-selected",String(active));
    });
    renderRecords();
  });
});

renderSummaryFieldOptions();
renderReviewControls();
renderRecords();
initializeRollPlanReference();
restoreTemporaryDraft();

window.__reportExtractionDebug = {
  getRecords: () => records.map(r=>structuredClone(r)),
  getLotRows: () => currentLotRows().map(row=>({
    lot:row.lot,
    complaintCount:row.complaints.size,
    complaints:[...row.complaints],
    families:[...row.families],
    symptoms:[...row.symptoms],
    statuses:[...row.statuses]
  })),
  getSelectedSheets: () => [...selectedExportSheets()],
  getLastBuiltSheetNames: () => [...lastBuiltSheetNames],
  getLastPreservedSheetNames: () => [...lastPreservedSheetNames],
  getSummaryDataset: () => summaryDataset.map(row=>structuredClone(row)),
  getSummarySources: () => summarySources.map(source=>({name:source.name,kind:source.kind,recordCount:source.records?.length||0,errorCount:source.errors?.length||0})),
  getReviewProfile: () => ({
    name:currentReviewProfile().name,
    columns:currentReviewProfile().columns.map(column=>({...column,selected:(reviewSelections.get(currentReviewProfile().name)||new Set()).has(column.id)}))
  }),
  collectComplaintDataset,
  parseRecord
};
