const DAYS = ["monday","tuesday","wednesday","thursday","friday","saturday","sunday"];
const DAY_LABELS = ["周一","周二","周三","周四","周五","周六","周日"];
const COLORS = ["DBEAFE","DCF5E7","FEF3C7","F3E8FF","FFE4E6","CCFBF1","FEE2E2","E0E7FF","FDE68A","D1FAE5"];
const BORDER = {
  top:{style:"thin",color:{argb:"FFB0B0B0"}},
  left:{style:"thin",color:{argb:"FFB0B0B0"}},
  bottom:{style:"thin",color:{argb:"FFB0B0B0"}},
  right:{style:"thin",color:{argb:"FFB0B0B0"}}
};

const ALIASES = {
  className:["className","courseName","curriculumName","lessonName","subjectName","teachingClassName","name"],
  teacher:["teacherName","teacherNames","teachers","teacher","instructorName","instructor","lecturerName"],
  weeks:["weeks","week","weekList","weekNums","weekNumbers","teachingWeeks","weekRange","courseWeeks","weekNo"],
  room:["classroomName","classroom","roomName","placeName","location","classroomAddress","room"],
  semester:["semesterId","semesterName","termId","termName","semester"]
};

const jsonInput = document.getElementById("jsonInput");
const generateBtn = document.getElementById("generateBtn");
const pasteBtn = document.getElementById("pasteBtn");
const statusEl = document.getElementById("status");
const diagnosticsEl = document.getElementById("diagnostics");
const bookmarkletLink = document.getElementById("bookmarkletLink");
const copyBookmarkletBtn = document.getElementById("copyBookmarkletBtn");

generateBtn.addEventListener("click", () => processAndGenerate(jsonInput.value, "手动粘贴"));

pasteBtn.addEventListener("click", async () => {
  try {
    const text = await navigator.clipboard.readText();
    if (!text.trim()) throw new Error("剪贴板为空");
    jsonInput.value = text;
    await processAndGenerate(text, "剪贴板");
  } catch (e) {
    setStatus(`读取剪贴板失败：${e.message}。可直接粘贴 JSON。`, "error");
  }
});

window.addEventListener("message", async event => {
  if (!isAllowedSchoolOrigin(event.origin)) return;
  if (!event.data || event.data.type !== "CCMU_SCHEDULE_DATA") return;

  try {
    event.source?.postMessage({type:"CCMU_SCHEDULE_ACK"}, event.origin);
  } catch (_) {}

  const payload = event.data.payload;
  jsonInput.value = typeof payload === "string" ? payload : JSON.stringify(payload);
  await processAndGenerate(payload, "教务系统一键导出");
});

setupBookmarklet();

function setStatus(message, type="") {
  statusEl.textContent = message;
  statusEl.className = type;
}

function setDiagnostics(lines=[]) {
  diagnosticsEl.hidden = !lines.length;
  diagnosticsEl.textContent = lines.join("\n");
}

function isAllowedSchoolOrigin(origin) {
  try {
    const host = new URL(origin).hostname.toLowerCase();
    return host === "graduate.ccmu.edu.cn" ||
      /^graduate(?:-\d+)?\.webvpn\.ccmu\.edu\.cn$/.test(host);
  } catch (_) {
    return false;
  }
}

async function processAndGenerate(input, source) {
  setStatus(`正在解析（来源：${source}）并生成课表...`, "info");
  setDiagnostics();

  try {
    if (typeof ExcelJS === "undefined") {
      throw new Error("ExcelJS 加载失败，请检查网络或改为仓库内本地引用");
    }

    const root = parseInput(input);
    const found = findBestSections(root);
    if (!found) {
      throw new Error("JSON 中未找到包含 monday~sunday 的课表节次数组，接口结构可能已变化");
    }

    const result = buildSchedule(found.sections);
    const lines = [
      `课表数组路径：${found.path || "(顶层)"}`,
      `识别到节次记录：${found.sections.length}`,
      `接口课程条目：${result.rawCount}`,
      `成功解析课程条目：${result.parsedCount}`,
      `最大周次：${result.maxWeek || 0}`
    ];
    if (result.exampleKeys.length) lines.push(`课程对象字段示例：${result.exampleKeys.join(", ")}`);
    if (result.warnings.length) lines.push(`提示：${result.warnings.slice(0,5).join("；")}`);
    setDiagnostics(lines);

    if (result.rawCount === 0) {
      throw new Error("接口返回了课表框架，但没有课程条目。请确认已选正确学期并点击“查询”；为避免生成损坏 XLSX，本次不会输出文件");
    }
    if (result.parsedCount === 0 || result.maxWeek < 1) {
      const keys = result.exampleKeys.length ? ` 当前课程对象字段：${result.exampleKeys.join(", ")}。` : "";
      throw new Error("发现课程对象，但无法识别课程名/教师/周次，可能是教务系统字段发生变化。" + keys);
    }

    const fileName = await generateExcel(result);
    setStatus(`课表生成成功：${fileName}`, "success");
  } catch (e) {
    console.error(e);
    const message = e instanceof SyntaxError ? "JSON 格式错误，请确认复制的是完整 Response" : e.message;
    setStatus(`生成失败：${message}`, "error");
  }
}

function parseInput(input) {
  if (input && typeof input === "object") return input;
  let text = String(input ?? "").trim();
  if (!text) throw new Error("JSON 内容不能为空");

  text = text.replace(/^\)\]\}',?\s*/, "");

  if (!text.startsWith("{") && !text.startsWith("[")) {
    const starts = [text.indexOf("{"), text.indexOf("[")].filter(x => x >= 0);
    const end = Math.max(text.lastIndexOf("}"), text.lastIndexOf("]"));
    if (starts.length && end > Math.min(...starts)) text = text.slice(Math.min(...starts), end + 1);
  }

  let value = JSON.parse(text);
  for (let i=0; i<2 && typeof value === "string"; i++) value = JSON.parse(value);
  return value;
}

function findBestSections(root) {
  let best = null;
  const seen = new Set();

  function score(arr) {
    if (!Array.isArray(arr) || !arr.length) return 0;
    let s = (arr.length >= 8 && arr.length <= 20) ? 20 : 0;
    for (const item of arr.slice(0,30)) {
      if (!item || typeof item !== "object" || Array.isArray(item)) continue;
      const dayCount = DAYS.filter(d => Array.isArray(item[d])).length;
      if (dayCount >= 3) s += 50 + dayCount * 5;
      if ("section" in item || "sectionName" in item || "key" in item) s += 5;
    }
    return s;
  }

  function visit(v, path="", depth=0) {
    if (depth > 6 || !v || typeof v !== "object" || seen.has(v)) return;
    seen.add(v);

    if (Array.isArray(v)) {
      const s = score(v);
      if (s > 0 && (!best || s > best.score)) best = {sections:v, path, score:s};
      v.slice(0,30).forEach((x,i) => visit(x, `${path}[${i}]`, depth+1));
      return;
    }

    Object.entries(v).forEach(([k,x]) => visit(x, path ? `${path}.${k}` : k, depth+1));
  }

  visit(root);
  return best;
}

function first(obj, keys) {
  for (const k of keys) {
    if (obj && obj[k] !== undefined && obj[k] !== null && obj[k] !== "") return obj[k];
  }
}

function periodIndex(section, fallback) {
  const label = first(section, ["section","sectionName","periodName","period","lessonNo"]);
  if (label !== undefined) {
    const m = String(label).match(/(\d{1,2})/);
    if (m) {
      const n = Number(m[1]);
      if (n >= 1 && n <= 12) return n - 1;
    }
  }

  const key = Number(section?.key);
  if (Number.isInteger(key)) {
    if (key >= 0 && key <= 11) return key;
    if (key >= 1 && key <= 12) return key - 1;
  }

  return fallback >= 0 && fallback <= 11 ? fallback : null;
}

function textOf(value, keys=["name","teacherName","className","value","label"]) {
  if (value === undefined || value === null) return "";
  if (Array.isArray(value)) {
    return [...new Set(value.map(v => textOf(v, keys)).filter(Boolean))].join("、");
  }
  if (typeof value === "object") {
    for (const k of keys) {
      const t = textOf(value[k], keys);
      if (t) return t;
    }
    return "";
  }
  return String(value).trim();
}

function parseWeeks(value) {
  if (value === undefined || value === null || value === "") return [];
  if (Array.isArray(value)) return [...new Set(value.flatMap(parseWeeks))].sort((a,b)=>a-b);

  if (typeof value === "object") {
    const nested = first(value, ["weeks","week","weekNo","weekNum","weekNumber","value","label"]);
    return nested !== undefined ? parseWeeks(nested) : [];
  }

  if (typeof value === "number") {
    return Number.isInteger(value) && value > 0 && value <= 60 ? [value] : [];
  }

  let s = String(value).trim();
  if (!s) return [];

  if ((s.startsWith("[") && s.endsWith("]")) || (s.startsWith('"') && s.endsWith('"'))) {
    try { return parseWeeks(JSON.parse(s)); } catch (_) {}
  }

  s = s.replace(/[，、；;]/g,",")
       .replace(/[～~—–至到]/g,"-")
       .replace(/第/g,"")
       .replace(/周/g,"");

  const out = new Set();
  for (const chunk of s.split(/[,\s]+/).filter(Boolean)) {
    const odd = /单/.test(chunk), even = /双/.test(chunk);
    const range = chunk.match(/(\d{1,2})\s*-\s*(\d{1,2})/);

    if (range) {
      let a = Number(range[1]), b = Number(range[2]);
      if (a > b) [a,b] = [b,a];
      for (let w=a; w<=b && w<=60; w++) {
        if (w < 1 || (odd && w%2===0) || (even && w%2!==0)) continue;
        out.add(w);
      }
    } else {
      (chunk.match(/\d{1,2}/g) || []).map(Number).forEach(w => {
        if (w >= 1 && w <= 60) out.add(w);
      });
    }
  }
  return [...out].sort((a,b)=>a-b);
}

function buildSchedule(sections) {
  const schedule = {};
  let maxWeek = 0, semester = "", rawCount = 0, parsedCount = 0;
  const warnings = [], keys = new Set();

  sections.forEach((section, fallback) => {
    if (!section || typeof section !== "object") return;
    const p = periodIndex(section, fallback);
    if (p === null) return;

    for (const day of DAYS) {
      const items = section[day];
      if (!Array.isArray(items)) continue;

      for (const course of items) {
        if (!course || typeof course !== "object") continue;
        rawCount++;
        Object.keys(course).slice(0,30).forEach(k => keys.add(k));

        const className = textOf(first(course, ALIASES.className));
        const teacherText = textOf(first(course, ALIASES.teacher), ["teacherName","name","label","value"]);
        const weeks = parseWeeks(first(course, ALIASES.weeks));
        const room = textOf(first(course, ALIASES.room)) || "线上教学";
        const semesterValue = first(course, ALIASES.semester);

        if (!className || !teacherText || !weeks.length) {
          if (warnings.length < 10) {
            const miss = [!className&&"课程名", !teacherText&&"教师", !weeks.length&&"周次"].filter(Boolean).join("/");
            warnings.push(`有课程条目缺少可识别的${miss}`);
          }
          continue;
        }

        if (!semester && semesterValue !== undefined) semester = textOf(semesterValue) || String(semesterValue);
        parsedCount++;
        maxWeek = Math.max(maxWeek, ...weeks);
        const teachers = [...new Set(teacherText.split(/[、,，;；/]+/).map(x=>x.trim()).filter(Boolean))];

        for (const week of weeks) {
          schedule[week] ??= {};
          schedule[week][day] ??= Array.from({length:12}, () => []);
          const cell = schedule[week][day][p];
          const existing = cell.find(x => x.className === className && x.room === room);

          if (existing) {
            teachers.forEach(t => { if (!existing.teachers.includes(t)) existing.teachers.push(t); });
          } else {
            cell.push({className, room, teachers:[...teachers]});
          }
        }
      }
    }
  });

  return {schedule, maxWeek, semester, rawCount, parsedCount,
          warnings:[...new Set(warnings)], exampleKeys:[...keys].slice(0,30)};
}

function formatCell(entries) {
  return (entries || []).map(x => `${x.className}\n${x.room}\n${x.teachers.join("、")}`).join("\n────────\n");
}

function signature(entries) {
  if (!entries?.length) return null;
  return entries.map(x => `${x.className}|${x.room}|${[...x.teachers].sort().join("、")}`).sort().join("||");
}

function courseColors(schedule, maxWeek) {
  const names = new Set();
  for (let w=1; w<=maxWeek; w++) for (const d of DAYS)
    for (const cell of schedule[w]?.[d] || []) for (const x of cell || []) names.add(x.className);
  const map = {}; let i=0;
  names.forEach(n => map[n] = COLORS[i++ % COLORS.length]);
  return map;
}

async function generateExcel(result) {
  const {schedule,maxWeek,semester} = result;
  if (!Number.isInteger(maxWeek) || maxWeek < 1) throw new Error("没有可生成的有效周次");

  const wb = new ExcelJS.Workbook();
  wb.creator = "CCMU Schedule";
  wb.created = new Date();
  wb.modified = new Date();

  const headers = ["时间/节次", ...DAY_LABELS];
  const times = [
    "8:00-8:45","8:45-9:30","9:45-10:30","10:30-11:15","11:25-12:10",
    "13:30-14:15","14:15-15:00","15:10-15:55","15:55-16:40",
    "18:00-18:45","18:45-19:30","19:30-20:15"
  ];
  const cmap = courseColors(schedule, maxWeek);

  for (let week=1; week<=maxWeek; week++) {
    const ws = wb.addWorksheet(`第${week}周`);
    ws.views = [{state:"frozen",xSplit:1,ySplit:1,topLeftCell:"B2"}];
    ws.addRow(headers);

    const sigs = [], entriesByRow = [];
    for (let p=0; p<12; p++) {
      const row = [`第${p+1}节\n${times[p]}`], rs=[null], re=[null];
      for (const day of DAYS) {
        const entries = schedule[week]?.[day]?.[p] || [];
        row.push(formatCell(entries)); rs.push(signature(entries)); re.push(entries);
      }
      ws.addRow(row); sigs.push(rs); entriesByRow.push(re);
    }

    for (let col=1; col<=7; col++) {
      let start=0;
      while (start<12) {
        const s=sigs[start][col];
        if (!s) { start++; continue; }
        let end=start;
        while (end+1<12 && sigs[end+1][col]===s) end++;
        if (end>start) ws.mergeCells(start+2,col+1,end+2,col+1);
        start=end+1;
      }
    }

    const header = ws.getRow(1); header.height=28;
    header.eachCell({includeEmpty:true}, cell => {
      cell.font={name:"等线",size:12,bold:true,color:{argb:"FF333333"}};
      cell.alignment={vertical:"middle",horizontal:"center",wrapText:true};
      cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFE8ECF0"}};
      cell.border=BORDER;
    });

    for (let r=0; r<12; r++) {
      const row=ws.getRow(r+2), tc=row.getCell(1);
      tc.font={name:"等线",size:10,bold:true,color:{argb:"FF555555"}};
      tc.alignment={vertical:"middle",horizontal:"center",wrapText:true};
      tc.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF5F5F5"}};
      tc.border=BORDER;

      for (let col=1; col<=7; col++) {
        const cell=row.getCell(col+1), entries=entriesByRow[r][col] || [];
        cell.alignment={vertical:"middle",horizontal:"center",wrapText:true};
        cell.border=BORDER; cell.font={name:"等线",size:11};
        if (entries.length) {
          cell.fill={type:"pattern",pattern:"solid",
            fgColor:{argb:"FF"+(cmap[entries[0].className] || "FFFFFF")}};
        }
      }
    }

    ws.columns.forEach((column,index) => {
      let max = index===0 ? 14 : 12;
      column.eachCell({includeEmpty:true}, cell => {
        for (const line of String(cell.value ?? "").split("\n")) {
          let n=0; for (const ch of line) n += ch.charCodeAt(0)>255 ? 2 : 1;
          max=Math.max(max,n);
        }
      });
      column.width=Math.min(50,Math.max(15,max*1.15));
    });

    for (let r=0; r<12; r++) {
      const row=ws.getRow(r+2); let lines=2;
      row.eachCell({includeEmpty:true}, cell => lines=Math.max(lines,String(cell.value ?? "").split("\n").length));
      row.height=Math.min(120,Math.max(38,lines*20));
    }

    ws.pageSetup={orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0,paperSize:9};
  }

  if (!wb.worksheets.length) throw new Error("生成结果没有任何工作表，已阻止输出");

  const buffer = await wb.xlsx.writeBuffer();
  const verify = new ExcelJS.Workbook();
  await verify.xlsx.load(buffer);
  if (!verify.worksheets.length || verify.worksheets.length !== wb.worksheets.length) {
    throw new Error("XLSX 自检失败，已阻止下载");
  }

  const safeSemester = String(semester || "").trim()
    .replace(/[\\/:*?"<>|\[\]]+/g,"_").replace(/\s+/g,"_").slice(0,80);
  const fileName = safeSemester ? `课表_${safeSemester}.xlsx` : "course_schedule.xlsx";

  const blob = new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  const url=URL.createObjectURL(blob), a=document.createElement("a");
  a.href=url; a.download=fileName; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),3000);
  return fileName;
}

function setupBookmarklet() {
  const generator = new URL(".", window.location.href);
  generator.search=""; generator.hash="";
  const code = buildBookmarklet(generator.href);
  bookmarkletLink.href = `javascript:${code}`;

  copyBookmarkletBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(`javascript:${code}`);
      setStatus("书签代码已复制。新建书签并粘贴到网址/URL 栏即可。", "success");
    } catch (_) {
      setStatus("浏览器不允许自动复制；直接把蓝色链接拖到书签栏即可。", "info");
    }
  });
}

function buildBookmarklet(generatorUrl) {
  const G = JSON.stringify(generatorUrl);
  const code = `(function(){
var G=${G},O=(new URL(G)).origin,W=window.open(G,"ccmuScheduleExporter"),D=false,T=null;
if(!W){alert("请允许本页打开弹窗后重试。");return;}
function M(u){return String(u||"").indexOf("queryStudentSchedule")>=0}
function P(x){
 if(D)return;
 try{
  var o=typeof x==="string"?JSON.parse(x):x;
  if(!o||typeof o!=="object")return;
  D=true;var m={type:"CCMU_SCHEDULE_DATA",payload:o};
  function s(){try{W.postMessage(m,O)}catch(e){}}
  s();T=setInterval(s,500);setTimeout(function(){clearInterval(T)},10000);
 }catch(e){}
}
window.addEventListener("message",function(e){
 if(e.origin===O&&e.data&&e.data.type==="CCMU_SCHEDULE_ACK"&&T)clearInterval(T);
});
var F=window.fetch;
if(F)window.fetch=async function(){
 var r=await F.apply(this,arguments);
 try{var a=arguments[0],u=typeof a==="string"?a:(a&&a.url);if(M(u))r.clone().text().then(P)}catch(e){}
 return r;
};
var X=XMLHttpRequest.prototype,OO=X.open,SS=X.send;
X.open=function(m,u){this.__ccmuScheduleUrl=u;return OO.apply(this,arguments)};
X.send=function(){
 if(M(this.__ccmuScheduleUrl))this.addEventListener("load",function(){
  try{P(this.responseType==="json"?this.response:this.responseText)}catch(e){}
 },{once:true});
 return SS.apply(this,arguments);
};
function Q(){
 var es=[].slice.call(document.querySelectorAll("button,a,[role=button],.el-button"));
 var b=es.find(function(e){return /^\\s*查询\\s*$/.test((e.innerText||e.textContent||"").trim())});
 if(b){b.click();return true}return false;
}
if(!Q())alert("监听已开启，请在课表页面点击一次“查询”。");
setTimeout(function(){if(!D)alert("暂未捕获到课表接口。请确认当前在课表页，并再点击一次“查询”。")},8000);
})()`;
  return code.replace(/\n+/g,"").replace(/\s{2,}/g," ");
}
