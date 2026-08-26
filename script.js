const jsonInput = document.getElementById("jsonInput");
const generateBtn = document.getElementById("generateBtn");
const pasteBtn = document.getElementById("pasteBtn");
const statusEl = document.getElementById("status");
const diagnosticsEl = document.getElementById("diagnostics");
const bookmarkletLink = document.getElementById("bookmarkletLink");
const copyBookmarkletBtn = document.getElementById("copyBookmarkletBtn");

const params = new URLSearchParams(window.location.search);
const receiverMode = params.get("receiver") === "1";
const receiverToken = params.get("token") || "";
const receiverSemester = params.get("semester") || "";

if (receiverMode) document.body.classList.add("receiver-mode");

generateBtn.addEventListener("click", async () => {
    await processAndGenerate(jsonInput.value, "手动粘贴");
});

pasteBtn.addEventListener("click", async () => {
    try {
        const text = await navigator.clipboard.readText();
        if (!text.trim()) throw new Error("剪贴板为空");
        jsonInput.value = text;
        await processAndGenerate(text, "剪贴板");
    } catch (error) {
        setStatus(`读取剪贴板失败：${error.message}。可直接粘贴 JSON。`, "error");
    }
});

window.addEventListener("message", async (event) => {
    if (!isAllowedSchoolOrigin(event.origin)) return;
    if (!event.data || event.data.type !== "CCMU_SCHEDULE_DATA") return;

    // V5：接收窗口和书签任务通过 token 一一绑定。
    // 旧导出任务即使稍后才返回，也不能进入本次生成流程。
    if (receiverToken && event.data.token !== receiverToken) return;

    try {
        event.source?.postMessage({
            type: "CCMU_SCHEDULE_ACK",
            token: receiverToken
        }, event.origin);
    } catch (_) {}

    const payload = event.data.payload;
    jsonInput.value = typeof payload === "string" ? payload : JSON.stringify(payload);

    const source = event.data.source === "dom-fallback"
        ? "教务系统 DOM 安全兜底"
        : "本次 queryStudentSchedule 响应";

    const outcome = await processAndGenerate(payload, source);

    try {
        if (outcome?.ok) {
            event.source?.postMessage({
                type: "CCMU_SCHEDULE_COMPLETE",
                token: receiverToken,
                semester: receiverSemester,
                fileName: outcome.fileName
            }, event.origin);
        } else {
            event.source?.postMessage({
                type: "CCMU_SCHEDULE_ERROR",
                token: receiverToken,
                semester: receiverSemester,
                message: outcome?.error || "未知错误"
            }, event.origin);
        }
    } catch (_) {}

    if (receiverMode) {
        setTimeout(() => {
            try { window.close(); } catch (_) {}
        }, outcome?.ok ? 1400 : 6000);
    }
});

window.addEventListener("load", () => {
    if (!receiverMode || !window.opener) return;

    setStatus(
        receiverSemester
            ? `接收窗口已就绪，等待学期 ${receiverSemester} 的本次查询响应…`
            : "接收窗口已就绪，等待本次查询响应…",
        "info"
    );

    try {
        window.opener.postMessage({
            type: "CCMU_SCHEDULE_READY",
            token: receiverToken
        }, "*");
        window.blur();
        window.opener.focus();
    } catch (_) {}
});

setupBookmarklet();

function setStatus(message, type = "") {
    statusEl.textContent = message;
    statusEl.className = type;
}

function setDiagnostics(lines = []) {
    if (!diagnosticsEl) return;
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
    setDiagnostics([]);

    try {
        if (typeof ExcelJS === "undefined") {
            throw new Error("ExcelJS 加载失败，请检查网络或改为仓库内本地引用");
        }

        // V5 按用户最初 script (1).js 的方式解析 JSON：
        // 手动输入时只做一次 JSON.parse；收到书签传来的对象时直接交给 generateExcel。
        const jsonData = (input && typeof input === "object")
            ? input
            : JSON.parse(String(input ?? "").trim());

        const sections = Array.isArray(jsonData?.data) ? jsonData.data : null;
        const rawCourseCount = sections
            ? sections.reduce((sum, section) => {
                return sum + ["monday","tuesday","wednesday","thursday","friday","saturday","sunday"]
                    .reduce((n, day) => n + (Array.isArray(section?.[day]) ? section[day].length : 0), 0);
            }, 0)
            : 0;

        setDiagnostics([
            "JSON 解析模式：原始 script (1).js 规则",
            "要求结构：顶层 data 数组",
            "课程字段：weeks / className / teacherName / classroomName / semesterId",
            `节次记录：${sections ? sections.length : 0}`,
            `课程条目：${rawCourseCount}`
        ]);

        const fileName = await generateExcel(jsonData);
        setStatus(`课表生成成功：${fileName}`, "success");
        return { ok: true, fileName };
    } catch (error) {
        console.error("生成课表时发生错误:", error);

        let message = error.message;
        if (error instanceof SyntaxError) {
            message = "JSON格式错误，请检查是否复制完整...";
        }

        setStatus(`生成失败：${message}`, "error");
        return { ok: false, error: message };
    }
}


// 为不同课程分配柔和的背景色
const COURSE_COLORS = [
    'DBEAFE', // 浅蓝
    'DCF5E7', // 浅绿
    'FEF3C7', // 浅黄
    'F3E8FF', // 浅紫
    'FFE4E6', // 浅粉
    'CCFBF1', // 浅青
    'FEE2E2', // 浅红
    'E0E7FF', // 浅靛蓝
    'FDE68A', // 浅橙黄
    'D1FAE5', // 浅翠绿
];

function getCourseColorMap(schedule, maxWeek, daysOfWeek) {
    const courseNames = new Set();
    for (let w = 1; w <= maxWeek; w++) {
        if (!schedule[w]) continue;
        for (const day of daysOfWeek) {
            if (!schedule[w][day]) continue;
            for (const cell of schedule[w][day]) {
                if (cell) courseNames.add(cell.className);
            }
        }
    }
    const colorMap = {};
    let i = 0;
    for (const name of courseNames) {
        colorMap[name] = COURSE_COLORS[i % COURSE_COLORS.length];
        i++;
    }
    return colorMap;
}

// 通用边框样式
const THIN_BORDER = {
    top:    { style: 'thin', color: { argb: 'FFB0B0B0' } },
    left:   { style: 'thin', color: { argb: 'FFB0B0B0' } },
    bottom: { style: 'thin', color: { argb: 'FFB0B0B0' } },
    right:  { style: 'thin', color: { argb: 'FFB0B0B0' } },
};

async function generateExcel(jsonData) {
    // ========== 1. 数据解析与重组 ==========
    const schedule = {};
    let maxWeek = 0;
    let semesterId = '';
    const daysOfWeek = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"];
    
    if (!jsonData.data || !Array.isArray(jsonData.data)) {
        throw new Error("JSON结构不符合预期，缺少顶层 'data' 数组");
    }
    
    jsonData.data.forEach((sectionData, sectionIndex) => {
        const currentPeriodIndex = sectionIndex;

        for (const dayName of daysOfWeek) {
            if (sectionData[dayName] && Array.isArray(sectionData[dayName])) {
                for (const course of sectionData[dayName]) {
                    if (!course.weeks || !course.className || !course.teacherName) continue;

                    // 【改进5】提取学期信息用于文件命名
                    if (!semesterId && course.semesterId) {
                        semesterId = course.semesterId;
                    }

                    const weekNums = String(course.weeks).split(',').map(Number);
                    if (weekNums.length > 0) {
                        const currentMax = Math.max(...weekNums);
                        if (currentMax > maxWeek) maxWeek = currentMax;
                    }

                    for (const week of weekNums) {
                        if (!schedule[week]) schedule[week] = {};
                        if (!schedule[week][dayName]) schedule[week][dayName] = Array(12).fill(null);

                        const cellData = schedule[week][dayName][currentPeriodIndex];

                        if (!cellData) {
                            schedule[week][dayName][currentPeriodIndex] = {
                                className: course.className,
                                classroomName: course.classroomName || '线上教学',
                                teachers: [course.teacherName]
                            };
                        } else {
                            if (!cellData.teachers.includes(course.teacherName)) {
                                cellData.teachers.push(course.teacherName);
                            }
                        }
                    }
                }
            }
        }
    });

    // V5 安全保护：解析规则保持原版不变，但不再允许 maxWeek=0 时写出空工作簿。
    // 这不是字段兼容逻辑；只是防止无有效课程时生成 Excel 无法正常打开的空 XLSX。
    if (maxWeek < 1) {
        throw new Error("未按原始 JSON 结构解析到有效课程数据，已阻止生成空工作簿");
    }

    // 获取课程→颜色映射
    const courseColorMap = getCourseColorMap(schedule, maxWeek, daysOfWeek);
    
    // ========== 2. 创建Excel工作簿 ==========
    const wb = new ExcelJS.Workbook();
    const headers = ["时间/节次", "周一", "周二", "周三", "周四", "周五", "周六", "周日"];
    const times = [
        "8:00-8:45", "8:45-9:30", "9:45-10:30", "10:30-11:15", "11:25-12:10",
        "13:30-14:15", "14:15-15:00", "15:10-15:55", "15:55-16:40",
        "18:00-18:45", "18:45-19:30", "19:30-20:15"
    ];

    for (let weekNum = 1; weekNum <= maxWeek; weekNum++) {
        const ws = wb.addWorksheet(`第${weekNum}周`);

        // 【改进3】冻结首行首列，滚动时始终可见
        ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];

        // 写入表头
        ws.addRow(headers);

        // 写入12节课数据，同时记录每个单元格对应的课程名（用于后续合并判断）
        const cellCourseInfo = []; // cellCourseInfo[行][列] = className 或 null

        for (let i = 0; i < 12; i++) {
            const rowHeader = `第${i + 1}节\n${times[i]}`;
            const rowData = [rowHeader];
            const rowCourseNames = [null]; // 第一列是时间列，不参与合并

            for (const dayName of daysOfWeek) {
                const daySchedule = (schedule[weekNum] && schedule[weekNum][dayName]) ? schedule[weekNum][dayName] : [];
                const cellData = daySchedule[i];

                if (cellData) {
                    const formattedString = `${cellData.className}\n${cellData.classroomName}\n${cellData.teachers.join('、')}`;
                    rowData.push(formattedString);
                    rowCourseNames.push(cellData.className);
                } else {
                    rowData.push("");
                    rowCourseNames.push(null);
                }
            }
            ws.addRow(rowData);
            cellCourseInfo.push(rowCourseNames);
        }

        // ========== 【改进1】纵向合并连续相同课程的单元格 ==========
        for (let col = 1; col <= 7; col++) {
            let mergeStart = 0;
            while (mergeStart < 12) {
                const courseName = cellCourseInfo[mergeStart][col];
                if (!courseName) {
                    mergeStart++;
                    continue;
                }
                // 找到连续相同课程的结束位置
                let mergeEnd = mergeStart;
                while (mergeEnd + 1 < 12 && cellCourseInfo[mergeEnd + 1][col] === courseName) {
                    mergeEnd++;
                }
                // 如果跨了多行，执行合并（Excel行号 = 数据行索引 + 2，因为第1行是表头）
                if (mergeEnd > mergeStart) {
                    const excelCol = col + 1; // ExcelJS列号从1开始，第1列是时间列
                    const startRow = mergeStart + 2;
                    const endRow = mergeEnd + 2;
                    ws.mergeCells(startRow, excelCol, endRow, excelCol);
                }
                mergeStart = mergeEnd + 1;
            }
        }

        // ========== 【改进2】设置样式：边框、背景色 ==========

        // --- 表头行样式 ---
        const headerRow = ws.getRow(1);
        headerRow.height = 28;
        headerRow.eachCell({ includeEmpty: true }, (cell) => {
            cell.font = { name: '等线', size: 12, bold: true, color: { argb: 'FF333333' } };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE8ECF0' },
            };
            cell.border = THIN_BORDER;
        });

        // --- 数据行样式 ---
        for (let rowIdx = 0; rowIdx < 12; rowIdx++) {
            const excelRowNum = rowIdx + 2;
            const row = ws.getRow(excelRowNum);

            // 时间列（第1列）：浅灰底、加粗、小号字
            const timeCell = row.getCell(1);
            timeCell.font = { name: '等线', size: 10, bold: true, color: { argb: 'FF555555' } };
            timeCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
            timeCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF5F5F5' },
            };
            timeCell.border = THIN_BORDER;

            // 课程列（第2~8列）：根据课程名上色
            for (let col = 1; col <= 7; col++) {
                const cell = row.getCell(col + 1);
                const courseName = cellCourseInfo[rowIdx][col];

                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = THIN_BORDER;

                if (courseName) {
                    cell.font = { name: '等线', size: 11 };
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF' + (courseColorMap[courseName] || 'FFFFFF') },
                    };
                } else {
                    cell.font = { name: '等线', size: 11, color: { argb: 'FFAAAAAA' } };
                }
            }
        }

        // ========== 动态调整列宽和行高 ==========

        // 列宽：按最长行内容自适应
        ws.columns.forEach(column => {
            let maxCharLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                const cellText = cell.value ? cell.value.toString() : '';
                const lines = cellText.split('\n');
                lines.forEach(line => {
                    let lineLength = 0;
                    for (let k = 0; k < line.length; k++) {
                        lineLength += line.charCodeAt(k) > 255 ? 2 : 1;
                    }
                    if (lineLength > maxCharLength) {
                        maxCharLength = lineLength;
                    }
                });
            });
            column.width = Math.max(15, maxCharLength * 1.2);
        });

        // 行高：按换行数自适应，合并单元格后保证最小行高
        for (let rowIdx = 0; rowIdx < 12; rowIdx++) {
            const excelRowNum = rowIdx + 2;
            const row = ws.getRow(excelRowNum);
            let maxLines = 1;
            row.eachCell({ includeEmpty: true }, cell => {
                const cellText = cell.value ? cell.value.toString() : '';
                const numLines = cellText.split('\n').length;
                if (numLines > maxLines) {
                    maxLines = numLines;
                }
            });
            row.height = Math.max(maxLines * 22, 30);
        }
    }

    // ========== 【改进5】生成带学期信息的文件名 ==========
    const fileName = semesterId ? `课表_${semesterId}.xlsx` : 'course_schedule.xlsx';

    // 生成文件前再次确认至少有一个工作表。
    if (!wb.worksheets.length) {
        throw new Error("生成结果没有任何工作表，已阻止输出");
    }

    // 生成文件并使用 ExcelJS 自身回读一次，拦截明显的 OOXML/ZIP 写出异常。
    const buffer = await wb.xlsx.writeBuffer();
    const verifyBook = new ExcelJS.Workbook();
    await verifyBook.xlsx.load(buffer);
    if (!verifyBook.worksheets.length || verifyBook.worksheets.length !== wb.worksheets.length) {
        throw new Error("XLSX 自检失败，已阻止下载");
    }

    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    const objectUrl = URL.createObjectURL(blob);
    link.href = objectUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(objectUrl), 3000);

    return fileName;
}


function setupBookmarklet() {
    if (!bookmarkletLink || !copyBookmarkletBtn) return;

    const generator = new URL(".", window.location.href);
    generator.search = "";
    generator.hash = "";

    const code = buildBookmarklet(generator.href);
    const fullCode = `javascript:${code}`;
    bookmarkletLink.href = fullCode;

    copyBookmarkletBtn.addEventListener("click", async () => {
        try {
            await navigator.clipboard.writeText(fullCode);
            setStatus("V5 书签代码已复制。请替换浏览器里旧的 CCMU 导出书签。", "success");
        } catch (_) {
            setStatus("浏览器不允许自动复制；请把蓝色“CCMU 导出已选学期”重新拖到书签栏。", "info");
        }
    });
}

function buildBookmarklet(generatorUrl) {
    const G = JSON.stringify(generatorUrl);

    const code = `(function(){
var OLD=window.__CCMU_SCHEDULE_EXPORT_TOKEN__;
if(OLD){NOTICE("已有导出任务正在运行，请等待完成或刷新课表页面后重试。","warn",5000);return}
var G=${G},O=(new URL(G)).origin,TOKEN="v5-"+Date.now().toString(36)+"-"+Math.random().toString(36).slice(2,10);
window.__CCMU_SCHEDULE_EXPORT_TOKEN__=TOKEN;

var SEM="",W=null,ACK_TIMER=null,TIMEOUT=null,POLL=null,DONE=false,START=0,ARMED=false,SEEN_REQUEST=false,LAST_CHANGE=0,BEFORE_SIG="",LAST_SIG="";
var OF=window.fetch,XP=XMLHttpRequest.prototype,OO=XP.open,OS=XP.send,FW=null,XO=null,XS=null;

function ACTIVE(){return window.__CCMU_SCHEDULE_EXPORT_TOKEN__===TOKEN}
function TXT(e){return String((e&&e.textContent)||"").replace(/[\\s\\u00a0]+/g,"").trim()}
function NOTICE(m,t,d){
 var id="__ccmu_schedule_notice__",n=document.getElementById(id);
 if(!n){
  n=document.createElement("div");n.id=id;
  n.style.cssText="position:fixed;z-index:2147483647;right:20px;top:20px;max-width:460px;padding:13px 16px;border-radius:9px;background:#23364d;color:#fff;font:14px/1.55 -apple-system,BlinkMacSystemFont,Segoe UI,Arial,sans-serif;box-shadow:0 8px 28px rgba(0,0,0,.25);white-space:pre-wrap;transition:opacity .2s";
  document.documentElement.appendChild(n)
 }
 n.textContent=m;
 n.style.background=t==="error"?"#b93636":t==="success"?"#218c55":t==="warn"?"#8a6418":"#23364d";
 n.style.opacity="1";
 clearTimeout(n.__timer);
 if(d)n.__timer=setTimeout(function(){n.style.opacity="0";setTimeout(function(){try{n.remove()}catch(e){}},250)},d)
}
function RESTORE(){
 try{if(FW&&window.fetch===FW)window.fetch=OF}catch(e){}
 try{if(XO&&XP.open===XO)XP.open=OO}catch(e){}
 try{if(XS&&XP.send===XS)XP.send=OS}catch(e){}
}
function CLEAN(){
 if(ACK_TIMER){clearInterval(ACK_TIMER);ACK_TIMER=null}
 if(TIMEOUT){clearTimeout(TIMEOUT);TIMEOUT=null}
 if(POLL){clearInterval(POLL);POLL=null}
 RESTORE();
 if(ACTIVE())delete window.__CCMU_SCHEDULE_EXPORT_TOKEN__
}
function FAIL(m){
 if(!ACTIVE())return;
 DONE=true;
 NOTICE(m,"error",9000);
 CLEAN();
 window.removeEventListener("message",MSG);
 try{if(W&&!W.closed)W.close()}catch(e){}
}
function SEMESTER(){
 var root=document.querySelector("div#semesterId.ant-select")||document.querySelector("#semesterId.ant-select")||document.querySelector("[id=semesterId].ant-select");
 if(root){
  var v=root.querySelector(".ant-select-selection-selected-value");
  var s=((v&&v.getAttribute("title"))||(v&&v.textContent)||"").trim();
  var m=s.match(/20\\d{2}-20\\d{2}-[12]/);
  if(m)return m[0]
 }
 var label=document.querySelector('label[for="semesterId"]');
 if(label){
  var item=label.closest(".ant-form-item"),v2=item&&item.querySelector(".ant-select-selection-selected-value");
  var s2=((v2&&v2.getAttribute("title"))||(v2&&v2.textContent)||"").trim();
  var m2=s2.match(/20\\d{2}-20\\d{2}-[12]/);
  if(m2)return m2[0]
 }
 return ""
}
function QUERY_BUTTON(){
 var es=[].slice.call(document.querySelectorAll("form button,button"));
 for(var i=0;i<es.length;i++){
  var e=es[i];
  if(e.disabled)continue;
  if(TXT(e)==="查询"){
   if(e.classList&&e.classList.contains("ant-btn-background-ghost"))continue;
   return e
  }
 }
 return null
}
function OPEN_RECEIVER(){
 var u=G+(G.indexOf("?")>=0?"&":"?")+"receiver=1&token="+encodeURIComponent(TOKEN)+"&semester="+encodeURIComponent(SEM)+"&_="+Date.now();
 var sw=460,sh=390,l=Math.max(0,(screen.availWidth||screen.width||1200)-sw-24),tp=55;
 W=window.open(u,"ccmuScheduleReceiver","popup=yes,width="+sw+",height="+sh+",left="+l+",top="+tp+",resizable=yes,scrollbars=yes");
 if(!W){
  FAIL("浏览器阻止了接收窗口。请允许学校课表页弹出窗口后重新点击书签。");
  return false
 }
 try{W.blur();window.focus()}catch(e){}
 return true
}
function IS_QUERY_URL(u){return /queryStudentSchedule/i.test(String(u||""))}
function REQUEST_TEXT(url,body){
 var s=String(url||"");
 try{
  if(typeof body==="string")s+=" "+body;
  else if(body instanceof URLSearchParams)s+=" "+body.toString();
  else if(typeof FormData!=="undefined"&&body instanceof FormData){
   body.forEach(function(v,k){s+=" "+k+"="+String(v)})
  }
 }catch(e){}
 return s
}
function REQUEST_SEMESTER_OK(url,body){
 var a=REQUEST_TEXT(url,body).match(/20\\d{2}-20\\d{2}-[12]/g)||[];
 if(!a.length)return true;
 return a.indexOf(SEM)>=0
}
function RESPONSE_SEMESTERS(root){
 var out=[],seen=[];
 function walk(v,d){
  if(d>6||v===null||v===undefined)return;
  if(typeof v!=="object")return;
  if(seen.indexOf(v)>=0)return;seen.push(v);
  if(Array.isArray(v)){for(var i=0;i<v.length&&i<200;i++)walk(v[i],d+1);return}
  if(v.semesterId!==undefined&&v.semesterId!==null){
   var m=String(v.semesterId).match(/20\\d{2}-20\\d{2}-[12]/);
   if(m&&out.indexOf(m[0])<0)out.push(m[0])
  }
  var ks=Object.keys(v);
  for(var j=0;j<ks.length&&j<80;j++)walk(v[ks[j]],d+1)
 }
 walk(root,0);return out
}
function PARSE_RESPONSE(x){
 if(x&&typeof x==="object")return x;
 var s=String(x==null?"":x).trim();
 if(!s)throw new Error("接口响应为空");
 return JSON.parse(s)
}
function ACCEPT(raw){
 if(DONE||!ACTIVE())return;
 var payload;
 try{payload=PARSE_RESPONSE(raw)}catch(e){FAIL("已捕获 queryStudentSchedule，但响应不是有效 JSON："+e.message);return}
 var rs=RESPONSE_SEMESTERS(payload);
 if(rs.length&&rs.indexOf(SEM)<0){
  FAIL("已捕获接口响应，但其中 semesterId 为 "+rs.join("、")+"，与当前已选学期 "+SEM+" 不一致。为避免导出旧数据，本次已终止。");
  return
 }
 SEND(payload,"network")
}
function INSTALL(){
 if(typeof OF==="function"){
  FW=function(){
   var args=arguments,input=args[0],init=args[1]||{},url=typeof input==="string"?input:(input&&input.url)||"";
   var candidate=ACTIVE()&&ARMED&&Date.now()>=START&&IS_QUERY_URL(url);
   if(candidate){
    SEEN_REQUEST=true;
    if(!REQUEST_SEMESTER_OK(url,init.body)){
     setTimeout(function(){FAIL("本次 queryStudentSchedule 请求中的学期与已选学期 "+SEM+" 不一致，已阻止导出。")},0)
    }
   }
   var p=OF.apply(this,args);
   if(candidate){
    p.then(function(resp){
     try{resp.clone().text().then(ACCEPT).catch(function(e){FAIL("读取 queryStudentSchedule 响应失败："+e.message)})}
     catch(e){FAIL("读取 queryStudentSchedule 响应失败："+e.message)}
    }).catch(function(){})
   }
   return p
  };
  window.fetch=FW
 }
 XO=function(method,url){
  this.__ccmuV5={url:String(url||""),method:String(method||"GET")};
  return OO.apply(this,arguments)
 };
 XS=function(body){
  var meta=this.__ccmuV5||{},candidate=ACTIVE()&&ARMED&&Date.now()>=START&&IS_QUERY_URL(meta.url);
  if(candidate){
   SEEN_REQUEST=true;
   if(!REQUEST_SEMESTER_OK(meta.url,body)){
    setTimeout(function(){FAIL("本次 queryStudentSchedule 请求中的学期与已选学期 "+SEM+" 不一致，已阻止导出。")},0)
   }
   this.addEventListener("load",function(){
    if(DONE||!ACTIVE())return;
    try{
     var raw=this.responseType==="json"?this.response:this.responseText;
     ACCEPT(raw)
    }catch(e){FAIL("读取 queryStudentSchedule 响应失败："+e.message)}
   },{once:true});
   this.addEventListener("error",function(){
    if(!DONE&&ACTIVE())FAIL("本次 queryStudentSchedule 请求失败，请检查网络后重试。")
   },{once:true})
  }
  return OS.apply(this,arguments)
 };
 XP.open=XO;XP.send=XS
}
function SIG(){
 var tb=document.querySelector(".scheduleTable .ant-table-tbody")||document.querySelector(".ant-table-tbody");
 return tb?String(tb.innerText||tb.textContent||"").replace(/\\s+/g,"").slice(0,30000):""
}
function BUSY(){return !!document.querySelector(".scheduleTable .ant-spin-spinning,.scheduleTable .ant-spin-dot-spin,.ant-spin-spinning")}
function ROOM(x){
 var s=String(x||"").trim(),m=s.match(/^(.*)\\[([^\\]]+)\\]\\s*$/);
 return m?{room:m[1].trim()||"线上教学",weeks:m[2].trim()}:{room:s||"线上教学",weeks:""}
}
function CELL(td){
 var labs=[].slice.call(td.querySelectorAll("label")),out=[],cur=null;
 function push(){
  if(!cur)return;
  if(cur.className&&cur.weeks){
   if(!cur.teacherName)cur.teacherName="教师未提供";
   if(!cur.classroomName)cur.classroomName="线上教学";
   cur.semesterId=SEM;out.push(cur)
  }
  cur=null
 }
 for(var i=0;i<labs.length;i++){
  var lab=labs[i],svg=lab.querySelector("svg[data-icon]"),typ=svg?svg.getAttribute("data-icon"):"",txt=(lab.textContent||"").trim();
  if(!txt)continue;
  if(typ==="calculator"){push();cur={className:txt,teacherName:"",classroomName:"",weeks:"",semesterId:SEM}}
  else if(typ==="user"&&cur){cur.teacherName=txt}
  else if(typ==="environment"&&cur){var rr=ROOM(txt);cur.classroomName=rr.room;cur.weeks=rr.weeks}
 }
 push();return out
}
function SCRAPE(){
 var body=document.querySelector(".scheduleTable .ant-table-tbody")||document.querySelector(".ant-table-tbody");
 if(!body)throw new Error("未找到课表表格");
 var rows=[].slice.call(body.querySelectorAll(":scope > tr"));
 if(!rows.length)rows=[].slice.call(body.querySelectorAll("tr"));
 if(!rows.length)throw new Error("课表表格中没有节次行");
 var days=["monday","tuesday","wednesday","thursday","friday","saturday","sunday"],sec=[],occ=[];
 for(var r=0;r<12;r++){
  sec.push({key:r,section:"第"+(r+1)+"节",monday:[],tuesday:[],wednesday:[],thursday:[],friday:[],saturday:[],sunday:[]});
  occ.push([false,false,false,false,false,false,false])
 }
 rows.forEach(function(tr,ri){
  var p=parseInt(tr.getAttribute("data-row-key"),10);
  if(!(p>=0&&p<12)){var first=tr.querySelector("td"),mm=(first&&first.textContent||"").match(/第\\s*(\\d+)\\s*节/);p=mm?parseInt(mm[1],10)-1:ri}
  if(!(p>=0&&p<12))return;
  var tds=[].slice.call(tr.children).filter(function(x){return x.tagName==="TD"});
  if(!tds.length)return;tds=tds.slice(1);var dc=0;
  tds.forEach(function(td){
   while(dc<7&&occ[p][dc])dc++;if(dc>=7)return;
   var courses=CELL(td),span=parseInt(td.getAttribute("rowspan")||"1",10);if(!(span>0))span=1;
   for(var k=0;k<span&&p+k<12;k++){
    occ[p+k][dc]=true;
    for(var z=0;z<courses.length;z++)sec[p+k][days[dc]].push(Object.assign({},courses[z]))
   }
   dc++
  })
 });
 var count=0;sec.forEach(function(s){days.forEach(function(d){count+=s[d].length})});
 if(!count)throw new Error("新表格已刷新，但没有识别到课程");
 return {code:200,data:sec}
}
function SEND(payload,source){
 if(DONE||!ACTIVE())return;
 DONE=true;
 NOTICE(source==="network"
  ?"已收到学期 "+SEM+" 本次查询的接口响应，正在生成 Excel…"
  :"未捕获接口，但已确认表格内容在本次查询后发生变化；正在使用安全兜底生成 Excel…","info");
 var m={type:"CCMU_SCHEDULE_DATA",token:TOKEN,semester:SEM,source:source==="network"?"network":"dom-fallback",payload:payload};
 function post(){try{if(W&&!W.closed)W.postMessage(m,O)}catch(e){}}
 post();ACK_TIMER=setInterval(post,350);
 setTimeout(function(){if(ACK_TIMER){clearInterval(ACK_TIMER);ACK_TIMER=null}},15000);
 if(POLL){clearInterval(POLL);POLL=null}
 if(TIMEOUT){clearTimeout(TIMEOUT);TIMEOUT=null}
 RESTORE();
 TIMEOUT=setTimeout(function(){
  if(ACTIVE())FAIL("生成器在 30 秒内没有返回完成状态。请确认下载是否成功；如未成功，请重新导出。")
 },30000)
}
function MSG(e){
 if(e.origin!==O||!e.data||e.data.token!==TOKEN)return;
 if(e.data.type==="CCMU_SCHEDULE_READY"){try{W.blur();window.focus()}catch(x){}}
 if(e.data.type==="CCMU_SCHEDULE_ACK"&&ACK_TIMER){clearInterval(ACK_TIMER);ACK_TIMER=null}
 if(e.data.type==="CCMU_SCHEDULE_COMPLETE"){
  NOTICE("导出完成："+(e.data.fileName||"课表.xlsx"),"success",5000);
  CLEAN();window.removeEventListener("message",MSG);try{window.focus()}catch(x){}
 }
 if(e.data.type==="CCMU_SCHEDULE_ERROR"){FAIL("生成失败："+(e.data.message||"未知错误"))}
}

SEM=SEMESTER();
if(!SEM){
 NOTICE("未读取到“学年学期”的已选值。请先手动选择学期，再点击导出书签。","error",9000);
 CLEAN();return
}
var Q=QUERY_BUTTON();
if(!Q){
 NOTICE("未找到课表页面的“查询”按钮。请确认当前位于“课程管理 / 课表查看”页面。","error",9000);
 CLEAN();return
}

window.addEventListener("message",MSG);
if(!OPEN_RECEIVER())return;

BEFORE_SIG=SIG();
LAST_SIG=BEFORE_SIG;
LAST_CHANGE=Date.now();
INSTALL();
START=Date.now();
ARMED=true;
NOTICE("V5 正在查询学期 "+SEM+"。只会使用本次点击后返回的 queryStudentSchedule 数据；旧课表不会参与生成。","info");

try{Q.click()}catch(e){FAIL("自动点击“查询”失败："+e.message);return}

POLL=setInterval(function(){
 if(DONE||!ACTIVE())return;
 var now=Date.now(),sig=SIG();
 if(sig!==LAST_SIG){
  LAST_SIG=sig;LAST_CHANGE=now
 }

 if(!SEEN_REQUEST&&now-START>10000&&sig&&sig!==BEFORE_SIG&&!BUSY()&&now-LAST_CHANGE>700){
  try{SEND(SCRAPE(),"dom-fallback")}
  catch(e){FAIL("未捕获 queryStudentSchedule，且 DOM 安全兜底失败："+e.message)}
 }
},200);

TIMEOUT=setTimeout(function(){
 if(DONE||!ACTIVE())return;
 var sig=SIG();
 if(!SEEN_REQUEST&&sig&&sig!==BEFORE_SIG&&!BUSY()){
  try{SEND(SCRAPE(),"dom-fallback");return}
  catch(e){}
 }
 if(SEEN_REQUEST){
  FAIL("已经捕获到本次 queryStudentSchedule 请求，但 30 秒内没有得到可用响应。为避免使用旧课表，本次未生成文件。")
 }else{
  FAIL("30 秒内未捕获本次 queryStudentSchedule，且无法确认课表已由新数据覆盖。为避免误导出旧学期，本次未生成文件。")
 }
},30000)
})()`;

    return code.replace(/\n+/g, "").replace(/\s{2,}/g, " ");
}
