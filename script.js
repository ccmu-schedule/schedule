document.getElementById('generateBtn').addEventListener('click', async () => {
    const jsonInput = document.getElementById('jsonInput').value;
    const statusEl = document.getElementById('status');

    if (!jsonInput.trim()) {
        statusEl.textContent = '错误：JSON内容不能为空！';
        statusEl.className = 'error';
        return;
    }

    statusEl.textContent = '正在解析和生成中，请稍候...';
    statusEl.className = '';

    try {
        const jsonData = JSON.parse(jsonInput.trim());
        
        await generateExcel(jsonData);
        
        statusEl.textContent = '课表生成成功！已开始下载...';
        statusEl.className = 'success';

    } catch (error) {
        console.error("生成课表时发生错误:", error);
        
        let errorMessage = error.message;
        
        if (error instanceof SyntaxError) {
            errorMessage = 'JSON格式错误，请检查是否复制完整...';
        }
        
        statusEl.textContent = `生成失败：${errorMessage}`;
        statusEl.className = 'error';
    }
});

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

    // 生成文件并触发下载
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
