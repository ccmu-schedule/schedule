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
        // 直接解析用户输入的原始文本，如果格式不正确，直接抛出错误
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

async function generateExcel(jsonData) {
    // 1. 数据解析与重组
    const schedule = {};
    let maxWeek = 0;
    const daysOfWeek = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"];
    
    if (!jsonData.data || !Array.isArray(jsonData.data)) {
        throw new Error("JSON结构不符合预期，缺少顶层 'data' 数组");
    }
    
    jsonData.data.forEach((sectionData, sectionIndex) => {
        const currentPeriodIndex = sectionIndex;

        for (const dayName of daysOfWeek) {
            if (sectionData[dayName] && Array.isArray(sectionData[dayName])) {
                for (const course of sectionData[dayName]) {
                    if (!course.weeks || !course.className || !course.classroomName || !course.teacherName) continue;

                    const weekNums = String(course.weeks).split(',').map(Number);
                    if (weekNums.length > 0) {
                        const currentMax = Math.max(...weekNums);
                        if (currentMax > maxWeek) maxWeek = currentMax;
                    }

                    for (const week of weekNums) {
                        if (!schedule[week]) schedule[week] = {};
                        if (!schedule[week][dayName]) schedule[week][dayName] = Array(12).fill(null); // Use null for empty slots

                        // *** 创建或更新结构化对象 ***
                        const cellData = schedule[week][dayName][currentPeriodIndex];

                        if (!cellData) {
                            // 如果单元格为空，创建新的对象
                            schedule[week][dayName][currentPeriodIndex] = {
                                className: course.className,
                                classroomName: course.classroomName,
                                teachers: [course.teacherName] // 上课教师数组
                            };
                        } else {
                            // 如果单元格已有对象，只更新教师列表
                            // 检查教师是否已存在，防止重复添加
                            if (!cellData.teachers.includes(course.teacherName)) {
                                cellData.teachers.push(course.teacherName);
                            }
                        }
                    }
                }
            }
        }
    });
    
    // 2. 创建和设置Excel工作簿
    const wb = new ExcelJS.Workbook();
    const headers = ["时间/节次", "周一", "周二", "周三", "周四", "周五", "周六", "周日"];
    const times = [
        "8:00-8:45", "8:45-9:30", "9:45-10:30", "10:30-11:15", "11:25-12:10",
        "13:30-14:15", "14:15-15:00", "15:10-15:55", "15:55-16:40",
        "18:00-18:45", "18:45-19:30", "19:30-20:15"
    ];

    for (let weekNum = 1; weekNum <= maxWeek; weekNum++) {
        const ws = wb.addWorksheet(`第${weekNum}周`);
        ws.addRow(headers);
        for (let i = 0; i < 12; i++) {
            const rowHeader = `第${i + 1}节\n${times[i]}`;
            const rowData = [rowHeader];
            for (const dayName of daysOfWeek) {
                const daySchedule = (schedule[weekNum] && schedule[weekNum][dayName]) ? schedule[weekNum][dayName] : [];
                const cellData = daySchedule[i]; 

                // *** 格式化对象为最终字符串 ***
                if (cellData) {
                    const formattedString = `${cellData.className}\n${cellData.classroomName}\n${cellData.teachers.join('、')}`;
                    rowData.push(formattedString);
                } else {
                    rowData.push(""); 
                }
            }
            ws.addRow(rowData);
        }

        // 3. 设置基础样式
        ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: true }, (cell) => {
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                if (rowNumber === 1) {
                    cell.font = { name: '等线', size: 12, bold: true };
                } else {
                    cell.font = { name: '等线', size: 11 };
                }
            });
        });
        
        // 4. 动态调整所有单元格大小
        // a. 自动调整列宽
        ws.columns.forEach(column => {
            let maxCharLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                const cellText = cell.value ? cell.value.toString() : '';
                const lines = cellText.split('\n');
                lines.forEach(line => {
                    let lineLength = 0;
                    for (let i = 0; i < line.length; i++) {
                        lineLength += line.charCodeAt(i) > 255 ? 2 : 1;
                    }
                    if (lineLength > maxCharLength) {
                        maxCharLength = lineLength;
                    }
                });
            });
            column.width = Math.max(15, maxCharLength * 1.2);
        });

        // b. 自动调整行高
        ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            let maxLines = 1;
            row.eachCell({ includeEmpty: true }, cell => {
                const cellText = cell.value ? cell.value.toString() : '';
                const numLines = cellText.split('\n').length;
                if (numLines > maxLines) {
                    maxLines = numLines;
                }
            });
            if (rowNumber === 1) {
                row.height = 25;
            } else {
                row.height = maxLines * 22;
            }
        });
    }

    // 5. 生成文件并触发下载
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'course_schedule.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
