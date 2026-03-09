// ==UserScript==
// @name         NUIT 课表提取与 CSV 导出
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  拦截 kbxx 数据，处理周数和节次，并一键导出为可以直接用 Excel 打开的 CSV 格式
// @author       Jimmy
// @match        *://jww.nuit.edu.cn/xsgrkbcx!xsAllKbList.action*
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function() {
    'use strict';

    // CSV 输出字段顺序
    const CSV_FIELDNAMES = ["课程名称", "星期", "开始节数", "结束节数", "老师", "地点", "周数"];

    // 1. 将周数字符串压缩为区间形式 (对应 Python 的 compress_weeks)
    function compressWeeks(weekStr) {
        if (!weekStr) return "";
        // 解析、去重并排序
        let weeks = [...new Set(weekStr.split(',').map(x => parseInt(x.trim())).filter(x => !isNaN(x)))];
        weeks.sort((a, b) => a - b);

        if (weeks.length === 0) return "";

        let ranges = [];
        let start = weeks[0];
        let prev = weeks[0];

        for (let i = 1; i < weeks.length; i++) {
            let week = weeks[i];
            if (week === prev + 1) {
                prev = week;
            } else {
                ranges.push([start, prev]);
                start = week;
                prev = week;
            }
        }
        ranges.push([start, prev]);

        return ranges.map(r => r[0] === r[1] ? `${r[0]}` : `${r[0]}-${r[1]}`).join('、');
    }

    // 2. 解析节次并构建最终数据 (对应 Python 的 build_result & transform_course_data)
    function processCourseData(originalData) {
        return originalData.map(course => {
            // 解析上课节次，例如 "01,02" -> [1, 2]
            let sections = (course.jcdm2 || "").split(',')
                .map(x => parseInt(x.trim()))
                .filter(x => !isNaN(x));

            let minSection = sections.length > 0 ? Math.min(...sections) : "未知";
            let maxSection = sections.length > 0 ? Math.max(...sections) : "未知";

            return {
                "课程名称": course.kcmc || "未知",
                "星期": course.xq || "未知",
                "开始节数": String(minSection),
                "结束节数": String(maxSection),
                "老师": course.teaxms || "未知",
                "地点": course.jxcdmcs || "未知",
                "周数": compressWeeks(course.zcs)
            };
        });
    }

    // 3. 保存为 CSV (对应 Python 的 save_to_csv)
    function saveToCSV(data) {
        // \uFEFF 是 UTF-8 的 BOM 头，加上它 Excel 打开才不会中文乱码
        let csvContent = "\uFEFF" + CSV_FIELDNAMES.join(",") + "\n";

        data.forEach(row => {
            let rowArray = CSV_FIELDNAMES.map(field => {
                let value = row[field] ? String(row[field]) : "";
                // 处理 CSV 中的转义：如果包含逗号、双引号或换行符，需要用双引号包围，并将内部双引号翻倍
                if (value.includes(',') || value.includes('"') || value.includes('\n')) {
                    value = `"${value.replace(/"/g, '""')}"`;
                }
                return value;
            });
            csvContent += rowArray.join(",") + "\n";
        });

        // 创建下载链接并触发下载
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", "课程表.csv");
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // 4. 拦截页面数据并创建按钮
    const htmlContent = document.documentElement.innerHTML;
    const regex = /var\s+kbxx\s*=\s*(\[[\s\S]*?\]);/;
    const match = htmlContent.match(regex);

    if (match && match[1]) {
        try {
            const kbxxData = JSON.parse(match[1]);

            // 创建导出按钮
            const btn = document.createElement('button');
            btn.innerText = "📅 导出课表 (CSV)";
            btn.style.cssText = `
                position: fixed;
                bottom: 20px;
                right: 20px;
                padding: 12px 20px;
                background-color: #28a745;
                color: #fff;
                border: none;
                border-radius: 8px;
                cursor: pointer;
                z-index: 9999;
                font-weight: bold;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            `;

            btn.onclick = function() {
                const processedData = processCourseData(kbxxData);
                console.log("处理后的数据:", processedData);
                saveToCSV(processedData);
            };

            document.body.appendChild(btn);

        } catch (e) {
            console.error("❌ 解析课表数据失败:", e);
        }
    } else {
        console.warn("⚠️ 未能在页面源码中找到 kbxx 数据。");
    }
})();
