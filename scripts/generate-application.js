#!/usr/bin/env node
/**
 * Dream Big 元大公益圓夢計畫申請書生成器
 * 
 * 使用方式：
 * node scripts/generate-application.js <json-data-file> <output-file>
 * 
 * JSON 資料檔案格式請參考 references/fields-guide.md
 */

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, WidthType, BorderStyle, ShadingType,
        HeadingLevel, PageBreak } = require('docx');
const fs = require('fs');

// 表格邊框樣式
const border = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };

// 標題儲存格樣式（淺藍色背景）
const headerCellStyle = {
    borders,
    shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 }
};

// 內容儲存格樣式
const contentCellStyle = {
    borders,
    shading: { fill: "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 }
};

// SDGs 選項
const sdgsOptions = [
    "1消除貧窮", "2消除飢餓", "3健康福祉", "4教育品質", "5性別平等", "6環境品質",
    "7可負擔能源", "8就業與經濟成長", "9永續運輸", "10減少不平等", "11永續城市",
    "12責任消費與生產", "13氣候行動", "14海洋生態", "15陸地生態", "16和平與正義制度",
    "17全球夥伴"
];

/**
 * 建立標題儲存格
 */
function createHeaderCell(text, width) {
    return new TableCell({
        ...headerCellStyle,
        width: { size: width, type: WidthType.DXA },
        children: [new Paragraph({
            children: [new TextRun({ text, bold: true, size: 22, font: "標楷體" })]
        })]
    });
}

/**
 * 建立內容儲存格
 */
function createContentCell(text, width) {
    const lines = text ? text.split('\n') : [''];
    return new TableCell({
        ...contentCellStyle,
        width: { size: width, type: WidthType.DXA },
        children: lines.map(line => new Paragraph({
            children: [new TextRun({ text: line, size: 22, font: "標楷體" })]
        }))
    });
}

/**
 * 建立計畫內容表
 */
function createProjectTable(data) {
    const tableWidth = 9360;
    const headerWidth = 3000;
    const contentWidth = 6360;

    const rows = [
        // 服務區域概況及需求說明
        {
            header: "服務區域概況及需求說明\n（請說明與計畫相關之區域現況及弱勢程度描述，以及為何需要Dream Big元大公益圓夢計畫協助，亦可提供相關統計數據輔助說明。）",
            content: data.serviceAreaDescription || ""
        },
        // 服務目的
        {
            header: "服務目的（預期藉由計畫解決之問題）",
            content: data.servicePurpose || ""
        },
        // 服務項目
        {
            header: "請條列單位目前「所有」服務項目與內容說明\n（除列出項目外，敬請量化說明服務內容，例如：每周針對偏鄉學童進行2次才藝課程）",
            content: data.currentServices || ""
        },
        // SDGs
        {
            header: "請勾選本次計畫可對應到之聯合國永續發展目標(Sustainable Development Goals, SDGs)",
            content: formatSDGs(data.sdgs || [])
        },
        // 申請項目詳述
        {
            header: "請詳述預計申請Dream Big元大公益圓夢計畫經費所進行的服務項目，包括內容規劃及時程",
            content: data.projectDetails || ""
        },
        // 過往效益
        {
            header: "呈上題，請說明此申請項目，過去所產生的「社會效益」或「正面影響力」為何，若無則可填無。",
            content: data.pastImpact || ""
        },
        // 企業連結
        {
            header: "期望元大金控集團暨子公司之企業志工、元大文教基金會或其他資源為您的計畫提供什麼協助或合作？預期如何與元大金控集團共同合作與連結？",
            content: data.corporateConnection || ""
        },
        // 預期成效
        {
            header: "預期計畫成效，請說明獲得Dream Big元大公益圓夢計畫之經費與志工等資源後，預計達成之至少三項目標。敬請以質化與量化資料呈現",
            content: data.expectedOutcomes || ""
        },
        // 成果發表
        {
            header: "請說明計畫成果預計發表方式與完成時間",
            content: data.presentationPlan || ""
        },
        // 宣傳資源
        {
            header: "單位現有之宣傳資源或管道，以及可因應此計畫使用之宣傳方式",
            content: data.promotionResources || ""
        }
    ];

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [headerWidth, contentWidth],
        rows: rows.map(row => new TableRow({
            children: [
                createHeaderCell(row.header, headerWidth),
                createContentCell(row.content, contentWidth)
            ]
        }))
    });
}

/**
 * 格式化 SDGs 選項
 */
function formatSDGs(selected) {
    return sdgsOptions.map(opt => {
        const num = parseInt(opt);
        const checked = selected.includes(num) ? "☑" : "□";
        return `${checked}${opt}`;
    }).join(" ");
}

/**
 * 建立預算表
 */
function createBudgetTable(budgetItems) {
    const colWidths = [2500, 1500, 1200, 1800, 2360];
    
    const headerRow = new TableRow({
        children: [
            createHeaderCell("項目", colWidths[0]),
            createHeaderCell("單價", colWidths[1]),
            createHeaderCell("數量", colWidths[2]),
            createHeaderCell("金額", colWidths[3]),
            createHeaderCell("備註", colWidths[4])
        ]
    });

    const dataRows = (budgetItems || []).map(item => new TableRow({
        children: [
            createContentCell(item.name || "", colWidths[0]),
            createContentCell(item.unitPrice ? item.unitPrice.toString() : "", colWidths[1]),
            createContentCell(item.quantity ? item.quantity.toString() : "", colWidths[2]),
            createContentCell(item.amount ? item.amount.toString() : "", colWidths[3]),
            createContentCell(item.note || "", colWidths[4])
        ]
    }));

    // 計算總計
    const total = (budgetItems || []).reduce((sum, item) => sum + (item.amount || 0), 0);
    
    const totalRow = new TableRow({
        children: [
            new TableCell({
                ...headerCellStyle,
                width: { size: colWidths[0] + colWidths[1] + colWidths[2], type: WidthType.DXA },
                columnSpan: 3,
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "預算金額總計", bold: true, size: 22, font: "標楷體" })]
                })]
            }),
            createContentCell(total.toString(), colWidths[3]),
            createContentCell("", colWidths[4])
        ]
    });

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: colWidths,
        rows: [headerRow, ...dataRows, totalRow]
    });
}

/**
 * 建立出席人員表
 */
function createAttendeeTable(attendees) {
    const colWidths = [2800, 1500, 1500, 1700, 1860];
    const meetings = [
        { key: "kickoff", name: "期初發表會", count: 2 },
        { key: "final", name: "期末發表會", count: 2 },
        { key: "progress1", name: "第一次\n計畫進度會議\n暨圓夢工作坊", count: 2 },
        { key: "progress2", name: "第二次\n計畫進度會議", count: 2 }
    ];

    const headerRow = new TableRow({
        children: [
            createHeaderCell("項目", colWidths[0]),
            createHeaderCell("姓名", colWidths[1]),
            createHeaderCell("職稱", colWidths[2]),
            createHeaderCell("電話", colWidths[3]),
            createHeaderCell("email", colWidths[4])
        ]
    });

    const dataRows = [];
    meetings.forEach(meeting => {
        const meetingAttendees = attendees?.[meeting.key] || [{}, {}];
        for (let i = 0; i < meeting.count; i++) {
            const person = meetingAttendees[i] || {};
            dataRows.push(new TableRow({
                children: [
                    createHeaderCell(i === 0 ? meeting.name : "", colWidths[0]),
                    createContentCell(person.name || "", colWidths[1]),
                    createContentCell(person.title || "", colWidths[2]),
                    createContentCell(person.phone || "", colWidths[3]),
                    createContentCell(person.email || "", colWidths[4])
                ]
            }));
        }
    });

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: colWidths,
        rows: [headerRow, ...dataRows]
    });
}

/**
 * 建立單位概況表
 */
function createOrgInfoTable(org) {
    const tableWidth = 9360;
    const col1 = 2000, col2 = 2680, col3 = 1500, col4 = 3180;

    const rows = [
        // 第一行：全銜
        new TableRow({
            children: [
                createHeaderCell("計畫申請人/單位\n(全銜)", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: org.fullName || "", size: 22, font: "標楷體" })]
                    })]
                })
            ]
        }),
        // 成立時間、核准機關
        new TableRow({
            children: [
                createHeaderCell("成立時間", col1),
                createContentCell(org.establishedDate || "", col2),
                createHeaderCell("核准機關", col3),
                createContentCell(org.approvalAuthority || "", col4)
            ]
        }),
        // 立案字號
        new TableRow({
            children: [
                createHeaderCell("立案字號", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: org.registrationNumber || "", size: 22, font: "標楷體" })]
                    })]
                })
            ]
        }),
        // 立案地址
        new TableRow({
            children: [
                createHeaderCell("立案地址", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: org.registrationAddress || "", size: 22, font: "標楷體" })]
                    })]
                })
            ]
        }),
        // 通訊地址
        new TableRow({
            children: [
                createHeaderCell("通訊地址", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: org.mailingAddress || "", size: 22, font: "標楷體" })]
                    })]
                })
            ]
        }),
        // 負責人
        new TableRow({
            children: [
                createHeaderCell("單位負責人姓名", col1),
                createContentCell(org.directorName || "", col2),
                createHeaderCell("職稱", col3),
                createContentCell(org.directorTitle || "", col4)
            ]
        }),
        // 聯絡人
        new TableRow({
            children: [
                createHeaderCell("聯絡人姓名", col1),
                createContentCell(org.contactName || "", col2),
                createHeaderCell("職稱", col3),
                createContentCell(org.contactTitle || "", col4)
            ]
        }),
        // 聯絡電話
        new TableRow({
            children: [
                createHeaderCell("聯絡人手機", col1),
                createContentCell(org.contactMobile || "", col2),
                createHeaderCell("室內電話", col3),
                createContentCell(org.contactPhone || "", col4)
            ]
        }),
        // 傳真與 Email
        new TableRow({
            children: [
                createHeaderCell("傳真", col1),
                createContentCell(org.fax || "", col2),
                createHeaderCell("聯絡人E-mail", col3),
                createContentCell(org.contactEmail || "", col4)
            ]
        }),
        // 官網/社群
        new TableRow({
            children: [
                createHeaderCell("單位官網/社群", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: org.website || "", size: 22, font: "標楷體" })]
                    })]
                })
            ]
        }),
        // 服務對象
        new TableRow({
            children: [
                createHeaderCell("單位主要服務對象", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: formatServiceTargets(org.serviceTargets), size: 22, font: "標楷體" })]
                    })]
                })
            ]
        }),
        // 經費說明
        new TableRow({
            children: [
                createHeaderCell("單位經費說明", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: formatFinancialInfo(org.financial), size: 22, font: "標楷體" })]
                    })]
                })
            ]
        }),
        // 人事概況
        new TableRow({
            children: [
                createHeaderCell("單位人事概況", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ 
                            text: `全職人員${org.fullTimeStaff || '___'}人 / 兼職人員${org.partTimeStaff || '___'}人 / 固定志工${org.volunteers || '___'}人`, 
                            size: 22, 
                            font: "標楷體" 
                        })]
                    })]
                })
            ]
        }),
        // 組織圖
        new TableRow({
            children: [
                createHeaderCell("單位組織圖", col1),
                new TableCell({
                    ...contentCellStyle,
                    width: { size: col2 + col3 + col4, type: WidthType.DXA },
                    columnSpan: 3,
                    children: [new Paragraph({
                        children: [new TextRun({ text: org.orgChart || "※請用文字說明、或可用圖表呈現，並清楚填寫姓名。", size: 22, font: "標楷體" })]
                    })]
                })
            ]
        })
    ];

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [col1, col2, col3, col4],
        rows
    });
}

/**
 * 格式化服務對象資訊
 */
function formatServiceTargets(targets) {
    if (!targets) return "";
    let result = [];
    if (targets.children) {
        result.push(`☑ 兒童${targets.children.count || '___'}人(0-12歲) - ${(targets.children.types || []).join('、')}`);
    }
    if (targets.youth) {
        result.push(`☑ 青少年${targets.youth.count || '___'}人(13-18歲) - ${(targets.youth.types || []).join('、')}`);
    }
    if (targets.adults) {
        result.push(`☑ 成人${targets.adults.count || '___'}人(19-65歲) - ${(targets.adults.types || []).join('、')}`);
    }
    if (targets.elderly) {
        result.push(`☑ 老人${targets.elderly.count || '___'}人(65歲以上) - ${(targets.elderly.types || []).join('、')}`);
    }
    if (targets.other) {
        result.push(`☑ 其他${targets.other.count || '___'}人 - ${targets.other.description || ''}`);
    }
    return result.join('\n');
}

/**
 * 格式化經費資訊
 */
function formatFinancialInfo(financial) {
    if (!financial) return "";
    let lines = [];
    if (financial.totalAssets) lines.push(`總資產或資本額 ${financial.totalAssets} 元`);
    if (financial.annualRevenue) lines.push(`最近一年年營收(含捐贈) ${financial.annualRevenue} 元`);
    
    let sources = [];
    if (financial.governmentGrant) sources.push(`政府補助 ${financial.governmentGrant}%`);
    if (financial.npoGrant) sources.push(`其他非營利組織 ${financial.npoGrant}%`);
    if (financial.serviceIncome) sources.push(`服務收費 ${financial.serviceIncome}%`);
    if (financial.businessIncome) sources.push(`社會事業收入 ${financial.businessIncome}%`);
    if (financial.corporateSponsorship) {
        sources.push(`企業贊助 ${financial.corporateSponsorship.percentage}%（${financial.corporateSponsorship.companies}）`);
    }
    if (financial.other) sources.push(`其他 ${financial.other.percentage}%（${financial.other.description}）`);
    
    if (sources.length > 0) {
        lines.push('經費來源：' + sources.join('、'));
    }
    return lines.join('\n');
}

/**
 * 主程式：生成文件
 */
async function generateDocument(data, outputPath) {
    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: { font: "標楷體", size: 24 }
                }
            },
            paragraphStyles: [
                {
                    id: "Title",
                    name: "Title",
                    basedOn: "Normal",
                    run: { size: 36, bold: true, font: "標楷體" },
                    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 200 } }
                },
                {
                    id: "Subtitle",
                    name: "Subtitle",
                    basedOn: "Normal",
                    run: { size: 28, bold: true, font: "標楷體" },
                    paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 400 } }
                }
            ]
        },
        sections: [{
            properties: {
                page: {
                    size: { width: 11906, height: 16838 },
                    margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
                }
            },
            children: [
                // 標題
                new Paragraph({
                    style: "Title",
                    children: [new TextRun({ text: "Dream Big元大公益圓夢計畫", bold: true })]
                }),
                new Paragraph({
                    style: "Subtitle",
                    children: [new TextRun({ text: "第九屆單位申請書", bold: true })]
                }),
                
                // 申請單位資訊
                new Paragraph({
                    spacing: { before: 400, after: 200 },
                    children: [new TextRun({ text: `申請單位全銜：${data.organization?.fullName || '________________________________________'}`, size: 24 })]
                }),
                new Paragraph({
                    spacing: { after: 400 },
                    children: [new TextRun({ text: `申請計畫：${data.projectName || '________________________________________'}`, size: 24 })]
                }),

                // 計畫內容表
                new Paragraph({
                    spacing: { before: 400, after: 200 },
                    children: [new TextRun({ text: "【計畫內容表】", bold: true, size: 28 })]
                }),
                createProjectTable(data),
                
                new Paragraph({ children: [new PageBreak()] }),

                // 預算表
                new Paragraph({
                    spacing: { before: 400, after: 200 },
                    children: [new TextRun({ text: "【計畫預算表】", bold: true, size: 28 })]
                }),
                new Paragraph({
                    spacing: { after: 100 },
                    children: [new TextRun({ text: "（含活動費用、場地租借、講師費用、交通費、雜支及預估往返台北市之會議出席車資等）", size: 20 })]
                }),
                createBudgetTable(data.budget),
                new Paragraph({
                    spacing: { before: 100 },
                    children: [new TextRun({ text: "【註】以上項目可視需求自行加列，相關會議與活動均依衛生福利部疾病管制署規定或實際狀況彈性辦理", size: 18 })]
                }),

                // 出席人員表
                new Paragraph({
                    spacing: { before: 400, after: 200 },
                    children: [new TextRun({ text: "【預計出席人員代表】", bold: true, size: 28 })]
                }),
                createAttendeeTable(data.attendees),
                new Paragraph({
                    spacing: { before: 100 },
                    children: [new TextRun({ text: "【註】相關會議與活動均依衛生福利部疾病管制署規定或實際狀況彈性辦理", size: 18 })]
                }),

                new Paragraph({ children: [new PageBreak()] }),

                // 單位概況表
                new Paragraph({
                    spacing: { before: 400, after: 200 },
                    children: [new TextRun({ text: "【單位概況表】", bold: true, size: 28 })]
                }),
                createOrgInfoTable(data.organization || {}),

                // 提醒事項
                new Paragraph({
                    spacing: { before: 600, after: 200 },
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "★ 提醒事項 ★", bold: true, size: 28 })]
                }),
                new Paragraph({
                    spacing: { after: 100 },
                    children: [new TextRun({ 
                        text: "1. 此表完成後檔名請以「單位名稱_Dream Big元大公益圓夢計畫」命名，若需附上佐證資料(佐證資料可為圖表、照片)，請妥善統整後壓縮為PDF格式，電子檔寄至專屬收件信箱 dreambig@in-harmony.com.tw", 
                        size: 20 
                    })]
                }),
                new Paragraph({
                    spacing: { after: 100 },
                    children: [new TextRun({ 
                        text: "2. 申請資料單封信件大小上限為10M，若有檔案較大之影音檔案，可上傳Google雲端硬碟後，將雲端連結附在E-mail信件文字中提供。", 
                        size: 20 
                    })]
                }),
                new Paragraph({
                    spacing: { after: 100 },
                    children: [new TextRun({ 
                        text: "3. 各項報名資料，應於114年9月1日16時00分以前繳交，以完成報名程序。", 
                        size: 20 
                    })]
                }),
                new Paragraph({
                    spacing: { after: 100 },
                    children: [new TextRun({ 
                        text: "4. 填答諮詢可來電本計畫協辦單位「合拍創意行銷有限公司」—Dream Big元大公益圓夢小組(02)6605-0633分機201莊小姐。", 
                        size: 20 
                    })]
                })
            ]
        }]
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outputPath, buffer);
    console.log(`申請書已生成：${outputPath}`);
}

// 執行
if (require.main === module) {
    const args = process.argv.slice(2);
    if (args.length < 2) {
        console.log("使用方式: node generate-application.js <json-data-file> <output-file>");
        console.log("範例: node generate-application.js application-data.json output.docx");
        process.exit(1);
    }
    
    const data = JSON.parse(fs.readFileSync(args[0], 'utf8'));
    generateDocument(data, args[1]).catch(console.error);
}

module.exports = { generateDocument, createProjectTable, createBudgetTable, createAttendeeTable, createOrgInfoTable };
