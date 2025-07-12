import * as fs from "fs";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } from "docx";
import mysql from "mysql2/promise";
import path from "path";

// Helper to generate a unique filename if file exists
function getUniqueFilename(baseName, ext = ".docx") {
    let filename = baseName + ext;
    let counter = 1;
    while (fs.existsSync(filename)) {
        filename = `${baseName}_${counter}${ext}`;
        counter++;
    }
    return filename;
}

// 1. Connect to MariaDB and fetch data
async function fetchData() {
    const connection = await mysql.createConnection({
        host: "45.150.24.108",
        port: 23306,
        user: "lesa_prod",
        password: "lesa_prod$34133$$",
        database: "lesaapp"
    });

    const [rows] = await connection.execute("Call get__refunded_orders_data(372)");
    const [rows2] = await connection.execute("Call get__refunded_order_payment_data(372)");
    // rows2[0] is the actual data for stored procedures in mysql2
    const paymentData = Array.isArray(rows2) && Array.isArray(rows2[0]) ? rows2[0][0] : (rows2[0] || {});
    await connection.end();
    // rows[0] is the actual data for stored procedures in mysql2
    return {
        dbRows: Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0] : rows,
        paymentData
    };
}

// 2. Generate docx from data
async function generateDocx() {
    // Fetch data from DB
    const { dbRows, paymentData } = await fetchData();

    // Example static data for landlord/tenant
    const landlord = {
        name: paymentData?.vendor_display_name,
        address: paymentData?.vendor_addres,
        phone: [paymentData?.vendor_phone, paymentData?.vendor_phone_second]
            .filter(p => p && p !== 'null')
            .join('\n')
    };
    const tenant = {
        name: paymentData?.customer_display_name,
        address: paymentData?.customer_addres,
        phone: [paymentData?.customer_phone, paymentData?.customer_phone_second]
            .filter(p => p && p !== 'null')
            .join('\n')
    };

    // Split data into yuborilganlar (is_refund == 0) and qaytganlar (is_refund == 1)
    const yuborilganlar = dbRows.filter(row => row.is_refund == null);
    const qaytganlar = dbRows.filter(row => row.is_refund == 1);

    // Helper to format bundle
    function bundleText(is_bundle) {
        return is_bundle == 1 ? " (to'plam)" : "";
    }

    // Helper to format number with "ta"
    function qtyText(qty) {
        return qty !== null && qty !== undefined ? `${qty} ta` : "";
    }

    // Helper to format price
    function priceText(price) {
        if (price === null || price === undefined) return "0 so'm";
        return `${Number(price).toLocaleString()} so'm`;
    }
    // Helper to format total price
    function totalPriceText(total) {
        if (total === null || total === undefined) return "0 so'm";
        return `${Number(total).toLocaleString()} so'm`;
    }
    // Helper to format date
    function dateText(date) {
        if (!date) return "";
        const d = new Date(date);
        if (isNaN(d)) return "";
        return d.toISOString().slice(0, 10);
    }

    // Helper to format used days
    function usedDaysText(days) {
        return days !== null && days !== undefined ? `${days} kun` : "";
    }

    // Yuborilganlar table rows (add yuborilgan sanasi)
    // Helper to format date as dd-mm--yyyy
    function dateTextDDMMYYYY(date) {
        if (!date) return "";
        const d = new Date(date);
        if (isNaN(d)) return "";
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();
        return `${day}-${month}-${year}`;
    }

    const yuborilganlarRows = yuborilganlar.map((row, idx) => [
        String(idx + 1),
        row.post_title + bundleText(row.is_bundle),
        qtyText(row.product_qty),
        dateTextDDMMYYYY(row.start_date), // yuborilgan sanasi
        priceText(row.price) // kunlik narxi
    ]);

    // Qaytganlar table rows (add qaytgan sanasi)
    const qaytganlarRows = qaytganlar.map((row, idx) => [
        String(idx + 1),
        row.post_title + bundleText(row.is_bundle),
        qtyText(Math.abs(row.product_qty)),
        dateTextDDMMYYYY(row.end_date), // qaytgan sanasi
        usedDaysText(row.used_days), // ishlatilgan kuni
        priceText(row.price), // kunlik narxi
        totalPriceText(row.price * row.used_days) // umumiy narxi
    ]);

    // Yo'qotilganlar table rows (is_refund == null && is_bundle == null)
    const yoqotilganlar = dbRows.filter(row => row.is_refund == null && row.is_bundle == null);
    const yoqotilganlarRows = yoqotilganlar.map((row, idx) => [
        String(idx + 1),
        row.post_title,
        qtyText(row.lost_qty),
        priceText(row.regular_price),
        totalPriceText(row.lost_qty * row.regular_price)
    ]);

    // Totals from paymentData
    // total_payment_amount as To'langan
    // delivery_price as yetkazib berish narxi
    // lost_debt_price as yo'qotilgan tovarlar uchun shtraf
    function formatSum(val) {
        if (val === null || val === undefined) return "0 so'm";
        return `${Number(val).toLocaleString()} so'm`;
    }
    const totals = [
        ["To'langan", formatSum(paymentData?.total_payment_amount)],
        ["Yetkazib berish narxi", formatSum(paymentData?.delivery_price)],
        ["Yo'qotilgan tovarlar uchun shtraf", formatSum(paymentData?.lost_debt_price)],
    ];

    const doc = new Document({
        sections: [
            {
                children: [
                    // Title
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: "TO'LOV QOG'OZI",
                                bold: true,
                                size: 36, // 18pt
                                font: "Times New Roman",
                            }),
                        ],
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 200 },
                    }),
                    // Header Table (Landlord & Tenant)
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            // First row: Section headers
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "Ijaraga beruvchi", bold: true, size: 24, font: "Times New Roman" }),
                                                ],
                                                alignment: AlignmentType.CENTER,
                                            }),
                                        ],
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: "Ijaraga oluvchi", bold: true, size: 24, font: "Times New Roman" }),
                                                ],
                                                alignment: AlignmentType.CENTER,
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            // Second row: Names
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: `« ${landlord.name} »`, bold: true, size: 24, font: "Times New Roman" }),
                                                ],
                                            }),
                                        ],
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: tenant.name, bold: true, size: 24, font: "Times New Roman" }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            // Third row: Addresses
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: `Manzili: ${landlord.address}`, size: 22, font: "Times New Roman" }),
                                                ],
                                            }),
                                        ],
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({ text: `Obyekt manzili: ${tenant.address}`, size: 22, font: "Times New Roman" }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            // Fifth row: Phones
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            ...landlord.phone.split('\n').map(phone =>
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({ text: "Tel.: ", size: 22, font: "Times New Roman" }),
                                                        new TextRun({ text: phone, bold: true, size: 22, font: "Times New Roman" }),
                                                    ],
                                                })
                                            ),
                                        ],
                                    }),
                                    new TableCell({
                                        children: [
                                            ...tenant.phone.split('\n').map(phone =>
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({ text: "Tel.: ", size: 22, font: "Times New Roman" }),
                                                        new TextRun({ text: phone, size: 22, font: "Times New Roman" }),
                                                    ],
                                                })
                                            ),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),
                    new Paragraph({ text: "" }), // Spacer

                    // Yuborilganlar Table
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Yuborilganlar", bold: true, size: 28, font: "Times New Roman" })
                        ],
                        spacing: { after: 100 }
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    ...["№", "Mahsulot nomi", "Soni", "Yuborilgan sanasi", "Kunlik narxi"].map(header =>
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({ text: header, bold: true, size: 22, font: "Times New Roman" })
                                                    ],
                                                    alignment: AlignmentType.CENTER,
                                                })
                                            ],
                                            shading: { fill: "D9D9D9" },
                                            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "000000" } }
                                        })
                                    )
                                ],
                                tableHeader: true,
                            }),
                            ...(yuborilganlarRows.length > 0
                                ? yuborilganlarRows.map(row =>
                                    new TableRow({
                                        children: row.map(val =>
                                            new TableCell({
                                                children: [
                                                    new Paragraph({
                                                        text: val,
                                                        alignment: AlignmentType.CENTER
                                                    })
                                                ]
                                            })
                                        )
                                    })
                                )
                                : [
                                    new TableRow({
                                        children: [
                                            new TableCell({
                                                children: [new Paragraph({ text: "Ma'lumot yo'q", alignment: AlignmentType.CENTER })],
                                                columnSpan: 6
                                            })
                                        ]
                                    })
                                ]
                            )
                        ]
                    }),
                    new Paragraph({ text: "" }), // Spacer

                    // Qaytganlar Table
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Qaytganlar", bold: true, size: 28, font: "Times New Roman" })
                        ],
                        spacing: { after: 100 }
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    ...["№", "Mahsulot nomi", "Soni", "Qaytgan sanasi", "Ishlatilgan kuni", "Kunlik narxi", "Umumiy narxi"].map(header =>
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({ text: header, bold: true, size: 22, font: "Times New Roman" })
                                                    ],
                                                    alignment: AlignmentType.CENTER,
                                                })
                                            ],
                                            shading: { fill: "F2DCDB" },
                                            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "000000" } }
                                        })
                                    )
                                ],
                                tableHeader: true,
                            }),
                            ...(qaytganlarRows.length > 0
                                ? qaytganlarRows.map(row =>
                                    new TableRow({
                                        children: row.map(val =>
                                            new TableCell({
                                                children: [
                                                    new Paragraph({
                                                        text: val,
                                                        alignment: AlignmentType.CENTER
                                                    })
                                                ]
                                            })
                                        )
                                    })
                                )
                                : [
                                    new TableRow({
                                        children: [
                                            new TableCell({
                                                children: [new Paragraph({ text: "Ma'lumot yo'q", alignment: AlignmentType.CENTER })],
                                                columnSpan: 6
                                            })
                                        ]
                                    })
                                ]
                            )
                        ]
                    }),
                    new Paragraph({ text: "" }), // Spacer

                    // Yo'qotilganlar Table
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Yo'qotilganlar", bold: true, size: 28, font: "Times New Roman" })
                        ],
                        spacing: { after: 100 }
                    }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    ...["№", "Mahsulot nomi", "Soni", "Narxi", "Umumiy narxi"].map(header =>
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({ text: header, bold: true, size: 22, font: "Times New Roman" })
                                                    ],
                                                    alignment: AlignmentType.CENTER,
                                                })
                                            ],
                                            shading: { fill: "FFE699" },
                                            borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "000000" } }
                                        })
                                    )
                                ],
                                tableHeader: true,
                            }),
                            ...(yoqotilganlarRows.length > 0
                                ? yoqotilganlarRows.map(row =>
                                    new TableRow({
                                        children: row.map(val =>
                                            new TableCell({
                                                children: [
                                                    new Paragraph({
                                                        text: val,
                                                        alignment: AlignmentType.CENTER
                                                    })
                                                ]
                                            })
                                        )
                                    })
                                )
                                : [
                                    new TableRow({
                                        children: [
                                            new TableCell({
                                                children: [new Paragraph({ text: "Ma'lumot yo'q", alignment: AlignmentType.CENTER })],
                                                columnSpan: 5
                                            })
                                        ]
                                    })
                                ]
                            )
                        ]
                    }),
                    new Paragraph({ text: "" }), // Spacer

                    // Totals Table
                    new Table({
                        width: { size: 50, type: WidthType.PERCENTAGE },
                        rows: [
                            ...totals.map(([label, value]) =>
                                new TableRow({
                                    children: [
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({ text: label, size: 22, font: "Times New Roman" })
                                                    ]
                                                })
                                            ]
                                        }),
                                        new TableCell({
                                            children: [
                                                new Paragraph({
                                                    children: [
                                                        new TextRun({ text: value, size: 22, font: "Times New Roman" })
                                                    ],
                                                    alignment: AlignmentType.RIGHT
                                                })
                                            ]
                                        }),
                                    ]
                                })
                            ),
                        ]
                    }),
                    new Paragraph({ text: "" }), // Spacer

                    // Payment info
                    new Paragraph({
                        text: paymentData?.vendor_card2_name && paymentData?.vendor_card2_number
                            ? `${paymentData.vendor_card2_name}    ${paymentData.vendor_card2_number}`
                            : ""
                    }),
                    new Paragraph({
                        text: paymentData?.vendor_card1_name && paymentData?.vendor_card1_number
                            ? `${paymentData.vendor_card1_name}    ${paymentData.vendor_card1_number}`
                            : ""
                    }),
                    new Paragraph({ text: "" }), // Add space before landlord name
                    new Paragraph({
                        text: landlord.name || ""
                    })
                ]
            }
        ]
    });

    const logid = Date.now();
    let baseName = `DynamicDocument_${logid}`;
    let filename = getUniqueFilename(baseName);
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filename, buffer);
    console.log(`Document created! logid: ${logid}, filename: ${filename}`);
}

generateDocx();
