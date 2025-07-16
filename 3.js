import * as fs from "fs";
import * as mysql from "mysql2/promise";
import {
    patchDocument,
    PatchType,
    TextRun,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    WidthType,
    AlignmentType,
    BorderStyle
} from "docx";
import { type } from "os";

// NOTE: The data from file_context_0 uses keys like post_title, product_qty, start_date, price, is_bundle, is_refund, etc.
//       There is no product_name, qty, created_at, price_per_day in the data. 
//       So, we must map to the correct keys.

async function generateSingleDocumentWithAllOrders() {
    try {
        const connection = await mysql.createConnection({
            host: "45.150.24.108",
            port: 23306,
            user: "lesa_prod",
            password: "lesa_prod$34133$$",
            database: "lesaapp",
        });

        const order_id = 372;

        const [rows] = await connection.execute(`Call get__refunded_orders_data(${order_id})`);

        // Uncomment for debugging:
        // console.dir(rows, { depth: null });

        const [rows2] = await connection.execute(`Call get__refunded_order_payment_data(${order_id})`);


        // Fetch all payments for this order
        const [paymentRows] = await connection.execute(`SELECT payment_amount, payment_type, date FROM app_order_payment WHERE order_id = ?`, [order_id]);
        // rows2[0] is the actual data for stored procedures in mysql2
        const paymentData = Array.isArray(rows2) && Array.isArray(rows2[0]) ? rows2[0][0] : (rows2[0] || {});

        // Helper to format values
        function bundleText(is_bundle) {
            return is_bundle == 1 ? " (to'plam)" : "";
        }
        function qtyText(qty) {
            return qty !== null && qty !== undefined ? `${qty} ta` : "";
        }
        function priceText(price) {
            if (price === null || price === undefined) return "0 so'm";
            // Format with thousands separator and "so'm"
            return Number(price).toLocaleString("ru-RU") + " so'm";
        }
        function totalPriceText(total) {
            if (total === null || total === undefined) return "0 so'm";
            return Number(total).toLocaleString("ru-RU") + " so'm";
        }
        function dateText(date) {
            if (!date) return "";
            // Format as YYYY-MM-DD
            const d = new Date(date);
            if (isNaN(d)) return String(date);
            return d.toISOString().slice(0, 10);
        }
        function dateTextDDMMYYYY(date) {
            if (!date) return "";
            const d = new Date(date);
            if (isNaN(d)) return "";
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = d.getFullYear();
            return `${day}-${month}-${year}`;
        }

        // rows[0] is the actual data for stored procedures in mysql2
        const dbRows = Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0] : rows;

        // --- Yuborilganlar (order_items_table) ---
        const yuborilganlar = dbRows.filter(row => row.is_refund == null);
        const yuborilganlarRows = yuborilganlar.map((row, idx) => [
            String(idx + 1),
            (row.post_title || "") + bundleText(row.is_bundle),
            qtyText(row.product_qty),
            dateTextDDMMYYYY(row.start_date),
            priceText(row.price)
        ]);
        const yuborilganlarTable = new Table({
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
                                shading: { fill: "a9cce3" },
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
                                    columnSpan: 5
                                })
                            ]
                        })
                    ]
                )
            ]
        });

        // --- Qaytganlar (refund_items_table) ---
        // Use correct keys: post_title, is_bundle, product_qty, end_date, used_days, price
        const qaytganlar = dbRows.filter(row => row.is_refund == 1);
        const qaytganlarRows = qaytganlar.map((row, idx) => [
            String(idx + 1),
            (row.post_title || "") + bundleText(row.is_bundle),
            qtyText(Math.abs(row.product_qty)),
            dateTextDDMMYYYY(row.end_date),
            row.used_days !== null && row.used_days !== undefined ? `${row.used_days} kun` : "",
            priceText(Math.abs(row.price)),
            totalPriceText(Math.abs(row.price) * row.used_days * Math.abs(row.product_qty))
        ]);
        const qaytganlarTable = new Table({
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
                                shading: { fill: "f9e79f" },
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
                                    columnSpan: 7
                                })
                            ]
                        })
                    ]
                )
            ]
        });

        // --- Yo'qotilganlar (lost_items_table) ---
        // is_refund == null && is_bundle == null
        const yoqotilganlar = dbRows.filter(row => row.is_refund == null && row.is_bundle == null);
        const yoqotilganlarRows = yoqotilganlar.map((row, idx) => [
            String(idx + 1),
            row.post_title || "",
            qtyText(row.lost_qty),
            priceText(row.regular_price),
            totalPriceText(row.lost_qty * row.regular_price)
        ]);
        const yoqotilganlarTable = new Table({
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
                                shading: { fill: "f1948a" },
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
        });

        // --- To'lovlar (payments_table) ---
        // First row: prepayment amount, then paymentRows, then jami
        const tolovlarRows = [];
        if (paymentData?.prepayment_amount) {
            tolovlarRows.push([
                priceText(paymentData.prepayment_amount),
                "Oldindan to'lov miqdori",
                paymentData?.prepayment_date ? dateTextDDMMYYYY(paymentData.prepayment_date) : ""
            ]);
        }
        let jami = 0;
        paymentRows.forEach(row => {
            const amount = Number(row.payment_amount) || 0;
            jami += amount;
            tolovlarRows.push([
                priceText(amount),
                row.payment_type || "",
                dateTextDDMMYYYY(row.date)
            ]);
        });
        if (tolovlarRows.length > 0) {
            tolovlarRows.push([
                priceText((Number(paymentData?.prepayment_amount) || 0) + jami),
                "Jami",
                ""
            ]);
        }
        const tolovlarTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        ...["To'lov miqdori", "To'lov turi", "Sana"].map(header =>
                            new TableCell({
                                children: [
                                    new Paragraph({
                                        children: [
                                            new TextRun({ text: header, bold: true, size: 22, font: "Times New Roman" })
                                        ],
                                        alignment: AlignmentType.CENTER,
                                    })
                                ],
                                shading: { fill: "abebc6" },
                                borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "000000" } }
                            })
                        )
                    ],
                    tableHeader: true,
                }),
                ...(tolovlarRows.length > 0
                    ? tolovlarRows.map(row =>
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
                                    columnSpan: 3
                                })
                            ]
                        })
                    ]
                )
            ]
        });

        // rows2data
        // Use destructuring and optional chaining for cleaner extraction
        const data = Array.isArray(rows2) && Array.isArray(rows2[0]) ? rows2[0][0] : rows2[0] || {};

        // Helper function to check for null or undefined
        function safeValue(val, fallback = "") {
            return val !== undefined && val !== null ? val : fallback;
        }

        // Helper function for numeric fields
        function safeNumber(val) {
            return val !== undefined && val !== null ? Number(val).toLocaleString("uz-UZ") : "0";
        }

        // Helper function to format phone numbers (Uzbekistan style: +998 XX XXX XX XX)
        function formatPhoneNumber(phone) {
            if (!phone) return "";
            // Remove all non-digit characters
            let digits = phone.replace(/\D/g, "");
            // If already starts with 998, keep it, else try to add
            if (digits.length === 9) {
                // Assume it's just the local part, add +998
                digits = "998" + digits;
            } else if (digits.length === 12 && digits.startsWith("998")) {
                // already correct
            } else if (digits.length === 13 && digits.startsWith("8")) {
                // Sometimes numbers start with 8, drop it and add 998
                digits = "998" + digits.slice(1);
            }
            if (digits.length !== 12) return phone; // fallback to original if not matching
            return `+${digits.slice(0,3)} ${digits.slice(3,5)} ${digits.slice(5,8)} ${digits.slice(8,10)} ${digits.slice(10,12)}`;
        }

        // Helper function to format card numbers (XXXX XXXX XXXX XXXX)
        function formatCardNumber(card) {
            if (!card) return "";
            let digits = card.replace(/\D/g, "");
            if (digits.length < 12) return card; // fallback if not enough digits
            // Group by 4
            return digits.replace(/(.{4})/g, "$1 ").trim();
        }

        // Patch all of these fields from the data object, checking for nulls
        const patches = {
            order_items_table: {
                type: PatchType.DOCUMENT,
                children: [yuborilganlarTable],
            },
            refund_items_table: {
                type: PatchType.DOCUMENT,
                children: [qaytganlarTable],
            },
            lost_items_table: {
                type: PatchType.DOCUMENT,
                children: [yoqotilganlarTable],
            },
            payments_table: {
                type: PatchType.DOCUMENT,
                children: [tolovlarTable],
            },
            // Numeric and string fields, all patched with null checks
            t_pay_amount: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeNumber(data.total_payment_amount),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            d_price: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeNumber(data.delivery_price),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            lost_debt_price: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeNumber(data.lost_debt_price),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            rental_debt_price: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeNumber(data.rental_debt_price),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            t_debt_price: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeNumber(data.total_debt_price),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            vendor_name: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeValue(data.vendor_display_name),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            vendor_adress: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeValue(data.vendor_addres),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            card1_name: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeValue(data.vendor_card1_name),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            card2_name: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeValue(data.vendor_card2_name),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            card1_number: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: formatCardNumber(data.vendor_card1_number),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            card2_number: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: formatCardNumber(data.vendor_card2_number),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            customer_name: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeValue(data.customer_display_name),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            customer_adress: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: safeValue(data.customer_addres),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            vendor_phone: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: formatPhoneNumber(data.vendor_phone),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            vendor_phone_second: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: formatPhoneNumber(data.vendor_phone_second),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            customer_phone: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: formatPhoneNumber(data.customer_phone),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
            customer_phone_second: {
                type: PatchType.PARAGRAPH,
                children: [
                    new TextRun({
                        text: formatPhoneNumber(data.customer_phone_second),
                        font: "Times New Roman",
                        size: 24
                    })
                ]
            },
        };

        const templateData = fs.readFileSync("./check_example.docx");
        const outputType = "nodebuffer";

        const docBuffer = await patchDocument(templateData, { outputType, patches });

        fs.writeFileSync("All_Orders_Report.docx", docBuffer);
        console.log("✅ Document generated: All_Orders_Report.docx");

        await connection.end();
    } catch (err) {
        console.error("❌ Error generating document:", err);
    }
}

generateSingleDocumentWithAllOrders();