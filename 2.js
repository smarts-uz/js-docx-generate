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
    BorderStyle,
    VerticalMergeType,
    VerticalAlign
} from "docx";
import { log } from "console";

async function generateSingleDocumentWithAllOrders() {
    try {
        const connection = await mysql.createConnection({
            host: "45.150.24.108",
            port: 23306,
            user: "lesa_prod",
            password: "lesa_prod$34133$$",
            database: "lesa_test",
        });

        const order_id = 400;

        const [rows] = await connection.execute(`Call get__refunded_orders_data(${order_id})`);

        const [rows2] = await connection.execute(`Call get__refunded_order_payment_data(${order_id})`);


        const [paymentRows] = await connection.execute(`SELECT payment_amount, payment_type, date FROM app_order_payment WHERE order_id = ?`, [order_id]);
        const paymentData = Array.isArray(rows2) && Array.isArray(rows2[0]) ? rows2[0][0] : (rows2[0] || {});

        function bundleText(is_bundle) {
            return is_bundle == 1 ? " (to'plam)" : "";
        }
        function qtyText(qty) {
            return qty !== null && qty !== undefined ? `${qty} ta` : "";
        }
        function priceText(price) {
            if (price === null || price === undefined) return "0 so'm";
            // Minglik ajratuvchi va "so'm" qo'shish
            return Number(price).toLocaleString("ru-RU") + " so'm";
        }
        function totalPriceText(total) {
            if (total === null || total === undefined) return "0 so'm";
            return Number(total).toLocaleString("ru-RU") + " so'm";
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
        function dateTextDDMMYYYYHHMM(date) {
            if (!date) return "";
            const d = new Date(date);
            if (isNaN(d)) return "";
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = d.getFullYear();
            const hours = String(d.getHours()).padStart(2, '0');
            const minutes = String(d.getMinutes()).padStart(2, '0');
            return `${day}-${month}-${year} ${hours}:${minutes}`;
        }

        const dbRows = Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0] : rows;

        const yuborilganlar = dbRows.filter(row => row.parent_order_id == null && row.is_bundle != 1);

        function getParentProductPrice(row, allRows) {
            if (row.parent_product_id) {
                const parent = allRows.find(r => r.order_item_id == row.parent_item_id);
                if (parent && parent.price !== undefined && parent.price !== null) {
                    return parent.price;
                }
            }
            return row.price;
        }
        function getParentBundleQty(row, allRows) {
            if (row.parent_product_id) {
                // Bir nechta parentlarni qty sini yig'indisini hisoblash
                const parents = allRows.filter(r => r.order_item_id == row.parent_item_id );
                if (parents.length > 0) {
                    return parents.reduce((sum, parent) => {
                        if (parent.product_qty !== undefined && parent.product_qty !== null) {
                            return sum + Number(parent.product_qty);
                        }
                        return sum;
                    }, 0);
                }
            }else if (row.parent_item_id==null) {
                return row.product_qty;
            }else{
                return 0;
            }
        }

        /**
         * Jadvaldagi qatorlarni bir nechta ustunlar bo'yicha birlashtirish uchun funksiya.
         * rows - birlashtiriladigan qatorlar (array of arrays)
         * fieldIndexes - qaysi ustun(lar) bo'yicha birlashtirish kerak (index yoki indexlar massivi)
         */
        function mergeRowsByFields(rows, fieldIndexes) {
            // fieldIndexes massiv emas bo'lsa, uni massivga aylantiramiz
            if (!Array.isArray(fieldIndexes)) fieldIndexes = [fieldIndexes];

            let prevValues = [];    // Oldingi qatorning birlashtiriladigan ustun(lar) qiymatlari
            let mergeStart = -1;    // Birlashtirish boshlanadigan qator indeksi
            let mergeCount = 0;     // Nechta qator birlashtirilayotganini hisoblash

            // Barcha qatorlarni aylanib chiqamiz
            for (let i = 0; i < rows.length; i++) {
                // Hozirgi qatorning birlashtiriladigan ustun(lar) qiymatlari
                const vals = fieldIndexes.map(idx => rows[i][idx]);

                // Agar oldingi qiymatlar mavjud bo'lsa va hozirgi qiymatlar hammasi bir xil bo'lsa
                if (
                    prevValues.length > 0 &&
                    vals.every((val, idx) => val === prevValues[idx])
                ) {
                    // Birlashtirish davom etadi
                    mergeCount++;
                } else {
                    // Agar oldingi birlashtirish bo'lgan bo'lsa, uni yakunlaymiz
                    if (mergeCount > 0) {
                        for (let j = mergeStart; j < mergeStart + mergeCount + 1; j++) {
                            fieldIndexes.forEach(idx => {
                                if (j === mergeStart) {
                                    // Birinchi qator - merge boshlanishi
                                    rows[j][idx] = { value: rows[j][idx], merge: "restart" };
                                } else {
                                    // Qolgan qatorlar - merge davom etadi
                                    rows[j][idx] = { value: "", merge: "continue" };
                                }
                            });
                        }
                    }
                    // Yangi qiymatlar uchun merge boshlaymiz
                    prevValues = vals;
                    mergeStart = i;
                    mergeCount = 0;
                }
            }
            // Oxirgi birlashtirishni ham tekshirib, yakunlaymiz
            if (mergeCount > 0) {
                for (let j = mergeStart; j < mergeStart + mergeCount + 1; j++) {
                    fieldIndexes.forEach(idx => {
                        if (j === mergeStart) {
                            rows[j][idx] = { value: rows[j][idx], merge: "restart" };
                        } else {
                            rows[j][idx] = { value: "", merge: "continue" };
                        }
                    });
                }
            }
            // Natijani konsolga chiqaramiz (debug uchun)
            console.log(rows);
            // Birlashtirilgan qatorlarni qaytaramiz
            return rows;
        }

        
        let yuborilganlarRows = yuborilganlar.map((row, idx) => {
            // parent_product_id bo'lsa, narxni parentdan olamiz
            const priceValue = getParentProductPrice(row, dbRows);
            return [
                String(idx + 1),
                (row.post_title || "") + bundleText(row.is_bundle),
                qtyText(row.product_qty),
                (row.parent_product_title == null || getParentBundleQty(row, dbRows) == null || getParentBundleQty(row, dbRows) == 0)
                    ? ""
                    : (row.parent_product_title + '\n' + qtyText(getParentBundleQty(row, dbRows))),
                // start_date bo'sh bo'lsa, "" chiqaramiz, aks holda formatlaymiz
                row.start_date ? dateTextDDMMYYYYHHMM(row.start_date) : "salom",
                priceText(priceValue),
                priceText(priceValue * getParentBundleQty(row, dbRows))
            ];
        });

        yuborilganlarRows = mergeRowsByFields(yuborilganlarRows, [3, 4, 5, 6]);

        // Merge qilib bo'lgandan keyin 5-ustunni (index 4) faqat dd-mm-yyyy formatga o'zgartiramiz
        yuborilganlarRows = yuborilganlarRows.map(row => {
            let newRow = [...row];
            if (newRow[4]) {
                // Agar qiymat object bo'lsa (merge uchun), value ni o'zgartiramiz
                if (typeof newRow[4] === "object" && newRow[4] !== null && "value" in newRow[4]) {
                    // value ichida sana va vaqt bo'lsa, faqat sanani ajratib olamiz
                    let value = newRow[4].value;
                    if (typeof value === "string" && value.match(/^\d{2}-\d{2}-\d{4} \d{2}:\d{2}$/)) {
                        // Masalan: 26-07-2025 07:29 => 26-07-2025
                        value = value.split(" ")[0];
                        newRow[4] = { ...newRow[4], value };
                    } else if (value) {
                        newRow[4] = { ...newRow[4], value: dateTextDDMMYYYY(value) };
                    } else {
                        newRow[4] = { ...newRow[4], value: "salom" };
                    }
                } else if (typeof newRow[4] === "string" && newRow[4].match(/^\d{2}-\d{2}-\d{4} \d{2}:\d{2}$/)) {
                    // Masalan: 26-07-2025 07:29 => 26-07-2025
                    newRow[4] = newRow[4].split(" ")[0];
                } else {
                    newRow[4] = newRow[4] ? dateTextDDMMYYYY(newRow[4]) : "salom";
                }
            }
            return newRow;
        });

        const yuborilganlarTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        ...["№", "Mahsulot nomi", "Soni","To'plam", "Yuborilgan sanasi", "Bir dona to'plam kunlik narxi","Umumiy bir kunlik narxi"].map(header =>
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
                            children: row.map((val, colIdx) => {
                                let cellProps = {};
                                let text = val;
                                if (
                                    (colIdx === 3 || colIdx === 4 || colIdx === 5 || colIdx === 6) &&
                                    typeof val === "object" &&
                                    val !== null &&
                                    val.merge
                                ) {
                                    cellProps.verticalMerge = val.merge === "restart"
                                        ? VerticalMergeType.RESTART
                                        : VerticalMergeType.CONTINUE;
                                    cellProps.verticalAlign = VerticalAlign.CENTER;
                                    text = val.value;
                                }
                                return new TableCell({
                                    ...cellProps,
                                    children: [
                                        new Paragraph({
                                            text: text,
                                            alignment: AlignmentType.CENTER,
                                            spacing: { before: 0, after: 0 },
                                        })
                                    ]
                                });
                            })
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
        });

        const qaytganlar = dbRows.filter(row => row.is_refund == 1);

        let qaytganlarRows = qaytganlar.map((row, idx) => {
            return [
                String(idx + 1),
                (row.post_title || "") + bundleText(row.is_bundle),
                qtyText(Math.abs(row.product_qty)),
                (row.parent_product_title == null || getParentBundleQty(row, dbRows) == null || getParentBundleQty(row, dbRows) == 0)
                    ? ""
                    : (row.parent_product_title + '\n' + qtyText(getParentBundleQty(row, dbRows))),
                dateTextDDMMYYYY(row.end_date),
                row.used_days !== null && row.used_days !== undefined ? `${row.used_days} kun` : "",
                priceText(Math.abs(row.price)),
                totalPriceText(Math.abs(row.price) * row.used_days * Math.abs(row.product_qty))
            ];
        });

        // 3-ustun (To'plam), 4-ustun (Qaytgan sanasi) va 6-ustun (Kunlik narxi) uchun merge bo'lsin
        qaytganlarRows = mergeRowsByFields(qaytganlarRows, [3, 4, 6]);

        const qaytganlarTable = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: [
                new TableRow({
                    children: [
                        ...["№", "Mahsulot nomi", "Soni","To'plam", "Qaytgan sanasi", "Ishlatilgan kuni", "Kunlik narxi", "Umumiy narxi"].map(header =>
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
                            children: row.map((val, colIdx) => {
                                let cellProps = {};
                                let text = val;
                                
                                // 3-ustun (To'plam), 4-ustun (Qaytgan sanasi) va 6-ustun (Kunlik narxi) uchun merge
                                if (
                                    (colIdx === 3 || colIdx === 4 || colIdx === 6) &&
                                    typeof val === "object" &&
                                    val !== null &&
                                    val.merge
                                ) {
                                    cellProps.verticalMerge = val.merge === "restart"
                                        ? VerticalMergeType.RESTART
                                        : VerticalMergeType.CONTINUE;
                                    cellProps.verticalAlign = VerticalAlign.CENTER;
                                    text = val.value;
                                }
                                return new TableCell({
                                    ...cellProps,
                                    children: [
                                        new Paragraph({
                                            text: text,
                                            alignment: AlignmentType.CENTER,
                                            spacing: { before: 0, after: 0 },
                                        })
                                    ]
                                });
                            })
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
            if (phone === null || phone === undefined) return " ";
            // Barcha raqam bo'lmagan belgilarni olib tashlash
            let digits = phone.replace(/\D/g, "");
            // Agar 9 ta raqam bo'lsa, +998 qo'shamiz
            if (digits.length === 9) {
                digits = "998" + digits;
            } else if (digits.length === 12 && digits.startsWith("998")) {
                // To'g'ri formatda
            } else if (digits.length === 13 && digits.startsWith("8")) {
                // Ba'zan 8 bilan boshlanadi, uni olib tashlab 998 qo'shamiz
                digits = "998" + digits.slice(1);
            }
            if (digits.length !== 12) return " "; // noto'g'ri bo'lsa bosh joy qaytariladi
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