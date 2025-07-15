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
        function dateText(date) {
            if (!date) return "";
            // Format as YYYY-MM-DD
            const d = new Date(date);
            if (isNaN(d)) return String(date);
            return d.toISOString().slice(0, 10);
        }

        // rows[0] is the actual data for stored procedures in mysql2
        const dbRows = Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0] : rows;

        // Only show items that are not refunded (is_refund == null)
        // file_context_0: is_refund == null means "yuborilganlar"
        // Use correct keys: post_title, is_bundle, product_qty, start_date, price
        const yuborilganlar = dbRows.filter(row => row.is_refund == null);

        // Prepare table rows
        const yuborilganlarRows = yuborilganlar.map((row, idx) => [
            String(idx + 1),
            (row.post_title || "") + bundleText(row.is_bundle),
            qtyText(row.product_qty),
            dateText(row.start_date),
            priceText(row.price)
        ]);

        // Build the table as in file_context_0
        const table = new Table({
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
                                    columnSpan: 5
                                })
                            ]
                        })
                    ]
                )
            ]
        });

        // Patch the document to insert the table
        const patches = {
            order_items_table: {
                type: PatchType.DOCUMENT,
                children: [table],
            }
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