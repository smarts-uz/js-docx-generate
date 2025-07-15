import * as fs from "fs";
import * as mysql from "mysql2/promise";
import {
    patchDocument,
    PatchType,
    TextRun,
    Paragraph,
    Table,
    TableRow,
    TableCell
} from "docx";

async function generateSingleDocumentWithAllOrders() {
    try {
        const connection = await mysql.createConnection({
            host: "45.150.24.108",
            port: 23306,
            user: "lesa_prod",
            password: "lesa_prod$34133$$",
            database: "lesaapp",
        });

        const [rows] = await connection.execute(
            `SELECT id, vendor_id, customer_id FROM wp_wc_orders`
        );

        if (!rows.length) {
            console.error("❌ No data found in wp_wc_orders table.");
            return;
        }

        // Create table header row
        const headerRow = new TableRow({
            children: [
                new TableCell({ children: [new Paragraph("Order ID")] }),
                new TableCell({ children: [new Paragraph("Vendor ID")] }),
                new TableCell({ children: [new Paragraph("Customer ID")] }),
            ]
        });

        // Create rows for each order
        const dataRows = rows.map(row => {
            return new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(String(row.order_id))] }),
                    new TableCell({ children: [new Paragraph(String(row.vendor_id))] }),
                    new TableCell({ children: [new Paragraph(String(row.customer_id))] }),
                ]
            });
        });

        // Combine header and data rows
        const table = new Table({
            rows: [headerRow, ...dataRows],
            width: {
                size: 100,
                type: "pct",
            },
        });

        // Patch the document to insert the table
        const patches = {
            orders_table: {
                type: PatchType.DOCUMENT,
                children: [table],
            }
        };

        const templateData = fs.readFileSync("./simple-template.docx");
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
