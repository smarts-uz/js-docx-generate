import * as fs from "fs";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, ShadingType } from "docx";
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

    const [rows] = await connection.execute("SELECT * FROM wp_users");
    console.log(rows);
    await connection.end();
    return rows;
}

// 2. Generate docx from data
async function generateDocx() {
    const data = await fetchData();

    // Generate a logid (timestamp-based for uniqueness)
    const logid = Date.now();
    let baseName = `DynamicDocument_${logid}`;
    let filename = getUniqueFilename(baseName);

    const doc = new Document({
        sections: [
            {
                properties: {},
                children: [
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            // Header row with style
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [new TextRun({ text: "ID", bold: true, color: "FFFFFF" })],
                                                alignment: AlignmentType.CENTER,
                                            })
                                        ],
                                        shading: { type: ShadingType.CLEAR, color: "auto", fill: "4472C4" },
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [new TextRun({ text: "Display Name", bold: true, color: "FFFFFF" })],
                                                alignment: AlignmentType.CENTER,
                                            })
                                        ],
                                        shading: { type: ShadingType.CLEAR, color: "auto", fill: "4472C4" },
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [new TextRun({ text: "Phone", bold: true, color: "FFFFFF" })],
                                                alignment: AlignmentType.CENTER,
                                            })
                                        ],
                                        shading: { type: ShadingType.CLEAR, color: "auto", fill: "4472C4" },
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                    }),
                                ],
                                tableHeader: true,
                            }),
                            // Data rows with style
                            ...data.map(row => new TableRow({
                                children: [
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [new TextRun({ text: String(row.ID || row.id || "") })],
                                                alignment: AlignmentType.CENTER,
                                            })
                                        ],
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [new TextRun({ text: String(row.display_name || "") })],
                                                alignment: AlignmentType.CENTER,
                                            })
                                        ],
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                    }),
                                    new TableCell({
                                        children: [
                                            new Paragraph({
                                                children: [new TextRun({ text: String(row.phone || "") })],
                                                alignment: AlignmentType.CENTER,
                                            })
                                        ],
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                    }),
                                ],
                            })),
                        ],
                        borders: {
                            top: { style: "single", size: 1, color: "000000" },
                            bottom: { style: "single", size: 1, color: "000000" },
                            left: { style: "single", size: 1, color: "000000" },
                            right: { style: "single", size: 1, color: "000000" },
                            insideHorizontal: { style: "single", size: 1, color: "000000" },
                            insideVertical: { style: "single", size: 1, color: "000000" },
                        },
                    })
                ],
            },
        ],
    });

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filename, buffer);
    console.log(`Document created! logid: ${logid}, filename: ${filename}`);
}

generateDocx();
