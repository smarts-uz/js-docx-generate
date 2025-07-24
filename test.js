import * as fs from "fs";
import {
    Document,
    Packer,
    TextRun,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    WidthType,
    AlignmentType,
    BorderStyle,
    VerticalMergeType
} from "docx";

// Statik userlar ro'yxati
const users = [
    { ism: "Ali", familiya: "Valiyev" },
    { ism: "Gulnoza", familiya: "Valiyev" },
    { ism: "Gulnoza", familiya: "qozi" },
    { ism: "Jasur", familiya: "qozi" }
];

// Jadval uchun sarlavhalar
const headers = ["№", "Ism", "Familiya"];

// Familiyasi yoki ismi bir xil bo'lganlarni birlashtirish uchun yordamchi funksiya
function mergeRowsByField(rows, fieldIndex) {
    let merged = [];
    let prevValue = null;
    let mergeStart = -1;
    let mergeCount = 0;

    for (let i = 0; i < rows.length; i++) {
        const val = rows[i][fieldIndex];
        if (val === prevValue) {
            mergeCount++;
        } else {
            if (mergeCount > 0) {
                // oldingi merge uchun belgila
                for (let j = mergeStart; j < mergeStart + mergeCount + 1; j++) {
                    if (j === mergeStart) {
                        rows[j][fieldIndex] = { value: rows[j][fieldIndex], merge: "restart" };
                    } else {
                        rows[j][fieldIndex] = { value: "", merge: "continue" };
                    }
                }
            }
            prevValue = val;
            mergeStart = i;
            mergeCount = 0;
        }
    }
    // Oxirgi merge uchun
    if (mergeCount > 0) {
        for (let j = mergeStart; j < mergeStart + mergeCount + 1; j++) {
            if (j === mergeStart) {
                rows[j][fieldIndex] = { value: rows[j][fieldIndex], merge: "restart" };
            } else {
                rows[j][fieldIndex] = { value: "", merge: "continue" };
            }
        }
    }
    return rows;
}

// Userlar ma'lumotlarini jadvalga tayyorlash
let userRows = users.map((user, idx) => [
    String(idx + 1),
    user.ism,
    user.familiya
]);

// Avval familiya bo'yicha tartiblash va birlashtirish
userRows.sort((a, b) => a[2].localeCompare(b[2]) || a[1].localeCompare(b[1]));
userRows = mergeRowsByField(userRows, 2); // familiya (index 2) bo'yicha merge

// Endi ism bo'yicha tartiblash va birlashtirish (faqat familiya merge bo'lmagan joylarda)
userRows = mergeRowsByField(userRows, 1); // ism (index 1) bo'yicha merge

const userTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    alignment: AlignmentType.CENTER, // Jadvalni sahifa markaziga joylash
    rows: [
        new TableRow({
            children: headers.map(header =>
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
            ),
            tableHeader: true,
        }),
        ...(userRows.length > 0
            ? userRows.map(row =>
                new TableRow({
                    children: row.map((val, colIdx) => {
                        let cellProps = {};
                        let text = val;
                        if (typeof val === "object" && val !== null && val.merge) {
                            cellProps.verticalMerge = val.merge === "restart"
                                ? VerticalMergeType.RESTART
                                : VerticalMergeType.CONTINUE;
                            text = val.value;
                        }
                        return new TableCell({
                            ...cellProps,
                            children: [
                                new Paragraph({
                                    text: text,
                                    alignment: AlignmentType.CENTER
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
                            columnSpan: 3
                        })
                    ]
                })
            ]
        )
    ]
});

// DOCX hujjat yaratish va jadvalni joylash
const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "Foydalanuvchilar ro'yxati", bold: true, size: 28, font: "Times New Roman" })],
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 300 }
                }),
                userTable
            ]
        }
    ]
});

// Jadvalni faylga yozish
Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("test.docx", buffer);
    console.log("✅ Hujjat yaratildi: test.docx");
});