import {
    TextRun,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    WidthType,
    AlignmentType,
    BorderStyle
} from "docx";

import {
    bundleText,
    qtyText,
    dateTextDDMMYYYY,
    priceText,
    totalPriceText
} from '../utils/formatters.js';

export function createYuborilganlarTable(dbRows) {
    const yuborilganlar = dbRows.filter(row => row.is_refund == null);
    const yuborilganlarRows = yuborilganlar.map((row, idx) => [
        String(idx + 1),
        (row.post_title || "") + bundleText(row.is_bundle),
        qtyText(row.product_qty),
        dateTextDDMMYYYY(row.start_date),
        priceText(row.price)
    ]);

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            createHeaderRow(["№", "Mahsulot nomi", "Soni", "Yuborilgan sanasi", "Kunlik narxi"], "a9cce3"),
            ...(yuborilganlarRows.length > 0
                ? yuborilganlarRows.map(row => createDataRow(row))
                : [createEmptyRow(5)])
        ]
    });
}

export function createQaytganlarTable(dbRows) {
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

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            createHeaderRow([
                "№", "Mahsulot nomi", "Soni", "Qaytgan sanasi",
                "Ishlatilgan kuni", "Kunlik narxi", "Umumiy narxi"
            ], "f9e79f"),
            ...(qaytganlarRows.length > 0
                ? qaytganlarRows.map(row => createDataRow(row))
                : [createEmptyRow(7)])
        ]
    });
}

export function createYoqotilganlarTable(dbRows) {
    const yoqotilganlar = dbRows.filter(row => row.is_refund == null && row.is_bundle == null);
    const yoqotilganlarRows = yoqotilganlar.map((row, idx) => [
        String(idx + 1),
        row.post_title || "",
        qtyText(row.lost_qty),
        priceText(row.regular_price),
        totalPriceText(row.lost_qty * row.regular_price)
    ]);

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            createHeaderRow(["№", "Mahsulot nomi", "Soni", "Narxi", "Umumiy narxi"], "f1948a"),
            ...(yoqotilganlarRows.length > 0
                ? yoqotilganlarRows.map(row => createDataRow(row))
                : [createEmptyRow(5)])
        ]
    });
}

export function createTolovlarTable(paymentData, paymentRows) {
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

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
            createHeaderRow(["To'lov miqdori", "To'lov turi", "Sana"], "abebc6"),
            ...(tolovlarRows.length > 0
                ? tolovlarRows.map(row => createDataRow(row))
                : [createEmptyRow(3)])
        ]
    });
}

function createHeaderRow(headers, shadingColor) {
    return new TableRow({
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
                shading: { fill: shadingColor },
                borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "000000" } }
            })
        ),
        tableHeader: true,
    });
}

function createDataRow(rowData) {
    return new TableRow({
        children: rowData.map(val =>
            new TableCell({
                children: [
                    new Paragraph({
                        text: val,
                        alignment: AlignmentType.CENTER
                    })
                ]
            })
        )
    });
}

function createEmptyRow(columnCount) {
    return new TableRow({
        children: [
            new TableCell({
                children: [new Paragraph({ text: "Ma'lumot yo'q", alignment: AlignmentType.CENTER })],
                columnSpan: columnCount
            })
        ]
    });
} 