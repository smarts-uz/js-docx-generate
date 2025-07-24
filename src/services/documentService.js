import * as fs from "fs";
import { patchDocument, PatchType, TextRun } from "docx";
import {
    createYuborilganlarTable,
    createQaytganlarTable,
    createYoqotilganlarTable,
    createTolovlarTable
} from '../templates/tableTemplates.js';
import {
    safeNumber,
    safeValue,
    formatPhoneNumber,
    formatCardNumber
} from '../utils/formatters.js';

export async function generateDocument(dbRows, paymentData, paymentRows) {
    try {
        const patches = {
            order_items_table: {
                type: PatchType.DOCUMENT,
                children: [createYuborilganlarTable(dbRows)],
            },
            refund_items_table: {
                type: PatchType.DOCUMENT,
                children: [createQaytganlarTable(dbRows)],
            },
            lost_items_table: {
                type: PatchType.DOCUMENT,
                children: [createYoqotilganlarTable(dbRows)],
            },
            payments_table: {
                type: PatchType.DOCUMENT,
                children: [createTolovlarTable(paymentData, paymentRows)],
            },
            ...createTextPatches(paymentData)
        };

        const templateData = fs.readFileSync("./check_example.docx");
        const outputType = "nodebuffer";
        const docBuffer = await patchDocument(templateData, { outputType, patches });

        fs.writeFileSync("All_Orders_Report.docx", docBuffer);
        console.log("✅ Document generated: All_Orders_Report.docx");
        return true;
    } catch (err) {
        console.error("❌ Error generating document:", err);
        return false;
    }
}

function createTextPatches(data) {
    return {
        t_pay_amount: createTextPatch(safeNumber(data.total_payment_amount)),
        d_price: createTextPatch(safeNumber(data.delivery_price)),
        lost_debt_price: createTextPatch(safeNumber(data.lost_debt_price)),
        rental_debt_price: createTextPatch(safeNumber(data.rental_debt_price)),
        t_debt_price: createTextPatch(safeNumber(data.total_debt_price)),
        vendor_name: createTextPatch(safeValue(data.vendor_display_name)),
        vendor_adress: createTextPatch(safeValue(data.vendor_addres)),
        card1_name: createTextPatch(safeValue(data.vendor_card1_name)),
        card2_name: createTextPatch(safeValue(data.vendor_card2_name)),
        card1_number: createTextPatch(formatCardNumber(data.vendor_card1_number)),
        card2_number: createTextPatch(formatCardNumber(data.vendor_card2_number)),
        customer_name: createTextPatch(safeValue(data.customer_display_name)),
        customer_adress: createTextPatch(safeValue(data.customer_addres)),
        vendor_phone: createTextPatch(formatPhoneNumber(data.vendor_phone)),
        vendor_phone_second: createTextPatch(formatPhoneNumber(data.vendor_phone_second)),
        customer_phone: createTextPatch(formatPhoneNumber(data.customer_phone)),
        customer_phone_second: createTextPatch(formatPhoneNumber(data.customer_phone_second)),
    };
}

function createTextPatch(text) {
    return {
        type: PatchType.PARAGRAPH,
        children: [
            new TextRun({
                text: text,
                font: "Times New Roman",
                size: 24
            })
        ]
    };
} 