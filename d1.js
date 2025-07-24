import {
    createConnection,
    fetchOrderData,
    fetchOrderPaymentData,
    fetchOrderPayments
} from './src/config/database.js';
import { generateDocument } from './src/services/documentService.js';

async function generateSingleDocumentWithAllOrders() {
    try {
        const connection = await createConnection();
        const order_id = 372;

        const dbRows = await fetchOrderData(connection, order_id);
        const paymentData = await fetchOrderPaymentData(connection, order_id);
        const paymentRows = await fetchOrderPayments(connection, order_id);

        await generateDocument(dbRows, paymentData, paymentRows);
        await connection.end();
    } catch (err) {
        console.error("❌ Error generating document:", err);
    }
}

generateSingleDocumentWithAllOrders();