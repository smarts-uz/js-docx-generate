import * as mysql from "mysql2/promise";

export const dbConfig = {
    host: "45.150.24.108",
    port: 23306,
    user: "lesa_prod",
    password: "lesa_prod$34133$$",
    database: "lesa_test",
};

export async function createConnection() {
    return await mysql.createConnection(dbConfig);
}

export async function fetchOrderData(connection, order_id) {
    const [rows] = await connection.execute(`Call get__refunded_orders_data(${order_id})`);
    return Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0] : rows;
}

export async function fetchOrderPaymentData(connection, order_id) {
    const [rows] = await connection.execute(`Call get__refunded_order_payment_data(${order_id})`);
    return Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0][0] : (rows[0] || {});
}

export async function fetchOrderPayments(connection, order_id) {
    const [rows] = await connection.execute(
        `SELECT payment_amount, payment_type, date FROM app_order_payment WHERE order_id = ?`, 
        [order_id]
    );
    return rows;
} 