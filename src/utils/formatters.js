export function bundleText(is_bundle) {
    return is_bundle == 1 ? " (to'plam)" : "";
}

export function qtyText(qty) {
    return qty !== null && qty !== undefined ? `${qty} ta` : "";
}

export function priceText(price) {
    if (price === null || price === undefined) return "0 so'm";
    return Number(price).toLocaleString("ru-RU") + " so'm";
}

export function totalPriceText(total) {
    if (total === null || total === undefined) return "0 so'm";
    return Number(total).toLocaleString("ru-RU") + " so'm";
}

export function dateText(date) {
    if (!date) return "";
    const d = new Date(date);
    if (isNaN(d)) return String(date);
    return d.toISOString().slice(0, 10);
}

export function dateTextDDMMYYYY(date) {
    if (!date) return "";
    const d = new Date(date);
    if (isNaN(d)) return "";
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}-${month}-${year}`;
}

export function safeValue(val, fallback = "") {
    return val !== undefined && val !== null ? val : fallback;
}

export function safeNumber(val) {
    return val !== undefined && val !== null ? Number(val).toLocaleString("uz-UZ") : "0";
}

export function formatPhoneNumber(phone) {
    if (!phone) return "";
    let digits = phone.replace(/\D/g, "");
    if (digits.length === 9) {
        digits = "998" + digits;
    } else if (digits.length === 12 && digits.startsWith("998")) {
        // already correct
    } else if (digits.length === 13 && digits.startsWith("8")) {
        digits = "998" + digits.slice(1);
    }
    if (digits.length !== 12) return phone;
    return `+${digits.slice(0,3)} ${digits.slice(3,5)} ${digits.slice(5,8)} ${digits.slice(8,10)} ${digits.slice(10,12)}`;
}

export function formatCardNumber(card) {
    if (!card) return "";
    let digits = card.replace(/\D/g, "");
    if (digits.length < 12) return card;
    return digits.replace(/(.{4})/g, "$1 ").trim();
} 