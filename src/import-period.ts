export function billingMonthFromText(value: string) {
  const match = String(value || "").match(
    /(?:חודש\s*החיוב|חיוב\s*בתאריך)\s*:?[\s\u200e\u200f]*(\d{1,2})[./-](\d{1,2})[./-](\d{4})/,
  );
  return match
    ? `${match[3]}${String(Number(match[2])).padStart(2, "0")}`
    : "";
}
