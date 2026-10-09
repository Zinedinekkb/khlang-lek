// ทุนซื้อต่อ กก. ของสินค้า: ใช้ราคาจากบิลซื้อเข้าล่าสุดก่อน ถ้าไม่มีใช้ ทุน/หน่วย ÷ กก./หน่วย
// buys ต้องเรียงใหม่สุดก่อน (แบบที่ store.sub ส่งมา)
export function buyLinePpk(l) {
  const price = +l.price || 0, qty = +l.qty || 0;
  if (l.priceBasis === "kg" && +l.unitPrice > 0) return +l.unitPrice;
  if (l.unit === "กก." && price > 0) return price;
  if (+l.kgTotal > 0 && price > 0 && qty > 0) return (price * qty) / +l.kgTotal;
  if (+l.kgPer > 0 && price > 0) return price / +l.kgPer;
  return 0;
}

export function itemCostPerKg(item, buys = []) {
  if (!item) return 0;
  for (const b of buys) {
    for (const l of b.lines || []) {
      if (item.id && l.itemId === item.id) {
        const p = buyLinePpk(l);
        if (p > 0) return Math.round(p * 100) / 100;
      }
    }
  }
  const cost = +item.cost || 0, kg = +item.kg || 0;
  return cost > 0 && kg > 0 ? Math.round((cost / kg) * 100) / 100 : 0;
}
