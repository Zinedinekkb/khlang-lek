import { describe, it, expect } from "vitest";
import { buyLinePpk, itemCostPerKg } from "../src/lib/cost.js";

describe("buyLinePpk", () => {
  it("ราคาคิดต่อ กก. จากหน้ารีวิวบิล", () => {
    expect(buyLinePpk({ priceBasis: "kg", unitPrice: 28.5, price: 300, qty: 10 })).toBe(28.5);
  });
  it("ซื้อเป็นหน่วย กก.", () => {
    expect(buyLinePpk({ unit: "กก.", price: 27, qty: 500 })).toBe(27);
  });
  it("ซื้อเป็นเส้น มีน้ำหนักรวม", () => {
    // 10 เส้น × 300 = 3000 บาท / 120 กก. = 25 บาท/กก.
    expect(buyLinePpk({ unit: "เส้น", price: 300, qty: 10, kgTotal: 120 })).toBe(25);
  });
  it("ซื้อเป็นเส้น มีแค่ กก./เส้น", () => {
    expect(buyLinePpk({ unit: "เส้น", price: 300, qty: 10, kgPer: 12 })).toBe(25);
  });
  it("ไม่มีข้อมูลน้ำหนัก = 0", () => {
    expect(buyLinePpk({ unit: "เส้น", price: 300, qty: 10 })).toBe(0);
  });
});

describe("itemCostPerKg", () => {
  const item = { id: "a", cost: 330, kg: 12 };
  it("ใช้บิลซื้อล่าสุดก่อน (buys เรียงใหม่สุดก่อน)", () => {
    const buys = [
      { lines: [{ itemId: "b", unit: "กก.", price: 99 }] },
      { lines: [{ itemId: "a", unit: "กก.", price: 26 }] },
      { lines: [{ itemId: "a", unit: "กก.", price: 24 }] },
    ];
    expect(itemCostPerKg(item, buys)).toBe(26);
  });
  it("ไม่มีบิล ใช้ ทุน/หน่วย ÷ กก./หน่วย", () => {
    expect(itemCostPerKg(item, [])).toBe(27.5);
  });
  it("ข้ามบรรทัดบิลที่คำนวณต่อ กก. ไม่ได้", () => {
    const buys = [{ lines: [{ itemId: "a", unit: "เส้น", price: 300, qty: 1 }] }];
    expect(itemCostPerKg(item, buys)).toBe(27.5);
  });
  it("ไม่มีน้ำหนักหรือไม่มีทุน = 0", () => {
    expect(itemCostPerKg({ id: "x", cost: 100, kg: 0 }, [])).toBe(0);
    expect(itemCostPerKg(null, [])).toBe(0);
  });
});
