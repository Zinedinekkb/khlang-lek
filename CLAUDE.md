# CLAUDE.md — คลังเหล็ก (ระบบสต็อก/ขาย/การเงินร้านเหล็ก)

> ไฟล์นี้ให้ Claude Code อ่านอัตโนมัติทุกครั้งที่เปิด session ใน repo นี้
> ข้อความที่เขียนว่า "ตรวจจาก repo" คือยังไม่ยืนยัน ให้เช็กกับโค้ดจริงก่อนเชื่อ แล้วแก้ไฟล์นี้ให้ตรง

## โปรเจ็คนี้คืออะไร
เว็บแอปจัดการสต็อก ซื้อเข้า ขาย ออกเอกสาร ลูกหนี้ และการเงินของร้านค้าเหล็ก UI ภาษาไทย ใช้บนมือถือเป็นหลัก
พอร์ตมาจากอาร์ติแฟกต์ใน claude.ai (ไฟล์เดียว ~2,750 บรรทัด) ต้นฉบับอยู่ที่ `reference/original.html` **ห้ามแก้ไฟล์นั้น**
แผนงานเต็มอยู่ที่ `docs/PLAN.md` (เฟส 1–10) — **ยังไม่มีไฟล์นี้ใน repo** ให้ใช้เป็นเกณฑ์ตรวจว่างานครบหรือยัง

## Stack
- Vite + vanilla JS (ยังไม่ใช่ React ตั้งใจไม่เขียนใหม่)
- Firebase: Firestore, Auth, Storage — ใช้ **compat SDK** (`firebase/compat/*` ใน `src/lib/firebase.js`)
- Vercel: deploy อัตโนมัติจาก `main`, มี serverless function `api/read-doc.js` เรียก Anthropic API อ่านบิล — **ยังไม่มีใน repo** (ไม่มีโฟลเดอร์ `api/`, capability `sample` ยังไม่ถูกต่อ)
- ไลบรารี: `xlsx` (SheetJS), `qrcode-generator`

## คำสั่งที่ใช้บ่อย
- `npm install`
- `npm run dev` — รันเครื่องตัวเอง
- `npm run build` — ต้องผ่านก่อนเปิด PR ทุกครั้ง
- `npm test` — ยังไม่มี (ต้องเพิ่ม Vitest สำหรับฟังก์ชันคำนวณ)

## กฎการทำงาน (สำคัญ)
1. **ห้าม push เข้า `main` ตรงๆ** Vercel deploy ขึ้นของจริงที่ลูกค้าใช้อยู่ทันที ให้ทำบน branch แยก (`fix/...`, `feat/...`) แล้วเปิด PR รอผมอนุมัติ Vercel จะมี preview URL ให้ทดสอบ
2. **ห้ามใส่ secret ในโค้ดหรือ commit** (`ANTHROPIC_API_KEY`, `FIREBASE_SERVICE_ACCOUNT`, `.env`) ค่าเหล่านี้อยู่ใน Vercel env เท่านั้น ใน client ใช้ได้เฉพาะ `VITE_FIREBASE_*`
3. **แก้ให้น้อยที่สุด** ห้ามเปลี่ยนดีไซน์ ข้อความภาษาไทย หรือตรรกะธุรกิจ โดยไม่ได้ขอ ถ้าเจอสิ่งที่ควรปรับ ให้รายงานก่อน ไม่ต้องแก้เอง
4. **ห้ามลบหรือแก้ข้อมูลใน Firestore ของจริง** ทดสอบกับ emulator หรือโปรเจ็ค Firebase แยก
5. ทุกอย่างที่แตะเงิน/สต็อกต้องมีเทสต์: คำนวณน้ำหนักเหล็ก, ตัดแบ่งเหล็ก (cut planner), ต้นทุน-กำไร, ยอดหนี้และการตัดชำระ (ใบเก่าสุดก่อน), เลขที่เอกสาร (`nextNo`)
6. commit message สั้น ชัดเจน หนึ่งเรื่องต่อหนึ่ง commit

## สถาปัตยกรรมที่ต้องรู้
- **ชั้น storage**: `store` มี `sub/set/update/del` path รูปแบบ `collection/id`; มี `localStore` (localStorage, ไว้ dev) และ `dbStore` (Firestore)
- **Collections**: `items, moves, buys, docs, cash, settings, closes, counts, prices, orders, customers` (+ `users`, และ `itemCosts` ถ้าแยกต้นทุนแล้ว)
- **Shim**: โค้ดเดิมเรียก `claude.use("db" | "user" | "assets" | "downloads" | "sample")` ซึ่งเดิมมาจาก claude.ai ตอนนี้ถูกแทนด้วย `window.claude` shim ที่ต่อ Firebase/Vercel (ดู `src/lib/claude-shim.js`)
- **สิทธิ์**: เจ้าของเห็นต้นทุน-กำไร ลูกน้องไม่เห็น ต้องบังคับที่ `firestore.rules` ด้วย ไม่ใช่แค่ซ่อนใน UI (CSS `.owner`, `isStaff()`)
- **ข้อมูลร้าน** (ชื่อ ที่อยู่ เบอร์ เลขผู้เสียภาษี) อยู่ใน `settings/shop` ห้ามฝังค่าของลูกค้ารายใดไว้ในโค้ดหรือในพรอมต์อ่านบิล

## งานแรกของ session แรก (รีวิวสถานะ)
1. อ่าน `docs/PLAN.md` แล้วเทียบกับโค้ดจริงทีละเฟส ทำตารางสรุป ✅ ทำแล้ว / ⚠️ ทำไม่ครบ / ❌ ยังไม่ทำ พร้อมไฟล์และบรรทัดที่อ้างอิง
2. รัน `npm install && npm run build` รายงาน error/warning
3. ตรวจ secret หลุด: `grep -rn "sk-ant\|private_key\|service_account" --exclude-dir=node_modules --exclude-dir=.git .` และตรวจ `dist/`
4. ตรวจ `firestore.rules` / `storage.rules` ว่าลูกน้องอ่าน collection ต้นทุน/กำไรไม่ได้ และคนไม่ล็อกอินอ่านอะไรไม่ได้
5. ตรวจว่า `DEFAULT_SHOP` (ชื่อร้าน/เบอร์ลูกค้ารายแรก) ถูกเอาออกจากโค้ดและพรอมต์แล้ว
6. เสนอรายการเทสต์ที่ควรเพิ่มตามลำดับความเสี่ยง แล้ว**รอผมเลือก**ก่อนเริ่มเขียน

## นิยามว่า "เสร็จ"
`npm run build` ผ่าน, เทสต์ผ่าน, ทดสอบบน Vercel preview แล้ว, ไม่มี secret ใน bundle, PR มีคำอธิบายว่าแก้อะไรและทดสอบอย่างไร
