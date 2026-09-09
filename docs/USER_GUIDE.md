# คู่มือผู้ใช้ PO Generator

## สำหรับฝ่ายจัดซื้อ

### สิ่งที่ต้องเตรียม

- ไฟล์แผนสั่งผลิต Excel ของเดือนนั้น
- ไฟล์ PO Template ล่าสุด

### วิธีทำงาน

1. เปิดหน้า PO Generator
2. เลือกไฟล์แผน
3. เลือก PO Template
4. เลือกเดือน เช่น M9 / M10 / M11
5. เลือก AiR / Freedom / Kids ตามที่ต้องการ
6. ตรวจจำนวน PO, จำนวนสินค้า และ Amount
7. Download PO Excel
8. หากต้องการ PDF ให้ใช้ Print / Save PDF

### การตรวจสอบก่อนส่ง Supplier

ตรวจอย่างน้อย 5 จุด:

1. Supplier ถูกต้อง
2. จำนวน Order/pc ตรงกับแผน
3. Order/pack ถูกต้อง
4. Unit price ถูกต้อง
5. Deposit / Payment Term / Shipment Date ถูกต้อง

ระบบไม่ควรแก้ Commercial Terms โดยอัตโนมัติ เพราะเงื่อนไขใน Template อาจแตกต่างจากแผน

## สำหรับผู้ดูแลระบบ

โครงสร้าง mapping หลักอยู่ใน `web/index.html` และ `generate_po_one_click.py` เมื่อมีการเปลี่ยนชื่อ sheet หรือโครงสร้างแผน ให้แก้ mapping และทดสอบด้วยไฟล์จริงก่อนเผยแพร่รุ่นใหม่
