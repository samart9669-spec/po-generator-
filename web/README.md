# PO Generator Web App

สำหรับผู้ใช้ทั่วไปที่ไม่ต้องติดตั้ง Python หรือรู้ GitHub

## วิธีใช้

1. เปิด `web/index.html` ผ่าน GitHub Pages หรือเว็บเซิร์ฟเวอร์ภายในบริษัท
2. เลือกไฟล์ **แผนสั่งผลิต .xlsx**
3. เลือกไฟล์ **PO Template .xlsx**
4. เลือกเดือนและรุ่นที่ต้องการ
5. กด **Generate PO**
6. ตรวจยอดสรุป
7. กด **Download PO Excel** หรือ **Print / Save PDF**

ไฟล์จะถูกประมวลผลใน browser ของผู้ใช้ ไม่ส่ง Excel ขึ้น server

## หมายเหตุ

- Web app ใช้ SheetJS จาก CDN
- Template จริงและข้อมูล Supplier ไม่ควร commit เข้า public repository
- ก่อนส่ง PO ให้ตรวจ Commercial Terms เช่น Deposit / Payment Term / Shipment Date
