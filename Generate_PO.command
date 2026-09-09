#!/bin/bash
cd "$(dirname "$0")"
if ! command -v python3 >/dev/null 2>&1; then
  osascript -e 'display dialog "ไม่พบ Python 3 ในเครื่องนี้ กรุณาติดตั้ง Python 3 ก่อนใช้งาน" buttons {"OK"} default button "OK"'
  exit 1
fi
python3 "$(pwd)/generate_po_one_click.py"
rc=$?
if [ $rc -eq 0 ]; then
  osascript -e 'display notification "สร้าง PO เสร็จแล้ว ดูในโฟลเดอร์ OUTPUT" with title "PO Generator"'
  open "$(pwd)/OUTPUT"
else
  osascript -e 'display dialog "สร้าง PO ไม่สำเร็จ กรุณารัน Generate_PO.command จาก Terminal เพื่อดู error" buttons {"OK"} default button "OK"'
fi
exit $rc
