# สร้าง template-data.js (base64 ของเทมเพลต .docx) — ต้องรันทุกครั้งที่แก้ไฟล์ในโฟลเดอร์ templates/
#   python tools/gen-template-data.py
# เหตุที่ต้องแปลงเป็น base64: หน้าเว็บเปิดจาก file:// ในเครื่องได้ด้วย (fetch ไฟล์ .docx ตรงๆ ไม่ได้)
import base64, json, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL = os.path.join(ROOT, "templates", "บันทึกออกตรวจงานจ้าง_template.docx")
OUT = os.path.join(ROOT, "template-data.js")

with open(TPL, "rb") as f:
    b64 = base64.b64encode(f.read()).decode("ascii")
with open(OUT, "w", encoding="utf-8", newline="\n") as f:
    f.write("// สร้างอัตโนมัติด้วย tools/gen-template-data.py — ห้ามแก้มือ\n")
    f.write("var TEMPLATE_INSPECT_BASE64 = %s;\n" % json.dumps(b64))
print("ok  %s -> template-data.js (%d bytes base64)" % (os.path.basename(TPL), len(b64)))
