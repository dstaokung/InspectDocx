// ---- รายชื่อเดือนไทย + ปี พ.ศ. ใช้ร่วมกันทุกหน้า (main + เอกสาร 1-5) ----
// หมายเหตุ: เดิมฟังก์ชันนี้แปลงเลขปีเป็น "เลขไทย" (๐-๙) ก่อนใส่ลงเอกสาร แต่เอกสารทุกฉบับใช้ฟอนต์
// "TH SarabunIT๙" ซึ่งจะแปลงตัวเลขอารบิก (0-9) ที่พิมพ์ไปเป็นตัวเลขไทยให้อัตโนมัติตอนแสดงผล/พิมพ์อยู่แล้ว
// เพื่อกันความผิดพลาด (เลขไทยจริงกับเลขอารบิกที่ฟอนต์แปลงให้ ปนกันแล้วแก้ยาก) จึงเปลี่ยนให้เก็บ/ส่งออก
// เป็น "เลขอารบิก" เสมอทั้งระบบ แล้วปล่อยให้ฟอนต์เป็นผู้แปลงเป็นเลขไทยตอนแสดงผลแทน — คง key `toThaiDigits`
// ไว้เผื่อโค้ดอื่นเรียกใช้ชื่อนี้อยู่ แต่เปลี่ยนพฤติกรรมเป็น "คืนค่าเป็นเลขอารบิกเดิม" (ไม่แปลงแล้ว)
window.ConsDocxMonthYear = (function () {
  var MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];

  function toThaiDigits(n) {
    return String(n);
  }

  function currentMonth() {
    return MONTHS[new Date().getMonth()];
  }

  function currentYearThai() {
    var beYear = new Date().getFullYear() + 543;
    return String(beYear);
  }

  // รายการปี พ.ศ. (เลขอารบิก) รอบๆ ปีปัจจุบัน ให้เลือกล่วงหน้า/ย้อนหลังได้
  function yearOptionsThai(spread) {
    spread = spread || 3;
    var centerBE = new Date().getFullYear() + 543;
    var out = [];
    for (var i = -spread; i <= spread; i++) {
      out.push(String(centerBE + i));
    }
    return out;
  }

  // เติม <option> ให้ select เดือน/ปี แล้วตั้งค่าเริ่มต้นเป็นเดือน/ปีปัจจุบันเสมอตอนสร้าง options ใหม่
  // (สำคัญ: <select> ที่เพิ่ง append option จะ auto-select ตัวเลือกแรกให้เองอยู่แล้ว ทำให้ sel.value
  // ไม่ใช่ค่าว่างตั้งแต่แรก เช็ค "if (!sel.value)" แล้วข้ามไปเลยไม่ตั้งเดือน/ปีปัจจุบันให้ตามที่ตั้งใจไว้
  // จึงต้องตั้งค่าตรงๆ เสมอ ส่วนค่าที่เคยบันทึกไว้ก่อนหน้า (autosave) หรือส่งมาจากหน้า main จะมาทับ
  // อีกทีในโค้ดที่เรียกหลังจากนี้อยู่แล้ว ไม่ต้องกลัวว่าจะไปทับค่าที่ผู้ใช้เคยเลือกไว้)
  function populateMonthSelect(sel) {
    if (!sel || sel.dataset.filled === "1") return;
    sel.dataset.filled = "1";
    MONTHS.forEach(function (m) {
      var opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      sel.appendChild(opt);
    });
    sel.value = currentMonth();
  }

  function populateYearSelect(sel, spread) {
    if (!sel || sel.dataset.filled === "1") return;
    sel.dataset.filled = "1";
    yearOptionsThai(spread).forEach(function (y) {
      var opt = document.createElement("option");
      opt.value = y;
      opt.textContent = "พ.ศ. " + y;
      sel.appendChild(opt);
    });
    sel.value = currentYearThai();
  }

  // ---- แปลงวันที่ที่ผู้ใช้พิมพ์แบบย่อ ให้เป็นวันที่เต็มภาษาไทย ----
  // "1/2/2569", "1-2-69", "01.02.2026", "2569-02-01", "1 ก.พ. 69", "1 กุมภา 2569" -> "1 กุมภาพันธ์ 2569"
  // ถ้าแกะไม่ออก (เช่น พิมพ์ข้อความอื่นปนมา) จะคืนค่าที่พิมพ์มาเดิม ไม่ไปแก้ให้เพี้ยน
  var MONTH_ABBR = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
                    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

  function arabicDigits(s) {
    return String(s).replace(/[๐-๙]/g, function (d) { return String("๐๑๒๓๔๕๖๗๘๙".indexOf(d)); });
  }

  // ปีที่พิมพ์มาอาจเป็น พ.ศ. เต็ม (2569), พ.ศ. ย่อ 2 หลัก (69) หรือ ค.ศ. (2026)
  function normalizeYear(y, fallbackYear) {
    if (y === undefined || y === null || y === "") return String(fallbackYear || "").trim();
    var n = parseInt(y, 10);
    if (isNaN(n)) return String(fallbackYear || "").trim();
    if (n < 100) return String(2500 + n);          // 69 -> 2569
    if (n >= 1900 && n <= 2200) return String(n + 543); // ค.ศ. -> พ.ศ.
    return String(n);
  }

  // หาเลขเดือนจากคำที่พิมพ์ (รับได้ทั้งชื่อเต็ม ชื่อย่อ มีจุด/ไม่มีจุด และพิมพ์ไม่ครบคำ)
  function monthIndexFromText(word) {
    var w = String(word).replace(/[.\s]/g, "");
    if (!w) return -1;
    for (var i = 0; i < MONTHS.length; i++) {
      if (MONTHS[i].indexOf(w) === 0) return i;                       // "กุมภา" -> กุมภาพันธ์
      if (MONTH_ABBR[i].replace(/\./g, "") === w) return i;           // "กพ" -> ก.พ.
    }
    return -1;
  }

  function formatThaiDate(raw, fallbackYear) {
    var s = arabicDigits(String(raw == null ? "" : raw)).trim().replace(/\s+/g, " ");
    if (!s) return "";
    var day = null, monthIdx = -1, year = "";

    var m = s.match(/^(\d{1,2})\s*[\/\-.]\s*(\d{1,2})(?:\s*[\/\-.]\s*(\d{1,4}))?$/);
    if (m) {                                   // 1/2/2569, 1-2-69, 1.2
      day = parseInt(m[1], 10); monthIdx = parseInt(m[2], 10) - 1; year = m[3];
    } else if ((m = s.match(/^(\d{4})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{1,2})$/))) {
      day = parseInt(m[3], 10); monthIdx = parseInt(m[2], 10) - 1; year = m[1]; // 2569-02-01
    } else if ((m = s.match(/^(\d{1,2})\s+([฀-๿. ]+?)(?:\s+(\d{1,4}))?$/))) {
      day = parseInt(m[1], 10); monthIdx = monthIndexFromText(m[2]); year = m[3]; // 1 ก.พ. 69
    } else {
      return s;
    }

    if (!(day >= 1 && day <= 31) || monthIdx < 0 || monthIdx > 11) return s;
    var y = normalizeYear(year, fallbackYear);
    return day + " " + MONTHS[monthIdx] + (y ? " " + y : "");
  }

  return {
    MONTHS: MONTHS,
    MONTH_ABBR: MONTH_ABBR,
    formatThaiDate: formatThaiDate,
    toThaiDigits: toThaiDigits,
    currentMonth: currentMonth,
    currentYearThai: currentYearThai,
    yearOptionsThai: yearOptionsThai,
    populateMonthSelect: populateMonthSelect,
    populateYearSelect: populateYearSelect,
  };
})();
