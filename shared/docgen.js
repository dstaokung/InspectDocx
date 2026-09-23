// ============================================================================
// docgen.js — เครื่องมือเติม [placeholder] ลงไฟล์ .docx (ใช้ได้ทั้งในเบราว์เซอร์และ Node)
//
// หลักการเดียวกับ ConsDocx: เทมเพลต .docx มีข้อความ [ชื่อ_placeholder] พิมพ์ไว้ธรรมดา
// โค้ดเปิดไฟล์ด้วย JSZip → แก้ word/document.xml → แทนที่ข้อความ → zip กลับเป็น .docx
//
// ค่าที่ส่งเข้า buildDocx():
//   "ข้อความ"                 → แทนที่ตามปกติ (รูปแบบตัวอักษรเดิมของเทมเพลต)
//   { text, isDefault:true }   → แทนที่ด้วยตัวอักษร "สีแดง" (ช่องที่ผู้ใช้ยังไม่กรอก ให้เห็นชัดว่าต้องแก้)
//   null                       → ลบทั้งย่อหน้าที่มี placeholder นั้นทิ้ง
// ============================================================================
(function (root) {
  "use strict";

  var DEFAULT_TEXT_COLOR = "FF0000";
  var BLANK = "..................";   // ข้อความตั้งต้นเมื่อช่องว่าง (ออกเป็นสีแดง เขียนมือเติมได้)

  function asDefault(text) { return { text: String(text), isDefault: true }; }

  function base64ToBytes(b64) {
    var bin = atob(b64);
    var out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function xmlEscape(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  function getParaText(para) {
    return (para.match(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g) || [])
      .map(function (t) { return t.replace(/<[^>]+>/g, ""); }).join("");
  }

  // รวม "[placeholder]" ที่ Word แยกไว้หลายรัน (เกิดเมื่อแก้ไขเทมเพลตทีละตัวอักษร) ให้อยู่รันเดียว
  function normalizePlaceholderRuns(xml) {
    xml = xml.replace(/<w:proofErr[^>]*\/>/g, "");
    var runRe = /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g;
    var rPrRe = /<w:rPr>[\s\S]*?<\/w:rPr>/;
    var textRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/;
    var result = "", lastIndex = 0, pending = null, m;
    while ((m = runRe.exec(xml)) !== null) {
      var gap = xml.slice(lastIndex, m.index);
      lastIndex = runRe.lastIndex;
      var runXml = m[0];
      if (pending && gap.indexOf("</w:p>") !== -1) { result += pending.raw; pending = null; }
      result += gap;
      var tm = runXml.match(textRe);
      if (!tm) {
        if (pending) { result += pending.raw; pending = null; }
        result += runXml;
        continue;
      }
      var text = tm[1];
      if (pending) {
        pending.text += text;
        pending.raw += runXml;
        if (text.indexOf("]") !== -1) {
          result += "<w:r>" + pending.rPr + '<w:t xml:space="preserve">' + pending.text + "</w:t></w:r>";
          pending = null;
        }
        continue;
      }
      var b = text.lastIndexOf("[");
      if (b !== -1 && text.slice(b).indexOf("]") === -1) {
        var rp = runXml.match(rPrRe);
        pending = { rPr: rp ? rp[0] : "", text: text, raw: runXml };
        continue;
      }
      result += runXml;
    }
    result += xml.slice(lastIndex);
    if (pending) result += pending.raw;
    return result;
  }

  // ใส่สีแดงลง <w:rPr> (ลำดับ element ใน rPr มีผลกับ Word: w:color ต้องมาก่อน w:sz/w:lang ฯลฯ)
  function rPrAsDefault(rPr) {
    var colorTag = '<w:color w:val="' + DEFAULT_TEXT_COLOR + '"/>';
    if (!rPr) return "<w:rPr>" + colorTag + "</w:rPr>";
    var out = rPr.replace(/<w:color[^>]*\/>/g, "");
    var anchor = out.match(/<w:(spacing|w|kern|position|sz|szCs|highlight|u|effect|bdr|shd|vertAlign|rtl|cs|lang)[ />]/);
    if (anchor) return out.slice(0, anchor.index) + colorTag + out.slice(anchor.index);
    return out.replace("</w:rPr>", colorTag + "</w:rPr>");
  }

  function replacePlaceholder(xml, placeholder, value, isDefault) {
    var esc = placeholder.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var re = new RegExp("<w:r(?:\\s[^>]*)?>((?:(?!</w:r>)[\\s\\S])*?)<w:t([^>]*)>([^<]*" + esc + "[^<]*)</w:t></w:r>", "g");
    var result = "", lastIndex = 0, m;
    var lines = String(value).split(/\n+/);
    while ((m = re.exec(xml)) !== null) {
      result += xml.slice(lastIndex, m.index);
      var rp = m[1].match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
      var rPr = rp ? rp[0] : "";
      var rPrVal = isDefault ? rPrAsDefault(rPr) : rPr;
      var parts = [];
      m[3].split(placeholder).forEach(function (seg, i) {
        if (i > 0) {
          parts.push("<w:r>" + rPrVal + '<w:t xml:space="preserve">' + xmlEscape(lines[0]) + "</w:t></w:r>");
          for (var j = 1; j < lines.length; j++) {
            parts.push("<w:r>" + rPrVal + '<w:br/><w:t xml:space="preserve">' + xmlEscape(lines[j]) + "</w:t></w:r>");
          }
        }
        if (seg) parts.push("<w:r>" + rPr + '<w:t xml:space="preserve">' + xmlEscape(seg) + "</w:t></w:r>");
      });
      result += parts.join("");
      lastIndex = m.index + m[0].length;
      re.lastIndex = lastIndex;
    }
    return result + xml.slice(lastIndex);
  }

  // ======================================================================== รูปถ่าย
  // เทมเพลตมีกรอบรูปเป็น SmartArt (รูปภาพ 2×2) — โปรแกรมอ่านตำแหน่ง/ขนาดกรอบรูปจาก
  // word/diagrams/drawing1.xml แล้วถ้าผู้ใช้ใส่รูปอย่างน้อย 1 รูป จะ "ลบ SmartArt ทิ้ง" แล้ววาง
  // รูปธรรมดา (floating picture) ลงตำแหน่งเดิมเป๊ะๆ แทน · ไม่ใส่รูปเลย = คง SmartArt ไว้ตามเดิม
  var DIAGRAM_DRAWING_RE = /<w:drawing>(?:(?!<\/w:drawing>)[\s\S])*?drawingml\/2006\/diagram(?:(?!<\/w:drawing>)[\s\S])*?<\/w:drawing>/;
  var MIN_SLOT_EMU = 1000000; // กรอบที่เล็กกว่านี้ (กล่องคำบรรยายเล็กๆ ใน SmartArt) ไม่นับเป็นช่องรูป
  // PNG สีขาว 1×1 — ใช้กับช่องที่ไม่ได้ใส่รูป (ได้กรอบเส้นเทาว่างๆ ไว้)
  var BLANK_PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4//8/AAX+Av4N70a4AAAAAElFTkSuQmCC";

  // คืน { slots:[{x,y,cx,cy}], frame:{hRel,hOff,vRel,vOff} } หรือ null ถ้าเทมเพลตไม่มี SmartArt รูป
  function readPhotoSlots(zip, docXml) {
    var m = docXml.match(DIAGRAM_DRAWING_RE);
    var dgm = zip.file("word/diagrams/drawing1.xml");
    if (!m || !dgm) return Promise.resolve(null);
    var d = m[0];
    var ph = d.match(/<wp:positionH relativeFrom="(\w+)">\s*<wp:posOffset>(-?\d+)<\/wp:posOffset>/);
    var pv = d.match(/<wp:positionV relativeFrom="(\w+)">\s*<wp:posOffset>(-?\d+)<\/wp:posOffset>/);
    if (!ph || !pv) return Promise.resolve(null);
    return dgm.async("string").then(function (dx) {
      var slots = [];
      var re = /<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/g, s;
      while ((s = re.exec(dx)) !== null) {
        var r = { x: +s[1], y: +s[2], cx: +s[3], cy: +s[4] };
        if (r.cx >= MIN_SLOT_EMU && r.cy >= MIN_SLOT_EMU) slots.push(r);
      }
      // เรียงซ้าย→ขวา บน→ล่าง (y ต่างกันไม่เกิน 0.5 ซม. ถือว่าแถวเดียวกัน)
      slots.sort(function (a, b) { return Math.abs(a.y - b.y) > 180000 ? a.y - b.y : a.x - b.x; });
      return { slots: slots, frame: { hRel: ph[1], hOff: +ph[2], vRel: pv[1], vOff: +pv[2] } };
    });
  }

  // ให้หน้าเว็บรู้จำนวนช่องรูป + สัดส่วนกว้าง:สูง (ไว้ครอปรูปให้พอดีกรอบ)
  function getPhotoSlots(JSZip, base64) {
    return JSZip.loadAsync(base64ToBytes(base64)).then(function (zip) {
      return zip.file("word/document.xml").async("string").then(function (xml) {
        return readPhotoSlots(zip, xml);
      });
    }).then(function (info) {
      return info ? info.slots.map(function (s) { return { cx: s.cx, cy: s.cy, aspect: s.cx / s.cy }; }) : [];
    });
  }

  function pictureXml(i, rId, slot, frame, name) {
    var A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    var PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
    return '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' + (251700000 + i) +
      '" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="' + frame.hRel + '"><wp:posOffset>' + (frame.hOff + slot.x) + '</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="' + frame.vRel + '"><wp:posOffset>' + (frame.vOff + slot.y) + '</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="' + slot.cx + '" cy="' + slot.cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
      '<wp:docPr id="' + (7000 + i) + '" name="' + name + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks ' + A + ' noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic ' + A + '><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic ' + PIC + '><pic:nvPicPr><pic:cNvPr id="0" name="' + name + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + slot.cx + '" cy="' + slot.cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="9525"><a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill></a:ln></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>';
  }

  // photos = [ Uint8Array(JPEG) | null, ... ] ตามลำดับช่อง (ที่หน้าเว็บครอปพอดีกรอบมาแล้ว)
  function insertPhotos(zip, xml, photos) {
    var has = (photos || []).some(function (p) { return p && p.length; });
    if (!has) return Promise.resolve(xml);
    return readPhotoSlots(zip, xml).then(function (info) {
      if (!info || !info.slots.length) return xml;
      return zip.file("word/_rels/document.xml.rels").async("string").then(function (rels) {
        var pics = info.slots.map(function (slot, i) {
          var p = photos[i];
          var ext = p && p.length ? "jpeg" : "png";
          var file = "inspect_photo" + (i + 1) + "." + ext;
          zip.file("word/media/" + file, p && p.length ? p : base64ToBytes(BLANK_PNG_B64), { createFolders: false });
          var rId = "rIdInspectPhoto" + (i + 1);
          rels = rels.replace("</Relationships>", '<Relationship Id="' + rId +
            '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + file + '"/></Relationships>');
          return pictureXml(i, rId, slot, info.frame, "รูปที่ " + (i + 1));
        });
        zip.file("word/_rels/document.xml.rels", rels);
        return zip.file("[Content_Types].xml").async("string").then(function (ct) {
          [["jpeg", "image/jpeg"], ["png", "image/png"]].forEach(function (e) {
            if (ct.indexOf('Extension="' + e[0] + '"') === -1)
              ct = ct.replace("<Default ", '<Default Extension="' + e[0] + '" ContentType="' + e[1] + '"/><Default ');
          });
          zip.file("[Content_Types].xml", ct);
          return xml.replace(DIAGRAM_DRAWING_RE, pics.join(""));
        });
      });
    });
  }

  // สร้าง .docx จาก base64 ของเทมเพลต + ค่า placeholder (+ รูป) → คืน Uint8Array
  function buildDocx(JSZip, base64, fields, photos) {
    return JSZip.loadAsync(base64ToBytes(base64)).then(function (zip) {
      return zip.file("word/document.xml").async("string").then(function (xml) {
        return insertPhotos(zip, xml, photos);
      }).then(function (xml) {
        xml = normalizePlaceholderRuns(xml);
        Object.keys(fields).forEach(function (ph) {
          var raw = fields[ph];
          if (raw === null) {
            xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, function (p) { return getParaText(p).indexOf(ph) !== -1 ? "" : p; });
          } else {
            var isDef = !!(raw && typeof raw === "object" && raw.isDefault);
            xml = replacePlaceholder(xml, ph, isDef ? raw.text : (raw == null ? "" : raw), isDef);
          }
        });
        zip.file("word/document.xml", xml);
        return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
      });
    });
  }

  // placeholder ที่ยังค้างในไฟล์ (ใช้ตอนทดสอบ)
  function leftoverPlaceholders(xml) {
    return (getParaTextAll(xml).match(/\[[^\]\s]{1,40}\]/g) || []);
  }
  function getParaTextAll(xml) {
    return (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map(getParaText).join("\n");
  }

  // ---------------------------------------------------------------- ตัวช่วยแปลงค่า
  function arabic(s) {
    return String(s == null ? "" : s).replace(/[๐-๙]/g, function (d) { return String("๐๑๒๓๔๕๖๗๘๙".indexOf(d)); });
  }

  function bahtText(numStr) {
    var cleaned = arabic(numStr).replace(/[,\s฿]/g, "").replace(/บาท$/, "");
    if (!cleaned || !/^\d*\.?\d*$/.test(cleaned) || cleaned === ".") return "";
    var sp = cleaned.split("."), intPart = sp[0], decPart = sp[1] || "";
    intPart = intPart.replace(/^0+(?=\d)/, "") || "0";
    var satang = parseInt((decPart + "00").slice(0, 2), 10) || 0;
    var digits = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
    var posName = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"];
    function readGroup(s, hasPrefix) {
      var out = "", n = s.length;
      for (var i = 0; i < n; i++) {
        var d = +s[i], p = n - i - 1;
        if (d === 0) continue;
        if (p === 0 && d === 1 && (n > 1 || hasPrefix)) out += "เอ็ด";
        else if (p === 1 && d === 2) out += "ยี่สิบ";
        else if (p === 1 && d === 1) out += "สิบ";
        else out += digits[d] + posName[p];
      }
      return out;
    }
    function readNumber(s) {
      s = s.replace(/^0+/, "");
      if (!s) return "";
      if (s.length <= 6) return readGroup(s, false);
      var head = s.slice(0, -6), tail = s.slice(-6).replace(/^0+/, "");
      return readNumber(head) + "ล้าน" + (tail ? readGroup(tail, true) : "");
    }
    var out = "";
    if (intPart !== "0") out += readNumber(intPart) + "บาท";
    else if (satang === 0) return "ศูนย์บาทถ้วน";
    if (satang > 0) out += readGroup(String(satang).padStart(2, "0").replace(/^0+/, ""), false) + "สตางค์";
    else out += "ถ้วน";
    return out;
  }

  // "4867000" → "4,867,000.00"   (คืนข้อความเดิมถ้าไม่ใช่ตัวเลข)
  function formatMoney(numStr) {
    var cleaned = arabic(numStr).replace(/[,\s฿]/g, "").replace(/บาท$/, "");
    if (!cleaned || !/^\d*\.?\d*$/.test(cleaned) || cleaned === ".") return String(numStr || "").trim();
    var n = Number(cleaned);
    return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  // "9" / "9.3" / "9:30" / "0930" / "9.30 น." → "09.30"
  function formatTime(raw) {
    var s = arabic(raw).trim().replace(/\s*น\.?$/, "").trim();
    if (!s) return "";
    var m = s.match(/^(\d{1,2})(?:\s*[.:]\s*(\d{1,2}))?$/) || s.match(/^(\d{2})(\d{2})$/);
    if (!m) return s;
    var h = parseInt(m[1], 10), mi = m[2] === undefined ? 0 : parseInt(m[2].length === 1 ? m[2] + "0" : m[2], 10);
    if (h > 23 || mi > 59) return s;
    return (h < 10 ? "0" : "") + h + "." + (mi < 10 ? "0" : "") + mi;
  }

  function thaiDate(raw) {
    var MY = root.ConsDocxMonthYear;
    var y = MY ? MY.currentYearThai() : "";
    return MY ? MY.formatThaiDate(raw, y) : String(raw || "").trim();
  }

  // ---------------------------------------------------------------- ข้อมูลฟอร์ม → placeholder
  // data = ค่าจากฟอร์ม (key ตรงกับ id ช่องใน index.html) — ดูรายละเอียดใน PLACEHOLDERS.md
  function buildFields(data) {
    function v(k) { return String(data[k] == null ? "" : data[k]).trim(); }
    function orBlank(s) { return s ? s : asDefault(BLANK); }
    var inspectDate = thaiDate(v("inspectDate"));
    var docDate = thaiDate(v("docDate")) || inspectDate;   // ไม่กรอกวันที่หนังสือ = ใช้วันที่ออกตรวจ
    var money = v("contractAmount");
    var f = {
      "[วันที่หนังสือ]": orBlank(docDate),
      "[ครั้งที่]": orBlank(arabic(v("round"))),
      "[เลขที่คำสั่ง]": orBlank(arabic(v("orderNo"))),
      "[วันที่คำสั่ง]": orBlank(thaiDate(v("orderDate"))),
      "[เลขที่สัญญา]": orBlank(arabic(v("contractNo"))),
      "[วันที่สัญญา]": orBlank(thaiDate(v("contractDate"))),
      "[วันเริ่มสัญญา]": orBlank(thaiDate(v("startDate"))),
      "[วันสิ้นสุดสัญญา]": orBlank(thaiDate(v("endDate"))),
      "[วงเงินตัวเลข]": orBlank(money ? formatMoney(money) : ""),
      "[วงเงินตัวอักษร]": orBlank(money ? bahtText(money) : ""),
      "[ผู้รับจ้าง]": orBlank(v("contractor")),
      "[ชื่อโครงการ]": orBlank(v("projectName").replace(/\s*\n\s*/g, " ")),
      "[วันที่ออกตรวจ]": orBlank(inspectDate),
      "[เวลาเริ่ม]": orBlank(formatTime(v("timeStart"))),
      "[เวลาสิ้นสุด]": orBlank(formatTime(v("timeEnd"))),
      "[ผลการตรวจ]": orBlank(v("result")),
    };
    for (var i = 1; i <= 3; i++) {
      // คำนำหน้า (ยศ) — ไม่มีก็ปล่อยว่างจริงๆ ไม่ใส่ข้อความสีแดง
      f["[คำนำหน้า" + i + "]"] = v("prefix" + i);
      f["[ชื่อกรรมการ" + i + "]"] = orBlank(v("name" + i));
    }
    return f;
  }

  var api = {
    BLANK: BLANK, asDefault: asDefault, buildDocx: buildDocx, buildFields: buildFields, getPhotoSlots: getPhotoSlots,
    normalizePlaceholderRuns: normalizePlaceholderRuns, replacePlaceholder: replacePlaceholder,
    leftoverPlaceholders: leftoverPlaceholders, getParaTextAll: getParaTextAll,
    bahtText: bahtText, formatMoney: formatMoney, formatTime: formatTime, thaiDate: thaiDate,
  };
  root.InspectDocGen = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
