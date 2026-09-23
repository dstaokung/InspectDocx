// ทดสอบสร้างเอกสารด้วยโค้ดจริง (shared/docgen.js) แล้วเช็คว่าไม่มี [placeholder] ค้าง
//   node tools/test-build.js            → สร้าง tools/_test-output-*.docx ไว้เปิดดู
const fs = require("fs"), path = require("path");
const ROOT = path.dirname(__dirname);
globalThis.window = globalThis;
require(path.join(ROOT, "shared/monthyear.js"));
const JSZip = require(path.join(ROOT, "shared/jszip.lib.js"));
const G = require(path.join(ROOT, "shared/docgen.js"));
eval(fs.readFileSync(path.join(ROOT, "template-data.js"), "utf8").replace(/^var /m, "globalThis."));

const FULL = {
  round: "3", docDate: "", inspectDate: "5/9/2568", timeStart: "9", timeEnd: "9.45",
  orderNo: "655/2568", orderDate: "7 พ.ค. 68",
  contractNo: "27/2568", contractDate: "1/7/2568", startDate: "2/7/68", endDate: "28/12/2568",
  contractAmount: "4867000", contractor: "บริษัท รุ่งสิริวัฒน์ จำกัด",
  projectName: "โครงการก่อสร้างถนน ค.ส.ล. พร้อมวางท่อระบายน้ำ ซอยหม้อน้ำ หมู่ที่ 2",
  prefix1: "", name1: "นางมาลัย  เลิศฤทธิ์",
  prefix2: "ว่าที่ร้อยตรี", name2: "คมสรร  เครือศรี",
  prefix3: "", name3: "นายกำชัย  ทองสุทธิ์",
  result: "ผู้รับจ้างสามารถวางบ่อพัก ค.ส.ล. และการวางท่อระบายน้ำ ค.ส.ล.",
};
const CASES = { full: FULL, empty: {}, photos: FULL };
// รูปทดสอบ: ใส่ช่อง 1, 2, 4 (เว้นช่อง 3 ว่าง ต้องได้กรอบเปล่า)
const PHOTOS = [1, 2, null, 3].map(n => n ? new Uint8Array(fs.readFileSync(path.join(__dirname, `fixtures/photo${n}.jpg`))) : null);

(async () => {
  let fail = 0;
  for (const [name, data] of Object.entries(CASES)) {
    const bytes = await G.buildDocx(JSZip, globalThis.TEMPLATE_INSPECT_BASE64, G.buildFields(data), name === "photos" ? PHOTOS : null);
    const out = path.join(__dirname, `_test-output-${name}.docx`);
    fs.writeFileSync(out, bytes);
    const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml").async("string");
    const left = G.leftoverPlaceholders(xml);
    console.log(`${left.length ? "FAIL" : "ok  "} ${name}: ${left.length ? "ค้าง " + left.join(" ") : "ไม่มี placeholder ค้าง"}`);
    if (left.length) fail++;
    const zip2 = await JSZip.loadAsync(bytes);
    const nPic = (xml.match(/<pic:pic /g) || []).length, hasDgm = xml.includes("drawingml/2006/diagram");
    const media = Object.keys(zip2.files).filter(f => f.startsWith("word/media/inspect_photo"));
    const expectPics = name === "photos" ? 4 : 0;
    const okPic = nPic === expectPics && hasDgm === (expectPics === 0) && media.length === expectPics;
    console.log(`${okPic ? "ok  " : "FAIL"} ${name}: รูป ${nPic} ช่อง, SmartArt ${hasDgm ? "คงไว้" : "ถูกแทนที่"}, media ${media.join(" ")}`);
    if (!okPic) fail++;
    if (name === "full" && process.argv.includes("-v")) console.log(G.getParaTextAll(xml).split("\n").filter(s => s.trim()).slice(0, 60).join("\n"));
  }
  process.exit(fail ? 1 : 0);
})();
