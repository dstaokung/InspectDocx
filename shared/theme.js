(function () {
  "use strict";

  var STORAGE_KEY = "consdocx_theme";
  var QUERY_PARAM = "theme";

  var THEMES = {
    blue: {
      label: "Light Mode",
      accent: "#2563eb", accentDark: "#1d4ed8", accentLight: "#93c5fd", accentBg: "#eff6ff", accentText: "#1e40af",
      pageBg: "#f2f4f7", cardBg: "#ffffff", border: "#e5e7eb", inputBorder: "#d1d5db",
      text: "#1f2937", textMuted: "#6b7280", textFaint: "#9ca3af",
      surface: "#f9fafb", inputBg: "#ffffff", inputText: "#1f2937",
    },
    dark: {
      label: "Dark Mode",
      accent: "#60a5fa", accentDark: "#3b82f6", accentLight: "#3b5a8a", accentBg: "#17233c", accentText: "#bfdbfe",
      pageBg: "#0f172a", cardBg: "#1e293b", border: "#334155", inputBorder: "#334155",
      text: "#e2e8f0", textMuted: "#94a3b8", textFaint: "#64748b",
      surface: "#243144", inputBg: "#1e293b", inputText: "#e2e8f0",
    },
  };
  var THEME_ORDER = ["blue", "dark"];
  var currentTheme = "blue";

  function getQueryTheme() {
    try {
      var params = new URLSearchParams(location.search);
      var t = params.get(QUERY_PARAM);
      if (t && THEMES[t]) return t;
    } catch (e) {}
    return null;
  }

  function getSavedTheme() {
    try {
      var t = localStorage.getItem(STORAGE_KEY);
      if (t && THEMES[t]) return t;
    } catch (e) {}
    return "blue";
  }

  function saveTheme(id) {
    try { localStorage.setItem(STORAGE_KEY, id); } catch (e) {}
  }

  function applyTheme(id, opts) {
    var th = THEMES[id] || THEMES.blue;
    id = THEMES[id] ? id : "blue";
    currentTheme = id;
    var root = document.documentElement.style;
    root.setProperty("--accent", th.accent);
    root.setProperty("--accent-dark", th.accentDark);
    root.setProperty("--accent-light", th.accentLight);
    root.setProperty("--accent-bg", th.accentBg);
    root.setProperty("--accent-text", th.accentText);
    root.setProperty("--page-bg", th.pageBg);
    root.setProperty("--card-bg", th.cardBg);
    root.setProperty("--border", th.border);
    root.setProperty("--input-border", th.inputBorder);
    root.setProperty("--text", th.text);
    root.setProperty("--text-muted", th.textMuted);
    root.setProperty("--text-faint", th.textFaint);
    root.setProperty("--surface", th.surface);
    root.setProperty("--input-bg", th.inputBg);
    root.setProperty("--input-text", th.inputText);
    document.documentElement.setAttribute("data-theme", id);

    if (!opts || !opts.silent) {
      try {
        window.dispatchEvent(new CustomEvent("consdocx:themechange", { detail: { theme: id } }));
      } catch (e) {}
    }
  }

  // ---- เลือกโทนเริ่มต้นตอนโหลดหน้า: ให้ query string ?theme=... (ถ้ามี ส่งมาจากลิงก์ของหน้าอื่น)
  // ชนะ localStorage ของหน้านี้เอง เพราะไฟล์ .html แต่ละไฟล์ที่เปิดแบบ file:// จะแยก localStorage
  // กันคนละก้อน (คนละ origin) ไม่แชร์กันเอง ต้องส่งค่าผ่าน URL แทนถึงจะพาโทนสีข้ามหน้าไปด้วยได้จริง ----
  var initial = getQueryTheme() || getSavedTheme();
  saveTheme(initial); // จำไว้ในหน้านี้เองด้วย เผื่อเปิดหน้านี้ตรงๆ ครั้งต่อไป
  applyTheme(initial, { silent: true }); // เงียบตอน initial กันหน้าอื่นที่ยังไม่ได้ผูก listener งง

  function withTheme(url) {
    var sep = url.indexOf("?") === -1 ? "?" : "&";
    return url + sep + QUERY_PARAM + "=" + encodeURIComponent(currentTheme);
  }

  window.ConsDocxTheme = {
    get: function () { return currentTheme; },
    withTheme: withTheme,
    THEME_ORDER: THEME_ORDER.slice(),
  };

  function buildSwitcher() {
    if (document.getElementById("themeSwitcher")) return;

    var wrap = document.createElement("div");
    wrap.id = "themeSwitcher";
    wrap.style.cssText =
      "position:fixed; bottom:20px; right:20px; z-index:9999; " +
      "font-family:'TH Sarabun New','Sarabun',sans-serif;";

    var btn = document.createElement("button");
    btn.type = "button";
    btn.title = "เปลี่ยนโทนสีโปรแกรม";
    btn.textContent = "🎨";
    btn.style.cssText =
      "width:46px; height:46px; border-radius:50%; border:1px solid #e5e7eb; " +
      "background:#fff; box-shadow:0 2px 10px rgba(0,0,0,.18); font-size:20px; " +
      "cursor:pointer; display:flex; align-items:center; justify-content:center; padding:0;";

    var popover = document.createElement("div");
    popover.style.cssText =
      "display:none; position:absolute; bottom:56px; right:0; background:#fff; " +
      "border:1px solid #e5e7eb; border-radius:10px; box-shadow:0 4px 18px rgba(0,0,0,.18); " +
      "padding:8px; min-width:170px;";

    var heading = document.createElement("div");
    heading.textContent = "โทนสีโปรแกรม";
    heading.style.cssText = "font-size:13px; font-weight:600; color:#6b7280; padding:4px 8px 6px;";
    popover.appendChild(heading);

    var hintNote = document.createElement("div");
    hintNote.textContent = "เลือกที่หน้านี้แล้วกดลิงก์ไปหน้าอื่นต่อ จะพาโทนสีติดไปด้วย";
    hintNote.style.cssText = "font-size:11px; color:#9ca3af; padding:0 8px 8px; max-width:180px;";
    popover.appendChild(hintNote);

    THEME_ORDER.forEach(function (id) {
      var th = THEMES[id];
      var row = document.createElement("button");
      row.type = "button";
      row.style.cssText =
        "display:flex; align-items:center; gap:8px; width:100%; padding:8px 10px; " +
        "border:none; background:#fff; border-radius:6px; cursor:pointer; font-size:14px; " +
        "color:#1f2937; text-align:left; font-family:inherit;";
      row.addEventListener("mouseenter", function () { row.style.background = "#f3f4f6"; });
      row.addEventListener("mouseleave", function () { row.style.background = "#fff"; });

      var dot = document.createElement("span");
      dot.style.cssText =
        "display:inline-block; width:16px; height:16px; border-radius:50%; " +
        "background:" + th.accent + "; flex-shrink:0; border:1px solid rgba(0,0,0,.08);";

      row.appendChild(dot);
      row.appendChild(document.createTextNode(th.label));
      row.addEventListener("click", function () {
        applyTheme(id);
        saveTheme(id);
        popover.style.display = "none";
      });
      popover.appendChild(row);
    });

    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      popover.style.display = popover.style.display === "none" ? "block" : "none";
    });
    document.addEventListener("click", function (e) {
      if (!wrap.contains(e.target)) popover.style.display = "none";
    });

    wrap.appendChild(popover);
    wrap.appendChild(btn);
    document.body.appendChild(wrap);
  }

  function init() {
    buildSwitcher();
  }
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }

  // สลับสีพร้อมกันทุกแท็บ/หน้าต่างที่เปิด "ไฟล์เดียวกัน" ค้างไว้ (คนละไฟล์ยังต้องพาไปผ่านลิงก์ตามปกติ)
  window.addEventListener("storage", function (e) {
    if (e.key === STORAGE_KEY && e.newValue && THEMES[e.newValue]) {
      applyTheme(e.newValue);
    }
  });
})();
