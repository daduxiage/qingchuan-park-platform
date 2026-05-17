const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
const IMG = "/Users/wangcaihua/WorkBuddy/20260427142154/slides/imgs";

// Color Palettes per role (same master but different accent)
const M = {
  bg: "F8FAFC", card: "FFFFFF", text: "1E293B", gray: "64748B",
  light: "E2E8F0", darkBg: "0F172A"
};
// Role colors
const R = {
  gov: { primary: "1E40AF", light: "DBEAFE", name: "政府领导层" },
  mgmt: { primary: "059669", light: "D1FAE5", name: "园区管理层" },
  ent: { primary: "D97706", light: "FEF3C7", name: "入驻企业" },
  pub: { primary: "7C3AED", light: "EDE9FE", name: "公众用户" }
};

let pn = 0;

function badge(sl, n, c) {
  sl.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.38, h: 0.38, fill: { color: c || "1E40AF" } });
  sl.addText(String(n), { x: 9.3, y: 5.1, w: 0.38, h: 0.38, fontSize: 10, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
}
function topBar(sl, c) { sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: c || "1E40AF" } }); }
function dot(sl, x, y, s, c) { sl.addShape(pres.shapes.OVAL, { x, y, w: s, h: s, fill: { color: c || "1E40AF" } }); }

function roleSlide(title, items, imgFile, role, accent) {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.bg };
  topBar(sl, accent);
  // Decorative dots
  for (var i = 0; i < 6; i++) dot(sl, 0.3 + i * 0.2, 0.07, 0.035, accent);
  // Role badge top-right
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.8, y: 0.12, w: 1.9, h: 0.28, fill: { color: accent }, rectRadius: 0.04 });
  sl.addText(role, { x: 7.8, y: 0.12, w: 1.9, h: 0.28, fontSize: 8, fontFace: "Microsoft YaHei", color: "FFFFFF", align: "center", valign: "middle" });
  // Title
  sl.addText(title, { x: 0.5, y: 0.15, w: 7, h: 0.5, fontSize: 20, fontFace: "Microsoft YaHei", color: M.text, bold: true });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.65, w: 0.8, h: 0.03, fill: { color: accent } });
  // Items
  var yPos = 0.9;
  items.forEach(function(item, i) {
    dot(sl, 0.5, yPos + 0.06, 0.07, accent);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: yPos - 0.02, w: 4.5, h: 0.3, fill: { color: M.card }, rectRadius: 0.04, shadow: { type: "outer", blur: 2, offset: 1, color: "000000", opacity: 0.03 } });
    sl.addText(item, { x: 0.85, y: yPos - 0.02, w: 4.2, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: M.text, valign: "middle" });
    yPos += 0.38;
  });
  imgFile = (/^IMG\//.test(imgFile)) ? IMG + imgFile.slice(3) : imgFile;
  if (imgFile && fs.existsSync(imgFile)) {
    sl.addImage({ path: imgFile, x: 5.4, y: 0.5, w: 4.3, h: 4.7, sizing: { type: "contain" } });
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 0.5, w: 4.3, h: 4.7, line: { color: M.light, width: 0.5 } });
  }
  badge(sl, pn, accent);
}

function chineseNum(n) {
  var map = ["\u2460","\u2461","\u2462","\u2463","\u2464","\u2465","\u2466","\u2467","\u2468","\u2469"];
  return map[n-1] || String(n);
}

// ====== 1. COVER ======
function s1() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.darkBg };
  topBar(sl, "F59E0B");
  // Large decorative circles
  sl.addShape(pres.shapes.OVAL, { x: 6.5, y: -2, w: 6, h: 6, fill: { color: "1E3A5F", transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -1, y: 2.5, w: 4, h: 4, fill: { color: "1E3A5F", transparency: 60 } });
  for (var i = 0; i < 10; i++) dot(sl, 0.3 + i * 0.22, 0.07, 0.04, "F59E0B");
  
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.8, w: 0.07, h: 2.8, fill: { color: "F59E0B" } });
  sl.addText("\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a", { x: 0.8, y: 0.8, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: "F59E0B" });
  sl.addText("\u667a\u6167\u56ed\u533a\u5e73\u53f0\u6c47\u62a5", { x: 0.8, y: 1.3, w: 8, h: 1.0, fontSize: 36, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addText("\u5206\u89d2\u8272\u529f\u80fd\u5c55\u793a \u00b7 \u5efa\u8bbe\u65b9\u6848", { x: 0.8, y: 2.3, w: 6, h: 0.4, fontSize: 14, fontFace: "Microsoft YaHei", color: "94A3B8" });
  
  // Role tags
  var roles = [
    ["\u653f\u5e9c\u9886\u5bfc\u5c42", "1E40AF"], ["\u56ed\u533a\u7ba1\u7406\u5c42", "059669"],
    ["\u5165\u9a7b\u4f01\u4e1a", "D97706"], ["\u516c\u4f17\u7528\u6237", "7C3AED"]
  ];
  roles.forEach(function(r, i) {
    var x = 0.8 + i * 2.15;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.2, w: 2.0, h: 0.35, fill: { color: r[1] }, rectRadius: 0.06 });
    sl.addText(r[0], { x, y: 3.2, w: 2.0, h: 0.35, fontSize: 9, fontFace: "Microsoft YaHei", color: "FFFFFF", align: "center", valign: "middle" });
  });
  
  // Real data
  sl.addText("\u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf\u7af9\u56ed\u9547", { x: 0.6, y: 3.85, w: 5, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: "94A3B8" });
  sl.addText("2025\u5e74\u524d\u4e09\u5b63\u5ea6GDP\u00a057.68\u4ebf\u5143 \u00b7 \u5168\u53bf\u589e\u957f15% \u00b7 2026\u5e745\u6708", { x: 0.6, y: 4.15, w: 8, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: "64748B" });
  
  // AI image on right
  if (fs.existsSync(IMG + "/ai_smart_park_dashboard.png")) {
    sl.addImage({ path: IMG + "/ai_smart_park_dashboard.png", x: 6.0, y: 3.5, w: 3.7, h: 1.9, sizing: { type: "cover", w: 3.7, h: 1.9 } });
  }
}

// ====== 2. TOC ======
function s2() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.bg };
  topBar(sl, "1E40AF");
  for (var i = 0; i < 8; i++) dot(sl, 0.3 + i * 0.2, 0.07, 0.035, "1E40AF");
  sl.addText("\u6c47\u62a5\u76ee\u5f55", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 22, fontFace: "Microsoft YaHei", color: M.text, bold: true });
  var sections = [
    ["\u9879\u76ee\u80cc\u666f\u4e0e\u9752\u5ddd\u6570\u636e", "1E40AF"],
    ["\u653f\u5e9c\u9886\u5bfc\u5c42\u89c6\u89d2", "1E40AF"],
    ["\u56ed\u533a\u7ba1\u7406\u5c42\u89c6\u89d2", "059669"],
    ["\u5165\u9a7b\u4f01\u4e1a\u89c6\u89d2", "D97706"],
    ["\u516c\u4f17\u7528\u6237\u89c6\u89d2", "7C3AED"],
    ["\u56db\u7c7b\u89d2\u8272\u529f\u80fd\u77e9\u9635", "F59E0B"],
    ["\u6295\u8d44\u4e0e\u6548\u76ca\u5206\u6790", "F59E0B"],
    ["\u672a\u6765\u5c55\u671b", "1E40AF"]
  ];
  sections.forEach(function(s, i) {
    var y = 1.0 + i * 0.5;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.42, fill: { color: M.card }, rectRadius: 0.06, shadow: { type: "outer", blur: 2, offset: 1, color: "000000", opacity: 0.03 } });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 0.32, h: 0.42, fill: { color: s[1] }, rectRadius: 0.03 });
    sl.addText(String(i+1), { x: 0.5, y, w: 0.32, h: 0.42, fontSize: 11, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    dot(sl, 9.1, y + 0.14, 0.06, s[1]);
    sl.addText(s[0], { x: 0.95, y, w: 7, h: 0.42, fontSize: 12, fontFace: "Microsoft YaHei", color: M.text, valign: "middle" });
  });
  badge(sl, pn, "1E40AF");
}

// ====== 3. Background with real data ======
function s3() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.bg };
  topBar(sl, "1E40AF");
  for (var i = 0; i < 8; i++) dot(sl, 0.3 + i * 0.2, 0.07, 0.035, "1E40AF");
  sl.addText("\u9879\u76ee\u80cc\u666f\u4e0e\u9752\u5ddd\u6570\u636e", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 22, fontFace: "Microsoft YaHei", color: M.text, bold: true });
  sl.addText("\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a \u00b7 \u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf", { x: 0.5, y: 0.65, w: 8, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: M.gray });
  
  // Location box
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.05, w: 9, h: 0.55, fill: { color: "DBEAFE" }, rectRadius: 0.06 });
  sl.addText("\u9752\u5ddd\u53bf\u4f4d\u4e8e\u56db\u5ddd\u7701\u5317\u90e8\u8fb9\u7f18\uff0c\u5ddd\u7518\u9655\u4e09\u7701\u7ed3\u5408\u90e8\uff0c\u7d20\u6709\u201c\u91d1\u4e09\u89d2\u201d\u4e4b\u79f0\u3002\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u89c4\u5212\u9762\u79ef2.1\u5e73\u65b9\u516c\u91cc\uff0c\u5df2\u5f00\u53d11.6\u5e73\u65b9\u516c\u91cc\u3002", { x: 0.7, y: 1.1, w: 8.6, h: 0.45, fontSize: 11, fontFace: "Microsoft YaHei", color: M.text, valign: "middle" });
  
  // Data cards with real Qingchuan data
  var data = [
    ["57.68\u4ebf", "2025\u524d\u4e09\u5b63\u5ea6GDP", "\u5168\u53bf\u589e\u957f15%", "1E40AF"],
    ["137\u4ebf", "\u53bf\u5c5e\u56fd\u4f01\u8d44\u4ea7\u603b\u989d", "\u8425\u4e1a\u6536\u51653.44\u4ebf", "059669"],
    ["45\u5bb6", "\u56fd\u5bb6\u79d1\u6280\u578b\u4e2d\u5c0f\u4f01\u4e1a", "\u65b0\u589e\u9ad8\u65b0\u6280\u672f\u4f018\u5bb6", "D97706"],
    ["47\u5bb6", "\u56ed\u533a\u5165\u9a7b\u4f01\u4e1a", "\u4ece\u4e1a\u4eba\u54583,000+", "7C3AED"]
  ];
  data.forEach(function(d, i) {
    var x = 0.5 + i * 2.3, y = 1.8;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.1, h: 1.4, fill: { color: M.card }, rectRadius: 0.08, shadow: { type: "outer", blur: 3, offset: 1, color: "000000", opacity: 0.03 } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.1, h: 0.04, fill: { color: d[3] } });
    sl.addText(d[0], { x, y: y + 0.15, w: 2.1, h: 0.45, fontSize: 24, fontFace: "Arial", color: d[3], bold: true, align: "center" });
    sl.addText(d[1], { x, y: y + 0.6, w: 2.1, h: 0.25, fontSize: 10, fontFace: "Microsoft YaHei", color: M.text, bold: true, align: "center" });
    sl.addText(d[2], { x: x + 0.1, y: y + 0.9, w: 1.9, h: 0.35, fontSize: 9, fontFace: "Microsoft YaHei", color: M.gray, align: "center" });
  });
  
  // Strategy
  sl.addText("\u6218\u7565\u5b9a\u4f4d\uff1a\u4ee5\u667a\u6167\u56ed\u533a\u5efa\u8bbe\u9a71\u52a8\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u80fd\u7ea7\u8dc3\u5347", { x: 0.5, y: 3.5, w: 9, h: 0.3, fontSize: 12, fontFace: "Microsoft YaHei", color: M.text, bold: true });
  var steps = [
    "\u8d4b\u80fd\u521b\u5efa\uff1a\u4e3a\u7533\u62a5\u7701\u7ea7\u7ecf\u5f00\u533a\u3001\u7eff\u8272\u5de5\u4e1a\u56ed\u533a\u63d0\u4f9b\u6570\u636e\u652f\u6491\u4e0e\u7ba1\u7406\u5de5\u5177",
    "\u8d4b\u80fd\u4f01\u4e1a\uff1a\u6253\u9020\u201c\u667a\u6539\u6570\u8f6c\u201d\u516c\u5171\u670d\u52a1\u5e73\u53f0\uff0c\u7834\u89e3\u4e2d\u5c0f\u4f01\u4e1a\u8f6c\u578b\u96be\u9898",
    "\u8d4b\u80fd\u51b3\u7b56\uff1a\u6784\u5efa\u56ed\u533a\u6570\u5b57\u5927\u8111\uff0c\u5b9e\u73b0\u4e00\u5c4f\u7edf\u89c8\u3001\u4e00\u56fe\u7edf\u7ba1"
  ];
  steps.forEach(function(s, i) {
    var y = 3.95 + i * 0.35;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.3, fill: { color: M.card }, rectRadius: 0.04 });
    dot(sl, 0.62, y + 0.07, 0.06, "1E40AF");
    sl.addText(s, { x: 0.8, y, w: 8.5, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: M.text, valign: "middle" });
  });
  badge(sl, pn, "1E40AF");
}

s1(); s2(); s3();

// ====== ROLE SECTIONS ======
// Government Role
roleSlide("\u653f\u5e9c\u9886\u5bfc\u5c42\u89c6\u89d2 \u2014 \u6570\u636e\u9a71\u52a8\u51b3\u7b56",
  ["AI\u62db\u5546\u670d\u52a1\uff1a\u57fa\u4e8e\u5927\u6570\u636e\u548cAI\u7b97\u6cd5\u7cbe\u51c6\u5339\u914d\u76ee\u6807\u4f01\u4e1a\uff0c\u63d0\u5347\u62db\u5546\u6548\u7387",
   "\u6570\u5b57\u6c99\u76d8\uff1a3D\u53ef\u89c6\u5316\u5c55\u793a\u56ed\u533a\u89c4\u5212\u4e0e\u571f\u5730\u8d44\u6e90\uff0c\u6c89\u6d78\u5f0f\u62db\u5546\u5c55\u793a",
   "\u4ea7\u4e1a\u94fe\u56fe\u8c31\uff1a\u53ef\u89c6\u5316\u5c55\u793a\u4ea7\u4e1a\u94fe\u4e0a\u4e0b\u6e38\u5173\u7cfb\uff0c\u7cbe\u51c6\u5f3a\u94fe\u8865\u94fe",
   "\u6570\u636e\u9a7e\u9a76\u8231\uff1a\u4e00\u5c4f\u7edf\u89c8\u56ed\u533a\u8fd0\u884c\u72b6\u6001\uff0c\u4e3a\u51b3\u7b56\u63d0\u4f9b\u6570\u636e\u652f\u6491",
   "\u6295\u8d44\u6982\u7b97\uff1a\u201c\u4e09\u6b65\u8d70\u201d\u521b\u5efa\u8def\u5f84\u4e0e\u9879\u76ee\u6295\u8d44\u76d1\u7763",
   "\u653f\u7b56\u5236\u5b9a\uff1a\u57fa\u4e8e\u6570\u636e\u5206\u6790\u5236\u5b9a\u60e0\u4f01\u653f\u7b56\uff0c\u7cbe\u51c6\u5bf9\u63a5\u7701\u5e02\u8d44\u91d1"],
  IMG + "/featured_supply.png", R.gov.name, R.gov.primary);

roleSlide("\u653f\u5e9c\u9886\u5bfc\u5c42\u89c6\u89d2 \u2014 \u521b\u5efa\u76d1\u7763\u4e0e\u89c4\u5212",
  ["\u4e09\u6b65\u8d70\u521b\u5efa\u8def\u5f84\uff1a\u7701\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a\u2192\u7eff\u8272\u5de5\u4e1a\u56ed\u533a\u2192\u56fd\u5bb6\u7ea7\u7ecf\u5f00\u533a",
   "\u73af\u4fdd\u76d1\u63a7\u6570\u636e\uff1a\u4e3a\u7533\u62a5\u7eff\u8272\u5de5\u4e1a\u56ed\u533a\u63d0\u4f9b\u771f\u5b9e\u3001\u53ef\u8ffd\u6eaf\u7684\u5728\u7ebf\u76d1\u6d4b\u6570\u636e\u8bc1\u636e",
   "\u80fd\u8017\u76d1\u6d4b\uff1a\u201c\u4ea9\u5747\u8bba\u82f1\u96c4\u201d\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u4e3a\u5de5\u4e1a\u7528\u5730\u8282\u7ea6\u7ea6\u675f\u63d0\u4f9b\u79d1\u5b66\u4f9d\u636e",
   "\u5b89\u5168\u76d1\u63a7\uff1a\u5929\u773c+\u54e8\u5175+\u6167\u773c\u5168\u57df\u8986\u76d6\uff0c\u4fdd\u969c\u56ed\u533a\u5b89\u5168\u8fd0\u884c",
   "\u6570\u636e\u4e2d\u5fc3\uff1a\u6c47\u805a\u56ed\u533a\u5168\u91cf\u6570\u636e\uff0c\u6784\u5efa\u7edf\u4e00\u6570\u636e\u6e56\uff0c\u652f\u6491\u7edf\u8ba1\u4e0e\u5206\u6790",
   "\u653f\u7b56\u5956\u52b1\uff1a\u6210\u529f\u521b\u5efa\u7eff\u8272\u56ed\u533a\u53ef\u83b7\u767e\u4e07\u7ea7\u4e00\u6b21\u6027\u5956\u52b1"],
  IMG + "/info_publish.png", R.gov.name, R.gov.primary);

// Management Role
roleSlide("\u56ed\u533a\u7ba1\u7406\u5c42\u89c6\u89d2 \u2014 \u8fd0\u8425\u7ba1\u7406\u4e0e\u76d1\u63a7",
  ["\u5168\u57df\u5b89\u5168\u76d1\u63a7\uff1a\u9053\u8def\u4ea4\u901a\u76d1\u63a7(\u5929\u773c)+\u5730\u707e\u76d1\u6d4b(\u54e8\u5175)+\u73af\u4fdd\u76d1\u63a7(\u6167\u773c)",
   "\u89c6\u9891\u76d1\u63a7\u603b\u89c8\uff1a\u7edf\u4e00\u89c6\u9891\u5e73\u53f0\uff0c16\u5bab\u683c\u667a\u80fd\u6392\u5217\uff0c\u8986\u76d6\u56ed\u533a\u5173\u952e\u533a\u57df",
   "\u5e94\u6025\u6307\u6325\u8054\u52a8\uff1a\u4e00\u952e\u8c03\u5ea6\u3001\u9884\u6848\u7ba1\u7406\u3001\u5b9e\u65f6\u901a\u8baf\uff0c\u54cd\u5e94\u6548\u7387\u63d0\u534760%",
   "\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3\uff1a\u5927\u5c4f\u6307\u6325+\u56ed\u5c0f\u4e8c\u56e2\u961f+3\u5e74\u957f\u6548\u670d\u52a1",
   "\u65e5\u5e38\u8fd0\u8425\uff1a\u6570\u5b57\u5316\u5ba1\u6279\u3001\u65e5\u5e38\u5de1\u68c0\u3001\u8d44\u4ea7\u7ba1\u7406\u3001\u6863\u6848\u7ba1\u7406\u5168\u6d41\u7a0b\u5728\u7ebf\u5316"],
  IMG + "/info_publish.png", R.mgmt.name, R.mgmt.primary);

roleSlide("\u56ed\u533a\u7ba1\u7406\u5c42\u89c6\u89d2 \u2014 \u80fd\u6548\u4e0e\u6570\u636e\u7ba1\u63a7",
  ["\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b\uff1a\u667a\u80fd\u8ba1\u91cf\u8868\u8ba1\u5b9e\u65f6\u91c7\u96c6\u80fd\u8017\u4e0e\u4ea7\u51fa\u6570\u636e",
   "\u4ea9\u5747\u8bba\u82f1\u96c4\uff1a\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u8d44\u6e90\u8981\u7d20\u5dee\u522b\u5316\u914d\u7f6e\uff0c\u5012\u903c\u4f01\u4e1a\u63d0\u8d28\u589e\u6548",
   "\u6570\u636e\u4e2d\u5fc3\uff1a12\u5927\u4e09\u65b9\u7cfb\u7edf\u6570\u636e\u4e92\u8054\u4e92\u901a\uff0c\u7edf\u4e00\u6570\u636e\u6807\u51c6",
   "\u5f02\u5e38\u9884\u8b66\uff1a\u7528\u80fd\u5f02\u5e38\u81ea\u52a8\u544a\u8b66\uff0c\u9632\u6b62\u80fd\u6e90\u6d6a\u8d39\uff0c\u78b3\u6392\u7ba1\u7406\u8ffd\u8e2a",
   "\u7efc\u5408\u7ba1\u7406\u6548\u7387\u63d0\u534730%\uff0c\u4ece\u88ab\u52a8\u54cd\u5e94\u8f6c\u5411\u4e3b\u52a8\u9884\u8b66"],
  IMG + "/parking_detail.png", R.mgmt.name, R.mgmt.primary);

// Enterprise Role
roleSlide("\u5165\u9a7b\u4f01\u4e1a\u89c6\u89d2 \u2014 \u4f01\u4e1a\u670d\u52a1\u4e0e\u8bc9\u6c42",
  ["\u8bc9\u6c42\u76f4\u8fbe\uff08\u968f\u624b\u62cd\uff09\uff1a\u6587\u5b57+\u56fe\u7247+\u5b9a\u4f4d\uff0c\u5feb\u901f\u63d0\u4ea4\u8bc9\u6c42\uff0c2\u5c0f\u65f6\u54cd\u5e94",
   "\u670d\u52a1\u76f4\u8fbe\uff1a\u5929\u7136\u6c14/\u6c34/\u7535\u8d39\u5728\u7ebf\u7f34\u8d39\uff0c\u4e00\u7ad9\u5f0f\u670d\u52a1",
   "\u56ed\u4f01\u4ea4\u4e92\uff1a\u5728\u7ebf\u6c9f\u901a\u3001\u901a\u77e5\u516c\u544a\u3001\u653f\u7b56\u63a8\u9001",
   "\u4f9b\u9700\u5bf9\u63a5\uff1a\u7269\u6d41\u4f9b\u9700\u3001\u7279\u8272\u4f9b\u9700\u667a\u80fd\u5339\u914d",
   "\u4f01\u4e1a\u5165\u9a7b\uff1a\u7ebf\u4e0a\u7533\u8bf7\u3001\u8bc1\u7167\u7ba1\u7406\u3001\u4fe1\u7528\u8bc4\u4ef7"],
  IMG + "/info_publish.png", R.ent.name, R.ent.primary);

roleSlide("\u5165\u9a7b\u4f01\u4e1a\u89c6\u89d2 \u2014 \u8d44\u6e90\u5bf9\u63a5\u4e0e\u653f\u7b56",
  ["\u7269\u6d41\u4f9b\u9700\uff1a\u56ed\u533a\u8d27\u6e90+\u8fd4\u7a0b\u7a7a\u8f66+\u8f66\u8f86\u8d44\u6e90\uff0c\u6bcf\u5f20\u5361\u7247\u5e26\u8054\u7cfb\u65b9\u5f0f",
   "\u7279\u8272\u4f9b\u9700\uff1a\u9752\u5ddd\u5c71\u73cd\u7b499\u79cd\u7279\u8272\u4ea7\u54c1\u5c55\u793a\u4e0e\u4f9b\u9700\u5bf9\u63a5",
   "\u653f\u7b56\u63a8\u9001\uff1a\u60e0\u4f01\u653f\u7b56\u7cbe\u51c6\u5339\u914d\u4e0e\u4e3b\u52a8\u63a8\u9001\uff0c\u8865\u8d34/\u51cf\u514d/\u4eba\u624d\u7b49",
   "\u667a\u6539\u6570\u8f6c\u8d4b\u80fd\uff1a\u5bf9\u63a5\u7701\u5e02\u6bcf\u5e74\u8d8520\u4ebf\u5143\u4e13\u9879\u8d44\u91d1\uff0c\u964d\u4f4e\u8f6c\u578b\u6210\u672c",
   "\u505c\u8f66\u670d\u52a1\uff1a\u9ad8\u5fb7\u5730\u56fe\u5bfc\u822a+\u5b9e\u65f6\u4f59\u4f4d+\u5728\u7ebf\u7f34\u8d39"],
  IMG + "/logistics_supply.png", R.ent.name, R.ent.primary);

roleSlide("\u5165\u9a7b\u4f01\u4e1a\u89c6\u89d2 \u2014 \u573a\u9986\u4e0e\u9910\u996e",
  ["\u573a\u9986\u670d\u52a1\uff1a\u591a\u529f\u80fd\u4f1a\u8bae\u5385/\u5c55\u5385/\u4f53\u80b2\u9986/\u57f9\u8bad\u6559\u5ba4\u7b498\u4e2a\u573a\u9986",
   "\u56ed\u533a\u9910\u996e\uff1a\u6d41\u52a8\u5361\u7247\u5c55\u793a6\u5bb6\u9910\u5385\uff0c20+\u9053\u83dc\u54c1\u4e00\u89c8",
   "\u7528\u5de5\u4fe1\u606f\uff1a\u62db\u8058+\u5c97\u4f4d\u5339\u914d\uff0c\u652f\u6301\u5206\u7c7b\u7b5b\u9009",
   "\u56ed\u533a\u6d3b\u52a8\uff1a\u6587\u4f53/\u57f9\u8bad/\u516c\u76ca/\u4ea4\u6d41\u7b498\u4e2a\u6d3b\u52a8\u53c2\u4e0e",
   "\u56ed\u533a\u516c\u544a\uff1a\u5b9e\u65f6\u901a\u77e5\u4e0e\u4fe1\u606f\u53d1\u5e03"],
  IMG + "/venue_service.png", R.ent.name, R.ent.primary);

// Public Role
roleSlide("\u516c\u4f17\u7528\u6237\u89c6\u89d2 \u2014 \u4fe1\u606f\u670d\u52a1\u4e0e\u53c2\u4e0e",
  ["\u56ed\u533a\u516c\u544a\uff1a\u5b9e\u65f6\u67e5\u770b\u56ed\u533a\u901a\u77e5\uff08\u65bd\u5de5\u3001\u68c0\u67e5\u3001\u7ef4\u62a4\u7b49\uff09",
   "\u56ed\u533a\u6d3b\u52a8\uff1a\u67e5\u770b\u5e76\u53c2\u4e0e\u6587\u4f53/\u516c\u76ca\u7b49\u6d3b\u52a8\uff0c\u8ddf\u8e2a\u6d3b\u52a8\u72b6\u6001",
   "\u7279\u8272\u4f9b\u9700\uff1a\u6d4f\u89c8\u9752\u5ddd\u5c71\u73cd\u7b49\u7279\u8272\u4ea7\u54c1\uff0c\u76f4\u63a5\u8054\u7cfb\u4f9b\u5e94\u5546",
   "\u7528\u5de5\u4fe1\u606f\uff1a\u67e5\u770b\u56ed\u533a\u4f01\u4e1a\u62db\u8058\u4fe1\u606f\uff0c\u5e97\u4f4d\u5339\u914d",
   "\u56ed\u533a\u9910\u996e\uff1a\u6d4f\u89c8\u9910\u5385\u4fe1\u606f\u548c\u83dc\u54c1\uff0c\u4e86\u89e3\u4ef7\u683c\u4e0e\u4f4d\u7f6e"],
  IMG + "/notice_detail.png", R.pub.name, R.pub.primary);

roleSlide("\u516c\u4f17\u7528\u6237\u89c6\u89d2 \u2014 \u751f\u6d3b\u670d\u52a1\u4e0e\u4e92\u52a8",
  ["\u56ed\u533a\u9910\u996e\uff1a6\u5bb6\u9910\u5385\u4fe1\u606f+\u4eca\u65e5\u83dc\u5355+\u4ef7\u683c+\u5206\u7c7b\u7b5b\u9009",
   "\u505c\u8f66\u670d\u52a1\uff1a\u5730\u56fe\u67e5\u770b\u505c\u8f66\u4f4d\u5b9e\u65f6\u4f59\u4f4d\uff0c\u5728\u7ebf\u7f34\u8d39",
   "\u573a\u9986\u670d\u52a1\uff1a\u6d4f\u89c8\u56ed\u533a\u573a\u9986\u4fe1\u606f\uff0c\u4e86\u89e3\u5f00\u653e\u65f6\u95f4\u4e0e\u7ea6\u70b9",
   "\u56ed\u533a\u653f\u7b56\uff1a\u67e5\u770b\u516c\u5f00\u653f\u7b56\u4fe1\u606f\uff0c\u4e86\u89e3\u56ed\u533a\u53d1\u5c55\u52a8\u6001",
   "\u4fe1\u606f\u516c\u5f00\u900f\u660e\uff0c\u589e\u5f3a\u516c\u4f17\u5f52\u5c5e\u611f\u4e0e\u83b7\u5f97\u611f"],
  IMG + "/catering.png", R.pub.name, R.pub.primary);

// === ROLE MATRIX ===
function sRoleMatrix() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.bg };
  topBar(sl, "F59E0B");
  for (var i = 0; i < 8; i++) dot(sl, 0.3 + i * 0.2, 0.07, 0.035, "F59E0B");
  sl.addText("\u56db\u7c7b\u89d2\u8272\u529f\u80fd\u77e9\u9635\u603b\u89c8", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 20, fontFace: "Microsoft YaHei", color: M.text, bold: true });
  sl.addText("\u6bcf\u4e2a\u89d2\u8272\u5728\u5404\u529f\u80fd\u6a21\u5757\u4e2d\u7684\u53c2\u4e0e\u5ea6\u4e00\u89c8", { x: 0.5, y: 0.65, w: 8, h: 0.25, fontSize: 10, fontFace: "Microsoft YaHei", color: M.gray });

  var cols = ["\u529f\u80fd\u6a21\u5757", "\u653f\u5e9c\u9886\u5bfc\u5c42", "\u56ed\u533a\u7ba1\u7406\u5c42", "\u5165\u9a7b\u4f01\u4e1a", "\u516c\u4f17\u7528\u6237"];
  var hColors = [M.text, R.gov.primary, R.mgmt.primary, R.ent.primary, R.pub.primary];
  cols.forEach(function(c, i) {
    var x = 0.4 + (i === 0 ? 0 : 1.5 + (i-1)*1.85);
    var w = i === 0 ? 1.3 : 1.85;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.05, w, h: 0.3, fill: { color: hColors[i] }, rectRadius: 0.03 });
    sl.addText(c, { x, y: 1.05, w, h: 0.3, fontSize: 9, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  });

  var rows = [
    ["AI\u62db\u5546\u00b7\u6570\u5b57\u6c99\u76d8", "\u25cf\u51b3\u7b56\u652f\u6491", "\u25cb\u6570\u636e\u7ba1\u7406", "\u25cf\u4fe1\u606f\u67e5\u770b", "\u2014"],
    ["\u5168\u57df\u5b89\u5168\u76d1\u63a7", "\u25cf\u76d1\u7ba1\u603b\u89c8", "\u25cf\u65e5\u5e38\u8fd0\u7ef4", "\u25cb\u5b89\u5168\u670d\u52a1", "\u25cf\u5b89\u5168\u4fe1\u606f"],
    ["\u4f01\u4e1a\u670d\u52a1\u00b7\u8bc9\u6c42", "\u25cf\u6548\u80fd\u5206\u6790", "\u25cf\u6d41\u7a0b\u7ba1\u7406", "\u25cf\u6838\u5fc3\u4f7f\u7528", "\u25cb\u6295\u8bc9\u5efa\u8bae"],
    ["\u7269\u6d41\u4f9b\u9700", "\u25cb\u4ea7\u4e1a\u6570\u636e", "\u25cf\u8fd0\u8425\u8c03\u5ea6", "\u25cf\u53d1\u5e03\u5bf9\u63a5", "\u25cb\u4fe1\u606f\u67e5\u770b"],
    ["\u7279\u8272\u4f9b\u9700", "\u25cf\u7ecf\u6d4e\u5c55\u793a", "\u25cb\u6570\u636e\u7edf\u8ba1", "\u25cf\u4ea7\u54c1\u53d1\u5e03", "\u25cf\u6d4f\u89c8\u8d2d\u4e70"],
    ["\u56ed\u533a\u653f\u7b56", "\u25cf\u5236\u5b9a\u53d1\u5e03", "\u25cf\u6267\u884c\u76d1\u7763", "\u25cf\u653f\u7b56\u7533\u62a5", "\u25cb\u653f\u7b56\u67e5\u9605"],
    ["\u9910\u996e\u00b7\u505c\u8f66\u00b7\u573a\u9986", "\u2014", "\u25cf\u8fd0\u8425\u7ba1\u7406", "\u25cf\u5728\u7ebf\u4f7f\u7528", "\u25cf\u9884\u7ea6\u9884\u8ba2"],
    ["\u6d3b\u52a8\u00b7\u516c\u544a\u00b7\u7528\u5de5", "\u25cf\u5b8f\u89c2\u6570\u636e", "\u25cf\u4fe1\u606f\u5ba1\u6838", "\u25cf\u53c2\u4e0e\u53d1\u5e03", "\u25cf\u83b7\u53d6\u4fe1\u606f"],
    ["\u6570\u636e\u4e2d\u5fc3\u00b7\u80fd\u6548", "\u25cf\u51b3\u7b56\u4f9d\u636e", "\u25cf\u76d1\u6d4b\u9884\u8b66", "\u25cb\u7528\u80fd\u7ba1\u7406", "\u2014"],
    ["\u6295\u8d44\u6982\u7b97", "\u25cf\u8d22\u653f\u5ba1\u6279", "\u25cb\u9884\u7b97\u53c2\u8003", "\u2014", "\u2014"]
  ];
  
  rows.forEach(function(row, ri) {
    var y = 1.4 + ri * 0.35;
    var bg = ri % 2 === 0 ? M.card : M.bg;
    row.forEach(function(cell, ci) {
      var x = 0.4 + (ci === 0 ? 0 : 1.5 + (ci-1)*1.85);
      var w = ci === 0 ? 1.3 : 1.85;
      sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.35, fill: { color: bg } });
      if (ci > 0 && cell !== "\u2014") {
        var isMain = cell.indexOf("\u25cf") === 0;
        dot(sl, x + 0.12, y + 0.1, 0.07, isMain ? "1E40AF" : M.gray);
      }
      sl.addText(cell, { x: x + (ci > 0 ? 0.3 : 0.08), y, w: w - (ci > 0 ? 0.4 : 0.12), h: 0.35, fontSize: ci === 0 ? 9 : 8, fontFace: "Microsoft YaHei", color: ci === 0 ? M.text : M.gray, valign: "middle" });
    });
  });
  
  // Legend
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.3, fill: { color: M.light }, rectRadius: 0.04 });
  dot(sl, 0.65, 5.06, 0.08, "1E40AF"); 
  sl.addText("\u25cf \u4e3b\u8981\u529f\u80fd", { x: 0.85, y: 5.0, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Microsoft YaHei", color: M.text, valign: "middle" });
  dot(sl, 2.4, 5.06, 0.08, M.gray); 
  sl.addText("\u25cb \u8f85\u52a9\u529f\u80fd", { x: 2.6, y: 5.0, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Microsoft YaHei", color: M.text, valign: "middle" });
  sl.addText("\u2014 \u65e0\u76f4\u63a5\u5173\u8054", { x: 4.2, y: 5.0, w: 2, h: 0.3, fontSize: 9, fontFace: "Microsoft YaHei", color: M.gray, valign: "middle" });
  badge(sl, pn, "F59E0B");
}

sRoleMatrix();

// === Investment & Benefits ===
function sInvest() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.bg };
  topBar(sl, "F59E0B");
  for (var i = 0; i < 8; i++) dot(sl, 0.3 + i * 0.2, 0.07, 0.035, "F59E0B");
  sl.addText("\u6295\u8d44\u4e0e\u6548\u76ca\u5206\u6790", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 22, fontFace: "Microsoft YaHei", color: M.text, bold: true });
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.5, y: 0.8, w: 5, h: 1.3, fill: { color: "0F172A" }, rectRadius: 0.1, shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.1 } });
  sl.addText("\u603b\u6295\u8d44\u6982\u7b97", { x: 2.5, y: 0.9, w: 5, h: 0.25, fontSize: 12, fontFace: "Microsoft YaHei", color: M.light, align: "center" });
  sl.addText("\u00a54,178,800.00", { x: 2.5, y: 1.15, w: 5, h: 0.7, fontSize: 30, fontFace: "Arial", color: "F59E0B", bold: true, align: "center" });
  var details = ["\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7","\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7","\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546","\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b","\u667a\u6539\u6570\u8f6c\u8d4b\u80fd\u5e73\u53f0","\u4e00\u4f53\u5316\u6570\u5b57\u5e73\u53f0","\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3"];
  details.forEach(function(d, i) {
    var y = 2.4 + i * 0.3;
    dot(sl, 1.0, y + 0.05, 0.06, "F59E0B");
    sl.addText(d, { x: 1.2, y, w: 4, h: 0.25, fontSize: 10, fontFace: "Microsoft YaHei", color: M.text });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.0, y: y + 0.04, w: 3.2, h: 0.12, fill: { color: M.light }, rectRadius: 0.06 });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.0, y: y + 0.04, w: 3.2 * (i+1)/7, h: 0.12, fill: { color: "F59E0B" }, rectRadius: 0.06 });
  });
  
  // Benefits
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.5, w: 9, h: 0.03, fill: { color: M.light } });
  var metrics = [["30%", "\u7ba1\u7406\u6548\u7387\u63d0\u5347", "1E40AF"], ["60%", "\u5e94\u6025\u54cd\u5e94\u63d0\u5347", "059669"], ["\u00a520\u4ebf+", "\u4e13\u9879\u8d44\u91d1\u5bf9\u63a5", "D97706"]];
  metrics.forEach(function(m, i) {
    var x = 0.8 + i * 3.0;
    sl.addText(m[0], { x, y: 4.6, w: 2.2, h: 0.35, fontSize: 22, fontFace: "Arial", color: m[2], bold: true, align: "center" });
    sl.addText(m[1], { x, y: 4.9, w: 2.2, h: 0.25, fontSize: 10, fontFace: "Microsoft YaHei", color: M.text, bold: true, align: "center" });
  });
  badge(sl, pn, "F59E0B");
}

sInvest();

// === Future ===
function sFuture() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.bg };
  topBar(sl, "1E40AF");
  for (var i = 0; i < 8; i++) dot(sl, 0.3 + i * 0.2, 0.07, 0.035, "1E40AF");
  sl.addText("\u672a\u6765\u5c55\u671b", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 22, fontFace: "Microsoft YaHei", color: M.text, bold: true });
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 0.5, y: 0.65, w: 8, h: 0.25, fontSize: 11, fontFace: "Microsoft YaHei", color: M.gray });
  var items = [
    "AIoT\u6df1\u5ea6\u878d\u5408\uff1a\u5f15\u5165\u7269\u8054\u7f51\u4f20\u611f\u8bbe\u5907\uff0c\u5b9e\u73b0\u4e07\u7269\u4e92\u8054\u3001\u5168\u57df\u611f\u77e5",
    "\u6570\u5b57\u5b6a\u751f\u56ed\u533a\uff1a1:1\u4e09\u7ef4\u5efa\u6a21\uff0c\u5168\u8981\u7d20\u3001\u5168\u65f6\u57df\u53ef\u89c6\u5316\u8fd0\u8425\u7ba1\u7406",
    "AI Agent\u667a\u80fd\u5316\uff1a\u667a\u80fd\u5ba2\u670d\u3001\u667a\u80fd\u8c03\u5ea6\u3001\u667a\u80fd\u9884\u8b66",
    "\u4ea7\u4e1a\u5927\u8111\uff1a\u57fa\u4e8e\u5927\u6570\u636e\u4e0eAI\u7b97\u6cd5\u8f85\u52a9\u4ea7\u4e1a\u62db\u5546\u51b3\u7b56",
    "\u8de8\u56ed\u533a\u534f\u540c\uff1a\u5bf9\u63a5\u7701\u5e02\u667a\u6167\u5e73\u53f0\uff0c\u5b9e\u73b0\u6570\u636e\u5171\u4eab\u4e0e\u4e1a\u52a1\u534f\u540c"
  ];
  items.forEach(function(item, i) {
    var y = 1.1 + i * 0.6;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.48, fill: { color: M.card }, rectRadius: 0.06, shadow: { type: "outer", blur: 2, offset: 1, color: "000000", opacity: 0.03 } });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y + 0.08, w: 0.05, h: 0.32, fill: { color: "1E40AF" } });
    sl.addText(item, { x: 0.75, y, w: 8.5, h: 0.48, fontSize: 11, fontFace: "Microsoft YaHei", color: M.text, valign: "middle" });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.3, w: 9, h: 0.02, fill: { color: M.light } });
  sl.addText("\u9a71\u52a8\u56ed\u533a\u6cbf\u7740\u201c\u7701\u7ea7\u5f00\u53d1\u533a\u2192\u7eff\u8272\u56ed\u533a\u2192\u56fd\u5bb6\u7ea7\u7ecf\u5f00\u533a\u57f9\u80b2\u5bf9\u8c61\u201d\u7684\u822a\u5411\u5168\u901f\u524d\u8fdb", { x: 0.5, y: 4.5, w: 9, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: M.gray, italic: true, align: "center" });
  badge(sl, pn, "1E40AF");
}

sFuture();

// === Thank You ===
function sThank() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: M.darkBg };
  topBar(sl, "F59E0B");
  sl.addShape(pres.shapes.OVAL, { x: 2.5, y: 0.2, w: 5, h: 5, fill: { color: "1E3A5F", transparency: 60 } });
  sl.addText("\u611f\u8c22\u804a\u542c", { x: 1, y: 1.5, w: 8, h: 1.0, fontSize: 36, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center" });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4, y: 2.5, w: 2, h: 0.04, fill: { color: "F59E0B" } });
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 1, y: 2.7, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: "F59E0B", align: "center" });
  ["\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u7ba1\u59d4\u4f1a", "\u6280\u672f\u652f\u6301\uff1a\u9752\u8fc5\u79d1\u6280", "\u5730\u5740\uff1a\u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf\u7af9\u56ed\u9547"].forEach(function(item, i) {
    sl.addText(item, { x: 1, y: 3.5 + i * 0.4, w: 8, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: "94A3B8", align: "center" });
  });
}

sThank();

pres.writeFile({ fileName: path.join(__dirname, "output", "zhuangzishang_role_ppt.pptx") })
  .then(function() { console.log("PPT OK: " + pn + " slides"); })
  .catch(function(e) { console.error("ERR:", e); });
