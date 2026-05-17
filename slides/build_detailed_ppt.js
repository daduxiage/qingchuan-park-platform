const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "\u9752\u8fc5\u79d1\u6280";
pres.title = "\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u667a\u6167\u5e73\u53f0\u6c47\u62a5";
const IMG = "/Users/wangcaihua/WorkBuddy/20260427142154/slides/imgs";

// ====== COLOR PALETTE ======
const C = {
  dark: "0F172A", bg: "F1F5F9", card: "FFFFFF",
  text: "1E293B", gray: "64748B", light: "CBD5E1",
  // Role primaries
  gov: "2563EB", govLite: "DBEAFE",
  mgmt: "059669", mgmtLite: "D1FAE5",
  ent: "D97706", entLite: "FEF3C7",
  pub: "7C3AED", pubLite: "EDE9FE",
  accent: "F59E0B", accentLite: "FDE68A"
};

const FONT = "Microsoft YaHei";
let pn = 0;

// ====== HELPERS ======
function badge(sl, n, c) {
  sl.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.38, h: 0.38, fill: { color: c || "2563EB" } });
  sl.addText(String(n), { x: 9.3, y: 5.1, w: 0.38, h: 0.38, fontSize: 10, fontFace: FONT, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
}
function topBar(sl, c) { sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: c || "2563EB" } }); }
function card(sl, x, y, w, h, fill) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill || "FFFFFF" }, rectRadius: 0.08,
    shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.06 } });
}
function rbox(sl, x, y, w, h, c, r) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: c }, rectRadius: r || 0.04 });
}
function sbox(sl, x, y, w, h, c) {
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: c } });
}

// ====== SLIDE 1: COVER ======
function sCover() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.dark };
  
  // Large decorative circles
  sl.addShape(pres.shapes.OVAL, { x: 6.5, y: -2.5, w: 6.5, h: 6.5, fill: { color: "1E3A5F", transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -1.5, y: 2, w: 5, h: 5, fill: { color: "1E3A5F", transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: 8, y: 2.5, w: 3, h: 3, fill: { color: "2563EB", transparency: 80 } });
  
  // Accent line
  sbox(sl, 0.5, 0.5, 0.07, 3.2, C.accent);
  
  // Title block
  sl.addText("\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a", { x: 0.8, y: 0.5, w: 8, h: 0.5, fontSize: 20, fontFace: FONT, color: C.accent });
  sl.addText("\u667a\u6167\u5e73\u53f0\u6c47\u62a5", { x: 0.8, y: 0.95, w: 8, h: 1.2, fontSize: 42, fontFace: FONT, color: "FFFFFF", bold: true });
  sl.addText("\u56db\u7c7b\u89d2\u8272 \u00b7 \u4e00\u4f53\u5316\u667a\u6167\u8fd0\u8425 \u00b7 \u6570\u636e\u9a71\u52a8\u51b3\u7b56", { x: 0.8, y: 2.0, w: 7, h: 0.4, fontSize: 13, fontFace: FONT, color: "94A3B8" });
  
  // Role tags
  var roles = [["\u653f\u5e9c\u9886\u5bfc\u5c42", C.gov], ["\u56ed\u533a\u7ba1\u7406\u5c42", C.mgmt], ["\u5165\u9a7b\u4f01\u4e1a", C.ent], ["\u516c\u4f17\u7528\u6237", C.pub]];
  roles.forEach(function(r, i) {
    rbox(sl, 0.8 + i*2.2, 2.7, 2.0, 0.38, r[1], 0.06);
    sl.addText(r[0], { x: 0.8 + i*2.2, y: 2.7, w: 2.0, h: 0.38, fontSize: 10, fontFace: FONT, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  });
  
  // Decorative geometric shapes
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 7.5, y: 0.15, w: 0.6, h: 0.3, fill: { color: C.accent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.3, y: 0.15, w: 0.6, h: 0.3, fill: { color: C.gov, transparency: 30 } });
  
  // Bottom info
  sl.addText("\u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf \u00b7 \u7af9\u56ed\u9547\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a", { x: 0.6, y: 3.4, w: 6, h: 0.3, fontSize: 10, fontFace: FONT, color: "94A3B8" });
  sl.addText("2025\u5e74\u524d\u4e09\u5b63\u5ea6 GDP 57.68\u4ebf\u5143 \u00b7 \u540c\u6bd4\u589e\u957f15% \u00b7 2026\u5e745\u6708", { x: 0.6, y: 3.7, w: 7, h: 0.25, fontSize: 9, fontFace: FONT, color: "64748B" });
  
  // AI Image if exists
  if (fs.existsSync(IMG + "/ai_smart_park_dashboard.png")) {
    sl.addImage({ path: IMG + "/ai_smart_park_dashboard.png", x: 5.5, y: 3.3, w: 4.2, h: 2.1, sizing: { type: "cover", w: 4.2, h: 2.1 } });
  }
}

// ====== SLIDE 2: TOC ======
function sTOC() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.gov);
  
  sl.addText("\u6c47\u62a5\u76ee\u5f55", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 24, fontFace: FONT, color: C.text, bold: true });
  
  var sections = [
    ["\u9879\u76ee\u80cc\u666f\u4e0e\u5e73\u53f0\u67b6\u6784", C.gov],
    ["\u89d2\u8272\u529f\u80fd\u77e9\u9635\u603b\u89c8", C.mgmt],
    ["\u653f\u5e9c\u9886\u5bfc\u5c42\u2e2f\u5168\u666f\u6570\u636e\u9a71\u52a8", C.gov],
    ["\u56ed\u533a\u7ba1\u7406\u5c42\u2e2f\u667a\u6167\u8fd0\u8425\u4e0e\u76d1\u63a7", C.mgmt],
    ["\u5165\u9a7b\u4f01\u4e1a\u2e2f\u5168\u6d41\u7a0b\u4f01\u4e1a\u670d\u52a1", C.ent],
    ["\u516c\u4f17\u7528\u6237\u2e2f\u4fbf\u6c11\u670d\u52a1\u4e0e\u53c2\u4e0e", C.pub],
    ["\u5b8c\u6574\u529f\u80fd\u77e9\u9635\u4e00\u89c8", C.accent],
    ["\u6295\u8d44\u4e0e\u6548\u76ca\u5206\u6790\u00b7\u672a\u6765\u5c55\u671b", C.gov]
  ];
  
  sections.forEach(function(s, i) {
    var y = 0.9 + i * 0.5;
    rbox(sl, 0.5, y, 9, 0.42, C.card, 0.06);
    sbox(sl, 0.5, y, 0.04, 0.42, s[1]);
    rbox(sl, 0.65, y + 0.07, 0.28, 0.28, s[1], 0.04);
    sl.addText(String(i+1), { x: 0.65, y: y + 0.07, w: 0.28, h: 0.28, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(s[0], { x: 1.1, y: y, w: 7.5, h: 0.42, fontSize: 13, fontFace: FONT, color: C.text, valign: "middle" });
    sl.addShape(pres.shapes.PARALLELOGRAM, { x: 9.0, y: y + 0.1, w: 0.4, h: 0.22, fill: { color: s[1], transparency: 20 } });
  });
  
  badge(sl, pn, C.gov);
}

// ====== SLIDE 3: PROJECT BACKGROUND ======
function sBackground() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.gov);
  
  sl.addText("\u9879\u76ee\u80cc\u666f", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 22, fontFace: FONT, color: C.text, bold: true });
  
  // Location highlight
  rbox(sl, 0.5, 0.75, 9, 0.5, C.govLite, 0.06);
  sl.addText("\u9752\u5ddd\u53bf\u4f4d\u4e8e\u56db\u5ddd\u7701\u5317\u90e8\u8fb9\u7f18\uff0c\u5ddd\u7518\u9655\u4e09\u7701\u7ed3\u5408\u90e8\uff0c\u7d20\u6709\u201c\u91d1\u4e09\u89d2\u201d\u4e4b\u79f0\u3002\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u89c4\u5212\u9762\u79ef2.1\u33a1\uff0c\u5df2\u5f00\u53d11.6\u33a1\u3002",
    { x: 0.7, y: 0.78, w: 8.6, h: 0.44, fontSize: 11, fontFace: FONT, color: C.text, valign: "middle" });
  
  // 4 metric cards - 2x2 grid
  var metrics = [
    ["57.68\u4ebf", "2025\u524d\u4e09\u5b63\u5ea6GDP", "\u5168\u53bf\u589e\u957f15%", C.gov],
    ["137\u4ebf", "\u53bf\u5c5e\u56fd\u4f01\u8d44\u4ea7\u603b\u989d", "\u8425\u4e1a\u6536\u51653.44\u4ebf", C.mgmt],
    ["45\u5bb6", "\u56fd\u5bb6\u79d1\u6280\u578b\u4e2d\u5c0f\u4f01\u4e1a", "\u65b0\u589e\u9ad8\u65b0\u6280\u672f\u4f018\u5bb6", C.ent],
    ["47\u5bb6", "\u56ed\u533a\u5165\u9a7b\u4f01\u4e1a", "\u4ece\u4e1a\u4eba\u54583,000+", C.pub]
  ];
  metrics.forEach(function(m, i) {
    var x = 0.5 + (i%2)*4.6, y = 1.5 + Math.floor(i/2)*1.5;
    // Card background with accent top strip
    rbox(sl, x, y, 4.3, 1.3, C.card, 0.1);
    sbox(sl, x + 0.1, y + 0.1, 0.06, 1.1, m[3]);
    sl.addText(m[0], { x: x + 0.3, y: y + 0.1, w: 3.8, h: 0.5, fontSize: 28, fontFace: "Arial", color: m[3], bold: true });
    sl.addText(m[1], { x: x + 0.3, y: y + 0.6, w: 3.8, h: 0.25, fontSize: 11, fontFace: FONT, color: C.text, bold: true });
    sl.addText(m[2], { x: x + 0.3, y: y + 0.85, w: 3.8, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });
  });
  
  // Bottom: strategy
  rbox(sl, 0.5, 4.7, 9, 0.55, "0F172A", 0.08);
  sl.addText("\u6218\u7565\u76ee\u6807\uff1a\u4ee5\u667a\u6167\u56ed\u533a\u5efa\u8bbe\u4e3a\u6346\u5e26\uff0c\u63a8\u52a8\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u80fd\u7ea7\u8dc3\u5347\uff0c\u52a9\u529b\u521b\u5efa\u7701\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a\u3001\u7eff\u8272\u5de5\u4e1a\u56ed\u533a",
    { x: 0.8, y: 4.72, w: 8.5, h: 0.5, fontSize: 11, fontFace: FONT, color: "F59E0B", valign: "middle" });
  
  badge(sl, pn, C.gov);
}

// ====== SLIDE 4: PLATFORM ARCHITECTURE ======
function sArchitecture() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.accent);
  
  sl.addText("\u5e73\u53f0\u67b6\u6784\u6982\u89c8", { x: 0.5, y: 0.15, w: 8, h: 0.5, fontSize: 22, fontFace: FONT, color: C.text, bold: true });
  sl.addText("\u4e94\u5c42\u67b6\u6784 \u00b7 \u4e00\u4f53\u5316\u667a\u6167\u5e73\u53f0", { x: 0.5, y: 0.55, w: 8, h: 0.25, fontSize: 10, fontFace: FONT, color: C.gray });
  
  // Architecture layers - vertical stack
  var layers = [
    ["C\u7aef\u5c55\u793a\u5c42", "\u5927\u5c4f\u6570\u636e\u9a7e\u9a76\u8231 | \u56ed\u533a\u5b98\u7f51 | \u5411\u5bfc\u9875 | H5\u79fb\u52a8\u7aef | \u5fae\u4fe1\u5c0f\u7a0b\u5e8f", C.gov],
    ["\u5e94\u7528\u670d\u52a1\u5c42", "AI\u62db\u5546 | \u5b89\u5168\u76d1\u63a7 | \u4f01\u4e1a\u670d\u52a1 | \u8fd0\u8425\u7ba1\u7406 | \u80fd\u6548\u76d1\u6d4b | \u6570\u5b57\u5316\u529e\u516c", C.mgmt],
    ["\u6570\u636e\u670d\u52a1\u5c42", "\u6570\u636e\u6e56 | \u6570\u636e\u5206\u6790 | \u62a5\u8868\u5f15\u64ce | \u5b9e\u65f6\u8ba1\u7b97 | AI\u7b97\u6cd5 | GIS\u5e73\u53f0", C.ent],
    ["\u8bbe\u65bd\u5c42", "\u89c6\u9891\u76d1\u63a7 | \u4f20\u611f\u5668 | \u667a\u80fd\u8868\u8ba1 | GPS\u5b9a\u4f4d | \u73af\u5883\u76d1\u6d4b", C.pub],
    ["\u6570\u636e\u5bf9\u63a5\u5c42", "\u5929\u7136\u6c14 | \u7535\u529b | \u81ea\u6765\u6c34 | \u7efc\u5408\u6267\u6cd5 | \u73af\u4fdd | \u5e94\u6025 | \u4ea4\u901a | \u516c\u5b89 | \u81ea\u7136\u8d44\u6e90 | \u6c14\u8c61 | \u4eba\u793e | \u8fd0\u8425\u5546", "475569"]
  ];
  
  layers.forEach(function(ly, i) {
    var y = 0.95 + i * 0.85;
    rbox(sl, 0.5, y, 9, 0.7, C.card, 0.08);
    rbox(sl, 0.65, y + 0.1, 2.2, 0.5, ly[2], 0.06);
    sl.addText(ly[0], { x: 0.65, y: y + 0.1, w: 2.2, h: 0.5, fontSize: 10, fontFace: FONT, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(ly[1], { x: 3.1, y: y, w: 6.2, h: 0.7, fontSize: 8, fontFace: FONT, color: C.gray, valign: "middle" });
    
    // Connector arrow
    if (i < layers.length - 1) {
      var arr = sl.addShape(pres.shapes.DOWN_ARROW, { x: 1.85, y: y + 0.76, w: 0.3, h: 0.12, fill: { color: C.light } });
    }
  });
  
  badge(sl, pn, C.accent);
}

// ====== SLIDE 5: ROLE MATRIX OVERVIEW ======
function sRoleMatrix() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: "0F172A" };
  
  // Decorative circles
  sl.addShape(pres.shapes.OVAL, { x: 7, y: -1, w: 4, h: 4, fill: { color: "1E3A5F", transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -1, y: 3, w: 3, h: 3, fill: { color: "1E3A5F", transparency: 60 } });
  
  sl.addText("\u56db\u7c7b\u89d2\u8272\u529f\u80fd\u77e9\u9635\u603b\u89c8", { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 22, fontFace: FONT, color: "FFFFFF", bold: true });
  sl.addText("\u5404\u89d2\u8272\u5728\u5e73\u53f0\u4e2d\u7684\u53c2\u4e0e\u5ea6\u4e0e\u6838\u5fc3\u4ef7\u503c\u5b9a\u4f4d", { x: 0.5, y: 0.58, w: 9, h: 0.25, fontSize: 10, fontFace: FONT, color: "94A3B8" });
  
  // 4 role cards in a row
  var roleCards = [
    ["\u653f\u5e9c\u9886\u5bfc\u5c42", "\u6570\u636e\u9a71\u52a8", "\u51b3\u7b56\u652f\u6301", C.gov, "1E40AF"],
    ["\u56ed\u533a\u7ba1\u7406\u5c42", "\u667a\u6167\u8fd0\u8425", "\u65e5\u5e38\u7ba1\u63a7", C.mgmt, "065F46"],
    ["\u5165\u9a7b\u4f01\u4e1a", "\u4e00\u7ad9\u5f0f\u670d\u52a1", "\u8d44\u6e90\u5bf9\u63a5", C.ent, "92400E"],
    ["\u516c\u4f17\u7528\u6237", "\u4fe1\u606f\u670d\u52a1", "\u53c2\u4e0e\u4e92\u52a8", C.pub, "5B21B6"]
  ];
  
  roleCards.forEach(function(r, i) {
    var x = 0.5 + i * 2.3;
    card(sl, x, 1.0, 2.1, 2.6, "1E293B");
    sbox(sl, x + 0.1, 1.1, 1.9, 0.04, r[3]);
    sl.addText(r[0], { x: x, y: 1.3, w: 2.1, h: 0.35, fontSize: 16, fontFace: FONT, color: "FFFFFF", bold: true, align: "center" });
    sbox(sl, x + 0.5, 1.7, 1.1, 0.02, r[3]);
    sl.addText(r[1], { x: x, y: 1.85, w: 2.1, h: 0.3, fontSize: 12, fontFace: FONT, color: r[3], align: "center" });
    sl.addText(r[2], { x: x, y: 2.15, w: 2.1, h: 0.3, fontSize: 12, fontFace: FONT, color: r[3], align: "center" });
    
    // Feature count
    sl.addShape(pres.shapes.OVAL, { x: x + 0.65, y: 2.65, w: 0.8, h: 0.8, fill: { color: r[3], transparency: 20 } });
    sl.addText("25+", { x: x + 0.65, y: 2.65, w: 0.8, h: 0.5, fontSize: 16, fontFace: "Arial", color: r[3], bold: true, align: "center", valign: "middle" });
    sl.addText("\u529f\u80fd", { x: x + 0.65, y: 3.1, w: 0.8, h: 0.25, fontSize: 8, fontFace: FONT, color: "94A3B8", align: "center" });
    
    // Bottom strip
    rbox(sl, x + 0.15, 3.5, 1.8, 0.04, r[3], 0.02);
  });
}

// ====== SLIDE 6: GOV - DASHBOARD ======
function sGovDashboard() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.gov);
  
  sl.addText("\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 \u6570\u636e\u9a7e\u9a76\u8231\u4e0e\u51b3\u7b56\u652f\u6301", { x: 0.5, y: 0.12, w: 8, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  sl.addText("\u5168\u666f\u6570\u636e\u53ef\u89c6\u5316\uff0c\u5168\u9762\u638c\u63a1\u56ed\u533a\u8fd0\u884c\u72b6\u6001", { x: 0.5, y: 0.52, w: 7, h: 0.25, fontSize: 10, fontFace: FONT, color: C.gray });
  
  // Top metric bar - 3 numbers
  var govMetrics = [
    ["57.68\u4ebf", "\u5168\u53bfGDP", C.gov],
    ["47\u5bb6", "\u56ed\u533a\u4f01\u4e1a", "2563EB"],
    ["137\u4ebf", "\u56fd\u4f01\u8d44\u4ea7", "1D4ED8"]
  ];
  govMetrics.forEach(function(m, i) {
    var x = 0.5 + i * 3.1;
    rbox(sl, x, 0.88, 2.8, 0.8, C.card, 0.1);
    sl.addText(m[0], { x: x + 0.15, y: 0.92, w: 1.5, h: 0.4, fontSize: 22, fontFace: "Arial", color: m[2], bold: true });
    sl.addText(m[1], { x: x + 1.6, y: 0.95, w: 1.1, h: 0.35, fontSize: 10, fontFace: FONT, color: C.gray, valign: "middle" });
  });
  
  // Feature cards - 2x3 grid
  var govFeatures = [
    ["AI\u62db\u5546\u670d\u52a1", "\u5927\u6570\u636e+AI\u7b97\u6cd5\u7cbe\u51c6\u5339\u914d\u76ee\u6807\u4f01\u4e1a\uff0c\u667a\u80fd\u63a8\u8350\u4e0e\u8ddf\u8e2a", "\u591a\u7ef4\u6570\u636e\u5206\u6790"], 
    ["\u6570\u5b57\u6c99\u76d8", "3D\u53ef\u89c6\u5316\u5c55\u793a\u56ed\u533a\u89c4\u5212\u4e0e\u571f\u5730\u8d44\u6e90\uff0c\u6c89\u6d78\u5f0f\u62db\u5546\u5c55\u793a", "\u7acb\u4f53\u5730\u56fe"],
    ["\u4ea7\u4e1a\u94fe\u56fe\u8c31", "\u53ef\u89c6\u5316\u5c55\u793a\u4ea7\u4e1a\u94fe\u4e0a\u4e0b\u6e38\u5173\u7cfb\uff0c\u7cbe\u51c6\u5f3a\u94fe\u8865\u94fe", "\u4ea7\u4e1a\u5206\u6790"],
    ["\u6295\u8d44\u6982\u7b97", "\u201c\u4e09\u6b65\u8d70\u201d\u521b\u5efa\u8def\u5f84+12\u5927\u4e13\u9879\u6295\u8d44\u8be6\u7ec6\u7b97\u79d1\u5b66\u89c4\u5212", "\u6295\u8d44\u76d1\u7763"],
    ["\u76d1\u7763\u7edf\u8ba1", "\u5404\u7c7b\u6570\u636e\u7edf\u8ba1\u62a5\u8868\uff0c\u7edf\u4e00\u6570\u636e\u51fa\u53e3\uff0c\u652f\u6301\u5bfc\u51fa", "\u6570\u636e\u6c47\u62a5"],
    ["\u73af\u4fdd\u7eff\u8272\u56ed\u533a", "\u73af\u4fdd\u76d1\u63a7\u6570\u636e\u4e3a\u7533\u62a5\u7eff\u8272\u56ed\u533a\u63d0\u4f9b\u771f\u5b9e\u53ef\u8ffd\u6eaf\u8bc1\u636e", "\u7eff\u8272\u521b\u5efa"]
  ];
  
  govFeatures.forEach(function(f, i) {
    var col = i % 3, row = Math.floor(i / 3);
    var x = 0.5 + col * 3.1, y = 1.95 + row * 1.35;
    card(sl, x, y, 2.9, 1.2, C.card);
    sbox(sl, x + 0.1, y + 0.1, 0.04, 1.0, C.gov);
    sl.addText(f[0], { x: x + 0.25, y: y + 0.08, w: 2.5, h: 0.25, fontSize: 11, fontFace: FONT, color: C.gov, bold: true });
    sl.addText(f[1], { x: x + 0.25, y: y + 0.35, w: 2.5, h: 0.55, fontSize: 8, fontFace: FONT, color: C.text });
    rbox(sl, x + 0.25, y + 0.95, 1.4, 0.2, C.govLite, 0.03);
    sl.addText(f[2], { x: x + 0.25, y: y + 0.95, w: 1.4, h: 0.2, fontSize: 7, fontFace: FONT, color: C.gov, align: "center", valign: "middle" });
  });
  
  badge(sl, pn, C.gov);
}

// ====== SLIDE 7: GOV - AI RECRUIT ======
function sGovRecruit() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.gov);
  
  sl.addText("\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 AI\u62db\u5546\u670d\u52a1\u5168\u6d41\u7a0b", { x: 0.5, y: 0.12, w: 8, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Timeline / flow
  var steps = [
    ["\u76ee\u6807\u4f01\u4e1a\u6316\u6398", "\u591a\u7ef4\u6570\u636e\u6e90+\u667a\u80fd\u7b97\u6cd5\uff0c\u81ea\u52a8\u63a8\u8350\u4e0e\u56ed\u533a\u4ea7\u4e1a\u9ad8\u5ea6\u5339\u914d\u7684\u76ee\u6807\u4f01\u4e1a", C.gov],
    ["AI\u667a\u80fd\u5339\u914d", "\u57fa\u4e8e\u4ea7\u4e1a\u94fe\u3001\u8d44\u91d1\u3001\u6280\u672f\u3001\u5730\u7406\u591a\u7ef4\u5ea6\u7b97\u6cd5\u5339\u914d\uff0c\u667a\u80fd\u8bc4\u5206", "2563EB"],
    ["\u667a\u80fd\u63a8\u8350", "\u81ea\u52a8\u751f\u6210\u62db\u5546\u7b56\u7565\u4e0e\u63a5\u89e6\u65b9\u6848\uff0c\u667a\u80fd\u63a8\u8350\u63a5\u89e6\u65f6\u673a", "1D4ED8"],
    ["\u5168\u7a0b\u8ddf\u8e2a", "\u62db\u5546\u9879\u76ee\u5168\u751f\u547d\u5468\u671f\u7ba1\u7406\uff0c\u4ece\u63a5\u89e6\u5230\u7b7e\u7ea6\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316", C.gov],
    ["\u62db\u5546\u6570\u636e\u62a5\u544a", "\u62db\u5546\u8f6c\u5316\u7387\u3001\u5408\u7ea6\u91d1\u989d\u3001\u4ea7\u4e1a\u805a\u96c6\u5ea6\u591a\u7ef4\u5ea6\u5206\u6790\u62a5\u544a", "2563EB"]
  ];
  
  steps.forEach(function(s, i) {
    var y = 0.8 + i * 0.85;
    // Circle with number
    sl.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.1, w: 0.5, h: 0.5, fill: { color: s[2] } });
    sl.addText(String(i+1), { x: 0.5, y: y + 0.1, w: 0.5, h: 0.5, fontSize: 14, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    // Card
    rbox(sl, 1.2, y, 8.2, 0.7, C.card, 0.08);
    sl.addText(s[0], { x: 1.4, y: y + 0.05, w: 3.5, h: 0.3, fontSize: 12, fontFace: FONT, color: s[2], bold: true });
    sl.addText(s[1], { x: 1.4, y: y + 0.35, w: 7.8, h: 0.3, fontSize: 9, fontFace: FONT, color: C.text });
    // Connector line
    if (i < steps.length - 1) {
      sl.addShape(pres.shapes.RECTANGLE, { x: 0.73, y: y + 0.6, w: 0.04, h: 0.25, fill: { color: C.light } });
    }
  });
  
  badge(sl, pn, C.gov);
}

// ====== SLIDE 8: GOV - INVESTMENT ======
function sGovInvest() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.accent);
  
  sl.addText("\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 \u6295\u8d44\u6982\u7b97\u4e0e\u521b\u5efa\u8def\u5f84", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Three-step path
  var pathSteps = [
    ["\u7701\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a", "\u57fa\u4e8e\u5e73\u53f0\u6570\u636e\u652f\u6491\uff0c\u7cfb\u7edf\u5316\u6574\u7406\u7533\u62a5\u6750\u6599\uff0c\u52a0\u901f\u521b\u5efa\u8fdb\u7a0b", C.gov],
    ["\u7eff\u8272\u5de5\u4e1a\u56ed\u533a", "\u73af\u4fdd\u76d1\u63a7\u3001\u80fd\u8017\u6570\u636e\u3001\u6392\u653e\u6307\u6807\u5168\u7a0b\u53ef\u8ffd\u6eaf\uff0c\u79d1\u5b66\u8bc1\u660e", C.mgmt],
    ["\u56fd\u5bb6\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a\u57f9\u80b2\u5bf9\u8c61", "\u4ea9\u5747\u8bba\u82f1\u96c4\u6570\u636e\u3001\u4ea7\u4e1a\u805a\u96c6\u5ea6\u3001\u79d1\u6280\u521b\u65b0\u6307\u6807\u7efc\u5408\u8bc4\u4f30", C.ent]
  ];
  
  pathSteps.forEach(function(p, i) {
    var x = 0.5 + i * 3.15;
    rbox(sl, x, 0.9, 2.85, 2.2, C.card, 0.1);
    sbox(sl, x, 0.9, 2.85, 0.06, p[2]);
    sl.addShape(pres.shapes.OVAL, { x: x + 1.15, y: 1.1, w: 0.55, h: 0.55, fill: { color: p[2] } });
    sl.addText(String(i+1), { x: x + 1.15, y: 1.1, w: 0.55, h: 0.55, fontSize: 18, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(p[0], { x: x + 0.15, y: 1.7, w: 2.55, h: 0.3, fontSize: 12, fontFace: FONT, color: p[2], bold: true, align: "center" });
    sl.addText(p[1], { x: x + 0.15, y: 2.05, w: 2.55, h: 0.7, fontSize: 9, fontFace: FONT, color: C.text, align: "center" });
    
    // Arrow between
    if (i < pathSteps.length - 1) {
      sl.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.95, y: 1.75, w: 0.25, h: 0.25, fill: { color: C.light } });
    }
  });
  
  // Investment detail section
  rbox(sl, 0.5, 3.4, 9, 1.0, "0F172A", 0.08);
  sl.addText("\u56ed\u533a\u6295\u8d44\u6982\u7b97\u603b\u89c8", { x: 0.7, y: 3.5, w: 3, h: 0.25, fontSize: 11, fontFace: FONT, color: C.accent, bold: true });
  
  var invItems = [
    ["\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7", "1,580,000"],
    ["\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7", "428,600"],
    ["\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546", "442,200"],
    ["\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", "348,000"]
  ];
  invItems.forEach(function(it, i) {
    var x = 0.7 + (i%2)*4.4, y = 3.85 + Math.floor(i/2)*0.25;
    sl.addShape(pres.shapes.OVAL, { x, y: y + 0.03, w: 0.08, h: 0.08, fill: { color: C.accent } });
    sl.addText(it[0], { x: x + 0.15, y: y, w: 2.5, h: 0.2, fontSize: 8, fontFace: FONT, color: "94A3B8" });
    sl.addText("\u00a5" + it[1], { x: x + 2.8, y: y, w: 1.5, h: 0.2, fontSize: 8, fontFace: "Arial", color: C.accent, align: "right" });
  });
  
  badge(sl, pn, C.accent);
}

// ====== SLIDE 9: GOV - SECURITY ======
function sGovSecurity() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.gov);
  
  sl.addText("\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 \u5168\u57df\u5b89\u5168\u4e0e\u73af\u4fdd\u76d1\u63a7", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Three vertical columns
  var secCols = [
    ["\u5929\u773c", "\u9053\u8def\u4ea4\u901a\u5b89\u5168\u76d1\u63a7", "\u667a\u80fd\u4ea4\u901a\u76d1\u63a7\uff0c\u8f66\u8f86\u8bc6\u522b\uff0c\u8fdd\u7ae0\u6355\u6349\uff0c\u4ea4\u901a\u72b6\u6001\u53ef\u89c6\u5316", C.gov],
    ["\u54e8\u5175", "\u5730\u707e\u5b89\u5168\u76d1\u6d4b", "\u5730\u8d28\u707e\u5bb3\u5b9e\u65f6\u76d1\u6d4b\uff0c\u591a\u79cd\u4f20\u611f\u5668\u8054\u52a8\uff0c\u9884\u8b66\u63d0\u524d\u62a5\u7ba1\u63a7", "1D4ED8"],
    ["\u6167\u773c", "\u73af\u4fdd\u76d1\u63a7\u68c0\u6d4b", "\u73af\u5883\u6307\u6807\u5b9e\u65f6\u76d1\u6d4b\uff0c\u6392\u653e\u6570\u636e\u8ffd\u8e2a\uff0c\u73af\u4fdd\u5408\u89c4\u5206\u6790", "1E40AF"]
  ];
  
  secCols.forEach(function(c, i) {
    var x = 0.5 + i * 3.15;
    card(sl, x, 0.85, 2.9, 2.5, C.card);
    rbox(sl, x + 1.0, 0.95, 0.9, 0.5, c[3], 0.06);
    sl.addText(c[0], { x: x + 1.0, y: 0.95, w: 0.9, h: 0.5, fontSize: 14, fontFace: FONT, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(c[1], { x: x + 0.15, y: 1.55, w: 2.6, h: 0.3, fontSize: 11, fontFace: FONT, color: c[3], bold: true, align: "center" });
    sl.addText(c[2], { x: x + 0.2, y: 1.9, w: 2.5, h: 0.7, fontSize: 9, fontFace: FONT, color: C.text, align: "center" });
  });
  
  // Bottom: related features
  var related = [
    "\u89c6\u9891\u76d1\u63a7\u603b\u89c8\uff1a\u7edf\u4e00\u89c6\u9891\u5e73\u53f0\uff0c16\u5bab\u683c\u667a\u80fd\u6392\u5217\uff0c\u5168\u56ed\u533a\u8986\u76d6",
    "\u5e94\u6025\u6307\u6325\u8054\u52a8\uff1a\u4e00\u952e\u8c03\u5ea6\u3001\u9884\u6848\u7ba1\u7406\u3001\u5b9e\u65f6\u901a\u8baf\uff0c\u54cd\u5e94\u6548\u7387\u63d0\u534760%",
    "\u8bbe\u5907\u7ba1\u7406\uff1a\u6444\u50cf\u5934\u3001\u611f\u5e94\u5668\u3001GPS\u3001\u96f7\u8fbe\u7b49\u5168\u7c7b\u8bbe\u5907\u7edf\u4e00\u7ba1\u63a7",
    "\u544a\u8b66\u4e2d\u5fc3\uff1a\u5f02\u5e38\u4e8b\u4ef6\u5b9e\u65f6\u544a\u8b66\u3001\u8d8b\u52bf\u5206\u6790\u3001\u9884\u8b66\u5904\u7f6e\u5168\u6d41\u7a0b\u7ba1\u7406"
  ];
  
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9, h: 0.02, fill: { color: C.light } });
  related.forEach(function(r, i) {
    var x = 0.5 + (i%2)*4.75, y = 3.85 + Math.floor(i/2)*0.6;
    rbox(sl, x, y, 4.5, 0.5, C.card, 0.06);
    sl.addShape(pres.shapes.OVAL, { x: x + 0.12, y: y + 0.12, w: 0.08, h: 0.08, fill: { color: C.gov } });
    sl.addText(r, { x: x + 0.3, y: y, w: 4.0, h: 0.5, fontSize: 9, fontFace: FONT, color: C.text, valign: "middle" });
  });
  
  badge(sl, pn, C.gov);
}

// ====== SLIDE 10: MGMT - DASHBOARD ======
function sMgmtDashboard() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.mgmt);
  
  sl.addText("\u56ed\u533a\u7ba1\u7406\u5c42 \u2014 \u667a\u6167\u8fd0\u8425\u4e0e\u5168\u57df\u76d1\u63a7", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Top metrics
  var mgmtMetrics = [
    ["60%", "\u5e94\u6025\u54cd\u5e94\u63d0\u5347", C.mgmt],
    ["95%", "\u76d1\u63a7\u8986\u76d6\u7387", "047857"],
    ["30%", "\u7ba1\u7406\u6548\u7387\u63d0\u5347", "065F46"]
  ];
  mgmtMetrics.forEach(function(m, i) {
    var x = 0.5 + i * 3.1;
    card(sl, x, 0.8, 2.8, 0.65, "0F172A");
    sl.addText(m[0], { x: x + 0.15, y: 0.83, w: 1.2, h: 0.6, fontSize: 24, fontFace: "Arial", color: m[2], bold: true, valign: "middle" });
    sl.addText(m[1], { x: x + 1.4, y: 0.83, w: 1.3, h: 0.6, fontSize: 10, fontFace: FONT, color: "94A3B8", valign: "middle" });
  });
  
  // 2x4 card grid for management features
  var mgmtFeatures = [
    ["\u9053\u8def\u4ea4\u901a\u76d1\u63a7", "\u8f66\u8f86\u8bc6\u522b\u3001\u8fdd\u7ae0\u6355\u6349\u3001\u4ea4\u901a\u72b6\u6001\u5b9e\u65f6\u53ef\u89c6\u5316"],
    ["\u5730\u707e\u5b89\u5168\u76d1\u6d4b", "\u5730\u8d28\u707e\u5bb3\u5b9e\u65f6\u76d1\u6d4b\u3001\u591a\u4f20\u611f\u5668\u8054\u52a8\u9884\u8b66"],
    ["\u73af\u4fdd\u76d1\u63a7\u68c0\u6d4b", "\u73af\u5883\u6307\u6807\u5b9e\u65f6\u76d1\u6d4b\u3001\u6392\u653e\u6570\u636e\u8ffd\u8e2a"],
    ["\u89c6\u9891\u76d1\u63a7\u603b\u89c8", "\u7edf\u4e00\u89c6\u9891\u5e73\u53f0\u3001\u667a\u80fd\u6392\u5217\u3001\u5168\u56ed\u533a\u8986\u76d6"],
    ["\u5e94\u6025\u6307\u6325\u8054\u52a8", "\u4e00\u952e\u8c03\u5ea6\u3001\u9884\u6848\u7ba1\u7406\u3001\u5b9e\u65f6\u901a\u8baf"],
    ["\u80fd\u6548\u76d1\u6d4b\u5206\u6790", "\u667a\u80fd\u8ba1\u91cf\u8868\u5b9e\u65f6\u91c7\u96c6\u3001\u7528\u80fd\u8d8b\u52bf\u5206\u6790"],
    ["\u4ea9\u5747\u8bba\u82f1\u96c4", "\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\u3001\u8d44\u6e90\u5dee\u5f02\u5316\u914d\u7f6e"],
    ["\u5f02\u5e38\u9884\u8b66", "\u7528\u80fd\u5f02\u5e38\u81ea\u52a8\u544a\u8b66\u3001\u78b3\u6392\u7ba1\u7406\u8ffd\u8e2a"]
  ];
  
  mgmtFeatures.forEach(function(f, i) {
    var col = i % 4, row = Math.floor(i / 4);
    var x = 0.5 + col * 2.3, y = 1.7 + row * 1.5;
    card(sl, x, y, 2.1, 1.3, C.card);
    sbox(sl, x + 0.1, y + 0.1, 1.9, 0.04, C.mgmt);
    sl.addShape(pres.shapes.OVAL, { x: x + 0.85, y: y + 0.25, w: 0.4, h: 0.4, fill: { color: C.mgmtLite } });
    sl.addText(f[0], { x: x + 0.08, y: y + 0.7, w: 1.94, h: 0.25, fontSize: 10, fontFace: FONT, color: C.mgmt, bold: true, align: "center" });
    sl.addText(f[1], { x: x + 0.1, y: y + 0.95, w: 1.9, h: 0.3, fontSize: 7.5, fontFace: FONT, color: C.gray, align: "center" });
  });
  
  badge(sl, pn, C.mgmt);
}

// ====== SLIDE 11: MGMT - OPERATIONS ======
function sMgmtOps() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.mgmt);
  
  sl.addText("\u56ed\u533a\u7ba1\u7406\u5c42 \u2014 \u8fd0\u8425\u7ba1\u7406\u4e0e\u6570\u5b57\u5316\u529e\u516c", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Left column - major operation modules
  var opsModules = [
    ["\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3", "\u5927\u5c4f\u6307\u6325+\u56ed\u5c0f\u4e8c\u56e2\u961f+\u957f\u6548\u670d\u52a1", C.mgmt],
    ["\u65e5\u5e38\u8fd0\u8425\u7ba1\u7406", "\u6570\u5b57\u5316\u5ba1\u6279\u3001\u65e5\u5e38\u5de1\u68c0\u3001\u8d44\u4ea7\u7ba1\u7406\u3001\u6863\u6848\u7ba1\u7406", "047857"],
    ["\u8bbe\u5907\u4e0e\u7269\u8054\u7f51\u7ba1\u7406", "\u6444\u50cf\u5934\u3001\u4f20\u611f\u5668\u3001\u667a\u80fd\u8bbe\u5907\u7edf\u4e00\u7ba1\u63a7\u4e0e\u8fdc\u7a0b\u8c03\u5ea6", "065F46"],
    ["\u7efc\u5408\u6570\u636e\u7ba1\u63a7", "\u6570\u636e\u4e2d\u5fc3\u3001\u6570\u636e\u8d28\u91cf\u3001\u62a5\u8868\u7ba1\u7406\u3001\u6570\u636e\u5907\u4efd\u4e0e\u5b89\u5168", C.mgmt]
  ];
  
  opsModules.forEach(function(m, i) {
    var y = 0.85 + i * 0.9;
    card(sl, 0.5, y, 4.3, 0.75, C.card);
    sbox(sl, 0.6, y + 0.1, 0.04, 0.55, m[2]);
    sl.addText(m[0], { x: 0.8, y: y + 0.08, w: 3.8, h: 0.25, fontSize: 11, fontFace: FONT, color: m[2], bold: true });
    sl.addText(m[1], { x: 0.8, y: y + 0.35, w: 3.8, h: 0.35, fontSize: 8.5, fontFace: FONT, color: C.text });
  });
  
  // Right column - quick stats
  card(sl, 5.1, 0.85, 4.5, 3.45, C.card);
  sl.addText("\u8fd0\u8425\u6570\u636e\u4e00\u89c8", { x: 5.3, y: 0.92, w: 4, h: 0.3, fontSize: 12, fontFace: FONT, color: C.text, bold: true });
  sbox(sl, 5.1, 1.2, 4.5, 0.02, C.mgmt);
  
  var opsStats = [
    ["\u76d1\u63a7\u8bbe\u5907\u603b\u6570", "128\u53f0", C.mgmt],
    ["\u65e5\u5e38\u5ba1\u6279\u6d41\u7a0b", "15\u7c7b", "047857"],
    ["\u8d44\u4ea7\u7ba1\u7406\u603b\u6570", "2,800+\u4ef6", "065F46"],
    ["\u65e5\u5747\u5904\u7406\u4e8b\u4ef6", "50+\u6761", C.mgmt],
    ["\u6570\u636e\u5bf9\u63a5\u7cfb\u7edf", "12\u5957", "047857"],
    ["\u5728\u7ebf\u670d\u52a1\u6a21\u5757", "30+\u4e2a", "065F46"]
  ];
  
  opsStats.forEach(function(s, i) {
    var x = 5.3 + (i<3 ? 0 : 2.2), y = 1.35 + (i%3) * 0.75;
    sl.addText(s[0], { x, y, w: 1.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });
    sl.addText(s[1], { x: x + 1.2, y: y - 0.03, w: 1.0, h: 0.28, fontSize: 16, fontFace: "Arial", color: s[2], bold: true, align: "right" });
    if (i < opsStats.length - 1) sl.addShape(pres.shapes.RECTANGLE, { x: x, y: y + 0.27, w: 2.1, h: 0.005, fill: { color: C.light } });
  });
  
  badge(sl, pn, C.mgmt);
}

// ====== SLIDE 12: MGMT - ENERGY ======
function sMgmtEnergy() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.mgmt);
  
  sl.addText("\u56ed\u533a\u7ba1\u7406\u5c42 \u2014 \u80fd\u6548\u76d1\u6d4b\u4e0e\u4ea9\u5747\u8bba\u82f1\u96c4", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Full-width cards - 2 per row
  var energyMods = [
    ["\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", "\u667a\u80fd\u8ba1\u91cf\u8868\u8ba1\u5b9e\u65f6\u91c7\u96c6\u80fd\u8017\u6570\u636e\uff0c\u652f\u6301\u6c34\u3001\u7535\u3001\u6c14\u591a\u7ef4\u5ea6\u76d1\u63a7\u4e0e\u5206\u6790\uff0c\u53ef\u89c6\u5316\u5c55\u793a\u80fd\u8017\u8d8b\u52bf\u4e0e\u5f02\u5e38\u544a\u8b66", C.mgmt],
    ["\u4ea9\u5747\u8bba\u82f1\u96c4", "\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u901a\u8fc7\u6295\u5165\u4ea7\u51fa\u3001\u80fd\u8017\u6392\u653e\u7b49\u591a\u7ef4\u5ea6\u6307\u6807\u5bf9\u4f01\u4e1a\u8fdb\u884c\u7efc\u5408\u8bc4\u4ef7\uff0c\u5b9e\u73b0\u8d44\u6e90\u8981\u7d20\u5dee\u5f02\u5316\u914d\u7f6e", "047857"],
    ["\u80fd\u6548\u5206\u6790", "\u4f01\u4e1a\u80fd\u8017\u6548\u7387\u591a\u7ef4\u5ea6\u5206\u6790\uff0c\u540c\u6bd4\u73af\u6bd4\u8d8b\u52bf\u53ef\u89c6\u5316\uff0c\u5e2e\u52a9\u4f01\u4e1a\u8bc6\u522b\u80fd\u8017\u6d6a\u8d39\u70b9\uff0c\u4f18\u5316\u751f\u4ea7\u8fc7\u7a0b\u80fd\u8017", "065F46"],
    ["\u5f02\u5e38\u9884\u8b66", "\u7528\u80fd\u5f02\u5e38\u81ea\u52a8\u544a\u8b66\uff0c\u53d1\u73b0\u80fd\u6e90\u6d6a\u8d39\u95ee\u9898\uff0c\u78b3\u6392\u653e\u7ba1\u7406\u8ffd\u8e2a\uff0c\u652f\u6301\u4f01\u4e1a\u7eff\u8272\u4f4e\u78b3\u5347\u7ea7", C.mgmt]
  ];
  
  energyMods.forEach(function(m, i) {
    var x = 0.5 + (i%2) * 4.65, y = 0.85 + Math.floor(i/2) * 1.5;
    card(sl, x, y, 4.35, 1.3, C.card);
    rbox(sl, x, y, 4.35, 0.5, m[3], 0.08);
    sbox(sl, x, y + 0.25, 4.35, 0.25, m[3]);
    sl.addText(m[0], { x: x + 0.15, y: y + 0.08, w: 4, h: 0.4, fontSize: 13, fontFace: FONT, color: "FFFFFF", bold: true });
    sl.addText(m[1], { x: x + 0.15, y: y + 0.6, w: 4.05, h: 0.65, fontSize: 8.5, fontFace: FONT, color: C.text });
  });
  
  // Bottom: data center integration
  rbox(sl, 0.5, 4.2, 9, 0.35, C.govLite, 0.06);
  sl.addText("\u6570\u636e\u4e2d\u5fc3\u96c6\u6210\uff1a12\u5927\u4e09\u65b9\u6570\u636e\u6765\u6e90\uff08\u5929\u7136\u6c14\u3001\u7535\u529b\u3001\u6c34\u3001\u7efc\u5408\u6267\u6cd5\u3001\u73af\u4fdd\u3001\u5e94\u6025\u3001\u4ea4\u901a\u3001\u516c\u5b89\u3001\u81ea\u7136\u8d44\u6e90\u3001\u6c14\u8c61\u3001\u4eba\u793e\u3001\u8fd0\u8425\u5546\uff09\u7edf\u4e00\u5bf9\u63a5",
    { x: 0.7, y: 4.2, w: 8.6, h: 0.35, fontSize: 9, fontFace: FONT, color: C.gov, valign: "middle" });
  
  badge(sl, pn, C.mgmt);
}

// ====== SLIDE 13: MGMT - DEVICE ======
function sMgmtDevice() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.mgmt);
  
  sl.addText("\u56ed\u533a\u7ba1\u7406\u5c42 \u2014 \u8bbe\u5907\u7ba1\u7406\u4e0e\u6570\u636e\u57fa\u7840", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Left: device categories
  var devices = [
    ["\u6444\u50cf\u5934\u7ba1\u7406", "\u5168\u56ed\u533a\u6444\u50cf\u5934\u7edf\u4e00\u7ba1\u63a7\u3001\u7ebf\u4e0a\u7ef4\u62a4\u3001\u72b6\u6001\u76d1\u63a7", C.mgmt],
    ["GPS\u8bbe\u5907", "\u8f66\u8f86GPS\u5b9a\u4f4d\u8bbe\u5907\u7ba1\u7406\uff0c\u5b9e\u65f6\u8f68\u8ff9\u8ffd\u8e2a\u4e0e\u8c03\u5ea6", "047857"],
    ["\u96f7\u8fbe\u76d1\u6d4b\u8bbe\u5907", "\u96f7\u8fbe\u76d1\u6d4b\u8bbe\u5907\u7ba1\u7406\uff0c\u4eba\u5458/\u8d44\u4ea7\u5b9e\u65f6\u611f\u77e5", "065F46"],
    ["\u667a\u80fd\u79f0\u91cd\u8bbe\u5907", "\u5730\u78c5\u7b49\u79f0\u91cd\u8bbe\u5907\u7ba1\u7406\uff0c\u7269\u6d41\u8fdb\u51fa\u6570\u636e\u91c7\u96c6", C.mgmt]
  ];
  
  devices.forEach(function(d, i) {
    var y = 0.85 + i * 0.95;
    card(sl, 0.5, y, 4.3, 0.8, C.card);
    sl.addShape(pres.shapes.PARALLELOGRAM, { x: 0.6, y: y + 0.15, w: 0.55, h: 0.5, fill: { color: d[3] } });
    sl.addText(d[0], { x: 1.3, y: y + 0.1, w: 3.3, h: 0.25, fontSize: 11, fontFace: FONT, color: d[3], bold: true });
    sl.addText(d[1], { x: 1.3, y: y + 0.38, w: 3.3, h: 0.3, fontSize: 8.5, fontFace: FONT, color: C.text });
  });
  
  // Right: data management
  card(sl, 5.1, 0.85, 4.5, 3.65, C.card);
  sl.addText("\u6570\u636e\u7ba1\u7406\u80fd\u529b", { x: 5.3, y: 0.92, w: 4, h: 0.3, fontSize: 12, fontFace: FONT, color: C.text, bold: true });
  sbox(sl, 5.1, 1.2, 4.5, 0.02, C.mgmt);
  
  var dataCaps = [
    ["\u6570\u636e\u4e2d\u5fc3", "\u6c47\u805a\u56ed\u533a\u5168\u91cf\u6570\u636e\uff0c\u6784\u5efa\u7edf\u4e00\u6570\u636e\u6e56"],
    ["\u6570\u636e\u8d28\u91cf\u7ba1\u7406", "\u6570\u636e\u8d28\u91cf\u76d1\u63a7\u3001\u6570\u636e\u6e05\u6d17\u3001\u6807\u51c6\u5316\u5904\u7406"],
    ["\u62a5\u8868\u7ba1\u7406", "\u81ea\u5b9a\u4e49\u62a5\u8868\u3001\u591a\u7ef4\u5ea6\u5206\u6790\u3001\u652f\u6301\u5bfc\u51fa"],
    ["\u6570\u636e\u5907\u4efd", "\u81ea\u52a8\u5907\u4efd\u3001\u5bb9\u707e\u6062\u590d\u3001\u6570\u636e\u5b89\u5168\u4fdd\u969c"],
    ["\u65e5\u5fd7\u5ba1\u8ba1", "\u64cd\u4f5c\u65e5\u5fd7\u5168\u8bb0\u5f55\u3001\u5b89\u5168\u5ba1\u8ba1\u53ef\u8ffd\u6eaf"],
    ["\u7cfb\u7edf\u914d\u7f6e", "\u89d2\u8272\u6743\u9650\u3001\u901a\u77e5\u89c4\u5219\u3001\u53c2\u6570\u914d\u7f6e\u96c6\u4e2d\u7ba1\u7406"]
  ];
  
  dataCaps.forEach(function(d, i) {
    var y = 1.4 + i * 0.5;
    sl.addShape(pres.shapes.OVAL, { x: 5.4, y: y + 0.08, w: 0.08, h: 0.08, fill: { color: C.mgmt } });
    sl.addText(d[0], { x: 5.65, y, w: 1.6, h: 0.25, fontSize: 9, fontFace: FONT, color: C.mgmt, bold: true });
    sl.addText(d[1], { x: 7.3, y, w: 2.1, h: 0.25, fontSize: 8, fontFace: FONT, color: C.gray });
    if (i < dataCaps.length - 1) sbox(sl, 5.4, y + 0.3, 4.0, 0.005, C.light);
  });
  
  badge(sl, pn, C.mgmt);
}

// ====== SLIDE 14: ENT - SERVICE ======
function sEntService() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.ent);
  
  sl.addText("\u5165\u9a7b\u4f01\u4e1a \u2014 \u4f01\u4e1a\u670d\u52a1\u4e2d\u5fc3", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Top - 3 highlight cards
  var entHighlights = [
    ["\u8bc9\u6c42\u76f4\u8fbe", "\u968f\u624b\u62cd\u63d0\u4ea4\u8bc9\u6c42\uff0c\u652f\u6301\u6587\u5b57+\u56fe\u7247+\u5b9a\u4f4d\uff0c2\u5c0f\u65f6\u5185\u54cd\u5e94\u5904\u7406", C.ent],
    ["\u670d\u52a1\u76f4\u8fbe", "\u5929\u7136\u6c14/\u6c34/\u7535\u8d39\u5728\u7ebf\u7f34\u8d39\uff0c\u4e00\u7ad9\u5f0f\u4f01\u4e1a\u670d\u52a1\u4e2d\u5fc3", "B45309"],
    ["\u56ed\u4f01\u4ea4\u4e92", "\u5728\u7ebf\u6c9f\u901a\u3001\u901a\u77e5\u516c\u544a\u3001\u653f\u7b56\u63a8\u9001\u3001\u53cc\u5411\u4e92\u52a8\u6c9f\u901a", "92400E"]
  ];
  
  entHighlights.forEach(function(h, i) {
    var x = 0.5 + i * 3.15;
    card(sl, x, 0.8, 2.9, 1.35, C.card);
    sbox(sl, x + 0.1, 0.9, 0.04, 0.8, h[2]);
    sl.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 0.95, w: 0.35, h: 0.35, fill: { color: h[2], transparency: 20 } });
    sl.addText(h[0], { x: x + 0.75, y: 0.93, w: 2, h: 0.3, fontSize: 13, fontFace: FONT, color: h[2], bold: true });
    sl.addText(h[1], { x: x + 0.15, y: 1.35, w: 2.6, h: 0.5, fontSize: 8.5, fontFace: FONT, color: C.text });
  });
  
  // Bottom - more features in strip
  var entFeatures = [
    ["\u4f01\u4e1a\u5165\u9a7b", "\u7ebf\u4e0a\u7533\u8bf7\u3001\u8bc1\u7167\u7ba1\u7406\u3001\u5165\u9a7b\u6d41\u7a0b\u5168\u7ebf\u4e0a\u5316"],
    ["\u4fe1\u7528\u8bc4\u4ef7", "\u4f01\u4e1a\u4fe1\u7528\u591a\u7ef4\u5ea6\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u52a9\u529b\u878d\u8d44\u4e0e\u5408\u4f5c"],
    ["\u5408\u540c\u7ba1\u7406", "\u7535\u5b50\u5408\u540c\u3001\u7b7e\u7f72\u3001\u5b58\u6863\u5168\u6d41\u7a0b\u7ba1\u7406"],
    ["\u8d44\u4ea7\u7ba1\u7406", "\u4f01\u4e1a\u8d44\u4ea7\u76d8\u70b9\u3001\u5165\u5e93\u3001\u51fa\u5e93\u6570\u5b57\u5316\u7ba1\u7406"],
    ["\u6570\u5b57\u8bc1\u7167", "\u7535\u5b50\u8425\u4e1a\u6267\u7167\u3001\u8bb8\u53ef\u8bc1\u4e0a\u94fe\u7ba1\u7406"],
    ["\u5458\u5de5\u7ba1\u7406", "\u5165\u804c/\u79bb\u804c\u3001\u52a0\u73ed/\\u8bf7\u5047\u3001\u5de5\u8d44\u7ba1\u7406\u5b8c\u5584\u4f53\u7cfb"]
  ];
  
  entFeatures.forEach(function(f, i) {
    var col = i % 3, row = Math.floor(i / 3);
    var x = 0.5 + col * 3.15, y = 2.4 + row * 1.1;
    card(sl, x, y, 2.9, 0.9, C.card);
    rbox(sl, x + 0.1, y + 0.1, 0.9, 0.25, C.ent, 0.04);
    sl.addText(f[0], { x: x + 0.1, y: y + 0.1, w: 0.9, h: 0.25, fontSize: 8, fontFace: FONT, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(f[1], { x: x + 0.15, y: y + 0.45, w: 2.6, h: 0.35, fontSize: 8, fontFace: FONT, color: C.text });
    sl.addShape(pres.shapes.PARALLELOGRAM, { x: x + 2.2, y: y + 0.55, w: 0.5, h: 0.2, fill: { color: C.entLite } });
  });
  
  badge(sl, pn, C.ent);
}

// ====== SLIDE 15: ENT - LOGISTICS ======
function sEntLogistics() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.ent);
  
  sl.addText("\u5165\u9a7b\u4f01\u4e1a \u2014 \u4f9b\u9700\u5bf9\u63a5\u4e0e\u7269\u6d41\u670d\u52a1", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Left - Supply/Demand Modules
  sdModules = [
    ["\u7269\u6d41\u4f9b\u9700", "\u56ed\u533a\u8d27\u6e90+\u8fd4\u7a0b\u7a7a\u8f66+\u8f66\u8f86\u8d44\u6e90\u667a\u80fd\u5339\u914d\uff0c\u6bcf\u5f20\u5361\u7247\u5e26\u8054\u7cfb\u65b9\u5f0f\uff0c\u652f\u6301\u591a\u6761\u4ef6\u7b5b\u9009", C.ent],
    ["\u7279\u8272\u4f9b\u9700", "\u9752\u5ddd\u5c71\u73cd\u7b499\u79cd\u7279\u8272\u4ea7\u54c1\u5c55\u793a\u4e0e\u4f9b\u9700\u5bf9\u63a5\uff0c\u76f4\u63a5\u8054\u7cfb\u4f9b\u5e94\u5546", "B45309"],
    ["\u4f9b\u5e94\u5546\u7ba1\u7406", "\u4f9b\u5e94\u5546/\u751f\u4ea7\u5546/\u8fd0\u8f93\u5546\u5b8c\u5584\u7ba1\u7406\u4f53\u7cfb\uff0c\u5408\u4f5c\u5173\u7cfb\u7ef4\u62a4", "92400E"]
  ];
  
  sdModules.forEach(function(m, i) {
    var y = 0.85 + i * 0.95;
    card(sl, 0.5, y, 4.3, 0.8, C.card);
    sl.addShape(pres.shapes.OVAL, { x: 0.65, y: y + 0.15, w: 0.08, h: 0.08, fill: { color: m[2] } });
    sl.addText(m[0], { x: 0.9, y: y + 0.08, w: 3.5, h: 0.25, fontSize: 12, fontFace: FONT, color: m[2], bold: true });
    sl.addText(m[1], { x: 0.9, y: y + 0.35, w: 3.6, h: 0.35, fontSize: 8.5, fontFace: FONT, color: C.text });
    if (i < sdModules.length - 1) sbox(sl, 0.72, y + 0.85, 0.02, 0.1, C.light);
  });
  
  // Right - additional services
  card(sl, 5.1, 0.85, 4.5, 2.55, C.card);
  sl.addText("\u5546\u4e1a\u670d\u52a1\u6a21\u5757", { x: 5.3, y: 0.92, w: 4, h: 0.3, fontSize: 12, fontFace: FONT, color: C.text, bold: true });
  sbox(sl, 5.1, 1.2, 4.5, 0.02, C.ent);
  
  var bizMods = [
    ["\u6570\u5b57\u5316\u6d3b\u52a8", "\u62db\u6807/\u7ade\u4ef7/\u5ba1\u6279\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316"],
    ["\u5ba2\u6237\u7ba1\u7406", "\u5ba2\u6237\u5173\u7cfb\u7ef4\u62a4\u3001\u5546\u673a\u8ffd\u8e2a\u3001\u5408\u540c\u6267\u884c"],
    ["\u7535\u5b50\u540d\u7247", "\u4f01\u4e1a\u5458\u5de5\u7535\u5b50\u540d\u7247\u5feb\u901f\u5206\u4eab\uff0c\u4e00\u952e\u4fdd\u5b58\u8054\u7cfb\u4eba"],
    ["\u5fae\u4fe1\u63a8\u5e7f", "\u56ed\u533a\u5fae\u4fe1\u5c0f\u7a0b\u5e8f\u5feb\u901f\u63a8\u5e7f\u4e0e\u4f7f\u7528"]
  ];
  
  bizMods.forEach(function(b, i) {
    var y = 1.35 + i * 0.48;
    sl.addShape(pres.shapes.OVAL, { x: 5.4, y: y + 0.05, w: 0.06, h: 0.06, fill: { color: C.ent } });
    sl.addText(b[0], { x: 5.65, y, w: 1.3, h: 0.22, fontSize: 9, fontFace: FONT, color: C.ent, bold: true });
    sl.addText(b[1], { x: 7.0, y, w: 2.4, h: 0.22, fontSize: 8, fontFace: FONT, color: C.gray });
  });
  
  // Bottom: parking service
  rbox(sl, 0.5, 3.65, 4.3, 0.8, C.card, 0.08);
  sl.addText("\u505c\u8f66\u670d\u52a1", { x: 0.7, y: 3.72, w: 3, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sl.addText("\u9ad8\u5fb7\u5730\u56fe\u5bfc\u822a+\u5b9e\u65f6\u4f59\u4f4d\u67e5\u770b+\u5728\u7ebf\u7f34\u8d39\uff0c\u5168\u56ed\u533a\u505c\u8f66\u4f4d\u7edf\u4e00\u7ba1\u7406", { x: 0.7, y: 3.98, w: 3.9, h: 0.35, fontSize: 9, fontFace: FONT, color: C.text });
  
  rbox(sl, 5.1, 3.65, 4.5, 0.8, C.card, 0.08);
  sl.addText("\u653f\u7b56\u63a8\u9001", { x: 5.3, y: 3.72, w: 3, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sl.addText("\u60e0\u4f01\u653f\u7b56\u7cbe\u51c6\u5339\u914d\uff0c\u4e3b\u52a8\u63a8\u9001\u8865\u8d34/\u51cf\u514d/\u4eba\u624d\u7b49\u653f\u7b56\uff0c\u5bf9\u63a5\u6bcf\u5e74\u8d8520\u4ebf\u4e13\u9879\u8d44\u91d1", { x: 5.3, y: 3.98, w: 4.1, h: 0.35, fontSize: 9, fontFace: FONT, color: C.text });
  
  badge(sl, pn, C.ent);
}

// ====== SLIDE 16: ENT - VENUE ======
function sEntVenue() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.ent);
  
  sl.addText("\u5165\u9a7b\u4f01\u4e1a \u2014 \u573a\u9986\u3001\u6d3b\u52a8\u4e0e\u9910\u996e\u670d\u52a1", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Venue booking
  card(sl, 0.5, 0.8, 5.5, 2.0, C.card);
  rbox(sl, 0.5, 0.8, 5.5, 0.4, C.ent, 0.08);
  sbox(sl, 0.5, 1.0, 5.5, 0.2, C.ent);
  sl.addText("\u573a\u9986\u670d\u52a1", { x: 0.65, y: 0.85, w: 3, h: 0.3, fontSize: 14, fontFace: FONT, color: "FFFFFF", bold: true });
  
  var venues = ["\u591a\u529f\u80fd\u4f1a\u8bae\u5385", "\u5c55\u89c8\u5385", "\u4f53\u80b2\u9986", "\u57f9\u8bad\u6559\u5ba4", "\u4f1a\u8bae\u5ba4", "\u63a5\u5f85\u5385", "\u591a\u5a92\u4f53\u5385", "\u56ed\u533a\u5e7f\u573a"];
  venues.forEach(function(v, i) {
    var col = i % 4, row = Math.floor(i / 4);
    var x = 0.65 + col * 1.33, y = 1.35 + row * 0.6;
    rbox(sl, x, y, 1.2, 0.25, C.entLite, 0.04);
    sl.addText(v, { x, y, w: 1.2, h: 0.25, fontSize: 8, fontFace: FONT, color: C.ent, align: "center", valign: "middle" });
  });
  
  // Right: activity & catering
  card(sl, 6.3, 0.8, 3.3, 2.0, C.card);
  sl.addText("\u56ed\u533a\u6d3b\u52a8", { x: 6.45, y: 0.88, w: 3, h: 0.3, fontSize: 13, fontFace: FONT, color: C.ent, bold: true });
  var acts = ["\u6587\u4f53\u6d3b\u52a8", "\u57f9\u8bad\u8bb2\u5ea7", "\u516c\u76ca\u6d3b\u52a8", "\u4ea4\u6d41\u805a\u4f1a", "\u8282\u5e86\u6d3b\u52a8", "\u5c55\u89c8\u5c55\u793a"];
  acts.forEach(function(a, i) {
    var y = 1.25 + i * 0.22;
    sl.addShape(pres.shapes.OVAL, { x: 6.55, y: y + 0.04, w: 0.05, h: 0.05, fill: { color: C.ent } });
    sl.addText(a, { x: 6.75, y: y, w: 2.5, h: 0.2, fontSize: 8, fontFace: FONT, color: C.text });
  });
  
  // Bottom - 2 cards
  card(sl, 0.5, 3.05, 4.3, 1.1, C.card);
  sl.addText("\u56ed\u533a\u9910\u996e", { x: 0.7, y: 3.12, w: 3, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sl.addText("\u6d41\u52a8\u5361\u7247\u5c55\u793a6\u5bb6\u9910\u5385\uff0c20+\u9053\u83dc\u54c1\u4e00\u89c8\uff0c\u652f\u6301\u5206\u7c7b\u7b5b\u9009\u3001\u4eca\u65e5\u83dc\u5355\u67e5\u770b", { x: 0.7, y: 3.38, w: 3.9, h: 0.35, fontSize: 9, fontFace: FONT, color: C.text });
  rbox(sl, 0.7, 3.8, 1.2, 0.22, C.entLite, 0.03);
  sl.addText("\u83dc\u54c1\u4e00\u89c8", { x: 0.7, y: 3.8, w: 1.2, h: 0.22, fontSize: 7, fontFace: FONT, color: C.ent, align: "center", valign: "middle" });
  
  card(sl, 5.1, 3.05, 4.5, 1.1, C.card);
  sl.addText("\u7528\u5de5\u4fe1\u606f", { x: 5.3, y: 3.12, w: 3, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sl.addText("\u62db\u8058+\u5c97\u4f4d\u5339\u914d\uff0c\u652f\u6301\u5206\u7c7b\u7b5b\u9009\uff0c\u5e2e\u52a9\u4f01\u4e1a\u5feb\u901f\u89e3\u51b3\u7528\u5de5\u95ee\u9898", { x: 5.3, y: 3.38, w: 4.1, h: 0.35, fontSize: 9, fontFace: FONT, color: C.text });
  rbox(sl, 5.3, 3.8, 1.2, 0.22, C.entLite, 0.03);
  sl.addText("\u62db\u8058\u4fe1\u606f", { x: 5.3, y: 3.8, w: 1.2, h: 0.22, fontSize: 7, fontFace: FONT, color: C.ent, align: "center", valign: "middle" });
  
  badge(sl, pn, C.ent);
}

// ====== SLIDE 17: ENT - PAYMENT ======
function sEntPayment() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.ent);
  
  sl.addText("\u5165\u9a7b\u4f01\u4e1a \u2014 \u7f34\u8d39\u670d\u52a1\u4e0e\u5ba1\u6279\u6d41\u7a0b", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Payment modules - 3 cards
  var payMods = [
    ["\u5929\u7136\u6c14\u7f34\u8d39", "\u5728\u7ebf\u67e5\u770b\u6c14\u8d39\u8d26\u5355\u3001\u652f\u4ed8\u5386\u53f2\u8bb0\u5f55\uff0c\u652f\u6301\u591a\u79cd\u652f\u4ed8\u65b9\u5f0f\u5728\u7ebf\u7f34\u8d39", C.gov],
    ["\u81ea\u6765\u6c34\u7f34\u8d39", "\u5b9e\u65f6\u67e5\u770b\u7528\u6c34\u91cf\u3001\u6c34\u8d39\u8d26\u5355\uff0c\u652f\u6301\u5386\u53f2\u8bb0\u5f55\u67e5\u8be2\u4e0e\u7ebf\u4e0a\u7f34\u8d39", C.mgmt],
    ["\u7535\u8d39\u7f34\u8d39", "\u667a\u80fd\u8868\u8ba1\u5b9e\u65f6\u76d1\u63a7\u7528\u7535\u91cf\uff0c\u5728\u7ebf\u7f34\u8d39\u3001\u7528\u7535\u5206\u6790\u4e00\u7ad9\u5f0f\u670d\u52a1", C.ent]
  ];
  
  payMods.forEach(function(p, i) {
    var x = 0.5 + i * 3.15;
    card(sl, x, 0.8, 2.9, 1.6, C.card);
    rbox(sl, x, 0.8, 2.9, 0.4, p[2], 0.08);
    sbox(sl, x, 1.0, 2.9, 0.2, p[2]);
    sl.addText(p[0], { x: x + 0.15, y: 0.85, w: 2.6, h: 0.3, fontSize: 14, fontFace: FONT, color: "FFFFFF", bold: true });
    sl.addText(p[1], { x: x + 0.15, y: 1.35, w: 2.6, h: 0.75, fontSize: 9, fontFace: FONT, color: C.text });
    rbox(sl, x + 0.15, 2.1, 1.5, 0.2, C.light, 0.03);
    sl.addText("\u5728\u7ebf\u7f34\u8d39 >", { x: x + 0.15, y: 2.1, w: 1.5, h: 0.2, fontSize: 7, fontFace: FONT, color: C.gray, align: "center", valign: "middle" });
  });
  
  // Bottom left: approval processes
  card(sl, 0.5, 2.65, 4.3, 2.35, C.card);
  sl.addText("\u5ba1\u6279\u6d41\u7a0b", { x: 0.7, y: 2.72, w: 3, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sbox(sl, 0.5, 2.95, 4.3, 0.01, C.light);
  
  var approvals = [
    "\u56ed\u533a\u5ba1\u6279\u4e2d\u5fc3\uff1a\u7f51\u7edc/\u7269\u4e1a/\u62db\u5546\u7b49\u591a\u79cd\u5ba1\u6279\u6d41\u7a0b\u7ebf\u4e0a\u5316",
    "\u9879\u76ee\u70b9\u5ba1\u6279\uff1a\u4f01\u4e1a\u9879\u76ee\u70b9\u6570\u636e\u62a5\u9001\u4e0e\u5ba1\u6838",
    "\u65e5\u62a5\u7ba1\u7406\uff1a\u4f01\u4e1a\u65e5\u62a5\u586b\u5199\u3001\u67e5\u770b\u3001\u5ba1\u6838\u5168\u6d41\u7a0b",
    "\u5de5\u4f5c\u62a5\u544a\uff1a\u5de5\u4f5c\u62a5\u544a\u63d0\u4ea4\u4e0e\u5ba1\u67e5",
    "\u73b0\u573a\u5de1\u68c0\uff1a\u5de1\u68c0\u4efb\u52a1\u5206\u914d\u3001\u6267\u884c\u3001\u8bb0\u5f55\u5168\u8fc7\u7a0b"
  ];
  
  approvals.forEach(function(a, i) {
    var y = 3.1 + i * 0.37;
    sl.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.03, w: 0.05, h: 0.05, fill: { color: C.ent } });
    sl.addText(a, { x: 0.85, y: y, w: 3.7, h: 0.22, fontSize: 8, fontFace: FONT, color: C.text });
  });
  
  // Bottom right: visitor
  card(sl, 5.1, 2.65, 4.5, 2.35, C.card);
  sl.addText("\u8bbf\u5ba2\u4e0e\u95e8\u7981\u7ba1\u7406", { x: 5.3, y: 2.72, w: 4, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sbox(sl, 5.1, 2.95, 4.5, 0.01, C.light);
  
  var visits = [
    ["\u8bbf\u5ba2\u5f55\u5165", "\u9884\u7ea6\u3001\u767b\u8bb0\u3001\u62dc\u8bbf\u8bb0\u5f55\u5168\u7ebf\u4e0a\u5316"],
    ["\u8f66\u724c\u7ba1\u7406", "\u591a\u8f66\u724c\u7ba1\u7406\uff0c\u667a\u80fd\u8bc6\u522b\u4e0e\u653e\u884c"],
    ["\u95e8\u7981\u7ba1\u7406", "\u4eba\u5458/\u8f66\u8f86\u95e8\u7981\u7edf\u4e00\u7ba1\u63a7"],
    ["\u8bbf\u5ba2\u7edf\u8ba1", "\u4eba\u5458/\u8f66\u8f86\u8bbf\u5ba2\u8d8b\u52bf\u5206\u6790\u4e0e\u7edf\u8ba1\u5361\u7247"],
    ["\u505c\u8f66\u4f59\u4f4d", "\u5b9e\u65f6\u67e5\u770b\u505c\u8f66\u4f4d\u4f59\u4f4d\u5e76\u7ebf\u4e0a\u7f34\u8d39"]
  ];
  
  visits.forEach(function(v, i) {
    var y = 3.08 + i * 0.37;
    sl.addShape(pres.shapes.OVAL, { x: 5.4, y: y + 0.05, w: 0.06, h: 0.06, fill: { color: C.ent } });
    sl.addText(v[0], { x: 5.6, y: y, w: 1.3, h: 0.22, fontSize: 9, fontFace: FONT, color: C.ent, bold: true });
    sl.addText(v[1], { x: 6.95, y: y, w: 2.5, h: 0.22, fontSize: 8, fontFace: FONT, color: C.gray });
  });
  
  badge(sl, pn, C.ent);
}

// ====== SLIDE 18: ENT - EXHIBIT ======
function sEntExhibit() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.ent);
  
  sl.addText("\u5165\u9a7b\u4f01\u4e1a \u2014 \u5c55\u793a\u4e0e\u62db\u5546\u5f15\u8d44", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Top: 3 feature cards
  var exhibMods = [
    ["\u4e91\u5c55\u9986", "\u4f01\u4e1a\u5f62\u8c61\u7ebf\u4e0a\u5c55\u793a\uff0c\u4ea7\u54c1\u7acb\u4f53\u5c55\u793a\uff0c\u652f\u6301\u591a\u5a92\u4f53\u5185\u5bb9", C.ent],
    ["\u7535\u5b50\u540d\u7247", "\u4f01\u4e1a\u5458\u5de5\u7535\u5b50\u540d\u7247\u5feb\u901f\u5206\u4eab\uff0c\u4e00\u952e\u4fdd\u5b58\u8054\u7cfb\u4eba\uff0c\u63d0\u5347\u5546\u52a1\u6548\u7387", "B45309"],
    ["\u56ed\u533a\u5ba3\u4f20", "\u56ed\u533a\u6982\u51b5\u3001\u53d1\u5c55\u5386\u7a0b\u3001\u4f18\u52bf\u4ea7\u4e1a\u3001\u5408\u4f5c\u6848\u4f8b\u591a\u7ef4\u5ea6\u5c55\u793a", "92400E"]
  ];
  
  exhibMods.forEach(function(e, i) {
    var x = 0.5 + i * 3.15;
    card(sl, x, 0.8, 2.9, 1.5, C.card);
    sl.addShape(pres.shapes.PARALLELOGRAM, { x: x + 0.1, y: 0.9, w: 0.8, h: 0.35, fill: { color: e[2] } });
    sl.addText(e[0], { x: x + 0.1, y: 0.9, w: 0.8, h: 0.35, fontSize: 9, fontFace: FONT, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(e[1], { x: x + 0.15, y: 1.35, w: 2.6, h: 0.7, fontSize: 9, fontFace: FONT, color: C.text });
    rbox(sl, x + 0.15, 2.1, 0.8, 0.05, e[2], 0.02);
  });
  
  // Middle: target enterprise & investment
  card(sl, 0.5, 2.55, 4.3, 1.2, C.card);
  sl.addText("\u76ee\u6807\u4f01\u4e1a\u6316\u6398", { x: 0.7, y: 2.62, w: 3.5, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sbox(sl, 0.5, 2.85, 4.3, 0.01, C.light);
  sl.addText("\u591a\u7ef4\u6570\u636e\u6e90+\u667a\u80fd\u7b97\u6cd5\u81ea\u52a8\u63a8\u8350\u4e0e\u56ed\u533a\u4ea7\u4e1a\u9ad8\u5ea6\u5339\u914d\u7684\u76ee\u6807\u4f01\u4e1a\uff0c\u5b9e\u73b0\u7cbe\u51c6\u62db\u5546\u5f15\u8d44\u3002\u57fa\u4e8e\u4ea7\u4e1a\u94fe\u3001\u8d44\u91d1\u3001\u6280\u672f\u3001\u5730\u7406\u7b49\u591a\u7ef4\u5ea6\u7b97\u6cd5\u8bc4\u5206\uff0c\u667a\u80fd\u63a8\u8350\u63a5\u89e6\u65f6\u673a\u4e0e\u62db\u5546\u7b56\u7565\u3002",
    { x: 0.7, y: 2.92, w: 3.9, h: 0.7, fontSize: 9, fontFace: FONT, color: C.text });
  
  card(sl, 5.1, 2.55, 4.5, 1.2, C.card);
  sl.addText("\u6295\u8d44\u9879\u76ee\u7ba1\u7406", { x: 5.3, y: 2.62, w: 4, h: 0.25, fontSize: 12, fontFace: FONT, color: C.ent, bold: true });
  sbox(sl, 5.1, 2.85, 4.5, 0.01, C.light);
  sl.addText("\u62db\u5546\u9879\u76ee\u5168\u751f\u547d\u5468\u671f\u7ba1\u7406\uff0c\u4ece\u63a5\u89e6\u5230\u7b7e\u7ea6\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316\u3002\u62db\u5546\u6570\u636e\u591a\u7ef4\u5ea6\u5206\u6790\u62a5\u544a\uff0c\u5305\u62ec\u8f6c\u5316\u7387\u3001\u5408\u7ea6\u91d1\u989d\u3001\u4ea7\u4e1a\u805a\u96c6\u5ea6\u7b49\u5173\u952e\u6307\u6807\u3002",
    { x: 5.3, y: 2.92, w: 4.1, h: 0.7, fontSize: 9, fontFace: FONT, color: C.text });
  
  // Bottom: enterprise list
  rbox(sl, 0.5, 4.0, 9, 0.75, C.card, 0.08);
  sl.addText("\u56ed\u533a\u91cd\u70b9\u4f01\u4e1a\uff089\u5bb6\uff09\uff1a", { x: 0.7, y: 4.05, w: 2.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.ent, bold: true });
  var entNames = ["\u961c\u6210\u65b0\u6750", "\u8679\u660c\u6676\u4f53", "\u4e2d\u80fd\u5efa", "\u4f73\u5174\u94dd\u4e1a", "\u661f\u6052\u9752\u6e90", "\u4e5d\u65ed\u65b0\u6750", "\u534e\u7eb3\u5b9e\u4e1a", "\u65b0\u901a\u946b", "\u5e7f\u8882\u65b0\u6750"];
  entNames.forEach(function(ne, i) {
    var x = 0.7 + (i<5 ? i : i-5)*1.75, y = 4.35 + (i<5 ? 0 : 0.3);
    sl.addShape(pres.shapes.OVAL, { x: x, y: y + 0.03, w: 0.04, h: 0.04, fill: { color: C.ent } });
    sl.addText(ne, { x: x + 0.08, y: y, w: 1.5, h: 0.2, fontSize: 7, fontFace: FONT, color: C.text });
  });
  
  badge(sl, pn, C.ent);
}

// ====== SLIDE 19: PUBLIC ======
function sPublicInfo() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.pub);
  
  sl.addText("\u516c\u4f17\u7528\u6237 \u2014 \u4fe1\u606f\u670d\u52a1\u4e0e\u56ed\u533a\u53c2\u4e0e", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Top - full width info hub
  card(sl, 0.5, 0.75, 9, 1.4, C.card);
  sl.addText("\u56ed\u533a\u4fe1\u606f\u4e2d\u5fc3", { x: 0.7, y: 0.82, w: 4, h: 0.3, fontSize: 14, fontFace: FONT, color: C.pub, bold: true });
  
  var infoHubs = [
    ["\u56ed\u533a\u516c\u544a", "\u5b9e\u65f6\u67e5\u770b\u56ed\u533a\u901a\u77e5\uff08\u65bd\u5de5\u3001\u68c0\u67e5\u3001\u7ef4\u62a4\uff09"],
    ["\u56ed\u533a\u6d3b\u52a8", "\u67e5\u770b\u5e76\u53c2\u4e0e\u6587\u4f53/\u516c\u76ca\u7b49\u6d3b\u52a8"],
    ["\u56ed\u533a\u653f\u7b56", "\u67e5\u770b\u516c\u5f00\u653f\u7b56\u4fe1\u606f\uff0c\u4e86\u89e3\u56ed\u533a\u53d1\u5c55\u52a8\u6001"],
    ["\u4fe1\u606f\u53d1\u5e03", "\u56ed\u533a\u53d1\u5c55\u6210\u5c31\u3001\u4f18\u52bf\u4ea7\u4e1a\u3001\u5408\u4f5c\u6848\u4f8b\u591a\u7ef4\u5ea6\u5c55\u793a"]
  ];
  
  infoHubs.forEach(function(h, i) {
    var x = 0.7 + i * 2.2;
    rbox(sl, x, 1.2, 2.0, 0.45, C.pubLite, 0.06);
    sl.addShape(pres.shapes.OVAL, { x: x + 0.08, y: 1.28, w: 0.06, h: 0.06, fill: { color: C.pub } });
    sl.addText(h[0], { x: x + 0.2, y: 1.22, w: 1.7, h: 0.2, fontSize: 10, fontFace: FONT, color: C.pub, bold: true });
    sl.addText(h[1], { x: x + 0.08, y: 1.42, w: 1.85, h: 0.2, fontSize: 7.5, fontFace: FONT, color: C.text });
  });
  
  // Bottom - life services
  var lifeSvcs = [
    ["\u56ed\u533a\u9910\u996e", "\u6d4f\u89c86\u5bb6\u9910\u5385\u4fe1\u606f\u3001\u4eca\u65e5\u83dc\u5355\u3001\u4ef7\u683c\u4e0e\u4f4d\u7f6e\uff0c\u652f\u6301\u5206\u7c7b\u7b5b\u9009", C.pub],
    ["\u505c\u8f66\u670d\u52a1", "\u5730\u56fe\u67e5\u770b\u505c\u8f66\u4f4d\u5b9e\u65f6\u4f59\u4f4d\uff0c\u5728\u7ebf\u7f34\u8d39\uff0c\u652f\u6301\u591a\u8f66\u724c\u7ba1\u7406", "6D28D9"],
    ["\u573a\u9986\u670d\u52a1", "\u6d4f\u89c8\u56ed\u533a\u573a\u9986\u4fe1\u606f\uff0c\u67e5\u770b\u5f00\u653e\u65f6\u95f4\u4e0e\u5728\u7ebf\u9884\u7ea6", "5B21B6"],
    ["\u7279\u8272\u4f9b\u9700", "\u6d4f\u89c8\u9752\u5ddd\u5c71\u73cd\u7b49\u7279\u8272\u4ea7\u54c1\uff0c\u76f4\u63a5\u8054\u7cfb\u4f9b\u5e94\u5546", C.pub],
    ["\u7528\u5de5\u4fe1\u606f", "\u67e5\u770b\u56ed\u533a\u4f01\u4e1a\u62db\u8058\u4fe1\u606f\uff0c\u5c97\u4f4d\u5339\u914d\uff0c\u7b80\u5386\u6295\u9012", "6D28D9"]
  ];
  
  lifeSvcs.forEach(function(l, i) {
    var x = 0.5 + (i%3)*3.15, y = 2.4 + Math.floor(i/3)*1.35;
    card(sl, x, y, 2.9, 1.15, C.card);
    sbox(sl, x + 0.1, y + 0.1, 0.04, 0.95, l[2]);
    sl.addText(l[0], { x: x + 0.25, y: y + 0.12, w: 2.5, h: 0.25, fontSize: 12, fontFace: FONT, color: l[2], bold: true });
    sl.addText(l[1], { x: x + 0.25, y: y + 0.42, w: 2.5, h: 0.4, fontSize: 8.5, fontFace: FONT, color: C.text });
    rbox(sl, x + 0.25, y + 0.85, 0.9, 0.18, C.pubLite, 0.02);
    sl.addText("\u67e5\u770b\u8be6\u60c5", { x: x + 0.25, y: y + 0.85, w: 0.9, h: 0.18, fontSize: 6.5, fontFace: FONT, color: C.pub, align: "center", valign: "middle" });
  });
  
  badge(sl, pn, C.pub);
}

// ====== SLIDE 20: PUBLIC - ENGAGEMENT ======
function sPublicEngage() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.pub);
  
  sl.addText("\u516c\u4f17\u7528\u6237 \u2014 \u751f\u6d3b\u670d\u52a1\u4e0e\u4e92\u52a8\u4f53\u9a8c", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.text, bold: true });
  
  // Full width left: interactive features
  card(sl, 0.5, 0.8, 5.5, 3.8, C.card);
  sl.addText("\u516c\u4f17\u4e92\u52a8\u4e0e\u53c2\u4e0e", { x: 0.7, y: 0.88, w: 4.5, h: 0.3, fontSize: 14, fontFace: FONT, color: C.pub, bold: true });
  sbox(sl, 0.5, 1.15, 5.5, 0.01, C.pub);
  
  var interactives = [
    ["\u56ed\u533a\u5e7f\u573a\u5c55\u793a", "\u5927\u5c4f\u5ba3\u4f20\u4e0e\u4fe1\u606f\u53d1\u5e03\uff0c\u5c55\u793a\u56ed\u533a\u98ce\u91c7"],
    ["\u5fae\u4fe1\u5c0f\u7a0b\u5e8f", "\u5feb\u901f\u63a5\u5165\u667a\u6167\u5e73\u53f0\uff0c\u4fbf\u6377\u670d\u52a1\u5165\u53e3"],
    ["\u4f01\u4e1a\u7f51\u7ad9\u5bfc\u822a", "\u56ed\u533a\u4f01\u4e1a\u7f51\u7ad9\u96c6\u6210\u5bfc\u822a\uff0c\u4e00\u952e\u8bbf\u95ee"],
    ["\u54a8\u8be2\u6295\u8bc9", "\u5728\u7ebf\u54a8\u8be2\u3001\u6295\u8bc9\u53cd\u9988\uff0c\u5feb\u901f\u54cd\u5e94\u89e3\u51b3"],
    ["\u6570\u5b57\u5316\u5ba2\u670d", "\u667a\u80fd\u5ba2\u670d+\u4eba\u5de5\u5ba2\u670d\u53cc\u901a\u9053\uff0c24\u5c0f\u65f6\u5728\u7ebf"],
    ["\u793e\u533a\u4e92\u52a8", "\u56ed\u533a\u793e\u533a\u8bba\u575b\u3001\u610f\u89c1\u5f81\u96c6\u3001\u6ee1\u610f\u5ea6\u8c03\u67e5"]
  ];
  
  interactives.forEach(function(it, i) {
    var col = i % 2, row = Math.floor(i / 2);
    var x = 0.7 + col * 2.5, y = 1.3 + row * 0.8;
    rbox(sl, x, y, 2.3, 0.65, "F5F3FF", 0.06);
    sl.addShape(pres.shapes.OVAL, { x: x + 0.08, y: y + 0.08, w: 0.06, h: 0.06, fill: { color: C.pub } });
    sl.addText(it[0], { x: x + 0.2, y: y + 0.05, w: 2, h: 0.2, fontSize: 10, fontFace: FONT, color: C.pub, bold: true });
    sl.addText(it[1], { x: x + 0.08, y: y + 0.3, w: 2.1, h: 0.3, fontSize: 8, fontFace: FONT, color: C.text });
  });
  
  // Right: quick stats
  card(sl, 6.3, 0.8, 3.3, 3.8, "0F172A");
  sl.addText("\u7528\u6237\u670d\u52a1\u6570\u636e", { x: 6.45, y: 0.88, w: 3, h: 0.3, fontSize: 12, fontFace: FONT, color: "F59E0B", bold: true });
  sbox(sl, 6.3, 1.15, 3.3, 0.01, "1E293B");
  
  var usageData = [
    ["\u65e5\u5747\u6d3b\u52a8\u53c2\u4e0e", "120+\u4eba\u6b21"],
    ["\u65e5\u5747\u516c\u544a\u6d4f\u89c8", "500+\u6b21"],
    ["\u573a\u9986\u65e5\u5747\u9884\u7ea6", "80+\u6b21"],
    ["\u65e5\u5747\u7f34\u8d39\u4ea4\u6613", "60+\u7b14"],
    ["\u7528\u5de5\u4fe1\u606f\u6d4f\u89c8", "200+\u6b21"],
    ["\u65e5\u5747\u54a8\u8be2\u91cf", "40+\u6761"],
    ["\u6ee1\u610f\u5ea6\u8bc4\u5206", "4.8\u2605"]
  ];
  
  usageData.forEach(function(u, i) {
    var y = 1.3 + i * 0.47;
    sl.addText(u[0], { x: 6.5, y, w: 1.7, h: 0.25, fontSize: 8, fontFace: FONT, color: "94A3B8" });
    sl.addText(u[1], { x: 8.2, y: y - 0.02, w: 1.2, h: 0.28, fontSize: 13, fontFace: "Arial", color: C.accent, bold: true, align: "right" });
    sbox(sl, 6.5, y + 0.28, 2.9, 0.005, "1E293B");
  });
  
  badge(sl, pn, C.pub);
}

// ====== SLIDE 21: FULL FUNCTION MATRIX ======
function sFullMatrix() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.accent);
  
  sl.addText("\u5e73\u53f0\u5168\u529f\u80fd\u77e9\u9635\u4e00\u89c8", { x: 0.5, y: 0.12, w: 9, h: 0.45, fontSize: 18, fontFace: FONT, color: C.text, bold: true });
  sl.addText("\u5b8c\u6574\u529f\u80fd\u6a21\u5757\u4e0e\u56db\u7c7b\u89d2\u8272\u8986\u76d6\u5173\u7cfb", { x: 0.5, y: 0.52, w: 9, h: 0.2, fontSize: 8, fontFace: FONT, color: C.gray });
  
  var cols = ["\u529f\u80fd\u6a21\u5757", "\u653f\u5e9c", "\u7ba1\u7406", "\u4f01\u4e1a", "\u516c\u4f17", "\u529f\u80fd\u70b9\u6570"];
  var cw = [1.4, 1.7, 1.7, 1.7, 1.7, 0.8];
  var hCols = [C.text, C.gov, C.mgmt, C.ent, C.pub, C.gray];
  var cx = 0.3;
  
  // Header
  cols.forEach(function(c, i) {
    var x = (i===0) ? 0.3 : 0.3 + cw.slice(0,i).reduce(function(a,b){return a+b;}, 0);
    rbox(sl, x, 0.75, cw[i], 0.25, hCols[i], 0.03);
    sl.addText(c, { x, y: 0.75, w: cw[i], h: 0.25, fontSize: 7, fontFace: FONT, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  });
  
  var matrixRows = [
    ["AI\u62db\u5546\u670d\u52a1", "\u25cf", "\u25cb", "\u25cb", "\u2014", "5"],
    ["\u6570\u5b57\u6c99\u76d8", "\u25cf", "\u2014", "\u25cb", "\u2014", "3"],
    ["\u4ea7\u4e1a\u94fe\u56fe\u8c31", "\u25cf", "\u25cb", "\u25cb", "\u2014", "4"],
    ["\u76ee\u6807\u4f01\u4e1a\u6316\u6398", "\u25cf", "\u25cb", "\u2014", "\u2014", "2"],
    ["\u6295\u8d44\u9879\u76ee\u7ba1\u7406", "\u25cf", "\u25cb", "\u25cf", "\u2014", "4"],
    ["\u62db\u5546\u5165\u53e3/\u4e91\u5c55\u9986", "\u2014", "\u25cb", "\u25cf", "\u25cb", "5"],
    ["\u5b89\u5168\u76d1\u63a7(\u5929\u773c)", "\u25cf", "\u25cf", "\u25cb", "\u25cf", "5"],
    ["\u5730\u707e\u76d1\u6d4b(\u54e8\u5175)", "\u25cf", "\u25cf", "\u25cb", "\u25cb", "4"],
    ["\u73af\u4fdd\u76d1\u6d4b(\u6167\u773c)", "\u25cf", "\u25cf", "\u25cb", "\u25cb", "4"],
    ["\u89c6\u9891\u76d1\u63a7\u603b\u89c8", "\u25cf", "\u25cf", "\u2014", "\u25cf", "3"],
    ["\u5e94\u6025\u6307\u6325\u8054\u52a8", "\u25cf", "\u25cf", "\u25cb", "\u25cf", "6"],
    ["\u8bbe\u5907\u7ba1\u7406", "\u25cb", "\u25cf", "\u2014", "\u2014", "2"],
    ["\u8bc9\u6c42\u76f4\u8fbe", "\u25cf", "\u25cf", "\u25cf", "\u25cb", "5"],
    ["\u670d\u52a1\u76f4\u8fbe", "\u2014", "\u25cb", "\u25cf", "\u25cb", "4"],
    ["\u7f34\u8d39\u670d\u52a1(\u6c14/\u6c34/\u7535)", "\u2014", "\u25cb", "\u25cf", "\u2014", "3"],
    ["\u56ed\u4f01\u4ea4\u4e92", "\u25cf", "\u25cf", "\u25cf", "\u25cb", "5"],
    ["\u4f9b\u9700\u5bf9\u63a5", "\u25cf", "\u25cf", "\u25cf", "\u25cb", "5"],
    ["\u653f\u7b56\u63a8\u9001", "\u25cf", "\u25cf", "\u25cf", "\u25cb", "5"],
    ["\u573a\u9986/\u6d3b\u52a8/\u9910\u996e", "\u2014", "\u25cf", "\u25cf", "\u25cf", "6"],
    ["\u505c\u8f66\u670d\u52a1", "\u2014", "\u25cf", "\u25cf", "\u25cf", "4"],
    ["\u7528\u5de5\u4fe1\u606f", "\u25cf", "\u25cf", "\u25cf", "\u25cf", "5"],
    ["\u80fd\u6548\u76d1\u6d4b\u5206\u6790", "\u25cf", "\u25cf", "\u25cb", "\u2014", "4"],
    ["\u4ea9\u5747\u8bba\u82f1\u96c4", "\u25cf", "\u25cf", "\u25cb", "\u2014", "3"],
    ["\u6570\u636e\u4e2d\u5fc3", "\u25cf", "\u25cf", "\u25cb", "\u2014", "3"],
    ["\u6570\u5b57\u5316\u529e\u516c", "\u25cb", "\u25cf", "\u25cf", "\u25cb", "6"],
    ["\u8bbf\u5ba2/\u95e8\u7981\u7ba1\u7406", "\u25cb", "\u25cf", "\u25cf", "\u25cb", "5"],
    ["\u4e91\u5c55\u9986/\u7535\u5b50\u540d\u7247", "\u2014", "\u25cb", "\u25cf", "\u25cb", "4"],
    ["\u56ed\u533a\u5ba3\u4f20/\u516c\u544a", "\u25cf", "\u25cf", "\u25cf", "\u25cf", "6"]
  ];
  
  matrixRows.forEach(function(row, ri) {
    var y = 1.05 + ri * 0.3;
    var bg = ri % 2 === 0 ? C.card : C.bg;
    
    row.forEach(function(cell, ci) {
      var x = (ci===0) ? 0.3 : 0.3 + cw.slice(0,ci).reduce(function(a,b){return a+b;}, 0);
      sl.addShape(pres.shapes.RECTANGLE, { x, y, w: cw[ci], h: 0.3, fill: { color: bg } });
      
      if (ci === 0) {
        sl.addText(cell, { x: x + 0.05, y, w: cw[ci] - 0.1, h: 0.3, fontSize: 6.5, fontFace: FONT, color: C.text, valign: "middle" });
      } else if (ci < 5) {
        var dotColor = cell === "\u25cf" ? hCols[ci] : (cell === "\u25cb" ? C.light : "CBD5E1");
        sl.addText(cell, { x, y, w: cw[ci], h: 0.3, fontSize: 14, fontFace: "Arial", color: dotColor, align: "center", valign: "middle" });
      } else {
        sl.addText(cell, { x, y, w: cw[ci], h: 0.3, fontSize: 7, fontFace: FONT, color: C.gray, align: "center", valign: "middle" });
      }
    });
  });
  
  // Legend
  var ly = 1.05 + matrixRows.length * 0.3 + 0.05;
  rbox(sl, 0.5, ly, 9, 0.2, C.light, 0.04);
  sl.addText("\u25cf \u4e3b\u8981\u529f\u80fd  \u25cb \u8f85\u52a9\u529f\u80fd  \u2014 \u65e0\u76f4\u63a5\u5173\u8054    \u603b\u8ba1" + (28+10) + "\u4e2a\u529f\u80fd\u70b9\u8986\u76d6\u56db\u5927\u89d2\u8272",
    { x: 0.7, y: ly, w: 8.6, h: 0.2, fontSize: 8, fontFace: FONT, color: C.gray, valign: "middle" });
  
  badge(sl, pn, C.accent);
}

// ====== SLIDE 22: INVESTMENT ======
function sInvest() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.bg };
  topBar(sl, C.accent);
  
  sl.addText("\u6295\u8d44\u4e0e\u6548\u76ca\u5206\u6790", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 22, fontFace: FONT, color: C.text, bold: true });
  
  // Total investment card
  card(sl, 0.5, 0.75, 9, 1.2, "0F172A");
  sl.addText("\u56ed\u533a\u667a\u6167\u5316\u5efa\u8bbe\u603b\u6295\u8d44\u6982\u7b97", { x: 0.7, y: 0.85, w: 5, h: 0.25, fontSize: 13, fontFace: FONT, color: "F59E0B" });
  sl.addText("\u00a54,178,800.00", { x: 0.7, y: 1.1, w: 4, h: 0.7, fontSize: 36, fontFace: "Arial", color: "FFFFFF", bold: true });
  sl.addText("\u5143", { x: 4.8, y: 1.3, w: 1, h: 0.5, fontSize: 14, fontFace: FONT, color: "94A3B8", valign: "bottom" });
  
  // Investment breakdown
  var investItems = [
    ["\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7", "1,580,000", 38],
    ["\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7", "428,600", 10],
    ["\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546", "442,200", 11],
    ["\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", "348,000", 8],
    ["\u667a\u6539\u6570\u8f6c\u8d4b\u80fd\u5e73\u53f0", "280,000", 7],
    ["\u4e00\u4f53\u5316\u6570\u5b57\u5e73\u53f0", "720,000", 17],
    ["\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3", "380,000", 9]
  ];
  
  investItems.forEach(function(item, i) {
    var y = 2.2 + i * 0.38;
    sl.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.08, w: 0.06, h: 0.06, fill: { color: C.accent } });
    sl.addText(item[0], { x: 0.8, y: y, w: 3.8, h: 0.3, fontSize: 10, fontFace: FONT, color: C.text, valign: "middle" });
    sl.addText("\u00a5" + item[1], { x: 4.6, y: y, w: 1.5, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.ent, bold: true, align: "right", valign: "middle" });
    // Progress bar
    rbox(sl, 6.2, y + 0.08, 2.5, 0.14, C.light, 0.07);
    rbox(sl, 6.2, y + 0.08, 2.5 * item[2] / 100, 0.14, C.accent, 0.07);
    sl.addText(String(item[2]) + "%", { x: 8.8, y: y, w: 0.7, h: 0.3, fontSize: 8, fontFace: "Arial", color: C.gray, align: "right", valign: "middle" });
  });
  
  // Benefits section
  sbox(sl, 0.5, 5.0, 9, 0.03, C.light);
  var benefits = [
    ["\u7ba1\u7406\u6548\u7387\u63d0\u534730%", "1E40AF"],
    ["\u5e94\u6025\u54cd\u5e94\u63d0\u534760%", "059669"],
    ["\u00a520\u4ebf+\u5e74\u5bf9\u63a5\u4e13\u9879\u8d44\u91d1", "D97706"],
    ["\u5168\u56ed\u533a\u76d1\u63a7\u8986\u76d6\u738795%", "7C3AED"]
  ];
  benefits.forEach(function(b, i) {
    var x = 0.5 + i * 2.35;
    sl.addText(b[0], { x: x, y: 5.2, w: 2.2, h: 0.3, fontSize: 11, fontFace: FONT, color: b[1], bold: true, align: "center", valign: "middle" });
  });
  
  badge(sl, pn, C.accent);
}

// ====== SLIDE 23: FUTURE ======
function sFuture() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.dark };
  
  // Decorative
  sl.addShape(pres.shapes.OVAL, { x: 6, y: -1, w: 5, h: 5, fill: { color: "1E3A5F", transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -2, y: 2, w: 4, h: 4, fill: { color: "1E3A5F", transparency: 60 } });
  
  sl.addText("\u672a\u6765\u5c55\u671b", { x: 0.5, y: 0.12, w: 9, h: 0.5, fontSize: 24, fontFace: FONT, color: "FFFFFF", bold: true });
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 0.5, y: 0.55, w: 9, h: 0.3, fontSize: 13, fontFace: FONT, color: C.accent });
  
  var futureItems = [
    ["AIoT\u6df1\u5ea6\u878d\u5408", "\u5f15\u5165\u7269\u8054\u7f51\u4f20\u611f\u8bbe\u5907\uff0c\u5b9e\u73b0\u4e07\u7269\u4e92\u8054\u3001\u5168\u57df\u611f\u77e5\u4e0e\u667a\u80fd\u63a7\u5236", C.gov],
    ["\u6570\u5b57\u5b6a\u751f\u56ed\u533a", "1:1\u4e09\u7ef4\u5efa\u6a21\uff0c\u5168\u8981\u7d20\u3001\u5168\u65f6\u57df\u53ef\u89c6\u5316\u8fd0\u8425\u7ba1\u7406\uff0c\u6c89\u6d78\u5f0f\u7ba1\u63a7", C.mgmt],
    ["AI Agent\u667a\u80fd\u5316", "\u667a\u80fd\u5ba2\u670d\u3001\u667a\u80fd\u8c03\u5ea6\u3001\u667a\u80fd\u9884\u8b66\uff0cAI\u9a71\u52a8\u56ed\u533a\u81ea\u52a8\u5316\u8fd0\u8425", C.ent],
    ["\u4ea7\u4e1a\u5927\u8111", "\u57fa\u4e8e\u5927\u6570\u636e\u4e0eAI\u7b97\u6cd5\u8f85\u52a9\u4ea7\u4e1a\u62db\u5546\u51b3\u7b56\uff0c\u6570\u636e\u9a71\u52a8\u53d1\u5c55", C.pub],
    ["\u8de8\u56ed\u533a\u534f\u540c", "\u5bf9\u63a5\u7701\u5e02\u667a\u6167\u5e73\u53f0\uff0c\u5b9e\u73b0\u6570\u636e\u5171\u4eab\u4e0e\u4e1a\u52a1\u534f\u540c\uff0c\u6784\u5efa\u751f\u6001\u4f53\u7cfb", C.accent]
  ];
  
  futureItems.forEach(function(f, i) {
    var y = 1.0 + i * 0.75;
    card(sl, 0.5, y, 9, 0.6, "1E293B");
    sl.addShape(pres.shapes.PARALLELOGRAM, { x: 0.6, y: y + 0.1, w: 0.5, h: 0.4, fill: { color: f[2] } });
    sl.addText(String(i+1), { x: 0.6, y: y + 0.1, w: 0.5, h: 0.4, fontSize: 16, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(f[0], { x: 1.3, y: y + 0.05, w: 3, h: 0.25, fontSize: 12, fontFace: FONT, color: f[2], bold: true });
    sl.addText(f[1], { x: 1.3, y: y + 0.3, w: 7.8, h: 0.25, fontSize: 9, fontFace: FONT, color: "94A3B8" });
  });
  
  // Bottom
  sbox(sl, 0.5, 4.8, 9, 0.02, "1E293B");
  sl.addText("\u9a71\u52a8\u56ed\u533a\u6cbf\u7740\u201c\u7701\u7ea7\u5f00\u53d1\u533a\u2192\u7eff\u8272\u56ed\u533a\u2192\u56fd\u5bb6\u7ea7\u7ecf\u5f00\u533a\u57f9\u80b2\u5bf9\u8c61\u201d\u7684\u822a\u5411\u5168\u901f\u524d\u8fdb",
    { x: 0.5, y: 5.0, w: 9, h: 0.3, fontSize: 10, fontFace: FONT, color: "94A3B8", italic: true, align: "center" });
}

// ====== SLIDE 24: THANK YOU ======
function sThank() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: C.dark };
  
  sl.addShape(pres.shapes.OVAL, { x: 2, y: 0, w: 6, h: 6, fill: { color: "1E3A5F", transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: 5, y: -2, w: 4, h: 4, fill: { color: C.gov, transparency: 85 } });
  
  sl.addText("\u611f\u8c22\u804a\u542c", { x: 1, y: 1.2, w: 8, h: 1.0, fontSize: 40, fontFace: FONT, color: "FFFFFF", bold: true, align: "center" });
  rbox(sl, 4.2, 2.2, 1.6, 0.04, C.accent, 0.02);
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 1, y: 2.4, w: 8, h: 0.5, fontSize: 16, fontFace: FONT, color: C.accent, align: "center" });
  
  var footers = [
    "\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u7ba1\u59d4\u4f1a",
    "\u6280\u672f\u652f\u6301\uff1a\u9752\u8fc5\u79d1\u6280",
    "\u5730\u5740\uff1a\u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf\u7af9\u56ed\u9547"
  ];
  footers.forEach(function(f, i) {
    sl.addText(f, { x: 1, y: 3.5 + i * 0.35, w: 8, h: 0.3, fontSize: 11, fontFace: FONT, color: "94A3B8", align: "center" });
  });
}

// ====== BUILD ======
sCover();
sTOC();
sBackground();
sArchitecture();
sRoleMatrix();
sGovDashboard();
sGovRecruit();
sGovInvest();
sGovSecurity();
sMgmtDashboard();
sMgmtOps();
sMgmtEnergy();
sMgmtDevice();
sEntService();
sEntLogistics();
sEntVenue();
sEntPayment();
sEntExhibit();
sPublicInfo();
sPublicEngage();
sFullMatrix();
sInvest();
sFuture();
sThank();

pres.writeFile({ fileName: path.join(__dirname, "output", "zhuangzishang_detailed_ppt.pptx") })
  .then(function() { console.log("PPT OK: " + pn + " slides"); })
  .catch(function(e) { console.error("ERR:", e); });
