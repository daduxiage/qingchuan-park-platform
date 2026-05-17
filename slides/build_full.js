
const pptxgen = require("pptxgenjs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

const T = {
  primary: "0A1628", secondary: "1B3A5C", accent: "2D8CF0",
  highlight: "00D4AA", light: "E8F0FE", bg: "FFFFFF",
  gray: "6B7A8D", darkText: "1A2332", cardBg: "F0F4F8"
};

function addBadge(slide, num) {
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: T.accent } });
  slide.addText(String(num), { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 11, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
}

function addTopBar(slide) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: T.accent } });
}

function addSectionAccent(slide, x, y, h) {
  slide.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.06, h: h, fill: { color: T.accent } });
}

function addCard(slide, x, y, w, h, opts) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: T.cardBg }, rectRadius: 0.08, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.06 } });
  if (opts.accentColor) {
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.05, fill: { color: opts.accentColor } });
  }
}

function addRoleCard(slide, x, y, w, h, icon, title, desc, color) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: T.secondary }, rectRadius: 0.08 });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.05, fill: { color } });
  slide.addText(icon, { x, y: y + 0.15, w, h: 0.45, fontSize: 22, align: "center" });
  slide.addText(title, { x, y: y + 0.6, w, h: 0.25, fontSize: 12, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center" });
  slide.addText(desc, { x: x + 0.1, y: y + 0.9, w: w - 0.2, h: h - 1.1, fontSize: 9, fontFace: "Microsoft YaHei", color: T.light, align: "center" });
}

let pn = 0;

// === COVER ===
function s1() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: T.primary };
  addTopBar(sl);
  sl.addShape(pres.shapes.OVAL, { x: 7.5, y: -1, w: 4, h: 4, fill: { color: T.secondary, transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -1.5, y: 3.5, w: 3, h: 3, fill: { color: T.secondary, transparency: 60 } });
  addSectionAccent(sl, 0.6, 1.0, 2.0);
  sl.addText("\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a", { x: 0.9, y: 1.0, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: T.highlight });
  sl.addText("\u667a\u6167\u56ed\u533a\u5efa\u8bbe\u9879\u76ee", { x: 0.9, y: 1.5, w: 8, h: 1.2, fontSize: 38, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addText("\u5efa\u8bbe\u65b9\u6848\u6c47\u62a5", { x: 0.9, y: 2.7, w: 6, h: 0.5, fontSize: 22, fontFace: "Microsoft YaHei", color: T.light });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 3.4, w: 2, h: 0.04, fill: { color: T.accent } });
  var aud = ["\u653f\u5e9c\u9886\u5bfc\u5c42", "\u56ed\u533a\u7ba1\u7406\u5c42", "\u5165\u9a7b\u4f01\u4e1a", "\u516c\u4f17\u7528\u6237"];
  aud.forEach(function(a, i) {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.9 + i * 2.1, y: 3.8, w: 1.9, h: 0.35, fill: { color: T.secondary }, rectRadius: 0.05 });
    sl.addText(a, { x: 0.9 + i * 2.1, y: 3.8, w: 1.9, h: 0.35, fontSize: 9, fontFace: "Microsoft YaHei", color: T.light, align: "center", valign: "middle" });
  });
  sl.addText("2026\u5e745\u6708", { x: 0.6, y: 5.0, w: 3, h: 0.4, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: T.accent } });
}

// === TOC ===
function s2() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: T.primary }; addTopBar(sl);
  sl.addText("\u6c47\u62a5\u76ee\u5f55", { x: 0.8, y: 0.3, w: 8, h: 0.6, fontSize: 24, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 0.9, w: 1.2, h: 0.04, fill: { color: T.accent } });
  var items = ["\u9879\u76ee\u6218\u7565\u5b9a\u4f4d", "\u4e09\u6b65\u8d70\u521b\u5efa\u8def\u5f84", "\u4e03\u5927\u5efa\u8bbe\u5de5\u7a0b", "\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7", "\u73af\u4fdd\u76d1\u63a7\u5de5\u7a0b", "AI\u62db\u5546\u4e0e\u5f62\u8c61\u5c55\u793a", "\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", "\u667a\u6539\u6570\u8f6c\u8d4b\u80fd\u5e73\u53f0", "\u4e00\u4f53\u5316\u6570\u5b57\u5e73\u53f0", "\u4f01\u4e1a\u670d\u52a1\u677f\u5757", "\u4fe1\u606f\u53d1\u5e03\u5b50\u7cfb\u7edf", "\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3", "\u6295\u8d44\u6982\u7b97", "\u9884\u671f\u6548\u76ca\u5206\u6790", "\u56db\u7c7b\u89d2\u8272\u4ef7\u503c", "\u672a\u6765\u5c55\u671b"];
  items.forEach(function(item, i) {
    var col = i % 2, row = Math.floor(i / 2);
    var x = 0.8 + col * 4.5, y = 1.3 + row * 0.55;
    sl.addShape(pres.shapes.OVAL, { x, y: y + 0.06, w: 0.28, h: 0.28, fill: { color: T.accent } });
    sl.addText(String(i + 1), { x, y: y + 0.06, w: 0.28, h: 0.28, fontSize: 10, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(item, { x: x + 0.38, y, w: 3.8, h: 0.4, fontSize: 12, fontFace: "Microsoft YaHei", color: "FFFFFF", valign: "middle" });
  });
  addBadge(sl, pn);
}

// === Slide 3: Strategy ===
function s3() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" }; addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("\u9879\u76ee\u6218\u7565\u5b9a\u4f4d", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("\u667a\u6167\u5316\u5efa\u8bbe\u662f\u56ed\u533a\u5b9e\u73b0\u8de8\u8d8a\u5f0f\u53d1\u5c55\u7684\u5fc5\u7531\u4e4b\u8def\u4e0e\u5173\u952e\u4e00\u8dc3", { x: 0.75, y: 0.8, w: 8.5, h: 0.4, fontSize: 12, fontFace: "Microsoft YaHei", color: T.accent });
  
  var cards = [
    { t: "\u8d4b\u80fd\u521b\u5efa", d: "\u4e3a\u7533\u62a5\u7701\u7ea7\u7ecf\u5f00\u533a\u3001\u7eff\u8272\u5de5\u4e1a\u56ed\u533a\u63d0\u4f9b\u6700\u76f4\u63a5\u7684\u6570\u636e\u652f\u6491\u4e0e\u7ba1\u7406\u5de5\u5177", c: T.accent },
    { t: "\u8d4b\u80fd\u4f01\u4e1a", d: "\u6253\u9020\u53bf\u57df\u667a\u6539\u6570\u8f6c\u516c\u5171\u670d\u52a1\u5e73\u53f0\uff0c\u7834\u89e3\u4e2d\u5c0f\u4f01\u4e1a\u8f6c\u578b\u96be\u9898", c: T.highlight },
    { t: "\u8d4b\u80fd\u51b3\u7b56", d: "\u6784\u5efa\u56ed\u533a\u6570\u5b57\u5927\u8111\uff0c\u5b9e\u73b0\u8fd0\u884c\u4e00\u5c4f\u7edf\u89c8\u3001\u4ea7\u4e1a\u4e00\u56fe\u7edf\u7ba1", c: "F59E0B" }
  ];
  cards.forEach(function(card, i) {
    var x = 0.5 + i * 3.1;
    addCard(sl, x, 1.4, 2.8, 2.2, { accentColor: card.c });
    sl.addText(card.t, { x: x + 0.2, y: 1.65, w: 2.4, h: 0.4, fontSize: 15, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
    sl.addText(card.d, { x: x + 0.2, y: 2.15, w: 2.4, h: 1.2, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  });

  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.9, w: 9, h: 0.04, fill: { color: T.light } });
  sl.addText("2025\u5e74\u5b9e\u73b0\u89c4\u4e0a\u5de5\u4e1a\u603b\u4ea7\u503c", { x: 0.5, y: 4.1, w: 5, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  sl.addText("41.6\u4ebf\u5143", { x: 4.0, y: 4.1, w: 2, h: 0.3, fontSize: 14, fontFace: "Microsoft YaHei", color: T.accent, bold: true });
  sl.addText("\u603b\u6295\u8d44\u6982\u7b97\uff1a\u00a54,178,800.00", { x: 0.5, y: 4.5, w: 6, h: 0.3, fontSize: 12, fontFace: "Microsoft YaHei", color: T.secondary, bold: true });
  addBadge(sl, pn);
}

// === Slide 4: Three Steps ===
function s4() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: T.primary }; addTopBar(sl);
  sl.addText("\u4e09\u6b65\u8d70\u521b\u5efa\u8def\u5f84", { x: 0.7, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  var steps = [
    { t: "\u7b2c\u4e00\u6b65", s: "\u7701\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a", d: "\u4ee5\u667a\u6167\u7ba1\u7406\u4e3a\u4eae\u70b9\uff0c\u4ea9\u5747\u8bba\u82f1\u96c4\u8bc4\u4ef7\u4f53\u7cfb\u3001\u5b89\u5168\u73af\u4fdd\u76d1\u63a7\u7f51\u7edc\u63d0\u4f9b\u771f\u5b9e\u6570\u636e\u4f50\u8bc1", c: T.accent },
    { t: "\u7b2c\u4e8c\u6b65", s: "\u7701\u7ea7/\u56fd\u5bb6\u7ea7\u7eff\u8272\u5de5\u4e1a\u56ed\u533a", d: "\u4ee5\u7eff\u8272\u6570\u636e\u4e3a\u5b9e\u8bc1\uff0c\u5168\u8986\u76d6\u80fd\u8017\u4e0e\u6c61\u67d3\u6e90\u76d1\u6d4b\u4f53\u7cfb\u81ea\u52a8\u751f\u6210\u7eff\u8272\u7ee9\u6548\u62a5\u544a", c: T.highlight },
    { t: "\u7b2c\u4e09\u6b65", s: "\u56fd\u5bb6\u7ea7\u7ecf\u6d4e\u6280\u672f\u5f00\u53d1\u533a", d: "\u4ee5\u521b\u65b0\u751f\u6001\u4e3a\u9a71\u52a8\uff0c\u667a\u6167\u62db\u5546\u7cfb\u7edf\u7cbe\u51c6\u8865\u5f3a\u5916\u8d44\u3001\u5916\u8d38\u3001\u7814\u53d1\u7b49\u77ed\u677f\u6307\u6807", c: "F59E0B" }
  ];
  steps.forEach(function(s, i) {
    var y = 1.1 + i * 1.4;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y, w: 8.6, h: 1.15, fill: { color: T.secondary }, rectRadius: 0.08 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.06, h: 1.15, fill: { color: s.c } });
    sl.addShape(pres.shapes.OVAL, { x: 1.1, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: s.c } });
    sl.addText(s.t, { x: 1.1, y: y + 0.3, w: 0.55, h: 0.55, fontSize: 10, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(s.s, { x: 1.85, y: y + 0.1, w: 3.5, h: 0.4, fontSize: 14, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
    sl.addText(s.d, { x: 1.85, y: y + 0.55, w: 7, h: 0.5, fontSize: 10, fontFace: "Microsoft YaHei", color: T.light });
    if (i < 2) sl.addShape(pres.shapes.LINE, { x: 0.97, y: y + 1.15, w: 0, h: 0.25, line: { color: T.accent, width: 2 } });
  });
  addBadge(sl, pn);
}

// === Slide 5: Seven Projects ===
function s5() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" }; addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("\u4e03\u5927\u5efa\u8bbe\u5de5\u7a0b", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("\u8986\u76d6\u62db\u5546\u3001\u4f01\u4e1a\u3001\u8fd0\u8425\u3001\u7ba1\u7406\u5168\u4e1a\u52a1\u7684\u4e00\u4f53\u5316\u667a\u6167\u56ed\u533a\u5e73\u53f0", { x: 0.75, y: 0.8, w: 8, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  var bgColors = [T.accent, T.highlight, "F59E0B", "EB5757", "9B51E0", "2D9CDB", "219653"];
  var projects = [
    { n: "01", t: "\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7", d: "\u667a\u80fd\u6444\u50cf\u673a+\u5730\u8d28\u4f20\u611f\u5668\uff0c\u5168\u8986\u76d6\u611f\u77e5\u7f51\u7edc" },
    { n: "02", t: "\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7", d: "\u6c34\u8d28\u7a7a\u6c14\u5728\u7ebf\u76d1\u6d4b\uff0c\u76d1\u6d4b-\u9884\u8b66-\u6eaf\u6e90\u95ed\u73af" },
    { n: "03", t: "\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546", d: "\u4e09\u7ef4\u6570\u5b57\u6c99\u76d8+\u6570\u636e\u9a7e\u9a76\u8231+AI\u62db\u5546\u5927\u8111" },
    { n: "04", t: "\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", d: "\u667a\u80fd\u8ba1\u91cf\u8868\u8ba1+\u4ea9\u5747\u8bba\u82f1\u96c4\u5927\u6570\u636e\u8bc4\u4ef7" },
    { n: "05", t: "\u667a\u6539\u6570\u8f6c\u4f01\u4e1a\u8d4b\u80fd\u5e73\u53f0", d: "\u653f\u7b56\u63a8\u9001+\u4f9b\u9700\u5bf9\u63a5+\u6570\u5b57\u5316\u8bca\u65ad\u5de5\u5177" },
    { n: "06", t: "\u4e00\u4f53\u5316\u6570\u5b57\u5e73\u53f0", d: "\u7269\u8054\u7f51\u5e73\u53f0+\u6570\u636e\u4e2d\u53f0+AI\u4e2d\u53f0+\u6570\u5b57\u5b6a\u751f" },
    { n: "07", t: "\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3", d: "\u5927\u5c4f\u6307\u6325+\u56ed\u5c0f\u4e8c\u56e2\u961f+3\u5e74\u957f\u6548\u670d\u52a1" }
  ];
  projects.forEach(function(p, i) {
    var col = i % 3, row = Math.floor(i / 3);
    var x = 0.5 + col * 3.1, y = 1.3 + row * 1.3;
    addCard(sl, x, y, 2.8, 1.1, { accentColor: bgColors[i] });
    sl.addText(p.n, { x: x + 0.15, y: y + 0.15, w: 0.35, h: 0.35, fontSize: 16, fontFace: "Arial", color: bgColors[i], bold: true });
    sl.addText(p.t, { x: x + 0.55, y: y + 0.12, w: 2.1, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
    sl.addText(p.d, { x: x + 0.15, y: y + 0.55, w: 2.5, h: 0.45, fontSize: 9, fontFace: "Microsoft YaHei", color: T.gray });
  });
  addBadge(sl, pn);
}

// === Detail slide helper ===
function sDetail(title, items, accentColor) {
  pn++; const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" }; addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText(title, { x: 0.75, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: 0.8, w: 1.0, h: 0.03, fill: { color: accentColor } });
  items.forEach(function(item, i) {
    var y = 1.15 + i * 0.7;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.6, fill: { color: T.cardBg }, rectRadius: 0.06 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.05, h: 0.6, fill: { color: accentColor } });
    var icons = ["\u2776", "\u2777", "\u2778", "\u2779", "\u277a"];
    sl.addText(icons[i] || "\u2022", { x: 0.7, y, w: 0.35, h: 0.6, fontSize: 12, fontFace: "Microsoft YaHei", color: accentColor, align: "center", valign: "middle" });
    if (typeof item === "string") {
      sl.addText(item, { x: 1.15, y, w: 8.1, h: 0.6, fontSize: 12, fontFace: "Microsoft YaHei", color: T.darkText, valign: "middle" });
    } else {
      sl.addText(item.t, { x: 1.15, y, w: 4, h: 0.28, fontSize: 12, fontFace: "Microsoft YaHei", color: T.primary, bold: true, valign: "middle" });
      sl.addText(item.d, { x: 1.15, y: y + 0.28, w: 8.1, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: T.gray, valign: "middle" });
    }
  });
  addBadge(sl, pn);
}

// === Integrated Platform Slide ===
function sPlatform() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: T.primary }; addTopBar(sl);
  sl.addText("\u4e00\u4f53\u5316\u6570\u5b57\u5e73\u53f0 - \u56db\u5927\u4e1a\u52a1\u7ebf", { x: 0.7, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addText("\u7269\u8054\u7f51\u5e73\u53f0 | \u6570\u636e\u4e2d\u53f0 | AI\u4e2d\u53f0 | \u6570\u5b57\u5b6a\u751f\u5e95\u5ea7", { x: 0.7, y: 0.8, w: 8, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: T.highlight });
  var biz = [
    { t: "\u4f01\u4e1a\u670d\u52a1\u677f\u5757", items: ["\u8bc9\u6c42\u76f4\u8fbe\uff08\u968f\u624b\u62cd\uff09", "\u670d\u52a1\u76f4\u8fbe", "\u56ed\u4f01\u4ea4\u4e92", "\u4f9b\u9700\u5bf9\u63a5"], c: T.accent },
    { t: "\u56ed\u533a\u8fd0\u8425\u677f\u5757", items: ["\u6570\u5b57\u5316\u5ba1\u6279", "\u65e5\u5e38\u5de1\u68c0", "\u8d44\u4ea7\u7ba1\u7406", "\u6863\u6848\u7ba1\u7406"], c: T.highlight },
    { t: "\u4fe1\u606f\u53d1\u5e03\u677f\u5757", items: ["\u7528\u5de5\u00b7\u7269\u6d41\u00b7\u7279\u8272\u4f9b\u9700", "\u653f\u7b56\u63a8\u9001\u00b7\u9910\u996e\u00b7\u505c\u8f66", "\u573a\u9986\u670d\u52a1\u00b7\u56ed\u533a\u6d3b\u52a8", "\u56ed\u533a\u516c\u544a"], c: "F59E0B" },
    { t: "\u6570\u636e\u652f\u6491\u677f\u5757", items: ["12\u5927\u4e09\u65b9\u7cfb\u7edf\u63a5\u5165", "\u7edf\u4e00\u6570\u636e\u6807\u51c6", "\u7edf\u4e00\u8ba4\u8bc1\u4e2d\u5fc3", "\u7edf\u4e00\u6d88\u606f\u63a8\u9001"], c: "EB5757" }
  ];
  biz.forEach(function(b, i) {
    var x = 0.5 + i * 2.35;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.3, w: 2.15, h: 3.5, fill: { color: T.secondary }, rectRadius: 0.08 });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.15, h: 0.4, fill: { color: b.c } });
    sl.addText(b.t, { x, y: 1.3, w: 2.15, h: 0.4, fontSize: 11, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    b.items.forEach(function(item, j) {
      sl.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 2.0 + j * 0.6, w: 0.08, h: 0.08, fill: { color: b.c } });
      sl.addText(item, { x: x + 0.4, y: 1.9 + j * 0.6, w: 1.6, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: "FFFFFF" });
    });
  });
  addBadge(sl, pn);
}

// === Investment Slide ===
function sInvestment() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" }; addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("\u9879\u76ee\u6295\u8d44\u6982\u7b97", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.5, y: 1.2, w: 5, h: 1.5, fill: { color: T.primary }, rectRadius: 0.1 });
  sl.addText("\u603b\u6295\u8d44\u6982\u7b97", { x: 2.5, y: 1.35, w: 5, h: 0.35, fontSize: 14, fontFace: "Microsoft YaHei", color: T.light, align: "center" });
  sl.addText("\u00a54,178,800.00", { x: 2.5, y: 1.7, w: 5, h: 0.8, fontSize: 32, fontFace: "Arial", color: T.highlight, bold: true, align: "center" });
  var details = ["\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7\u5de5\u7a0b", "\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7\u5de5\u7a0b", "\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546\u5de5\u7a0b", "\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b\u5de5\u7a0b", "\u667a\u6539\u6570\u8f6c\u8d4b\u80fd\u5e73\u53f0", "\u4e00\u4f53\u5316\u6570\u5b57\u5e73\u53f0", "\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3"];
  details.forEach(function(d, i) {
    var y = 3.0 + i * 0.3;
    sl.addShape(pres.shapes.OVAL, { x: 1.5, y: y + 0.08, w: 0.08, h: 0.08, fill: { color: T.accent } });
    sl.addText(d, { x: 1.75, y, w: 6, h: 0.28, fontSize: 10, fontFace: "Microsoft YaHei", color: T.darkText });
  });
  addBadge(sl, pn);
}

// === Benefits Slide ===
function sBenefits() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" }; addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("\u9884\u671f\u6548\u76ca\u5206\u6790", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  var metrics = [
    { n: "30%", l: "\u7ba1\u7406\u6548\u7387\u63d0\u5347", d: "\u7efc\u5408\u7ba1\u7406\u6548\u7387\u63d0\u534730%\u4ee5\u4e0a", c: T.accent },
    { n: "60%", l: "\u5e94\u6025\u54cd\u5e94\u63d0\u5347", d: "\u5b89\u5168\u73af\u4fdd\u4e8b\u4ef6\u4ece\u88ab\u52a8\u54cd\u5e94\u8f6c\u5411\u4e3b\u52a8\u9884\u8b66\uff0c\u6548\u7387\u63d0\u534760%", c: T.highlight },
    { n: "\u00a520\u4ebf+", l: "\u4e13\u9879\u8d44\u91d1\u5bf9\u63a5", d: "\u5bf9\u63a5\u7701\u7ea7\u667a\u6539\u6570\u8f6c\u4e13\u9879\u8d44\u91d1\u53ca\u91d1\u878d\u4ea7\u54c1", c: "F59E0B" },
    { n: "\u767e\u4e07\u7ea7", l: "\u521b\u5efa\u5956\u52b1", d: "\u6210\u529f\u521b\u5efa\u7eff\u8272\u56ed\u533a\u53ef\u83b7\u767e\u4e07\u7ea7\u4e00\u6b21\u6027\u5956\u52b1", c: "EB5757" }
  ];
  metrics.forEach(function(m, i) {
    var x = 0.5 + i * 2.35;
    addCard(sl, x, 1.2, 2.15, 2.0, {});
    sl.addText(m.n, { x, y: 1.4, w: 2.15, h: 0.6, fontSize: 28, fontFace: "Arial", color: m.c, bold: true, align: "center" });
    sl.addText(m.l, { x, y: 2.0, w: 2.15, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.primary, bold: true, align: "center" });
    sl.addText(m.d, { x: x + 0.1, y: 2.35, w: 1.95, h: 0.6, fontSize: 9, fontFace: "Microsoft YaHei", color: T.gray, align: "center" });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9, h: 0.04, fill: { color: T.light } });
  sl.addText("\u5b9e\u65bd\u7b56\u7565\uff1a\u6025\u7528\u5148\u884c\u3001\u5206\u6b65\u4e0a\u7ebf", { x: 0.5, y: 3.7, w: 9, h: 0.35, fontSize: 13, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("\u4f18\u5148\u5b8c\u6210\u73af\u4fdd\u3001\u5b89\u9632\u3001\u62db\u5546\u5c55\u793a\u7b49\u652f\u6491\u521b\u5efa\u7533\u62a5\u7684\u6838\u5fc3\u529f\u80fd\uff0c\u786e\u4fdd\u9879\u76ee\u6210\u679c\u80fd\u5feb\u901f\u670d\u52a1\u4e8e\u7701\u7ea7\u5f00\u53d1\u533a\u7684\u7533\u62a5\u5de5\u4f5c\u3002", { x: 0.5, y: 4.1, w: 9, h: 0.5, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  addBadge(sl, pn);
}

// === Roles Slide ===
function sRoles() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: T.primary }; addTopBar(sl);
  sl.addText("\u9762\u5411\u56db\u7c7b\u89d2\u8272\u7684\u6838\u5fc3\u4ef7\u503c", { x: 0.7, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  var roles = [
    { t: "\u653f\u5e9c\u9886\u5bfc\u5c42", d: "\u6570\u636e\u9a71\u52a8\u51b3\u7b56\uff0c\u63d0\u5347\u56ed\u533a\u6cbb\u7406\u80fd\u529b\u73b0\u4ee3\u5316\u6c34\u5e73\uff1b\u6210\u529f\u521b\u5efa\u7701\u7ea7\u7ecf\u5f00\u533a/\u7eff\u8272\u56ed\u533a\uff0c\u4e89\u53d6\u767e\u4e07\u7ea7\u653f\u7b56\u5956\u52b1", c: T.accent, i: "\ud83c\udfdb" },
    { t: "\u56ed\u533a\u7ba1\u7406\u5c42", d: "\u8fd0\u8425\u6548\u7387\u63d0\u534730%\uff0c\u5e94\u6025\u54cd\u5e94\u63d0\u901f60%\uff1b\u4e00\u5c4f\u7edf\u89c8\u3001\u4e00\u56fe\u7edf\u7ba1\uff0c\u544a\u522b\u4fe1\u606f\u5b64\u5c9b", c: T.highlight, i: "\ud83d\udcca" },
    { t: "\u5165\u9a7b\u4f01\u4e1a", d: "\u4e00\u7ad9\u5f0f\u670d\u52a1\uff0c\u8bc9\u6c42\u76f4\u8fbe2\u5c0f\u65f6\u54cd\u5e94\uff1b\u5bf9\u63a5\u8d8520\u4ebf\u667a\u6539\u6570\u8f6c\u4e13\u9879\u8d44\u91d1\uff0c\u964d\u4f4e\u8f6c\u578b\u6210\u672c", c: "F59E0B", i: "\ud83c\udfe2" },
    { t: "\u516c\u4f17\u7528\u6237", d: "\u4fe1\u606f\u516c\u5f00\u900f\u660e\uff0c\u56ed\u533a\u52a8\u6001\u4e00\u624b\u638c\u63e1\uff1b\u53c2\u4e0e\u56ed\u533a\u5171\u5efa\uff0c\u589e\u5f3a\u5f52\u5c5e\u611f\u4e0e\u83b7\u5f97\u611f", c: "EB5757", i: "\ud83d\udc65" }
  ];
  roles.forEach(function(r, i) {
    addRoleCard(sl, 0.4 + i * 2.35, 1.1, 2.2, 3.8, r.i, r.t, r.d, r.c);
  });
  addBadge(sl, pn);
}

// === Future Slide ===
function sFuture() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" }; addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("\u672a\u6765\u5c55\u671b", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 0.75, y: 0.8, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: T.accent });
  var items = ["AIoT\u6df1\u5ea6\u878d\u5408\uff1a\u5f15\u5165\u7269\u8054\u7f51\u4f20\u611f\u8bbe\u5907\uff0c\u5b9e\u73b0\u4e07\u7269\u4e92\u8054\u3001\u5168\u57df\u611f\u77e5", "\u6570\u5b57\u5b6a\u751f\u56ed\u533a\uff1a1:1\u4e09\u7ef4\u5efa\u6a21\uff0c\u5168\u8981\u7d20\u3001\u5168\u65f6\u57df\u53ef\u89c6\u5316\u8fd0\u8425\u7ba1\u7406", "AI Agent\u667a\u80fd\u5316\uff1a\u667a\u80fd\u5ba2\u670d\u3001\u667a\u80fd\u8c03\u5ea6\u3001\u667a\u80fd\u9884\u8b66\uff0c\u4ece\u4eba\u627e\u6570\u636e\u5230\u6570\u636e\u627e\u4eba", "\u4ea7\u4e1a\u5927\u8111\uff1a\u57fa\u4e8e\u5927\u6570\u636e\u4e0eAI\u7b97\u6cd5\u8f85\u52a9\u4ea7\u4e1a\u62db\u5546\u51b3\u7b56\uff0c\u7cbe\u51c6\u5f3a\u94fe\u8865\u94fe", "\u8de8\u56ed\u533a\u534f\u540c\uff1a\u5bf9\u63a5\u7701\u5e02\u667a\u6167\u5e73\u53f0\uff0c\u5b9e\u73b0\u6570\u636e\u5171\u4eab\u4e0e\u4e1a\u52a1\u534f\u540c"];
  items.forEach(function(item, i) {
    var y = 1.4 + i * 0.7;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.55, fill: { color: T.cardBg }, rectRadius: 0.06 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y + 0.1, w: 0.05, h: 0.35, fill: { color: T.accent } });
    sl.addText(item, { x: 0.75, y, w: 8.5, h: 0.55, fontSize: 12, fontFace: "Microsoft YaHei", color: T.darkText, valign: "middle" });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 9, h: 0.02, fill: { color: T.light } });
  sl.addText("\u9a71\u52a8\u56ed\u533a\u6cbf\u7740\u7701\u7ea7\u5f00\u53d1\u533a\u2192\u7eff\u8272\u56ed\u533a\u2192\u56fd\u5bb6\u7ea7\u7ecf\u5f00\u533a\u57f9\u80b2\u5bf9\u8c61\u7684\u822a\u5411\u5168\u901f\u524d\u8fdb", { x: 0.5, y: 4.85, w: 9, h: 0.3, fontSize: 12, fontFace: "Microsoft YaHei", color: T.secondary, italic: true, align: "center" });
  addBadge(sl, pn);
}

// === Thank You ===
function sThankYou() {
  pn++; const sl = pres.addSlide();
  sl.background = { color: T.primary }; addTopBar(sl);
  sl.addShape(pres.shapes.OVAL, { x: 3, y: 0.5, w: 4, h: 4, fill: { color: T.secondary, transparency: 60 } });
  sl.addText("\u611f\u8c22\u804a\u542c", { x: 1, y: 1.5, w: 8, h: 1.0, fontSize: 38, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center" });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4, y: 2.5, w: 2, h: 0.04, fill: { color: T.accent } });
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 1, y: 2.7, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: T.highlight, align: "center" });
  var info = ["\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u7ba1\u59d4\u4f1a", "\u6280\u672f\u652f\u6301\uff1a\u9752\u8fc5\u79d1\u6280", "\u8054\u7cfb\u7535\u8bdd\uff1a0839-xxxxxxx", "\u90ae\u7bb1\uff1apark@zhuangzishang.gov.cn"];
  info.forEach(function(item, i) {
    sl.addText(item, { x: 1, y: 3.5 + i * 0.4, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: T.light, align: "center" });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: T.accent } });
}

// === Build ===
s1(); s2(); s3(); s4(); s5();
sDetail("\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7", [
  { t: "\u667a\u80fd\u6444\u50cf\u673a\u8986\u76d6", d: "\u90e8\u7f72\u667a\u80fd\u6444\u50cf\u673a\u4e0e\u4e13\u4e1a\u5730\u8d28\u4f20\u611f\u5668\uff0c\u6784\u5efa\u5168\u8986\u76d6\u611f\u77e5\u7f51\u7edc" },
  { t: "\u81ea\u52a8\u8bc6\u522b\u4e0e\u5b9e\u65f6\u9884\u8b66", d: "\u5b9e\u73b0\u5bf9\u4ea4\u901a\u8fdd\u89c4\u3001\u5730\u8d28\u707e\u5bb3\u98ce\u9669\u7684\u81ea\u52a8\u8bc6\u522b\u4e0e\u5b9e\u65f6\u9884\u8b66" },
  { t: "\u5b89\u5168\u8fd0\u884c\u5e95\u7ebf", d: "\u7b51\u7262\u56ed\u533a\u5b89\u5168\u8fd0\u884c\u5e95\u7ebf\uff0c\u4e3a\u521b\u5efa\u8bc4\u5ba1\u63d0\u4f9b\u5b89\u5168\u7a33\u5b9a\u7684\u575a\u5b9e\u57fa\u7840" },
  { t: "\u8d4b\u80fd\u521b\u5efa", d: "\u4e3a\u7701\u7ea7\u7ecf\u5f00\u533a\u53d1\u5c55\u8d28\u91cf\u4e0e\u7eff\u8272\u56ed\u533a\u57fa\u7840\u8bbe\u65bd\u5b89\u5168\u63d0\u4f9b\u786c\u6027\u652f\u6491" }
], T.accent);

sDetail("\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7\u5de5\u7a0b", [
  { t: "\u5728\u7ebf\u76d1\u6d4b\u8bbe\u5907\u5e03\u8bbe", d: "\u5728\u5173\u952e\u6392\u6c61\u53e3\u53ca\u73af\u5883\u654f\u611f\u533a\u57df\u5e03\u8bbe\u6c34\u8d28\u3001\u7a7a\u6c14\u5728\u7ebf\u76d1\u6d4b\u8bbe\u5907" },
  { t: "\u95ed\u73af\u7ba1\u7406\u4f53\u7cfb", d: "\u5efa\u7acb\u76d1\u6d4b-\u9884\u8b66-\u6eaf\u6e90\u7684\u95ed\u73af\u7ba1\u7406\u4f53\u7cfb" },
  { t: "\u8fde\u7eed\u6570\u636e\u8bc1\u636e", d: "\u4e3a\u6c61\u67d3\u7269\u7a33\u5b9a\u8fbe\u6807\u6392\u653e\u7387\u63d0\u4f9b\u8fde\u7eed\u3001\u771f\u5b9e\u3001\u53ef\u8ffd\u6eaf\u7684\u5728\u7ebf\u76d1\u6d4b\u6570\u636e" },
  { t: "\u7eff\u8272\u56ed\u533a\u6838\u5fc3\u652f\u6491", d: "\u662f\u7533\u62a5\u7eff\u8272\u5de5\u4e1a\u56ed\u533a\u6700\u5177\u8bf4\u670d\u529b\u7684\u6570\u636e\u8bc1\u636e" }
], T.highlight);

sDetail("\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546\u5de5\u7a0b", [
  { t: "\u4e09\u7ef4\u6570\u5b57\u6c99\u76d8", d: "\u5efa\u8bbe\u4e09\u7ef4\u6570\u5b57\u6c99\u76d8\u3001\u6570\u636e\u9a7e\u9a76\u8231\uff0c\u6253\u9020\u6c89\u6d78\u5f0f\u73b0\u4ee3\u5316\u5c55\u793a\u7a97\u53e3" },
  { t: "\u6570\u5b9e\u62db\u5546\u5927\u8111", d: "\u8fd0\u7528\u5927\u6570\u636e\u4e0eAI\u6280\u672f\u7cbe\u51c6\u7ed8\u5236\u4ea7\u4e1a\u94fe\u56fe\u8c31\uff0c\u667a\u80fd\u6316\u6398\u5339\u914d\u62db\u5546\u76ee\u6807" },
  { t: "\u5bf9\u5916\u5f00\u653e\u5f62\u8c61\u63d0\u5347", d: "\u76f4\u63a5\u670d\u52a1\u4e8e\u56fd\u5bb6\u7ea7\u7ecf\u5f00\u533a\u5bf9\u5916\u5f00\u653e\u5f62\u8c61\u63d0\u5347" },
  { t: "\u4ea7\u4e1a\u7ed3\u6784\u4f18\u5316", d: "\u9a71\u52a8\u9ad8\u8d28\u91cf\u9879\u76ee\u843d\u5730\uff0c\u52a9\u529b\u4f18\u5316\u4ea7\u4e1a\u7ed3\u6784\uff0c\u8865\u5f3a\u79d1\u6280\u521b\u65b0\u6307\u6807" }
], "F59E0B");

sDetail("\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b\u4e0e\u4ea9\u5747\u8bba\u82f1\u96c4", [
  { t: "\u667a\u80fd\u8ba1\u91cf\u8868\u8ba1", d: "\u63a8\u52a8\u91cd\u70b9\u4f01\u4e1a\u5b89\u88c5\u667a\u80fd\u8ba1\u91cf\u8868\u8ba1\uff0c\u5b9e\u65f6\u91c7\u96c6\u80fd\u8017\u4e0e\u4ea7\u51fa\u6570\u636e" },
  { t: "\u4ea9\u5747\u8bba\u82f1\u96c4\u8bc4\u4ef7\u4f53\u7cfb", d: "\u6784\u5efa\u4ea9\u5747\u8bba\u82f1\u96c4\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u5b9e\u73b0\u8d44\u6e90\u8981\u7d20\u5dee\u522b\u5316\u914d\u7f6e" },
  { t: "\u7eff\u8272\u56ed\u533a\u6838\u7b97\u57fa\u7840", d: "\u4e3a\u7eff\u8272\u56ed\u533a\u80fd\u6e90\u4e0e\u8d44\u6e90\u5229\u7528\u7eff\u8272\u5316\u6307\u6807\u63d0\u4f9b\u7cbe\u51c6\u6838\u7b97\u57fa\u7840" },
  { t: "\u5012\u903c\u4f01\u4e1a\u63d0\u8d28\u589e\u6548", d: "\u4ee5\u6570\u636e\u9a71\u52a8\u7684\u65b9\u5f0f\u5012\u903c\u4f01\u4e1a\u63d0\u8d28\u589e\u6548\u3001\u8f6c\u578b\u5347\u7ea7" }
], "EB5757");

sDetail("\u667a\u6539\u6570\u8f6c\u4f01\u4e1a\u8d4b\u80fd\u5e73\u53f0", [
  { t: "\u653f\u7b56\u7cbe\u51c6\u63a8\u9001", d: "\u96c6\u6210\u653f\u7b56\u7cbe\u51c6\u63a8\u9001\u3001\u4f9b\u9700\u667a\u80fd\u5bf9\u63a5\u3001\u8f7b\u91cf\u5316\u6570\u5b57\u5316\u8bca\u65ad\u5de5\u5177\u7b49\u529f\u80fd" },
  { t: "\u4e13\u9879\u8d44\u91d1\u5bf9\u63a5", d: "\u5e2e\u52a9\u4f01\u4e1a\u5bf9\u63a5\u7701\u3001\u5e02\u6bcf\u5e74\u8d8520\u4ebf\u5143\u7684\u667a\u6539\u6570\u8f6c\u4e13\u9879\u8d44\u91d1" },
  { t: "\u7834\u89e3\u8f6c\u578b\u56f0\u5883", d: "\u7834\u89e3\u4f01\u4e1a\u4e0d\u6562\u8f6c\u3001\u4e0d\u4f1a\u8f6c\u3001\u4e0d\u80fd\u8f6c\u3001\u4e0d\u613f\u8f6c\u7684\u56f0\u5883" },
  { t: "\u4ea7\u4e1a\u57fa\u7840\u5942\u5b9e", d: "\u6253\u9020\u56ed\u533a\u4f18\u8d28\u670d\u52a1\u751f\u6001\uff0c\u5942\u5b9e\u6240\u6709\u521b\u5efa\u76ee\u6807\u7684\u4ea7\u4e1a\u57fa\u7840" }
], "9B51E0");

sPlatform();
sInvestment();
sBenefits();
sRoles();
sFuture();
sThankYou();

pres.writeFile({ fileName: path.join(__dirname, "output", "zhuangzishang_platform_report.pptx") })
  .then(function() { console.log("PPT OK"); })
  .catch(function(e) { console.error("ERR:", e); });
