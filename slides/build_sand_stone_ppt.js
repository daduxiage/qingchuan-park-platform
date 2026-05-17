const pptxgen = require("pptxgenjs");
const path = require("path");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "\u9752\u8fc5\u79d1\u6280";
pres.title = "\u4e24\u6cb3\u53e3\u7802\u77f3\u667a\u80fd\u5efa\u9020\u57fa\u5730\u667a\u6167\u5316\u7ba1\u63a7\u5e73\u53f0";

// COLOR
const C = {
  dark: "0F172A", bg: "F8FAFC", card: "FFFFFF",
  text: "1E293B", gray: "64748B", subtxt: "94A3B8",
  line: "E2E8F0",
  gov: "2563EB", govL: "EFF6FF",
  mgmt: "059669", mgmtL: "ECFDF5",
  ent: "D97706", entL: "FFFBEB",
  pub: "7C3AED", pubL: "F5F3FF",
  acc: "F59E0B", accL: "FEF3C7",
};
const F = "Microsoft YaHei";
let p = 0;

// Helpers
function badge(sl, n, c) {
  var sz = n >= 10 ? 7 : 9;
  sl.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.35, h: 0.35, fill: { color: c || C.gov } });
  sl.addText(String(n), { x: 9.3, y: 5.1, w: 0.35, h: 0.35, fontSize: sz, fontFace: F, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
}
function tb(sl, c) { sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.035, fill: { color: c || C.gov } }); }
function sco(sl, x, y, s, c) { sl.addShape(pres.shapes.OVAL, { x, y, w: s, h: s, fill: { color: c } }); }
function scr(sl, x, y, w, h, c, r) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: c }, rectRadius: r||0.06,
    shadow: { type: "outer", blur: 4, offset: 1.5, color: "000000", opacity: 0.04 } });
}
function srec(sl, x, y, w, h, c) { sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: c } }); }
function titleDeco(sl, title, subtitle, accent) {
  var c = accent || C.gov;
  sl.addText(title, { x: 0.5, y: 0.12, w: 8.5, h: 0.5, fontSize: 20, fontFace: F, color: C.text, bold: true });
  srec(sl, 0.5, 0.6, 0.7, 0.025, c);
  if (subtitle) sl.addText(subtitle, { x: 0.5, y: 0.66, w: 8.5, h: 0.22, fontSize: 9, fontFace: F, color: C.gray });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.2, y: 0.08, w: 0.5, h: 0.22, fill: { color: c, transparency: 25 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.8, y: 0.08, w: 0.5, h: 0.22, fill: { color: c, transparency: 45 } });
}
function roleSlide(title, items, role, accent) {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, accent);
  titleDeco(sl, title, "", accent);
  sl.addText(role, { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: accent, align: "right" });
  items.forEach(function(item, i) {
    var col = i % 2, row = Math.floor(i / 2);
    var x = 0.5 + col * 4.65, y = 1.05 + row * 0.85;
    scr(sl, x, y, 4.35, 0.7, C.card, 0.08);
    srec(sl, x, y, 0.04, 0.7, accent);
    sco(sl, x+0.15, y+0.15, 0.15, accent);
    sl.addText(String(i+1), {x:x+0.15, y:y+0.15, w:0.15, h:0.15, fontSize:7, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(item.title || item, {x:x+0.4, y:y+0.05, w:3.8, h:0.25, fontSize:10, fontFace:F, color:accent, bold:true});
    if (item.desc) sl.addText(item.desc, {x:x+0.4, y:y+0.32, w:3.8, h:0.32, fontSize:7.5, fontFace:F, color:C.text});
    if (item.tag) {
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+3.2, y: y+0.07, w: 1.0, h: 0.18, fill: { color: C.line }, rectRadius: 0.02 });
      sl.addText(item.tag, {x:x+3.2, y:y+0.07, w:1.0, h:0.18, fontSize:6, fontFace:F, color:accent, align:"center", valign:"middle"});
    }
  });
  badge(sl, p, accent);
}

// ====== S1 COVER ======
function s1() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  // Cover parallelograms
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 7.6, y: 0.1, w: 0.5, h: 0.25, fill: { color: C.mgmt, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.2, y: 0.1, w: 0.5, h: 0.25, fill: { color: C.ent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.8, y: 0.1, w: 0.5, h: 0.25, fill: { color: C.pub, transparency: 30 } });
  srec(sl, 0.5, 0.5, 0.06, 2.8, C.acc);
  sl.addText("\u4e24\u6cb3\u53e3\u7802\u77f3\u667a\u80fd\u5efa\u9020\u57fa\u5730", { x: 0.8, y: 0.5, w: 8, h: 0.35, fontSize: 16, fontFace: F, color: C.gray });
  sl.addText("\u667a\u6167\u5316\u7ba1\u63a7\u5e73\u53f0\u4ea7\u54c1\u4ecb\u7ecd", { x: 0.8, y: 0.85, w: 7, h: 1.0, fontSize: 36, fontFace: F, color: C.text, bold: true });
  sl.addText("BIM+GIS \u00b7 \u6570\u5b57\u5b6a\u751f \u00b7 AI\u89c6\u9891 \u00b7 DCS\u81ea\u63a7", { x: 0.8, y: 1.7, w: 6, h: 0.35, fontSize: 12, fontFace: F, color: C.gray });
  // Tags
  var tags = [[C.gov,"\u667a\u6167\u7ba1\u63a7\u5e73\u53f0"],[C.mgmt,"\u6570\u5b57\u5b6a\u751f"],[C.ent,"AI\u89c6\u9891\u5206\u6790"],[C.pub,"DCS\u81ea\u52a8\u63a7\u5236"]];
  tags.forEach(function(t,i){
    var x=0.8+i*2.2;
    scr(sl, x, 2.3, 2.0, 0.38, t[0], 0.06);
    sl.addText(t[1],{x,y:2.3,w:2.0,h:0.38,fontSize:10,fontFace:F,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
  });
  scr(sl, 0.5, 3.0, 9, 0.7, C.card, 0.1);
  srec(sl, 0.5, 3.55, 9, 0.15, C.acc);
  var info = [["\u7802\u77f3\u884c\u4e1a\u667a\u80fd\u5316","\u5168\u6d41\u7a0b\u667a\u6167\u7ba1\u63a7"],["\u6570\u5b57\u5b6a\u751f\u5efa\u6a21","3D\u53ef\u89c6\u5316\u8fd0\u8425"],["AI\u89c6\u9891\u5206\u6790","\u667a\u80fd\u5b89\u5168\u76d1\u63a7"]];
  info.forEach(function(n,i){
    var x=0.8+i*3.0;
    sl.addText(n[0],{x,y:3.1,w:2.8,h:0.2,fontSize:10,fontFace:F,color:C.text,bold:true,align:"center"});
    sl.addText(n[1],{x,y:3.3,w:2.8,h:0.2,fontSize:9,fontFace:F,color:C.gray,align:"center"});
  });
}

// ====== S2 TOC ======
function s2() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  sl.addText("\u4ea7\u54c1\u76ee\u5f55", { x: 0.5, y: 0.1, w: 5, h: 0.5, fontSize: 22, fontFace: F, color: C.text, bold: true });
  var ses = [
    ["01", "\u4ea7\u54c1\u6982\u8ff0", "\u80cc\u666f\u4e0e\u67b6\u6784", C.gov],
    ["02", "\u6838\u5fc3\u6a21\u5757", "BIM+GIS / \u6570\u5b57\u5b6a\u751f / AI / DCS", C.acc],
    ["03", "\u786c\u4ef6\u8bbe\u5907", "\u76d1\u63a7 / \u670d\u52a1\u5668 / \u7f51\u7edc", C.mgmt],
    ["04", "\u5e94\u7528\u573a\u666f", "\u91c7\u8fd0\u9500\u5168\u6d41\u7a0b\u7ba1\u63a7", C.pub],
    ["05", "\u6848\u4f8b\u4e00\u89c8", "\u9879\u76ee\u6570\u636e\u4e0e\u5ba2\u6237", C.ent],
    ["06", "\u6848\u4f8b\u8be6\u89e3", "\u4ee3\u8868\u6027\u9879\u76ee\u5c55\u793a", C.gov],
    ["07", "\u4ea7\u54c1\u6536\u76ca", "\u4ef7\u503c\u4e0e\u6548\u76ca\u5206\u6790", C.mgmt],
    ["08", "\u672a\u6765\u5c55\u671b", "\u53d1\u5c55\u89c4\u5212", C.acc],
  ];
  ses.forEach(function(s, i){
    var x = 0.5 + (i%4) * 2.3, y = 0.8 + Math.floor(i/4) * 2.0;
    scr(sl, x, y, 2.15, 1.8, C.card, 0.12);
    srec(sl, x, y, 2.15, 0.06, s[3]);
    sco(sl, x+0.1, y+0.2, 0.46, s[3]);
    sl.addText(s[0], {x:x+0.1, y:y+0.2, w:0.46, h:0.46, fontSize:15, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(s[1], {x:x+0.65, y:y+0.18, w:1.4, h:0.25, fontSize:13, fontFace:F, color:s[3], bold:true});
    sl.addText(s[2], {x:x+0.65, y:y+0.43, w:1.4, h:0.2, fontSize:9, fontFace:F, color:C.gray});
    srec(sl, x+0.1, y+0.8, 1.95, 0.005, C.line);
    var feats = s[2].split(" ");
    feats.forEach(function(f,fi){ sl.addText("\u25b8 "+f, {x:x+0.12, y:y+0.9+fi*0.32, w:1.9, h:0.3, fontSize:8, fontFace:F, color:C.gray}); });
  });
  badge(sl, p, C.gov);
}

// ====== S3 OVERVIEW ======
function s3() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u4ea7\u54c1\u6982\u8ff0", "\u4e24\u6cb3\u53e3\u7802\u77f3\u667a\u80fd\u5efa\u9020\u57fa\u5730\u667a\u6167\u5316\u7ba1\u63a7\u5e73\u53f0", C.gov);
  scr(sl, 0.5, 0.95, 9, 0.4, C.govL, 0.06);
  sl.addText("\u5e73\u53f0\u81f4\u529b\u4e8e\u6784\u5efa\u7802\u77f3\u884c\u4e1a\u5168\u6d41\u7a0b\u667a\u6167\u7ba1\u63a7\u4f53\u7cfb\uff0c\u96c6\u6210BIM+GIS\u6570\u5b57\u5b6a\u751f\u3001AI\u89c6\u9891\u5206\u6790\u3001DCS\u96c6\u6563\u63a7\u5236\u3001\u7269\u8054\u7f51\u76d1\u63a7\u7b49\u6838\u5fc3\u80fd\u529b\uff0c\u5b9e\u73b0\u4ece\u91c7\u77ff\u3001\u751f\u4ea7\u3001\u8fd0\u8f93\u5230\u9500\u552e\u7684\u5168\u94fe\u6761\u6570\u5b57\u5316\u7ba1\u63a7\u3002",
    { x: 0.7, y: 0.96, w: 8.6, h: 0.38, fontSize: 10, fontFace: F, color: C.text, valign: "middle" });
  var md = [
    ["7\u5927", "\u6838\u5fc3\u6a21\u5757", "BIM/GIS/AI/DCS\u7b49", C.gov],
    ["18\u4e2a", "\u786e\u8ba4\u6848\u4f8b", "\u7d2f\u8ba1\u591a\u5e74\u5b9e\u8df5", C.mgmt],
    ["10+\u5bb6", "\u5408\u4f5c\u5ba2\u6237", "\u56fd\u4f01/\u6c11\u4f01\u5168\u8986\u76d6", C.ent],
    ["516%\u589e\u5f3a", "\u7ba1\u63a7\u80fd\u529b", "\u6570\u5b57\u5316\u8f6c\u578b\u5347\u7ea7", C.pub],
  ];
  md.forEach(function(m, i){
    var x = 0.5 + i * 2.3, y = 1.6;
    scr(sl, x, y, 2.1, 1.15, C.card, 0.1);
    sl.addText(m[0], {x, y:y+0.08, w:2.1, h:0.4, fontSize:22, fontFace:"Arial", color:m[3], bold:true, align:"center"});
    sl.addText(m[1], {x, y:y+0.48, w:2.1, h:0.2, fontSize:10, fontFace:F, color:C.text, bold:true, align:"center"});
    sl.addText(m[2], {x, y:y+0.72, w:2.1, h:0.2, fontSize:8, fontFace:F, color:C.gray, align:"center"});
    srec(sl, x+0.2, y+1.05, 1.7, 0.02, m[3]);
  });
  // Architecture
  sl.addText("\u5e73\u53f0\u67b6\u6784", { x: 0.5, y: 3.0, w: 5, h: 0.25, fontSize: 11, fontFace: F, color: C.text, bold: true });
  var arch = [
    ["\u5c55\u793a\u5c42", "\u6570\u5b57\u5b6a\u751f3D\u53ef\u89c6\u5316 | \u5927\u5c4f\u6570\u636e\u9a7e\u9a76\u8231 | \u79fb\u52a8\u7aefAPP", C.gov],
    ["\u5e94\u7528\u5c42", "AI\u89c6\u9891\u5206\u6790 | DCS\u96c6\u6563\u63a7\u5236 | \u751f\u4ea7\u7ba1\u7406 | \u5b89\u5168\u76d1\u63a7", C.mgmt],
    ["\u6570\u636e\u5c42", "\u6570\u5b57\u5b6a\u751f\u5f15\u64ce | \u7269\u8054\u7f51\u6570\u636e\u6e56 | \u89c6\u9891\u5206\u6790\u5e73\u53f0 | GIS\u5730\u7406\u4fe1\u606f", C.ent],
    ["\u8bbe\u65bd\u5c42", "\u6444\u50cf\u5934\u7f51\u7edc | \u4f20\u611f\u5668 | \u667a\u80fd\u8868\u8ba1 | \u8fb9\u7f18\u8ba1\u7b97\u7bb1 | \u96f7\u8fbe\u6d4b\u91cf", C.pub],
    ["\u5bf9\u63a5\u5c42", "\u7701\u6c34\u5229\u5927\u6570\u636e\u5e73\u53f0 | \u7535\u5b50\u91c7\u8fd0\u5355 | \u751f\u4ea7\u7cfb\u7edf | \u8d22\u52a1\u7cfb\u7edf", "475569"],
  ];
  arch.forEach(function(a, i){
    var y = 3.35 + i * 0.38;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.32, fill: { color: a[2] }, rectRadius: 0.04,
      shadow: { type: "outer", blur: 2, offset: 1, color: "000000", opacity: 0.03 } });
    sl.addText(a[0], { x: 0.6, y, w: 1.2, h: 0.32, fontSize: 8, fontFace: F, color: "FFFFFF", bold: true, valign: "middle" });
    sl.addText(a[1], { x: 1.9, y, w: 7.4, h: 0.32, fontSize: 7, fontFace: F, color: "FFFFFF", valign: "middle" });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.8, y: y+0.02, w: 0.03, h: 0.28, fill: { color: "FFFFFF", transparency: 50 }, rectRadius: 0.01 });
  });
  badge(sl, p, C.gov);
}

// ====== S4 CORE MODULES ======
function s4() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  titleDeco(sl, "\u6838\u5fc3\u529f\u80fd\u6a21\u5757", "\u4e94\u5927\u6838\u5fc3\u80fd\u529b\u9a71\u52a8\u7802\u77f3\u884c\u4e1a\u667a\u80fd\u5316\u5347\u7ea7", C.acc);
  var feats = [
    ["BIM+GIS\u6570\u5b57\u5b6a\u751f", "3D\u53ef\u89c6\u5316\u8fd0\u8425",
     "\u901a\u8fc7BIM\u5efa\u6a21\u4e0eGIS\u5730\u7406\u4fe1\u606f\u878d\u5408\uff0c\u6784\u5efa\u7802\u77f3\u5382\u533a1:1\u4e09\u7ef4\u6570\u5b57\u5b6a\u751f\u4f53\uff0c\u5b9e\u73b0\u8bbe\u5907\u72b6\u6001\u3001\u751f\u4ea7\u8fc7\u7a0b\u3001\u73af\u5883\u6570\u636e\u7684\u5168\u8981\u7d20\u53ef\u89c6\u5316\u4e0e\u5b9e\u65f6\u76d1\u63a7",
     C.gov],
    ["DCS\u96c6\u6563\u63a7\u5236\u7cfb\u7edf", "\u751f\u4ea7\u8fc7\u7a0b\u81ea\u52a8\u5316",
     "\u96c6\u6210\u7802\u77f3\u751f\u4ea7\u7ebf\u5404\u73af\u8282\u63a7\u5236\u7cfb\u7edf\uff0c\u5b9e\u73b0\u7834\u788e\u3001\u7b5b\u5206\u3001\u6d17\u6da4\u3001\u4f20\u9001\u7b49\u5de5\u5e8f\u7684\u96c6\u4e2d\u76d1\u63a7\u4e0e\u8fdc\u7a0b\u8c03\u63a7\uff0c\u63d0\u5347\u751f\u4ea7\u6548\u7387\u4e0e\u5b89\u5168\u6027",
     C.mgmt],
    ["AI\u89c6\u9891\u5206\u6790", "\u667a\u80fd\u5b89\u5168\u76d1\u63a7",
     "\u57fa\u4e8e\u6df1\u5ea6\u5b66\u4e60\u7684\u89c6\u9891\u5206\u6790\u5f15\u64ce\uff0c\u5b9e\u73b0\u8fdd\u7ae0\u4f5c\u4e1a\u8bc6\u522b\u3001\u5b89\u5168\u5e3d\u68c0\u6d4b\u3001\u8f66\u8f86\u8bc6\u522b\u3001\u8fb9\u754c\u8d8a\u754c\u544a\u8b66\u3001\u706b\u707e\u68c0\u6d4b\u7b49\u591a\u79cdAI\u80fd\u529b",
     C.ent],
    ["AI\u667a\u80fd\u7b97\u6cd5\u5e73\u53f0", "\u591a\u573a\u666f\u7b97\u6cd5\u5f15\u64ce",
     "\u63d0\u4f9b\u8d85\u8fc77\u79cd\u7802\u77f3\u884c\u4e1a\u4e13\u7528AI\u7b97\u6cd5\uff0c\u5305\u62ec\u8f66\u8f86\u8bc6\u522b\u3001\u4eba\u5458\u8bc6\u522b\u3001\u5f02\u5e38\u4e8b\u4ef6\u3001\u73af\u5883\u68c0\u6d4b\u7b49\uff0c\u652f\u6301\u8fb9\u7f18\u8ba1\u7b97\u4e0e\u4e91\u7aef\u534f\u540c",
     C.pub],
    ["\u7269\u8054\u7f51\u76d1\u63a7\u4e0e\u8bbe\u5907\u7ba1\u7406", "\u5168\u57df\u611f\u77e5\u8054\u52a8",
     "\u96c6\u6210\u6444\u50cf\u5934\u3001\u96f7\u8fbe\u3001\u667a\u80fd\u8868\u8ba1\u3001\u73af\u5883\u4f20\u611f\u5668\u7b49\u8bbe\u5907\uff0c\u5b9e\u73b0\u5bf9\u7802\u77f3\u5382\u533a\u4eba\u3001\u8f66\u3001\u7269\u3001\u73af\u5883\u7684\u5168\u65b9\u4f4d\u5b9e\u65f6\u611f\u77e5\u4e0e\u667a\u80fd\u7ba1\u7406",
     C.acc],
  ];
  feats.forEach(function(f, i){
    var x = 0.5 + (i%2) * 4.65, y = 0.95 + Math.floor(i/3) * 1.85;
    var h = (i < 4) ? 1.65 : 1.65;
    scr(sl, x, y, 4.35, h, C.card, 0.12);
    srec(sl, x, y, 4.35, 0.05, f[3]);
    sco(sl, x+0.15, y+0.15, 0.4, f[3]);
    sl.addText(String(i+1), {x:x+0.15, y:y+0.15, w:0.4, h:0.4, fontSize:15, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(f[0], {x:x+0.65, y:y+0.12, w:3.5, h:0.22, fontSize:11, fontFace:F, color:f[3], bold:true});
    sl.addText(f[1], {x:x+0.65, y:y+0.34, w:3.5, h:0.18, fontSize:8, fontFace:F, color:C.gray});
    srec(sl, x+0.15, y+0.6, 4.05, 0.005, C.line);
    sl.addText(f[2], {x:x+0.15, y:y+0.7, w:4.05, h:0.75, fontSize:7.5, fontFace:F, color:C.text});
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.15, y: y+h-0.25, w: 1.2, h: 0.18, fill: { color: f[3] }, rectRadius: 0.02 });
    sl.addText("\u6838\u5fc3\u80fd\u529b", { x: x+0.15, y: y+h-0.25, w: 1.2, h: 0.18, fontSize: 6.5, fontFace: F, color: "FFFFFF", align: "center", valign: "middle" });
  });
  badge(sl, p, C.acc);
}

// ====== S5 HARDWARE ======
function s5() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.mgmt);
  titleDeco(sl, "\u786c\u4ef6\u8bbe\u5907\u4e0e\u57fa\u7840\u8bbe\u65bd", "\u4e09\u5927\u786c\u4ef6\u4f53\u7cfb\u652f\u6491\u5e73\u53f0\u8fd0\u884c", C.mgmt);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 50 } });
  // 3 columns
  var cols = [
    ["\u76d1\u63a7\u4e0e\u611f\u77e5\u5c42", "\u5b89\u5168\u76d1\u63a7\u4e0e\u73af\u5883\u611f\u77e5",
     ["\u9ad8\u6e05\u6444\u50cf\u5934\uff087\u53f0\uff09\uff1a400\u4e07\u50cf\u7d20\uff0cIP67\u9632\u62a4", "\u667a\u80fd\u8fb9\u7f18\u8ba1\u7b97\u7bb1\uff0812\u53f0\uff09\uff1a\u652f\u6301AI\u63a8\u7406", "\u96f7\u8fbe\u6d4b\u91cf\u8bbe\u5907\uff1a\u5806\u6599\u573a\u4e09\u7ef4\u626b\u63cf", "\u73af\u5883\u4f20\u611f\u5668\uff1aDCS\u96c6\u6210\u76d1\u6d4b", "\u667a\u80fd\u8868\u8ba1\uff1a\u80fd\u8017/\u6c34\u8d28/PH\u76d1\u6d4b"],
     C.gov],
    ["\u8fb9\u7f18\u8ba1\u7b97\u5c42", "AI\u7b97\u529b\u4e0e\u6570\u636e\u5904\u7406",
     ["AI\u8fb9\u7f18\u8ba1\u7b97\u670d\u52a1\u5668\uff1aIntel Xeon 4314 + Tesla T4 GPU", "\u5185\u5b5832GB DDR5\uff0c4TB\u5b58\u50a8\uff0c\u652f\u6301RAID", "\u53ef\u540c\u65f6\u5904\u7406\u591a\u8def\u89c6\u9891\u6d41AI\u5206\u6790", "\u652f\u6301\u672c\u5730\u5316\u90e8\u7f72\uff0c\u4f4e\u5ef6\u8fdf\u54cd\u5e94"],
     C.mgmt],
    ["\u7f51\u7edc\u4e0e\u901a\u4fe1\u5c42", "\u6570\u636e\u4f20\u8f93\u4e0e\u8bbe\u5907\u5bf9\u63a5",
     ["24\u53e3\u5343\u5146\u4ea4\u6362\u673a\uff1a\u652f\u6301\u5168\u7f51\u7ba1\u7406", "\u5de5\u4e1a\u7f51\u5173\uff1a\u591a\u534f\u8bae\u5bf9\u63a5DCS\u53ca\u7269\u8054\u8bbe\u5907", "\u63a7\u5236\u67dc\u67dc\uff08600*1000*2000mm\uff09", "\u652f\u6301\u4ee5\u592a\u7f51/5G/\u7ec4\u6001\u7f51\u591a\u79cd\u901a\u4fe1\u65b9\u5f0f"],
     C.ent]
  ];
  cols.forEach(function(c, i){
    var x = 0.5 + i * 3.15;
    scr(sl, x, 0.85, 2.9, 3.5, C.card, 0.12);
    srec(sl, x, 0.85, 2.9, 0.06, c[3]);
    sl.addText(c[0], {x, y:0.95, w:2.9, h:0.25, fontSize:12, fontFace:F, color:c[3], bold:true, align:"center"});
    sl.addText(c[1], {x, y:1.2, w:2.9, h:0.2, fontSize:8.5, fontFace:F, color:C.gray, align:"center"});
    srec(sl, x+0.3, 1.45, 2.3, 0.015, c[3]);
    c[2].forEach(function(d, di){
      var dy = 1.6 + di * 0.42;
      sco(sl, x+0.12, dy+0.05, 0.05, c[3]);
      sl.addText(d, {x:x+0.25, y:dy, w:2.5, h:0.35, fontSize:7.5, fontFace:F, color:C.text});
    });
  });
  badge(sl, p, C.mgmt);
}

// ====== S6 APPLICATION SCENARIOS ======
function s6() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.pub);
  titleDeco(sl, "\u5e94\u7528\u573a\u666f", "\u8986\u76d6\u7802\u77f3\u91c7\u8fd0\u9500\u5168\u6d41\u7a0b\u667a\u80fd\u5316\u7ba1\u63a7", C.pub);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.pub, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.pub, transparency: 50 } });
  var scenes = [
    ["\u751f\u4ea7\u8fc7\u7a0b\u76d1\u63a7", "\u5b9e\u65f6\u76d1\u63a7\u7802\u77f3\u7834\u788e\u3001\u7b5b\u5206\u3001\u6d17\u6da4\u7b49\u5de5\u5e8f\u8fd0\u884c\u72b6\u6001\uff0c\u53ef\u89c6\u5316\u5c55\u793a\u751f\u4ea7\u7ebf\u8fd0\u884c\u53c2\u6570", C.pub],
    ["\u5b89\u5168\u76d1\u63a7\u4e0e\u9884\u8b66", "\u89c6\u9891AI\u5206\u6790\u8bc6\u522b\u8fdd\u7ae0\u4f5c\u4e1a\uff0c\u8fb9\u754c\u8d8a\u754c\u544a\u8b66\uff0c\u667a\u80fd\u68c0\u6d4b\u5b89\u5168\u5e3d\u3001\u706b\u707e\u7b49\u5f02\u5e38", "6D28D9"],
    ["\u8f66\u8f86\u4e0e\u7269\u6d41\u7ba1\u7406", "\u8f66\u8f86\u8bc6\u522b\u3001\u79f0\u91cd\u6570\u636e\u8054\u52a8\u3001\u8f66\u8f86\u8f68\u8ff9\u8ffd\u8e2a\uff0c\u652f\u6301\u7535\u5b50\u91c7\u8fd0\u5355\u8054\u52a8", "5B21B6"],
    ["\u73af\u5883\u76d1\u6d4b\u4e0e\u8d44\u6e90\u7ba1\u7406", "\u5c18\u5c18\u3001\u566a\u58f0\u3001\u6c34\u8d28PH\u7b49\u73af\u5883\u6307\u6807\u5b9e\u65f6\u76d1\u6d4b\uff0c\u7802\u77f3\u50a8\u91cf\u96f7\u8fbe\u626b\u63cf\u81ea\u52a8\u8ba1\u7b97", C.pub],
    ["\u91c7\u8fd0\u5355\u6570\u5b57\u5316", "\u5bf9\u63a5\u7701\u6c34\u5229\u5927\u6570\u636e\u5e73\u53f0\uff0c\u7535\u5b50\u91c7\u8fd0\u5355\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316\u7ba1\u7406", "6D28D9"],
    ["\u8bbe\u5907\u7ef4\u62a4\u4e0e\u8fdc\u7a0b\u8bca\u65ad", "\u8bbe\u5907\u8fd0\u884c\u72b6\u6001\u5b9e\u65f6\u76d1\u63a7\uff0c\u6545\u969c\u9884\u8b66\u3001\u8fdc\u7a0b\u8bca\u65ad\u4e0e\u7ef4\u62a4\u63d0\u793a", C.pub]
  ];
  scenes.forEach(function(s, i){
    var col = i % 2, row = Math.floor(i / 2);
    var x = 0.5 + col * 4.65, y = 0.85 + row * 1.1;
    scr(sl, x, y, 4.35, 0.95, C.card, 0.1);
    srec(sl, x, y, 0.04, 0.95, s[2]);
    // Icon number
    sco(sl, x+0.12, y+0.15, 0.3, s[2]);
    sl.addText(String(i+1), {x:x+0.12, y:y+0.15, w:0.3, h:0.3, fontSize:12, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(s[0], {x:x+0.5, y:y+0.1, w:3.7, h:0.22, fontSize:11, fontFace:F, color:s[2], bold:true});
    sl.addText(s[1], {x:x+0.12, y:y+0.5, w:4.1, h:0.4, fontSize:8, fontFace:F, color:C.text});
  });
  badge(sl, p, C.pub);
}

// ====== S7 CASE SUMMARY (timeline style) ======
function s7() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.ent);
  titleDeco(sl, "\u6848\u4f8b\u4e00\u89c8 \u2014 \u9879\u76ee\u65f6\u95f4\u7ebf", "\u7d2f\u8ba118\u4e2a\u786e\u8ba4\u6848\u4f8b\uff0c\u8986\u76d618\u5e74\u81f325\u5e74", C.ent);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 50 } });
  // Year sections
  var years = [
    ["2022\u5e74", "\u7802\u77f3\u6570\u636e\u5bf9\u63a5\u7f51\u5173\u9879\u76ee", C.ent],
    ["2023\u5e74", "\u5927\u6570\u636e\u4e91\u5e73\u53f0\u8fd0\u7ef4\u670d\u52a1", "B45309"],
    ["2024\u5e74", "\u79ef\u7d2f12\u4e2a\u9879\u76ee", "92400E"],
    ["2025\u5e74", "\u89c6\u9891\u76d1\u63a7+\u7535\u5b50\u56f4\u680f+\u96f7\u8fbe\u6d4b\u91cf\u7b49", C.ent],
  ];
  years.forEach(function(y, i){
    var x = 0.5 + i * 2.3;
    scr(sl, x, 0.85, 2.1, 1.0, C.card, 0.1);
    srec(sl, x, 0.85, 2.1, 0.04, y[2]);
    sl.addText(y[0], {x, y:0.92, w:2.1, h:0.2, fontSize:14, fontFace:F, color:y[2], bold:true, align:"center"});
    sl.addText(y[1], {x, y:1.15, w:2.1, h:0.2, fontSize:8.5, fontFace:F, color:C.gray, align:"center"});
    if (i < years.length-1) sl.addShape(pres.shapes.RIGHT_ARROW, { x: x+2.2, y: 1.2, w: 0.15, h: 0.12, fill: { color: C.line } });
  });
  // Key cases
  var cases = [
    ["\u6210\u90fd\u7a7a\u6e2f\u5174\u57ce\u5efa\u6750", "\u7802\u77f3\u5806\u573a\u667a\u6167\u5316\u7ba1\u7406\u5e73\u53f0\u5efa\u8bbe", "2022"],
    ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u957f\u6c5f\u6cb3\u9053\u7802\u77f3\u7535\u5b50\u91c7\u8fd0\u7ba1\u7406\u5355\u670d\u52a1", "2025"],
    ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u5806\u6599\u573a\u96f7\u8fbe\u6d4b\u91cf\u9879\u76ee", "2025"],
    ["\u4e2d\u56fd\u79fb\u52a8\u5fb7\u9633\u5206\u516c\u53f8", "\u7ef5\u7af9\u5e02\u91c7\u7802\u89c6\u9891\u76d1\u63a7\u548c\u7535\u5b50\u56f4\u680f", "2025"],
    ["\u6210\u90fd\u6e58\u6c5f\u73af\u5883\u65b0\u6750\u6599", "\u8fdb\u9500\u7cfb\u7edf\u786c\u4ef6\u8bbe\u5907\u91c7\u8d2d", "2025"],
    ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u96f6\u78b3\u8fd0\u8f93\u8f66\u8f86\u5b9a\u4f4d\u4e0e\u8f68\u8ff9\u5f00\u53d1", "2024"],
  ];
  cases.forEach(function(cs, i){
    var col = i % 3, row = Math.floor(i / 3);
    var x = 0.5 + col * 3.1, y = 2.1 + row * 1.0;
    scr(sl, x, y, 2.9, 0.8, C.card, 0.08);
    srec(sl, x, y, 0.03, 0.8, C.ent);
    sco(sl, x+0.1, y+0.1, 0.06, C.ent);
    sl.addText(cs[0], {x:x+0.22, y:y+0.06, w:2.6, h:0.2, fontSize:8.5, fontFace:F, color:C.ent, bold:true});
    sl.addText(cs[1], {x:x+0.1, y:y+0.28, w:2.7, h:0.3, fontSize:7.5, fontFace:F, color:C.text});
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.1, y: y+0.55, w: 0.6, h: 0.16, fill: { color: C.entL }, rectRadius: 0.02 });
    sl.addText(cs[2], {x:x+0.1, y:y+0.55, w:0.6, h:0.16, fontSize:6.5, fontFace:F, color:C.ent, align:"center", valign:"middle"});
  });
  badge(sl, p, C.ent);
}

// ====== S8 CASE DETAIL - SERVICE TYPES ======
function s8() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u6848\u4f8b\u8be6\u89e3 \u2014 \u670d\u52a1\u7c7b\u578b\u4e00\u89c8", "\u56db\u5927\u670d\u52a1\u7c7b\u578b\u8986\u76d6\u5ba2\u6237\u5168\u751f\u547d\u5468\u671f\u9700\u6c42", C.gov);
  // Service type cards - 2x2
  var svcs = [
    ["\u5e73\u53f0\u5b9e\u65bd\u670d\u52a1", C.gov,
     ["\u7802\u77f3\u5806\u573a\u667a\u6167\u5316\u7ba1\u7406\u5e73\u53f0\u5efa\u8bbe", "\u89c6\u9891\u76d1\u63a7\u548c\u7535\u5b50\u56f4\u680f\u9879\u76ee", "\u8fdb\u9500\u7cfb\u7edf\u786c\u4ef6\u8bbe\u5907\u91c7\u8d2d", "\u5806\u6599\u573a\u96f7\u8fbe\u6d4b\u91cf\u9879\u76ee", "\u96f6\u78b3\u8fd0\u8f93\u8f66\u8f86\u5b9a\u4f4d\u4e0e\u8f68\u8ff9\u5f00\u53d1"]],
    ["\u76d1\u7ba1\u63a5\u5165\u670d\u52a1", C.mgmt,
     ["\u7701\u6c34\u5229\u5927\u6570\u636e\u5e73\u53f0\u63a5\u5165", "\u7535\u5b50\u91c7\u8fd0\u5355\u5728\u7ebf\u76d1\u7ba1\u5b89\u88c5", "\u7802\u77f3\u91c7\u8fd0\u9500\u73af\u8282\u5728\u7ebf\u76d1\u7ba1", "\u6570\u636e\u63a5\u5165\u4e0e\u5e73\u53f0\u5bf9\u63a5"]],
    ["\u8fd0\u7ef4\u670d\u52a1", C.ent,
     ["7802\u77f3\u8d44\u6e90\u7ecf\u8425\u7ba1\u7406\u5927\u6570\u636e\u5e73\u53f0\u8fd0\u7ef4", "\u5e73\u53f0\u5b89\u5168\u7ec4\u4ef6\u642d\u5efa\u4e0e\u7b49\u4fdd\u8bc4\u6d4b", "\u957f\u671f\u8fd0\u7ef4\u670d\u52a1\u5408\u4f5c"]],
    ["\u7efc\u5408\u670d\u52a1", C.pub,
     ["\u5b89\u5168\u7ec4\u4ef6\u642d\u5efa\u4e0e\u7b49\u7ea7\u4fdd\u62a4\u4e8c\u7ea7\u8bc4\u6d4b", "\u6570\u636e\u63a5\u5165\u4e0e\u5e73\u53f0\u5b9e\u65bd\u4e00\u4f53\u5316\u670d\u52a1"]],
  ];
  svcs.forEach(function(s, i){
    var x = 0.5 + (i%2)*4.65, y = 0.85 + Math.floor(i/2)*1.8;
    scr(sl, x, y, 4.35, 1.6, C.card, 0.12);
    srec(sl, x, y, 4.35, 0.05, s[1]);
    sl.addText(s[0], {x:x+0.12, y:y+0.1, w:4, h:0.22, fontSize:12, fontFace:F, color:s[1], bold:true});
    s[2].forEach(function(d, di){
      var dy = y + 0.4 + di * 0.26;
      sco(sl, x+0.12, dy+0.04, 0.04, s[1]);
      sl.addText(d, {x:x+0.22, y:dy, w:3.9, h:0.22, fontSize:7.5, fontFace:F, color:C.text});
    });
  });
  badge(sl, p, C.gov);
}

// ====== S9 CASE DETAIL - ORGANIZATION GRID ======
function s9() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u6848\u4f8b\u8be6\u89e3 \u2014 \u4ee3\u8868\u6027\u5ba2\u6237\u4e0e\u9879\u76ee", "\u670d\u52a1\u8d85\u8fc710\u5bb6\u5ba2\u6237\uff0c\u6db5\u76d6\u56fd\u4f01\u3001\u6c11\u4f01\u3001\u8fd0\u8425\u5546\u7b49\u591a\u7c7b\u5ba2\u6237", C.gov);
  // Top row - key clients
  var clients = [
    ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u6700\u9ad8\u9891\u5ba2\u6237", "6\u9879\u5408\u4f5c", C.gov],
    ["\u4e2d\u56fd\u79fb\u52a8\u5fb7\u9633", "\u8fd0\u8425\u5546\u5408\u4f5c", "\u89c6\u9891\u56f4\u680f", C.mgmt],
    ["\u6210\u90fd\u7a7a\u6e2f\u5174\u57ce", "\u5927\u578b\u56fd\u4f01", "\u667a\u6167\u7802\u77f3", C.ent],
    ["\u6210\u90fd\u6e58\u6c5f\u73af\u5883", "\u6c11\u4f01\u5ba2\u6237", "\u8fdb\u9500\u7cfb\u7edf", C.pub],
  ];
  clients.forEach(function(cl, i){
    var x = 0.5 + i * 2.3;
    scr(sl, x, 0.85, 2.1, 0.9, C.card, 0.1);
    srec(sl, x, 0.85, 2.1, 0.04, cl[3]);
    sl.addText(cl[0], {x, y:0.92, w:2.1, h:0.22, fontSize:9, fontFace:F, color:cl[3], bold:true, align:"center"});
    sl.addText(cl[1], {x, y:1.15, w:2.1, h:0.18, fontSize:7.5, fontFace:F, color:C.gray, align:"center"});
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.5, y: 1.45, w: 1.1, h: 0.18, fill: { color: C.line }, rectRadius: 0.02 });
    sl.addText(cl[2], {x:x+0.5, y:1.45, w:1.1, h:0.18, fontSize:6.5, fontFace:F, color:cl[3], align:"center", valign:"middle"});
  });
  // Project list
  scr(sl, 0.5, 2.0, 9, 3.0, C.card, 0.1);
  sl.addText("\u90e8\u5206\u4ee3\u8868\u6027\u9879\u76ee\u5217\u8868", {x:0.7, y:2.05, w:4, h:0.22, fontSize:11, fontFace:F, color:C.gov, bold:true});
  srec(sl, 0.5, 2.25, 9, 0.01, C.line);
  var items = [
    ["\u7ef5\u7af9\u5e02\u91c7\u7802\u89c6\u9891\u76d1\u63a7\u548c\u7535\u5b50\u56f4\u680f\u9879\u76ee", "\u4e2d\u56fd\u79fb\u52a8\u5fb7\u9633", "\u5e73\u53f0\u5b9e\u65bd+\u76d1\u7ba1\u63a5\u5165", C.gov],
    ["\u957f\u6c5f\u6cb3\u9053\u7802\u77f3\u7535\u5b50\u91c7\u8fd0\u7ba1\u7406\u5355\u670d\u52a1", "\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u5e73\u53f0\u5b9e\u65bd+\u76d1\u7ba1\u63a5\u5165", C.mgmt],
    ["\u5806\u6599\u573a\u96f7\u8fbe\u6d4b\u91cf\u9879\u76ee", "\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u5e73\u53f0\u5b9e\u65bd", C.ent],
    ["\u7802\u77f3\u5806\u573a\u667a\u6167\u5316\u7ba1\u7406\u5e73\u53f0\u5efa\u8bbe", "\u6210\u90fd\u7a7a\u6e2f\u5174\u57ce\u5efa\u6750", "\u5e73\u53f0\u5b9e\u65bd", C.pub],
    ["\u7802\u77f3\u8d44\u6e90\u7ecf\u8425\u7ba1\u7406\u5927\u6570\u636e\u5e73\u53f0\u8fd0\u7ef4", "\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u8fd0\u7ef4\u670d\u52a1", C.gov],
    ["\u96f6\u78b3\u8fd0\u8f93\u8f66\u8f86\u5b9a\u4f4d\u4e0e\u8f68\u8ff9\u5f00\u53d1", "\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u5e73\u53f0\u5b9e\u65bd", C.mgmt],
    ["\u5b89\u5168\u7ec4\u4ef6\u642d\u5efa\u4e0e\u7b49\u7ea7\u4fdd\u62a4\u8bc4\u6d4b", "\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u7efc\u5408\u670d\u52a1", C.ent],
    ["\u7701\u6c34\u5229\u5927\u6570\u636e\u5e73\u53f0\u63a5\u5165\u670d\u52a1", "\u5efa\u6295\u5efa\u6750\u7f57\u6c5f", "\u76d1\u7ba1\u63a5\u5165", C.pub],
  ];
  items.forEach(function(it, i){
    var y = 2.4 + i * 0.32;
    sco(sl, 0.6, y+0.05, 0.04, it[3]);
    sl.addText(it[0], {x:0.72, y:y, w:4.0, h:0.22, fontSize:7.5, fontFace:F, color:C.text});
    sl.addText(it[1], {x:4.8, y:y, w:1.8, h:0.22, fontSize:7, fontFace:F, color:C.gray});
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: y+0.02, w: 1.2, h: 0.17, fill: { color: C.line }, rectRadius: 0.02 });
    sl.addText(it[2], {x:6.7, y:y+0.02, w:1.2, h:0.17, fontSize:6, fontFace:F, color:it[3], align:"center", valign:"middle"});
  });
  badge(sl, p, C.gov);
}

// ====== S10 BENEFITS ======
function s10() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.mgmt);
  titleDeco(sl, "\u4ea7\u54c1\u6536\u76ca", "\u667a\u6167\u5316\u7ba1\u63a7\u5e73\u53f0\u5e26\u6765\u7684\u591a\u7ef4\u5ea6\u4ef7\u503c\u4e0e\u6548\u76ca\u63d0\u5347", C.mgmt);
  sco(sl, 9.2, 0.15, 0.08, C.mgmtL);
  sco(sl, 9.4, 0.15, 0.05, C.mgmt);
  var benefits = [
    [C.gov, "\u751f\u4ea7\u6548\u7387\u63d0\u5347", "+30~50%",
     ["DCS\u96c6\u6563\u63a7\u5236\u5b9e\u73b0\u751f\u4ea7\u7ebf\u81ea\u52a8\u5316\u8fd0\u884c", "\u751f\u4ea7\u8fc7\u7a0b\u53ef\u89c6\u5316\u51cf\u5c11\u505c\u673a\u65f6\u95f4", "AI\u89c6\u9891\u5206\u6790\u51cf\u5c11\u5b89\u5168\u4e8b\u6545\u98ce\u9669", "\u6570\u636e\u81ea\u52a8\u91c7\u96c6\u51cf\u5c11\u4eba\u5de5\u8bb0\u5f55"]],
    [C.mgmt, "\u8fd0\u8425\u6210\u672c\u4f18\u5316", "\u00a5\u7ea625%\u8282\u7701",
     ["\u8bbe\u5907\u8fdc\u7a0b\u7ba1\u63a7\u51cf\u5c11\u73b0\u573a\u4eba\u5458\u914d\u7f6e", "\u80fd\u8017\u76d1\u63a7\u51cf\u5c11\u80fd\u6e90\u6d6a\u8d39", "\u667a\u80fd\u9884\u8b66\u51cf\u5c11\u8bbe\u5907\u6545\u969c\u635f\u5931", "\u6570\u5b57\u5316\u7ba1\u7406\u51cf\u5c11\u7eb8\u8d28\u6d41\u7a0b\u6210\u672c"]],
    [C.ent, "\u7ba1\u63a7\u80fd\u529b\u63d0\u5347", "\u5168\u6d41\u7a0b\u53ef\u89c6\u5316",
     ["\u91c7\u8fd0\u9500\u5168\u94fe\u6761\u6570\u5b57\u5316\u7ba1\u63a7", "\u7701\u5e73\u53f0\u6570\u636e\u5b9e\u65f6\u63a5\u5165", "\u7535\u5b50\u91c7\u8fd0\u5355\u5168\u7ebf\u4e0a\u5316\u7ba1\u7406", "\u591a\u573a\u666fAI\u667a\u80fd\u9884\u8b66\u4e0e\u5e94\u5bf9"]],
  ];
  benefits.forEach(function(b, i){
    var x = 0.5 + i * 3.15;
    scr(sl, x, 0.95, 2.9, 4.0, C.card, 0.12);
    srec(sl, x, 0.95, 2.9, 0.07, b[0]);
    sco(sl, x+1.1, 1.15, 0.7, b[0]);
    sl.addText(b[2], {x:x+1.1, y:1.15, w:0.7, h:0.7, fontSize:18, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(b[1], {x:x+0.15, y:1.9, w:2.6, h:0.25, fontSize:13, fontFace:F, color:b[0], bold:true, align:"center"});
    srec(sl, x+0.5, 2.2, 1.9, 0.02, b[0]);
    b[3].forEach(function(d, di){
      var dy = 2.35 + di * 0.4;
      sco(sl, x+0.2, dy+0.04, 0.05, b[0]);
      sl.addText(d, {x:x+0.35, y:dy, w:2.4, h:0.35, fontSize:8, fontFace:F, color:C.text});
    });
  });
  scr(sl, 0.5, 5.1, 9, 0.25, C.govL, 0.05);
  sl.addText("\u7efc\u5408\u6548\u76ca\uff1a\u5e73\u53f0\u5efa\u8bbe\u5e26\u6765\u751f\u4ea7\u6548\u7387\u3001\u8fd0\u8425\u6210\u672c\u3001\u7ba1\u63a7\u80fd\u529b\u4e09\u7ef4\u5ea6\u63d0\u5347\uff0c\u77ed\u671f\u6295\u5165\u957f\u671f\u56de\u62a5",
    { x: 0.7, y: 5.12, w: 8.6, h: 0.2, fontSize: 9, fontFace: F, color: C.gov, bold: true });
  badge(sl, p, C.mgmt);
}

// ====== S11 FUTURE ======
function s11() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u672a\u6765\u5c55\u671b", "\u6301\u7eed\u8d4b\u80fd\u7802\u77f3\u884c\u4e1a\u6570\u5b57\u5316\u8f6c\u578b", C.gov);
  var fs = [
    ["AI\u591a\u573a\u666f\u6df1\u5316", "\u5f15\u5165\u66f4\u591a\u884c\u4e1a\u4e13\u7528AI\u7b97\u6cd5\uff0c\u8986\u76d6\u8d28\u91cf\u68c0\u6d4b\u3001\u8bbe\u5907\u9884\u77e5\u7ef4\u62a4\u3001\u80fd\u8017\u4f18\u5316\u7b49\u573a\u666f", C.gov],
    ["\u4ea7\u4e1a\u94fe\u6570\u5b57\u5316\u534f\u540c", "\u6253\u901a\u7802\u77f3\u4ea7\u4e1a\u94fe\u4e0a\u4e0b\u6e38\u6570\u636e\uff0c\u5b9e\u73b0\u4ece\u77ff\u5c71\u5230\u5de5\u5730\u7684\u5168\u94fe\u6761\u6570\u5b57\u5316\u534f\u540c", C.mgmt],
    ["5G+\u8fb9\u7f18\u8ba1\u7b97\u6df1\u5316", "\u6784\u5efa\u9ad8\u901f\u3001\u4f4e\u5ef6\u8fdf\u7684\u7269\u8054\u7f51\u901a\u4fe1\u7f51\u7edc\uff0c\u652f\u6491\u66f4\u591a\u5b9e\u65f6\u6027\u8981\u6c42\u9ad8\u7684\u5e94\u7528\u573a\u666f", C.ent],
    ["\u7701\u5e02\u5e73\u53f0\u6df1\u5ea6\u5bf9\u63a5", "\u8fdb\u4e00\u6b65\u6df1\u5316\u4e0e\u7701\u6c34\u5229\u3001\u76d1\u7ba1\u90e8\u95e8\u6570\u636e\u5e73\u53f0\u7684\u5bf9\u63a5\uff0c\u652f\u6491\u76d1\u7ba1\u8981\u6c42", C.pub],
    ["\u6570\u5b57\u5b6a\u751f\u5168\u573a\u666f\u8986\u76d6", "\u4ece\u751f\u4ea7\u533a\u6269\u5c55\u5230\u6574\u4e2a\u7802\u77f3\u57fa\u5730\uff0c\u5b9e\u73b01:1\u6570\u5b57\u5b6a\u751f\u5168\u8986\u76d6", C.acc],
  ];
  fs.forEach(function(f, i){
    var y = 1.0 + i * 0.72;
    scr(sl, 0.5, y, 9, 0.58, C.card, 0.08);
    srec(sl, 0.5, y, 0.04, 0.58, f[2]);
    sl.addShape(pres.shapes.PARALLELOGRAM, { x: 0.65, y: y+0.1, w: 0.45, h: 0.38, fill: { color: f[2] } });
    sl.addText(String(i+1), {x:0.65, y:y+0.1, w:0.45, h:0.38, fontSize:14, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(f[0], {x:1.25, y:y+0.06, w:2.5, h:0.22, fontSize:11, fontFace:F, color:f[2], bold:true});
    sl.addText(f[1], {x:1.25, y:y+0.3, w:7.8, h:0.25, fontSize:8.5, fontFace:F, color:C.text});
  });
  badge(sl, p, C.gov);
}

// ====== S12 THANK YOU ======
function s12() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  sl.addText("\u611f\u8c22\u804a\u542c", { x: 1, y: 1.2, w: 8, h: 0.8, fontSize: 34, fontFace: F, color: C.text, bold: true, align: "center" });
  srec(sl, 4.2, 2.0, 1.6, 0.03, C.acc);
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 1, y: 2.15, w: 8, h: 0.4, fontSize: 14, fontFace: F, color: C.acc, align: "center" });
  var fts = ["\u4e24\u6cb3\u53e3\u7802\u77f3\u667a\u80fd\u5efa\u9020\u57fa\u5730", "\u6280\u672f\u652f\u6301\uff1a\u9752\u8fc5\u79d1\u6280", "\u5ba2\u6237\u6848\u4f8b\uff1a\u7ef5\u7af9\u5efa\u6295\u5efa\u6750 / \u4e2d\u56fd\u79fb\u52a8 / \u6210\u90fd\u7a7a\u6e2f\u5174\u57ce"];
  fts.forEach(function(f, i){
    sl.addText(f, { x: 1, y: 3.4 + i*0.35, w: 8, h: 0.3, fontSize: 10, fontFace: F, color: C.gray, align: "center" });
  });
}

// BUILD
s1(); s2(); s3(); s4(); s5(); s6();
s7(); s8(); s9(); s10(); s11(); s12();

pres.writeFile({ fileName: path.join(__dirname, "output", "lianghekou_sand_stone_ppt.pptx") })
  .then(function() { console.log("PPT OK: " + p + " slides"); })
  .catch(function(e) { console.error("ERR:", e); });
