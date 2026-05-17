const pptxgen = require("pptxgenjs");
const path = require("path");
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "\u9752\u8fc5\u79d1\u6280";
pres.title = "\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u667a\u6167\u5e73\u53f0\u4ea7\u54c1\u4ecb\u7ecd";

// ====== COLOR ======
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
  sl.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.35, h: 0.35, fill: { color: c || C.gov } });
  sl.addText(String(n), { x: 9.3, y: 5.1, w: 0.35, h: 0.35, fontSize: 9, fontFace: F, color: "FFFFFF", bold: true, align: "center", valign: "middle" });
}
function tb(sl, c) { sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.035, fill: { color: c || C.gov } }); }
function sco(sl, x, y, s, c) { sl.addShape(pres.shapes.OVAL, { x, y, w: s, h: s, fill: { color: c } }); }
function scr(sl, x, y, w, h, c, r) {
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: c }, rectRadius: r||0.06,
    shadow: { type: "outer", blur: 4, offset: 1.5, color: "000000", opacity: 0.04 } });
}
function srec(sl, x, y, w, h, c) { sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: c } }); }

// Decor: accent shape on title
function titleDeco(sl, title, subtitle, accent) {
  var c = accent || C.gov;
  sl.addText(title, { x: 0.5, y: 0.12, w: 8.5, h: 0.5, fontSize: 20, fontFace: F, color: C.text, bold: true });
  srec(sl, 0.5, 0.6, 0.7, 0.025, c);
  if (subtitle) sl.addText(subtitle, { x: 0.5, y: 0.66, w: 8.5, h: 0.22, fontSize: 9, fontFace: F, color: C.gray });
  // Decorative parallelograms top-right
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.2, y: 0.08, w: 0.5, h: 0.22, fill: { color: c, transparency: 25 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.8, y: 0.08, w: 0.5, h: 0.22, fill: { color: c, transparency: 45 } });
}

// ====== S1: COVER (light) ======
function s1() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  // Decorative light elements
  sl.addShape(pres.shapes.OVAL, { x: 6.5, y: -1.2, w: 4.5, h: 4.5, fill: { color: C.govL } });
  sl.addShape(pres.shapes.OVAL, { x: 8, y: -0.5, w: 2.5, h: 2.5, fill: { color: C.mgmtL } });
  sl.addShape(pres.shapes.OVAL, { x: -0.8, y: 3, w: 3, h: 3, fill: { color: C.accL } });
  // Decorative parallelograms
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 7.0, y: 0.1, w: 0.5, h: 0.25, fill: { color: C.gov, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 7.6, y: 0.1, w: 0.5, h: 0.25, fill: { color: C.mgmt, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.2, y: 0.1, w: 0.5, h: 0.25, fill: { color: C.ent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.8, y: 0.1, w: 0.5, h: 0.25, fill: { color: C.pub, transparency: 30 } });
  // Left accent bar
  srec(sl, 0.5, 0.5, 0.06, 2.8, C.acc);
  sl.addText("\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a", { x: 0.8, y: 0.5, w: 8, h: 0.35, fontSize: 16, fontFace: F, color: C.gray });
  sl.addText("\u667a\u6167\u5e73\u53f0\u4ea7\u54c1\u4ecb\u7ecd", { x: 0.8, y: 0.85, w: 7, h: 1.0, fontSize: 36, fontFace: F, color: C.text, bold: true });
  sl.addText("\u56db\u7c7b\u89d2\u8272 \u00b7 \u5168\u6a21\u5757\u667a\u6167\u5316 \u00b7 \u6570\u636e\u9a71\u52a8\u51b3\u7b56", { x: 0.8, y: 1.7, w: 6, h: 0.35, fontSize: 12, fontFace: F, color: C.gray });
  // Role tags
  var rs = [[C.gov,"\u653f\u5e9c\u9886\u5bfc\u5c42"],[C.mgmt,"\u56ed\u533a\u7ba1\u7406\u5c42"],[C.ent,"\u5165\u9a7b\u4f01\u4e1a"],[C.pub,"\u516c\u4f17\u7528\u6237"]];
  rs.forEach(function(r,i){
    var x=0.8+i*2.2;
    scr(sl, x, 2.3, 2.0, 0.38, r[0], 0.06);
    sl.addText(r[1],{x,y:2.3,w:2.0,h:0.38,fontSize:10,fontFace:F,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
  });
  // Bottom info card
  scr(sl, 0.5, 3.0, 9, 0.7, C.card, 0.1);
  srec(sl, 0.5, 3.55, 9, 0.15, C.acc);
  var infos = [["\u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf","\u7af9\u56ed\u9547\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a"],["89+\u529f\u80fd\u6a21\u5757","12\u5927\u6570\u636e\u5bf9\u63a5"],["47\u5bb6\u5165\u9a7b\u4f01\u4e1a","3,000+\u4ece\u4e1a\u4eba\u5458"]];
  infos.forEach(function(n,i){
    var x=0.8+i*3.0;
    sl.addText(n[0],{x,y:3.1,w:2.8,h:0.2,fontSize:10,fontFace:F,color:C.text,bold:true,align:"center"});
    sl.addText(n[1],{x,y:3.3,w:2.8,h:0.2,fontSize:9,fontFace:F,color:C.gray,align:"center"});
  });
}

// ====== S2: TOC (grid) ======
function s2() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  sl.addText("\u4ea7\u54c1\u76ee\u5f55", { x: 0.5, y: 0.1, w: 5, h: 0.5, fontSize: 22, fontFace: F, color: C.text, bold: true });
  // Grid layout: 4 columns x 2 rows
  var ses = [
    ["01", "\u4ea7\u54c1\u6982\u8ff0", "\u80cc\u666f\u4e0e\u67b6\u6784", C.gov],
    ["02", "\u4ea7\u54c1\u7279\u70b9", "\u6838\u5fc3\u4f18\u52bf", C.acc],
    ["03", "\u4ea7\u54c1\u6536\u76ca", "\u4ef7\u503c\u4e0e\u6548\u76ca", C.mgmt],
    ["04", "\u89d2\u8272\u529f\u80fd\u603b\u89c8", "\u56db\u7c7b\u89d2\u8272\u5b9a\u4f4d", C.pub],
    ["05", "\u653f\u5e9c\u9886\u5bfc\u5c42", "\u6570\u636e\u51b3\u7b56 \u00b7 AI\u62db\u5546", C.gov],
    ["06", "\u56ed\u533a\u7ba1\u7406\u5c42", "\u667a\u6167\u8fd0\u8425 \u00b7 \u5168\u57df\u76d1\u63a7", C.mgmt],
    ["07", "\u5165\u9a7b\u4f01\u4e1a", "\u4e00\u7ad9\u5f0f\u670d\u52a1 \u00b7 \u8d44\u6e90\u5bf9\u63a5", C.ent],
    ["08", "\u516c\u4f17\u7528\u6237", "\u4fbf\u6c11\u670d\u52a1 \u00b7 \u56ed\u533a\u53c2\u4e0e", C.pub]
  ];
  ses.forEach(function(s, i){
    var x = 0.5 + (i%4) * 2.3, y = 0.8 + Math.floor(i/4) * 2.0;
    scr(sl, x, y, 2.15, 1.8, C.card, 0.12);
    srec(sl, x, y, 2.15, 0.06, s[3]);
    sco(sl, x+0.1, y+0.2, 0.46, s[3]);
    sl.addText(s[0], {x:x+0.1, y:y+0.2, w:0.46, h:0.46, fontSize:15, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(s[1], {x:x+0.65, y:y+0.18, w:1.4, h:0.25, fontSize:13, fontFace:F, color:s[3], bold:true});
    sl.addText(s[2], {x:x+0.65, y:y+0.43, w:1.4, h:0.2, fontSize:9, fontFace:F, color:C.gray});
    // Dotted line
    srec(sl, x+0.1, y+0.8, 1.95, 0.005, C.line);
    // Feature preview
    var feats = s[2].split(" ");
    feats.forEach(function(f,fi){ sl.addText("\u25b8 "+f, {x:x+0.12, y:y+0.9+fi*0.32, w:1.9, h:0.3, fontSize:8, fontFace:F, color:C.gray}); });
  });
  badge(sl, p, C.gov);
}

// ====== S3: PRODUCT OVERVIEW ======
function s3() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u4ea7\u54c1\u6982\u8ff0", "\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u667a\u6167\u5e73\u53f0", C.gov);
  // Location
  scr(sl, 0.5, 0.95, 9, 0.4, C.govL, 0.06);
  sl.addText("\u9752\u5ddd\u53bf\u4f4d\u4e8e\u56db\u5ddd\u7701\u5317\u90e8\u8fb9\u7f18\uff0c\u5ddd\u7518\u9655\u4e09\u7701\u7ed3\u5408\u90e8\uff0c\u7d20\u6709\u201c\u91d1\u4e09\u89d2\u201d\u4e4b\u79f0\u3002\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u89c4\u5212\u9762\u79ef2.1\u33a1\uff0c\u5df2\u5f00\u53d11.6\u33a1\u3002",
    { x: 0.7, y: 0.96, w: 8.6, h: 0.38, fontSize: 10, fontFace: F, color: C.text, valign: "middle" });
  // 4 metrics
  var md = [
    ["57.68\u4ebf", "\u5168\u53bfGDP", "\u524d\u4e09\u5b63\u540c\u6bd4\u589e\u957f15%", C.gov],
    ["137\u4ebf", "\u56fd\u4f01\u8d44\u4ea7", "\u53bf\u5c5e\u56fd\u4f01\u8425\u65363.44\u4ebf", C.mgmt],
    ["45\u5bb6", "\u79d1\u6280\u4e2d\u5c0f\u4f01", "\u65b0\u589e\u9ad8\u65b0\u6280\u672f\u4f018\u5bb6", C.ent],
    ["47\u5bb6", "\u56ed\u533a\u5165\u9a7b\u4f01\u4e1a", "\u4ece\u4e1a\u4eba\u54583,000+", C.pub]
  ];
  md.forEach(function(m, i){
    var x = 0.5 + i * 2.3, y = 1.6;
    scr(sl, x, y, 2.1, 1.15, C.card, 0.1);
    sl.addText(m[0], {x, y:y+0.08, w:2.1, h:0.4, fontSize:22, fontFace:"Arial", color:m[3], bold:true, align:"center"});
    sl.addText(m[1], {x, y:y+0.48, w:2.1, h:0.2, fontSize:10, fontFace:F, color:C.text, bold:true, align:"center"});
    sl.addText(m[2], {x, y:y+0.72, w:2.1, h:0.2, fontSize:8, fontFace:F, color:C.gray, align:"center"});
    srec(sl, x+0.2, y+1.05, 1.7, 0.02, m[3]);
  });
  // Architecture - 5 layers
  sl.addText("\u5e73\u53f0\u67b6\u6784\uff1a\u4e94\u5c42\u4e00\u4f53\u5316\u8bbe\u8ba1", { x: 0.5, y: 3.0, w: 5, h: 0.25, fontSize: 11, fontFace: F, color: C.text, bold: true });
  var arch = [
    ["C\u7aef\u5c55\u793a\u5c42", "\u5927\u5c4f\u6570\u636e\u9a7e\u9a76\u8231 | \u56ed\u533a\u5b98\u7f51 | \u5411\u5bfc\u9875 | H5\u79fb\u52a8\u7aef | \u5fae\u4fe1\u5c0f\u7a0b\u5e8f", C.gov],
    ["\u5e94\u7528\u670d\u52a1\u5c42", "AI\u62db\u5546 | \u5b89\u5168\u76d1\u63a7 | \u4f01\u4e1a\u670d\u52a1 | \u8fd0\u8425\u7ba1\u7406 | \u80fd\u6548\u76d1\u6d4b", C.mgmt],
    ["\u6570\u636e\u670d\u52a1\u5c42", "\u6570\u636e\u6e56 | \u5206\u6790\u5f15\u64ce | AI\u7b97\u6cd5 | GIS\u5e73\u53f0 | \u5b9e\u65f6\u8ba1\u7b97", C.ent],
    ["\u8bbe\u65bd\u7269\u8054\u5c42", "\u89c6\u9891\u76d1\u63a7 | \u4f20\u611f\u5668\u7f51\u7edc | \u667a\u80fd\u8868\u8ba1 | GPS\u5b9a\u4f4d | \u73af\u5883\u76d1\u6d4b", C.pub],
    ["\u6570\u636e\u5bf9\u63a5\u5c42", "\u5929\u7136\u6c14\u00b7\u7535\u529b\u00b7\u6c34\u00b7\u7efc\u5408\u6267\u6cd5\u00b7\u73af\u4fdd\u00b7\u5e94\u6025\u00b7\u4ea4\u901a\u00b7\u516c\u5b89\u00b7\u81ea\u7136\u8d44\u6e90\u00b7\u6c14\u8c61\u00b7\u4eba\u793e\u00b7\u8fd0\u8425\u5546", "475569"]
  ];
  arch.forEach(function(a, i){
    var y = 3.35 + i * 0.38;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.32, fill: { color: a[2] }, rectRadius: 0.04,
      shadow: { type: "outer", blur: 2, offset: 1, color: "000000", opacity: 0.03 } });
    sl.addText(a[0], { x: 0.6, y, w: 1.5, h: 0.32, fontSize: 8, fontFace: F, color: "FFFFFF", bold: true, valign: "middle" });
    sl.addText(a[1], { x: 2.2, y, w: 7.1, h: 0.32, fontSize: 7, fontFace: F, color: "FFFFFF", valign: "middle" });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.95, y: y+0.02, w: 0.03, h: 0.28, fill: { color: "FFFFFF", transparency: 50 }, rectRadius: 0.01 });
  });
  badge(sl, p, C.gov);
}

// ====== S4: PRODUCT FEATURES ======
function s4() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  titleDeco(sl, "\u4ea7\u54c1\u7279\u70b9", "\u56db\u5927\u6838\u5fc3\u4f18\u52bf\u9a71\u52a8\u56ed\u533a\u667a\u6167\u5316\u5347\u7ea7", C.acc);
  var feats = [
    ["\u591a\u89d2\u8272\u8986\u76d6", "\u5168\u529f\u80fd\u670d\u52a1",
     "\u4e3a\u56ed\u533a\u653f\u5e9c\u3001\u7ba1\u7406\u65b9\u3001\u4f01\u4e1a\u3001\u516c\u4f17\u56db\u7c7b\u89d2\u8272\u63d0\u4f9b\u5b9a\u5236\u5316\u529f\u80fd\u4e0e\u754c\u9762\uff0c\u6bcf\u4e2a\u89d2\u8272\u72ec\u7acb\u6743\u9650\u4f53\u7cfb\uff0c\u4e00\u4e2a\u5e73\u53f0\u6ee1\u8db3\u6240\u6709\u56ed\u533a\u4eba\u5458\u9700\u6c42",
     C.gov],
    ["\u6570\u636e\u9a71\u52a8\u51b3\u7b56", "\u5168\u666f\u6570\u636e\u53ef\u89c6\u5316",
     "\u6c47\u805a\u56ed\u533a\u5168\u91cf\u6570\u636e\uff0c\u63d0\u4f9b\u5b9e\u65f6\u6570\u636e\u5206\u6790\u4e0e\u53ef\u89c6\u5316\u5448\u73b0\uff0c\u8d4b\u80fd\u7ba1\u7406\u8005\u505a\u51fa\u79d1\u5b66\u51b3\u7b56\uff0c\u8ba9\u6570\u636e\u8bf4\u8bdd",
     C.mgmt],
    ["\u5168\u6a21\u5757\u96c6\u6210", "\u4e00\u4f53\u5316\u667a\u6167\u5e73\u53f0",
     "\u6db5\u76d6AI\u62db\u5546\u3001\u5168\u57df\u5b89\u5168\u76d1\u63a7\u3001\u4f01\u4e1a\u670d\u52a1\u3001\u80fd\u6548\u76d1\u6d4b\u3001\u6570\u5b57\u5316\u529e\u516c\u7b4989+\u529f\u80fd\u6a21\u5757\uff0c\u4e00\u4e2a\u5e73\u53f0\u5b8c\u6210\u6240\u6709\u5de5\u4f5c",
     C.ent],
    ["\u591a\u6570\u636e\u6e90\u5bf9\u63a5", "\u7edf\u4e00\u6570\u636e\u6807\u51c6",
     "\u5bf9\u63a512\u5927\u7cfb\u7edf\u6570\u636e\u6e90\uff08\u6c14\u3001\u7535\u3001\u6c34\u3001\u7efc\u5408\u6267\u6cd5\u3001\u73af\u4fdd\u3001\u5e94\u6025\u3001\u4ea4\u901a\u3001\u516c\u5b89\u3001\u81ea\u7136\u8d44\u6e90\u3001\u6c14\u8c61\u3001\u4eba\u793e\u3001\u8fd0\u8425\u5546\uff09\uff0c\u6253\u7834\u6570\u636e\u5b64\u5c9b",
     C.pub]
  ];
  feats.forEach(function(f, i){
    var x = 0.5 + (i%2) * 4.65, y = 0.95 + Math.floor(i/2) * 2.0;
    scr(sl, x, y, 4.35, 1.8, C.card, 0.12);
    srec(sl, x, y, 4.35, 0.05, f[3]);
    sco(sl, x+0.15, y+0.2, 0.5, f[3]);
    sl.addText(String(i+1), {x:x+0.15, y:y+0.2, w:0.5, h:0.5, fontSize:18, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(f[0], {x:x+0.75, y:y+0.18, w:3.4, h:0.25, fontSize:14, fontFace:F, color:f[3], bold:true});
    sl.addText(f[1], {x:x+0.75, y:y+0.42, w:3.4, h:0.2, fontSize:9, fontFace:F, color:C.gray});
    srec(sl, x+0.15, y+0.75, 4.05, 0.005, C.line);
    sl.addText(f[2], {x:x+0.15, y:y+0.85, w:4.05, h:0.75, fontSize:8.5, fontFace:F, color:C.text});
    // Bottom tag
    var tagColor = f[3];
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.15, y: y+1.55, w: 1.3, h: 0.18, fill: { color: tagColor }, rectRadius: 0.02 });
    sl.addText("\u6838\u5fc3\u4f18\u52bf", { x: x+0.15, y: y+1.55, w: 1.3, h: 0.18, fontSize: 6.5, fontFace: F, color: "FFFFFF", align: "center", valign: "middle" });
  });
  badge(sl, p, C.acc);
}

// ====== S5: BENEFITS ======
function s5() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.mgmt);
  titleDeco(sl, "\u4ea7\u54c1\u6536\u76ca", "\u667a\u6167\u5e73\u53f0\u5e26\u6765\u7684\u591a\u7ef4\u5ea6\u4ef7\u503c\u4e0e\u6548\u76ca\u63d0\u5347", C.mgmt);
  // Decorative elements
  sco(sl, 9.2, 0.15, 0.08, C.mgmtL);
  sco(sl, 9.4, 0.15, 0.05, C.mgmt);
  // 3 big benefit cards
  var benefits = [
    [C.gov, "\u7ba1\u7406\u6548\u7387\u63d0\u5347", "+30~60%",
     ["\u5e94\u6025\u54cd\u5e94\u65f6\u95f4\u7f29\u77ed60%", "\u65e5\u5e38\u5ba1\u6279\u6d41\u7a0b\u7ebf\u4e0a\u5316\uff0c\u8282\u770150%+\u65f6\u95f4", "\u8bbe\u5907\u8fdc\u7a0b\u7ba1\u63a7\uff0c\u7ef4\u62a4\u6210\u672c\u964d\u4f4e40%", "\u6570\u636e\u81ea\u52a8\u6c47\u62a5\uff0c\u51cf\u5c11\u4eba\u5de5\u7edf\u8ba1\u5de5\u4f5c"]],
    [C.mgmt, "\u8fd0\u8425\u6210\u672c\u4f18\u5316", "\u00a5\u7ea629%\u8282\u7701",
     ["\u80fd\u8017\u76d1\u63a7\u51cf\u5c11\u80fd\u6e90\u6d6a\u8d39\u652f\u51fa25~35%", "\u5728\u7ebf\u670d\u52a1\u51cf\u5c11\u4eba\u5de5\u7ebf\u4e0b\u670d\u52a1\u6210\u672c", "\u4f9b\u9700\u5bf9\u63a5\u63d0\u5347\u7269\u6d41\u8fd0\u4f5c\u6548\u7387\uff0c\u964d\u4f4e\u7a7a\u9a76\u7387", "\u6570\u5b57\u5316\u529e\u516c\u51cf\u5c11\u7eb8\u5f20\u6d88\u8017300%"]],
    [C.ent, "\u5546\u4e1a\u4ef7\u503c\u63d0\u5347", "\u00a520\u4ebf+\u5e74",
     ["\u5bf9\u63a5\u7701\u5e02\u6bcf\u5e74\u8d8520\u4ebf\u667a\u6539\u6570\u8f6c\u4e13\u9879\u8d44\u91d1", "AI\u7cbe\u51c6\u62db\u5546\u5f15\u8d44\u63d0\u5347\u9879\u76ee\u626e\u6c14\u7387", "\u7269\u6d41\u4f9b\u9700\u5bf9\u63a5\u964d\u4f4e\u8fd0\u8f93\u6210\u672c", "\u653f\u7b56\u7cbe\u51c6\u63a8\u9001\u786e\u4fdd\u8865\u8d34\u4e0d\u6f0f\u62ff"]]
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
  // Bottom summary
  scr(sl, 0.5, 5.1, 9, 0.25, C.govL, 0.05);
  sl.addText("\u7efc\u5408\u6548\u76ca\uff1a\u56ed\u533a\u667a\u6167\u5316\u5efa\u8bbe\u5e26\u6765\u7ba1\u7406\u6548\u7387\u3001\u8fd0\u8425\u6210\u672c\u3001\u5546\u4e1a\u4ef7\u503c\u4e09\u7ef4\u5ea6\u63d0\u5347\uff0c\u77ed\u671f\u6295\u5165\u957f\u671f\u56de\u62a5\uff0c\u6301\u7eed\u8d4b\u80fd\u56ed\u533a\u53d1\u5c55",
    { x: 0.7, y: 5.12, w: 8.6, h: 0.2, fontSize: 9, fontFace: F, color: C.gov, bold: true });
  badge(sl, p, C.mgmt);
}

// ====== S6: ROLE OVERVIEW ======
function s6() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  sl.addText("\u56db\u7c7b\u89d2\u8272\u529f\u80fd\u603b\u89c8", { x: 0.5, y: 0.12, w: 8, h: 0.5, fontSize: 20, fontFace: F, color: C.text, bold: true });
  sl.addText("\u5e73\u53f0\u4e3a\u56ed\u533a\u56db\u7c7b\u6838\u5fc3\u7528\u6237\u63d0\u4f9b\u5b9a\u5236\u5316\u667a\u6167\u670d\u52a1", { x: 0.5, y: 0.52, w: 8, h: 0.22, fontSize: 9, fontFace: F, color: C.gray });
  // Decorative dots
  sco(sl, 9.0, 0.12, 0.1, C.govL);
  sco(sl, 9.2, 0.12, 0.06, C.gov);
  // 4 role cards
  var rc = [
    ["\u653f\u5e9c\u9886\u5bfc\u5c42", C.gov, "0F172A", "\u6570\u636e\u9a71\u52a8\u51b3\u7b56", ["AI\u62db\u5546\u670d\u52a1", "\u6570\u5b57\u6c99\u76d8", "\u4ea7\u4e1a\u94fe\u56fe\u8c31", "\u6295\u8d44\u6982\u7b97", "\u76d1\u7763\u7edf\u8ba1", "\u5168\u57df\u5b89\u5168\u76d1\u63a7"]],
    ["\u56ed\u533a\u7ba1\u7406\u5c42", C.mgmt, "064E3B", "\u667a\u6167\u8fd0\u8425\u7ba1\u63a7", ["\u9053\u8def\u4ea4\u901a\u76d1\u63a7", "\u5730\u707e\u5b89\u5168\u76d1\u6d4b", "\u73af\u4fdd\u76d1\u63a7\u68c0\u6d4b", "\u89c6\u9891\u76d1\u63a7\u603b\u89c8", "\u5e94\u6025\u6307\u6325\u8054\u52a8", "\u80fd\u6548\u76d1\u6d4b\u5206\u6790"]],
    ["\u5165\u9a7b\u4f01\u4e1a", C.ent, "78350F", "\u4e00\u7ad9\u5f0f\u4f01\u4e1a\u670d\u52a1", ["\u8bc9\u6c42\u76f4\u8fbe", "\u670d\u52a1\u76f4\u8fbe", "\u4f9b\u9700\u5bf9\u63a5", "\u7f34\u8d39\u670d\u52a1", "\u573a\u9986\u9910\u996e", "\u505c\u8f66\u7528\u5de5"]],
    ["\u516c\u4f17\u7528\u6237", C.pub, "4C1D95", "\u4fbf\u6c11\u4fe1\u606f\u670d\u52a1", ["\u56ed\u533a\u516c\u544a", "\u56ed\u533a\u6d3b\u52a8", "\u7279\u8272\u4f9b\u9700", "\u56ed\u533a\u9910\u996e", "\u505c\u8f66\u670d\u52a1", "\u5fae\u4fe1\u5c0f\u7a0b\u5e8f"]]
  ];
  var roleLights = [C.govL, C.mgmtL, C.entL, C.pubL];
  rc.forEach(function(r, i){
    var x = 0.5 + i * 2.3;
    scr(sl, x, 0.85, 2.1, 3.9, C.card, 0.12);
    srec(sl, x, 0.85, 2.1, 0.06, r[1]);
    sl.addText(r[0], {x, y:1.0, w:2.1, h:0.3, fontSize:13, fontFace:F, color:r[1], bold:true, align:"center"});
    srec(sl, x+0.3, 1.35, 1.5, 0.02, r[1]);
    sl.addText(r[3], {x, y:1.45, w:2.1, h:0.2, fontSize:9, fontFace:F, color:C.gray, align:"center"});
    // Feature list
    r[4].forEach(function(f, fi){
      var fy = 1.8 + fi * 0.42;
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.1, y: fy, w: 1.9, h: 0.35, fill: { color: roleLights[i] }, rectRadius: 0.04 });
      sco(sl, x+0.2, fy+0.08, 0.06, r[1]);
      sl.addText(f, {x:x+0.35, y:fy, w:1.6, h:0.35, fontSize:8, fontFace:F, color:C.text, valign:"middle"});
    });
    // Bottom tag
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.3, y: 4.05, w: 1.5, h: 0.22, fill: { color: r[2] }, rectRadius: 0.03 });
    sl.addText("25+\u529f\u80fd", {x:x+0.3, y:4.05, w:1.5, h:0.22, fontSize:7, fontFace:F, color:"FFFFFF", align:"center", valign:"middle"});
  });
  badge(sl, p, C.acc);
}

// ====== ROLE SLIDE HELPER ======
function roleSlide(title, items, role, accent) {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, accent);
  titleDeco(sl, title, "", accent);
  // Decorative parallelogram
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: accent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: accent, transparency: 50 } });
  sl.addText(role, { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: accent, align: "right" });
  items.forEach(function(item, i) {
    var col = i % 2, row = Math.floor(i / 2);
    var x = 0.5 + col * 4.65, y = 1.05 + row * 0.85;
    scr(sl, x, y, 4.35, 0.7, C.card, 0.08);
    srec(sl, x, y, 0.04, 0.7, accent);
    // Number badge
    sco(sl, x+0.15, y+0.15, 0.15, accent);
    sl.addText(String(i+1), {x:x+0.15, y:y+0.15, w:0.15, h:0.15, fontSize:7, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(item.title || item, {x:x+0.4, y:y+0.05, w:3.8, h:0.25, fontSize:10, fontFace:F, color:accent, bold:true});
    if (item.desc) sl.addText(item.desc, {x:x+0.4, y:y+0.32, w:3.8, h:0.32, fontSize:7.5, fontFace:F, color:C.text});
    if (item.tag) {
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+3.2, y: y+0.07, w: 1.0, h: 0.18, fill: { color: "E2E8F0" }, rectRadius: 0.02 });
      sl.addText(item.tag, {x:x+3.2, y:y+0.07, w:1.0, h:0.18, fontSize:6, fontFace:F, color:accent, align:"center", valign:"middle"});
    }
  });
  badge(sl, p, accent);
}

// ====== GOV SLIDES ======
function s7() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 \u6570\u636e\u9a7e\u9a76\u8231\u4e0e\u51b3\u7b56\u652f\u6301", "", C.gov);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.gov, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.gov, transparency: 50 } });
  sl.addText(C.gov+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.gov, align: "right" });
  // Top metrics
  var gms = [["57.68\u4ebf","\u5168\u53bfGDP",C.gov],["47\u5bb6","\u56ed\u533a\u4f01\u4e1a","2563EB"],["137\u4ebf","\u56fd\u4f01\u8d44\u4ea7","1D4ED8"]];
  gms.forEach(function(m,i){
    var x=0.5+i*3.1;
    scr(sl, x, 0.85, 2.8, 0.6, C.card, 0.08);
    sl.addText(m[0],{x:x+0.12,y:0.87,w:1.3,h:0.55,fontSize:20,fontFace:"Arial",color:m[2],bold:true,valign:"middle"});
    sl.addText(m[1],{x:x+1.5,y:0.87,w:1.2,h:0.55,fontSize:10,fontFace:F,color:C.gray,valign:"middle"});
  });
  // Feature grid 2x3
  var gfs = [
    ["AI\u62db\u5546\u670d\u52a1", "\u5927\u6570\u636e+AI\u7cbe\u51c6\u5339\u914d\u76ee\u6807\u4f01\u4e1a", "\u591a\u7ef4\u6570\u636e\u5206\u6790"],
    ["\u6570\u5b57\u6c99\u76d8", "3D\u53ef\u89c6\u5316\u5c55\u793a\u56ed\u533a\u89c4\u5212\u4e0e\u571f\u5730\u8d44\u6e90", "\u6c89\u6d78\u5f0f\u62db\u5546"],
    ["\u4ea7\u4e1a\u94fe\u56fe\u8c31", "\u53ef\u89c6\u5316\u5c55\u793a\u4ea7\u4e1a\u94fe\u4e0a\u4e0b\u6e38\u5173\u7cfb", "\u5f3a\u94fe\u8865\u94fe"],
    ["\u6295\u8d44\u6982\u7b97", "\u4e09\u6b65\u8d70\u521b\u5efa\u8def\u5f84+12\u5927\u4e13\u9879\u6295\u8d44\u8be6\u7b97", "\u79d1\u5b66\u89c4\u5212"],
    ["\u76d1\u7763\u7edf\u8ba1", "\u7edf\u4e00\u6570\u636e\u51fa\u53e3\uff0c\u652f\u6301\u591a\u79cd\u683c\u5f0f\u5bfc\u51fa", "\u6570\u636e\u6c47\u62a5"],
    ["\u7eff\u8272\u56ed\u533a", "\u73af\u4fdd\u76d1\u63a7\u6570\u636e\u53ef\u8ffd\u6eaf\uff0c\u52a9\u529b\u7eff\u8272\u521b\u5efa", "\u7eff\u8272\u521b\u5efa"]
  ];
  gfs.forEach(function(f, i){
    var col=i%3, row=Math.floor(i/3);
    var x=0.5+col*3.1, y=1.7+row*1.3;
    scr(sl, x, y, 2.9, 1.1, C.card, 0.1);
    srec(sl, x, y, 2.9, 0.04, C.gov);
    sl.addText(f[0], {x:x+0.12, y:y+0.1, w:2.7, h:0.22, fontSize:10, fontFace:F, color:C.gov, bold:true});
    sl.addText(f[1], {x:x+0.12, y:y+0.35, w:2.7, h:0.45, fontSize:7.5, fontFace:F, color:C.text});
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.12, y: y+0.82, w: 1.2, h: 0.18, fill: { color: C.govL }, rectRadius: 0.02 });
    sl.addText(f[2], {x:x+0.12, y:y+0.82, w:1.2, h:0.18, fontSize:6.5, fontFace:F, color:C.gov, align:"center", valign:"middle"});
  });
  badge(sl, p, C.gov);
}

// ====== S8: GOV - AI RECRUIT (timeline) ======
function s8() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 AI\u62db\u5546\u5168\u6d41\u7a0b", "", C.gov);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.gov, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.gov, transparency: 50 } });
  sl.addText(C.gov+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.gov, align: "right" });
  // Timeline
  var steps = [
    ["\u76ee\u6807\u4f01\u4e1a\u6316\u6398", "\u591a\u7ef4\u6570\u636e\u6e90+\u667a\u80fd\u7b97\u6cd5\uff0c\u81ea\u52a8\u63a8\u8350\u4e0e\u56ed\u533a\u4ea7\u4e1a\u9ad8\u5ea6\u5339\u914d\u7684\u76ee\u6807\u4f01\u4e1a", C.gov],
    ["AI\u667a\u80fd\u5339\u914d", "\u57fa\u4e8e\u4ea7\u4e1a\u94fe\u3001\u8d44\u91d1\u3001\u6280\u672f\u3001\u5730\u7406\u591a\u7ef4\u5ea6\u7b97\u6cd5\u5339\u914d\uff0c\u667a\u80fd\u8bc4\u5206", "2563EB"],
    ["\u667a\u80fd\u63a8\u8350", "\u81ea\u52a8\u751f\u6210\u62db\u5546\u7b56\u7565\u4e0e\u63a5\u89e6\u65b9\u6848\uff0c\u667a\u80fd\u63a8\u8350\u63a5\u89e6\u65f6\u673a", "1D4ED8"],
    ["\u5168\u7a0b\u8ddf\u8e2a", "\u62db\u5546\u9879\u76ee\u5168\u751f\u547d\u5468\u671f\u7ba1\u7406\uff0c\u4ece\u63a5\u89e6\u5230\u7b7e\u7ea6\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316", C.gov],
    ["\u6570\u636e\u62a5\u544a", "\u62db\u5546\u8f6c\u5316\u7387\u3001\u5408\u7ea6\u91d1\u989d\u3001\u4ea7\u4e1a\u805a\u96c6\u5ea6\u591a\u7ef4\u5ea6\u5206\u6790\u62a5\u544a", "2563EB"]
  ];
  steps.forEach(function(s, i){
    var y = 0.85 + i * 0.85;
    sco(sl, 0.5, y+0.08, 0.44, s[2]);
    sl.addText(String(i+1), {x:0.5, y:y+0.08, w:0.44, h:0.44, fontSize:13, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    scr(sl, 1.1, y, 8.3, 0.65, C.card, 0.08);
    sl.addText(s[0], {x:1.25, y:y+0.04, w:3, h:0.25, fontSize:11, fontFace:F, color:s[2], bold:true});
    sl.addText(s[1], {x:1.25, y:y+0.32, w:7.8, h:0.28, fontSize:8, fontFace:F, color:C.text});
    if (i < steps.length-1) srec(sl, 0.7, y+0.52, 0.03, 0.33, C.line);
    // Right tag
    var tags = ["\u6570\u636e\u9a71\u52a8","\u667a\u80fd\u7b97\u6cd5","\u81ea\u52a8\u5316","\u5168\u6d41\u7a0b","\u5206\u6790\u62a5\u544a"];
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.2, y: y+0.08, w: 1.0, h: 0.17, fill: { color: C.govL }, rectRadius: 0.02 });
    sl.addText(tags[i], {x:8.2, y:y+0.08, w:1.0, h:0.17, fontSize:6, fontFace:F, color:C.gov, align:"center", valign:"middle"});
  });
  badge(sl, p, C.gov);
}

// ====== S9: GOV - INVEST ======
function s9() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  titleDeco(sl, "\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 \u6295\u8d44\u6982\u7b97\u4e0e\u521b\u5efa\u8def\u5f84", "", C.acc);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.acc, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.acc, transparency: 50 } });
  sl.addText(C.gov+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.gov, align: "right" });
  // Three-step path
  var paths = [
    ["\u7701\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a", "\u57fa\u4e8e\u5e73\u53f0\u6570\u636e\u652f\u6491\uff0c\u7cfb\u7edf\u5316\u6574\u7406\u7533\u62a5\u6750\u6599\uff0c\u52a0\u901f\u521b\u5efa\u8fdb\u7a0b", C.gov],
    ["\u7eff\u8272\u5de5\u4e1a\u56ed\u533a", "\u73af\u4fdd\u76d1\u63a7\u3001\u80fd\u8017\u6570\u636e\u3001\u6392\u653e\u6307\u6807\u5168\u7a0b\u53ef\u8ffd\u6eaf\uff0c\u79d1\u5b66\u8bc1\u660e", C.mgmt],
    ["\u56fd\u5bb6\u7ea7\u57f9\u80b2\u5bf9\u8c61", "\u4ea9\u5747\u8bba\u82f1\u96c4\u6570\u636e\u3001\u4ea7\u4e1a\u805a\u96c6\u5ea6\u3001\u79d1\u6280\u521b\u65b0\u6307\u6807\u7efc\u5408\u8bc4\u4f30", C.ent]
  ];
  paths.forEach(function(pt, i){
    var x = 0.5 + i * 3.15;
    scr(sl, x, 0.9, 2.85, 1.8, C.card, 0.12);
    srec(sl, x, 0.9, 2.85, 0.06, pt[2]);
    sco(sl, x+1.15, 1.1, 0.55, pt[2]);
    sl.addText(String(i+1), {x:x+1.15, y:1.1, w:0.55, h:0.55, fontSize:18, fontFace:"Arial", color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(pt[0], {x:x+0.1, y:1.7, w:2.65, h:0.25, fontSize:11, fontFace:F, color:pt[2], bold:true, align:"center"});
    sl.addText(pt[1], {x:x+0.1, y:1.95, w:2.65, h:0.55, fontSize:8, fontFace:F, color:C.text, align:"center"});
    if (i < paths.length-1) sl.addShape(pres.shapes.RIGHT_ARROW, { x: x+2.95, y: 1.65, w: 0.25, h: 0.18, fill: { color: C.line } });
  });
  // Investment detail
  scr(sl, 0.5, 2.95, 9, 1.1, C.card, 0.1);
  sl.addText("\u667a\u6167\u5e73\u53f0\u6295\u8d44\u6982\u7b97\u603b\u89c8", { x: 0.7, y: 3.0, w: 4, h: 0.25, fontSize: 11, fontFace: F, color: C.acc, bold: true });
  var invs = [
    ["\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7", "1,580,000"],
    ["\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7", "428,600"],
    ["\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546", "442,200"],
    ["\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", "348,000"]
  ];
  invs.forEach(function(inv, i){
    var x = 0.7 + (i%2)*4.2, y = 3.3 + Math.floor(i/2)*0.32;
    sco(sl, x, y+0.03, 0.06, C.acc);
    sl.addText(inv[0], {x:x+0.12, y:y, w:2.5, h:0.22, fontSize:7.5, fontFace:F, color:C.text});

    sl.addText("\u00a5" + inv[1], {x:x+2.7, y:y, w:1.3, h:0.22, fontSize:7.5, fontFace:"Arial", color:C.acc, align:"right"});
  });
  // Total
  scr(sl, 6.3, 2.95, 3.2, 1.1, "0F172A", 0.1);
  sl.addText("\u603b\u6295\u8d44\u6982\u7b97", {x:6.4, y:3.0, w:3, h:0.2, fontSize:9, fontFace:F, color:"94A3B8", align:"center"});
  sl.addText("\u00a54,178,800", {x:6.4, y:3.2, w:3, h:0.4, fontSize:22, fontFace:"Arial", color:C.acc, bold:true, align:"center"});
  sl.addText("\u5143", {x:6.4, y:3.55, w:3, h:0.2, fontSize:9, fontFace:F, color:"94A3B8", align:"center"});
  // Bottom: benefits
  srec(sl, 0.5, 4.3, 9, 0.02, C.line);
  var gb = [["30%","\u7ba1\u7406\u6548\u7387\u63d0\u5347",C.gov],["60%","\u5e94\u6025\u54cd\u5e94\u63d0\u5347",C.mgmt],["\u00a520\u4ebf+","\u5e74\u5bf9\u63a5\u4e13\u9879\u8d44\u91d1",C.ent]];
  gb.forEach(function(b,i){
    var x=0.8+i*2.9;
    sl.addText(b[0],{x,y:4.4,w:2.2,h:0.25,fontSize:18,fontFace:"Arial",color:b[2],bold:true,align:"center"});
    sl.addText(b[1],{x,y:4.65,w:2.2,h:0.2,fontSize:8,fontFace:F,color:C.text,align:"center"});
  });
  badge(sl, p, C.acc);
}

// ====== S10: GOV - SECURITY ======
function s10() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u653f\u5e9c\u9886\u5bfc\u5c42 \u2014 \u5168\u57df\u5b89\u5168\u76d1\u63a7\u4f53\u7cfb", "", C.gov);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.gov, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.gov, transparency: 50 } });
  sl.addText(C.gov+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.gov, align: "right" });
  // 3 columns
  var sc = [
    ["\u5929\u773c", "\u9053\u8def\u4ea4\u901a\u5b89\u5168\u76d1\u63a7", "\u667a\u80fd\u4ea4\u901a\u76d1\u63a7\u00b7\u8f66\u8f86\u8bc6\u522b\u00b7\u8fdd\u7ae0\u6355\u6349\u00b7\u4ea4\u901a\u72b6\u6001\u53ef\u89c6\u5316", C.gov],
    ["\u54e8\u5175", "\u5730\u707e\u5b89\u5168\u76d1\u6d4b", "\u5730\u8d28\u707e\u5bb3\u5b9e\u65f6\u76d1\u6d4b\u00b7\u591a\u4f20\u611f\u5668\u8054\u52a8\u00b7\u9884\u8b66\u63d0\u524d\u62a5\u7ba1\u63a7", "1D4ED8"],
    ["\u6167\u773c", "\u73af\u4fdd\u76d1\u63a7\u68c0\u6d4b", "\u73af\u5883\u6307\u6807\u5b9e\u65f6\u76d1\u6d4b\u00b7\u6392\u653e\u6570\u636e\u8ffd\u8e2a\u00b7\u73af\u4fdd\u5408\u89c4\u5206\u6790", "1E40AF"]
  ];
  sc.forEach(function(c, i){
    var x = 0.5 + i * 3.15;
    scr(sl, x, 0.85, 2.9, 2.1, C.card, 0.12);
    srec(sl, x, 0.85, 2.9, 0.05, c[3]);
    // Big icon circle
    sco(sl, x+1.05, 1.0, 0.8, c[3]);
    sl.addText(c[0], {x:x+1.05, y:1.0, w:0.8, h:0.8, fontSize:16, fontFace:F, color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(c[1], {x:x+0.1, y:1.9, w:2.7, h:0.25, fontSize:11, fontFace:F, color:c[3], bold:true, align:"center"});
    sl.addText(c[2], {x:x+0.1, y:2.2, w:2.7, h:0.55, fontSize:8, fontFace:F, color:C.text, align:"center"});
  });
  // Related features
  var rels = [
    "\u89c6\u9891\u76d1\u63a7\u603b\u89c8\uff1a\u7edf\u4e00\u89c6\u9891\u5e73\u53f0\uff0c16\u5bab\u683c\u667a\u80fd\u6392\u5217\uff0c\u5168\u56ed\u533a\u8986\u76d6",
    "\u5e94\u6025\u6307\u6325\u8054\u52a8\uff1a\u4e00\u952e\u8c03\u5ea6\u3001\u9884\u6848\u7ba1\u7406\u3001\u5b9e\u65f6\u901a\u8baf\uff0c\u54cd\u5e94\u6548\u7387\u63d0\u534760%",
    "\u8bbe\u5907\u7ba1\u7406\uff1a\u6444\u50cf\u5934\u3001\u4f20\u611f\u5668\u3001GPS\u3001\u96f7\u8fbe\u7b49\u5168\u7c7b\u8bbe\u5907\u7edf\u4e00\u7ba1\u63a7",
    "\u544a\u8b66\u4e2d\u5fc3\uff1a\u5f02\u5e38\u4e8b\u4ef6\u5b9e\u65f6\u544a\u8b66\u3001\u8d8b\u52bf\u5206\u6790\u3001\u5904\u7f6e\u5168\u6d41\u7a0b"
  ];
  rels.forEach(function(r, i){
    var x = 0.5 + (i%2)*4.75, y = 3.2 + Math.floor(i/2)*0.5;
    scr(sl, x, y, 4.5, 0.4, C.card, 0.06);
    sco(sl, x+0.1, y+0.1, 0.06, C.gov);
    sl.addText(r, {x:x+0.22, y:y, w:4.1, h:0.4, fontSize:8, fontFace:F, color:C.text, valign:"middle"});
  });
  badge(sl, p, C.gov);
}

// ====== S11: MGMT - DASHBOARD ======
function s11() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.mgmt);
  titleDeco(sl, "\u56ed\u533a\u7ba1\u7406\u5c42 \u2014 \u667a\u6167\u8fd0\u8425\u4e0e\u5168\u57df\u76d1\u63a7", "", C.mgmt);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 50 } });
  sl.addText(C.mgmt+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.mgmt, align: "right" });
  // Top metrics
  var mms = [["60%","\u5e94\u6025\u54cd\u5e94\u63d0\u5347",C.mgmt],["95%","\u76d1\u63a7\u8986\u76d6\u7387","047857"],["30%","\u7ba1\u7406\u6548\u7387\u63d0\u5347","065F46"]];
  mms.forEach(function(m,i){
    var x=0.5+i*3.1;
    scr(sl, x, 0.85, 2.8, 0.55, "0F172A", 0.08);
    sl.addText(m[0],{x:x+0.12,y:0.87,w:1.2,h:0.5,fontSize:22,fontFace:"Arial",color:m[2],bold:true,valign:"middle"});
    sl.addText(m[1],{x:x+1.4,y:0.87,w:1.3,h:0.5,fontSize:10,fontFace:F,color:"94A3B8",valign:"middle"});
  });
  // 2x4 grid
  var mfs = [
    ["\u9053\u8def\u4ea4\u901a\u76d1\u63a7", "\u8f66\u8f86\u8bc6\u522b\u3001\u8fdd\u7ae0\u6355\u6349\u3001\u4ea4\u901a\u72b6\u6001\u5b9e\u65f6\u53ef\u89c6\u5316"],
    ["\u5730\u707e\u5b89\u5168\u76d1\u6d4b", "\u5730\u8d28\u707e\u5bb3\u5b9e\u65f6\u76d1\u6d4b\u3001\u591a\u4f20\u611f\u5668\u8054\u52a8\u9884\u8b66"],
    ["\u73af\u4fdd\u76d1\u63a7\u68c0\u6d4b", "\u73af\u5883\u6307\u6807\u5b9e\u65f6\u76d1\u6d4b\u3001\u6392\u653e\u6570\u636e\u8ffd\u8e2a"],
    ["\u89c6\u9891\u76d1\u63a7\u603b\u89c8", "\u7edf\u4e00\u89c6\u9891\u5e73\u53f0\u3001\u667a\u80fd\u6392\u5217\u3001\u5168\u56ed\u533a\u8986\u76d6"],
    ["\u5e94\u6025\u6307\u6325\u8054\u52a8", "\u4e00\u952e\u8c03\u5ea6\u3001\u9884\u6848\u7ba1\u7406\u3001\u5b9e\u65f6\u901a\u8baf"],
    ["\u80fd\u6548\u76d1\u6d4b\u5206\u6790", "\u667a\u80fd\u8ba1\u91cf\u8868\u5b9e\u65f6\u91c7\u96c6\u3001\u7528\u80fd\u8d8b\u52bf\u5206\u6790"],
    ["\u4ea9\u5747\u8bba\u82f1\u96c4", "\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\u3001\u8d44\u6e90\u5dee\u5f02\u5316\u914d\u7f6e"],
    ["\u5f02\u5e38\u9884\u8b66", "\u7528\u80fd\u5f02\u5e38\u81ea\u52a8\u544a\u8b66\u3001\u78b3\u6392\u7ba1\u7406\u8ffd\u8e2a"]
  ];
  mfs.forEach(function(f, i){
    var col=i%4, row=Math.floor(i/4);
    var x=0.5+col*2.3, y=1.6+row*1.45;
    scr(sl, x, y, 2.1, 1.25, C.card, 0.1);
    srec(sl, x+0.1, y+0.1, 1.9, 0.04, C.mgmt);
    sco(sl, x+0.85, y+0.25, 0.35, C.mgmtL);
    sl.addText(f[0],{x:x+0.08,y:y+0.65,w:1.94,h:0.22,fontSize:9.5,fontFace:F,color:C.mgmt,bold:true,align:"center"});
    sl.addText(f[1],{x:x+0.08,y:y+0.88,w:1.94,h:0.3,fontSize:7,fontFace:F,color:C.gray,align:"center"});
  });
  badge(sl, p, C.mgmt);
}

// ====== S12: MGMT - OPS ======
function s12() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.mgmt);
  titleDeco(sl, "\u56ed\u533a\u7ba1\u7406\u5c42 \u2014 \u8fd0\u8425\u7ba1\u7406\u4e0e\u6570\u5b57\u5316\u529e\u516c", "", C.mgmt);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 50 } });
  sl.addText(C.mgmt+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.mgmt, align: "right" });
  // Left: modules
  var mMods = [
    ["\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3", "\u5927\u5c4f\u6307\u6325+\u56ed\u5c0f\u4e8c\u56e2\u961f+3\u5e74\u957f\u6548\u670d\u52a1", C.mgmt],
    ["\u65e5\u5e38\u8fd0\u8425\u7ba1\u7406", "\u6570\u5b57\u5316\u5ba1\u6279\u3001\u65e5\u5e38\u5de1\u68c0\u3001\u8d44\u4ea7\u7ba1\u7406\u3001\u6863\u6848\u7ba1\u7406", "047857"],
    ["\u8bbe\u5907\u4e0e\u7269\u8054\u7f51", "\u6444\u50cf\u5934\u3001\u4f20\u611f\u5668\u3001\u667a\u80fd\u8bbe\u5907\u7edf\u4e00\u7ba1\u63a7\u4e0e\u8fdc\u7a0b\u8c03\u5ea6", "065F46"],
    ["\u7efc\u5408\u6570\u636e\u7ba1\u63a7", "\u6570\u636e\u4e2d\u5fc3\u3001\u6570\u636e\u8d28\u91cf\u3001\u62a5\u8868\u7ba1\u7406\u3001\u6570\u636e\u5907\u4efd\u4e0e\u5b89\u5168", C.mgmt]
  ];
  mMods.forEach(function(m, i){
    var y = 0.85 + i * 0.85;
    scr(sl, 0.5, y, 4.5, 0.7, C.card, 0.08);
    srec(sl, 0.55, y+0.08, 0.03, 0.54, m[2]);
    sl.addText(m[0], {x:0.7, y:y+0.06, w:3.8, h:0.22, fontSize:10, fontFace:F, color:m[2], bold:true});
    sl.addText(m[1], {x:0.7, y:y+0.3, w:3.8, h:0.35, fontSize:7.5, fontFace:F, color:C.text});
    // Right indicator
    sco(sl, 4.6, y+0.28, 0.04, m[2]);
  });
  // Right: stats
  scr(sl, 5.3, 0.85, 4.3, 3.25, C.card, 0.1);
  sl.addText("\u8fd0\u8425\u6570\u636e\u4e00\u89c8", {x:5.5, y:0.92, w:3.8, h:0.25, fontSize:11, fontFace:F, color:C.text, bold:true});
  srec(sl, 5.3, 1.15, 4.3, 0.015, C.mgmt);
  var os = [
    ["\u76d1\u63a7\u8bbe\u5907\u603b\u6570","128\u53f0",C.mgmt],
    ["\u65e5\u5e38\u5ba1\u6279\u6d41\u7a0b","15\u7c7b","047857"],
    ["\u8d44\u4ea7\u7ba1\u7406\u603b\u6570","2,800+\u4ef6","065F46"],
    ["\u65e5\u5747\u5904\u7406\u4e8b\u4ef6","50+\u6761",C.mgmt],
    ["\u6570\u636e\u5bf9\u63a5\u7cfb\u7edf","12\u5957","047857"],
    ["\u5728\u7ebf\u670d\u52a1\u6a21\u5757","35+\u4e2a","065F46"]
  ];
  os.forEach(function(s, i){
    var x=5.5+(i<3?0:2.1), y=1.25+(i%3)*0.65;
    sl.addText(s[0],{x,y:y,w:1.5,h:0.2,fontSize:8.5,fontFace:F,color:C.gray});
    sl.addText(s[1],{x:x+1.3,y:y-0.02,w:0.8,h:0.24,fontSize:14,fontFace:"Arial",color:s[2],bold:true,align:"right"});
    if (i<5) srec(sl, x, y+0.27, 2.1, 0.005, C.line);
  });
  badge(sl, p, C.mgmt);
}

// ====== S13: MGMT - ENERGY ======
function s13() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.mgmt);
  titleDeco(sl, "\u56ed\u533a\u7ba1\u7406\u5c42 \u2014 \u80fd\u6548\u76d1\u6d4b\u4e0e\u4ea9\u5747\u8bba\u82f1\u96c4", "", C.mgmt);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.mgmt, transparency: 50 } });
  sl.addText(C.mgmt+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.mgmt, align: "right" });
  // 2x2 grid - full-width
  var ems = [
    ["\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", "\u667a\u80fd\u8ba1\u91cf\u8868\u5b9e\u65f6\u91c7\u96c6\u80fd\u8017\u6570\u636e\uff0c\u652f\u6301\u6c34\u3001\u7535\u3001\u6c14\u591a\u7ef4\u5ea6\u76d1\u63a7\uff0c\u53ef\u89c6\u5316\u5c55\u793a\u80fd\u8017\u8d8b\u52bf\u4e0e\u5f02\u5e38\u544a\u8b66", C.mgmt],
    ["\u4ea9\u5747\u8bba\u82f1\u96c4", "\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u6295\u5165\u4ea7\u51fa\u3001\u80fd\u8017\u6392\u653e\u7b49\u591a\u7ef4\u5ea6\u6307\u6807\u7efc\u5408\u8bc4\u4ef7\uff0c\u5b9e\u73b0\u8d44\u6e90\u8981\u7d20\u5dee\u5f02\u5316\u914d\u7f6e", "047857"],
    ["\u80fd\u6548\u5206\u6790", "\u4f01\u4e1a\u80fd\u8017\u6548\u7387\u591a\u7ef4\u5ea6\u5206\u6790\uff0c\u540c\u6bd4\u73af\u6bd4\u8d8b\u52bf\u53ef\u89c6\u5316\uff0c\u8bc6\u522b\u80fd\u8017\u6d6a\u8d39\u70b9\uff0c\u4f18\u5316\u751f\u4ea7\u8fc7\u7a0b", "065F46"],
    ["\u5f02\u5e38\u9884\u8b66", "\u7528\u80fd\u5f02\u5e38\u81ea\u52a8\u544a\u8b66\uff0c\u53d1\u73b0\u80fd\u6e90\u6d6a\u8d39\u95ee\u9898\uff0c\u78b3\u6392\u653e\u7ba1\u7406\u8ffd\u8e2a\uff0c\u652f\u6301\u4f01\u4e1a\u7eff\u8272\u4f4e\u78b3\u5347\u7ea7", C.mgmt]
  ];
  ems.forEach(function(e, i){
    var x=0.5+(i%2)*4.65, y=0.85+Math.floor(i/2)*1.3;
    scr(sl, x, y, 4.35, 1.1, C.card, 0.1);
    srec(sl, x, y, 4.35, 0.04, e[2]);
    // Accent number
    sco(sl, x+0.12, y+0.15, 0.08, e[2]);
    sl.addText(e[0], {x:x+0.3, y:y+0.1, w:3.9, h:0.22, fontSize:11, fontFace:F, color:e[2], bold:true});
    sl.addText(e[1], {x:x+0.12, y:y+0.38, w:4.1, h:0.65, fontSize:7.5, fontFace:F, color:C.text});
  });
  // Data integration
  scr(sl, 0.5, 3.7, 9, 0.3, C.govL, 0.05);
  sl.addText("\u6570\u636e\u4e2d\u5fc3\u96c6\u6210\uff1a12\u5927\u4e09\u65b9\u6570\u636e\u6765\u6e90\uff08\u5929\u7136\u6c14\u3001\u7535\u529b\u3001\u6c34\u3001\u7efc\u5408\u6267\u6cd5\u3001\u73af\u4fdd\u3001\u5e94\u6025\u3001\u4ea4\u901a\u3001\u516c\u5b89\u3001\u81ea\u7136\u8d44\u6e90\u3001\u6c14\u8c61\u3001\u4eba\u793e\u3001\u8fd0\u8425\u5546\uff09\u7edf\u4e00\u5bf9\u63a5",
    { x: 0.7, y: 3.72, w: 8.6, h: 0.26, fontSize: 8.5, fontFace: F, color: C.gov, valign: "middle" });
  // Device mgmt summary
  scr(sl, 0.5, 4.15, 9, 0.65, C.card, 0.08);
  sl.addText("\u8bbe\u5907\u7ba1\u7406\u80fd\u529b", {x:0.7, y:4.2, w:3, h:0.22, fontSize:10, fontFace:F, color:C.mgmt, bold:true});
  var ds = ["\u6444\u50cf\u5934\u7ba1\u7406","GPS\u8bbe\u5907","\u96f7\u8fbe\u76d1\u6d4b","\u667a\u80fd\u79f0\u91cd","\u6570\u636e\u5e93\u7ba1\u7406","\u65e5\u5fd7\u5ba1\u8ba1"];
  ds.forEach(function(d,i){
    var x=0.7+i*1.5;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.5, w: 1.3, h: 0.22, fill: { color: C.mgmtL }, rectRadius: 0.03 });
    sl.addText(d, {x, y:4.5, w:1.3, h:0.22, fontSize:7, fontFace:F, color:C.mgmt, align:"center", valign:"middle"});
  });
  badge(sl, p, C.mgmt);
}

// ====== S14: ENT - SERVICE ======
function s14() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.ent);
  titleDeco(sl, "\u5165\u9a7b\u4f01\u4e1a \u2014 \u4f01\u4e1a\u670d\u52a1\u4e2d\u5fc3", "", C.ent);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 50 } });
  sl.addText(C.ent+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.ent, align: "right" });
  // 3 highlights
  var ehs = [
    ["\u8bc9\u6c42\u76f4\u8fbe", "\u968f\u624b\u62cd\u63d0\u4ea4\u8bc9\u6c42\uff0c\u652f\u6301\u6587\u5b57+\u56fe\u7247+\u5b9a\u4f4d\uff0c2\u5c0f\u65f6\u5185\u54cd\u5e94", C.ent],
    ["\u670d\u52a1\u76f4\u8fbe", "\u5929\u7136\u6c14/\u6c34/\u7535\u8d39\u5728\u7ebf\u7f34\u8d39\uff0c\u4e00\u7ad9\u5f0f\u4f01\u4e1a\u670d\u52a1", "B45309"],
    ["\u56ed\u4f01\u4ea4\u4e92", "\u5728\u7ebf\u6c9f\u901a\u3001\u901a\u77e5\u516c\u544a\u3001\u653f\u7b56\u63a8\u9001\u3001\u53cc\u5411\u4e92\u52a8", "92400E"]
  ];
  ehs.forEach(function(e, i){
    var x = 0.5 + i * 3.15;
    scr(sl, x, 0.8, 2.9, 1.2, C.card, 0.1);
    srec(sl, x, 0.8, 2.9, 0.04, e[2]);
    sco(sl, x+0.2, 0.92, 0.35, C.entL);
    sl.addText(e[0], {x:x+0.65, y:0.9, w:2.1, h:0.25, fontSize:12, fontFace:F, color:e[2], bold:true});
    sl.addText(e[1], {x:x+0.12, y:1.25, w:2.65, h:0.55, fontSize:8.5, fontFace:F, color:C.text});
  });
  // Features grid
  var efs = [
    ["\u4f01\u4e1a\u5165\u9a7b", "\u7ebf\u4e0a\u7533\u8bf7\u3001\u8bc1\u7167\u7ba1\u7406\u3001\u5165\u9a7b\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316"],
    ["\u4fe1\u7528\u8bc4\u4ef7", "\u4f01\u4e1a\u4fe1\u7528\u591a\u7ef4\u5ea6\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u52a9\u529b\u878d\u8d44\u5408\u4f5c"],
    ["\u5408\u540c\u7ba1\u7406", "\u7535\u5b50\u5408\u540c\u3001\u7b7e\u7f72\u3001\u5b58\u6863\u5168\u6d41\u7a0b\u7ba1\u7406"],
    ["\u8d44\u4ea7\u7ba1\u7406", "\u8d44\u4ea7\u76d8\u70b9\u3001\u5165\u5e93\u3001\u51fa\u5e93\u6570\u5b57\u5316\u7ba1\u7406"],
    ["\u6570\u5b57\u8bc1\u7167", "\u7535\u5b50\u8425\u4e1a\u6267\u7167\u3001\u8bb8\u53ef\u8bc1\u4e0a\u94fe\u7ba1\u7406"],
    ["\u5458\u5de5\u7ba1\u7406", "\u5165\u804c/\u79bb\u804c\u3001\u52a0\u73ed/\u8bf7\u5047\u3001\u5de5\u8d44\u7ba1\u7406\u5b8c\u5584\u4f53\u7cfb"]
  ];
  efs.forEach(function(f, i){
    var col=i%3, row=Math.floor(i/3);
    var x=0.5+col*3.15, y=2.2+row*1.0;
    scr(sl, x, y, 2.9, 0.8, C.card, 0.08);
    scr(sl, x+0.1, y+0.08, 0.8, 0.22, C.ent, 0.03);
    sl.addText(f[0], {x:x+0.1, y:y+0.08, w:0.8, h:0.22, fontSize:7, fontFace:F, color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(f[1], {x:x+0.1, y:y+0.38, w:2.7, h:0.35, fontSize:7.5, fontFace:F, color:C.text});
    sco(sl, x+2.5, y+0.5, 0.04, C.ent);
  });
  badge(sl, p, C.ent);
}

// ====== S15: ENT - LOGISTICS ======
function s15() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.ent);
  titleDeco(sl, "\u5165\u9a7b\u4f01\u4e1a \u2014 \u4f9b\u9700\u5bf9\u63a5\u4e0e\u7269\u6d41\u670d\u52a1", "", C.ent);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 50 } });
  sl.addText(C.ent+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.ent, align: "right" });
  // Left: modules
  var sds = [
    ["\u7269\u6d41\u4f9b\u9700", "\u56ed\u533a\u8d27\u6e90+\u8fd4\u7a0b\u7a7a\u8f66+\u8f66\u8f86\u8d44\u6e90\u667a\u80fd\u5339\u914d\uff0c\u6bcf\u5f20\u5361\u7247\u5e26\u8054\u7cfb\u65b9\u5f0f\uff0c\u652f\u6301\u591a\u6761\u4ef6\u7b5b\u9009", C.ent],
    ["\u7279\u8272\u4f9b\u9700", "\u9752\u5ddd\u5c71\u73cd\u7b499\u79cd\u7279\u8272\u4ea7\u54c1\u5c55\u793a\u4e0e\u4f9b\u9700\u5bf9\u63a5\uff0c\u76f4\u63a5\u8054\u7cfb\u4f9b\u5e94\u5546", "B45309"],
    ["\u4f9b\u5e94\u5546\u7ba1\u7406", "\u4f9b\u5e94\u5546/\u751f\u4ea7\u5546/\u8fd0\u8f93\u5546\u5b8c\u5584\u7ba1\u7406\u4f53\u7cfb\uff0c\u5408\u4f5c\u5173\u7cfb\u7ef4\u62a4", "92400E"]
  ];
  sds.forEach(function(s, i){
    var y = 0.85 + i * 0.8;
    scr(sl, 0.5, y, 4.5, 0.65, C.card, 0.08);
    sco(sl, 0.6, y+0.1, 0.06, s[2]);
    sl.addText(s[0], {x:0.78, y:y+0.05, w:3.6, h:0.22, fontSize:10, fontFace:F, color:s[2], bold:true});
    sl.addText(s[1], {x:0.78, y:y+0.28, w:3.6, h:0.32, fontSize:7.5, fontFace:F, color:C.text});
    if (i < sds.length-1) srec(sl, 0.75, y+0.68, 0.02, 0.15, C.line);
  });
  // Right: services
  scr(sl, 5.3, 0.85, 4.3, 2.55, C.card, 0.1);
  sl.addText("\u5546\u4e1a\u670d\u52a1\u6a21\u5757", {x:5.5, y:0.92, w:3.8, h:0.22, fontSize:11, fontFace:F, color:C.text, bold:true});
  srec(sl, 5.3, 1.12, 4.3, 0.01, C.ent);
  var bms = [
    ["\u6570\u5b57\u5316\u6d3b\u52a8", "\u62db\u6807/\u7ade\u4ef7/\u5ba1\u6279\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316"],
    ["\u5ba2\u6237\u7ba1\u7406", "\u5ba2\u6237\u5173\u7cfb\u7ef4\u62a4\u3001\u5546\u673a\u8ffd\u8e2a\u3001\u5408\u540c\u6267\u884c"],
    ["\u7535\u5b50\u540d\u7247", "\u4f01\u4e1a\u5458\u5de5\u7535\u5b50\u540d\u7247\u5feb\u901f\u5206\u4eab\uff0c\u4e00\u952e\u4fdd\u5b58"],
    ["\u5fae\u4fe1\u63a8\u5e7f", "\u56ed\u533a\u5fae\u4fe1\u5c0f\u7a0b\u5e8f\u5feb\u901f\u63a8\u5e7f\u4e0e\u4f7f\u7528"]
  ];
  bms.forEach(function(b, i){
    var y = 1.25 + i * 0.55;
    sco(sl, 5.5, y+0.04, 0.05, C.ent);
    sl.addText(b[0], {x:5.7, y:y, w:1.4, h:0.22, fontSize:8.5, fontFace:F, color:C.ent, bold:true});
    sl.addText(b[1], {x:7.1, y:y, w:2.4, h:0.22, fontSize:7.5, fontFace:F, color:C.gray});
    srec(sl, 5.5, y+0.28, 4.0, 0.005, C.line);
  });
  // Bottom: parking & policy
  scr(sl, 0.5, 3.65, 4.5, 0.7, C.card, 0.08);
  sl.addText("\u505c\u8f66\u670d\u52a1", {x:0.7, y:3.7, w:3, h:0.22, fontSize:10, fontFace:F, color:C.ent, bold:true});
  sl.addText("\u9ad8\u5fb7\u5730\u56fe\u5bfc\u822a+\u5b9e\u65f6\u4f59\u4f4d+\u5728\u7ebf\u7f34\u8d39\uff0c\u5168\u56ed\u533a\u505c\u8f66\u4f4d\u7edf\u4e00\u7ba1\u7406", {x:0.7, y:3.92, w:4.1, h:0.3, fontSize:7.5, fontFace:F, color:C.text});
  scr(sl, 5.3, 3.65, 4.3, 0.7, C.card, 0.08);
  sl.addText("\u653f\u7b56\u63a8\u9001", {x:5.5, y:3.7, w:3, h:0.22, fontSize:10, fontFace:F, color:C.ent, bold:true});
  sl.addText("\u60e0\u4f01\u653f\u7b56\u7cbe\u51c6\u5339\u914d\uff0c\u4e3b\u52a8\u63a8\u9001\u8865\u8d34/\u51cf\u514d/\u4eba\u624d\u7b49\u653f\u7b56", {x:5.5, y:3.92, w:4.1, h:0.3, fontSize:7.5, fontFace:F, color:C.text});
  badge(sl, p, C.ent);
}

// ====== S16: ENT - VENUE ======
function s16() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.ent);
  titleDeco(sl, "\u5165\u9a7b\u4f01\u4e1a \u2014 \u573a\u9986\u3001\u6d3b\u52a8\u4e0e\u9910\u996e\u670d\u52a1", "", C.ent);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 50 } });
  sl.addText(C.ent+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.ent, align: "right" });
  // Venue
  scr(sl, 0.5, 0.8, 5.2, 1.8, C.card, 0.1);
  srec(sl, 0.5, 0.8, 5.2, 0.35, C.ent);
  sl.addText("\u573a\u9986\u670d\u52a1 (8\u5927\u573a\u9986)", {x:0.7, y:0.83, w:4, h:0.28, fontSize:12, fontFace:F, color:"FFFFFF", bold:true});
  var vs = ["\u591a\u529f\u80fd\u4f1a\u8bae\u5385","\u5c55\u89c8\u5385","\u4f53\u80b2\u9986","\u57f9\u8bad\u6559\u5ba4","\u4f1a\u8bae\u5ba4","\u63a5\u5f85\u5385","\u591a\u5a92\u4f53\u5385","\u56ed\u533a\u5e7f\u573a"];
  vs.forEach(function(v,i){
    var x=0.6+(i%4)*1.25, y=1.25+Math.floor(i/4)*0.48;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.1, h: 0.22, fill: { color: C.entL }, rectRadius: 0.03 });
    sl.addText(v, {x, y, w:1.1, h:0.22, fontSize:7, fontFace:F, color:C.ent, align:"center", valign:"middle"});
  });
  // Activities
  scr(sl, 5.9, 0.8, 3.7, 1.8, C.card, 0.1);
  sl.addText("\u56ed\u533a\u6d3b\u52a8", {x:6.1, y:0.88, w:3, h:0.25, fontSize:11, fontFace:F, color:C.ent, bold:true});
  var acts = ["\u6587\u4f53\u6d3b\u52a8","\u57f9\u8bad\u8bb2\u5ea7","\u516c\u76ca\u6d3b\u52a8","\u4ea4\u6d41\u805a\u4f1a","\u8282\u5e86\u6d3b\u52a8","\u5c55\u89c8\u5c55\u793a"];
  acts.forEach(function(a,i){
    var y=1.2+i*0.22;
    sco(sl, 6.2, y+0.05, 0.045, C.ent);
    sl.addText(a, {x:6.35, y:y, w:3, h:0.2, fontSize:7.5, fontFace:F, color:C.text});
  });
  // Bottom: catering + employment
  scr(sl, 0.5, 2.85, 4.5, 0.9, C.card, 0.08);
  sl.addText("\u56ed\u533a\u9910\u996e", {x:0.7, y:2.9, w:3, h:0.22, fontSize:10, fontFace:F, color:C.ent, bold:true});
  sl.addText("6\u5bb6\u9910\u5385\uff0c20+\u9053\u83dc\u54c1\u4e00\u89c8\uff0c\u652f\u6301\u5206\u7c7b\u7b5b\u9009\u3001\u4eca\u65e5\u83dc\u5355\u67e5\u770b", {x:0.7, y:3.15, w:4.1, h:0.3, fontSize:8, fontFace:F, color:C.text});
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y: 3.48, w: 1.0, h: 0.18, fill: { color: C.entL }, rectRadius: 0.02 });
  sl.addText("\u83dc\u54c1\u4e00\u89c8", {x:0.7, y:3.48, w:1.0, h:0.18, fontSize:6.5, fontFace:F, color:C.ent, align:"center", valign:"middle"});
  scr(sl, 5.3, 2.85, 4.3, 0.9, C.card, 0.08);
  sl.addText("\u7528\u5de5\u4fe1\u606f", {x:5.5, y:2.9, w:3, h:0.22, fontSize:10, fontFace:F, color:C.ent, bold:true});
  sl.addText("\u62db\u8058+\u5c97\u4f4d\u5339\u914d\uff0c\u652f\u6301\u5206\u7c7b\u7b5b\u9009\uff0c\u5e2e\u52a9\u4f01\u4e1a\u5feb\u901f\u89e3\u51b3\u7528\u5de5", {x:5.5, y:3.15, w:4.1, h:0.3, fontSize:8, fontFace:F, color:C.text});
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.5, y: 3.48, w: 1.0, h: 0.18, fill: { color: C.entL }, rectRadius: 0.02 });
  sl.addText("\u62db\u8058\u4fe1\u606f", {x:5.5, y:3.48, w:1.0, h:0.18, fontSize:6.5, fontFace:F, color:C.ent, align:"center", valign:"middle"});
  // Visitor & payment shortcut
  scr(sl, 0.5, 4.0, 9, 0.7, C.card, 0.08);
  srec(sl, 0.5, 4.0, 9, 0.04, C.ent);
  sl.addText("\u5176\u4ed6\u670d\u52a1\uff1a", {x:0.7, y:4.1, w:1.5, h:0.22, fontSize:9, fontFace:F, color:C.ent, bold:true});
  var ots = ["\u5929\u7136\u6c14\u7f34\u8d39","\u7535\u8d39\u7f34\u8d39","\u81ea\u6765\u6c34\u7f34\u8d39","\u8bbf\u5ba2\u5f55\u5165","\u95e8\u7981\u7ba1\u7406","\u5ba1\u6279\u6d41\u7a0b","\u65e5\u62a5\u586b\u5199","\u573a\u9986\u9884\u7ea6"];
  ots.forEach(function(o,i){
    var x=2.3+i*0.9, y=4.1;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.85, h: 0.18, fill: { color: C.entL }, rectRadius: 0.02 });
    sl.addText(o, {x, y, w:0.85, h:0.18, fontSize:5.5, fontFace:F, color:C.ent, align:"center", valign:"middle"});
  });
  badge(sl, p, C.ent);
}

// ====== S17: ENT - EXHIBIT ======
function s17() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.ent);
  titleDeco(sl, "\u5165\u9a7b\u4f01\u4e1a \u2014 \u5c55\u793a\u4e0e\u62db\u5546\u5f15\u8d44", "", C.ent);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.ent, transparency: 50 } });
  sl.addText(C.ent+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.ent, align: "right" });
  // 3 top cards
  var exs = [
    ["\u4e91\u5c55\u9986", "\u4f01\u4e1a\u5f62\u8c61\u7ebf\u4e0a\u5c55\u793a\uff0c\u4ea7\u54c1\u7acb\u4f53\u5c55\u793a\uff0c\u652f\u6301\u591a\u5a92\u4f53\u5185\u5bb9", C.ent],
    ["\u7535\u5b50\u540d\u7247", "\u4f01\u4e1a\u5458\u5de5\u7535\u5b50\u540d\u7247\u5feb\u901f\u5206\u4eab\uff0c\u4e00\u952e\u4fdd\u5b58\u8054\u7cfb\u4eba", "B45309"],
    ["\u56ed\u533a\u5ba3\u4f20", "\u56ed\u533a\u6982\u51b5\u3001\u53d1\u5c55\u5386\u7a0b\u3001\u4f18\u52bf\u4ea7\u4e1a\u591a\u7ef4\u5ea6\u5c55\u793a", "92400E"]
  ];
  exs.forEach(function(e, i){
    var x=0.5+i*3.15;
    scr(sl, x, 0.8, 2.9, 1.3, C.card, 0.1);
    sl.addShape(pres.shapes.PARALLELOGRAM, { x: x+0.1, y: 0.9, w: 0.7, h: 0.28, fill: { color: e[2] } });
    sl.addText(e[0], {x:x+0.1, y:0.9, w:0.7, h:0.28, fontSize:8, fontFace:F, color:"FFFFFF", bold:true, align:"center", valign:"middle"});
    sl.addText(e[1], {x:x+0.15, y:1.25, w:2.6, h:0.65, fontSize:8.5, fontFace:F, color:C.text});
    srec(sl, x+0.15, 1.95, 0.8, 0.03, e[2]);
  });
  // Target enterprise & investment
  scr(sl, 0.5, 2.3, 4.5, 1.1, C.card, 0.08);
  sl.addText("\u76ee\u6807\u4f01\u4e1a\u6316\u6398", {x:0.7, y:2.35, w:3.5, h:0.22, fontSize:10, fontFace:F, color:C.ent, bold:true});
  srec(sl, 0.5, 2.55, 4.5, 0.01, C.line);
  sl.addText("\u591a\u7ef4\u6570\u636e\u6e90+\u667a\u80fd\u7b97\u6cd5\u81ea\u52a8\u63a8\u8350\u4e0e\u56ed\u533a\u4ea7\u4e1a\u9ad8\u5ea6\u5339\u914d\u7684\u76ee\u6807\u4f01\u4e1a\uff0c\u5b9e\u73b0\u7cbe\u51c6\u62db\u5546\u5f15\u8d44\u3002\u57fa\u4e8e\u4ea7\u4e1a\u94fe\u3001\u8d44\u91d1\u3001\u6280\u672f\u3001\u5730\u7406\u7b49\u591a\u7ef4\u5ea6\u7b97\u6cd5\u8bc4\u5206\u3002",
    {x:0.7, y:2.62, w:4.1, h:0.6, fontSize:7.5, fontFace:F, color:C.text});
  scr(sl, 5.3, 2.3, 4.3, 1.1, C.card, 0.08);
  sl.addText("\u6295\u8d44\u9879\u76ee\u7ba1\u7406", {x:5.5, y:2.35, w:3.8, h:0.22, fontSize:10, fontFace:F, color:C.ent, bold:true});
  srec(sl, 5.3, 2.55, 4.3, 0.01, C.line);
  sl.addText("\u62db\u5546\u9879\u76ee\u5168\u751f\u547d\u5468\u671f\u7ba1\u7406\uff0c\u4ece\u63a5\u89e6\u5230\u7b7e\u7ea6\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316\u3002\u5305\u62ec\u8f6c\u5316\u7387\u3001\u5408\u7ea6\u91d1\u989d\u3001\u4ea7\u4e1a\u805a\u96c6\u5ea6\u7b49\u5173\u952e\u6307\u6807\u3002",
    {x:5.5, y:2.62, w:4.0, h:0.6, fontSize:7.5, fontFace:F, color:C.text});
  // Enterprise list
  scr(sl, 0.5, 3.65, 9, 0.65, C.card, 0.08);
  sl.addText("\u56ed\u533a\u91cd\u70b9\u4f01\u4e1a\uff089\u5bb6\uff09\uff1a", {x:0.7, y:3.7, w:3, h:0.22, fontSize:9, fontFace:F, color:C.ent, bold:true});
  var ents = ["\u961c\u6210\u65b0\u6750","\u8679\u660c\u6676\u4f53","\u4e2d\u80fd\u5efa","\u4f73\u5174\u94dd\u4e1a","\u661f\u6052\u9752\u6e90","\u4e5d\u65ed\u65b0\u6750","\u534e\u7eb3\u5b9e\u4e1a","\u65b0\u901a\u946b","\u5e7f\u8882\u65b0\u6750"];
  ents.forEach(function(en, i){
    var x=0.7+(i<5?i:i-5)*1.75, y=3.95+(i<5?0:0.28);
    sco(sl, x, y+0.03, 0.035, C.ent);
    sl.addText(en, {x:x+0.06, y:y, w:1.5, h:0.2, fontSize:6.5, fontFace:F, color:C.text});
  });
  badge(sl, p, C.ent);
}

// ====== S18: PUBLIC ======
function s18() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.pub);
  titleDeco(sl, "\u516c\u4f17\u7528\u6237 \u2014 \u4fe1\u606f\u670d\u52a1\u4e0e\u56ed\u533a\u53c2\u4e0e", "", C.pub);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.pub, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.pub, transparency: 50 } });
  sl.addText(C.pub+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.pub, align: "right" });
  // Info center
  scr(sl, 0.5, 0.85, 9, 1.2, C.card, 0.1);
  sl.addText("\u56ed\u533a\u4fe1\u606f\u4e2d\u5fc3", {x:0.7, y:0.9, w:4, h:0.25, fontSize:12, fontFace:F, color:C.pub, bold:true});
  var ih = [["\u56ed\u533a\u516c\u544a","\u5b9e\u65f6\u67e5\u770b\u56ed\u533a\u901a\u77e5"],["\u56ed\u533a\u6d3b\u52a8","\u67e5\u770b\u5e76\u53c2\u4e0e\u56ed\u533a\u6d3b\u52a8"],["\u56ed\u533a\u653f\u7b56","\u67e5\u770b\u516c\u5f00\u653f\u7b56\u4fe1\u606f"],["\u4fe1\u606f\u53d1\u5e03","\u56ed\u533a\u53d1\u5c55\u6210\u5c31\u591a\u7ef4\u5ea6\u5c55\u793a"]];
  ih.forEach(function(h, i){
    var x=0.7+i*2.2;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.2, w: 2.0, h: 0.35, fill: { color: C.pubL }, rectRadius: 0.05 });
    sco(sl, x+0.08, 1.27, 0.06, C.pub);
    sl.addText(h[0], {x:x+0.2, y:1.2, w:1.7, h:0.18, fontSize:9, fontFace:F, color:C.pub, bold:true});
    sl.addText(h[1], {x:x+0.05, y:1.38, w:1.9, h:0.18, fontSize:7, fontFace:F, color:C.text});
  });
  // Life services
  var lss = [
    ["\u56ed\u533a\u9910\u996e", "6\u5bb6\u9910\u5385\u00b7\u4eca\u65e5\u83dc\u5355\u00b7\u5206\u7c7b\u7b5b\u9009", C.pub],
    ["\u505c\u8f66\u670d\u52a1", "\u5730\u56fe\u67e5\u770b\u4f59\u4f4d\u00b7\u5728\u7ebf\u7f34\u8d39", "6D28D9"],
    ["\u573a\u9986\u670d\u52a1", "\u6d4f\u89c8\u573a\u9986\u00b7\u5728\u7ebf\u9884\u7ea6", "5B21B6"],
    ["\u7279\u8272\u4f9b\u9700", "\u9752\u5ddd\u5c71\u73cd\u4e0e\u7279\u8272\u4ea7\u54c1", C.pub],
    ["\u7528\u5de5\u4fe1\u606f", "\u62db\u8058\u4fe1\u606f\u00b7\u5c97\u4f4d\u5339\u914d", "6D28D9"]
  ];
  lss.forEach(function(l, i){
    var x=0.5+(i%3)*3.15, y=2.3+Math.floor(i/3)*1.1;
    scr(sl, x, y, 2.9, 0.9, C.card, 0.1);
    srec(sl, x+0.08, y+0.08, 0.03, 0.74, l[2]);
    sl.addText(l[0], {x:x+0.2, y:y+0.1, w:2.5, h:0.22, fontSize:10, fontFace:F, color:l[2], bold:true});
    sl.addText(l[1], {x:x+0.2, y:y+0.38, w:2.5, h:0.35, fontSize:8, fontFace:F, color:C.text});
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x+0.2, y: y+0.65, w: 0.7, h: 0.16, fill: { color: C.pubL }, rectRadius: 0.02 });
    sl.addText("\u67e5\u770b >", {x:x+0.2, y:y+0.65, w:0.7, h:0.16, fontSize:6, fontFace:F, color:C.pub, align:"center", valign:"middle"});
  });
  badge(sl, p, C.pub);
}

// ====== S19: PUBLIC - ENGAGEMENT ======
function s19() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.pub);
  titleDeco(sl, "\u516c\u4f17\u7528\u6237 \u2014 \u4e92\u52a8\u4f53\u9a8c\u4e0e\u670d\u52a1\u6570\u636e", "", C.pub);
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.0, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.pub, transparency: 30 } });
  sl.addShape(pres.shapes.PARALLELOGRAM, { x: 8.7, y: 0.1, w: 0.6, h: 0.3, fill: { color: C.pub, transparency: 50 } });
  sl.addText(C.pub+"\u2014", { x: 7.5, y: 0.55, w: 2.0, h: 0.2, fontSize: 8, fontFace: F, color: C.pub, align: "right" });
  // Interactive modules grid 2x3
  scr(sl, 0.5, 0.85, 5.3, 4.0, C.card, 0.1);
  sl.addText("\u516c\u4f17\u4e92\u52a8\u4e0e\u53c2\u4e0e", {x:0.7, y:0.92, w:4.5, h:0.25, fontSize:12, fontFace:F, color:C.pub, bold:true});
  srec(sl, 0.5, 1.14, 5.3, 0.01, C.pub);
  var ints = [
    ["\u56ed\u533a\u5e7f\u573a\u5c55\u793a", "\u5927\u5c4f\u5ba3\u4f20\u4e0e\u4fe1\u606f\u53d1\u5e03"],
    ["\u5fae\u4fe1\u5c0f\u7a0b\u5e8f", "\u5feb\u901f\u63a5\u5165\u667a\u6167\u5e73\u53f0"],
    ["\u4f01\u4e1a\u7f51\u7ad9\u5bfc\u822a", "\u56ed\u533a\u4f01\u4e1a\u7f51\u7ad9\u96c6\u6210"],
    ["\u54a8\u8be2\u6295\u8bc9", "\u5728\u7ebf\u54a8\u8be2\u3001\u6295\u8bc9\u53cd\u9988"],
    ["\u6570\u5b57\u5316\u5ba2\u670d", "\u667a\u80fd\u5ba2\u670d+24h\u4eba\u5de5\u5ba2\u670d"],
    ["\u793e\u533a\u4e92\u52a8", "\u8bba\u575b\u3001\u610f\u89c1\u5f81\u96c6\u3001\u6ee1\u610f\u5ea6"]
  ];
  ints.forEach(function(it, i){
    var col=i%2, row=Math.floor(i/2);
    var x=0.6+col*2.4, y=1.3+row*0.85;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.2, h: 0.65, fill: { color: C.pubL }, rectRadius: 0.06 });
    sco(sl, x+0.08, y+0.08, 0.05, C.pub);
    sl.addText(it[0], {x:x+0.18, y:y+0.05, w:1.9, h:0.22, fontSize:9, fontFace:F, color:C.pub, bold:true});
    sl.addText(it[1], {x:x+0.08, y:y+0.3, w:2.0, h:0.28, fontSize:7.5, fontFace:F, color:C.text});
  });
  // Right: stats
  scr(sl, 6.0, 0.85, 3.6, 4.0, C.card, 0.1);
  srec(sl, 6.0, 0.85, 3.6, 0.06, C.pub);
  sl.addText("\u670d\u52a1\u6570\u636e\u4e00\u89c8", {x:6.15, y:0.95, w:3.2, h:0.22, fontSize:10, fontFace:F, color:"FFFFFF", bold:true, align:"center"});
  var us = [["\u65e5\u5747\u6d3b\u52a8\u53c2\u4e0e","120+\u4eba\u6b21"],["\u65e5\u5747\u516c\u544a\u6d4f\u89c8","500+\u6b21"],["\u573a\u9986\u65e5\u5747\u9884\u7ea6","80+\u6b21"],["\u65e5\u5747\u7f34\u8d39\u4ea4\u6613","60+\u7b14"],["\u7528\u5de5\u4fe1\u606f\u6d4f\u89c8","200+\u6b21"],["\u65e5\u5747\u54a8\u8be2\u91cf","40+\u6761"],["\u6ee1\u610f\u5ea6\u8bc4\u5206","4.8\u2605"]];
  us.forEach(function(u,i){
    var y=1.1+i*0.38;
    sco(sl, 6.15, y+0.05, 0.04, C.pub);
    sl.addText(u[0],{x:6.3,y,w:1.5,h:0.2,fontSize:7.5,fontFace:F,color:C.gray});
    sl.addText(u[1],{x:7.7,y:y-0.02,w:1.2,h:0.22,fontSize:12,fontFace:"Arial",color:C.acc,bold:true,align:"right"});
    if (i<6) srec(sl, 6.15, y+0.25, 3.2, 0.005, C.line);
  });
  badge(sl, p, C.pub);
}

// ====== S20: FULL MATRIX (compressed) ======
function s20() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  sl.addText("\u5e73\u53f0\u5168\u529f\u80fd\u77e9\u9635\u4e00\u89c8", { x: 0.5, y: 0.1, w: 8, h: 0.4, fontSize: 18, fontFace: F, color: C.text, bold: true });
  sl.addText("\u5b8c\u6574\u529f\u80fd\u6a21\u5757\u4e0e\u56db\u7c7b\u89d2\u8272\u8986\u76d6\u5173\u7cfb\u4e00\u89c8\u8868", { x: 0.5, y: 0.46, w: 8, h: 0.18, fontSize: 7.5, fontFace: F, color: C.gray });
  var cols = ["\u529f\u80fd\u6a21\u5757","\u653f\u5e9c","\u7ba1\u7406","\u4f01\u4e1a","\u516c\u4f17","\u70b9"],
      cw = [1.5, 1.65, 1.65, 1.65, 1.65, 0.5],
      hc = [C.text, C.gov, C.mgmt, C.ent, C.pub, C.gray];
  cols.forEach(function(c,i){
    var x = (i===0) ? 0.3 : 0.3 + cw.slice(0,i).reduce(function(a,b){return a+b;}, 0);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 0.7, w: cw[i], h: 0.2, fill: { color: hc[i] }, rectRadius: 0.02 });
    sl.addText(c, {x, y:0.7, w:cw[i], h:0.2, fontSize:6.5, fontFace:F, color:"FFFFFF", bold:true, align:"center", valign:"middle"});
  });
  var rows = [
    ["AI\u62db\u5546\u670d\u52a1",1,0,0,0,5],["\u6570\u5b57\u6c99\u76d8",1,0,0,0,3],
    ["\u4ea7\u4e1a\u94fe\u56fe\u8c31",1,0,0,0,4],["\u76ee\u6807\u4f01\u4e1a\u6316\u6398",1,0,0,0,2],
    ["\u6295\u8d44\u9879\u76ee\u7ba1\u7406",1,0,1,0,4],["\u62db\u5546\u5165\u53e3/\u4e91\u5c55\u9986",0,0,1,0,5],
    ["\u5b89\u5168\u76d1\u63a7(\u5929\u773c)",1,1,0,1,5],["\u5730\u707e\u76d1\u6d4b(\u54e8\u5175)",1,1,0,0,4],
    ["\u73af\u4fdd\u76d1\u6d4b(\u6167\u773c)",1,1,0,0,4],["\u89c6\u9891\u76d1\u63a7\u603b\u89c8",1,1,0,1,3],
    ["\u5e94\u6025\u6307\u6325\u8054\u52a8",1,1,0,1,6],["\u8bbe\u5907\u7ba1\u7406",0,1,0,0,2],
    ["\u8bc9\u6c42\u76f4\u8fbe",1,1,1,0,5],["\u670d\u52a1\u76f4\u8fbe",0,0,1,0,4],
    ["\u7f34\u8d39(\u6c14/\u6c34/\u7535)",0,0,1,0,3],["\u56ed\u4f01\u4ea4\u4e92",1,1,1,0,5],
    ["\u4f9b\u9700\u5bf9\u63a5",1,1,1,0,5],["\u653f\u7b56\u63a8\u9001",1,1,1,0,5],
    ["\u573a\u9986/\u6d3b\u52a8/\u9910\u996e",0,1,1,1,6],["\u505c\u8f66\u670d\u52a1",0,1,1,1,4],
    ["\u7528\u5de5\u4fe1\u606f",1,1,1,1,5],["\u80fd\u6548\u76d1\u6d4b\u5206\u6790",1,1,0,0,4],
    ["\u4ea9\u5747\u8bba\u82f1\u96c4",1,1,0,0,3],["\u6570\u636e\u4e2d\u5fc3",1,1,0,0,3],
    ["\u6570\u5b57\u5316\u529e\u516c",0,1,1,0,6],["\u8bbf\u5ba2/\u95e8\u7981\u7ba1\u7406",0,1,1,0,5],
    ["\u56ed\u533a\u5ba3\u4f20/\u516c\u544a",1,1,1,1,6],["\u6570\u5b57\u5316\u6d3b\u52a8",0,1,1,0,4]
  ];
  rows.forEach(function(row, ri){
    var y = 0.93 + ri * 0.143;
    var bg = ri % 2 === 0 ? C.card : C.bg;
    row.forEach(function(cell, ci){
      var x = (ci===0) ? 0.3 : 0.3 + cw.slice(0,ci).reduce(function(a,b){return a+b;}, 0);
      sl.addShape(pres.shapes.RECTANGLE, { x, y, w: cw[ci], h: 0.143, fill: { color: bg } });
      if (ci === 0) {
        sl.addText(cell, {x: x+0.03, y, w: cw[ci]-0.06, h: 0.143, fontSize:5.5, fontFace:F, color:C.text, valign:"middle"});
      } else if (ci < 5) {
        var m = cell === 1 ? "\u25cf" : "\u2014";
        var mc = cell === 1 ? hc[ci] : C.line;
        sl.addText(m, {x, y, w: cw[ci], h: 0.143, fontSize:8, fontFace:"Arial", color:mc, align:"center", valign:"middle"});
      } else {
        sl.addText(String(cell), {x, y, w: cw[ci], h: 0.143, fontSize:5.5, fontFace:F, color:C.gray, align:"center", valign:"middle"});
      }
    });
  });
  var ly = 0.93 + rows.length * 0.143 + 0.05;
  scr(sl, 0.5, ly, 9, 0.16, "0F172A", 0.04);
  sl.addText("\u25cf \u4e3b\u8981\u529f\u80fd  \u2014 \u65e0\u76f4\u63a5\u5173\u8054   |   \u603b\u8ba128\u4e2a\u529f\u80fd\u6a21\u5757\u8986\u76d6\u56db\u5927\u89d2\u8272",
    { x: 0.7, y: ly, w: 8.6, h: 0.16, fontSize: 7, fontFace: F, color: "94A3B8", valign: "middle" });
  badge(sl, p, C.acc);
}

// ====== S21: INVESTMENT ======
function s21() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  titleDeco(sl, "\u6295\u8d44\u4e0e\u6548\u76ca\u5206\u6790", "", C.acc);
  // Total investment
  scr(sl, 0.5, 0.75, 9, 1.0, C.card, 0.12);
  srec(sl, 0.5, 0.75, 9, 0.06, C.acc);
  sl.addText("\u667a\u6167\u5e73\u53f0\u5efa\u8bbe\u603b\u6295\u8d44\u6982\u7b97", {x:0.7, y:0.85, w:5, h:0.2, fontSize:11, fontFace:F, color:C.acc, bold:true});
  sl.addText("\u00a54,178,800.00", {x:0.7, y:1.05, w:4, h:0.55, fontSize:30, fontFace:"Arial", color:C.text, bold:true});
  sl.addText("\u5143", {x:4.8, y:1.2, w:1, h:0.4, fontSize:12, fontFace:F, color:C.gray, valign:"bottom"});
  // Breakdown
  var items = [
    ["\u9053\u8def\u4ea4\u901a\u5b89\u5168\u4e0e\u5730\u707e\u76d1\u63a7","1,580,000",38],
    ["\u91cd\u70b9\u53e3\u5b50\u73af\u4fdd\u76d1\u63a7","428,600",10],
    ["\u56ed\u533a\u5f62\u8c61\u5c55\u793a\u4e0eAI\u62db\u5546","442,200",11],
    ["\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b","348,000",8],
    ["\u667a\u6539\u6570\u8f6c\u8d4b\u80fd\u5e73\u53f0","280,000",7],
    ["\u4e00\u4f53\u5316\u6570\u5b57\u5e73\u53f0","720,000",17],
    ["\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3","380,000",9]
  ];
  items.forEach(function(it, i){
    var y = 1.95 + i * 0.38;
    sco(sl, 0.6, y+0.08, 0.05, C.acc);
    sl.addText(it[0], {x:0.75, y:y, w:3.5, h:0.25, fontSize:9, fontFace:F, color:C.text, valign:"middle"});
    sl.addText("\u00a5"+it[1], {x:4.3, y:y, w:1.3, h:0.25, fontSize:8, fontFace:"Arial", color:C.ent, bold:true, align:"right", valign:"middle"});
    // Progress bar
    scr(sl, 5.7, y+0.06, 2.5, 0.12, C.line, 0.06);
    scr(sl, 5.7, y+0.06, 2.5*it[2]/100, 0.12, C.acc, 0.06);
    sl.addText(String(it[2]) + "%", {x:8.3, y:y, w:0.7, h:0.25, fontSize:7, fontFace:"Arial", color:C.gray, align:"right", valign:"middle"});
  });
  // Benefits
  srec(sl, 0.5, 4.7, 9, 0.02, C.line);
  var bs = [["\u7ba1\u7406\u6548\u7387\u63d0\u534730%",C.gov],["\u5e94\u6025\u54cd\u5e94\u63d0\u534760%",C.mgmt],["\u00a520\u4ebf+\u5e74\u5bf9\u63a5\u8d44\u91d1",C.ent],["\u76d1\u63a7\u8986\u76d6\u738795%",C.pub]];
  bs.forEach(function(b,i){
    var x=0.5+i*2.35;
    sl.addText(b[0],{x,y:4.85,w:2.2,h:0.25,fontSize:11,fontFace:F,color:b[1],bold:true,align:"center"});
  });
  badge(sl, p, C.acc);
}

// ====== S22: FUTURE ======
function s22() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.gov);
  titleDeco(sl, "\u672a\u6765\u5c55\u671b", "\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", C.gov);
  var fs = [
    ["AIoT\u6df1\u5ea6\u878d\u5408", "\u5f15\u5165\u7269\u8054\u7f51\u4f20\u611f\u8bbe\u5907\uff0c\u5b9e\u73b0\u4e07\u7269\u4e92\u8054\u3001\u5168\u57df\u611f\u77e5\u4e0e\u667a\u80fd\u63a7\u5236", C.gov],
    ["\u6570\u5b57\u5b6a\u751f\u56ed\u533a", "1:1\u4e09\u7ef4\u5efa\u6a21\uff0c\u5168\u8981\u7d20\u3001\u5168\u65f6\u57df\u53ef\u89c6\u5316\u8fd0\u8425\u7ba1\u7406", C.mgmt],
    ["AI Agent\u667a\u80fd\u5316", "\u667a\u80fd\u5ba2\u670d\u3001\u667a\u80fd\u8c03\u5ea6\u3001\u667a\u80fd\u9884\u8b66\uff0cAI\u9a71\u52a8\u56ed\u533a\u81ea\u52a8\u5316\u8fd0\u8425", C.ent],
    ["\u4ea7\u4e1a\u5927\u8111", "\u57fa\u4e8e\u5927\u6570\u636e\u4e0eAI\u7b97\u6cd5\u8f85\u52a9\u4ea7\u4e1a\u62db\u5546\u51b3\u7b56", C.pub],
    ["\u8de8\u56ed\u533a\u534f\u540c", "\u5bf9\u63a5\u7701\u5e02\u667a\u6167\u5e73\u53f0\uff0c\u5b9e\u73b0\u6570\u636e\u5171\u4eab\u4e0e\u4e1a\u52a1\u534f\u540c", C.acc]
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
  srec(sl, 0.5, 4.7, 9, 0.015, C.line);
  sl.addText("\u9a71\u52a8\u56ed\u533a\u6cbf\u7740\u201c\u7701\u7ea7\u5f00\u53d1\u533a\u2192\u7eff\u8272\u56ed\u533a\u2192\u56fd\u5bb6\u7ea7\u7ecf\u5f00\u533a\u57f9\u80b2\u5bf9\u8c61\u201d\u7684\u822a\u5411\u5168\u901f\u524d\u8fdb",
    { x: 0.5, y: 4.85, w: 9, h: 0.25, fontSize: 9, fontFace: F, color: C.gray, italic: true, align: "center" });
  badge(sl, p, C.gov);
}

// ====== S23: THANK YOU ======
function s23() {
  p++; var sl = pres.addSlide(); sl.background = { color: C.bg };
  tb(sl, C.acc);
  sco(sl, 3, 0.5, 4, C.govL);
  sco(sl, -1, 2, 3, C.accL);
  sl.addText("\u611f\u8c22\u804a\u542c", { x: 1, y: 1.2, w: 8, h: 0.8, fontSize: 34, fontFace: F, color: C.text, bold: true, align: "center" });
  srec(sl, 4.2, 2.0, 1.6, 0.03, C.acc);
  sl.addText("\u4ee5\u6570\u5b57\u4e3a\u5e06\uff0c\u4ee5\u667a\u6167\u4e3a\u6868", { x: 1, y: 2.15, w: 8, h: 0.4, fontSize: 14, fontFace: F, color: C.acc, align: "center" });
  var fts = ["\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u7ba1\u59d4\u4f1a", "\u6280\u672f\u652f\u6301\uff1a\u9752\u8fc5\u79d1\u6280", "\u5730\u5740\uff1a\u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf\u7af9\u56ed\u9547"];
  fts.forEach(function(f, i){
    sl.addText(f, { x: 1, y: 3.4 + i*0.35, w: 8, h: 0.3, fontSize: 10, fontFace: F, color: C.gray, align: "center" });
  });
}

// ====== BUILD ======
s1(); s2(); s3(); s4(); s5(); s6();
s7(); s8(); s9(); s10();
s11(); s12(); s13();
s14(); s15(); s16(); s17();
s18(); s19();
s20();
s21();
s22();
s23();

pres.writeFile({ fileName: path.join(__dirname, "output", "zhuangzishang_product_ppt.pptx") })
  .then(function() { console.log("PPT OK: " + p + " slides"); })
  .catch(function(e) { console.error("ERR:", e); });
