
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
const IMG = "/Users/wangcaihua/WorkBuddy/20260427142154/slides/imgs";

const T = {
  primary: "1B4332", secondary: "2D6A4F", accent: "40916C",
  blue: "2D8CF0", sky: "E8F4FD", gold: "F59E0B",
  lightBg: "F8FAFC", cardBg: "FFFFFF",
  gray: "64748B", dark: "1E293B"
};

let pn = 0;
function badge(sl, n, c) {
  sl.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: c || T.accent } });
  sl.addText(String(n), { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 11, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
}

function dot(sl, x, y, size, color) {
  sl.addShape(pres.shapes.OVAL, { x, y, w: size, h: size, fill: { color: color || T.accent } });
}

function sectionTitle(sl, title, sub) {
  sl.addText(title, { x: 0.6, y: 0.2, w: 8, h: 0.55, fontSize: 22, fontFace: "Microsoft YaHei", color: T.dark, bold: true });
  if (sub) sl.addText(sub, { x: 0.6, y: 0.75, w: 8, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  dot(sl, 0.6, 0.75, 0.06, T.accent);
}

function featureSlide(title, items, imgFile, accent, roleTag) {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  // Decorative top dots
  for (var i = 0; i < 5; i++) dot(sl, 0.3 + i * 0.2, 0.08, 0.04, T.accent, 50);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: accent || T.accent } });
  sectionTitle(sl, title);
  
  var yPos = 1.15;
  items.forEach(function(item, i) {
    dot(sl, 0.6, yPos + 0.08, 0.08, accent || T.accent);
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: yPos - 0.05, w: 4.8, h: 0.38, fill: { color: "FFFFFF" }, rectRadius: 0.04, shadow: { type: "outer", blur: 2, offset: 1, color: "000000", opacity: 0.04 } });
    sl.addText(item, { x: 1.0, y: yPos - 0.05, w: 4.5, h: 0.38, fontSize: 11, fontFace: "Microsoft YaHei", color: T.dark, valign: "middle" });
    yPos += 0.45;
  });
  
  if (imgFile && fs.existsSync(imgFile)) {
    sl.addImage({ path: imgFile, x: 5.6, y: 0.5, w: 4.2, h: 4.5 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: 0.5, w: 4.2, h: 4.5, line: { color: T.gray, width: 0.5 }, lineDash: "dash" });
  }
  if (roleTag) {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.0, y: 4.7, w: 2.6, h: 0.28, fill: { color: accent || T.accent }, rectRadius: 0.05 });
    sl.addText(roleTag, { x: 7.0, y: 4.7, w: 2.6, h: 0.28, fontSize: 8, fontFace: "Microsoft YaHei", color: "FFFFFF", align: "center", valign: "middle" });
  }
  badge(sl, pn, accent);
}

// === SLIDE 1: COVER ===
function s1() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.primary };
  // Decorative circles
  sl.addShape(pres.shapes.OVAL, { x: 7, y: -1.5, w: 5, h: 5, fill: { color: T.secondary, transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -1, y: 3, w: 3, h: 3, fill: { color: T.secondary, transparency: 60 } });
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.05, T.accent);
  
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.9, w: 0.08, h: 2.5, fill: { color: T.accent } });
  sl.addText("青川县庄子上工业园区", { x: 0.9, y: 0.9, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: T.gold });
  sl.addText("智慧园区平台汇报", { x: 0.9, y: 1.4, w: 8, h: 1.2, fontSize: 38, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addText("建设方案 · 功能展示 · 角色价值", { x: 0.9, y: 2.6, w: 6, h: 0.4, fontSize: 16, fontFace: "Microsoft YaHei", color: T.sky });
  
  var aud = ["政府领导层", "园区管理层", "入驻企业", "公众用户"];
  aud.forEach(function(a, i) {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.9 + i * 2.1, y: 3.6, w: 1.9, h: 0.35, fill: { color: T.secondary }, rectRadius: 0.08 });
    sl.addText(a, { x: 0.9 + i * 2.1, y: 3.6, w: 1.9, h: 0.35, fontSize: 9, fontFace: "Microsoft YaHei", color: "FFFFFF", align: "center", valign: "middle" });
  });
  
  // Location & date
  sl.addText("四川省广元市青川县竹园镇", { x: 0.6, y: 4.2, w: 5, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: T.gray });
  sl.addText("2026年5月", { x: 0.6, y: 4.5, w: 3, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  
  // AI image on the right side
  if (fs.existsSync(IMG + "/ai_smart_park_dashboard.png")) {
    sl.addImage({ path: IMG + "/ai_smart_park_dashboard.png", x: 5.5, y: 2.8, w: 4.2, h: 2.4, sizing: { type: "cover", w: 4.2, h: 2.4 } });
  }
}

// === SLIDE 2: TOC ===
function s2() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.accent);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.accent } });
  sectionTitle(sl, "汇报目录");
  var items = ["项目背景与战略定位", "三步走创建路径", "七大建设工程", "平台架构总览", "AI招商服务", "全域安全监控", "企业服务", "物流·特色·政策", "餐饮·停车·场馆", "活动·公告·用工", "数据中心与能效", "投资概算", "预期效益", "四类角色功能矩阵", "未来展望"];
  items.forEach(function(item, i) {
    var col = i % 2, row = Math.floor(i / 2);
    var x = 0.6 + col * 4.6, y = 1.1 + row * 0.48;
    dot(sl, x, y + 0.06, 0.12, T.accent);
    sl.addText(String(i + 1), { x: x - 0.02, y: y + 0.01, w: 0.16, h: 0.16, fontSize: 8, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(item, { x: x + 0.2, y, w: 4.2, h: 0.2, fontSize: 11, fontFace: "Microsoft YaHei", color: T.dark, valign: "middle" });
  });
  badge(sl, pn);
}

s1(); s2();

// === SLIDE 3: Background + Location ===
function s3() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.accent);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.accent } });
  sectionTitle(sl, "项目背景与战略定位", "青川县庄子上工业园区 · 四川省广元市");
  
  // Location info box
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.15, w: 9, h: 0.65, fill: { color: T.sky }, rectRadius: 0.06 });
  dot(sl, 0.65, 1.25, 0.1, T.blue);
  sl.addText("青川县位于四川省北部边缘，川甘陕三省结合部，素有\u201c金三角\u201d之称", { x: 0.9, y: 1.2, w: 8.4, h: 0.25, fontSize: 11, fontFace: "Microsoft YaHei", color: T.dark });
  sl.addText("庄子上工业园区规划面积2.1平方公里，已开发1.6平方公里", { x: 0.9, y: 1.45, w: 8.4, h: 0.25, fontSize: 11, fontFace: "Microsoft YaHei", color: T.dark });
  
  // Data cards
  var data = [["41.6亿", "规上工业总产值", "2025年"], ["47家", "入驻企业", "新增5家/年"], ["3,000+", "从业人员", "带动就业"], ["7大", "产业门类", "建材·新能源·新材料"]];
  data.forEach(function(d, i) {
    var x = 0.5 + i * 2.3, y = 2.1;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.1, h: 1.3, fill: { color: "FFFFFF" }, rectRadius: 0.08, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.04 } });
    dot(sl, x + 0.15, y + 0.15, 0.06, T.accent);
    sl.addText(d[0], { x: x + 0.1, y: y + 0.3, w: 1.9, h: 0.45, fontSize: 26, fontFace: "Arial", color: T.accent, bold: true, align: "center" });
    sl.addText(d[1], { x: x + 0.1, y: y + 0.7, w: 1.9, h: 0.25, fontSize: 11, fontFace: "Microsoft YaHei", color: T.dark, bold: true, align: "center" });
    sl.addText(d[2], { x: x + 0.1, y: y + 0.95, w: 1.9, h: 0.2, fontSize: 9, fontFace: "Microsoft YaHei", color: T.gray, align: "center" });
  });
  
  // Strategy
  sl.addText("战略定位：以智慧园区建设驱动庄子上工业园区能级跃升", { x: 0.5, y: 3.7, w: 9, h: 0.3, fontSize: 13, fontFace: "Microsoft YaHei", color: T.dark, bold: true });
  ["赋能创建：为申报省级经开区、绿色工业园区提供数据支撑", "赋能企业：打造\u2018智改数转\u2019公共服务平台，破解转型难题", "赋能决策：构建园区数字大脑，实现一屏统览、一图统管"].forEach(function(item, i) {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.1 + i * 0.35, w: 9, h: 0.3, fill: { color: "FFFFFF" }, rectRadius: 0.04 });
    dot(sl, 0.62, 4.14 + i * 0.35, 0.06, T.accent);
    sl.addText(item, { x: 0.8, y: 4.1 + i * 0.35, w: 8.5, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: T.dark, valign: "middle" });
  });
  
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.0, y: 4.7, w: 2.6, h: 0.28, fill: { color: T.accent }, rectRadius: 0.05 });
  sl.addText("面向 政府领导层", { x: 7.0, y: 4.7, w: 2.6, h: 0.28, fontSize: 8, fontFace: "Microsoft YaHei", color: "FFFFFF", align: "center", valign: "middle" });
  badge(sl, pn);
}

s3();

// === SLIDE 4: Three Steps ===
function s4() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.primary };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.gold);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.gold } });
  sl.addText("三步走创建路径", { x: 0.6, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addText("智慧园区项目贯穿始终的核心支撑", { x: 0.6, y: 0.8, w: 8, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  
  var steps = [["STEP 1", "省级经济开发区", "以\u2018智慧管理\u2019为亮点，亩均论英雄评价体系、安全环保监控网络提供数据佐证", T.accent],
    ["STEP 2", "省级/国家级绿色工业园区", "以\u2018绿色数据\u2019为实证，全覆盖能耗与污染源监测体系自动生成绿色绩效报告", T.blue],
    ["STEP 3", "国家级经济技术开发区", "以\u2018创新生态\u2019为驱动，智慧招商系统精准补强外资、外贸、研发等短板指标", T.gold]];
  steps.forEach(function(s, i) {
    var y = 1.3 + i * 1.3;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 8.8, h: 1.05, fill: { color: T.secondary }, rectRadius: 0.08 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 0.06, h: 1.05, fill: { color: s[3] } });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.85, y: y + 0.2, w: 0.8, h: 0.65, fill: { color: s[3] }, rectRadius: 0.06 });
    sl.addText(s[0], { x: 0.85, y: y + 0.2, w: 0.8, h: 0.65, fontSize: 8, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(s[1], { x: 1.85, y: y + 0.08, w: 4, h: 0.35, fontSize: 14, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
    sl.addText(s[2], { x: 1.85, y: y + 0.45, w: 7.2, h: 0.5, fontSize: 10, fontFace: "Microsoft YaHei", color: T.sky });
    if (i < 2) sl.addShape(pres.shapes.LINE, { x: 0.97, y: y + 1.05, w: 0, h: 0.25, line: { color: s[3], width: 2 } });
  });
  badge(sl, pn, T.gold);
}

// === SLIDE 5: 7 Projects ===
function s5() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.accent);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.accent } });
  sectionTitle(sl, "七大建设工程", "覆盖招商、企业、运营、管理全业务");
  
  var bgC = [T.accent, T.blue, T.gold, "EB5757", "9B51E0", "2D9CDB", "219653"];
  var pjs = [["01","道路交通安全与地灾监控"], ["02","重点口子环保监控"], ["03","园区形象展示与AI招商"],
    ["04","企业用能监测"], ["05","智改数转企业赋能"], ["06","一体化数字平台"], ["07","运营服务中心"]];
  pjs.forEach(function(p, i) {
    var col = i % 3, row = Math.floor(i / 3);
    var x = 0.5 + col * 3.1, y = 1.2 + row * 1.3;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.8, h: 1.05, fill: { color: "FFFFFF" }, rectRadius: 0.08, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.04 } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 0.04, fill: { color: bgC[i] } });
    sl.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.15, w: 0.35, h: 0.35, fill: { color: bgC[i] } });
    sl.addText(p[0], { x: x + 0.15, y: y + 0.15, w: 0.35, h: 0.35, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(p[1], { x: x + 0.6, y: y + 0.15, w: 2.0, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: T.dark, bold: true, valign: "middle" });
    // Dots below
    for (var j = 0; j < 3; j++) dot(sl, x + 0.3 + j * 0.15, y + 0.65, 0.04, bgC[i]);
  });
  badge(sl, pn);
}

s4(); s5();

// === SLIDE 6: Platform Architecture ===
function s6() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.blue);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.blue } });
  sectionTitle(sl, "一体化数字平台架构", "物联网 · 数据中台 · AI中台 · 数字孪生");
  var biz = [["企业服务", T.accent, ["诉求直达（随手拍）","服务直达","园企交互","供需对接"]],
    ["园区运营", T.blue, ["数字化审批","日常巡检","资产管理","档案管理"]],
    ["信息发布", T.gold, ["用工·物流·特色","政策·餐饮·停车","场馆·活动","园区公告"]],
    ["数据支撑", "EB5757", ["12大三方系统","统一数据标准","统一认证中心","统一消息推送"]]];
  biz.forEach(function(b, i) {
    var x = 0.4 + i * 2.35;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.2, w: 2.15, h: 3.6, fill: { color: "FFFFFF" }, rectRadius: 0.08, shadow: { type: "outer", blur: 3, offset: 1, color: "000000", opacity: 0.04 } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.15, h: 0.04, fill: { color: b[1] } });
    sl.addText(b[0], { x, y: 1.3, w: 2.15, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: T.dark, bold: true, align: "center" });
    b[2].forEach(function(item, j) {
      dot(sl, x + 0.2, 1.85 + j * 0.6, 0.06, b[1]);
      sl.addText(item, { x: x + 0.35, y: 1.75 + j * 0.6, w: 1.6, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: T.gray });
    });
  });
  // Shape accent lines
  sl.addShape(pres.shapes.LINE, { x: 0.4, y: 4.8, w: 9.2, h: 0, line: { color: T.lightBg, width: 1 } });
  badge(sl, pn, T.blue);
}

s6();

// Feature slides with screenshots
featureSlide("AI招商 · 智能招商大脑", 
  ["AI精准匹配目标企业，基于产业链图谱智能推荐", "3D数字沙盘展示园区规划与土地资源", "全生命周期企业档案管理与画像分析", "招商项目全流程数字化跟踪（线索→签约）", "云展馆与电子名片，提升对外招商形象"],
  IMG + "/featured_supply.png", T.accent, "面向 政府领导层");

featureSlide("全域安全监控 · 天眼+哨兵+慧眼",
  ["天眼：道路交通安全监控，智能摄像机全时覆盖", "哨兵：地灾安全监测，滑坡、沉降实时预警", "慧眼：环保监控检测，废气废水排放在线监测", "视频监控总览：统一视频平台，16宫格智能排列", "应急指挥联动：一键调度、预案管理、实时通讯"],
  IMG + "/info_publish.png", T.blue, "面向 园区管理层");

featureSlide("企业服务 · 诉求直达与服务直达",
  ["诉求直达（随手拍）：文字+图片+定位快速提交", "服务直达：天然气/水/电费在线缴纳", "园企交互：在线沟通、通知公告、政策推送", "供需对接：物流供需、特色供需智能匹配", "全程透明：诉求进度实时可查、结果可评价"],
  IMG + "/info_publish.png", T.accent, "面向 入驻企业");

featureSlide("物流供需 · 货源+车源+返程空车",
  ["园区货源：企业货物运输需求发布与对接", "返程空车：返程车辆信息展示，提高配货率", "车辆资源：物流企业运力资源发布", "每张卡片带联系方式、有效期、路线信息", "三种类型卡片：园区货源/返程空车/车辆资源"],
  IMG + "/logistics_supply.png", T.gold, "面向 入驻企业");

featureSlide("特色供需 · 青川山珍展示",
  ["9种青川特色农产品（黑木耳/天麻/竹荪等）", "水平自动轮播画廊，hover暂停", "每张卡片：联系方式、价格、标签、有效期", "青川黑木耳为国家地理标志产品", "支持分类筛选与关键词搜索"],
  IMG + "/featured_supply.png", "9B51E0", "面向 公众用户");

featureSlide("园区政策 · 惠企精准推送",
  ["政策分类：补贴申报/收费减免/人才引进/税务", "精准推送：根据企业类型自动匹配政策", "15+条政策数据，支持搜索与筛选", "热度标签（热门/新/执行中）", "对接省级'智改数转'专项资金超20亿"],
  IMG + "/policy_detail.png", T.accent, "面向 入驻企业");

featureSlide("园区餐饮 · 餐厅展示+今日菜单",
  ["6家园区餐厅水平滑动卡片展示", "每家餐厅带SVG建筑背景和招牌标签", "20+道菜品，4列网格展示", "支持按套餐/面食/小炒/汤品/饮品分类", "左侧边栏菜品分类快速筛选"],
  IMG + "/catering.png", "EB5757", "面向 公众用户");

featureSlide("停车服务 · 地图导航+实时余位",
  ["高德地图可视化展示停车位分布", "实时余位数据，绿/黄/红状态标识", "多个停车场切换查看", "在线缴费与充值功能", "青川县竹园镇本地停车场数据"],
  IMG + "/parking_detail.png", "2D9CDB", "面向 入驻企业");

featureSlide("场馆服务 · 会议/体育/培训/展示",
  ["8个场馆卡片，2列网格展示", "每个场馆：图标+名称+位置+设施+价格", "支持按会议/体育/培训/展示分类筛选", "'立即预约'一键操作", "C区综合楼、B区运动中心等本地场馆"],
  IMG + "/venue_service.png", "219653", "面向 入驻企业");

featureSlide("园区活动 · 文体/培训/公益/交流",
  ["8个活动卡片，渐变背景+emoji图标", "状态标签：进行中/即将开始/筹备中", "活动名+描述+地点+时间+标签", "支持按类别筛选与关键词搜索", "篮球联赛、安全竞赛、植树公益等活动"],
  IMG + "/activities_detail.png", "9B51E0", "面向 公众用户");

featureSlide("园区公告 · 实时通知发布",
  ["10条公告数据，支持分类筛选", "标签：置顶/安全/新/通知", "每条公告带标题+描述+日期", "支持关键词搜索", "主干道施工、消防检查、网络升级等实时通知"],
  IMG + "/notice_detail.png", T.blue, "面向 公众用户");

featureSlide("用工信息 · 招聘+岗位匹配",
  ["8张岗位卡片，4列网格展示", "每张卡片：岗位名+状态+公司+薪资+招聘人数", "状态标签：招聘中/已招满", "岗位：研发工程师/操作工/品质主管等", "阜成新材、虹禾晶、中能建等园区企业"],
  IMG + "/employment_detail.png", T.accent, "面向 公众用户");

// === Role Feature Matrix ===
function sRoleMatrix() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.gold);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.gold } });
  sl.addText("四类角色功能矩阵", { x: 0.6, y: 0.2, w: 8, h: 0.5, fontSize: 22, fontFace: "Microsoft YaHei", color: T.dark, bold: true });
  sl.addText("每个角色可获取的功能全景图", { x: 0.6, y: 0.7, w: 8, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  
  // Header row
  var cols = ["功能模块", "政府领导层", "园区管理层", "入驻企业", "公众用户"];
  var headerC = [T.dark, T.accent, T.blue, T.gold, "EB5757"];
  cols.forEach(function(c, i) {
    var x = 0.4 + i * 1.85;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: (i==0?0.4:x), y: 1.1, w: (i==0?1.65:1.85), h: 0.35, fill: { color: headerC[i] }, rectRadius: 0.04 });
    sl.addText(c, { x: (i==0?0.4:x), y: 1.1, w: (i==0?1.65:1.85), h: 0.35, fontSize: (i==0?10:9), fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  });
  
  // Data rows
  var rows = [
    ["AI招商·数字沙盘", "●决策支撑", "●数据管理", "○信息查看", "—"],
    ["安全监控·应急", "●监管总览", "●日常运维", "○安全服务", "●安全信息"],
    ["企业服务·诉求", "●效能分析", "●流程管理", "●核心使用", "○投诉建议"],
    ["物流供需", "○产业数据", "●运营调度", "●发布对接", "○信息查看"],
    ["特色供需", "●经济展示", "○数据统计", "●产品发布", "●浏览购买"],
    ["园区政策", "●制定发布", "●执行监督", "●政策申报", "○政策查阅"],
    ["餐饮·停车·场馆", "—", "●运营管理", "●在线使用", "●预约预订"],
    ["活动·公告·用工", "●宏观数据", "●信息审核", "●参与发布", "●获取信息"],
    ["数据中心·能效", "●决策依据", "●监测预警", "○用能管理", "—"],
    ["投资概算", "●财政审批", "○预算参考", "—", "—"]
  ];
  
  rows.forEach(function(row, ri) {
    var y = 1.55 + ri * 0.34;
    var bg = ri % 2 === 0 ? "FFFFFF" : T.lightBg;
    row.forEach(function(cell, ci) {
      var x = ci === 0 ? 0.4 : (0.4 + (ci === 0 ? 0 : 1.65 + (ci-1) * 1.85));
      var w = ci === 0 ? 1.65 : 1.85;
      var isMain = cell.indexOf("●") === 0;
      var isView = cell.indexOf("○") === 0;
      sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.34, fill: { color: bg } });
      if (ci > 0 && (isMain || isView)) {
        var dotC = isMain ? T.accent : T.gray;
        dot(sl, x + 0.15, y + 0.1, 0.08, dotC);
      }
      sl.addText(cell, { x: x + (ci > 0 ? 0.35 : 0.1), y, w: w - (ci > 0 ? 0.45 : 0.15), h: 0.34, fontSize: (ci === 0 ? 10 : 9), fontFace: "Microsoft YaHei", color: (ci === 0 ? T.dark : T.gray), valign: "middle" });
    });
  });
  
  // Legend
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 5.0, w: 9, h: 0.3, fill: { color: T.sky }, rectRadius: 0.04 });
  dot(sl, 0.65, 5.06, 0.08, T.accent); sl.addText("● 主要功能  ", { x: 0.85, y: 5.0, w: 2, h: 0.3, fontSize: 9, fontFace: "Microsoft YaHei", color: T.dark, valign: "middle" });
  dot(sl, 2.8, 5.06, 0.08, T.gray); sl.addText("○ 辅助功能  ", { x: 3.0, y: 5.0, w: 2, h: 0.3, fontSize: 9, fontFace: "Microsoft YaHei", color: T.dark, valign: "middle" });
  sl.addText("— 无直接关联", { x: 5.0, y: 5.0, w: 2, h: 0.3, fontSize: 9, fontFace: "Microsoft YaHei", color: T.gray, valign: "middle" });
  badge(sl, pn, T.gold);
}

sRoleMatrix();

// === Investment ===
function sInvest() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.accent);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.accent } });
  sectionTitle(sl, "项目投资概算");
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.5, y: 1.1, w: 5, h: 1.4, fill: { color: T.primary }, rectRadius: 0.1, shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.1 } });
  sl.addText("总投资概算", { x: 2.5, y: 1.2, w: 5, h: 0.3, fontSize: 13, fontFace: "Microsoft YaHei", color: T.lightBg, align: "center" });
  sl.addText("¥4,178,800.00", { x: 2.5, y: 1.5, w: 5, h: 0.8, fontSize: 32, fontFace: "Arial", color: T.gold, bold: true, align: "center" });
  var details = ["道路交通安全与地灾监控","重点口子环保监控","园区形象展示与AI招商","企业用能监测","智改数转赋能平台","一体化数字平台","运营服务中心"];
  details.forEach(function(d, i) {
    var y = 2.8 + i * 0.3;
    dot(sl, 1.5, y + 0.06, 0.07, T.accent);
    sl.addText(d, { x: 1.75, y, w: 5, h: 0.28, fontSize: 10, fontFace: "Microsoft YaHei", color: T.dark });
    // Progress bar
    var pct = 0.05 + Math.random() * 0.2;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: y + 0.06, w: 2.8, h: 0.14, fill: { color: T.lightBg }, rectRadius: 0.07 });
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: y + 0.06, w: 2.8 * (i+1)/7, h: 0.14, fill: { color: T.accent }, rectRadius: 0.07 });
  });
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.0, y: 4.7, w: 2.6, h: 0.28, fill: { color: T.accent }, rectRadius: 0.05 });
  sl.addText("面向 政府领导层", { x: 7.0, y: 4.7, w: 2.6, h: 0.28, fontSize: 8, fontFace: "Microsoft YaHei", color: "FFFFFF", align: "center", valign: "middle" });
  badge(sl, pn);
}

// === Benefits ===
function sBenefits() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.gold);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.gold } });
  sectionTitle(sl, "预期效益分析");
  var metrics = [["30%", "管理效率提升", T.accent], ["60%", "应急响应提升", T.blue], ["¥20亿+", "专项资金对接", T.gold], ["百万级", "创建奖励", "EB5757"]];
  metrics.forEach(function(m, i) {
    var x = 0.5 + i * 2.35;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.2, w: 2.15, h: 1.8, fill: { color: "FFFFFF" }, rectRadius: 0.08, shadow: { type: "outer", blur: 4, offset: 1, color: "000000", opacity: 0.04 } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.2, w: 2.15, h: 0.04, fill: { color: m[2] } });
    sl.addText(m[0], { x, y: 1.4, w: 2.15, h: 0.5, fontSize: 28, fontFace: "Arial", color: m[2], bold: true, align: "center" });
    sl.addText(m[1], { x, y: 1.9, w: 2.15, h: 0.25, fontSize: 11, fontFace: "Microsoft YaHei", color: T.dark, bold: true, align: "center" });
  });
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.3, w: 9, h: 1.3, fill: { color: T.sky }, rectRadius: 0.08 });
  sl.addText("实施策略：急用先行、分步上线", { x: 0.7, y: 3.4, w: 8.5, h: 0.35, fontSize: 13, fontFace: "Microsoft YaHei", color: T.dark, bold: true });
  sl.addText("优先完成环保、安防、招商展示等支撑创建申报的核心功能，确保项目成果能快速服务于省级开发区申报。预计综合管理效率提升30%以上，安全环保事件从被动响应转向主动预警，应急响应效率提升60%以上。", { x: 0.7, y: 3.8, w: 8.5, h: 0.7, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  badge(sl, pn, T.gold);
}

sInvest(); sBenefits();

// === Future ===
function sFuture() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.lightBg };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.04, T.blue);
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: T.blue } });
  sectionTitle(sl, "未来展望", "以数字为帆，以智慧为桨");
  var items = ["AIoT深度融合：引入物联网传感设备，万物互联、全域感知","数字孪生园区：1:1三维建模，全要素可视化运营","AI Agent智能化：智能客服、调度、预警","产业大脑：大数据+AI辅助招商决策","跨园区协同：对接省市智慧平台"];
  items.forEach(function(item, i) {
    var y = 1.3 + i * 0.65;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.5, fill: { color: "FFFFFF" }, rectRadius: 0.06, shadow: { type: "outer", blur: 2, offset: 1, color: "000000", opacity: 0.03 } });
    dot(sl, 0.65, y + 0.12, 0.08, T.blue);
    sl.addText(item, { x: 0.9, y, w: 8.3, h: 0.5, fontSize: 12, fontFace: "Microsoft YaHei", color: T.dark, valign: "middle" });
  });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.6, w: 9, h: 0.02, fill: { color: T.lightBg } });
  sl.addText("驱动园区沿着\u2018省级开发区→绿色园区→国家级经开区培育对象\u2019的航向全速前进", { x: 0.5, y: 4.7, w: 9, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.secondary, italic: true, align: "center" });
  badge(sl, pn, T.blue);
}

sFuture();

// === Thank You ===
function sThank() {
  pn++; var sl = pres.addSlide();
  sl.background = { color: T.primary };
  for (var i = 0; i < 8; i++) dot(sl, 0.4 + i * 0.25, 0.08, 0.05, T.gold);
  sl.addShape(pres.shapes.OVAL, { x: 3, y: 0.5, w: 4, h: 4, fill: { color: T.secondary, transparency: 60 } });
  sl.addText("感谢聆听", { x: 1, y: 1.5, w: 8, h: 1.0, fontSize: 38, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center" });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4, y: 2.5, w: 2, h: 0.04, fill: { color: T.gold } });
  sl.addText("以数字为帆，以智慧为桨", { x: 1, y: 2.7, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: T.gold, align: "center" });
  ["青川县庄子上工业园区管委会","技术支持：青迅科技","地址：四川省广元市青川县竹园镇"].forEach(function(item, i) {
    sl.addText(item, { x: 1, y: 3.5 + i * 0.4, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: T.lightBg, align: "center" });
  });
}

sThank();

pres.writeFile({ fileName: path.join(__dirname, "output", "zhuangzishang_full_report.pptx") })
  .then(function() { console.log("PPT OK: " + pn + " slides"); })
  .catch(function(e) { console.error("ERR:", e); });
