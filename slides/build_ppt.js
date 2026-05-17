const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

const T = { // Theme - Deep Blue Tech
  primary: "0A1628",
  secondary: "1B3A5C",
  accent: "2D8CF0",
  highlight: "00D4AA",
  light: "E8F0FE",
  bg: "FFFFFF",
  gray: "6B7A8D",
  darkText: "1A2332",
  cardBg: "F0F4F8"
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

let pageNum = 0;

// ====== Slide 1: Cover ======
function slide1() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: T.primary };
  addTopBar(sl);
  // Decorative circles
  sl.addShape(pres.shapes.OVAL, { x: 7.5, y: -1, w: 4, h: 4, fill: { color: T.secondary, transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -1.5, y: 3.5, w: 3, h: 3, fill: { color: T.secondary, transparency: 60 } });
  // Left accent
  addSectionAccent(sl, 0.6, 1.0, 2.0);
  sl.addText("青川县庄子上工业园区", { x: 0.9, y: 1.0, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: T.highlight });
  sl.addText("智慧园区建设项目", { x: 0.9, y: 1.5, w: 8, h: 1.2, fontSize: 38, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addText("建设方案汇报", { x: 0.9, y: 2.7, w: 6, h: 0.5, fontSize: 22, fontFace: "Microsoft YaHei", color: T.light });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 3.4, w: 2, h: 0.04, fill: { color: T.accent } });
  // Audience labels
  const audiences = ["面向政府领导层", "面向园区管理层", "面向入驻企业", "面向公众用户"];
  audiences.forEach((a, i) => {
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.9 + i * 2.1, y: 3.8, w: 1.9, h: 0.35, fill: { color: T.secondary }, rectRadius: 0.05 });
    sl.addText(a, { x: 0.9 + i * 2.1, y: 3.8, w: 1.9, h: 0.35, fontSize: 9, fontFace: "Microsoft YaHei", color: T.light, align: "center", valign: "middle" });
  });
  sl.addText("2026年5月", { x: 0.6, y: 5.0, w: 3, h: 0.4, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  // Bottom
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: T.accent } });
}

// ====== Slide 2: TOC ======
function slide2() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: T.primary };
  addTopBar(sl);
  sl.addText("汇报目录", { x: 0.8, y: 0.3, w: 8, h: 0.6, fontSize: 24, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 0.9, w: 1.2, h: 0.04, fill: { color: T.accent } });
  const items = [
    "项目战略定位", "三步走创建路径", "七大建设工程", "道路交通安全与地灾监控",
    "环保监控工程", "AI招商与形象展示", "企业用能监测", "智改数转赋能平台",
    "一体化数字平台", "企业服务板块", "信息发布子系统", "运营服务中心",
    "投资概算", "预期效益分析", "四类角色价值", "未来展望"
  ];
  items.forEach((item, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.8 + col * 4.5, y = 1.3 + row * 0.55;
    sl.addShape(pres.shapes.OVAL, { x, y: y + 0.06, w: 0.28, h: 0.28, fill: { color: T.accent } });
    sl.addText(String(i + 1), { x, y: y + 0.06, w: 0.28, h: 0.28, fontSize: 10, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(item, { x: x + 0.38, y, w: 3.8, h: 0.4, fontSize: 12, fontFace: "Microsoft YaHei", color: "FFFFFF", valign: "middle" });
  });
  addBadge(sl, pageNum);
}

// ====== Slide 3: Project Strategy ======
function slide3() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" };
  addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("项目战略定位", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("智慧化建设是园区实现跨越式发展的"必由之路"与"关键一跃"", { x: 0.75, y: 0.8, w: 8.5, h: 0.4, fontSize: 12, fontFace: "Microsoft YaHei", color: T.accent });
  
  const cards = [
    { title: "赋能创建", desc: "为申报省级经开区、绿色工业园区提供最直接的数据支撑与管理工具", color: T.accent },
    { title: "赋能企业", desc: "打造县域"智改数转"公共服务平台，破解中小企业转型难题", color: T.highlight },
    { title: "赋能决策", desc: "构建园区"数字大脑"，实现运行一屏统览、产业一图统管", color: "F59E0B" }
  ];
  cards.forEach((card, i) => {
    const x = 0.5 + i * 3.1;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.4, w: 2.8, h: 2.2, fill: { color: T.cardBg }, rectRadius: 0.1, shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.08 } });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.4, w: 2.8, h: 0.06, fill: { color: card.color } });
    sl.addText(card.title, { x: x + 0.2, y: 1.65, w: 2.4, h: 0.4, fontSize: 15, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
    sl.addText(card.desc, { x: x + 0.2, y: 2.15, w: 2.4, h: 1.2, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  });

  // Stats
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.9, w: 9, h: 0.04, fill: { color: T.light } });
  sl.addText("产业基础：2025年实现规上工业总产值", { x: 0.5, y: 4.1, w: 5, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  sl.addText("41.6亿元", { x: 4.6, y: 4.1, w: 2, h: 0.3, fontSize: 14, fontFace: "Microsoft YaHei", color: T.accent, bold: true });
  sl.addText("，主导产业集聚效应初显", { x: 6.0, y: 4.1, w: 3, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  
  sl.addText("总投资概算：¥4,178,800.00", { x: 0.5, y: 4.5, w: 6, h: 0.3, fontSize: 12, fontFace: "Microsoft YaHei", color: T.secondary, bold: true });
  addBadge(sl, pageNum);
}

// ====== Slide 4: Three-Step Path ======
function slide4() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: T.primary };
  addTopBar(sl);
  sl.addText(""三步走"创建路径", { x: 0.7, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  
  const steps = [
    { title: "第一步", subtitle: "省级经济开发区", desc: "以"智慧管理"为亮点，亩均论英雄评价体系、安全环保监控网络提供真实数据佐证", color: T.accent },
    { title: "第二步", subtitle: "省级/国家级绿色工业园区", desc: "以"绿色数据"为实证，全覆盖能耗与污染源监测体系自动生成绿色绩效报告", color: T.highlight },
    { title: "第三步", subtitle: "国家级经济技术开发区", desc: "以"创新生态"为驱动，智慧招商系统精准补强外资、外贸、研发等短板指标", color: "F59E0B" }
  ];
  
  steps.forEach((s, i) => {
    const y = 1.1 + i * 1.4;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.7, y, w: 8.6, h: 1.15, fill: { color: T.secondary }, rectRadius: 0.08 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.7, y, w: 0.06, h: 1.15, fill: { color: s.color } });
    sl.addShape(pres.shapes.OVAL, { x: 1.1, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: s.color } });
    sl.addText(s.title, { x: 1.1, y: y + 0.3, w: 0.55, h: 0.55, fontSize: 10, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    sl.addText(s.subtitle, { x: 1.85, y: y + 0.1, w: 3.5, h: 0.4, fontSize: 14, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
    sl.addText(s.desc, { x: 1.85, y: y + 0.55, w: 7, h: 0.5, fontSize: 10, fontFace: "Microsoft YaHei", color: T.light });
    if (i < 2) {
      sl.addShape(pres.shapes.LINE, { x: 0.97, y: y + 1.15, w: 0, h: 0.25, line: { color: T.accent, width: 2 } });
    }
  });
  addBadge(sl, pageNum);
}

// ====== Slide 5: Seven Projects Overview ======
function slide5() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" };
  addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("七大建设工程", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("覆盖"招商、企业、运营、管理"全业务的一体化智慧园区平台", { x: 0.75, y: 0.8, w: 8, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });

  const projects = [
    { n: "01", t: "道路交通安全与地灾监控", d: "智能摄像机+地质传感器，全覆盖感知网络" },
    { n: "02", t: "重点口子环保监控", d: "水质空气在线监测，监测-预警-溯源闭环" },
    { n: "03", t: "园区形象展示与AI招商", d: "三维数字沙盘+数据驾驶舱+AI招商大脑" },
    { n: "04", t: "企业用能监测", d: "智能计量表计+亩均论英雄大数据评价" },
    { n: "05", t: "智改数转企业赋能平台", d: "政策推送+供需对接+数字化诊断工具" },
    { n: "06", t: "一体化数字平台", d: "物联网平台+数据中台+AI中台+数字孪生" },
    { n: "07", t: "运营服务中心", d: "大屏指挥+园小二团队+3年长效服务" }
  ];

  projects.forEach((p, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.1, y = 1.3 + row * 1.3;
    const bgColors = [T.accent, T.highlight, "F59E0B", "EB5757", "9B51E0", "2D9CDB", "219653"];
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.8, h: 1.1, fill: { color: T.cardBg }, rectRadius: 0.08 });
    sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.8, h: 0.05, fill: { color: bgColors[i] } });
    sl.addText(p.n, { x: x + 0.15, y: y + 0.15, w: 0.35, h: 0.35, fontSize: 16, fontFace: "Arial", color: bgColors[i], bold: true });
    sl.addText(p.t, { x: x + 0.55, y: y + 0.12, w: 2.1, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
    sl.addText(p.d, { x: x + 0.15, y: y + 0.55, w: 2.5, h: 0.45, fontSize: 9, fontFace: "Microsoft YaHei", color: T.gray });
  });
  addBadge(sl, pageNum);
}

// ====== Slides 6-12: Detailed project slides ======
function slideDetail(title, items, accentColor) {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" };
  addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText(title, { x: 0.75, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: 0.8, w: 1.0, h: 0.03, fill: { color: accentColor } });

  items.forEach((item, i) => {
    const y = 1.15 + i * 0.7;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.6, fill: { color: T.cardBg }, rectRadius: 0.06 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.05, h: 0.6, fill: { color: accentColor } });
    const icon = ["❶", "❷", "❸", "❹", "❺"][i] || "•";
    sl.addText(icon, { x: 0.7, y, w: 0.35, h: 0.6, fontSize: 12, fontFace: "Microsoft YaHei", color: accentColor, align: "center", valign: "middle" });
    if (typeof item === "string") {
      sl.addText(item, { x: 1.15, y, w: 8.1, h: 0.6, fontSize: 12, fontFace: "Microsoft YaHei", color: T.darkText, valign: "middle" });
    } else {
      sl.addText(item.title, { x: 1.15, y, w: 4, h: 0.3, fontSize: 12, fontFace: "Microsoft YaHei", color: T.primary, bold: true, valign: "middle" });
      sl.addText(item.desc, { x: 1.15, y: y + 0.28, w: 8.1, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: T.gray, valign: "middle" });
    }
  });
  addBadge(sl, pageNum);
}

slide1();
slide2();
slide3();
slide4();
slide5();
slideDetail("道路交通安全与地灾监控", [
  { title: "智能摄像机覆盖", desc: "部署智能摄像机与专业地质传感器，构建全覆盖感知网络" },
  { title: "自动识别与实时预警", desc: "实现对交通违规、地质灾害风险的自动识别与实时预警" },
  { title: "安全运行底线", desc: "筑牢园区安全运行底线，为创建评审提供"安全稳定"的坚实基础" },
  { title: "赋能创建", desc: "为省级经开区"发展质量"与绿色园区"基础设施"安全提供硬性支撑" }
], T.accent);

slideDetail("重点口子环保监控工程", [
  { title: "在线监测设备布设", desc: "在关键排污口及环境敏感区域布设水质、空气在线监测设备" },
  { title: "闭环管理体系", desc: "建立"监测-预警-溯源"的闭环管理体系" },
  { title: "连续数据证据", desc: "为污染物稳定达标排放率提供连续、真实、可追溯的在线监测数据" },
  { title: "绿色园区核心支撑", desc: "是申报绿色工业园区最具说服力的"数据证据"" }
], T.highlight);

slideDetail("园区形象展示与AI招商工程", [
  { title: "三维数字沙盘", desc: "建设三维数字沙盘、数据驾驶舱，打造沉浸式现代化展示窗口" },
  { title: "数实招商大脑", desc: "运用大数据与AI技术精准绘制产业链图谱，智能挖掘匹配招商目标" },
  { title: "对外开放形象提升", desc: "直接服务于国家级经开区"对外开放"形象提升" },
  { title: "产业结构优化", desc: "驱动高质量项目落地，助力优化产业结构，补强科技创新指标" }
], "F59E0B");

slideDetail("企业用能监测与"亩均论英雄"", [
  { title: "智能计量表计", desc: "推动重点企业安装智能计量表计，实时采集能耗与产出数据" },
  { title: "亩均论英雄评价体系", desc: "构建"亩均论英雄"大数据评价体系，实现资源要素差别化配置" },
  { title: "绿色园区核算基础", desc: "为绿色园区"能源与资源利用绿色化"指标提供精准核算基础" },
  { title: "倒逼企业提质增效", desc: "以数据驱动的方式倒逼企业提质增效、转型升级" }
], "EB5757");

slideDetail(""智改数转"企业赋能平台", [
  { title: "政策精准推送", desc: "集成政策精准推送、供需智能对接、轻量化数字化诊断工具等功能" },
  { title: "专项资金对接", desc: "帮助企业对接省、市每年超20亿元的"智改数转"专项资金" },
  { title: "破解转型困境", desc: "破解企业"不敢转、不会转、不能转、不愿转"的困境" },
  { title: "产业基础夯实", desc: "打造园区优质服务生态，夯实所有创建目标的产业基础" }
], "9B51E0");

// ====== Slide: Four Main Business Lines ======
function slide13() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: T.primary };
  addTopBar(sl);
  sl.addText("一体化数字平台 - 四大业务线", { x: 0.7, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  sl.addText("物联网平台 | 数据中台 | AI中台 | 数字孪生底座", { x: 0.7, y: 0.8, w: 8, h: 0.35, fontSize: 11, fontFace: "Microsoft YaHei", color: T.highlight });

  const biz = [
    { title: "企业服务板块", items: ["诉求直达（随手拍）", "服务直达", "园企交互", "供需对接"], color: T.accent },
    { title: "园区运营板块", items: ["数字化审批", "日常巡检", "资产管理", "档案管理"], color: T.highlight },
    { title: "信息发布板块", items: ["用工·物流·特色供需", "政策推送·餐饮·停车", "场馆服务·园区活动", "园区公告"], color: "F59E0B" },
    { title: "数据支撑板块", items: ["12大三方系统接入", "统一数据标准", "统一认证中心", "统一消息推送"], color: "EB5757" }
  ];

  biz.forEach((b, i) => {
    const x = 0.5 + i * 2.35;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.3, w: 2.15, h: 3.5, fill: { color: T.secondary }, rectRadius: 0.08 });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.3, w: 2.15, h: 0.4, fill: { color: b.color } });
    sl.addText(b.title, { x, y: 1.3, w: 2.15, h: 0.4, fontSize: 11, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
    b.items.forEach((item, j) => {
      sl.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 2.0 + j * 0.6, w: 0.08, h: 0.08, fill: { color: b.color } });
      sl.addText(item, { x: x + 0.4, y: 1.9 + j * 0.6, w: 1.6, h: 0.3, fontSize: 10, fontFace: "Microsoft YaHei", color: "FFFFFF" });
    });
  });
  addBadge(sl, pageNum);
}

slide13();

// ====== Slide: Investment ======
function slideInvestment() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" };
  addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("项目投资概算", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  
  sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.5, y: 1.2, w: 5, h: 1.5, fill: { color: T.primary }, rectRadius: 0.1 });
  sl.addText("总投资概算", { x: 2.5, y: 1.35, w: 5, h: 0.35, fontSize: 14, fontFace: "Microsoft YaHei", color: T.light, align: "center" });
  sl.addText("¥4,178,800.00", { x: 2.5, y: 1.7, w: 5, h: 0.8, fontSize: 32, fontFace: "Arial", color: T.highlight, bold: true, align: "center" });

  const details = [
    "道路交通安全与地灾监控", "重点口子环保监控", "园区形象展示与AI招商",
    "企业用能监测", "智改数转赋能平台", "一体化数字平台", "运营服务中心"
  ];
  details.forEach((d, i) => {
    const y = 3.0 + i * 0.3;
    sl.addShape(pres.shapes.OVAL, { x: 1.5, y: y + 0.08, w: 0.08, h: 0.08, fill: { color: T.accent } });
    sl.addText(d, { x: 1.75, y, w: 6, h: 0.28, fontSize: 10, fontFace: "Microsoft YaHei", color: T.darkText });
  });
  addBadge(sl, pageNum);
}

slideInvestment();

// ====== Slide: Expected Benefits ======
function slideBenefits() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" };
  addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("预期效益分析", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });

  const metrics = [
    { num: "30%", label: "管理效率提升", desc: "综合管理效率提升30%以上", color: T.accent },
    { num: "60%", label: "应急响应提升", desc: "安全环保事件从被动响应转向主动预警，效率提升60%", color: T.highlight },
    { num: "¥20亿+", label: "专项资金对接", desc: "对接省级"智改数转"专项资金及金融产品", color: "F59E0B" },
    { num: "百万级", label: "创建奖励", desc: "成功创建绿色园区可获百万级一次性奖励", color: "EB5757" }
  ];

  metrics.forEach((m, i) => {
    const x = 0.5 + i * 2.35;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.2, w: 2.15, h: 2.0, fill: { color: T.cardBg }, rectRadius: 0.1 });
    sl.addText(m.num, { x, y: 1.4, w: 2.15, h: 0.6, fontSize: 28, fontFace: "Arial", color: m.color, bold: true, align: "center" });
    sl.addText(m.label, { x, y: 2.0, w: 2.15, h: 0.3, fontSize: 11, fontFace: "Microsoft YaHei", color: T.primary, bold: true, align: "center" });
    sl.addText(m.desc, { x: x + 0.1, y: 2.35, w: 1.95, h: 0.6, fontSize: 9, fontFace: "Microsoft YaHei", color: T.gray, align: "center" });
  });

  // Implementation approach
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9, h: 0.04, fill: { color: T.light } });
  sl.addText("实施策略："急用先行、分步上线"", { x: 0.5, y: 3.7, w: 9, h: 0.35, fontSize: 13, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("优先完成环保、安防、招商展示等支撑创建申报的核心功能，确保项目成果能快速服务于省级开发区的申报工作。", { x: 0.5, y: 4.1, w: 9, h: 0.5, fontSize: 11, fontFace: "Microsoft YaHei", color: T.gray });
  addBadge(sl, pageNum);
}

slideBenefits();

// ====== Slide: Role-Specific Value ======
function slideRoles() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: T.primary };
  addTopBar(sl);
  sl.addText("面向四类角色的核心价值", { x: 0.7, y: 0.2, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });

  const roles = [
    { title: "政府领导层", desc: "数据驱动决策，提升园区治理能力现代化水平；成功创建省级经开区/绿色园区，争取百万级政策奖励", color: T.accent, icon: "🏛" },
    { title: "园区管理层", desc: "运营效率提升30%，应急响应提速60%；一屏统览、一图统管，告别信息孤岛", color: T.highlight, icon: "📊" },
    { title: "入驻企业", desc: "一站式服务，诉求直达2小时响应；对接超20亿智改数转专项资金，降低转型成本", color: "F59E0B", icon: "🏢" },
    { title: "公众用户", desc: "信息公开透明，园区动态一手掌握；参与园区共建，增强归属感与获得感", color: "EB5757", icon: "👥" }
  ];

  roles.forEach((r, i) => {
    const x = 0.4 + i * 2.35;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.1, w: 2.2, h: 3.8, fill: { color: T.secondary }, rectRadius: 0.08 });
    sl.addShape(pres.shapes.RECTANGLE, { x, y: 1.1, w: 2.2, h: 0.06, fill: { color: r.color } });
    sl.addText(r.icon, { x, y: 1.3, w: 2.2, h: 0.5, fontSize: 24, align: "center" });
    sl.addText(r.title, { x, y: 1.8, w: 2.2, h: 0.3, fontSize: 13, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center" });
    sl.addText(r.desc, { x: x + 0.15, y: 2.2, w: 1.9, h: 2.5, fontSize: 10, fontFace: "Microsoft YaHei", color: T.light, align: "center" });
  });
  addBadge(sl, pageNum);
}

slideRoles();

// ====== Slide: Future Outlook ======
function slideFuture() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: "FFFFFF" };
  addTopBar(sl);
  addSectionAccent(sl, 0.5, 0.25, 0.5);
  sl.addText("未来展望", { x: 0.75, y: 0.2, w: 6, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: T.primary, bold: true });
  sl.addText("以"数字"为帆，以"智慧"为桨", { x: 0.75, y: 0.8, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: T.accent });

  const items = [
    "AIoT深度融合：引入物联网传感设备，实现万物互联、全域感知",
    "数字孪生园区：1:1三维建模，全要素、全时域可视化运营管理",
    "AI Agent智能化：智能客服、智能调度、智能预警，从人找数据到数据找人",
    "产业大脑：基于大数据与AI算法辅助产业招商决策，精准强链补链",
    "跨园区协同：对接省市智慧平台，实现数据共享与业务协同"
  ];

  items.forEach((item, i) => {
    const y = 1.4 + i * 0.7;
    sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.55, fill: { color: T.cardBg }, rectRadius: 0.06 });
    sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y + 0.1, w: 0.05, h: 0.35, fill: { color: T.accent } });
    sl.addText(item, { x: 0.75, y, w: 8.5, h: 0.55, fontSize: 12, fontFace: "Microsoft YaHei", color: T.darkText, valign: "middle" });
  });

  // Quote
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.7, w: 9, h: 0.02, fill: { color: T.light } });
  sl.addText("驱动园区沿着"省级开发区→绿色园区→国家级经开区培育对象"的航向全速前进", { x: 0.5, y: 4.85, w: 9, h: 0.3, fontSize: 12, fontFace: "Microsoft YaHei", color: T.secondary, italic: true, align: "center" });
  addBadge(sl, pageNum);
}

slideFuture();

// ====== Slide: Thank You ======
function slideThankYou() {
  pageNum++;
  const sl = pres.addSlide();
  sl.background = { color: T.primary };
  addTopBar(sl);
  sl.addShape(pres.shapes.OVAL, { x: 3, y: 0.5, w: 4, h: 4, fill: { color: T.secondary, transparency: 60 } });
  sl.addShape(pres.shapes.OVAL, { x: -1, y: 4, w: 3, h: 3, fill: { color: T.secondary, transparency: 60 } });
  
  sl.addText("感谢聆听", { x: 1, y: 1.5, w: 8, h: 1.0, fontSize: 38, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center" });
  sl.addShape(pres.shapes.RECTANGLE, { x: 4, y: 2.5, w: 2, h: 0.04, fill: { color: T.accent } });
  sl.addText("以"数字"为帆，以"智慧"为桨", { x: 1, y: 2.7, w: 8, h: 0.5, fontSize: 16, fontFace: "Microsoft YaHei", color: T.highlight, align: "center" });
  
  const info = ["青川县庄子上工业园区管委会", "技术支持：青迅科技", "联系电话：0839-xxxxxxx", "邮箱：park@zhuangzishang.gov.cn"];
  info.forEach((item, i) => {
    sl.addText(item, { x: 1, y: 3.5 + i * 0.4, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: T.light, align: "center" });
  });
  
  sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: T.accent } });
}

slideThankYou();

// Write file
const outputPath = path.join(__dirname, "output", "zhuangzishang_platform_report.pptx");
pres.writeFile({ fileName: outputPath })
  .then(() => console.log("PPT generated successfully: " + outputPath))
  .catch(err => console.error("Error:", err));
