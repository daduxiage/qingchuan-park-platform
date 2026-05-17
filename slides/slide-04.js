const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'text', index: 04, title: '平台整体架构' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addText("平台整体架构", { x: 0.6, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: theme.primary, bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.9, w: 1.0, h: 0.04, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.35, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("一云一网一平台：统一云底座 + 园区通信网 + 智慧平台", { x: 0.9, y: 1.3, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.05, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("四层架构：基础设施层、数据层、平台层、应用层", { x: 0.9, y: 2.0, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.75, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("九大核心应用模块：AI招商、安全监控、企业服务等", { x: 0.9, y: 2.7, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.4499999999999993, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("多端覆盖：Web管理端 + 企业端 + 公众端 + 大屏展示", { x: 0.9, y: 3.3999999999999995, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.1499999999999995, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("统一认证中心、统一数据标准、统一消息推送", { x: 0.9, y: 4.1, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("04", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-04-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
