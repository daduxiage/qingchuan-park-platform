const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'text', index: 12, title: '物流供需平台' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addText("物流供需平台", { x: 0.6, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: theme.primary, bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.9, w: 1.0, h: 0.04, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.35, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("园区货源：园区企业货物运输需求发布", { x: 0.9, y: 1.3, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.05, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("返程空车：返程车辆信息展示，提高回程配货率", { x: 0.9, y: 2.0, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.75, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("车辆资源：物流企业运力资源对接", { x: 0.9, y: 2.7, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.4499999999999993, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("每张卡片展示：货物信息、路线、联系人电话", { x: 0.9, y: 3.3999999999999995, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.1499999999999995, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("AI智能匹配：自动推送匹配的货源与运力信息", { x: 0.9, y: 4.1, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("12", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-12-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
