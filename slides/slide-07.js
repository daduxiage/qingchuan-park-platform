const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'section', index: 07, title: '全域安全监控' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.8, w: 0.08, h: 1.5, fill: { color: theme.accent } });
  slide.addText("全域安全监控", { x: 0.9, y: 1.8, w: 7, h: 1.0, fontSize: 30, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 2.9, w: 1.5, h: 0.04, fill: { color: theme.accent } });
  slide.addText("• 道路交通安全监控（天眼）：园区道路全时监控", { x: 0.9, y: 3.3, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 地灾安全监测（哨兵）：地质滑坡、沉降实时预警", { x: 0.9, y: 3.6999999999999997, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 环保监控检测（慧眼）：废气废水排放在线监测", { x: 0.9, y: 4.1, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 视频监控总览：统一视频平台，覆盖园区关键区域", { x: 0.9, y: 4.5, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 应急指挥联动：突发事件快速响应与资源调度", { x: 0.9, y: 4.9, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("07", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-07-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
