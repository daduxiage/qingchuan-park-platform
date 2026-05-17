const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'section', index: 05, title: 'AI招商服务' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.8, w: 0.08, h: 1.5, fill: { color: theme.accent } });
  slide.addText("AI招商服务", { x: 0.9, y: 1.8, w: 7, h: 1.0, fontSize: 30, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 2.9, w: 1.5, h: 0.04, fill: { color: theme.accent } });
  slide.addText("• AI智能招商：基于大数据和AI算法精准匹配目标企业", { x: 0.9, y: 3.3, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 数字沙盘：3D可视化展示园区规划与土地资源", { x: 0.9, y: 3.6999999999999997, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 企业档案：全生命周期企业信息管理与画像分析", { x: 0.9, y: 4.1, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 产业链图谱：可视化展示产业链上下游关系", { x: 0.9, y: 4.5, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 招商项目管理：从线索到签约全流程数字化跟踪", { x: 0.9, y: 4.9, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("05", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-05-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
