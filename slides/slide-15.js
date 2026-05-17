const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'section', index: 15, title: '平台价值总结' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.8, w: 0.08, h: 1.5, fill: { color: theme.accent } });
  slide.addText("平台价值总结", { x: 0.9, y: 1.8, w: 7, h: 1.0, fontSize: 30, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 2.9, w: 1.5, h: 0.04, fill: { color: theme.accent } });
  slide.addText("• 政府领导层：数据驱动决策，提升园区治理能力", { x: 0.9, y: 3.3, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 园区管理层：运营效率提升30%，管理成本降低15%-20%", { x: 0.9, y: 3.6999999999999997, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 入驻企业：一站式服务，诉求响应时间缩短至2小时", { x: 0.9, y: 4.1, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 公众用户：信息公开透明，增强园区参与感", { x: 0.9, y: 4.5, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 生态效益：能耗降低约18%，碳排放监测全覆盖", { x: 0.9, y: 4.9, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("15", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-15-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
