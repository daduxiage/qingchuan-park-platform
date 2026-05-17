const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'text', index: 10, title: '诉求直达与服务直达' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addText("诉求直达与服务直达", { x: 0.6, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: theme.primary, bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.9, w: 1.0, h: 0.04, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.35, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("随手拍功能：文字+图片+定位，快速提交诉求", { x: 0.9, y: 1.3, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.05, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("诉求分类：物业报修、环境问题、投诉建议等", { x: 0.9, y: 2.0, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.75, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("自动派单：基于AI规则智能分配到责任部门", { x: 0.9, y: 2.7, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.4499999999999993, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("服务大厅：水、电、气、停车等一站式在线缴费", { x: 0.9, y: 3.3999999999999995, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.1499999999999995, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("全程透明：诉求进度实时可查，处理结果可评价", { x: 0.9, y: 4.1, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("10", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-10-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
