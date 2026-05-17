const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'text', index: 06, title: '数字沙盘与企业档案' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addText("数字沙盘与企业档案", { x: 0.6, y: 0.3, w: 8, h: 0.6, fontSize: 22, fontFace: "Microsoft YaHei", color: theme.primary, bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.9, w: 1.0, h: 0.04, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.35, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("3D数字沙盘：园区全景三维可视化，支持多视角漫游", { x: 0.9, y: 1.3, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.05, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("土地资源展示：标注可用地块、面积、规划用途", { x: 0.9, y: 2.0, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.75, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("企业画像：多维数据分析展示企业经营、信用、创新力", { x: 0.9, y: 2.7, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 3.4499999999999993, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("产业分析：自动识别产业链关键环节与缺口", { x: 0.9, y: 3.3999999999999995, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.1499999999999995, w: 0.12, h: 0.12, fill: { color: theme.accent } });
  slide.addText("招商门户：对外展示园区优势，吸引优质企业入驻", { x: 0.9, y: 4.1, w: 8.2, h: 0.5, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.secondary });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("06", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-06-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
