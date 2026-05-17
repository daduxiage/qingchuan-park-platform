const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'section', index: 09, title: '企业服务体系' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.8, w: 0.08, h: 1.5, fill: { color: theme.accent } });
  slide.addText("企业服务体系", { x: 0.9, y: 1.8, w: 7, h: 1.0, fontSize: 30, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 2.9, w: 1.5, h: 0.04, fill: { color: theme.accent } });
  slide.addText("• 诉求直达（随手拍）：企业问题一键提交、全程跟踪", { x: 0.9, y: 3.3, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 服务直达：天然气/水/电费在线缴纳", { x: 0.9, y: 3.6999999999999997, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 园企交互：在线沟通、通知公告、政策推送", { x: 0.9, y: 4.1, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 供需对接：物流供需、特色供需智能匹配", { x: 0.9, y: 4.5, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 企业入驻：线上申请、证照管理、信用评价", { x: 0.9, y: 4.9, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("09", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-09-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
