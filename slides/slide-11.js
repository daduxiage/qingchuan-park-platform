const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'section', index: 11, title: '运营管理平台' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.8, w: 0.08, h: 1.5, fill: { color: theme.accent } });
  slide.addText("运营管理平台", { x: 0.9, y: 1.8, w: 7, h: 1.0, fontSize: 30, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 2.9, w: 1.5, h: 0.04, fill: { color: theme.accent } });
  slide.addText("• 物流供需：园区货源+返程空车+车辆资源信息匹配", { x: 0.9, y: 3.3, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 特色供需：青川山珍等特色产品展示与供需对接", { x: 0.9, y: 3.6999999999999997, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 政策推送：惠企政策精准匹配与主动推送", { x: 0.9, y: 4.1, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 数字化活动：园区文体、培训、交流活动管理", { x: 0.9, y: 4.5, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addText("• 餐饮服务：园区食堂菜单、在线订餐、商务用餐", { x: 0.9, y: 4.9, w: 8, h: 0.35, fontSize: 12, fontFace: "Microsoft YaHei", color: theme.light, valign: "middle" });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("11", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 12, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-11-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
