const pptxgen = require("pptxgenjs");
const slideConfig = { type: 'end', index: 17, title: '感谢聆听' };
function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });
  slide.addText("感谢聆听", { x: 0.5, y: 1.5, w: 9, h: 1.0, fontSize: 36, fontFace: "Microsoft YaHei", color: "FFFFFF", bold: true, align: "center" });
  slide.addText("感谢聆听", { x: 0.5, y: 2.5, w: 9, h: 0.4, fontSize: 14, fontFace: "Microsoft YaHei", color: theme.light, align: "center" });
  slide.addText("青川县庄子上工业园区管委会", { x: 0.5, y: 3.2, w: 9, h: 0.4, fontSize: 13, fontFace: "Microsoft YaHei", color: theme.light, align: "center" });
  slide.addText("技术支持：青迅科技", { x: 0.5, y: 3.6500000000000004, w: 9, h: 0.4, fontSize: 13, fontFace: "Microsoft YaHei", color: theme.light, align: "center" });
  slide.addText("联系电话：0839-xxxxxxx", { x: 0.5, y: 4.1000000000000005, w: 9, h: 0.4, fontSize: 13, fontFace: "Microsoft YaHei", color: theme.light, align: "center" });
  slide.addText("邮箱：park@zhuangzishang.gov.cn", { x: 0.5, y: 4.550000000000001, w: 9, h: 0.4, fontSize: 13, fontFace: "Microsoft YaHei", color: theme.light, align: "center" });
  slide.addShape(pres.shapes.OVAL, { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fill: { color: theme.accent } });
  slide.addText("END", { x: 9.3, y: 5.1, w: 0.4, h: 0.4, fontSize: 10, fontFace: "Arial", color: "FFFFFF", bold: true, align: "center", valign: "middle" });
  return slide;
}
if (require.main === module) { const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9'; const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" }; createSlide(pres, theme); pres.writeFile({ fileName: "slides/output/slide-17-preview.pptx" }); }
module.exports = { createSlide, slideConfig };
