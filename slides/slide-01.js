const pptxgen = require("pptxgenjs");

const slideConfig = { type: 'cover', index: 1, title: '封面' };

function createSlide(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };

  // Decorative top bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: theme.accent } });

  // Left accent line
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.2, w: 0.08, h: 1.8, fill: { color: theme.accent } });

  // Main title
  slide.addText("青川县庄子上\n智慧工业园区平台", {
    x: 0.9, y: 1.2, w: 8, h: 2.0,
    fontSize: 36, fontFace: "Microsoft YaHei",
    color: "FFFFFF", bold: true, lineSpacingMultiple: 1.3
  });

  // Subtitle
  slide.addText("汇报介绍 · 智慧赋能产业升级", {
    x: 0.9, y: 3.2, w: 8, h: 0.5,
    fontSize: 16, fontFace: "Microsoft YaHei",
    color: theme.light
  });

  // Decorative elements
  slide.addShape(pres.shapes.OVAL, { x: 8.5, y: 0.5, w: 2, h: 2, fill: { color: theme.secondary, transparency: 70 } });
  slide.addShape(pres.shapes.OVAL, { x: -0.5, y: 4, w: 1.5, h: 1.5, fill: { color: theme.secondary, transparency: 70 } });

  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.2, w: 10, h: 0.425, fill: { color: theme.accent } });
  slide.addText("2026年5月", {
    x: 0.6, y: 5.2, w: 3, h: 0.425,
    fontSize: 11, fontFace: "Microsoft YaHei", color: "FFFFFF", valign: "middle"
  });

  return slide;
}

if (require.main === module) {
  const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9';
  const theme = { primary: "1B4332", secondary: "40916C", accent: "52B788", light: "D8F3DC", bg: "F5FBF7" };
  createSlide(pres, theme);
  pres.writeFile({ fileName: "slides/output/slide-01-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
