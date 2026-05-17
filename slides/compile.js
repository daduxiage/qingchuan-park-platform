const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

const theme = {
  primary: "1B4332",
  secondary: "40916C",
  accent: "52B788",
  light: "D8F3DC",
  bg: "F5FBF7"
};

// Helper to add page badge
for (let i = 1; i <= 17; i++) {
  const num = String(i).padStart(2, '0');
  const slideModule = require(`./slide-${num}.js`);
  slideModule.createSlide(pres, theme);
}

pres.writeFile({ fileName: './output/zhuangzishang_smart_park_report.pptx' })
  .then(() => console.log('PPTX generated: output/zhuangzishang_smart_park_report.pptx'))
  .catch(err => console.error('Error:', err));
