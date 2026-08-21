const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

const { FaRecycle, FaCarBattery, FaMapMarkerAlt, FaExclamationTriangle, FaChartBar, FaHeadset, FaUserCog, FaChartLine } = require('react-icons/fa');

const OUT_DIR = path.join(__dirname, 'icons');
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR);

async function makeIcon(Comp, name, color, size = 256) {
  let svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { size })
  );
  // react-icons uses fill="currentColor"; replace explicitly since librsvg
  // does not reliably resolve CSS currentColor from an inline style attr.
  svg = svg.replace(/currentColor/g, `#${color}`);
  if (!svg.includes('xmlns=')) {
    svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
  }
  const outPath = path.join(OUT_DIR, `${name}.png`);
  await sharp(Buffer.from(svg)).resize(size, size).png().toFile(outPath);
  console.log('wrote', outPath);
}

(async () => {
  await makeIcon(FaRecycle, 'recycle_white', 'FFFFFF');
  await makeIcon(FaRecycle, 'recycle_blue', '060E9F');
  await makeIcon(FaCarBattery, 'battery_white', 'FFFFFF');
  await makeIcon(FaCarBattery, 'battery_orange', 'FF5000');
  await makeIcon(FaMapMarkerAlt, 'pin_gray', '888888');
  await makeIcon(FaMapMarkerAlt, 'pin_blue', '060E9F');
  await makeIcon(FaExclamationTriangle, 'warn', '6B4800');
  await makeIcon(FaChartBar, 'chart_blue', '060E9F');
  await makeIcon(FaChartBar, 'chart_white', 'FFFFFF');
  await makeIcon(FaHeadset, 'headset_blue', '060E9F');
  await makeIcon(FaExclamationTriangle, 'warn_orange', 'FF5000');
  await makeIcon(FaMapMarkerAlt, 'pin_darkblue', '0076A9');
  await makeIcon(FaUserCog, 'usercog_gold', 'B8860B');
  await makeIcon(FaChartLine, 'chartline_white', 'FFFFFF');
})();
