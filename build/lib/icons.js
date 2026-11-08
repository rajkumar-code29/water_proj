// Renders Lucide icons (via react-icons) to PNG data URIs for pptxgenjs.
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");

const cache = new Map();

async function icon(name, hex = "FFFFFF", px = 256) {
  const key = `${name}-${hex}-${px}`;
  if (cache.has(key)) return cache.get(key);
  const Comp = lu[name];
  if (!Comp) throw new Error(`Unknown icon ${name}`);
  let svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { size: px, color: `#${hex}` }));
  svg = svg.replace(/currentColor/g, `#${hex}`);
  if (!svg.includes("xmlns=")) svg = svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
  const buf = await sharp(Buffer.from(svg), { density: 300 }).resize(px, px).png().toBuffer();
  const uri = "image/png;base64," + buf.toString("base64");
  cache.set(key, uri);
  return uri;
}

module.exports = { icon };
