// Writes the deck's own colour scheme into ppt/theme/theme1.xml after pptxgenjs has saved the file.
// pptxgenjs only sets theme fonts, so without this step every scheme colour falls back to Office's palette.
// Also fixes two pptxgenjs quirks that make PowerPoint offer to "repair" a file.
const fs = require("fs");
const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

function escapeAttr(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

async function applyTheme(deckPath, theme) {
  for (const k of SLOTS) {
    if (!/^[0-9A-F]{6}$/i.test(theme.colors[k] || "")) throw new Error(`theme.colors.${k} must be 6 hex digits`);
  }
  const zip = await JSZip.loadAsync(fs.readFileSync(deckPath));
  const themePart = "ppt/theme/theme1.xml";
  const name = escapeAttr(theme.name);
  const scheme =
    `<a:clrScheme name="${name}">` +
    SLOTS.map((k) => `<a:${k}><a:srgbClr val="${theme.colors[k].toUpperCase()}"/></a:${k}>`).join("") +
    `</a:clrScheme>`;
  let xml = await zip.file(themePart).async("string");
  xml = xml
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${name}"`);
  zip.file(themePart, xml);

  for (const partName of Object.keys(zip.files)) {
    if (!partName.endsWith(".xml")) continue;
    let part = await zip.file(partName).async("string");
    const bad = part.match(/<a:srgbClr val="((?![0-9A-Fa-f]{6}")[^"]*)"/);
    if (bad) throw new Error(`${partName}: srgbClr "${bad[1]}" is not hex. A scheme colour was passed to a hex-only option.`);
    if (part.includes('<c:crossesAt val="autoZero"/>')) {
      part = part.split('<c:crossesAt val="autoZero"/>').join('<c:crosses val="autoZero"/>');
      zip.file(partName, part);
    }
  }
  // Folder entries are not OPC parts; drop them.
  for (const n of Object.keys(zip.files)) if (zip.files[n].dir) delete zip.files[n];
  fs.writeFileSync(deckPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

module.exports = { applyTheme };
