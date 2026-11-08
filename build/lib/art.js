// SVG illustrations for the deck, rasterised to PNG with sharp.
// Text labels are NOT drawn here: they are added as native, editable PowerPoint text on top.
const sharp = require("sharp");

async function svgToPng(svg, vbWidth, targetPx) {
  const density = Math.max(72, Math.round((72 * targetPx) / vbWidth));
  const buf = await sharp(Buffer.from(svg), { density }).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------------------------------------------------------------- Lake (schematic, not to scale)
const LAKE = {
  vbW: 640,
  vbH: 600,
  path:
    "M 490 120 L 425 495 C 415 525, 380 535, 350 520 C 310 500, 270 490, 235 460 C 200 430, 195 400, 160 380 " +
    "C 120 355, 115 310, 135 275 C 150 245, 140 210, 170 180 C 200 150, 240 145, 270 125 " +
    "C 310 100, 360 90, 410 95 C 445 98, 475 105, 490 120 Z",
  // Key points in viewBox units (used to place native labels/overlays)
  pts: {
    buddha: [375, 340],
    kukatpallyIn: [295, 118],
    picketIn: [455, 108],
    balkapurIn: [128, 320],
    banjaraIn: [291, 494],
    outflow: [482, 185],
    sanjeevaiah: [168, 178],
    lumbini: [430, 500],
    ntr: [350, 522],
    tankBundMid: [455, 320],
  },
  // MLD (study basis) for arrow widths
  flows: { kukatpally: 70, balkapur: 13.3, banjara: 6, picket: 5.7 },
};

function arrow(x1, y1, x2, y2, width, color, opacity = 1) {
  const ang = Math.atan2(y2 - y1, x2 - x1);
  const head = Math.max(16, width * 1.9);
  const bx = x2 - Math.cos(ang) * head;
  const by = y2 - Math.sin(ang) * head;
  const hw = Math.max(11, width * 1.15);
  const px = -Math.sin(ang) * hw;
  const py = Math.cos(ang) * hw;
  return (
    `<line x1="${x1}" y1="${y1}" x2="${bx}" y2="${by}" stroke="#${color}" stroke-width="${width}" stroke-linecap="round" opacity="${opacity}"/>` +
    `<polygon points="${x2},${y2} ${bx + px},${by + py} ${bx - px},${by - py}" fill="#${color}" opacity="${opacity}"/>`
  );
}

function lakeSvg({ polluted = true } = {}) {
  const water = polluted
    ? `<linearGradient id="w" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#8FA552"/><stop offset="0.55" stop-color="#5E8A5E"/><stop offset="1" stop-color="#3F6F66"/></linearGradient>`
    : `<linearGradient id="w" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#7FDCD0"/><stop offset="0.55" stop-color="#2FAFAA"/><stop offset="1" stop-color="#0E7C86"/></linearGradient>`;
  const blotches = polluted
    ? [
        [230, 200, 60, 26], [180, 300, 40, 22], [330, 160, 50, 18], [300, 430, 55, 20], [400, 450, 30, 14], [250, 360, 34, 14],
      ]
        .map(([cx, cy, rx, ry]) => `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="#C9D65A" opacity="0.35"/>`)
        .join("")
    : [
        [260, 210, 70, 3], [330, 260, 90, 3], [220, 330, 60, 3], [300, 400, 80, 3], [380, 180, 50, 3],
      ]
        .map(([cx, cy, rx]) => `<path d="M ${cx - rx} ${cy} q ${rx / 4} -8 ${rx / 2} 0 t ${rx / 2} 0 t ${rx / 2} 0 t ${rx / 2} 0" stroke="#FFFFFF" stroke-width="3" fill="none" opacity="0.45"/>`)
        .join("");

  const inflowColor = polluted ? "7A5C2E" : "0E8A8F";
  const w = (mld) => 4 + mld * 0.24;
  const p = LAKE.pts;
  const inflows = polluted
    ? arrow(283, 8, p.kukatpallyIn[0], p.kukatpallyIn[1] + 6, w(70), inflowColor, 0.92) +
      arrow(505, 18, p.picketIn[0], p.picketIn[1] + 4, w(5.7), inflowColor, 0.92) +
      arrow(18, 322, p.balkapurIn[0] + 6, p.balkapurIn[1], w(13.3), inflowColor, 0.92) +
      arrow(225, 588, p.banjaraIn[0] - 2, p.banjaraIn[1] + 4, w(6), inflowColor, 0.92)
    : arrow(283, 8, p.kukatpallyIn[0], p.kukatpallyIn[1] + 6, 16, "36C2B4", 0.95);
  const outflow = arrow(p.outflow[0] + 4, p.outflow[1], 625, p.outflow[1], 8, polluted ? "6E7F55" : "36C2B4", 0.9);

  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${LAKE.vbW} ${LAKE.vbH}" width="${LAKE.vbW}" height="${LAKE.vbH}">
  <defs>${water}</defs>
  <path d="${LAKE.path}" fill="url(#w)" stroke="#0B3A44" stroke-width="3"/>
  <g clip-path="none">${blotches}</g>
  <line x1="490" y1="120" x2="425" y2="495" stroke="#5E7076" stroke-width="11" stroke-linecap="round"/>
  <line x1="490" y1="120" x2="425" y2="495" stroke="#FFFFFF" stroke-width="2" stroke-dasharray="10 9"/>
  ${inflows}
  ${outflow}
  <ellipse cx="${p.buddha[0]}" cy="${p.buddha[1] + 6}" rx="15" ry="8" fill="#8C8C7A" stroke="#0B3A44" stroke-width="2"/>
  <circle cx="${p.buddha[0]}" cy="${p.buddha[1] - 4}" r="7" fill="#FFFFFF" stroke="#0B3A44" stroke-width="2.5"/>
</svg>`;
}

// ---------------------------------------------------------------- Boat side elevation
function boatSideSvg({ dark = false } = {}) {
  const S = dark ? "#BFEDE8" : "#0B3A44"; // stroke
  const hull = dark ? "none" : "#F2A900";
  const frame = dark ? "none" : "#5E7076";
  const panel = dark ? "none" : "#17323A";
  const bin = dark ? "none" : "#0E8A8F";
  const conv = dark ? "none" : "#9DB0B4";
  const box = dark ? "none" : "#DDE7E6";
  const sw = dark ? 3 : 3;
  const waterFill = dark ? "#36C2B4" : "#36C2B4";
  const waterOpacity = dark ? 0.12 : 0.3;
  const cells = Array.from({ length: 6 }, (_, i) => {
    const x = 175 + (i + 1) * (590 / 6);
    const y1 = 112 - ((x - 175) / 590) * 12;
    return i < 5 ? `<line x1="${x}" y1="${y1}" x2="${x}" y2="${y1 + 14}" stroke="${dark ? S : "#6FA9B1"}" stroke-width="2"/>` : "";
  }).join("");
  const cleats = Array.from({ length: 9 }, (_, i) => {
    const t = (i + 1) / 10;
    const x = 950 + (715 - 950) * t;
    const y = 345 + (182 - 345) * t;
    return `<line x1="${x}" y1="${y}" x2="${x - 11}" y2="${y + 11}" stroke="${S}" stroke-width="2"/>`;
  }).join("");
  const sun = dark ? `<circle cx="905" cy="80" r="46" fill="#F2A900"/>` : "";
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 470" width="1000" height="470">
  ${sun}
  <!-- canopy posts + mast -->
  <line x1="205" y1="254" x2="205" y2="121" stroke="${S}" stroke-width="${sw + 2}"/>
  <line x1="455" y1="254" x2="455" y2="116" stroke="${S}" stroke-width="${sw + 2}"/>
  <line x1="700" y1="254" x2="700" y2="111" stroke="${S}" stroke-width="${sw + 2}"/>
  <line x1="775" y1="254" x2="775" y2="72" stroke="${S}" stroke-width="${sw + 1}"/>
  <!-- solar canopy -->
  <polygon points="175,112 765,100 765,114 175,126" fill="${panel}" stroke="${S}" stroke-width="${sw}"/>
  ${cells}
  <!-- mast head: GNSS dome, camera, nav light -->
  <circle cx="775" cy="60" r="13" fill="${dark ? "none" : "#FFFFFF"}" stroke="${S}" stroke-width="${sw}"/>
  <rect x="784" y="76" width="34" height="18" rx="3" fill="${dark ? "none" : "#10242A"}" stroke="${S}" stroke-width="${sw - 1}"/>
  <circle cx="775" cy="94" r="5" fill="#36C2B4"/>
  <!-- debris bin -->
  <rect x="545" y="176" width="160" height="78" rx="4" fill="${bin}" stroke="${S}" stroke-width="${sw}"/>
  <line x1="565" y1="186" x2="565" y2="246" stroke="${dark ? S : "#7FD3CB"}" stroke-width="2"/>
  <line x1="595" y1="186" x2="595" y2="246" stroke="${dark ? S : "#7FD3CB"}" stroke-width="2"/>
  <line x1="625" y1="186" x2="625" y2="246" stroke="${dark ? S : "#7FD3CB"}" stroke-width="2"/>
  <line x1="655" y1="186" x2="655" y2="246" stroke="${dark ? S : "#7FD3CB"}" stroke-width="2"/>
  <line x1="685" y1="186" x2="685" y2="246" stroke="${dark ? S : "#7FD3CB"}" stroke-width="2"/>
  <!-- electronics box -->
  <rect x="300" y="214" width="110" height="40" rx="4" fill="${box}" stroke="${S}" stroke-width="${sw}"/>
  <!-- conveyor -->
  <polygon points="950,345 935,360 700,197 715,182" fill="${conv}" stroke="${S}" stroke-width="${sw}"/>
  ${cleats}
  <!-- deck frame -->
  <rect x="160" y="254" width="700" height="14" fill="${frame}" stroke="${S}" stroke-width="${sw}"/>
  <!-- hull -->
  <path d="M150 268 L790 268 C850 268 885 285 895 305 C880 345 840 362 790 362 L185 362 C160 362 150 345 150 325 Z" fill="${hull}" stroke="${S}" stroke-width="${sw}"/>
  <!-- battery bay (dashed) with 4 packs -->
  <rect x="290" y="284" width="230" height="60" rx="4" fill="${dark ? "none" : "#FFFFFF"}" fill-opacity="0.55" stroke="${S}" stroke-width="2" stroke-dasharray="8 6"/>
  <rect x="302" y="294" width="46" height="40" rx="3" fill="${dark ? "none" : "#0B3A44"}" stroke="${S}" stroke-width="2"/>
  <rect x="356" y="294" width="46" height="40" rx="3" fill="${dark ? "none" : "#0B3A44"}" stroke="${S}" stroke-width="2"/>
  <rect x="410" y="294" width="46" height="40" rx="3" fill="${dark ? "none" : "#0B3A44"}" stroke="${S}" stroke-width="2"/>
  <rect x="464" y="294" width="46" height="40" rx="3" fill="${dark ? "none" : "#0B3A44"}" stroke="${S}" stroke-width="2"/>
  <!-- rim-drive thruster -->
  <line x1="215" y1="362" x2="215" y2="376" stroke="${S}" stroke-width="${sw + 2}"/>
  <circle cx="215" cy="394" r="19" fill="none" stroke="${S}" stroke-width="6"/>
  <circle cx="215" cy="394" r="5" fill="${S}"/>
  <!-- sonde -->
  <line x1="620" y1="362" x2="620" y2="420" stroke="${S}" stroke-width="2.5"/>
  <rect x="611" y="418" width="18" height="36" rx="5" fill="${dark ? "none" : "#0E8A8F"}" stroke="${S}" stroke-width="${sw - 1}"/>
  <!-- water -->
  <rect x="0" y="312" width="1000" height="158" fill="${waterFill}" opacity="${waterOpacity}"/>
  <path d="M0 312 q 25 -8 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0 t 50 0" stroke="${dark ? "#36C2B4" : "#0E8A8F"}" stroke-width="3" fill="none"/>
</svg>`;
}

// Points of interest on the side view (viewBox units) for numbered callouts
const BOAT_SIDE_PTS = {
  canopy: [330, 92],
  mast: [810, 50],
  bin: [625, 215],
  conveyor: [822, 236],
  electronics: [355, 234],
  battery: [405, 314],
  thruster: [258, 400],
  sonde: [655, 436],
  hull: [820, 330],
};

// ---------------------------------------------------------------- Boat top plan
function boatTopSvg() {
  const S = "#0B3A44";
  const pontoon = (y) =>
    `<path d="M150 ${y} L800 ${y} C860 ${y} 900 ${y + 18} 915 ${y + 27.5} C900 ${y + 37} 860 ${y + 55} 800 ${y + 55} L170 ${y + 55} C158 ${y + 55} 150 ${y + 47} 150 ${y + 35} L150 ${y + 20} C150 ${y + 8} 158 ${y} 170 ${y} Z" fill="#F2A900" stroke="${S}" stroke-width="3"/>`;
  const cleats = Array.from({ length: 11 }, (_, i) => `<line x1="${735 + i * 18}" y1="168" x2="${735 + i * 18}" y2="312" stroke="${S}" stroke-width="2"/>`).join("");
  const panels = [];
  for (let c = 0; c < 3; c++) for (let r = 0; r < 2; r++) panels.push(`<rect x="${185 + c * 193}" y="${40 + r * 200}" width="183" height="190" rx="4" fill="none" stroke="#6FA9B1" stroke-width="2" stroke-dasharray="6 6"/>`);
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 540" width="1000" height="540">
  <rect x="0" y="0" width="1000" height="540" fill="none"/>
  <!-- cross beams -->
  <rect x="222" y="90" width="20" height="300" fill="#5E7076" stroke="${S}" stroke-width="2"/>
  <rect x="445" y="90" width="20" height="300" fill="#5E7076" stroke="${S}" stroke-width="2"/>
  <rect x="690" y="90" width="20" height="300" fill="#5E7076" stroke="${S}" stroke-width="2"/>
  ${pontoon(40)}
  ${pontoon(385)}
  <!-- payload bay highlight -->
  <rect x="530" y="150" width="425" height="180" rx="10" fill="#F2A900" fill-opacity="0.12" stroke="#C98C00" stroke-width="3" stroke-dasharray="12 8"/>
  <!-- conveyor -->
  <rect x="725" y="165" width="210" height="150" rx="4" fill="#9DB0B4" stroke="${S}" stroke-width="3"/>
  ${cleats}
  <!-- bin -->
  <rect x="545" y="175" width="160" height="130" rx="4" fill="#0E8A8F" stroke="${S}" stroke-width="3"/>
  <!-- electronics -->
  <rect x="300" y="200" width="110" height="80" rx="4" fill="#DDE7E6" stroke="${S}" stroke-width="3"/>
  <!-- mast head -->
  <circle cx="775" cy="240" r="13" fill="#FFFFFF" stroke="${S}" stroke-width="3"/>
  <!-- thrusters -->
  <ellipse cx="136" cy="67" rx="12" ry="22" fill="none" stroke="${S}" stroke-width="5"/>
  <ellipse cx="136" cy="412" rx="12" ry="22" fill="none" stroke="${S}" stroke-width="5"/>
  <!-- solar canopy outline -->
  ${panels.join("")}
  <!-- dimension: LOA -->
  <line x1="124" y1="490" x2="915" y2="490" stroke="${S}" stroke-width="2"/>
  <line x1="124" y1="478" x2="124" y2="502" stroke="${S}" stroke-width="2"/>
  <line x1="915" y1="478" x2="915" y2="502" stroke="${S}" stroke-width="2"/>
  <!-- dimension: beam -->
  <line x1="70" y1="40" x2="70" y2="440" stroke="${S}" stroke-width="2"/>
  <line x1="58" y1="40" x2="82" y2="40" stroke="${S}" stroke-width="2"/>
  <line x1="58" y1="440" x2="82" y2="440" stroke="${S}" stroke-width="2"/>
  <!-- dimension: intake -->
  <line x1="965" y1="95" x2="965" y2="385" stroke="#C98C00" stroke-width="2"/>
  <line x1="953" y1="95" x2="977" y2="95" stroke="#C98C00" stroke-width="2"/>
  <line x1="953" y1="385" x2="977" y2="385" stroke="#C98C00" stroke-width="2"/>
</svg>`;
}

module.exports = {
  LAKE,
  BOAT_SIDE_PTS,
  lakePng: (opts) => svgToPng(lakeSvg(opts), LAKE.vbW, 1600),
  boatSidePng: (opts) => svgToPng(boatSideSvg(opts), 1000, 2400),
  boatTopPng: () => svgToPng(boatTopSvg(), 1000, 2000),
};
