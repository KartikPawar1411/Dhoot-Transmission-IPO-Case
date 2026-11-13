// Writes a colour scheme and heading/body fonts into every theme part of a .pptx.
// pptxgenjs cannot set theme colours itself, so scheme colours used in the deck
// would otherwise resolve to Office's stock palette.
const fs = require("fs");
const JSZip = require("jszip");

const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

async function applyTheme(file, theme) {
  for (const k of SLOTS) {
    if (!/^[0-9A-Fa-f]{6}$/.test(theme.colors[k] || "")) throw new Error(`theme colour ${k} must be 6-digit hex`);
  }
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const parts = Object.keys(zip.files).filter((p) => /^ppt\/theme\/theme\d+\.xml$/.test(p));
  for (const p of parts) {
    let xml = await zip.file(p).async("string");
    xml = xml.replace(/<a:clrScheme name="[^"]*">/, `<a:clrScheme name="${esc(theme.name)}">`);
    for (const k of SLOTS) {
      xml = xml.replace(new RegExp(`<a:${k}>[\\s\\S]*?</a:${k}>`), `<a:${k}><a:srgbClr val="${theme.colors[k].toUpperCase()}"/></a:${k}>`);
    }
    xml = xml.replace(/(<a:majorFont>\s*<a:latin typeface=")[^"]*(")/, `$1${esc(theme.headFontFace)}$2`);
    xml = xml.replace(/(<a:minorFont>\s*<a:latin typeface=")[^"]*(")/, `$1${esc(theme.bodyFontFace)}$2`);
    zip.file(p, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

module.exports = { applyTheme };
