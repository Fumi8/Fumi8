// 共通：テーマ・アイコン・テキストヘルパー
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const THEME = {
  name: "Sankei Ink",
  headFontFace: "Meiryo",
  bodyFontFace: "Meiryo",
  colors: {
    dk1: "1F2933", // 墨
    lt1: "FFFFFF",
    dk2: "2D4059", // 藍
    lt2: "F2F4F7", // 淡いグレー
    accent1: "D9482B", // 朱（赤ペン）
    accent2: "2D4059",
    accent3: "6B7785", // グレー
    accent4: "E3A13B", // 黄土
    accent5: "3C8D7F", // 青緑
    accent6: "C9D1DB",
    hlink: "2D4059",
    folHlink: "6B7785",
  },
};
const HEX = THEME.colors;

const iconCache = {};
async function icon(name, color) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const Comp = fa[name];
  if (!Comp) throw new Error("icon not found: " + name);
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: 256 })
  );
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  iconCache[key] = "image/png;base64," + buf.toString("base64");
  return iconCache[key];
}

// テキストボックス（行間を pt 固定にして、環境によるフォント差を抑える）
function txt(slide, text, o) {
  const size = o.fontSize || 14;
  const opts = Object.assign(
    {
      isTextBox: true,
      fontSize: size,
      margin: 0.04,
      valign: "top",
      lineSpacing: Math.round(size * 1.4),
      fit: "none",
    },
    o
  );
  slide.addText(text, opts);
}

module.exports = { THEME, HEX, icon, txt };
