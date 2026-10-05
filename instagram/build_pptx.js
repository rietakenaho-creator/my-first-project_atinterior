// Instagram 固定投稿3件 × 4枚を1枚のスライドにまとめる
// 実行: NODE_PATH=<pptxgenjs の node_modules> APPLY_THEME=<apply_theme.js> node build_pptx.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME);

const OUT = path.join(__dirname, "Instagram固定投稿イメージ.pptx");
const img = (f) => path.join(__dirname, "out", f);

const THEME = {
  name: "HouseWork",
  headFontFace: "Meiryo",
  bodyFontFace: "Meiryo",
  colors: {
    dk1: "231F1C", lt1: "FFFFFF", dk2: "7B6A56", lt2: "F4F1ED",
    accent1: "F39800", accent2: "7B6A56", accent3: "231F1C",
    accent4: "C9BBA8", accent5: "D8CFC4", accent6: "06C755",
    hlink: "F39800", folHlink: "7B6A56",
  },
};

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "Instagram 固定投稿イメージ";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "HW_OVERVIEW",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.6,
        fontSize: 24, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
  });

  const s = pres.addSlide({ masterName: "HW_OVERVIEW" });
  s.addText("Instagram 固定投稿イメージ", { placeholder: "title" });

  const posts = [
    ["固定投稿 1", "ハウスワークの家 vol.1", "コンセプト・デザイン"],
    ["固定投稿 2", "ハウスワークの家 vol.2", "暮らし方の提案"],
    ["固定投稿 3", "モデルハウス見学会", "古河展示場・実例紹介"],
  ];
  const ih = 1.85, iw = ih * 4 / 5, gap = 0.15, x0 = 2.9, y0 = 1.15;
  posts.forEach(([no, name, sub], r) => {
    const y = y0 + r * (ih + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 2.2, h: ih, rectRadius: 0.1,
      fill: { color: C.background2 }, line: { type: "none" }, objectName: `row-${r + 1}-card` });
    s.addText([
      { text: no, options: { fontSize: 11, bold: true, color: C.accent1, breakLine: true } },
      { text: name, options: { fontSize: 15, bold: true, color: C.text1, breakLine: true } },
      { text: sub, options: { fontSize: 11, color: C.text2 } },
    ], { x: 0.7, y: y + 0.15, w: 1.85, h: ih - 0.3, valign: "middle", margin: 0, paraSpaceAfter: 4,
      isTextBox: true, objectName: `row-${r + 1}-label` });
    for (let i = 0; i < 4; i++) {
      s.addImage({ path: img(`p${r + 1}-${i + 1}.jpg`), x: x0 + i * (iw + gap), y, w: iw, h: ih,
        objectName: `post-${r + 1}-${i + 1}` });
    }
  });

  // 右側：プロフィール画面（1080 x 2129）
  const ph = 5.85, pw = ph * 1080 / 2129;
  s.addImage({ path: img("profile.jpg"), x: 12.83 - pw, y: y0, w: pw, h: ph,
    shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 6, offset: 2, angle: 90 }, objectName: "profile" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
