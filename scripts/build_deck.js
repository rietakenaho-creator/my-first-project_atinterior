const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/5e082871-4813-41d1-9e93-3211e2eb9a88_0c465b0e-60d5-49f8-8580-157b5ec9e08f/pptx/scripts/apply_theme.js");

const OUT = "/home/user/my-first-project_atinterior/reform_strategy_2026.pptx";
const FONT = "Yu Gothic";
const HEX = {
  dk1: "1F2624", lt1: "FFFFFF", dk2: "26403A", lt2: "EEF2F0",
  accent1: "C0623B", accent2: "7E9C8A", accent3: "D9A441", accent4: "4F6D7A",
  accent5: "A8B9B0", accent6: "8A5A44", hlink: "4F6D7A", folHlink: "8A5A44",
};
const THEME = { name: "Reform Strategy", headFontFace: FONT, bodyFontFace: FONT, colors: HEX };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
pres.title = "2026 リフォーム市場分析と事業戦略";
const C = pres.SchemeColor;
const W = 13.333;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 11.7, h: 1.6, fontSize: 40, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 11.7, h: 1.2, fontSize: 20, color: C.accent5, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 11.7, h: 1.2, fontSize: 36, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.95, w: 11.7, h: 0.9, fontSize: 18, color: C.accent5, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.75, fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.1, w: 12.1, h: 0.5, fontSize: 15, color: C.accent1, valign: "top", margin: 0 }, text: "" } },
    { text: { text: "2026 リフォーム市場分析と事業戦略", options: { x: 0.6, y: 7.05, w: 8, h: 0.3, fontSize: 9, color: "7A8580", margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.3, fontSize: 9, color: "7A8580", align: "right" },
});

// ---------- helpers ----------
let objN = 0;
const nm = (p) => `${p}-${++objN}`;
function content(section, title, lead) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (lead) s.addText(lead, { placeholder: "body" });
  return s;
}
function card(s, x, y, w, h, fill) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 }, objectName: nm("card") });
}
function txt(s, text, o) {
  s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top", objectName: nm("text") }, o));
}
function stat(s, x, y, w, big, label, color) {
  card(s, x, y, w, 1.75);
  txt(s, big, { x: x + 0.25, y: y + 0.2, w: w - 0.5, h: 0.85, fontSize: big.length > 7 ? 28 : 34, bold: true, color: color || C.accent1, valign: "middle" });
  txt(s, label, { x: x + 0.25, y: y + 1.05, w: w - 0.5, h: 0.6, fontSize: 12, color: C.text1 });
}
function table(s, rows, o) {
  const head = rows[0].map((t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: o.hs || 12, valign: "middle" } }));
  const body = rows.slice(1).map((r, i) => r.map((t, j) => {
    const em = typeof t === "object";
    return { text: em ? t.t : t, options: { fontSize: o.fs || 12, color: HEX.dk1, bold: em || (o.boldFirst && j === 0), fill: { color: i % 2 ? HEX.lt2 : HEX.lt1 }, valign: "middle" } };
  }));
  s.addTable([head, ...body], Object.assign({ border: { type: "solid", pt: 0.5, color: "D5DCD8" }, margin: 0.06, objectName: nm("table") }, o.pos));
}
function source(s, t) {
  txt(s, t, { x: 0.6, y: 6.68, w: 12.1, h: 0.3, fontSize: 9, color: "6B7671" });
}
const chartBase = {
  catAxisLabelColor: "4A5450", valAxisLabelColor: "4A5450", catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  valGridLine: { color: "E3E8E5", size: 0.75 }, catGridLine: { style: "none" }, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
};

// ================= 1. Title =================
pres.addSection({ title: "表紙" });
let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "表紙" });
s.addText("2026 リフォーム市場分析と事業戦略", { placeholder: "title" });
s.addText("新築をあきらめた中間層に、“予算内で最適な住まい”を。\n水回りを入口に、コーディネートで選ばれる会社へ", { placeholder: "body" });
s.addNotes("本資料は公的統計・調査機関のデータと、それをもとにした試算で構成しています。試算値はスライド内に明記しています。");

// ================= 2. Summary =================
pres.addSection({ title: "サマリー" });
s = content("サマリー", "エグゼクティブサマリー", "新築は縮小、リフォームは単価で伸びる。勝負は「1件あたりの提案価値」");
stat(s, 0.6, 1.75, 2.9, "71.1万戸", "2025年度の新設住宅着工\n前年度比 −12.9%");
stat(s, 3.67, 1.75, 2.9, "7.7兆円", "2026年のリフォーム市場予測\n単価上昇で微増");
stat(s, 6.74, 1.75, 2.9, "約1.5倍", "リフォーム予算291万円→\n実際の費用434万円", C.accent4);
stat(s, 9.81, 1.75, 2.9, "約8割", "Z世代の「インテリア迷子」が\nAIでのコーディネートを希望", C.accent4);
card(s, 0.6, 3.85, 12.1, 2.6, C.text2);
txt(s, "結論", { x: 1.0, y: 4.05, w: 3, h: 0.4, fontSize: 14, bold: true, color: C.accent3 });
txt(s, [
  { text: "ターゲット：築10〜15年の持家（戸建て・マンション）と中古購入者", options: { bullet: true, breakLine: true } },
  { text: "入口：水回り4点（約200万円）で接点をつくる", options: { bullet: true, breakLine: true } },
  { text: "利益：コーディネート（内装・家具）と大型リノベで伸ばす", options: { bullet: true, breakLine: true } },
  { text: "差別化：AIで用意し、Riekoが「判断」する ―「買う前ジャッジ」", options: { bullet: true } },
], { x: 1.0, y: 4.5, w: 11.3, h: 1.8, fontSize: 17, color: C.background1, paraSpaceAfter: 6 });

// ================= Section: market =================
pres.addSection({ title: "市場" });
s = pres.addSlide({ masterName: "SECTION", sectionTitle: "市場" });
s.addText("1. 市場：新築からリフォームへ", { placeholder: "title" });
s.addText("新築着工・リフォーム市場・家計・地域の動き", { placeholder: "body" });

// 3. National starts
s = content("市場", "新築着工は2030年度でも2024年度の水準に戻らない", "2026年度以降の増加は反動減からの「揺り戻し」。持家は2040年度に14万戸へ");
s.addChart(pres.charts.BAR, [{ name: "着工戸数（万戸）", labels: ["2024", "2025", "2026", "2027", "2028", "2029", "2030"], values: [81.6, 71.1, 73, 75, 77, 78.5, 80] }], Object.assign({}, chartBase, {
  x: 0.6, y: 1.75, w: 8.2, h: 4.8, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.accent2], valAxisMinVal: 0, valAxisMaxVal: 90,
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 11, dataLabelColor: "1F2624", showLegend: false,
  showTitle: true, title: "新設住宅着工戸数（全国・年度、万戸）", titleFontSize: 12, titleColor: "4A5450",
  objectName: "chart-national-starts",
}));
card(s, 9.1, 1.75, 3.6, 4.8);
txt(s, [
  { text: "実績", options: { bold: true, color: C.accent1, breakLine: true } },
  { text: "2024年度 81.6万戸（駆け込み）→ 2025年度 71.1万戸", options: { breakLine: true } },
  { text: " ", options: { breakLine: true } },
  { text: "予測（野村総研）", options: { bold: true, color: C.accent1, breakLine: true } },
  { text: "2026年度 73万戸、2027年度 75万戸、2030年度 80万戸がピーク、2040年度 61万戸", options: { breakLine: true } },
  { text: " ", options: { breakLine: true } },
  { text: "示唆", options: { bold: true, color: C.accent1, breakLine: true } },
  { text: "新築を建てられない層が「今の家を直す」側に回る" },
], { x: 9.35, y: 1.95, w: 3.1, h: 4.4, fontSize: 13 });
source(s, "出典：国交省 住宅着工統計、野村総合研究所（2026年6月）。2028・2029年度は2027年度と2030年度の間を均等に補間した値。");

// 4. Reform market & orders
s = content("市場", "リフォーム市場は伸びるが、主因は単価の上昇", "件数は横ばい〜微減。売上は「客単価」と「家具の同時購入」で伸ばす");
stat(s, 0.6, 1.75, 3.0, "7.51兆円", "2025年 リフォーム市場\n前年比 +2.5%（矢野経済研究所）");
stat(s, 0.6, 3.7, 3.0, "747万件", "2025年度 住宅リフォーム受注件数（約）\n国交省 建築物リフォーム調査", C.accent4);
s.addChart(pres.charts.LINE, [
  { name: "強気（横ばい）", labels: ["2025", "2026", "2027", "2028", "2029", "2030"], values: [747, 747, 747, 747, 747, 747] },
  { name: "標準", labels: ["2025", "2026", "2027", "2028", "2029", "2030"], values: [747, 710, 703, 696, 689, 682] },
  { name: "弱気", labels: ["2025", "2026", "2027", "2028", "2029", "2030"], values: [747, 672, 659, 646, 633, 620] },
], Object.assign({}, chartBase, {
  x: 3.9, y: 1.75, w: 8.8, h: 4.8, chartColors: [HEX.accent2, HEX.dk2, HEX.accent1], lineSize: 2.5, lineDataSymbolSize: 7,
  valAxisMinVal: 580, valAxisMaxVal: 780, showLegend: true, legendPos: "b", legendFontSize: 11,
  showTitle: true, title: "住宅リフォーム受注件数の見通し（万件）― 2026年度以降は試算", titleFontSize: 12, titleColor: "4A5450",
  objectName: "chart-orders",
}));
source(s, "試算：標準＝2026年度 −5%（同年4〜6月の住宅受注高 −7.6%）、以降は年 −1%（市場 +1.5% − 単価 +2.5%）。747万件は二次情報で、国交省の原資料での確認を推奨。");

// 5. Households
s = content("市場", "家計：平均は上がっても、多くの世帯は豊かさを実感できない", "新築をあきらめた中間層が、最大のボリュームゾーンになる");
stat(s, 0.6, 1.75, 2.9, "451万円", "世帯所得の中央値\n（平均は575万円）");
stat(s, 3.67, 1.75, 2.9, "61.5%", "平均所得を下回る世帯の割合");
stat(s, 6.74, 1.75, 2.9, "−1.3%", "2025年の実質賃金\n4年連続のマイナス");
stat(s, 9.81, 1.75, 2.9, "−30%", "持家の新築着工（約）\n2021年比（2025年度）");
table(s, [
  ["世帯タイプ", "行動", "当社にとっての意味"],
  ["高所得層（年収1,000万円〜）", "新築と高額リノベの両方を検討", "デザイン性・提案力で選ばれる"],
  [{ t: "中間層（400〜800万円）" }, { t: "新築をあきらめ「今の家を直す」" }, { t: "最大のボリュームゾーン" }],
  ["低所得層・高齢世帯", "壊れたら直す（必要最低限）", "水回り・給湯器の交換需要"],
], { pos: { x: 0.6, y: 3.85, w: 12.1, colW: [3.6, 4.4, 4.1], rowH: 0.5 }, fs: 13, hs: 13, boldFirst: true });
source(s, "出典：厚労省 2025年国民生活基礎調査（2024年の所得）、毎月勤労統計、国交省 住宅着工統計");

// 6. Regional
s = content("市場", "地域：大都市圏は横ばいに近く、地方の減少が最大", "首都圏は「中古＋リノベ＋インテリア」、地方は「水回り・断熱の交換」");
s.addChart(pres.charts.BAR, [{ name: "前年比（%）", labels: ["近畿圏", "首都圏", "中部圏", "その他地域"], values: [-1.6, -5.9, -7.1, -9.2] }], Object.assign({}, chartBase, {
  x: 0.6, y: 1.75, w: 5.4, h: 4.8, barDir: "bar", chartColors: [HEX.accent4], valAxisMinVal: -12, valAxisMaxVal: 0, catAxisOrientation: "maxMin", catAxisLabelPos: "low",
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0\"%\"", dataLabelFontSize: 12, dataLabelColor: "1F2624", showLegend: false,
  showTitle: true, title: "2025年 新設住宅着工の前年比（全国 −6.5%）", titleFontSize: 12, titleColor: "4A5450",
  objectName: "chart-regional",
}));
table(s, [
  ["地域", "主なリフォーム需要", "打ち手"],
  ["首都圏", "中古マンション購入＋リノベ、水回り一式", "購入・リノベ・家具をまとめて提案"],
  ["近畿圏", "マンションの内装・水回り", "入居時オプションと家具"],
  ["中部圏", "戸建ての水回り・外装、二世帯化", "外装と水回りのセット"],
  ["地方", "給湯器・浴室、断熱窓、実家の改修", "補助金活用、定額パック"],
], { pos: { x: 6.3, y: 1.75, w: 6.4, colW: [1.2, 2.7, 2.5], rowH: 0.62 }, fs: 11.5, hs: 12, boldFirst: true });
source(s, "出典：国交省 住宅着工統計（2025年暦年）。その他地域は全国計から三大都市圏を差し引いて算出。需要と打ち手は仮説。");

// 7. SWOT
s = content("市場", "SWOT（2026年時点）", "強みは「内装・家具・設備をまとめて提案できる」こと");
const sw = [
  ["S 強み", ["インテリアとリフォームを一体で提案", "店舗で実物を見せられる", "家具・内装・設備をまとめて扱える"], C.accent2],
  ["W 弱み", ["職人・施工力に限りがある", "資材高で見積もりが上がりやすい", "提案から成約まで時間がかかる"], C.accent5],
  ["O 機会", ["新築離れで「直して住む」へ移行", "築20年超の水回り交換時期", "省エネ・断熱の補助金", "情報はあるが「選べない」人が多い"], C.accent3],
  ["T 脅威", ["実質賃金のマイナスで予算が抑えられる", "ホームセンター・ネット専業の低価格", "職人の高齢化・人手不足", "金利上昇"], C.accent1],
];
sw.forEach(([h, items, col], i) => {
  const x = 0.6 + (i % 2) * 6.15, y = 1.75 + Math.floor(i / 2) * 2.5;
  card(s, x, y, 5.95, 2.3);
  s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: y + 0.22, w: 0.5, h: 0.5, fill: { color: col }, line: { color: col }, objectName: nm("dot") });
  txt(s, h.slice(0, 1), { x: x + 0.25, y: y + 0.22, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(s, h.slice(2), { x: x + 0.9, y: y + 0.22, w: 4.5, h: 0.5, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
  txt(s, items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < items.length - 1 } })), { x: x + 0.35, y: y + 0.85, w: 5.4, h: 1.35, fontSize: 13, paraSpaceAfter: 2 });
});

// ================= Section: customers =================
pres.addSection({ title: "顧客と予算" });
s = pres.addSlide({ masterName: "SECTION", sectionTitle: "顧客と予算" });
s.addText("2. 顧客と予算：水回りが入口になる", { placeholder: "title" });
s.addText("中古購入の増加、水回りの交換サイクル、払える予算", { placeholder: "body" });

// 9. Used homes
s = content("顧客と予算", "中古住宅の購入は増え、購入時の約7割がリフォームする", "新築が高くなるほど「中古＋リフォーム」が選ばれる（特に首都圏）");
stat(s, 0.6, 1.75, 3.9, "4.9万件", "2025年 首都圏の中古マンション成約\n前年比 +31.9%（3年連続の増加）");
stat(s, 4.71, 1.75, 3.9, "2.2万件", "2025年 首都圏の中古戸建て成約\n前年比 +52.5%");
stat(s, 8.82, 1.75, 3.9, "約7割", "中古購入と同時に\n何らかのリフォームを実施", C.accent4);
card(s, 0.6, 3.75, 12.1, 2.7);
txt(s, "読み取れること", { x: 0.95, y: 3.95, w: 6, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
txt(s, [
  { text: "中古マンションを選んだ理由は「予算的に手頃だった」が75.7%", options: { bullet: true, breakLine: true } },
  { text: "全国の既存住宅流通量は2030年に19万戸へ増える見込み（野村総研）", options: { bullet: true, breakLine: true } },
  { text: "購入前から相談に乗れれば、水回り・内装・家具をまとめて受注できる", options: { bullet: true, breakLine: true } },
  { text: "不動産会社との提携が、第2ターゲットの入口になる", options: { bullet: true } },
], { x: 0.95, y: 4.45, w: 11.4, h: 1.9, fontSize: 15, paraSpaceAfter: 6 });
source(s, "出典：東日本レインズ（2025年）、国交省 令和6年度 住宅市場動向調査、野村総合研究所");

// 10. Funnel
s = content("顧客と予算", "水回りを入口にした受注は全国で年 約32万件", "そのうち約5万件が大型リノベに進む（全国・試算）");
const yrs = ["2024", "2025", "2026", "2027", "2028", "2029", "2030"];
s.addChart(pres.charts.BAR, [
  { name: "築10〜15年の持家 水回り", labels: yrs, values: [19.3, 19.8, 19.9, 19.8, 19.7, 19.3, 19.2] },
  { name: "中古購入時のリフォーム", labels: yrs, values: [11.5, 12.6, 12.8, 12.9, 13.1, 13.2, 13.3] },
], Object.assign({}, chartBase, {
  x: 0.6, y: 1.75, w: 8.2, h: 4.8, barDir: "col", barGrouping: "stacked", barGapWidthPct: 55, chartColors: [HEX.dk2, HEX.accent2],
  showValue: true, dataLabelPosition: "ctr", dataLabelFormatCode: "0.0", dataLabelFontSize: 11, dataLabelColor: "FFFFFF", valAxisMaxVal: 40,
  showLegend: true, legendPos: "b", legendFontSize: 11,
  showTitle: true, title: "入口の件数（万件/年）", titleFontSize: 12, titleColor: "4A5450", objectName: "chart-funnel",
}));
stat(s, 9.1, 1.75, 3.6, "約32万件", "入口（水回り＋中古購入時）\n市場規模 年 約6,500億円");
stat(s, 9.1, 3.7, 3.6, "約5万件", "大型リノベに進む件数（15%）\n市場規模 年 約2,300億円", C.accent4);
txt(s, "築10〜15年の持家は約330万戸。年6%が改修（5年間で持家の28.8%）", { x: 9.1, y: 5.6, w: 3.6, h: 0.9, fontSize: 11, color: "4A5450" });
source(s, "試算：住宅・土地統計調査2023、国交省 住宅市場動向調査、野村総研の既存住宅流通量予測をもとに算出。築年別の戸数は着工数の概数、転換率15%は仮定。");

// 11. Budget
s = content("顧客と予算", "水回りの予算は「約200万円」が妥当", "10年使う前提なら、年収500〜600万円の世帯でも月2万円以内");
table(s, [
  ["世帯年収", "手取り（概算）", "出せる上限（10年割で手取りの5%）", "月あたり"],
  ["500万円", "約390万円", { t: "約195万円" }, "約1.6万円"],
  ["600万円", "約460万円", { t: "約230万円" }, "約1.9万円"],
  ["700万円", "約530万円", "約265万円", "約2.2万円"],
  ["800万円", "約600万円", "約300万円", "約2.5万円"],
], { pos: { x: 0.6, y: 1.75, w: 7.3, colW: [1.6, 1.8, 2.5, 1.4], rowH: 0.55 }, fs: 13, hs: 12, boldFirst: true });
s.addChart(pres.charts.LINE, [
  { name: "水回り4点（ミドル）", labels: yrs, values: [190, 195, 200, 205, 210, 215, 220] },
  { name: "大型リノベ（平均）", labels: yrs, values: [434, 445, 456, 467, 479, 491, 503] },
], Object.assign({}, chartBase, {
  x: 8.2, y: 1.75, w: 4.5, h: 4.8, chartColors: [HEX.dk2, HEX.accent1], lineSize: 2.5, lineDataSymbolSize: 6, valAxisMinVal: 0, valAxisMaxVal: 550,
  showLegend: true, legendPos: "b", legendFontSize: 10, catAxisLabelFontSize: 9,
  showTitle: true, title: "1件あたりの予算（万円・試算）", titleFontSize: 11, titleColor: "4A5450", objectName: "chart-budget",
}));
card(s, 0.6, 4.75, 7.3, 1.75, C.text2);
txt(s, [
  { text: "水回り4点の相場は120〜250万円（ミドルグレードで約200万円）", options: { bullet: true, breakLine: true } },
  { text: "年収700万円以上なら、内装（床・壁）も加えて250〜300万円まで提案できる", options: { bullet: true } },
], { x: 0.9, y: 4.95, w: 6.8, h: 1.4, fontSize: 14, color: C.background1, paraSpaceAfter: 6 });
source(s, "試算：手取りの5%を上限と仮定。予算は2024年の相場（水回り190万円、リ推協の実施平均434万円）から年2.5%の単価上昇で延伸。");

// ================= Section: strategy =================
pres.addSection({ title: "戦略" });
s = pres.addSlide({ masterName: "SECTION", sectionTitle: "戦略" });
s.addText("3. 戦略：水回りで接点、コーディネートで利益", { placeholder: "title" });
s.addText("ターゲット、5年ロードマップ、利益モデル", { placeholder: "body" });

// 13. Targets
s = content("戦略", "ターゲットは「築10〜15年の持家」に絞る", "第2に中古購入者。高齢世帯は接点として使い、最安値層は追わない");
const tg = [
  ["◎", "最優先", "築10〜15年の持家", "40〜50代・年収600〜800万円\n全国 年約20万件\n水回り約200万円＋内装・家具", C.accent1],
  ["○", "第2", "中古購入者", "30〜40代・首都圏で増加\n全国 年約13万件\n1件400〜500万円＋家具", C.accent4],
  ["△", "維持", "築20年超・高齢世帯", "給湯器・トイレの交換\n20〜50万円\n集客の接点として使う", C.accent2],
  ["×", "対象外", "単品の最安値層", "ホームセンター・ネット専業との価格競争になる", C.accent5],
];
tg.forEach(([mk, lv, who, desc, col], i) => {
  const x = 0.6 + i * 3.08;
  card(s, x, 1.75, 2.9, 4.7);
  s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 1.95, w: 0.8, h: 0.8, fill: { color: col }, line: { color: col }, objectName: nm("mark") });
  txt(s, mk, { x: x + 0.2, y: 1.95, w: 0.8, h: 0.8, fontSize: 26, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(s, lv, { x: x + 1.15, y: 2.05, w: 1.6, h: 0.6, fontSize: 16, bold: true, color: C.accent1, valign: "middle" });
  txt(s, who, { x: x + 0.2, y: 3.0, w: 2.5, h: 0.8, fontSize: 18, bold: true, color: C.text2 });
  txt(s, desc, { x: x + 0.2, y: 3.85, w: 2.6, h: 2.4, fontSize: 12.5 });
});

// 14. Roadmap
s = content("戦略", "5年ロードマップ：入口 → 単価アップ → 利益の柱", "毎年1つずつサービスを積み上げる");
const rm = [
  ["2026", "水回り4点の定額パック", "150／200／250万円の3グレード＋無料のコーディネート診断", "水回り 60件"],
  ["2027", "コーディネートの商品化", "床・壁・照明・家具のセット提案", "追加率 40%"],
  ["2028", "中古購入＋リノベ", "不動産会社と提携し、購入前から見積もり", "中古リノベ 25件"],
  ["2029", "大型リノベ・断熱改修", "LDK・間取り変更、補助金の活用", "大型転換 18%"],
  ["2030", "住まい計画・リピート", "10年計画で外装・給湯器を次回提案", "リピート率 30%"],
];
s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.35, w: 11.5, h: 0, line: { color: HEX.accent5, width: 2 }, objectName: "timeline" });
rm.forEach(([y, t, d, k], i) => {
  const x = 0.6 + i * 2.46;
  s.addShape(pres.shapes.OVAL, { x: x + 0.85, y: 2.05, w: 0.6, h: 0.6, fill: { color: i === 0 ? HEX.accent1 : HEX.dk2 }, line: { color: HEX.lt1, width: 2 }, objectName: nm("node") });
  txt(s, y, { x, y: 1.6, w: 2.3, h: 0.4, fontSize: 15, bold: true, color: C.text2, align: "center" });
  card(s, x, 2.95, 2.3, 3.5);
  txt(s, t, { x: x + 0.15, y: 3.1, w: 2.0, h: 0.8, fontSize: 15, bold: true, color: C.accent1 });
  txt(s, d, { x: x + 0.15, y: 3.95, w: 2.0, h: 1.5, fontSize: 12 });
  txt(s, k, { x: x + 0.15, y: 5.7, w: 2.0, h: 0.5, fontSize: 13, bold: true, color: C.text2 });
});

// 15. Profit model
s = content("戦略", "利益は水回りの「後」で出す", "同じ水回り100件でも、コーディネートと大型リノベで粗利は約1.6倍");
table(s, [
  ["段階", "単価", "粗利率", "1件の粗利"],
  ["接点（給湯器・トイレ）", "20〜50万円", "15〜20%", "4〜10万円"],
  ["入口（水回り4点）", "約200万円", "28%", "約56万円"],
  [{ t: "追加（床・壁・照明・家具）" }, "約80万円", { t: "33%" }, "約26万円"],
  ["大型リノベ", "約470万円", "30%", "約141万円"],
], { pos: { x: 0.6, y: 1.75, w: 6.6, colW: [2.6, 1.4, 1.1, 1.5], rowH: 0.55 }, fs: 13, hs: 12, boldFirst: true });
s.addChart(pres.charts.BAR, [
  { name: "水回り", labels: ["水回りだけ", "コーディネート＋大型"], values: [5600, 5600] },
  { name: "内装・家具の追加", labels: ["水回りだけ", "コーディネート＋大型"], values: [0, 1056] },
  { name: "大型リノベ", labels: ["水回りだけ", "コーディネート＋大型"], values: [0, 2115] },
], Object.assign({}, chartBase, {
  x: 7.5, y: 1.75, w: 5.2, h: 4.8, barDir: "col", barGrouping: "stacked", barGapWidthPct: 70, chartColors: [HEX.dk2, HEX.accent3, HEX.accent1],
  showLegend: true, legendPos: "b", legendFontSize: 10, valAxisMaxVal: 10000,
  showTitle: true, title: "水回り100件あたりの粗利（万円）", titleFontSize: 12, titleColor: "4A5450", objectName: "chart-profit",
}));
card(s, 0.6, 4.85, 6.6, 1.65, C.text2);
txt(s, [
  { text: "約8,770万円", options: { fontSize: 30, bold: true, color: C.accent3, breakLine: true } },
  { text: "粗利（水回りだけなら5,600万円）。売上 約3.0億円・粗利率29%", options: { fontSize: 13, color: C.background1 } },
], { x: 0.9, y: 5.0, w: 6.1, h: 1.4 });
source(s, "試算：内装・家具の追加率40%、大型への転換率15%を仮定。粗利率は業界平均（約30%、中小15〜25%）をもとに設定。");

// 16. Store plan
s = content("戦略", "1店舗の5年計画（例）：2030年に売上7.9億円・粗利2.3億円", "件数・追加率・転換率を毎年少しずつ上げる");
s.addChart(pres.charts.BAR, [
  { name: "売上（億円）", labels: ["2026", "2027", "2028", "2029", "2030"], values: [1.62, 2.88, 4.57, 6.33, 7.87] },
  { name: "粗利（億円）", labels: ["2026", "2027", "2028", "2029", "2030"], values: [0.47, 0.84, 1.34, 1.87, 2.33] },
], Object.assign({}, chartBase, {
  x: 0.6, y: 1.75, w: 6.6, h: 4.8, barDir: "col", barGapWidthPct: 50, chartColors: [HEX.accent2, HEX.accent1],
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", dataLabelFontSize: 10, dataLabelColor: "1F2624",
  showLegend: true, legendPos: "b", legendFontSize: 11, valAxisMaxVal: 9,
  showTitle: true, title: "売上と粗利（億円）", titleFontSize: 12, titleColor: "4A5450", objectName: "chart-store",
}));
table(s, [
  ["年", "水回り", "追加率", "大型転換", "中古リノベ"],
  ["2026", "60件", "30%", "10%", "0件"],
  ["2027", "80件", "40%", "12%", "10件"],
  ["2028", "100件", "50%", "15%", "25件"],
  ["2029", "120件", "55%", "18%", "40件"],
  ["2030", "140件", "60%", "20%", "50件"],
], { pos: { x: 7.5, y: 1.75, w: 5.2, colW: [0.8, 1.1, 1.0, 1.1, 1.2], rowH: 0.5 }, fs: 13, hs: 12, boldFirst: true });
txt(s, "単価：水回り約200万円・追加約80万円・大型約456万円・中古リノベ約500万円（年2.5%上昇）。件数・率は目標値の例。年140件には施工体制の拡大が必要。", { x: 7.5, y: 5.0, w: 5.2, h: 1.4, fontSize: 11, color: "4A5450" });

// ================= Section: coordination =================
pres.addSection({ title: "コーディネート事業" });
s = pres.addSlide({ masterName: "SECTION", sectionTitle: "コーディネート事業" });
s.addText("4. コーディネート事業：「判断」を売る", { placeholder: "title" });
s.addText("AIが選択肢を用意し、プロが判断する", { placeholder: "body" });

// 18. Insight
s = content("コーディネート事業", "選択肢はAIで手に入る。足りないのは「判断」と「安心」", "AIが普及するほど、お金を払う対象は「案づくり」から「お墨付き」へ移る");
const steps = [
  ["情報を集める", "SNS・AIで画像も商品も大量に見られる", true],
  ["案をつくる", "AIで部屋のイメージも家具の候補も出せる", true],
  ["選ぶ・決める", "どれが自分の部屋に合うか判断できない", false],
  ["失敗を避ける", "15万円のソファは返品しにくい", false],
  ["確信を持つ", "「プロが良いと言った」後押しが欲しい", false],
];
steps.forEach(([h, d, ai], i) => {
  const x = 0.6 + i * 2.46;
  card(s, x, 1.85, 2.3, 2.9, ai ? C.background2 : C.text2);
  txt(s, ai ? "AIでできる" : "AIではできない", { x: x + 0.15, y: 2.0, w: 2.0, h: 0.4, fontSize: 12, bold: true, color: ai ? C.accent2 : C.accent3 });
  txt(s, h, { x: x + 0.15, y: 2.45, w: 2.0, h: 0.6, fontSize: 18, bold: true, color: ai ? C.text2 : C.background1 });
  txt(s, d, { x: x + 0.15, y: 3.15, w: 2.0, h: 1.4, fontSize: 12.5, color: ai ? C.text1 : C.background1 });
});
stat(s, 0.6, 4.95, 3.9, "半数以上", "今のインテリアに不満\n理由の4割超が「色やテイストの不統一」");
stat(s, 4.71, 4.95, 3.9, "7割以上", "インテリアにこだわりたい", C.accent4);
stat(s, 8.82, 4.95, 3.9, "約8割", "Z世代の「インテリア迷子」が\nAIでのコーディネートを希望", C.accent4);
txt(s, "出典：PR TIMES 住まいのインテリア意識調査、Z世代ほか1,000人調査（2026年）", { x: 0.6, y: 6.75, w: 12.1, h: 0.25, fontSize: 9, color: "6B7671" });

// 19. Judge service
s = content("コーディネート事業", "サービス案「買う前ジャッジ」", "その家具、あなたの部屋に合うか。買う前にプロが判定します。");
// mock card
card(s, 0.6, 1.75, 5.2, 4.75, C.text2);
txt(s, "判定カード（届くもの）", { x: 0.9, y: 1.9, w: 4.6, h: 0.4, fontSize: 13, bold: true, color: C.accent3 });
txt(s, "○", { x: 0.9, y: 2.35, w: 1.1, h: 1.1, fontSize: 60, bold: true, color: C.background1, valign: "middle" });
txt(s, "色を変えればOK", { x: 2.0, y: 2.55, w: 3.6, h: 0.7, fontSize: 20, bold: true, color: C.background1, valign: "middle" });
txt(s, [
  { text: "色：床の木目と同系色で沈む → 一段明るいグレーへ", options: { bullet: true, breakLine: true } },
  { text: "サイズ：幅180cmで動線もOK", options: { bullet: true, breakLine: true } },
  { text: "テイスト：ナチュラルに合う", options: { bullet: true, breakLine: true } },
  { text: "代わりの候補3点（価格・リンク付き）", options: { bullet: true, breakLine: true } },
  { text: "置き方：窓に背を向けて配置", options: { bullet: true } },
], { x: 0.9, y: 3.6, w: 4.7, h: 2.7, fontSize: 13, color: C.background1, paraSpaceAfter: 5 });
const dz = [
  ["LINEで送るだけ", "部屋の写真2枚・寸法・迷っている家具。面談なし"],
  ["答えは1枚のカード", "◎○△×と理由3行。48時間以内に返信"],
  ["1点3,300円から", "部屋まるごと9,800円。家具の失敗より安い"],
  ["購入時に全額充当", "提案した家具を買えば実質0円"],
  ["合わなければ再判定無料", "判定どおりに買って合わなければ次回無料"],
];
dz.forEach(([h, d], i) => {
  const y = 1.75 + i * 0.97;
  s.addShape(pres.shapes.OVAL, { x: 6.1, y: y + 0.12, w: 0.55, h: 0.55, fill: { color: HEX.accent1 }, line: { color: HEX.accent1 }, objectName: nm("num") });
  txt(s, String(i + 1), { x: 6.1, y: y + 0.12, w: 0.55, h: 0.55, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
  txt(s, h, { x: 6.85, y: y + 0.05, w: 5.8, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
  txt(s, d, { x: 6.85, y: y + 0.45, w: 5.8, h: 0.4, fontSize: 13 });
});

// 20. Monosashi + AI
s = content("コーディネート事業", "Riekoのものさし × AIで、工数を約3分の1に", "AIはRiekoが選んだ商品の中からしか選ばない ＝ センスが保たれる");
const ms = [["色の比率", "70：25：5"], ["素材の数", "3種類まで"], ["木の色", "2トーンまで"], ["高さ", "低・中・高を三角形に"], ["余白", "床の3割を空ける"]];
txt(s, "Riekoのものさし（例）", { x: 0.6, y: 1.75, w: 5, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
ms.forEach(([h, d], i) => {
  const y = 2.25 + i * 0.82;
  card(s, 0.6, y, 5.3, 0.68);
  txt(s, h, { x: 0.85, y, w: 1.8, h: 0.68, fontSize: 14, bold: true, color: C.text2, valign: "middle" });
  txt(s, d, { x: 2.7, y, w: 3.1, h: 0.68, fontSize: 14, valign: "middle" });
});
s.addChart(pres.charts.BAR, [
  { name: "現状（時間）", labels: ["イメージ作成", "家具の選定", "リスト・見積もり", "修正対応", "事前ヒアリング", "面談"], values: [3.0, 2.0, 1.0, 1.0, 0.5, 1.0] },
  { name: "AI活用後（時間）", labels: ["イメージ作成", "家具の選定", "リスト・見積もり", "修正対応", "事前ヒアリング", "面談"], values: [0.5, 0.5, 0.17, 0.5, 0, 1.0] },
], Object.assign({}, chartBase, {
  x: 6.3, y: 1.75, w: 6.4, h: 3.7, barDir: "bar", barGapWidthPct: 40, chartColors: [HEX.accent5, HEX.accent1], catAxisOrientation: "maxMin",
  showLegend: true, legendPos: "b", legendFontSize: 10, catAxisLabelFontSize: 10,
  showTitle: true, title: "1件あたりの工数（時間）", titleFontSize: 12, titleColor: "4A5450", objectName: "chart-hours",
}));
card(s, 6.3, 5.6, 6.4, 0.95, C.text2);
txt(s, [
  { text: "約8.5時間 → 約2.7時間", options: { bold: true, color: C.accent3, fontSize: 20 } },
  { text: "　面談と判断は人が担う", options: { color: C.background1, fontSize: 13 } },
], { x: 6.55, y: 5.6, w: 6.0, h: 0.95, valign: "middle" });
txt(s, "現状の工数は一般的な流れからの見積もり。", { x: 0.6, y: 6.45, w: 5.3, h: 0.3, fontSize: 10, color: "6B7671" });

// 21. Segments
s = content("コーディネート事業", "ターゲット別の商品：男性向け・女性向け・50歳から", "共通する本音は「何がいまいちなのか、自分ではわからない」");
const seg = [
  ["50歳から", "これからの20年の部屋", "これからを心地よく。でも大きなお金はかけたくない", "家具を「残す・直す・手放す」で判定／空いた子ども部屋を書斎に／訪問あり", "22,000円 → 5万円", "「減らして、ととのえる。」", "優先 1"],
  ["男性向け", "部屋の正解パッケージ", "正解を知りたい。でも聞くのは恥ずかしい。時間はかけたくない", "理由を数字で示す判定／まとめて買える購入リスト／組み立て代行", "3,300円 → 29,800円", "「センスはいらない。正解だけ。」", "優先 2"],
  ["女性向け", "好きをまとめるジャッジ", "好きな画像は保存しているのに、家ではまとまらない", "保存画像からテイストを読み解く／今ある家具を活かす／夫婦の折衷案", "9,800円 → 3万円", "「まとめ方を教えます。」", "優先 3"],
];
seg.forEach(([who, name, honne, what, price, copy, pr], i) => {
  const x = 0.6 + i * 4.1;
  card(s, x, 1.75, 3.9, 4.75);
  txt(s, pr, { x: x + 2.6, y: 1.9, w: 1.1, h: 0.35, fontSize: 11, bold: true, color: C.accent1, align: "right" });
  txt(s, who, { x: x + 0.25, y: 1.9, w: 2.3, h: 0.35, fontSize: 13, bold: true, color: C.accent4 });
  txt(s, name, { x: x + 0.25, y: 2.3, w: 3.45, h: 0.5, fontSize: 18, bold: true, color: C.text2 });
  txt(s, "本音：" + honne, { x: x + 0.25, y: 2.9, w: 3.45, h: 0.9, fontSize: 12, italic: true, color: "4A5450" });
  txt(s, what, { x: x + 0.25, y: 3.85, w: 3.45, h: 1.3, fontSize: 12.5 });
  txt(s, price, { x: x + 0.25, y: 5.2, w: 3.45, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
  txt(s, copy, { x: x + 0.25, y: 5.65, w: 3.45, h: 0.7, fontSize: 13, bold: true, color: C.text2 });
});

// 22. Monthly scenario
s = content("コーディネート事業", "月300万円は「判定＋プラン＋家具販売」で届く", "コーディネート料だけでは月49件が必要。家具販売と組み合わせて件数を抑える");
table(s, [
  ["メニュー", "料金", "月の件数", "収入", "工数"],
  ["買う前ジャッジ（部屋まるごと）", "9,800円", "40件", "約39万円", "20時間"],
  ["ライト（1部屋）", "3万円", "10件", "30万円", "20時間"],
  ["スタンダード（LDK）", "8万円", "8件", "64万円", "28時間"],
  ["プレミアム（複数の部屋・訪問）", "20万円", "2件", "40万円", "16時間"],
  ["家具販売の粗利", "購入18件×40万円×30%", "―", "約216万円", "―"],
  [{ t: "合計" }, "", { t: "60件" }, { t: "約390万円" }, { t: "84時間" }],
], { pos: { x: 0.6, y: 1.75, w: 8.0, colW: [2.9, 1.9, 1.0, 1.2, 1.0], rowH: 0.56 }, fs: 12.5, hs: 12, boldFirst: true });
stat(s, 8.9, 1.75, 3.8, "約390万円", "月の収益（コーディネート料＋家具粗利）\n購入率が半分でも約340万円");
stat(s, 8.9, 3.7, 3.8, "84時間", "Riekoの月の作業時間\nAIなしなら同規模で200時間超", C.accent4);
source(s, "試算：家具購入は診断の20%・プランの50%と仮定。収益＝コーディネート料＋家具販売の粗利。");

// 23. Closing
pres.addSection({ title: "まとめ" });
s = pres.addSlide({ masterName: "TITLE", sectionTitle: "まとめ" });
s.addText("水回りで出会い、判断で選ばれ、\n住まい全体で長く付き合う", { placeholder: "title" });
s.addText("次の一歩：① 自社実績で試算を置き換える　② 「Riekoのものさし」と商品マスターをつくる　③ 買う前ジャッジを月10件で試す", { placeholder: "body" });

// 24. Sources
s = content("まとめ", "出典", null);
txt(s, [
  "矢野経済研究所「住宅リフォーム市場に関する調査（2026年）」",
  "野村総合研究所「2026〜2040年度の新設住宅着工戸数・リフォーム市場規模予測」、既存住宅流通量予測",
  "国土交通省「住宅着工統計」「建築物リフォーム・リニューアル調査」「令和6年度 住宅市場動向調査」",
  "厚生労働省「2025年 国民生活基礎調査」「毎月勤労統計調査」",
  "総務省「令和5年 住宅・土地統計調査」",
  "住宅リフォーム推進協議会「住宅リフォームに関する消費者実態調査」",
  "東日本不動産流通機構（東日本レインズ）「首都圏不動産流通市場の動向（2025年）」",
  "建設経済研究所「建設経済モデルによる建設投資の見通し」",
  "PR TIMES「住まいのインテリアに関する意識調査」「Z世代の部屋づくりとインテリアに関する実態調査」",
  "LIFULL HOME'S（TSMC・ラピダス周辺の賃貸市場）、ミツモア（コーディネーター料金相場）、ANDPAD（リフォーム会社の利益率）",
].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })), { x: 0.6, y: 1.4, w: 12.1, h: 4.8, fontSize: 13, paraSpaceAfter: 5 });
txt(s, "※「試算」と記載した数値は、上記データをもとにした仮説値です。自社の実績で検証してください。", { x: 0.6, y: 6.3, w: 12.1, h: 0.4, fontSize: 12, color: C.accent1 });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
