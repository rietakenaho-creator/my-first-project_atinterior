// LINE公式アカウント 開設手順書（株式会社ハウスワーク様向け）を生成する
// 実行: NODE_PATH=<pptxgenjs等のnode_modules> node build.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");
const { applyTheme } = require(process.env.APPLY_THEME);

const OUT = path.join(__dirname, "LINE公式アカウント開設手順書.pptx");
const LOGO = path.join(__dirname, "logo.png");
const RICHMENU = path.join(__dirname, "../line-richmenu/richmenu.jpg");

const HEX = {
  black: "231F1C", orange: "F39800", taupe: "7B6A56", white: "FFFFFF",
  tint: "F4F1ED", line: "D8CFC4", sand: "C9BBA8", green: "06C755",
};
const THEME = {
  name: "HouseWork",
  headFontFace: "Meiryo",
  bodyFontFace: "Meiryo",
  colors: {
    dk1: HEX.black, lt1: HEX.white, dk2: HEX.taupe, lt2: HEX.tint,
    accent1: HEX.orange, accent2: HEX.taupe, accent3: HEX.black,
    accent4: HEX.sand, accent5: HEX.line, accent6: HEX.green,
    hlink: HEX.orange, folHlink: HEX.taupe,
  },
};

async function icon(name, hex, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(fa[name], { color: "#" + hex, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "LINE公式アカウント 開設手順書";
  pres.author = "株式会社ハウスワーク";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  // ---------- レイアウト ----------
  pres.defineSlideMaster({
    title: "HW_DARK",
    background: { color: C.text1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.1, w: 7.6, h: 1.9,
        fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.25, w: 7.6, h: 1.3,
        fontSize: 18, color: C.accent4, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "HW_CONTENT",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 10.2, h: 0.8,
        fontSize: 28, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "LINE公式アカウント 開設手順書", options: { x: 0.6, y: 6.95, w: 6, h: 0.3,
        fontSize: 10, color: C.text2, margin: 0 } } },
      { image: { path: LOGO, x: 11.85, y: 6.8, w: 0.83, h: 0.55 } },
    ],
    slideNumber: { x: 11.1, y: 6.95, w: 0.6, h: 0.3, fontSize: 10, color: C.text2 },
  });

  const ICON = {};
  for (const [k, n, c] of [
    ["mail", "FaEnvelope", HEX.orange], ["key", "FaKey", HEX.orange], ["building", "FaBuilding", HEX.orange],
    ["phone", "FaMobileScreen", HEX.orange], ["image", "FaImage", HEX.orange], ["laptop", "FaLaptop", HEX.orange],
    ["globe", "FaGlobe", HEX.white], ["at", "FaAt", HEX.white], ["id", "FaIdCard", HEX.white],
    ["pen", "FaPenToSquare", HEX.white], ["check", "FaCircleCheck", HEX.white], ["users", "FaUserPlus", HEX.white],
    ["bulb", "FaLightbulb", HEX.orange], ["checkG", "FaCircleCheck", HEX.green], ["clock", "FaClock", HEX.sand],
    ["user", "FaCircleUser", HEX.orange], ["comment", "FaCommentDots", HEX.orange], ["grid", "FaTableCellsLarge", HEX.orange],
    ["shield", "FaShieldHalved", HEX.orange], ["q", "FaCircleQuestion", HEX.orange],
  ]) ICON[k] = await icon(n, c);

  // ---------- 共通パーツ ----------
  const addTitle = (s, step, text) =>
    s.addText(step ? [{ text: step + "　", options: { color: C.accent1 } }, { text }] : text, { placeholder: "title" });

  // 当日一緒に進めるためのチェック欄
  const addCheck = (s) =>
    s.addText("☐ 完了", { x: 11.45, y: 0.5, w: 1.25, h: 0.55, fontSize: 14, bold: true, color: C.text1,
      align: "center", valign: "middle", shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1,
      fill: { color: C.background2 }, line: { color: C.accent5, width: 1 }, isTextBox: true, objectName: "check" });

  const addSteps = (s, items, x = 0.6, y = 1.6, w = 5.9, gap = 0.95) =>
    items.forEach((t, i) => {
      const yy = y + i * gap;
      s.addText(String(i + 1), { x, y: yy, w: 0.48, h: 0.48, shape: pres.shapes.OVAL, fill: { color: C.accent1 },
        color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0,
        isTextBox: true, objectName: `step-${i + 1}-num` });
      s.addText(t, { x: x + 0.7, y: yy - 0.12, w: w - 0.7, h: 0.75, fontSize: 16, color: C.text1,
        valign: "middle", margin: 0, isTextBox: true, objectName: `step-${i + 1}-text` });
    });

  const addPoint = (s, text, x = 0.6, y = 5.1, w = 5.9, h = 1.35) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.background2 },
      line: { type: "none" }, objectName: "point-box" });
    s.addImage({ data: ICON.bulb, x: x + 0.3, y: y + 0.3, w: 0.42, h: 0.42, objectName: "point-icon" });
    s.addText([{ text: "ポイント", options: { bold: true, color: C.accent1, breakLine: true } }, { text }],
      { x: x + 0.95, y: y + 0.15, w: w - 1.2, h: h - 0.3, fontSize: 14, color: C.text1, valign: "middle",
        margin: 0, isTextBox: true, objectName: "point-text" });
  };

  // 画面イメージ（ブラウザ枠）
  const addScreen = (s, url, x = 7.0, y = 1.45, w = 5.7, h = 4.75) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: C.background1 },
      line: { color: C.accent4, width: 1.25 }, shadow: { type: "outer", color: "000000", opacity: 0.12, blur: 8, offset: 3, angle: 90 },
      objectName: "screen-frame" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: y + 0.18, w: w - 0.4, h: 0.4, rectRadius: 0.2,
      fill: { color: C.background2 }, line: { type: "none" }, objectName: "screen-url-bar" });
    s.addText(url, { x: x + 0.4, y: y + 0.18, w: w - 0.8, h: 0.4, fontSize: 12, color: C.text2, valign: "middle",
      margin: 0, isTextBox: true, objectName: "screen-url" });
    s.addText("※画面はイメージです。実際の表示や文言は変わる場合があります", { x, y: y + h + 0.1, w, h: 0.3,
      fontSize: 10, color: C.text2, margin: 0, isTextBox: true, objectName: "screen-note" });
    return { x: x + 0.4, y: y + 0.8, w: w - 0.8 };
  };
  const btn = (s, text, x, y, w, h, opt = {}) =>
    s.addText(text, { x, y, w, h, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08,
      fill: { color: opt.fill || C.accent6 }, line: opt.line || { type: "none" }, color: opt.color || C.background1,
      fontSize: opt.fontSize || 14, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true,
      objectName: opt.name || "mock-button" });
  const field = (s, label, value, x, y, w) => {
    s.addText(label, { x, y, w, h: 0.3, fontSize: 12, color: C.text2, margin: 0, isTextBox: true, objectName: "mock-label" });
    s.addText(value, { x, y: y + 0.32, w, h: 0.45, fontSize: 13, color: C.text1, margin: [0, 8, 0, 8], valign: "middle",
      shape: pres.shapes.RECTANGLE, fill: { color: C.background1 }, line: { color: C.accent4, width: 1 },
      isTextBox: true, objectName: "mock-field" });
  };
  const highlight = (s, x, y, w, h) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x - 0.08, y: y - 0.08, w: w + 0.16, h: h + 0.16, rectRadius: 0.12,
      fill: { type: "none" }, line: { color: C.accent1, width: 3 }, objectName: "mock-highlight" });

  // ---------- 1. 表紙 ----------
  pres.addSection({ title: "はじめに" });
  let s = pres.addSlide({ masterName: "HW_DARK", sectionTitle: "はじめに" });
  s.addText("LINE公式アカウント\n開設手順書", { placeholder: "title" });
  s.addText("株式会社ハウスワーク 様\n所要時間 約15分・開設は無料です", { placeholder: "body" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.0, y: 2.1, w: 3.5, h: 3.3, rectRadius: 0.15,
    fill: { color: C.background1 }, line: { type: "none" }, objectName: "logo-card" });
  s.addImage({ path: LOGO, x: 9.3, y: 2.75, w: 2.9, h: 1.92, objectName: "logo" });
  s.addNotes("本日はこの資料を見ながら、一緒にLINE公式アカウントを開設します。15分ほどで完了します。");

  // ---------- 2. 全体の流れ ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "はじめに" });
  addTitle(s, null, "開設までの流れ");
  const flow = [
    ["globe", "申込みページを開く", "1分"], ["at", "メールで登録", "2分"],
    ["id", "ビジネスIDを作成", "3分"], ["pen", "情報を入力", "5分"],
    ["check", "開設完了", "1分"], ["users", "制作担当を招待", "3分"],
  ];
  const fw = 1.95, fgap = 0.15, fx0 = (13.333 - (flow.length * fw + (flow.length - 1) * fgap)) / 2;
  flow.forEach(([ic, label, min], i) => {
    const x = fx0 + i * (fw + fgap);
    if (i < flow.length - 1)
      s.addShape(pres.shapes.LINE, { x: x + fw / 2 + 0.6, y: 2.75, w: fw + fgap - 1.2, h: 0,
        line: { color: C.accent4, width: 2, endArrowType: "triangle" }, objectName: `flow-arrow-${i + 1}` });
    s.addShape(pres.shapes.OVAL, { x: x + fw / 2 - 0.55, y: 2.2, w: 1.1, h: 1.1,
      fill: { color: i === flow.length - 1 ? C.accent1 : C.text1 }, line: { type: "none" }, objectName: `flow-circle-${i + 1}` });
    s.addImage({ data: ICON[ic], x: x + fw / 2 - 0.25, y: 2.5, w: 0.5, h: 0.5, objectName: `flow-icon-${i + 1}` });
    s.addText(`STEP ${i + 1}`, { x, y: 3.55, w: fw, h: 0.35, fontSize: 12, bold: true, color: C.accent1,
      align: "center", margin: 0, isTextBox: true, objectName: `flow-step-${i + 1}` });
    s.addText(label, { x, y: 3.9, w: fw, h: 0.75, fontSize: 14, bold: true, color: C.text1,
      align: "center", valign: "top", margin: 0, isTextBox: true, objectName: `flow-label-${i + 1}` });
    s.addText(min, { x, y: 4.35, w: fw, h: 0.35, fontSize: 12, color: C.text2, align: "center", margin: 0,
      isTextBox: true, objectName: `flow-time-${i + 1}` });
  });
  s.addText("STEP 1〜5 で開設完了。STEP 6 で制作担当を招待すると、プロフィールやリッチメニューの設定を代わりに進められます。",
    { x: 1.2, y: 5.55, w: 10.9, h: 0.7, fontSize: 15, color: C.text1, align: "center", valign: "middle",
      shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, fill: { color: C.background2 }, margin: 0.1,
      isTextBox: true, objectName: "flow-summary" });
  s.addNotes("全体で15分ほどです。STEP6の招待までしていただければ、残りの設定はこちらで進めます。");

  // ---------- 3. 準備するもの ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "はじめに" });
  addTitle(s, null, "事前に準備するもの");
  addCheck(s);
  const prep = [
    ["mail", "共有メールアドレス", "info@ など。当日メールを受信できる状態にしておきます"],
    ["key", "パスワード", "8文字以上の英数字。社内で保管してください"],
    ["building", "会社情報", "会社名・住所・電話番号・ホームページURL"],
    ["laptop", "パソコン", "Chrome などのブラウザで操作します"],
    ["phone", "スマートフォン", "LINEアプリが入ったもの。確認や動作チェックに使います"],
    ["image", "ロゴ画像", "プロフィール設定で使います（後日でも可）"],
  ];
  const cw = 3.85, ch = 2.15, cgx = 0.3, cgy = 0.3, cx0 = 0.6, cy0 = 1.55;
  prep.forEach(([ic, head, body], i) => {
    const x = cx0 + (i % 3) * (cw + cgx), y = cy0 + Math.floor(i / 3) * (ch + cgy);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.12, fill: { color: C.background2 },
      line: { type: "none" }, objectName: `prep-card-${i + 1}` });
    s.addImage({ data: ICON[ic], x: x + 0.35, y: y + 0.35, w: 0.5, h: 0.5, objectName: `prep-icon-${i + 1}` });
    s.addText(head, { x: x + 1.05, y: y + 0.3, w: cw - 1.3, h: 0.6, fontSize: 17, bold: true, color: C.text1,
      valign: "middle", margin: 0, isTextBox: true, objectName: `prep-head-${i + 1}` });
    s.addText(body, { x: x + 0.35, y: y + 1.05, w: cw - 0.7, h: 0.9, fontSize: 14, color: C.text1,
      valign: "top", margin: 0, isTextBox: true, objectName: `prep-body-${i + 1}` });
  });
  s.addNotes("メールアドレスは個人用ではなく、会社で共有しているものがおすすめです。担当者が替わっても引き継げます。");

  // ---------- 4. STEP1 ----------
  pres.addSection({ title: "開設手順" });
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設手順" });
  addTitle(s, "STEP 1", "申込みページを開く");
  addCheck(s);
  addSteps(s, [
    "パソコンのブラウザで「entry.line.biz」を開く\n（「LINE公式アカウント 開設」で検索してもOK）",
    "「LINE公式アカウントの開設（無料）」をクリック",
  ]);
  addPoint(s, "開設・月額基本料は無料のプランから始められます。", 0.6, 3.75, 5.9, 1.2);
  let a = addScreen(s, "https://entry.line.biz/");
  s.addText("LINE公式アカウント", { x: a.x, y: a.y + 0.2, w: a.w, h: 0.6, fontSize: 22, bold: true, color: C.text1,
    margin: 0, isTextBox: true, objectName: "mock-heading" });
  s.addText("お店や企業のアカウントを作って、友だちに情報を届けよう", { x: a.x, y: a.y + 0.85, w: a.w, h: 0.5,
    fontSize: 13, color: C.text2, margin: 0, isTextBox: true, objectName: "mock-sub" });
  btn(s, "LINE公式アカウントの開設（無料）", a.x + 0.3, a.y + 1.9, a.w - 0.6, 0.65);
  highlight(s, a.x + 0.3, a.y + 1.9, a.w - 0.6, 0.65);
  s.addNotes("ブラウザで entry.line.biz を開き、緑の「LINE公式アカウントの開設（無料）」ボタンを押してください。");

  // ---------- 5. STEP2 ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設手順" });
  addTitle(s, "STEP 2", "メールアドレスで登録する");
  addCheck(s);
  addSteps(s, [
    "「メールアドレスで登録」を選ぶ",
    "会社の共有メールアドレスを入力",
    "「登録用のリンクを送信」をクリック",
  ]);
  addPoint(s, "「LINEアカウントで登録」は選ばないでください。会社のメールで作ると、担当者が替わっても引き継げます。",
    0.6, 4.6, 5.9, 1.6);
  a = addScreen(s, "LINEビジネスID");
  s.addText("LINEビジネスIDを作成", { x: a.x, y: a.y + 0.1, w: a.w, h: 0.5, fontSize: 18, bold: true, color: C.text1,
    margin: 0, isTextBox: true, objectName: "mock-heading" });
  btn(s, "LINEアカウントで登録", a.x, a.y + 0.8, a.w, 0.55, { fill: C.background2, color: C.text2, name: "mock-button-line" });
  btn(s, "メールアドレスで登録", a.x, a.y + 1.55, a.w, 0.55);
  highlight(s, a.x, a.y + 1.55, a.w, 0.55);
  field(s, "メールアドレス", "info@house-work.jp（例）", a.x, a.y + 2.35, a.w);
  btn(s, "登録用のリンクを送信", a.x + 0.8, a.y + 3.3, a.w - 1.6, 0.5, { fill: C.text1, name: "mock-button-send" });
  s.addNotes("必ず「メールアドレスで登録」を選びます。入力したアドレスに登録用のリンクが届きます。");

  // ---------- 6. STEP3 ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設手順" });
  addTitle(s, "STEP 3", "ビジネスIDを作成する");
  addCheck(s);
  addSteps(s, [
    "届いたメールの「登録画面に進む」をクリック",
    "名前とパスワードを入力",
    "「登録」→ 内容を確認して「登録」",
    "「サービスに移動」をクリック",
  ], 0.6, 1.6, 5.9, 0.85);
  addPoint(s, "メールが届かないときは、迷惑メールフォルダを確認してください。", 0.6, 5.1, 5.9, 1.2);
  a = addScreen(s, "LINEビジネスID｜登録");
  field(s, "名前", "株式会社ハウスワーク", a.x, a.y + 0.05, a.w);
  field(s, "パスワード", "●●●●●●●●", a.x, a.y + 0.95, a.w);
  field(s, "パスワード（確認用）", "●●●●●●●●", a.x, a.y + 1.85, a.w);
  btn(s, "登録", a.x + 1.2, a.y + 2.95, a.w - 2.4, 0.55);
  highlight(s, a.x + 1.2, a.y + 2.95, a.w - 2.4, 0.55);
  s.addNotes("パスワードはこの場で社内の管理表などに控えておいてください。");

  // ---------- 7. STEP4 ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設手順" });
  addTitle(s, "STEP 4", "アカウント情報を入力する");
  addCheck(s);
  const hdr = { bold: true, color: HEX.white, fill: { color: HEX.black } };
  const rows = [
    ["アカウント名", "ハウスワーク｜小山市の注文住宅"],
    ["メールアドレス", "会社の共有メールアドレス"],
    ["会社・事業者の所在国・地域", "日本"],
    ["会社・事業者名", "株式会社ハウスワーク"],
    ["業種", "住宅・不動産 → 工務店・住宅メーカーに近い項目"],
    ["運用目的", "集客・販促、情報発信 など"],
    ["主な使い方", "メッセージ配信 など"],
  ];
  s.addTable([
    [{ text: "項目", options: hdr }, { text: "入力内容", options: hdr }],
    ...rows.map(([k, v], i) => [
      { text: k, options: { bold: true, fill: { color: i % 2 ? HEX.white : HEX.tint } } },
      { text: v, options: { fill: { color: i % 2 ? HEX.white : HEX.tint } } },
    ]),
  ], { x: 0.6, y: 1.5, w: 8.1, colW: [2.9, 5.2], rowH: 0.56, fontSize: 14, color: HEX.black, valign: "middle",
    margin: [0.04, 0.12, 0.04, 0.12], border: { type: "solid", pt: 0.75, color: HEX.line }, objectName: "account-table" });
  addPoint(s, "アカウント名は、友だちのトーク一覧に表示される名前です。", 9.0, 1.5, 3.7, 2.0);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.0, y: 3.75, w: 3.7, h: 2.3, rectRadius: 0.12,
    fill: { color: C.background1 }, line: { color: C.accent5, width: 1 }, objectName: "name-ideas-box" });
  s.addText([
    { text: "アカウント名の候補", options: { bold: true, color: C.accent1, breakLine: true } },
    { text: "ハウスワーク", options: { bullet: true, breakLine: true } },
    { text: "HOUSE WORK（ハウスワーク）", options: { bullet: true, breakLine: true } },
    { text: "ハウスワーク｜注文住宅・古河展示場", options: { bullet: true } },
  ], { x: 9.2, y: 3.9, w: 3.35, h: 2.0, fontSize: 13, color: C.text1, valign: "top", margin: 0,
    paraSpaceAfter: 4, isTextBox: true, objectName: "name-ideas" });
  s.addNotes("表のとおりに入力します。業種は「住宅・不動産」を選び、小業種は工務店・住宅メーカーに近いものを選んでください。");

  // ---------- 8. STEP5 ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設手順" });
  addTitle(s, "STEP 5", "確認して開設完了");
  addCheck(s);
  addSteps(s, [
    "入力内容を確認して「完了」をクリック",
    "「LINE Official Account Manager へ」をクリック",
    "規約などの同意画面が出たら、内容を確認して同意",
  ]);
  addPoint(s, "管理画面（manager.line.biz）をブックマークしておくと、次回からすぐに開けます。", 0.6, 4.6, 5.9, 1.6);
  a = addScreen(s, "LINE公式アカウント｜作成完了");
  s.addImage({ data: ICON.checkG, x: a.x + a.w / 2 - 0.45, y: a.y + 0.3, w: 0.9, h: 0.9, objectName: "mock-done-icon" });
  s.addText("LINE公式アカウントの作成が\n完了しました", { x: a.x, y: a.y + 1.4, w: a.w, h: 0.9, fontSize: 17, bold: true,
    color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "mock-done-text" });
  btn(s, "LINE Official Account Manager へ", a.x + 0.3, a.y + 2.75, a.w - 0.6, 0.6);
  highlight(s, a.x + 0.3, a.y + 2.75, a.w - 0.6, 0.6);
  s.addNotes("ここで開設は完了です。続けて管理画面から制作担当を招待します。");

  // ---------- 9. STEP6 ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設手順" });
  addTitle(s, "STEP 6", "制作担当をメンバーに招待する");
  addCheck(s);
  addSteps(s, [
    "管理画面の右上「設定」をクリック",
    "左のメニューから「権限」を選ぶ",
    "「メンバーを追加」→ 権限を「管理者」にして「URLを発行」",
    "発行されたURLを制作担当へ送る（有効期限は24時間）",
  ], 0.6, 1.6, 5.9, 0.85);
  addPoint(s, "パスワードを共有せずに、安全に一緒に運用できます。", 0.6, 5.1, 5.9, 1.2);
  a = addScreen(s, "manager.line.biz｜設定");
  ["アカウント設定", "権限", "応答設定", "Messaging API"].forEach((t, i) => {
    s.addText(t, { x: a.x, y: a.y + 0.1 + i * 0.55, w: 1.9, h: 0.45, fontSize: 13, bold: i === 1,
      color: i === 1 ? C.accent1 : C.text2, valign: "middle", margin: [0, 8, 0, 8],
      fill: { color: i === 1 ? C.background2 : C.background1 }, isTextBox: true, objectName: `mock-menu-${i + 1}` });
  });
  if (true) highlight(s, a.x, a.y + 0.65, 1.9, 0.45);
  s.addText("権限", { x: a.x + 2.2, y: a.y + 0.1, w: 2.7, h: 0.45, fontSize: 17, bold: true, color: C.text1,
    margin: 0, isTextBox: true, objectName: "mock-perm-heading" });
  btn(s, "メンバーを追加", a.x + 2.2, a.y + 0.75, 2.7, 0.5, { fill: C.text1, name: "mock-button-add" });
  s.addText("権限：管理者", { x: a.x + 2.2, y: a.y + 1.5, w: 2.7, h: 0.45, fontSize: 13, color: C.text1,
    margin: [0, 8, 0, 8], valign: "middle", shape: pres.shapes.RECTANGLE, fill: { color: C.background1 },
    line: { color: C.accent4, width: 1 }, isTextBox: true, objectName: "mock-role" });
  btn(s, "URLを発行", a.x + 2.2, a.y + 2.2, 2.7, 0.5, { name: "mock-button-url" });
  highlight(s, a.x + 2.2, a.y + 2.2, 2.7, 0.5);
  s.addNotes("発行したURLをメールやLINEで制作担当に送ってください。24時間以内に承認します。");

  // ---------- 10. 困ったとき ----------
  pres.addSection({ title: "開設後" });
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設後" });
  addTitle(s, null, "困ったときは");
  const faq = [
    ["メールが届かない", "迷惑メールフォルダを確認し、もう一度「登録用のリンクを送信」を押してください。"],
    ["すでにビジネスIDがある", "「LINEビジネスIDでログイン」から進めば、新しく作る必要はありません。"],
    ["アカウント名を変えたい", "管理画面の「設定」から後で変更できます。"],
    ["料金が気になる", "無料プランで開始できます。配信数が増えたら有料プランを検討します。"],
  ];
  const qw = 5.9, qh = 2.25;
  faq.forEach(([q, ans], i) => {
    const x = 0.6 + (i % 2) * (qw + 0.3), y = 1.55 + Math.floor(i / 2) * (qh + 0.3);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: qw, h: qh, rectRadius: 0.12, fill: { color: C.background2 },
      line: { type: "none" }, objectName: `faq-card-${i + 1}` });
    s.addImage({ data: ICON.q, x: x + 0.35, y: y + 0.35, w: 0.45, h: 0.45, objectName: `faq-icon-${i + 1}` });
    s.addText(q, { x: x + 1.0, y: y + 0.3, w: qw - 1.3, h: 0.55, fontSize: 18, bold: true, color: C.text1,
      valign: "middle", margin: 0, isTextBox: true, objectName: `faq-q-${i + 1}` });
    s.addText(ans, { x: x + 1.0, y: y + 1.0, w: qw - 1.3, h: 1.0, fontSize: 14, color: C.text1, valign: "top",
      margin: 0, isTextBox: true, objectName: `faq-a-${i + 1}` });
  });

  // ---------- 11. 開設後の設定 ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設後" });
  addTitle(s, null, "開設後の設定は制作側で進めます");
  const after = [
    ["user", "プロフィール", "ロゴ・営業時間・住所・ホームページを登録"],
    ["comment", "あいさつメッセージ", "友だち追加された直後に届くメッセージ"],
    ["grid", "リッチメニュー", "トーク画面の下に出るメニュー（次のページ）"],
    ["shield", "認証済アカウント申請", "青いバッジが付き、LINE内の検索に表示される"],
  ];
  after.forEach(([ic, head, body], i) => {
    const y = 1.55 + i * 1.2;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.85, h: 0.85, fill: { color: C.background2 }, line: { type: "none" },
      objectName: `after-circle-${i + 1}` });
    s.addImage({ data: ICON[ic], x: 0.81, y: y + 0.21, w: 0.43, h: 0.43, objectName: `after-icon-${i + 1}` });
    s.addText(head, { x: 1.7, y: y - 0.05, w: 4.5, h: 0.5, fontSize: 18, bold: true, color: C.text1, valign: "middle",
      margin: 0, isTextBox: true, objectName: `after-head-${i + 1}` });
    s.addText(body, { x: 1.7, y: y + 0.42, w: 4.8, h: 0.45, fontSize: 14, color: C.text2, valign: "middle",
      margin: 0, isTextBox: true, objectName: `after-body-${i + 1}` });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.0, y: 1.55, w: 5.7, h: 3.3, rectRadius: 0.12,
    fill: { color: C.text1 }, line: { type: "none" }, objectName: "ask-box" });
  s.addText([
    { text: "お願いしたいこと", options: { bold: true, color: C.accent1, fontSize: 18, breakLine: true } },
    { text: "各ボタンのリンク先URL", options: { bullet: true, breakLine: true } },
    { text: "　建築実例・古河展示場 Room Tour", options: { breakLine: true, color: C.accent4 } },
    { text: "　資料請求フォーム／見学予約フォーム", options: { breakLine: true, color: C.accent4 } },
    { text: "営業時間・定休日", options: { bullet: true, breakLine: true } },
    { text: "展示場や施工の写真（あれば）", options: { bullet: true } },
  ], { x: 7.4, y: 1.85, w: 5.0, h: 2.8, fontSize: 16, color: C.background1, valign: "top", margin: 0,
    paraSpaceAfter: 8, isTextBox: true, objectName: "ask-text" });

  // ---------- 12. リッチメニュー ----------
  s = pres.addSlide({ masterName: "HW_CONTENT", sectionTitle: "開設後" });
  addTitle(s, null, "リッチメニューでできること");
  // スマホ枠
  const px = 0.9, py = 1.3, pw = 4.3, ph = 5.45;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: ph, rectRadius: 0.35, fill: { color: C.text1 },
    line: { type: "none" }, objectName: "phone-body" });
  const sx = px + 0.18, sw = pw - 0.36, sy = py + 0.35;
  s.addShape(pres.shapes.RECTANGLE, { x: sx, y: sy, w: sw, h: ph - 0.7, fill: { color: "8CABD9" }, line: { type: "none" },
    objectName: "phone-screen" });
  s.addText("ハウスワーク", { x: sx, y: sy, w: sw, h: 0.42, fontSize: 12, bold: true, color: C.background1, align: "center",
    valign: "middle", fill: { color: "6E8DBB" }, margin: 0, isTextBox: true, objectName: "phone-header" });
  s.addText("友だち追加ありがとうございます！", { x: sx + 0.2, y: sy + 0.6, w: 2.9, h: 0.5, fontSize: 11, color: C.text1,
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, fill: { color: C.background1 }, valign: "middle",
    margin: [0, 8, 0, 8], isTextBox: true, objectName: "phone-message" });
  const mh = sw * 1686 / 2500, my = sy + (ph - 0.7) - mh;
  s.addImage({ path: RICHMENU, x: sx, y: my, w: sw, h: mh, objectName: "richmenu-image" });
  // 番号バッジ（メニュー上）
  const bo = 0.26; // 角からの距離
  const badges = [
    [sx + sw - bo, my + bo], [sx + sw / 3 - bo, my + mh / 2 + bo],
    [sx + sw * 2 / 3 - bo, my + mh / 2 + bo], [sx + sw - bo, my + mh / 2 + bo],
  ];
  badges.forEach(([bx, by], i) =>
    s.addText(String(i + 1), { x: bx - 0.19, y: by - 0.19, w: 0.38, h: 0.38, shape: pres.shapes.OVAL,
      fill: { color: C.accent1 }, line: { color: C.background1, width: 1.5 }, color: C.background1, fontSize: 13,
      bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `menu-badge-${i + 1}` }));
  // 説明
  const menu = [
    ["ハウスワークの家", "選べるデザイン注文住宅。会社と家づくりの紹介"],
    ["建築実例", "施工事例と ROOM TOUR 動画が見られる"],
    ["家づくり資料プレゼント", "家づくりの資料を申し込める"],
    ["見学予約", "展示場の見学をその場で予約できる"],
  ];
  menu.forEach(([head, body], i) => {
    const y = 1.45 + i * 1.3;
    s.addText(String(i + 1), { x: 6.0, y, w: 0.5, h: 0.5, shape: pres.shapes.OVAL, fill: { color: C.accent1 },
      color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true,
      objectName: `menu-num-${i + 1}` });
    s.addText(head, { x: 6.75, y: y - 0.02, w: 5.9, h: 0.5, fontSize: 19, bold: true, color: C.text1, valign: "middle",
      margin: 0, isTextBox: true, objectName: `menu-head-${i + 1}` });
    s.addText(body, { x: 6.75, y: y + 0.5, w: 5.9, h: 0.6, fontSize: 14, color: C.text2, valign: "top",
      margin: 0, isTextBox: true, objectName: `menu-body-${i + 1}` });
  });
  s.addNotes("友だち追加すると、トーク画面の下にこのメニューが表示されます。ボタンを押すだけで各ページが開きます。");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
