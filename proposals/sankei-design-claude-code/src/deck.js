// ① 詳細スライド（18枚）
const pptxgen = require("pptxgenjs");
const path = require("path");
const { THEME, HEX, icon, txt } = require("./common");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const DATE = "2026年11月";
const AS_OF = "2026年11月時点";

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Claude Code 業務活用のご提案（サンケイデザイン株式会社様）";
  pres.subject = "Claude Code の業務活用のご紹介とご提案";
  pres.author = "【作成者名】";
  const C = pres.SchemeColor;

  const W = 13.333, MX = 0.6, CW = W - MX * 2;
  const FOOT = "サンケイデザイン株式会社様　Claude Code 業務活用のご提案";

  // ---------- レイアウト ----------
  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: C.text1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 2.1, w: 9.3, h: 1.0, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: MX, y: 4.6, w: 9.3, h: 0.5, fontSize: 16, color: C.accent6, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.4, 0.6, 0.7, 0.6],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.32, w: CW, h: 0.62, fontSize: 26, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { rect: { x: MX, y: 1.1, w: 0.78, h: 0.4, fill: { color: C.accent1 }, rectRadius: 0.08 } },
      { text: { text: "結論", options: { x: MX, y: 1.1, w: 0.78, h: 0.4, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 } } },
      { placeholder: { options: { name: "lead", type: "body", x: MX + 0.95, y: 1.04, w: CW - 0.95, h: 0.52, fontSize: 16, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
      { text: { text: FOOT, options: { x: MX, y: 7.02, w: 8, h: 0.3, fontSize: 9, color: C.accent3, margin: 0, valign: "middle" } } },
    ],
    slideNumber: { x: W - MX - 0.6, y: 7.02, w: 0.6, h: 0.3, fontSize: 9, color: C.accent3, align: "right" },
  });
  // CONTENT の rectRadius は ROUNDED_RECTANGLE でないと無視されるため、角丸タグは不要なら矩形のままでよい

  let sec = "";
  function section(t) { pres.addSection({ title: t }); sec = t; }
  function content(title, lead) {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
    s.addText(title, { placeholder: "title" });
    s.addText(lead, { placeholder: "lead" });
    return s;
  }
  function card(s, x, y, w, h, fill, name) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill || C.background2 }, line: { type: "none" }, rectRadius: 0.1, objectName: name || "card" });
  }
  async function iconCircle(s, name, x, y, d, bg, fgHex) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: "icon-bg" });
    const pad = d * 0.24;
    s.addImage({ data: await icon(name, fgHex), x: x + pad, y: y + pad, w: d - pad * 2, h: d - pad * 2, altText: name });
  }
  function numCircle(s, n, x, y, d, bg) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg || C.accent1 }, line: { type: "none" }, objectName: "num" });
    txt(s, String(n), { x, y, w: d, h: d, fontSize: Math.round(d * 30), bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, lineSpacing: null });
  }
  function arrow(s, x, y, w, h, color) {
    s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w, h, fill: { color: color || C.accent6 }, line: { type: "none" }, objectName: "arrow" });
  }
  function downArrow(s, x, y, w, h, color) {
    s.addShape(pres.shapes.DOWN_ARROW, { x, y, w, h, fill: { color: color || C.accent6 }, line: { type: "none" }, objectName: "arrow" });
  }
  // 表：ヘッダー行は藍、本文は白と淡色の交互
  function table(s, header, rows, o) {
    const fs = o.fontSize || 12;
    const head = header.map((h) => ({ text: h, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: fs, valign: "middle" } }));
    const body = rows.map((r, i) =>
      r.map((c, j) => {
        const cell = typeof c === "string" ? { text: c, options: {} } : c;
        cell.options = Object.assign({ fontSize: fs, color: HEX.dk1, fill: { color: i % 2 ? HEX.lt2 : HEX.lt1 }, valign: "middle", bold: j === 0 && o.boldFirst !== false }, cell.options);
        return cell;
      })
    );
    s.addTable([head, ...body], {
      x: o.x, y: o.y, w: o.w, colW: o.colW, rowH: o.rowH,
      border: { type: "solid", pt: 0.75, color: HEX.accent6 },
      margin: [0.05, 0.08, 0.05, 0.08], fontFace: THEME.bodyFontFace,
      autoPage: false,
    });
  }
  const NOTE = (s, t) => s.addNotes(t);

  // ======================= 1. 表紙 =======================
  section("はじめに");
  {
    const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: sec });
    txt(s, "サンケイデザイン株式会社様", { x: MX, y: 1.45, w: 9, h: 0.5, fontSize: 18, bold: true, color: C.accent1, margin: 0 });
    s.addText("Claude Code 業務活用のご提案", { placeholder: "title" });
    txt(s, "人を増やさずに、できる仕事の幅を広げる", { x: MX, y: 3.25, w: 9.3, h: 0.6, fontSize: 24, color: C.background1, margin: 0 });
    s.addText(`${DATE}　作成：【作成者名】（SecondDesk 北大路堀川 利用者）`, { placeholder: "body" });
    // 右側のモチーフ：校正記号風の円とアイコン
    s.addShape(pres.shapes.OVAL, { x: 10.15, y: 1.9, w: 2.6, h: 2.6, fill: { color: C.text2 }, line: { type: "none" }, objectName: "motif" });
    s.addShape(pres.shapes.OVAL, { x: 11.55, y: 3.85, w: 1.2, h: 1.2, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "motif" });
    s.addImage({ data: await icon("FaFolderOpen", HEX.lt1), x: 10.85, y: 2.6, w: 1.2, h: 1.2, altText: "フォルダのアイコン" });
    s.addImage({ data: await icon("FaPenFancy", HEX.lt1), x: 11.85, y: 4.15, w: 0.6, h: 0.6, altText: "ペンのアイコン" });
    txt(s, "本資料は、貴社が運営されるコワーキングスペースを利用させていただいている外部の立場から、ご参考としてご提案するものです。料金・機能は変更されることがあるため、最新情報は各公式ページをご確認ください。", { x: MX, y: 6.15, w: 11.5, h: 0.7, fontSize: 11, color: C.accent6, margin: 0 });
    NOTE(s, "本日はお時間をいただきありがとうございます。SecondDesk 北大路堀川でいつもお世話になっている【作成者名】です。今日は、AIのClaude Code（クロード・コード）を貴社の業務にどう活かせるかを、外部の立場からご提案させていただきます。結論から申し上げると、「人を増やさずに、対応できる仕事の幅を広げる」道具としてご検討いただきたい、というのが本日の趣旨です。資料は会議後にも読み返せるよう、文字を多めにしています。");
  }

  // ======================= 2. エグゼクティブサマリー =======================
  {
    const s = content("エグゼクティブサマリー", "まずは月数千円〜1万円台の規模で、2つの業務に絞った試験導入から始めることをご提案します");
    const items = [
      ["FaRobot", "何ができるか", "パソコンの中の案件フォルダを開き、原稿のチェック、議事録、Excel・PowerPoint・HTMLの作成や修正まで、指示ひとつで作業を進める「AIアシスタント」です。"],
      ["FaLayerGroup", "業務の何が変わるか", "校正の一次チェック、議事録から制作指示書・見積下書きまでの書類づくり、Web更新などの「下ごしらえ」をAIに任せ、人は判断と仕上げに集中できます。"],
      ["FaYenSign", "費用の目安", "個人向け Pro は月額約$20、法人向け Team は1席あたり月額約$25（年払いは$20）から。3名で始めても月額$60〜75程度です（税抜・USD、" + AS_OF + "）。"],
      ["FaRocket", "次の一歩", "1〜2か月、少人数で「原稿校正」と「議事録・制作準備」の2業務に絞って試験導入し、削減時間を実測してから本格導入を判断していただく流れをご提案します。"],
    ];
    const cw = (CW - 0.3 * 3) / 4;
    for (let i = 0; i < 4; i++) {
      const x = MX + i * (cw + 0.3), y = 1.85;
      card(s, x, y, cw, 4.25);
      await iconCircle(s, items[i][0], x + 0.25, y + 0.25, 0.75, i === 3 ? C.accent1 : C.text2, HEX.lt1);
      txt(s, items[i][1], { x: x + 0.25, y: y + 1.15, w: cw - 0.5, h: 0.45, fontSize: 16, bold: true, color: C.text1, margin: 0 });
      txt(s, items[i][2], { x: x + 0.25, y: y + 1.65, w: cw - 0.5, h: 2.5, fontSize: 13, color: C.text1, margin: 0 });
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 6.25, w: CW, h: 0.65, fill: { color: C.background1 }, line: { color: C.accent6, width: 1 }, rectRadius: 0.1, objectName: "note-frame" });
    txt(s, "価値の中心：少人数・兼務の体制のまま、制作本数や事業の幅を広げること。AIは「下ごしらえ」と「たたき台」まで、最終判断と品質責任は人が持つ前提です。", { x: MX + 0.25, y: 6.28, w: CW - 0.5, h: 0.6, fontSize: 13, color: C.text2, valign: "middle", margin: 0 });
    NOTE(s, "最初に結論を4点でお伝えします。1つ目、Claude Codeはパソコンの中のファイルを直接扱い、作業まで進めてくれるAIです。2つ目、貴社では校正の一次チェックや、打ち合わせ後の書類づくりなどの「下ごしらえ」を任せることで、人は判断と仕上げに集中できます。3つ目、費用は1人あたり月額20〜25ドル程度からで、3名でも月1万円前後の規模です。4つ目、いきなり全社導入ではなく、1〜2か月の試験導入で効果を実測していただくことをご提案します。なお、AIは間違えることもあるため、最終確認は必ず人が行う前提です。");
  }

  // ======================= 3. Claude Code とは =======================
  section("Claude Code とは");
  {
    const s = content("Claude Code とは（一言で）", "「相談相手」ではなく、自分の机に座って実際に作業までしてくれるアシスタントです");
    const colW = (CW - 0.4) / 2;
    const cols = [
      ["FaComments", "普段の ChatGPT（チャット）", "窓口で相談に乗ってくれる人", ["質問すると、文章で答えてくれる", "資料は人が窓口まで持って行き、見せる（コピー＆貼り付け、添付）", "回答を受け取ったら、自分の机に戻って反映するのは人"], C.accent3],
      ["FaDesktop", "Claude Code", "自分の机に座り、フォルダを開いて作業までしてくれるアシスタント", ["案件フォルダを自分で開いて、原稿を読む", "修正案の一覧（Excel）や議事録（Word）などを、フォルダに直接作って保存する", "大事な操作の前には「実行してよいですか？」と確認してくれる"], C.accent1],
    ];
    for (let i = 0; i < 2; i++) {
      const x = MX + i * (colW + 0.4), y = 1.85, c = cols[i];
      card(s, x, y, colW, 4.0);
      await iconCircle(s, c[0], x + 0.3, y + 0.3, 0.9, c[4], HEX.lt1);
      txt(s, c[1], { x: x + 1.4, y: y + 0.3, w: colW - 1.7, h: 0.4, fontSize: 14, color: C.accent3, bold: true, margin: 0 });
      txt(s, c[2], { x: x + 1.4, y: y + 0.7, w: colW - 1.7, h: 0.8, fontSize: 18, bold: true, color: C.text1, margin: 0 });
      txt(s, c[3].map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < c[3].length - 1 } })), { x: x + 0.35, y: y + 1.65, w: colW - 0.7, h: 2.2, fontSize: 14, color: C.text1, paraSpaceAfter: 6, margin: 0 });
    }
    txt(s, [
      { text: "もう少し正確に言うと：", options: { bold: true, color: C.text2 } },
      { text: "Claude Code は、Anthropic（アンソロピック）社のAI「Claude」を、パソコンの中で「作業する人」として動かすためのソフトです。こうした「目的を伝えると、手順を考えて自分で作業を進めるAI」を「エージェント」と呼びます。もともとは開発者向けの道具ですが、現在はデスクトップアプリからも使え、文章・表・資料づくりなど事務や制作の作業にも使えます。" },
    ], { x: MX, y: 6.0, w: CW, h: 0.9, fontSize: 13, color: C.text1, margin: 0 });
    NOTE(s, "たとえ話でご説明します。普段のChatGPTは「窓口で相談に乗ってくれる人」です。資料を窓口まで持って行って見せ、答えをもらって、自分の机に戻って反映するのは人です。一方Claude Codeは「自分の机に座ってくれるアシスタント」です。机の上の案件フォルダを自分で開き、原稿を読み、修正案の一覧や議事録をフォルダに直接作って保存してくれます。ただし勝手に何でもするわけではなく、大事な操作の前には「やってよいですか」と確認を求めてきます。このように、目的を伝えると手順を考えて作業まで進めるAIを「エージェント」と呼びます。");
  }

  // ======================= 4. ローカル環境で動くとは =======================
  {
    const s = content("「ローカル環境で動く」とは", "人がファイルとチャットの間を行き来する手間がなくなり、一言の指示で作業が完結します");
    // Before
    const by = 1.8, bh = 2.8, half = (CW - 0.4) / 2;
    card(s, MX, by, half, bh);
    txt(s, "Before：人が「仲介役」になる", { x: MX + 0.25, y: by + 0.15, w: half - 0.5, h: 0.4, fontSize: 15, bold: true, color: C.accent3, margin: 0 });
    const steps = ["原稿ファイル\nを開く", "チャット\nに貼る", "回答を\nコピー", "元のファイルに\n貼って保存"];
    const sw = 1.25, gap = (half - 0.5 - sw * 4) / 3;
    for (let i = 0; i < 4; i++) {
      const x = MX + 0.25 + i * (sw + gap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: by + 0.75, w: sw, h: 1.0, fill: { color: C.background1 }, line: { color: C.accent6, width: 1 }, rectRadius: 0.08, objectName: "step" });
      txt(s, steps[i], { x: x + 0.05, y: by + 0.75, w: sw - 0.1, h: 1.0, fontSize: 11, color: C.text1, align: "center", valign: "middle", margin: 0 });
      if (i < 3) arrow(s, x + sw + gap / 2 - 0.12, by + 1.13, 0.24, 0.24, C.accent6);
    }
    txt(s, "ファイルが10本あれば、この往復を10回。貼り間違い・反映漏れの心配も人が負います。", { x: MX + 0.25, y: by + 1.95, w: half - 0.5, h: 0.7, fontSize: 13, color: C.text1, margin: 0 });
    // After
    const ax = MX + half + 0.4;
    card(s, ax, by, half, bh, "FCEDE9");
    txt(s, "After：一言で完結する", { x: ax + 0.25, y: by + 0.15, w: half - 0.5, h: 0.4, fontSize: 15, bold: true, color: C.accent1, margin: 0 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ax + 0.25, y: by + 0.7, w: half - 0.5, h: 1.05, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.25 }, rectRadius: 0.1, objectName: "prompt" });
    txt(s, "「この案件フォルダの原稿を、表記ルールでチェックして、修正案を一覧（Excel）にして」", { x: ax + 0.45, y: by + 0.72, w: half - 0.9, h: 1.0, fontSize: 14, bold: true, color: C.text1, valign: "middle", margin: 0 });
    txt(s, "AIがフォルダ内の原稿を自分で開いて読み、修正案の一覧ファイルを同じフォルダに保存します。人は一覧を見て採否を決めるだけです。", { x: ax + 0.25, y: by + 1.95, w: half - 0.5, h: 0.7, fontSize: 13, color: C.text1, margin: 0 });
    // ループ図
    txt(s, "作業の進み方（1つの指示の中で、このループを自動で回します）", { x: MX, y: 4.72, w: CW, h: 0.35, fontSize: 14, bold: true, color: C.text2, margin: 0 });
    const loop = [["指示", "人が目的を伝える"], ["計画", "AIが手順を考える"], ["許可を求める", "変更の前に確認"], ["実行", "読む・書く・作る"], ["確認・報告", "結果を見直し、人に報告"]];
    const lw = 2.05, lg = (CW - lw * 5) / 4, ly = 5.15;
    for (let i = 0; i < 5; i++) {
      const x = MX + i * (lw + lg);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: ly, w: lw, h: 1.05, fill: { color: i === 2 ? C.accent1 : C.text2 }, line: { type: "none" }, rectRadius: 0.1, objectName: "loop-step" });
      txt(s, [{ text: loop[i][0], options: { bold: true, fontSize: 15, breakLine: true } }, { text: loop[i][1], options: { fontSize: 11 } }], { x: x + 0.08, y: ly + 0.08, w: lw - 0.16, h: 0.9, color: C.background1, align: "center", valign: "middle", margin: 0, lineSpacing: 19 });
      if (i < 4) arrow(s, x + lw + lg / 2 - 0.14, ly + 0.38, 0.28, 0.28, C.accent6);
    }
    // 戻り矢印（報告→次の指示）
    s.addShape(pres.shapes.LINE, { x: MX + lw / 2, y: 6.48, w: CW - lw, h: 0, line: { color: C.accent3, width: 1.25, dashType: "dash", beginArrowType: "triangle" }, objectName: "loop-back" });
    s.addShape(pres.shapes.LINE, { x: MX + lw / 2, y: 6.2, w: 0, h: 0.28, line: { color: C.accent3, width: 1.25, dashType: "dash" }, objectName: "loop-back" });
    s.addShape(pres.shapes.LINE, { x: MX + CW - lw / 2, y: 6.2, w: 0, h: 0.28, line: { color: C.accent3, width: 1.25, dashType: "dash" }, objectName: "loop-back" });
    txt(s, "報告を見て、人が次の指示や修正を出す", { x: MX + 4.2, y: 6.5, w: 4.5, h: 0.32, fontSize: 11, color: C.accent3, align: "center", margin: 0 });
    NOTE(s, "「ローカル環境で動く」とは、AIが皆さんのパソコンの中、つまり手元のフォルダで作業するという意味です。左のBeforeをご覧ください。普段のチャットAIでは、ファイルを開いて、チャットに貼って、回答をコピーして、元のファイルに貼って保存する。人がファイルとAIの間の「仲介役」になっています。右のAfterでは「この案件フォルダの原稿をチェックして、修正案を一覧にして」と一言伝えるだけで、AIが自分でファイルを開き、結果をフォルダに保存します。下の図が作業の進み方です。指示を受けたAIは手順を考え、ファイルを変更する前には許可を求め、実行し、結果を見直して報告します。赤色の「許可を求める」がポイントで、勝手に作業が進まない仕組みになっています。");
  }

  // ======================= 5. 比較A =======================
  {
    const s = content("比較A：普段の ChatGPT（チャット利用）との違い", "違いは「答えを返す」か「作業を終わらせる」か。手元のファイルを直接扱えることが最大の差です");
    table(s, ["観点", "普段の ChatGPT（チャット利用）", "Claude Code"], [
      ["ファイルの扱い", "原稿をコピー＆貼り付け、またはファイルを添付して渡す", "案件フォルダを直接読み書きする（貼り付け不要）"],
      ["作業の単位", "1回の質問に1回の回答（1往復）が基本", "「読む→直す→保存→確認」の複数ステップを最後まで進める"],
      ["成果物", "主に画面上のテキスト。ファイルを作れる場合も、ダウンロードして保存し直す", "Word・Excel・PowerPoint・CSV・HTML などの実ファイルを、フォルダに直接作成・更新"],
      ["繰り返し作業", "毎回同じ指示を書く（カスタム指示やプロジェクト機能で一部は保存可）", "表記ルールや手順を CLAUDE.md・Skills としてファイルに保存し、チームで共有"],
    ], { x: MX, y: 1.85, w: CW, colW: [2.2, 4.95, 4.98], rowH: [0.45, 0.85, 0.85, 0.95, 0.95], fontSize: 13 });
    await iconCircle(s, "FaBalanceScale", MX, 6.0, 0.7, C.background2, HEX.dk2);
    txt(s, [
      { text: "公平のために：", options: { bold: true, color: C.text2 } },
      { text: "ちょっとした相談、アイデア出し、調べものは、普段のチャットで十分なことも多くあります。Claude Code は「ファイルを伴う作業」をまとめて任せたいときに力を発揮する道具、と考えるのが適切です。" },
    ], { x: MX + 0.9, y: 5.95, w: CW - 0.9, h: 0.85, fontSize: 13, color: C.text1, valign: "middle", margin: 0 });
    NOTE(s, "まず、皆さんがよくご存じの、普段のChatGPTのチャット利用と比べます。一番大きな違いはファイルの扱いです。チャットでは原稿を貼り付けたり添付したりして渡しますが、Claude Codeはフォルダを直接読み書きします。また、チャットは1回の質問に1回の答えが基本ですが、Claude Codeは「読む、直す、保存する、確認する」という複数の手順を最後まで進めます。成果物も、画面上の文章ではなく、ExcelやWordなどの実際のファイルとしてフォルダに残ります。繰り返しの作業は、ルールをファイルとして保存して共有できます。ただし、ちょっとした相談やアイデア出しはチャットで十分です。道具の得意分野が違う、という理解が正確です。");
  }

  // ======================= 6. 比較B =======================
  {
    const s = content("比較B：ChatGPT のエージェント機能との違い", "優劣ではなく使い分け。手元のファイル作業は Claude Code、Web上の作業は ChatGPT が得意です");
    table(s, ["観点", "ChatGPT のエージェント機能（Agent モード／Codex）", "Claude Code"], [
      ["どこで動くか", "Agent モード：クラウド上の仮想コンピューターでWeb閲覧や操作を行う【要確認】\nCodex：クラウドの隔離環境のほか、手元で動くCLI・IDE拡張もある（主に開発者向け）", "自分のPCのフォルダで動く（デスクトップアプリ／ターミナル）。必要に応じてクラウドで動かすWeb版もある"],
      ["手元の案件ファイルとの距離", "クラウド側では、必要なファイルをアップロードして渡す", "PC内の案件フォルダをそのまま扱える（アップロード不要）"],
      ["作業ルールの保存・共有", "カスタム指示、プロジェクト、GPTs などで保存", "CLAUDE.md（ルール）と Skills（手順書）をファイルで保存し、フォルダごと共有"],
      ["外部サービスとの連携", "コネクタ（アプリ連携）で Google Drive 等と連携【要確認】", "MCP／コネクタで Google Drive・Notion 等と連携"],
      ["得意な業務", "Webサイトを巡回する調べもの、Web上の手続きや入力、普段使いの幅広い用途", "手元の原稿・データ・Webサイトのファイルを、まとめて読み・直し・作る作業"],
    ], { x: MX, y: 1.85, w: CW, colW: [2.2, 5.2, 4.73], rowH: [0.42, 1.12, 0.62, 0.62, 0.62, 0.72], fontSize: 12 });
    txt(s, [
      { text: "ChatGPT の強み：", options: { bold: true, color: C.text2 } },
      { text: "Web上の調べものや操作、画像生成など用途が幅広く、すでに使い慣れている方が多い点です。両方を目的に応じて使い分けるのが現実的です。※ChatGPT の機能は変更が多いため、最新は OpenAI 公式情報をご確認ください。" },
    ], { x: MX, y: 6.18, w: CW, h: 0.7, fontSize: 12, color: C.text1, margin: 0 });
    NOTE(s, "次に、ChatGPTにも「エージェント」と呼ばれる機能があります。Agentモードや、開発者向けのCodexです。公平にお伝えすると、ChatGPTのAgentモードはクラウド上の仮想のコンピューターで、Webサイトを調べたり操作したりするのが得意です。Codexには手元で動く版もありますが、主に開発者向けです。Claude Codeは手元のフォルダで動くのが基本なので、案件ファイルをアップロードせずにそのまま扱えるのが強みです。どちらが優れているかではなく、Web上の作業はChatGPT、手元のファイルを扱う作業はClaude Code、という使い分けをお勧めします。なお、ChatGPT側の仕様は変更が多いため、一部に要確認を付けています。");
  }

  // ======================= 7. メリット =======================
  section("貴社での活用");
  {
    const s = content("貴社で導入するメリット", "少人数・兼務の体制のまま、対応できる仕事の量と幅を広げ、品質とノウハウを会社に残せます");
    const items = [
      ["FaUsers", "対応範囲の拡大", "印刷・出版・Web・動画・イベントを少人数で兼務する体制でも、下ごしらえをAIに任せることで、人を増やさずに制作本数や事業の幅を広げられます。"],
      ["FaCheckCircle", "品質の標準化", "貴社の表記ルールやクライアントごとの決まり（社名表記、数字の書き方など）を覚えさせることで、誰が作業しても同じ基準でチェックできます。"],
      ["FaBook", "属人化の解消", "ベテランの方の校正の勘所や制作の段取りを Skills（手順書）として残せます。担当者が不在でも、新しい方でも同じ手順で作業できます。"],
      ["FaPlug", "既存ツール連携と管理機能", "Google Drive・Notion 等と連携可能です。法人プランでは、ユーザー管理や、入力データを学習に使わない条件（既定）などが用意されています。"],
    ];
    const cw = (CW - 0.4) / 2, ch = 2.15;
    for (let i = 0; i < 4; i++) {
      const x = MX + (i % 2) * (cw + 0.4), y = 1.85 + Math.floor(i / 2) * (ch + 0.3);
      card(s, x, y, cw, ch);
      await iconCircle(s, items[i][0], x + 0.3, y + 0.3, 0.85, C.text2, HEX.lt1);
      txt(s, items[i][1], { x: x + 1.4, y: y + 0.3, w: cw - 1.7, h: 0.45, fontSize: 17, bold: true, color: C.text1, margin: 0 });
      txt(s, items[i][2], { x: x + 1.4, y: y + 0.8, w: cw - 1.7, h: 1.3, fontSize: 13, color: C.text1, margin: 0 });
    }
    txt(s, "※効果は業務内容や使い方によって大きく変わります。まずは試験導入で、貴社の業務での効果を確かめることをお勧めします。", { x: MX, y: 6.6, w: CW, h: 0.3, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "貴社で導入するメリットを4つにまとめました。1つ目が一番大きなポイントで、少人数で多くの事業を兼務されている中で、下ごしらえをAIに任せることで、人を増やさずに対応できる範囲を広げられます。2つ目は品質の標準化です。表記ルールやクライアントごとの決まりを覚えさせれば、誰が作業しても同じ基準でチェックできます。3つ目は属人化の解消です。ベテランの方のノウハウを手順書として残せます。4つ目は、Google DriveやNotionなど既存のツールとの連携と、法人プランの管理機能です。ただし、効果は使い方次第ですので、まず試して確かめることをお勧めします。");
  }

  // ======================= 8. ユースケース① =======================
  {
    const s = content("ユースケース一覧①：印刷・出版・広告制作", "「読む・比べる・まとめる・書類にする」作業は、ほぼすべてAIの下ごしらえの対象になります");
    table(s, ["業務", "依頼文の例", "Before → After", "効果の目安（仮）"], [
      ["入稿前の原稿校正・表記ゆれチェック", "「このチラシ原稿5本を表記ルールでチェックし、修正案をExcelに」", "1本ずつ目視で確認 → AIの指摘一覧を見て採否を判断", "一次チェックの時間を半分程度に"],
      ["チラシ・広告のコピー案、見出し案", "「春の新生活セールのチラシ見出しを、トーン別に10案」", "白紙から考える → たたき台から選んで磨く", "案出しの時間を短縮"],
      ["出版物の原稿整理・目次や索引", "「この原稿フォルダから目次案と索引語の候補を作って」", "全章を読み返して手作業で拾う → 候補一覧を確認・修正", "拾い出し作業を大幅に短縮"],
      ["打ち合わせ議事録の作成", "「この文字起こしから、決定事項・ToDo・担当者を整理して」", "録音を聞き直してまとめる → 下書きを確認して共有", "1件30〜60分 → 10分程度の確認"],
      ["企画書・提案書（PowerPoint）", "「この議事録をもとに、提案書のたたき台を10枚で」", "構成から自分で作る → 構成と文章のたたき台を調整", "初稿づくりを短縮"],
      ["見積書・制作指示書の下書き", "「過去の見積Excelを参考に、名刺500部×3種の見積下書きを」", "過去案件を探して転記 → 下書きの金額・条件を確認", "転記・入力の手間を削減"],
    ], { x: MX, y: 1.85, w: CW, colW: [2.55, 3.85, 3.55, 2.18], rowH: [0.42, 0.72, 0.72, 0.72, 0.72, 0.72, 0.72], fontSize: 12 });
    txt(s, "※効果の目安は一般的な想定（仮置き）です。金額・数量・日付などの最終確認は必ず人が行ってください。", { x: MX, y: 6.6, w: CW, h: 0.3, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "ここからは具体的な使い道です。まず、貴社の中心である印刷・出版・広告制作の分野です。入稿前の原稿チェック、チラシの見出し案、出版物の目次や索引の候補づくり、打ち合わせの議事録、提案書のたたき台、見積書や制作指示書の下書きなど、「読む・比べる・まとめる・書類にする」作業は、ほぼすべてAIに下ごしらえを任せられます。右端の効果の目安はあくまで仮置きで、実際の効果は試験導入で測っていただく想定です。特に見積の金額や部数、納期などは、AIの下書きを必ず人が確認してください。");
  }

  // ======================= 9. ユースケース② =======================
  {
    const s = content("ユースケース一覧②：Web・動画・イベント・コワーキング運営", "サブ事業や日々の運営業務まで、同じ道具一つで幅広く下ごしらえを任せられます");
    table(s, ["業務", "依頼文の例", "Before → After", "効果の目安（仮）"], [
      ["Webサイトの更新・修正", "「お知らせページに年末年始の営業案内を追加して」", "HTMLを手で編集 → AIが直接編集し、人が表示を確認", "小さな更新を短時間で"],
      ["動画の構成案・台本・字幕", "「この企画メモから60秒の店舗紹介動画の構成と台本を」", "白紙から構成 → たたき台を調整（字幕文も同時に）", "構成・台本の初稿を短縮"],
      ["イベントの企画書・進行表", "「来月の地域イベントの運営マニュアルと進行表を作って」", "過去資料を探して作り直す → 過去資料を参考に下書き", "書類一式の作成を短縮"],
      ["地域情報の記事・SNS投稿文", "「この取材メモから記事と、Instagram用の投稿文3案を」", "媒体ごとに書き分け → 1つの素材から複数分を下書き", "書き分けの手間を削減"],
      ["コワーキングの案内文・告知文", "「SecondDeskの会員向けに、交流イベントの告知文を」", "毎回ゼロから作成 → 型に沿って下書き", "定型文の作成を短縮"],
      ["利用データ（CSV）の集計・グラフ化", "「この利用データCSVを月別に集計して、Excelでグラフに」", "手作業で集計 → 集計表とグラフをExcelで出力", "月次集計を短時間で"],
      ["案件ファイルの整理・リネーム", "「このフォルダの画像を『日付_案件名_連番』に名前変更して」", "1つずつ手作業 → 一括で整理（実行前に確認）", "単純作業をほぼ自動化"],
      ["よく使う作業の定型化（Skills）", "「今の校正手順を、次回から使えるSkillにまとめて」", "毎回指示を書く → 一言で同じ手順を再現", "2回目以降の指示の手間を削減"],
    ], { x: MX, y: 1.75, w: CW, colW: [2.55, 3.85, 3.55, 2.18], rowH: [0.36, 0.54, 0.54, 0.54, 0.54, 0.54, 0.54, 0.54, 0.54], fontSize: 11 });
    txt(s, "※会員情報・個人情報を含むデータは、扱うフォルダを分けるなどの運用ルールを決めてから使うことをお勧めします（スライド16参照）。", { x: MX, y: 6.68, w: CW, h: 0.3, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "次はWeb、動画、イベント、地域情報の発信、そしてコワーキングスペースの運営です。Webサイトの小さな更新はHTMLのファイルを直接編集できますし、動画の台本や字幕テキスト、イベントの運営マニュアルや進行表、記事とSNS投稿の書き分けなども任せられます。SecondDeskの運営では、会員向けの案内文や、利用データの集計とグラフ化にも使えます。ファイル名の一括変更のような単純作業も得意です。最後の「Skills」は、一度うまくいった手順を保存して、次から一言で再現できるようにする仕組みです。会員情報など個人情報を扱う場合は、先に運用ルールを決めておくことをお勧めします。");
  }

  // ======================= 10. 深掘り① =======================
  {
    const s = content("深掘り①：入稿前の原稿校正・表記統一", "AIが「指摘と修正案」まで、人が「最終判断」。一次チェックの時間を減らし、見落としも防ぎます");
    // 左：流れ（縦）
    const fx = MX, fw = 4.0;
    const flow = [["表記ルールを登録", "貴社・クライアントの表記ルールを CLAUDE.md や Skill に書いておく（初回のみ）"], ["フォルダ内を一括チェック", "案件フォルダの複数原稿をAIが順に読み、表記ゆれ・誤字・ルール違反を探す"], ["修正案の一覧を出力", "ファイル名・行・原文・修正案・理由を Excel にまとめて保存"], ["人が最終確認", "担当者が一覧を見て採否を判断し、原稿に反映（反映もAIに指示可能）"]];
    for (let i = 0; i < 4; i++) {
      const y = 1.85 + i * 1.22;
      numCircle(s, i + 1, fx, y + 0.05, 0.5, i === 3 ? C.accent1 : C.text2);
      txt(s, flow[i][0], { x: fx + 0.65, y, w: fw - 0.65, h: 0.36, fontSize: 14, bold: true, color: C.text1, margin: 0 });
      txt(s, flow[i][1], { x: fx + 0.65, y: y + 0.38, w: fw - 0.65, h: 0.75, fontSize: 12, color: C.text1, margin: 0 });
    }
    // 右上：依頼文
    const rx = MX + fw + 0.4, rw = CW - fw - 0.4;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rx, y: 1.85, w: rw, h: 1.05, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.25 }, rectRadius: 0.1, objectName: "prompt" });
    txt(s, [{ text: "依頼文の例　", options: { bold: true, color: C.accent1, fontSize: 12 } }, { text: "「『○○不動産様_春チラシ』フォルダの原稿を、クライアント表記ルールでチェックして。修正案は『校正チェック.xlsx』に一覧にして。原稿そのものはまだ書き換えないで」", options: { fontSize: 13, color: C.text1 } }], { x: rx + 0.2, y: 1.9, w: rw - 0.4, h: 0.95, valign: "middle", margin: 0 });
    // 右中：出力イメージ（表）
    txt(s, "出力イメージ（校正チェック.xlsx）", { x: rx, y: 3.05, w: rw, h: 0.32, fontSize: 12, bold: true, color: C.text2, margin: 0 });
    table(s, ["ファイル", "行", "原文", "修正案", "理由"], [
      ["表面.docx", "12", "お問合せ", "お問い合わせ", "送り仮名ルール"],
      ["表面.docx", "18", "１０％OFF", "10%OFF", "数字は半角"],
      ["裏面.docx", "5", "（株）○○", "株式会社○○", "社名は正式表記"],
    ], { x: rx, y: 3.4, w: rw, colW: [1.45, 0.55, 1.75, 1.75, rw - 5.5], rowH: [0.34, 0.34, 0.34, 0.34], fontSize: 11, boldFirst: false });
    // 右下：Before/After
    const by = 5.0, bw = (rw - 0.3) / 2;
    card(s, rx, by, bw, 1.25);
    card(s, rx + bw + 0.3, by, bw, 1.25, "FCEDE9");
    txt(s, [{ text: "Before（例）", options: { fontSize: 12, bold: true, color: C.accent3, breakLine: true } }, { text: "約3時間", options: { fontSize: 26, bold: true, color: C.text1, breakLine: true } }, { text: "原稿10本を1本ずつ目視で確認", options: { fontSize: 11, color: C.text1 } }], { x: rx + 0.2, y: by + 0.08, w: bw - 0.4, h: 1.1, margin: 0, lineSpacing: null });
    txt(s, [{ text: "After（例）", options: { fontSize: 12, bold: true, color: C.accent1, breakLine: true } }, { text: "約1時間", options: { fontSize: 26, bold: true, color: C.text1, breakLine: true } }, { text: "AIの一覧作成は数分＋人の確認", options: { fontSize: 11, color: C.text1 } }], { x: rx + bw + 0.5, y: by + 0.08, w: bw - 0.4, h: 1.1, margin: 0, lineSpacing: null });
    txt(s, "※時間は仮の例です。AIは見落としや誤った指摘をすることもあるため、色校正・文字校正の最終責任は人が持ちます。", { x: MX, y: 6.5, w: CW, h: 0.35, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "1つ目の深掘りは、入稿前の原稿校正と表記統一です。まず、貴社やクライアントの表記ルールを、CLAUDE.mdというルールファイルやSkillとして一度書いておきます。あとは「このフォルダの原稿をチェックして、修正案を一覧にして」と頼むと、AIがフォルダ内の原稿を順に読み、表記ゆれや誤字を探して、ファイル名、行、原文、修正案、理由をExcelにまとめます。人はその一覧を見て、採用するかどうかを判断します。「原稿そのものはまだ書き換えないで」と指示しておけば、勝手に原稿が変わることはありません。時間は仮の例ですが、10本で3時間かかっていた一次チェックが1時間程度になるイメージです。AIは指摘と修正案まで、最終判断は人、という役割分担が大前提です。");
  }

  // ======================= 11. 深掘り② =======================
  {
    const s = content("深掘り②：打ち合わせから制作準備まで", "1つの依頼で、議事録・制作指示書・見積下書き・共有用の要約まで、書類を「つなげて」作れます");
    const flow = [["FaMicrophone", "文字起こし", "録音を文字起こししたテキスト（※）"], ["FaClipboardList", "議事録", "決定事項・ToDo・担当者・期限を整理"], ["FaFileWord", "制作指示書", "仕様・サイズ・部数・納期・入稿データの一覧"], ["FaFileExcel", "見積の下書き", "過去の見積を参考に、項目と数量を入力"], ["FaShareAlt", "社内共有", "要点の要約と、共有フォルダへの保存"]];
    const fw = 2.05, fg = (CW - fw * 5) / 4, fy = 1.85;
    for (let i = 0; i < 5; i++) {
      const x = MX + i * (fw + fg);
      card(s, x, fy, fw, 2.0);
      await iconCircle(s, flow[i][0], x + (fw - 0.7) / 2, fy + 0.2, 0.7, i === 0 ? C.accent3 : C.text2, HEX.lt1);
      txt(s, flow[i][1], { x: x + 0.1, y: fy + 1.0, w: fw - 0.2, h: 0.35, fontSize: 14, bold: true, color: C.text1, align: "center", margin: 0 });
      txt(s, flow[i][2], { x: x + 0.15, y: fy + 1.36, w: fw - 0.3, h: 0.6, fontSize: 11, color: C.text1, margin: 0 });
      if (i < 4) arrow(s, x + fw + fg / 2 - 0.14, fy + 0.85, 0.28, 0.3, C.accent1);
    }
    // 依頼文
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 4.1, w: 7.4, h: 2.15, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.25 }, rectRadius: 0.1, objectName: "prompt" });
    txt(s, [
      { text: "依頼文の例", options: { bold: true, color: C.accent1, fontSize: 12, breakLine: true } },
      { text: "「『△△商店様_周年イベント』フォルダの打ち合わせ文字起こしから、①議事録（決定事項・ToDo・担当者）、②制作指示書（ポスターB2・チラシA4・SPグッズ）、③過去の見積を参考にした見積の下書きを作って。最後に、社内共有用の5行要約もお願いします」", options: { fontSize: 13, color: C.text1 } },
    ], { x: MX + 0.25, y: 4.2, w: 6.9, h: 1.95, margin: 0 });
    // Before/After
    const bx = MX + 7.7, bw = CW - 7.7;
    card(s, bx, 4.1, bw, 1.0);
    card(s, bx, 5.25, bw, 1.0, "FCEDE9");
    txt(s, [{ text: "Before（例）　", options: { fontSize: 12, bold: true, color: C.accent3 } }, { text: "約3時間", options: { fontSize: 22, bold: true, color: C.text1, breakLine: true } }, { text: "議事録・指示書・見積を別々に作成", options: { fontSize: 11, color: C.text1 } }], { x: bx + 0.25, y: 4.15, w: bw - 0.5, h: 0.9, valign: "middle", margin: 0, lineSpacing: null });
    txt(s, [{ text: "After（例）　", options: { fontSize: 12, bold: true, color: C.accent1 } }, { text: "約1時間", options: { fontSize: 22, bold: true, color: C.text1, breakLine: true } }, { text: "下書き一式の確認・修正が中心に", options: { fontSize: 11, color: C.text1 } }], { x: bx + 0.25, y: 5.3, w: bw - 0.5, h: 0.9, valign: "middle", margin: 0, lineSpacing: null });
    txt(s, "※文字起こしは、録音の文字起こし機能を持つ別のツールで用意する想定です【要確認：使用ツール】。時間は仮の例で、金額・納期は必ず人が確認します。", { x: MX, y: 6.42, w: CW, h: 0.45, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "2つ目の深掘りは、打ち合わせから制作準備までの流れです。ここがエージェントらしさが一番伝わる部分です。打ち合わせの文字起こしを渡し、「議事録、制作指示書、見積の下書き、社内共有用の要約まで作って」と一度頼むだけで、AIが順番に書類を作ってフォルダに保存します。チャットでは書類ごとに貼り付けと指示を繰り返す必要がありますが、Claude Codeは前の書類の内容を引き継いで、次の書類につなげて作れます。時間は仮の例ですが、3時間かかっていた作業が、下書きの確認と修正を中心とした1時間程度になるイメージです。なお、録音から文字起こしをする部分は、別のツールを使う想定です。");
  }

  // ======================= 12. 最近の主な機能 =======================
  section("機能と費用");
  {
    const s = content("最近の主な機能", "ルールを覚える・手順を型にする・外部とつながる仕組みがそろい、業務に組み込みやすくなっています");
    const feats = [
      ["FaFileAlt", "CLAUDE.md", "ルールを覚えさせる", "表記ルールや作業の決まりを書いたファイル。毎回の作業の最初に自動で読み込まれます"],
      ["FaBook", "Skills", "業務の型化", "「校正の手順」などを手順書として保存。一言、または必要な場面で自動的に使われます"],
      ["FaSitemap", "サブエージェント", "作業の分担", "調べもの専門などの「助手」に作業を分け、本体は要点だけを受け取ります"],
      ["FaPlug", "MCP／コネクタ", "外部サービス連携", "外部サービスとつなぐ共通の規格（差し込み口）。Google Drive・Notion 等と連携できます"],
      ["FaMap", "Plan モード", "先に計画だけ確認", "ファイルを変更せずに、作業の計画だけを提案させ、人が確認してから実行できます"],
      ["FaUserCheck", "許可の設定", "勝手に動かない", "ファイル変更やコマンド実行の前に確認を求めるかどうかを、段階的に設定できます"],
      ["FaDesktop", "デスクトップアプリ", "画面操作で使える", "Windows／Mac のアプリの「Code」タブから利用。コマンド入力は不要です"],
      ["FaCloud", "Web版", "PCを閉じても進む", "クラウド上で作業を実行。主にGitHub上のファイルが対象で、開発・Web制作向け"],
    ];
    const cw = (CW - 0.3 * 3) / 4, ch = 2.25;
    for (let i = 0; i < 8; i++) {
      const x = MX + (i % 4) * (cw + 0.3), y = 1.8 + Math.floor(i / 4) * (ch + 0.22);
      card(s, x, y, cw, ch);
      await iconCircle(s, feats[i][0], x + 0.2, y + 0.2, 0.62, C.text2, HEX.lt1);
      txt(s, feats[i][1], { x: x + 0.95, y: y + 0.18, w: cw - 1.1, h: 0.35, fontSize: 13, bold: true, color: C.text1, margin: 0 });
      txt(s, feats[i][2], { x: x + 0.95, y: y + 0.52, w: cw - 1.1, h: 0.32, fontSize: 12, bold: true, color: C.accent1, margin: 0 });
      txt(s, feats[i][3], { x: x + 0.2, y: y + 0.95, w: cw - 0.4, h: 1.25, fontSize: 12, color: C.text1, margin: 0 });
    }
    txt(s, `※${AS_OF}に公式ドキュメントで確認した内容の概要です。詳細・最新情報は https://code.claude.com/docs をご確認ください。`, { x: MX, y: 6.62, w: CW, h: 0.3, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "最近の主な機能を、何がうれしいかと合わせてご紹介します。CLAUDE.mdはルールを覚えさせるファイルで、作業のたびに最初に読み込まれます。Skillsは業務の手順書です。サブエージェントは作業を分担する助手です。MCPやコネクタは、外部のサービスとつなぐ共通の差し込み口のようなもので、Google DriveやNotionと連携できます。Planモードは、ファイルを変更せずに計画だけを先に見せてもらう機能で、慣れないうちは安心です。許可の設定で、勝手に動かないようにもできます。使い方は、デスクトップアプリとWeb版があります。Web版は主に開発やWeb制作向けです。ここに載せたのは公式ドキュメントで確認できたものの概要で、詳細は公式ページをご覧ください。");
  }

  // ======================= 13. 料金プラン =======================
  {
    const s = content("料金プランの概要", "少人数の貴社には、まず2〜3名で Team（Standard席）または Pro から始める形が合っています");
    table(s, ["プラン", "対象", "月額の目安（1人あたり）", "Claude Code", "最低人数", "特徴"], [
      ["Pro", "個人", "$20（年払いで$17相当）", "○", "1名", "まず個人で試すのに最適"],
      ["Max", "個人（多く使う方）", "$100（Proの5倍の利用量）／上位$200【要確認：名称・倍率】", "○", "1名", "長時間・大量の作業向け"],
      ["Team Standard", "法人", "$25（年払いで$20）", "○", "2〜150名", "管理機能。学習利用なし（既定）"],
      ["Team Premium", "法人", "$125（年払いで$100）", "○", "同上", "Standardの5倍の量。席の混在可"],
      ["Enterprise", "法人（大規模）", "席料$20（年払い）＋利用量（API料金）", "○", "20名〜", "SSO・監査ログ等の高度な管理"],
      ["API 従量課金", "開発・自動化", "使った分だけ（例：Sonnet 5.5 は100万トークンあたり入力$2・出力$10）", "○", "—", "利用量で費用が変動する"],
      [{ text: "参考：ChatGPT Business", options: { color: HEX.accent3 } }, "法人", "$25（年払いで$20）【要確認】", "Codex含む【要確認】", "2名〜【要確認】", "OpenAI の法人プラン"],
    ], { x: MX, y: 1.8, w: CW, colW: [2.05, 1.4, 3.75, 1.15, 1.5, 2.28], rowH: [0.36, 0.46, 0.56, 0.56, 0.46, 0.46, 0.56, 0.5], fontSize: 11 });
    card(s, MX, 6.08, CW, 0.82);
    await iconCircle(s, "FaLightbulb", MX + 0.15, 6.16, 0.66, C.accent1, HEX.lt1);
    txt(s, [
      { text: "始め方の例：", options: { bold: true, color: C.text2 } },
      { text: "試験導入は Team Standard を2〜3名・月払い（月$50〜75）で開始。クライアント原稿を扱うため、学習利用なし（既定）と管理機能がある法人プランをお勧めします。" },
      { text: `\n${AS_OF}・税抜・USD。最新の料金は公式ページ（https://claude.com/pricing）をご確認ください。`, options: { fontSize: 10, color: C.accent3 } },
    ], { x: MX + 1.0, y: 6.1, w: CW - 1.15, h: 0.78, fontSize: 12, color: C.text1, valign: "middle", margin: 0, lineSpacing: null });
    NOTE(s, "料金の概要です。数字は2026年11月時点の公式ページを元にした税抜・米ドルの目安で、変更されることがあります。個人向けのProは月20ドル、たくさん使う方向けのMaxは月100ドルからです。法人向けのTeamは2名から契約でき、Standard席が月25ドル、年払いなら20ドルです。より多く使う方にはPremium席があり、席の種類は混在できます。Enterpriseは20名からの大規模向けです。API従量課金は使った分だけですが、費用が読みにくいので、最初は定額のプランをお勧めします。参考としてChatGPTの法人プランも並べましたが、こちらは公式ページで確認しきれていないため要確認です。貴社の場合、クライアントの未公開原稿を扱うため、入力内容が学習に使われない法人プランのTeamで、2〜3名から始めるのがよいと考えます。");
  }

  // ======================= 14. 費用対効果 =======================
  {
    const s = content("費用対効果の試算（イメージ）", "仮置きの前提でも、月額費用に対して数倍の時間削減効果が見込めます。実数は試験導入で測ります");
    // 左：前提
    const lw = 5.6;
    txt(s, "試算の前提（すべて仮置き・貴社の実数に差し替え）", { x: MX, y: 1.8, w: lw, h: 0.35, fontSize: 14, bold: true, color: C.text2, margin: 0 });
    table(s, ["項目", "仮の値"], [
      ["利用人数", "3名"],
      ["プラン", "Team Standard（月払い $25／人）"],
      ["為替", "1ドル＝150円（仮）"],
      ["月額費用", "$75 ≒ 約11,250円（税抜）"],
      ["削減時間（1人・月）", "校正4h＋議事録・制作準備3h＋資料作成3h＝10h"],
      ["時間単価", "3,000円／時間（人件費の目安・仮）"],
    ], { x: MX, y: 2.2, w: lw, colW: [2.0, lw - 2.0], rowH: [0.38, 0.48, 0.48, 0.48, 0.48, 0.62, 0.48], fontSize: 12 });
    txt(s, "計算式：3名 × 10時間 × 3,000円 ＝ 月90,000円相当\n控えめに半分（1人5時間）でも 月45,000円相当", { x: MX, y: 5.75, w: lw, h: 0.8, fontSize: 13, bold: true, color: C.text1, margin: 0 });
    // 右：チャート
    const cx = MX + lw + 0.4, cwid = CW - lw - 0.4;
    s.addChart(pres.charts.BAR, [{ name: "金額（円／月）", labels: ["月額費用", "削減効果（控えめ：5h／人）", "削減効果（想定：10h／人）"], values: [11250, 45000, 90000] }], {
      x: cx, y: 1.8, w: cwid, h: 4.3, barDir: "bar",
      chartColors: [HEX.accent3, HEX.dk2, HEX.accent1],
      showTitle: true, title: "月額費用と削減効果（金額換算）の比較", titleFontSize: 14, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0\"円\"", dataLabelFontSize: 12, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
      catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
      valAxisHidden: true, valAxisMaxVal: 110000, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      showLegend: false, barGapWidthPct: 60,
    });
    txt(s, "※最初の1か月は操作に慣れる期間のため、効果は小さめに見込むのが現実的です。削減時間は試験導入時に「作業前後の時間」を記録して実測することをお勧めします。", { x: cx, y: 6.15, w: cwid, h: 0.7, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "費用対効果の試算です。ここに置いた数字はすべて仮置きですので、貴社の実際の人数や時間単価に差し替えてご覧ください。例として、3名がTeamのStandard席を月払いで使うと月75ドル、1ドル150円で約1万1千円です。1人あたり、校正で4時間、議事録や制作準備で3時間、資料作成で3時間、合計月10時間の削減ができたとすると、時間単価3千円で月9万円相当になります。控えめに半分の5時間としても4万5千円相当です。ただし、最初の1か月は慣れる期間ですので、効果は小さめに見込んでください。試験導入では、作業の前後の時間を記録して、実際の効果を測ることをお勧めします。");
  }

  // ======================= 15. セットアップ =======================
  section("導入にあたって");
  {
    const s = content("セットアップ方法（Windows）：2つの使い方", "デザイナー・編集者の方はデスクトップアプリ、Web制作や自動化の担当の方はターミナル版がお勧めです");
    const colW = (CW - 2.2 - 0.0) / 2;
    table(s, ["", "デスクトップアプリ", "ターミナル（CLI）"], [
      ["向いている人", "デザイナー・編集者・営業など、画面操作で使いたい方", "Web制作の担当者、細かい自動化をしたい方"],
      ["使い方のイメージ", "アプリの「Code」タブで作業フォルダを選び、話しかけるように指示", "黒い画面（PowerShell 等）で案件フォルダに移動し、claude と入力して指示"],
      ["手順の概要", "① 公式サイトから Windows 版をダウンロード\n② インストーラーを実行\n③ 有料プランのアカウントでサインイン\n④「Code」タブを開き、作業フォルダを選ぶ", "① PowerShell でインストール用コマンドを実行\n② 新しい画面で claude --version を入力して確認\n③ 案件フォルダで claude と入力して起動\n④ 初回はブラウザでログイン"],
      ["難易度", "★☆☆　やさしい（コマンド入力なし）", "★★☆　ふつう（コマンドを数行入力）"],
      ["必要なもの", "有料プラン（Pro／Max／Team など）", "有料プラン。Windows 10（1809以降）、メモリ4GB以上。Git for Windows は推奨（必須ではない）"],
    ], { x: MX, y: 1.8, w: CW, colW: [2.2, colW, colW], rowH: [0.4, 0.62, 0.72, 1.55, 0.5, 0.72], fontSize: 12 });
    await iconCircle(s, "FaInfoCircle", MX, 6.42, 0.45, C.background2, HEX.dk2);
    txt(s, "どちらも同じアカウント・同じプランで使えます。詳細な手順は付録（スライド18）と公式ドキュメント https://code.claude.com/docs/ja/setup をご確認ください。", { x: MX + 0.6, y: 6.4, w: CW - 0.6, h: 0.5, fontSize: 12, color: C.text1, valign: "middle", margin: 0 });
    NOTE(s, "セットアップ方法です。Windowsでは2つの使い方があります。デスクトップアプリは、普段のアプリと同じように画面操作で使えるので、デザイナーや編集者の方に向いています。ダウンロードしてインストールし、サインインして「Code」タブを開くだけです。ターミナル版は、黒い画面にコマンドを入力して使う方法で、Web制作の担当の方や、細かい自動化をしたい方に向いています。インストールはコマンドを1行入力するだけで、管理者の権限は不要です。どちらも有料プランが必要で、同じアカウントで両方使えます。試験導入では、まずデスクトップアプリから始めるのがお勧めです。");
  }

  // ======================= 16. セキュリティ =======================
  {
    const s = content("セキュリティと安全な使い方", "仕組み（許可・管理・学習利用なし）と運用ルール（フォルダ分け・人の最終確認）の両輪で安全に使えます");
    const items = [
      ["FaHandPaper", "実行前に許可を求める", "ファイルの変更やコマンドの実行の前に「よいですか？」と確認を求めます。初めは確認ありの設定で使うことをお勧めします。"],
      ["FaFolder", "作業フォルダを分ける", "AIに開かせるのは、作業用にコピーした案件フォルダだけにします。マスターデータや他案件のフォルダとは分けて運用します。"],
      ["FaUserShield", "未公開原稿・個人情報の扱い", "クライアントの未公開原稿、顧客の個人情報、コワーキング会員情報は、扱ってよい範囲を社内ルールで決めてから使います。"],
      ["FaDatabase", "入力データの学習利用", "Team・Enterprise などの法人向け契約では、既定で学習に使われません。個人向け（Pro／Max）は設定で選べるため、確認が必要です。"],
      ["FaUserCog", "管理者による利用制御", "Team では管理者がメンバーの追加・削除や利用状況の確認、コネクタの管理などを行えます（詳細は公式ページ）。"],
      ["FaEye", "最終成果物は人が確認", "AIは誤ることがあります。入稿データ、見積金額、Webの公開前の内容は、必ず人が確認してから外に出す運用にします。"],
    ];
    const cw = (CW - 0.3 * 2) / 3, ch = 2.2;
    for (let i = 0; i < 6; i++) {
      const x = MX + (i % 3) * (cw + 0.3), y = 1.8 + Math.floor(i / 3) * (ch + 0.25);
      card(s, x, y, cw, ch);
      await iconCircle(s, items[i][0], x + 0.25, y + 0.25, 0.65, i === 5 ? C.accent1 : C.text2, HEX.lt1);
      txt(s, items[i][1], { x: x + 1.05, y: y + 0.3, w: cw - 1.25, h: 0.6, fontSize: 15, bold: true, color: C.text1, valign: "middle", margin: 0 });
      txt(s, items[i][2], { x: x + 0.25, y: y + 1.0, w: cw - 0.5, h: 1.15, fontSize: 12, color: C.text1, margin: 0 });
    }
    txt(s, "※データの取り扱い・保存期間の詳細は公式ドキュメント（https://code.claude.com/docs/ja/data-usage）をご確認ください。", { x: MX, y: 6.65, w: CW, h: 0.28, fontSize: 11, color: C.accent3, margin: 0 });
    NOTE(s, "セキュリティと安全な使い方です。仕組みの面では、AIはファイルの変更などの前に許可を求めますし、法人向けの契約では入力した内容が既定でAIの学習に使われません。個人向けのProやMaxでは設定で選べるので、確認が必要です。管理者がメンバーや利用状況を管理することもできます。運用の面では、AIに開かせるフォルダを作業用に分けること、クライアントの未公開原稿や会員情報など個人情報の扱いを先にルール化すること、そして最終成果物は必ず人が確認してから外に出すこと、この3点をお勧めします。");
  }

  // ======================= 17. ご提案 =======================
  {
    const s = content("ご提案：導入ステップと次の一歩", "まずは1〜2か月、2業務に絞った少人数の試験導入から。効果を実測してから広げる進め方をご提案します");
    const ph = [
      ["フェーズ1", "試験導入", "1〜2か月", ["2〜3名で開始", "対象は「原稿校正」と「議事録・制作準備」の2業務", "作業前後の時間を記録"]],
      ["フェーズ2", "整備と展開", "3〜6か月目", ["表記ルール（CLAUDE.md）と Skills を整備", "Web更新・動画・イベント業務へ展開", "社内の使い方ルールを文書化"]],
      ["フェーズ3", "本格導入", "7か月目〜", ["利用人数・プランを見直し", "効果測定（削減時間・制作本数）", "定着した業務の型を全員で共有"]],
    ];
    const pw = (CW - 0.6) / 3, py = 1.8, phh = 2.75;
    for (let i = 0; i < 3; i++) {
      const x = MX + i * (pw + 0.3);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: py, w: pw, h: 0.75, fill: { color: i === 0 ? C.accent1 : C.text2 }, line: { type: "none" }, rectRadius: 0.1, objectName: "phase-head" });
      txt(s, [{ text: ph[i][0] + "　", options: { fontSize: 12 } }, { text: ph[i][1], options: { fontSize: 17, bold: true } }, { text: "　" + ph[i][2], options: { fontSize: 12 } }], { x: x + 0.2, y: py, w: pw - 0.4, h: 0.75, color: C.background1, valign: "middle", margin: 0, lineSpacing: null });
      card(s, x, py + 0.85, pw, phh - 0.85);
      txt(s, ph[i][3].map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < 2 } })), { x: x + 0.25, y: py + 1.0, w: pw - 0.5, h: phh - 1.1, fontSize: 13, color: C.text1, paraSpaceAfter: 6, margin: 0 });
    }
    // 次の一歩
    const ny = 4.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: ny, w: CW, h: 2.05, fill: { color: "FCEDE9" }, line: { type: "none" }, rectRadius: 0.1, objectName: "next" });
    await iconCircle(s, "FaHandshake", MX + 0.3, ny + 0.3, 0.8, C.accent1, HEX.lt1);
    txt(s, "次の一歩として、ご検討いただきたいこと", { x: MX + 1.3, y: ny + 0.25, w: CW - 1.6, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0 });
    txt(s, [
      { text: "試験導入の「メンバー（2〜3名）」「対象業務（2つ）」「期間（1〜2か月）」のご検討", options: { bullet: { type: "number" }, breakLine: true } },
      { text: "実際の業務に近い題材（過去のチラシ原稿など）を使った、30〜60分のデモ・ハンズオンの実施", options: { bullet: { type: "number" }, breakLine: true } },
      { text: "選択肢として、SecondDesk 北大路堀川を会場にした体験会（社内向け、または会員のクリエイター向け）の開催", options: { bullet: { type: "number" } } },
    ], { x: MX + 1.3, y: ny + 0.72, w: CW - 1.6, h: 1.25, fontSize: 13, color: C.text1, paraSpaceAfter: 3, margin: 0 });
    NOTE(s, "最後にご提案です。3つのフェーズで進めることをお勧めします。フェーズ1は1〜2か月の試験導入で、2〜3名、対象業務を「原稿校正」と「議事録・制作準備」の2つに絞ります。作業の前後の時間を記録しておくと、効果を判断しやすくなります。フェーズ2で表記ルールやSkillsを整え、Web、動画、イベントの業務に広げます。フェーズ3で本格導入と効果測定を行います。次の一歩として、試験導入のメンバー、対象業務、期間をご検討いただけますと幸いです。また、過去の原稿などを使ったデモやハンズオンもご用意できます。SecondDeskを会場にした体験会は、社内向けだけでなく会員のクリエイターの方向けのイベントにもなり得るので、選択肢の一つとしてご提案します。");
  }

  // ======================= 18. 付録 =======================
  section("付録");
  {
    const s = content("付録：用語集・Windows での詳しいセットアップ手順・参考リンク", "わからない言葉や詳しい手順は、このページと公式ドキュメントをご参照ください");
    const c1 = 4.3, c2 = 4.6, c3 = CW - c1 - c2 - 0.6, y = 1.8, h = 5.05;
    card(s, MX, y, c1, h);
    card(s, MX + c1 + 0.3, y, c2, h);
    card(s, MX + c1 + c2 + 0.6, y, c3, h);
    const head = (t, x, w) => txt(s, t, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.35, fontSize: 14, bold: true, color: C.text2, margin: 0 });
    head("用語集", MX, c1);
    head("Windows での詳しいセットアップ手順", MX + c1 + 0.3, c2);
    head("参考リンク", MX + c1 + c2 + 0.6, c3);
    const terms = [
      ["エージェント", "目的を伝えると、手順を考えて作業まで進めるAI"],
      ["ローカル環境", "自分のパソコンの中（手元のフォルダ）"],
      ["CLI／ターミナル", "文字（コマンド）で操作する黒い画面"],
      ["CLAUDE.md", "AIに守ってほしいルールを書くファイル"],
      ["Skills", "業務の手順書。一言で同じ作業を再現"],
      ["MCP", "外部サービスとつなぐ共通の差し込み口"],
      ["トークン", "AIが文章を数える単位。API料金の基準"],
      ["Plan モード", "変更せず、計画だけ先に見せるモード"],
    ];
    txt(s, terms.flatMap((t, i) => [{ text: t[0] + "：", options: { bold: true } }, { text: t[1], options: { breakLine: i < terms.length - 1 } }]), { x: MX + 0.2, y: y + 0.6, w: c1 - 0.4, h: h - 0.75, fontSize: 11, color: C.text1, paraSpaceAfter: 5, margin: 0, lineSpacing: 16 });
    const setup = [
      { text: "A. デスクトップアプリ", options: { bold: true, breakLine: true } },
      { text: "① 公式サイトから Windows 版（x64／ARM64）をダウンロード", options: { breakLine: true } },
      { text: "② インストーラーを実行 → Claude を起動してサインイン", options: { breakLine: true } },
      { text: "③「Code」タブ → 作業フォルダを選んで開始", options: { breakLine: true } },
      { text: "B. ターミナル（PowerShell）", options: { bold: true, breakLine: true } },
      { text: "① 次のコマンドを実行（管理者権限は不要）", options: { breakLine: true } },
      { text: "irm https://claude.ai/install.ps1 | iex", options: { fontFace: "Roboto Mono", color: C.accent1, breakLine: true } },
      { text: "　（WinGet の場合：winget install Anthropic.ClaudeCode）", options: { breakLine: true } },
      { text: "② 新しい画面で claude --version を実行して確認", options: { breakLine: true } },
      { text: "③ cd で案件フォルダへ移動し、claude と入力", options: { breakLine: true } },
      { text: "④ 初回はブラウザが開くのでログイン", options: { breakLine: true } },
      { text: "※ Git for Windows は推奨（必須ではない）", options: {} },
    ];
    txt(s, setup, { x: MX + c1 + 0.5, y: y + 0.6, w: c2 - 0.4, h: h - 0.75, fontSize: 11, color: C.text1, paraSpaceAfter: 3, margin: 0, lineSpacing: 16 });
    const links = [
      ["公式ドキュメント（日本語）", "https://code.claude.com/docs/ja"],
      ["セットアップ手順", "https://code.claude.com/docs/ja/setup"],
      ["デスクトップアプリ", "https://code.claude.com/docs/ja/desktop"],
      ["データの取り扱い", "https://code.claude.com/docs/ja/data-usage"],
      ["料金ページ", "https://claude.com/pricing"],
      ["ChatGPT の料金（参考）", "https://chatgpt.com/pricing"],
    ];
    txt(s, links.flatMap((l, i) => [{ text: l[0], options: { bold: true, breakLine: true } }, { text: l[1], options: { hyperlink: { url: l[1] }, color: C.accent1, breakLine: i < links.length - 1 } }]), { x: MX + c1 + c2 + 0.8, y: y + 0.6, w: c3 - 0.4, h: h - 0.75, fontSize: 11, color: C.text1, paraSpaceAfter: 5, margin: 0, lineSpacing: 16 });
    NOTE(s, "付録です。用語集、Windowsでの詳しいセットアップ手順、参考リンクをまとめました。会議後に読み返す際や、実際に試していただく際にご参照ください。手順や料金は変更されることがあるため、最新の情報は公式ドキュメントと料金ページをご確認ください。本日はありがとうございました。");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
