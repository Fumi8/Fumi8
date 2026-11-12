// ② 1枚サマリー（A4横）
const pptxgen = require("pptxgenjs");
const { THEME, HEX, icon, txt } = require("./common");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");
const OUT = process.argv[2] || "summary.pptx";

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name: "A4L", width: 11.69, height: 8.27 });
  pres.layout = "A4L";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Claude Code 業務活用のご提案（1枚サマリー）";
  pres.author = "【作成者名】";
  const C = pres.SchemeColor;
  const W = 11.69, MX = 0.45, CW = W - MX * 2;

  pres.defineSlideMaster({
    title: "ONEPAGER",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: MX, y: 0.3, w: 7.3, h: 0.55, fontSize: 22, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    ],
  });
  pres.addSection({ title: "サマリー" });
  const s = pres.addSlide({ masterName: "ONEPAGER", sectionTitle: "サマリー" });
  s.addText("Claude Code 業務活用のご提案　1枚サマリー", { placeholder: "title" });
  txt(s, "サンケイデザイン株式会社様\n2026年11月｜作成：【作成者名】", { x: W - MX - 3.3, y: 0.3, w: 3.3, h: 0.55, fontSize: 10, color: C.accent3, align: "right", valign: "middle", margin: 0 });

  const card = (x, y, w, h, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill || C.background2 }, line: { type: "none" }, rectRadius: 0.08, objectName: "card" });
  const head = (t, x, y, w) => txt(s, t, { x, y, w, h: 0.3, fontSize: 12, bold: true, color: C.text2, margin: 0 });
  const circ = async (name, x, y, d, bg) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: "icon-bg" });
    const p = d * 0.24;
    s.addImage({ data: await icon(name, HEX.lt1), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name });
  };

  // 結論
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: 0.98, w: CW, h: 0.55, fill: { color: C.text1 }, line: { type: "none" }, rectRadius: 0.08, objectName: "conclusion" });
  txt(s, [{ text: "結論　", options: { color: C.accent1, bold: true } }, { text: "Claude Code は、PCの案件フォルダで原稿チェックや書類づくりまで進めるAIです。少人数のまま対応範囲を広げる道具として、2業務に絞った試験導入をご提案します。", options: { color: C.background1, bold: true } }], { x: MX + 0.2, y: 0.98, w: CW - 0.4, h: 0.55, fontSize: 12, valign: "middle", margin: 0, lineSpacing: 16 });

  // 1段目：とは／違い／費用
  const y1 = 1.7, h1 = 2.8, c1 = 3.7, c2 = 3.55, c3 = CW - c1 - c2 - 0.4;
  const x1 = MX, x2 = MX + c1 + 0.2, x3 = x2 + c2 + 0.2;
  card(x1, y1, c1, h1); card(x2, y1, c2, h1); card(x3, y1, c3, h1);
  head("Claude Code とは", x1 + 0.2, y1 + 0.12, c1 - 0.4);
  txt(s, "ChatGPT が「窓口で相談に乗ってくれる人」なら、Claude Code は「自分の机に座り、フォルダを開いて作業までしてくれるアシスタント」です。", { x: x1 + 0.2, y: y1 + 0.45, w: c1 - 0.4, h: 0.85, fontSize: 10.5, color: C.text1, margin: 0 });
  // 小さな図：指示 → AI（PC内）→ 案件フォルダ
  const dy = y1 + 1.45, bw = 0.95, gap = (c1 - 0.4 - bw * 3) / 2;
  const boxes = [["FaUser", "人が指示", C.accent3], ["FaRobot", "AIが計画・実行", C.text2], ["FaFolderOpen", "案件フォルダに保存", C.accent1]];
  for (let i = 0; i < 3; i++) {
    const bx = x1 + 0.2 + i * (bw + gap);
    await circ(boxes[i][0], bx + (bw - 0.55) / 2, dy, 0.55, boxes[i][2]);
    txt(s, boxes[i][1], { x: bx - 0.1, y: dy + 0.6, w: bw + 0.2, h: 0.4, fontSize: 9, color: C.text1, align: "center", margin: 0, lineSpacing: 11 });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: bx + bw + gap / 2 - 0.1, y: dy + 0.18, w: 0.2, h: 0.2, fill: { color: C.accent6 }, line: { type: "none" }, objectName: "arrow" });
  }
  txt(s, "※ファイル変更などの前には「実行してよいか」を人に確認", { x: x1 + 0.2, y: y1 + h1 - 0.25, w: c1 - 0.4, h: 0.2, fontSize: 8.5, color: C.accent3, margin: 0 });

  head("ChatGPT との違い（3点）", x2 + 0.2, y1 + 0.12, c2 - 0.4);
  const diffs = [["手元のファイルを直接扱う", "コピペや添付ではなく、案件フォルダを読み書き"], ["作業を最後まで進める", "読む→直す→保存→確認を1つの指示で"], ["ルールを保存・共有できる", "表記ルールや手順をファイルで残し、誰でも同じ品質に"]];
  diffs.forEach((d, i) => {
    const yy = y1 + 0.5 + i * 0.68;
    s.addShape(pres.shapes.OVAL, { x: x2 + 0.2, y: yy + 0.02, w: 0.3, h: 0.3, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "num" });
    txt(s, String(i + 1), { x: x2 + 0.2, y: yy + 0.02, w: 0.3, h: 0.3, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, lineSpacing: null });
    txt(s, [{ text: d[0], options: { bold: true, breakLine: true } }, { text: d[1], options: { fontSize: 9.5 } }], { x: x2 + 0.6, y: yy, w: c2 - 0.8, h: 0.62, fontSize: 10.5, color: C.text1, margin: 0, lineSpacing: 14 });
  });
  txt(s, "※Web上の調べもの等は ChatGPT も得意です", { x: x2 + 0.2, y: y1 + h1 - 0.25, w: c2 - 0.4, h: 0.2, fontSize: 8.5, color: C.accent3, margin: 0 });

  head("費用の目安（1人あたり月額）", x3 + 0.2, y1 + 0.12, c3 - 0.4);
  const price = [["Pro（個人）", "$20"], ["Max（個人・多く使う方）", "$100〜"], ["Team Standard（法人2名〜）", "$25"], ["Enterprise（法人20名〜）", "個別"]];
  price.forEach((p, i) => {
    const yy = y1 + 0.48 + i * 0.38;
    txt(s, p[0], { x: x3 + 0.2, y: yy, w: c3 - 1.15, h: 0.35, fontSize: 9.5, color: C.text1, valign: "middle", margin: 0 });
    txt(s, p[1], { x: x3 + c3 - 1.0, y: yy, w: 0.8, h: 0.35, fontSize: 13, bold: true, color: i === 2 ? C.accent1 : C.text1, align: "right", valign: "middle", margin: 0 });
  });
  txt(s, "例：3名で Team Standard なら月$75（約1.1万円・1ドル150円で仮計算）", { x: x3 + 0.2, y: y1 + 2.0, w: c3 - 0.4, h: 0.4, fontSize: 9, bold: true, color: C.text2, margin: 0, lineSpacing: 12 });
  txt(s, "2026年11月時点・税抜・USD。詳細は公式ページ（claude.com/pricing）参照", { x: x3 + 0.2, y: y1 + h1 - 0.38, w: c3 - 0.4, h: 0.32, lineSpacing: 11, fontSize: 8.5, color: C.accent3, margin: 0 });

  // 2段目：ユースケース
  const y2 = y1 + h1 + 0.15, h2 = 1.45;
  card(MX, y2, CW, h2);
  head("貴社での主なユースケース", MX + 0.2, y2 + 0.12, 5);
  const uc = [["FaSpellCheck", "原稿校正", "表記ルールで一括チェック、修正案をExcelに"], ["FaClipboardList", "議事録・制作準備", "議事録→制作指示書→見積下書き"], ["FaFilePowerpoint", "企画書・提案書", "議事録から提案書のたたき台"], ["FaGlobe", "Web更新", "HTMLを直接編集し、人が表示確認"], ["FaVideo", "動画・イベント", "台本・字幕・進行表・マニュアル"], ["FaUsers", "コワーキング運営", "会員向け告知文、利用データの集計"]];
  const uw = (CW - 0.4) / 6;
  for (let i = 0; i < 6; i++) {
    const ux = MX + 0.2 + i * uw;
    await circ(uc[i][0], ux + 0.05, y2 + 0.5, 0.5, i < 2 ? C.accent1 : C.text2);
    txt(s, uc[i][1], { x: ux + 0.62, y: y2 + 0.5, w: uw - 0.7, h: 0.5, fontSize: 10.5, bold: true, color: C.text1, valign: "middle", margin: 0, lineSpacing: 13 });
    txt(s, uc[i][2], { x: ux + 0.05, y: y2 + 0.98, w: uw - 0.15, h: 0.42, fontSize: 9, color: C.text1, margin: 0, lineSpacing: 12 });
  }

  // 3段目：導入ステップ＋次の一歩
  const y3 = y2 + h2 + 0.15, h3 = 8.27 - 0.42 - y3;
  const sw = 6.6;
  card(MX, y3, sw, h3);
  head("導入ステップ（3フェーズ）", MX + 0.2, y3 + 0.12, sw - 0.4);
  const ph = [["1", "試験導入", "1〜2か月", "2〜3名で「原稿校正」「議事録・制作準備」の2業務に絞り、削減時間を実測"], ["2", "整備と展開", "3〜6か月", "表記ルール・Skills（手順書）を整え、Web・動画・イベントへ展開"], ["3", "本格導入", "7か月目〜", "人数・プランを見直し、効果測定と業務の型の共有"]];
  const pw = (sw - 0.4 - 0.3) / 3;
  ph.forEach((p, i) => {
    const px = MX + 0.2 + i * (pw + 0.15), py = y3 + 0.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: px, y: py, w: pw, h: 0.42, fill: { color: i === 0 ? C.accent1 : C.text2 }, line: { type: "none" }, rectRadius: 0.06, objectName: "phase" });
    txt(s, [{ text: `${p[0]}. ${p[1]}`, options: { bold: true, fontSize: 11 } }, { text: `　${p[2]}`, options: { fontSize: 8.5 } }], { x: px + 0.1, y: py, w: pw - 0.2, h: 0.42, color: C.background1, valign: "middle", margin: 0, lineSpacing: null });
    txt(s, p[3], { x: px + 0.05, y: py + 0.5, w: pw - 0.1, h: h3 - 1.05, fontSize: 9.5, color: C.text1, margin: 0, lineSpacing: 13 });
  });

  const nx = MX + sw + 0.2, nw = CW - sw - 0.2;
  card(nx, y3, nw, h3, "FCEDE9");
  await circ("FaHandshake", nx + 0.2, y3 + 0.12, 0.42, C.accent1);
  txt(s, "次の一歩（ご提案）", { x: nx + 0.75, y: y3 + 0.17, w: nw - 0.9, h: 0.32, fontSize: 12, bold: true, color: C.text1, margin: 0 });
  txt(s, [
    { text: "試験導入のメンバー・対象業務・期間のご検討", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "過去の原稿を使ったデモ・ハンズオン（30〜60分）", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "SecondDesk 北大路堀川での体験会（選択肢）", options: { bullet: { type: "number" } } },
  ], { x: nx + 0.2, y: y3 + 0.65, w: nw - 0.4, h: h3 - 0.75, fontSize: 9.5, color: C.text1, paraSpaceAfter: 4, margin: 0, lineSpacing: 14 });

  txt(s, "※AIの出力は「たたき台」で、入稿データ・見積金額などは必ず人が最終確認します。最新の料金・機能は claude.com/pricing・code.claude.com/docs をご確認ください。", { x: MX, y: 8.27 - 0.34, w: CW, h: 0.25, fontSize: 8, color: C.accent3, margin: 0 });
  s.addNotes("1枚サマリーです。結論、Claude Code とは、ChatGPT との違い、費用の目安、主なユースケース、導入ステップ、次の一歩を1枚にまとめています。詳細は18枚の詳細資料をご覧ください。");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
