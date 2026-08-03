/**
 * 株式会社M&A承継機構 会社紹介パンフレット（東京商工リサーチ提携先向け）
 * A3二つ折り／仕上がりA4×4ページ
 *   スライド1（外面）= P4 裏表紙 | P1 表紙
 *   スライド2（内面）= P2        | P3
 *
 * トンマナはコーポレートPPTテンプレート SL_2026 のテーマに準拠。
 */
const pptx = require("pptxgenjs");

// ---------------------------------------------------------------- ブランド
const NAVY = "0B3041";
const NAVY_DEEP = "051B26";
const MINT = "54CCB3";
const MINT_DK = "2E9C87";
const BLUE = "0889C9";
const PALE = "E1F3FB";
const COOL = "E3E7E9";
const GRAY = "5A6670";
const GRAY_LT = "8A959D";
const LINE = "D8E0E4";
const BG = "FFFFFF";
const W = "FFFFFF";

const JPM = "游ゴシック Medium"; // 見出し
const JP = "游ゴシック"; // 本文
const EN = "Arial"; // 欧文ラベル・数字

// ---------------------------------------------------------------- 版面
const SW = 16.5354, SH = 11.6929; // A3横
const PW = SW / 2, PH = SH; // 仕上がりA4縦
const M = 0.62; // 天地左右マージン
const CW = PW - M * 2; // 段幅 7.0277
const L = 0, R = PW; // 左ページ／右ページの原点

const A = "assets/";

const deck = new pptx();
deck.defineLayout({ name: "A3", width: SW, height: SH });
deck.layout = "A3";
deck.author = "株式会社M&A承継機構";
deck.title = "M&A承継機構 会社紹介パンフレット";

// ---------------------------------------------------------------- 部品
const tx = (s, o) => s.addText.bind(s);

/** 英字の小見出し（花弁マーク付き） */
function eyebrow(s, x, y, label, color, markImg) {
  s.addImage({ path: A + (markImg || "mark_mint.png"), x: x, y: y + 0.022, w: 0.105, h: 0.105 });
  s.addText(label, {
    x: x + 0.17, y: y - 0.03, w: 4.2, h: 0.2, margin: 0,
    fontFace: EN, fontSize: 8, bold: true, color: color || MINT_DK, charSpacing: 2.4,
  });
}

/** ページ番号 */
function folio(s, x, n) {
  s.addText(n, {
    x: x + CW - 0.8, y: M - 0.06, w: 0.8, h: 0.24, margin: 0,
    fontFace: EN, fontSize: 10, color: "BCC6CC", align: "right",
  });
}

/** 罫線 */
function rule(s, x, y, w, color, h) {
  s.addShape(deck.ShapeType.rect, { x, y, w, h: h || 0.008, fill: { color: color || LINE }, line: { none: true } });
}

/** 見出し（日本語・大） */
function headline(s, x, y, lines, size, color) {
  s.addText(lines.map((t, i) => ({ text: t, options: { breakLine: i < lines.length - 1 } })), {
    x, y, w: CW, h: 0.52 * lines.length + 0.12, margin: 0,
    fontFace: JPM, fontSize: size || 19, bold: true, color: color || NAVY, lineSpacingMultiple: 1.34,
  });
}

/** セクション見出し（日本語・中） */
function sectionHead(s, x, y, jp, en) {
  s.addText(jp, {
    x, y, w: CW * 0.6, h: 0.26, margin: 0,
    fontFace: JPM, fontSize: 12.5, bold: true, color: NAVY,
  });
  if (en) {
    s.addText(en, {
      x: x + CW * 0.6, y: y + 0.045, w: CW * 0.4, h: 0.22, margin: 0,
      fontFace: EN, fontSize: 7.5, color: GRAY_LT, charSpacing: 1.8, align: "right",
    });
  }
  rule(s, x, y + 0.30, CW, LINE);
}

/** カード（枠のみ・面のみ） */
function card(s, x, y, w, h, opt) {
  opt = opt || {};
  s.addShape(deck.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.05,
    fill: { color: opt.fill || W },
    line: opt.noLine ? { none: true } : { color: opt.line || LINE, width: 0.6 },
    shadow: opt.shadow
      ? { type: "outer", color: "0B3041", opacity: 0.07, blur: 8, offset: 1.5, angle: 90 }
      : undefined,
  });
}

// ================================================================ スライド1：外面
const s1 = deck.addSlide();
s1.background = { color: BG };

// ---------------------------------------------------------------- P1 表紙（右）
{
  const x = R + M;

  // 全面をヒーローで裁ち落とす。スクリムは上下2枚 —
  // 中央に光を残しつつ、文字が乗る天地は必ず沈める。
  s1.addImage({ path: A + "hero_cover.jpg", x: R, y: 0, w: PW, h: PH, sizing: { type: "cover", w: PW, h: PH } });
  s1.addImage({ path: A + "scrim_top.png", x: R, y: 0, w: PW, h: PH * 0.42 });
  s1.addImage({ path: A + "scrim_cover.png", x: R, y: 0, w: PW, h: PH });

  s1.addImage({ path: A + "logo_knockout.png", x: x, y: M - 0.02, w: 2.52, h: 0.50 });
  s1.addText("COMPANY PROFILE", {
    x: x + CW - 3.0, y: M + 0.11, w: 3.0, h: 0.22, margin: 0,
    fontFace: EN, fontSize: 8, bold: true, color: "9FB4BD", charSpacing: 2.2, align: "right",
  });

  eyebrow(s1, x, 5.30, "M&A ADVISORY & BUSINESS SUCCESSION", MINT, "mark_mint.png");

  s1.addText(
    [
      { text: "会社の未来を、", options: { breakLine: true } },
      { text: "次の成長へつなぐ。", options: {} },
    ],
    {
      x, y: 5.70, w: CW, h: 2.10, margin: 0,
      fontFace: JPM, fontSize: 37, bold: true, color: W, lineSpacingMultiple: 1.26,
    }
  );

  s1.addText(
    "事業承継、成長戦略、資本提携。\n経営者の想いに寄り添い、最適な選択肢をご提案します。",
    {
      x, y: 7.94, w: CW, h: 0.80, margin: 0,
      fontFace: JP, fontSize: 11, color: "C6D7DE", lineSpacingMultiple: 1.58,
    }
  );

  s1.addImage({ path: A + "rule_accent.png", x, y: 9.02, w: 1.30, h: 0.030 });

  s1.addText("企業価値の最大化を通じて、経済に正しい循環を。", {
    x, y: 9.26, w: CW, h: 0.32, margin: 0,
    fontFace: JPM, fontSize: 12.5, bold: true, color: MINT,
  });

  s1.addText(
    [
      { text: "株式会社TWOSTONE&Sons", options: { bold: true, color: W } },
      { text: "（東証グロース・証券コード 7352）グループ", options: { color: "A9BEC7" } },
    ],
    { x, y: 9.86, w: CW, h: 0.24, margin: 0, fontFace: JP, fontSize: 8.5 }
  );

  s1.addShape(deck.ShapeType.roundRect, {
    x, y: PH - M - 0.44, w: 4.92, h: 0.44, rectRadius: 0.22,
    fill: { color: "FFFFFF", transparency: 88 }, line: { color: "FFFFFF", width: 0.6, transparency: 58 },
  });
  s1.addText("本資料は東京商工リサーチのご担当者様を通じてご案内しています", {
    x: x + 0.24, y: PH - M - 0.44, w: 4.5, h: 0.44, margin: 0,
    fontFace: JP, fontSize: 8.5, color: W, valign: "middle",
  });
}

// ---------------------------------------------------------------- P4 裏表紙（左）
{
  const x = L + M;
  let y = M;

  eyebrow(s1, x, y, "CONTACT & COMPANY");
  folio(s1, x, "04");
  y += 0.52;

  headline(s1, x, y, ["まだ具体的でない段階から、", "お気軽にご相談ください。"], 19);
  y += 1.16;

  s1.addText(
    "ご相談いただいたからといって、M&Aを進めなければならないわけではありません。まずは会社の現状と、考えられる選択肢を一緒に整理するところから始めます。",
    { x, y, w: CW, h: 0.62, margin: 0, fontFace: JP, fontSize: 9, color: GRAY, lineSpacingMultiple: 1.55 }
  );
  y += 0.78;

  // ご相談の流れ
  sectionHead(s1, x, y, "ご相談の流れ", "HOW TO START");
  y += 0.56;

  const steps = [
    ["東京商工リサーチの", "ご担当者様へご相談"],
    ["M&A承継機構との", "初回面談"],
    ["現状・ご意向・", "選択肢の整理"],
    ["必要に応じて企業評価・", "具体的なご提案"],
  ];
  const colW = CW / 4;
  // 連結線
  rule(s1, x + colW / 2, y + 0.185, colW * 3, COOL, 0.022);
  steps.forEach((sp, i) => {
    const cx = x + colW * i + colW / 2;
    const last = i === steps.length - 1;
    s1.addShape(deck.ShapeType.ellipse, {
      x: cx - 0.195, y: y, w: 0.39, h: 0.39,
      fill: { color: last ? MINT_DK : W }, line: { color: last ? MINT_DK : COOL, width: 1.1 },
    });
    s1.addText(String(i + 1), {
      x: cx - 0.195, y: y, w: 0.39, h: 0.39, margin: 0,
      fontFace: EN, fontSize: 11, bold: true, color: last ? W : NAVY, align: "center", valign: "middle",
    });
    s1.addText(sp.join("\n"), {
      x: cx - colW / 2 + 0.04, y: y + 0.50, w: colW - 0.08, h: 0.58, margin: 0,
      fontFace: JP, fontSize: 8.5, color: NAVY, align: "center", lineSpacingMultiple: 1.32,
    });
  });
  y += 1.22;

  // 3つの安心
  const promises = [
    ["初回相談無料", "企業価値評価のみのご依頼も可能"],
    ["完全成功報酬", "ご成約まで費用は発生しません"],
    ["秘密厳守", "ISO/IEC 27001 認証取得"],
  ];
  const pw = (CW - 0.24) / 3;
  promises.forEach((p, i) => {
    const px = x + (pw + 0.12) * i;
    card(s1, px, y, pw, 0.86, { fill: "F5FAFB", line: "DCE8EC" });
    s1.addText(p[0], {
      x: px, y: y + 0.11, w: pw, h: 0.26, margin: 0,
      fontFace: JPM, fontSize: 10.5, bold: true, color: NAVY, align: "center",
    });
    s1.addText(p[1], {
      x: px + 0.08, y: y + 0.40, w: pw - 0.16, h: 0.36, margin: 0,
      fontFace: JP, fontSize: 7.5, color: GRAY, align: "center", lineSpacingMultiple: 1.3,
    });
  });
  y += 1.10;

  // 会社概要
  sectionHead(s1, x, y, "会社概要", "COMPANY");
  y += 0.50;

  const rows = [
    ["会社名", "株式会社M&A承継機構（M&A SUCCESSION GROUP INC.）"],
    ["所在地", "東京都港区港南1-8-15 Wビル13F"],
    ["代表者", "代表取締役CEO　西村 将明"],
    ["事業内容", "M&A仲介・アドバイザリー業務／財務・会計アドバイザリー業務／事業承継コンサルティング"],
    ["従業員数", "70名（2026年6月時点）"],
    ["グループ", "株式会社TWOSTONE&Sons（東証グロース・証券コード7352）のグループ会社\nグループ連結売上高 18,077百万円（2025年8月期実績）／連結従業員 750名（2025年8月時点）"],
    ["登録・許認可", "中小企業庁 M&A支援機関登録制度 登録企業／一般社団法人M&A支援機関協会 正会員\n宅地建物取引業免許 東京都知事(1)第112845号／ISO/IEC 27001:2022 認証取得"],
  ];
  const labW = 1.16;
  rows.forEach((r) => {
    const two = r[1].indexOf("\n") >= 0;
    const rh = two ? 0.44 : 0.27;
    s1.addText(r[0], {
      x, y, w: labW, h: rh, margin: 0,
      fontFace: JP, fontSize: 8, color: GRAY_LT, valign: "top",
    });
    s1.addText(r[1], {
      x: x + labW, y, w: CW - labW, h: rh, margin: 0,
      fontFace: JP, fontSize: 8, color: NAVY, valign: "top", lineSpacingMultiple: 1.34,
    });
    y += rh;
    rule(s1, x, y - 0.045, CW, "EDF1F3");
  });

  // CONTACT
  const cy = PH - M - 1.90;
  s1.addImage({ path: A + "band_contact.jpg", x: L, y: cy, w: PW, h: 1.90 + M, sizing: { type: "cover", w: PW, h: 1.90 + M } });
  s1.addText("CONTACT", {
    x, y: cy + 0.34, w: 3.0, h: 0.2, margin: 0,
    fontFace: EN, fontSize: 8, bold: true, color: MINT, charSpacing: 2.4,
  });
  s1.addText(
    [
      { text: "ご相談は、", options: {} },
      { text: "東京商工リサーチのご担当者様", options: { color: MINT } },
      { text: "まで", options: {} },
      { text: "\nお申し付けください。", options: {} },
    ],
    {
      x, y: cy + 0.64, w: CW - 1.55, h: 0.76, margin: 0,
      fontFace: JPM, fontSize: 12.5, bold: true, color: W, lineSpacingMultiple: 1.28,
    }
  );
  s1.addText("https://ma-sg.co.jp/", {
    x, y: cy + 1.46, w: CW - 1.55, h: 0.26, margin: 0,
    fontFace: EN, fontSize: 9.5, color: MINT, charSpacing: 0.6,
  });
  s1.addShape(deck.ShapeType.rect, {
    x: x + CW - 1.30, y: cy + 0.42, w: 1.30, h: 1.30,
    fill: { color: "FFFFFF", transparency: 88 }, line: { color: "FFFFFF", width: 0.6, transparency: 60 },
  });
  s1.addText("QRコード\n（差し替え）", {
    x: x + CW - 1.30, y: cy + 0.42, w: 1.30, h: 1.30, margin: 0,
    fontFace: JP, fontSize: 7.5, color: W, align: "center", valign: "middle", lineSpacingMultiple: 1.3,
  });
}

// 中央折り線
[s1].forEach((s) => {});

// ================================================================ スライド2：内面
const s2 = deck.addSlide();
s2.background = { color: BG };

// ---------------------------------------------------------------- P2（左）
{
  const x = L + M;
  let y = M;

  eyebrow(s2, x, y, "YOUR CHALLENGE");
  folio(s2, x, "02");
  y += 0.52;

  headline(s2, x, y, ["会社の未来について、", "このようなお悩みはありませんか。"], 19);
  y += 1.16;

  s2.addText("ひとつでも当てはまるものがあれば、いまが会社の将来を考えるタイミングかもしれません。", {
    x, y, w: CW, h: 0.24, margin: 0, fontFace: JP, fontSize: 9, color: GRAY,
  });
  y += 0.40;

  const worries = [
    "後継者が決まっていない",
    "従業員や取引先を守りたい",
    "自社の強みを次の世代に残したい",
    "大手企業との連携で事業を成長させたい",
    "一部事業の譲渡や整理を検討している",
    "会社の価値や今後の選択肢を知りたい",
  ];
  const wcw = CW / 2, wrh = 0.46;
  rule(s2, x, y, CW, LINE);
  worries.forEach((t, i) => {
    const c = i % 2, r = Math.floor(i / 2);
    const wx = x + wcw * c, wy = y + 0.10 + wrh * r;
    s2.addText(String(i + 1).padStart(2, "0"), {
      x: wx, y: wy, w: 0.42, h: 0.30, margin: 0,
      fontFace: EN, fontSize: 11.5, bold: true, color: MINT_DK,
    });
    s2.addText(t, {
      x: wx + 0.44, y: wy + 0.015, w: wcw - 0.50, h: 0.30, margin: 0,
      fontFace: JP, fontSize: 9.5, color: NAVY, valign: "middle",
    });
    if (c === 1) rule(s2, x, y + 0.10 + wrh * (r + 1) - 0.06, CW, "EDF1F3");
  });
  y += 0.10 + wrh * 3 + 0.18;

  // 転換メッセージ
  s2.addImage({ path: A + "panel_dark.jpg", x, y, w: CW, h: 1.42, sizing: { type: "cover", w: CW, h: 1.42 } });
  s2.addImage({ path: A + "scrim_panel.png", x, y, w: CW, h: 1.42 });
  s2.addText("大切なのは、M&Aをするかどうかを急いで決めることではありません。", {
    x: x + 0.34, y: y + 0.22, w: CW - 0.68, h: 0.30, margin: 0,
    fontFace: JPM, fontSize: 12, bold: true, color: W,
  });
  s2.addText(
    "まずは自社の価値と、考えられる選択肢を知ることから始まります。M&Aは会社を手放すためだけの手段ではなく、事業承継・成長戦略としての買収・資本提携・不採算事業の整理・事業再生など、会社の未来を考えるための選択肢のひとつです。",
    {
      x: x + 0.34, y: y + 0.60, w: CW - 0.68, h: 0.76, margin: 0,
      fontFace: JP, fontSize: 8.5, color: "B9CDD6", lineSpacingMultiple: 1.48,
    }
  );
  y += 1.64;

  // 選ばれる理由
  sectionHead(s2, x, y, "選ばれる理由", "WHY WE ARE CHOSEN");
  y += 0.50;

  const reasons = [
    ["01", "実績豊富なプロフェッショナル", "在籍アドバイザーの70%が大手M&A仲介会社の出身。豊富な実務経験に基づく専門性で、早期の成約を実現します。"],
    ["02", "上場企業グループの安心感", "東証グロース上場・TWOSTONE&Sonsのグループ会社。確かな経営基盤と情報管理体制のもとで対応します。"],
    ["03", "複雑な案件にも対応する実行力", "会社分割・二段階譲渡・私的再生など、他社が敬遠する難易度の高い案件にも知見をもって伴走します。"],
    ["04", "M&A以外の経営課題にも対応", "資金調達・財務改善から成長支援まで、M&Aを前提としないご相談も承ります。"],
  ];
  const rcw = (CW - 0.20) / 2, rch = 1.36;
  reasons.forEach((rr, i) => {
    const c = i % 2, r = Math.floor(i / 2);
    const rx = x + (rcw + 0.20) * c, ry = y + (rch + 0.16) * r;
    card(s2, rx, ry, rcw, rch, { shadow: true });
    s2.addText(rr[0], {
      x: rx + 0.24, y: ry + 0.15, w: 1.0, h: 0.36, margin: 0,
      fontFace: EN, fontSize: 19, bold: true, color: MINT,
    });
    s2.addText(rr[1], {
      x: rx + 0.24, y: ry + 0.55, w: rcw - 0.48, h: 0.26, margin: 0,
      fontFace: JPM, fontSize: 10, bold: true, color: NAVY,
    });
    s2.addText(rr[2], {
      x: rx + 0.24, y: ry + 0.84, w: rcw - 0.48, h: 0.50, margin: 0,
      fontFace: JP, fontSize: 8, color: GRAY, lineSpacingMultiple: 1.42,
    });
  });
  y += rch * 2 + 0.16 + 0.24;

  // 実績バンド
  const bh = 0.86;
  s2.addImage({ path: A + "band_stat.jpg", x, y, w: CW, h: bh, sizing: { type: "cover", w: CW, h: bh } });
  const stats = [
    ["650", "件+", "累計成約実績※"],
    ["70", "%", "大手M&A仲介会社出身※"],
    ["100", "件+", "常時受託している譲渡案件"],
  ];
  const sw = CW / 3;
  stats.forEach((st, i) => {
    const sx = x + sw * i;
    if (i > 0) {
      s2.addShape(deck.ShapeType.rect, {
        x: sx, y: y + 0.20, w: 0.008, h: bh - 0.40,
        fill: { color: "FFFFFF", transparency: 70 }, line: { none: true },
      });
    }
    s2.addText(
      [
        { text: st[0], options: { fontFace: EN, fontSize: 19, bold: true, color: W } },
        { text: st[1], options: { fontFace: EN, fontSize: 9.5, bold: true, color: W } },
      ],
      { x: sx, y: y + 0.14, w: sw, h: 0.32, margin: 0, align: "center" }
    );
    s2.addText(st[2], {
      x: sx + 0.04, y: y + 0.50, w: sw - 0.08, h: 0.22, margin: 0,
      fontFace: JP, fontSize: 7.5, color: "CFE7E4", align: "center",
    });
  });
  y += bh + 0.14;

  [
    ["支援領域｜", "事業承継・カーブアウト・私的再生・スポンサー探索・業界再編・不動産M&A・スタートアップExit"],
    ["業種例｜", "訪問介護／人材サービス／外食（上場）／自動車部品／建材製造／不動産／広告／エンタメ ほか"],
  ].forEach((ln, i) => {
    s2.addText(
      [
        { text: ln[0], options: { bold: true, color: NAVY } },
        { text: ln[1], options: { color: GRAY } },
      ],
      { x, y: y + 0.25 * i, w: CW, h: 0.24, margin: 0, fontFace: JP, fontSize: 7.5, valign: "middle" }
    );
  });
  y += 0.56;

  s2.addText("※ 在籍アドバイザーの過去の成約実績・出身実績を含みます。", {
    x, y, w: CW, h: 0.2, margin: 0, fontFace: JP, fontSize: 6.8, color: GRAY_LT,
  });
}

// ---------------------------------------------------------------- P3（右）
{
  const x = R + M;
  let y = M;

  eyebrow(s2, x, y, "SERVICE & FEE");
  folio(s2, x, "03");
  y += 0.52;

  headline(s2, x, y, ["M&A仲介にとどまらず、", "経営課題にワンストップで。"], 19);
  y += 1.16;

  // 中核サービス
  s2.addImage({ path: A + "band_service.jpg", x, y, w: CW, h: 1.00, sizing: { type: "cover", w: CW, h: 1.00 } });
  s2.addImage({ path: A + "scrim_panel.png", x, y, w: CW, h: 1.00 });
  s2.addText("M&A仲介・アドバイザリー", {
    x: x + 0.34, y: y + 0.17, w: CW - 0.68, h: 0.28, margin: 0,
    fontFace: JPM, fontSize: 12, bold: true, color: W,
  });
  s2.addText(
    "譲渡（事業承継・成長戦略）と譲受（買収・資本提携）の双方を支援。株式譲渡・事業譲渡に加え、会社分割や二段階譲渡など目的に応じたスキームを設計します。",
    {
      x: x + 0.34, y: y + 0.48, w: CW - 0.68, h: 0.46, margin: 0,
      fontFace: JP, fontSize: 8.5, color: "BDD1DA", lineSpacingMultiple: 1.42,
    }
  );
  y += 1.20;

  // ワンストップ支援
  s2.addText("M&A以外の経営課題も、提携先ネットワークでワンストップにご相談いただけます。", {
    x, y, w: CW, h: 0.24, margin: 0, fontFace: JP, fontSize: 8.5, color: GRAY,
  });
  y += 0.34;

  const svc = [
    ["財務・資金繰り支援", ["資金調達・資金繰り改善", "財務コンサルティング", "事業再生・スポンサー探索"]],
    ["成長支援", ["成長戦略・DX・IT支援", "人材・組織構築", "BPO・業務効率化"]],
    ["譲渡後のサポート", ["資産運用・保険", "不動産のご紹介", "税務・専門家との連携"]],
  ];
  const scw = (CW - 0.24) / 3;
  svc.forEach((sv, i) => {
    const sx = x + (scw + 0.12) * i;
    card(s2, sx, y, scw, 1.32, { fill: "F7FAFB", line: "E2EAEE" });
    s2.addText(sv[0], {
      x: sx + 0.16, y: y + 0.13, w: scw - 0.32, h: 0.24, margin: 0,
      fontFace: JPM, fontSize: 9.5, bold: true, color: NAVY,
    });
    rule(s2, sx + 0.16, y + 0.41, scw - 0.32, "D3E2E7");
    s2.addText(sv[1].join("\n"), {
      x: sx + 0.16, y: y + 0.48, w: scw - 0.32, h: 0.78, margin: 0,
      fontFace: JP, fontSize: 8, color: GRAY, lineSpacingMultiple: 1.5,
    });
  });
  y += 1.32 + 0.30;

  // 料金
  sectionHead(s2, x, y, "料金体系｜完全成功報酬型", "FEE");
  y += 0.52;

  s2.addText("ご成約まで、費用は一切いただきません。", {
    x, y, w: CW, h: 0.28, margin: 0,
    fontFace: JPM, fontSize: 12.5, bold: true, color: NAVY,
  });
  y += 0.38;

  const frees = ["着手金", "中間報酬", "企業評価料", "案件化料"];
  const fw = (CW - 0.27) / 4;
  frees.forEach((f, i) => {
    const fx = x + (fw + 0.09) * i;
    s2.addShape(deck.ShapeType.roundRect, {
      x: fx, y, w: fw, h: 0.62, rectRadius: 0.05,
      fill: { color: "EEF9F6" }, line: { color: MINT, width: 0.8 },
    });
    s2.addText(
      [
        { text: f, options: { fontFace: JP, fontSize: 8.5, color: NAVY, breakLine: true } },
        { text: "無料", options: { fontFace: JPM, fontSize: 11.5, bold: true, color: MINT_DK } },
      ],
      { x: fx, y: y + 0.055, w: fw, h: 0.52, margin: 0, align: "center", lineSpacingMultiple: 1.16 }
    );
  });
  y += 0.80;

  // 他社比較
  const head = ["", "当社", "大手A社", "大手B社"];
  const body = [
    ["料金体系", "完全成功報酬", "着手金＋中間報酬\n＋成功報酬", "中間報酬\n＋成功報酬"],
    ["着手金", "無料", "あり", "無料"],
    ["中間報酬", "無料", "あり", "あり"],
    ["株価評価・案件化料", "無料", "着手金に含む", "無料"],
    ["最低報酬", "1,500万円", "2,000万円", "2,500万円"],
  ];
  const c0 = 1.72, cN = (CW - c0) / 3;
  const hRow = 0.34;

  // ヘッダ
  s2.addShape(deck.ShapeType.rect, { x: x + c0, y, w: cN, h: hRow, fill: { color: NAVY }, line: { none: true } });
  head.forEach((h, i) => {
    if (i === 0) return;
    const hx = x + c0 + cN * (i - 1);
    s2.addText(h, {
      x: hx, y, w: cN, h: hRow, margin: 0,
      fontFace: JPM, fontSize: 9, bold: true, color: i === 1 ? W : NAVY,
      align: "center", valign: "middle",
    });
  });
  let ty = y + hRow;

  body.forEach((row, ri) => {
    const two = row.some((c) => c.indexOf("\n") >= 0);
    const rh = two ? 0.50 : 0.36;
    const isLast = ri === body.length - 1;
    // 当社列を淡いブランド面で強調
    s2.addShape(deck.ShapeType.rect, {
      x: x + c0, y: ty, w: cN, h: rh,
      fill: { color: isLast ? "DFF4EE" : "F1FAF8" }, line: { none: true },
    });
    s2.addText(row[0], {
      x, y: ty, w: c0 - 0.12, h: rh, margin: 0,
      fontFace: JP, fontSize: 8.5, color: GRAY, valign: "middle",
    });
    row.slice(1).forEach((c, ci) => {
      s2.addText(c, {
        x: x + c0 + cN * ci, y: ty, w: cN, h: rh, margin: 0,
        fontFace: ci === 0 ? JPM : JP,
        fontSize: isLast ? (ci === 0 ? 11 : 9) : 8.5,
        bold: ci === 0,
        color: ci === 0 ? (isLast ? MINT_DK : NAVY) : GRAY,
        align: "center", valign: "middle", lineSpacingMultiple: 1.24,
      });
    });
    ty += rh;
    rule(s2, x, ty, CW, "E4EBEE");
  });
  // 当社列の輪郭
  s2.addShape(deck.ShapeType.rect, {
    x: x + c0, y, w: cN, h: ty - y,
    fill: { type: "solid", color: "FFFFFF", transparency: 100 }, line: { color: MINT_DK, width: 1.2 },
  });
  y = ty + 0.16;

  s2.addText(
    "※ 成功報酬は譲渡企業の時価総資産額（営業権を含む）に対するレーマン方式。5億円以下の部分5%／5億円超10億円以下の部分4%／10億円超50億円以下の部分3%／50億円超100億円以下の部分2%／100億円超の部分1%（消費税別）。合併・会社分割・業務提携等は形態・規模に応じてお見積りいたします。\n※ 他社条件は当社調べ。2026年8月時点の公表情報等に基づく一例であり、各社の最新の料金体系とは異なる場合があります。",
    { x, y, w: CW, h: 0.60, margin: 0, fontFace: JP, fontSize: 6.8, color: GRAY_LT, lineSpacingMultiple: 1.35 }
  );
  y += 0.70;

  // 支援事例（匿名化）
  sectionHead(s2, x, y, "支援事例（匿名化）", "CASE");
  y += 0.42;
  const cases = [
    ["事業承継", "後継者不在の企業の株式譲渡。条件面を丁寧に整理し、会社と従業員の将来を守るかたちで円満に承継。"],
    ["成長型M&A", "スタートアップの株式譲渡。成長性を評価に織り込み、株式譲渡とアーンアウトの組合せで約6カ月で成約。"],
    ["資金調達支援", "食品卸の運転資金確保。M&Aによらない本業支援として、ご相談から約2カ月で調達を完了。"],
  ];
  cases.forEach((cc) => {
    s2.addShape(deck.ShapeType.roundRect, {
      x, y: y + 0.015, w: 0.92, h: 0.22, rectRadius: 0.04,
      fill: { color: "EEF9F6" }, line: { none: true },
    });
    s2.addText(cc[0], {
      x, y: y + 0.015, w: 0.92, h: 0.22, margin: 0,
      fontFace: JPM, fontSize: 7, bold: true, color: MINT_DK, align: "center", valign: "middle",
    });
    s2.addText(cc[1], {
      x: x + 1.04, y, w: CW - 1.04, h: 0.26, margin: 0,
      fontFace: JP, fontSize: 7.8, color: GRAY, valign: "middle",
    });
    y += 0.285;
  });
}

// ---------------------------------------------------------------- 折り線ガイド
[s1, s2].forEach((s) => {
  s.addText("中央折り線", {
    x: PW - 0.62, y: 0.10, w: 1.24, h: 0.18, margin: 0,
    fontFace: JP, fontSize: 6.5, color: "C8D0D4", align: "center",
  });
  s.addShape(deck.ShapeType.rect, {
    x: PW - 0.0025, y: 0.34, w: 0.005, h: SH - 0.68,
    fill: { color: "DDE4E7" }, line: { none: true },
  });
});

deck.writeFile({ fileName: "MAsg_brochure.pptx" }).then(() => console.log("built"));
