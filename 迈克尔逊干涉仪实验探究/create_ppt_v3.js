const pptxgen = require("pptxgenjs");

// ─── Color Palette ───
const C = {
  darkBg:    "0D1B2A",
  primary:   "1B3A5C",
  secondary: "2980B9",
  accent:    "CC2233",
  gold:      "D4A017",
  lightBg:   "F0F4F8",
  cardBg:    "FFFFFF",
  text:      "1E293B",
  muted:     "64748B",
  white:     "FFFFFF",
  codeBg:    "F5F7FA",
};

const FONT_H = "Georgia";
const FONT_B = "Calibri";
const MARGIN = 0.5;

const makeShadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.10 });

function addSlideNumber(slide, num, total, dark = false) {
  slide.addText(`${num} / ${total}`, {
    x: 9.0, y: 5.25, w: 0.8, h: 0.3,
    fontSize: 9, fontFace: FONT_B, color: dark ? "8899AA" : C.muted,
    align: "right", margin: 0,
  });
}
function addFooterBar(slide, dark = false) {
  slide.addShape("rect", { x: 0, y: 5.35, w: 10, h: 0.275, fill: { color: dark ? C.primary : C.secondary } });
}
function addSectionTitle(slide, titleText, dark = false) {
  slide.addText(titleText, {
    x: MARGIN, y: 0.25, w: 8.6, h: 0.5,
    fontSize: 28, fontFace: FONT_H, color: dark ? C.white : C.text, bold: true, margin: 0,
  });
  slide.addShape("rect", { x: MARGIN, y: 0.82, w: 1.5, h: 0.04, fill: { color: dark ? C.gold : C.accent } });
}
function addIconBox(slide, symbol, x, y, w, h, bg) {
  slide.addShape("rect", { x, y, w, h, fill: { color: bg }, rectRadius: 0.05 });
  slide.addText(symbol, { x, y, w, h, fontSize: w*26, fontFace: "Segoe UI Symbol", color: C.white, align: "center", valign: "middle", margin: 0 });
}
function addMathBox(slide, lines, x, y, w, h, bg = C.codeBg) {
  slide.addShape("rect", { x, y, w, h, fill: { color: bg } });
  slide.addText(lines, { x: x+0.1, y: y+0.08, w: w-0.2, h: h-0.16, fontSize: 12, fontFace: "Cambria Math", color: C.text, lineSpacingMultiple: 1.25, valign: "top", margin: 0 });
}
function addHighlight(slide, text, x, y, w, h, color) {
  slide.addShape("rect", { x, y, w, h, fill: { color: color } });
  slide.addText(text, { x: x+0.15, y, w: w-0.3, h, fontSize: 11, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
}

async function main() {
  const TOTAL = 15;
  let pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "实验探究答辩";
  pres.title = "迈克尔逊干涉仪测折射率实验探究";

  // ==================== SLIDE 1: COVER ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
    s.addText("迈克尔逊干涉仪\n测折射率实验探究", {
      x: 0.8, y: 1.0, w: 8.4, h: 2.0,
      fontSize: 42, fontFace: FONT_H, color: C.white, bold: true,
      align: "left", lineSpacingMultiple: 1.2, margin: 0,
    });
    s.addText("答辩报告", {
      x: 0.8, y: 3.2, w: 8.4, h: 0.6,
      fontSize: 22, fontFace: FONT_B, color: C.gold, align: "left", margin: 0,
    });
    s.addShape("rect", { x: 0.8, y: 3.0, w: 2.0, h: 0.03, fill: { color: C.accent } });
    s.addText("物理实验 · 迈克尔逊干涉仪 · 折射率测量", {
      x: 0.8, y: 4.6, w: 8.4, h: 0.35,
      fontSize: 12, fontFace: FONT_B, color: "8899AA", margin: 0,
    });
    addSlideNumber(s, 1, TOTAL, true);
  }

  // ==================== SLIDE 2: OUTLINE ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "目录");
    addFooterBar(s);
    const items = [
      { num: "01", title: "实验背景、目的与仪器", desc: "研究动机与实验配置" },
      { num: "02", title: "实验现象总览", desc: "红光 & 白光两种实验流程与观测对比" },
      { num: "03", title: "核心原理：形态判据", desc: "两虚点光源模型与 Δ(θ)≈2h cos θ" },
      { num: "04", title: "红光原理：虚像位移补偿", desc: "减速效应 vs 视深效应，补偿逻辑" },
      { num: "05", title: "白光原理：零光程差补偿", desc: "OPD≈0 约束与折射率推导" },
      { num: "06", title: "数据处理、误差分析与结论", desc: "完整计算、误差讨论与最终结论" },
    ];
    items.forEach((item, i) => {
      const col = i < 3 ? 0 : 1;
      const row = i % 3;
      const x0 = MARGIN + col * 4.6;
      const y0 = 1.35 + row * 1.2;
      s.addText(item.num, { x: x0, y: y0, w: 0.65, h: 0.9, fontSize: 26, fontFace: FONT_H, color: C.accent, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(item.title, { x: x0+0.75, y: y0+0.05, w: 3.5, h: 0.4, fontSize: 15, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
      s.addText(item.desc, { x: x0+0.75, y: y0+0.45, w: 3.5, h: 0.35, fontSize: 10, fontFace: FONT_B, color: C.muted, margin: 0 });
    });
    addSlideNumber(s, 2, TOTAL);
  }

  // ==================== SLIDE 3: 实验目的与仪器 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "实验目的与仪器");
    addFooterBar(s);

    // Left: purpose
    s.addShape("rect", { x: MARGIN, y: 1.25, w: 4.3, h: 1.4, fill: { color: C.cardBg }, shadow: makeShadow() });
    addIconBox(s, "🎯", MARGIN+0.2, 1.4, 0.35, 0.35, C.accent);
    s.addText("实验目的", { x: MARGIN+0.65, y: 1.4, w: 3.0, h: 0.35, fontSize: 15, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "1.", options: { bold:true, color:C.accent, fontSize:12 } },
      { text: " 用迈克尔逊干涉仪测量两种不同厚度与类型玻璃片的折射率。\n", options: { fontSize:12, color:C.text } },
      { text: "2.", options: { bold:true, color:C.accent, fontSize:12 } },
      { text: " 探究不同光源下插入玻璃片后，动镜 M1 移动方向差异的物理原因，并给出对应折射率计算公式。", options: { fontSize:12, color:C.text } },
    ], { x: MARGIN+0.2, y: 1.85, w: 3.9, h: 0.8, fontFace: FONT_B, lineSpacingMultiple: 1.3, margin: 0 });

    // Right: equipment list
    s.addShape("rect", { x: 5.3, y: 1.25, w: 4.3, h: 1.4, fill: { color: C.cardBg }, shadow: makeShadow() });
    addIconBox(s, "⚙", 5.5, 1.4, 0.35, 0.35, C.primary);
    s.addText("实验仪器", { x: 5.95, y: 1.4, w: 3.0, h: 0.35, fontSize: 15, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText("迈克尔逊干涉仪、氦氖激光器、短焦距透镜、扩束镜、毛玻璃片、白光光源、观察屏", {
      x: 5.5, y: 1.85, w: 3.9, h: 0.8, fontSize: 12, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.3, margin: 0
    });

    // Bottom: configuration diagram (text-based)
    s.addShape("rect", { x: MARGIN, y: 2.85, w: 9.0, h: 2.35, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("实验配置", { x: MARGIN+0.2, y: 2.95, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });

    const boxH = 0.32;
    const boxY = 3.45;
    s.addShape("rect", { x: 0.8, y: boxY, w: 0.7, h: boxH, fill: { color: C.secondary }, rectRadius: 0.03 });
    s.addText("光源 S", { x: 0.8, y: boxY, w: 0.7, h: boxH, fontSize: 8, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText("→", { x: 1.55, y: boxY, w: 0.3, h: boxH, fontSize: 10, color: C.muted, align: "center", valign: "middle", margin: 0 });
    s.addShape("rect", { x: 1.9, y: boxY, w: 0.55, h: boxH, fill: { color: C.primary } });
    s.addText("G₁", { x: 1.9, y: boxY, w: 0.55, h: boxH, fontSize: 9, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText("—", { x: 2.5, y: boxY, w: 0.25, h: boxH, fontSize: 10, color: C.muted, align: "center", valign: "middle", margin: 0 });
    s.addShape("rect", { x: 2.8, y: boxY, w: 0.55, h: boxH, fill: { color: C.primary } });
    s.addText("G₂", { x: 2.8, y: boxY, w: 0.55, h: boxH, fontSize: 9, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText("→", { x: 3.4, y: boxY, w: 0.3, h: boxH, fontSize: 10, color: C.muted, align: "center", valign: "middle", margin: 0 });
    s.addShape("rect", { x: 3.75, y: boxY, w: 0.7, h: boxH, fill: { color: C.secondary }, rectRadius: 0.03 });
    s.addText("M₂(定镜)", { x: 3.75, y: boxY, w: 0.7, h: boxH, fontSize: 8, color: C.white, align: "center", valign: "middle", margin: 0 });

    s.addText("↓", { x: 1.9, y: boxY+0.35, w: 0.55, h: 0.25, fontSize: 10, color: C.muted, align: "center", margin: 0 });
    s.addShape("rect", { x: 1.9, y: boxY+0.6, w: 0.55, h: boxH, fill: { color: C.gold } });
    s.addText("玻璃片 d", { x: 1.9, y: boxY+0.6, w: 0.55, h: boxH, fontSize: 8, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText("↓", { x: 1.9, y: boxY+0.95, w: 0.55, h: 0.25, fontSize: 10, color: C.muted, align: "center", margin: 0 });
    s.addShape("rect", { x: 1.9, y: boxY+1.2, w: 0.55, h: boxH, fill: { color: C.accent }, rectRadius: 0.03 });
    s.addText("M₁(动镜)", { x: 1.9, y: boxY+1.2, w: 0.55, h: boxH, fontSize: 8, color: C.white, align: "center", valign: "middle", margin: 0 });

    s.addText([
      { text: "元件", options: { bold:true, fontSize:9, color:C.primary } },
      { text: "\t位置\t\t说明\n", options: { fontSize:9, color:C.muted } },
      { text: "M₁", options: { bold:true, fontSize:9, color:C.accent } },
      { text: "\t上方垂直臂\t可移动镜，沿垂直方向移动\n", options: { fontSize:9, color:C.text } },
      { text: "M₂", options: { bold:true, fontSize:9, color:C.secondary } },
      { text: "\t右侧水平臂\t固定镜，位置不变\n", options: { fontSize:9, color:C.text } },
      { text: "G₁", options: { bold:true, fontSize:9, color:C.primary } },
      { text: "\t左侧45°平板\t分光板，半透半反\n", options: { fontSize:9, color:C.text } },
      { text: "G₂", options: { bold:true, fontSize:9, color:C.primary } },
      { text: "\t右侧45°平板\t补偿板，材质厚度同G₁", options: { fontSize:9, color:C.text } },
    ], { x: 5.0, y: 3.0, w: 4.5, h: 2.0, fontFace: FONT_B, lineSpacingMultiple: 1.3, margin: 0 });

    addSlideNumber(s, 3, TOTAL);
  }

  // ==================== SLIDE 4: 实验现象总览（合并红光&白光）====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "实验现象总览：红光 vs 白光");
    addFooterBar(s);

    // Left: Red Laser
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 3.85, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 0.45, fill: { color: C.accent } });
    s.addText("红色激光实验", { x: MARGIN, y: 1.15, w: 4.4, h: 0.45, fontSize: 15, fontFace: FONT_B, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

    const rSteps = [
      { n:"1", t:"光源准直", d:"He-Ne激光 + 扩束镜" },
      { n:"2", t:"调出条纹", d:"微调M₁/M₂ → 同心圆环" },
      { n:"3", t:"参考状态", d:"调至平行直条纹，记d₀" },
      { n:"4", t:"插入玻璃", d:"垂直插入厚玻璃板(7mm)" },
      { n:"5", t:"补偿调整", d:"M₁远离分光板 → 恢复平行条纹，记d₁" },
    ];
    rSteps.forEach((st, i) => {
      const y0 = 1.75 + i*0.6;
      s.addShape("ellipse", { x: MARGIN+0.2, y: y0, w: 0.3, h: 0.3, fill: { color: C.accent } });
      s.addText(st.n, { x: MARGIN+0.2, y: y0, w: 0.3, h: 0.3, fontSize: 12, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(st.t, { x: MARGIN+0.6, y: y0, w: 1.4, h: 0.28, fontSize: 11, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
      s.addText(st.d, { x: MARGIN+0.6, y: y0+0.26, w: 3.5, h: 0.28, fontSize: 10, fontFace: FONT_B, color: C.muted, margin: 0 });
    });

    s.addShape("rect", { x: MARGIN+0.2, y: 4.65, w: 4.0, h: 0.28, fill: { color: "FDE8E8" } });
    s.addText("现象：插入后条纹变回圆环；M₁远离可恢复", { x: MARGIN+0.2, y: 4.65, w: 4.0, h: 0.28, fontSize: 10, fontFace: FONT_B, color: C.accent, align: "center", valign: "middle", margin: 0 });

    // Right: White Light
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 3.85, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 0.45, fill: { color: C.gold } });
    s.addText("白色光源实验", { x: 5.1, y: 1.15, w: 4.4, h: 0.45, fontSize: 15, fontFace: FONT_B, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });

    const wSteps = [
      { n:"1", t:"切换光源", d:"白光 + 毛玻璃形成扩展光源" },
      { n:"2", t:"寻迹条纹", d:"缓慢移动M₁ → 中心彩色条纹" },
      { n:"3", t:"参考状态", d:"条纹最清晰时，记d₀" },
      { n:"4", t:"插入玻璃", d:"垂直插入薄玻璃片(0.3mm)" },
      { n:"5", t:"补偿调整", d:"M₁靠近分光板 → 恢复彩条，记d₁" },
    ];
    wSteps.forEach((st, i) => {
      const y0 = 1.75 + i*0.6;
      s.addShape("ellipse", { x: 5.3, y: y0, w: 0.3, h: 0.3, fill: { color: C.gold } });
      s.addText(st.n, { x: 5.3, y: y0, w: 0.3, h: 0.3, fontSize: 12, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(st.t, { x: 5.7, y: y0, w: 1.4, h: 0.28, fontSize: 11, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
      s.addText(st.d, { x: 5.7, y: y0+0.26, w: 3.5, h: 0.28, fontSize: 10, fontFace: FONT_B, color: C.muted, margin: 0 });
    });

    s.addShape("rect", { x: 5.3, y: 4.65, w: 4.0, h: 0.28, fill: { color: "FFF8E1" } });
    s.addText("现象：插入后彩色条纹迅速消失；M₁靠近可恢复", { x: 5.3, y: 4.65, w: 4.0, h: 0.28, fontSize: 10, fontFace: FONT_B, color: C.gold, align: "center", valign: "middle", margin: 0 });

    // Bottom formulas
    s.addShape("rect", { x: MARGIN, y: 5.05, w: 4.4, h: 0.22, fill: { color: C.codeBg } });
    s.addText("n = d / (d − Δx)", { x: MARGIN, y: 5.05, w: 4.4, h: 0.22, fontSize: 13, fontFace: "Cambria Math", color: C.accent, bold: true, italic: true, align: "center", valign: "middle", margin: 0 });
    s.addShape("rect", { x: 5.1, y: 5.05, w: 4.4, h: 0.22, fill: { color: C.codeBg } });
    s.addText("n = 1 + Δx / d", { x: 5.1, y: 5.05, w: 4.4, h: 0.22, fontSize: 13, fontFace: "Cambria Math", color: C.gold, bold: true, italic: true, align: "center", valign: "middle", margin: 0 });

    addSlideNumber(s, 4, TOTAL);
  }

  // ==================== SLIDE 5: 核心先导 — 两个必须区分的概念 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "核心先导：两个必须区分的概念");
    addFooterBar(s);

    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.3, h: 3.85, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 0.08, h: 3.85, fill: { color: C.accent } });
    s.addText("光程差 Δ 本身", { x: MARGIN+0.25, y: 1.3, w: 3.8, h: 0.35, fontSize: 18, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText([
      { text: "决定屏上某一点的亮暗（干涉级次）", options: { fontSize: 12, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "插入玻璃后追加的固定偏移：", options: { fontSize: 11, color: C.muted, breakLine: true } },
      { text: "2d(n−1)", options: { fontSize: 16, color: C.accent, bold: true, italic: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "对所有 θ 都相同，求导后为零，不改变条纹形态。", options: { fontSize: 11, color: C.text } },
    ], { x: MARGIN+0.25, y: 1.75, w: 3.8, h: 2.8, fontFace: FONT_B, lineSpacingMultiple: 1.2, valign: "top", margin: 0 });

    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.3, h: 3.85, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addShape("rect", { x: 5.1, y: 1.15, w: 0.08, h: 3.85, fill: { color: C.secondary } });
    s.addText("光程差随角度变化率 ∂Δ/∂θ", { x: 5.35, y: 1.3, w: 3.8, h: 0.35, fontSize: 18, fontFace: FONT_B, color: C.secondary, bold: true, margin: 0 });
    s.addText([
      { text: "决定条纹的整体形态", options: { fontSize: 12, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "= 0  →  光程差在全屏均匀", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "      →  平直条纹（等厚干涉）", options: { fontSize: 11, color: C.secondary, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "≠ 0  →  光程差随 θ 变化", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "      →  同心圆环（等倾干涉）", options: { fontSize: 11, color: C.secondary, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "这一区分是理解「为何激光实验条纹形态与绝对光程差无关」的钥匙。", options: { fontSize: 11, color: C.text, bold: true } },
    ], { x: 5.35, y: 1.75, w: 3.8, h: 2.8, fontFace: FONT_B, lineSpacingMultiple: 1.2, valign: "top", margin: 0 });

    addSlideNumber(s, 5, TOTAL);
  }

  // ==================== SLIDE 6: 两虚点光源模型 + 核心公式推导 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "核心原理：两虚点光源模型与 Δ(θ)≈2h cosθ");
    addFooterBar(s);

    // Left: derivation (key points only)
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.5, h: 3.85, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("关键推导", { x: MARGIN+0.2, y: 1.25, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });

    s.addText([
      { text: "两虚光源 S₁、S₂ 沿光轴相距 2h，屏距为 L（L ≫ 2h）。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 5, breakLine: true } },
      { text: "精确光程差：", options: { fontSize: 10, color: C.text, breakLine: true } },
    ], { x: MARGIN+0.2, y: 1.6, w: 4.1, h: 0.8, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    addMathBox(s, [
      { text: "Δ(θ) = √[(L+2h)²+ρ²] − √[L²+ρ²]", options: { fontSize: 11, color: C.text } },
    ], MARGIN+0.2, 2.2, 4.1, 0.35, C.codeBg);

    s.addText([
      { text: "", options: { fontSize: 5, breakLine: true } },
      { text: "分子有理化 + 近似（L ≫ 2h，略去 4h²）：", options: { fontSize: 10, color: C.text, breakLine: true } },
    ], { x: MARGIN+0.2, y: 2.6, w: 4.1, h: 0.5, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    addMathBox(s, [
      { text: "Δ(θ) ≈ 2h · L/√(L²+ρ²) = 2h cosθ", options: { fontSize: 13, color: C.accent, bold: true, italic: true } },
    ], MARGIN+0.2, 3.05, 4.1, 0.4, "FDE8E8");

    s.addText([
      { text: "", options: { fontSize: 5, breakLine: true } },
      { text: "总光程差（含劈尖角 α）：", options: { fontSize: 10, color: C.text, breakLine: true } },
    ], { x: MARGIN+0.2, y: 3.5, w: 4.1, h: 0.4, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    addMathBox(s, [
      { text: "Δ总(θ) ≈ 2h cosθ + 2xα", options: { fontSize: 12, color: C.text } },
    ], MARGIN+0.2, 3.85, 4.1, 0.35, C.codeBg);

    s.addText([
      { text: "h → 0 时，Δ总 ≈ 2xα，平直条纹由劈尖角主导。", options: { fontSize: 10, color: C.muted, breakLine: true } },
    ], { x: MARGIN+0.2, y: 4.25, w: 4.1, h: 0.4, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    // Right: image from md
    s.addShape("rect", { x: 5.2, y: 1.15, w: 4.3, h: 3.85, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("示意图（来自报告）", { x: 5.4, y: 1.25, w: 3.9, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addImage({
      path: "d:/project/迈克尔逊干涉仪实验探究/文件2/interference.png",
      x: 5.3, y: 1.6, w: 4.1, h: 2.2,
      sizing: { type: "contain", w: 4.1, h: 2.2 }
    });
    s.addText([
      { text: "S₁: 虚光源₁ (M₁臂)", options: { fontSize: 10, color: C.accent, breakLine: true } },
      { text: "S₂: 虚光源₂ (M₂臂)", options: { fontSize: 10, color: C.secondary, breakLine: true } },
      { text: "h = 0 时两虚光源重合 → 平直条纹", options: { fontSize: 10, color: C.text, bold: true } },
    ], { x: 5.4, y: 3.9, w: 3.9, h: 0.9, fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0 });

    addSlideNumber(s, 6, TOTAL);
  }

  // ==================== SLIDE 7: 形态判据 + 劈尖干涉图 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "核心原理：形态判据与物理图像");
    addFooterBar(s);

    s.addShape("rect", { x: MARGIN, y: 1.15, w: 9.0, h: 0.75, fill: { color: "FDE8E8" } });
    s.addText("核心结论", { x: MARGIN+0.2, y: 1.2, w: 1.5, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText("条纹是圆环还是平直，唯一取决于 h 是否为零 —— 与绝对光程差的大小毫无关系。实验中用红光调出平直条纹，实质上是将 M1 调到 h = 0，即两虚光源在空间上重合。",
      { x: MARGIN+0.2, y: 1.5, w: 8.6, h: 0.35, fontSize: 11, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    // Three columns
    const cols = [
      { title: "h = 0（参考状态）", color: C.secondary, lines: [
        "• 两虚光源几何重合",
        "• ∂Δ/∂θ = 0",
        "• 发散角影响消除",
        "• 平直条纹（等厚）",
        "",
        "此时若两镜面间存在",
        "微小夹角（劈尖角 α），",
        "偏离中心 x 处附加",
        "光程差 = 2xα"
      ]},
      { title: "插入玻璃后（h ≠ 0）", color: C.accent, lines: [
        "• 视深效应使虚像前移",
        "• ∂Δ/∂θ ≠ 0",
        "• 光程差随 θ 变化",
        "• 同心圆环（等倾）",
        "",
        "固定光程偏移 2d(n−1)",
        "对所有 θ 都相同，",
        "对 ∂Δ/∂θ 毫无贡献",
        "因此不改变条纹形态"
      ]},
      { title: "补偿后（h → 0）", color: "27AE60", lines: [
        "• M1 后退抵消视深位移",
        "• 虚像重新几何重合",
        "• ∂Δ/∂θ = 0",
        "• 平直条纹恢复",
        "",
        "注意：后退距离 Δx",
        "纯粹是让虚像退回",
        "h = 0，完全不涉及",
        "对固定光程偏移的补偿"
      ]},
    ];
    cols.forEach((col, i) => {
      const x0 = MARGIN + i * 3.1;
      s.addShape("rect", { x: x0, y: 2.05, w: 2.9, h: 2.55, fill: { color: C.cardBg }, shadow: makeShadow() });
      s.addText(col.title, { x: x0+0.15, y: 2.15, w: 2.6, h: 0.3, fontSize: 12, fontFace: FONT_B, color: col.color, bold: true, margin: 0 });
      s.addText(col.lines.join("\n"), { x: x0+0.15, y: 2.5, w: 2.6, h: 2.0, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
    });

    // Bottom: 劈尖干涉 image
    s.addShape("rect", { x: MARGIN, y: 4.75, w: 9.0, h: 0.5, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addImage({
      path: "d:/project/迈克尔逊干涉仪实验探究/f4052a3a-3c7f-49b1-8fa6-93036d5ef15b.png",
      x: MARGIN+0.2, y: 4.8, w: 1.8, h: 0.4,
      sizing: { type: "contain", w: 1.8, h: 0.4 }
    });
    s.addText("劈尖干涉示意图（来自报告）：h→0 时，平直条纹由劈尖角 α 主导，偏离中心 x 处附加光程差 2xα。",
      { x: MARGIN+2.1, y: 4.8, w: 6.7, h: 0.4, fontSize: 10, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });

    addSlideNumber(s, 7, TOTAL);
  }

  // ==================== SLIDE 8: 红光原理 — 双重效应与补偿 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "红光原理：视深效应与补偿逻辑");
    addFooterBar(s);

    // Table
    const header = [
      { text: "效应", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
      { text: "机制", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
      { text: "对 Δ(θ) 的影响", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
      { text: "对 ∂Δ/∂θ 的影响", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
    ];
    const rows = [
      [
        { text: "减速效应", options: { fontSize:11, bold:true, color:C.text, align:"center" } },
        { text: "光在玻璃中速度变慢", options: { fontSize:11, color:C.text, align:"center" } },
        { text: "追加固定偏移 2d(n−1)", options: { fontSize:11, color:C.text, align:"center" } },
        { text: "为零（常数求导）", options: { fontSize:11, color:C.text, align:"center" } },
      ],
      [
        { text: "视深效应", options: { fontSize:11, bold:true, color:C.accent, align:"center" } },
        { text: "折射使 M1 虚像位置前移", options: { fontSize:11, color:C.text, align:"center" } },
        { text: "改变 h（h 不再为 0）", options: { fontSize:11, color:C.accent, align:"center" } },
        { text: "非零（h≠0）", options: { fontSize:11, color:C.accent, bold:true, align:"center" } },
      ],
    ];
    s.addTable([header, ...rows], {
      x: MARGIN, y: 1.15, w: 9.0,
      colW: [1.5, 2.5, 2.5, 2.5],
      rowH: [0.4, 0.4, 0.4],
      border: { pt: 0.5, color: "DDDDDD" },
      margin: [3, 5, 3, 5],
    });

    // Left: apparent depth
    s.addShape("rect", { x: MARGIN, y: 2.25, w: 4.4, h: 2.3, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("视深效应", { x: MARGIN+0.2, y: 2.35, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText("虚像前移量：",
      { x: MARGIN+0.2, y: 2.7, w: 4.0, h: 0.3, fontSize: 11, fontFace: FONT_B, color: C.text, margin: 0 });
    addMathBox(s, [
      { text: "Δx' = d(1 − 1/n)", options: { fontSize: 16, color: C.accent, bold: true, italic: true } },
    ], MARGIN+0.2, 3.05, 4.0, 0.45, C.codeBg);
    s.addText("直接破坏 h=0 → 条纹退化为同心圆环",
      { x: MARGIN+0.2, y: 3.55, w: 4.0, h: 0.35, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0 });

    // Right: compensation
    s.addShape("rect", { x: 5.1, y: 2.25, w: 4.4, h: 2.3, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("补偿逻辑", { x: 5.3, y: 2.35, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "M1 向后（远离分光板）移动 Δx = Δx'", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "→ 虚像重新几何重合 → h = 0", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "→ 平直条纹恢复", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 5, breakLine: true } },
      { text: "不涉及对固定光程偏移 2d(n−1) 的补偿", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "激光相干长度 ≫ 该偏移量", options: { fontSize: 10, color: C.muted } },
    ], { x: 5.3, y: 2.7, w: 4.0, h: 1.7, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    // Bottom formula
    s.addShape("rect", { x: MARGIN, y: 4.7, w: 9.0, h: 0.55, fill: { color: "FDE8E8" } });
    s.addText("折射率公式", { x: MARGIN+0.2, y: 4.75, w: 1.5, h: 0.25, fontSize: 12, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    addMathBox(s, [
      { text: "n = d / (d − Δx)", options: { fontSize: 18, color: C.accent, bold: true, italic: true } },
    ], MARGIN+3.5, 4.7, 3.0, 0.55, "FDE8E8");

    addSlideNumber(s, 8, TOTAL);
  }

  // ==================== SLIDE 9: 白光原理 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "白光原理：零光程差约束与补偿推导");
    addFooterBar(s);

    // Left: coherence constraint
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText('白光的「光程极零」强制约', { x: MARGIN+0.2, y: 1.25, w: 3.8, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "白光属宽频带复色光，相干长度仅微米量级。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "核心要求：OPD ≈ 0", options: { fontSize: 12, color: C.gold, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "若 OPD 不近零，视场中根本不存在任何条纹，讨论形态无从谈起。", options: { fontSize: 10, color: C.text } },
    ], { x: MARGIN+0.2, y: 1.65, w: 4.0, h: 1.5, fontFace: FONT_B, lineSpacingMultiple: 1.2, valign: "top", margin: 0 });

    // Right: OPD jump
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("插入玻片引发的光程跳变", { x: 5.3, y: 1.25, w: 3.8, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText("光束两次穿过厚度 d 的玻璃片，该臂总光程跳变：",
      { x: 5.3, y: 1.65, w: 4.0, h: 0.35, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    addMathBox(s, [
      { text: "Δ(OPD) = 2d(n−1)", options: { fontSize: 16, color: C.gold, bold: true, italic: true } },
    ], 5.3, 2.05, 4.0, 0.45, C.codeBg);

    s.addText([
      { text: "推导：", options: { bold:true, fontSize:10, color:C.text, breakLine:true } },
      { text: "• 单程光程增量 = nd − d = d(n−1)", options: { fontSize:10, color:C.text, breakLine:true } },
      { text: "• 往返总增量 = 2d(n−1)", options: { fontSize:10, color:C.text, breakLine:true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "这一跳变远超白光相干长度，彩色条纹迅速消失。", options: { fontSize: 10, color: C.accent, bold: true } },
    ], { x: 5.3, y: 2.55, w: 4.0, h: 1.0, fontFace: FONT_B, lineSpacingMultiple: 1.15, margin: 0 });

    // Bottom: compensation
    s.addShape("rect", { x: MARGIN, y: 3.9, w: 9.0, h: 1.3, fill: { color: "FFF8E1" } });
    s.addText("补偿推导", { x: MARGIN+0.2, y: 3.95, w: 1.5, h: 0.25, fontSize: 12, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    s.addText("M1 向分光板方向推进 Δx → 光程减少 2Δx。由补偿守恒：",
      { x: MARGIN+0.2, y: 4.2, w: 4.5, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0 });
    addMathBox(s, [
      { text: "2Δx = 2d(n−1)", options: { fontSize: 14, color: C.gold, bold: true, italic: true } },
    ], MARGIN+4.8, 3.95, 2.5, 0.5, "FFF8E1");
    addMathBox(s, [
      { text: "n = 1 + Δx / d", options: { fontSize: 16, color: C.gold, bold: true, italic: true } },
    ], MARGIN+7.4, 3.95, 1.5, 0.5, "FFF8E1");
    s.addText("白光仅要求 OPD≈0，不要求 h=0。补偿后残留 h≈0.06mm，条纹曲率极小，近似平行。",
      { x: MARGIN+0.2, y: 4.7, w: 8.6, h: 0.35, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0 });

    addSlideNumber(s, 9, TOTAL);
  }

  // ==================== SLIDE 10: 原理对比与厚度选择 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "原理对比与玻璃厚度选择的物理必然性");
    addFooterBar(s);

    // Comparison table
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 9.0, h: 2.0, fill: { color: C.cardBg }, shadow: makeShadow() });
    const compHeader = [
      { text: "", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11 } },
      { text: "红光实验", options: { fill:{color:C.accent}, color:C.white, bold:true, fontSize:11, align:"center" } },
      { text: "白光实验", options: { fill:{color:C.gold}, color:C.white, bold:true, fontSize:11, align:"center" } },
    ];
    const compRows = [
      [
        { text: "物理机制", options: { bold:true, fontSize:10, color:C.text } },
        { text: "视深效应 → 虚像位移", options: { fontSize:10, color:C.text, align:"center" } },
        { text: "光程跳变 → OPD补偿", options: { fontSize:10, color:C.text, align:"center" } },
      ],
      [
        { text: "M1 移动方向", options: { bold:true, fontSize:10, color:C.text } },
        { text: "远离分光板 ←", options: { fontSize:10, color:C.accent, bold:true, align:"center" } },
        { text: "→ 靠近分光板", options: { fontSize:10, color:C.gold, bold:true, align:"center" } },
      ],
      [
        { text: "玻璃厚度", options: { bold:true, fontSize:10, color:C.text } },
        { text: "7 mm（厚）", options: { fontSize:10, color:C.text, align:"center" } },
        { text: "0.3 mm（薄）", options: { fontSize:10, color:C.text, align:"center" } },
      ],
      [
        { text: "折射率公式", options: { bold:true, fontSize:10, color:C.text } },
        { text: "n = d/(d−Δx)", options: { fontSize:10, color:C.accent, italic:true, align:"center" } },
        { text: "n = 1+Δx/d", options: { fontSize:10, color:C.gold, italic:true, align:"center" } },
      ],
      [
        { text: "关键约束", options: { bold:true, fontSize:10, color:C.text } },
        { text: "h = 0（几何条件）", options: { fontSize:10, color:C.text, align:"center" } },
        { text: "OPD ≈ 0（相干条件）", options: { fontSize:10, color:C.text, align:"center" } },
      ],
    ];
    s.addTable([compHeader, ...compRows], {
      x: MARGIN+0.2, y: 1.3, w: 8.6,
      colW: [1.6, 3.5, 3.5],
      rowH: [0.35, 0.3, 0.3, 0.3, 0.3, 0.3],
      border: { pt: 0.5, color: "DDDDDD" },
      margin: [2, 5, 2, 5],
    });

    // Bottom: thickness analysis
    s.addShape("rect", { x: MARGIN, y: 3.3, w: 9.0, h: 1.8, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("玻璃厚度选择的物理必然性", { x: MARGIN+0.2, y: 3.35, w: 4.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });

    s.addText([
      { text: "激光实验（d = 7mm）：", options: { bold: true, fontSize: 11, color: C.accent, breakLine: true } },
      { text: "虚像位移 Δx' = 7×(1−1/1.5) ≈ 2.33 mm，千分尺上 2mm 量级位移易于精确读取。玻璃越厚信号越大，且激光不在乎相伴的固定光程偏移——", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "厚玻璃在激光实验中纯属优势。", options: { fontSize: 10, color: C.accent, bold: true } },
    ], { x: MARGIN+0.2, y: 3.7, w: 4.2, h: 1.2, fontFace: FONT_B, lineSpacingMultiple: 1.15, margin: 0 });

    s.addText([
      { text: "白光实验（d = 0.3mm）：", options: { bold: true, fontSize: 11, color: C.gold, breakLine: true } },
      { text: "OPD 跳变 = 2×0.3×0.5 = 0.3 mm，补偿位移 ≈ 0.15 mm，仍在可追踪范围。若用 7mm 厚玻璃，OPD 跳变将达 7mm 量级——远超白光微米级相干长度，实验上几乎不可能重新捕捉彩色条纹。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "白光实验的玻璃必须薄到让 OPD 跳变不超出可操作范围。", options: { fontSize: 10, color: C.gold, bold: true } },
    ], { x: MARGIN+4.7, y: 3.7, w: 4.2, h: 1.2, fontFace: FONT_B, lineSpacingMultiple: 1.15, margin: 0 });

    addSlideNumber(s, 10, TOTAL);
  }

  // ==================== SLIDE 11: 数据处理 红光 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "数据处理：红光实验（He-Ne 激光）");
    addFooterBar(s);

    // Left: raw data
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 2.2, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("原始数据（d = 7.0 mm）", { x: MARGIN+0.2, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });

    const dataHeader = [
      { text: "次数", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
      { text: "d₀ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
      { text: "d₁ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
      { text: "Δxᵢ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
    ];
    const dataRows = [
      [{ text:"1", options:{fontSize:9, align:"center"} }, { text:"54.32107", options:{fontSize:9, align:"center"} }, { text:"57.11739", options:{fontSize:9, align:"center"} }, { text:"2.79632", options:{fontSize:9, align:"center", color:C.accent} }],
      [{ text:"2", options:{fontSize:9, align:"center"} }, { text:"54.28965", options:{fontSize:9, align:"center"} }, { text:"57.09380", options:{fontSize:9, align:"center"} }, { text:"2.80415", options:{fontSize:9, align:"center", color:C.accent} }],
      [{ text:"3", options:{fontSize:9, align:"center"} }, { text:"54.35782", options:{fontSize:9, align:"center"} }, { text:"57.14769", options:{fontSize:9, align:"center"} }, { text:"2.78987", options:{fontSize:9, align:"center", color:C.accent} }],
    ];
    s.addTable([dataHeader, ...dataRows], {
      x: MARGIN+0.2, y: 1.55, w: 4.0,
      colW: [0.7, 1.3, 1.3, 1.3],
      rowH: [0.28, 0.28, 0.28, 0.28],
      border: { pt: 0.5, color: "DDDDDD" },
      margin: [1, 3, 1, 3],
    });
    s.addText("平均值 Δx̄ = 2.79678 mm", { x: MARGIN+0.2, y: 2.55, w: 3.0, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });

    // Right: calculation
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.2, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("计算过程", { x: 5.3, y: 1.2, w: 2.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "样本标准差：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "s_Δx = 0.00716 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "A类标准不确定度：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(Δx̄) = s_Δx / √3 = 0.00413 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "由 n = d/(d−Δx) 得：", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "n_红 = 7.0 / (7.0−2.79678) ≈ 1.6655", options: { fontSize: 11, color: C.accent, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "不确定度传播：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(n) = |∂n/∂Δx|·u(Δx̄) = d/(d−Δx̄)² · u(Δx̄)", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "u(n_红) ≈ 0.0016", options: { fontSize: 10, color: C.accent, bold: true } },
    ], { x: 5.3, y: 1.55, w: 4.0, h: 1.7, fontFace: FONT_B, lineSpacingMultiple: 1.1, valign: "top", margin: 0 });

    // Bottom result
    s.addShape("rect", { x: MARGIN, y: 3.55, w: 9.0, h: 1.15, fill: { color: "FDE8E8" } });
    s.addText("红光实验结果", { x: MARGIN+0.2, y: 3.6, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    addMathBox(s, [
      { text: "n_红 = 1.6655 ± 0.0016", options: { fontSize: 24, color: C.accent, bold: true, italic: true } },
    ], MARGIN+3.0, 3.6, 4.0, 1.0, "FDE8E8");

    addSlideNumber(s, 11, TOTAL);
  }

  // ==================== SLIDE 12: 数据处理 白光 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "数据处理：白光实验");
    addFooterBar(s);

    // Left: raw data
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 2.2, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("原始数据（d = 0.3 mm）", { x: MARGIN+0.2, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });

    const dataHeader2 = [
      { text: "次数", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
      { text: "d₀ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
      { text: "d₁ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
      { text: "Δxᵢ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
    ];
    const dataRows2 = [
      [{ text:"1", options:{fontSize:9, align:"center"} }, { text:"51.94480", options:{fontSize:9, align:"center"} }, { text:"51.77279", options:{fontSize:9, align:"center"} }, { text:"0.17201", options:{fontSize:9, align:"center", color:C.gold} }],
      [{ text:"2", options:{fontSize:9, align:"center"} }, { text:"51.93326", options:{fontSize:9, align:"center"} }, { text:"51.76271", options:{fontSize:9, align:"center"} }, { text:"0.17055", options:{fontSize:9, align:"center", color:C.gold} }],
      [{ text:"3", options:{fontSize:9, align:"center"} }, { text:"51.95813", options:{fontSize:9, align:"center"} }, { text:"51.78424", options:{fontSize:9, align:"center"} }, { text:"0.17389", options:{fontSize:9, align:"center", color:C.gold} }],
    ];
    s.addTable([dataHeader2, ...dataRows2], {
      x: MARGIN+0.2, y: 1.55, w: 4.0,
      colW: [0.7, 1.3, 1.3, 1.3],
      rowH: [0.28, 0.28, 0.28, 0.28],
      border: { pt: 0.5, color: "DDDDDD" },
      margin: [1, 3, 1, 3],
    });
    s.addText("平均值 Δx̄ = 0.17215 mm", { x: MARGIN+0.2, y: 2.55, w: 3.0, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });

    // Right: calculation
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.2, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("计算过程", { x: 5.3, y: 1.2, w: 2.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "样本标准差：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "s_Δx ≈ 0.00168 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "A类标准不确定度：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(Δx̄) = s_Δx / √3 ≈ 0.00097 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "由 n = 1 + Δx/d 得：", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "n_白 = 1 + 0.17215/0.3 ≈ 1.5738", options: { fontSize: 11, color: C.gold, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "不确定度传播：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(n) = u(Δx̄) / d ≈ 0.00097 / 0.3 ≈ 0.0032", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "u(n_白) ≈ 0.0032", options: { fontSize: 10, color: C.gold, bold: true } },
    ], { x: 5.3, y: 1.55, w: 4.0, h: 1.7, fontFace: FONT_B, lineSpacingMultiple: 1.1, valign: "top", margin: 0 });

    // Bottom result
    s.addShape("rect", { x: MARGIN, y: 3.55, w: 9.0, h: 1.15, fill: { color: "FFF8E1" } });
    s.addText("白光实验结果", { x: MARGIN+0.2, y: 3.6, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    addMathBox(s, [
      { text: "n_白 = 1.5738 ± 0.0032", options: { fontSize: 24, color: C.gold, bold: true, italic: true } },
    ], MARGIN+3.0, 3.6, 4.0, 1.0, "FFF8E1");

    addSlideNumber(s, 12, TOTAL);
  }

  // ==================== SLIDE 13: 误差分析 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "误差分析");
    addFooterBar(s);

    // Red errors
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.3, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("红光实验误差源", { x: MARGIN+0.2, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText([
      { text: "• 条纹判断误差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  「平直程度」依赖主观判断", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "• 千分尺读数误差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  回差或视差影响 Δx", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "• 玻璃板放置偏差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  未严格垂直入射 → 系统偏差", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "• 厚度 d 测量误差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  d 偏差会被公式放大或缩小", options: { fontSize: 10, color: C.muted } },
    ], { x: MARGIN+0.2, y: 1.55, w: 3.9, h: 2.1, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    // White errors
    s.addShape("rect", { x: 5.0, y: 1.15, w: 4.3, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("白光实验误差源", { x: 5.2, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    s.addText([
      { text: "• 相干条件更苛刻：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  条纹消失/重现判定更敏感", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "• 光源对准不稳：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  扩展光源偏差 → 彩条不清晰", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "• 薄片厚度偏差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  d仅0.3mm, 微小偏差被显著放大", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "  例如 d 偏差 0.02mm 即可导致 n 变化约 0.1", options: { fontSize: 10, color: C.muted } },
    ], { x: 5.2, y: 1.55, w: 3.9, h: 2.1, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    // Bottom: overall judgment
    s.addShape("rect", { x: MARGIN, y: 3.9, w: 9.0, h: 1.05, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("总体判断", { x: MARGIN+0.2, y: 3.95, w: 1.5, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText("n_红 = 1.6655 ± 0.0016，n_白 = 1.5738 ± 0.0032，两者均落在常见玻璃折射率范围内。微小差异可能来源于两块玻璃材质不同，也可能来自两种原理各自的系统偏差。两个实验从不同物理机制出发，均得到了量级合理的结果，相互印证了两种补偿逻辑（远离 vs 靠近）的正确性。",
      { x: MARGIN+0.2, y: 4.25, w: 8.6, h: 0.65, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    addSlideNumber(s, 13, TOTAL);
  }

  // ==================== SLIDE 14: 结论 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
    s.addText("结论", { x: MARGIN, y: 0.25, w: 8.6, h: 0.55, fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0 });

    const conclusions = [
      {
        num: "01",
        title: "红光实验结论",
        text: "插入厚玻璃板后需将 M1 远离分光板，以补偿视深效应并恢复等厚干涉的几何条件（h=0）。折射率公式 n = d/(d−Δx)，计算得 n_红 = 1.6655 ± 0.0016。",
      },
      {
        num: "02",
        title: "白光实验结论",
        text: "插入薄玻璃片后需将 M1 向分光板方向移动，以补偿增加的光程差并回到 OPD ≈ 0。折射率公式 n = 1 + Δx/d，计算得 n_白 = 1.5738 ± 0.0032。",
      },
      {
        num: "03",
        title: "方向差异的物理根源",
        text: "激光相干长度极长，只响应虚像间距 h（几何条件 ∂Δ/∂θ）；白光相干长度极短，只响应绝对光程差 OPD ≈ 0。两种光源对干涉条件的选择性不同，导致了完全相反的补偿方向。",
      },
    ];

    conclusions.forEach((c, i) => {
      const y0 = 1.1 + i * 1.3;
      s.addShape("rect", { x: MARGIN, y: y0+0.05, w: 0.55, h: 0.55, fill: { color: C.primary } });
      s.addText(c.num, { x: MARGIN, y: y0+0.05, w: 0.55, h: 0.55, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(c.title, { x: MARGIN+0.75, y: y0, w: 8.0, h: 0.35, fontSize: 15, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
      s.addText(c.text, { x: MARGIN+0.75, y: y0+0.38, w: 8.0, h: 0.55, fontSize: 11, fontFace: FONT_B, color: "D0D8E0", lineSpacingMultiple: 1.3, margin: 0 });
    });

    addSlideNumber(s, 14, TOTAL, true);
  }

  // ==================== SLIDE 15: 致谢 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
    s.addText("谢谢！", { x: 1.0, y: 1.4, w: 8.0, h: 1.2, fontSize: 52, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addShape("rect", { x: 3.2, y: 2.8, w: 3.6, h: 0.03, fill: { color: C.gold } });
    s.addText("感谢聆听 · 欢迎提问", { x: 1.0, y: 3.1, w: 8.0, h: 0.6, fontSize: 18, fontFace: FONT_B, color: "AABBCC", align: "center", margin: 0 });
    s.addText("迈克尔逊干涉仪测折射率实验探究", { x: 1.0, y: 4.1, w: 8.0, h: 0.4, fontSize: 13, fontFace: FONT_B, color: "667788", align: "center", margin: 0 });
    addSlideNumber(s, 15, TOTAL, true);
  }

  // ─── Write file ───
  const outPath = "d:/project/迈克尔逊干涉仪实验探究/答辩PPT_V3.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("PPT saved to: " + outPath);
}

main().catch(err => { console.error(err); process.exit(1); });
