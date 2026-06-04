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
  const TOTAL = 18;
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
      { num: "02", title: "实验操作与现象", desc: "红光 & 白光两种实验流程与观测" },
      { num: "03", title: "核心原理：形态判据", desc: "两虚点光源模型与 Δ(θ)≈2h cos θ 推导" },
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

    // Draw boxes for components
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

    // Vertical arm
    s.addText("↓", { x: 1.9, y: boxY+0.35, w: 0.55, h: 0.25, fontSize: 10, color: C.muted, align: "center", margin: 0 });
    s.addShape("rect", { x: 1.9, y: boxY+0.6, w: 0.55, h: boxH, fill: { color: C.gold } });
    s.addText("玻璃片 d", { x: 1.9, y: boxY+0.6, w: 0.55, h: boxH, fontSize: 8, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText("↓", { x: 1.9, y: boxY+0.95, w: 0.55, h: 0.25, fontSize: 10, color: C.muted, align: "center", margin: 0 });
    s.addShape("rect", { x: 1.9, y: boxY+1.2, w: 0.55, h: boxH, fill: { color: C.accent }, rectRadius: 0.03 });
    s.addText("M₁(动镜)", { x: 1.9, y: boxY+1.2, w: 0.55, h: boxH, fontSize: 8, color: C.white, align: "center", valign: "middle", margin: 0 });

    // Table
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

  // ==================== SLIDE 4: 红光实验操作与现象 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "红光实验：操作与现象");
    addFooterBar(s);

    const steps = [
      { title: "光源与准直", desc: "开启氦氖激光，光束通过扩束镜，形成近似点光源与发散光束。" },
      { title: "条纹调出", desc: "微调 M1 与 M2 的角度，使观察屏上出现同心圆环条纹。" },
      { title: "参考状态记录", desc: "不插入玻璃板，调整 M1 使圆环逐渐过渡为平行直条纹，记录位置 d₀。" },
      { title: "插入厚玻璃板", desc: "在 M1 臂中垂直插入厚玻璃板（d=7mm），保持玻璃板面与光束基本垂直。" },
      { title: "条纹变化观察", desc: "条纹重新变成同心圆环条纹。" },
      { title: "补偿调整", desc: "将 M1 向远离分光板方向移动，使平行条纹重新出现且清晰，记录 d₁。" },
    ];
    steps.forEach((st, i) => {
      const col = i < 3 ? 0 : 1;
      const row = i % 3;
      const x0 = MARGIN + col * 4.6;
      const y0 = 1.25 + row * 1.15;
      s.addShape("ellipse", { x: x0, y: y0+0.05, w: 0.35, h: 0.35, fill: { color: C.accent } });
      s.addText(String(i+1), { x: x0, y: y0+0.05, w: 0.35, h: 0.35, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(st.title, { x: x0+0.45, y: y0, w: 3.8, h: 0.3, fontSize: 12, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
      s.addText(st.desc, { x: x0+0.45, y: y0+0.28, w: 3.8, h: 0.5, fontSize: 10, fontFace: FONT_B, color: C.muted, lineSpacingMultiple: 1.2, margin: 0 });
    });

    // Bottom: observation + formula
    s.addShape("rect", { x: MARGIN, y: 4.55, w: 9.0, h: 0.7, fill: { color: "FDE8E8" } });
    s.addText("观测现象", { x: MARGIN+0.2, y: 4.6, w: 1.2, h: 0.25, fontSize: 12, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText("插入玻璃板后，条纹重新变成同心圆环；将 M1 向远离分光板方向移动，可重新找到清晰平行条纹。",
      { x: MARGIN+0.2, y: 4.85, w: 6.0, h: 0.35, fontSize: 11, fontFace: FONT_B, color: C.text, margin: 0 });
    s.addText("n = d / (d \u2212 \u0394x)", { x: 7.5, y: 4.55, w: 2.0, h: 0.7, fontSize: 16, fontFace: "Cambria Math", color: C.accent, bold: true, italic: true, align: "center", valign: "middle", margin: 0 });

    addSlideNumber(s, 4, TOTAL);
  }

  // ==================== SLIDE 5: 白光实验操作与现象 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "白光实验：操作与现象");
    addFooterBar(s);

    const steps = [
      { title: "光源切换", desc: "更换为白光光源，加入毛玻璃片以形成扩展光源，调整对准使光束进入干涉仪。" },
      { title: "条纹寻迹", desc: "缓慢移动 M1，在视场中寻找中心彩色条纹，并使其清晰可辨。" },
      { title: "参考状态记录", desc: "确认条纹最清晰时的 M1 位置，记录为参考读数 d₀。" },
      { title: "插入薄玻璃片", desc: "在 M1 臂中垂直插入薄玻璃片（d=0.3mm）。" },
      { title: "条纹变化观察", desc: "彩色条纹迅速消失。" },
      { title: "补偿调整", desc: "缓慢将 M1 向靠近分光板方向移动，直至彩色条纹重新出现，记录 d₁。" },
    ];
    steps.forEach((st, i) => {
      const col = i < 3 ? 0 : 1;
      const row = i % 3;
      const x0 = MARGIN + col * 4.6;
      const y0 = 1.25 + row * 1.15;
      s.addShape("ellipse", { x: x0, y: y0+0.05, w: 0.35, h: 0.35, fill: { color: C.gold } });
      s.addText(String(i+1), { x: x0, y: y0+0.05, w: 0.35, h: 0.35, fontSize: 14, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(st.title, { x: x0+0.45, y: y0, w: 3.8, h: 0.3, fontSize: 12, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
      s.addText(st.desc, { x: x0+0.45, y: y0+0.28, w: 3.8, h: 0.5, fontSize: 10, fontFace: FONT_B, color: C.muted, lineSpacingMultiple: 1.2, margin: 0 });
    });

    s.addShape("rect", { x: MARGIN, y: 4.55, w: 9.0, h: 0.7, fill: { color: "FFF8E1" } });
    s.addText("观测现象", { x: MARGIN+0.2, y: 4.6, w: 1.2, h: 0.25, fontSize: 12, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    s.addText("插入玻璃片后彩色条纹会消失；通过将 M1 向靠近分光板方向移动，可重新找到清晰彩色平行条纹。",
      { x: MARGIN+0.2, y: 4.85, w: 6.0, h: 0.35, fontSize: 11, fontFace: FONT_B, color: C.text, margin: 0 });
    s.addText("n = 1 + \u0394x / d", { x: 7.5, y: 4.55, w: 2.0, h: 0.7, fontSize: 16, fontFace: "Cambria Math", color: C.gold, bold: true, italic: true, align: "center", valign: "middle", margin: 0 });

    addSlideNumber(s, 5, TOTAL);
  }

  // ==================== SLIDE 6: 核心原理(1) 两虚点光源模型 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "核心原理(1)：两虚点光源模型");
    addFooterBar(s);

    s.addShape("rect", { x: MARGIN, y: 1.15, w: 9.0, h: 1.0, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("关键区分", { x: MARGIN+0.2, y: 1.2, w: 1.5, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText([
      { text: "光程差 \u0394 本身", options: { bold:true, color:C.accent, fontSize:11 } },
      { text: " \u2192 决定屏上某一点的亮暗（干涉级次）\n", options: { fontSize:11, color:C.text } },
      { text: "\u2202\u0394/\u2202\u03B8", options: { bold:true, color:C.secondary, fontSize:11 } },
      { text: " \u2192 决定条纹的整体形态。=0 为平直条纹，\u22600 为同心圆环。", options: { fontSize:11, color:C.text } },
    ], { x: MARGIN+0.2, y: 1.55, w: 8.6, h: 0.55, fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0 });

    // Left: model description
    s.addShape("rect", { x: MARGIN, y: 2.3, w: 4.5, h: 2.9, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("几何模型", { x: MARGIN+0.2, y: 2.4, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "经分光板 G₁ 分光后，同一光源在 M₁ 与 M₂ 后各形成一个虚点光源 S₁、S₂，两者沿光轴方向相距 2h（h 为动镜 M₁ 与定镜虚像 M₂' 之间的间距）。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "观察屏距 S₂ 为 L（L \u226B 2h），屏上一点 P 到轴距离为 \u03C1，视角 \u03B8 满足 cos\u03B8 = L / \u221a(L²+\u03C1²)。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "两束光的光程差精确为：", options: { fontSize: 10, color: C.text } },
    ], { x: MARGIN+0.2, y: 2.75, w: 4.1, h: 1.3, fontFace: FONT_B, lineSpacingMultiple: 1.2, valign: "top", margin: 0 });

    addMathBox(s, [
      { text: "\u0394(\u03B8) = \u221a[(L+2h)\u00B2 + \u03C1\u00B2] \u2212 \u221a[L\u00B2 + \u03C1\u00B2]", options: { fontSize: 13, color: C.text } },
    ], MARGIN+0.2, 3.85, 4.1, 0.45, C.codeBg);

    s.addText("将分子有理化，分子分母同乘共轭式：", {
      x: MARGIN+0.2, y: 4.4, w: 4.1, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.muted, margin: 0
    });

    addMathBox(s, [
      { text: "\u0394(\u03B8) = [(L+2h)\u00B2+\u03C1\u00B2 \u2212 L\u00B2\u2212\u03C1\u00B2] / [\u221a(L+2h)\u00B2+\u03C1\u00B2 + \u221aL\u00B2+\u03C1\u00B2]", options: { fontSize: 11, color: C.text } },
      { text: "     = (4Lh + 4h\u00B2) / [\u221a(L+2h)\u00B2+\u03C1\u00B2 + \u221aL\u00B2+\u03C1\u00B2]", options: { fontSize: 11, color: C.text } },
    ], MARGIN+0.2, 4.7, 4.1, 0.55, C.codeBg);

    // Right: diagram
    s.addShape("rect", { x: 5.2, y: 2.3, w: 4.3, h: 2.9, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("示意图", { x: 5.4, y: 2.4, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });

    // Draw S1 and S2
    s.addShape("ellipse", { x: 5.6, y: 3.0, w: 0.2, h: 0.2, fill: { color: C.accent } });
    s.addText("S₁", { x: 5.6, y: 2.82, w: 0.2, h: 0.18, fontSize: 9, fontFace: FONT_B, color: C.accent, align: "center", margin: 0 });
    s.addShape("ellipse", { x: 6.8, y: 3.0, w: 0.2, h: 0.2, fill: { color: C.secondary } });
    s.addText("S₂", { x: 6.8, y: 2.82, w: 0.2, h: 0.18, fontSize: 9, fontFace: FONT_B, color: C.secondary, align: "center", margin: 0 });
    s.addText("\u2190 2h \u2192", { x: 5.85, y: 2.95, w: 0.9, h: 0.2, fontSize: 10, fontFace: FONT_B, color: C.muted, align: "center", margin: 0 });
    s.addShape("line", { x: 5.85, y: 3.12, w: 0.9, h: 0, line: { color: C.muted, width: 1, dashType: "dash" } });

    s.addText("→", { x: 7.1, y: 2.95, w: 0.3, h: 0.2, fontSize: 12, color: C.muted, align: "center", margin: 0 });
    s.addText("L", { x: 7.35, y: 2.95, w: 0.2, h: 0.2, fontSize: 10, fontFace: FONT_B, color: C.muted, align: "center", margin: 0 });
    s.addShape("rect", { x: 7.8, y: 2.8, w: 0.06, h: 0.6, fill: { color: C.muted } });
    s.addText("屏", { x: 7.9, y: 3.0, w: 0.3, h: 0.2, fontSize: 10, fontFace: FONT_B, color: C.muted, margin: 0 });

    s.addShape("ellipse", { x: 7.95, y: 3.5, w: 0.08, h: 0.08, fill: { color: C.text } });
    s.addText("P (\u03C1)", { x: 8.05, y: 3.4, w: 0.4, h: 0.18, fontSize: 9, fontFace: FONT_B, color: C.text, margin: 0 });
    s.addShape("line", { x: 7.83, y: 3.12, w: 0.12, h: 0.38, line: { color: C.muted, width: 1 } });
    s.addText("\u03B8", { x: 7.6, y: 3.35, w: 0.25, h: 0.2, fontSize: 9, fontFace: FONT_B, color: C.muted, align: "center", margin: 0 });

    s.addText([
      { text: "S₁: 虚光源₁ (M₁臂)", options: { fontSize: 10, color: C.accent, breakLine: true } },
      { text: "S₂: 虚光源₂ (M₂臂)", options: { fontSize: 10, color: C.secondary, breakLine: true } },
      { text: "L \u226B 2h，\u03C1 为屏上点到轴距离", options: { fontSize: 10, color: C.muted } },
    ], { x: 5.4, y: 3.8, w: 3.9, h: 0.8, fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0 });

    addSlideNumber(s, 6, TOTAL);
  }

  // ==================== SLIDE 7: 核心原理(2) Δ(θ)≈2h cosθ 推导 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "核心原理(2)：\u0394(\u03B8) \u2248 2h cos\u03B8 推导");
    addFooterBar(s);

    // Step 1
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.3, h: 2.0, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("Step 1：近似处理", { x: MARGIN+0.2, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.primary, bold: true, margin: 0 });
    s.addText("利用 L \u226B 2h，分母中 (L+2h)\u00B2 \u2248 L\u00B2，故：", {
      x: MARGIN+0.2, y: 1.55, w: 3.9, h: 0.3, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0
    });
    addMathBox(s, [
      { text: "\u221a(L+2h)\u00B2+\u03C1\u00B2 + \u221aL\u00B2+\u03C1\u00B2 \u2248 2\u221aL\u00B2+\u03C1\u00B2", options: { fontSize: 11, color: C.text } },
    ], MARGIN+0.2, 1.9, 3.9, 0.35, C.codeBg);
    s.addText("同时分子中 4h\u00B2 项可略去：", {
      x: MARGIN+0.2, y: 2.3, w: 3.9, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0
    });
    addMathBox(s, [
      { text: "\u0394(\u03B8) \u2248 4Lh / (2\u221aL\u00B2+\u03C1\u00B2)", options: { fontSize: 11, color: C.text } },
      { text: "     = 2h \u00B7 L / \u221aL\u00B2+\u03C1\u00B2", options: { fontSize: 11, color: C.text } },
    ], MARGIN+0.2, 2.6, 3.9, 0.5, C.codeBg);

    // Step 2
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.3, h: 2.0, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("Step 2：核心公式", { x: 5.3, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText("由 cos\u03B8 = L / \u221a(L\u00B2+\u03C1\u00B2)，代入得：", {
      x: 5.3, y: 1.55, w: 3.9, h: 0.3, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0
    });
    addMathBox(s, [
      { text: "\u0394(\u03B8) \u2248 2h cos\u03B8", options: { fontSize: 18, color: C.accent, bold: true, italic: true } },
    ], 5.3, 1.9, 3.9, 0.55, "FDE8E8");
    s.addText("这是理解整个实验形态判据的核心公式。", {
      x: 5.3, y: 2.55, w: 3.9, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.muted, margin: 0
    });

    // Step 3: angle dependence
    s.addShape("rect", { x: MARGIN, y: 3.3, w: 9.0, h: 1.85, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("角度依赖性分析", { x: MARGIN+0.2, y: 3.35, w: 3.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });

    // Left column
    s.addText([
      { text: "h 较大时", options: { bold:true, color:C.accent, fontSize:12, breakLine:true } },
      { text: "\u2202\u0394/\u2202\u03B8 = \u22122h sin\u03B8 \u2260 0", options: { fontSize:11, color:C.text, breakLine:true } },
      { text: "光程差随 \u03B8 变化，同一 \u03B8 的圆上 \u0394 相同", options: { fontSize:11, color:C.text, breakLine:true } },
      { text: "\u2192 同心圆环（等倾干涉）", options: { fontSize:11, color:C.accent, bold:true } },
    ], { x: MARGIN+0.2, y: 3.7, w: 4.0, h: 1.3, fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0 });

    // Right column
    s.addText([
      { text: "h \u2192 0 时", options: { bold:true, color:C.secondary, fontSize:12, breakLine:true } },
      { text: "\u2202\u0394/\u2202\u03B8 = 0", options: { fontSize:11, color:C.text, breakLine:true } },
      { text: "光程差对所有 \u03B8 均相同，发散角影响被消除", options: { fontSize:11, color:C.text, breakLine:true } },
      { text: "\u2192 平直条纹（等厚干涉）", options: { fontSize:11, color:C.secondary, bold:true } },
    ], { x: MARGIN+4.6, y: 3.7, w: 4.0, h: 1.3, fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0 });

    // Separator
    s.addShape("line", { x: 5.0, y: 3.7, w: 0, h: 1.2, line: { color: "DDDDDD", width: 1 } });

    addSlideNumber(s, 7, TOTAL);
  }

  // ==================== SLIDE 8: 核心原理(3) 形态判据 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "核心原理(3)：形态判据与物理图像");
    addFooterBar(s);

    s.addShape("rect", { x: MARGIN, y: 1.15, w: 9.0, h: 0.85, fill: { color: "FDE8E8" } });
    s.addText("核心结论", { x: MARGIN+0.2, y: 1.2, w: 1.5, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText("条纹是圆环还是平直，唯一取决于 h 是否为零 —— 与绝对光程差的大小毫无关系。实验中用红光调出平直条纹，实质上是将 M1 调到 h = 0，即两虚光源在空间上重合。",
      { x: MARGIN+0.2, y: 1.5, w: 8.6, h: 0.4, fontSize: 11, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    // Three columns
    const cols = [
      { title: "h = 0（参考状态）", color: C.secondary, lines: [
        "\u2022 两虚光源几何重合",
        "\u2022 \u2202\u0394/\u2202\u03B8 = 0",
        "\u2022 发散角影响消除",
        "\u2022 平直条纹（等厚）",
        "",
        "此时若两镜面间存在",
        "微小夹角（劈尖角 \u03B1），",
        "偏离中心 x 处附加",
        "光程差 = 2x\u03B1"
      ]},
      { title: "插入玻璃后（h \u2260 0）", color: C.accent, lines: [
        "\u2022 视深效应使虚像前移",
        "\u2022 \u2202\u0394/\u2202\u03B8 \u2260 0",
        "\u2022 光程差随 \u03B8 变化",
        "\u2022 同心圆环（等倾）",
        "",
        "固定光程偏移 2d(n\u22121)",
        "对所有 \u03B8 都相同，",
        "对 \u2202\u0394/\u2202\u03B8 毫无贡献",
        "因此不改变条纹形态"
      ]},
      { title: "补偿后（h \u2192 0）", color: "27AE60", lines: [
        "\u2022 M1 后退抵消视深位移",
        "\u2022 虚像重新几何重合",
        "\u2022 \u2202\u0394/\u2202\u03B8 = 0",
        "\u2022 平直条纹恢复",
        "",
        "注意：后退距离 \u0394x",
        "纯粹是让虚像退回",
        "h = 0，完全不涉及",
        "对固定光程偏移的补偿"
      ]},
    ];
    cols.forEach((col, i) => {
      const x0 = MARGIN + i * 3.1;
      s.addShape("rect", { x: x0, y: 2.15, w: 2.9, h: 3.0, fill: { color: C.cardBg }, shadow: makeShadow() });
      s.addText(col.title, { x: x0+0.15, y: 2.25, w: 2.6, h: 0.3, fontSize: 12, fontFace: FONT_B, color: col.color, bold: true, margin: 0 });
      s.addText(col.lines.join("\n"), { x: x0+0.15, y: 2.6, w: 2.6, h: 2.4, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
    });

    addSlideNumber(s, 8, TOTAL);
  }

  // ==================== SLIDE 9: 红光原理(1) 双重效应 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "红光原理(1)：插入玻璃的双重效应");
    addFooterBar(s);

    s.addShape("rect", { x: MARGIN, y: 1.15, w: 9.0, h: 0.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("在未插玻璃时 h = 0，虚像重合与光程相等两个条件自然同时满足。垂直插入厚度 d、折射率 n 的厚玻璃板后，玻璃对光同时产生两种效应：",
      { x: MARGIN+0.2, y: 1.2, w: 8.6, h: 0.5, fontSize: 11, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    // Table
    const header = [
      { text: "效应", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
      { text: "机制", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
      { text: "对 \u0394(\u03B8) 的影响", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
      { text: "对 \u2202\u0394/\u2202\u03B8 的影响", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:11, align:"center" } },
    ];
    const rows = [
      [
        { text: "减速效应", options: { fontSize:11, bold:true, color:C.text, align:"center" } },
        { text: "光在玻璃中速度变慢", options: { fontSize:11, color:C.text, align:"center" } },
        { text: "追加固定偏移 2d(n\u22121)", options: { fontSize:11, color:C.text, align:"center" } },
        { text: "为零（常数求导）", options: { fontSize:11, color:C.text, align:"center" } },
      ],
      [
        { text: "视深效应", options: { fontSize:11, bold:true, color:C.accent, align:"center" } },
        { text: "折射使 M1 虚像位置前移", options: { fontSize:11, color:C.text, align:"center" } },
        { text: "改变 h（h 不再为 0）", options: { fontSize:11, color:C.accent, align:"center" } },
        { text: "非零（h\u22600）", options: { fontSize:11, color:C.accent, bold:true, align:"center" } },
      ],
    ];
    s.addTable([header, ...rows], {
      x: MARGIN, y: 1.9, w: 9.0,
      colW: [1.5, 2.5, 2.5, 2.5],
      rowH: [0.4, 0.4, 0.4],
      border: { pt: 0.5, color: "DDDDDD" },
      margin: [3, 5, 3, 5],
    });

    // Key point
    s.addShape("rect", { x: MARGIN, y: 3.05, w: 9.0, h: 1.5, fill: { color: "FDE8E8" } });
    s.addText("关键洞察", { x: MARGIN+0.2, y: 3.1, w: 1.5, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText([
      { text: "玻璃引入的固定光程偏移 2d(n\u22121) 对所有 \u03B8 都相同，对 \u2202\u0394/\u2202\u03B8 毫无贡献，因此不改变条纹形态。", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "激光的相干长度长达数十厘米，这个固定偏移完全在其容忍范围内——即便绝对光程差已远非零，激光仍能保持清晰干涉。", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "而视深效应则完全不同：它直接破坏了 h = 0 的条件，使 \u2202\u0394/\u2202\u03B8 不再为零，条纹从平直退化为同心圆环。", options: { fontSize: 11, color: C.accent, bold: true } },
    ], { x: MARGIN+0.2, y: 3.4, w: 8.6, h: 1.1, fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0 });

    addSlideNumber(s, 9, TOTAL);
  }

  // ==================== SLIDE 10: 红光原理(2) 视深效应与补偿 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "红光原理(2)：视深效应与补偿推导");
    addFooterBar(s);

    // Left: apparent depth
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 2.3, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("视深效应", { x: MARGIN+0.2, y: 1.2, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText("在近垂直入射条件下透过厚玻璃看 M1，其虚像会向观察者方向（分光板 G₁ 方向）跃迁前移：",
      { x: MARGIN+0.2, y: 1.55, w: 4.0, h: 0.55, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
    addMathBox(s, [
      { text: "\u0394x' = d(1 \u2212 1/n)", options: { fontSize: 16, color: C.accent, bold: true, italic: true } },
    ], MARGIN+0.2, 2.15, 4.0, 0.45, C.codeBg);
    s.addText("这直接破坏了 h = 0 的条件——两虚光源在几何上重新分离，\u2202\u0394/\u2202\u03B8 不再为零，条纹从平直退化为同心圆环。",
      { x: MARGIN+0.2, y: 2.7, w: 4.0, h: 0.6, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    // Right: compensation
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.3, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("补偿逻辑", { x: 5.3, y: 1.2, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText("条纹形态只取决于 h，要恢复平直条纹只需让 h 重新归零。加玻璃使 M1 的虚像前移了 \u0394x'，因此在物理空间中把 M1 向后（远离分光板方向）移动同样的距离 \u0394x：",
      { x: 5.3, y: 1.55, w: 4.0, h: 0.7, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
    addMathBox(s, [
      { text: "\u0394x = \u0394x' = d(1 \u2212 1/n)", options: { fontSize: 14, color: C.text, bold: true, italic: true } },
    ], 5.3, 2.3, 4.0, 0.4, C.codeBg);
    s.addText("注意：此移动完全不涉及对固定光程偏移 2d(n\u22121) 的补偿——激光的相干长度远大于该偏移量。",
      { x: 5.3, y: 2.8, w: 4.0, h: 0.5, fontSize: 10, fontFace: FONT_B, color: C.muted, lineSpacingMultiple: 1.2, margin: 0 });

    // Bottom: formula derivation
    s.addShape("rect", { x: MARGIN, y: 3.6, w: 9.0, h: 1.5, fill: { color: "FDE8E8" } });
    s.addText("折射率公式推导", { x: MARGIN+0.2, y: 3.65, w: 2.5, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText("由 \u0394x = d(1 \u2212 1/n)，整理：",
      { x: MARGIN+0.2, y: 3.95, w: 2.5, h: 0.25, fontSize: 11, fontFace: FONT_B, color: C.text, margin: 0 });
    addMathBox(s, [
      { text: "\u0394x = d \u2212 d/n", options: { fontSize: 12, color: C.text } },
      { text: "d/n = d \u2212 \u0394x", options: { fontSize: 12, color: C.text } },
      { text: "n = d / (d \u2212 \u0394x)", options: { fontSize: 16, color: C.accent, bold: true, italic: true } },
    ], MARGIN+3.0, 3.65, 3.5, 1.35, C.codeBg);
    s.addText("这就是红光实验中折射率的计算公式。",
      { x: MARGIN+6.8, y: 4.2, w: 2.0, h: 0.6, fontSize: 11, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });

    addSlideNumber(s, 10, TOTAL);
  }

  // ==================== SLIDE 11: 白光原理(1) 零光程差约束 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "白光原理(1)：零光程差约束与光程跳变");
    addFooterBar(s);

    // Left: coherence constraint
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText('白光的「光程极零」强制约', { x: MARGIN+0.2, y: 1.2, w: 3.8, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "白光属于宽频带复色光，各色波长干涉条纹相互错位中和，相干长度仅微米量级。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "核心要求：两臂光程的绝对差值趋近于零", options: { fontSize: 11, color: C.gold, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "只有当各色条纹在视场中心同相位叠加，即 OPD \u2248 0 时，才能在中心区域观察到少量彩色干涉条纹。", options: { fontSize: 10, color: C.text } },
    ], { x: MARGIN+0.2, y: 1.55, w: 4.0, h: 1.5, fontFace: FONT_B, lineSpacingMultiple: 1.2, valign: "top", margin: 0 });

    // Right: OPD jump
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("插入玻片引发的单向光程延长", { x: 5.3, y: 1.2, w: 3.8, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText("光束两次穿过厚度 d 的玻璃片（射入并经动镜反射折返），该臂的总光程跳变为：",
      { x: 5.3, y: 1.55, w: 4.0, h: 0.5, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    addMathBox(s, [
      { text: "\u0394(OPD) = 2d(n\u22121)", options: { fontSize: 16, color: C.gold, bold: true, italic: true } },
    ], 5.3, 2.1, 4.0, 0.45, C.codeBg);

    s.addText([
      { text: "推导：", options: { bold:true, fontSize:10, color:C.text, breakLine:true } },
      { text: "\u2022 单程光程增量 = nd \u2212 d = d(n\u22121)", options: { fontSize:10, color:C.text, breakLine:true } },
      { text: "\u2022 往返总增量 = 2d(n\u22121)", options: { fontSize:10, color:C.text, breakLine:true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "这一跳变远超白光的相干长度，使原有的彩色条纹迅速消失。", options: { fontSize: 10, color: C.accent, bold: true } },
    ], { x: 5.3, y: 2.65, w: 4.0, h: 1.0, fontFace: FONT_B, lineSpacingMultiple: 1.15, margin: 0 });

    // Bottom: residual h note
    s.addShape("rect", { x: MARGIN, y: 3.9, w: 9.0, h: 1.2, fill: { color: "FFF8E1" } });
    s.addText("重要细节", { x: MARGIN+0.2, y: 3.95, w: 1.5, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    s.addText("白光补偿仅恢复了 OPD \u2248 0，并未同时使两虚光源的间距 h 归零。残留 h = \u0394x \u2212 d(1\u22121/n) = d(n\u22122+1/n)，对 d=0.3mm、n\u22481.57 约为 0.06mm。此残留间距使 \u0394(\u03B8) 随 \u03B8 变化极为缓慢——视场内光程差变化不足半个波长——因此条纹曲率极小，在劈尖角主导下呈现近似平行的彩色外观。",
      { x: MARGIN+0.2, y: 4.25, w: 8.6, h: 0.8, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    addSlideNumber(s, 11, TOTAL);
  }

  // ==================== SLIDE 12: 白光原理(2) 补偿推导 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "白光原理(2)：补偿动作与折射率推导");
    addFooterBar(s);

    // Left: compensation
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.4, h: 2.8, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("补偿动作", { x: MARGIN+0.2, y: 1.2, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    s.addText("要使 OPD 重新归零以重现彩色条纹，须在物理上缩短该臂的光程。将 M1 向分光板方向推进（即靠近），在空气段消去冗余距离。",
      { x: MARGIN+0.2, y: 1.55, w: 4.0, h: 0.65, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });
    s.addText("设移动距离为 \u0394x，对应的物理光程减少量为 2\u0394x，由光程补偿守恒条件：",
      { x: MARGIN+0.2, y: 2.25, w: 4.0, h: 0.4, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0 });
    addMathBox(s, [
      { text: "2\u0394x = 2d(n\u22121)", options: { fontSize: 16, color: C.gold, bold: true, italic: true } },
    ], MARGIN+0.2, 2.7, 4.0, 0.45, C.codeBg);
    s.addText("化简即得折射率公式：",
      { x: MARGIN+0.2, y: 3.25, w: 4.0, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.text, margin: 0 });
    addMathBox(s, [
      { text: "n = 1 + \u0394x / d", options: { fontSize: 18, color: C.gold, bold: true, italic: true } },
    ], MARGIN+0.2, 3.55, 4.0, 0.45, "FFF8E1");

    // Right: verification
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.8, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("物理验证", { x: 5.3, y: 1.2, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "该结论与低相干光源「零光程差」条件一致，也与白光彩条仅在 OPD \u2248 0 附近出现的事实相符。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "关键区别：", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u2022 红光：只补偿虚像位移（几何条件 h=0）", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "\u2022 白光：只补偿光程跳变（OPD=0）", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "两种光源对干涉条件的「选择性」不同，导致了完全相反的补偿方向。", options: { fontSize: 10, color: C.gold, bold: true } },
    ], { x: 5.3, y: 1.55, w: 4.0, h: 1.9, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    // Bottom: direction comparison
    s.addShape("rect", { x: MARGIN, y: 4.1, w: 9.0, h: 1.0, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("补偿方向对比", { x: MARGIN+0.2, y: 4.15, w: 2.0, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText("红光：M1 远离分光板 \u2190 | 白光：M1 靠近分光板 \u2192",
      { x: MARGIN+0.2, y: 4.5, w: 4.5, h: 0.35, fontSize: 13, fontFace: FONT_B, color: C.text, valign: "middle", margin: 0 });
    s.addText("两种操作方向完全相反的现象看似矛盾，实则是不同光源对干涉条件的「选择性」不同所致。",
      { x: MARGIN+5.0, y: 4.5, w: 3.8, h: 0.35, fontSize: 10, fontFace: FONT_B, color: C.muted, valign: "middle", margin: 0 });

    addSlideNumber(s, 12, TOTAL);
  }

  // ==================== SLIDE 13: 原理对比与厚度选择 ====================
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
        { text: "视深效应 \u2192 虚像位移", options: { fontSize:10, color:C.text, align:"center" } },
        { text: "光程跳变 \u2192 OPD补偿", options: { fontSize:10, color:C.text, align:"center" } },
      ],
      [
        { text: "M1 移动方向", options: { bold:true, fontSize:10, color:C.text } },
        { text: "远离分光板 \u2190", options: { fontSize:10, color:C.accent, bold:true, align:"center" } },
        { text: "\u2192 靠近分光板", options: { fontSize:10, color:C.gold, bold:true, align:"center" } },
      ],
      [
        { text: "玻璃厚度", options: { bold:true, fontSize:10, color:C.text } },
        { text: "7 mm（厚）", options: { fontSize:10, color:C.text, align:"center" } },
        { text: "0.3 mm（薄）", options: { fontSize:10, color:C.text, align:"center" } },
      ],
      [
        { text: "折射率公式", options: { bold:true, fontSize:10, color:C.text } },
        { text: "n = d/(d\u2212\u0394x)", options: { fontSize:10, color:C.accent, italic:true, align:"center" } },
        { text: "n = 1+\u0394x/d", options: { fontSize:10, color:C.gold, italic:true, align:"center" } },
      ],
      [
        { text: "关键约束", options: { bold:true, fontSize:10, color:C.text } },
        { text: "h = 0（几何条件）", options: { fontSize:10, color:C.text, align:"center" } },
        { text: "OPD \u2248 0（相干条件）", options: { fontSize:10, color:C.text, align:"center" } },
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
      { text: "虚像位移 \u0394x' = 7\u00D7(1\u22121/1.5) \u2248 2.33 mm，千分尺上 2mm 量级位移易于精确读取。玻璃越厚信号越大，且激光不在乎相伴的固定光程偏移——", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "厚玻璃在激光实验中纯属优势。", options: { fontSize: 10, color: C.accent, bold: true } },
    ], { x: MARGIN+0.2, y: 3.7, w: 4.2, h: 1.2, fontFace: FONT_B, lineSpacingMultiple: 1.15, margin: 0 });

    s.addText([
      { text: "白光实验（d = 0.3mm）：", options: { bold: true, fontSize: 11, color: C.gold, breakLine: true } },
      { text: "OPD 跳变 = 2\u00D70.3\u00D70.5 = 0.3 mm，补偿位移 \u2248 0.15 mm，仍在可追踪范围。若用 7mm 厚玻璃，OPD 跳变将达 7mm 量级——远超白光微米级相干长度，实验上几乎不可能重新捕捉彩色条纹。", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "白光实验的玻璃必须薄到让 OPD 跳变不超出可操作范围。", options: { fontSize: 10, color: C.gold, bold: true } },
    ], { x: MARGIN+4.7, y: 3.7, w: 4.2, h: 1.2, fontFace: FONT_B, lineSpacingMultiple: 1.15, margin: 0 });

    addSlideNumber(s, 13, TOTAL);
  }

  // ==================== SLIDE 14: 数据处理 红光 ====================
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
      { text: "\u0394xᵢ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
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
    s.addText("平均值 \u0394x\u0304 = 2.79678 mm", { x: MARGIN+0.2, y: 2.55, w: 3.0, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });

    // Right: calculation
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.2, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("计算过程", { x: 5.3, y: 1.2, w: 2.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "样本标准差：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "s\u0394x = 0.00716 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "A类标准不确定度：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(\u0394x\u0304) = s\u0394x / \u221a3 = 0.00413 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "由 n = d/(d\u2212\u0394x) 得：", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "n\u7EA2 = 7.0 / (7.0\u22122.79678) \u2248 1.6655", options: { fontSize: 11, color: C.accent, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "不确定度传播：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(n) = |\u2202n/\u2202\u0394x|\u00B7u(\u0394x\u0304) = d/(d\u2212\u0394x\u0304)\u00B2 \u00B7 u(\u0394x\u0304)", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "u(n\u7EA2) \u2248 0.0016", options: { fontSize: 10, color: C.accent, bold: true } },
    ], { x: 5.3, y: 1.55, w: 4.0, h: 1.7, fontFace: FONT_B, lineSpacingMultiple: 1.1, valign: "top", margin: 0 });

    // Bottom result
    s.addShape("rect", { x: MARGIN, y: 3.55, w: 9.0, h: 1.15, fill: { color: "FDE8E8" } });
    s.addText("红光实验结果", { x: MARGIN+0.2, y: 3.6, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    addMathBox(s, [
      { text: "n\u7EA2 = 1.6655 \u00B1 0.0016", options: { fontSize: 24, color: C.accent, bold: true, italic: true } },
    ], MARGIN+3.0, 3.6, 4.0, 1.0, "FDE8E8");

    addSlideNumber(s, 14, TOTAL);
  }

  // ==================== SLIDE 15: 数据处理 白光 ====================
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
      { text: "\u0394xᵢ (mm)", options: { fill:{color:C.primary}, color:C.white, bold:true, fontSize:9, align:"center" } },
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
    s.addText("平均值 \u0394x\u0304 = 0.17215 mm", { x: MARGIN+0.2, y: 2.55, w: 3.0, h: 0.25, fontSize: 10, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });

    // Right: calculation
    s.addShape("rect", { x: 5.1, y: 1.15, w: 4.4, h: 2.2, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("计算过程", { x: 5.3, y: 1.2, w: 2.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText([
      { text: "样本标准差：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "s\u0394x = 0.00168 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "A类标准不确定度：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(\u0394x\u0304) = s\u0394x / \u221a3 = 0.00097 mm", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "由 n = 1 + \u0394x/d 得：", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "n\u767D = 1 + 0.17215/0.3 \u2248 1.5738", options: { fontSize: 11, color: C.gold, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "不确定度传播：", options: { fontSize: 10, color: C.text, breakLine: false } },
      { text: "u(n) = u(\u0394x\u0304) / d = 0.00097 / 0.3 \u2248 0.0032", options: { fontSize: 10, color: C.text, breakLine: true } },
      { text: "u(n\u767D) \u2248 0.0032", options: { fontSize: 10, color: C.gold, bold: true } },
    ], { x: 5.3, y: 1.55, w: 4.0, h: 1.7, fontFace: FONT_B, lineSpacingMultiple: 1.1, valign: "top", margin: 0 });

    // Bottom result
    s.addShape("rect", { x: MARGIN, y: 3.55, w: 9.0, h: 1.15, fill: { color: "FFF8E1" } });
    s.addText("白光实验结果", { x: MARGIN+0.2, y: 3.6, w: 2.0, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    addMathBox(s, [
      { text: "n\u767D = 1.5738 \u00B1 0.0032", options: { fontSize: 24, color: C.gold, bold: true, italic: true } },
    ], MARGIN+3.0, 3.6, 4.0, 1.0, "FFF8E1");

    addSlideNumber(s, 15, TOTAL);
  }

  // ==================== SLIDE 16: 误差分析 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "误差分析");
    addFooterBar(s);

    // Red errors
    s.addShape("rect", { x: MARGIN, y: 1.15, w: 4.3, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("红光实验误差源", { x: MARGIN+0.2, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.accent, bold: true, margin: 0 });
    s.addText([
      { text: "\u2022 条纹判断误差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  「平直程度」依赖操作者主观判断，容易把「接近平直」当作「完全平直」。", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "\u2022 千分尺读数误差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  M1 位移需反复微调，若读数时存在回差或视差，会直接影响 \u0394x 平均值。", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "\u2022 玻璃板放置偏差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  若厚玻璃板未严格保持垂直入射，虚像位移不再只由厚度决定，引入系统偏差。", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "\u2022 厚度测量误差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  玻璃厚度 d 偏差会按 n = d/(d\u2212\u0394x) 被放大或缩小。", options: { fontSize: 10, color: C.muted } },
    ], { x: MARGIN+0.2, y: 1.55, w: 3.9, h: 2.1, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    // White errors
    s.addShape("rect", { x: 5.0, y: 1.15, w: 4.3, h: 2.6, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("白光实验误差源", { x: 5.2, y: 1.2, w: 3.0, h: 0.3, fontSize: 13, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
    s.addText([
      { text: "\u2022 相干条件更苛刻：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  白光条纹只在极窄的 OPD\u22480 区域出现，补偿位置更易受主观判断影响。", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "\u2022 光源与对准不稳定：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  稍有偏差就可能导致彩色条纹不清晰，影响 M1 最终停点。", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "\u2022 薄片厚度偏差：", options: { bold: true, fontSize: 10, color: C.text, breakLine: true } },
      { text: "  d = 0.3mm 较小，厚度标称值与实际值的微小偏差会按 n=1+\u0394x/d 被显著放大。", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "  例如 d 偏差 0.02mm 即可导致 n 变化约 0.1。", options: { fontSize: 10, color: C.muted } },
    ], { x: 5.2, y: 1.55, w: 3.9, h: 2.1, fontFace: FONT_B, lineSpacingMultiple: 1.15, valign: "top", margin: 0 });

    // Bottom: overall judgment
    s.addShape("rect", { x: MARGIN, y: 3.9, w: 9.0, h: 1.05, fill: { color: C.cardBg }, shadow: makeShadow() });
    s.addText("总体判断", { x: MARGIN+0.2, y: 3.95, w: 1.5, h: 0.25, fontSize: 13, fontFace: FONT_B, color: C.text, bold: true, margin: 0 });
    s.addText("n\u7EA2 = 1.6655 \u00B1 0.0016，n\u767D = 1.5738 \u00B1 0.0032，两者均落在常见玻璃折射率范围内。微小差异可能来源于两块玻璃材质不同，也可能来自两种原理各自的系统偏差。两个实验从不同物理机制出发，均得到了量级合理的结果，相互印证了两种补偿逻辑（远离 vs 靠近）的正确性。",
      { x: MARGIN+0.2, y: 4.25, w: 8.6, h: 0.65, fontSize: 10, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0 });

    addSlideNumber(s, 16, TOTAL);
  }

  // ==================== SLIDE 17: 结论 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
    s.addText("结论", { x: MARGIN, y: 0.25, w: 8.6, h: 0.55, fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0 });

    const conclusions = [
      {
        num: "01",
        title: "红光实验结论",
        text: "插入厚玻璃板后需将 M1 远离分光板，以补偿视深效应并恢复等厚干涉的几何条件（h=0）。折射率公式 n = d/(d\u2212\u0394x)，计算得 n\u7EA2 = 1.6655 \u00B1 0.0016。",
      },
      {
        num: "02",
        title: "白光实验结论",
        text: "插入薄玻璃片后需将 M1 向分光板方向移动，以补偿增加的光程差并回到 OPD \u2248 0。折射率公式 n = 1 + \u0394x/d，计算得 n\u767D = 1.5738 \u00B1 0.0032。",
      },
      {
        num: "03",
        title: "方向差异的物理根源",
        text: "激光相干长度极长，只响应虚像间距 h（几何条件 \u2202\u0394/\u2202\u03B8）；白光相干长度极短，只响应绝对光程差 OPD \u2248 0。两种光源对干涉条件的选择性不同，导致了完全相反的补偿方向。",
      },
    ];

    conclusions.forEach((c, i) => {
      const y0 = 1.1 + i * 1.3;
      s.addShape("rect", { x: MARGIN, y: y0+0.05, w: 0.55, h: 0.55, fill: { color: C.primary } });
      s.addText(c.num, { x: MARGIN, y: y0+0.05, w: 0.55, h: 0.55, fontSize: 18, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(c.title, { x: MARGIN+0.75, y: y0, w: 8.0, h: 0.35, fontSize: 15, fontFace: FONT_B, color: C.gold, bold: true, margin: 0 });
      s.addText(c.text, { x: MARGIN+0.75, y: y0+0.38, w: 8.0, h: 0.55, fontSize: 11, fontFace: FONT_B, color: "D0D8E0", lineSpacingMultiple: 1.3, margin: 0 });
    });

    addSlideNumber(s, 17, TOTAL, true);
  }

  // ==================== SLIDE 18: 致谢 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
    s.addText("谢谢！", { x: 1.0, y: 1.4, w: 8.0, h: 1.2, fontSize: 52, fontFace: FONT_H, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addShape("rect", { x: 3.2, y: 2.8, w: 3.6, h: 0.03, fill: { color: C.gold } });
    s.addText("感谢聆听 · 欢迎提问", { x: 1.0, y: 3.1, w: 8.0, h: 0.6, fontSize: 18, fontFace: FONT_B, color: "AABBCC", align: "center", margin: 0 });
    s.addText("迈克尔逊干涉仪测折射率实验探究", { x: 1.0, y: 4.1, w: 8.0, h: 0.4, fontSize: 13, fontFace: FONT_B, color: "667788", align: "center", margin: 0 });
    addSlideNumber(s, 18, TOTAL, true);
  }

  // ─── Write file ───
  const outPath = "d:/project/迈克尔逊干涉仪实验探究/答辩PPT_V2.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("PPT saved to: " + outPath);
}

main().catch(err => { console.error(err); process.exit(1); });
