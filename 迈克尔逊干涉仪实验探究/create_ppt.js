const pptxgen = require("pptxgenjs");

// ─── Color Palette ───
const C = {
  darkBg:    "0D1B2A",
  primary:   "1B3A5C",
  secondary: "2980B9",
  accent:    "CC2233",  // laser red
  gold:      "D4A017",
  lightBg:   "F0F4F8",
  cardBg:    "FFFFFF",
  text:      "1E293B",
  muted:     "64748B",
  white:     "FFFFFF",
};

const FONT_H = "Georgia";
const FONT_B = "Calibri";
const MARGIN = 0.7;

// ─── Reusable helpers ───
const makeShadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.10 });

function addSlideNumber(slide, num, total, dark = false) {
  slide.addText(`${num} / ${total}`, {
    x: 9.0, y: 5.25, w: 0.8, h: 0.3,
    fontSize: 9, fontFace: FONT_B, color: dark ? "8899AA" : C.muted,
    align: "right", margin: 0,
  });
}

function addFooterBar(slide, dark = false) {
  slide.addShape("rect", {
    x: 0, y: 5.35, w: 10, h: 0.275,
    fill: { color: dark ? C.primary : C.secondary },
  });
}

function addSectionTitle(slide, titleText, dark = false) {
  slide.addText(titleText, {
    x: MARGIN, y: 0.3, w: 8.6, h: 0.55,
    fontSize: 30, fontFace: FONT_H, color: dark ? C.white : C.text,
    bold: true, margin: 0,
  });
  slide.addShape("rect", {
    x: MARGIN, y: 0.92, w: 1.5, h: 0.04,
    fill: { color: dark ? C.gold : C.accent },
  });
}

// Native icon: colored rounded square with unicode symbol
function addIcon(slide, symbol, x, y, w, h, bgColor, textColor) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: bgColor },
    rectRadius: 0.05,
  });
  slide.addText(symbol, {
    x, y, w, h,
    fontSize: w * 28,  // scale font with box size
    fontFace: "Segoe UI Symbol",
    color: textColor || C.white,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Main ───
async function main() {
  const TOTAL = 14;
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
      x: 0.8, y: 1.2, w: 8.4, h: 2.0,
      fontSize: 42, fontFace: FONT_H, color: C.white, bold: true,
      align: "left", lineSpacingMultiple: 1.2, margin: 0,
    });
    s.addText("答辩报告", {
      x: 0.8, y: 3.4, w: 8.4, h: 0.6,
      fontSize: 22, fontFace: FONT_B, color: C.gold,
      align: "left", margin: 0,
    });
    s.addShape("rect", { x: 0.8, y: 3.2, w: 2.0, h: 0.03, fill: { color: C.accent } });
    s.addText("物理实验 · 迈克尔逊干涉仪 · 折射率测量", {
      x: 0.8, y: 4.7, w: 8.4, h: 0.35,
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
      { num: "01", title: "实验背景与目的", desc: "研究动机与实验目标" },
      { num: "02", title: "实验装置与光路", desc: "迈克尔逊干涉仪结构与配置" },
      { num: "03", title: "实验操作与现象", desc: "红光 & 白光两种实验流程" },
      { num: "04", title: "核心原理探究", desc: "虚像位移 vs 零光程差补偿" },
      { num: "05", title: "数据处理与结果", desc: "测量数据、折射率计算与不确定度" },
      { num: "06", title: "误差分析与结论", desc: "误差来源讨论与最终结论" },
    ];
    items.forEach((item, i) => {
      const col = i < 3 ? 0 : 1;
      const row = i % 3;
      const x0 = MARGIN + col * 4.5;
      const y0 = 1.4 + row * 1.2;
      s.addText(item.num, {
        x: x0, y: y0, w: 0.65, h: 0.9,
        fontSize: 28, fontFace: FONT_H, color: C.accent, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(item.title, {
        x: x0 + 0.75, y: y0 + 0.05, w: 3.5, h: 0.45,
        fontSize: 16, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
      });
      s.addText(item.desc, {
        x: x0 + 0.75, y: y0 + 0.48, w: 3.5, h: 0.35,
        fontSize: 11, fontFace: FONT_B, color: C.muted, margin: 0,
      });
    });
    addSlideNumber(s, 2, TOTAL);
  }

  // ==================== SLIDE 3: 实验背景与目的 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "实验背景与目的");
    addFooterBar(s);

    // left card: background
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 4.0, h: 3.6,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    addIcon(s, "🔍", MARGIN + 0.3, 1.55, 0.4, 0.4, C.secondary, C.white);
    s.addText("研究背景", {
      x: MARGIN + 0.85, y: 1.55, w: 2.8, h: 0.4,
      fontSize: 17, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    s.addText([
      { text: "迈克尔逊干涉仪是精密光学测量的经典工具，利用双光束干涉原理实现高精度位移与折射率测量。", options: { fontSize: 12, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "核心问题：在干涉仪一臂插入玻璃片后，条纹形态发生变化。不同光源（激光/白光）下，恢复条纹时动镜M1的移动方向居然相反——这一现象背后的物理本质是什么？", options: { fontSize: 12, color: C.text } },
    ], {
      x: MARGIN + 0.3, y: 2.15, w: 3.4, h: 2.6,
      fontFace: FONT_B, lineSpacingMultiple: 1.4, valign: "top", margin: 0,
    });

    // right card: objectives
    s.addShape("rect", {
      x: 5.3, y: 1.35, w: 4.0, h: 3.6,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    addIcon(s, "🎯", 5.6, 1.55, 0.4, 0.4, C.accent, C.white);
    s.addText("实验目的", {
      x: 6.15, y: 1.55, w: 2.8, h: 0.4,
      fontSize: 17, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    s.addText([
      { text: "1.", options: { bold: true, color: C.accent, fontSize: 13, breakLine: false } },
      { text: " 使用迈克尔逊干涉仪测量两种不同厚度与类型玻璃片的折射率。", options: { fontSize: 13, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 8, breakLine: true } },
      { text: "2.", options: { bold: true, color: C.accent, fontSize: 13, breakLine: false } },
      { text: " 探究不同光源下插入玻璃片后，动镜M1的移动方向差异及其物理原因，并给出对应折射率计算公式。", options: { fontSize: 13, color: C.text } },
    ], {
      x: 5.6, y: 2.15, w: 3.4, h: 2.6,
      fontFace: FONT_B, lineSpacingMultiple: 1.4, valign: "top", margin: 0,
    });
    addSlideNumber(s, 3, TOTAL);
  }

  // ==================== SLIDE 4: 实验装置 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "实验装置与光路");
    addFooterBar(s);

    // Left: equipment list card
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 3.7, h: 3.6,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    addIcon(s, "⚙", MARGIN + 0.3, 1.55, 0.4, 0.4, C.primary, C.white);
    s.addText("实验仪器", {
      x: MARGIN + 0.85, y: 1.55, w: 2.5, h: 0.4,
      fontSize: 17, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    const equipment = [
      "迈克尔逊干涉仪",
      "氦氖激光器（红光，λ≈632.8nm）",
      "短焦距透镜 + 扩束镜",
      "毛玻璃片",
      "白光光源",
      "观察屏",
      "厚玻璃板（7mm）& 薄玻璃片（0.3mm）",
    ];
    const eqItems = equipment.map((eq, i) => ({
      text: "• " + eq,
      options: { fontSize: 12, color: C.text, breakLine: i < equipment.length - 1 },
    }));
    s.addText(eqItems, {
      x: MARGIN + 0.3, y: 2.15, w: 3.1, h: 2.6,
      fontFace: FONT_B, lineSpacingMultiple: 1.4, valign: "top", margin: 0,
    });

    // Right: optical path schematic
    s.addShape("rect", {
      x: 5.05, y: 1.35, w: 4.25, h: 3.6,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("光路示意", {
      x: 5.35, y: 1.55, w: 3.6, h: 0.4,
      fontSize: 17, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });

    // Source S
    s.addShape("rect", { x: 5.3, y: 2.0, w: 0.7, h: 0.35, fill: { color: C.secondary }, rectRadius: 0.05 });
    s.addText("光源 S", { x: 5.3, y: 2.0, w: 0.7, h: 0.35, fontSize: 9, fontFace: FONT_B, color: C.white, align: "center", valign: "middle", margin: 0 });
    // G1
    s.addShape("rect", { x: 6.1, y: 2.0, w: 0.55, h: 0.35, fill: { color: C.primary } });
    s.addText("G\u2081", { x: 6.1, y: 2.0, w: 0.55, h: 0.35, fontSize: 9, fontFace: FONT_B, color: C.white, align: "center", valign: "middle", margin: 0 });
    // G2
    s.addShape("rect", { x: 6.8, y: 2.0, w: 0.55, h: 0.35, fill: { color: C.primary } });
    s.addText("G\u2082", { x: 6.8, y: 2.0, w: 0.55, h: 0.35, fontSize: 9, fontFace: FONT_B, color: C.white, align: "center", valign: "middle", margin: 0 });
    // M1 (top, movable)
    s.addShape("rect", { x: 6.65, y: 2.7, w: 0.7, h: 0.55, fill: { color: C.accent }, rectRadius: 0.05 });
    s.addText("M\u2081\n(动镜)", { x: 6.65, y: 2.7, w: 0.7, h: 0.55, fontSize: 8, fontFace: FONT_B, color: C.white, align: "center", valign: "middle", margin: 0 });
    // M2 (right, fixed)
    s.addShape("rect", { x: 7.8, y: 2.7, w: 0.7, h: 0.55, fill: { color: C.secondary }, rectRadius: 0.05 });
    s.addText("M\u2082\n(定镜)", { x: 7.8, y: 2.7, w: 0.7, h: 0.55, fontSize: 8, fontFace: FONT_B, color: C.white, align: "center", valign: "middle", margin: 0 });
    // Glass plate between G2 and M2
    s.addShape("rect", { x: 7.5, y: 2.85, w: 0.2, h: 0.25, fill: { color: C.gold } });
    s.addText("玻璃片", { x: 7.5, y: 3.15, w: 0.8, h: 0.2, fontSize: 8, fontFace: FONT_B, color: C.gold, margin: 0 });

    // Screen
    s.addShape("rect", { x: 5.3, y: 3.3, w: 0.08, h: 0.7, fill: { color: C.muted } });
    s.addText("观察屏", { x: 5.3, y: 4.05, w: 0.7, h: 0.2, fontSize: 8, fontFace: FONT_B, color: C.muted, margin: 0 });

    // Light ray arrows (simplified)
    s.addShape("line", { x: 6.0, y: 2.18, w: 0.65, h: 0.3, line: { color: C.accent, width: 1.5 } });
    s.addShape("line", { x: 6.0, y: 2.18, w: 0.1, h: -0.8, line: { color: C.accent, width: 1.5 } });

    s.addText([
      { text: "核心测量：一臂插入玻璃片 → 条纹变形 → 移动M1恢复条纹 → 记录Δx → 计算n", options: { fontSize: 11, color: C.text, bold: true } },
    ], {
      x: 5.35, y: 4.4, w: 3.6, h: 0.4,
      fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0,
    });

    addSlideNumber(s, 4, TOTAL);
  }

  // ==================== SLIDE 5: 红光实验 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "实验操作与现象 —— 红色激光");
    addFooterBar(s);

    const steps = [
      { title: "准直扩束", desc: "He-Ne激光 + 扩束镜\n形成点光源" },
      { title: "调出条纹", desc: "微调M1/M2角度\n出现同心圆环" },
      { title: "参考状态", desc: "调至平行直条纹\n记录d\u2080" },
      { title: "插入玻璃", desc: "插入7mm厚玻璃板\n条纹变回圆环" },
      { title: "补偿调整", desc: "M1远离分光板\n恢复平行条纹 \u2192 d\u2081" },
    ];
    steps.forEach((st, i) => {
      const sx = MARGIN + i * 1.75;
      s.addShape("ellipse", {
        x: sx + 0.4, y: 1.35, w: 0.5, h: 0.5,
        fill: { color: i === 0 ? C.accent : (i === 4 ? C.accent : C.secondary) },
      });
      s.addText(String(i + 1), {
        x: sx + 0.4, y: 1.35, w: 0.5, h: 0.5,
        fontSize: 18, fontFace: FONT_H, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      if (i < 4) {
        s.addText("\u2192", {
          x: sx + 0.95, y: 1.35, w: 0.75, h: 0.5,
          fontSize: 18, fontFace: FONT_B, color: C.muted,
          align: "center", valign: "middle", margin: 0,
        });
      }
      s.addText(st.title, {
        x: sx, y: 2.0, w: 1.55, h: 0.35,
        fontSize: 12, fontFace: FONT_B, color: C.text, bold: true,
        align: "center", margin: 0,
      });
      s.addText(st.desc, {
        x: sx, y: 2.35, w: 1.55, h: 0.6,
        fontSize: 10, fontFace: FONT_B, color: C.muted,
        align: "center", lineSpacingMultiple: 1.3, margin: 0,
      });
    });

    // observation callout
    s.addShape("rect", {
      x: MARGIN, y: 3.2, w: 8.6, h: 1.2,
      fill: { color: "FDE8E8" },
    });
    s.addText("观测现象", {
      x: MARGIN + 0.3, y: 3.3, w: 2.0, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0,
    });
    // Arrow icon -> use unicode
    s.addText("\u279C", {
      x: MARGIN + 7.8, y: 3.55, w: 0.5, h: 0.4,
      fontSize: 24, fontFace: "Segoe UI Symbol", color: C.accent,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("插入玻璃板后，条纹重新变成同心圆环；将 M1 向远离分光板方向移动，可重新找到清晰平行条纹。", {
      x: MARGIN + 0.3, y: 3.7, w: 7.5, h: 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.text, margin: 0,
    });

    // Formula hint
    s.addShape("rect", {
      x: MARGIN, y: 4.55, w: 8.6, h: 0.55,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText([
      { text: "折射率公式：", options: { fontSize: 13, color: C.text, bold: true } },
      { text: "n = d / (d \u2212 \u0394x)", options: { fontSize: 14, color: C.accent, bold: true, italic: true } },
      { text: "    其中 \u0394x 为 M1 远离分光板的位移量", options: { fontSize: 11, color: C.muted } },
    ], {
      x: MARGIN + 0.3, y: 4.55, w: 8.0, h: 0.55,
      fontFace: FONT_B, valign: "middle", margin: 0,
    });
    addSlideNumber(s, 5, TOTAL);
  }

  // ==================== SLIDE 6: 白光实验 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "实验操作与现象 —— 白色光源");
    addFooterBar(s);

    const steps = [
      { title: "切换光源", desc: "白光 + 毛玻璃\n形成扩展光源" },
      { title: "寻迹条纹", desc: "缓慢移动M1\n找到彩色条纹" },
      { title: "参考状态", desc: "条纹最清晰时\n记录d\u2080" },
      { title: "插入玻璃", desc: "插入0.3mm薄片\n彩条消失" },
      { title: "补偿调整", desc: "M1靠近分光板\n恢复彩条 \u2192 d\u2081" },
    ];
    steps.forEach((st, i) => {
      const sx = MARGIN + i * 1.75;
      s.addShape("ellipse", {
        x: sx + 0.4, y: 1.35, w: 0.5, h: 0.5,
        fill: { color: i === 0 ? C.gold : (i === 4 ? C.gold : C.secondary) },
      });
      s.addText(String(i + 1), {
        x: sx + 0.4, y: 1.35, w: 0.5, h: 0.5,
        fontSize: 18, fontFace: FONT_H, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      if (i < 4) {
        s.addText("\u2192", {
          x: sx + 0.95, y: 1.35, w: 0.75, h: 0.5,
          fontSize: 18, fontFace: FONT_B, color: C.muted,
          align: "center", valign: "middle", margin: 0,
        });
      }
      s.addText(st.title, {
        x: sx, y: 2.0, w: 1.55, h: 0.35,
        fontSize: 12, fontFace: FONT_B, color: C.text, bold: true, align: "center", margin: 0,
      });
      s.addText(st.desc, {
        x: sx, y: 2.35, w: 1.55, h: 0.6,
        fontSize: 10, fontFace: FONT_B, color: C.muted, align: "center",
        lineSpacingMultiple: 1.3, margin: 0,
      });
    });

    // observation callout
    s.addShape("rect", {
      x: MARGIN, y: 3.2, w: 8.6, h: 1.2,
      fill: { color: "FFF8E1" },
    });
    s.addText("观测现象", {
      x: MARGIN + 0.3, y: 3.3, w: 2.0, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.gold, bold: true, margin: 0,
    });
    s.addText("\u279C", {
      x: MARGIN + 7.8, y: 3.55, w: 0.5, h: 0.4,
      fontSize: 24, fontFace: "Segoe UI Symbol", color: C.gold,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("插入玻璃片后彩色条纹迅速消失；通过将 M1 向靠近分光板方向移动，可重新找到清晰彩色平行条纹。", {
      x: MARGIN + 0.3, y: 3.7, w: 7.5, h: 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.text, margin: 0,
    });

    s.addShape("rect", {
      x: MARGIN, y: 4.55, w: 8.6, h: 0.55,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText([
      { text: "折射率公式：", options: { fontSize: 13, color: C.text, bold: true } },
      { text: "n = 1 + \u0394x / d", options: { fontSize: 14, color: C.gold, bold: true, italic: true } },
      { text: "    其中 \u0394x 为 M1 靠近分光板的位移量", options: { fontSize: 11, color: C.muted } },
    ], {
      x: MARGIN + 0.3, y: 4.55, w: 8.0, h: 0.55,
      fontFace: FONT_B, valign: "middle", margin: 0,
    });
    addSlideNumber(s, 6, TOTAL);
  }

  // ==================== SLIDE 7: 核心原理 —— 条纹形态判据 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "核心原理：条纹形态判据");
    addFooterBar(s);

    // Key concept box
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 8.6, h: 1.1,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    addIcon(s, "\uD83D\uDCA1", MARGIN + 0.3, 1.55, 0.4, 0.4, C.gold, C.white);
    s.addText("关键区分", {
      x: MARGIN + 0.85, y: 1.5, w: 3.0, h: 0.4,
      fontSize: 17, fontFace: FONT_B, color: C.accent, bold: true, margin: 0,
    });
    s.addText([
      { text: "光程差 \u0394 本身", options: { bold: true, color: C.accent, fontSize: 13, breakLine: false } },
      { text: " \u2192 决定亮暗（干涉级次）", options: { fontSize: 13, color: C.text, breakLine: true } },
      { text: "\u2202\u0394/\u2202\u03B8", options: { bold: true, color: C.secondary, fontSize: 13, breakLine: false } },
      { text: " \u2192 决定条纹形态：=0为平直条纹，\u22600为同心圆环", options: { fontSize: 13, color: C.text } },
    ], {
      x: MARGIN + 0.85, y: 1.9, w: 7.5, h: 0.55,
      fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0,
    });

    // Two-source model (left)
    s.addShape("rect", {
      x: MARGIN, y: 2.7, w: 4.0, h: 2.35,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("两虚点光源模型", {
      x: MARGIN + 0.3, y: 2.85, w: 3.4, h: 0.35,
      fontSize: 15, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    // simplified diagram
    s.addShape("ellipse", { x: MARGIN + 0.5, y: 3.5, w: 0.25, h: 0.25, fill: { color: C.accent } });
    s.addText("S\u2081", { x: MARGIN + 0.5, y: 3.3, w: 0.25, h: 0.2, fontSize: 8, fontFace: FONT_B, color: C.accent, align: "center", margin: 0 });
    s.addShape("ellipse", { x: MARGIN + 2.0, y: 3.5, w: 0.25, h: 0.25, fill: { color: C.secondary } });
    s.addText("S\u2082", { x: MARGIN + 2.0, y: 3.3, w: 0.25, h: 0.2, fontSize: 8, fontFace: FONT_B, color: C.secondary, align: "center", margin: 0 });
    s.addText("\u2190 2h \u2192", {
      x: MARGIN + 0.8, y: 3.4, w: 1.2, h: 0.2,
      fontSize: 9, fontFace: FONT_B, color: C.muted, align: "center", margin: 0,
    });
    s.addShape("line", { x: MARGIN + 0.75, y: 3.62, w: 1.25, h: 0, line: { color: C.muted, width: 1, dashType: "dash" } });
    // screen
    s.addShape("rect", { x: MARGIN + 2.9, y: 3.15, w: 0.08, h: 1.0, fill: { color: C.muted } });
    s.addText("屏", { x: MARGIN + 3.05, y: 3.5, w: 0.4, h: 0.3, fontSize: 9, fontFace: FONT_B, color: C.muted, margin: 0 });

    s.addText([
      { text: "\u0394(\u03B8) \u2248 2h cos\u03B8", options: { fontSize: 14, color: C.accent, bold: true, italic: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "\u2022 h 较大 \u2192 \u2202\u0394/\u2202\u03B8 \u2260 0 \u2192 同心圆环", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u2022 h \u2192 0 \u2192 \u2202\u0394/\u2202\u03B8 = 0 \u2192 平直条纹（等厚）", options: { fontSize: 11, color: C.text } },
    ], {
      x: MARGIN + 0.3, y: 4.2, w: 3.4, h: 0.8,
      fontFace: FONT_B, lineSpacingMultiple: 1.2, margin: 0,
    });

    // Right side: morphology judgment
    s.addShape("rect", {
      x: 5.3, y: 2.7, w: 4.0, h: 2.35,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("形态判据总结", {
      x: 5.6, y: 2.85, w: 3.4, h: 0.35,
      fontSize: 15, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    s.addText([
      { text: "条纹是圆环还是平直，唯一取决于 h 是否为零 —— 与绝对光程差的大小毫无关系。", options: { fontSize: 12, color: C.text, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "实验操作中调出平直条纹 = 将 M1 调至 h = 0，即两虚光源在空间上重合。", options: { fontSize: 12, color: C.text } },
    ], {
      x: 5.6, y: 3.35, w: 3.4, h: 1.5,
      fontFace: FONT_B, lineSpacingMultiple: 1.3, valign: "top", margin: 0,
    });

    addSlideNumber(s, 7, TOTAL);
  }

  // ==================== SLIDE 8: 原理一 —— 红光实验（视深效应） ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "原理一：红光实验 —— 虚像位移补偿");
    addFooterBar(s);

    // Left: effects table
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 4.4, h: 2.0,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("插入玻璃的双重效应", {
      x: MARGIN + 0.3, y: 1.5, w: 3.8, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    const tblHeader = [
      { text: "效应", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, align: "center" } },
      { text: "机制", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, align: "center" } },
      { text: "对 \u2202\u0394/\u2202\u03B8", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, align: "center" } },
    ];
    const tblRows = [
      [
        { text: "减速效应", options: { fontSize: 10, bold: true, color: C.text, align: "center" } },
        { text: "光速变慢\u2192固定光程偏移", options: { fontSize: 10, color: C.text, align: "center" } },
        { text: "零（常数求导）", options: { fontSize: 10, color: C.text, align: "center" } },
      ],
      [
        { text: "视深效应", options: { fontSize: 10, bold: true, color: C.accent, align: "center" } },
        { text: "折射\u2192M1虚像前移", options: { fontSize: 10, color: C.text, align: "center" } },
        { text: "非零（h\u22600）", options: { fontSize: 10, color: C.accent, bold: true, align: "center" } },
      ],
    ];
    s.addTable([tblHeader, ...tblRows], {
      x: MARGIN + 0.3, y: 2.0, w: 3.8,
      colW: [1.1, 1.5, 1.2],
      rowH: [0.35, 0.35, 0.35],
      border: { pt: 0.5, color: "DDDDDD" },
      margin: [2, 4, 2, 4],
    });

    // Right: compensation logic
    s.addShape("rect", {
      x: 5.05, y: 1.35, w: 4.25, h: 2.0,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("补偿逻辑", {
      x: 5.35, y: 1.5, w: 3.6, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    s.addText([
      { text: "虚像前移量：\u0394x' = d(1 \u2212 1/n)", options: { fontSize: 12, color: C.accent, bold: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "\u2192 M1 向后移动相同距离 \u0394x = \u0394x'", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u2192 h 重新归零 \u2192 平直条纹恢复", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "n = d / (d \u2212 \u0394x)", options: { fontSize: 14, color: C.accent, bold: true, italic: true } },
    ], {
      x: 5.35, y: 2.0, w: 3.6, h: 1.2,
      fontFace: FONT_B, lineSpacingMultiple: 1.2, valign: "top", margin: 0,
    });

    // Bottom: key insight
    s.addShape("rect", {
      x: MARGIN, y: 3.6, w: 8.6, h: 1.5,
      fill: { color: "FDE8E8" },
    });
    addIcon(s, "\uD83D\uDCA1", MARGIN + 0.3, 3.75, 0.35, 0.35, C.accent, C.white);
    s.addText("核心洞察", {
      x: MARGIN + 0.8, y: 3.75, w: 3.0, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0,
    });
    s.addText([
      { text: "激光相干长度极长（数十厘米），对固定光程偏移 2d(n\u22121) 完全不敏感。", options: { fontSize: 12, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "条纹形态只响应于虚像间距 h 的变化。M1 后退纯粹是为了让虚像在几何上退回 h=0，完全无关光程差补偿。", options: { fontSize: 12, color: C.text } },
    ], {
      x: MARGIN + 0.8, y: 4.2, w: 7.8, h: 0.85,
      fontFace: FONT_B, lineSpacingMultiple: 1.3, margin: 0,
    });

    addSlideNumber(s, 8, TOTAL);
  }

  // ==================== SLIDE 9: 原理二 —— 白光实验（零光程差） ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "原理二：白光实验 —— 零光程差补偿");
    addFooterBar(s);

    // Left: coherence constraint
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 4.4, h: 2.3,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText('白光的「光程极零」强制约', {
      x: MARGIN + 0.3, y: 1.5, w: 3.8, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    s.addText([
      { text: "白光为宽频带复色光，相干长度仅微米量级。", options: { fontSize: 12, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "核心要求：OPD \u2248 0", options: { fontSize: 13, color: C.gold, bold: true, breakLine: true } },
      { text: "各色条纹须在视场中心同相位叠加，才能在中心区域观察到少量彩色干涉条纹。", options: { fontSize: 12, color: C.text } },
    ], {
      x: MARGIN + 0.3, y: 2.0, w: 3.8, h: 1.5,
      fontFace: FONT_B, lineSpacingMultiple: 1.3, valign: "top", margin: 0,
    });

    // Right: compensation derivation
    s.addShape("rect", {
      x: 5.05, y: 1.35, w: 4.25, h: 2.3,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("补偿推导", {
      x: 5.35, y: 1.5, w: 3.6, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    s.addText([
      { text: "插入玻璃 \u2192 OPD 跳变：", options: { fontSize: 11, color: C.text, breakLine: false } },
      { text: "\u0394(OPD) = 2d(n\u22121)", options: { fontSize: 12, color: C.gold, bold: true, italic: true, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "M1 推进 \u0394x \u2192 光程减少 2\u0394x", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "补偿守恒：2\u0394x = 2d(n\u22121)", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 6, breakLine: true } },
      { text: "n = 1 + \u0394x / d", options: { fontSize: 14, color: C.gold, bold: true, italic: true } },
    ], {
      x: 5.35, y: 2.0, w: 3.6, h: 1.5,
      fontFace: FONT_B, lineSpacingMultiple: 1.2, valign: "top", margin: 0,
    });

    // Bottom: key insight
    s.addShape("rect", {
      x: MARGIN, y: 3.9, w: 8.6, h: 1.15,
      fill: { color: "FFF8E1" },
    });
    addIcon(s, "\uD83D\uDCA1", MARGIN + 0.3, 4.05, 0.35, 0.35, C.gold, C.white);
    s.addText("核心洞察", {
      x: MARGIN + 0.8, y: 4.05, w: 3.0, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.gold, bold: true, margin: 0,
    });
    s.addText("白光仅要求 OPD \u2248 0，不要求 h = 0。补偿后残留 h \u2248 0.06mm 使条纹曲率极小，在劈尖角主导下呈现近似平行的彩色外观。", {
      x: MARGIN + 0.8, y: 4.42, w: 7.8, h: 0.55,
      fontSize: 12, fontFace: FONT_B, color: C.text, lineSpacingMultiple: 1.2, margin: 0,
    });
    addSlideNumber(s, 9, TOTAL);
  }

  // ==================== SLIDE 10: 原理对比 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "原理对比与实验设计呼应");
    addFooterBar(s);

    // Comparison table
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 8.6, h: 2.1,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    const compHeader = [
      { text: "", options: { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 11 } },
      { text: "红光实验", options: { fill: { color: C.accent }, color: C.white, bold: true, fontSize: 11, align: "center" } },
      { text: "白光实验", options: { fill: { color: C.gold }, color: C.white, bold: true, fontSize: 11, align: "center" } },
    ];
    const compRows = [
      [
        { text: "物理机制", options: { bold: true, fontSize: 10, color: C.text } },
        { text: "视深效应 \u2192 虚像位移", options: { fontSize: 10, color: C.text, align: "center" } },
        { text: "光程跳变 \u2192 OPD补偿", options: { fontSize: 10, color: C.text, align: "center" } },
      ],
      [
        { text: "M1 移动方向", options: { bold: true, fontSize: 10, color: C.text } },
        { text: "远离分光板 \u2190", options: { fontSize: 10, color: C.accent, bold: true, align: "center" } },
        { text: "\u2192 靠近分光板", options: { fontSize: 10, color: C.gold, bold: true, align: "center" } },
      ],
      [
        { text: "玻璃厚度", options: { bold: true, fontSize: 10, color: C.text } },
        { text: "7 mm（厚）", options: { fontSize: 10, color: C.text, align: "center" } },
        { text: "0.3 mm（薄）", options: { fontSize: 10, color: C.text, align: "center" } },
      ],
      [
        { text: "公式", options: { bold: true, fontSize: 10, color: C.text } },
        { text: "n = d/(d\u2212\u0394x)", options: { fontSize: 10, color: C.accent, italic: true, align: "center" } },
        { text: "n = 1+\u0394x/d", options: { fontSize: 10, color: C.gold, italic: true, align: "center" } },
      ],
    ];
    s.addTable([compHeader, ...compRows], {
      x: MARGIN + 0.3, y: 1.55, w: 8.0,
      colW: [1.6, 3.2, 3.2],
      rowH: [0.4, 0.35, 0.35, 0.35, 0.35],
      border: { pt: 0.5, color: "DDDDDD" },
      margin: [2, 6, 2, 6],
    });

    // Bottom insight: thickness choice
    s.addShape("rect", {
      x: MARGIN, y: 3.7, w: 8.6, h: 1.4,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    addIcon(s, "\u2194", MARGIN + 0.3, 3.85, 0.35, 0.35, C.secondary, C.white);
    s.addText("玻璃厚度选择的物理必然性", {
      x: MARGIN + 0.8, y: 3.85, w: 7.4, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.text, bold: true, margin: 0,
    });
    s.addText([
      { text: "红光实验（厚 7mm）：虚像位移 \u2248 2.33mm，信号大、易测量。激光不计较相伴的 7mm 级光程偏移。", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "", options: { fontSize: 5, breakLine: true } },
      { text: "白光实验（薄 0.3mm）：OPD 跳变 \u2248 0.3mm，补偿位移 \u2248 0.15mm 可追踪。若用 7mm 厚玻璃，OPD 跳变远超微米级相干长度，实验上不可能重新捕捉彩色条纹。", options: { fontSize: 11, color: C.text } },
    ], {
      x: MARGIN + 0.8, y: 4.25, w: 7.8, h: 0.8,
      fontFace: FONT_B, lineSpacingMultiple: 1.3, margin: 0,
    });

    addSlideNumber(s, 10, TOTAL);
  }

  // ==================== SLIDE 11: 数据处理与结果 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "数据处理与结果");
    addFooterBar(s);

    // Red laser data + result
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 3.9, h: 3.55,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("红光实验 (d = 7.0 mm)", {
      x: MARGIN + 0.3, y: 1.5, w: 3.3, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0,
    });
    s.addText([
      { text: "\u0394x\u2081 = 2.79632 mm", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u0394x\u2082 = 2.80415 mm", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u0394x\u2083 = 2.78987 mm", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "均值 \u0394x\u0304 = 2.79678 mm", options: { fontSize: 11, color: C.text, bold: true, breakLine: true } },
      { text: "s_\u0394x = 0.00716 mm", options: { fontSize: 11, color: C.muted, breakLine: true } },
      { text: "u(\u0394x\u0304) = 0.00413 mm", options: { fontSize: 11, color: C.muted } },
    ], {
      x: MARGIN + 0.3, y: 2.0, w: 3.3, h: 2.0,
      fontFace: FONT_B, lineSpacingMultiple: 1.4, valign: "top", margin: 0,
    });
    s.addShape("rect", {
      x: MARGIN + 0.3, y: 4.1, w: 3.3, h: 0.6,
      fill: { color: "FDE8E8" },
    });
    s.addText([
      { text: "n", options: { fontSize: 16, color: C.accent, bold: true, italic: true, breakLine: false } },
      { text: "\u7EA2", options: { fontSize: 12, color: C.accent, bold: true, breakLine: false } },
      { text: " = 1.6655 \u00B1 0.0016", options: { fontSize: 13, color: C.accent, bold: true } },
    ], {
      x: MARGIN + 0.3, y: 4.1, w: 3.3, h: 0.6,
      fontFace: FONT_B, align: "center", valign: "middle", margin: 0,
    });

    // White light data + result
    s.addShape("rect", {
      x: 5.4, y: 1.35, w: 3.9, h: 3.55,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("白光实验 (d = 0.3 mm)", {
      x: 5.7, y: 1.5, w: 3.3, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.gold, bold: true, margin: 0,
    });
    s.addText([
      { text: "\u0394x\u2081 = 0.17201 mm", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u0394x\u2082 = 0.17055 mm", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u0394x\u2083 = 0.17389 mm", options: { fontSize: 11, color: C.text, breakLine: true } },
      { text: "均值 \u0394x\u0304 = 0.17215 mm", options: { fontSize: 11, color: C.text, bold: true, breakLine: true } },
      { text: "s_\u0394x = 0.00168 mm", options: { fontSize: 11, color: C.muted, breakLine: true } },
      { text: "u(\u0394x\u0304) = 0.00097 mm", options: { fontSize: 11, color: C.muted } },
    ], {
      x: 5.7, y: 2.0, w: 3.3, h: 2.0,
      fontFace: FONT_B, lineSpacingMultiple: 1.4, valign: "top", margin: 0,
    });
    s.addShape("rect", {
      x: 5.7, y: 4.1, w: 3.3, h: 0.6,
      fill: { color: "FFF8E1" },
    });
    s.addText([
      { text: "n", options: { fontSize: 16, color: C.gold, bold: true, italic: true, breakLine: false } },
      { text: "\u767D", options: { fontSize: 12, color: C.gold, bold: true, breakLine: false } },
      { text: " = 1.5738 \u00B1 0.0032", options: { fontSize: 13, color: C.gold, bold: true } },
    ], {
      x: 5.7, y: 4.1, w: 3.3, h: 0.6,
      fontFace: FONT_B, align: "center", valign: "middle", margin: 0,
    });

    addSlideNumber(s, 11, TOTAL);
  }

  // ==================== SLIDE 12: 误差分析 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.lightBg };
    addSectionTitle(s, "误差分析");
    addFooterBar(s);

    // Red laser errors
    s.addShape("rect", {
      x: MARGIN, y: 1.35, w: 4.0, h: 2.7,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("红光实验误差源", {
      x: MARGIN + 0.3, y: 1.5, w: 3.4, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.accent, bold: true, margin: 0,
    });
    s.addText([
      { text: "条纹判断误差", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "\u300C平直程度\u300D依赖主观判断", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "千分尺读数误差", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "回差或视差影响 \u0394x", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "玻璃板放置偏差", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "未严格垂直入射 \u2192 系统偏差", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "厚度 d 测量误差", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "d 偏差会被放大或缩小", options: { fontSize: 10, color: C.muted } },
    ], {
      x: MARGIN + 0.3, y: 2.0, w: 3.4, h: 1.9,
      fontFace: FONT_B, lineSpacingMultiple: 1.25, valign: "top", margin: 0,
    });

    // White light errors
    s.addShape("rect", {
      x: 5.3, y: 1.35, w: 4.0, h: 2.7,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    s.addText("白光实验误差源", {
      x: 5.6, y: 1.5, w: 3.4, h: 0.35,
      fontSize: 14, fontFace: FONT_B, color: C.gold, bold: true, margin: 0,
    });
    s.addText([
      { text: "相干条件更苛刻", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "条纹消失/重现判定更敏感", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "光源对准不稳", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "扩展光源偏差 \u2192 彩条不清晰", options: { fontSize: 10, color: C.muted, breakLine: true } },
      { text: "薄片厚度偏差", options: { bold: true, fontSize: 11, color: C.text, breakLine: true } },
      { text: "d仅0.3mm, 微小偏差被显著放大", options: { fontSize: 10, color: C.muted } },
    ], {
      x: 5.6, y: 2.0, w: 3.4, h: 1.9,
      fontFace: FONT_B, lineSpacingMultiple: 1.25, valign: "top", margin: 0,
    });

    // Overall judgment
    s.addShape("rect", {
      x: MARGIN, y: 4.3, w: 8.6, h: 0.75,
      fill: { color: C.cardBg }, shadow: makeShadow(),
    });
    addIcon(s, "\u2714", MARGIN + 0.3, 4.47, 0.35, 0.35, "27AE60", C.white);
    s.addText("两者均落在常见玻璃折射率范围内，从不同物理机制出发得到量级合理的结果，相互印证了两种补偿逻辑的正确性。", {
      x: MARGIN + 0.8, y: 4.3, w: 7.8, h: 0.75,
      fontSize: 12, fontFace: FONT_B, color: C.text, valign: "middle", lineSpacingMultiple: 1.3, margin: 0,
    });

    addSlideNumber(s, 12, TOTAL);
  }

  // ==================== SLIDE 13: 结论 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

    s.addText("结论", {
      x: MARGIN, y: 0.3, w: 8.6, h: 0.6,
      fontSize: 32, fontFace: FONT_H, color: C.white, bold: true, margin: 0,
    });

    const conclusions = [
      {
        num: "01",
        title: "红光实验",
        text: "需将 M1 远离分光板以补偿视深效应，恢复等厚干涉几何条件。n = d/(d\u2212\u0394x)，n\u7EA2 = 1.6655 \u00B1 0.0016",
      },
      {
        num: "02",
        title: "白光实验",
        text: "需将 M1 靠近分光板以补偿光程跳变，回归 OPD \u2248 0。n = 1+\u0394x/d，n\u767D = 1.5738 \u00B1 0.0032",
      },
      {
        num: "03",
        title: "方向差异的根源",
        text: "激光相干长度极长，只响应虚像间距 h（几何条件）；白光相干长度极短，只响应绝对光程差 OPD \u2248 0。两种光源对干涉条件的选择性不同导致了完全相反的补偿方向。",
      },
    ];

    conclusions.forEach((c, i) => {
      const y0 = 1.3 + i * 1.3;
      s.addShape("rect", {
        x: MARGIN, y: y0 + 0.05, w: 0.55, h: 0.55,
        fill: { color: C.primary },
      });
      s.addText(c.num, {
        x: MARGIN, y: y0 + 0.05, w: 0.55, h: 0.55,
        fontSize: 20, fontFace: FONT_H, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(c.title, {
        x: MARGIN + 0.75, y: y0, w: 8.0, h: 0.35,
        fontSize: 16, fontFace: FONT_B, color: C.gold, bold: true, margin: 0,
      });
      s.addText(c.text, {
        x: MARGIN + 0.75, y: y0 + 0.38, w: 8.0, h: 0.55,
        fontSize: 12, fontFace: FONT_B, color: "D0D8E0", lineSpacingMultiple: 1.3, margin: 0,
      });
    });

    addSlideNumber(s, 13, TOTAL, true);
  }

  // ==================== SLIDE 14: 致谢 ====================
  {
    let s = pres.addSlide();
    s.background = { color: C.darkBg };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });

    s.addText("谢谢！", {
      x: 1.0, y: 1.5, w: 8.0, h: 1.2,
      fontSize: 52, fontFace: FONT_H, color: C.white, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addShape("rect", { x: 3.2, y: 2.9, w: 3.6, h: 0.03, fill: { color: C.gold } });
    s.addText("感谢聆听 · 欢迎提问", {
      x: 1.0, y: 3.2, w: 8.0, h: 0.6,
      fontSize: 18, fontFace: FONT_B, color: "AABBCC",
      align: "center", margin: 0,
    });
    s.addText("迈克尔逊干涉仪测折射率实验探究", {
      x: 1.0, y: 4.2, w: 8.0, h: 0.4,
      fontSize: 13, fontFace: FONT_B, color: "667788",
      align: "center", margin: 0,
    });
    addSlideNumber(s, 14, TOTAL, true);
  }

  // ─── Write file ───
  const outPath = "d:/project/迈克尔逊干涉仪实验探究/答辩PPT.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log("PPT saved to: " + outPath);
}

main().catch(err => { console.error(err); process.exit(1); });
