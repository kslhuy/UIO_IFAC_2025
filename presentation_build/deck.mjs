import pptxgen from "pptxgenjs";
import fs from "node:fs/promises";
import path from "node:path";

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Q.H. Nguyen et al.";
pptx.company = "Université de Lorraine · CRAN · SEGULA Engineering";
pptx.subject = "15-minute conference presentation";
pptx.title = "Learning with Unknown Input Observers for Robust Nonlinear Estimation";
pptx.lang = "en-US";
pptx.theme = {
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  lang: "en-US",
};
pptx.defineLayout({ name: "IFAC_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "IFAC_WIDE";
pptx.margin = 0;

const root = path.resolve("..");
const output = path.join(root, "UIO_IFAC_15min_presentation.pptx");
const paperRef = "User-provided manuscript: ifac/UIO_IFAC.tex";
const C = {
  ink: "111827",
  muted: "52606D",
  faint: "E8EDF1",
  panel: "F3F5F7",
  blue: "1677C8",
  blueSoft: "DCEEF9",
  green: "2FAE75",
  purple: "9B6CC1",
  orange: "EE9A17",
  red: "E5543D",
  white: "FFFFFF",
};
const FONT = "Arial";
const M = 0.55;
const W = 13.333;

function addText(slide, text, x, y, w, h, options = {}) {
  slide.addText(text, {
    x, y, w, h,
    margin: options.margin ?? 0,
    // Cambria Math is not embedded reliably by the local PowerPoint PDF renderer;
    // Arial renders the Unicode mathematical notation consistently on all slides.
    fontFace: options.fontFace === "Cambria Math" ? FONT : (options.fontFace ?? FONT),
    fontSize: options.fontSize ?? 19,
    color: options.color ?? C.ink,
    bold: options.bold ?? false,
    italic: options.italic ?? false,
    align: options.align ?? "left",
    valign: options.valign ?? "top",
    breakLine: false,
    fit: "shrink",
    paraSpaceAfterPt: options.paraSpaceAfterPt ?? 0,
    bullet: options.bullet,
    transparency: options.transparency,
    isTextBox: true,
  });
}

function rect(slide, x, y, w, h, fill, line = null, radius = false) {
  slide.addShape(radius ? pptx.ShapeType.roundRect : pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: fill },
    line: line ? { color: line.color ?? line, width: line.width ?? 0.8 } : { color: fill, transparency: 100 },
    radius: radius ? 0.06 : undefined,
  });
}

function line(slide, x, y, w, h, color = C.faint, width = 1, arrow = false) {
  slide.addShape(pptx.ShapeType.line, {
    x, y, w, h,
    line: {
      color,
      width,
      beginArrowType: "none",
      endArrowType: arrow ? "triangle" : "none",
    },
  });
}

function addHeader(slide, n, section) {
  addText(slide, "LEARNING WITH UNKNOWN INPUT OBSERVERS", M, 0.22, 5.6, 0.22, {
    fontSize: 9.5, color: C.muted, bold: true,
  });
  if (section) {
    addText(slide, section.toUpperCase(), 8.25, 0.22, 4.05, 0.22, {
      fontSize: 9.5, color: C.muted, bold: true, align: "right",
    });
  }
  line(slide, M, 0.52, 12.23, 0, C.faint, 0.7);
  addText(slide, String(n).padStart(2, "0"), 12.28, 7.06, 0.45, 0.2, {
    fontSize: 10, color: C.muted, align: "right",
  });
}

function addTitle(slide, title, subtitle = "") {
  addText(slide, title, M, 0.7, 12.15, 0.52, { fontSize: 32, bold: true, color: C.ink });
  if (subtitle) addText(slide, subtitle, M, 1.23, 12.0, 0.3, { fontSize: 15.5, color: C.muted });
}

function addNotes(slide, timing, cue, sources = []) {
  const sourceList = [paperRef, ...sources].map((s) => `- ${s}`).join("\n");
  slide.addNotes(`Timing: ${timing}\nPresenter cue: ${cue}\n\n[Sources]\n${sourceList}`);
}

function boxLabel(slide, text, x, y, w, color = C.blue) {
  rect(slide, x, y, w, 0.28, color);
  addText(slide, text.toUpperCase(), x + 0.1, y + 0.06, w - 0.2, 0.14, {
    fontSize: 9.5, color: C.white, bold: true, align: "center",
  });
}

function bulletList(slide, items, x, y, w, h, options = {}) {
  const gap = options.gap ?? 0.54;
  items.forEach((item, i) => {
    rect(slide, x, y + i * gap + 0.11, 0.08, 0.08, options.dotColor ?? C.blue);
    addText(slide, item, x + 0.2, y + i * gap, w - 0.2, options.height ?? Math.max(0.4, gap - 0.08), {
      fontSize: options.fontSize ?? 18.5, color: options.color ?? C.ink,
    });
  });
}

function image(slide, file, x, y, w, h, altText) {
  slide.addImage({ path: path.join(root, file), x, y, w, h, altText, transparency: 0 });
}

// 1 — Cover
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  rect(slide, 8.85, 0, 4.483, 7.5, C.blue);
  rect(slide, 9.27, 0.56, 3.65, 6.38, "0D5B9E");
  addText(slide, "IFAC 2025", M, 0.58, 2.2, 0.25, { fontSize: 12, bold: true, color: C.blue });
  addText(slide, "Learning with\nUnknown Input Observers\nfor Robust Nonlinear Estimation", M, 1.2, 7.75, 2.26, {
    fontSize: 34, bold: true, color: C.ink, valign: "mid",
  });
  line(slide, M, 3.86, 6.48, 0, C.blue, 2.5);
  addText(slide, "Q.H. Nguyen · L. Qi · A. Zemouche · H. Rafaralahy · M. Haddad", M, 4.18, 7.25, 0.34, {
    fontSize: 17, color: C.ink,
  });
  addText(slide, "Université de Lorraine · CRAN · SEGULA Engineering", M, 4.62, 6.55, 0.28, {
    fontSize: 15, color: C.muted,
  });
  addText(slide, "Physics-bound estimation\n+ online residual learning", 9.72, 1.55, 2.82, 1.1, {
    fontSize: 25, color: C.white, bold: true, valign: "mid",
  });
  line(slide, 9.72, 3.08, 2.2, 0, C.blueSoft, 1.4);
  addText(slide, "A teacher–student observer architecture for vehicle systems with unmodeled dynamics.", 9.72, 3.34, 2.64, 1.0, {
    fontSize: 16.5, color: C.blueSoft,
  });
  addText(slide, "15-minute presentation", 9.72, 6.35, 2.4, 0.25, { fontSize: 13, color: C.white });
  addNotes(slide, "0:35", "Frame the talk around one question: how can online learning help estimate unknown vehicle dynamics without giving up observer guarantees?");
}

// 2 — Motivation
{
  const slide = pptx.addSlide();
  addHeader(slide, 2, "Motivation");
  addTitle(slide, "Unknown dynamics challenge model-based estimation", "Vehicle dynamics are known only approximately—and the uncertainty affects estimation and control.");
  rect(slide, 0.55, 1.83, 6.05, 4.62, C.panel);
  boxLabel(slide, "Nonlinear plant", 0.82, 2.16, 1.72, C.ink);
  addText(slide, "ẋ = φ(x, u) + B μₓ\ny = Cx", 0.85, 2.75, 5.15, 0.9, { fontFace: "Cambria Math", fontSize: 29, color: C.ink });
  addText(slide, "μₓ captures unmodeled dynamics\nand external disturbances", 0.85, 4.05, 4.65, 0.76, {
    fontSize: 23, bold: true, color: C.blue,
  });
  addText(slide, "Examples in a vehicle: tire-road effects, aerodynamic mismatch, parameter variations.", 0.85, 5.16, 5.05, 0.62, {
    fontSize: 18, color: C.muted,
  });
  line(slide, 6.9, 2.1, 0, 3.78, C.faint, 1);
  addText(slide, "Why not simply learn μₓ?", 7.35, 2.12, 4.95, 0.38, { fontSize: 25, bold: true });
  bulletList(slide, [
    "Pure online learning can be unpredictable under distribution shifts.",
    "Safety-critical estimation needs a bounded-error reference.",
    "Classical UIOs often fail for the available sensor set.",
  ], 7.42, 2.86, 4.85, 2.2, { fontSize: 18.5, gap: 0.78 });
  rect(slide, 7.35, 5.55, 4.85, 0.84, C.blueSoft);
  addText(slide, "Goal: learn the repeatable part of μₓ while retaining an observer-based guarantee.", 7.62, 5.72, 4.35, 0.5, { fontSize: 15.5, bold: true, color: "105B94", align: "center", valign: "mid" });
  addNotes(slide, "1:05", "Introduce μₓ as the mismatch between the nominal vehicle model and the road. The point is not that neural networks are bad; it is that their online behaviour needs a physics-based reference.");
}

// 3 — Rank limitation
{
  const slide = pptx.addSlide();
  addHeader(slide, 3, "Motivation");
  addTitle(slide, "Classical UIOs can exclude practical sensor suites", "Exact unknown-input decoupling requires a rank condition that may not hold with the sensors already on the vehicle.");
  rect(slide, 0.55, 1.82, 12.2, 1.16, C.ink);
  addText(slide, "rank(CB) = rank(B)", 0.9, 2.1, 4.05, 0.45, { fontFace: "Cambria Math", fontSize: 31, color: C.white, bold: true });
  addText(slide, "Necessary condition for direct unknown-input estimation", 5.12, 2.2, 6.9, 0.24, { fontSize: 18.5, color: C.white });
  rect(slide, 0.55, 3.46, 5.85, 2.36, C.panel);
  boxLabel(slide, "If the condition holds", 0.82, 3.75, 2.04, C.green);
  bulletList(slide, ["A classical UIO can decouple the unknown input.", "In the ideal case, estimation errors converge exponentially."], 0.88, 4.24, 5.0, 1.1, { dotColor: C.green, fontSize: 18.5, gap: 0.63 });
  rect(slide, 6.88, 3.46, 5.87, 2.36, "FBECE8");
  boxLabel(slide, "If it fails", 7.15, 3.75, 1.45, C.red);
  bulletList(slide, ["Sensor limitations can prevent exact decoupling.", "Common workarounds amplify noise, add sensors, or restrict disturbance dynamics."], 7.21, 4.24, 5.0, 1.1, { dotColor: C.red, fontSize: 18.5, gap: 0.63 });
  addText(slide, "This motivates a bounded first estimate rather than demanding exactness immediately.", 0.55, 6.10, 11.6, 0.48, { fontSize: 18.5, bold: true, color: C.blue, valign: "mid" });
  addNotes(slide, "1:00", "Explain why the rank condition is the core practical obstacle. The method starts by accepting that exact decoupling is not always possible, then controls the resulting estimation error.");
}

// 4 — Regularization
{
  const slide = pptx.addSlide();
  addHeader(slide, 4, "Layer 1 · Regularized UIO");
  addTitle(slide, "Regularization yields a controlled H¹ error bound", "A rank-restoring rewrite is algebraically equivalent to the original system; the induced terms are treated as bounded disturbances.");
  // Connectors are placed before nodes so they stay behind the content.
  line(slide, 3.48, 3.9, 0.78, 0, C.blue, 2, true);
  line(slide, 7.55, 3.9, 0.78, 0, C.blue, 2, true);
  rect(slide, 0.55, 2.52, 2.92, 2.58, C.panel, { color: C.faint, width: 0.8 });
  rect(slide, 4.27, 2.52, 3.28, 2.58, C.blueSoft, { color: "B8DDF3", width: 0.8 });
  rect(slide, 8.33, 2.52, 4.42, 2.58, C.panel, { color: C.faint, width: 0.8 });
  boxLabel(slide, "Original system", 0.84, 2.81, 1.63, C.ink);
  addText(slide, "ẋ = φ(x,u) + Bμₓ\ny = Cx", 0.84, 3.42, 2.02, 0.76, { fontFace: "Cambria Math", fontSize: 23 });
  boxLabel(slide, "Regularize", 4.59, 2.81, 2.0, C.blue);
  addText(slide, "Eδ₁ μ̇ₓ  and  Dδ₂ μₓ\nare added and subtracted", 4.6, 3.42, 2.58, 0.76, { fontSize: 19.5, bold: true, color: "105B94", align: "center" });
  boxLabel(slide, "Observer design", 8.66, 2.81, 1.58, C.green);
  addText(slide, "The remaining terms become ωδ.\nLMI design minimizes their effect\non the augmented error.", 8.66, 3.37, 3.38, 1.02, { fontSize: 19.3 });
  rect(slide, 0.55, 5.72, 12.2, 0.62, C.ink);
  addText(slide, "‖ẑ − z‖H¹  ≤  λδ ‖μₓ‖H¹ + β‖ξ₀‖    ·    λδ is explicitly driven by δ₁ and δ₂", 0.9, 5.91, 11.5, 0.24, {
    fontFace: "Cambria Math", fontSize: 19.5, color: C.white, align: "center",
  });
  addNotes(slide, "1:15", "Walk through the principle, not the full derivation. E and D are selected to recover the descriptor rank condition, while δ₁ and δ₂ set the size of the perturbation entering the H¹ bound.");
}

// 5 — Generalized UIO
{
  const slide = pptx.addSlide();
  addHeader(slide, 5, "Layer 1 · Generalized UIO");
  addTitle(slide, "Output derivatives recover the strong case", "The generalized observer augments the model with the measured output and a filtered derivative state.");
  // Connectors first.
  line(slide, 2.94, 3.6, 0.56, 0, C.muted, 1.4, true);
  line(slide, 6.26, 3.6, 0.56, 0, C.muted, 1.4, true);
  line(slide, 9.56, 3.6, 0.55, 0, C.muted, 1.4, true);
  rect(slide, 0.62, 2.84, 2.32, 1.55, C.panel, { color: C.faint, width: 0.8 });
  rect(slide, 3.5, 2.84, 2.76, 1.55, C.blueSoft, { color: "B8DDF3", width: 0.8 });
  rect(slide, 6.82, 2.84, 2.74, 1.55, C.panel, { color: C.faint, width: 0.8 });
  rect(slide, 10.11, 2.84, 2.62, 1.55, "EAF7F0", { color: "B8E4CC", width: 0.8 });
  addText(slide, "Measured output\ny = Cx", 0.9, 3.22, 1.75, 0.6, { fontSize: 23, bold: true, align: "center" });
  addText(slide, "Augmented states\nχ₁ = y,  χ₂ = ẏ", 3.8, 3.17, 2.15, 0.65, { fontFace: "Cambria Math", fontSize: 22, bold: true, align: "center", color: "105B94" });
  addText(slide, "Generalized\nregularized model", 7.08, 3.22, 2.22, 0.6, { fontSize: 22, bold: true, align: "center" });
  addText(slide, "Generalized UIO\nẑ = [χ̂₁, χ̂₂, x̂, μ̂ₓ]ᵀ", 10.38, 3.17, 2.08, 0.65, { fontFace: "Cambria Math", fontSize: 19, bold: true, align: "center", color: "1A7750" });
  rect(slide, 0.62, 5.12, 12.11, 0.82, C.blueSoft);
  addText(slide, "When rank(CB) = rank(B), choose δ₁ = δ₂ = 0: the generalized UIO recovers exponential convergence. Otherwise, it retains the adjustable H¹ bound.", 0.93, 5.36, 11.5, 0.28, { fontSize: 18.5, bold: true, color: "105B94", align: "center" });
  addText(slide, "χ̂₂ acts as a filtered estimate of ẏ, avoiding a posteriori numerical differentiation of noisy measurements.", 0.62, 6.08, 11.8, 0.32, { fontSize: 15.5, color: C.muted, align: "center" });
  addNotes(slide, "1:20", "This is the key extension. Do not claim that the method differentiates noisy data afterwards: the augmented observer estimates the derivative state internally. Emphasize the recovery of exponential convergence in the rank-satisfied case.");
}

// 6 — LMI guarantee
{
  const slide = pptx.addSlide();
  addHeader(slide, 6, "Layer 1 · Guarantee");
  addTitle(slide, "LMIs certify a bounded physics-based teacher signal", "The Layer 1 observer is designed over a polytopic representation of the nonlinear Jacobian.");
  rect(slide, 0.55, 1.9, 4.1, 3.95, C.panel);
  boxLabel(slide, "Design ingredients", 0.85, 2.22, 1.8, C.ink);
  bulletList(slide, [
    "Lipschitz nonlinearity → polytope vertices Aⱼ.",
    "Decision variables: P ≻ 0 and Q.",
    "Gain: K = P⁻¹Qᵀ.",
  ], 0.92, 2.87, 3.3, 1.8, { fontSize: 16.2, gap: 0.84, height: 0.7 });
  rect(slide, 4.93, 1.9, 7.82, 3.95, C.ink);
  addText(slide, "Certified conclusion", 5.36, 2.29, 3.9, 0.34, { fontSize: 24, bold: true, color: C.white });
  line(slide, 5.36, 2.86, 6.76, 0, "52606D", 0.8);
  addText(slide, "‖ζ − ζ̂‖H¹ ≤ λδ ‖μₓ‖H¹ + β‖ξ₀‖", 5.36, 3.26, 6.76, 0.55, { fontFace: "Cambria Math", fontSize: 28, color: C.white, bold: true, align: "center" });
  addText(slide, "The bound is robust to the regularization disturbance and initial error.", 5.58, 4.15, 6.3, 0.36, { fontSize: 19, color: C.blueSoft, align: "center" });
  addText(slide, "Layer 1 is not the final learned model—it is the safe supervisory reference for Layer 2.", 0.55, 6.14, 12.1, 0.52, { fontSize: 18.5, bold: true, color: C.blue, align: "center", valign: "mid" });
  addNotes(slide, "1:00", "Keep this at the level of the result. The LMI converts the nonlinear error analysis into a tractable design condition and certifies the UIO estimates used as targets later.");
}

// 7 — Teacher–student architecture
{
  const slide = pptx.addSlide();
  addHeader(slide, 7, "Layer 2 · Learning architecture");
  addTitle(slide, "Layer 2 learns a predictive residual model", "The neural adaptive observer uses the UIO output as a physics-consistent proxy target during online adaptation.");
  // Connectors first.
  line(slide, 2.4, 3.8, 0.5, 0, C.muted, 1.7, true);
  line(slide, 5.47, 3.8, 0.48, 0, C.muted, 1.7, true);
  line(slide, 8.63, 3.8, 0.48, 0, C.muted, 1.7, true);
  line(slide, 5.0, 4.9, 0, 0.66, C.blue, 1.4, true);
  rect(slide, 0.56, 2.95, 1.84, 1.72, C.panel, { color: C.faint, width: 0.8 });
  rect(slide, 2.9, 2.95, 2.57, 1.72, C.blueSoft, { color: "B8DDF3", width: 0.8 });
  rect(slide, 5.95, 2.95, 2.68, 1.72, "EAF7F0", { color: "B8E4CC", width: 0.8 });
  rect(slide, 9.11, 2.95, 3.62, 1.72, C.panel, { color: C.faint, width: 0.8 });
  addText(slide, "Vehicle\nplant", 0.87, 3.47, 1.22, 0.6, { fontSize: 24, bold: true, align: "center" });
  addText(slide, "Generalized UIO\nLayer 1", 3.3, 3.3, 1.8, 0.66, { fontSize: 24, bold: true, align: "center", color: "105B94" });
  addText(slide, "μθ(·)\nNeural model", 6.32, 3.28, 1.94, 0.68, { fontFace: "Cambria Math", fontSize: 24, bold: true, align: "center", color: "1A7750" });
  addText(slide, "Neural adaptive observer\nLayer 2", 9.46, 3.3, 2.92, 0.66, { fontSize: 24, bold: true, align: "center" });
  addText(slide, "y, u", 2.45, 3.38, 0.4, 0.2, { fontSize: 15, color: C.muted, align: "center" });
  addText(slide, "x̂UIO, μ̂ₓ", 5.41, 3.36, 0.56, 0.3, { fontFace: "Cambria Math", fontSize: 13.5, color: C.muted, align: "center" });
  addText(slide, "μθ", 8.68, 3.37, 0.38, 0.2, { fontFace: "Cambria Math", fontSize: 15, color: C.muted, align: "center" });
  rect(slide, 3.14, 5.40, 3.72, 0.68, C.blueSoft);
  addText(slide, "Composite loss regularizes Layer 2\nwith the Layer 1 reference.", 3.42, 5.56, 3.17, 0.34, { fontSize: 16.6, bold: true, color: "105B94", align: "center", valign: "mid" });
  addText(slide, "Outcome: Layer 1 gives bounded residuals; Layer 2 learns the repeatable state-dependent component.", 0.56, 6.32, 12.1, 0.42, { fontSize: 15.6, color: C.muted, align: "center", valign: "mid" });
  addNotes(slide, "1:20", "This is the central system diagram. Explain that the UIO is the teacher signal, not a frozen baseline: it remains in the loop to regularize online learning. The controller ultimately uses the Layer 2 state estimate.");
}

// 8 — Loss and neural observer guarantee
{
  const slide = pptx.addSlide();
  addHeader(slide, 8, "Layer 2 · Objective");
  addTitle(slide, "The loss keeps learning tethered to physics", "The loss combines task performance with measurement consistency and UIO supervision.");
  rect(slide, 0.55, 1.91, 12.2, 1.22, C.ink);
  addText(slide, "L = tracking error + output error + UIO consistency + residual regularization", 0.92, 2.31, 11.48, 0.36, { fontFace: "Cambria Math", fontSize: 25, color: C.white, bold: true, align: "center" });
  const terms = [
    ["Track", "x̂NN stays near xref", C.blue],
    ["Measure", "yNN stays near y", C.green],
    ["Match", "x̂NN stays near x̂UIO", C.purple],
    ["Learn", "μθ stays near μ̂ₓ", C.orange],
  ];
  terms.forEach((term, i) => {
    const x = 0.55 + i * 3.07;
    rect(slide, x, 3.74, 2.85, 1.8, C.panel);
    rect(slide, x, 3.74, 2.85, 0.22, term[2]);
    addText(slide, term[0], x + 0.2, 4.19, 2.45, 0.32, { fontSize: 23, bold: true, color: term[2], align: "center" });
    addText(slide, term[1], x + 0.22, 4.78, 2.41, 0.42, { fontSize: 17.5, align: "center", color: C.ink });
  });
  rect(slide, 0.55, 6.02, 12.2, 0.54, C.blueSoft);
  addText(slide, "For a given approximation mismatch Δx, an LMI certifies: ‖εNN‖L² ≤ λNN‖Δx‖L² + βNN‖εNN(0)‖.", 0.88, 6.18, 11.55, 0.21, { fontFace: "Cambria Math", fontSize: 18, bold: true, color: "105B94", align: "center" });
  addNotes(slide, "1:10", "Connect each loss term to a role. The UIO estimate acts as a proxy label because ground-truth μₓ is unavailable online. The L2 result isolates how estimation quality depends on the residual approximation error.");
}

// 9 — Vehicle setup
{
  const slide = pptx.addSlide();
  addHeader(slide, 9, "Vehicle study");
  addTitle(slide, "Vehicle study: state and tire-force reconstruction", "A reduced vehicle model is simulated with tire-force residuals and parametric mismatch.");
  rect(slide, 0.55, 1.94, 4.0, 3.98, C.panel);
  boxLabel(slide, "System variables", 0.88, 2.27, 1.63, C.ink);
  addText(slide, "State", 0.9, 2.88, 0.75, 0.28, { fontSize: 20, bold: true, color: C.muted });
  addText(slide, "x = [vx, vy, ψ, r]ᵀ", 1.86, 2.87, 2.2, 0.3, { fontFace: "Cambria Math", fontSize: 20.5 });
  addText(slide, "Input", 0.9, 3.53, 0.75, 0.28, { fontSize: 20, bold: true, color: C.muted });
  addText(slide, "u = [a, δ]ᵀ", 1.86, 3.52, 2.2, 0.3, { fontFace: "Cambria Math", fontSize: 20.5 });
  addText(slide, "Measured", 0.9, 4.18, 1.2, 0.28, { fontSize: 18, bold: true, color: C.muted });
  addText(slide, "y = [vx, ψ, r]ᵀ", 2.13, 4.17, 1.93, 0.3, { fontFace: "Cambria Math", fontSize: 20.5 });
  addText(slide, "Unknown", 0.9, 4.83, 1.2, 0.28, { fontSize: 18, bold: true, color: C.muted });
  addText(slide, "μₓ = [fyf, fyr]ᵀ", 2.13, 4.82, 1.93, 0.3, { fontFace: "Cambria Math", fontSize: 20.5, color: C.red, bold: true });
  // Simple vehicle diagram: connectors first, then nodes.
  line(slide, 6.0, 3.58, 0.95, 0, C.muted, 1.5, true);
  line(slide, 8.65, 3.58, 0.95, 0, C.muted, 1.5, true);
  rect(slide, 5.05, 2.63, 0.95, 1.85, C.blueSoft, { color: "B8DDF3", width: 0.8 });
  rect(slide, 6.95, 2.63, 1.7, 1.85, C.ink);
  rect(slide, 9.6, 2.63, 2.5, 1.85, "EAF7F0", { color: "B8E4CC", width: 0.8 });
  addText(slide, "y", 5.18, 3.30, 0.68, 0.36, { fontFace: "Cambria Math", fontSize: 27, bold: true, color: "105B94", align: "center" });
  addText(slide, "Vehicle\nmodel", 7.28, 3.1, 1.04, 0.6, { fontSize: 23, bold: true, color: C.white, align: "center" });
  addText(slide, "Two-layer\nestimator", 10.04, 3.1, 1.62, 0.6, { fontSize: 22, bold: true, color: "1A7750", align: "center" });
  addText(slide, "Reconstruct: unmeasured lateral velocity vy and tire-force residuals fyf, fyr.", 5.05, 5.35, 6.92, 0.35, { fontSize: 21, bold: true, color: C.blue, align: "center" });
  addNotes(slide, "0:55", "State only what is measured versus reconstructed. The same controller and reference are used for both observer layers; Layer 2 supplies feedback for trajectory tracking.");
}

// 10 — Tracking result
{
  const slide = pptx.addSlide();
  addHeader(slide, 10, "Vehicle study · Results");
  addTitle(slide, "Layer 2 supports close tracking and smooth control", "The Layer 2 estimates provide useful feedback for the trajectory-tracking controller in the simulated maneuver.");
  rect(slide, 0.55, 1.72, 8.25, 5.1, C.panel, { color: C.faint, width: 0.8 });
  image(slide, "img/traj_control_py.png", 0.68, 1.84, 7.98, 4.79, "Vehicle trajectory, tracking error, steering angle, and throttle input");
  rect(slide, 9.18, 1.92, 3.48, 1.55, C.blue);
  addText(slide, "0.44 cm", 9.48, 2.26, 2.88, 0.43, { fontSize: 31, bold: true, color: C.white, align: "center" });
  addText(slide, "mean position error", 9.48, 2.82, 2.88, 0.22, { fontSize: 15, color: C.white, align: "center" });
  bulletList(slide, [
    "Estimated trajectory follows GPS closely.",
    "Steering and throttle profiles remain smooth and physically consistent.",
  ], 9.32, 4.0, 3.18, 1.2, { fontSize: 17.6, gap: 0.76 });
  addText(slide, "The position-error metric is reported directly in the simulation figure.", 9.3, 6.16, 3.16, 0.36, { fontSize: 13.7, italic: true, color: C.muted, align: "center" });
  addNotes(slide, "1:05", "Begin with the top-left trajectory and compare the estimates visually. Then use the mean position error from the figure to anchor the claim. The bottom panels show that the corresponding input profiles stay well behaved.", ["User-provided simulation figure: img/traj_control_py.png"]);
}

// 11 — State result
{
  const slide = pptx.addSlide();
  addHeader(slide, 11, "Vehicle study · Results");
  addTitle(slide, "The neural layer preserves state-estimation fidelity", "Measured states are recovered closely, while the unmeasured lateral velocity is reconstructed through vehicle-dynamics coupling.");
  rect(slide, 0.55, 1.72, 8.25, 5.1, C.panel, { color: C.faint, width: 0.8 });
  image(slide, "img/State_est_py.png", 0.68, 1.84, 7.98, 4.79, "State estimation plots for longitudinal and lateral velocity, yaw rate, and yaw angle");
  rect(slide, 9.12, 2.0, 3.55, 1.17, "EAF7F0");
  addText(slide, "Measured: vx, ψ, r", 9.36, 2.29, 3.05, 0.25, { fontFace: "Cambria Math", fontSize: 21.5, bold: true, color: "1A7750", align: "center" });
  rect(slide, 9.12, 3.48, 3.55, 1.17, C.blueSoft);
  addText(slide, "Reconstructed: vy", 9.36, 3.78, 3.05, 0.25, { fontFace: "Cambria Math", fontSize: 21.5, bold: true, color: "105B94", align: "center" });
  addText(slide, "Layer 1 and neural estimates remain close to the true trajectories; the difficult unmeasured channel is recovered via the model coupling.", 9.2, 5.25, 3.35, 0.88, { fontSize: 17.4, color: C.ink, align: "center", valign: "mid" });
  addNotes(slide, "1:00", "Use the lateral-velocity panel to make the key point: vy is not directly measured. The Layer 2 neural estimate stays close to the true trajectory while retaining the UIO reference structure.", ["User-provided simulation figure: img/State_est_py.png"]);
}

// 12 — Force residual result
{
  const slide = pptx.addSlide();
  addHeader(slide, 12, "Vehicle study · Results");
  addTitle(slide, "Learning improves tire-force residual reconstruction", "Both layers are bounded; Layer 2 converges faster and tracks the simulated residuals more closely.");
  rect(slide, 0.55, 1.72, 8.25, 5.1, C.panel, { color: C.faint, width: 0.8 });
  image(slide, "img/tireforce_slip_py.png", 0.68, 1.84, 7.98, 4.79, "Rear and front tire residual and slip angle estimation plots");
  rect(slide, 9.12, 1.96, 3.55, 1.14, C.purple);
  addText(slide, "Layer 1 · UIO", 9.38, 2.27, 3.03, 0.24, { fontSize: 22, bold: true, color: C.white, align: "center" });
  addText(slide, "bounded physics reference", 9.38, 2.61, 3.03, 0.2, { fontSize: 14.5, color: C.white, align: "center" });
  rect(slide, 9.12, 3.42, 3.55, 1.14, C.blue);
  addText(slide, "Layer 2 · Neural", 9.38, 3.73, 3.03, 0.24, { fontSize: 22, bold: true, color: C.white, align: "center" });
  addText(slide, "faster convergence and closer tracking", 9.38, 4.07, 3.03, 0.2, { fontSize: 14.5, color: C.white, align: "center" });
  addText(slide, "The learned map refines an instantaneous residual estimate into a predictive state-dependent correction.", 9.18, 5.18, 3.34, 0.9, { fontSize: 17.2, bold: true, color: C.ink, align: "center", valign: "mid" });
  addNotes(slide, "1:20", "This is the payoff slide. Compare purple Layer 1 versus blue neural estimates against the green truth for the front and rear residuals. State the conclusion cautiously: it is demonstrated in this simulation.", ["User-provided simulation figure: img/tireforce_slip_py.png"]);
}

// 13 — Close
{
  const slide = pptx.addSlide();
  addHeader(slide, 13, "Conclusion");
  addText(slide, "Safe learning for unknown vehicle dynamics", 0.55, 0.95, 10.8, 0.55, { fontSize: 34, bold: true });
  const points = [
    ["1", "Rank-relaxed UIO", "A generalized, regularized observer produces bounded estimates even when exact unknown-input decoupling is unavailable.", C.blue],
    ["2", "UIO-supervised learning", "The neural observer learns the repeatable part of the unknown input while remaining consistent with the physics-based reference.", C.green],
    ["3", "Vehicle evidence", "Simulation shows accurate trajectory tracking, state reconstruction, and improved tire-force residual estimation.", C.orange],
  ];
  points.forEach((p, i) => {
    const x = 0.55 + i * 4.08;
    rect(slide, x, 2.1, 3.62, 3.25, C.panel);
    rect(slide, x + 0.25, 2.4, 0.58, 0.58, p[3], null, true);
    addText(slide, p[0], x + 0.25, 2.54, 0.58, 0.18, { fontSize: 18, color: C.white, bold: true, align: "center" });
    addText(slide, p[1], x + 1.02, 2.44, 2.3, 0.32, { fontSize: 22, bold: true, color: p[3] });
    addText(slide, p[2], x + 0.3, 3.42, 3.02, 1.14, { fontSize: 18.2, color: C.ink, align: "center", valign: "mid" });
  });
  rect(slide, 0.55, 6.05, 12.2, 0.62, C.ink);
  addText(slide, "Next: validation with real driving data and a discrete-time embedded implementation.", 0.86, 6.24, 11.58, 0.22, { fontSize: 20, color: C.white, bold: true, align: "center" });
  addText(slide, "Thank you", 0.55, 6.91, 3.0, 0.25, { fontSize: 15.5, color: C.muted });
  addNotes(slide, "0:45", "Return to the opening question. The contribution is a learning system that does not discard physics: the UIO provides a certified teacher signal, and Layer 2 turns it into a predictive residual model. Invite questions.");
}

await fs.writeFile("source-notes.txt", `${paperRef}\nFigures used: img/traj_control_py.png, img/State_est_py.png, img/tireforce_slip_py.png\nAll claims and figures in the deck are derived from the user-provided manuscript and local simulation outputs.\n`);
await pptx.writeFile({ fileName: output });
console.log(output);
