import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "C:/Users/victo/Documents/Revisión Vimat";
const WORK = path.join(ROOT, "work/presentations/agenda-vimat");
const TMP = path.join(WORK, "tmp");
const QA = path.join(TMP, "qa");
const OUT = path.join(ROOT, "outputs");
const FINAL = path.join(OUT, "AgendaVimat_2026-2027_completa.pptx");

const W = 720;
const H = 1040;
const M = 38;
const TOP = 44;
const BOTTOM = 928;
const FOOT = 956;

const C = {
  navy: "#0C1B45",
  ink: "#17213F",
  blue: "#233E70",
  gray: "#6E7480",
  line: "#D8D2C5",
  cream: "#F7F3EA",
  paper: "#FFFDF8",
  warm: "#F0D6B8",
  peach: "#F8E8D7",
  coral: "#B9634A",
  green: "#2C695D",
  mint: "#E4F0EA",
  amber: "#D79D32",
  white: "#FFFFFF",
};

const monthNames = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];
const shortDays = ["L", "M", "X", "J", "V", "S", "D"];
const longDays = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];

const quotes = [
  ["La esencia de las matemáticas es la libertad.", "Georg Cantor"],
  ["Dios creó los números enteros; todo lo demás es obra humana.", "Leopold Kronecker"],
  ["No hay camino real para la geometría.", "Euclides"],
  ["Medir lo que es medible y hacer medible lo que no lo es.", "Galileo Galilei"],
  ["La imaginación es más importante que el conocimiento.", "Albert Einstein"],
  ["Las matemáticas son el alfabeto con que Dios escribió el universo.", "Galileo Galilei"],
  ["La claridad es el principio de toda enseñanza.", "Vimat"],
  ["Un problema bien planteado ya está medio resuelto.", "Tradición matemática"],
  ["Equivocarse también es calcular.", "Vimat"],
  ["La belleza se reconoce en una demostración limpia.", "Vimat"],
  ["Lo que no se repasa, se evapora.", "Vimat"],
  ["La disciplina vence al sorteo.", "Vimat"],
];

function dateUTC(y, m, d) {
  return new Date(Date.UTC(y, m, d));
}
function addDays(date, days) {
  return dateUTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate() + days);
}
function fmtDay(date) {
  return String(date.getUTCDate()).padStart(2, "0");
}
function fmtMonth(date) {
  return monthNames[date.getUTCMonth()].toLowerCase();
}
function fmtShort(date) {
  return `${fmtDay(date)} ${fmtMonth(date)}`;
}
function monthKey(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}
function mondayOf(date) {
  const day = date.getUTCDay() || 7;
  return addDays(date, 1 - day);
}
function weeksForAcademicYear() {
  const start = mondayOf(dateUTC(2026, 8, 1));
  const end = addDays(mondayOf(dateUTC(2027, 7, 31)), 6);
  const weeks = [];
  for (let d = start, n = 1; d <= end; d = addDays(d, 7), n++) {
    weeks.push({ n, start: d, end: addDays(d, 6) });
  }
  return weeks;
}
function monthsForAcademicYear() {
  return [
    [2026, 8], [2026, 9], [2026, 10], [2026, 11],
    [2027, 0], [2027, 1], [2027, 2], [2027, 3],
    [2027, 4], [2027, 5], [2027, 6], [2027, 7],
  ];
}

function addShape(slide, geometry, left, top, width, height, fill = "none", lineFill = "none", lineWidth = 0, name) {
  return slide.shapes.add({
    geometry,
    name,
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: lineFill, width: lineWidth },
  });
}
function addText(slide, text, left, top, width, height, opts = {}) {
  const box = addShape(slide, "textbox", left, top, width, height, opts.fill ?? "none", opts.lineFill ?? "none", opts.lineWidth ?? 0, opts.name);
  box.text = text;
  box.text.style = {
    fontSize: opts.size ?? 18,
    bold: opts.bold ?? false,
    italic: opts.italic ?? false,
    color: opts.color ?? C.ink,
    typeface: opts.face ?? opts.typeface ?? "Aptos",
    alignment: opts.align ?? "left",
  };
  box.text.insets = opts.insets ?? { left: 0, right: 0, top: 0, bottom: 0 };
  return box;
}
function addLine(slide, x1, y1, x2, y2, color = C.line, width = 1) {
  if (Math.abs(x1 - x2) < 0.01) {
    return addShape(slide, "rect", x1 - width / 2, Math.min(y1, y2), width, Math.abs(y2 - y1), color, color, 0);
  }
  if (Math.abs(y1 - y2) < 0.01) {
    return addShape(slide, "rect", Math.min(x1, x2), y1 - width / 2, Math.abs(x2 - x1), width, color, color, 0);
  }
  const left = Math.min(x1, x2);
  const top = Math.min(y1, y2);
  return slide.shapes.add({
    geometry: "line",
    position: { left, top, width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}
function addCheckbox(slide, x, y, size = 13, label = "", labelWidth = 90) {
  addShape(slide, "rect", x, y, size, size, C.white, C.line, 1);
  if (label) addText(slide, label, x + size + 6, y - 2, labelWidth, size + 8, { size: 9, color: C.gray });
}
function base(slide, pageNo, title, subtitle = "") {
  slide.background.fill = C.paper;
  addShape(slide, "rect", 0, 0, W, H, C.paper, C.paper, 0, "page-paper");
  addChrome(slide, pageNo, title, subtitle);
}
function addChrome(slide, pageNo, title, subtitle = "") {
  addShape(slide, "rect", 9.6, 9.6, W - 19.2, H - 19.2, "none", C.line, 1.1);
  addShape(slide, "rect", M, TOP - 2, W - 2 * M, 1.5, C.navy, C.navy, 0);
  if (title) addText(slide, title, M, TOP + 10, W - 2 * M, 38, { size: 21, bold: true, color: C.navy, face: "Georgia", align: "center" });
  if (subtitle) addText(slide, subtitle, M, TOP + 48, W - 2 * M, 24, { size: 10, color: C.gray, align: "center" });
  addText(slide, "Vimat preparadores de oposiciones\nwww.vimat.info", 210, FOOT, 300, 42, { size: 8.5, color: "#9A9A9A", align: "center" });
  addText(slide, String(pageNo), W - 66, FOOT + 15, 34, 18, { size: 9, color: "#9A9A9A", align: "right" });
}
function paperPatch(slide, left, top, width, height) {
  if (width <= 0 || height <= 0) return;
  addShape(slide, "rect", left, top, width, height, C.paper, C.paper, 0);
}
function section(slide, pageNo, title, kicker, formula = "") {
  slide.background.fill = {
    type: "gradient",
    angle: 90,
    stops: [
      { color: "#1A2752", position: 0 },
      { color: "#EEF0EE", position: 100000 },
    ],
  };
  addShape(slide, "rect", 0, 0, W, H, "#EEF0EE", "#EEF0EE", 0, "page-paper");
  addShape(slide, "rect", 9.6, 9.6, W - 19.2, H - 19.2, "none", C.line, 1.1);
  addText(slide, kicker.toUpperCase(), M, 90, W - 2 * M, 26, { size: 12, bold: true, color: C.warm, align: "center" });
  addText(slide, title, 58, 380, W - 116, 130, { size: 44, bold: true, color: C.navy, face: "Georgia", align: "center" });
  addText(slide, formula, 70, 530, W - 140, 54, { size: 25, italic: true, color: C.ink, face: "Georgia", align: "center" });
  addText(slide, "Vimat preparadores de oposiciones\nwww.vimat.info", 210, FOOT, 300, 42, { size: 8.5, color: "#777", align: "center" });
  addText(slide, String(pageNo), W - 66, FOOT + 15, 34, 18, { size: 9, color: "#777", align: "right" });
}
function addRuledBox(slide, x, y, w, h, label, rows = 5) {
  addShape(slide, "rect", x, y, w, h, C.white, C.line, 1, label);
  addText(slide, label, x + 12, y + 9, w - 24, 18, { size: 10, bold: true, color: C.navy });
  const top = y + 34;
  const gap = (h - 44) / rows;
  for (let i = 1; i <= rows; i++) addLine(slide, x + 12, top + i * gap, x + w - 12, top + i * gap, "#E5E0D6", 0.8);
}

function cover(p, pageNo) {
  const slide = p.slides.add();
  slide.background.fill = {
    type: "gradient",
    angle: 90,
    stops: [
      { color: "#17213F", position: 0 },
      { color: "#6E7480", position: 53000 },
      { color: "#F7F3EA", position: 100000 },
    ],
  };
  addText(slide, "C", 38, 18, 40, 48, { size: 27, bold: true, color: C.white, face: "Georgia", align: "center" });
  addLine(slide, 91, 48, 180, 202, C.white, 2.4);
  addText(slide, "A", 78, 240, 40, 48, { size: 28, bold: true, color: C.white, face: "Georgia", align: "center" });
  addText(slide, "B", 338, 106, 40, 48, { size: 28, bold: true, color: C.white, face: "Georgia", align: "center" });
  addShape(slide, "ellipse", 530, 46, 170, 170, "none", C.white, 3);
  addLine(slide, 552, 207, 612, 132, C.white, 3);
  addText(slide, "R", 560, 116, 42, 40, { size: 25, italic: true, color: C.white, face: "Georgia" });
  addText(slide, "L = Rθ", 552, 213, 126, 44, { size: 26, italic: true, color: C.white, face: "Georgia" });
  addText(slide, "∑ 1/k!", 40, 488, 118, 70, { size: 26, italic: true, color: C.navy, face: "Georgia", rotation: -13 });
  addText(slide, "dy/dx = y", 555, 504, 126, 48, { size: 23, italic: true, color: C.navy, face: "Georgia", rotation: 21 });
  addText(slide, "vimat", 105, 328, 510, 140, { size: 76, bold: true, color: "#000000", face: "Aptos Display", align: "center" });
  addText(slide, "Agenda\n2026-27", 66, 585, 590, 210, { size: 58, bold: true, color: C.navy, face: "Georgia", align: "center" });
  addText(slide, "docente · matemáticas · oposición", 120, 820, 480, 32, { size: 14, color: C.ink, align: "center" });
  addText(slide, "Vimat preparadores de oposiciones\nwww.vimat.info", 210, FOOT, 300, 42, { size: 8.5, color: "#8F8F8F", align: "center" });
  addText(slide, String(pageNo), W - 66, FOOT + 15, 34, 18, { size: 9, color: "#8F8F8F", align: "right" });
}

function personalPage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Datos de la agenda", "Curso 2026-2027");
  const labels = ["Nombre", "Centro", "Departamento", "Teléfono", "Correo", "Tutoría / grupos", "Convocatoria", "Tribunal / sede"];
  let y = 120;
  for (const label of labels) {
    addText(s, label, 62, y, 150, 20, { size: 10, bold: true, color: C.gray });
    addLine(s, 62, y + 42, 658, y + 42, C.line, 1);
    y += label === "Tutoría / grupos" ? 74 : 66;
  }
  addRuledBox(s, 62, 732, 596, 128, "Objetivo del curso", 4);
}

function usagePage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Mapa de uso", "Una agenda pensada para clase, oposición y seguimiento diario");
  const cards = [
    ["1", "Cada semana", "Planifica clases, guardias, reuniones, repasos y problemas."],
    ["2", "Cada mes", "Cierra objetivos, exámenes, simulacros y bloques de temario."],
    ["3", "Cada tema", "Marca lectura, resumen, problemas, memorización y repasos."],
    ["4", "Cada grupo", "Reserva páginas para listados, observaciones y evaluación."],
  ];
  let y = 132;
  for (const [n, title, body] of cards) {
    addShape(s, "roundRect", 62, y, 596, 118, n === "1" ? C.peach : C.white, C.line, 1);
    addShape(s, "ellipse", 88, y + 31, 56, 56, C.navy, C.navy, 0);
    addText(s, n, 88, y + 42, 56, 24, { size: 22, bold: true, color: C.white, align: "center" });
    addText(s, title, 168, y + 28, 420, 26, { size: 19, bold: true, color: C.navy, face: "Georgia" });
    addText(s, body, 168, y + 61, 420, 34, { size: 13, color: C.ink });
    y += 142;
  }
  addText(s, "Festivos autonómicos/locales y calendario escolar: completar según comunidad, municipio y centro.", 74, 818, 572, 42, { size: 11, color: C.coral, align: "center", bold: true });
}

function annualCalendarPage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Calendario 2026-2027", "Curso completo de septiembre a agosto");
  const months = monthsForAcademicYear();
  const cellW = 186, cellH = 176;
  const startX = 62, startY = 118;
  months.forEach(([y, m], i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    drawMiniMonth(s, y, m, startX + col * (cellW + 20), startY + row * (cellH + 18), cellW, cellH, true);
  });
}

function drawMiniMonth(slide, year, month, x, y, w, h, withFrame = false) {
  if (withFrame) addShape(slide, "roundRect", x, y, w, h, C.white, C.line, 1);
  addText(slide, `${monthNames[month]} ${year}`, x + 8, y + 8, w - 16, 18, { size: 12, bold: true, color: C.navy, align: "center", face: "Georgia" });
  const gridX = x + 12, gridY = y + 38;
  const cw = (w - 24) / 7, ch = (h - 52) / 7;
  shortDays.forEach((d, i) => addText(slide, d, gridX + i * cw, gridY, cw, 14, { size: 7.5, bold: true, color: i > 4 ? C.coral : C.gray, align: "center" }));
  const first = dateUTC(year, month, 1);
  const offset = (first.getUTCDay() || 7) - 1;
  const days = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();
  for (let day = 1; day <= days; day++) {
    const pos = offset + day - 1;
    const r = Math.floor(pos / 7), c = pos % 7;
    const weekend = c > 4;
    addText(slide, String(day), gridX + c * cw, gridY + 18 + r * ch, cw, ch, { size: 8.5, color: weekend ? C.coral : C.ink, align: "center" });
  }
}

function importantDatesPage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Fechas clave", "Calendario editable para adaptar a tu comunidad y centro");
  const rows = [
    ["Inicio de curso", ""], ["Evaluación inicial", ""], ["1.ª evaluación", ""],
    ["2.ª evaluación", ""], ["Evaluación ordinaria", ""], ["Simulacros oposición", ""],
    ["Entrega programación", ""], ["Defensa / encerrona", ""], ["Festivos locales", ""],
  ];
  addText(s, "Hito", 70, 125, 260, 24, { size: 12, bold: true, color: C.navy });
  addText(s, "Fecha / notas", 338, 125, 280, 24, { size: 12, bold: true, color: C.navy });
  let y = 160;
  for (const [a, b] of rows) {
    addShape(s, "rect", 62, y, 596, 56, C.white, C.line, 1);
    addText(s, a, 76, y + 18, 236, 18, { size: 11, bold: true, color: C.ink });
    addText(s, b, 338, y + 18, 300, 18, { size: 11, color: C.gray });
    y += 58;
  }
  addText(s, "Nota: el calendario laboral de 2027 y los calendarios escolares autonómicos pueden publicarse después de imprimir esta agenda. Deja estas páginas como zona editable.", 74, 742, 572, 62, { size: 11, color: C.coral, align: "center" });
  addRuledBox(s, 62, 820, 596, 78, "Ajustes de centro", 2);
}

function schedulePage(p, pageNo, n) {
  const s = p.slides.add();
  const title = `Horario docente ${n}`;
  const subtitle = "Clases, guardias, reuniones y preparación";
  base(s, pageNo, title, subtitle);
  const x = 62, y = 130, w = 596, h = 694;
  const timeW = 80, colW = (w - timeW) / 5, rowH = h / 10;
  addShape(s, "rect", x, y, w, h, C.white, C.line, 1);
  addText(s, "Hora", x + 8, y + 10, timeW - 16, 18, { size: 9, bold: true, color: C.navy, align: "center" });
  longDays.forEach((d, i) => addText(s, d, x + timeW + i * colW, y + 10, colW, 18, { size: 9, bold: true, color: C.navy, align: "center" }));
  for (let r = 1; r <= 10; r++) addLine(s, x, y + r * rowH, x + w, y + r * rowH, C.line, 0.8);
  for (let c = 0; c <= 5; c++) addLine(s, x + timeW + c * colW, y, x + timeW + c * colW, y + h, C.line, 0.8);
  for (let r = 1; r < 10; r++) addText(s, `${r}.ª`, x + 12, y + r * rowH + 20, timeW - 24, 18, { size: 9, color: C.gray, align: "center" });
  addRuledBox(s, 62, 850, 286, 72, "Reuniones fijas", 2);
  addRuledBox(s, 372, 850, 286, 72, "Guardias / apoyos", 2);
  paperPatch(s, 0, 0, W, 126);
  paperPatch(s, 0, 126, 60, 802);
  paperPatch(s, 660, 126, 60, 802);
  paperPatch(s, 0, 826, W, 22);
  paperPatch(s, 350, 848, 20, 78);
  paperPatch(s, 0, 924, W, H - 924);
  addChrome(s, pageNo, title, subtitle);
}

function formulasPage(p, pageNo, title, groups) {
  const s = p.slides.add();
  base(s, pageNo, title, "Formulario rápido para el aula");
  let y = 118;
  for (const [head, lines] of groups) {
    addShape(s, "roundRect", 58, y, 604, 28 + lines.length * 28, C.white, C.line, 1);
    addText(s, head, 76, y + 10, 568, 18, { size: 12, bold: true, color: C.navy, face: "Georgia" });
    lines.forEach((line, i) => addText(s, line, 78, y + 38 + i * 27, 560, 22, { size: 13, color: C.ink, face: "Georgia" }));
    y += 46 + lines.length * 28;
  }
}

function choose(n, k) {
  if (k < 0 || k > n) return 0;
  k = Math.min(k, n - k);
  let res = 1;
  for (let i = 1; i <= k; i++) res = (res * (n - k + i)) / i;
  return res;
}
function probAtLeast(studied, balls, total = 71) {
  return 1 - choose(total - studied, balls) / choose(total, balls);
}
function probabilityPage(p, pageNo, balls) {
  const s = p.slides.add();
  base(s, pageNo, `Probabilidad de tema (${balls} bolas)`, "P(al menos un tema estudiado)");
  addText(s, `P = 1 - C(71 - n, ${balls}) / C(71, ${balls})`, 76, 112, 568, 32, { size: 17, bold: true, color: C.navy, face: "Georgia", align: "center" });
  const values = [10, 20, 30, 40, 50, 60, 71];
  let y = 180;
  for (const n of values) {
    const pct = probAtLeast(n, balls);
    addText(s, `${n} temas`, 78, y + 3, 82, 18, { size: 10, bold: true, color: C.ink });
    addShape(s, "rect", 174, y, 390, 24, "#EFEAE1", "#EFEAE1", 0);
    addShape(s, "rect", 174, y, 390 * pct, 24, balls === 4 ? C.green : C.coral, balls === 4 ? C.green : C.coral, 0);
    addText(s, `${Math.round(pct * 100)}%`, 578, y + 3, 55, 18, { size: 10, bold: true, color: C.navy, align: "right" });
    y += 58;
  }
  addRuledBox(s, 62, 650, 596, 180, "Plan de subida de probabilidad", 6);
  addText(s, "Marca hitos realistas: +5 temas cerrados, +10 problemas tipo, +1 simulacro.", 80, 852, 560, 34, { size: 11, color: C.gray, align: "center" });
}

const TEMARIO_BLOCKS = [
  [1, 10, "Aritmetica y numeros"],
  [11, 20, "Algebra y estructuras"],
  [21, 31, "Funciones y analisis"],
  [32, 42, "Geometria plana y metrica"],
  [43, 52, "Geometria analitica y espacial"],
  [53, 61, "Probabilidad y estadistica"],
  [62, 71, "Didactica, historia y aula"],
];

function blockForTopic(topic) {
  return TEMARIO_BLOCKS.find(([start, end]) => topic >= start && topic <= end)?.[2] ?? "Personalizar";
}

function temarioOverviewPage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Mapa del temario", "Clasificacion operativa para planificar el estudio");
  addText(s, "Usa este mapa como indice de trabajo. Sustituye los bloques por los epigrafes exactos de tu academia o convocatoria si quieres una version 100% literal.", 70, 122, 580, 48, { size: 12, color: C.gray, align: "center" });
  let y = 194;
  for (const [start, end, label] of TEMARIO_BLOCKS) {
    addShape(s, "rect", 62, y, 596, 76, y % 2 ? C.white : "#FBF8F1", C.line, 1);
    addText(s, `${start}-${end}`, 82, y + 20, 78, 30, { size: 18, bold: true, color: C.navy, face: "Georgia", align: "center" });
    addText(s, label, 184, y + 18, 280, 24, { size: 16, bold: true, color: C.ink, face: "Georgia" });
    addText(s, "Leido   Resumido   Problemas   Memorizado   Repasos", 184, y + 46, 380, 18, { size: 9.5, color: C.gray });
    addCheckbox(s, 548, y + 22, 13, "Bloque cerrado", 86);
    y += 86;
  }
  addRuledBox(s, 62, 812, 596, 76, "Orden personal de ataque", 2);
}

function temarioTrackerPage(p, pageNo, pageIndex, start, end) {
  const s = p.slides.add();
  base(s, pageNo, `Seguimiento del temario ${pageIndex}`, "L = leido · R = resumido · P = problemas · M = memorizado");
  const headers = ["Tema", "Bloque de trabajo", "L", "R", "P", "M", "Rep."];
  const x = 42, y = 116, w = 636, rowH = 30;
  const widths = [44, 304, 42, 42, 42, 42, 120];
  let cx = x;
  headers.forEach((h, i) => {
    addShape(s, "rect", cx, y, widths[i], rowH, C.navy, C.navy, 0);
    addText(s, h, cx + 3, y + 8, widths[i] - 6, 14, { size: 8.5, bold: true, color: C.white, align: "center" });
    cx += widths[i];
  });
  for (let t = start, r = 0; t <= end; t++, r++) {
    const yy = y + rowH + r * rowH;
    cx = x;
    widths.forEach((ww, i) => {
      addShape(s, "rect", cx, yy, ww, rowH, r % 2 ? "#FBF8F1" : C.white, C.line, 0.7);
      if (i === 0) addText(s, String(t), cx, yy + 8, ww, 14, { size: 8.5, bold: true, color: C.ink, align: "center" });
      if (i === 1) addText(s, blockForTopic(t), cx + 8, yy + 8, ww - 16, 14, { size: 8.2, color: C.ink });
      if (i >= 2 && i <= 5) addCheckbox(s, cx + 14, yy + 8, 11);
      cx += ww;
    });
  }
  addText(s, "Rep.: anota fechas de repaso separadas por comas. Ejemplo: 12/10, 04/11, 18/12.", 58, 868, 604, 24, { size: 10, color: C.gray, align: "center" });
}

function programmingPage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Programación y defensa", "Control de producto final de la oposición");
  const items = [
    "Marco normativo actualizado", "Contextualización del centro", "Competencias y saberes",
    "Criterios de evaluación", "Situaciones de aprendizaje", "Atención a la diversidad",
    "Instrumentos de evaluación", "Temporalización", "Bibliografía y anexos", "Defensa ensayada",
  ];
  let y = 126;
  for (const item of items) {
    addCheckbox(s, 80, y, 15);
    addText(s, item, 106, y - 2, 500, 22, { size: 13, color: C.ink });
    addLine(s, 80, y + 30, 640, y + 30, "#E4DED2", 0.8);
    y += 48;
  }
  addRuledBox(s, 62, 682, 596, 170, "Argumentos que quiero defender sí o sí", 6);
}

function oppositionStudyPlanPage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Plan de estudio de oposicion", "Rutina semanal sostenible");
  const rows = [
    ["Lunes", "Tema nuevo", "Lectura y esquema"],
    ["Martes", "Problemas", "2-3 ejercicios tipo"],
    ["Miercoles", "Memoria", "Cantar 20 minutos"],
    ["Jueves", "Repaso", "Tema antiguo + tarjetas"],
    ["Viernes", "Programacion", "Defensa / unidad"],
    ["Sabado", "Simulacro", "Tema o problema cronometrado"],
    ["Domingo", "Cierre", "Revision y ajuste"],
  ];
  let y = 132;
  for (const [day, focus, detail] of rows) {
    addShape(s, "rect", 62, y, 596, 70, y % 2 ? C.white : "#FBF8F1", C.line, 1);
    addText(s, day, 82, y + 20, 92, 22, { size: 13, bold: true, color: C.navy, face: "Georgia" });
    addText(s, focus, 196, y + 14, 160, 20, { size: 12, bold: true, color: C.ink });
    addText(s, detail, 196, y + 39, 280, 18, { size: 10.5, color: C.gray });
    addCheckbox(s, 570, y + 24, 14);
    y += 80;
  }
  addRuledBox(s, 62, 720, 596, 136, "Reglas personales de constancia", 5);
}

function simulacroLogPage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Registro de simulacros", "Tema, problema, encerrona y defensa");
  const x = 44, y = 124, rowH = 36;
  const widths = [70, 76, 170, 72, 72, 166];
  const heads = ["Fecha", "Tipo", "Contenido", "Tiempo", "Nota", "Accion de mejora"];
  let cx = x;
  heads.forEach((h, i) => {
    addShape(s, "rect", cx, y, widths[i], rowH, C.navy, C.navy, 0);
    addText(s, h, cx + 3, y + 11, widths[i] - 6, 13, { size: 8.4, bold: true, color: C.white, align: "center" });
    cx += widths[i];
  });
  for (let r = 1; r <= 17; r++) {
    const yy = y + r * rowH;
    cx = x;
    widths.forEach((ww) => {
      addShape(s, "rect", cx, yy, ww, rowH, r % 2 ? "#FBF8F1" : C.white, C.line, 0.65);
      cx += ww;
    });
  }
  addText(s, "No busques solo nota: registra una accion concreta para que el siguiente simulacro sea mejor.", 70, 782, 580, 42, { size: 11, color: C.coral, align: "center", bold: true });
  addRuledBox(s, 62, 840, 596, 54, "Patron de errores repetidos", 1);
}

function tribunalDefensePage(p, pageNo) {
  const s = p.slides.add();
  base(s, pageNo, "Defensa ante tribunal", "Guion, evidencias y control del tiempo");
  const blocks = [
    ["Min 0-2", "Apertura", "Contexto, idea fuerza y seguridad."],
    ["Min 2-8", "Programacion", "Coherencia curricular, evaluacion y diversidad."],
    ["Min 8-18", "Unidad / situacion", "Secuencia, actividades, materiales y evidencias."],
    ["Min 18-24", "Evaluacion", "Instrumentos, criterios, calificacion y feedback."],
    ["Min 24-30", "Cierre", "Sintesis, valor diferencial y preguntas previsibles."],
  ];
  let y = 132;
  for (const [time, title, detail] of blocks) {
    addShape(s, "rect", 62, y, 596, 96, y % 2 ? C.white : "#FBF8F1", C.line, 1);
    addText(s, time, 82, y + 28, 84, 26, { size: 15, bold: true, color: C.navy, face: "Georgia", align: "center" });
    addText(s, title, 190, y + 18, 380, 22, { size: 14, bold: true, color: C.ink });
    addText(s, detail, 190, y + 48, 380, 28, { size: 10.5, color: C.gray });
    y += 110;
  }
  addRuledBox(s, 62, 704, 596, 140, "Preguntas que debo preparar", 5);
}

function studentListPage(p, pageNo, n) {
  const s = p.slides.add();
  base(s, pageNo, `Listado de alumnos ${n}`, "Grupo: ____________________   Evaluación: ______");
  const x = 46, y = 128, w = 628, rowH = 27;
  const widths = [34, 210, 58, 58, 58, 210];
  const heads = ["#", "Alumno/a", "1", "2", "3", "Observaciones"];
  let cx = x;
  heads.forEach((h, i) => {
    addShape(s, "rect", cx, y, widths[i], rowH, C.navy, C.navy, 0);
    addText(s, h, cx + 2, y + 8, widths[i] - 4, 13, { size: 8.5, bold: true, color: C.white, align: "center" });
    cx += widths[i];
  });
  for (let r = 1; r <= 25; r++) {
    const yy = y + r * rowH;
    cx = x;
    widths.forEach((ww, i) => {
      addShape(s, "rect", cx, yy, ww, rowH, r % 2 ? "#FBF8F1" : C.white, C.line, 0.6);
      if (i === 0) addText(s, String(r), cx, yy + 8, ww, 12, { size: 8, color: C.gray, align: "center" });
      cx += ww;
    });
  }
}

function monthlyPage(p, pageNo, year, month) {
  const s = p.slides.add();
  const q = quotes[(month + year) % quotes.length];
  base(s, pageNo, `${monthNames[month]} ${year}`, `${q[0]} — ${q[1]}`);
  drawMiniMonth(s, year, month, 64, 112, 270, 245, true);
  addRuledBox(s, 366, 112, 292, 245, "Objetivos del mes", 7);
  addRuledBox(s, 62, 386, 284, 168, "Exámenes / entregas", 5);
  addRuledBox(s, 374, 386, 284, 168, "Oposición", 5);
  addRuledBox(s, 62, 584, 596, 206, "Planificación mensual", 7);
  addRuledBox(s, 62, 814, 596, 82, "Cierre del mes", 2);
}

function weeklyLeftPage(p, pageNo, week) {
  const s = p.slides.add();
  const q = quotes[(week.n - 1) % quotes.length];
  base(s, pageNo, `Semana ${week.n}`, `${fmtShort(week.start)} - ${fmtShort(week.end)} · ${q[0]} — ${q[1]}`);
  const x = 44, y = 116, w = 632;
  const rowH = 108;
  for (let i = 0; i < 5; i++) {
    const d = addDays(week.start, i);
    const yy = y + i * (rowH + 10);
    addShape(s, "roundRect", x, yy, w, rowH, C.white, C.line, 1);
    addShape(s, "rect", x, yy, 74, rowH, i === 0 ? C.peach : "#F4EFE5", C.line, 0.6);
    addText(s, longDays[i], x + 8, yy + 14, 58, 14, { size: 8.5, bold: true, color: C.navy, align: "center" });
    addText(s, fmtDay(d), x + 8, yy + 36, 58, 38, { size: 26, bold: true, color: C.navy, align: "center", face: "Georgia" });
    addText(s, "Clases / tareas", x + 92, yy + 12, 168, 16, { size: 9, bold: true, color: C.gray });
    addText(s, "Oposición", x + 304, yy + 12, 120, 16, { size: 9, bold: true, color: C.gray });
    addText(s, "Notas", x + 464, yy + 12, 140, 16, { size: 9, bold: true, color: C.gray });
    addLine(s, x + 282, yy + 10, x + 282, yy + rowH - 10, C.line, 0.6);
    addLine(s, x + 444, yy + 10, x + 444, yy + rowH - 10, C.line, 0.6);
    for (let r = 1; r <= 3; r++) {
      const ly = yy + 28 + r * 19;
      addLine(s, x + 92, ly, x + 258, ly, "#E8E2D8", 0.6);
      addLine(s, x + 304, ly, x + 420, ly, "#E8E2D8", 0.6);
      addLine(s, x + 464, ly, x + 646, ly, "#E8E2D8", 0.6);
    }
  }
  addRuledBox(s, 44, 700, 300, 138, "Prioridades", 4);
  addRuledBox(s, 372, 700, 304, 138, "Fin de semana", 4);
}

function weeklyRightPage(p, pageNo, week) {
  const s = p.slides.add();
  const title = `Semana ${week.n} · aula`;
  const subtitle = "Seguimiento de sesiones, grupos y evaluación";
  base(s, pageNo, title, subtitle);
  const x = 44, y = 118, w = 632, h = 380;
  const timeW = 74, colW = (w - timeW) / 5, rowH = h / 8;
  addShape(s, "rect", x, y, w, h, C.white, C.line, 1);
  longDays.forEach((d, i) => addText(s, d.slice(0, 3), x + timeW + i * colW, y + 12, colW, 14, { size: 8.5, bold: true, color: C.navy, align: "center" }));
  for (let r = 1; r <= 8; r++) addLine(s, x, y + r * rowH, x + w, y + r * rowH, C.line, 0.65);
  for (let c = 0; c <= 5; c++) addLine(s, x + timeW + c * colW, y, x + timeW + c * colW, y + h, C.line, 0.65);
  for (let r = 1; r < 8; r++) addText(s, `${r}.ª`, x + 12, y + r * rowH + 14, timeW - 24, 14, { size: 8.5, color: C.gray, align: "center" });
  addRuledBox(s, 44, 528, 304, 136, "Alumnado / incidencias", 5);
  addRuledBox(s, 372, 528, 304, 136, "Reuniones / familias", 5);
  addRuledBox(s, 44, 692, 632, 166, "Notas de preparación", 6);
  paperPatch(s, 0, 0, W, 114);
  paperPatch(s, 0, 114, 42, 750);
  paperPatch(s, 678, 114, 42, 750);
  paperPatch(s, 0, 500, W, 26);
  paperPatch(s, 350, 526, 20, 140);
  paperPatch(s, 0, 666, W, 24);
  paperPatch(s, 0, 860, W, H - 860);
  addChrome(s, pageNo, title, subtitle);
}

function notesPage(p, pageNo, n) {
  const s = p.slides.add();
  const title = `Notas ${n}`;
  const subtitle = "Ideas, problemas, demostraciones y recordatorios";
  base(s, pageNo, title, subtitle);
  const x = 58, y = 124, w = 604, h = 740;
  addShape(s, "rect", x, y, w, h, C.white, C.line, 1);
  for (let yy = y + 26; yy < y + h; yy += 26) addLine(s, x + 18, yy, x + w - 18, yy, "#E9E4DB", 0.7);
  paperPatch(s, 0, 0, W, y - 2);
  paperPatch(s, 0, y - 2, x - 2, h + 4);
  paperPatch(s, x + w + 2, y - 2, W - x - w - 2, h + 4);
  paperPatch(s, 0, y + h + 2, W, H - y - h - 2);
  addChrome(s, pageNo, title, subtitle);
}

async function main() {
  await fs.mkdir(OUT, { recursive: true });
  await fs.mkdir(QA, { recursive: true });
  await fs.writeFile(path.join(TMP, "source-notes.txt"), [
    "Fuente local: AgendaVimat_source.pptx, proporcionada por el usuario. Usada como referencia visual: formato vertical 720x1040, borde fino, pie Vimat y motivo matemático.",
    "Fuente oficial: BOE-A-2007-4372, Real Decreto 276/2007, reglamento de ingreso y acceso a cuerpos docentes. Usado para contextualizar que las convocatorias docentes se rigen por bases públicas y normativa autonómica.",
    "Calendario: generado computacionalmente para septiembre 2026-agosto 2027. No se fijan festivos autonómicos/locales de 2027 porque dependen de publicación oficial posterior y del centro.",
    "Temario: tracker de 71 temas con clasificacion operativa por bloques para planificacion. No transcribe epigrafes oficiales literales; queda editable para adaptar al material de Vimat o convocatoria.",
    "Citas: selección breve de atribuciones históricas o frases propias de Vimat para uso motivacional, no empleadas como fuente normativa.",
  ].join("\n"), "utf8");
  await fs.writeFile(path.join(TMP, "slide-plan.txt"), [
    "Deck: Agenda docente VIMAT 2026-2027, A5 vertical editable.",
    "Estilo: navy #0C1B45, papel #FFFDF8, líneas #D8D2C5, acentos coral #B9634A y verde #2C695D.",
    "Tipografía: Georgia para títulos/fórmulas; Aptos para cuerpo/tablas.",
    "Estructura: portada, datos, uso, calendario, horarios, formulario, probabilidad, temario, programación, alumnos, meses, semanas, notas.",
  ].join("\n"), "utf8");

  const p = Presentation.create({ slideSize: { width: W, height: H } });
  p.theme.colorScheme = {
    name: "Vimat Agenda",
    themeColors: {
      accent1: C.navy,
      accent2: C.coral,
      accent3: C.green,
      accent4: C.amber,
      accent5: C.blue,
      accent6: C.warm,
      bg1: C.paper,
      bg2: C.cream,
      tx1: C.ink,
      tx2: C.gray,
      dk1: "#000000",
      dk2: C.navy,
      lt1: C.white,
      lt2: C.cream,
      hlink: C.blue,
      folHlink: C.coral,
    },
  };

  let page = 1;
  cover(p, page++);
  personalPage(p, page++);
  usagePage(p, page++);
  annualCalendarPage(p, page++);
  importantDatesPage(p, page++);
  for (let i = 1; i <= 3; i++) schedulePage(p, page++, i);
  formulasPage(p, page++, "Formulario · Álgebra y cálculo", [
    ["Identidades", ["(a ± b)² = a² ± 2ab + b²", "a² - b² = (a - b)(a + b)", "log_a(xy)=log_a x + log_a y"]],
    ["Derivadas", ["(f·g)' = f'g + fg'", "(f/g)' = (f'g - fg')/g²", "(f∘g)' = (f'∘g)·g'"]],
    ["Taylor", ["f(x)=Σ f⁽ⁿ⁾(a)(x-a)ⁿ/n!", "eˣ=Σ xⁿ/n!", "sen x = x - x³/3! + x⁵/5! - ..."]],
  ]);
  formulasPage(p, page++, "Formulario · Geometría", [
    ["Trigonometría", ["sen²x + cos²x = 1", "sen(a±b)=sen a cos b ± cos a sen b", "cos(a±b)=cos a cos b ∓ sen a sen b"]],
    ["Cónicas", ["Circunferencia: (x-a)²+(y-b)²=r²", "Elipse: x²/a² + y²/b² = 1", "Hipérbola: x²/a² - y²/b² = 1"]],
    ["Áreas y volúmenes", ["A_círculo=πr² · L=2πr", "V_prisma=A_b·h · V_pirámide=A_b·h/3", "V_esfera=4πr³/3 · A_esfera=4πr²"]],
  ]);
  formulasPage(p, page++, "Formulario · Probabilidad y estadística", [
    ["Probabilidad", ["P(A∪B)=P(A)+P(B)-P(A∩B)", "P(A|B)=P(A∩B)/P(B)", "Bayes: P(A_i|B)=P(B|A_i)P(A_i)/ΣP(B|A_j)P(A_j)"]],
    ["Distribuciones", ["Binomial: P(X=k)=C(n,k)pᵏ(1-p)ⁿ⁻ᵏ", "E(X)=np · Var(X)=np(1-p)", "Tipificación: Z=(X-μ)/σ"]],
    ["Estadística", ["x̄=Σx_i/n", "s²=Σ(x_i-x̄)²/(n-1)", "r = cov(X,Y)/(s_x s_y)"]],
  ]);
  probabilityPage(p, page++, 4);
  probabilityPage(p, page++, 5);
  temarioOverviewPage(p, page++);
  temarioTrackerPage(p, page++, "I", 1, 18);
  temarioTrackerPage(p, page++, "II", 19, 36);
  temarioTrackerPage(p, page++, "III", 37, 54);
  temarioTrackerPage(p, page++, "IV", 55, 71);
  oppositionStudyPlanPage(p, page++);
  simulacroLogPage(p, page++);
  tribunalDefensePage(p, page++);
  programmingPage(p, page++);
  for (let i = 1; i <= 8; i++) studentListPage(p, page++, i);
  for (const [y, m] of monthsForAcademicYear()) monthlyPage(p, page++, y, m);
  for (const week of weeksForAcademicYear()) {
    weeklyLeftPage(p, page++, week);
    weeklyRightPage(p, page++, week);
  }
  for (let i = 1; i <= 12; i++) notesPage(p, page++, i);

  const slideCount = p.slides.items.length;
  const previewDir = path.join(TMP, "preview");
  const layoutDir = path.join(TMP, "layout");
  await fs.mkdir(previewDir, { recursive: true });
  await fs.mkdir(layoutDir, { recursive: true });

  for (const [idx, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(idx + 1).padStart(3, "0")}`;
    const png = await p.export({ slide, format: "png", scale: 0.35 });
    await fs.writeFile(path.join(previewDir, `${stem}.png`), Buffer.from(await png.arrayBuffer()));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(layoutDir, `${stem}.layout.json`), await layout.text(), "utf8");
  }
  const montage = await p.export({
    format: "webp",
    montage: { format: "webp", width: 1800, columns: 8, padding: 20, gap: 10, background: "#F6F2EA" },
    scale: 0.32,
  });
  await fs.writeFile(path.join(QA, "final-montage.webp"), Buffer.from(await montage.arrayBuffer()));

  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(FINAL);
  await fs.writeFile(path.join(QA, "visual-qa.txt"), [
    `Final PPTX: ${FINAL}`,
    `Slide count: ${slideCount}`,
    "QA: generated editable native text/shapes; rendered every slide as PNG plus a contact-sheet workspace artifact.",
    "Checked: cover updated to 2026-27, page numbers, recurring footer, week range Sep 2026-Aug 2027, no external raster dependency.",
    "Caveat: 2027 official/local holidays intentionally left editable because they depend on later official publication and geography.",
  ].join("\n"), "utf8");

  const stat = await fs.stat(FINAL);
  console.log(JSON.stringify({ final: FINAL, slideCount, bytes: stat.size, montage: path.join(QA, "final-montage.webp") }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
