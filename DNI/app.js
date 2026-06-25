const DNI_ASPECT_RATIO = 85.6 / 53.98;
const MIN_IMAGE_PIXELS = 700 * 430;
const WARN_IMAGE_PIXELS = 1050 * 650;
const MAX_CANVAS_SIDE = 900;

const state = {
  items: [],
  report: null,
};

const els = {
  dropzone: document.querySelector("#dropzone"),
  fileInput: document.querySelector("#file-input"),
  queueBody: document.querySelector("#queue-body"),
  fileCount: document.querySelector("#file-count"),
  checkButton: document.querySelector("#check-button"),
  clearButton: document.querySelector("#clear-button"),
  progressText: document.querySelector("#progress-text"),
  progressPercent: document.querySelector("#progress-percent"),
  progressBar: document.querySelector("#progress-bar"),
  resultsPanel: document.querySelector("#results-panel"),
  summaryGrid: document.querySelector("#summary-grid"),
  resultBody: document.querySelector("#result-body"),
  downloadHtml: document.querySelector("#download-html"),
  downloadCsv: document.querySelector("#download-csv"),
  downloadJson: document.querySelector("#download-json"),
  sideTemplate: document.querySelector("#side-select-template"),
};

els.fileInput.addEventListener("change", (event) => addFiles(event.target.files));

["dragenter", "dragover"].forEach((name) => {
  els.dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    els.dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((name) => {
  els.dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    els.dropzone.classList.remove("dragover");
  });
});

els.dropzone.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));
els.checkButton.addEventListener("click", runChecks);
els.clearButton.addEventListener("click", clearAll);
els.downloadHtml.addEventListener("click", () => downloadReport("html"));
els.downloadCsv.addEventListener("click", () => downloadReport("csv"));
els.downloadJson.addEventListener("click", () => downloadReport("json"));

function addFiles(fileList) {
  const incoming = Array.from(fileList || []);
  if (!incoming.length) return;

  for (const file of incoming) {
    state.items.push({
      id: crypto.randomUUID(),
      file,
      side: "auto",
      group: normalizeGroupName(file.name),
      status: "Pendiente",
      result: null,
    });
  }

  state.report = null;
  els.resultsPanel.hidden = true;
  els.fileInput.value = "";
  renderQueue();
  setProgress(0, "Listo para comprobar");
}

function clearAll() {
  state.items = [];
  state.report = null;
  els.resultsPanel.hidden = true;
  renderQueue();
  setProgress(0, "Sin documentos");
}

function renderQueue() {
  els.queueBody.textContent = "";
  els.fileCount.textContent = `${state.items.length} ${state.items.length === 1 ? "archivo" : "archivos"}`;
  els.checkButton.disabled = state.items.length === 0;
  els.clearButton.disabled = state.items.length === 0;

  if (!state.items.length) {
    const tr = document.createElement("tr");
    tr.className = "empty-row";
    tr.innerHTML = `<td colspan="5">No hay documentos cargados.</td>`;
    els.queueBody.append(tr);
    return;
  }

  for (const item of state.items) {
    const tr = document.createElement("tr");
    const nameTd = document.createElement("td");
    nameTd.innerHTML = `<span class="file-name"></span><span class="file-sub"></span>`;
    nameTd.querySelector(".file-name").textContent = item.file.name;
    nameTd.querySelector(".file-sub").textContent = item.group;

    const sideTd = document.createElement("td");
    const select = els.sideTemplate.content.firstElementChild.cloneNode(true);
    select.value = item.side;
    select.addEventListener("change", () => {
      item.side = select.value;
      item.result = null;
      state.report = null;
      els.resultsPanel.hidden = true;
    });
    sideTd.append(select);

    const typeTd = document.createElement("td");
    typeTd.textContent = item.file.type || extensionOf(item.file.name).toUpperCase() || "Desconocido";

    const sizeTd = document.createElement("td");
    sizeTd.textContent = formatBytes(item.file.size);

    const statusTd = document.createElement("td");
    statusTd.append(statusBadge(item.status, item.result?.verdict));

    tr.append(nameTd, sideTd, typeTd, sizeTd, statusTd);
    els.queueBody.append(tr);
  }
}

async function runChecks() {
  if (!state.items.length) return;

  els.checkButton.disabled = true;
  state.report = null;
  els.resultsPanel.hidden = true;
  const seenHashes = new Map();

  for (let index = 0; index < state.items.length; index += 1) {
    const item = state.items[index];
    item.status = "Analizando";
    renderQueue();
    setProgress((index / state.items.length) * 100, `Analizando ${index + 1} de ${state.items.length}`);

    try {
      item.result = await analyzeFile(item, seenHashes);
      item.status = "Completado";
    } catch (error) {
      item.result = {
        verdict: "no valido",
        score: 0,
        signals: [`Error de lectura: ${error.message || "no se pudo analizar"}`],
        technical: {},
      };
      item.status = "Completado";
    }

    renderQueue();
    await nextFrame();
  }

  state.report = buildReport();
  renderReport(state.report);
  setProgress(100, "Comprobación finalizada");
  els.checkButton.disabled = false;
}

async function analyzeFile(item, seenHashes) {
  const file = item.file;
  const [header, hash] = await Promise.all([readHeader(file, 64), sha256(file)]);
  const detected = detectFormat(header, file);
  const signals = [];
  const technical = {
    hash,
    detectedFormat: detected.label,
    extension: extensionOf(file.name),
    size: file.size,
    selectedSide: item.side,
  };

  let score = 100;

  if (seenHashes.has(hash)) {
    score -= 18;
    signals.push(`Duplicado de ${seenHashes.get(hash)}`);
  } else {
    seenHashes.set(hash, file.name);
  }

  if (file.size < 25_000) {
    score -= 28;
    signals.push("Archivo demasiado pequeño para un escaneo fiable");
  }

  if (file.size > 35_000_000) {
    score -= 10;
    signals.push("Archivo muy pesado; conviene revisar origen y compresión");
  }

  if (!detected.supported) {
    technical.blockingIssue = true;
    return finishVerdict(25, ["Formato no soportado o firma de archivo no reconocida"], technical);
  }

  if (detected.kind === "pdf") {
    return analyzePdf(file, header, score, signals, technical, item);
  }

  if (detected.kind === "image") {
    return analyzeImage(file, score, signals, technical, item);
  }

  return finishVerdict(35, ["El formato se reconoce, pero el navegador no puede inspeccionar su contenido"], technical);
}

async function analyzePdf(file, header, score, signals, technical, item) {
  const rawText = await file.text();
  const pdfText = await extractPdfText(file);
  const text = `${rawText}\n${pdfText.text || ""}`;
  const trimmed = rawText.slice(0, 12);
  const pageMatches = rawText.match(/\/Type\s*\/Page\b/g) || [];
  const pageCount = pdfText.pages || pageMatches.length || (rawText.includes("/Pages") ? 1 : 0);
  const imageCount = (rawText.match(/\/Subtype\s*\/Image\b/g) || []).length;
  const dniCheck = extractDniCandidates(text);
  const dniNumbers = dniCheck.valid;
  const invalidDniNumbers = dniCheck.invalid;
  const hasDniNumber = dniNumbers.length > 0;
  const hasMrz = /IDESP|DNI|ESP[A-Z<]{2,}|\d{8}[A-Z]<{1,}/i.test(text);
  const hasScript = /\/JavaScript|\/JS|\/OpenAction|\/AA\b/i.test(text);
  const encrypted = /\/Encrypt\b/i.test(text);
  const nameSide = detectSideFromName(file.name);

  technical.pages = pageCount;
  technical.embeddedImages = imageCount;
  technical.hasReadableText = hasDniNumber || hasMrz;
  technical.dniNumbers = dniNumbers;
  technical.invalidDniNumbers = invalidDniNumbers;
  technical.ignoredDniLikeCandidates = dniCheck.ignored;
  technical.identityNumberStatus = invalidDniNumbers.length ? "invalid" : hasDniNumber ? "valid" : hasMrz ? "mrz-detected" : "unknown";
  technical.pdfHeader = trimmed;
  technical.pdfTextExtracted = Boolean(pdfText.text);

  if (!startsWithBytes(header, [0x25, 0x50, 0x44, 0x46])) {
    score -= 45;
    signals.push("El PDF no empieza con la firma estándar %PDF");
  }

  if (encrypted) {
    score -= 35;
    technical.blockingIssue = true;
    signals.push("PDF cifrado; no se puede inspeccionar el contenido offline");
  }

  if (hasScript) {
    score -= 25;
    signals.push("PDF con acciones o JavaScript incrustado");
  }

  if (imageCount >= 2) {
    technical.containsFrontAndBackCandidate = true;
    signals.push(`PDF con ${imageCount} imÃ¡genes incrustadas; se considera posible frente y reverso en el mismo archivo`);
  }

  if (pageCount === 0) {
    score -= 25;
    signals.push("No se ha podido estimar el número de páginas");
  } else if (pageCount === 1) {
    score -= 10;
    signals.push("PDF de una sola página; puede faltar frente o reverso");
  } else if (pageCount >= 2) {
    technical.containsFrontAndBackCandidate = true;
    signals.push(`PDF con ${pageCount} páginas detectadas; se considera posible frente y reverso en el mismo archivo`);
  }

  if (invalidDniNumbers.length) {
    score = Math.min(score - 60, 25);
    technical.blockingIssue = true;
    technical.criticalIssue = true;
    signals.push(`DNI con letra de control incorrecta: ${invalidDniNumbers.join(", ")}`);
  } else if (hasDniNumber) {
    score += 8;
    signals.push("DNI con letra de control válida detectado");
  }

  if (dniCheck.ignored.length) {
    signals.push(`Candidatos tipo DNI ignorados por parecer fechas/metadatos: ${dniCheck.ignored.join(", ")}`);
  }

  if (hasDniNumber || hasMrz) {
    score += 5;
    signals.push("Texto compatible con DNI/MRZ detectado");
  } else {
    signals.push("PDF sin texto legible de DNI; puede ser un escaneo sin OCR");
  }

  technical.detectedSide = resolveDetectedSide(item.side, inferPdfSide({ pageCount, imageCount, hasMrz, hasDniNumber, nameSide }));
  signals.push(`Cara detectada: ${sideLabel(technical.detectedSide)}`);

  return finishVerdict(score, signals, technical, "El contenido visual de un PDF escaneado requiere OCR o renderizado PDF local.");
}

async function analyzeImage(file, score, signals, technical, item) {
  let bitmap;
  try {
    bitmap = await createImageBitmap(file);
  } catch {
    return finishVerdict(20, ["El navegador no puede decodificar esta imagen"], technical);
  }

  const width = bitmap.width;
  const height = bitmap.height;
  const pixels = width * height;
  const aspect = Math.max(width, height) / Math.min(width, height);
  const nameSide = detectSideFromName(file.name);
  const aspectCombinedLayout = detectCombinedSidesLayout(aspect, pixels);
  technical.width = width;
  technical.height = height;
  technical.aspect = Number(aspect.toFixed(3));

  if (pixels < MIN_IMAGE_PIXELS) {
    score -= 35;
    signals.push(`Resolución baja (${width} x ${height})`);
  } else if (pixels < WARN_IMAGE_PIXELS) {
    score -= 12;
    signals.push(`Resolución justa (${width} x ${height})`);
  } else {
    signals.push(`Resolución suficiente (${width} x ${height})`);
  }

  const sample = sampleImage(bitmap);
  Object.assign(technical, sample.technical);
  bitmap.close?.();

  const regionCombinedLayout = pixels >= MIN_IMAGE_PIXELS && sample.technical.cardRegionCount >= 2
    ? `${sample.technical.cardRegionCount} zonas tipo DNI detectadas`
    : "";
  const combinedSidesLayout = regionCombinedLayout || (pixels >= MIN_IMAGE_PIXELS ? aspectCombinedLayout : "") || (nameSide === "both" ? "nombre de archivo" : "");
  technical.containsFrontAndBackCandidate = Boolean(combinedSidesLayout);
  if (combinedSidesLayout) {
    technical.combinedSidesLayout = combinedSidesLayout;
  }

  if (combinedSidesLayout) {
    score += 8;
    signals.push(`Posible frente y reverso en la misma imagen (${combinedSidesLayout})`);
  } else {
    const aspectDelta = Math.abs(aspect - DNI_ASPECT_RATIO);
    if (aspectDelta > 0.28) {
      score -= 22;
      signals.push("Proporción alejada del formato físico del DNI");
    } else if (aspectDelta > 0.16) {
      score -= 8;
      signals.push("Proporción compatible con margen de escaneo amplio");
    } else {
      signals.push("Proporción compatible con DNI");
    }
  }

  if (sample.technical.contrast < 28) {
    score -= combinedSidesLayout ? 8 : 18;
    signals.push("Contraste bajo; posible escaneo borroso o lavado");
  } else if (sample.technical.contrast < 42) {
    score -= combinedSidesLayout ? 3 : 8;
    signals.push("Contraste moderado");
  }

  if (sample.technical.clippedRatio > 0.42) {
    score -= combinedSidesLayout ? 4 : 12;
    signals.push("Exceso de zonas quemadas u oscuras");
  }

  const edgeThreshold = sample.technical.cardRegionCount >= 1 ? 0.02 : 0.035;
  if (sample.technical.edgeDensity < edgeThreshold) {
    score -= combinedSidesLayout ? 6 : 18;
    signals.push("Poco detalle fino; posible imagen desenfocada");
  } else {
    signals.push("Detalle visual suficiente");
  }

  if (sample.technical.saturation < 9) {
    score -= combinedSidesLayout ? 2 : 8;
    signals.push("Imagen casi monocroma; revisar si procede del original");
  }

  technical.detectedSide = resolveDetectedSide(item.side, inferImageSide({
    combinedSidesLayout,
    nameSide,
    mrzLikeScore: sample.technical.mrzLikeScore,
    photoLikeScore: sample.technical.photoLikeScore,
  }));
  technical.identityNumberStatus = "unknown";
  signals.push(`Cara detectada: ${sideLabel(technical.detectedSide)}`);

  if (!signals.length) {
    signals.push("No se han detectado incidencias técnicas relevantes");
  }

  return finishVerdict(score, signals, technical);
}

function sampleImage(bitmap) {
  const scale = Math.min(1, MAX_CANVAS_SIDE / Math.max(bitmap.width, bitmap.height));
  const width = Math.max(1, Math.round(bitmap.width * scale));
  const height = Math.max(1, Math.round(bitmap.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(bitmap, 0, 0, width, height);
  const { data } = ctx.getImageData(0, 0, width, height);

  let sum = 0;
  let sumSq = 0;
  let clipped = 0;
  let saturationSum = 0;
  const grays = new Float32Array(width * height);

  for (let i = 0, p = 0; i < data.length; i += 4, p += 1) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    const gray = 0.299 * r + 0.587 * g + 0.114 * b;
    grays[p] = gray;
    sum += gray;
    sumSq += gray * gray;
    if (gray < 8 || gray > 247) clipped += 1;
    saturationSum += Math.max(r, g, b) - Math.min(r, g, b);
  }

  const count = width * height;
  const mean = sum / count;
  const variance = Math.max(0, sumSq / count - mean * mean);
  const contrast = Math.sqrt(variance);

  let edgeHits = 0;
  let edgeSamples = 0;
  const step = Math.max(1, Math.floor(Math.min(width, height) / 240));
  for (let y = step; y < height - step; y += step) {
    for (let x = step; x < width - step; x += step) {
      const center = grays[y * width + x];
      const dx = Math.abs(center - grays[y * width + x + step]);
      const dy = Math.abs(center - grays[(y + step) * width + x]);
      if (dx + dy > 42) edgeHits += 1;
      edgeSamples += 1;
    }
  }

  return {
    technical: {
      brightness: Number(mean.toFixed(1)),
      contrast: Number(contrast.toFixed(1)),
      clippedRatio: Number((clipped / count).toFixed(3)),
      saturation: Number((saturationSum / count).toFixed(1)),
      edgeDensity: Number((edgeHits / Math.max(1, edgeSamples)).toFixed(3)),
      ...detectVisualDocumentSignals(data, width, height, grays),
    },
  };
}

function finishVerdict(rawScore, signals, technical, limitation = "") {
  const score = Math.max(0, Math.min(100, Math.round(rawScore)));
  let verdict = "valido";
  if (score < 58) verdict = "no valido";
  else if (score < 78) verdict = "revisar";

  if (limitation) {
    technical.limitation = limitation;
  }

  return {
    verdict,
    score,
    signals: signals.length ? signals : ["Sin incidencias técnicas relevantes"],
    technical,
  };
}

function buildReport() {
  const groups = new Map();
  for (const item of state.items) {
    const key = item.group || item.file.name;
    if (!groups.has(key)) {
      groups.set(key, { name: key, items: [] });
    }
    groups.get(key).items.push(item);
  }

  const documents = Array.from(groups.values()).map((group) => {
    const scores = group.items.map((item) => item.result?.score ?? 0);
    const minScore = Math.min(...scores);
    const resolvedSides = group.items.map((item) => item.result?.technical?.detectedSide || resolveDetectedSide(item.side, "unknown"));
    const hasFront = resolvedSides.includes("front");
    const hasBack = resolvedSides.includes("back");
    const hasBothSelected = group.items.some((item) => item.side === "both");
    const hasBothDetected = resolvedSides.includes("both") || group.items.some((item) => item.result?.technical?.containsFrontAndBackCandidate);
    const hasCompleteDocument = (hasFront && hasBack) || hasBothSelected || hasBothDetected;
    const hasBlockingIssue = group.items.some((item) => item.result?.technical?.blockingIssue);
    const hasCriticalIssue = group.items.some((item) => item.result?.technical?.criticalIssue);
    const hasVerifiedIdentity = group.items.some((item) => {
      const status = item.result?.technical?.identityNumberStatus;
      return status === "valid" || status === "mrz-detected";
    });
    const hasInvalidIdentity = group.items.some((item) => item.result?.technical?.identityNumberStatus === "invalid");
    const hasPdfScanWithoutText = group.items.some((item) => {
      const technical = item.result?.technical || {};
      return technical.detectedFormat === "PDF" && technical.containsFrontAndBackCandidate && technical.identityNumberStatus === "unknown";
    });
    const hasLowQualityUnverifiedImage = group.items.some((item) => {
      const technical = item.result?.technical || {};
      const pixels = (technical.width || 0) * (technical.height || 0);
      return technical.identityNumberStatus === "unknown" && technical.detectedFormat !== "PDF" && pixels > 0 && pixels < MIN_IMAGE_PIXELS;
    });
    const signals = group.items.flatMap((item) => item.result?.signals.map((signal) => `${item.file.name}: ${signal}`) || []);
    let score = minScore;

    if (hasBothSelected) {
      signals.unshift("El archivo se ha marcado como frente y reverso en el mismo documento");
    } else if (hasBothDetected) {
      signals.unshift("La aplicación ha identificado frente y reverso dentro del mismo documento");
    }

    if (!hasCompleteDocument) {
      score = Math.min(score, 72);
      signals.unshift("No se ha identificado pareja frente/reverso completa");
    }

    if (hasCompleteDocument && hasVerifiedIdentity && !hasBlockingIssue) {
      score = Math.min(100, score + 6);
    } else if (hasCompleteDocument && !hasVerifiedIdentity && !hasBlockingIssue) {
      if (hasPdfScanWithoutText) {
        score = Math.max(62, Math.min(score, 68));
        signals.unshift("PDF completo, pero sin OCR suficiente para validar número/letra; queda a revisión");
      } else {
        score = Math.min(score, 58);
        signals.unshift("No se ha podido validar número/letra del DNI; no se marca como correcto");
      }
    }

    if (!hasVerifiedIdentity && !hasBlockingIssue && hasLowQualityUnverifiedImage) {
      score = Math.min(score, 45);
      signals.unshift("Imagen con resolución insuficiente y sin número/letra verificable");
    }

    if (hasInvalidIdentity || hasCriticalIssue) {
      score = Math.min(score, 25);
      signals.unshift("Incoherencia crítica en número/letra del DNI");
    } else if (hasBlockingIssue) {
      score = Math.min(score, 45);
    }

    const verdict = score < 58 ? "no valido" : score < 78 ? "revisar" : "valido";

    return {
      name: group.name,
      verdict,
      score,
      files: group.items.map((item) => item.file.name),
      signals,
    };
  });

  const summary = {
    totalFiles: state.items.length,
    totalDocuments: documents.length,
    valid: documents.filter((doc) => doc.verdict === "valido").length,
    review: documents.filter((doc) => doc.verdict === "revisar").length,
    invalid: documents.filter((doc) => doc.verdict === "no valido").length,
    generatedAt: new Date().toISOString(),
  };

  return { summary, documents, files: state.items.map(fileReportRow) };
}

function fileReportRow(item) {
  return {
    name: item.file.name,
    group: item.group,
    side: item.side,
    detectedSide: item.result?.technical?.detectedSide || "unknown",
    size: item.file.size,
    verdict: item.result?.verdict || "sin resultado",
    score: item.result?.score ?? 0,
    signals: item.result?.signals || [],
    technical: item.result?.technical || {},
  };
}

function renderReport(report) {
  els.resultsPanel.hidden = false;
  els.summaryGrid.textContent = "";
  els.resultBody.textContent = "";

  const metrics = [
    ["Archivos", report.summary.totalFiles],
    ["Documentos", report.summary.totalDocuments],
    ["Válidos", report.summary.valid],
    ["A revisar", report.summary.review + report.summary.invalid],
  ];

  for (const [label, value] of metrics) {
    const div = document.createElement("div");
    div.className = "metric";
    div.innerHTML = `<strong></strong><span></span>`;
    div.querySelector("strong").textContent = value;
    div.querySelector("span").textContent = label;
    els.summaryGrid.append(div);
  }

  for (const doc of report.documents) {
    const tr = document.createElement("tr");
    const nameTd = document.createElement("td");
    nameTd.innerHTML = `<span class="file-name"></span><span class="file-sub"></span>`;
    nameTd.querySelector(".file-name").textContent = doc.name;
    nameTd.querySelector(".file-sub").textContent = doc.files.join(", ");

    const verdictTd = document.createElement("td");
    verdictTd.append(statusBadge("Completado", doc.verdict));

    const scoreTd = document.createElement("td");
    scoreTd.textContent = `${doc.score}/100`;

    const signalsTd = document.createElement("td");
    const list = document.createElement("ul");
    list.className = "signal-list";
    for (const signal of doc.signals.slice(0, 6)) {
      const li = document.createElement("li");
      li.textContent = signal;
      list.append(li);
    }
    if (doc.signals.length > 6) {
      const li = document.createElement("li");
      li.textContent = `${doc.signals.length - 6} señales adicionales en el informe descargable`;
      list.append(li);
    }
    signalsTd.append(list);

    tr.append(nameTd, verdictTd, scoreTd, signalsTd);
    els.resultBody.append(tr);
  }
}

function statusBadge(status, verdict) {
  const span = document.createElement("span");
  span.className = "badge neutral";
  span.textContent = status;

  if (verdict === "valido") {
    span.className = "badge ok";
    span.textContent = "Correcto";
  } else if (verdict === "revisar") {
    span.className = "badge warn";
    span.textContent = "Revisar";
  } else if (verdict === "no valido") {
    span.className = "badge bad";
    span.textContent = "No válido";
  } else if (status === "Analizando") {
    span.className = "badge warn";
  }

  return span;
}

function downloadReport(type) {
  if (!state.report) return;
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  if (type === "json") {
    downloadBlob(JSON.stringify(state.report, null, 2), `informe-dni-${timestamp}.json`, "application/json");
  } else if (type === "csv") {
    downloadBlob(toCsv(state.report.files), `informe-dni-${timestamp}.csv`, "text/csv;charset=utf-8");
  } else {
    downloadBlob(toHtmlReport(state.report), `informe-dni-${timestamp}.html`, "text/html;charset=utf-8");
  }
}

function toCsv(rows) {
  const headers = ["archivo", "documento", "seleccion", "cara_detectada", "tamano", "veredicto", "puntuacion", "senales"];
  const lines = [headers.join(",")];
  for (const row of rows) {
    lines.push([
      row.name,
      row.group,
      row.side,
      row.detectedSide,
      row.size,
      row.verdict,
      row.score,
      row.signals.join(" | "),
    ].map(csvCell).join(","));
  }
  return lines.join("\n");
}

function toHtmlReport(report) {
  const rows = report.documents.map((doc) => `
    <tr>
      <td>${escapeHtml(doc.name)}</td>
      <td>${escapeHtml(labelVerdict(doc.verdict))}</td>
      <td>${doc.score}/100</td>
      <td><ul>${doc.signals.map((signal) => `<li>${escapeHtml(signal)}</li>`).join("")}</ul></td>
    </tr>`).join("");

  return `<!doctype html>
<html lang="es">
<meta charset="utf-8">
<title>Informe de verificación DNI</title>
<style>
body{font-family:Arial,sans-serif;margin:28px;color:#182026}
table{width:100%;border-collapse:collapse}th,td{border:1px solid #d9e0e5;padding:9px;text-align:left;vertical-align:top}
th{background:#eef3f6}.meta{color:#62717c}
</style>
<h1>Informe de verificación DNI</h1>
<p class="meta">Generado: ${escapeHtml(new Date(report.summary.generatedAt).toLocaleString("es-ES"))}</p>
<p>Archivos: ${report.summary.totalFiles} | Documentos: ${report.summary.totalDocuments} | Correctos: ${report.summary.valid} | Revisar: ${report.summary.review} | No válidos: ${report.summary.invalid}</p>
<table>
  <thead><tr><th>Documento</th><th>Veredicto</th><th>Puntuación</th><th>Señales</th></tr></thead>
  <tbody>${rows}</tbody>
</table>
<p class="meta">Resultado offline basado en reglas técnicas. No sustituye una validación oficial de identidad.</p>
</html>`;
}

function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.append(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function detectSideFromName(name) {
  const clean = stripDiacritics(name).toLowerCase();
  if (/\b(completo|ambas|ambos|dos caras|frente y reverso|anverso y reverso|front back|front and back|both)\b/.test(clean)) return "both";
  if (/\b(frente|anverso|delante|front|cara-a|cara_a)\b/.test(clean)) return "front";
  if (/\b(reverso|dorso|detras|trasera|posterior|back|cara-b|cara_b)\b/.test(clean)) return "back";
  return "unknown";
}

function normalizeGroupName(name) {
  const base = name.replace(/\.[^.]+$/, "");
  return stripDiacritics(base)
    .toLowerCase()
    .replace(/\b(completo|ambas|ambos|dos caras|frente y reverso|anverso y reverso|front back|front and back|both|frente|anverso|delante|front|reverso|dorso|detras|trasera|posterior|back|cara[-_\s]?[ab])\b/g, "")
    .replace(/[_\-.]+/g, " ")
    .replace(/\s+/g, " ")
    .trim() || base;
}

function detectFormat(bytes, file) {
  const ext = extensionOf(file.name);
  if (startsWithBytes(bytes, [0x25, 0x50, 0x44, 0x46])) return { kind: "pdf", label: "PDF", supported: true };
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47])) return { kind: "image", label: "PNG", supported: true };
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) return { kind: "image", label: "JPEG", supported: true };
  if (startsWithBytes(bytes, [0x47, 0x49, 0x46, 0x38])) return { kind: "image", label: "GIF", supported: true };
  if (startsWithBytes(bytes, [0x42, 0x4d])) return { kind: "image", label: "BMP", supported: true };
  if (String.fromCharCode(...bytes.slice(8, 12)) === "WEBP") return { kind: "image", label: "WEBP", supported: true };
  if (["tif", "tiff", "heic", "heif"].includes(ext)) return { kind: "other", label: ext.toUpperCase(), supported: true };
  if (file.type.startsWith("image/")) return { kind: "image", label: file.type, supported: true };
  return { kind: "unknown", label: ext.toUpperCase() || "Desconocido", supported: false };
}

async function readHeader(file, length) {
  return new Uint8Array(await file.slice(0, length).arrayBuffer());
}

async function sha256(file) {
  const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
  return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function startsWithBytes(bytes, signature) {
  return signature.every((byte, index) => bytes[index] === byte);
}

function extensionOf(name) {
  const match = /\.([^.]+)$/.exec(name);
  return match ? match[1].toLowerCase() : "";
}

function stripDiacritics(value) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB"];
  let size = bytes / 1024;
  let unit = units.shift();
  while (size >= 1024 && units.length) {
    size /= 1024;
    unit = units.shift();
  }
  return `${size.toFixed(size >= 10 ? 1 : 2)} ${unit}`;
}

function setProgress(percent, text) {
  const rounded = Math.round(percent);
  els.progressBar.value = rounded;
  els.progressPercent.textContent = `${rounded}%`;
  els.progressText.textContent = text;
}

function nextFrame() {
  return new Promise((resolve) => requestAnimationFrame(resolve));
}

function csvCell(value) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function labelVerdict(verdict) {
  if (verdict === "valido") return "Correcto";
  if (verdict === "revisar") return "Revisar";
  return "No válido";
}

function detectCombinedSidesLayout(aspect, pixels) {
  if (pixels < MIN_IMAGE_PIXELS) return "";
  if (aspect >= 2.65 && aspect <= 3.55) return "dos caras en horizontal";
  if (aspect >= 1.12 && aspect <= 1.34) return "dos caras en vertical";
  if (aspect >= 1.35 && aspect <= 1.55) return "escaneo tipo A4 con posible doble cara";
  return "";
}

async function extractPdfText(file) {
  try {
    if (!window.pdfjsLibPromise) {
      window.pdfjsLibPromise = import("./vendor/pdf.min.mjs").then((module) => {
        module.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.mjs";
        return module;
      });
    }
    const pdfjsLib = await window.pdfjsLibPromise;
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const chunks = [];
    const maxPages = Math.min(pdf.numPages, 6);
    for (let pageNumber = 1; pageNumber <= maxPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      chunks.push(content.items.map((item) => item.str || "").join(" "));
    }
    return { text: chunks.join("\n"), pages: pdf.numPages };
  } catch {
    return { text: "", pages: 0 };
  }
}

function inferPdfSide({ pageCount, imageCount, hasMrz, hasDniNumber, nameSide }) {
  if (pageCount >= 2 || imageCount >= 2 || nameSide === "both") return "both";
  if (nameSide !== "unknown") return nameSide;
  if (hasMrz) return "back";
  if (hasDniNumber) return "front";
  return "unknown";
}

function inferImageSide({ combinedSidesLayout, nameSide, mrzLikeScore, photoLikeScore }) {
  if (combinedSidesLayout || nameSide === "both") return "both";
  if (nameSide !== "unknown") return nameSide;
  if (mrzLikeScore > 0.16) return "back";
  if (photoLikeScore > 0.2) return "front";
  return "unknown";
}

function resolveDetectedSide(selectedSide, inferredSide) {
  return selectedSide && selectedSide !== "auto" ? selectedSide : inferredSide;
}

function sideLabel(side) {
  if (side === "front") return "solo frente";
  if (side === "back") return "solo reverso";
  if (side === "both") return "frente y reverso";
  return "no determinada";
}

function detectVisualDocumentSignals(data, width, height, grays) {
  const regions = detectCardLikeRegions(data, width, height);
  return {
    cardRegionCount: regions.length,
    cardRegionLayout: regions.length >= 2 ? describeRegionLayout(regions) : "",
    mrzLikeScore: Number(detectMrzLikeRows(grays, width, height).toFixed(3)),
    photoLikeScore: Number(detectPhotoLikeArea(data, width, height).toFixed(3)),
  };
}

function detectCardLikeRegions(data, width, height) {
  const step = Math.max(1, Math.ceil(Math.max(width, height) / 280));
  const gridWidth = Math.floor(width / step);
  const gridHeight = Math.floor(height / step);
  const total = gridWidth * gridHeight;
  const background = averageCornerColor(data, width, height);
  const bgGray = 0.299 * background.r + 0.587 * background.g + 0.114 * background.b;
  const mask = new Uint8Array(total);
  const seen = new Uint8Array(total);

  for (let gy = 0; gy < gridHeight; gy += 1) {
    for (let gx = 0; gx < gridWidth; gx += 1) {
      const x = Math.min(width - 1, gx * step);
      const y = Math.min(height - 1, gy * step);
      const idx = (y * width + x) * 4;
      const r = data[idx];
      const g = data[idx + 1];
      const b = data[idx + 2];
      const gray = 0.299 * r + 0.587 * g + 0.114 * b;
      const colorDelta = Math.abs(r - background.r) + Math.abs(g - background.g) + Math.abs(b - background.b);
      if ((colorDelta > 46 || gray < bgGray - 20) && gray < 248) {
        mask[gy * gridWidth + gx] = 1;
      }
    }
  }

  const regions = [];
  const stack = [];
  for (let i = 0; i < total; i += 1) {
    if (!mask[i] || seen[i]) continue;
    seen[i] = 1;
    stack.push(i);
    let area = 0;
    let minX = gridWidth;
    let maxX = 0;
    let minY = gridHeight;
    let maxY = 0;

    while (stack.length) {
      const current = stack.pop();
      const x = current % gridWidth;
      const y = Math.floor(current / gridWidth);
      area += 1;
      minX = Math.min(minX, x);
      maxX = Math.max(maxX, x);
      minY = Math.min(minY, y);
      maxY = Math.max(maxY, y);

      const neighbors = [current - 1, current + 1, current - gridWidth, current + gridWidth];
      for (const next of neighbors) {
        if (next < 0 || next >= total || seen[next] || !mask[next]) continue;
        const nx = next % gridWidth;
        if (Math.abs(nx - x) > 1) continue;
        seen[next] = 1;
        stack.push(next);
      }
    }

    const boxWidth = maxX - minX + 1;
    const boxHeight = maxY - minY + 1;
    const areaRatio = area / total;
    const boxAreaRatio = (boxWidth * boxHeight) / total;
    const aspect = Math.max(boxWidth, boxHeight) / Math.max(1, Math.min(boxWidth, boxHeight));

    if (areaRatio > 0.012 && boxAreaRatio > 0.03 && boxWidth > 16 && boxHeight > 10 && aspect >= 1.15 && aspect <= 2.45) {
      regions.push({ area, minX, maxX, minY, maxY, width: boxWidth, height: boxHeight });
    }
  }

  return regions.sort((a, b) => b.area - a.area).slice(0, 4);
}

function averageCornerColor(data, width, height) {
  const points = [
    [0.05, 0.05],
    [0.95, 0.05],
    [0.05, 0.95],
    [0.95, 0.95],
  ];
  let r = 0;
  let g = 0;
  let b = 0;
  for (const [px, py] of points) {
    const x = Math.min(width - 1, Math.max(0, Math.round(width * px)));
    const y = Math.min(height - 1, Math.max(0, Math.round(height * py)));
    const idx = (y * width + x) * 4;
    r += data[idx];
    g += data[idx + 1];
    b += data[idx + 2];
  }
  return { r: r / points.length, g: g / points.length, b: b / points.length };
}

function describeRegionLayout(regions) {
  const first = regions[0];
  const second = regions[1];
  const horizontalGap = Math.abs((first.minX + first.maxX) / 2 - (second.minX + second.maxX) / 2);
  const verticalGap = Math.abs((first.minY + first.maxY) / 2 - (second.minY + second.maxY) / 2);
  return horizontalGap >= verticalGap ? "dos caras separadas en horizontal" : "dos caras separadas en vertical";
}

function detectMrzLikeRows(grays, width, height) {
  let hitRows = 0;
  let sampledRows = 0;
  const startY = Math.floor(height * 0.55);
  for (let y = startY; y < height; y += 2) {
    let dark = 0;
    for (let x = Math.floor(width * 0.08); x < width * 0.92; x += 3) {
      if (grays[y * width + x] < 95) dark += 1;
    }
    const samples = Math.max(1, Math.floor((width * 0.84) / 3));
    if (dark / samples > 0.18) hitRows += 1;
    sampledRows += 1;
  }
  return hitRows / Math.max(1, sampledRows);
}

function detectPhotoLikeArea(data, width, height) {
  let content = 0;
  let samples = 0;
  const xStart = Math.floor(width * 0.05);
  const xEnd = Math.floor(width * 0.45);
  const yStart = Math.floor(height * 0.18);
  const yEnd = Math.floor(height * 0.82);
  const step = Math.max(2, Math.floor(Math.min(width, height) / 160));
  for (let y = yStart; y < yEnd; y += step) {
    for (let x = xStart; x < xEnd; x += step) {
      const idx = (y * width + x) * 4;
      const r = data[idx];
      const g = data[idx + 1];
      const b = data[idx + 2];
      const gray = 0.299 * r + 0.587 * g + 0.114 * b;
      const saturation = Math.max(r, g, b) - Math.min(r, g, b);
      if (gray > 35 && gray < 235 && saturation > 12) content += 1;
      samples += 1;
    }
  }
  return content / Math.max(1, samples);
}

function extractDniCandidates(text) {
  const clean = text.toUpperCase();
  const candidates = new Map();
  const compact = /\b\d{8}[A-Z]\b/g;
  let match = compact.exec(clean);
  while (match) {
    candidates.set(match[0], getDniCandidateContext(clean, match.index, match[0].length));
    match = compact.exec(clean);
  }

  const separated = /(^|[^\d])(\d[\d\s.\-]{6,18}\d)\s*[-.\s]?\s*([A-Z])(?=$|[^A-Z])/g;
  match = separated.exec(clean);
  while (match) {
    const digits = match[2].replace(/\D/g, "");
    if (digits.length === 8) {
      candidates.set(`${digits}${match[3]}`, getDniCandidateContext(clean, match.index, match[0].length));
    }
    match = separated.exec(clean);
  }

  const valid = [];
  const invalid = [];
  const ignored = [];
  for (const [candidate, context] of candidates) {
    const validLetter = isValidDniLetter(candidate);
    if (validLetter) {
      valid.push(candidate);
    } else if (isStrongDniContext(context) && !looksLikeDateCandidate(candidate, context)) {
      invalid.push(candidate);
    } else {
      ignored.push(candidate);
    }
  }

  return {
    valid: Array.from(new Set(valid)),
    invalid: Array.from(new Set(invalid)),
    ignored: Array.from(new Set(ignored)),
  };
}

function getDniCandidateContext(text, index, length) {
  return text.slice(Math.max(0, index - 45), Math.min(text.length, index + length + 45));
}

function isStrongDniContext(context) {
  return /\b(DNI|NIF|NUMERO|N[ºO]\.?|DOCUMENTO|IDENTIDAD|IDESP|ESP)\b/.test(stripDiacritics(context));
}

function looksLikeDateCandidate(candidate, context) {
  const digits = candidate.slice(0, 8);
  const year = Number(digits.slice(0, 4));
  const month = Number(digits.slice(4, 6));
  const day = Number(digits.slice(6, 8));
  const yyyymmdd = year >= 1900 && year <= 2099 && month >= 1 && month <= 12 && day >= 1 && day <= 31;
  const timestampContext = /D:?\d{8}|CREATIONDATE|MODDATE|FECHA|DATE|TIME|TIMESTAMP/.test(context);
  return yyyymmdd && (candidate.endsWith("T") || timestampContext);
}

function isValidDniLetter(value) {
  const letters = "TRWAGMYFPDXBNJZSQVHLCKE";
  const number = Number(value.slice(0, 8));
  return letters[number % 23] === value.slice(8).toUpperCase();
}
