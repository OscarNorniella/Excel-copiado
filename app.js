/**
 * SuperCopy Pro - Complemento para Excel
 * Exporta selecciones de Excel con formato adaptado para WhatsApp, Imagen y Correo.
 */

let appInitialized = false;
let toastTimeout;

function initApp() {
  if (appInitialized) return;
  appInitialized = true;
  setupEventListeners();
  loadCurrentSelection();
}

document.addEventListener("DOMContentLoaded", initApp);
if (document.readyState === "complete" || document.readyState === "interactive") initApp();

function setupEventListeners() {
  document.getElementById("btn-refresh")?.addEventListener("click", loadCurrentSelection);
  document.getElementById("btn-copy-chat")?.addEventListener("click", () => handleCopy("chat"));
  document.getElementById("btn-copy-image")?.addEventListener("click", () => handleCopy("image"));
  document.getElementById("btn-copy-email")?.addEventListener("click", () => handleCopy("email"));

  if (typeof Office !== "undefined") {
    Office.onReady(() => {
      loadCurrentSelection();
      if (typeof Excel !== "undefined") {
        Excel.run(async context => {
          context.workbook.onSelectionChanged.add(onSelectionChange);
          await context.sync();
        }).catch(() => {});
      }
    });
  }
}

async function onSelectionChange() {
  await loadCurrentSelection();
}

async function getSelectionData() {
  try {
    return await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "text", "values", "rowCount", "columnCount"]);
      await context.sync();
      if (range.rowCount === 0 || range.columnCount === 0) return null;

      const headerCells = [];
      for (let c = 0; c < range.columnCount; c++) {
        const cell = range.getCell(0, c);
        cell.load(["format/fill/color", "format/font/color", "format/font/bold"]);
        headerCells.push(cell);
      }
      await context.sync();

      const matrix = range.text;
      const headers = matrix[0].map(value => String(value || ""));
      const rows = matrix.slice(1).map(row => row.map(value => String(value || "")));
      const headerStyles = headerCells.map(cell => {
        const rawBg = cell.format.fill.color;
        const rawFont = cell.format.font.color;
        const bg = rawBg && rawBg !== "" && rawBg.toLowerCase() !== "#00000000" ? rawBg : "#0f172a";
        return {
          bg,
          color: rawFont && rawFont !== "" ? rawFont : (isLightColor(bg) ? "#0f172a" : "#ffffff"),
          bold: cell.format.font.bold
        };
      });

      return {
        address: range.address,
        rowCount: range.rowCount,
        colCount: range.columnCount,
        headers,
        headerStyles,
        rows,
        rawValues: range.values
      };
    });
  } catch (error) {
    console.error("Error al leer la selección:", error);
    return null;
  }
}

async function loadCurrentSelection() {
  const addressElem = document.getElementById("range-address");
  const dimElem = document.getElementById("range-dimensions");
  if (!addressElem || !dimElem) return;

  const data = await getSelectionData();
  if (!data) {
    addressElem.textContent = "Sin selección";
    dimElem.textContent = "0 filas × 0 columnas";
    return;
  }

  addressElem.textContent = data.address.split("!").pop() || data.address;
  dimElem.textContent = `${data.rowCount} fila${data.rowCount > 1 ? "s" : ""} × ${data.colCount} col${data.colCount > 1 ? "s" : ""}`;
  renderPreview(data);
}

async function handleCopy(mode) {
  const data = await getSelectionData();
  if (!data || data.rowCount === 0) {
    showToast("Selecciona un rango con datos en Excel primero");
    return;
  }

  try {
    if (mode === "chat") {
      await copyForChat(data);
      showToast("Copiado para WhatsApp / Chat");
    } else if (mode === "image") {
      await copyAsImage(data);
      showToast("Copiado como Imagen al portapapeles");
    } else if (mode === "email") {
      await copyForEmail(data);
      showToast("Copiado como Tabla para Correo / HTML");
    }
  } catch (error) {
    console.error("Error al copiar:", error);
    showToast(error?.message || "Error al copiar. Intenta de nuevo.");
  }
}

function normalizeCellValue(value) {
  return String(value ?? "")
    .replace(/\r\n/g, " ").replace(/\r/g, " ").replace(/\n/g, " ").trim();
}

function getLongestLineLength(value) {
  return normalizeCellValue(value).length;
}

function calculateColumnWidths(data, options = {}) {
  const minWidth = options.minWidth ?? 5;
  const horizontalPadding = options.horizontalPadding ?? 2;
  const maxWidth = options.maxWidth ?? null;
  const widths = [];

  for (let columnIndex = 0; columnIndex < data.colCount; columnIndex++) {
    let longestValue = getLongestLineLength(data.headers[columnIndex] ?? "");
    for (const row of data.rows) {
      longestValue = Math.max(longestValue, getLongestLineLength(row[columnIndex] ?? ""));
    }

    const calculatedWidth = Math.max(longestValue + horizontalPadding * 2, minWidth);
    widths.push(maxWidth === null ? calculatedWidth : Math.min(calculatedWidth, maxWidth));
  }
  return widths;
}

function fitText(value, width) {
  const text = normalizeCellValue(value);
  if (text.length <= width) return text;
  if (width <= 3) return text.substring(0, width);
  return `${text.substring(0, width - 3)}...`;
}

function centerText(value, width) {
  const text = fitText(value, width);
  const remaining = Math.max(width - text.length, 0);
  const leftPadding = Math.floor(remaining / 2);
  return " ".repeat(leftPadding) + text + " ".repeat(remaining - leftPadding);
}

function leftText(value, width) {
  return fitText(value, width).padEnd(width, " ");
}

/**
 * Copia la tabla para WhatsApp/Chat sin limitar artificialmente las columnas.
 * Antes maxWidth: 40 truncaba cualquier celda de más de 40 caracteres.
 */
async function copyForChat(data) {
  const colWidths = calculateColumnWidths(data, {
    minWidth: 7,
    horizontalPadding: 1,
    maxWidth: null
  });

  const cleanAddress = data.address.split("!").pop() || data.address;
  const primaryBg = data.headerStyles[0]?.bg || "#0f172a";
  const title = `${getColorBadge(primaryBg)} *TABLA EXCEL (${cleanAddress})*`;
  const headerLine = data.headers.map((header, index) => centerText(header, colWidths[index])).join(" | ");
  const separatorLine = colWidths.map(width => "=".repeat(width)).join("=+=");
  const dataLines = data.rows.map(row => row.map((cell, index) => leftText(cell, colWidths[index])).join(" | "));

  await copyTextToClipboard([
    title,
    "",
    "```text",
    headerLine,
    separatorLine,
    ...dataLines,
    "```"
  ].join("\n"));
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch (error) {
      console.warn("clipboard.writeText falló, intentando fallback:", error);
    }
  }

  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.setAttribute("readonly", "");
  textarea.style.position = "fixed";
  textarea.style.top = "-9999px";
  textarea.style.left = "-9999px";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();

  let success = false;
  try {
    success = document.execCommand("copy");
  } finally {
    document.body.removeChild(textarea);
  }
  if (!success) throw new Error("Este entorno no permite copiar al portapapeles. Prueba otra versión de Excel o revisa los permisos.");
}

/* Mantiene las operaciones de imagen y correo existentes. */
async function copyAsImage(data) {
  if (!navigator.clipboard || typeof navigator.clipboard.write !== "function" || typeof ClipboardItem === "undefined") {
    throw new Error("Este entorno de Excel no admite copiar imágenes al portapapeles.");
  }

  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  if (!context) throw new Error("No se pudo crear el contexto gráfico del canvas.");

  const scale = 2;
  const padding = 24;
  const rowHeight = 36;
  const widths = data.headers.map((header, index) => {
    const values = [header, ...data.rows.map(row => row[index] ?? "")];
    context.font = "13px sans-serif";
    return Math.min(500, Math.max(80, Math.ceil(Math.max(...values.map(value => context.measureText(normalizeCellValue(value)).width)) + 32)));
  });
  const tableWidth = widths.reduce((sum, width) => sum + width, 0);
  const totalWidth = tableWidth + padding * 2;
  const totalHeight = (data.rows.length + 1) * rowHeight + padding * 2;
  canvas.width = totalWidth * scale;
  canvas.height = totalHeight * scale;
  context.setTransform(scale, 0, 0, scale, 0, 0);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, totalWidth, totalHeight);

  let x = padding;
  context.font = "bold 13px sans-serif";
  data.headers.forEach((header, index) => {
    context.fillStyle = data.headerStyles[index]?.bg || "#0f172a";
    context.fillRect(x, padding, widths[index], rowHeight);
    context.fillStyle = data.headerStyles[index]?.color || "#ffffff";
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillText(normalizeCellValue(header), x + widths[index] / 2, padding + rowHeight / 2, widths[index] - 8);
    x += widths[index];
  });

  context.font = "13px sans-serif";
  data.rows.forEach((row, rowIndex) => {
    x = padding;
    const y = padding + (rowIndex + 1) * rowHeight;
    context.fillStyle = rowIndex % 2 ? "#f8fafc" : "#ffffff";
    context.fillRect(padding, y, tableWidth, rowHeight);
    context.fillStyle = "#1e293b";
    context.textAlign = "left";
    row.forEach((cell, index) => {
      context.fillText(normalizeCellValue(cell), x + 8, y + rowHeight / 2, widths[index] - 16);
      x += widths[index];
    });
  });

  await new Promise((resolve, reject) => canvas.toBlob(async blob => {
    if (!blob) return reject(new Error("No se pudo generar el blob de la imagen."));
    try {
      await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
      resolve();
    } catch (error) {
      reject(error);
    }
  }, "image/png", 1.0));
}

async function copyForEmail(data) {
  let html = '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px"><thead><tr>';
  data.headers.forEach((header, index) => {
    const style = data.headerStyles[index] || { bg: "#0f172a", color: "#ffffff", bold: true };
    html += `<th style="background-color:${style.bg};color:${style.color};padding:10px 14px;border:1px solid ${style.bg}">${escapeHtml(normalizeCellValue(header))}</th>`;
  });
  html += "</tr></thead><tbody>";
  data.rows.forEach((row, rowIndex) => {
    html += `<tr style="background-color:${rowIndex % 2 ? "#f8fafc" : "#ffffff"}">`;
    row.forEach(cell => { html += `<td style="padding:8px 14px;border:1px solid #e2e8f0;white-space:nowrap">${escapeHtml(normalizeCellValue(cell))}</td>`; });
    html += "</tr>";
  });
  html += "</tbody></table>";
  const plainText = [data.headers, ...data.rows].map(row => row.map(normalizeCellValue).join("\t")).join("\n");

  if (navigator.clipboard && typeof navigator.clipboard.write === "function" && typeof ClipboardItem !== "undefined") {
    try {
      await navigator.clipboard.write([new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([plainText], { type: "text/plain" })
      })]);
      return;
    } catch (error) {
      console.warn("Fallo al copiar HTML. Intentando fallback de texto.", error);
    }
  }
  await copyTextToClipboard(plainText);
}

function renderPreview(data) {
  const container = document.getElementById("preview-container");
  if (!container) return;
  let html = '<div class="preview-table-wrapper"><table class="preview-table"><thead><tr>';
  data.headers.forEach((header, index) => {
    const style = data.headerStyles[index] || { bg: "#0f172a", color: "#ffffff" };
    html += `<th style="background-color:${style.bg};color:${style.color};text-align:center;white-space:nowrap">${escapeHtml(normalizeCellValue(header))}</th>`;
  });
  html += "</tr></thead><tbody>";
  data.rows.slice(0, 5).forEach((row, rowIndex) => {
    html += `<tr style="background-color:${rowIndex % 2 ? "#f8fafc" : "#ffffff"}">`;
    row.forEach(cell => { html += `<td style="white-space:nowrap">${escapeHtml(normalizeCellValue(cell))}</td>`; });
    html += "</tr>";
  });
  if (data.rows.length > 5) html += `<tr><td colspan="${data.colCount}">... y ${data.rows.length - 5} filas más</td></tr>`;
  container.innerHTML = html + "</tbody></table></div>";
}

function showToast(message) {
  const toast = document.getElementById("toast");
  const toastMsg = document.getElementById("toast-message");
  if (!toast || !toastMsg) return;
  toastMsg.textContent = message;
  toast.classList.remove("hidden");
  if (toastTimeout) clearTimeout(toastTimeout);
  toastTimeout = window.setTimeout(() => toast.classList.add("hidden"), 2500);
}

function isLightColor(hexColor) {
  if (!hexColor) return false;
  const hex = hexColor.replace("#", "");
  if (hex.length !== 6) return false;
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.65;
}

function getColorBadge(hexColor) {
  if (!hexColor) return "⬛";
  const hex = hexColor.replace("#", "").toUpperCase();
  if (hex.length !== 6) return "⬛";
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  if (r < 60 && g < 60 && b < 60) return "⬛";
  if (r > 200 && g > 200 && b > 200) return "⬜";
  if (g > r + 25 && g > b) return "🟩";
  if (b > r + 20 && b > g) return "🟦";
  if (r > 180 && g > 120 && b < 80) return "🟧";
  if (r > 170 && g < 80 && b < 80) return "🟥";
  if (r > 120 && b > 120 && g < 100) return "🟪";
  return "⬛";
}

function isNumeric(val) {
  if (!val) return false;
  const clean = String(val).trim().replace(/[\$,€£%\s]/g, "").replace(/,/g, "");
  return !isNaN(Number(clean)) && clean !== "";
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = String(text ?? "");
  return div.innerHTML;
}
