/**
 * SuperCopy Pro - Complemento para Excel
 * Exporta selecciones de Excel con formato adaptado para WhatsApp, Imagen y Correo.
 * Respeta estilos de fondo y color de fuente del encabezado de Excel.
 */

let appInitialized = false;

function initApp() {
  if (appInitialized) return;
  appInitialized = true;

  setupEventListeners();
  loadCurrentSelection();
}

document.addEventListener("DOMContentLoaded", initApp);

if (document.readyState === "complete" || document.readyState === "interactive") {
  initApp();
}

function setupEventListeners() {
  const btnRefresh = document.getElementById("btn-refresh");
  const btnChat = document.getElementById("btn-copy-chat");
  const btnImage = document.getElementById("btn-copy-image");
  const btnEmail = document.getElementById("btn-copy-email");

  btnRefresh?.addEventListener("click", () => loadCurrentSelection());
  btnChat?.addEventListener("click", () => handleCopy("chat"));
  btnImage?.addEventListener("click", () => handleCopy("image"));
  btnEmail?.addEventListener("click", () => handleCopy("email"));

  if (typeof Office !== "undefined") {
    Office.onReady(() => {
      loadCurrentSelection();

      if (typeof Excel !== "undefined") {
        Excel.run(async context => {
          context.workbook.onSelectionChanged.add(onSelectionChange);
          await context.sync();
        }).catch(() => {
          // El evento ya puede estar registrado o no estar soportado en el entorno.
        });
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

      if (range.rowCount === 0 || range.columnCount === 0) {
        return null;
      }

      const headerCells = [];
      for (let c = 0; c < range.columnCount; c++) {
        const cell = range.getCell(0, c);
        cell.load(["format/fill/color", "format/font/color", "format/font/bold"]);
        headerCells.push(cell);
      }

      await context.sync();

      const matrix = range.text;
      const rawMatrix = range.values;
      const headers = matrix[0].map(h => String(h || ""));
      const rows = matrix.slice(1).map(row => row.map(cell => String(cell || "")));

      const headerStyles = headerCells.map(cell => {
        const rawBg = cell.format.fill.color;
        const rawFont = cell.format.font.color;
        const isBold = cell.format.font.bold;

        let bg = "#0f172a";
        if (rawBg && rawBg !== "" && rawBg.toLowerCase() !== "#00000000") {
          bg = rawBg;
        }

        let color = "#ffffff";
        if (rawFont && rawFont !== "") {
          color = rawFont;
        } else {
          color = isLightColor(bg) ? "#0f172a" : "#ffffff";
        }

        return {
          bg,
          color,
          bold: isBold
        };
      });

      return {
        address: range.address,
        rowCount: range.rowCount,
        colCount: range.columnCount,
        headers,
        headerStyles,
        rows,
        rawValues: rawMatrix
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

/**
 * Normaliza el valor para evitar saltos de línea y espacios extra.
 */
function normalizeCellValue(value) {
  return String(value ?? "")
    .replace(/\r\n/g, " ")
    .replace(/\r/g, " ")
    .replace(/\n/g, " ")
    .trim();
}

/**
 * Devuelve la línea más larga de un texto para calcular ancho.
 */
function getLongestLineLength(value) {
  const text = normalizeCellValue(value);
  return text.length;
}

/**
 * Calcula el ancho de cada columna para el texto (modo WhatsApp / Chat).
 */
function calculateColumnWidths(data, options = {}) {
  const minWidth = options.minWidth ?? 5;
  const horizontalPadding = options.horizontalPadding ?? 2;
  const maxWidth = options.maxWidth ?? 40;

  const widths = [];

  for (let columnIndex = 0; columnIndex < data.colCount; columnIndex++) {
    let longestValue = 0;

    const headerValue = data.headers[columnIndex] ?? "";
    longestValue = Math.max(
      longestValue,
      getLongestLineLength(headerValue)
    );

    for (const row of data.rows) {
      const value = row[columnIndex] ?? "";
      longestValue = Math.max(
        longestValue,
        getLongestLineLength(value)
      );
    }

    const calculatedWidth = longestValue + horizontalPadding * 2;

    widths.push(Math.min(Math.max(calculatedWidth, minWidth), maxWidth));
  }

  return widths;
}

/**
 * Normaliza los datos para formato de imagen (Canvas).
 */
function normalizeImageCellValue(value) {
  return String(value ?? "")
    .replace(/\r\n/g, " ")
    .replace(/\r/g, " ")
    .replace(/\n/g, " ")
    .trim();
}

/**
 * Calcula el ancho real de columnas usando measureText() en píxeles.
 * Esto es importante porque Canvas trabaja en píxeles, no en caracteres.
 */
function calculateImageColumnWidths(data, context, fonts) {
  const minimumWidth = 80;
  const maximumWidth = 360;
  const horizontalPadding = 28;

  const widths = [];

  for (let columnIndex = 0; columnIndex < data.colCount; columnIndex++) {
    let maximumTextWidth = 0;

    context.font = fonts.header;
    const headerValue = normalizeImageCellValue(data.headers[columnIndex]);
    maximumTextWidth = Math.max(
      maximumTextWidth,
      context.measureText(headerValue).width
    );

    context.font = fonts.body;

    for (const row of data.rows) {
      const cellValue = normalizeImageCellValue(row[columnIndex]);
      maximumTextWidth = Math.max(
        maximumTextWidth,
        context.measureText(cellValue).width
      );
    }

    const calculatedWidth = Math.ceil(maximumTextWidth + horizontalPadding);

    widths.push(
      Math.min(Math.max(calculatedWidth, minimumWidth), maximumWidth)
    );
  }

  return widths;
}

/**
 * Ajusta texto para que no se salga del ancho disponible.
 */
function fitCanvasText(text, maxWidth, context) {
  const value = normalizeImageCellValue(text);

  if (context.measureText(value).width <= maxWidth) {
    return value;
  }

  const ellipsis = "...";
  const ellipsisWidth = context.measureText(ellipsis).width;

  if (ellipsisWidth >= maxWidth) {
    return ellipsis;
  }

  let result = "";

  for (const character of value) {
    const candidate = result + character;

    if (context.measureText(candidate).width + ellipsisWidth > maxWidth) {
      break;
    }

    result = candidate;
  }

  return `${result.trimEnd()}${ellipsis}`;
}

/**
 * Recorta texto largo para la tabla de WhatsApp.
 */
function fitText(value, width) {
  const text = normalizeCellValue(value);

  if (text.length <= width) {
    return text;
  }

  if (width <= 3) {
    return text.substring(0, width);
  }

  return `${text.substring(0, width - 3)}...`;
}

/**
 * Centra texto dentro del ancho de la columna.
 */
function centerText(value, width) {
  const text = fitText(value, width);
  const remaining = Math.max(width - text.length, 0);

  const leftPadding = Math.floor(remaining / 2);
  const rightPadding = remaining - leftPadding;

  return " ".repeat(leftPadding) + text + " ".repeat(rightPadding);
}

/**
 * Alinea a la izquierda dentro del ancho de la columna.
 */
function leftText(value, width) {
  const text = fitText(value, width);
  return text.padEnd(width, " ");
}

/**
 * Copia texto con fallback para navegadores/Excel que no permiten navigator.clipboard.
 */
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
  } catch (error) {
    console.error("Fallo en execCommand('copy'):", error);
    success = false;
  } finally {
    document.body.removeChild(textarea);
  }

  if (!success) {
    throw new Error("Este entorno no permite copiar al portapapeles. Prueba otra versión de Excel o revisa los permisos.");
  }
}

/**
 * 1. FORMATO WHATSAPP / CHAT
 * - ancho automático por contenido
 * - encabezados centrados
 * - filas alineadas a la izquierda
 */
async function copyForChat(data) {
  const colWidths = calculateColumnWidths(data, {
    minWidth: 7,
    horizontalPadding: 1,
    maxWidth: 40
  });

  const cleanAddress = data.address.split("!").pop() || data.address;
  const primaryBg = data.headerStyles[0]?.bg || "#0f172a";
  const badge = getColorBadge(primaryBg);
  const title = `${badge} *TABLA EXCEL (${cleanAddress})*`;

  const headerLine = data.headers
    .map((header, index) => centerText(header, colWidths[index]))
    .join(" | ");

  const separatorLine = colWidths
    .map(width => "=".repeat(width))
    .join("=+=");

  const dataLines = data.rows.map(row => {
    const normalizedRow = row.map(value => normalizeCellValue(value));

    return normalizedRow
      .map((cell, index) => leftText(cell, colWidths[index]))
      .join(" | ");
  });

  const output = [
    title,
    "",
    "```text",
    headerLine,
    separatorLine,
    ...dataLines,
    "```"
  ].join("\n");

  await copyTextToClipboard(output);
}

/**
 * 2. FORMATO IMAGEN EJECUTIVA
 * - ancho automático en píxeles según contenido real
 * - encabezados centrados
 * - filas alineadas a la izquierda
 */
async function copyAsImage(data) {
  if (
    !navigator.clipboard ||
    typeof navigator.clipboard.write !== "function" ||
    typeof ClipboardItem === "undefined"
  ) {
    throw new Error("Este entorno de Excel no admite copiar imágenes al portapapeles.");
  }

  const colCount = data.colCount;

  const paddingX = 32;
  const paddingY = 28;
  const headerHeight = 48;
  const rowHeight = 38;

  const fonts = {
    body: "13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    header: "600 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
  };

  const measuringCanvas = document.createElement("canvas");
  const measuringContext = measuringCanvas.getContext("2d");

  if (!measuringContext) {
    throw new Error("No se pudo iniciar el contexto de medición del canvas.");
  }

  const columnWidths = calculateImageColumnWidths(data, measuringContext, fonts);

  const tableWidth = columnWidths.reduce((total, width) => total + width, 0);
  const totalWidth = tableWidth + paddingX * 2;
  const totalHeight = paddingY * 2 + headerHeight + data.rows.length * rowHeight;

  const scale = 2;

  const canvas = document.createElement("canvas");
  canvas.width = totalWidth * scale;
  canvas.height = totalHeight * scale;

  const context = canvas.getContext("2d");

  if (!context) {
    throw new Error("No se pudo crear el contexto gráfico del canvas.");
  }

  context.scale(scale, scale);

  context.fillStyle = "#f8fafc";
  context.fillRect(0, 0, totalWidth, totalHeight);

  const cardX = paddingX;
  const cardY = paddingY;
  const cardWidth = tableWidth;
  const cardHeight = headerHeight + data.rows.length * rowHeight;

  const radius = 10;

  drawRoundedRect(context, cardX, cardY, cardWidth, cardHeight, radius);
  context.fillStyle = "#ffffff";
  context.fill();

  context.strokeStyle = "#e2e8f0";
  context.lineWidth = 1;
  context.stroke();

  context.save();

  drawRoundedRect(context, cardX, cardY, cardWidth, cardHeight, radius);
  context.clip();

  let currentX = cardX;

  for (let columnIndex = 0; columnIndex < colCount; columnIndex++) {
    const columnWidth = columnWidths[columnIndex];
    const style = data.headerStyles[columnIndex] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };

    context.fillStyle = style.bg;
    context.fillRect(currentX, cardY, columnWidth, headerHeight);

    currentX += columnWidth;
  }

  context.font = fonts.header;
  context.textBaseline = "middle";
  context.textAlign = "center";

  currentX = cardX;

  for (let columnIndex = 0; columnIndex < colCount; columnIndex++) {
    const columnWidth = columnWidths[columnIndex];
    const style = data.headerStyles[columnIndex] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };

    const headerText = fitCanvasText(
      data.headers[columnIndex],
      columnWidth - 20,
      context
    );

    context.fillStyle = style.color;
    context.fillText(
      headerText,
      currentX + columnWidth / 2,
      cardY + headerHeight / 2
    );

    currentX += columnWidth;
  }

  context.font = fonts.body;
  context.textBaseline = "middle";
  context.textAlign = "left";

  for (let rowIndex = 0; rowIndex < data.rows.length; rowIndex++) {
    const rowY = cardY + headerHeight + rowIndex * rowHeight;
    const isEvenRow = rowIndex % 2 === 1;

    context.fillStyle = isEvenRow ? "#f8fafc" : "#ffffff";
    context.fillRect(cardX, rowY, cardWidth, rowHeight);

    context.strokeStyle = "#f1f5f9";
    context.beginPath();
    context.moveTo(cardX, rowY);
    context.lineTo(cardX + cardWidth, rowY);
    context.stroke();

    context.fillStyle = "#1e293b";
    currentX = cardX;

    for (let columnIndex = 0; columnIndex < colCount; columnIndex++) {
      const columnWidth = columnWidths[columnIndex];
      const availableTextWidth = columnWidth - 28;

      const cellText = fitCanvasText(
        data.rows[rowIndex][columnIndex],
        availableTextWidth,
        context
      );

      context.fillText(
        cellText,
        currentX + 14,
        rowY + rowHeight / 2
      );

      currentX += columnWidth;
    }
  }

  context.restore();

  return new Promise((resolve, reject) => {
    canvas.toBlob(async blob => {
      if (!blob) {
        reject(new Error("No se pudo generar el blob de la imagen."));
        return;
      }

      try {
        const item = new ClipboardItem({
          "image/png": blob
        });

        await navigator.clipboard.write([item]);
        resolve();
      } catch (error) {
        reject(error);
      }
    }, "image/png");
  });
}

/**
 * 3. FORMATO CORREO / OUTLOOK / HTML
 * - ancho automático
 * - encabezados centrados
 * - filas alineadas a la izquierda
 */
async function copyForEmail(data) {
  const tableStyle = [
    "border-collapse: collapse",
    "table-layout: auto",
    "width: max-content",
    "max-width: 100%",
    "font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    "font-size: 13px",
    "color: #1e293b",
    "margin: 8px 0"
  ].join("; ");

  let html = `<table style="${tableStyle}"><thead><tr>`;

  data.headers.forEach((h, idx) => {
    const style = data.headerStyles[idx] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };

    const fontWeight = style.bold ? "700" : "600";
    const thStyle = [
      `background-color: ${style.bg}`,
      `color: ${style.color}`,
      `font-weight: ${fontWeight}`,
      "padding: 10px 14px",
      "text-align: center",
      "vertical-align: middle",
      "white-space: nowrap",
      `border: 1px solid ${style.bg}`
    ].join("; ");

    html += `<th style="${thStyle}">${escapeHtml(normalizeCellValue(h))}</th>`;
  });

  html += `</tr></thead><tbody>`;

  data.rows.forEach((row, rIdx) => {
    const bg = rIdx % 2 === 1 ? "#f8fafc" : "#ffffff";
    html += `<tr style="background-color: ${bg};">`;

    row.forEach(cell => {
      const tdStyle = [
        "padding: 8px 14px",
        "border: 1px solid #e2e8f0",
        "text-align: left",
        "vertical-align: middle",
        "white-space: nowrap"
      ].join("; ");

      html += `<td style="${tdStyle}">${escapeHtml(normalizeCellValue(cell))}</td>`;
    });

    html += `</tr>`;
  });

  html += `</tbody></table>`;

  const plainText = [
    data.headers.map(normalizeCellValue).join("\t"),
    ...data.rows.map(r => r.map(normalizeCellValue).join("\t"))
  ].join("\n");

  if (
    navigator.clipboard &&
    typeof navigator.clipboard.write === "function" &&
    typeof ClipboardItem !== "undefined"
  ) {
    try {
      const blobHtml = new Blob([html], { type: "text/html" });
      const blobText = new Blob([plainText], { type: "text/plain" });

      const item = new ClipboardItem({
        "text/html": blobHtml,
        "text/plain": blobText
      });

      await navigator.clipboard.write([item]);
      return;
    } catch (error) {
      console.warn("Fallo al copiar HTML con ClipboardItem. Intentando fallback de texto.", error);
    }
  }

  await copyTextToClipboard(plainText);
}

/**
 * Renderiza la vista previa del contenido en el panel.
 */
function renderPreview(data) {
  const container = document.getElementById("preview-container");
  if (!container) return;

  if (data.rowCount === 0) {
    container.innerHTML = `<p class="preview-placeholder">Selecciona un rango en Excel para ver la previsualización.</p>`;
    return;
  }

  let html = `
    <div class="preview-table-wrapper">
      <table class="preview-table">
        <thead>
          <tr>
  `;

  data.headers.forEach((h, idx) => {
    const style = data.headerStyles[idx] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };

    html += `
      <th
        style="
          background-color: ${style.bg};
          color: ${style.color};
          text-align: center;
          vertical-align: middle;
          white-space: nowrap;
        "
      >
        ${escapeHtml(normalizeCellValue(h))}
      </th>
    `;
  });

  html += `
          </tr>
        </thead>
        <tbody>
  `;

  data.rows.slice(0, 5).forEach((row, rowIndex) => {
    const bg = rowIndex % 2 === 1 ? "#f8fafc" : "#ffffff";

    html += `<tr style="background-color: ${bg};">`;
    row.forEach(cell => {
      html += `
        <td style="text-align: left; vertical-align: middle; white-space: nowrap;">
          ${escapeHtml(normalizeCellValue(cell))}
        </td>
      `;
    });
    html += `</tr>`;
  });

  if (data.rows.length > 5) {
    html += `<tr><td colspan="${data.colCount}" style="text-align: center; color: #94a3b8; font-style: italic;">... y ${data.rows.length - 5} filas más</td></tr>`;
  }

  html += `
        </tbody>
      </table>
    </div>
  `;

  container.innerHTML = html;
}

/**
 * Dibuja un rectángulo con esquinas redondeadas en Canvas.
 */
function drawRoundedRect(ctx, x, y, w, h, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + w - radius, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + radius);
  ctx.lineTo(x + w, y + h - radius);
  ctx.quadraticCurveTo(x + w, y + h, x + w - radius, y + h);
  ctx.lineTo(x + radius, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

/**
 * Muestra notificación flotante.
 */
let toastTimeout;

function showToast(message) {
  const toast = document.getElementById("toast");
  const toastMsg = document.getElementById("toast-message");

  if (!toast || !toastMsg) return;

  toastMsg.textContent = message;
  toast.classList.remove("hidden");

  if (toastTimeout) {
    clearTimeout(toastTimeout);
  }

  toastTimeout = window.setTimeout(() => {
    toast.classList.add("hidden");
  }, 2500);
}

/**
 * Determina si un color hexadecimal es claro para ajustar el contraste.
 */
function isLightColor(hexColor) {
  if (!hexColor) return false;

  const hex = hexColor.replace("#", "");
  if (hex.length !== 6) return false;

  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;

  return luminance > 0.65;
}

/**
 * Mapea el color de fondo dominante de Excel a un badge para WhatsApp.
 */
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

/**
 * Detecta si una cadena representa un valor numérico, moneda o porcentaje.
 */
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