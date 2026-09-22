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
 * Calcula el ancho de cada columna según el texto más largo
 * entre encabezados y filas.
 */
function calculateColumnWidths(data, options = {}) {
  const minWidth = options.minWidth ?? 80;
  const padding = options.padding ?? 28;

  const widths = [];

  for (let columnIndex = 0; columnIndex < data.colCount; columnIndex++) {
    let maxLength = 0;

    const headerValue = String(data.headers[columnIndex] ?? "");
    maxLength = Math.max(maxLength, headerValue.length);

    for (const row of data.rows) {
      const value = String(row[columnIndex] ?? "");
      maxLength = Math.max(maxLength, value.length);
    }

    widths.push(Math.max(minWidth, maxLength + padding));
  }

  return widths;
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
 * - encabezado centrado
 * - datos alineados a la izquierda
 */
async function copyForChat(data) {
  const colWidths = calculateColumnWidths(data, {
    minWidth: 10,
    padding: 4
  });

  const centerText = (value, width) => {
    const text = String(value ?? "");
    const remaining = Math.max(width - text.length, 0);
    const leftPadding = Math.floor(remaining / 2);
    const rightPadding = remaining - leftPadding;

    return " ".repeat(leftPadding) + text + " ".repeat(rightPadding);
  };

  const leftText = (value, width) => {
    const text = String(value ?? "");
    return text.padEnd(width, " ");
  };

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

  const dataLines = data.rows.map(row =>
    row
      .map((cell, index) => leftText(cell, colWidths[index]))
      .join(" | ")
  );

  const output = [
    title,
    "```",
    headerLine,
    separatorLine,
    ...dataLines,
    "```"
  ].join("\n");

  await copyTextToClipboard(output);
}

/**
 * 2. FORMATO IMAGEN EJECUTIVA
 * - ancho automático por contenido
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

  const colWidths = calculateColumnWidths(data, {
    minWidth: 80,
    padding: 28
  });

  const colCount = data.colCount;
  const paddingX = 32;
  const paddingY = 28;
  const rowHeight = 36;
  const headerHeight = 46;
  const fontBody = "13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
  const fontHeader = "600 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";

  const tempCanvas = document.createElement("canvas");
  const tempCtx = tempCanvas.getContext("2d");

  if (!tempCtx) {
    throw new Error("No se pudo iniciar el contexto 2D del canvas.");
  }

  const tableWidth = colWidths.reduce((a, b) => a + b, 0);
  const totalWidth = tableWidth + paddingX * 2;
  const totalHeight = paddingY * 2 + headerHeight + data.rows.length * rowHeight;

  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = totalWidth * scale;
  canvas.height = totalHeight * scale;

  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("Error de contexto en canvas final.");
  }

  ctx.scale(scale, scale);

  ctx.fillStyle = "#f8fafc";
  ctx.fillRect(0, 0, totalWidth, totalHeight);

  const cardX = paddingX;
  const cardY = paddingY;
  const cardW = tableWidth;
  const cardH = headerHeight + data.rows.length * rowHeight;
  const radius = 10;

  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, radius);
  ctx.fillStyle = "#ffffff";
  ctx.fill();
  ctx.strokeStyle = "#e2e8f0";
  ctx.lineWidth = 1;
  ctx.stroke();

  ctx.save();
  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, radius);
  ctx.clip();

  let curX = cardX;
  for (let c = 0; c < colCount; c++) {
    const width = colWidths[c];
    const style = data.headerStyles[c] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };

    ctx.fillStyle = style.bg;
    ctx.fillRect(curX, cardY, width, headerHeight);
    curX += width;
  }

  ctx.font = fontHeader;
  ctx.textBaseline = "middle";
  ctx.textAlign = "center";

  curX = cardX;
  for (let c = 0; c < colCount; c++) {
    const width = colWidths[c];
    const text = data.headers[c] || "";
    const style = data.headerStyles[c] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };

    ctx.fillStyle = style.color;
    ctx.fillText(text, curX + width / 2, cardY + headerHeight / 2);
    curX += width;
  }

  ctx.font = fontBody;
  ctx.textAlign = "left";

  for (let r = 0; r < data.rows.length; r++) {
    const rowY = cardY + headerHeight + r * rowHeight;
    const isEven = r % 2 === 1;

    ctx.fillStyle = isEven ? "#f8fafc" : "#ffffff";
    ctx.fillRect(cardX, rowY, cardW, rowHeight);

    ctx.strokeStyle = "#f1f5f9";
    ctx.beginPath();
    ctx.moveTo(cardX, rowY);
    ctx.lineTo(cardX + cardW, rowY);
    ctx.stroke();

    ctx.fillStyle = "#1e293b";
    curX = cardX;

    for (let c = 0; c < colCount; c++) {
      const width = colWidths[c];
      const val = String(data.rows[r][c] ?? "");
      ctx.fillText(val, curX + 14, rowY + rowHeight / 2);
      curX += width;
    }
  }

  ctx.restore();

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
      } catch (err) {
        reject(err);
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

    html += `<th style="${thStyle}">${escapeHtml(h)}</th>`;
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

      html += `<td style="${tdStyle}">${escapeHtml(cell)}</td>`;
    });

    html += `</tr>`;
  });

  html += `</tbody></table>`;

  const plainText = [data.headers.join("\t"), ...data.rows.map(r => r.join("\t"))].join("\n");

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
        ${escapeHtml(h)}
      </th>
    `;
  });

  html += `
          </tr>
        </thead>
        <tbody>
  `;

  data.rows.slice(0, 5).forEach(r => {
    html += `<tr>`;
    r.forEach(c => {
      html += `
        <td style="text-align: left; vertical-align: middle; white-space: nowrap;">
          ${escapeHtml(c)}
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
 * Dibuja un rectángulo con esquinas redondeadas en Canvas
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
 * Muestra notificación flotante
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