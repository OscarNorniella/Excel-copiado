/**
 * SuperCopy Pro - Complemento para Script Lab
 * Exporta selecciones de Excel con formato adaptado para WhatsApp, Imagen y Correo.
 * Respeta estilos de fondo y color de fuente del encabezado de Excel.
 */

// Inicialización de eventos al cargar el script
document.addEventListener("DOMContentLoaded", () => {
  setupEventListeners();
  loadCurrentSelection();
});

// Respaldo directo si el DOM ya estuviera listo en el runner de Script Lab
if (document.readyState === "complete" || document.readyState === "interactive") {
  setupEventListeners();
  loadCurrentSelection();
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

  // Suscribirse al cambio de selección activo en Excel si Office está disponible
  if (typeof Office !== "undefined") {
    Office.onReady(() => {
      loadCurrentSelection();
      Excel.run(async context => {
        context.workbook.onSelectionChanged.add(onSelectionChange);
        await context.sync();
      }).catch(() => {
        // En caso de que el evento ya esté suscrito o no sea soportado en el entorno
      });
    });
  }
}
async function onSelectionChange() {
  await loadCurrentSelection();
}

/**
 * Lee la selección actual de Excel con sus valores formateados y estilos de encabezado
 */
async function getSelectionData() {
  try {
    return await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "text", "values", "rowCount", "columnCount"]);
      await context.sync();
      if (range.rowCount === 0 || range.columnCount === 0) {
        return null;
      }

      // Cargar formato de las celdas del header (fila 0 de la selección)
      const headerCells = [];
      for (let c = 0; c < range.columnCount; c++) {
        const cell = range.getCell(0, c);
        cell.load(["format/fill/color", "format/font/color", "format/font/bold"]);
        headerCells.push(cell);
      }
      await context.sync();

      // 'text' contiene el texto ya formateado (fechas, monedas, etc.)
      const matrix = range.text;
      const rawMatrix = range.values;
      const headers = matrix[0].map(h => String(h || ""));
      const rows = matrix.slice(1).map(row => row.map(cell => String(cell || "")));

      // Extraer y normalizar colores de fondo y texto del encabezado
      const headerStyles = headerCells.map(cell => {
        const rawBg = cell.format.fill.color;
        const rawFont = cell.format.font.color;
        const isBold = cell.format.font.bold;
        let bg = "#0f172a"; // Color ejecutivo predeterminado
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

/**
 * Actualiza los indicadores visuales en la interfaz del panel
 */
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

/**
 * Manejador principal para cada una de las 3 opciones de copiado
 */
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
    showToast("Error al copiar. Intenta de nuevo.");
  }
}

/**
 * 1. FORMATO WHATSAPP / CHAT
 * WhatsApp es texto plano. Destacamos el encabezado mediante:
 * - Indicador temático según el color de fondo del header en Excel (🟦, 🟩, ⬛, etc.)
 * - Separador de doble línea (===+===) que resalta visualmente la cabecera en el bloque monoespaciado.
 */
async function copyForChat(data) {
  const allRows = [data.headers, ...data.rows];
  const colWidths = [];
  for (let c = 0; c < data.colCount; c++) {
    let max = 0;
    for (let r = 0; r < allRows.length; r++) {
      const len = (allRows[r][c] || "").length;
      if (len > max) max = len;
    }
    colWidths.push(Math.max(max, 3));
  }
  const formatRow = row => {
    return row.map((cell, idx) => {
      const width = colWidths[idx];
      const val = cell || "";
      const isNum = isNumeric(val);
      return isNum ? val.padStart(width, " ") : val.padEnd(width, " ");
    }).join(" | ");
  };
  const cleanAddress = data.address.split("!").pop() || data.address;
  const primaryBg = data.headerStyles[0]?.bg || "#0f172a";
  const badge = getColorBadge(primaryBg);
  const title = `${badge} *TABLA EXCEL (${cleanAddress})*`;
  const headerLine = formatRow(data.headers);
  // Separador de doble línea para denotar visualmente el peso del encabezado
  const separatorLine = colWidths.map(w => "=".repeat(w)).join("=+=");
  const dataLines = data.rows.map(row => formatRow(row));
  const textTable = [headerLine, separatorLine, ...dataLines].join("\n");
  const output = `${title}\n\`\`\`\n${textTable}\n\`\`\``;
  await navigator.clipboard.writeText(output);
}

/**
 * 2. FORMATO IMAGEN EJECUTIVA
 * Renderiza la tabla sobre un Canvas de alta resolución respetando el fondo y fuente del header de Excel.
 */
async function copyAsImage(data) {
  const colCount = data.colCount;

  // Medidas del canvas
  const paddingX = 32;
  const paddingY = 28;
  const rowHeight = 36;
  const headerHeight = 42;
  const fontBody = "13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";
  const fontHeader = "600 13px -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";

  // Canvas auxiliar para calcular anchos de texto
  const tempCanvas = document.createElement("canvas");
  const tempCtx = tempCanvas.getContext("2d");
  if (!tempCtx) throw new Error("No se pudo iniciar el contexto 2D");
  tempCtx.font = fontHeader;
  const colWidths = [];
  for (let c = 0; c < colCount; c++) {
    let max = tempCtx.measureText(data.headers[c] || "").width;
    tempCtx.font = fontBody;
    for (let r = 0; r < data.rows.length; r++) {
      const w = tempCtx.measureText(data.rows[r][c] || "").width;
      if (w > max) max = w;
    }
    colWidths.push(Math.ceil(max) + 28); // margen interno de columna
  }
  const tableWidth = colWidths.reduce((a, b) => a + b, 0);
  const totalWidth = tableWidth + paddingX * 2;
  const totalHeight = paddingY * 2 + headerHeight + data.rows.length * rowHeight;

  // Escala para alta densidad (Retina 2x)
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = totalWidth * scale;
  canvas.height = totalHeight * scale;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("Error de contexto en canvas final");
  ctx.scale(scale, scale);

  // Fondo exterior elegante
  ctx.fillStyle = "#f8fafc";
  ctx.fillRect(0, 0, totalWidth, totalHeight);

  // Contenedor tipo tarjeta redondeada
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

  // Clip para que las esquinas de filas respeten el radio
  ctx.save();
  drawRoundedRect(ctx, cardX, cardY, cardW, cardH, radius);
  ctx.clip();

  // 1. Cabecera (Fondo dinámico según cada celda del header de Excel)
  let curX = cardX;
  for (let c = 0; c < colCount; c++) {
    const w = colWidths[c];
    const style = data.headerStyles[c] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    ctx.fillStyle = style.bg;
    ctx.fillRect(curX, cardY, w, headerHeight);
    curX += w;
  }
  ctx.font = fontHeader;
  ctx.textBaseline = "middle";

  // Texto del encabezado con el color de fuente configurado
  curX = cardX;
  for (let c = 0; c < colCount; c++) {
    const w = colWidths[c];
    const text = data.headers[c] || "";
    const style = data.headerStyles[c] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    const isNum = data.rows.length > 0 && isNumeric(data.rows[0][c]);
    ctx.fillStyle = style.color;
    if (isNum) {
      ctx.textAlign = "right";
      ctx.fillText(text, curX + w - 14, cardY + headerHeight / 2);
    } else {
      ctx.textAlign = "left";
      ctx.fillText(text, curX + 14, cardY + headerHeight / 2);
    }
    curX += w;
  }

  // 2. Filas de datos
  ctx.font = fontBody;
  for (let r = 0; r < data.rows.length; r++) {
    const rowY = cardY + headerHeight + r * rowHeight;
    const isEven = r % 2 === 1;

    // Fila cebra sutil
    ctx.fillStyle = isEven ? "#f8fafc" : "#ffffff";
    ctx.fillRect(cardX, rowY, cardW, rowHeight);

    // Línea divisoria
    ctx.strokeStyle = "#f1f5f9";
    ctx.beginPath();
    ctx.moveTo(cardX, rowY);
    ctx.lineTo(cardX + cardW, rowY);
    ctx.stroke();

    // Texto de celdas
    ctx.fillStyle = "#1e293b";
    curX = cardX;
    for (let c = 0; c < colCount; c++) {
      const w = colWidths[c];
      const val = data.rows[r][c] || "";
      const isNum = isNumeric(val);
      if (isNum) {
        ctx.textAlign = "right";
        ctx.fillText(val, curX + w - 14, rowY + rowHeight / 2);
      } else {
        ctx.textAlign = "left";
        ctx.fillText(val, curX + 14, rowY + rowHeight / 2);
      }
      curX += w;
    }
  }
  ctx.restore();

  // Exportar a Blob y copiar al portapapeles
  return new Promise((resolve, reject) => {
    canvas.toBlob(async blob => {
      if (!blob) {
        reject(new Error("No se pudo generar el blob de la imagen"));
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
 * Genera una tabla con estilos CSS inline respetando fielmente el fondo y color de texto del header de Excel.
 */
async function copyForEmail(data) {
  const tableStyle = "border-collapse: collapse; width: 100%; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; font-size: 13px; color: #1e293b; margin: 8px 0;";
  let html = `<table style="${tableStyle}"><thead><tr>`;
  data.headers.forEach((h, idx) => {
    const style = data.headerStyles[idx] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    const isNum = data.rows.length > 0 && isNumeric(data.rows[0][idx]);
    const align = isNum ? "right" : "left";
    const fontWeight = style.bold ? "700" : "600";
    const thStyle = `background-color: ${style.bg}; color: ${style.color}; font-weight: ${fontWeight}; padding: 10px 14px; text-align: ${align}; border: 1px solid ${style.bg};`;
    html += `<th style="${thStyle}">${escapeHtml(h)}</th>`;
  });
  html += `</tr></thead><tbody>`;
  data.rows.forEach((row, rIdx) => {
    const bg = rIdx % 2 === 1 ? "#f8fafc" : "#ffffff";
    html += `<tr style="background-color: ${bg};">`;
    row.forEach(cell => {
      const isNum = isNumeric(cell);
      const align = isNum ? "right" : "left";
      const tdStyle = `padding: 8px 14px; border: 1px solid #e2e8f0; text-align: ${align};`;
      html += `<td style="${tdStyle}">${escapeHtml(cell)}</td>`;
    });
    html += `</tr>`;
  });
  html += `</tbody></table>`;

  // Fallback a texto plano
  const plainText = [data.headers.join("\t"), ...data.rows.map(r => r.join("\t"))].join("\n");
  const blobHtml = new Blob([html], {
    type: "text/html"
  });
  const blobText = new Blob([plainText], {
    type: "text/plain"
  });
  const item = new ClipboardItem({
    "text/html": blobHtml,
    "text/plain": blobText
  });
  await navigator.clipboard.write([item]);
}

/**
 * Renderiza la vista previa en el panel respetando los colores del header
 */
function renderPreview(data) {
  const container = document.getElementById("preview-container");
  if (!container) return;
  if (data.rowCount === 0) {
    container.innerHTML = `<p class="preview-placeholder">Selecciona un rango en Excel para ver la previsualización.</p>`;
    return;
  }
  let html = `<table class="preview-table"><thead><tr>`;
  data.headers.forEach((h, idx) => {
    const style = data.headerStyles[idx] || {
      bg: "#0f172a",
      color: "#ffffff",
      bold: true
    };
    html += `<th style="background-color: ${style.bg}; color: ${style.color};">${escapeHtml(h)}</th>`;
  });
  html += `</tr></thead><tbody>`;
  data.rows.slice(0, 5).forEach(r => {
    html += `<tr>`;
    r.forEach(c => {
      html += `<td>${escapeHtml(c)}</td>`;
    });
    html += `</tr>`;
  });
  if (data.rows.length > 5) {
    html += `<tr><td colspan="${data.colCount}" style="text-align: center; color: #94a3b8; font-style: italic;">... y ${data.rows.length - 5} filas más</td></tr>`;
  }
  html += `</tbody></table>`;
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
 * Muestra notificación flotante (Toast)
 */
let toastTimeout;
function showToast(message) {
  const toast = document.getElementById("toast");
  const toastMsg = document.getElementById("toast-message");
  if (!toast || !toastMsg) return;
  toastMsg.textContent = message;
  toast.classList.remove("hidden");
  if (toastTimeout) clearTimeout(toastTimeout);
  toastTimeout = window.setTimeout(() => {
    toast.classList.add("hidden");
  }, 2500);
}

/**
 * Determina si un color hexadecimal es claro para ajustar el contraste de fuente
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
 * Mapea el color de fondo dominante de Excel a un badge discreto para WhatsApp
 */
function getColorBadge(hexColor) {
  if (!hexColor) return "⬛";
  const hex = hexColor.replace("#", "").toUpperCase();
  if (hex.length !== 6) return "⬛";
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  if (r < 60 && g < 60 && b < 60) return "⬛"; // Oscuro / Negro
  if (r > 200 && g > 200 && b > 200) return "⬜"; // Claro / Blanco
  if (g > r + 25 && g > b) return "🟩"; // Verde
  if (b > r + 20 && b > g) return "🟦"; // Azul
  if (r > 180 && g > 120 && b < 80) return "🟧"; // Naranja / Ámbar
  if (r > 170 && g < 80 && b < 80) return "🟥"; // Rojo
  if (r > 120 && b > 120 && g < 100) return "🟪"; // Morado
  return "⬛";
}

/**
 * Detecta si una cadena representa un valor numérico, moneda o porcentaje
 */
function isNumeric(val) {
  if (!val) return false;
  const clean = val.trim().replace(/[\$,€£%\s]/g, "").replace(/,/g, "");
  return !isNaN(Number(clean)) && clean !== "";
}
function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}