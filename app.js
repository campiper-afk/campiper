// Configuración
const SPRINT_MASTER_EMAIL = "sprintmaster@tudominio.com"; // Reemplaza con el correo real
const SHEET_NAME = "Kanban";

/**
 * Función GET para leer todas las tarjetas al cargar la página frontend.
 */
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Si la hoja está vacía (solo encabezados)
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const headers = data.shift(); 
  const cards = data.map(row => ({
    id: row[0],
    titulo: row[1],
    descripcion: row[2],
    responsable: row[3],
    fechaInicio: row[4],
    fechaFin: row[5],
    url: row[6],
    estado: row[7]
  }));
  
  return ContentService.createTextOutput(JSON.stringify(cards))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Función POST para recibir actualizaciones desde el frontend (Drag & Drop o Creación).
 * Nota: El frontend debe enviar 'text/plain' para evitar errores de CORS en el preflight de Apps Script.
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    
    if (action === "updateStatus") {
      return actualizarEstadoTarjeta(payload.id, payload.newStatus);
    } else if (action === "createCard") {
      return crearTarjeta(payload.card);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function actualizarEstadoTarjeta(id, nuevoEstado) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) { 
      const estadoAnterior = data[i][7];
      const fila = i + 1;
      
      // Actualizar la columna 'Estado' (Columna H = 8)
      sheet.getRange(fila, 8).setValue(nuevoEstado);
      
      // Lógica de Notificación: Disparar correo solo si cambia a "En revisión"
      if (nuevoEstado === "En revisión" && estadoAnterior !== "En revisión") {
        enviarNotificacionRevision(data[i]);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Tarjeta no encontrada' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function crearTarjeta(card) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const id = new Date().getTime(); // ID único basado en timestamp
  
  sheet.appendRow([
    id,
    card.titulo,
    card.descripcion,
    card.responsable,
    card.fechaInicio,
    card.fechaFin,
    card.url,
    "Backlog" // Estado inicial por defecto
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function enviarNotificacionRevision(datosFila) {
  const titulo = datosFila[1];
  const responsable = datosFila[3];
  const url = datosFila[6] || "No se adjuntó URL";
  
  const asunto = `[LMS Kanban] Tarea Lista para Revisión: ${titulo}`;
  const cuerpoHTML = `
    <h2>Una tarea ha pasado a etapa de revisión</h2>
    <p>El miembro del equipo <strong>${responsable}</strong> ha movido una tarea a la columna "En revisión".</p>
    <ul>
      <li><strong>Tarea:</strong> ${titulo}</li>
      <li><strong>Entregable (URL):</strong> <a href="${url}">${url}</a></li>
    </ul>
    <p>Por favor, valida el entregable para mover la tarjeta a "Finalizado".</p>
  `;
  
  MailApp.sendEmail({
    to: SPRINT_MASTER_EMAIL,
    subject: asunto,
    htmlBody: cuerpoHTML
  });
}
