# Citar-trabajadores-por-Whatsapp
Genera un panel donde toma los datos de los trabajadores que  vienen al dia siguiente, donde se pueden ir llevando el control de quien esta citado.
/**
 * Panel de citación de conductores por WhatsApp
 * -----------------------------------------------
 * Instrucciones de instalación:
 * 1. Abre tu Google Sheet "Hoja ruta".
 * 2. Menú Extensiones > Apps Script.
 * 3. Borra TODO el contenido de tu archivo .gs actual (el que ya
 *    tenías pegado) y pega TODO el contenido de este archivo en su lugar.
 * 4. Asegúrate de que existe un archivo HTML llamado exactamente
 *    "Sidebar" (sin extensión) con el contenido de Sidebar.html.
 * 5. Guarda (icono disquete) y vuelve a la hoja de cálculo.
 * 6. Recarga la hoja (F5).
 * 7. Cada día, antes de empezar a citar: menú Citaciones > Reiniciar
 *    estado (nuevo día). Esto borra los "Enviado" del día anterior.
 * 8. Luego: Citaciones > Abrir panel de envío.
 */

// Nombre exacto de la pestaña donde está la lista de conductores
const NOMBRE_HOJA = 'CONVOCATORIA';

// Número de columna (A=1, B=2, C=3...) de cada dato
const COL_NOMBRE = 1;
const COL_TELEFONO = 2;
const COL_HORA = 3;
const COL_ESTADO = 5;

// Prefijo de país a anteponer al teléfono si no lo lleva ya
const PREFIJO_PAIS = '34';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Citaciones')
    .addItem('Abrir panel de envío', 'showSidebar')
    .addItem('Reiniciar estado (nuevo día)', 'reiniciarEstado')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Citación de conductores');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getPendingDrivers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  if (!sheet) {
    throw new Error('No se encuentra la pestaña "' + NOMBRE_HOJA + '"');
  }
  const data = sheet.getDataRange().getValues();
  const drivers = [];

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    const nombre = fila[COL_NOMBRE - 1];
    const telefono = fila[COL_TELEFONO - 1];
    const hora = fila[COL_HORA - 1];
    const estado = fila[COL_ESTADO - 1];

    // Ya no exigimos teléfono para mostrar al conductor: si falta,
    // se muestra igualmente pero con un aviso (ver faltaTelefono abajo).
    if (!nombre) continue;
    if (estado === 'Enviado') continue;

    let horaTexto = hora;
    if (hora instanceof Date) {
      const zonaHoja = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      horaTexto = Utilities.formatDate(hora, zonaHoja, 'HH:mm');
    }

    let telefonoLimpio = '';
    let faltaTelefono = true;
    if (telefono) {
      const soloDigitos = telefono.toString().replace(/\D/g, '');
      if (soloDigitos) {
        faltaTelefono = false;
        telefonoLimpio = soloDigitos.startsWith(PREFIJO_PAIS)
          ? soloDigitos
          : PREFIJO_PAIS + soloDigitos;
      }
    }

    drivers.push({
      row: i + 1,
      nombre: nombre,
      telefono: telefonoLimpio,
      hora: horaTexto || '',
      faltaTelefono: faltaTelefono
    });
  }

  return drivers;
}

function marcarEnviado(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  sheet.getRange(row, COL_ESTADO).setValue('Enviado');
  return true;
}

/**
 * Borra la columna ESTADO de todos los conductores.
 * Hay que ejecutarlo cada día (o cada vez que se vuelca una nueva
 * lista de conductores en CONVOCATORIA) para que no queden marcados
 * como "Enviado" los conductores del día anterior.
 */
function reiniciarEstado() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Reiniciar citación',
    '¿Borrar el estado "Enviado" de todos los conductores para empezar un día nuevo?',
    ui.ButtonSet.YES_NO
  );
  if (respuesta !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  const ultimaFila = sheet.getLastRow();
  if (ultimaFila > 1) {
    sheet.getRange(2, COL_ESTADO, ultimaFila - 1, 1).clearContent();
  }
  ui.alert('Listo, el estado se ha reiniciado.');
}
