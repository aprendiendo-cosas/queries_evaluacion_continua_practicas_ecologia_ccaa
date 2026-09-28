/**
 * Backend para el visor analítico de respuestas.
 * Conexión externa a la base de datos de prácticas.
 */
const CONFIG_APP = {
  // ID extraído de: https://docs.google.com/spreadsheets/d/1lc-ke9_Es8hTT1VdwxF_yJpqioBZ8rdrH4dujU1HktU/edit
  SOURCE_SPREADSHEET_ID: "1lc-ke9_Es8hTT1VdwxF_yJpqioBZ8rdrH4dujU1HktU",
  SHEETS: {
    PRACTICAS: "dicc_practicas",
    ENUNCIADOS_ALT1: "dicc_enunciados_detallados_x_pr",
    ENUNCIADOS_ALT2: "dicc_enunciados_detallados_x_practica",
    ESTUDIANTES: "dicc_estudiantes",
    RESPUESTAS: "respuestas"
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Visor de Respuestas - Ecología CCAA")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Abre y retorna la hoja de cálculo remota donde residen los datos
 */
function getSourceSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG_APP.SOURCE_SPREADSHEET_ID);
}

/**
 * Localiza la hoja de enunciados tolerando variaciones en el nombre de pestaña
 */
function getSheetEnunciados(ss) {
  let sheet = ss.getSheetByName(CONFIG_APP.SHEETS.ENUNCIADOS_ALT1);
  if (!sheet) {
    sheet = ss.getSheetByName(CONFIG_APP.SHEETS.ENUNCIADOS_ALT2);
  }
  return sheet;
}

/**
 * Devuelve el listado de prácticas disponibles (id y nombre)
 */
function obtenerPracticas() {
  const ss = getSourceSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_APP.SHEETS.PRACTICAS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const practicas = [];
  
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    const nombre = String(data[i][1]).trim();
    if (id !== "") {
      practicas.push({ id: id, nombre: nombre });
    }
  }
  return practicas;
}

/**
 * Devuelve las sesiones numéricas disponibles para una práctica dada
 */
function obtenerSesiones(idPractica) {
  const ss = getSourceSpreadsheet();
  const sheet = getSheetEnunciados(ss);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const sesiones = new Set();
  const targetId = String(idPractica).trim();

  for (let i = 1; i < data.length; i++) {
    const rowPractica = String(data[i][1]).trim();
    const rowSesion = data[i][2];
    if (rowPractica === targetId && rowSesion !== "") {
      sesiones.add(Number(rowSesion));
    }
  }
  return Array.from(sesiones).sort((a, b) => a - b);
}

/**
 * Devuelve las preguntas correspondientes a una práctica y sesión específicas
 */
function obtenerPreguntas(idPractica, sesion) {
  const ss = getSourceSpreadsheet();
  const sheet = getSheetEnunciados(ss);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const targetId = String(idPractica).trim();
  const targetSesion = parseInt(sesion, 10);
  const preguntas = [];

  for (let i = 1; i < data.length; i++) {
    const rowIdEnunciado = String(data[i][0]).trim();
    const rowPractica = String(data[i][1]).trim();
    const rowSesion = parseInt(data[i][2], 10);
    const rowTexto = String(data[i][4]).trim();

    if (rowPractica === targetId && rowSesion === targetSesion) {
      preguntas.push({
        id_enunciado: rowIdEnunciado,
        enunciado: rowTexto
      });
    }
  }
  return preguntas;
}

/**
 * Cruza las respuestas con los datos de estudiante y las ordena por el campo 'GM'
 */
function consultarRespuestas(idEnunciado) {
  const ss = getSourceSpreadsheet();
  const targetEnunciado = String(idEnunciado).trim();

  // 1. Diccionario de estudiantes: Col A = Nombre, Col B = NIE, Col D = GM
  const sheetEst = ss.getSheetByName(CONFIG_APP.SHEETS.ESTUDIANTES);
  const mapaEstudiantes = new Map();
  if (sheetEst) {
    const dataEst = sheetEst.getDataRange().getValues();
    for (let i = 1; i < dataEst.length; i++) {
      const nombre = String(dataEst[i][0]).trim();
      const nie = String(dataEst[i][1]).trim();
      const gm = String(dataEst[i][3]).trim();
      if (nie !== "") {
        mapaEstudiantes.set(nie, {
          nombre: nombre || "Desconocido",
          gm: gm || "Sin grupo"
        });
      }
    }
  }

  // 2. Filtrado de respuestas: Col D = NIE, Col E = id_enunciado, Col F = respuesta_larga
  const sheetResp = ss.getSheetByName(CONFIG_APP.SHEETS.RESPUESTAS);
  if (!sheetResp) return [];

  const dataResp = sheetResp.getDataRange().getValues();
  const resultados = [];

  for (let i = 1; i < dataResp.length; i++) {
    const rowEnunciado = String(dataResp[i][4]).trim();
    if (rowEnunciado === targetEnunciado) {
      const nie = String(dataResp[i][3]).trim();
      const textoRespuesta = String(dataResp[i][5]).trim();
      const infoEstudiante = mapaEstudiantes.get(nie) || {
        nombre: "Estudiante no registrado (" + nie + ")",
        gm: "Sin grupo"
      };

      resultados.push({
        nie: nie,
        nombre: infoEstudiante.nombre,
        gm: infoEstudiante.gm,
        respuesta: textoRespuesta
      });
    }
  }

  // 3. Ordenación determinista: primero por GM (natural sort) y luego por Nombre
  resultados.sort((a, b) => {
    const cmpGm = a.gm.localeCompare(b.gm, undefined, { numeric: true, sensitivity: "base" });
    if (cmpGm !== 0) return cmpGm;
    return a.nombre.localeCompare(b.nombre, undefined, { sensitivity: "base" });
  });

  return resultados;
}