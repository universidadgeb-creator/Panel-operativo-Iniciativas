// ================================================================
// GEB Iniciativas Humanas — Web App
// Despliega como: Ejecutar como "Yo", Acceso "Cualquier persona"
// ================================================================

const SS = SpreadsheetApp.openById("1fS5qeoB1ViuCUP4HOQ1zTJgh4tfWUkObvb5LMr3to5A");

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const tipo    = payload.tipo;

    if (tipo === "asistencia")         return handleAsistencia(payload);
    if (tipo === "examen")             return handleExamen(payload);
    if (tipo === "donacion")           return handleDonacion(payload);
    if (tipo === "ap_asistencia")      return handleApAsistencia(payload);
    if (tipo === "ap_cambioModalidad") return handleApCambioModalidad(payload);
    if (tipo === "ap_sesionesI43")     return handleApSesionesI43(payload);
    if (tipo === "ra_reaccion")        return handleRaReaccion(payload);

    return resp({ ok: false, error: "Tipo desconocido: " + tipo });
  } catch(err) {
    return resp({ ok: false, error: err.message });
  }
}

// ── Asistencia INEA ─────────────────────────────────────────────
// Cols: A=Nombre, B=Fecha_Clase, C=Tema, D=Asistio, E=Notas
function handleAsistencia(p) {
  const ws = SS.getSheetByName("EG_Asistencias");
  if (!ws) return resp({ ok: false, error: "Pestaña EG_Asistencias no encontrada" });
  const nombres = (p.asistentes || []);
  nombres.forEach(function(nombre) {
    ws.appendRow([
      nombre,
      p.fecha  || "",
      p.tema   || "",
      "Sí",
      p.notas  || ""
    ]);
  });
  return resp({ ok: true });
}

// ── Examen ──────────────────────────────────────────────────────
// Cols: A=Nombre, B=Modalidad, C=Tipo, D=Tema_Materia, E=Fecha_Examen, F=Resultado, G=Calificacion, H=Notas
function handleExamen(p) {
  const ws = SS.getSheetByName("EG_Examenes");
  if (!ws) return resp({ ok: false, error: "Pestaña EG_Examenes no encontrada" });
  ws.appendRow([
    p.nombre       || "",
    p.modalidad    || "",
    p.tipoExamen   || "",
    p.tema         || "",
    p.fecha        || "",
    p.resultado    || "",
    p.calificacion || "",
    p.notas        || ""
  ]);
  return resp({ ok: true });
}

// ── Donación ────────────────────────────────────────────────────
// Cols: A=Nombre, B=Mes, C=Año, D=Dono, E=Fecha_Donacion, F=Talon_Recibido, G=Notas
function handleDonacion(p) {
  const ws = SS.getSheetByName("SV_Historial");
  if (!ws) return resp({ ok: false, error: "Pestaña SV_Historial no encontrada" });
  ws.appendRow([
    p.nombre        || "",
    p.mes           || "",
    p.año           || "",
    p.dono          || "",
    p.fechaDonacion || "",
    p.talon         || "",
    p.notas         || ""
  ]);
  return resp({ ok: true });
}

// ── Atención Psicológica: Registrar asistencia ──────────────────
// Cols AP_Sesiones: A=Nombre, B=Num_Sesion, C=Fecha_Programada, D=Hora_Programada,
//                   E=Modalidad, F=Asistio, G=Cambio_Modalidad, H=Notas_Sesion
function handleApAsistencia(p) {
  var ws = SS.getSheetByName("AP_Sesiones");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Sesiones no encontrada" });

  var nombre    = (p.nombre || "").trim();
  var numSesion = String(p.numSesion || "").trim();
  var data      = ws.getDataRange().getValues();
  var found     = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()
        && String(data[i][1]).trim() === numSesion) {
      ws.getRange(i + 1, 6).setValue(p.asistio || "");
      ws.getRange(i + 1, 8).setValue(p.notas || "");
      found = true;
      break;
    }
  }

  if (!found) {
    ws.appendRow([nombre, numSesion, "", "", "", p.asistio || "", "", p.notas || ""]);
  }

  // Actualizar Sesiones_Tomadas y Estado en AP_Inscritos
  var wsIns = SS.getSheetByName("AP_Inscritos");
  if (wsIns) {
    var insData = wsIns.getDataRange().getValues();
    for (var j = 1; j < insData.length; j++) {
      if (String(insData[j][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        // Contar sesiones con Asistio = Sí
        var sesData = ws.getDataRange().getValues();
        var count = 0;
        for (var k = 1; k < sesData.length; k++) {
          if (String(sesData[k][0]).trim().toLowerCase() === nombre.toLowerCase()) {
            var a = String(sesData[k][5]).trim().toLowerCase();
            if (a === "sí" || a === "si") count++;
          }
        }
        wsIns.getRange(j + 1, 11).setValue(count); // Sesiones_Tomadas (col K, índice 10)

        var totalSesiones = parseInt(insData[j][9]) || 5;
        if (count >= totalSesiones) {
          wsIns.getRange(j + 1, 8).setValue("Completó");
        } else if (p.asistio === "No") {
          wsIns.getRange(j + 1, 8).setValue("Sesión pendiente");
        }
        break;
      }
    }
  }

  return resp({ ok: true });
}

// ── Atención Psicológica: Cambio de modalidad ───────────────────
function handleApCambioModalidad(p) {
  var ws = SS.getSheetByName("AP_Sesiones");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Sesiones no encontrada" });

  var nombre    = (p.nombre || "").trim();
  var numSesion = String(p.numSesion || "").trim();
  var data      = ws.getDataRange().getValues();
  var found     = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()
        && String(data[i][1]).trim() === numSesion) {
      ws.getRange(i + 1, 5).setValue(p.modalidad || "");
      ws.getRange(i + 1, 7).setValue("Sí");
      found = true;
      break;
    }
  }

  if (!found) {
    ws.appendRow([nombre, numSesion, "", "", p.modalidad || "", "", "Sí", ""]);
  }

  return resp({ ok: true });
}

// ── Atención Psicológica: Sesiones Impulso43 ────────────────────
function handleApSesionesI43(p) {
  var ws = SS.getSheetByName("AP_Inscritos");
  if (!ws) return resp({ ok: false, error: "Pestaña AP_Inscritos no encontrada" });

  var nombre = (p.nombre || "").trim();
  var data   = ws.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
      ws.getRange(i + 1, 11).setValue(parseInt(p.sesionesTomadas) || 0);
      break;
    }
  }

  return resp({ ok: true });
}

// ── Reto Ahorro: Registrar reacción semanal ─────────────────────
// Cols RA_Semanas: A=Nombre, B=Semana, C=Reacciono, D=Notas
function handleRaReaccion(p) {
  var ws = SS.getSheetByName("RA_Semanas");
  if (!ws) return resp({ ok: false, error: "Pestaña RA_Semanas no encontrada" });

  var nombre  = (p.nombre || "").trim();
  var semana  = String(p.semana || "").trim();
  var data    = ws.getDataRange().getValues();
  var found   = false;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === nombre.toLowerCase()
        && String(data[i][1]).trim() === semana) {
      ws.getRange(i + 1, 3).setValue(p.reacciono || "");
      ws.getRange(i + 1, 4).setValue(p.notas || "");
      found = true;
      break;
    }
  }

  if (!found) {
    ws.appendRow([nombre, semana, p.reacciono || "", p.notas || ""]);
  }

  // Actualizar Semana_Actual en RA_Inscritos si avanzó
  var wsIns = SS.getSheetByName("RA_Inscritos");
  if (wsIns) {
    var insData = wsIns.getDataRange().getValues();
    for (var j = 1; j < insData.length; j++) {
      if (String(insData[j][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        var semActual = parseInt(insData[j][4]) || 0;
        if (parseInt(semana) > semActual) {
          wsIns.getRange(j + 1, 5).setValue(parseInt(semana));
        }
        break;
      }
    }
  }

  return resp({ ok: true });
}

// ── Respuesta CORS ───────────────────────────────────────────────
function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET para verificar que el Web App está activo
function doGet() {
  return resp({ ok: true, msg: "GEB Web App activo" });
}
