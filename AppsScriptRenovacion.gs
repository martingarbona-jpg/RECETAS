/*
 * Integracion:
 * 1. Agregar en el dispatcher de doPost:
 *      case "renovarReceta":
 *        return responderJson(renovarReceta(data));
 * 2. Agregar en el dispatcher de doGet:
 *      case "reportes":
 *        return responderJson(obtenerReportes());
 * 3. Si los nombres difieren, ajustar solamente HOJA_RECETAS y HOJA_HISTORIAL.
 * 4. Usar cicloVigenteDuplicado(...) en la validacion de guardarReceta.
 * 5. En buscar/vencidas/porVencer no filtrar solo estado ACTIVA:
 *    incluir RENOVADA, excluir ELIMINADA y aplicar ordenarCiclosParaRespuesta.
 */

var HOJA_RECETAS = "Recetas";
var HOJA_HISTORIAL = "Historial";
var ESTADO_ACTIVO = "ACTIVA";
var ESTADO_RENOVADO = "RENOVADA";
var MAX_RECETAS_CICLO = 5;

function obtenerReportes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var recetas = obtenerTabla_(ss.getSheetByName(HOJA_RECETAS));
  var historial = obtenerTabla_(ss.getSheetByName(HOJA_HISTORIAL));
  var emisionesPorReceta = {};

  historial.filas.forEach(function(fila) {
    var id = String(valorCampo_(fila.objeto, ["id receta", "idreceta", "id"]) || "");
    if (!id) return;
    emisionesPorReceta[id] = (emisionesPorReceta[id] || 0) + 1;
  });

  return {
    ok: true,
    reportes: {
      recetas: recetas.filas.map(function(fila) {
        return normalizarRecetaParaReporte_(fila.objeto, emisionesPorReceta);
      }),
      historial: historial.filas.map(function(fila) {
        return normalizarHistorialParaReporte_(fila.objeto, recetas);
      })
    }
  };
}

function normalizarRecetaParaReporte_(objeto, emisionesPorReceta) {
  var id = String(valorCampo_(objeto, ["id", "id receta", "idreceta"]) || "");
  var fechaOriginal = fechaCampo_(objeto, ["fecha receta original", "fecharecetaoriginal"]);
  var estado = mayusculas_(valorCampo_(objeto, ["estado"])) || ESTADO_ACTIVO;
  var emitidas = emisionesPorReceta[id] || 0;
  var vencimiento = fechaOriginal ? sumarMeses_(fechaOriginal, MAX_RECETAS_CICLO) : null;
  var diasRestantes = vencimiento ? Math.ceil((finDelDia_(vencimiento) - new Date()) / 86400000) : 0;

  return {
    id: id,
    idRecetaAnterior: valorCampo_(objeto, ["id receta anterior", "idrecetaanterior"]),
    paciente: valorCampo_(objeto, ["paciente", "nombre", "nombre y apellido"]),
    dni: valorCampo_(objeto, ["dni"]),
    droga: valorCampo_(objeto, ["droga", "medicamento"]),
    fechaRecetaOriginal: fechaOriginal ? fechaOriginal.toISOString() : "",
    estado: estado,
    recetasEmitidas: emitidas,
    recetasRestantes: Math.max(0, MAX_RECETAS_CICLO - emitidas),
    diasRestantes: diasRestantes
  };
}

function normalizarHistorialParaReporte_(objeto, recetas) {
  var idReceta = String(valorCampo_(objeto, ["id receta", "idreceta", "id"]) || "");
  var receta = buscarFilaPorId_(recetas, idReceta);

  return {
    idReceta: idReceta,
    paciente: valorCampo_(objeto, ["paciente", "nombre", "nombre y apellido"]) ||
      (receta ? valorCampo_(receta.objeto, ["paciente", "nombre", "nombre y apellido"]) : ""),
    dni: valorCampo_(objeto, ["dni"]) ||
      (receta ? valorCampo_(receta.objeto, ["dni"]) : ""),
    droga: valorCampo_(objeto, ["droga", "medicamento"]) ||
      (receta ? valorCampo_(receta.objeto, ["droga", "medicamento"]) : ""),
    fechaEmision: fechaCampo_(objeto, ["fecha emision", "fechaemision", "fecha"]) || "",
    mesCorrespondiente: valorCampo_(objeto, ["mes correspondiente", "mescorrespondiente"])
  };
}

function renovarReceta(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    data = data || {};

    var idAnterior = String(data.idRecetaAnterior || "").trim();
    var fechaNueva = parsearFechaLocal_(data.fechaRecetaOriginal);
    var emitidasPrevias = Number(data.emitidasPrevias);

    if (!idAnterior || !fechaNueva) {
      return { ok: false, mensaje: "Faltan el ciclo anterior o la nueva fecha de receta." };
    }

    if (!esEnteroEntre_(emitidasPrevias, 0, MAX_RECETAS_CICLO)) {
      return { ok: false, mensaje: "Las recetas emitidas previamente deben estar entre 0 y 5." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var recetas = obtenerTabla_(ss.getSheetByName(HOJA_RECETAS));
    var historial = obtenerTabla_(ss.getSheetByName(HOJA_HISTORIAL));
    var anterior = buscarFilaPorId_(recetas, idAnterior);

    if (!anterior) {
      return { ok: false, mensaje: "No se encontro la receta anterior." };
    }

    var estadoAnterior = mayusculas_(valorCampo_(anterior.objeto, ["estado"]));
    if (estadoAnterior === ESTADO_RENOVADO || estadoAnterior === "ELIMINADA") {
      return { ok: false, mensaje: "Este ciclo ya no esta activo." };
    }

    var emitidasAnteriores = contarEmisiones_(historial, idAnterior);
    var fechaAnterior = fechaCampo_(anterior.objeto, ["fecha receta original", "fecharecetaoriginal"]);
    var vencida = fechaAnterior ? finDelDia_(sumarMeses_(fechaAnterior, 5)) < new Date() : false;

    if (!vencida && emitidasAnteriores < MAX_RECETAS_CICLO) {
      return { ok: false, mensaje: "El tratamiento todavia tiene un ciclo vigente." };
    }

    var dni = valorCampo_(anterior.objeto, ["dni"]);
    var droga = valorCampo_(anterior.objeto, ["droga", "medicamento"]);

    if (cicloVigenteDuplicado_(recetas, historial, dni, droga, idAnterior)) {
      return { ok: false, mensaje: "Ya existe otro ciclo vigente para este DNI y droga." };
    }

    asegurarColumnas_(recetas.hoja, ["ID Receta Anterior", "ID Receta Siguiente"]);
    recetas = obtenerTabla_(recetas.hoja);
    anterior = buscarFilaPorId_(recetas, idAnterior);

    var nuevoId = "REC-" + Date.now() + "-R";
    var nuevaFila = anterior.valores.slice();

    asignarCampo_(recetas, nuevaFila, ["id", "id receta", "idreceta"], nuevoId);
    asignarCampo_(recetas, nuevaFila, ["fecha carga", "fechacarga"], new Date());
    asignarCampo_(recetas, nuevaFila, ["fecha receta original", "fecharecetaoriginal"], fechaNueva);
    asignarCampo_(recetas, nuevaFila, ["estado"], ESTADO_ACTIVO);
    asignarCampo_(recetas, nuevaFila, ["id receta anterior", "idrecetaanterior"], idAnterior);
    asignarCampo_(recetas, nuevaFila, ["id receta siguiente", "idrecetasiguiente"], "");

    recetas.hoja.appendRow(nuevaFila);

    actualizarCampoFila_(recetas, anterior.numeroFila, ["estado"], ESTADO_RENOVADO);
    actualizarCampoFila_(
      recetas,
      anterior.numeroFila,
      ["id receta siguiente", "idrecetasiguiente"],
      nuevoId
    );

    registrarEmitidasPrevias_(
      historial,
      nuevoId,
      fechaNueva,
      emitidasPrevias,
      valorCampo_(anterior.objeto, ["paciente", "nombre", "nombre y apellido"]),
      droga
    );

    return {
      ok: true,
      mensaje: "Tratamiento renovado correctamente.",
      idReceta: nuevoId,
      idRecetaAnterior: idAnterior
    };
  } catch (error) {
    console.error(error);
    return { ok: false, mensaje: "No se pudo renovar el tratamiento: " + error.message };
  } finally {
    lock.releaseLock();
  }
}

/*
 * Reemplaza el bloqueo por simple coincidencia DNI/droga al guardar.
 * Devuelve true solo si existe otro ciclo realmente vigente.
 */
function cicloVigenteDuplicado(dni, droga) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return cicloVigenteDuplicado_(
    obtenerTabla_(ss.getSheetByName(HOJA_RECETAS)),
    obtenerTabla_(ss.getSheetByName(HOJA_HISTORIAL)),
    dni,
    droga,
    ""
  );
}

function cicloVigenteDuplicado_(recetas, historial, dni, droga, excluirId) {
  var dniNormalizado = normalizar_(dni);
  var drogaNormalizada = normalizar_(droga);
  var hoy = new Date();

  return recetas.filas.some(function(fila) {
    var id = String(valorCampo_(fila.objeto, ["id", "id receta", "idreceta"]) || "");
    var estado = mayusculas_(valorCampo_(fila.objeto, ["estado"]));
    var fecha = fechaCampo_(fila.objeto, ["fecha receta original", "fecharecetaoriginal"]);
    var mismoTratamiento =
      normalizar_(valorCampo_(fila.objeto, ["dni"])) === dniNormalizado &&
      normalizar_(valorCampo_(fila.objeto, ["droga", "medicamento"])) === drogaNormalizada;

    if (!mismoTratamiento || id === excluirId) return false;
    if (estado === ESTADO_RENOVADO || estado === "ELIMINADA") return false;
    if (!fecha || finDelDia_(sumarMeses_(fecha, 5)) < hoy) return false;

    return contarEmisiones_(historial, id) < MAX_RECETAS_CICLO;
  });
}

/*
 * Aplicar al final de buscar, vencidas y porVencer antes de responder.
 * El ciclo vigente queda primero y luego se ordena por fecha descendente.
 */
function ordenarCiclosParaRespuesta(resultados) {
  return (resultados || []).sort(function(a, b) {
    var inactivoA = ["RENOVADA", "ELIMINADA"].indexOf(mayusculas_(a.estado)) >= 0;
    var inactivoB = ["RENOVADA", "ELIMINADA"].indexOf(mayusculas_(b.estado)) >= 0;

    if (inactivoA !== inactivoB) return inactivoA ? 1 : -1;
    return new Date(b.fechaRecetaOriginal || 0) - new Date(a.fechaRecetaOriginal || 0);
  });
}

function registrarEmitidasPrevias_(historial, idReceta, fechaOriginal, cantidad, paciente, droga) {
  for (var i = 0; i < cantidad; i++) {
    var fila = new Array(historial.encabezados.length).fill("");
    var mes = new Date(fechaOriginal.getFullYear(), fechaOriginal.getMonth() + i, 1);

    asignarCampo_(historial, fila, ["id receta", "idreceta", "id"], idReceta);
    asignarCampo_(historial, fila, ["fecha emision", "fechaemision"], new Date());
    asignarCampo_(historial, fila, ["mes correspondiente", "mescorrespondiente"], mes);
    asignarCampo_(historial, fila, ["paciente", "nombre"], paciente);
    asignarCampo_(historial, fila, ["droga", "medicamento"], droga);

    historial.hoja.appendRow(fila);
  }
}

function contarEmisiones_(historial, idReceta) {
  return historial.filas.filter(function(fila) {
    return String(valorCampo_(fila.objeto, ["id receta", "idreceta", "id"]) || "") === idReceta;
  }).length;
}

function obtenerTabla_(hoja) {
  if (!hoja) throw new Error("No se encontro una de las hojas configuradas.");

  var valores = hoja.getDataRange().getValues();
  var encabezados = valores.length ? valores[0] : [];
  var claves = encabezados.map(normalizarEncabezado_);
  var filas = [];

  for (var i = 1; i < valores.length; i++) {
    var objeto = {};
    claves.forEach(function(clave, indice) {
      objeto[clave] = valores[i][indice];
    });
    filas.push({ numeroFila: i + 1, valores: valores[i], objeto: objeto });
  }

  return { hoja: hoja, encabezados: encabezados, claves: claves, filas: filas };
}

function buscarFilaPorId_(tabla, id) {
  return tabla.filas.find(function(fila) {
    return String(valorCampo_(fila.objeto, ["id", "id receta", "idreceta"]) || "") === id;
  });
}

function asegurarColumnas_(hoja, nombres) {
  var ultimaColumna = Math.max(hoja.getLastColumn(), 1);
  var encabezados = hoja.getRange(1, 1, 1, ultimaColumna).getValues()[0];
  var existentes = encabezados.map(normalizarEncabezado_);

  nombres.forEach(function(nombre) {
    if (existentes.indexOf(normalizarEncabezado_(nombre)) === -1) {
      hoja.getRange(1, encabezados.length + 1).setValue(nombre);
      encabezados.push(nombre);
      existentes.push(normalizarEncabezado_(nombre));
    }
  });
}

function asignarCampo_(tabla, fila, aliases, valor) {
  var indice = indiceCampo_(tabla, aliases);
  if (indice >= 0) fila[indice] = valor;
}

function actualizarCampoFila_(tabla, numeroFila, aliases, valor) {
  var indice = indiceCampo_(tabla, aliases);
  if (indice < 0) throw new Error("Falta la columna " + aliases[0] + ".");
  tabla.hoja.getRange(numeroFila, indice + 1).setValue(valor);
}

function indiceCampo_(tabla, aliases) {
  var normalizados = aliases.map(normalizarEncabezado_);
  for (var i = 0; i < tabla.claves.length; i++) {
    if (normalizados.indexOf(tabla.claves[i]) >= 0) return i;
  }
  return -1;
}

function valorCampo_(objeto, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var clave = normalizarEncabezado_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(objeto, clave)) return objeto[clave];
  }
  return "";
}

function fechaCampo_(objeto, aliases) {
  var valor = valorCampo_(objeto, aliases);
  if (valor instanceof Date && !isNaN(valor)) return valor;
  var fecha = valor ? new Date(valor) : null;
  return fecha && !isNaN(fecha) ? fecha : null;
}

function parsearFechaLocal_(valor) {
  var partes = String(valor || "").split("-");
  if (partes.length !== 3) return null;

  var fecha = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
  return isNaN(fecha) ? null : fecha;
}

function sumarMeses_(fecha, meses) {
  var copia = new Date(fecha);
  copia.setMonth(copia.getMonth() + meses);
  return copia;
}

function finDelDia_(fecha) {
  var copia = new Date(fecha);
  copia.setHours(23, 59, 59, 999);
  return copia;
}

function esEnteroEntre_(valor, minimo, maximo) {
  return Number.isInteger(valor) && valor >= minimo && valor <= maximo;
}

function normalizarEncabezado_(valor) {
  return normalizar_(valor).replace(/\s+/g, "");
}

function normalizar_(valor) {
  return String(valor == null ? "" : valor)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function mayusculas_(valor) {
  return String(valor || "").trim().toUpperCase();
}
