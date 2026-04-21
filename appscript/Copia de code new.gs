// //© Banco Davivienda S.A. 2021. Se prohíbe su uso o reproducción sin previa autorización del Banco Davivienda S.A

// var parametros, indicesColumnas;
// /**
// * Crea un menu UI en la hoja de calculo
// */
// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu("ArtBot Onboarding")
//     // .addItem("Importar datos Taleo", 'copyInfoFromInsumoToPlano')
//     .addItem("Leer correo archivos funcionario", 'leerCorreo')
//     .addItem("Leer correo archivos responsable", 'leerCorreoConsultor')
//     .addItem("Leer correo Usuario y Red", 'leerCorreoUsuarioRed')
//     .addItem("Leer correo Reporte Taleo", 'leerCorreoBaseTaleo')
//     .addItem("Leer correo Reporte Requisisiones Taleo", 'leerCorreoHistoricoRequisisionesTaleo')


//     .addSeparator()
//     .addItem("Registrar respuesta Formulario Kit", 'registrarFormularioKit')
//     .addSeparator()
//     .addItem("Instalar ArtBot", 'instalar')
//     .addToUi();
// }

// function instalar() {
//   parametros = obtenerParametros()
//   desinstalar();

//   ScriptApp.newTrigger("copyInfoFromInsumoToPlano").timeBased().everyDays(1).atHour(20).create();
//   ScriptApp.newTrigger("leerCorreo").timeBased().everyHours(1).create();
//   ScriptApp.newTrigger("leerCorreoConsultor").timeBased().everyHours(1).create();
//   ScriptApp.newTrigger("registrarFormularioKit").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onFormSubmit().create();
//   ScriptApp.newTrigger("leerCorreoConsultor").timeBased().everyDays(1).atHour(23).create();
//   ScriptApp.newTrigger("envioEmailCumplimientoCursos").timeBased().everyDays(1).atHour(8).create();
//   ScriptApp.newTrigger("envioNotificacionesEncuestaYNPS").timeBased().everyDays(1).atHour(8).create();
//   ScriptApp.newTrigger("encontrarPorcentajesCursos").timeBased().everyDays(1).atHour(21).create();
//   ScriptApp.newTrigger("leerCorreoUsuarioRed").timeBased().everyHours(1).create()
//   //   // ScriptApp.newTrigger("leerCorreoBaseTaleo").timeBased().everyDays(1).atHour(20).create();
//   ScriptApp.newTrigger("encontrarRolesInstalar").timeBased().everyDays(1).atHour(1).create();
//   ScriptApp.newTrigger("leerCorreoHistoricoRequisisionesTaleo").timeBased().everyDays(1).atHour(21).create();

//   ScriptApp.newTrigger("leerCorreoBaseTaleo").timeBased().everyHours(1).create();

//   ScriptApp.newTrigger("gestionAprendiz").timeBased().everyDays(1).atHour(8).create();
//   ScriptApp.newTrigger("envioCorreoPorcentajeCursos").timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(15).create();
//   ScriptApp.newTrigger("envioCorreoReContratacionesAlistamiento").timeBased().everyDays(1).atHour(8).create();
//   ScriptApp.newTrigger("identifiacionAprendiz").timeBased().everyDays(1).atHour(8).create();

//   //ScriptApp.newTrigger("realizarBackUp").timeBased().everyDays(1).atHour(23).create();
//   Browser.msgBox("Instalación Completada!", "Todos los activadores se crearon exitosamente", Browser.Buttons.OK);
// }

// function desinstalar() {
//   var triggers = ScriptApp.getProjectTriggers();
//   for (var i in triggers) {
//     ScriptApp.deleteTrigger(triggers[i]);
//   }
// }

// function instalarback() {
//   ScriptApp.newTrigger("realizarBackUp").timeBased().everyDays(1).atHour(23).create();
// }



// /**
// * Procesa la solicitud HTTP GET proveniente de la página HTML
// * @param {Event} e Contiene datos de la solicitud HTTP GET
// * @return HTMLOutput 
// */
// function doGet(e) {
//   parametros = obtenerParametros();
//   var template = HtmlService.createTemplateFromFile("Onboarding-start");
//   for (var iContratacion = 0; iContratacion < 6; iContratacion++) {
//     template["nota" + (iContratacion + 1) + "Contratacion"] = obtenerParametro("nota" + (iContratacion + 1) + "Contratacion");
//   }
//   for (var iAlistamiento = 0; iAlistamiento < 4; iAlistamiento++) {
//     template["nota" + (iAlistamiento + 1) + "Alistamiento"] = obtenerParametro("nota" + (iAlistamiento + 1) + "Alistamiento");
//   }
//   for (var iAprendizaje = 0; iAprendizaje < 3; iAprendizaje++) {
//     template["nota" + (iAprendizaje + 1) + "Aprendizaje"] = obtenerParametro("nota" + (iAprendizaje + 1) + "Aprendizaje");
//   }
//   for (var iCarnetizacion = 0; iCarnetizacion < 3; iCarnetizacion++) {
//     template["nota" + (iCarnetizacion + 1) + "Carnetizacion"] = obtenerParametro("nota" + (iCarnetizacion + 1) + "Carnetizacion");
//   }
//   template["notaCasoDesistido"] = obtenerParametro("notaCasoDesistido");
//   template["notaEnivoCarnetizacion"] = obtenerParametro("notaEnivoCarnetizacion");

//   template = template.evaluate()
//     .setTitle("Onboarding")
//     //.setFaviconUrl("https://i.ibb.co/TgCNRZj/davivienda-icon.png")
//     .setFaviconUrl("https://i.ibb.co/gyhHKfh/minuaturaDav.png")
//     .addMetaTag('viewport', 'width=device-width, initial-scale=1')

//   return template;
// }

// /**
// * Obtiene el contenido HTML del archivo especificado
// * @param {String} filename Nombre del archivo
// * @return {Stirng} Contenido HTML de un archivo especificado dentro del Proyecto App Script
// */
// function include(filename) {
//   return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
// }

// /**
// * Inicializa la vista web con las listas y llaves
// * @return{Object} Contiene los datos a cargar en la vista web
// */
// function inicializarWeb(area) {
//   var registros = obtenerRegistros(area);
//   return { registros: registros };
// }

// /**
// * Obtiene los indices de las columnas del consolidado
// */
// function obtenerIndicesColumnas() {

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var ultimaColumna = hojaConsolidado.getLastColumn();
//   var hojaParametros = parametros.spsParametros.getSheetByName("Parametros");
//   var listaParametros = hojaParametros.getRange("A1:A" + hojaParametros.getLastRow()).getValues().toString().split(",");
//   var filaParametros = listaParametros.indexOf("Columnas Consolidado [Ordenadas]") + 1;



//   var nombreCol, propiedad, numeroCol;
//   var indicesConsolidado = {};

//   for (var col = 1; col <= ultimaColumna; col++, filaParametros++) {
//     if (!parametros.datosParametros[filaParametros]) break;
//     propiedad = parametros.datosParametros[filaParametros][0];
//     indicesConsolidado[propiedad] = col;

//   }

//   return {
//     consolidado: indicesConsolidado
//   };

// }

// /*
// * Obtiene el archivo de parametros
// * @return {Object} Datos de paramatros
// * @throw {Error} Reporta que no se pudo encontrar el archivo Parametros
// */
// function obtenerParametros() {

//   var idArchivoParametros = CacheService.getScriptCache().get("idArchivoParametros");
//   var archivoParametros, spsParametros;

//   if (!idArchivoParametros) {

//     var archivo = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
//     var carpeta;
//     try {

//       carpeta = archivo.getParents().next();
//       archivoParametros = carpeta.searchFolders("title contains 'Fuentes'").next().searchFiles("title contains 'Parametros'").next();
//       CacheService.getScriptCache().put("idArchivoParametros", archivoParametros.getId());
//       spsParametros = SpreadsheetApp.open(archivoParametros);

//     } catch (error) {

//       throw "No se pudo obtener el archivo \"Parametros\". Verifique el nombre de las carpetas y del archivo. " + error;

//     }
//   } else {

//     spsParametros = SpreadsheetApp.openById(idArchivoParametros);

//   }

//   var sheetParametros = spsParametros.getSheetByName("Parametros");
//   var datosParametrosArr = sheetParametros.getDataRange().getValues();
//   var datosParametros = datosParametrosArr;

//   return {
//     spsParametros: spsParametros,
//     datosParametros: datosParametros
//   };
// }


// /**
// * Obtiene el valor de un parámetro de la hoja de parámetros
// * @param {String} parametro Nombre del parámetro a buscar
// * @return {String} Valor del parámetro buscado
// * @throw {Error} Reporta que el parámetro no se encontró
// */
// function obtenerParametro(parametro) {

//   var valor = CigoApp.buscarPorLlave(parametro, parametros.datosParametros, 1, 2);
//   if (valor == null) {
//     throw "El parámetro " + parametro + " no se encuentra en la hoja de parámetros";
//   }

//   return valor;
// }

// /**
// * Obtiene los emails que son permitidos para Onboarding
// */
// function obtenerEmailsPermitidos(area) {

//   var datosAreas = parametros.spsParametros.getSheetByName("Permisos").getDataRange().getValues();
//   var correosPermitidos = CigoApp.buscarPorLlave(area, datosAreas, 1, -1);
//   if (!correosPermitidos) throw "No se encontró el area " + area;
//   var correoUsuario = Session.getActiveUser().getEmail();
//   if (correosPermitidos[1].indexOf(correoUsuario) == -1) throw "No tienes permisos para ingresar a la herramienta";

// }

// /**
// * Obtener información de la pestaña Hoja 1 de la Sheet Insumo TALEO
// * para insertar información en la pestaña Definitiva de la Sheet Plano
// */
// function copyInfoFromInsumoToPlano() {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas()
  
//   var hojaConsolidado = SpreadsheetApp.openById(obtenerParametro("idHojaInsumo")).getSheetByName(obtenerParametro("nombreHojaTaleo"));
//   var datosSps = hojaConsolidado.getDataRange().getValues();
//   var hojaDestino = SpreadsheetApp.openById(obtenerParametro("idHojaPlano")).getSheetByName("Definitiva");
//   var registros = [];
//   var registrosAgregar = [];

//   for (var fila = 1; fila < datosSps.length; fila++) {

//     if (datosSps[fila][55] != "X") {

//       var id = "ONB" + new Date().getTime();
//       Utilities.sleep(1);
//       var registroFila = [id, datosSps[fila][0], "Sin Gestion", "Sin Gestion", "", "", "Sin Gestion", "", "", "Sin Gestion", "", "", datosSps[fila][1], datosSps[fila][2], datosSps[fila][3], datosSps[fila][4], datosSps[fila][5], datosSps[fila][6], datosSps[fila][7], datosSps[fila][8], datosSps[fila][9], datosSps[fila][10], datosSps[fila][11], datosSps[fila][12], datosSps[fila][13], datosSps[fila][14], datosSps[fila][15], datosSps[fila][16], datosSps[fila][17], datosSps[fila][18], datosSps[fila][19], datosSps[fila][20], datosSps[fila][21], datosSps[fila][22], datosSps[fila][23], datosSps[fila][24], datosSps[fila][25], datosSps[fila][26], datosSps[fila][27], datosSps[fila][indicesColumnas.consolidado.catCargo - 1], datosSps[fila][29], datosSps[fila][30], datosSps[fila][46], datosSps[fila][48], "", datosSps[fila][31], datosSps[fila][32], datosSps[fila][33], datosSps[fila][34], datosSps[fila][35], datosSps[fila][36], datosSps[fila][37], datosSps[fila][38], datosSps[fila][39], datosSps[fila][40], datosSps[fila][41], datosSps[fila][42], datosSps[fila][43], datosSps[fila][44], datosSps[fila][45], "", "X", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", datosSps[fila][indicesColumnas.consolidado.correo - 1], "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", datosSps[fila][54], "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "113", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", datosSps[fila][49], datosSps[fila][50], datosSps[fila][51], datosSps[fila][52], "", "", "", ""]
//       registrosAgregar.push(registroFila);

//       hojaConsolidado.getRange(fila + 1, 56).setValue("X");

//     }
//   }

//   if (registrosAgregar.length == 0) return;
//   var ultimaFila = hojaDestino.getLastRow()
//   hojaDestino.getRange(ultimaFila + 1, 1, registrosAgregar.length, registrosAgregar[0].length).setValues(registrosAgregar);

//   //Copia las formulas de las columnas especificadas en la fila del nuevo registro
//   var columnasFormulas = obtenerParametro("columnasFormulas").toString().split(",");

//   if (columnasFormulas != "") {
//     var rangoFormula;
//     for (var i = 0; i < columnasFormulas.length; i++) {
//       rangoFormula = hojaDestino.getRange((columnasFormulas[i].trim()) + 2);
//       var numeroColumna = CigoApp.obtenerNumeroColumna(columnasFormulas[i].trim());
//       // rangoFormula.copyTo(hojaConsolidado.getRange((columnasFormulas[i].trim()) + ultimaFila ));
//       rangoFormula.copyTo(hojaDestino.getRange(ultimaFila + 1, numeroColumna, registrosAgregar.length, 1));
//     }
//   }

// }

// const nombreIndicesColumnasArea = {
//   "Contratacion": { indiceGestion: "gestionContratacion", indiceFechaGestion: "horaGestionCont", indiceGestionSucursal: 2 },
//   "Aprendizaje": { indiceGestion: "gestionAprendizaje", indiceFechaGestion: "horaGestionApr", indiceGestionSucursal: 3 },
//   "Carnetizacion": { indiceGestion: "gestionCarnetizacion", indiceFechaGestion: "horaGestionCarnetizacion", indiceGestionSucursal: 5 },
//   "Alistamiento": { indiceGestion: "gestionAlistamiento", indiceFechaGestion: "horaGestionAlist", indiceGestionSucursal: 7 },
// }

// /**
// * Inicializa la vista web de acuerdo al estado especificado en los parametros de la URL
// * @param {Object} Datos provenientes de la vista Web
// * @return {Object} Contiene los datos a cargar en la vista web
// */
// function inicializarWebGestion(area) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var datosConsolidado = hojaConsolidado.getDataRange().getValues();
//   var hojaSucursales = sps.getSheetByName(obtenerParametro("nombreHojaSucursales"));
//   var datosSucursales = hojaSucursales.getDataRange().getValues();
//   var urlPaginaHome = obtenerParametro("urlPaginaHome");

//   var datosColumnasTotal = parametros.spsParametros.getSheetByName("Columnas Visualizar").getDataRange().getValues();

//   var datosColumnas, datosColumnasSeleccion;
//   if (area == "Lider") {
//     datosColumnas = datosColumnasTotal.filter(function (r) {
//       return r[0] == "lider";
//     });

//     datosColumnasSeleccion = datosColumnasTotal.filter(function (r) {
//       return r[0] == "seleccion";
//     });

//     for (var i in datosColumnasSeleccion) {
//       datosColumnasSeleccion[i][1] = CigoApp.obtenerNumeroColumna(datosColumnasSeleccion[i][1]);
//     }

//   } else {
//     datosColumnas = datosColumnasTotal.filter(function (r) {
//       return r[0] == "gestion";
//     });
//   }

//   for (var i in datosColumnas) {
//     datosColumnas[i][1] = CigoApp.obtenerNumeroColumna(datosColumnas[i][1]);
//   }

//   var registrosWeb = [];
//   // var registrosFuncionarios = [];
//   var correoUsuario = Session.getActiveUser().getEmail();
//   var estado, estadoContratacion;
//   var codigosSucursal = "";

//   obtenerEmailsPermitidos(area);

//   const infoIndices = nombreIndicesColumnasArea[area]

//   if (!infoIndices) throw "No se encuentra el area " + infoIndices
//   let columna = indicesColumnas.consolidado[infoIndices.indiceGestion];
//   let columnaFechaGestion = indicesColumnas.consolidado[infoIndices.indiceFechaGestion]

//   for (let filaSuc in datosSucursales) {
//     var resultado = datosSucursales[filaSuc][infoIndices.indiceGestionSucursal].indexOf(correoUsuario);
//     if (resultado >= 0) {
//       var sucursal = datosSucursales[filaSuc][0];
//       codigosSucursal = sucursal + "," + codigosSucursal;
//     }
//   }

//   // console.log(codigosSucursal)
//   // var resultado = datosSucursales[filaSuc][infoIndices.indiceGestionSucursal].indexOf(correoUsuario)


//   // switch (area) {
//   //   case "Contratacion":

//   //     for (var filaSuc in datosSucursales) {
//   //       var resultado = datosSucursales[filaSuc][2].indexOf(correoUsuario);
//   //       if (resultado >= 0) {
//   //         var sucursal = datosSucursales[filaSuc][0];
//   //         codigosSucursal = sucursal + "," + codigosSucursal;
//   //       }
//   //     }
//   //     break;
//   //   case "Carnetizacion":
//   //     columna = indicesColumnas.consolidado.gestionCarnetizacion;
//   //     columnaFechaGestion = indicesColumnas.consolidado.horaGestionCarnetizacion;
//   //     break;
//   //   case "Alistamiento":
//   //     columna = indicesColumnas.consolidado.gestionAlistamiento;
//   //     columnaFechaGestion = indicesColumnas.consolidado.horaGestionAlist;
//   //     break;
//   //   case "Aprendizaje":
//   //     columna = indicesColumnas.consolidado.gestionAprendizaje;
//   //     columnaFechaGestion = indicesColumnas.consolidado.horaGestionApr;
//   //     for (var filaSuc in datosSucursales) {
//   //       var resultado = datosSucursales[filaSuc][3].indexOf(correoUsuario);
//   //       if (resultado >= 0) {
//   //         var sucursal = datosSucursales[filaSuc][0];
//   //         codigosSucursal = sucursal + "," + codigosSucursal;
//   //       }
//   //     }
//   //     break;
//   //   case "Lider":
//   //     break;
//   //   default:
//   //     throw "No se encontró la columna gestion para el area: " + area;
//   //     break;
//   // }

//   var fechaCorte = new Date();
//   fechaCorte.setDate(fechaCorte.getDate() - 90);

//   for (var fila in datosConsolidado) {
//     if (area == "Lider") {
//       columnaLider = indicesColumnas.consolidado.correoLider;
//       if (datosConsolidado[fila][columnaLider - 1] == correoUsuario && datosConsolidado[fila][indicesColumnas.consolidado.estadoGeneral - 1] != "Desistido") {
//         registrosWeb.push(datosConsolidado[fila]);
//       }
//     } else {
//       estado = datosConsolidado[fila][columna - 1]
//       estado = estado.toString().toUpperCase() == "NO APLICA" ? "Gestionado" : datosConsolidado[fila][columna - 1]
//       fechaHoraGestion = datosConsolidado[fila][columnaFechaGestion - 1];
//       var resultado = codigosSucursal.indexOf(datosConsolidado[fila][indicesColumnas.consolidado.codigoSucursal - 1])

//       if (resultado == -1) continue

//       if (area == "Contratacion") {
//         if (estado == "Sin Gestion" || estado == "Pendiente") {
//           // var resultado = codigosSucursal.indexOf(datosConsolidado[fila][indicesColumnas.consolidado.codigoSucursal - 1]);
//           // if (resultado >= 0) {
//           registrosWeb.push(datosConsolidado[fila]);
//           // }
//         } else if (estado == "Gestionado" && (fechaHoraGestion >= fechaCorte)) {
//           // var resultado = codigosSucursal.indexOf(datosConsolidado[fila][indicesColumnas.consolidado.codigoSucursal - 1]);
//           // if (resultado >= 0) {
//           registrosWeb.push(datosConsolidado[fila]);
//           // }
//         }
//       } else {
//         estadoContratacion = datosConsolidado[fila][indicesColumnas.consolidado.gestionContratacion - 1];
//         estadoContratacion = estadoContratacion == "No Aplica" ? "Gestionado" : estadoContratacion
//         var reContratacion = datosConsolidado[fila][indicesColumnas.consolidado.reContratacion - 1];

//         if (estadoContratacion == "Gestionado" && (estado == "Pendiente" || estado == "Sin Gestion" || estado == "Pendiente Gestión")) {
//           if ((area == "Aprendizaje" && reContratacion.toString().toUpperCase() != "SI" && codigosSucursal && codigosSucursal.indexOf(datosConsolidado[fila][42]) != -1 && estado != "Sin Gestion") || area != "Aprendizaje")

//             registrosWeb.push(datosConsolidado[fila]);

//         } else if (estadoContratacion == "Gestionado" && estado == "Gestionado" && fechaHoraGestion >= fechaCorte) {
//           if ((area == "Aprendizaje" && reContratacion.toString().toUpperCase() != "SI" && codigosSucursal && codigosSucursal.indexOf(datosConsolidado[fila][42]) != -1 && estado != "Sin Gestion") || area != "Aprendizaje")

//             registrosWeb.push(datosConsolidado[fila]);

//         } else if ((area == "Carnetizacion") && (reContratacion.toString().toUpperCase() != "SI") && (estado == "Pendiente" || estado == "Pendiente Gestión")) {

//           registrosWeb.push(datosConsolidado[fila]);

//         } else if ((area == "Carnetizacion") && reContratacion.toString().toUpperCase() != "SI" && estado == "Gestionado" && fechaHoraGestion >= fechaCorte) {

//           registrosWeb.push(datosConsolidado[fila]);

//         } else if (area == "Alistamiento" && (estado == "Pendiente" || estado == "Pendiente Gestión" || estado == "Gestionado")) {

//           registrosWeb.push(datosConsolidado[fila]);
//         } else if (area == "Aprendizaje" && reContratacion.toString().toUpperCase() != "SI" && codigosSucursal && codigosSucursal.indexOf(datosConsolidado[fila][42]) != -1 && (estado == "Pendiente" || estado == "Pendiente Gestión" || estado == "Gestionado")) {

//           registrosWeb.push(datosConsolidado[fila]);
//         }
//       }
//     }
//   }

//   for (var f in registrosWeb) {
//     var valorFechaContratacion = registrosWeb[f][indicesColumnas.consolidado.citaContratacion - 1].toString()
//     // var valorFechaExpe = registrosWeb[f][indicesColumnas.consolidado.fechaExpCedula - 1].toString()

//     for (var c = 0; c < registrosWeb[f].length; c++) {
//       registrosWeb[f][c] = formatearFecha(registrosWeb[f][c]);
//     }
//     registrosWeb[f][indicesColumnas.consolidado.citaContratacion - 1] = valorFechaContratacion
//     // registrosWeb[f][indicesColumnas.consolidado.fechaExpCedula - 1] = valorFechaExpe
//   }

//   var encabezados = datosConsolidado[0];

//   return {
//     registros: registrosWeb,
//     encabezados: encabezados,
//     datosVistas: datosColumnas,
//     indicesColumnas: indicesColumnas,
//     listas: obtenerListas(),
//     urlPaginaHome: urlPaginaHome,

//     datosVistasSeleccion: datosColumnasSeleccion
//   };
// }



// function formatearFecha(fecha, zonaHoraria, formato) {
//   zonaHoraria = zonaHoraria || "America/Bogota";
//   formato = formato || "dd/MM/yyyy";

//   if (fecha instanceof Date) return Utilities.formatDate(fecha, zonaHoraria, formato);
//   return fecha;
// }


// /**
// * Metodo encargado de realizar la inserción de la información editable para cada
// * una de las areas que se pueden logear a la aplicación
// * @param {Object} formulario proveniente de la vista Web
// * @throw {Error} Reporta que el id no se encuentra en el origen de datos
// */
// function registrarFormularioGestion(formulario) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + formulario.id;

//   var fechaActual = new Date();

//   var registroAnterior = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];

//   if (formulario.gestionado == "Gestionado") {
//     validacionBotones(formulario, registroAnterior)
//   }

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   if (formulario.gestionado == "No aplica") {
//     const indiceColumnaArea = nombreIndicesColumnasArea[formulario.area].indiceGestion
//     const indiceFechaGestion = nombreIndicesColumnasArea[formulario.area].indiceFechaGestion
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado[indiceColumnaArea]).setValue("No aplica")
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado[indiceFechaGestion]).setValue(new Date())
//     return
//   }

//   if (registroAnterior[indicesColumnas.consolidado.fotoCorrecta - 1] != "Si" && formulario.fotoCorrecta == "Si") {
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["fechaFotoCorrecta"]).setValue(fechaActual);
//   }

//   var columnaGestion, columnaFechaGestion, columnaCorreoGestion, correoNotificacion;
//   switch (formulario.area) {
//     case "Contratacion":
//       columnaGestion = indicesColumnas.consolidado.gestionContratacion;
//       columnaFechaGestion = indicesColumnas.consolidado.horaGestionCont;
//       columnaCorreoGestion = indicesColumnas.consolidado.correoGestionCont;
//       correosNotificacion = obtenerParametro("correosNotificacionContratacion");

//       if (registroAnterior[indicesColumnas.consolidado.fechaInicioAlistamiento - 1] == "" && formulario.gestionado == "Gestionado") {
//         hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaInicioAlistamiento).setValue(fechaActual);
//         hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.gestionAlistamiento).setValue("Pendiente Gestión");
//       }

//       if (registroAnterior[indicesColumnas.consolidado.fechaInicioCarnetizacion - 1] == "" && formulario.gestionado == "Gestionado") {
//         hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaInicioCarnetizacion).setValue(fechaActual);
//         hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.gestionCarnetizacion).setValue("Pendiente Gestión");
//       }
//       break;
//     case "Carnetizacion":
//       columnaGestion = indicesColumnas.consolidado.gestionCarnetizacion;
//       columnaFechaGestion = indicesColumnas.consolidado.horaGestionCarnetizacion;
//       columnaCorreoGestion = indicesColumnas.consolidado.correoGestionCarnetizacion;
//       break;
//     case "Alistamiento":
//       columnaGestion = indicesColumnas.consolidado.gestionAlistamiento;
//       columnaFechaGestion = indicesColumnas.consolidado.horaGestionAlist;
//       columnaCorreoGestion = indicesColumnas.consolidado.correoGestionAlist;
//       correosNotificacion = obtenerParametro("correosNotificacionAlistamiento");

//       if (registroAnterior[indicesColumnas.consolidado.fechaInicioAprendizaje - 1] == "" && formulario.gestionado == "Gestionado") {
//         hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaInicioAprendizaje).setValue(fechaActual);
//         if (registroAnterior[indicesColumnas.consolidado.horaGestionApr - 1] == "") {
//           hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.gestionAprendizaje).setValue("Pendiente Gestión");
//         }
//       }

//       break;
//     case "Aprendizaje":
//       columnaGestion = indicesColumnas.consolidado.gestionAprendizaje;
//       columnaFechaGestion = indicesColumnas.consolidado.horaGestionApr;
//       columnaCorreoGestion = indicesColumnas.consolidado.correoGestionApr;
//       break;
//   }

//   if (registroAnterior[indicesColumnas.consolidado.usuarioXplora - 1] != "" && registroAnterior[indicesColumnas.consolidado.fechaXplora - 1] == "") {
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaXplora).setValue(fechaActual);
//   }

//   if (registroAnterior[indicesColumnas.consolidado.fechaInicioCarnetizacion - 1] == "" &&
//     formulario.area == "Contratación" && formulario.gestionado == "Gestionado") {
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaInicioCarnetizacion).setValue(fechaActual);
//   }

//   if (formulario.gestionado != "Gestionado" && formulario.gestionado != "Pendiente" && formulario.gestionado != "Gestionar") {
//     formulario.gestionado = "Pendiente";
//     hojaConsolidado.getRange(filaConsolidado, columnaGestion).setValue(formulario.gestionado);
//   } else {
//     hojaConsolidado.getRange(filaConsolidado, columnaGestion).setValue(formulario.gestionado);
//   }

//   hojaConsolidado.getRange(filaConsolidado, columnaFechaGestion).setValue(fechaActual)
//   hojaConsolidado.getRange(filaConsolidado, columnaCorreoGestion).setValue(Session.getActiveUser().getEmail());

// }

// function guardarFormulario(fila, hoja, formulario) {
//   for (var nombreCampo in formulario) {
//     if (!indicesColumnas.consolidado[nombreCampo]) continue;
//     if (!formulario[nombreCampo]) continue;
//     hoja.getRange(fila, indicesColumnas.consolidado[nombreCampo]).setValue(formulario[nombreCampo]);
//   }
// }

// function validacionBotones(formulario, fila) {
//   var mensajeErroneo = ""
//   switch (formulario.area) {
//     case "Contratacion":
//       var envioCarnetizacion = fila[indicesColumnas.consolidado.envioCarnetizacion - 1]
//       var fechaFirmContr = fila[indicesColumnas.consolidado.fechaFirmContr - 1]
//       var gestionAlistamiento = fila[indicesColumnas.consolidado.gestionAlistamiento - 1]
//       var envioCorreoLiderPadrino = fila[indicesColumnas.consolidado.envioCorreoLiderPadrino - 1]
//       if (envioCarnetizacion != "Si") {
//         mensajeErroneo += "<p> Falta dar clic en el botón enviar de la sección 2 </p>"
//       }
//       if (fechaFirmContr == "") {
//         mensajeErroneo += "<p> Falta dar clic en el botón enviar de la sección 4 </p>"
//       }
//       if (gestionAlistamiento == "") {
//         mensajeErroneo += "<p> Falta dar clic en el botón enviar alistamiento de la sección 4</p>"
//       }
//       if (envioCorreoLiderPadrino != "Si") {
//         mensajeErroneo += "<p> Falta dar clic en el botón enviar de la sección 5 </p>"
//       }
//       break;
//     case "Alistamiento":
//       var envioCorreoasuntoNotificacionPasoFlujo = fila[indicesColumnas.consolidado.envioCorreoasuntoNotificacionPasoFlujo - 1]
//       var masivos = fila[indicesColumnas.consolidado.masivos - 1]
//       if (envioCorreoasuntoNotificacionPasoFlujo != "Si" && masivos == "No") {
//         mensajeErroneo += "<p> Falta dar clic en el botón enviar notificación al usuario y activar aprendizaje de la sección 2 </p>"
//       }

//       break;
//     case "Aprendizaje":
//       var envioCorreoAprendizajeCargado = fila[indicesColumnas.consolidado.envioCorreoAprendizajeCargado - 1]
//       if (envioCorreoAprendizajeCargado != "Si") {
//         mensajeErroneo += "<p> Falta dar clic en el botón enviar de la sección 2 </p>"
//       };
//       break;
//   }
//   if (mensajeErroneo.length > 0) {
//     throw mensajeErroneo
//   }
// }

// // function enviarCorreoAccesos(id) {

// //   parametros = obtenerParametros();
// //   indicesColumnas = obtenerIndicesColumnas();

// //   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
// //   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
// //   var idsConsolidado = hojaConsolidado.getRange(1,indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
// //   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado,1);

// //   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
// //   var correo = infoFila[indicesColumnas.consolidado.correoLider-1];

// //   if (!correo) throw "El correo del Lider no está definido";
// //   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

// //   var asuntoCorreosOK = CigoApp.reemplazarValores(obtenerParametro("asuntoCorreosOK"), infoFila);
// //   var asuntoMensajesOK = CigoApp.reemplazarValores(obtenerParametro("mensajeCorreoOK"), infoFila);

// //   CigoApp.enviarCorreo(correo, asuntoCorreosOK, asuntoMensajesOK, {});

// // }

// function enviarCorreoEntregaKit(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   // var asuntoKitMicroinformatica = obtenerComunicacionPorFilial("asuntoKitMicroinformatica", filial)
//   // asuntoKitMicroinformatica = CigoApp.reemplazarValores(asuntoKitMicroinformatica, infoFila);

//   // var mensajeKitMicroinformatica = obtenerComunicacionPorFilial("mensajeKitMicroinformatica", filial)
//   // mensajeKitMicroinformatica = CigoApp.reemplazarValores(mensajeKitMicroinformatica, infoFila);

//   // var inlineImages = obtenerImagenesCorreo("mensajeKitMicroinformatica", filial);

//   // CigoApp.enviarCorreo(correo, asuntoKitMicroinformatica, mensajeKitMicroinformatica, {});
//   // enviarCorreo(correo, asuntoKitMicroinformatica, mensajeKitMicroinformatica, {inlineImages:inlineImages});
// }

// function enviarCorreoInduccion(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;


//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoCitacionInduccion = obtenerComunicacionPorFilial("asuntoCitacionInduccion", filial)
//   asuntoCitacionInduccion = CigoApp.reemplazarValores(asuntoCitacionInduccion, infoFila);

//   var mensajeCitacionInduccion = obtenerComunicacionPorFilial("mensajeCitacionInduccion", filial)
//   mensajeCitacionInduccion = CigoApp.reemplazarValores(mensajeCitacionInduccion, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeCitacionInduccion", filial);

//   // CigoApp.enviarCorreo(correo, asuntoCitacionInduccion, mensajeCitacionInduccion, {});
//   // enviarCorreo(correo, asuntoCitacionInduccion, mensajeCitacionInduccion, {inlineImages:inlineImages});

// }

// function enviarCorreoInduccionFinalizada(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoInduccionFinalizada = obtenerComunicacionPorFilial("asuntoInduccionFinalizada", filial)
//   asuntoInduccionFinalizada = CigoApp.reemplazarValores(asuntoInduccionFinalizada, infoFila);

//   var mensajeInduccionFinalizada = CigoApp.reemplazarValores(obtenerParametro("mensajeInduccionFinalizada"), infoFila);
//   mensajeInduccionFinalizada = CigoApp.reemplazarValores(mensajeInduccionFinalizada, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeInduccionFinalizada", filial);

//   // var archivoAdjunto = DriveApp.getFileById(CigoApp.obtenerIdUrl(obtenerParametro("urlArchivoProtocoloInduccion")));
//   // var blobArchivo = archivoAdjunto.getBlob();


//   // CigoApp.enviarCorreo(correo, asuntoInduccionFinalizada, mensajeInduccionFinalizada, {attachments:[blobArchivo]});
//   // enviarCorreo(correo, asuntoInduccionFinalizada, mensajeInduccionFinalizada, {attachments:[blobArchivo], inlineImages: inlineImages});
// }

// function cargueDocExitoso(formulario) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + formulario.id;

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];
//   var genero = infoFila[indicesColumnas.consolidado.genero - 1];
//   var tipoContrato = infoFila[indicesColumnas.consolidado.tipoContrato - 1];
//   var tipoSalario = infoFila[indicesColumnas.consolidado.tipoSalario - 1];
//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   if (!correo) throw "El correo externo no está definido";


//   var checksNoValidar = [];
//   switch (tipoContrato) {
//     case "Aprendiz sena":
//     case "Practicante universitario":
//       checksNoValidar = ["checkCerFondoPensiones", "checkCerFondoCesantias", "checkActaPregrado", "checkDiplomaPregrado", "checkTarjetaProfesional"];
//       break;
//     case "Profesional en práctica":
//     case "Prestación de servicios":
//       checksNoValidar = ["checkActaPregrado", "checkDiplomaPregrado", "checkTarjetaProfesional"];
//       break;
//     case "Término fijo":
//       checksNoValidar = [];
//       break;
//     case "Término indefinido":
//       if (tipoSalario == "Integral") checksNoValidar = ["checkCerFondoCesantias"];
//       break;
//   }

//   if (formulario.documCompletos == "NO") {

//     var documentosFaltantes = [];

//     if (formulario.checkAfiliacionEPS != "on" && checksNoValidar.indexOf("checkAfiliacionEPS") == -1)
//       documentosFaltantes.push("Certificado de afiliación EPS");

//     if (formulario.checkConsAfilADRES != "on" && checksNoValidar.indexOf("checkConsAfilADRES") == -1)
//       documentosFaltantes.push("Constancia de afiliación ADRES");

//     if (formulario.checkCerFondoPensiones != "on" && checksNoValidar.indexOf("checkCerFondoPensiones") == -1)
//       documentosFaltantes.push("Certificado Fondo de Pensiones");

//     if (formulario.checkCerFondoCesantias != "on" && checksNoValidar.indexOf("checkCerFondoCesantias") == -1)
//       documentosFaltantes.push("Certificado Fondo de Cesantías");

//     if (formulario.checkDocIdentidad150 != "on" && checksNoValidar.indexOf("checkDocIdentidad150") == -1)
//       documentosFaltantes.push("Documento de Identidad al 150%");

//     if (formulario.checkLibretaMilitar != "on" && genero == "Hombre" && checksNoValidar.indexOf("checkLibretaMilitar") == -1)
//       documentosFaltantes.push("Fotocopia de la Libreta Militar al 150% o certificación provisional");

//     if (formulario.checkDiplomaActaBachiller != "on" && checksNoValidar.indexOf("checkDiplomaActaBachiller") == -1)
//       documentosFaltantes.push("Acta Grado Bachiller");

//     if (formulario.checkDiplomaBachiller != "on" && checksNoValidar.indexOf("checkDiplomaBachiller") == -1)
//       documentosFaltantes.push("Diploma Grado Bachiller");

//     if (formulario.checkActaPregrado != "on" && checksNoValidar.indexOf("checkActaPregrado") == -1)
//       documentosFaltantes.push("Acta Grado Pregrado");

//     if (formulario.checkDiplomaPregrado != "on" && checksNoValidar.indexOf("checkDiplomaPregrado") == -1)
//       documentosFaltantes.push("Diploma Grado Pregrado");

//     if (formulario.checkTarjetaProfesional != "on" && checksNoValidar.indexOf("checkTarjetaProfesional") == -1)
//       documentosFaltantes.push("Tarjeta Profesional");

//     if (formulario.checkCertLaborales != "on" && checksNoValidar.indexOf("checkCertLaborales") == -1)
//       documentosFaltantes.push("Certificados Laborales");

//     if (formulario.checkFoto3x4 != "on" && checksNoValidar.indexOf("checkFoto3x4") == -1)
//       documentosFaltantes.push("Foto de acuerdo a los lineamientos <a href='https://comunicaciones.davivienda.com/carnet-davivienda'>aquí</a>");

//     if (formulario.checkRegistroNacimiento != "on" && checksNoValidar.indexOf("checkRegistroNacimiento") == -1)
//       documentosFaltantes.push("Registro de nacimiento");

//     if (formulario.checkCertFinCasa != "on" && checksNoValidar.indexOf("checkCertFinCasa") == -1)
//       documentosFaltantes.push("Certificado del Curso Mis Finanzas en Casa");

//     if (formulario.checkCertCuentaDAMAS != "on" && checksNoValidar.indexOf("checkCertCuentaDAMAS") == -1)
//       documentosFaltantes.push("Certificado cuenta nómina DAMAS");

//     //if (formulario.checkCartaPresentacion != "on") documentosFaltantes += "Carta de Presentación, ";
//     //if (formulario.checkContratoTrabajo != "on") documentosFaltantes += "Contrato de Trabajo Firmado, ";
//     //if (formulario.checkAfiliacionSaludEPS != "on") documentosFaltantes += "Afiliación a Salud (E.P.S), ";
//     //if (formulario.tipoContrato != "Aprendiz DAV"){
//     //  if (formulario.checkFondoPensiones != "on") documentosFaltantes += "Fondo de pensiones, ";
//     //}
//     //if (formulario.checkAfiliacionCaja != "on") documentosFaltantes += "Afiliación a la Caja de Compensación Familiar, ";
//     //if (formulario.checkCartaPactoCol != "on") documentosFaltantes += "Carta Adhesión al Pacto Colectivo, ";
//     //if (formulario.checkAfiliacionFondo != "on") documentosFaltantes += "Afiliación a Fondo de Empleados. ";
//     var textoDocsFaltantes = "<ul>";
//     for (var i = 0; i < documentosFaltantes.length; i++) {
//       textoDocsFaltantes += "<li>" + documentosFaltantes[i] + "</li>";
//     }
//     textoDocsFaltantes += "</ul>";


//     var asuntoCargueErroneo = obtenerComunicacionPorFilial("asuntoCargueErroneo", filial)
//     var asuntoCargueErroneo = CigoApp.reemplazarValores(asuntoCargueErroneo, infoFila);

//     var mensajeCargueErroneo = obtenerComunicacionPorFilial("mensajeCargueErroneo", filial)
//     mensajeCargueErroneo = CigoApp.reemplazarValores(mensajeCargueErroneo, infoFila);
//     mensajeCargueErroneo = mensajeCargueErroneo.replace("[documentosFaltantes]", textoDocsFaltantes);

//     var inlineImages = obtenerImagenesCorreo("mensajeCargueErroneo", filial);

//     // CigoApp.enviarCorreo(correo, asuntoCargueErroneo, mensajeCargueErroneo, {});
//     enviarCorreo(correo, asuntoCargueErroneo, mensajeCargueErroneo, { inlineImages: inlineImages });
//     registrarLogCorreo(infoFila, correo, asuntoCargueErroneo);

//   } else {

//     // var asuntoAsignacionContrato = CigoApp.reemplazarValores(obtenerParametro("asuntoAsignacionContrato"), infoFila);
//     // var mensajeAsignacionContrato = CigoApp.reemplazarValores(obtenerParametro("mensajeAsignacionContrato"), infoFila);

//     // CigoApp.enviarCorreo(correo, asuntoAsignacionContrato, mensajeAsignacionContrato, {});

//     envioArchivoProveedor(formulario, infoFila, filaConsolidado, hojaConsolidado);
//   }


//   const nombreCompañia = !abreviacionFiliales[filial] ? filial : abreviacionFiliales[filial]

//   console.log(nombreCompañia, filial, "S")

//   if (nombreCompañia == "BETA" || nombreCompañia == "GAMMA" || nombreCompañia == "CORREDORES") {
//     let asuntoPasoFlujo = obtenerComunicacionPorFilial("asuntoNotificacionPasoFlujo", filial);
//     asuntoPasoFlujo = CigoApp.reemplazarValores(asuntoPasoFlujo, infoFila);

//     let mensajePasoFlujo = obtenerComunicacionPorFilial("mensajeNotificacionPasoFlujo", filial);
//     mensajePasoFlujo = CigoApp.reemplazarValores(mensajePasoFlujo, infoFila);

//     const inlineImages = obtenerImagenesCorreo("mensajeNotificacionPasoFlujo", filial)

//     let correosCarnetizacion = obtenerParametro(parametrosCorreoComunicacionPasoFlujo[nombreCompañia].carnetizacion)
//     correosCarnetizacion = CigoApp.reemplazarValores(correosCarnetizacion, infoFila)

//     enviarCorreo(correosCarnetizacion, asuntoPasoFlujo, mensajePasoFlujo, { inlineImages: inlineImages })
//   }
// }

// function asignacionContrato(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoAsignacionContrato = obtenerComunicacionPorFilial("asuntoAsignacionContrato", filial);
//   asuntoAsignacionContrato = CigoApp.reemplazarValores(asuntoAsignacionContrato, infoFila);

//   var mensajeAsignacionContrato = obtenerComunicacionPorFilial("mensajeAsignacionContrato", filial)
//   mensajeAsignacionContrato = CigoApp.reemplazarValores(mensajeAsignacionContrato, infoFila)

//   var archivoAdjunto = DriveApp.getFileById(CigoApp.obtenerIdUrl(obtenerParametro("urlsArchivosContrato")));
//   var blobArchivo = archivoAdjunto.getBlob();

//   var inlineImages = obtenerImagenesCorreo("mensajeAsignacionContrato", filial);

//   // CigoApp.enviarCorreo(correo, asuntoInduccionFinalizada, mensajeInduccionFinalizada, {attachments:[blobArchivo]});
//   enviarCorreo(correo, asuntoAsignacionContrato, mensajeAsignacionContrato, { attachments: [blobArchivo], inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correo, asuntoAsignacionContrato);

// }

// function firmaContrato(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   // var asuntoFirmaContrato = CigoApp.reemplazarValores(obtenerParametro("asuntoFirmaContrato"), infoFila);
//   // var mensajeFirmaContrato = CigoApp.reemplazarValores(obtenerParametro("mensajeFirmaContrato"), infoFila);

//   // var inlineImages = obtenerImagenesCorreo("mensajeFirmaContrato");

//   // CigoApp.enviarCorreo(correo, asuntoFirmaContrato, mensajeFirmaContrato, {});
//   // enviarCorreo(correo, asuntoFirmaContrato, mensajeFirmaContrato, {inlineImages: inlineImages});

// }

// function fechaIngreso(formulario) {
//   var id = formulario.id;

//   //if (!fechaFirmContr) throw "Ingresa la fecha y hora de firma del contrato";

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];
//   if (!correo) throw "El correo externo no está definido";

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoFechaIngreso = obtenerComunicacionPorFilial("asuntoFechaIngreso", filial);
//   asuntoFechaIngreso = CigoApp.reemplazarValores(asuntoFechaIngreso, infoFila);

//   var mensajeFechaIngreso = obtenerComunicacionPorFilial("mensajeFechaIngreso", filial);
//   mensajeFechaIngreso = CigoApp.reemplazarValores(mensajeFechaIngreso, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeFechaIngreso", filial);

//   // CigoApp.enviarCorreo(correo, asuntoFechaIngreso, mensajeFechaIngreso, {});
//   enviarCorreo(correo, asuntoFechaIngreso, mensajeFechaIngreso, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correo, asuntoFechaIngreso);

// }


// function registroMasivos(tabla) {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var numeroGrupos = tabla.auxContadorGrupoMasivo
//   var idsUnicos = tabla.arrayIdUnico.split(",")
//   var correoResponsable = Session.getActiveUser().getEmail()
//   for (var i = 1; i < numeroGrupos; i++) {
//     var fechaHoraAgenda = tabla["fecha-grupo-" + i]
//     if (!fechaHoraAgenda) throw "Ingresa la fecha y hora de citacion del grupo: " + i + ""

//     fechaHoraAgendaAux = fechaHoraAgenda.split("T");
//     var fechaString = fechaHoraAgendaAux[0].split("-");
//     var horaString = fechaHoraAgendaAux[1].split(":");

//     var anio = Number(fechaString[0]);
//     var mes = Number(fechaString[1]) - 1;
//     var dia = Number(fechaString[2]);

//     var hora = Number(horaString[0]);
//     var minuto = Number(horaString[1]);
//     var seg = 0;

//     var fechaInicial = new Date(anio, mes, dia, hora, minuto, seg);
//     var fechaFinal = new Date(anio, mes, dia, hora, minuto + 20, seg);

//     var fechaInicialTxt = formatearFecha(fechaFinal, "America/Bogota", "yyyy-MM-dd hh:mm a");
//     var correosInvitados = []

//     for (var j = 0; j < idsUnicos.length; j++) {
//       if (!tabla["" + idsUnicos[j] + "-" + i]) continue
//       var filaConsolidado = CigoApp.obtenerFila(idsUnicos[j], idsConsolidado, 1);
//       if (!filaConsolidado) throw "No se encontró el registro con el ID " + idsUnicos[j];
//       var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//       infoFila[indicesColumnas.consolidado.citaContratacion - 1] = fechaHoraAgenda;
//       infoFila[indicesColumnas.consolidado.respContrato - 1] = correoResponsable;

//       var sueldo = infoFila[indicesColumnas.consolidado.salario - 1].toString();
//       sueldo = sueldo.replace(/\D/g, "");
//       var sueldoTexto = NUMEROATEXTO(sueldo, "PESOS MCTE");

//       infoFila[indicesColumnas.consolidado.sueldoTexto - 1] = sueldoTexto;
//       hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.sueldoTexto).setValue(sueldoTexto);

//       hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.citaContratacion).setValue(fechaInicialTxt);
//       hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.respContrato).setValue(correoResponsable);

//       let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//       var asuntoAsignacionContrato = obtenerComunicacionPorFilial("asuntoAsignacionContrato", filial);
//       asuntoAsignacionContrato = CigoApp.reemplazarValores(asuntoAsignacionContrato, infoFila);

//       var mensajeAsignacionContrato = obtenerComunicacionPorFilial("mensajeAsignacionContrato", filial);
//       mensajeAsignacionContrato = CigoApp.reemplazarValores(mensajeAsignacionContrato, infoFila);

//       var inlineImages = {}

//       var correoFuncionario = infoFila[indicesColumnas.consolidado.correo - 1];

//       enviarCorreo(correoFuncionario, asuntoAsignacionContrato, mensajeAsignacionContrato, { inlineImages: inlineImages });

//       correosInvitados.push(correoFuncionario)
//       registrarLogCorreo(correoFuncionario, correoFuncionario, asuntoAsignacionContrato);

//       var asuntoContrato = obtenerComunicacionPorFilial("asuntoContrato", filial);
//       asuntoContrato = CigoApp.reemplazarValores(asuntoContrato, infoFila);

//       var mensajeContrato = obtenerComunicacionPorFilial("mesajeContrato", filial);
//       mensajeContrato = CigoApp.reemplazarValores(mensajeContrato, infoFila);

//       var inlineImages = obtenerImagenesCorreo("mesajeContrato", filial)

//       var archivos = combinarCampos(infoFila);
//       // var archivosDeEmail = [];

//       var inlineImages = obtenerImagenesCorreo("mesajeContrato");
//       enviarCorreo(correoResponsable, asuntoContrato, mensajeContrato, { attachments: archivos, inlineImages: inlineImages });
//       registrarLogCorreo(infoFila, correoResponsable, asuntoContrato);

//     }
//     var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioContratacion"), infoFila);
//     var descripcionEvento = obtenerParametro("descripcionEventoContratacionMasivo")

//     correosInvitados.push(correoResponsable)
//     var opciones = {
//       description: descripcionEvento,
//       sendInvites: true,
//       guests: correosInvitados.toString()
//     };

//     var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioContratacion"));
//     calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);

//     //var asunto = CigoApp.reemplazarValores(obtenerParametro("mensajeAsignacionMasivo"), infoFila);
//     //var mensaje = CigoApp.reemplazarValores(obtenerParametro("asuntoAsignacionMasivo"), infoFila);
//     //var inlineImages = obtenerImagenesCorreo("");
//     //enviarCorreo(correoResponsable, asunto, mensaje, {inlineImages: inlineImages});
//   }
// }

// //Citacion
// function citarContratacion(id, citaContratacion, correoResponsable) {
//   if (!citaContratacion) throw "Ingresa la fecha y hora de la citación";
//   if (!correoResponsable) throw "Ingresa el correo del responsable de la contratación";

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;
//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];

//   citaContratacion = citaContratacion.split("T");
//   var fechaString = citaContratacion[0].split("-");
//   var horaString = citaContratacion[1].split(":");

//   var anio = Number(fechaString[0]);
//   var mes = Number(fechaString[1]) - 1;
//   var dia = Number(fechaString[2]);

//   var hora = Number(horaString[0]);
//   var minuto = Number(horaString[1]);
//   var seg = 0;

//   var fechaInicial = new Date(anio, mes, dia, hora, minuto, seg);
//   var fechaFinal = new Date(anio, mes, dia, hora, minuto + 20, seg);

//   var fechaInicialTxt = formatearFecha(fechaFinal, "America/Bogota", "yyyy-MM-dd hh:mm a");

//   infoFila[indicesColumnas.consolidado.citaContratacion - 1] = citaContratacion;
//   infoFila[indicesColumnas.consolidado.respContrato - 1] = correoResponsable;

//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.citaContratacion).setValue(fechaInicialTxt);
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.respContrato).setValue(correoResponsable);


//   var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioContratacion"), infoFila);
//   var descripcionEvento = CigoApp.reemplazarValores(obtenerParametro("descripcionEventoContratacion"), infoFila);

//   var correosInvitacion = [
//     infoFila[indicesColumnas.consolidado.correo - 1],
//     correoResponsable
//   ];

//   var opciones = {
//     description: descripcionEvento,
//     sendInvites: true,
//     guests: correosInvitacion.toString()
//   };

//   var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioContratacion"));
//   calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoAsignacionContrato = obtenerComunicacionPorFilial("asuntoAsignacionContrato", filial);
//   asuntoAsignacionContrato = CigoApp.reemplazarValores(asuntoAsignacionContrato, infoFila);

//   var mensajeAsignacionContrato = obtenerComunicacionPorFilial("mensajeAsignacionContrato", filial);
//   mensajeAsignacionContrato = CigoApp.reemplazarValores(mensajeAsignacionContrato, infoFila);


//   var correoFuncionario = infoFila[indicesColumnas.consolidado.correo - 1];

//   var inlineImages = {}

//   // CigoApp.enviarCorreo(correoFuncionario, asuntoAsignacionContrato, mensajeAsignacionContrato, {});
//   enviarCorreo(correoFuncionario, asuntoAsignacionContrato, mensajeAsignacionContrato, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correoFuncionario, asuntoAsignacionContrato);

// }

// function hvCreadaSARH(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoFechaIngreso = obtenerComunicacionPorFilial("asuntoFechaIngreso", filial);
//   var asuntoFechaIngreso = CigoApp.reemplazarValores(asuntoFechaIngreso, infoFila);

//   var mensajeFechaIngreso = obtenerComunicacionPorFilial("mensajeFechaIngreso", filial);
//   mensajeFechaIngreso = CigoApp.reemplazarValores(mensajeFechaIngreso, infoFila);
//   // var mensajeFechaIngreso = "A este email le ahce falta el cuerpo (no se encuentra definido)";

//   var inlineImages = obtenerImagenesCorreo("mensajeFechaIngreso", filial);

//   // CigoApp.enviarCorreo(correo, asuntoFechaIngreso, mensajeFechaIngreso, {});
//   enviarCorreo(correo, asuntoFechaIngreso, mensajeFechaIngreso, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correo, asuntoFechaIngreso);

// }

// function despachoKit(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   // var asuntoDespachoKit = CigoApp.reemplazarValores(obtenerParametro("asuntoDespachoKit"), infoFila);
//   // var mensajeDespachoKit = CigoApp.reemplazarValores(obtenerParametro("mensajeDespachoKit"), infoFila);

//   // var inlineImages = obtenerImagenesCorreo("mensajeDespachoKit");

//   // CigoApp.enviarCorreo(correo, asuntoDespachoKit, mensajeDespachoKit, {});
//   // enviarCorreo(correo, asuntoDespachoKit, mensajeDespachoKit, {inlineImages: inlineImages});

// }

// function envioAlistamiento(formulario) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correo) throw "El correo del Lider no está definido";


//   // var asuntoGestionadoContratacion = CigoApp.reemplazarValores(obtenerParametro("asuntoGestionadoContratacion"), infoFila);
//   // var mensajeGestionadoContratacion = CigoApp.reemplazarValores(obtenerParametro("mensajeGestionadoContratacion"), infoFila);

//   //CigoApp.enviarCorreo(correo, asuntoGestionadoContratacion, mensajeGestionadoContratacion, {});

//   // for (var nombreCampo in formulario) {
//   //   if (!indicesColumnas.consolidado[nombreCampo]) continue;
//   //   hojaConsolidado.getRange(filaConsolidado,indicesColumnas.consolidado[nombreCampo]).setValue(formulario[nombreCampo]);
//   // }

//   var columnaGestion, columnaFechaGestion, columnaCorreoGestion;
//   if (formulario.area == "Contratacion") {
//     columnaGestion = indicesColumnas.consolidado.gestionAlistamiento;
//     columnaFechaGestion = indicesColumnas.consolidado.horaGestionAlist;
//     columnaCorreoGestion = indicesColumnas.consolidado.correoGestionAlist;

//     hojaConsolidado.getRange(filaConsolidado, columnaGestion).setValue("Pendiente Gestión")
//     // hojaConsolidado.getRange(filaConsolidado, columnaFechaGestion).setValue(new Date())
//     // hojaConsolidado.getRange(filaConsolidado, columnaCorreoGestion).setValue(Session.getActiveUser().getEmail());

//     columnaGestion = indicesColumnas.consolidado.gestionCarnetizacion;
//     columnaFechaGestion = indicesColumnas.consolidado.horaGestionCarnetizacion;
//     columnaCorreoGestion = indicesColumnas.consolidado.correoGestionCarnetizacion;

//     if (infoFila[indicesColumnas.consolidado.fechaInicioAlistamiento - 1] == "") {
//       hojaConsolidado.getRange(filaConsolidado, columnaGestion).setValue("Pendiente Gestión");
//       hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaInicioAlistamiento).setValue(new Date());
//     }
//     // hojaConsolidado.getRange(filaConsolidado, columnaFechaGestion).setValue(new Date())
//     // hojaConsolidado.getRange(filaConsolidado, columnaCorreoGestion).setValue(Session.getActiveUser().getEmail());
//   }

//   const filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   const nombreCompañia = !abreviacionFiliales[filial] ? filial : abreviacionFiliales[filial]

//   if (nombreCompañia == "BETA" || nombreCompañia == "GAMMA" || nombreCompañia == "CORREDORES") {
//     let asuntoPasoFlujo = obtenerComunicacionPorFilial("asuntoNotificacionPasoFlujo", filial);
//     asuntoPasoFlujo = CigoApp.reemplazarValores(asuntoPasoFlujo, infoFila);

//     let mensajePasoFlujo = obtenerComunicacionPorFilial("mensajeNotificacionPasoFlujo", filial);
//     mensajePasoFlujo = CigoApp.reemplazarValores(mensajePasoFlujo, infoFila);

//     const inlineImages = obtenerImagenesCorreo("mensajeNotificacionPasoFlujo", filial)

//     let correosAlistamientoGestionUsuario = obtenerParametro(parametrosCorreoComunicacionPasoFlujo[nombreCompañia].gestionUsuarioAlistamiento)
//     correosAlistamientoGestionUsuario = CigoApp.reemplazarValores(correosAlistamientoGestionUsuario, infoFila)

//     enviarCorreo(correosAlistamientoGestionUsuario, asuntoPasoFlujo, mensajePasoFlujo, { inlineImages: inlineImages })
//   }
// }

// function envioContratacion(formulario) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaIngreso).setValue(formulario.fechaIngreso)
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.citaInduccion).setValue(formulario.citaInduccion)
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.idHojaDeVida).setValue(formulario.idHojaDeVida)


//   var reContratacion = hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.reContratacion).getValue()
//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];
//   var correoPadrino = infoFila[indicesColumnas.consolidado.correoPadrino - 1];
//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];
//   if (infoFila[indicesColumnas.consolidado.masivos - 1] != "Si") {

//     correos = (reContratacion == "Si") ? correoLider : correoLider + "," + correoPadrino

//     if (!correo) throw "El correo externo no está definido";
//     if (!correoLider) throw "El correo del Lider no está definido";
//     if (!correoPadrino) throw "El correo del Padrino no está definido";
//     if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//     let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//     // console.log(filial)

//     var asuntoIntContratado = obtenerComunicacionPorFilial("asuntoIntContratado", filial);
//     asuntoIntContratado = obtenerComunicacionPorFilial(asuntoIntContratado, infoFila);

//     var mensajeIntContratado = obtenerComunicacionPorFilial("mensajeIntContratado", filial)
//     mensajeIntContratado = CigoApp.reemplazarValores(mensajeIntContratado, infoFila)

//     var inlineImages = obtenerImagenesCorreo("mensajeIntContratado", filial);


//     //CigoApp.enviarCorreo(correos, asuntoIntContratado, mensajeIntContratado, {});
//     enviarCorreo(correos, asuntoIntContratado, mensajeIntContratado, { inlineImages: inlineImages });
//     registrarLogCorreo(infoFila, correos.toString(), asuntoIntContratado)
//   }
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.envioCorreoLiderPadrino).setValue("Si")


//   //var asuntoFinalizacionContratacion = CigoApp.reemplazarValores(obtenerParametro("asuntoFinalizacionContratacion"), infoFila);
//   //var mensajeFinalizacionContratacion = CigoApp.reemplazarValores(obtenerParametro("mensajeFinalizacionContratacion"), infoFila);
//   //registrarLogCorreo(infoFila,correo.toString(),asuntoFinalizacionContratacion)

//   //CigoApp.enviarCorreo(correo, asuntoFinalizacionContratacion, mensajeFinalizacionContratacion, {});

// }

// function envioFinalizacionContratacion(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];

//   if (!correo) throw "El correo externo no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   // var asuntoFinalizacionContratacion = CigoApp.reemplazarValores(obtenerParametro("asuntoExperiencia"), infoFila);
//   // var mensajeFinalizacionContratacion = CigoApp.reemplazarValores(obtenerParametro("mensajeExperienciaAlistamiento"), infoFila);

//   // CigoApp.enviarCorreo(correo, asuntoFinalizacionContratacion, mensajeFinalizacionContratacion, {});

// }

// function envioArchivoProveedor(formulario, infoFila, filaConsolidado, hojaConsolidado) {

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));

//   var valorGestion = infoFila[indicesColumnas.consolidado.envioCarnetizacion - 1]
//   if (valorGestion == "Si") throw "El usuario ya se encuentra en gestion de carnetizacion"

//   const codigoSucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1]

//   var hojaSucursales = sps.getSheetByName(obtenerParametro("nombreHojaSucursales"))
//   var datosSucursales = hojaSucursales.getDataRange().getValues()

//   const filaSucursal = CigoApp.buscarPorLlave(codigoSucursal, datosSucursales, 1, -1)

//   var correoCarnetizacion = filaSucursal[5]
//   var correosProveedores = filaSucursal[6]

//   // var correoCarnetizacion = obtenerParametro("carnetizacion");

//   var reContratacion = infoFila[indicesColumnas.consolidado.reContratacion - 1]
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.envioCarnetizacion).setValue("Si")

//   if (reContratacion.toString().toUpperCase() != "NO") return

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   if (!correosProveedores) throw "No se encontraron parametrizados los emails de los proveedores.";
//   if (!correoCarnetizacion) throw "No se encontró parametrizado el email de carnetización.";

//   var asuntoResponsableProveedor = obtenerComunicacionPorFilial("asuntoResponsableProveedor", filial)
//   asuntoResponsableProveedor = CigoApp.reemplazarValores(asuntoResponsableProveedor, infoFila);

//   var mensajeResponsableProveedor = obtenerComunicacionPorFilial("mensajeResponsableProveedor", filial)
//   mensajeResponsableProveedor = CigoApp.reemplazarValores(mensajeResponsableProveedor, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeResponsableProveedor", filial);

//   enviarCorreo(correosProveedores, asuntoResponsableProveedor, mensajeResponsableProveedor, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correosProveedores.toString(), asuntoResponsableProveedor);


//   // var asuntoResponsableCarnetizacion = obtenerComunicacionPorFilial("asuntoResponsableCarnetizacion", filial)
//   // asuntoResponsableCarnetizacion = CigoApp.reemplazarValores(asuntoResponsableCarnetizacion, infoFila);

//   // var mensajeResponsableCarnetizacion = obtenerComunicacionPorFilial("mensajeResponsableCarnetizacion", filial);
//   // mensajeResponsableCarnetizacion = CigoApp.reemplazarValores(mensajeResponsableCarnetizacion, infoFila);

//   // var inlineImages = obtenerImagenesCorreo("mensajeResponsableCarnetizacion", filial);

//   // enviarCorreo(correoCarnetizacion, asuntoResponsableCarnetizacion, mensajeResponsableCarnetizacion, { inlineImages: inlineImages });
//   // registrarLogCorreo(infoFila, correoCarnetizacion.toString(), asuntoResponsableCarnetizacion);

//   // hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaIngresoDocumentos).setValue(new Date());
//   for (var nombreCampo in formulario) {
//     if (!indicesColumnas.consolidado[nombreCampo]) continue;
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado[nombreCampo]).setValue(formulario[nombreCampo]);
//   }

//   var columnaGestion, columnaFechaGestion, columnaCorreoGestion;
//   if (formulario.area == "Contratacion") {
//     columnaGestion = indicesColumnas.consolidado.gestionCarnetizacion;
//     columnaFechaGestion = indicesColumnas.consolidado.horaGestionCarnetizacion;
//     columnaCorreoGestion = indicesColumnas.consolidado.correoGestionCarnetizacion;
//     if (infoFila[columnaFechaGestion - 1] == "") {
//       hojaConsolidado.getRange(filaConsolidado, columnaGestion).setValue("Pendiente Gestión")
//       hojaConsolidado.getRange(filaConsolidado, columnaFechaGestion).setValue(new Date())
//       hojaConsolidado.getRange(filaConsolidado, columnaCorreoGestion).setValue(Session.getActiveUser().getEmail());
//       hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaInicioCarnetizacion).setValue(new Date());
//     }
//   }
// }



// function envioCorreoGestionUsuario(formulario) {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.usuarioCorreo).setValue(formulario.usuarioCorreo);
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaEnvioCorreoGestionUsuario).setValue(new Date())

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];

//   const codigoSucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1]

//   var hojaSucursales = sps.getSheetByName(obtenerParametro("nombreHojaSucursales"))
//   var datosSucursales = hojaSucursales.getDataRange().getValues()

//   var correoGestionUsuario = CigoApp.buscarPorLlave(codigoSucursal, datosSucursales, 1, 5)

//   var correoContratacion = filaConsolidado[indicesColumnas.consolidado.correoGestionCont - 1]
//   var correo = [correoGestionUsuario, correoContratacion]

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoCorreoGestionUsuarios = obtenerComunicacionPorFilial("asuntoCorreoGestionUsuarios", filial);
//   asuntoCorreoGestionUsuarios = CigoApp.reemplazarValores(asuntoCorreoGestionUsuarios, infoFila);

//   if (infoFila[indicesColumnas.consolidado.cargoCritico - 1] == "No") {
//     asuntoCorreoGestionUsuarios = asuntoCorreoGestionUsuarios.replace(/Critico/g, "")
//   } else {
//     asuntoCorreoGestionUsuarios = asuntoCorreoGestionUsuarios.replace(/Critico/g, "Cargo critico")
//   }

//   var mensajeAsuntoGestionUsuarios = obtenerComunicacionPorFilial("mensajeAsuntoGestionUsuarios", filial);
//   mensajeAsuntoGestionUsuarios = CigoApp.reemplazarValores(mensajeAsuntoGestionUsuarios, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeAsuntoGestionUsuarios", filial)


//   //CigoApp.enviarCorreo(correoGestionUsuario, asunto, mensaje, {noReply:true});
//   enviarCorreo(correo, asuntoCorreoGestionUsuarios, mensajeAsuntoGestionUsuarios, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correo, asuntoCorreoGestionUsuarios);
// }


// function envioCorreoGestionUsuario2(formulario) {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.usuarioCorreo).setValue(formulario.usuarioCorreo2);
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaEnvioCorreoGestionUsuario).setValue(new Date())

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];

//   const codigoSucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1]

//   var hojaSucursales = sps.getSheetByName(obtenerParametro("nombreHojaSucursales"))
//   var datosSucursales = hojaSucursales.getDataRange().getValues()

//   var correoGestionUsuario = CigoApp.buscarPorLlave(codigoSucursal, datosSucursales, 1, 5)

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoCorreoGestionUsuarios = obtenerComunicacionPorFilial("asuntoCorreoGestionUsuarios", filial);
//   asuntoCorreoGestionUsuarios = CigoApp.reemplazarValores(asuntoCorreoGestionUsuarios, infoFila);

//   if (infoFila[indicesColumnas.consolidado.cargoCritico - 1] == "No") {
//     asuntoCorreoGestionUsuarios = asuntoCorreoGestionUsuarios.replace(/Critico/g, "")
//   } else {
//     asuntoCorreoGestionUsuarios = asuntoCorreoGestionUsuarios.replace(/Critico/g, "Cargo critico")
//   }

//   var mensajeAsuntoGestionUsuarios = obtenerComunicacionPorFilial("mensajeAsuntoGestionUsuarios", filial);
//   mensajeAsuntoGestionUsuarios = CigoApp.reemplazarValores(mensajeAsuntoGestionUsuarios, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeAsuntoGestionUsuarios", filial)

//   //CigoApp.enviarCorreo(correoGestionUsuario, asunto, mensaje, {noReply:true});
//   enviarCorreo(correoGestionUsuario, asuntoCorreoGestionUsuarios, mensajeAsuntoGestionUsuarios, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correoGestionUsuario, asuntoCorreoGestionUsuarios);
// }

// function envioContrato(formulario) {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);


//   formulario.fechaGenerarContrato = new Date()

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   // hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.respContrato).setValue(formulario.respContrato);
//   // hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaGenerarContrato).setValue(new Date())

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];

//   formulario.sueldoTexto = infoFila[indicesColumnas.consolidado.salario - 1].toString();
//   formulario.sueldoTexto = formulario.sueldoTexto.replace(/\D/g, "");
//   formulario.sueldoTexto = NUMEROATEXTO(formulario.sueldoTexto, "PESOS MCTE");


//   formulario.hkLetras = infoFila[indicesColumnas.consolidado.salarioRemunerativo - 1].toString();
//   console.log(formulario.hkLetras)
//   formulario.hkLetras = formulario.hkLetras.replace(/\D/g, "");
//   formulario.hkLetras = NUMEROATEXTO(formulario.hkLetras, "PESOS MCTE");

//   formulario.hmLetras = infoFila[indicesColumnas.consolidado.salarioFactorPrestacional - 1].toString();
//   console.log(formulario.hmLetras)
//   formulario.hmLetras = formulario.hmLetras.replace(/\D/g, "");
//   formulario.hmLetras = NUMEROATEXTO(formulario.hmLetras, "PESOS MCTE"); 


//   infoFila[indicesColumnas.consolidado.sueldoTexto - 1] = formulario.sueldoTexto;
//   infoFila[indicesColumnas.consolidado.hkLetras - 1] = formulario.hkLetras;
//   infoFila[indicesColumnas.consolidado.hmLetras - 1] = formulario.hmLetras;

//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.sueldoTexto).setValue(formulario.sueldoTexto);
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.hkLetras).setValue(formulario.hkLetras);
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.hmLetras).setValue(formulario.hmLetras);

//   var responsableContrato = formulario.respContrato;
//   if (!responsableContrato) throw "El correo externo no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + formulario.id;

//   var filial = infoFila[indicesColumnas.consolidado.compania - 1]
 
//   var idCarpeta = CigoApp.obtenerIdUrl(obtenerParametro("contratosCarpeta"));
//   var carpetaContratos = DriveApp.getFolderById(idCarpeta);
//   // carpetaContratos.createFile(formulario.contrato);

//   var asuntoContrato = obtenerComunicacionPorFilial("asuntoContrato", filial);
//   asuntoContrato = CigoApp.reemplazarValores(asuntoContrato, infoFila);

//   var mensajeContrato = obtenerComunicacionPorFilial("mesajeContrato", filial);
//   mensajeContrato = CigoApp.reemplazarValores(mensajeContrato, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mesajeContrato", filial)

//   var archivos = combinarCampos(infoFila);

//   // var asuntoContrato = CigoApp.reemplazarValores(obtenerParametro("asuntoContrato"), infoFila);
//   // var mesajeContrato = CigoApp.reemplazarValores(obtenerParametro("mesajeContrato"), infoFila);

//   // combinarCampos(infoFila);

//   // for (var j = 0; j < archivos.length; j++){
//   //   var blob = archivos[j].getBlob();
//   //   archivosDeEmail.push(blob);
//   // }

//   // var blob = archivo.getBlob();

//   // var inlineImages = obtenerImagenesCorreo("mesajeContrato");


//   // // CigoApp.enviarCorreo(responsableContrato, asuntoContrato, mesajeContrato, {attachments:archivos});
//   enviarCorreo(responsableContrato, asuntoContrato, mensajeContrato, {attachments:archivos, inlineImages: inlineImages});
//   registrarLogCorreo(infoFila, responsableContrato, asuntoContrato);
// }

// function enviarMicroinformatica(formulario) {

//   var id = formulario.id;

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var hojaSucursales = sps.getSheetByName(obtenerParametro("nombreHojaSucursales"))
//   var valorSucursal = hojaSucursales.getRange(1, 1, hojaSucursales.getLastRow(), hojaSucursales.getLastColumn()).getValues()
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();


//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   var correoGU = hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.correoElectronico).getValue();
//   var usuarioGU = hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.usuarioRed).getValue();
//   if (correoGU == "" || usuarioGU == "") {
//     throw "El correo electronico y usuario de red de la sección 1 deben estar diligenciados"
//   }
//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];

//   var codigoSucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1]
//   var correosSucursal

//   for (var registro in valorSucursal) {
//     if (valorSucursal[registro][0] != codigoSucursal) continue;
//     correosSucursal = valorSucursal[registro][3]
//   }

//   var correoExterno = infoFila[indicesColumnas.consolidado.correo - 1]
//   if (!correoExterno) throw "El correo externo no está definido";

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   let asuntoKit = "", mensajeKit = "", inlineImages = {}

//   if (formulario.tipoKit != "No aplica") {
//     asuntoKit = obtenerComunicacionPorFilial("asuntoSorpresa", filial);
//     asuntoKit = CigoApp.reemplazarValores(asuntoKit, infoFila);

//     mensajeKit = obtenerComunicacionPorFilial("mensajeSorpresa", filial);
//     mensajeKit = CigoApp.reemplazarValores(mensajeKit, infoFila);
//     inlineImages = obtenerImagenesCorreo("mensajeSorpresa", filial);

//   } else {
//     asuntoKit = obtenerComunicacionPorFilial("asuntoKitNoDespachado", filial);
//     asuntoKit = CigoApp.reemplazarValores(asuntoKit, infoFila);

//     mensajeKit = obtenerComunicacionPorFilial("mensajeKitNoDespachado", filial);
//     mensajeKit = CigoApp.reemplazarValores(mensajeKit, infoFila);

//     inlineImages = obtenerImagenesCorreo("mensajeKitNoDespachado", filial);
//   }

//   enviarCorreo(correoExterno, asuntoKit, mensajeKit, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correoExterno.toString(), asuntoKit);

//   // let asuntoPasoFlujo = obtenerComunicacionPorFilial("asuntoNotificacionPasoFlujo", filial);
//   // asuntoPasoFlujo = CigoApp.reemplazarValores(asuntoPasoFlujo, infoFila);

//   // let mensajePasoFlujo = obtenerComunicacionPorFilial("mensajeNotificacionPasoFlujo", filial);
//   // mensajePasoFlujo = CigoApp.reemplazarValores(mensajePasoFlujo, infoFila);

//   // inlineImages = obtenerImagenesCorreo("mensajeKitNoDespachado", filial);

//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.envioCorreoasuntoNotificacionPasoFlujo).setValue("Si")


//   // enviarCorreo(correosSucursal, asuntoPasoFlujo, mensajePasoFlujo, { inlineImages: inlineImages });
//   // registrarLogCorreo(infoFila, correosSucursal.toString(), asuntoPasoFlujo);

//   enviarConfirmacion(formulario);
// }

// const parametrosCorreoComunicacionPasoFlujo = {
//   "BETA": {
//     "carnetizacion": "correosNotificacionPasoFlujoBetaCarnetizacion",
//     "gestionUsuarioAlistamiento": "correosNotificacionPasoFlujoBetaAlistamientoGestionUsuario",
//   },
//   "GAMMA": {
//     "carnetizacion": "correosNotificacionPasoFlujoGammaCarnetizacion",
//     "gestionUsuarioAlistamiento": "correosNotificacionPasoFlujoGammaAlistamientoGestionUsuario",
//   },
//   "CORREDORES": {
//     "carnetizacion": "correosNotificacionPasoFlujoCorredoresCarnetizacion",
//     "gestionUsuarioAlistamiento": "correosNotificacionPasoFlujoCorredoresAlistamientoGestionUsuario",
//   }
// }



// function enviarConfirmacion(formulario) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];
//   var correoPadrino = infoFila[indicesColumnas.consolidado.correoPadrino - 1];
//   var correos = correoLider + "," + correoPadrino;

//   if (!correoLider) throw "El correo del Lider no está definido";
//   if (!correoPadrino) throw "El correo del Padrino no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + formulario.id;

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   let asuntoExperiencia = obtenerComunicacionPorFilial("asuntoCorreosOK", filial);
//   asuntoExperiencia = CigoApp.reemplazarValores(asuntoExperiencia, infoFila);

//   let mensajeExperiencia = obtenerComunicacionPorFilial("mensajeCorreoOK", filial);
//   mensajeExperiencia = CigoApp.reemplazarValores(mensajeExperiencia, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeCorreoOK", filial);

//   enviarCorreo(correos, asuntoExperiencia, mensajeExperiencia, { inlineImages: inlineImages });

//   for (var nombreCampo in formulario) {
//     if (!indicesColumnas.consolidado[nombreCampo]) continue;
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado[nombreCampo]).setValue(formulario[nombreCampo]);
//   }

//   var columnaGestion
//   if (formulario.area == "Alistamiento") {
//     columnaGestion = indicesColumnas.consolidado.gestionAprendizaje;

//     var fechaAprendizaje = infoFila[indicesColumnas.consolidado.fechaInicioAprendizaje - 1]
//     var fechaFinAprendizaje = infoFila[indicesColumnas.consolidado.horaGestionApr - 1]
//     if (fechaAprendizaje == "") {
//       hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaInicioAprendizaje).setValue(new Date());
//       if (fechaFinAprendizaje == "") hojaConsolidado.getRange(filaConsolidado, columnaGestion).setValue("Pendiente Gestión");
//     }
//   }
//   envioFinalizacionContratacion(formulario.id);
// }

// function enviarRecordatorioConfirmacion(formulario) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + formulario.id;

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];

//   if (!correo) throw "El correo no está definido";

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   let asuntoRecordatorio = obtenerComunicacionPorFilial("asuntoRecordatorioConfirmacion", filial);
//   asuntoRecordatorio = CigoApp.reemplazarValores(asuntoRecordatorio, infoFila);

//   var mensajeRecordatorio = obtenerComunicacionPorFilial("mensajeRecordatorioConfirmacion", filial);
//   mensajeRecordatorio = CigoApp.reemplazarValores(mensajeRecordatorio, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeRecordatorioConfirmacion", infoFila);

//   // CigoApp.enviarCorreo(correo, asunto, mensaje, {});
//   enviarCorreo(correo, asuntoRecordatorio, mensajeRecordatorio, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correo, asuntoRecordatorio);
// }

// function enviarExperiencia(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];
//   var correoPadrino = infoFila[indicesColumnas.consolidado.correoPadrino - 1];
//   var correos = correoLider + "," + correoPadrino;

//   if (!correoLider) throw "El correo del Lider no está definido";
//   if (!correoPadrino) throw "El correo del Padrino no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   // var asuntoCorreosOK = CigoApp.reemplazarValores(obtenerParametro("asuntoCorreosOK"), infoFila);
//   // var mensajeCorreoOK = CigoApp.reemplazarValores(obtenerParametro("mensajeCorreoOK"), infoFila);

//   // var inlineImages = obtenerImagenesCorreo("mensajeCorreoOK");

//   // CigoApp.enviarCorreo(correos, asuntoCorreosOK, mensajeCorreoOK, {});
//   // enviarCorreo(correos, asuntoCorreosOK, mensajeCorreoOK, {inlineImages: inlineImages});

// }

// function enviarCorreoInduccion2(formulario) {
//   var id = formulario.id;
//   var citacionRealInduccion = formulario.citacionRealInduccion;
//   var contenidoCargado = formulario.contenidoCargado;

//   if (!citacionRealInduccion) throw "Ingresa la fecha y hora de la inducción";

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   guardarFormulario(filaConsolidado, hojaConsolidado, formulario)

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];

//   var fechaString = citacionRealInduccion.split("-");


//   var anio = Number(fechaString[0]);
//   var mes = Number(fechaString[1]) - 1;
//   var dia = Number(fechaString[2]);

//   var tipoContrato = infoFila[indicesColumnas.consolidado.tipoContrato - 1];
//   var sucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1];

//   if (tipoContrato != "Aprendiz sena" || (tipoContrato == "Aprendiz Sena" && sucursal == 70)) {

//     var fechaInicial = new Date(anio, mes, dia, 8, 30, 0);
//     var fechaFinal = new Date(anio, mes, dia, 9, 0, 0);
//     var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioLider"), infoFila);
//     var descripcionEvento = CigoApp.reemplazarValores(obtenerParametro("descripcionEventoLider"), infoFila);
//     var correosInvitacion = []
//     if (infoFila[indicesColumnas.consolidado.correoElectronico - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correoElectronico - 1].toUpperCase() != "") {
//       correosInvitacion.push(infoFila[indicesColumnas.consolidado.correoElectronico - 1])
//     }
//     if (infoFila[indicesColumnas.consolidado.correo - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correo - 1].toUpperCase() != "") {
//       correosInvitacion.push(infoFila[indicesColumnas.consolidado.correo - 1])
//     }
//     if (infoFila[indicesColumnas.consolidado.correoLider - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correoLider - 1].toUpperCase() != "") {
//       correosInvitacion.push(infoFila[indicesColumnas.consolidado.correoLider - 1])
//     }
//     var opciones = {
//       description: descripcionEvento,
//       sendInvites: true,
//       guests: correosInvitacion.toString()
//     };

//     var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioInduccion"));
//     var calendarEvent = calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);


//     var fechaInicial = new Date(anio, mes, dia, 9, 00, 0);
//     var fechaFinal = new Date(anio, mes, dia, 9, 30, 0);
//     var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioPadrino"), infoFila);
//     var descripcionEvento = CigoApp.reemplazarValores(obtenerParametro("descripcionEventoPadrino"), infoFila);
//     var correosInvitacion = []
//     if (infoFila[indicesColumnas.consolidado.correoElectronico - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correoElectronico - 1].toUpperCase() != "") {
//       correosInvitacion.push(infoFila[indicesColumnas.consolidado.correoElectronico - 1])
//     }
//     if (infoFila[indicesColumnas.consolidado.correo - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correo - 1].toUpperCase() != "") {
//       correosInvitacion.push(infoFila[indicesColumnas.consolidado.correo - 1])
//     }
//     if (infoFila[indicesColumnas.consolidado.correoPadrino - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correoPadrino - 1].toUpperCase() != "") {
//       correosInvitacion.push(infoFila[indicesColumnas.consolidado.correoPadrino - 1])
//     }



//     var opciones = {
//       description: descripcionEvento,
//       sendInvites: true,
//       guests: correosInvitacion.toString()
//     };

//     var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioInduccion"));
//     var calendarEvent = calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);

//   }//IF



//   //
//   var fechaInicial = new Date(anio, mes, dia, 10, 00, 0);
//   var fechaFinal = new Date(anio, mes, dia, 15, 00, 0);
//   var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioInduccion"), infoFila);
//   var descripcionEvento = CigoApp.reemplazarValores(obtenerParametro("descripcionEventoInduccion"), infoFila);
//   var correosInvitacion = []
//   if (infoFila[indicesColumnas.consolidado.correoElectronico - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correoElectronico - 1].toUpperCase() != "") {
//     correosInvitacion.push(infoFila[indicesColumnas.consolidado.correoElectronico - 1])
//   }
//   if (infoFila[indicesColumnas.consolidado.correo - 1].toUpperCase() != "N/A" && infoFila[indicesColumnas.consolidado.correo - 1].toUpperCase() != "") {
//     correosInvitacion.push(infoFila[indicesColumnas.consolidado.correo - 1])
//   }




//   var opciones = {
//     description: descripcionEvento,
//     sendInvites: true,
//     guests: correosInvitacion.toString()
//   };
//   if (contenidoCargado != "No") {
//     var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioInduccion"));
//     var calendarEvent = calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);
//   }

//   // var idCarpetaAdjuntos = CigoApp.obtenerIdUrl(obtenerParametro("idCarpetaAdjuntos"));
//   // var carpetaAdjuntos = <App.getFolderById(idCarpetaAdjuntos);
//   // var claveAdjuntos = obtenerParametro("claveAdjuntos");
//   // console.log("idCarpetaAdjuntos es: " + idCarpetaAdjuntos);
//   // console.log("claveAdjuntos es: " + claveAdjuntos);
//   // archivo = carpetaAdjuntos.searchFiles("title contains '" + claveAdjuntos + "'");
//   // console.log("archivo es: " + archivo);
//   // if(archivo.hasNext()){
//   //   archivo = archivo.next();
//   // } 
//   // var adjunto = archivo.getBlob();
//   // var opciones = {attachments: [adjunto], noReply: parametros.noReply};

//   var opciones = {};
//   var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];
//   var correoPadrino = infoFila[indicesColumnas.consolidado.correoPadrino - 1];
//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];
//   // var correos = correoLider + "," + correoPadrino;
//   var masivo = infoFila[indicesColumnas.consolidado.masivos - 1]

//   if (infoFila[indicesColumnas.consolidado.masivos - 1] == "No" && !correoPadrino) throw "El correo del Padrino no está definido"
//   if (!correoLider) throw "El correo del Lider no está definido";
//   if (!correo) throw "El correo externo no está definido";

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   if (contenidoCargado != "No") {
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.envioCorreoAprendizajeCargado).setValue("Si")

//     if (masivo == "Si") return

//     var asuntoCitacionInduccion = obtenerComunicacionPorFilial("asuntoCitacionInduccion", filial);
//     asuntoCitacionInduccion = CigoApp.reemplazarValores(asuntoCitacionInduccion, infoFila);

//     var mensajeCitacionInduccion = obtenerComunicacionPorFilial("mensajeCitacionInduccion", filial);
//     mensajeCitacionInduccion = CigoApp.reemplazarValores(mensajeCitacionInduccion, infoFila);

//     var inlineImages = obtenerImagenesCorreo("mensajeCitacionInduccion", filial);

//     // CigoApp.enviarCorreo(correoLider, asuntoCitacionInduccion, mensajeCitacionInduccion, {});
//     enviarCorreo(correoLider, asuntoCitacionInduccion, mensajeCitacionInduccion, { inlineImages: inlineImages });

//     // (8var asuntoPlanPadrino = CigoApp.reemplazarValores(obtenerParametro("asuntoPlanPadrino"), infoFila);
//     // var mensajePlanPadrino = CigoApp.reemplazarValores(obtenerParametro("mensajePlanPadrino"), infoFila);

//     // var inlineImages = obtenerImagenesCorreo("mensajePlanPadrino");

//     // CigoApp.enviarCorreo(correoPadrino, asuntoPlanPadrino, mensajePlanPadrino, {});
//     //enviarCorreo(correoPadrino, asuntoPlanPadrino, mensajePlanPadrino, {inlineImages: inlineImages});
//     //registrarLogCorreo(infoFila, correoPadrino, asuntoPlanPadrino);

//     var asuntoAprendizajeSeleccionado = obtenerComunicacionPorFilial("asuntoAprendizajeSeleccionado", filial);
//     asuntoAprendizajeSeleccionado = CigoApp.reemplazarValores(asuntoAprendizajeSeleccionado, infoFila);

//     var mensajeAprendizajeSeleccionado = obtenerComunicacionPorFilial("mensajeAprendizajeSeleccionado", filial);
//     mensajeAprendizajeSeleccionado = CigoApp.reemplazarValores(mensajeAprendizajeSeleccionado, infoFila);

//     var inlineImages = obtenerImagenesCorreo("mensajeAprendizajeSeleccionado", filial);

//     // CigoApp.enviarCorreo(correo, asuntoAprendizajeSeleccionado, mensajeAprendizajeSeleccionado, {});
//     enviarCorreo(correo, asuntoAprendizajeSeleccionado, mensajeAprendizajeSeleccionado, { inlineImages: inlineImages });
//     registrarLogCorreo(infoFila, correo, asuntoAprendizajeSeleccionado);


//   } else {
//     if (masivo == "Si") return

//     var asuntoCitacionInduccion = obtenerComunicacionPorFilial("asuntoCitacionInduccionContCargadoNo", filial);
//     asuntoCitacionInduccion = CigoApp.reemplazarValores(asuntoCitacionInduccion, infoFila);

//     var mensajeCitacionInduccion = obtenerComunicacionPorFilial("mensajeCitacionInduccionContCargadoNo", filial);
//     mensajeCitacionInduccion = CigoApp.reemplazarValores(mensajeCitacionInduccion, infoFila);

//     var inlineImages = obtenerImagenesCorreo("mensajeCitacionInduccionContCargadoNo", filial);

//     // CigoApp.enviarCorreo(correoLider, asuntoCitacionInduccion, mensajeCitacionInduccion, {});
//     enviarCorreo(correoLider, asuntoCitacionInduccion, mensajeCitacionInduccion, { inlineImages: inlineImages });
//     registrarLogCorreo(infoFila, correoLider, asuntoCitacionInduccion);

//     var asuntoPlanPadrino = obtenerComunicacionPorFilial("asuntoPlanPadrinoContCargadoNo", filial);
//     asuntoPlanPadrino = CigoApp.reemplazarValores(asuntoPlanPadrino, infoFila);

//     var mensajePlanPadrino = obtenerComunicacionPorFilial("mensajePlanPadrinoContCargadoNo", filial);
//     mensajePlanPadrino = CigoApp.reemplazarValores(mensajePlanPadrino, infoFila);

//     var inlineImages = obtenerImagenesCorreo("mensajePlanPadrinoContCargadoNo", filial);

//     // CigoApp.enviarCorreo(correoPadrino, asuntoPlanPadrino, mensajePlanPadrino, {});
//     enviarCorreo(correoPadrino, asuntoPlanPadrino, mensajePlanPadrino, { inlineImages: inlineImages });
//     registrarLogCorreo(infoFila, correoPadrino, asuntoPlanPadrino);

//     var asuntoAprendizajeSeleccionado = obtenerComunicacionPorFilial("asuntoAprendizajeSeleccionadoContCargadoNo", filial);
//     asuntoAprendizajeSeleccionado = CigoApp.reemplazarValores(asuntoAprendizajeSeleccionado, infoFila);

//     var mensajeAprendizajeSeleccionado = obtenerComunicacionPorFilial("mensajeAprendizajeSeleccionadoContCargadoNo", filial);
//     mensajeAprendizajeSeleccionado = CigoApp.reemplazarValores(mensajeAprendizajeSeleccionado, infoFila);

//     var inlineImages = obtenerImagenesCorreo("mensajeAprendizajeSeleccionadoContCargadoNo", filial);

//     // CigoApp.enviarCorreo(correo, asuntoAprendizajeSeleccionado, mensajeAprendizajeSeleccionado, {});
//     enviarCorreo(correo, asuntoAprendizajeSeleccionado, mensajeAprendizajeSeleccionado, { inlineImages: inlineImages });
//     registrarLogCorreo(infoFila, correo, asuntoAprendizajeSeleccionado);
//   }
// }


// function agendarInduccion(id, citacionRealInduccion) {
//   if (!citacionRealInduccion) throw "Ingresa la fecha y hora de la inducción";

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;
//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];

//   var fechaString = citacionRealInduccion.split("-");


//   var anio = Number(fechaString[0]);
//   var mes = Number(fechaString[1]) - 1;
//   var dia = Number(fechaString[2]);


//   var fechaInicial = new Date(anio, mes, dia, 8, 30, 0);
//   var fechaFinal = new Date(anio, mes, dia, 9, 0, 0);
//   var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioLider"), infoFila);
//   var descripcionEvento = CigoApp.reemplazarValores(obtenerParametro("descripcionEventoLider"), infoFila);

//   var correosInvitacion = [
//     infoFila[indicesColumnas.consolidado.correoElectronico - 1],
//     infoFila[indicesColumnas.consolidado.correo - 1],
//     infoFila[indicesColumnas.consolidado.correoLider - 1]
//   ];
//   var opciones = {
//     description: descripcionEvento,
//     sendInvites: true,
//     guests: correosInvitacion.toString()
//   };

//   var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioInduccion"));
//   var calendarEvent = calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);


//   var fechaInicial = new Date(anio, mes, dia, 9, 00, 0);
//   var fechaFinal = new Date(anio, mes, dia, 9, 30, 0);
//   var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioPadrino"), infoFila);
//   var descripcionEvento = CigoApp.reemplazarValores(obtenerParametro("descripcionEventoPadrino"), infoFila);

//   var correosInvitacion = [
//     infoFila[indicesColumnas.consolidado.correoElectronico - 1],
//     infoFila[indicesColumnas.consolidado.correoPadrino - 1],
//     infoFila[indicesColumnas.consolidado.correo - 1],
//   ];

//   var opciones = {
//     description: descripcionEvento,
//     sendInvites: true,
//     guests: correosInvitacion.toString()
//   };

//   var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioInduccion"));
//   var calendarEvent = calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);

//   //
//   var fechaInicial = new Date(anio, mes, dia, 10, 00, 0);
//   var fechaFinal = new Date(anio, mes, dia, 15, 00, 0);
//   var tituloEvento = CigoApp.reemplazarValores(obtenerParametro("tituloCalendarioInduccion"), infoFila);
//   var descripcionEvento = CigoApp.reemplazarValores(obtenerParametro("descripcionEventoInduccion"), infoFila);

//   var correosInvitacion = [
//     infoFila[indicesColumnas.consolidado.correoElectronico - 1],
//     infoFila[indicesColumnas.consolidado.correo - 1]
//   ];

//   var opciones = {
//     description: descripcionEvento,
//     sendInvites: true,
//     guests: correosInvitacion.toString()
//   };

//   var calendario = CalendarApp.getCalendarById(obtenerParametro("idCalendarioInduccion"));
//   var calendarEvent = calendario.createEvent(tituloEvento, fechaInicial, fechaFinal, opciones);

// }

// function enviarXplora(formulario) {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correoLider) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + formulario.id;

//   // var opcion = formulario.fechaInduccionSelect;

//   // if (opcion == "SI") {
//   //   var asuntoFinInduccionSI = CigoApp.reemplazarValores(obtenerParametro("asuntoFinInduccionSI"), infoFila);
//   //   var mensajeFinInduccionSI = CigoApp.reemplazarValores(obtenerParametro("mensajeFinInduccionSI"), infoFila);

//   //   var inlineImages = obtenerImagenesCorreo("mensajeFinInduccionSI");


//   //   // CigoApp.enviarCorreo(correoLider, asuntoFinInduccionSI, mensajeFinInduccionSI, {});
//   //   // enviarCorreo(correoLider, asuntoFinInduccionSI, mensajeFinInduccionSI, {inlineImages: inlineImages});
//   // } else {
//   //   var asuntoFinInduccionNO = CigoApp.reemplazarValores(obtenerParametro("asuntoFinInduccionNO"), infoFila);
//   //   var mensajeFinInduccionNO = CigoApp.reemplazarValores(obtenerParametro("mensajeFinInduccionNO"), infoFila);

//   //   var inlineImages = obtenerImagenesCorreo("mensajeFinInduccionNO");


//   //   // CigoApp.enviarCorreo(correoLider, asuntoFinInduccionNO, mensajeFinInduccionNO, {});
//   //   // enviarCorreo(correoLider, asuntoFinInduccionNO, mensajeFinInduccionNO, {inlineImages: inlineImages});
//   // }

// }

// function enviarPlanPadrino(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correoPadrino = infoFila[indicesColumnas.consolidado.correoPadrino - 1];

//   if (!correoPadrino) throw "El correo del Padrino no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;
//   if (infoFila[indicesColumnas.consolidado.masivos - 1] == "Si") return

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoPlanPadrino = obtenerComunicacionPorFilial("asuntoPlanPadrino", filial);
//   asuntoPlanPadrino = CigoApp.reemplazarValores(asuntoPlanPadrino, infoFila);

//   var mensajePlanPadrino = obtenerComunicacionPorFilial("mensajePlanPadrino", filial);
//   mensajePlanPadrino = CigoApp.reemplazarValores(mensajePlanPadrino, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajePlanPadrino", filial);

//   // CigoApp.enviarCorreo(correoPadrino, asuntoPlanPadrino, mensajePlanPadrino, {});
//   enviarCorreo(correoPadrino, asuntoPlanPadrino, mensajePlanPadrino, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correoPadrino, asuntoPlanPadrino);
// }

// function enviarExperienciaAprendizaje(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correoLider) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   // var asuntoExperiencia = CigoApp.reemplazarValores(obtenerParametro("asuntoExperiencia"), infoFila);
//   // var mensajeExperiencia = CigoApp.reemplazarValores(obtenerParametro("mensajeExperiencia"), infoFila);

//   // CigoApp.enviarCorreo(correoLider, asuntoExperiencia, mensajeExperiencia, {});

// }

// function enviarNPS(id) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
//   var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];

//   if (!correoLider) throw "El correo del Lider no está definido";
//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   // var asuntoNPS = CigoApp.reemplazarValores(obtenerParametro("asuntoNPS"), infoFila);
//   // var mensajeNPS = CigoApp.reemplazarValores(obtenerParametro("mensajeNPS"), infoFila);

//   // var inlineImages = obtenerImagenesCorreo("mensajeNPS");


//   // CigoApp.enviarCorreo(correoLider, asuntoNPS, mensajeNPS, {});
//   // enviarCorreo(correoLider, asuntoNPS, mensajeNPS, {inlineImages: inlineImages});

// }

// function notificarAfiliaciones(formulario) {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];

//   var datosComunicaciones = parametros.spsParametros.getSheetByName("Comunicacion Afiliaciones").getDataRange().getValues()

//   var llaveComunicaciones = datosComunicaciones.reduce((llaves, fila) => {
//     var campo = fila[0].trim()
//     var valor = fila[1].trim()
//     if (!llaves[campo]) {
//       llaves[campo] = {}
//     }
//     llaves[campo][valor] = fila

//     return llaves
//   }, {})

//   var sucursalesNotifiacion1 = obtenerParametro("sucursalesNotificacion1").split(",")
//   var sucuraslesNotifiacion2 = obtenerParametro("sucursalesNotificacion2").split(",")

//   var sucursalRegistro = infoFila[indicesColumnas.consolidado.codigoSucursal - 1]

//   var valorCampo, infoCorreo, correoContratacion, columnaCorreo = 0

//   if (sucursalesNotifiacion1.indexOf(sucursalRegistro) > -1) {
//     columnaCorreo = 5
//   }
//   if (sucuraslesNotifiacion2.indexOf(sucursalRegistro) > -1) {
//     columnaCorreo = 6
//   }

//   if (columnaCorreo == 0) throw "No esta paremetrizado las notificaciones paara la sucursal " + sucursalRegistro

//   var carpeteUsuario = DriveApp.getFolderById(CigoApp.obtenerIdUrl(infoFila[indicesColumnas.consolidado.linkCarpDigital - 1]))
//   var docCedula = carpeteUsuario.searchFiles("title contains 'Cedula'").next().getBlob()

//   correoContratacion = infoFila[indicesColumnas.consolidado.correoGestionCont - 1]

//   if (formulario.arl.toString().toUpperCase() != "NO APLICA") {
//     valorCampo = formulario.arl.trim()
//     infoCorreo = llaveComunicaciones.arl[valorCampo]

//     if (!infoCorreo) throw "No se ha encontrado la ARL " + valorCampo

//     enviarCorreoNotificacionPorAfiliacion(infoCorreo, infoFila, columnaCorreo, correoContratacion, docCedula)
//   }
//   if (formulario.cajaCompensacion.toString().toUpperCase() != "NO APLICA") {
//     valorCampo = formulario.cajaCompensacion.trim()
//     infoCorreo = llaveComunicaciones.cajaCompensacion[valorCampo]

//     if (!infoCorreo) throw "No se ha encontrado la caja de compensacion " + valorCampo

//     enviarCorreoNotificacionPorAfiliacion(infoCorreo, infoFila, columnaCorreo, correoContratacion, docCedula)
//   }
//   if (formulario.eps.toString().toUpperCase() != "NO APLICA") {
//     valorCampo = formulario.eps.trim()
//     infoCorreo = llaveComunicaciones.eps[valorCampo]

//     if (!infoCorreo) throw "No se ha encontrado la EPS " + valorCampo

//     enviarCorreoNotificacionPorAfiliacion(infoCorreo, infoFila, columnaCorreo, correoContratacion, docCedula)
//   }

//   if (formulario.notificacionCesantias && formulario.fondoCesantias.toString().toUpperCase() != "NO APLICA") {
//     valorCampo = formulario.fondoCesantias.trim()
//     infoCorreo = llaveComunicaciones.fondoCesantias[valorCampo]

//     if (!infoCorreo) throw "No se ha encontrado el fondo de cesantias " + valorCampo

//     enviarCorreoNotificacionPorAfiliacion(infoCorreo, infoFila, columnaCorreo, correoContratacion, docCedula)
//   }
//   if (formulario.notificacionPensiones && formulario.afp.toString().toUpperCase() != "NO APLICA") {
//     valorCampo = formulario.afp.trim()
//     infoCorreo = llaveComunicaciones.afp[valorCampo]

//     if (!infoCorreo) throw "No se ha encontrado el fondo de pensiones " + valorCampo

//     enviarCorreoNotificacionPorAfiliacion(infoCorreo, infoFila, columnaCorreo, correoContratacion, docCedula)
//   }
// }

// function enviarCorreoNotificacionPorAfiliacion(infoCorreo, infoFila, columnaCorreo, correoContratacion, docCedula) {
//   // console.log(infoCorreo)
//   var correo = infoCorreo[columnaCorreo]

//   // let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asunto = infoCorreo[3].trim()
//   if (asunto == "") return
//   asunto = CigoApp.reemplazarValores(asunto, infoFila)

//   var mensaje = infoCorreo[4].trim()
//   if (mensaje == "") return
//   mensaje = CigoApp.reemplazarValores(mensaje, infoFila)

//   CigoApp.enviarCorreo(correo, asunto, mensaje, { noReply: false, cc: correoContratacion, attachments: [docCedula] })
// }

// function enviarRecordatorioFoto(formulario) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["fotoCorrecta"]).setValue(formulario.fotoCorrecta);
//   if (formulario.fotoCorrecta == "Si") {
//     hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["fechaFotoCorrecta"]).setValue(new Date());
//   }
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.fechaEnvioRecordatorioFoto).setValue(new Date());

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoRecordatorioFoto = obtenerComunicacionPorFilial("asuntoRecordatorioFoto", filial);
//   asuntoRecordatorioFoto = CigoApp.reemplazarValores(asuntoRecordatorioFoto, infoFila);

//   var mensajeRecordatorioFoto = obtenerComunicacionPorFilial("mensajeRecordatorioFoto", infoFila);
//   mensajeRecordatorioFoto = CigoApp.reemplazarValores(mensajeRecordatorioFoto, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeRecordatorioFoto", filial);

//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];

//   enviarCorreo(correo, asuntoRecordatorioFoto, mensajeRecordatorioFoto, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correo, asuntoRecordatorioFoto);
// }

// function enviarCodigoAccessMobile(formulario) {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   if (formulario.codigoCarnetDigital == "") throw "Por favor complete el campo Código autorización carné digital"

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(formulario.id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;

//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado.codigoCarnetDigital).setValue(formulario.codigoCarnetDigital);
//   SpreadsheetApp.flush()

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getDisplayValues()[0];

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoCodigoAccess = obtenerComunicacionPorFilial("asuntoCodigoAccess", filial);
//   asuntoCodigoAccess = CigoApp.reemplazarValores(asuntoCodigoAccess, infoFila);

//   var mensajeCodigoAccess = obtenerComunicacionPorFilial("mensajeCodigoAccess", filial);
//   mensajeCodigoAccess = CigoApp.reemplazarValores(mensajeCodigoAccess, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeCodigoAccess", filial);

//   var correo = infoFila[indicesColumnas.consolidado.correo - 1];

//   enviarCorreo(correo, asuntoCodigoAccess, mensajeCodigoAccess, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correo, asuntoCodigoAccess);

// }

// function enviarCorreoDesiste(id, motivo) {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var idsConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.id, hojaConsolidado.getLastRow(), 1).getValues();
//   var filaConsolidado = CigoApp.obtenerFila(id, idsConsolidado, 1);

//   if (!filaConsolidado) throw "No se encontró el registro con el ID " + id;
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["motivoDesiste"]).setValue(motivo);
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["fechaDesiste"]).setValue(new Date());
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["gestionContratacion"]).setValue("Desistido");
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["gestionAlistamiento"]).setValue("Desistido");
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["gestionAprendizaje"]).setValue("Desistido");
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["gestionCarnetizacion"]).setValue("Desistido");
//   hojaConsolidado.getRange(filaConsolidado, indicesColumnas.consolidado["estadoGeneral"]).setValue("Desistido");

//   var infoFila = hojaConsolidado.getRange(filaConsolidado, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];

//   var correoConsultor = infoFila[indicesColumnas.consolidado["correoConsultor"] - 1];
//   if (!correoConsultor) throw "El correo del consultor del contrato no está definido";

//   var correoLider = infoFila[indicesColumnas.consolidado["correoLider"] - 1];
//   var correoAlistamiento = infoFila[indicesColumnas.consolidado["correoGestionAlist"] - 1]
//   var correoCarnetizacion = infoFila[indicesColumnas.consolidado["correoGestionCarnetizacion"] - 1]
//   var correoApredizaje = infoFila[indicesColumnas.consolidado["correoGestionApr"] - 1]
//   var correoConsultor = infoFila[indicesColumnas.consolidado.correoConsultor - 1]

//   var correos = [correoConsultor, correoLider, correoApredizaje, correoCarnetizacion, correoAlistamiento, correoConsultor];

//   let filial = infoFila[indicesColumnas.consolidado.compania - 1]

//   var asuntoProcesoDesistido = obtenerComunicacionPorFilial("asuntoProcesoDesistido", filial);
//   asuntoProcesoDesistido = CigoApp.reemplazarValores(asuntoProcesoDesistido, infoFila);

//   var mensajeProcesoDesistido = obtenerComunicacionPorFilial("mensajeProcesoDesistido", filial);
//   mensajeProcesoDesistido = CigoApp.reemplazarValores(mensajeProcesoDesistido, infoFila);

//   var inlineImages = obtenerImagenesCorreo("mensajeProcesoDesistido", filial);

//   enviarCorreo(correos.toString(), asuntoProcesoDesistido, mensajeProcesoDesistido, { inlineImages: inlineImages });
//   registrarLogCorreo(infoFila, correos.toString(), asuntoProcesoDesistido);
// }

// function registrarLogCorreo(registro, destinatario, asunto) {
//   var spsLog = SpreadsheetApp.openById(obtenerParametro("idHojaPlano")).getSheetByName(obtenerParametro("nombreHojaLogCorreos"))
//   CigoApp.removerFiltro(spsLog);

//   spsLog.appendRow([
//     registro[indicesColumnas.consolidado.numCedula - 1],
//     new Date(),
//     destinatario,
//     asunto
//   ]);
// }

// /**
// * Obtiene las listas desde la hoja de parámetros
// */
// function obtenerListas() {
//   var listas = parametros.spsParametros.getSheetByName("Listas").getDataRange().getValues();
//   return listas;
// }

// function envioEmailCumplimientoCursos() {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var datosConsolidado = hojaConsolidado.getDataRange().getValues();
//   datosConsolidado.shift()

//   columnaCitacionRealInduccion = indicesColumnas.consolidado.citacionRealInduccion;
//   columnaCorreoElectronico = indicesColumnas.consolidado.correoElectronico;
//   columnaCorreoLider = indicesColumnas.consolidado.correoLider;
//   columnaCorreoPadrino = indicesColumnas.consolidado.correoPadrino;
//   columnaPorcentajeCursos = indicesColumnas.consolidado.porcentajeFinalizacionInduccion;

//   for (var fila in datosConsolidado) {
//     if (datosConsolidado[fila][indicesColumnas.consolidado.estadoGeneral - 1] == "Desistido") continue;

//     citaRealInduccion = datosConsolidado[fila][columnaCitacionRealInduccion - 1];
//     correoSeleccionado = datosConsolidado[fila][columnaCorreoElectronico - 1];
//     correoLider = datosConsolidado[fila][columnaCorreoLider - 1];
//     correoPadrino = datosConsolidado[fila][columnaCorreoPadrino - 1];

//     var porcentajeCursos = datosConsolidado[fila][columnaPorcentajeCursos - 1];
//     if (porcentajeCursos == "") continue;

//     if (porcentajeCursos < 100 && porcentajeCursos > 0) {
//       if (!citaRealInduccion) {
//         continue;
//       } else if (citaRealInduccion != "Citación Real de Inducción") {
//         console.log("cita real de induccion es: " + citaRealInduccion);
//         console.log("el correo del seleccionado es: " + correoSeleccionado);
//         console.log("el correo del lider es: " + correoLider);
//         console.log("el correo del padrino es: " + correoPadrino);
//         console.log("-------------------");
//         if (correoSeleccionado == "") {
//           console.log("finaliza proceso porque no se encuentra correo del seleccionado");
//           continue;
//         } else if (correoLider == "") {
//           console.log("finaliza proceso porque no se encuentra correo del lider");
//           continue;
//         } else if (correoPadrino == "") {
//           console.log("finaliza proceso porque no se encuentra correo del padrino");
//           continue;
//         }
//         var fechaActual = new Date();
//         var correosQuinceDias = [correoLider, correoPadrino];
//         var correosOchoDias = [correoSeleccionado];
//         sumarDias(citaRealInduccion, 14);

//         if (formatearFecha(fechaActual) == formatearFecha(citaRealInduccion)) {

//           // var asuntoCumpCursosQuinceDias = CigoApp.reemplazarValores(obtenerParametro("asuntoCumpCursosQuinceDias"), datosConsolidado[fila]);
//           // var mensajeCumpCursosQuinceDias = CigoApp.reemplazarValores(obtenerParametro("mensajeCumpCursosQuinceDias"), datosConsolidado[fila]);

//           // var inlineImages = obtenerImagenesCorreo("mensajeCumpCursosQuinceDias");

//           // CigoApp.enviarCorreo(correosQuinceDias, asuntoCumpCursosQuinceDias, mensajeCumpCursosQuinceDias, {});
//           // enviarCorreo(correosQuinceDias, asuntoCumpCursosQuinceDias, mensajeCumpCursosQuinceDias, {inlineImages: inlineImages});

//         } else {
//           sumarDias(citaRealInduccion, -7);
//           if (formatearFecha(fechaActual) == formatearFecha(citaRealInduccion)) {

//             let filial = datosConsolidado[fila][indicesColumnas.consolidado.compania - 1]

//             var asuntoCumpCursosOchoDias = obtenerComunicacionPorFilial("asuntoCumpCursosOchoDias", filial);
//             asuntoCumpCursosOchoDias = CigoApp.reemplazarValores(asuntoCumpCursosOchoDias, datosConsolidado[fila]);

//             var mensajeCumpCursosOchoDias = obtenerComunicacionPorFilial("mensajeCumpCursosOchoDias", filial);
//             mensajeCumpCursosOchoDias = CigoApp.reemplazarValores(mensajeCumpCursosOchoDias, datosConsolidado[fila]);

//             var inlineImages = obtenerImagenesCorreo("mensajeCumpCursosOchoDias", filial);

//             // CigoApp.enviarCorreo(correosOchoDias, asuntoCumpCursosOchoDias, mensajeCumpCursosOchoDias, {});
//             enviarCorreo(correosOchoDias, asuntoCumpCursosOchoDias, mensajeCumpCursosOchoDias, { inlineImages: inlineImages });
//             registrarLogCorreo(datosConsolidado[fila], correosOchoDias, asuntoCumpCursosOchoDias);

//           }
//         }
//       }
//     }
//   }
// }



// function sumarDias(fecha, dias) {
//   fecha.setDate(fecha.getDate() + dias);
//   return fecha;
// }

// function realizarBackUp() {

//   parametros = obtenerParametros()
//   indicesColumnas = obtenerIndicesColumnas()

//   actualizarInformacionBackUp()
//   eliminarInfoConsolidado()
// }

// function actualizarInformacionBackUp() {
//   var sps = SpreadsheetApp.getActiveSpreadsheet();
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var datosConsolidado = hojaConsolidado.getDataRange().getValues();


//   var spsBackUp = SpreadsheetApp.openById(CigoApp.obtenerIdUrl(obtenerParametro("urlBackUp")))
//   var hojaConsolidadoBackUp = spsBackUp.getSheetByName(obtenerParametro("nombreHojaBackUp"));
//   var datosConsolidadoBackUp = hojaConsolidadoBackUp.getDataRange().getValues();

//   var columnasExtraer = obtenerParametro("ColumnasExtraer").split(",")
//   // var columnasAgregar = obtenerParametro("ColumnasAgregar").split(",")

//   columnasExtraer = CigoApp.obtenerNumeroColumna(columnasExtraer)

//   for (var fila in datosConsolidado) {
//     var datosFilaExtraer = []
//     for (var k = 0; k < columnasExtraer.length; k++) {
//       let columna = !datosConsolidado[fila][columnasExtraer[k] - 1] ? "" : datosConsolidado[fila][columnasExtraer[k] - 1]
//       datosFilaExtraer.push(columna)
//     }
//     var filaBackUp = CigoApp.obtenerFila(datosConsolidado[fila][indicesColumnas.consolidado.id - 1], datosConsolidadoBackUp, 1)

//     if (filaBackUp) {
//       datosConsolidadoBackUp[filaBackUp - 1] = datosFilaExtraer
//     } else {
//       datosConsolidadoBackUp.push(datosFilaExtraer)
//     }
//   }

//   // console.log(datosConsolidadoBackUp[0])s
//   hojaConsolidadoBackUp.getRange(1, 1, datosConsolidadoBackUp.length, datosConsolidadoBackUp[0].length).setValues(datosConsolidadoBackUp)//
// }

// function eliminarInfoConsolidado() {

//   parametros = obtenerParametros()
//   indicesColumnas = obtenerIndicesColumnas()

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var datosConsolidado = hojaConsolidado.getDataRange().getValues()
//   // var datosConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.horaGestionApr, hojaConsolidado.getLastRow(), 1).getValues();
//   // var contadorFilaEliminadas = 0
//   var fechaActual = new Date()

//   const columnaFechaAprendizaje = indicesColumnas.consolidado.horaGestionApr
//   const columnaEstadoGeneral = indicesColumnas.consolidado.estadoGeneral

//   for (var fila = datosConsolidado.length; fila > 1; fila--) {

//     const estadoGeneral = datosConsolidado[fila - 1][columnaEstadoGeneral - 1].toString().toUpperCase()

//     if (estadoGeneral == "DESISTIDO" || estadoGeneral == "NO APLICA") {
//       hojaConsolidado.deleteRow(fila)
//       continue
//     }

//     let fechaGestionAprendizaje = datosConsolidado[fila - 1][columnaFechaAprendizaje - 1]
//     if (!(fechaGestionAprendizaje instanceof Date)) continue

//     sumarDias(fechaGestionAprendizaje, 90);
//     if (Math.round((fechaActual.getTime() - fechaGestionAprendizaje.getTime()) / (1000 * 60 * 60 * 24) < 0)) continue

//     hojaConsolidado.deleteRow(fila)
//   }
// }

// function envioNotificacionesEncuestaYNPS() {
//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
//   var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var datosConsolidado = hojaConsolidado.getDataRange().getValues();

//   columnaFechaIngreso = indicesColumnas.consolidado.fechaInicioLabo;
//   columnaCorreoConsultor = indicesColumnas.consolidado.correoConsultor;
//   columnaCorreoLider = indicesColumnas.consolidado.correoLider;
//   columnaCorreoPadrino = indicesColumnas.consolidado.correoPadrino;
//   columnaCorreo = indicesColumnas.consolidado.correoElectronico;

//   for (var fila in datosConsolidado) {
//     fechaIngreso = datosConsolidado[fila][columnaFechaIngreso - 1];
//     correoLider = datosConsolidado[fila][columnaCorreoLider - 1];
//     correo = datosConsolidado[fila][columnaCorreo - 1];
//     let filial = datosConsolidado[fila][indicesColumnas.consolidado.compania - 1]

//     if (!fechaIngreso) {
//       console.log("vacio");
//       continue;
//     } else if (fechaIngreso != "Fecha Ingreso") {
//       console.log("valor: " + fechaIngreso);
//       var fechaActual = new Date();
//       sumarDias(fechaIngreso, 7);

//       if (formatearFecha(fechaActual) == formatearFecha(fechaIngreso)) {

//         var asuntoNPS = obtenerComunicacionPorFilial("asuntoNPS", filial);
//         asuntoNPS = CigoApp.reemplazarValores(asuntoNPS, datosConsolidado[fila]);

//         var mensajeNPS = obtenerComunicacionPorFilial("mensajeNPS", filial);
//         mensajeNPS = CigoApp.reemplazarValores(mensajeNPS, datosConsolidado[fila]);

//         var inlineImages = obtenerImagenesCorreo("mensajeNPS", filial);

//         // CigoApp.enviarCorreo(correoLider, asuntoNPS, mensajeNPS, {});
//         enviarCorreo(correo, asuntoNPS, mensajeNPS, { inlineImages: inlineImages });
//         registrarLogCorreo(datosConsolidado[fila], correo, asuntoNPS);

//         var asuntoExperiencia = obtenerComunicacionPorFilial("asuntoNPSLider", filial);
//         asuntoExperiencia = CigoApp.reemplazarValores(asuntoExperiencia, datosConsolidado[fila]);

//         var mensajeExperiencia = obtenerComunicacionPorFilial("mensajeNPSLider", filial);
//         mensajeExperiencia = CigoApp.reemplazarValores(mensajeExperiencia, datosConsolidado[fila]);

//         var inlineImages = obtenerImagenesCorreo("mensajeNPSLider", filial);

//         // CigoApp.enviarCorreo(correoLider, asuntoExperiencia, mensajeExperiencia, {});
//         enviarCorreo(correoLider, asuntoExperiencia, mensajeExperiencia, { inlineImages: inlineImages });
//         registrarLogCorreo(datosConsolidado[fila], correoLider, asuntoExperiencia);

//       } else {
//         sumarDias(fechaIngreso, -11);
//         if (formatearFecha(fechaActual) == formatearFecha(fechaIngreso)) {

//           // var asuntoExperiencia = CigoApp.reemplazarValores(obtenerParametro("asuntoNPSLider"), datosConsolidado[fila]);
//           // var mensajeExperiencia = CigoApp.reemplazarValores(obtenerParametro("mensajeNPSLider"), datosConsolidado[fila]);

//           // var inlineImages = obtenerImagenesCorreo("mensajeNPSLider");

//           // // CigoApp.enviarCorreo(correoLider, asuntoExperiencia, mensajeExperiencia, {});
//           // enviarCorreo(correoLider, asuntoExperiencia, mensajeExperiencia, {inlineImages: inlineImages});

//         } else {
//           console.log("No se debe enviar email");
//         }
//       }
//     }
//   }
// }

// function envioCorreoReContratacionesAlistamiento() {
//   parametros = obtenerParametros()
//   indicesColumnas = obtenerIndicesColumnas()
//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
//   var nombreHojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"))
//   var datosConsolidado = nombreHojaConsolidado.getDataRange().getDisplayValues()
//   for (var i = 1; i < datosConsolidado.length; i++) {
//     if (datosConsolidado[i][indicesColumnas.consolidado.reContratacion - 1] == "Si" && datosConsolidado[i][indicesColumnas.consolidado.estadoGeneral - 1] != "Desistido" && datosConsolidado[i][indicesColumnas.consolidado.correoElectronico - 1] != "" && datosConsolidado[i][indicesColumnas.consolidado.usuarioRed - 1] != "" && datosConsolidado[i][indicesColumnas.consolidado.correoRecontratacionesOk - 1] != "Si") {

//       nombreHojaConsolidado.getRange(i + 1, [indicesColumnas.consolidado.correoRecontratacionesOk]).setValue("Si")
//       var correoLider = datosConsolidado[i][indicesColumnas.consolidado.correoLider - 1];
//       var correoPadrino = datosConsolidado[i][indicesColumnas.consolidado.correoPadrino - 1]
//       var correos = correoLider + "," + correoPadrino

//       let filial = datosConsolidado[i][indicesColumnas.consolidado.compania - 1]

//       var asunto = obtenerComunicacionPorFilial("asuntoCorreosOK", filial)
//       asunto = CigoApp.reemplazarValores(asunto, datosConsolidado[i])

//       var mensaje = obtenerComunicacionPorFilial("mensajeCorreoOK", filial)
//       mensaje = CigoApp.reemplazarValores(mensaje, datosConsolidado[i])

//       var inlineImages = obtenerImagenesCorreo("mensajeCorreoOK")
//       enviarCorreo(correos, asunto, mensaje, { inlineImages: inlineImages });
//       registrarLogCorreo(datosConsolidado[i], correoLider, asunto);
//     }
//   }
// }


// function envioCorreoPorcentajeCursos() {
//   parametros = obtenerParametros()
//   indicesColumnas = obtenerIndicesColumnas()
//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
//   var nombreHojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"))
//   var datosConsolidado = nombreHojaConsolidado.getRange(2, 1, nombreHojaConsolidado.getLastRow() - 1, nombreHojaConsolidado.getLastColumn()).getDisplayValues()

//   var fechaActual = new Date()
//   for (var i in datosConsolidado) {
//     if (datosConsolidado[i][indicesColumnas.consolidado.porcentajeFinalizacionInduccion - 1] != 100 && datosConsolidado[i][indicesColumnas.consolidado.estadoGeneral - 1] != "Desistido") {
//       var fechaCitacion = new Date(datosConsolidado[i][indicesColumnas.consolidado.citaContratacion - 1])
//       var diaDiferencia = Math.round((fechaActual.getTime() - fechaCitacion.getTime()) / (1000 * 60 * 60 * 24))

//       let filial = datosConsolidado[i][indicesColumnas.consolidado.compania - 1]

//       if (diaDiferencia > 6 && diaDiferencia <= 14) {
//         var asunto = obtenerComunicacionPorFilial("asuntoFinInduccionNO", filial)
//         asunto = CigoApp.reemplazarValores(asunto, datosConsolidado[i])

//         var mensaje = obtenerComunicacionPorFilial("mensajeFinInduccionNO", filial)
//         mensaje = CigoApp.reemplazarValores(mensaje, datosConsolidado[i])

//         var inlineImages = obtenerImagenesCorreo("mensajeFinInduccionNO", filial)

//         var correoPersonal = datosConsolidado[i][indicesColumnas.consolidado.correo - 1]

//         enviarCorreo(correoPersonal, asunto, mensaje, { inlineImages: inlineImages });
//         registrarLogCorreo(datosConsolidado[i], correo, asunto);

//       } else if (diaDiferencia > 14) {

//         var asunto = obtenerComunicacionPorFilial("asuntoCumpCursosQuinceDias", filial)
//         asunto = CigoApp.reemplazarValores(asunto, datosConsolidado[i])

//         var mensaje = obtenerComunicacionPorFilial("mensajeCumpCursosQuinceDias", filial)
//         vmensaje = CigoApp.reemplazarValores(mensaje, datosConsolidado[i])

//         var inlineImages = obtenerImagenesCorreo("mensajeCumpCursosQuinceDias", filial)
//         var correoLider = datosConsolidado[i][indicesColumnas.consolidado.correoLider - 1]

//         enviarCorreo(correoLider, asunto, mensaje, { inlineImages: inlineImages });
//         registrarLogCorreo(datosConsolidado[i], correo, asunto);
//       }
//     }
//   }
// }


// function identifiacionAprendiz() {
//   parametros = obtenerParametros()
//   indicesColumnas = obtenerIndicesColumnas()
//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
//   var nombreHojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"))
//   var hojaSucursales = sps.getSheetByName(obtenerParametro("nombreHojaSucursales"))
//   var valorSucursal = hojaSucursales.getDataRange().getValues()
//   var datosConsolidado = nombreHojaConsolidado.getDataRange().getDisplayValues()


//   for (var i = 1; i < datosConsolidado.length; i++) {

//     if ((datosConsolidado[i][indicesColumnas.consolidado.catCargo - 1].toLowerCase().indexOf("aprendiz") != -1 || datosConsolidado[i][indicesColumnas.consolidado.catCargo - 1].toLowerCase().indexOf("sena") != -1 || datosConsolidado[i][indicesColumnas.consolidado.catCargo - 1].toLowerCase().indexOf("semillero") != -1 || datosConsolidado[i][indicesColumnas.consolidado.catCargo - 1].toLowerCase().indexOf("fuerza comercial") != -1 || datosConsolidado[i][indicesColumnas.consolidado.catCargo - 1].toLowerCase().indexOf("formación") != -1) && datosConsolidado[i][indicesColumnas.consolidado.estadoGeneral - 1] != "Desistido") {

//       if (datosConsolidado[i][indices.consolidado.correoElectronico - 1] != "" && datosConsolidado[i][indices.consolidado.usuarioRed - 1] != "" && datosConsolidado[i][indicesColumnas.consolidado.envioCorreoSorpresaAuto - 1] == "") {
//         for (var registro in valorSucursal) {
//           if (valorSucursal[registro][0] != datosConsolidado[i][indicesColumnas.consolidado.codigoSucursal - 1]) continue;
//           correosSucursal = valorSucursal[registro][2]
//         }

//         let filial = datosConsolidado[i][indicesColumnas.consolidado.compania - 1]

//         nombreHojaConsolidado.getRange(i + 1, indicesColumnas.consolidado.envioCorreoSorpresaAuto).setValue(new Date())

//         var correoPersonal = datosConsolidado[i][indicesColumnas.consolidado.correo - 1]

//         var asunto = obtenerImagenesCorreo("asuntoKitNoDespachado", filial);
//         asunto = CigoApp.reemplazarValores(asunto, datosConsolidado[i]);

//         var mensaje = obtenerImagenesCorreo("mensajeKitNoDespachado", filial);
//         mensaje = CigoApp.reemplazarValores(mensaje, datosConsolidado[i]);

//         var inlineImages = obtenerImagenesCorreo("mensajeKitNoDespachado");

//         enviarCorreo(correoPersonal, asunto, mensaje, { inlineImages: inlineImages });
//         registrarLogCorreo(datosConsolidado[i], correo.toString(), asunto)
//       }
//     }
//   }
// }


// function encontrarPorcentajesCursos() {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();
//   var hojaConsolidado = SpreadsheetApp.openById(obtenerParametro("idHojaPlano")).getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var datosSps = hojaConsolidado.getDataRange().getValues();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaReporteAprendizaje"));
//   var hojaInduccionCorporativa = sps.getSheetByName(obtenerParametro("nombreHojaInduccionCorportiva"));
//   var datosInduccionCorporativa = hojaInduccionCorporativa.getDataRange().getValues();
//   var hojaAntisobornoAnticorrupcion = sps.getSheetByName(obtenerParametro("nombreHojaAntisobornoAnticorrupcion"));
//   var datosAntisobornoAnticorrupcion = hojaAntisobornoAnticorrupcion.getDataRange().getValues();
//   var hojaCovid19 = sps.getSheetByName(obtenerParametro("nombreHojaCovid19"));
//   var datosCovid19 = hojaCovid19.getDataRange().getValues();
//   var hojaDataCracks = sps.getSheetByName(obtenerParametro("nombreHojaDataCracks"));
//   var datosDataCracks = hojaDataCracks.getDataRange().getValues();
//   var hojaSarlaft = sps.getSheetByName(obtenerParametro("nombreHojaSarlaft"));
//   var datosSarlaft = hojaSarlaft.getDataRange().getValues();
//   var hojaFatca = sps.getSheetByName(obtenerParametro("nombreHojaFatca"));
//   var datosFatca = hojaFatca.getDataRange().getValues();
//   var hojaTBC = sps.getSheetByName(obtenerParametro("nombreHojaTBC"));
//   var datosTBC = hojaTBC.getDataRange().getValues();
//   var hojaUnMundo = sps.getSheetByName(obtenerParametro("nombreHojaUnMundo"));
//   var datosUnMundo = hojaUnMundo.getDataRange().getValues();
//   var hojaVivimosLosCinco = sps.getSheetByName(obtenerParametro("nombreHojaVivimosLosCinco"));
//   var datosVivimosLosCinco = hojaVivimosLosCinco.getDataRange().getValues();

//   for (var fila = 1; fila < datosSps.length; fila++) {
//     var cedula = datosSps[fila][indicesColumnas.consolidado.numCedula - 1];
//     console.log(cedula)
//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosInduccionCorporativa, 1, 6);

//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.teEstabamosEsperando).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosAntisobornoAnticorrupcion, 1, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.aprendeYPasaLaVoz).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosCovid19, 4, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.covidEnDavivienda).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosDataCracks, 1, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.dataCracks).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosSarlaft, 1, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.reentrenamientoSARLAFT).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosFatca, 1, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.fortaleceTuConocimiento).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosTBC, 1, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.todoBajoControl).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosUnMundo, 1, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.unMundo).setValue(porcentaje)
//     }

//     var porcentaje = CigoApp.buscarPorLlave(cedula, datosVivimosLosCinco, 4, 6);
//     if (porcentaje != null) {
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.vivimosYCuidamosLosCinco).setValue(porcentaje)
//     }

//   }

// }

// function encontrarRolesInstalar() {

//   parametros = obtenerParametros();
//   indicesColumnas = obtenerIndicesColumnas();
//   var hojaConsolidado = SpreadsheetApp.openById(obtenerParametro("idHojaPlano")).getSheetByName(obtenerParametro("nombreHojaPlano"));
//   var datosSps = hojaConsolidado.getDataRange().getValues();

//   var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlantaSemanalOnBoarding"));
//   var hojaAplicativosRol = sps.getSheetByName(obtenerParametro("nombreHojaAplicativosRol"));
//   var aplicativosRol = hojaAplicativosRol.getDataRange().getValues();

//   for (var fila = 1; fila < datosSps.length; fila++) {
//     var codigo = datosSps[fila][191];
//     var acumuladorRol = "";
//     if (codigo != "#N/A" && codigo != "") {
//       var codProcesado = codigo.replace(/\s+/g, '');
//       for (var filaRol = 1; filaRol < aplicativosRol.length; filaRol++) {
//         var valorActual = aplicativosRol[filaRol][1];
//         var resultado = valorActual.indexOf(codProcesado);
//         if (resultado >= 0) {
//           var rol = aplicativosRol[filaRol][0];
//           acumuladorRol = acumuladorRol + " // " + rol;
//           console.log("acumuladorRol para el valor : " + codProcesado + " es : " + acumuladorRol);
//         }
//       }
//       hojaConsolidado.getRange(fila + 1, indicesColumnas.consolidado.aplicativosInstalar).setValue(acumuladorRol)
//     }
//   }
// }

