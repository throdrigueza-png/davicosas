

function leerCorreo() {
  parametros = obtenerParametros();
  indicesColumnas = obtenerIndicesColumnas();

  var fechaHoy = new Date()
  var fechaAyer = sumarDias(fechaHoy, -1)
  var fechaCorte = formatearFecha(fechaAyer, "America/Bogota", "yyyy/MM/dd")

  var threads = GmailApp.search('is:unread after:' + fechaCorte);

  var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
  var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
  var ultimaColumna = hojaConsolidado.getLastColumn();

  var correosConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.correo, hojaConsolidado.getLastRow(), 1).getValues();
  const indicesPorCorreo = obtenerIndicesPorCorreo(correosConsolidado)
  // console.log("FILA", indicesPorCorreo["EDWARIOSTO@GMAIL.COM"])

  var carpetaRepositorio = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaRepositorioArchivos")));

  var hojaSucursales = sps.getSheetByName(obtenerParametro("nombreHojaSucursales"));

  var dataSucursales = hojaSucursales.getDataRange().getDisplayValues();



  for (var i = 0; i < threads.length; i++) {

    var msgs = threads[i].getMessages();
    var msg = msgs[msgs.length - 1];

    var remitente = msg.getFrom();
    var adjuntos = msg.getAttachments();

    // console.log(remitente)

    var correoRemitente = remitente.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi);
    if (correoRemitente.length == 0) continue

    var fila = indicesPorCorreo[correoRemitente.toString().trim().toUpperCase()]
    if (!fila) continue;

    var infoFila = hojaConsolidado.getRange(fila, 1, 1, ultimaColumna).getValues()[0];

    var cedula = infoFila[indicesColumnas.consolidado.numCedula - 1];
    var compania = infoFila[indicesColumnas.consolidado.compania - 1]
    // var codigo_sucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1]

    var codigoSucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1];
    var carpetaRegistro = CigoApp.crearObtenerCarpeta(carpetaRepositorio, [codigoSucursal, cedula]);


    // for (var j = 0; j < adjuntos.length; j++) {
    //   carpetaRegistro.createFile(adjuntos[j]);
    // }

    var urlCarpeta = carpetaRegistro.getUrl();
    hojaConsolidado.getRange(fila, indicesColumnas.consolidado.linkCarpDigital).setValue(urlCarpeta);
    infoFila[indicesColumnas.consolidado.linkCarpDigital - 1] = urlCarpeta;

    var existeUrl = infoFila[indicesColumnas.consolidado.linkCarpDigital - 1];
    existeUrl = existeUrl == "" ? false : true;

    var correoSucursal = CigoApp.buscarPorLlave(codigoSucursal, dataSucursales, 1, 3);

    if (existeUrl) {

      var asuntoContratacion = obtenerComunicacionPorFilial("asuntoReenvioaContratacion", compania);
      asuntoContratacion = CigoApp.reemplazarValores(asuntoContratacion, infoFila);

      var mensajeContratacion = obtenerComunicacionPorFilial("mensajeReenvioaContratacion", compania);
      mensajeContratacion = CigoApp.reemplazarValores(mensajeContratacion, infoFila);
    } else {
      var asuntoContratacion = obtenerComunicacionPorFilial("asuntoInicialContratacion", compania);
      asuntoContratacion = CigoApp.reemplazarValores(asuntoContratacion, infoFila);

      var mensajeContratacion = obtenerComunicacionPorFilial("mensajeInicialContratacion", compania);
      mensajeContratacion = CigoApp.reemplazarValores(mensajeContratacion, infoFila);
    }

    var inlineImages = obtenerImagenesCorreo("mensajeInicialContratacion", compania);

    // if (correoSucursal)CigoApp.enviarCorreo(correoSucursal, auxAsunto, auxMensaje, {});
    if (correoSucursal) enviarCorreo(correoSucursal, asuntoContratacion, mensajeContratacion, { inlineImages: inlineImages });

    console.log("Mensaje leido de " + remitente + "\nCarpeta creada: " + urlCarpeta)

    msg.markRead();
    threads[i].markRead();
  }
}


function obtenerIndicesPorCorreo(correos) {
  const indicesPorCorreo = correos.reduce((accumuladoCorreos, filaCorreo, indice) => {

    const correo = filaCorreo[0].toString().toUpperCase()

    accumuladoCorreos[correo] = indice + 1

    return accumuladoCorreos
  }, {})
  return indicesPorCorreo
}

function leerCorreoConsultor() {
  parametros = obtenerParametros();
  indicesColumnas = obtenerIndicesColumnas();

  var etiquetaSeleccionados = obtenerParametro("etiquetaSeleccion");

  var threads = GmailApp.search('is:unread is:' + etiquetaSeleccionados);


  var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
  var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
  var cedulasConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.numCedula, hojaConsolidado.getLastRow(), 1).getValues();


  var carpetaRepositorio = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaRepositorioArchivos")));
  var ultimaColumna = hojaConsolidado.getLastColumn();



  var hojaFiliales = parametros.spsParametros.getSheetByName(obtenerParametro("nombreHojaFiliales"));
  var dataFiliales = hojaFiliales.getDataRange().getDisplayValues();

  let cedulas_no_encontradas = []

  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    var msg = msgs[msgs.length - 1];
    var asunto = msg.getSubject().trim();
    // var cuerpo = msg.getPlainBody();
    // var fechaCorreo = msg.getDate();
    var remitente = msg.getFrom();
    var adjuntos = msg.getAttachments();

    // var correoRemitente = remitente.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi);
    // if (correoRemitente.length == 0) {
    //   console.warn("No se encontró el correo en el remitente: " + remitente);
    //   continue;
    // }


    var cedulaAsunto = asunto.split(" ")[2]
    if (!cedulaAsunto) continue;
    cedulaAsunto = cedulaAsunto.toString().trim();
    cedulaAsunto = cedulaAsunto.replace(/\D/gi, "");


    var fila = CigoApp.obtenerFila(cedulaAsunto, cedulasConsolidado, 1);
    if (!fila) {
      cedulas_no_encontradas.push(cedulaAsunto)
      // console.warn("No se encontró el registro con la cedula " + cedulaAsunto);
      continue;
    }

    var infoFila = hojaConsolidado.getRange(fila, 1, 1, ultimaColumna).getValues()[0];
    var cedula = infoFila[indicesColumnas.consolidado.numCedula - 1];
    var compania = infoFila[indicesColumnas.consolidado.compania - 1]

    var codigo_sucursal = infoFila[indicesColumnas.consolidado.codigoSucursal - 1]

    var carpetaRegistro = CigoApp.crearObtenerCarpeta(carpetaRepositorio, [codigo_sucursal, cedula]);

    for (var j = 0; j < adjuntos.length; j++) {
      carpetaRegistro.createFile(adjuntos[j]);
    }

    var urlCarpeta = carpetaRegistro.getUrl();
    hojaConsolidado.getRange(fila, indicesColumnas.consolidado.linkCarpDigital).setValue(urlCarpeta);
    infoFila[indicesColumnas.consolidado.linkCarpDigital - 1] = urlCarpeta;

    var compania = infoFila[indicesColumnas.consolidado.compania - 1];

    // msg.markRead();
    // threads[i].markRead();

    console.log("Mensaje leido de " + remitente + "\nCarpeta creada: " + urlCarpeta + "\Numero de documento: " + cedula)
    msg.markRead();
    threads[i].markRead();


    // if (compania != "DAVIVIENDA - HCM") {
    //   var datosFilial = CigoApp.buscarPorLlave(compania, dataFiliales, 1, -1);
    //   if (!datosFilial) continue;
    //   var correoFilial = datosFilial[1];
    //   if (!correoFilial) continue;


    //   var asuntoFilial = obtenerComunicacionPorFilial("asuntoFilial", compania);
    //   asuntoFilial = CigoApp.reemplazarValores(asuntoFilial, infoFila);

    //   var mensajeFilial = obtenerComunicacionPorFilial("mensajeFilial", compania);
    //   mensajeFilial = CigoApp.reemplazarValores(mensajeFilial, infoFila)

    //   var inlineImages = obtenerImagenesCorreo("mensajeEnvioFilial", compania);

    //   enviarCorreo(correoFilial, asuntoFilial, mensajeFilial, { inlineImages: inlineImages });
    // }
  }
  console.warn(`Cedulas no encontradas: ${cedulas_no_encontradas}`)

}

function leerCorreoUsuarioRed() {
  parametros = obtenerParametros();
  indicesColumnas = obtenerIndicesColumnas();

  var correoRemitenteFiltro = obtenerParametro("correoRemitenteUsuarios");
  var threads = GmailApp.search('is:unread from: ' + correoRemitenteFiltro);

  var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
  var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
  var cedulasConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.numCedula, hojaConsolidado.getLastRow(), 1).getValues();

  let lista_correos_no_encontrados = []
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    var msg = msgs[msgs.length - 1];
    var asunto = msg.getSubject().trim();
    var cedula = asunto.split("Alta usuario")

    if (!cedula[1]) {
      console.warn("No se encontró la cedula en el asunto: " + asunto)
      continue;
    }
    cedula = cedula[1].toString().trim();

    var cuerpo = msg.getPlainBody();
    var usuarioRed = cuerpo.split("Usuario de red directorio activo:")[1].split("\n")[0];
    usuarioRed = usuarioRed.replace(/\./g, "");
    var cuentaCorreo = cuerpo.split("Cuenta de correo:")[1].split("\n")[0];

    // var fechaCorreo = msg.getDate();
    // var remitente = msg.getFrom();
    // var adjuntos = msg.getAttachments();

    // var correoRemitente = remitente.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi);
    // if (correoRemitente.length == 0) {
    //   console.warn("No se encontró el correo en el remitente: " + remitente);
    //   continue;
    // }


    var fila = CigoApp.obtenerFila(cedula, cedulasConsolidado, 1);
    if (!fila) {
      lista_correos_no_encontrados.push(cedula)
      // msg.markRead();
      // threads[i].markRead();
      // console.warn("No se encontró el registro con la cedula :" + cedula);
      continue;
    }

    hojaConsolidado.getRange(fila, indicesColumnas.consolidado.usuarioRed).setValue(usuarioRed);
    hojaConsolidado.getRange(fila, indicesColumnas.consolidado.correoElectronico).setValue(cuentaCorreo);


    // msg.markRead();
    // threads[i].markRead();
    console.log("Mensaje leido cedula " + cedula + "\nUsuario Red: " + usuarioRed + "\nCuenta Correo: " + cuentaCorreo)
    msg.markRead();
    threads[i].markRead();
  }

  console.warn(`Lista de correos no encontrados en el consolidado: ${lista_correos_no_encontrados} `)
}

function leerCorreoBaseTaleo() {

  parametros = obtenerParametros();
  indicesColumnas = obtenerIndicesColumnas();
  console.time("Actualizar Backup")
  // actualizarInformacionBackUp()
  console.timeEnd("Actualizar Backup")

  
  var fechaActual = new Date().getHours()
  if (fechaActual == 23) return

  var fechaHoy = new Date()
  var fechaAyer = sumarDias(fechaHoy, 0)
  var fechaCorte = formatearFecha(fechaAyer, "America/Bogota", "yyyy/MM/dd")

  var correoRemitenteFiltro = obtenerParametro("correoRemitenteBaseTaleo");
  var threads = GmailApp.search('is:unread has:attachment after:' + fechaCorte + 'from:' + correoRemitenteFiltro);
  // var threads = GmailApp.search('is:unread is:Prueba');

  var carpetaArchivosTaleo = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaArchivosTaleo")));

  console.log("threads.length", threads.length);

  // return;


  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var iMsg = 0; iMsg < msgs.length; iMsg++) {

      var msg = msgs[iMsg];
      if(!msg.isUnread()) continue;

      var asunto = msg.getSubject().trim();

      var fechaCorreo = msg.getDate();
      var remitente = msg.getFrom();
      var adjuntos = msg.getAttachments();

      var archivo = adjuntos[0];

      try {
        console.log("Procesando thread", i, "message", iMsg)
        var archivoSheets = crearArchivoSheets(archivo, carpetaArchivosTaleo);
        if (!archivoSheets) continue;
        var sps = SpreadsheetApp.openById(archivoSheets.id).getSheets()[0];
        var datosTaleo = sps.getDataRange().getValues();
        datosTaleo.shift();

        console.log("Registros Totales", datosTaleo.length)

        console.time("Agregar datos")
        agregarDatosPlano(datosTaleo);
        console.timeEnd("Agregar datos")
      } catch (e) {
        console.warn("thread", i, "message", iMsg, "Error", e)
      }

      console.log("Mensaje leido base Taleo");

      msg.markRead();
    }
    threads[i].markRead();
  }

}

var idReqCedulaAgregar = [], datosConsolidado, datosBackUp;
function agregarDatosPlano(datosTaleo) {

  if (!parametros) parametros = obtenerParametros()

  if (!datosConsolidado || !datosBackUp) {
    var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
    var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
    var datosConsolidado = hojaConsolidado.getDataRange().getValues()

    var spsBackUp = SpreadsheetApp.openByUrl(obtenerParametro("urlBackUp"))
    var hojaConsolidadoBackUp = spsBackUp.getSheetByName(obtenerParametro("nombreHojaBackUp"))
    var datosBackUp = hojaConsolidadoBackUp.getDataRange().getValues()
  }
  

  // var indicesBackUp = datosBackUp.reduce(function(indices, fila){
  //   var cedula = fila[50]
  //   var requision = fila[1]
  //   if(!indices[requision]){
  //     indices[requision] = {}
  //   }
  //   indices[requision][cedula] = true
  //   return indices
  // }, {})

  // var datosConsolidado = hojaConsolidado.getRange(1,1,hojaConsolidado.getLastRow(), indicesColumnas.consolidado.numCedula).getValues();
  // datosConsolidado.shift();

  var registrosAgregar;

  if (datosBackUp.length == 0) {
    registrosAgregar = datosTaleo;
  } else {

    // var datosConsolidado = ordenarDatos(datosConsolidado, indicesColumnas.consolidado.numCedula, false);

    registrosAgregar = [];

    var cedulaTaleo, idReqTaleo, datosFuncionario;
    for (var iTaleo in datosTaleo) {

      cedulaTaleo = datosTaleo[iTaleo][36];
      idReqTaleo = datosTaleo[iTaleo][0];
      var idReqCedula = idReqTaleo + "-" + cedulaTaleo;
      if (!cedulaTaleo || !idReqTaleo) continue;
      encontrado = false
      // var encontrado = !indicesBackUp[idReqTaleo][cedulaTaleo] ? false : truex
      for (var iCon in datosBackUp) {
        var cedulaConsolidado = datosBackUp[iCon][50];
        var idReqConsolidado = datosBackUp[iCon][1];

        if (cedulaTaleo == cedulaConsolidado && idReqTaleo == idReqConsolidado) {
          encontrado = true;
          break;
        }
      }

      if (idReqCedulaAgregar.includes(idReqCedula)) encontrado = true
      else idReqCedulaAgregar.push(idReqCedula)

      if (!encontrado) registrosAgregar.push(datosTaleo[iTaleo]);
    }

  }

  // ---------------------------------

  // var cedulasConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.numCedula, hojaConsolidado.getLastRow(), 1).getValues();
  // cedulasConsolidado.shift();

  // var registrosAgregar;

  // if (cedulasConsolidado.length == 0) {
  //   registrosAgregar = datosTaleo;
  // } else {
  //   var cedulasConsolidado = ordenarDatos(cedulasConsolidado, 1, false);

  //   registrosAgregar = [];

  //   var cedulaTaleo, datosFuncionario;
  //   for (var iTaleo in datosTaleo) {
  //     cedulaTaleo = datosTaleo[iTaleo][36];
  //     if(!cedulaTaleo) continue;
  //     datosFuncionario = buscarPorLlaveBinario(cedulaTaleo, cedulasConsolidado,1,1,false);

  //     if (datosFuncionario != null) continue;
  //     registrosAgregar.push(datosTaleo[iTaleo]);

  //   }

  // }

  // ---------------------------------

  console.log("Cantidad Datos Agregar", registrosAgregar.length);

  if (registrosAgregar.length == 0) return;

  var datosAgregarPlano = organizarDatosAgregar(registrosAgregar);

  var ultimaFila = hojaConsolidado.getLastRow()
  hojaConsolidado.getRange(ultimaFila + 1, 1, datosAgregarPlano.length, datosAgregarPlano[0].length).setValues(datosAgregarPlano);
  datosConsolidado = datosConsolidado.concat(datosAgregarPlano)

  //Copia las formulas de las columnas especificadas en la fila del nuevo registro
  var columnasFormulas = obtenerParametro("columnasFormulas").toString().split(",");

  if (columnasFormulas != "") {
    var rangoFormula;
    for (var i = 0; i < columnasFormulas.length; i++) {
      rangoFormula = hojaConsolidado.getRange((columnasFormulas[i].trim()) + 2);
      var numeroColumna = obtenerNumeroColumna(columnasFormulas[i].trim());
      // rangoFormula.copyTo(hojaConsolidado.getRange((columnasFormulas[i].trim()) + ultimaFila ));
      rangoFormula.copyTo(hojaConsolidado.getRange(ultimaFila + 1, numeroColumna, registrosAgregar.length, 1));
    }
  }
}


function organizarDatosAgregar(registrosAgregar) {

  const carpeta_por_identificacion = obtenerUltimaCarpetaPorIdentificacion()
  var carpetaRepositorio = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaRepositorioArchivos")));

  var datosColumnas = parametros.spsParametros.getSheetByName("Columnas Taleo - Plano").getDataRange().getValues();
  datosColumnas.shift();

  var fechaActual = new Date();

  for (var f in datosColumnas) {
    for (var c in datosColumnas[f]) {
      datosColumnas[f][c] = obtenerNumeroColumna(datosColumnas[f][c]);
    }
  }

  var registrosOrdenados = [];
  var registro;

  for (var iTaleo in registrosAgregar) {
    registro = [];
    for (var iCol in datosColumnas) {
      var columnaPlano = datosColumnas[iCol][1];
      var columnaTaleo = datosColumnas[iCol][0];

      // var salario = 

      if (columnaPlano == 2) registro[columnaPlano - 1] = "'" + registrosAgregar[iTaleo][columnaTaleo - 1]
      else if (columnaPlano == indicesColumnas.consolidado.fechaExpCedula) {
        var fechaFormatoMesDiaAño = registrosAgregar[iTaleo][columnaTaleo - 1].split("-")
        registro[columnaPlano - 1] = fechaFormatoMesDiaAño[1] + "/" + fechaFormatoMesDiaAño[0] + "/" + fechaFormatoMesDiaAño[2]
      } else {
        registro[columnaPlano - 1] = registrosAgregar[iTaleo][columnaTaleo - 1]
      }

    }

    var gestion = "Sin Gestion";
    // if (registro[indicesColumnas.consolidado.compania - 1] == "DAVIVIENDA - HCM") gestion = "Sin Gestion"
    // else gestion = "No Aplica";
    registro[indicesColumnas.consolidado.id - 1]
    registro[indicesColumnas.consolidado.id - 1] = "ONB" + new Date().getTime();
    registro[indicesColumnas.consolidado.estadoGeneral - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionContratacion - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionAlistamiento - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionAprendizaje - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionCarnetizacion - 1] = gestion;
    registro[indicesColumnas.consolidado.fechaIngresoBase - 1] = fechaActual;

    Utilities.sleep(2);

    for (var r = 0; r < registro.length; r++) {
      if (!registro[r]) registro[r] = "";
    }


    var cedula = registro[indicesColumnas.consolidado.numCedula - 1];
    var codigo_sucursal = registro[indicesColumnas.consolidado.codigoSucursal - 1]

    if (!cedula || !codigo_sucursal) {
      registrosOrdenados.push(registro)
      continue
    }

    const nueva_carpeta = CigoApp.crearObtenerCarpeta(carpetaRepositorio, [codigo_sucursal, cedula])
    registro[indicesColumnas.consolidado.fechaIngresoBase - 1] = fechaActual;
    registro[indicesColumnas.consolidado.linkCarpDigital - 1] = nueva_carpeta.getUrl();

    registrosOrdenados.push(registro);

    if (!carpeta_por_identificacion[cedula]) continue

    // try {
    const carpeta_existente = DriveApp.getFolderById(CigoApp.obtenerIdUrl(carpeta_por_identificacion[cedula]))
    copiarCarpeta(carpeta_existente, nueva_carpeta)

    // } catch (e) {
    //   console.warn(e)
    // }

  }

  return registrosOrdenados;
}

function leerCorreoHistoricoRequisisionesTaleo() {
  parametros = obtenerParametros();
  indicesColumnas = obtenerIndicesColumnas();

  var etiquetaHistoricoRequisiciones = obtenerParametro("etiquetaHistoricoRequisiciones");
  var threads = GmailApp.search('is:unread is:' + etiquetaHistoricoRequisiciones);
  // var threads = GmailApp.search('is:unread is:Prueba');

  var carpetaArchivosTaleo = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaArchivosTaleo")));

  console.log("threads.length", threads.length);

  var spsHistorico = SpreadsheetApp.openByUrl(obtenerParametro("archivoHistoricoRequisisiones")).getSheets()[0];


  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    var msg = msgs[msgs.length - 1];
    var asunto = msg.getSubject().trim();

    var fechaCorreo = msg.getDate();
    var remitente = msg.getFrom();
    var adjuntos = msg.getAttachments();

    var archivo = adjuntos[0];

    var archivoSheets = crearArchivoSheets(archivo, carpetaArchivosTaleo);
    if (!archivoSheets) continue;
    var sps = SpreadsheetApp.openById(archivoSheets.id).getSheets()[0];
    var datosTaleo = sps.getDataRange().getValues();
    if (datosTaleo.length == 0) continue;
    datosTaleo.shift();
    datosTaleo.shift();
    datosTaleo.shift();
    if (datosTaleo.length == 0) continue;

    // spsHistorico.getDataRange().clear();
    spsHistorico.getRange(2, 1, datosTaleo.length, datosTaleo[0].length).clear().setValues(datosTaleo);



    msg.markRead();
    threads[i].markRead();

    console.log("Mensaje leido base Histórico Requisiciones");
  }
}
