function inicializarWebCreacion() {
  parametros = obtenerParametros()
  indicesColumnas = obtenerIndicesColumnas()

  obtenerEmailsPermitidos("Ingreso");

  return {
    indicesColumnas,
    listas: obtenerListas(),
    urlPaginaHomeGlobal: obtenerParametro("urlPaginaHome")
  }
}

function cargarRegistroIndividual(formulario) {
  parametros = obtenerParametros()
  indicesColumnas = obtenerIndicesColumnas()

  let registro = []
  for (const indice in indicesColumnas.consolidado) {
    const valor = formulario[indice] || ""
    registro[indicesColumnas.consolidado[indice] - 1] = valor
  }

  for (const col in formulario) {
    if (formulario[col]) continue
    formulario[col] = ""
  }

  //Indica que este registro fue añadido Manualmente desde Onboarding ya que se requiere para otros procesos
  registro[indicesColumnas.consolidado.esManual - 1] = true
  //Indica el correo que registro esta persona
  registro[indicesColumnas.consolidado.registradoPor - 1] = Session.getActiveUser().getEmail()

  const gestion = "Sin Gestion"
  registro[indicesColumnas.consolidado.id - 1] = `ONB${new Date().getTime()}`
  registro[indicesColumnas.consolidado.estadoGeneral - 1] = gestion;
  registro[indicesColumnas.consolidado.gestionContratacion - 1] = gestion;
  registro[indicesColumnas.consolidado.gestionAlistamiento - 1] = gestion;
  registro[indicesColumnas.consolidado.gestionAprendizaje - 1] = gestion;
  registro[indicesColumnas.consolidado.gestionCarnetizacion - 1] = gestion;
  registro[indicesColumnas.consolidado.fechaIngresoBase - 1] = new Date();

  // const es_sena = registro[indicesColumnas.consolidado.catCargo - 1].toString().toUpperCase().trim() === "APRENDIZ SENA"
  //   ? true
  //   : false

  // if (es_sena) {
  //   const fecha_inicio_lectiva = registro[indicesColumnas.consolidado.fechaIniEtapaLecti - 1]
  //   const fecha_fin_lectiva = registro[indicesColumnas.consolidado.fechaFinEtapaLecti - 1]
  //   const fecha_inicio_productiva = registro[indicesColumnas.consolidado.fechaIniEtapaPrac - 1]
  //   const fecha_fin_productiva = registro[indicesColumnas.consolidado.fechaFinEtapaPrac - 1]

  //   if (fecha_inicio_lectiva && fecha_fin_lectiva && fecha_inicio_productiva && fecha_fin_productiva) {
  //     registro[indicesColumnas.consolidado.gestionAlistamiento - 1] = "No aplica";
  //     registro[indicesColumnas.consolidado.gestionAprendizaje - 1] = "No aplica";
  //     registro[indicesColumnas.consolidado.gestionCarnetizacion - 1] = "No aplica";
  //   }
  // }


  const sps_consolidado = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
  const hoja_consolidado = sps_consolidado.getSheetByName(obtenerParametro("nombreHojaPlano"))
  const ultima_columna_consolidado = hoja_consolidado.getLastRow()

  hoja_consolidado.getRange(ultima_columna_consolidado + 1, 1, 1, registro.length).setValues([registro])
}


function cargarRegistrosMasivos(formulario) {
  parametros = obtenerParametros()
  indicesColumnas = obtenerIndicesColumnas()

  if (!formulario.plano_masivo) throw "No has cargado ningun documento"


  const columnas_extraer_crear_masivos = CigoApp.obtenerNumeroColumna(obtenerParametro("columnasExtraerCrearMasivos").split(","))
  // const columnas_agregar_crear_masivos = CigoApp.obtenerNumeroColumna(obtenerParametro("columnasAgregarCrearMasivos").split(","))

  const carpeta_auxiliar = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaAuxiliarGenerarArchivos")))
  const archivo_hoja_calculo = convertirDocumentoFormatoDrive(carpeta_auxiliar, formulario.plano_masivo)

  const sps_archivos_nuevos = SpreadsheetApp.openById(archivo_hoja_calculo.getId())
  const hoja_plano = sps_archivos_nuevos.getSheets()[0]
  let datos_a_subir = hoja_plano.getDataRange().getValues()
  datos_a_subir.shift()

  const carpetas_por_identificacion = obtenerUltimaCarpetaPorIdentificacion()

  datos_a_subir = datos_a_subir.reduce((registros_validos, registro_sin_ordenar, index) => {

    let registro = []

    registro_sin_ordenar.forEach(function (columna, indice) {
      registro[columnas_extraer_crear_masivos[indice] - 1] = columna
    })

    for (var i in registro) {
      if (registro[i]) continue
      registro[i] = ""
    }

    // registro = ordenarRegistro(registro, columnas_agregar_crear_masivos, columnas_extraer_crear_masivos)
    if (!registro[indicesColumnas.consolidado.numCedula - 1]) return registros_validos

    //Indica que este registro fue añadido Manuealmente desde Onboarding ya que se requiere para otros procesos
    registro[indicesColumnas.consolidado.esManual - 1] = true
    //Indica el correo que registro esta persona
    registro[indicesColumnas.consolidado.registradoPor - 1] = Session.getActiveUser().getEmail()

    const numero_identificacion = registro[indicesColumnas.consolidado.numCedula - 1]
    if (!numero_identificacion) return registros_validos


    //Inicializa el registro
    const gestion = "Sin Gestion"
    registro[indicesColumnas.consolidado.id - 1] = `ONB${new Date().getTime()}${index}`
    registro[indicesColumnas.consolidado.estadoGeneral - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionContratacion - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionAlistamiento - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionAprendizaje - 1] = gestion;
    registro[indicesColumnas.consolidado.gestionCarnetizacion - 1] = gestion;
    registro[indicesColumnas.consolidado.fechaIngresoBase - 1] = new Date();


    //Si no tiene carpeta es porque esta persona nunca ha existido en TALENT ON
    const url_carpeta = carpetas_por_identificacion[numero_identificacion]
    console.log(url_carpeta)
    if (!url_carpeta) return registros_validos

    registro[indicesColumnas.consolidado.linkCarpDigital - 1] = url_carpeta

    registros_validos.push(registro)

    return registros_validos
  }, [])


  const sps_consolidado = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
  const hoja_consolidado = sps_consolidado.getSheetByName(obtenerParametro("nombreHojaPlano"))
  const ultima_columna_consolidado = hoja_consolidado.getLastRow()
  hoja_consolidado.getRange(ultima_columna_consolidado + 1, 1, datos_a_subir.length, datos_a_subir[0].length).setValues(datos_a_subir)

}

function busquedaMasiva(identificaciones_usuarios) {

  const correo_usuario = Session.getActiveUser().getEmail()

  parametros = obtenerParametros()
  indicesColumnas = obtenerIndicesColumnas()

  const sps_consolidado = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
  const hoja_consolidado = sps_consolidado.getSheetByName(obtenerParametro("nombreHojaPlano"))

  const ultima_columna_consolidado = hoja_consolidado.getLastColumn()
  const columna_cedula_consolidado = hoja_consolidado.getRange(1, indicesColumnas.consolidado.numCedula, ultima_columna_consolidado, 1).getValues()


  const sps_back_up = SpreadsheetApp.openByUrl(obtenerParametro("urlBackUp"))
  const hoja_back_up = sps_back_up.getSheetByName(obtenerParametro("nombreHojaBackUp"))

  const ultima_columna_back_up = hoja_consolidado.getLastColumn()
  const columna_cedula_back_up = hoja_back_up.getRange(1, indicesColumnas.consolidado.numCedula, ultima_columna_back_up, 1).getValues()


  const columnas_extraer_crear_masivos = CigoApp.obtenerNumeroColumna(obtenerParametro("columnasExtraerCrearMasivos").split(","))
  // const columnas_agregar_crear_masivos = CigoApp.obtenerNumeroColumna(obtenerParametro("columnasAgregarCrearMasivos").split(","))

  const indices_num_documento_con = obtenerIndicesIdentificacion(columna_cedula_consolidado)
  const indices_num_documento_back = obtenerIndicesIdentificacion(columna_cedula_back_up)

  const columnas_extraer_backup = CigoApp.obtenerNumeroColumna(obtenerParametro("ColumnasExtraer").split(","))

  let registros_encontrados = identificaciones_usuarios.reduce((registros, identificacion) => {

    let registro = []
    if (indices_num_documento_con[identificacion]) {
      registro = hoja_consolidado.getRange(indices_num_documento_con[identificacion], 1, 1, ultima_columna_consolidado).getValues()[0]
    } else if (indices_num_documento_back[identificacion]) {
      registro = hoja_back_up.getRange(indices_num_documento_back[identificacion], 1, 1, ultima_columna_back_up).getValues()[0]
      registro = ordenarRegistro(registro, columnas_extraer_backup)
    }

    if (registro.length === 0) return registros

    registro = ordenarRegistro(registro, columnas_extraer_crear_masivos)
    registros.push(registro)

    return registros
  }, []);

  if (registros_encontrados.length === 0) "No hemos encontrado ningun numero de documento en la base"


  const carpeta_auxiliar = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaAuxiliarGenerarArchivos")))

  const spread_para_copiar = DriveApp.getFileById(CigoApp.obtenerIdUrl(obtenerParametro('urlPlantillaCrearMasivos')))
  const archivo_copia = spread_para_copiar.makeCopy(`Plano nuevos ingresos`, carpeta_auxiliar)
  const sps_plano = SpreadsheetApp.openByUrl(archivo_copia.getUrl())
  const hoja_añadir = sps_plano.getSheets()[0]

  let encabezados = hoja_consolidado.getRange(1, 1, 1, ultima_columna_consolidado).getValues()[0]
  encabezados = ordenarRegistro(encabezados, columnas_extraer_crear_masivos)

  registros_encontrados.unshift(encabezados)
  hoja_añadir.getRange(1, 1, registros_encontrados.length, registros_encontrados[0].length).setValues(registros_encontrados)
  SpreadsheetApp.flush()

  const blob_excel = convertirAExcel(sps_plano.getId(), carpeta_auxiliar)
  archivo_copia.setTrashed(true)

  const asunto = obtenerParametro("asuntoCrearCamposMasivo")
  const mensaje = obtenerParametro("mensajeCrearCamposMasivo")

  CigoApp.enviarCorreo(correo_usuario, asunto, mensaje, {
    noReply: true,
    attachments: [blob_excel]
  })
}

function busquedaIndividual(identificacion_usuario) {
  if (!identificacion_usuario) "Debes diligenciar el campo: Numero de documento"


  parametros = obtenerParametros()
  indicesColumnas = obtenerIndicesColumnas()

  const sps_consolidado = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
  const hoja_consolidado = sps_consolidado.getSheetByName(obtenerParametro("nombreHojaPlano"))

  const ultima_columna_consolidado = hoja_consolidado.getLastColumn()
  const ultima_fila_consolidado = hoja_consolidado.getLastRow()
  const columna_cedula_consolidado = hoja_consolidado.getRange(1, indicesColumnas.consolidado.numCedula, ultima_fila_consolidado, 1).getValues()

  let indices_num_documento_con = obtenerIndicesIdentificacion(columna_cedula_consolidado)

  var posicion_usuario = indices_num_documento_con[identificacion_usuario]
  if (posicion_usuario) {
    const registro = hoja_consolidado.getRange(posicion_usuario, 1, 1, ultima_columna_consolidado).getValues()[0]
    
    for(const col in registro){
      const valor = registro[col]
      const valorFormateado = CigoApp.formatearFecha(valor)
      registro[col] =  valorFormateado
    }
    console.log(registro)
    return registro
  }

  const sps_back_up = SpreadsheetApp.openByUrl(obtenerParametro("urlBackUp"))
  const hoja_back_up = sps_back_up.getSheetByName(obtenerParametro("nombreHojaBackUp"))

  let ultima_columna_back_up = hoja_consolidado.getLastColumn()
  const columna_cedula_back_up = hoja_back_up.getRange(1, indicesColumnas.consolidado.numCedula, ultima_columna_back_up, 1).getValues()

  const columnas_extraer_backup = CigoApp.obtenerNumeroColumna(obtenerParametro("ColumnasExtraer").split(","))

  let indices_num_documento_back = obtenerIndicesIdentificacion(columna_cedula_back_up)

  posicion_usuario = indices_num_documento_back[identificacion_usuario]
  if (posicion_usuario) {
    let registro = hoja_back_up.getRange(posicion_usuario, 1, 1, ultima_columna_back_up).getValues()[0]

    for(const col in registro){
      const valor = registro[col]
      const valorFormateado = CigoApp.formatearFecha(valor)
      registro[col] =  valorFormateado
    }

    registro = ordenarRegistro(registro, columnas_extraer_backup)
    console.log(registro)
    return registro_ordenado
  }

  throw `No se ha encontrado el usuario con numero identificacion: ${identificacion_usuario}`

}



function obtenerIndicesIdentificacion(datos) {
  return datos.reduce((indices, registro, index) => {
    const identificacion = registro[0]
    indices[identificacion] = index + 1
    return indices
  }, {})
}



function ordenarRegistro(registro_a_ordenar, columnas_antiguas, columnas_nuevas) {
  const registro_ordenado = [];
  const usar_nuevas_posiciones = Boolean(columnas_nuevas);

  for (let i = 0; i < columnas_antiguas.length; i++) {
    const antigua_posicion = columnas_antiguas[i] - 1;
    const nueva_posicion = usar_nuevas_posiciones ? columnas_nuevas[i] - 1 : i;

    registro_ordenado[nueva_posicion] = registro_a_ordenar[antigua_posicion];
  }

  return registro_ordenado;
}




/**
 * Trasnforma un documento a un formato compatible con la GSUIT 
 * @param {DriveApp} carpeta - Carpeta donde estaran adjuntos los documentos
 * @param {BlobSource} blob - Documento a convertir a excel
 * @return {DriveApp.File} Documento compatible con DriveApp
 */
function convertirDocumentoFormatoDrive(carpeta, blob) {

  try {
    const recurso = {
      "title": blob.getName(),
      "parents": [{ "id": carpeta.getId() }]
    };

    const hoja_de_calculo = Drive.Files.insert(recurso, blob, {
      "convert": true,
      "supportsAllDrives": true
    });
    return hoja_de_calculo
  } catch (error) {
    throw `Ha ocurrido un error al subir el documento, vuelve a intentarlo mas tarde`
  }
}

