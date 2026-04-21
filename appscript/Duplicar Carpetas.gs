//© Banco Davivienda S.A. 2021. Se prohíbe su uso o reproducción sin previa autorización del Banco Davivienda S.A

/**
* Copia todos los archivos y carpetas que contiene la carpeta origen en la carpeta destino
*/
function copiarCarpeta(carpeta_origen, carpeta_destino){  
  var archivos = carpeta_origen.getFiles();
  var carpetas = carpeta_origen.getFolders();
  
  copiarArchivos(archivos,carpeta_destino);
  copiarCarpetas(carpetas,carpeta_destino);
}



/**
* Copia los archivos enviados en el parametros archivos en la carpeta destino
* @param {Iterator} archivos Iterador que contiene los archivos a ser copiados
* @param {Folder} destino Carpeta destino en donde se copiaran los archivos
*/
function copiarArchivos(archivos, destino){
  while(archivos.hasNext()){
    var archivo =  archivos.next();
    try{
      archivo.makeCopy(archivo.getName(), destino);
    } catch(e) {
      Browser.inputBox("Archivo no copiado", "No se pudo copiar el archivo " + archivo.getName() + ". El copiado continuará.", Browser.Buttons.OK);
    }
  }
}


/**
* Crea las carpetas enviadas en el parametro carpetas en la carpeta destino. En Gsuite las carpetas no se pueden copiar por los que se deben crear
* @param {Iterator} carpetas Iterador que contiene las carpetas a ser copiadas
* @param {Folder} destino Carpeta destino en donde se copiaran/crearan las carpetas
*/
function copiarCarpetas(carpetas,destino){
  while(carpetas.hasNext()){
    var carpetaOrigen = carpetas.next();
    var carpetaDestino = destino.createFolder(carpetaOrigen.getName());
    var archivosOrigen = carpetaOrigen.getFiles();
    copiarArchivos(archivosOrigen,carpetaDestino);
    copiarCarpetas(carpetaOrigen.getFolders(),carpetaDestino);
  }
}

 * Obtiene la ultima carpeta asignada a un numero de identificacion
 * @return {Object} ultima_carpeta_identificacion
 */
function obtenerUltimaCarpetaPorIdentificacion() {
  parametros = parametros || obtenerParametros()
  indicesColumnas = indicesColumnas || obtenerIndicesColumnas()

  const sps_backup = SpreadsheetApp.openByUrl(obtenerParametro("urlBackUp"))
  const hoja_backup = sps_backup.getSheetByName(obtenerParametro("nombreHojaBackUp"))
  const datos_backup = hoja_backup.getDataRange().getValues()

  let ultima_carpeta_identificacion = datos_backup.reduce((carpetas_identificacion, registro) => {
    const identificacion = registro[50]
    const url_carpeta = registro[109]
    if (!url_carpeta) return carpetas_identificacion

    carpetas_identificacion[identificacion] = url_carpeta
    return carpetas_identificacion
  }, {})

  const sps_consolidado = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"))
  const hoja_consolidado = sps_consolidado.getSheetByName(obtenerParametro("nombreHojaPlano"))
  const datos_consolidado = hoja_consolidado.getDataRange().getValues()

  ultima_carpeta_identificacion = datos_consolidado.reduce((carpetas_identificacion, registro) => {
    const identificacion = registro[indicesColumnas.consolidado.numCedula - 1]
    const url_carpeta = registro[indicesColumnas.consolidado.linkCarpDigital - 1]
    if (!url_carpeta) return carpetas_identificacion

    carpetas_identificacion[identificacion] = url_carpeta
    return carpetas_identificacion

  }, ultima_carpeta_identificacion)

  return ultima_carpeta_identificacion
}
