
/**
* Ordena un array de acuerdo a la columna especificada
* @param {Array} datos Datos a ordenar
* @param {Integer} columnaOrden Numero de la columna por la que se debe ordener. Empezando desde 1
* @param {Boolean} numerico Indica si la columna por la que se debe ordenar ordenar es númerica o no
* @return {Array} Array con las coincidencias encontradas
*/
function ordenarDatos(datos, columnaOrden, numerico) {

  columnaOrden = columnaOrden - 1;

  // var tipoPrimerValor = typeof datos[0][columnaOrden];
  // var numericoPrimerValor = tipoPrimerValor == "number" ? true: false;

  // numerico = numerico || numericoPrimerValor;


  if (numerico) {
    datos.sort(function (a, b) {
      return a[columnaOrden] - b[columnaOrden];
    });

  } else {

    datos.sort(function (a, b) {
      if (!a[columnaOrden]) a[columnaOrden] = "";
      if (!b[columnaOrden]) b[columnaOrden] = "";

      var bandA = a[columnaOrden].toString().toUpperCase();
      var bandB = b[columnaOrden].toString().toUpperCase();

      var comparison = 0;
      if (bandA > bandB) {
        comparison = 1;
      } else if (bandA < bandB) {
        comparison = -1;
      }
      return comparison;
    });

  }

  return datos;
}

/**
* Busca en la columna colBus del rango de datos la llave especificada y devuelve el valor en la columna colRes
* @param {String} llave Llave a buscar
* @param {Array} datos Datos ordenados en donde se debe buscar
* @param {Integer} colBus Numero de la columna en donde se busca la llave
* @param {Integer} colRes Numero de la columna en donde se encuentra el valor a devolver. -1 para devolver los datos de toda la fila
* @param {Boolean} numerico Indica si la columna por la llave de busqueda es numerica o no
* @return {Array} Array con la coincidencia encontrada
*/
function buscarPorLlaveBinario(llave, datos, colBus, colRes, numerico) {

  // var tipoLlave = typeof llave;
  // var llaveNumerica = tipoLlave == "number" ? true : false;

  // numerico = numerico || llaveNumerica;

  if (numerico) llave = Number(llave);
  else llave = llave.toString();


  var startIndex = 0;
  var endIndex = datos.length - 1;
  var compare;
  while (startIndex <= endIndex) {
    var middleIndex = Math.floor((startIndex + endIndex) / 2);

    if (numerico) compare = Number(datos[middleIndex][colBus - 1]);
    else compare = datos[middleIndex][colBus - 1].toString();

    if (llave == compare) {
      if (colRes == -1) return datos[middleIndex]
      else return datos[middleIndex][colRes - 1];
    } else if (llave > compare) {
      startIndex = middleIndex + 1;
    } else if (llave < compare) {
      endIndex = middleIndex - 1;
    } else {
      return null;
    }
  }

  return null;
}



function crearArchivoSheets(blob, carpeta) {
  var tipo = blob.getContentType();
  // console.log("tipo", tipo);
  // console.log("MICROSOFT_EXCEL", MimeType.MICROSOFT_EXCEL);
  // console.log("MICROSOFT_EXCEL_LEGACY", MimeType.MICROSOFT_EXCEL_LEGACY);

  // if (tipo != MimeType.MICROSOFT_EXCEL && tipo != MimeType.MICROSOFT_EXCEL_LEGACY) return null;

  var recurso = {
    "title": blob.getName(),
    "parents": [{ "id": carpeta.getId() }]
  };

  var archivo = Drive.Files.insert(recurso, blob, {
    "convert": true
  });

  return archivo;
}

/**
* Obtiene el numero de la columna especificada
* @param {String} columna Letra de la columna
* @return {Integer} Numero de la columna especificada
*/
function obtenerNumeroColumna(columna) {
  if (typeof columna == "string") {
    // return SpreadsheetApp.getActiveSheet().getRange(columna.trim() + "1").getColumn();
    columna = columna.toUpperCase();
    var numeroColumna = 0, longitud = columna.length;
    for (var i = 0; i < longitud; i++) {
      numeroColumna += (columna.charCodeAt(i) - 64) * Math.pow(26, longitud - i - 1);
    }
    return numeroColumna;
  }

  if (typeof columna == "object") {
    var numerosColumnas = [];
    for (var i = 0; i < columna.length; i++) {
      numerosColumnas.push(obtenerNumeroColumna(columna[i]));
    }
    return numerosColumnas;
  }

}


// }
/**
* Genera el documento con los datos enviados, las columnas y los tags especificados en la hoja de parametros
* @param {Array} registro Contiene los datos que se deben escribir en la plantilla
* @return {String} URL del PDF generado
*/
function combinarCampos(registro) {
  parametros = obtenerParametros()
  var idPlantilla;
  var opcion = registro[indicesColumnas.consolidado.tipoContrato - 1];
  var opcionTipoSalario = registro[indicesColumnas.consolidado.tipoSalario - 1];
  var compañia = registro[indicesColumnas.consolidado.compania - 1]

  console.log(opcion, opcionTipoSalario)

  var datosPlantillas = parametros.spsParametros.getSheetByName("MOD Formatos Contratos").getDataRange().getValues();

  if (opcion == "Término indefinido" && opcionTipoSalario != "Integral") {
    opcion = "Término indefinido - Normal"
  } else if (opcion == "Término indefinido" && opcionTipoSalario == "Integral") {
    opcion = "Término indefinido - Integral"
  }

  var filaContratroAsignado = datosPlantillas.filter(function (elemento) {
    if (elemento[0] == compañia && elemento[1] == opcion) return true
    else return false
  })

  if (filaContratroAsignado.length == 0) throw "No se encontró el tipo de contrato: " + opcion + " de la compañia: " + compañia

  var urlPlantilla = filaContratroAsignado[0][2]
  var arregloUrls = urlPlantilla.split(",");

  var cedula = registro[indicesColumnas.consolidado.numCedula - 1]
  var compania = registro[indicesColumnas.consolidado.compania - 1]
  var codigo_sucursal = registro[indicesColumnas.consolidado.codigoSucursal - 1]

  var carpetaRepositorio = DriveApp.getFolderById(CigoApp.obtenerIdUrl(obtenerParametro("carpetaRepositorioArchivos")));
  var urlCarpetaRegistro = registro[indicesColumnas.consolidado.linkCarpDigital - 1];

  var carpetaRegistro;
  if (!urlCarpetaRegistro) {
    carpetaRegistro = CigoApp.crearObtenerCarpeta(carpetaRepositorio, [codigo_sucursal, cedula])
  } else {
    carpetaRegistro = DriveApp.getFolderById(CigoApp.obtenerIdUrl(urlCarpetaRegistro));
  }

  var archivosReturn = [];

  // console.log(arregloUrls)

  for (var j = 0; j < arregloUrls.length; j++) {
    var idPlantillaValor = arregloUrls[j];
    console.log(idPlantillaValor)
    var idPlantilla = CigoApp.obtenerIdUrl(idPlantillaValor);
    console.log(idPlantilla)

    var archivoTemplatePlantilla = DriveApp.getFileById(idPlantilla);

    var nombreArchivo = archivoTemplatePlantilla.getName() + " - " + cedula;

    var archivoPlantilla = archivoTemplatePlantilla.makeCopy(nombreArchivo, carpetaRegistro)
    var datosCampos = parametros.spsParametros.getSheetByName("Relación Campos").getDataRange().getValues();

    var documentoPlantilla = DocumentApp.openById(archivoPlantilla.getId());
    var body = documentoPlantilla.getBody();

    var numeroColumna, tag, valorColumna;
    for (var i = 1; i < datosCampos.length; i++) {

      numeroColumna = CigoApp.obtenerNumeroColumna(datosCampos[i][0]);
      valorColumna = registro[numeroColumna - 1];
      valorColumna = valorColumna || ""

      if (valorColumna instanceof Date) {
        valorColumna = Utilities.formatDate(valorColumna, "America/Bogota", "dd/MM/yyyy");
      }
      console.log(tag, valorColumna)
      tag = datosCampos[i][1];
      body.replaceText(tag, valorColumna);

    }

    documentoPlantilla.saveAndClose();

    var blobPDF = DriveApp.getFileById(archivoPlantilla.getId()).getAs(MimeType.PDF);
    var archivoPDF = carpetaRegistro.createFile(blobPDF);
    archivoPDF.setName(nombreArchivo);

    archivosReturn.push(blobPDF);

    archivoPlantilla.setTrashed(true);
  }
  return archivosReturn;
}


function cambiarPermisosCarpetas() {
  if (!parametros) parametros = obtenerParametros()

  const sps_permisos = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ojO-uPn-dn-0Qv9zuPBbISq66G5jnGmzzXoKcZ4gWY0/edit?usp=sharing")
  const nombre_hoja = sps_permisos.getSheetByName("Permisos")
  let info_accesos = nombre_hoja.getDataRange().getValues()
  info_accesos.shift()


  const url_carpeta_repo = obtenerParametro("carpetaRepositorioArchivos")
  const carpeta = DriveApp.getFolderById(CigoApp.obtenerIdUrl(url_carpeta_repo))

  for (let info_permisos of info_accesos) {

    const carpeta_sucursal = CigoApp.crearObtenerCarpeta(carpeta, [info_permisos[0]])
    const correos = info_permisos[2].split(",")

    var permisos_edicion = carpeta_sucursal.getEditors();

    for (var i = 0; i < permisos_edicion.length; i++) {
      var user = permisos_edicion[i];
      if (correos.indexOf(user.getEmail()) === -1) {
        carpeta_sucursal.removeEditor(user);
      }
    }

    var permisos_vista = carpeta_sucursal.getViewers()

    for (var i = 0; i < permisos_vista.length; i++) {
      var user = permisos_vista[i];
      if (correos.indexOf(user.getEmail()) === -1) {
        carpeta_sucursal.removeViewer(user);
      }
    }

    for (const correo of correos) {
      try {
        if (permisos_edicion.includes(correo)) continue
        console.log(correo)
        carpeta_sucursal.addEditor(correo);
      } catch (error) {
        continue
      }
    }
  }
}

function registrarFormularioKit(e) {

  if (!e) return;
  if (e.range.columnStart == e.range.columnEnd) return;

  parametros = obtenerParametros();
  indicesColumnas = obtenerIndicesColumnas();

  var sps = SpreadsheetApp.openById(obtenerParametro("idHojaPlano"));
  var hojaConsolidado = sps.getSheetByName(obtenerParametro("nombreHojaPlano"));
  var cedulasConsolidado = hojaConsolidado.getRange(1, indicesColumnas.consolidado.numCedula, hojaConsolidado.getLastRow(), 1).getValues();

  var respuesta = e.values;
  var cedula = respuesta[1];

  var fila = CigoApp.obtenerFila(cedula, cedulasConsolidado, 1);
  if (!fila) {
    console.warn("No se encontró el registro con la cédula " + cedula);
    return;
  }

  if (respuesta[2] == "No" || respuesta[4] == "No") {
    var infoFila = hojaConsolidado.getRange(fila, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
    var correo = infoFila[indicesColumnas.consolidado.correo - 1];

    if (!correo) throw "El correo no está definido";

    // var asuntoRespuestaNOFormularioKit = CigoApp.reemplazarValores(obtenerParametro("asuntoRespuestaNOFormularioKit"), infoFila);
    // var mensajeRespuestaNOFormularioKit = CigoApp.reemplazarValores(obtenerParametro("mensajeRespuestaNOFormularioKit"), infoFila);

    // CigoApp.enviarCorreo(correo, asuntoRespuestaNOFormularioKit, mensajeRespuestaNOFormularioKit, {});
  } else {

    var infoFila = hojaConsolidado.getRange(fila, 1, 1, hojaConsolidado.getLastColumn()).getValues()[0];
    var correoLider = infoFila[indicesColumnas.consolidado.correoLider - 1];

    if (correoLider) {
      // var asuntoCorreosOK = CigoApp.reemplazarValores(obtenerParametro("asuntoCorreosOK"), infoFila);
      // var mensajeCorreoOK = CigoApp.reemplazarValores(obtenerParametro("mensajeCorreoOK"), infoFila);

      // CigoApp.enviarCorreo(correoLider, asuntoCorreosOK, mensajeCorreoOK, {});
    }

    hojaConsolidado.getRange(fila, indicesColumnas.consolidado.kitRecibido).setValue(respuesta[2])
    hojaConsolidado.getRange(fila, indicesColumnas.consolidado.fechaRecibidoKIT).setValue(respuesta[3])
    hojaConsolidado.getRange(fila, indicesColumnas.consolidado.accesosCorreoOK).setValue(respuesta[4])


    columnaGestion = indicesColumnas.consolidado.gestionAlistamiento;
    columnaFechaGestion = indicesColumnas.consolidado.horaGestionAlist;
    columnaCorreoGestion = indicesColumnas.consolidado.correoGestionAlist;

    hojaConsolidado.getRange(fila, columnaGestion).setValue("Gestionado");
    hojaConsolidado.getRange(fila, columnaFechaGestion).setValue(new Date());
    hojaConsolidado.getRange(fila, columnaCorreoGestion).setValue(Session.getActiveUser().getEmail());
  }
}

function correccionFecha() {
  parametros = obtenerParametros()
  indicesColumnas = obtenerIndicesColumnas()

  const hoja_plano = SpreadsheetApp.getActive().getSheetByName("Definitiva")
  const datos_plano = hoja_plano.getDataRange().getValues()

  const hoja_taleo = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Anu4nmNyupIx6cgHKR-u637VkWLoxnad3apE46k-Gkw/edit#gid=910746765").getSheetByName("Sheet1")

  const datos_taleo = hoja_taleo.getDataRange().getDisplayValues()

  const indices_fila_cedula = obteneFilasPorCedula(datos_plano)
  const columna_fecha_expedicion_cedula_plano = obtenerColumnaExpedicionCedula(datos_plano)

  console.log(columna_fecha_expedicion_cedula_plano.length)

  const fila_expedicion_cedula_taleo = 44
  const fila_cedula_taleo = 36

  const columna_expedicion_cedula_formateada = datos_taleo.reduce((columna_expedicion_cedula, fila_taleo) => {
    const cedula = fila_taleo[fila_cedula_taleo].trim()

    if (!indices_fila_cedula[cedula]) return columna_expedicion_cedula

    let expedicion_cedula_formateada = fila_taleo[fila_expedicion_cedula_taleo]

    if (expedicion_cedula_formateada != "") {
      expedicion_cedula_formateada = expedicion_cedula_formateada.split("-")
      expedicion_cedula_formateada = `${expedicion_cedula_formateada[1]}/${expedicion_cedula_formateada[0]}/${expedicion_cedula_formateada[2]}`
    } else {
      expedicion_cedula_formateada = ""
    }

    // console.log(expedicion_cedula_formateada)
    for (let fila of indices_fila_cedula[cedula]) {
      columna_fecha_expedicion_cedula_plano[fila] = [expedicion_cedula_formateada]
    }
    return columna_fecha_expedicion_cedula_plano

  }, columna_fecha_expedicion_cedula_plano)

  hoja_plano.getRange(1, 59, columna_expedicion_cedula_formateada.length, 1).setValues(columna_expedicion_cedula_formateada)
  // console.log(columna_expedicion_cedula_formateada)
}

function obtenerColumnaExpedicionCedula(datosPlano) {

  const columna_expedicion_cedula = datosPlano.reduce((acumulado_columna_cedula, fila_actual, index) => {
    const fecha_expedicion_cedula = fila_actual[indicesColumnas.consolidado.fechaExpCedula - 1]

    acumulado_columna_cedula[index] = [fecha_expedicion_cedula]

    return acumulado_columna_cedula
  }, [])

  return columna_expedicion_cedula
}

function obteneFilasPorCedula(datosPlano) {

  const indices_fila_cedula = datosPlano.reduce((acumulado_filas_cedula, fila_actual, index) => {
    const cedula = fila_actual[indicesColumnas.consolidado.numCedula - 1].toString().trim()

    if (!acumulado_filas_cedula[cedula]) {
      acumulado_filas_cedula[cedula] = []
    }

    acumulado_filas_cedula[cedula].push(index)

    return acumulado_filas_cedula
  }, {})
  return indices_fila_cedula
}/**
* Convierte el valor ingresado en letras
*
* @param {A3} valor Valor a convertir
* @param {"PESOS"} textoFinal Texto agregar al final del resultado
* @return Numero en letras
* @customfunction
*
*/
function NUMEROATEXTO(valor, textoFinal) {
  textoFinal = textoFinal || "";
  textoFinal = textoFinal.toString().toUpperCase();

  if (isNaN(valor)) throw "Ingresa un valor numérico";

  return numeroALetras(Number(valor), textoFinal);
}

/*************************************************************/
// NumeroALetras
// The MIT License (MIT)
// 
// Copyright (c) 2015 Luis Alfredo Chee 
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
// 
// @author Rodolfo Carmona
// @contributor Jean (jpbadoino@gmail.com)
/*************************************************************/
function Unidades(num){

    switch(num)
    {
        case 1: return "UN";
        case 2: return "DOS";
        case 3: return "TRES";
        case 4: return "CUATRO";
        case 5: return "CINCO";
        case 6: return "SEIS";
        case 7: return "SIETE";
        case 8: return "OCHO";
        case 9: return "NUEVE";
    }

    return "";
}//Unidades()

function Decenas(num){

    decena = Math.floor(num/10);
    unidad = num - (decena * 10);

    switch(decena)
    {
        case 1:
            switch(unidad)
            {
                case 0: return "DIEZ";
                case 1: return "ONCE";
                case 2: return "DOCE";
                case 3: return "TRECE";
                case 4: return "CATORCE";
                case 5: return "QUINCE";
                default: return "DIECI" + Unidades(unidad);
            }
        case 2:
            switch(unidad)
            {
                case 0: return "VEINTE";
                default: return "VEINTI" + Unidades(unidad);
            }
        case 3: return DecenasY("TREINTA", unidad);
        case 4: return DecenasY("CUARENTA", unidad);
        case 5: return DecenasY("CINCUENTA", unidad);
        case 6: return DecenasY("SESENTA", unidad);
        case 7: return DecenasY("SETENTA", unidad);
        case 8: return DecenasY("OCHENTA", unidad);
        case 9: return DecenasY("NOVENTA", unidad);
        case 0: return Unidades(unidad);
    }
}//Unidades()

function DecenasY(strSin, numUnidades) {
    if (numUnidades > 0)
    return strSin + " Y " + Unidades(numUnidades)

    return strSin;
}//DecenasY()

function Centenas(num) {
    centenas = Math.floor(num / 100);
    decenas = num - (centenas * 100);

    switch(centenas)
    {
        case 1:
            if (decenas > 0)
                return "CIENTO " + Decenas(decenas);
            return "CIEN";
        case 2: return "DOSCIENTOS " + Decenas(decenas);
        case 3: return "TRESCIENTOS " + Decenas(decenas);
        case 4: return "CUATROCIENTOS " + Decenas(decenas);
        case 5: return "QUINIENTOS " + Decenas(decenas);
        case 6: return "SEISCIENTOS " + Decenas(decenas);
        case 7: return "SETECIENTOS " + Decenas(decenas);
        case 8: return "OCHOCIENTOS " + Decenas(decenas);
        case 9: return "NOVECIENTOS " + Decenas(decenas);
    }

    return Decenas(decenas);
}//Centenas()

function Seccion(num, divisor, strSingular, strPlural) {
    cientos = Math.floor(num / divisor)
    resto = num - (cientos * divisor)

    letras = "";

    if (cientos > 0)
        if (cientos > 1)
            letras = Centenas(cientos) + " " + strPlural;
        else
            letras = strSingular;

    if (resto > 0)
        letras += "";

    return letras;
}//Seccion()

function Miles(num) {
    divisor = 1000;
    cientos = Math.floor(num / divisor)
    resto = num - (cientos * divisor)

    strMiles = Seccion(num, divisor, "UN MIL", "MIL");
    strCentenas = Centenas(resto);

    if(strMiles == "")
        return strCentenas;

    return strMiles + " " + strCentenas;
}//Miles()

function Millones(num) {
    divisor = 1000000;
    cientos = Math.floor(num / divisor)
    resto = num - (cientos * divisor)

    strMillones = Seccion(num, divisor, "UN MILLON", "MILLONES");
    strMiles = Miles(resto);

    if(strMillones == "")
        return strMiles;

    return strMillones + " " + strMiles;
}//Millones()


function numeroALetras(num, textoFinal) {
    var data = {
        numero: num,
        enteros: Math.floor(num),
        centavos: (((Math.round(num * 100)) - (Math.floor(num) * 100))),
        letrasCentavos: "",
        letrasMonedaPlural: textoFinal,//"PESOS", 'Dólares', 'Bolívares', 'etcs'
        letrasMonedaSingular: textoFinal, //"PESO", 'Dólar', 'Bolivar', 'etc'

        letrasMonedaCentavoPlural: "",
        letrasMonedaCentavoSingular: ""
    };

    if (data.centavos > 0) {
        data.letrasCentavos = "PUNTO " + (function (){
            if (data.centavos == 1)
                return Millones(data.centavos);
            else
                return Millones(data.centavos);
            })();
    };

    if(data.enteros == 0)
        return "CERO " + data.letrasMonedaPlural + " " + data.letrasCentavos;
    if (data.enteros == 1)
        return Millones(data.enteros) + " " + data.letrasCentavos + " " + data.letrasMonedaSingular;
    else
        return Millones(data.enteros) + " " + data.letrasCentavos + " " + data.letrasMonedaPlural;
}

function numeroALetras_OLD(num) {
    var data = {
        numero: num,
        enteros: Math.floor(num),
        centavos: (((Math.round(num * 100)) - (Math.floor(num) * 100))),
        letrasCentavos: "",
        letrasMonedaPlural: 'PESOS',//"PESOS", 'Dólares', 'Bolívares', 'etcs'
        letrasMonedaSingular: 'PESO', //"PESO", 'Dólar', 'Bolivar', 'etc'

        letrasMonedaCentavoPlural: "CENTAVOS",
        letrasMonedaCentavoSingular: "CENTAVO"
    };

    if (data.centavos > 0) {
        data.letrasCentavos = "CON " + (function (){
            if (data.centavos == 1)
                return Millones(data.centavos) + " " + data.letrasMonedaCentavoSingular;
            else
                return Millones(data.centavos) + " " + data.letrasMonedaCentavoPlural;
            })();
    };

    if(data.enteros == 0)
        return "CERO " + data.letrasMonedaPlural + " " + data.letrasCentavos;
    if (data.enteros == 1)
        return Millones(data.enteros) + " " + data.letrasMonedaSingular + " " + data.letrasCentavos;
    else
        return Millones(data.enteros) + " " + data.letrasMonedaPlural + " " + data.letrasCentavos;
}


