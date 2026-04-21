//© Banco Davivienda S.A. 2021. Se prohíbe su uso o reproducción sin previa autorización del Banco Davivienda S.A

const abreviacionFiliales = {
  "Davivienda": "DAVIVIENDA",
  "DAVIVIENDA - HCM": "DAVIVIENDA",
  "BETA - HCM": "BETA",
  "Beta": "BETA",
  "GAMMA - HCM": "GAMMA",
  "Gamma": "GAMMA",
  "Fiduciaria Davivienda": "FIDUCIARIA",
  "FIDUCIARIA - HCM": "FIDUCIARIA",
  "CORREDORES DAVIVIENDA": "CORREDORES",
  "CORREDORES - HCM": "CORREDORES",
  "Corporación Financiera": "FINANCIERA",
  "CORPORACION FINANCIERA - HCM": "FINANCIERA"
}

function testCorreoImagenes() {
  parametros = obtenerParametros();
  indicesColumnas = obtenerIndicesColumnas();

  var filial = 'DAVIVIENDA'
  var infoFila= []


  var asuntoCitacionInduccion = obtenerComunicacionPorFilial("asuntoCitacionInduccionContCargadoNo", filial);
  asuntoCitacionInduccion = CigoApp.reemplazarValores(asuntoCitacionInduccion, infoFila);

  var mensajeCitacionInduccion = obtenerComunicacionPorFilial("mensajeCitacionInduccionContCargadoNo", filial);
  mensajeCitacionInduccion = CigoApp.reemplazarValores(mensajeCitacionInduccion, infoFila);

  var inlineImages = obtenerImagenesCorreo("mensajeCitacionInduccionContCargadoNo", filial);
  console.log(mensajeCitacionInduccion)

  enviarCorreo("santiago.aguilarsoto@davivienda.com", asuntoCitacionInduccion, mensajeCitacionInduccion, { noReply: true, inlineImages: inlineImages });

}

let infoPorComunicaciones
function obtenerComunicacionPorFilial(comunicacion, filial) {

  filial = filial.toString().toUpperCase().trim()

  if (!infoPorComunicaciones) {
    infoPorComunicaciones = obtenerDatosPorComunicaciones()
  }

  filial = !abreviacionFiliales[filial] ? filial : abreviacionFiliales[filial]

  const filiales = infoPorComunicaciones.Filial
  const comunicacionPorFilial = infoPorComunicaciones[comunicacion][filiales.indexOf(filial)]

  if (!comunicacionPorFilial) throw `Ha ocurrido un problema, no existe la parametrizacion para la filial ${filial}`

  return comunicacionPorFilial
}

function obtenerDatosPorComunicaciones() {
  parametros = obtenerParametros()
  indicesColumnas = obtenerIndicesColumnas()

  datosComunicaciones = parametros.spsParametros.getSheetByName("Comunicaciones").getDataRange().getValues()

  datosComunicacionPorLlave = datosComunicaciones.reduce((llavesPorComunicacion, detalleComunicacion) => {

    const llaveComunicacion = detalleComunicacion[0]

    llavesPorComunicacion[llaveComunicacion] = detalleComunicacion
    return llavesPorComunicacion

  }, {})

  return datosComunicacionPorLlave
}

function obtenerImagenesCorreo(mensaje, filial) {

  try {
    var imagenes = obtenerComunicacionPorFilial("imagenes-" + mensaje, filial);
  } catch (e) {
    console.warn(e);
    return {};
  }

  imagenes = imagenes.split(",");

  var inlineImages = {};
  for (var i = 0; i < imagenes.length; i++) {
    console.log(imagenes[i])
    var blobImagen = DriveApp.getFileById(CigoApp.obtenerIdUrl(imagenes[i].trim())).getBlob();
    inlineImages["img" + (i + 1)] = blobImagen;
  }

  console.log(inlineImages)

  return inlineImages;
}


/**
* Envia un correo electrónico con los parámetros dados
* @param {String} correo Correo del destinatario
* @param {String} asunto Asunto del mensaje
* @param {String} mensaje Mensaje a enviar
* @param {Object} opciones Opciones en formato JSON. Las opciones son las mismas disponibles en https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
*/
function enviarCorreo(correo, asunto, mensaje, opciones) {
  var template = HtmlService.createTemplateFromFile("CorreoHTML");
  template.mensaje = mensaje;
  var body = template.evaluate().getContent();

  if (typeof opciones === "string") {
    opciones = {};
  }

  opciones.htmlBody = body;
  GmailApp.sendEmail(correo, asunto, "", opciones);

  // try {
  //   var blobImage = DriveApp.getFileById("1Kye2eX5LFWGlYvZVC7cbJ9Jvm6vrtvPk").getBlob();
  //   if (!opciones.inlineImages) opciones.inlineImages = {logoDav: blobImage};
  //   else {
  //     if (!opciones.inlineImages.logoDav) opciones.inlineImages.logoDav = blobImage;
  //   }

  // } catch (e) {
  //   console.warn("No se pudo cargar la imagen en el correo. \n" + JSON.stringify(e));
  // } finally {
  //   opciones.htmlBody = body;
  //   GmailApp.sendEmail(correo, asunto,"", opciones);
  // }

}
