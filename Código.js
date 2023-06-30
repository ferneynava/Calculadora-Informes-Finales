/* eslint-disable no-var */
/* eslint-disable space-before-function-paren */
/* eslint-disable prefer-const */
/* eslint-disable no-undef */
/* eslint-disable no-unused-vars */
var objectAsignatura = {
  ART: 'ArtesArtística',
  ART_PK: 'ArtesArtística',
  ART_K: 'ArtesArtística',
  ART_T: 'ArtesArtística',
  BOI: 'Ciencias Naturales y Ed ambientalBiología',
  CN: 'Ciencias Naturales y Ed ambientalCiencias Naturales',
  CN_PK: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
  CN_K: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
  ETIC: 'ÉticaÉtica',
  E: 'Educación FísicaEd. Física',
  ERE: 'EREERE',
  E_PK: 'Educación FísicaEd. Física',
  E_K: 'Educación FísicaEd. Física',
  E_T: 'Educación FísicaEd. Física',
  ETIC_PK: 'DimensionesDimensión Ética',
  ETIC_K: 'DimensionesDimensión Ética',
  ETIC_T: 'DimensionesDimensión Ética',
  ERE_PK: 'DimensionesDimensión Espiritual ERE',
  ERE_K: 'DimensionesDimensión Espiritual ERE',
  ERE_T: 'DimensionesDimensión Espiritual ERE',
  FIS: 'Ciencias Naturales y Ed ambientalFísica',
  FILO: 'FilosofíaFilosofía',
  INV: 'Ciencias SocialesInvestigación',
  ING: 'InglésInglés',
  ING_PK: 'DimensionesDimensión Comunicativa Inglés',
  ING_K: 'DimensionesDimensión Comunicativa Inglés',
  ING_T: 'DimensionesDimensión Comunicativa Inglés',
  ICFES: 'Preparación ICFESPreicfes',
  LC: 'Lengua CastellanaLengua Castellana',
  LC_PK: 'DimensionesDimensión Comunicativa Lengua Castellana',
  LC_K: 'DimensionesDimensión Comunicativa Lengua Castellana',
  LC_T: 'DimensionesDimensión Comunicativa Lengua Castellana',
  MAT: 'MatemáticasMatemáticas',
  MATH: 'InglésMath',
  MUS: 'ArtesMúsica',
  MUS_PK: 'ArtesMúsica',
  MUS_K: 'ArtesMúsica',
  MUS_T: 'ArtesMúsica',
  MAT_PK: 'DimensionesDimensión congnitiva Matemáticas',
  MAT_K: 'DimensionesDimensión congnitiva Matemáticas',
  MAT_T: 'DimensionesDimensión congnitiva Matemáticas',
  MATH_PK: 'InglésMath',
  MATH_K: 'InglésMath',
  MATH_T: 'InglésMath',
  PL: 'InglésPlan lector',
  QUI: 'Ciencias Naturales y Ed ambientalQuímica',
  SOC: 'Ciencias SocialesC. Sociales',
  SCI: 'InglésScience',
  SOC_PK: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
  SOC_K: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
  SOC_T: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
  SCI_PK: 'InglésScience',
  SCI_K: 'InglésScience',
  SCI_T: 'InglésScience',
  TEC: 'TecnologíaTecnología',
  TEC_PK: 'DimensionesDimensión cognitiva Tecnología',
  TEC_K: 'DimensionesDimensión cognitiva Tecnología',
  TEC_T: 'DimensionesDimensión cognitiva Tecnología'
}

var objectGrado = {
  PK: 'PJº',
  K: 'Jº',
  T: 'TRº',
  '01': '1º',
  '02': '2º',
  '03': '3º',
  '04': '4º',
  '05': '5º',
  '06': '6º',
  '07': '7º',
  '08': '8º',
  '09': '9º',

  10: '10º',

  11: '11º'
}

var objectPeriodos = {
  '01': '1P',
  '02': '2P',
  '03': '3P',
  '04': '4P',
  Final: '5P'
}

var DatabaseDoc = function () {
  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerData = SpreadsheetApp.openById(`${idBaseData}`).getActiveSheet()
  let data = agregarObtenerData.getDataRange().getValues()
  data.shift()
  let Docid = data.map((arrays) => {
    arrays.splice(1, 1)
    return arrays
  })

  return Object.fromEntries(Docid)
}

var DatabaseEva = function () {
  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerData = SpreadsheetApp.openById(`${idBaseData}`).getSheetByName('Evaluaciones')
  let data = agregarObtenerData.getRange(2, 1, 17, 2).getValues()
  let dataEva = data.filter(dat => dat.some(element => element !== ''))
  return Object.fromEntries(dataEva)
}

var objectGradoID = {
  '00_K': '1V9DxP4MrvWV24c2v0_bAxUkYrsdXLFMp3EMe81O8Kvs',
  '00_PK': '1r_FOk72TfYaikMKWF0zSkQ7qP7_CnGlEiz3Rit4p7XQ',
  '00_T': '1ieXf0KSn4ybzFF2fjebpvrvVsWdbl9AB_kh5fh_cxC4',
  '01': '1Rc6KLWI7HbMRTbBqW2n4wpfNDJnvS0Ax_kXvmrivOBs',
  '02': '1o__2-AOkTrvcrwkYL_QTFQTofhHj5M5jij2REzZRoJU',
  '03': '1IDnC0cTgTMjNB57gTb2IrxDNjp3ixYQGEyS4rXxXdRE',
  '04': '12GrBIsDXLrWbDMDMtW6s2e4MYS9bJHJYB5nY-pgPRKs',
  '05': '1eXiA69pEr8yZdHM3bTxfZ2f9KTzNkNn55NI_ThdPbmY',
  '06': '1pZ-cz_Bmq_9784pFXsOtSVo-yMYgFQmqOjk9DwlsmSs',
  '07': '1PphndEzhN0wod3BxzNfPdS8jlpwzJtbZvFamXX2aZEI',
  '08': '1SXa5D3HMhdg4nIIa_cpsy1IwYX3_MWXvIcKFZ2VsBCc',
  '09': '1hmiODqziuLO7Q2_sdpZph7x-yQHASo6X2lWtQth9SFg',
  // eslint-disable-next-line quote-props
  '10': '1y5sd7uTfZDLF1J2Z75r9eLGM27TYoS8NFjqi_k0QODQ',
  // eslint-disable-next-line quote-props
  '11': '1YRpY7FLwnUYW8_1eEanLaBmv8BiPlWhc5_8H3-cfHpE'
}

var objectPeriodoCur = {
  '01': 5,
  '02': 35,
  '03': 65,
  '04': 95,
  Final: 125
}

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Envio de datos MauxiChia')
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent()
}

function docentes(object) {
  let objectDocente = DatabaseDoc()
  const periodoIngre = object.periodo
  const periodo = objectPeriodos[periodoIngre] ?? 'N/A'

  const asignatura = object.asignatura
  const arrayAsignatutas = asignatura.split('_')
  let materia = arrayAsignatutas[0]
  let grado = arrayAsignatutas[1]
  if (grado === 'FIS') {
    grado = arrayAsignatutas[2]
  }
  const gradoPlaCompleta = objectGrado[grado] ?? 'N/A'

  if (grado.includes('PK') || grado.includes('K') || grado.includes('T')) {
    materia = materia + '_' + grado
  }

  const id = object.idnotas
  const hoja = SpreadsheetApp.openById(`${id}`).getSheetByName(gradoPlaCompleta)
  const dataCompleta = hoja.getDataRange().getValues()

  const docente = object.idCalculadora
  const idDocente = objectDocente[docente] ?? 'N/A'

  const calculadora = SpreadsheetApp.openById(`${idDocente}`).getSheetByName(`Datos ${periodo}`)
  let dataCalculadora = calculadora.getDataRange().getValues()
  let arrayAsignatura = dataCalculadora[2]

  const numeroEstuden = dataCompleta.map(filaNom => filaNom[0])
  numeroEstuden.shift()
  const arrayNombres = numeroEstuden.filter(eliminar => eliminar !== '')

  let celdaColumnaAsig = arrayAsignatura.indexOf(asignatura)
  const asignaturaPlaCompleta = objectAsignatura[materia] ?? 'N/A'
  const arrayAsignaturasPlaCompleta = dataCompleta[0]
  let celdaColumna = arrayAsignaturasPlaCompleta.indexOf(asignaturaPlaCompleta)

  let n1 = 2
  let c1 = 9

  if (periodo === '1P') {
    for (let i = 0; i < arrayNombres.length; i++) {
      let notas = hoja.getRange(n1, celdaColumna + 1).getValues()
      n1 += 5
      c1 += 1
      calculadora.getRange(c1, celdaColumnaAsig + 1).setValue(notas)
    }
  }

  let n2 = 3
  let c2 = 9

  if (periodo === '2P') {
    for (let i = 0; i < arrayNombres.length; i++) {
      let notas = hoja.getRange(n2, celdaColumna + 1).getValues()
      n2 += 5
      c2 += 1
      calculadora.getRange(c2, celdaColumnaAsig + 1).setValue(notas)
    }
  }

  let n3 = 4
  let c3 = 9

  if (periodo === '3P') {
    for (let i = 0; i < arrayNombres.length; i++) {
      let notas = hoja.getRange(n3, celdaColumna + 1).getValues()
      n3 += 5
      c3 += 1
      calculadora.getRange(c3, celdaColumnaAsig + 1).setValue(notas)
    }
  }

  let n4 = 5
  let c4 = 9

  if (periodo === '4P') {
    for (let i = 0; i < arrayNombres.length; i++) {
      let notas = hoja.getRange(n4, celdaColumna + 1).getValues()
      n4 += 5
      c4 += 1
      calculadora.getRange(c4, celdaColumnaAsig + 1).setValue(notas)
    }
  }

  let n5 = 6
  let c5 = 9

  if (periodo === '5P') {
    for (let i = 0; i < arrayNombres.length; i++) {
      let notas = hoja.getRange(n5, celdaColumna + 1).getValues()
      n5 += 5
      c5 += 1
      calculadora.getRange(c5, celdaColumnaAsig + 1).setValue(notas)
    }
  }
}

function hojasData(object) {
  const id = object.inputId
  const datos = SpreadsheetApp.openById(`${id}`)
  let hojas = datos.getSheets()
  let totalArray = hojas.length

  let array = []
  for (let i = 0; i < totalArray; i++) {
    let hojasReco = hojas[i].getName()
    array.push(hojasReco)
  }
  return array
}

function evaluaciones(object) {
  let objectEvaMaterias = DatabaseEva()
  let idPlanillaNotas = object.idnotasEva
  let materiaGrado = object.asignaturaEva
  let evaMateria = object.EvMateria
  let evaGrado = object.EvGrados
  let columna = object.columa
  let numeroEs = object.numeroEst
  let mayusColumna = columna.toUpperCase()
  let fila = object.fila

  const arrayevaMateria = evaMateria.split('_')
  let Eva = arrayevaMateria[0]
  let Mate = arrayevaMateria[1]

  let materiaPlaEva = Mate + '_' + evaGrado

  const planillaNotas = SpreadsheetApp.openById(idPlanillaNotas).getSheetByName(materiaGrado)

  let hojaEva = objectEvaMaterias[evaMateria]

  const periodoEvaIngre = object.periodoEva
  const periodoEva = objectPeriodos[periodoEvaIngre] ?? 'N/A'
  const calculadoraEva = SpreadsheetApp.openById(hojaEva).getSheetByName(`Datos ${periodoEva}`)
  let dataCalculadoraEva = calculadoraEva.getDataRange().getValues()
  let arrayAsignaturaEva = dataCalculadoraEva[2]
  let celdaColumnaAsig = arrayAsignaturaEva.indexOf(materiaPlaEva)

  let filaNum = Number(fila)
  let c5 = 9

  for (let i = 0; i < numeroEs; i++) {
    let filaString = filaNum.toString()
    const dataCompletaNotas = planillaNotas.getRange(mayusColumna + filaString).getValues()
    filaNum += 1
    c5 += 1
    calculadoraEva.getRange(c5, celdaColumnaAsig + 1).setValue(dataCompletaNotas)
  }
}

function curso(object) {
  let objectDocente = DatabaseDoc()
  let docenteCur = object.idCalculadoraDoc
  let asignaturaGraCur = object.AsignaturaGrad
  let PeriodoCur = object.periodoGrad
  let GradoCur = object.grado
  let MateriaGradoCur = object.materiaGrado

  const idDocenteCur = objectDocente[docenteCur] ?? 'N/A'
  let idPeriodoCur = objectPeriodoCur[PeriodoCur] ?? 'N/A'
  const idGrado = objectGradoID[GradoCur] ?? 'N/A'

  const calculadoraDoc = SpreadsheetApp.openById(`${idDocenteCur}`).getSheetByName('Análisis')
  let dataCalculadoraDoc = calculadoraDoc.getDataRange().getValues()
  let arrayasignaturaGra = dataCalculadoraDoc[3]
  let celdaColumnaAsig = arrayasignaturaGra.indexOf(asignaturaGraCur)

  let dataColum = calculadoraDoc.getRange(idPeriodoCur, celdaColumnaAsig + 1, 7).getValues()

  const planillasDoc = SpreadsheetApp.openById(idGrado).getSheetByName('Análisis')
  let dataPlanillaDoc = planillasDoc.getDataRange().getValues()

  let arrayAsignaturaDoc = dataPlanillaDoc[3]
  const arrayAsignatutasDoc = MateriaGradoCur.split('_')
  let materiaCur = arrayAsignatutasDoc[0]
  let fis = arrayAsignatutasDoc[1]
  if (fis === 'FIS') {
    materiaCur = materiaCur + '_' + fis
  }

  let celdaColumnaAsigDoc = arrayAsignaturaDoc.indexOf(materiaCur)

  for (let i = 0; i < 7; i++) {
    planillasDoc.getRange(idPeriodoCur, celdaColumnaAsigDoc + 1).setValue(dataColum[i])
    idPeriodoCur += 1
  }
}

function directivas (object) {
  let objectDocente = DatabaseDoc()
  var docenteCambio = object.cambioDoc
  const idDocenteCambi = objectDocente[docenteCambio]
  const nuevaDoc = object.nombreDocNuv

  SpreadsheetApp.openById(`${idDocenteCambi}`).setName(nuevaDoc)

  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerData = SpreadsheetApp.openById(`${idBaseData}`).getActiveSheet()
  let data = agregarObtenerData.getRange(2, 1, 30, 1).getValues()

  let nombreApellidos = SpreadsheetApp.openById(`${idDocenteCambi}`).getActiveSheet()
  nombreApellidos.getRange(1, 7).setValue(object.nombreApellido)

  for (let i = 0; i < data.length; i++) {
    if (data[i].includes(docenteCambio)) {
      let fila = i
      agregarObtenerData.getRange(fila + 2, 1).setValue(nuevaDoc)
    }
  }
}

function directivaEliminar (object) {
  let objectDocente = DatabaseDoc()
  let docenteEliminar = object.eliminarDoc
  const idDocenteEli = objectDocente[docenteEliminar]

  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerData = SpreadsheetApp.openById(`${idBaseData}`).getActiveSheet()
  let data = agregarObtenerData.getRange(2, 1, 30, 1).getValues()
  let archivo = DriveApp.getFileById(idDocenteEli)
  archivo.setTrashed(true)
  for (let i = 0; i < data.length; i++) {
    if (data[i].includes(docenteEliminar)) {
      let position = i
      agregarObtenerData.deleteRow(position + 2)
    }
  }
}

function agregar (object) {
  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerData = SpreadsheetApp.openById(`${idBaseData}`).getActiveSheet()
  let archivo = DriveApp.getFileById(`${object.InputIdArch}`)
  archivo.setName(object.InputDocNom)
  let nombreDocente = SpreadsheetApp.openById(`${object.InputIdArch}`).getActiveSheet()
  nombreDocente.getRange(1, 7).setValue(object.InputNombA)
  agregarObtenerData.appendRow([`${object.InputDocNom}`, `${object.InputNombA}`, `${object.InputIdArch}`])
}

function agregarEva (object) {
  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerData = SpreadsheetApp.openById(`${idBaseData}`).getSheetByName('Evaluaciones')
  let archivo = DriveApp.getFileById(`${object.inputArchivoCrea}`)
  archivo.setName(object.inputEvaArea)
  let nombreÁrea = SpreadsheetApp.openById(`${object.inputArchivoCrea}`).getActiveSheet()
  nombreÁrea.getRange(1, 7).setValue(object.inputNombreArea)
  agregarObtenerData.appendRow([`${object.inputEvaArea}`, `${object.inputArchivoCrea}`])
}

function directivaEliminarEva (object) {
  let objectEvaluacion = DatabaseEva()
  let evaluacionEliminar = object.eliminarEva
  const idDocenteEli = objectEvaluacion[evaluacionEliminar]

  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerData = SpreadsheetApp.openById(`${idBaseData}`).getSheetByName('Evaluaciones')
  let data = agregarObtenerData.getRange(2, 1, 17, 1).getValues()
  let archivo = DriveApp.getFileById(idDocenteEli)
  archivo.setTrashed(true)
  for (let i = 0; i < data.length; i++) {
    if (data[i].includes(evaluacionEliminar)) {
      let position = i
      agregarObtenerData.deleteRow(position + 2)
    }
  }
}

function leerBasedeDatos () {
  const idBaseData = '1XqiRk7cVqoYiWNFdZ6ePYrjdsmFZugDzGGtpw5ouR9U'
  let agregarObtenerDataDoc = SpreadsheetApp.openById(`${idBaseData}`).getSheetByName('Docente')
  let agregarObtenerDataEva = SpreadsheetApp.openById(`${idBaseData}`).getSheetByName('Evaluaciones')
  let dataBaseDocnom = agregarObtenerDataDoc.getRange(2, 1, 30, 1).getValues()
  let dataBaseEvaMat = agregarObtenerDataEva.getRange(2, 1, 17, 1).getValues()
  let objec = {
    dataBaseDoc: dataBaseDocnom,
    dataBaseEva: dataBaseEvaMat
  }
  return objec
}
