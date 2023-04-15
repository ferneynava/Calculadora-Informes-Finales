/* eslint-disable space-before-function-paren */
/* eslint-disable prefer-const */
/* eslint-disable no-undef */
/* eslint-disable no-unused-vars */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Envio de datos MauxiChia')
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent()
}

function docentes(numero) {
  const objectAsignatura = {
    MAT: 'MatemáticasMatemáticas',
    TEC: 'TecnologíaTecnología',
    LC: 'Lengua CastellanaLengua Castellana',
    ETIC: 'ÉticaÉtica',
    ART: 'ArtesArtística',
    SOC: 'Ciencias SocialesC. Sociales',
    INV: 'Ciencias SocialesInvestigación',
    E: 'Educación FísicaEd. Física',
    SCI: 'InglésScience',
    MATH: 'InglésMath',
    PL: 'InglésPlan lector',
    CN: 'Ciencias Naturales y Ed ambientalCiencias Naturales',
    BOI: 'Ciencias Naturales y Ed ambientalBiología',
    ING: 'InglésInglés',
    FIS: 'Ciencias Naturales y Ed ambientalFísica',
    QUI: 'Ciencias Naturales y Ed ambientalQuímica',
    ICFES: 'Preparación ICFESPreicfes',
    MUS: 'ArtesMúsica',
    ERE: 'EREERE',
    FILO: 'FilosofíaFilosofía',
    TEC_PK: 'DimensionesDimensión cognitiva Tecnología',
    TEC_K: 'DimensionesDimensión cognitiva Tecnología',
    TEC_T: 'DimensionesDimensión cognitiva Tecnología',
    ART_PK: 'ArtesArtística',
    ART_K: 'ArtesArtística',
    ART_T: 'ArtesArtística',
    E_PK: 'Educación FísicaEd. Física',
    E_K: 'Educación FísicaEd. Física',
    E_T: 'Educación FísicaEd. Física',
    CN_PK: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
    CN_K: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
    ETIC_PK: 'DimensionesDimensión Ética',
    ETIC_K: 'DimensionesDimensión Ética',
    ETIC_T: 'DimensionesDimensión Ética',
    MUS_PK: 'ArtesMúsica',
    MUS_K: 'ArtesMúsica',
    MUS_T: 'ArtesMúsica',
    LC_PK: 'DimensionesDimensión Comunicativa Lengua Castellana',
    LC_K: 'DimensionesDimensión Comunicativa Lengua Castellana',
    LC_T: 'DimensionesDimensión Comunicativa Lengua Castellana',
    MAT_PK: 'DimensionesDimensión congnitiva Matemáticas',
    MAT_K: 'DimensionesDimensión congnitiva Matemáticas',
    MAT_T: 'DimensionesDimensión congnitiva Matemáticas',
    SOC_PK: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
    SOC_K: 'DimensionesDimensión Socioafectiva Ciencias Sociales y Naturales',
    ING_PK: 'DimensionesDimensión Comunicativa Inglés',
    ING_K: 'DimensionesDimensión Comunicativa Inglés',
    ING_T: 'DimensionesDimensión Comunicativa Inglés',
    SCI_PK: 'InglésScience',
    SCI_K: 'InglésScience',
    SCI_T: 'InglésScience',
    MATH_PK: 'InglésMath',
    MATH_K: 'InglésMath',
    MATH_T: 'InglésMath',
    ERE_PK: 'DimensionesDimensión Espiritual ERE',
    ERE_K: 'DimensionesDimensión Espiritual ERE',
    ERE_T: 'DimensionesDimensión Espiritual ERE'
  }

  const objectGrado = {
    PK: 'PJº',
    K: 'Jº',
    T: 'Tº',
    '01': '1º',
    '02': '2º',
    '03': '3º',
    '04': '4º',
    '05': '5º',
    '06': '6º',
    '07': '7º',
    '08': '8º',
    '09': '9º',
    // eslint-disable-next-line quote-props
    '10': '10º',
    // eslint-disable-next-line quote-props
    '11': '11º'
  }

  const objectPeriodos = {
    '01': '1P',
    '02': '2P',
    '03': '3P',
    '04': '4P',
    Final: '5P'
  }

  const objectDocente = {
    Doc_Amanda: '1XB1-v-m9lBm1scgXfX1sn7XK2OAmjsX4AVVHZUZ6Jhw',
    Doc_Angela: '1kwM2CVB8Rby7PGHdZhq9xXZ7KhAOSQRueuXxL8JpPKQ',
    Doc_Daniela: '1sOq0XaFy7_fJlOZQnWAFqJnFtyJq4NJUyQffMV-c6J8',
    Doc_Erika: '1XZGOUBAg6lOViB23YMpB3zLvXZSbxcygxJhTja5AiLU',
    Doc_Ferney: '18t8aEDKxRe9m6MydLV4VQlyNZc6LSQy8Ul5OvhtZVsI',
    Doc_John: '1GUZJvDYGAdy4H8UpNFPVJTO-gUb0Nd8yLpf7J90tauI',
    Doc_JuanPablo: '1SlCg1RXGajBBe4HEGR3Vda2kUd_lcM0D3atuzrccXdY',
    Doc_July: '1L9EUDudzmn8l0_9wdXkE-g_uNWwoCl5fDfPeX217-E0',
    Doc_Karyna: '1sejVpoJcWCsGCl44UPnHtf9Dk5esjlbbh2KF7Xln9ng',
    Doc_Katherine: 'hqPA5kedWn7yommZgBVK68JKs9Vpz_ppeQHwKY',
    Doc_Kevin: '1aB8F9iyw6V7J5O2aSI9fjrnovMqDT3zJ6aLA9id6X5w',
    Doc_Luisa: '13Ymh4ornfqJTfA4SDrqRGjtyGgIxLR5VRG-XAJFwDVI',
    Doc_Mateo: '1DldUG-_A_baDFpF5IbnJeP72sG1MxYagV4ZJn7A9oxQ',
    Doc_Oscar: '1zIJSceuJSs9sh3E5GPgiEef87AHc5gBRXr_bfbdbN7k',
    Doc_Patricia: '1oOVTRGI6y2o4qevMv8nzbdPurAjO5q8AZWqoLMulxY8',
    Doc_SorJaneth: '14cMmtnrNBf900U7TrXm1zO28QfJkRVVgtF4JDUT0HUk',
    Doc_Victoria: '1LANDaIZfbe9oGAJitER2Mpf0azZkcV6lvuwYQDfu6c4'
  }

  const periodoIngre = numero.periodo
  const periodo = objectPeriodos[periodoIngre] ?? 'N/A'

  const asignatura = numero.asignatura
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

  const id = numero.idnotas
  const hoja = SpreadsheetApp.openById(`${id}`).getSheetByName(gradoPlaCompleta)
  const dataCompleta = hoja.getDataRange().getValues()

  const docente = numero.idCalculadora
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

function hojasData(numeros) {
  const id = numeros.inputId
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

function evaluaciones(numeros) {
  let idPlanillaNotas = numeros.idnotasEva
  let materiaGrado = numeros.asignaturaEva
  let evaMateria = numeros.EvMateria
  let evaGrado = numeros.EvGrados
  let columna = numeros.columa
  let numeroEs = numeros.numeroEst
  let mayusColumna = columna.toUpperCase()
  let fila = numeros.fila

  const arrayevaMateria = evaMateria.split('_')
  let Eva = arrayevaMateria[0]
  let Mate = arrayevaMateria[1]

  const objectPeriodosEva = {
    '01': '1P',
    '02': '2P',
    '03': '3P',
    '04': '4P',
    Final: '5P'
  }

  let materiaPlaEva = Mate + '_' + evaGrado

  const objectEvaMaterias = {
    Eva_MAT: '1cyDyae-stuOSTLje0npIokayvJEPZzTPxpB-jhTIp38',
    Eva_SOC: '17fN-tEadh6WLUHXGzAZcVI7uN0Lv9F_Slv8g--CM8xo',
    Eva_LC: '1OgnITsMPKNIB4yVW5v1U0IzI0LJlng6z-JU954PxuVA',
    Eva_ING: '1bWSc3ctSnsxSYU-Sqot-lG_Dxi9HMZomBmue0hahUuo',
    Eva_CN: '14xV5uEn42TSzsumBxl1OrKYpQbV52sqdIl7TFq8Og1I'
  }

  const planillaNotas = SpreadsheetApp.openById(idPlanillaNotas).getSheetByName(materiaGrado)

  let hojaEva = objectEvaMaterias[evaMateria]

  const periodoEvaIngre = numeros.periodoEva
  const periodoEva = objectPeriodosEva[periodoEvaIngre] ?? 'N/A'
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

function curso(numeros) {
  const objectGrado = {
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

  const objectDocenteCur = {
    Doc_Amanda: '1XB1-v-m9lBm1scgXfX1sn7XK2OAmjsX4AVVHZUZ6Jhw',
    Doc_Angela: '1kwM2CVB8Rby7PGHdZhq9xXZ7KhAOSQRueuXxL8JpPKQ',
    Doc_Daniela: '1sOq0XaFy7_fJlOZQnWAFqJnFtyJq4NJUyQffMV-c6J8',
    Doc_Erika: '1XZGOUBAg6lOViB23YMpB3zLvXZSbxcygxJhTja5AiLU',
    Doc_Ferney: '18t8aEDKxRe9m6MydLV4VQlyNZc6LSQy8Ul5OvhtZVsI',
    Doc_John: '1GUZJvDYGAdy4H8UpNFPVJTO-gUb0Nd8yLpf7J90tauI',
    Doc_JuanPablo: '1SlCg1RXGajBBe4HEGR3Vda2kUd_lcM0D3atuzrccXdY',
    Doc_July: '1L9EUDudzmn8l0_9wdXkE-g_uNWwoCl5fDfPeX217-E0',
    Doc_Karyna: '1sejVpoJcWCsGCl44UPnHtf9Dk5esjlbbh2KF7Xln9ng',
    Doc_Katherine: 'hqPA5kedWn7yommZgBVK68JKs9Vpz_ppeQHwKY',
    Doc_Kevin: '1aB8F9iyw6V7J5O2aSI9fjrnovMqDT3zJ6aLA9id6X5w',
    Doc_Luisa: '13Ymh4ornfqJTfA4SDrqRGjtyGgIxLR5VRG-XAJFwDVI',
    Doc_Mateo: '1DldUG-_A_baDFpF5IbnJeP72sG1MxYagV4ZJn7A9oxQ',
    Doc_Oscar: '1zIJSceuJSs9sh3E5GPgiEef87AHc5gBRXr_bfbdbN7k',
    Doc_Patricia: '1oOVTRGI6y2o4qevMv8nzbdPurAjO5q8AZWqoLMulxY8',
    Doc_SorJaneth: '14cMmtnrNBf900U7TrXm1zO28QfJkRVVgtF4JDUT0HUk',
    Doc_Victoria: '1LANDaIZfbe9oGAJitER2Mpf0azZkcV6lvuwYQDfu6c4'
  }

  const objectPeriodoCur = {
    '01': 5,
    '02': 35,
    '03': 65,
    '04': 95,
    Final: 125
  }

  let docenteCur = numeros.idCalculadoraDoc
  let asignaturaGraCur = numeros.AsignaturaGrad
  let PeriodoCur = numeros.periodoGrad
  let GradoCur = numeros.grado
  let MateriaGradoCur = numeros.materiaGrado

  const idDocenteCur = objectDocenteCur[docenteCur] ?? 'N/A'
  let idPeriodoCur = objectPeriodoCur[PeriodoCur] ?? 'N/A'
  const idGrado = objectGrado[GradoCur] ?? 'N/A'

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
