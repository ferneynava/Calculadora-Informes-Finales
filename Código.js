/* eslint-disable space-before-function-paren */
/* eslint-disable prefer-const */
/* eslint-disable no-undef */
/* eslint-disable no-unused-vars */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
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

  const idCal = numero.idcalculadora
  const calculadora = SpreadsheetApp.openById(`${idCal}`).getSheetByName(`Datos ${periodo}`)
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

function evaluaciones(numeros) {
  Logger.log(numeros)
}
