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

function main(numero) {
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
  const periodos = numero.periodo

  const idCal = numero.idcalculadora
  const calculadora = SpreadsheetApp.openById(`${idCal}`).getSheetByName('Datos 1P')
  let dataCalculadora = calculadora.getDataRange().getValues()
  let arrayAsignatura = dataCalculadora[2]

  const numeroEstuden = dataCompleta.map(filaNom => filaNom[0])
  numeroEstuden.shift()
  const arrayNombres = numeroEstuden.filter(eliminar => eliminar !== '')

  let celdaColumnaAsig = arrayAsignatura.indexOf(asignatura)
  const asignaturaPlaCompleta = objectAsignatura[materia] ?? 'N/A'
  const arrayAsignaturasPlaCompleta = dataCompleta[0]
  let celdaColumna = arrayAsignaturasPlaCompleta.indexOf(asignaturaPlaCompleta)
  let n = 1
  let c = 9
  if (periodos === '01') {
    for (let i = 0; i < arrayNombres.length; i++) {
      n = n + 5
      let notas = hoja.getRange(n, celdaColumna + 1).getValues()
      c = c + 1
      calculadora.getRange(c, celdaColumnaAsig + 1).setValue(notas)
    }
  }// let data = dataNombre.map(fila => fila[1]).filter(periodo => periodo == periodos);
}
