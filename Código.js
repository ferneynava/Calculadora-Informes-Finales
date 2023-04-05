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
  const id = numero.idnotas
  const hoja = SpreadsheetApp.openById(`${id}`).getSheetByName('3º')
  const dataCompleta = hoja.getDataRange().getValues()
  const asignatura = numero.asignatura
  const arrayAsignatutas = asignatura.split('_')
  const materia = arrayAsignatutas[0]

  const objectAsignatura = {
    MAT: 'MatemáticasMatemáticas',
    TEC: 'TecnologíaTecnología',
    LC: 'Lengua CastellanaLengua Castellana',
    ETIC: 'ÉticaÉtica'
  }

  const idCal = numero.idcalculadora
  const calculadora = SpreadsheetApp.openById(`${idCal}`).getSheetByName('Datos 1P')
  let dataCalculadora = calculadora.getDataRange().getValues()
  let arrayAsignatura = dataCalculadora[2]

  const numeroEstuden = dataCompleta.map(filaNom => filaNom[0])
  numeroEstuden.shift()
  const arrayNombres = numeroEstuden.filter(eliminar => eliminar !== '')

  const periodos = numero.periodo
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
