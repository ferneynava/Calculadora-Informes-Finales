/* eslint-disable prefer-const */
/* eslint-disable no-undef */
/* eslint-disable no-unused-vars */
function doGet () {
  return HtmlService.createTemplateFromFile('index').evaluate()
}

function include (filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent()
}

function main (numero) {
  const hoja = SpreadsheetApp.openById('1J6iynccfCQ2cQkkM_FQtZ_HhtS1njq3OWM2Oud885V4').getSheetByName('3º')
  const dataNombre = hoja.getDataRange().getValues()
  const calculadora = SpreadsheetApp.openById('18t8aEDKxRe9m6MydLV4VQlyNZc6LSQy8Ul5OvhtZVsI').getSheetByName('Datos 1P')
  let dataCalculadora = calculadora.getDataRange().getValues()

  const numeroEstuden = dataNombre.map(filaNom => filaNom[0])
  numeroEstuden.shift()
  const arrayNombres = numeroEstuden.filter(eliminar => eliminar !== '')

  const arrayNotasPri = []
  const asignatura = numero.asignatura
  const arrayAsignatutas = asignatura.split('_')
  const materia = arrayAsignatutas[0]

  const objectAsignatura = {
    MAT: 'MatemáticasMatemáticas',
    TEC: 'TecnologíaTecnología',
    LC: 'Lengua CastellanaLengua Castellana',
    ETIC: 'ÉticaÉtica'
  }

  const asignaturaPlaUnificada = objectAsignatura[materia] ?? 'N/A'
  const arrayAsignaturasPlaUnificada = dataNombre[0]
  let celdaColumna = arrayAsignaturasPlaUnificada.indexOf(asignaturaPlaUnificada)

  let arrayNotas = dataNombre.map(filaNota => filaNota[celdaColumna])
  arrayNotas.shift()
  const periodos = numero.periodo

  if (periodos === '01') {
    const contador = 5
    for (let step = 0; step <= arrayNombres.length; step++) {
      const resultado = contador * step
      arrayNotasPri.push(arrayNotas[resultado])
    }
  }
  arrayNotasPri.pop()

  Logger.log(arrayNotasPri)
  // let data = dataNombre.map(fila => fila[1]).filter(periodo => periodo == periodos);
}
