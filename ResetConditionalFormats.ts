
function main(workbook: ExcelScript.Workbook) {
  // Obtener la celda activa y la hoja de cálculo.
  const hoja = workbook.getActiveWorksheet();

  // Obtiene todos los formatos condicionales de la hoja
  const range = hoja.getRange("A1:XDF1048576");
  const formatosCondicionales = range.getConditionalFormats();

  // Elimina cada uno
  //formatosCondicionales.forEach(formato => formato.delete());
  if (formatosCondicionales.length > 0) {
    for (let i = formatosCondicionales.length - 1; i >= 0; i--) {
      formatosCondicionales[i].delete();
    }
  }

  //Añado regla formato Descarg a la línea
  let formula = '=OR(ISNUMBER(SEARCH("PARCIAL",$G1)),ISNUMBER(SEARCH("DESCARG",$G1)))';
  /*
  *  Parámetros para la función (libro, rango, fórmula?, color de celda?, color de fuente?, rango 2?, negrita?)
  *  negrita debe de ser true or false.
  *  "?" significa opcional. Si no se define poner "undefined"
  */
  setNewCustomRule(workbook, "A1:F1048576", formula, "#F1A983", undefined, "H1:Z1048576");

  //Añado regla Brida a rangos G:G y Z:Z
  setContainTextRule(workbook, "G:G", "BRIDA", "#00B0F0")
  setContainTextRule(workbook, "Z:Z", "BRIDA", "#00B0F0")

  //Añado regla de proceso de descarga
  setContainTextRule(workbook, "G:G", "PROCESO DE DESCARGA", "#FFFF00")

  //Añade la fórmula de "descarg" en toda la línea
  formula = '=AND($Q1<>"",ISNUMBER(SEARCH("DESCARG",$G1)))';
  setNewCustomRule(workbook, "G:G", formula, "#F1A983",);

  //Añade regla de fórmula ver a selectivos
  formula = '=$M1="VERDE"'
  setNewCustomRule(workbook, "G:G", formula, "#8ED973", undefined, "M:M")


  //Añade la fórmula color rojo y negrita para los selectivos rojos
  formula = '=OR($M1="ROJO",$M1="VERDE POTESTAD")'
  setNewCustomRule(workbook, "A1:Z1048576", formula, undefined, "#FF0000", undefined, true);

  //Añado fórmula repetir valores en rangos F:F y L:L
  setRepeatRule(workbook, "F:F", "#FFC7CE", "#9C0006");
  setRepeatRule(workbook, "L:L", "#FFC7CE", "#9C0006");

  setContainTextRule(workbook, "G:G", "EN TRÁNSITO", "#FFC000")

}

function setNewCustomRule(workbook: ExcelScript.Workbook, rango1: string, formula?: string, cellColor?: string, fontColor?: string, rango2?: string, bold?: boolean) {
  const hoja = workbook.getActiveWorksheet();

  //Validación de los rangos
  const rangoTotal = rango1 == undefined ? hoja.getRange(rango2) :
    rango2 == undefined ? hoja.getRange(rango1) : hoja.getRanges(rango1 + "," + rango2);

  //Creación del formato condicional según fórmula
  let nuevaRegla: ExcelScript.ConditionalFormat;
  if (formula != undefined) {
    nuevaRegla = rangoTotal.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    //podría haber ocupado el getCustom() en la asignación de la variable, pero lo entiendo una vez ya lo hice :)

    nuevaRegla.getCustom().getRule().setFormula(formula);
    //Los formatos personalizados por fórmulas no necesitan un "setRule()", sólo se obtiene el custom, la regla y se asigna la fórmula, nada más.
  }
  //relleno de celda
  if (cellColor != undefined)
    nuevaRegla.getCustom().getFormat().getFill().setColor(cellColor);

  //color de celda
  if (fontColor != undefined)
    nuevaRegla.getCustom().getFormat().getFont().setColor(fontColor);

  //negrita
  if (bold)
    nuevaRegla.getCustom().getFormat().getFont().setBold(bold);
}

function setRepeatRule(workbook: ExcelScript.Workbook, range: string, cellColor?: string, fontColor?: string, bold?: boolean) {
  const hoja = workbook.getActiveWorksheet();

  //creo el formato condicional y lo asigno según criterios predefinidos "Resaltar duplicados"
  const formatoRepeat = hoja.getRange(range).addConditionalFormat(ExcelScript.ConditionalFormatType.presetCriteria).getPreset();
  //getPreset es el método que me permite añadir formato a las celdas y permite añadir las reglas al rango.


  if (cellColor != undefined)
    formatoRepeat.getFormat().getFill().setColor(cellColor);

  if (fontColor != undefined)
    formatoRepeat.getFormat().getFont().setColor(fontColor);

  if (bold)
    formatoRepeat.getFormat().getFont().setBold(bold);

  //duplicateRule permite crear la regla, que luego será añadida como tal a formatoRepeat1, que ya posee el getPreset()
  const duplicateRule: ExcelScript.ConditionalPresetCriteriaRule = { criterion: ExcelScript.ConditionalFormatPresetCriterion.duplicateValues };
  //Parece que no se puede añadir este tipo de valores a variables no tipadas, propio de Typescript.

  //Asigno la regla al rango.
  formatoRepeat.setRule(duplicateRule);
}

function setContainTextRule(workbook: ExcelScript.Workbook, range: string, texto: string, cellColor?: string, fontColor?: string, bold?: boolean) {
  const hoja = workbook.getActiveWorksheet();

  //Código para aplicar formato a celdas que contienen texto.
  const formatoContainText = hoja.getRange(range).addConditionalFormat(ExcelScript.ConditionalFormatType.containsText).getTextComparison();
  //getTextComparison() es semejante a getPreset(), como defino que tipo de formato condicional es, necesito hacer uso de él.

  if (cellColor != undefined)
    formatoContainText.getFormat().getFill().setColor(cellColor);

  if (fontColor != undefined)
    formatoContainText.getFormat().getFont().setColor(fontColor);

  if (bold)
    formatoContainText.getFormat().getFont().setBold(bold);

  //Creo la regla y la añado al rango deseado.
  const textRule: ExcelScript.ConditionalTextComparisonRule = { operator: ExcelScript.ConditionalTextOperator.contains, text: texto }
  formatoContainText.setRule(textRule);
}


