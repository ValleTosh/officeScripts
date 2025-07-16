
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

  /*
  *  Parámetros para la función (libro, rango, fórmula?, color de celda?, color de fuente?, rango 2?, negrita?)
  *  negrita debe de ser true or false.
  *  "?" significa opcional. Si no se define poner "undefined"
  */

  //Añado regla valores en 0 se ven en blanco.
  setCellValueRule(workbook, "J:J", "0", undefined, "#FFFFFF");

  //Las celdas en blanco les quito el formato condicional
  setPresetRule(workbook, "I:I", "blank", "#FFFFFF");

  //Si no contienen "Received" se pone en amarillo
  setContainTextRule(workbook, "I3:I10014", "Received_notContains", "#FFFF00");

  //Si ya se subió la declaración al FDM4 se quita el formato para la celda.
  let formula = '=$AG1 <> ""';
  setNewCustomRule(workbook, "Q:Q", formula, "clear");

  //Si no se ha subido la factura
  formula = '=$S3="NO"';
  setNewCustomRule(workbook, "A3:T10014", formula, "#C6E0B4");

  //Si no se ha enviado por correo el ASN.
  formula = '=$R3="NO"';
  setNewCustomRule(workbook, "A3:T10014", formula, "#FFFF00");

  //La columna declaración se pone en amarillo cuando se pone datos en la columna de precinto.
  formula = '=E1 <> ""';
  setNewCustomRule(workbook, "Q:Q", formula, "#FFFF00");

  //La letra se pone en rojo cuando encuentra la palabra "ROJO"
  setContainTextRule(workbook, "I:I", "Rojo_contains", undefined, "#FF0000", true);

  //Repetición precinto, factura y declaración
  setPresetRule(workbook, "F:F", "repeat", "#FFC7CE", "#9C0006");
  setPresetRule(workbook, "L:L", "repeat", "#FFC7CE", "#9C0006");
  setPresetRule(workbook, "E:E", "repeat", "#FFC7CE", "#9C0006");
  //#C6E0B4 - verde
  //#000000 - negro
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
  if (cellColor != undefined) {
    if (cellColor == "clear") nuevaRegla.getCustom().getFormat().getFill().setColor("#FFFFFF");
    else nuevaRegla.getCustom().getFormat().getFill().setColor(cellColor);
  }


  //color de celda
  if (fontColor != undefined)
    nuevaRegla.getCustom().getFormat().getFont().setColor(fontColor);

  //negrita
  if (bold)
    nuevaRegla.getCustom().getFormat().getFont().setBold(bold);
}

function setPresetRule(workbook: ExcelScript.Workbook, range: string, preset: string, cellColor?: string, fontColor?: string, bold?: boolean) {
  const hoja = workbook.getActiveWorksheet();

  //creo el formato condicional y lo asigno según criterios predefinidos "Resaltar duplicados"
  const formato = hoja.getRange(range).addConditionalFormat(ExcelScript.ConditionalFormatType.presetCriteria).getPreset();
  //getPreset es el método que me permite añadir formato a las celdas y permite añadir las reglas al rango.


  if (cellColor != undefined)
    formato.getFormat().getFill().setColor(cellColor);

  if (fontColor != undefined)
    formato.getFormat().getFont().setColor(fontColor);

  if (bold)
    formato.getFormat().getFont().setBold(bold);

  //duplicateRule permite crear la regla, que luego será añadida como tal a formatoRepeat1, que ya posee el getPreset()
  let rule: ExcelScript.ConditionalPresetCriteriaRule;

  switch (preset) {
    case "repeat": rule = { criterion: ExcelScript.ConditionalFormatPresetCriterion.duplicateValues }; break;
    case "blank": rule = { criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks }; break;
    //case "": rule = { criterion: ExcelScript.ConditionalFormatPresetCriterion }; break;
  }

  //Asigno la regla al rango.
  formato.setRule(rule);
}

function setContainTextRule(workbook: ExcelScript.Workbook, range: string, texto_cond: string, cellColor?: string, fontColor?: string, bold?: boolean) {
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

  const cond = texto_cond.split("_")[1];
  const texto = texto_cond.split("_")[0];

  let textRule: ExcelScript.ConditionalTextComparisonRule;
  //Creo la regla y la añado al rango deseado.
  switch (cond) {
    case "contains": textRule = { operator: ExcelScript.ConditionalTextOperator.contains, text: texto }; break;
    case "notContains": textRule = { operator: ExcelScript.ConditionalTextOperator.notContains, text: texto }; break;
  }
  formatoContainText.setRule(textRule);
}

function setCellValueRule(workbook: ExcelScript.Workbook, range: string, cellValue: string, cellColor?: string, fontColor?: string, bold?: boolean) {
  const hoja = workbook.getActiveWorksheet();

  //Código para aplicar formato a celdas que contienen texto.
  const formatoCellValue = hoja.getRange(range).addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue).getCellValue();
  //getCellValue() es semejante a getPreset(), como defino que tipo de formato condicional es, necesito hacer uso de él.

  if (cellColor != undefined)
    formatoCellValue.getFormat().getFill().setColor(cellColor);

  if (fontColor != undefined)
    formatoCellValue.getFormat().getFont().setColor(fontColor);

  if (bold)
    formatoCellValue.getFormat().getFont().setBold(bold);

  //Creo la regla y la añado al rango deseado.
  const valueRule: ExcelScript.ConditionalCellValueRule = { operator: ExcelScript.ConditionalCellValueOperator.equalTo, formula1: cellValue };
  formatoCellValue.setRule(valueRule);
}

function setBlankValue(workbook: ExcelScript.Workbook, range: string, cellColor?: string, fontColor?: string, bold?: boolean) {
  const hoja = workbook.getActiveWorksheet();

  //creo el formato condicional y lo asigno según criterios predefinidos "Resaltar duplicados"
  const formatoBlank = hoja.getRange(range).addConditionalFormat(ExcelScript.ConditionalFormatType.presetCriteria).getPreset();
  //getPreset es el método que me permite añadir formato a las celdas y permite añadir las reglas al rango.

  if (cellColor != undefined)
    formatoBlank.getFormat().getFill().setColor(cellColor);

  if (fontColor != undefined)
    formatoBlank.getFormat().getFont().setColor(fontColor);

  if (bold)
    formatoBlank.getFormat().getFont().setBold(bold);

  //blankRule permite crear la regla, que luego será añadida como tal a formatoRepeat1, que ya posee el getPreset()
  const blankRule: ExcelScript.ConditionalPresetCriteriaRule = { criterion: ExcelScript.ConditionalFormatPresetCriterion.blanks };
  //Parece que no se puede añadir este tipo de valores a variables no tipadas, propio de Typescript.

  //Asigno la regla al rango.
  formatoBlank.setRule(blankRule);
}