function main(workbook: ExcelScript.Workbook) {
  // Get the active cell and worksheet.
  let sheet = workbook.getActiveWorksheet();

  //Consigo el mes y el día actual, el cual es el nombre de la tabla.
  const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  let fecha = new Date();
  let tableName = meses[fecha.getMonth()].substring(0, 3) + "_" + fecha.getDate();
  //console.log(tableName); //Imprime el nombre de la tabla.

  //Obtengo los valores de la tabla.
  let validation: ExcelScript.Table;
  do {//loop de validación
    validation = sheet.getTable(tableName);//Valido si la tabla existe.
    if (!validation) {//Si no existe...
      let ayer: Date = new Date(fecha);
      ayer.setDate(fecha.getDate() - 1);
      tableName = meses[ayer.getMonth()].substring(0, 3) + "_" + (ayer.getDate());
      //console.log(tableName);
      fecha = ayer;
    }
  } while (!validation);//repito el proceso hasta que la tabla exista.

  let table = validation.getRange().getValues();//Obtengo finalmente los valores de la tabla existente.

  return { table }//retorno como un objeto el valor de la tabla.
  //

}