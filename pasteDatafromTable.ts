function main(workbook: ExcelScript.Workbook, inputData: string) {
  // TODO: Write code or use the Insert action button below.
  let data:string[][] = JSON.parse(inputData);
  
  let sheet = workbook.getActiveWorksheet();
  let formulas = getAdds();
  let rangeForm = sheet.getRange("A1").getAbsoluteResizedRange(formulas.length, formulas[0].length);
  rangeForm.setFormulas(formulas);

  sheet.getRange("A2:V50").clear(ExcelScript.ClearApplyTo.contents);
  let rangePaste = sheet.getRange("A2").getAbsoluteResizedRange(data.length, data[0].length)
  
  //sheet.getRange("A2");
  rangePaste.setValues(data);

  const date = new Date();
  sheet.getRange("L1").setValue("ACTUALIZADO A LAS: " + date.toLocaleTimeString('es-ES', { timeZone: 'America/Managua' }));
}

function getAdds() {
  const date = new Date();
  const dia = date.getDate();
  const mes = date.getMonth() + 1;
  const year = date.getFullYear()
  /*const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const tableName = meses[date.getMonth()].substring(0, 3) + "_" + date.getDate();*/
  let filas:string[][] = [];
  let formulas: string[] = [];
  
  formulas[0] = `REPORTE DE IMPORTACIONES BC LOGISTICS ${dia}/${mes}/${year}`;
  formulas[1] = "";
  formulas[2] = "";
  formulas[4] = "";
  formulas[3] = '="CORTE VERDE: "&COUNTIFS(G3:G50,"CORTE PENDIENTE",M3:M50,"VERDE")&" | BRIDA: "&COUNTIF(G3:G50,"*BRIDA*")';
  formulas[5] = '="CORTE ROJO: "&COUNTIFS(G3:G50,"CORTE PENDIENTE",M3:M50,"ROJO")';
  formulas[7] = '="DESCARGANDO: "&COUNTIF(G3:G50,"EN PROCESO DE DESCARGA")';
  formulas[6] = '="AFORO: "&COUNTIF(G3:G50,"AFORO")';
  formulas[8] = '="LIQ: "&COUNTIF(G3:G50,"LIQUIDADA")&" / MOD: "&COUNTIF(G3:G50,"MÓDULO")';
  formulas[9] = `="DESCARGADO: "&COUNTIFS(G3:G50,"DESCARGADO*",R3:R50,NUMBERVALUE("${dia}/${mes}/${year}"))`;
  formulas[10] = "";
  //formulas[11] = "ACTUALIZADO A LAS: " + date.toLocaleTimeString();
  filas.push(formulas);
  //console.log(filas);
  return filas;
}