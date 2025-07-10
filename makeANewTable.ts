function main(workbook: ExcelScript.Workbook) {
    // Obtener la hoja de cálculo.
    let sheet = workbook.getActiveWorksheet();
    workbook
    const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

    let hoy: Date = new Date(2025, 0, 21);
    //let tableName = meses[ayer.getMonth()].substring(0, 3) + "_" + ayer.getDate();

    let tablas = workbook.getTables();

    const sheetName: string = sheet.getName();
    if (sheetName.indexOf(meses[hoy.getMonth()].toUpperCase()) == -1) {
        sheet = makeNewSheet(workbook, meses[hoy.getMonth()], sheet.getPosition());
    }

    const newTableName = `${meses[hoy.getMonth()].substring(0, 3)}_${hoy.getDate()}`;

    let exists = tablas.some(tabla => newTableName == tabla.getName())
    let table: string[][] | boolean;
    if (exists) {
        return console.log("Tabla ya creada.");
    } else {
        table = getOldData(meses, hoy, tablas);
    }

    if(!table){
        return console.log("Tabla no encontrada.")
    }

    let newTable = makeNewTable(table);

    //table = tablas.find(tabla => tableName == tabla.getName())

    /*if (table)
        console.log(table.getName())
    else console.log("Tabla no encontrada")*/

    //console.log(validation.getRange().getAddress())
    /*const sheetName: string = sheet.getName();
    let newTable: string[][] = [];


    hoy = new Date();
    
 
    newTable = makeNewTable(workbook, table);
    let range = sheet.getUsedRange()
    
    let formulas = getAdds(meses, hoy, newTableName);
    let formulasRange: ExcelScript.Range;
    if (range) {
        range = sheet.getRange("A" + (sheet.getUsedRange().getRowCount() + 2)).getAbsoluteResizedRange(newTable.length, newTable[0].length);
        formulasRange = sheet.getRange("A" + (sheet.getUsedRange().getRowCount() + 1)).getAbsoluteResizedRange(formulas.length, formulas[0].length);
    } else {
        range = sheet.getRange("A2").getAbsoluteResizedRange(newTable.length, newTable[0].length);
        formulasRange = sheet.getRange("A1").getAbsoluteResizedRange(formulas.length, formulas[0].length);
    }
 
    try {
        sheet.addTable(range, true).setName(newTableName);
    } catch (error) {
        console.log("Nombre de tabla ya existente", error)
        return;
    }
    
    const excelTable = sheet.getTable(newTableName)
    excelTable.setPredefinedTableStyle(null);
    excelTable.getRange().copyFrom(validation.getRange(), ExcelScript.RangeCopyType.formats)
    range.setFormulas(newTable);
    formulasRange.setFormulas(formulas);
    //formulasRange.copyFrom();
    
 
    let format = excelTable.getRange().getFormat();
    format.getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous); // Top border
    format.getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous); // Bottom border
    format.getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous); // Left border
    format.getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous); // Right border*/

}


function makeNewSheet(workbook: ExcelScript.Workbook, mes: string, position: number) {
    let sheet = workbook.addWorksheet(mes.toUpperCase());
    sheet.setPosition(position + 1);
    sheet.activate();
    return sheet;
}

function getOldData(meses: string[], hoy: Date, tablas: ExcelScript.Table[]) {
    let ayer: Date;
    let tableName: string;
    let exists: ExcelScript.Table;

    for (let i = 0; i < tablas.length; i++) {
        ayer = new Date(hoy);
        ayer.setDate(hoy.getDate() - 1);
        tableName = meses[ayer.getMonth()].substring(0, 3) + "_" + (ayer.getDate());
        exists = tablas.find(tabla => tabla.getName() == tableName);
        hoy = ayer;
        if (exists) break;
    }
    if (exists) {
        let table = exists.getRange().getFormulas();
        return table;
    }else {
        return false;
    }

}

//type TablaMixta = (string | number | boolean)[][];

function makeNewTable(table: string[][] /*TablaMixta*/) {
    //let newSheet = workbook.getActiveWorksheet();

    let newTable: string[][] = [];

    for (let fila of table) {
        if (fila[6].toString().indexOf("DESCARG") == -1) {
            newTable.push(fila);
        }
    }
    //console.log(newTable);
    return newTable;
}

function getAdds(meses: string[], date: Date, tableName: string) {
    const dia = date.getDate();
    const mes = meses[date.getMonth()].toUpperCase();
    const year = date.getFullYear()

    let filas: string[][] = [];
    let formulas: string[] = [];

    formulas[0] = `${dia}/${date.getMonth() + 1}/${year}`;
    formulas[1] = `REPORTE DE IMPORTACIONES BC LOGISTICS ${dia} de ${mes} del ${year}`;
    formulas[2] = "";
    formulas[3] = "";
    formulas[4] = "";
    formulas[5] = "";
    formulas[6] = `="CORTE VERDE: "&COUNTIFS(${tableName}[ESTATUS],"CORTE PENDIENTE",${tableName}[Selectivo  DGA],"VERDE")&" | BRIDA: "&COUNTIF(${tableName}[ESTATUS],"* BRIDA*")`;
    formulas[7] = `="CORTE ROJO: "&COUNTIFS(${tableName}[ESTATUS],"CORTE PENDIENTE",${tableName}[Selectivo  DGA],"ROJO")`;
    formulas[8] = 'Priority';
    formulas[9] = `="DESCARGANDO: "&COUNTIF(${tableName}[ESTATUS],"EN PROCESO DE DESCARGA")`;
    formulas[10] = "";
    formulas[11] = `="AFORO: "&COUNTIF(${tableName}[ESTATUS],"AFORO")`;
    formulas[12] = `="LIQ: "&COUNTIF(${tableName}[ESTATUS],"LIQUIDADA")&" / MOD: "&COUNTIF(${tableName}[ESTATUS],"MÓDULO")`;
    formulas[13] = `="DESCARGADO: "&COUNTIF(${tableName}[ESTATUS],"DESCARGADO*")`;

    //formulas[11] = "ACTUALIZADO A LAS: " + date.toLocaleTimeString();
    filas.push(formulas);
    //console.log(filas);
    return filas;
}