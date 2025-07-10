function main(workbook: ExcelScript.Workbook) {
    // Obtener la hoja de cálculo.
    let sheet = workbook.getActiveWorksheet();
    sheet.getTable("ClassA").getRangeBetweenHeaderAndTotal().clear(ExcelScript.ClearApplyTo.contents);
    sheet.getTable("IRRS").getRangeBetweenHeaderAndTotal().clear(ExcelScript.ClearApplyTo.contents);
    workbook.getWorksheet("GFI").getRange("A2:K40").clear(ExcelScript.ClearApplyTo.contents);
}