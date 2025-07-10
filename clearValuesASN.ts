function main(workbook: ExcelScript.Workbook) {
    // Obtener la hoja de cálculo.
    let sheet = workbook.getActiveWorksheet();
    sheet.getRangeByIndexes(1, 0, sheet.getUsedRange().getRowCount(), 3).clear();
}