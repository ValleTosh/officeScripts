function main(workbook: ExcelScript.Workbook) {
    const hoja = workbook.getActiveWorksheet();

    // Define los rangos
    const rango1 = hoja.getRange("A1:F1048576");
    const rango2 = hoja.getRange("H1:Z1048576");
    const rangoTotal = hoja.getRanges(rango1.getAddress() + "," + rango2.getAddress());

    // Fórmula del formato condicional
    const formula = '=OR(ISNUMBER(SEARCH("PARCIAL",$G1)),ISNUMBER(SEARCH("DESCARG",$G1)))';

    // Crea la nueva regla
    const nuevaRegla = rangoTotal.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    nuevaRegla.getCustom().getRule().setFormula(formula);
    nuevaRegla.getCustom().getFormat().getFill().setColor("#F1A983");

    // Obtener todas las reglas existentes
    const reglas = hoja.getUsedRange().getConditionalFormats();

    // Si hay más de 4 reglas, mover las que están en la posición 5 o superior hacia abajo
    if (reglas.length >= 5) {
        for (let i = reglas.length - 1; i >= 4; i--) {
            reglas[i].setPriority(i + 2); // Mover hacia abajo
        }
    }

    // Establecer la prioridad deseada
    nuevaRegla.setPriority(5);


    //formato.getFont().setColor("black");
    //formato.getFont().setBold(true);


    /* const hoja = workbook.getActiveWorksheet();
     const celda = hoja.getRange("G59");
 
     // Obtener el color de relleno
     const colorRelleno = celda.getFormat().getFill().getColor();
 
     // Mostrar el color en la consola (solo visible en el entorno de scripts)
     console.log("Color de relleno de A1:", colorRelleno);*/


    /* // Obtiene todos los formatos condicionales de la hoja
     const range = hoja.getUsedRange();
     const formatosCondicionales = range. getConditionalFormats();
 
     formatosCondicionales.forEach((format, index) => {
         console.log(`Conditional Format ${index + 1}:`);
         console.log(`Type: ${format.getType()}`);
         console.log(`Priority: ${format.getPriority()}`);
         console.log(`Rule: ${JSON.stringify(format.getCustom().getRule().getFormula())}`);
     });*/
    // Elimina cada uno
    //formatosCondicionales.forEach(formato => formato.delete());

}