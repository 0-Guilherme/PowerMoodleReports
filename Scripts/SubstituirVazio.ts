function main(workbook: ExcelScript.Workbook) {
    // Obter a planilha "Sheet1", nome da planilha que é exportada no reatório do moodle
    const ws = workbook.getWorksheet("Sheet1");

    // Encontrar a última linha com dados na coluna I - Coluna a qual devo remover os espaços em branco
    const ultimaLinha = ws.getRange("I:I").getUsedRange().getRowCount();

    // Loop através das linhas da coluna I e J
    for (let i = 0; i < ultimaLinha; i++) { // i começa em 0 para corresponder à linha 1 no Excel
        // Verificar e substituir células vazias na coluna I
        const valorCelulaI = ws.getRange(`I${i + 1}`).getValue(); // A linha 1 no Excel corresponde ao índice 0 no TypeScript
        if (valorCelulaI === "" || valorCelulaI === "Não") {
            ws.getRange(`I${i + 1}`).setValue(0);
        }

        const valorCelulaII = ws.getRange(`F${i + 1}`).getValue(); // A linha 1 no Excel corresponde ao índice 0 no TypeScript
        if (valorCelulaII === "") {
            ws.getRange(`F${i + 1}`).setValue("VAZIO");
        }
    }
}
