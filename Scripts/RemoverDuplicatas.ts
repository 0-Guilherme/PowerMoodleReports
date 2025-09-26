function main(workbook: ExcelScript.Workbook) {
    // Obtenha a planilha chamada "Sheet1"
    let sheet = workbook.getWorksheet("Sheet1");

    // Obtenha o intervalo usado dinamicamente
    let usedRange = sheet.getUsedRange();
    let values = usedRange.getValues();
    let rowCount = values.length;
    let columnCount = values[0].length;

    // Adicione uma coluna auxiliar dinamicamente, proporcional ao número de colunas utilizadas no relatorio
    let auxiliaryColumnIndex = columnCount; // Índice da nova coluna

    // Criar uma concatenacao para remover linhas duplicadas, neste caso vamos concatenar o ID unico dos cursos com o e-mail unico dos usuários
    // Iterar pelas linhas para concatenar valores de B e K na nova coluna auxiliar
    for (let i = 1; i < rowCount; i++) {
        let valueB = values[i][1]?.toString(); // Converter para string
        let valueK = values[i][10]?.toString(); // Converter para string
        if (valueB && valueK) { // Certifique-se de que as colunas têm valores válidos
            values[i][auxiliaryColumnIndex] = `${valueB} ${valueK}`;
        } else {
            values[i][auxiliaryColumnIndex] = ""; // Em caso de algum problema no relatorio, envolvendo dados do usuarop, preencher com vazio caso não haja valores
        }
    }

    // Criar um mapa para identificar duplicatas com base na coluna auxiliar
    let uniqueValues: Map<string, boolean> = new Map();
    let rowsToDelete: number[] = []; // Índices das linhas duplicadas a serem deletadas

    for (let i = 1; i < rowCount; i++) {
        let key: string = values[i][auxiliaryColumnIndex]?.toString(); // Usar a coluna auxiliar como chave
        if (key && !uniqueValues.has(key)) {
            uniqueValues.set(key, true); // Adiciona chave ao mapa
        } else {
            rowsToDelete.push(i); // Marca a linha como duplicada
        }
    }

    // Deletar as linhas duplicadas da planilha
    for (let i = rowsToDelete.length - 1; i >= 0; i--) { // Iterar de trás para frente para evitar problemas de índice
        let rowIndex = rowsToDelete[i];
        sheet.getRange(`${rowIndex + 1}:${rowIndex + 1}`).delete(ExcelScript.DeleteShiftDirection.up);
    }

    // Autoajustar a largura das colunas após a remoção das linhas, para melhor leitura por parte do usuario
    sheet.getUsedRange().getFormat().autofitColumns();
}
