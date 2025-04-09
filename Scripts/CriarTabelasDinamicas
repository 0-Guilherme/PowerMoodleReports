function main(workbook: ExcelScript.Workbook) {
    // Selecionar a planilha pra fazer as tabelas dinamicas e caso ja tenha sido feita alguma outra coluna neste relatorio por algum teste deve verificar o nome e remove-la
    let planilhas = workbook.getWorksheets();
    planilhas.forEach(planilha => {
        let nomePlanilha = planilha.getName();
        if (nomePlanilha === "Relatorio de Usuarios" || nomePlanilha === "Sheet2" || nomePlanilha === "Relatorios do SAE" || nomePlanilha === "Tabela de Lotacoes") {
            planilha.delete();
        }
    });

    // Adicionar a primeira planilha e a primeira tabela dinâmica
    // Esta planilha usa os nomes dos cursos como linhas; os tipos de usuarios como colunas; e para a contagem usaremos o nome do usuario
    let planilha1 = workbook.addWorksheet("Relatorio de Usuarios");
    let tabela1 = workbook.getTable("Tabela1");
    let pivotTable1 = workbook.addPivotTable("Relatorio de Usuarios", tabela1, planilha1.getRange("A1:A20"));
    pivotTable1.addRowHierarchy(pivotTable1.getHierarchy("Nome do Curso/Evento"));
    pivotTable1.addColumnHierarchy(pivotTable1.getHierarchy("Tipo de Usuario"));
    pivotTable1.addDataHierarchy(pivotTable1.getHierarchy("Nome Usuario"));

    // Adicionar a segunda planilha
    // Estaremos filtrando esta tabela para os usarios do SAE, e utilizaremos essa tabela de referencia para a segunda tabela dinamica
    let planilha12 = workbook.addWorksheet("Sheet2");
    tabela1.getColumnByName("Usuario do SAE").getFilter().applyValuesFilter(["Sim"]);
    let dadosVisiveis = tabela1.getRange().getVisibleView().getValues();
    let rangeFiltrado = planilha12.getRange(`A1:${String.fromCharCode(65 + dadosVisiveis[0].length - 1)}${dadosVisiveis.length}`);
    rangeFiltrado.setValues(dadosVisiveis);
    let tabelaFiltrada = planilha12.addTable(rangeFiltrado.getAddress(), true);

    // Criar a segunda tabela dinâmica
    // Esta planilha usa os nomes dos cursos como linhas; os tipos de usuarios como colunas; e para a contagem usaremos a coluna Usuario do SAE
    let planilha2 = workbook.addWorksheet("Relatorios do SAE");
    let pivotTable2 = workbook.addPivotTable("Relatorios do SAE", tabelaFiltrada, planilha2.getRange("A1:A20"));
    pivotTable2.addRowHierarchy(pivotTable2.getHierarchy("Nome do Curso/Evento"));
    pivotTable2.addColumnHierarchy(pivotTable2.getHierarchy("Tipo de Usuario"));
    pivotTable2.addDataHierarchy(pivotTable2.getHierarchy("Usuario do SAE"));

    // Adicionar a terceira planilha e tabela dinâmica
    // Esta planilha serve para fazer a contagem de usuarios distribuidos pelo seu departamento. Para isso é utilizado a coluna "Lotacao" para linhas e para a contagem de dados
    let planilha3 = workbook.addWorksheet("Tabela de Lotacoes");
    let pivotTable3 = workbook.addPivotTable("Tabela de Lotacoes", tabela1, planilha3.getRange("A1:A20"));
    pivotTable3.addRowHierarchy(pivotTable3.getHierarchy("Lotacao"));
    pivotTable3.addDataHierarchy(pivotTable3.getHierarchy("Lotacao"));

    // Ocultar Sheet 2
    // Nao tem necessidade do usuario ter acesso a tabela de referencia para o SAE, visto que a tabela inicial possui a possibilidade de filtros/ ou ver os dados por inteiro
    let selectedSheet = workbook.getWorksheet("Sheet2");
    selectedSheet.setVisibility(ExcelScript.SheetVisibility.hidden);

    // Desmarcar Filtros
    // Para deixar tabela inicial totalmente visivel
	  tabela1.getColumnByName("Usuario do SAE").getFilter().clear();
}
