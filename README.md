# PowerMoodleReports

## Resume
This project leverages automated processes powered by Power Automate to efficiently format and generate dynamic tables. It is designed to create customized reports derived directly from Moodle, ensuring a streamlined and tailored reporting experience.

## Resumo
Este projeto aproveita processos automatizados alimentados pelo Power Automate para formatar e gerar tabelas dinâmicas de maneira eficiente. Ele foi projetado para criar relatórios personalizados diretamente extraídos do Moodle, garantindo uma experiência de geração de relatórios mais simplificada e ajustada.

---

## Softwares Utilizados

- **Power Automate**
- **Moodle**
- **Excel**

---

## Descrição

O projeto consiste em personalizar relatórios no Moodle, definir o público-alvo, configurar o ciclo do cronograma e o método de envio do relatório.  
Em seguida, utiliza o Power Automate para receber os relatórios, processá-los e encaminhá-los ao público-alvo designado.

---

## Configurações do Moodle

Para configuração dos Relatórios Personalizados:

**Caminho no Moodle:**  
Administração do Site > Relatórios > Construtor de Relatórios > Relatórios Personalizados

1. **Novo Relatório**:
   - **Nome**: Preencha o nome do relatório.
   - **Fonte do Relatório**: Escolha a base de origem dos dados.
   - **Tags**: Adicione, se necessário.
   - **Opções**: Habilite "Incluir Configuração Padrão" e "Remover Qualquer Linha Duplicada", se aplicável.

2. **Configuração do Relatório**:
   - Configure e renomeie as colunas que serão utilizadas.
   - Adicione as condições adequadas ao relatório.

**Exemplo de Configuração Utilizada**:
- **Nome**: Relatorio_ID_1
- **Fonte**: Participantes do Curso
- **Colunas**: Nome do Curso, ID do Curso, Nome de Usuário, Tipo de Usuário, Cargo, Departamento, Conclusão de Curso, Classificação de Usuário Externo, E-mail.
- **Condições**:
  - Número de Identificação do Curso > Inicia com: > 1
  - Conclusão de Curso > Tempo Concluído > Anterior > 1 Semana(s)

**Preview e Download**:
- Clique em "Ver Prévia" para verificar os dados do relatório.
- Para baixar o relatório, role até o final da página, selecione o tipo de arquivo em "Baixar dados da tabela como" e clique em "Download".

---

## Configuração do Power Automate
O arquivo exportado com a ordem do fluxo está disponível na pasta PowerAutomateFlux

### Etapas:

1. **Gatilho**:  
   Quando um novo e-mail é recebido (v3) - Outlook Business:
   - **De**: noreply@moodle.com  
   - **Incluir Anexos**: Sim  
   - **Importância**: Qualquer  
   - **Somente com Anexos**: Sim  
   - **Pasta**: Inbox (configurada para receber os relatórios)

2. **Loop: Para Cada**:  
   - **Criar Arquivo** - Outlook for Business:
     - **Caminho da Pasta**: `/Relatorios_Moodle_1-3`
     - **Nome do Arquivo**: `FuncaoDataFormatada_NomeVinculado`
       - `FuncaoDataFormatada`: Utilizar o código: `formatDateTime(triggerOutputs()?['body/receivedDateTime'], 'dd-MM-yyyy_HH-mm')`
       - `NomeVinculado`: Nome do relatório recebido no Outlook
     - **Conteúdo do Arquivo**: `ContentBytes` do anexo do Outlook

   - **Criar Tabela** - Excel Business:
     - **Localização**: OneDrive for Business
     - **Biblioteca de Documentos**: OneDrive
     - **Arquivo**: `/Relatorios_Moodle_1-3/NomeVinculado`
     - **Intervalo de Tabela**: `A:K`

   - **Listar Linhas Presentes em uma Tabela** - Excel Business:
     - Configurações semelhantes ao passo anterior, focando na `Tabela1`.

   - **Scripts**:
     - **RemoverDuplicatas**
     - **SubstituirVazio**
     - **CriarTabelasDinamicas**

3. **Condição SE**:
   - **Nome do Arquivo Contém ID_1**:
     - Verdadeiro: Enviar e-mail com ID=1
   - **Nome do Arquivo Contém ID_2**:
     - Verdadeiro: Enviar e-mail com ID=2
   - **Se Falso**: Enviar e-mail com ID=3.

---

## Scripts Utilizados

Os scripts estão localizados dentro da pasta "Scripts". São eles:

- **RemoverDuplicatas**: Remove linhas duplicadas no Excel.
- **SubstituirVazio**: Substitui células vazias por valores padrão.
- **CriarTabelasDinamicas**: Gera tabelas dinâmicas automaticamente.

---

## Referências

- [Documentação do Moodle](https://docs.moodle.org)
- [Documentação do Power Automate](https://learn.microsoft.com/power-automate)
- [Documentação do Excel Online](https://support.microsoft.com/excel)

        

 

