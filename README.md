# PowerMoodleReports
Resume: This project leverages automated processes powered by Power Automate to efficiently format and generate dynamic tables. It is designed to create customized reports derived directly from Moodle, ensuring a streamlined and tailored reporting experience.
Resumo: Este projeto aproveita processos automatizados alimentados pelo Power Automate para formatar e gerar tabelas dinâmicas de maneira eficiente. Ele foi projetado para criar relatórios personalizados diretamente extraídos do Moodle, garantindo uma experiência de geração de relatórios mais simplificada e ajustada
---
Softwares Utilizados:

- Power Automate
- Moodle
- Excel
---
Descrição:

O projeto consiste em personalizar relatórios no Moodle, definir o público-alvo, configurar o ciclo do cronograma e o método de envio do relatório;
A próxima etapa é utilizar o power automate para receber os relatórios, processa-los e fazer o encaminhamento para o público-alvo.
---
Configurações do Moodle
Para configuração dos Relatórios Personalizados:

Caminho no Moodle: Administração do Site > Relatórios > Construtor de Relatórios > Relatórios Personalizados

Selecionar > Novo Relatório
  - Preencher: Nome (do Relatório)    
  - Selecionar: Fonte do Relatório (qual será a base de origem dos dados)
  - Tags se necessário
  - Opções que podem ser utilizadas: Incluir Configuração padrão e Remover qualquer linha duplicada

Com o Relatório aberto, selecionar e configurar quais colunas devem ser utilizadas, renomeando-as quando for necessário.

Após, configurar as condições para o relatório.

Configuração Utilizada:
  - Nome: Relatorio_ID_1
  - Fonte: Participantes do Curso
  - Colunas: Nome do Curso, ID do Curso, Nome de Usuário, Tipo de Usuário, Cargo, Departamento, Conclusão de Curso, Classificação de Usuário Externo, E-mail.

  - Condições:
    * Número de Idenificação do Curso >  Inicia com: > 1
    * Conclusão de Curso > Tempo Concluído > Anterior > 1 > Semana(s)
    * Desta forma o relatório ira rodar para cursos que iniciam o seu ID com o digito 1 e o periodo será para a ultima semana.

> A partir daqui é possível ver uma prévia do relatório, clicando em > Ver Prévia
> Caso queira obter a prévia do relatório, é possível ir até o final da página, em "Baixar dados da tabela como" selecionando o tipo de arquivo e clicando em Download.
---
Configuração do PowerAutomate
*Deixarei uma seção separada para os scripts que estão citados neste campo.

Etapas:
  - Gatilho: Quando um novo email é rebido (v3) - Outlook Business
      =De: Colocar o e-mail de origem dos relatórios ex: noreply@moodle.com
      =Incluir Anexos: Sim
      =Importância: Qualquer
      =Somente com Anexos: Sim
      =Pasta: Inbox (Pasta configurada para receber os relatórios)
  - Loop: Para Cada (For Each):
      - Criar Arquivo - Outlook For Business
          Caminho da Pasta: Selecionar o caminho para a pasta que vai conter os relatórios do moodle dentro do onedrive. No caso: /Relatorios_Moodle_1-3
          Nome do Arquivo: FuncaoDataFormatada_NomeVinculado
            > FuncaoDataFormatada = Este código serve para formatar o nome do arquivo a ser gerado com a forma de data e hora. Clicar em inserir função e utilizar o seguinte código: formatDateTime(triggerOutputs()?['body/receivedDateTime'], 'dd-MM-yyyy_HH-mm')
            > NomeVinculado = Selecionar vinculo para o link "Criar Arquivo - Outlook For Business" desta forma irá usar de referencia o nome do relatório recebido no outlook
          Conteúdo do Arquivo: ContentBytes (do arquivo em anexo do Outlook)
            > 
      - Criar Tabela - Excel Business
      - Listar Linhas Presentes em uma Tabela - Excel Business
      - Script SubstituirVazio - Excel Business
      - Script CriarTabelasDinamicas - Excel Business
      - Obter Conteúdo de Arquivo - OneDrive for Business
      - Condição SE: Nome do Arquivo Contém ID_1
            - Se Verdadeiro: Enviar e-mail com ID=1
                Utilizar Enviar um e-mail - Outlook Business
            - Se Falso
                - Consição SE: Nome do Arquivo Contém ID_2
                      - Se Verdadeiro: Enviar e-mail com ID=2
                          Utilizar Enviar um e-mail - Outlook Business
                      - Se Falso: Enviar e-mail com ID=3
                          Utilizar Enviar um e-mail - Outlook Business

        

 

