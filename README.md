# Projetos_VBA


This is a automation project using VBA. The idea behind this project was to save time and increase production, bellow there is a brief explanation about the project,
i divided in 4 steps for a better understanding.

1° Automatização da planilha de controle de navios\
Os 4 arquivos abaixo pertencem ao mesmo projeto.\
EnviarEmail\
ExportarArquivo\
LimparBase\
RodarRelatorio\
1° clear spreadsheets data - The file "LimparBase" its a simple macro that cleans the currently spreadsheets before receiving the new data.\
2° Run report - After we have insert the new data in the spreadsheet we use the the  macro named "RodarRelatorio" to run the report and to update or main control sheet.\
3° Export file - The macro named "ExportarArquivo" has the objective to extract information from our main control sheet (that was updated with the previous macro) and automatically update another sheet, where we can better analyze the data.\
4° EnviarEmail - Essa macro tem como funcionalidade o envio automatico de email, selecionando casos pendentes de acordo com determinados filtros pré-estabelecidos.\
Todo esse processo era feito de forma manual, onde um analista dedicava cerca de uma hora e meia por dia para executar tais tarefas. Com a implementação das macros
o tempo para realizar a mesma atividade foi reduzido para poucos minutos.
