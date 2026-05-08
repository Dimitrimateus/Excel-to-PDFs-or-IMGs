# Excel-to-PDFs-or-IMGs
Esta automacao foi projetada para segmentar bases de dados e gerar arquivos individuais (PDF ou PNG) com base em filtros de uma coluna. O script identifica automaticamente os itens da coluna escolhida e exporta um relatorio para cada um.

Customizacao e Configuracoes Rapidas
O codigo foi estruturado para ser facilmente editavel. No topo do modulo, voce encontrara a secao denominada CONFIGURACOES RAPIDAS. Alterando os valores nesta parte, voce modifica o comportamento de todo o programa sem precisar mexer na logica complexa:

DATA_SHEET: Altere o texto entre aspas para o nome exato da aba onde estao seus dados (Ex: "Dados_Vendas").

ZONE_COLUMN: Altere o numero para indicar qual coluna deve ser filtrada. (Ex: Use 1 para Coluna A, 2 para Coluna B, etc).

HEADER_ROWS: Define quais linhas serao repetidas no topo de cada pagina do PDF (Ex: "$1:$2" se o seu cabecalho ocupar duas linhas).

LIMITE_LINHAS: Define o ponto de corte. Se o filtro resultar em menos linhas que este numero, o Excel gera uma Imagem (PNG). Se for maior, gera um PDF.

Funcionalidades
Exportacao Inteligente: Seleciona o formato (Imagem ou PDF) com base no tamanho do relatorio.

Algoritmo de Captura: Evita falhas visuais e celulas pretas ao gerar imagens, garantindo fidelidade ao que e visto no Excel.

Tratamento de Nomes: Remove automaticamente caracteres invalidos do Windows para evitar erros de salvamento.

Como Utilizar
Abra o Editor VBA (ALT + F11) e cole o codigo em um Modulo.

Ajuste as Configuracoes Rapidas no topo do codigo conforme a sua necessidade.

Retorne ao Excel e execute a macro (ALT + F8).

Insira um prefixo para os arquivos quando solicitado (Ex: "Vendas_Junho").

Selecione a pasta onde os arquivos serao armazenados.

Requisitos
Sistema Operacional Windows.

Microsoft Excel configurado para permitir a execucao de Macros (.xlsm ou .xlsb).

Notas Tecnicas
A ferramenta utiliza um objeto de dicionario para garantir que cada item da coluna de filtro seja processado apenas uma vez, otimizando o tempo de execucao em bases de dados com milhares de linhas. A geracao de imagem e feita atraves de um grafico temporario, o que garante que a resolucao nao dependa do zoom da tela do usuario.

Licenca
Este projeto esta sob a licenca MIT. Sinta-se livre para utilizar e modificar conforme sua necessidade profissional.
