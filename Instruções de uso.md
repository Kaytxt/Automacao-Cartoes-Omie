##### **Guia de Utilização: Automatizador de Extratos de Cartão de Crédito**



Este guia descreve como usar o seu aplicativo para automatizar a inserção de dados de extratos de cartão de crédito em sua planilha de contas a pagar, com a nova funcionalidade de conciliação com a API Omie.



###### **1. Requisitos e Preparação**

Antes de usar o programa, garanta que a planilha base e os arquivos de credenciais estejam no local correto.



**Caminho da Planilha Base:**

O programa está configurado para buscar a planilha no seguinte caminho:

C:\\Bitrix24\\Aurora Hotel\\Automação\\Omie\_Contas\_Pagar\_v1\_1\_5.xlsx

Verifique se o arquivo existe exatamente nesta pasta. Se não for encontrado, o programa exibirá um erro.



**Arquivo de Credenciais (.json):**

O programa busca as chaves de acesso da Omie na pasta credenciais. Certifique-se de que o arquivo JSON do cliente selecionado (aurora\_hotel.json, elias\_carnes.json, etc.) esteja presente e com o formato correto.



**Novo Arquivo Gerado:**

Após o processamento, uma nova planilha será criada na sua Área de Trabalho com os dados do extrato já conciliados. A planilha original não será alterada.



###### **2. Interface do Programa (Tela Principal)**

Ao abrir o aplicativo, você verá uma tela inicial organizada com os seguintes campos:



Selecione o cliente: Um menu para escolher o cliente que será processado. A seleção do cliente é essencial para carregar as credenciais de API corretas.



Selecione o banco: Escolha o banco do extrato. Os formatos aceitos são .ofx (para Sicoob) e .pdf (para os demais bancos).



Arquivo de Extrato: Clique em "Procurar..." para selecionar o arquivo do extrato em seu computador.



Conta Corrente: Digite o nome da conta corrente para que o valor seja inserido na coluna E da planilha.



Data de Vencimento: Insira a data de vencimento da fatura no formato DD/MM/AAAA.



Status do Processamento: A área de texto na parte inferior mostrará o status e as mensagens do programa.



###### **3. Processo de Conciliação e Salvamento**

O fluxo do programa agora tem uma etapa intermediária para garantir a precisão dos dados.



Preencha os Campos: Na tela principal, preencha todos os campos e clique em "Processar".



Início da Conciliação: O programa irá:



Extrair os dados do extrato.



Conectar-se à API da Omie para buscar a lista de fornecedores e categorias.



Tentar fazer uma conciliação automática, comparando as descrições do extrato com os nomes dos fornecedores da Omie.



Tela de Conciliação Manual: Se o programa não conseguir conciliar 100% dos itens automaticamente, uma nova janela de "Conciliação Manual" será aberta.



Tabela do Extrato: À esquerda, você verá uma tabela com os itens que precisam de sua atenção. As colunas "Fornecedor Omie" e "Categoria Omie" estarão em branco para esses itens.



Abas de Conciliação: À direita, você encontrará duas abas:



Fornecedores: Uma lista completa de todos os clientes da Omie para o cliente selecionado. Use o campo de "Pesquisar" ou role a lista para encontrar e selecionar o fornecedor correto.



Categorias: Uma lista de todas as categorias cadastradas na Omie. A categoria padrão "Cartão de Credito" estará disponível para seleção.



Como Usar:



Dê um clique duplo na linha da tabela do extrato que você quer editar.



O programa irá selecionar automaticamente a aba correspondente para você.



Dê um clique duplo no nome do fornecedor ou categoria na lista à direita para preencher a coluna correta na tabela do extrato.



Ao terminar, clique em "Salvar e Fechar".



Geração da Planilha: O programa irá gerar a nova planilha na sua Área de Trabalho com todos os dados preenchidos, incluindo as suas correções manuais.



###### **4. Mapeamento de Colunas**

O programa preenche a planilha com os dados do extrato da seguinte forma:



Coluna C: Fornecedor (Nome conciliado da Omie ou a descrição original)



Coluna D: Categoria (Valor selecionado na tela de conciliação manual)



Coluna E: Conta Corrente (Valor informado por você)



Coluna F: Valor da Conta (Valor da transação)



Coluna J: Data de Registro (Data da compra)



Coluna K: Data de Vencimento (Valor informado por você)

