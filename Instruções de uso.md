## **Guia de Utilização: Automatizador de Extratos de Cartão de Crédito**



Este guia descreve como usar o aplicativo executável para automatizar a inserção de dados de extratos de cartão de crédito em sua planilha de contas a pagar.



###### **1. Requisitos e Preparação**

Antes de usar o programa, garanta que a planilha base esteja no local correto. O executável foi configurado para buscar a planilha no seguinte caminho:



###### **Caminho da Planilha Base:**

C:\\Bitrix24\\Aurora Hotel\\Automação\\Omie\_Contas\_Pagar\_v1\_1\_5.xlsx



* **Verifique o Caminho:** Certifique-se de que o arquivo Omie\_Contas\_Pagar\_v1\_1\_5.xlsx exista exatamente nessa pasta. Se ele não for encontrado, o programa exibirá um erro.



* **Novo Arquivo Gerado:** Após o processamento, o programa criará uma cópia da planilha com os dados do extrato na sua Área de Trabalho. A planilha original na pasta Automação não será alterada.



###### **2. Interface do Programa**

Ao abrir o executável, você verá uma janela simples com os seguintes campos:



**Selecione o banco:** Um menu suspenso para escolher o banco do extrato. As opções disponíveis são Sicoob, Banco do Brasil, Caixa e Itaú.



**Arquivo de Extrato:** Um campo de texto onde você deve inserir o caminho do arquivo do extrato do cartão de crédito. Você pode clicar no botão "Procurar..." para navegar e selecionar o arquivo no seu computador. Os formatos de arquivo aceitos são .ofx (para Sicoob) e .pdf (para os demais bancos).



**Conta Corrente:** Digite o nome da sua conta corrente associada a este cartão. O valor será inserido na coluna E da planilha.



**Data de Vencimento:** Insira a data de vencimento da fatura no formato DD/MM/AAAA. O valor será inserido na coluna K da planilha.



###### **3. Processando os Dados**

Preencha os campos: Siga as instruções da seção anterior para selecionar o banco, o arquivo de extrato, a conta corrente e a data de vencimento.



Clique em "Processar": Após preencher todos os campos, clique no botão "Processar".



Aguarde o Processamento: O programa começará a ler o extrato, extrair as transações e inserir os dados na nova planilha.



**Verifique o Resultado:**



Se o processamento for bem-sucedido, uma mensagem de sucesso será exibida, informando quantas transações foram inseridas e o nome do novo arquivo gerado.



Em caso de erro (por exemplo, se o arquivo base não for encontrado ou o formato do extrato for incompatível), uma mensagem de erro será exibida.



###### **4. Funcionamento Interno e Mapeamento de Colunas**

O programa foi desenvolvido para preencher a planilha com os dados do extrato de forma automatizada. Os dados extraídos são mapeados para as seguintes colunas na planilha:



Coluna C: Fornecedor (descrição da compra)



Coluna D: Categoria (sempre "Cartão de Credito")



Coluna E: Conta Corrente (valor informado por você)



Coluna F: Valor da Conta (valor da transação)



Coluna J: Data de Registro (data da compra)



Coluna K: Data de Vencimento (valor informado por você)



A inserção de dados começa a partir da linha 6, e novas linhas são adicionadas para cada transação encontrada no extrato. A formatação original da planilha base é preservada.

