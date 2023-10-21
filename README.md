# Automacao_compras
 
Automacao de compras com Selenium e Chrome

Dados fornecidos via Excel, para extração e tratamento (via pandas) e utilização para pesquisar nos sites de compra:
GoogleShopping
Buscapé

Após ler e tratar os dados da planilha, inicia-se o loop de busca de todos os produtos e coleta das informações necessárias que são : Preço, Nome do anuncio e link de compra.

Logo após tratar o preço (dentro do range solicitado pelo cliente na planilha)
Anúncio sem os termos banidos pelo cliente(tambem informado na planilha)

Criar uma tabela (via Tabulate) e encaminhar via e-mail automaticamente.

Neste caso foi selecionado a opção de enviar um email de cada produto, com ambos market place.
