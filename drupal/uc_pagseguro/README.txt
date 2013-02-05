;$Id: README.txt,v 1.5 2010/11/05 16:09:00 wundo Exp $

UOL PagSeguro Payment Module for Drupal UberCart

OBS: Instructions in Portuguese-BR, since this module and this payment gateway 
just works for Brazil currently.

Sumário
1. Instalação
2. Configuração


1. INSTALAÇÃO
Copie a pasta "uc_pagseguro" no diretório de módulos do Drupal (normalmente em 
/sites/all/modules). Ative-o na URL /admin/build/modules, em "Ubercart - 
payment", item UOL PagSeguro.

Para maior comodidade, ative o "Retorno Automático de Dados" do PagSeguro. 
Veja procedimento em "Token gerado pelo PagSeguro" abaixo. 



2. CONFIGURAÇÃO
Vá para a URL admin/store/settings/checkout -> Edit -> Address Fields. Configure
o título dos campos da seguinte maneira:

First Name -> Nome
Last Name -> Sobrenome
Phone Number -> Telefone para Contato
Company -> Empresa
Street Address 1 -> Endereço e Número
Street Address 2 -> Complemento e Bairro
City -> Cidade
State/Province -> Estado
Country -> País
Postal Code -> CEP
Address -> Em branco
Street Address -> Em branco

Não é recomendável mudar a ordem ou título dos campos pois este módulo cria 
instruções próprias para campos específicos. É recomendável que todos os 
campos sejam marcados como obrigatórios, menos o de Empresa. Todas estas 
informações serão repasadas para o PagSeguro quando o processo de pagamento for
iniciado e o ideal seria que o usuário tenha que preencher dados no PagSeguro o 
menos possível.

Opcionalmente para melhor apresentação, considerando que o PagSeguro só trabalha
para o Brasil no momento, não deixe de importar e configurar o país na URL 
admin/store/settings/countries, desabilitando as opções padrão (Estados Unidos e
Canadá).

Vá para a URL /admin/store/settings/payment -> Edit -> Payment Methods. Ative o 
PagSeguro e coloque o peso de acordo com a ordem em que deseja que apareça para 
o usuário.

Em "PagSeguro settings":

- Email cadastrado no PagSeguro
O e-mail utilizado no sistema PagSeguro para recebimento de valores.


- Token gerado pelo PagSeguro
Para acertar este item, primeiro acesse no sistema do PagSeguro a página de 
configuração "Meus Dados" e de lá, o item "Retorno Automático de Dados"
(https://pagseguro.uol.com.br/Security/ConfiguracoesWeb/RetornoAutomatico.aspx).
Ative a URL de retorno e neste campo coloque o endereço do seu site acrescido de
"/cart/pagseguro/complete". Exemplo:

http://www.meusite.com.br/cart/pagseguro/complete

A seguir, clique no botão "Gerar" do item 2 ("Token para Conferência de 
Segurança"). Ele irá gerar uma sequência alfa-numérica de 32 caracteres ao lado 
do botão clicado. Após aparecer o aviso de "Dados salvos com sucesso", copie 
esta sequência no campo disponibilizado para a configuração de seu Drupal 
Ubercart, neste campo.

A geração e utilização do Token é opcional, mas deixar de acertar o Retorno
Automático de Dados no PagSeguro implica em ter de atualizar os passos de 
acompanhamento do pedido manualmente no Ubercart. 


- Cálculo de frete pelo Ubercart
Para utilizar esta opção, é preciso realizar uma configuração antes.
No PagSeguro, vá em "Meus Dados" e de lá, para o item "Preferências Web e Frete"
(https://pagseguro.uol.com.br/Security/WebPagamentos/ConfigWebPagto.aspx). Em 
"Definição do cálculo do frete", escolha "Frete Fixo sem Desconto".

OBS: Não é aconselhável utilizar a opção "Frete Fixo com Desconto". Por 
incompatibilidades entre o Ubercart e o PagSeguro, o valor total do frete 
calculado pelo Ubercart é dividido igualmente entre os produtos comprados pelo
usuário e repassado ao PagSeguro, não importando o peso e nem a quantidade, o 
que pode gerar desacordos no cálculo do novo frete no PagSeguro.


- Texto do botão de revisão do pedido
Insira um texto para colocar no botão de enviar pedido, na tela de revisão do 
pedido (URL cart/checkout/review) antes de o usuário ser desviado para o 
pagamento no PagSeguro.


