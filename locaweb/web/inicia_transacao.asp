<%
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Loja Exemplo Locaweb
' Versão: 6.5
' Data: 12/09/06
' Arquivo: inicia_transacao.asp
' Versão do arquivo: 0.0
' Data da ultima atualização: 21/10/08
'
'-----------------------------------------------------------------------------
' Licença Código Livre: http://comercio.Locaweb.com.br/gpl/gpl.txt
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-

rodape = "no"
navegacaocompra = "fim"
page = "inicia_transacao"
passo=3
%>
<script language="JavaScript">
    // Inibe o botão voltar
    history.go(+1)
</script>
<!--#INCLUDE FILE="funcoes/funcoes_grava_transacao.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_cartao.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_usuario.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_endereco.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_uteis.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_mail.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_md5.asp"-->
<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
        <td colspan="3" valign="top" height="30"><!--#INCLUDE FILE="cabecalho.asp" --></td>
    </tr>
    <tr>
        <%If navegacaocompra = "fim" Then%>
        <td valign="top" height="10%" width="10%" class="TBLlatesquerda"><!--#INCLUDE FILE="menu_poscarrinho.asp" --></td>
        <%Else%>
        <td valign="top" height="10%" width="10%" class="TBLlatesquerda"><!--#INCLUDE FILE="menu.asp" --></td>
        <%End if%>
        <td valign="top" height="95%">
            <%
            permissao="read"
            page = "iniciatransacao"
            readonly = "readonly"

            If Request.form("acao") = "" Then

                ' Verifica se o pedido não está sendo refeito com outra forma de pagamento
                If Session("novoPedido") = "yes" Then

                    ' Resgata a forma de pagamento utilizada na última tentativa
                    forma_pagto = Pega_DadoBanco("Pedidos","forma_pagamento","sessionID","'"&Session("id_transacao")&"'")

                    ' Zera as sessões
                    Session("novoPedido") = Empty
                    Session("resgistroPedido") = Empty
                    Session("resgistroPedidoItem") = Empty
                    Session("registrado") = Empty
                    Session("codigo_pedido") = Empty

                End If

                ' Verifica se o número de parcelas está puro
                If Instr(Request("dados_adicionais"),"|") = 0 And Len(Request("dados_adicionais")) <> 0 Then
                    ' Atualiza o número de parcelas no XML do pedido
                    Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","num_parcelas",Request("dados_adicionais"))
                End If

                'Resgata a forma de pagamento escolhida
                sForma_pagamento = pegaValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","forma_pagamento")
                
                ' Verifica a forma de pagamento escolhida
                If sForma_pagamento = "Amex" Or sForma_pagamento = "Mastercard" Or sForma_pagamento = "Diners" Or sForma_pagamento = "Visa" Or sForma_pagamento = "Paggo" Then
                    
                    ' Pega a atual taxa do pedido, caso exista.
                    sTaxaPedido = Trim(pegaValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","tipo_taxa_adicional"))

                    ' Verifica se o parcelado está disponível para essa forma de pagamento e se não há taxas já aplicadas ao pedido
                    If pegaValorAtrib(Application("XMLMeiosPagamentos"),"configuracao/pagto[@nome_pagto='"&sForma_pagamento&"']","permite_parcelamento") = "sim" And sTaxaPedido = "" Then
                        ' Verifica o tipo de parcelado configurado para essa forma de pagamento
                        If pegaValorAtrib(Application("XMLMeiosPagamentos"),"configuracao/pagto[@nome_pagto='"&sForma_pagamento&"']","juros") = "lojista" Then ' Juros do Lojista
                        
                            ' Formata o número de parcelas(retira o zero à esquerda)
                            If Left(Request("dados_adicionais"),1) = "0" Then
                                nNumParcela = Right(Request("dados_adicionais"),1)
                            Else
                                nNumParcela = Request("dados_adicionais")
                            End If

                            ' Resgata o tipo e valor da taxa da opção de parcelamento escolhida
                            sTipotaxa = pegaValorAtrib(Application("XMLMeiosPagamentos"),"configuracao/pagto[@nome_pagto='"&sForma_pagamento&"']","parc"&nNumParcela)
                            ' Resgata o valor total do pedido
                            currTotalPedido = pegaValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","valor_total")
                            ' Verifica o tipo de taxa utilizado no parcelamento escolhido
                            If sTipotaxa = "Desconto" Then ' Desconto
                                nValortaxa = pegaValorAtrib(Application("XMLMeiosPagamentos"),"configuracao/pagto[@nome_pagto='"&sForma_pagamento&"']","taxa_desconto")
                                currValorCalc = calculaValorTaxa(currTotalPedido,nValortaxa,"Desconto")
                            ElseIf sTipotaxa = "Com juros" Then ' Com Juros
                                nValortaxa = pegaValorAtrib(Application("XMLMeiosPagamentos"),"configuracao/pagto[@nome_pagto='"&sForma_pagamento&"']","taxa_juros")
                                currValorCalc = calculaValorTaxa(currTotalPedido,nValortaxa,"Acrescimo")
                            Else ' Sem Juros
                                nValortaxa = 0
                            End If

                            ' Atualiza o valor total da transação no XML do pedido, caso necessário
                            If currValorCalc <> "" Then
                                Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","valor_total",currValorCalc)
                            End If
                            ' Atualiza no XML do pedido o tipo e taxa do parcelamento
                            Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","tipo_taxa_adicional",sTipotaxa)
                            Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","taxa_adicional",nValortaxa)

                        Else ' Juros do Emissor

                            ' Atualiza no XML do pedido o tipo do parcelamento
                            Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","tipo_taxa_adicional","Juros do Emissor")

                        End If
                    End If
                End If

                'Grava os dados do pedido no banco de dados
                Call CarregaGrava_dados_pedido(session("id_transacao"), objXML, objRoot,VarAdicional)

            End If

            Set RS_pega_codigo = CreateObject("ADODB.Recordset")
            Set RS_pega_codigo.ActiveConnection = Conexao
            RS_pega_codigo.CursorLocation = 3
            RS_pega_codigo.CursorType = 0
            RS_pega_codigo.LockType =  1

            If Application("TipoBanco") = "mysql" Then
				QueryPegaCodPedido = "SELECT codigo_pedido FROM Pedidos WHERE sessionID = '" & session("id_transacao") & "' ORDER BY codigo_pedido DESC LIMIT 0,1"
			Else
				QueryPegaCodPedido = "SELECT TOP 1 codigo_pedido FROM Pedidos WHERE sessionID = '" & session("id_transacao") & "' ORDER BY codigo_pedido DESC"
			End If

			RS_pega_codigo.Open QueryPegaCodPedido, Conexao

            Session("codigo_pedido") = RS_pega_codigo("codigo_pedido")

            RS_pega_codigo.Close
            Set RS_pega_codigo = Nothing
            %>
            <table width="100%" border=0 align="center" cellpadding="0" cellspacing="10">
                <tr>
                    <td align="center" height="18" valign="middle"><!--#INCLUDE FILE="barra_passoapasso.asp" --></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td align="center"  valign="top">
                        <table width="100%" border=0 align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                            <%
                            'Abre conexao ao XML dos meios de pagto.
                            Call abre_ArquivoXML(Application("XMLMeiosPagamentos"),FctobjXML,FctobjRoot) 
                            Set configuracao = FctobjRoot.selectSingleNode("configuracao/pagto[@nome_pagto='"&Session("forma_pagamento")&"']")

                            'Abre conexao ao XML do pedido.
                            Call abre_xmlpedido(session("id_transacao"), objXML, objRoot) 
                            Set raiz_dados_pedido = objRoot.selectSingleNode("dados_pedido[@id_transacao="&session("id_transacao")&"]")
                            
                            'Verifica a forma de pagamento escolhida
                            If Session("forma_pagamento") = "Visa" Then
                                'Verifica a solução definida para uso da transação (VBV ou MOSET).
                                If configuracao.getAttribute("VisanetTipo") = "MOSET" Then
                                    varTXTtransacao = "Informe os dados de seu cartão Visa para prosseguir com a compra:"
                                Else
                                    varTXTtransacao = Application("InitTxtAvisoAguarde")
                                End If
                            Else
                                varTXTtransacao = Application("InitTxtAvisoAguarde")
                            End If
                            %>
                            <tr>
                                <td><B><%= varTXTtransacao%></B></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                                <td  height="40" align="center" bgcolor="#FFFFFF"> 
                                    <%
                                    ' Opção de pagamento Boleto Bancário / Depósito / CobreBem / Unibanco
                                    If Session("forma_pagamento") = "Boleto" Or Session("forma_pagamento") = "Deposito" Or Session("forma_pagamento") = "CobreBem" Or Session("forma_pagamento") = "Unibanco" Then

                                        Response.redirect "recibo.asp?lang=" & varLang

                                    ' Opção de pagamento Visa
                                    ElseIf Session("forma_pagamento") = "Visa" Or Session("forma_pagamento") = "VisaElectron" Then

                                        If Session("forma_pagamento") = "Visa" Then
                                            If Request("dados_adicionais") = "01" Or Request("dados_adicionais") = "" Then
                                                codigo_pagto = "10"
                                                session("codigo_pagamento") = codigo_pagto & "01"
                                            Else
                                                ' Verifica o tipo de juros configurado
                                                If configuracao.getAttribute("juros") = "emissor" Then ' Juros do emissor
                                                    codigo_pagto = "30"
                                                Else ' Juros do lojista
                                                    codigo_pagto = "20"
                                                End If
                                                session("codigo_pagamento") = codigo_pagto & Request("dados_adicionais")
                                            End If
                                        ElseIf Session("forma_pagamento") = "VisaElectron" Then
                                            codigo_pagto = configuracao.getAttribute("VisanetCodPagamento")
                                            session("codigo_pagamento") = codigo_pagto & "01"
                                        End If

                                        total = Replace(Replace(raiz_dados_pedido.getAttribute("valor_total"),",",""),".","")

                                        'Verifica a solução definida para uso da transação (VBV ou MOSET).
                                        If configuracao.getAttribute("modulo") = "VISAVBV" Then
                                            
                                            'Gerar TID -
                                            'O primiero parâmetro se refere ao númeto de afiliação da loja junto à Visanet
                                            'O segundo parâmetro se refere ao tipo de pagamento e o prazo
                                            TID = GerarTid_VBV(configuracao.getAttribute("VisanetID"),session("codigo_pagamento"))

                                            Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")

                                            Set raiz_dados_pedido = objRoot.selectSingleNode("dados_pedido[@id_transacao="&session("id_transacao")&"]")
                                            Set states = objXML.getElementsByTagName("dados_pedido[@id_transacao="&session("id_transacao")&"]/produto") 
                                            n_states = states.length
                                            for i = 0 to (n_states - 1)
                                                Set pagto = states.item(i)
                                                
                                                vSplitCodPed = Split(pagto.getAttribute("codigo_produto"),"_")

                                                varArrayCodPed = varArrayCodPed & vSplitCodPed(0) & ","

                                                Set pagto = Nothing
                                            next

                                            raiz_dados_pedido.setAttribute "codigo_pedido",Session("codigo_pedido")   
                                            'Essa operação é necessária para salvar o c´dogio do pedido no arquivo XML desta transação.
                                            objXML.save(Application("DiretorioPedidos")&session("id_transacao")&".xml")


                                            Set states = Nothing
                                            
                                            'Grava os dados iniciais da transação no banco de dados
                                            Call GravaTransacaoInicialVisa(Session("codigo_pedido"),TID,total,TIPOCARTAO,Request("dados_adicionais"),codigo_pagto,"VBV",configuracao.getAttribute("ambiente"),configuracao.getAttribute("IdentificacaoLocaweb"))

                                            If Trim(Session("razaosocial_cobranca")) <> "" Then
                                                varOrder = Session("razaosocial_cobranca") & "|" & Session("cnpj_cobranca") & "|" & Session("inscricaoestadual_cobranca") & "|"
                                            Else
                                                varOrder = Session("nome_cobranca") & "|" & Session("rg_cobranca") & "|" & Session("cpf_cobranca") & "|"
                                            End If
                                            
                                            varOrder = varOrder & Session("logradouro_cobranca") & "," & Session("numero_cobranca") & "-" & Session("complemento_cobranca") & "|" & Session("cep_cobranca") & "|" & Session("cidade_cobranca") & "|" & Session("estado_cobranca") & "|" & Session("pais_cobranca") & "|CODPRO:" & mid(varArrayCodPed,1,LEN(varArrayCodPed)-1)
                                            varOrder = mid(varOrder,1,1024)

                                            'Monta os dados postados à operadora
                                            valores = "PosicaoDadosVisanet=0"
                                            
                                            valores = valores & "&identificacao=" & configuracao.getAttribute("IdentificacaoLocaweb")
                                            valores = valores & "&modulo=" & configuracao.getAttribute("modulo")
                                            valores = valores & "&ambiente=" & configuracao.getAttribute("ambiente")

                                            valores = valores & "&visa_antipopup=" & configuracao.getAttribute("VisaNetAntiPopup") 
                                            valores = valores & "&tid=" & TID 
                                            valores = valores & "&price=" & total 
                                            valores = valores & "&language=" & Left(session("requestIdioma"),2)
                                            valores = valores & "&order=" & varOrder
                                            valores = valores & "&orderid=" & Session("codigo_pedido") 
                                            valores = valores & "&free=" & session("requestIdioma")
                                            valores = valores & "&damount=R$" & raiz_dados_pedido.getAttribute("valor_total") 
                                            valores = valores & "&authenttype=" & configuracao.getAttribute("VisanetAuthentType")
                                            
                                            objSrvHTTP.open "POST", Application("URLRecebeDadosVisaVBV"), False

                                            objSrvHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

                                            objSrvHTTP.send valores

                                            If objSrvHTTP.Status = 200 Then
                                                response.write objSrvHTTP.responseText
                                            Else
                                                Response.write "Error: (" & objSrvHTTP.Status & ") " & objSrvHTTP.statusText
                                            End If

                                            Set objSrvHTTP = Nothing

                                        Else
                                            
											' Verifica se foi postado os dados do cartão
											If Request("dados_cartao") <> "" Then

												'Gerar TID -
												'O primiero parâmetro se refere ao númeto de afiliação da loja junto à Visanet
												'O segundo parâmetro se refere ao tipo de pagamento e o prazo
												TID = GerarTid_MOSET(configuracao.getAttribute("VisanetID"),session("codigo_pagamento"))

												Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")

												Set raiz_dados_pedido = objRoot.selectSingleNode("dados_pedido[@id_transacao="&session("id_transacao")&"]")
												Set states = objXML.getElementsByTagName("dados_pedido[@id_transacao="&session("id_transacao")&"]/produto") 
												n_states = states.length
												for i = 0 to (n_states - 1)
													Set pagto = states.item(i)
													
													vSplitCodPed = Split(pagto.getAttribute("codigo_produto"),"_")

													varArrayCodPed = varArrayCodPed & vSplitCodPed(0) & ","

													Set pagto = Nothing
												next

												Set states = Nothing
												
												If Request("dados_adicionais") = "" Then
													NUMPARCELAS = "1"
												Else
													NUMPARCELAS = Request("dados_adicionais")
												End If

												'Grava os dados iniciais da transação no banco de dados
												Call GravaTransacaoInicialVisa(Session("codigo_pedido"),TID,total,TIPOCARTAO,NUMPARCELAS,codigo_pagto,"MOSET",configuracao.getAttribute("ambiente"),configuracao.getAttribute("IdentificacaoLocaweb"))

												If Trim(Session("razaosocial_cobranca")) <> "" Then
													varOrder = Session("razaosocial_cobranca") & "|" & Session("cnpj_cobranca") & "|" & Session("inscricaoestadual_cobranca") & "|"
												Else
													varOrder = Session("nome_cobranca") & "|" & Session("rg_cobranca") & "|" & Session("cpf_cobranca") & "|"
												End If
												
												varOrder = varOrder & Session("logradouro_cobranca") & "," & Session("numero_cobranca") & "-" & Session("complemento_cobranca") & "|" & Session("cep_cobranca") & "|" & Session("cidade_cobranca") & "|" & Session("estado_cobranca") & "|" & Session("pais_cobranca") & "|CODPRO:" & Mid(varArrayCodPed,1,Len(varArrayCodPed)-1)
												varOrder = Mid(varOrder,1,1024)

												dadosCartao = Split(Request("dados_cartao"),"|")
												numCartao = Trim(dadosCartao(0))
												cvvCartao = Trim(dadosCartao(1))
												expCartao = Trim(dadosCartao(2))
												
												' Parâmetros obrigatórios
												parametros = parametros & "identificacao=" & configuracao.getAttribute("IdentificacaoLocaweb")
												parametros = parametros & "&modulo=" & configuracao.getAttribute("modulo")
												parametros = parametros & "&operacao=Pagamento"
												parametros = parametros & "&ambiente=" & configuracao.getAttribute("ambiente")
												parametros = parametros & "&tid=" & TID
												parametros = parametros & "&ccn=" & numCartao
												parametros = parametros & "&cvv2=" & cvvCartao
												parametros = parametros & "&exp=" & expCartao
												parametros = parametros & "&price=" & total
												parametros = parametros & "&orderid=" & Session("codigo_pedido")
												parametros = parametros & "&order=" & varOrder
												parametros = parametros & "&free=" & session("requestIdioma")

												' URL de acesso ao Gateway Locaweb
												urlLocaWebCE = Application("URLRecebeDadosVisaMOSET")

												' Instancia o objeto HttpRequest. 
												Set objSrvHTTP = Server.CreateObject("MSXML2.XMLHTTP.3.0") 

												' Informe o método e a URL a ser capturada 
												objSrvHTTP.open "POST", urlLocawebCE, false 

												' Com o método setRequestHeader informamos o cabeçalho HTTP 
												objSrvHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 

												' O método Send envia a solicitação HTTP e exibe o conteúdo da página 
												objSrvHTTP.Send(parametros)

												' Verificando se a busca foi bem sucedida 
												If objSrvHTTP.statusText = "OK" Then
													Set objXmlDoc = Server.CreateObject("MSXML2.DOMDocument")
													objXmlDoc.loadXML(BinaryToString(objSrvHTTP.responseBody))
															
													' Verificando se o retorno foi bem sucedido 
													If TypeName(objXmlDoc) = "DOMDocument" Then 

														' Recupera os parâmetros e dá valores a variáveis para serem usadas na página 
														erro = objXmlDoc.selectSingleNode("//LocaWebCE//erro").text
														erro = DecodeUTF8(erro)
														origemErro = objXmlDoc.selectSingleNode("//LocaWebCE//origemErro").text

														' Se não ocorreu erro recupera parâmetros        
														If erro = "" Then
															' Resgata os dados de retorno da transação
															idReqLocaWeb = objXmlDoc.selectSingleNode("//LocaWebCE//idReqLocaWeb").text
															operacao = objXmlDoc.selectSingleNode("//LocaWebCE//operacao").text
															RETlr = objXmlDoc.selectSingleNode("//LocaWebCE//LR").text
															RETars = objXmlDoc.selectSingleNode("//LocaWebCE//ARS").text
															RETars = DecodeUTF8(RETars)
															RETtid = objXmlDoc.selectSingleNode("//LocaWebCE//TID").text
															RETorderid = objXmlDoc.selectSingleNode("//LocaWebCE//ORDERID").text
															RETprice = objXmlDoc.selectSingleNode("//LocaWebCE//PRICE").text
															RETarp = objXmlDoc.selectSingleNode("//LocaWebCE//ARP").text
															RETbank = objXmlDoc.selectSingleNode("//LocaWebCE//BANK").text
															RETfree = objXmlDoc.selectSingleNode("//LocaWebCE//FREE").text
															RETfree = DecodeUTF8(RETfree)

															'Grava os dados da transação Visanet no banco de dados
															Call GravaTransacaoFinalVisa(Session("codigo_pedido"),RETprice,RETtid,RETlr,RETarp,RETfree,PAN,RETbank,RETars,AUTHENT,idReqLocaWeb)

															Response.redirect "recibo.asp?lang=" & varLang & "&tid=" & RETtid
														Else
															' Exibe a mensagem de erro
															Response.write "Erro: " & erro & "<br>"
															Response.write "Origem do erro: " & origemErro & "<br>"
															Response.write "<br>" & Application("InitTxtInstrAlterarPedido") & " <a href=""" & Application("URLloja") & "/carrinho.asp?lang=" & varLang & "&mode=changeMeioPagto"">" & Application("InitTxtCliqueAqui") & "</a>"
														End If

												   End If
												Else
													' Exibe a mensagem de erro
													Response.write "Error: (" & objSrvHTTP.Status & ") " & objSrvHTTP.statusText & "<br>"
													Response.write "<br>" & Application("InitTxtInstrAlterarPedido") & " <a href=""" & Application("URLloja") & "/carrinho.asp?lang=" & varLang & "&mode=changeMeioPagto"">" & Application("InitTxtCliqueAqui") & "</a>"
												End If

												Set objSrvHTTP = Nothing 
												Set objXmlDoc = Nothing

											Else

												' Exibe a mensagem de erro
												Response.write "<b>Erro: " & Application("InitTxtProbTransacao") & "</b>"
												Response.write "<br>" & Application("InitTxtInstrAlterarPedido") & " <a href=""" & Application("URLloja") & "/carrinho.asp?lang=" & varLang & "&mode=changeMeioPagto"">" & Application("InitTxtCliqueAqui") & "</a>"

											End If

										End If

                                    ' Opção de pagamento Redecard (Mastercard / Diners)
                                    ElseIf Session("forma_pagamento") = "Mastercard" Or Session("forma_pagamento") = "Diners" Then

                                        If Request("dados_adicionais") = "01" Or Request("dados_adicionais") = "" Then
                                            parcelas = "00"
                                            juros = "0"
                                        Else
                                            parcelas = Request("dados_adicionais")
                                            ' Verifica a cobrança de juros
                                            If configuracao.getAttribute("juros") = "emissor" Then ' Juros do Emissor
                                                juros = "1"
                                            Else ' Juros do lojista
                                                juros = "0"
                                            End If
                                        End If

                                        total = Replace(Replace(raiz_dados_pedido.getAttribute("valor_total"),",",""),".","")


                                        'Grava os dados iniciais da transação no banco de dados
                                        Call GravaTransacaoInicialRedecard(Session("codigo_pedido"),UCASE(Session("forma_pagamento")),parcelas,juros)
                                        
                                        Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
                                        
                                        'Monta os dados postados à operadora
                                        valores = "identificacao=" & configuracao.getAttribute("IdentificacaoLocaweb")
                                        valores = valores & "&modulo=" & configuracao.getAttribute("modulo") 
                                        valores = valores & "&ambiente=" & configuracao.getAttribute("ambiente") 
                                        
                                        valores = valores & "&valor=" & total 
                                        valores = valores & "&pedido=" & Session("codigo_pedido") 
                                        valores = valores & "&pax1=" & session("requestIdioma") 

                                        If configuracao.getAttribute("RedeCardAVS") = "1" Then
                                            valores = valores & "&AVS=S" 
                                        End if

                                        valores = valores & "&parcelas=" & parcelas 
                                        valores = valores & "&juros=" & juros 
                                        valores = valores & "&BANDEIRA=" & UCASE(Session("forma_pagamento"))
                                        valores = valores & "&RedecardIdioma=" & Left(session("requestIdioma"),2)

                                        objSrvHTTP.open "POST", Application("URLRedecard"), False

                                        objSrvHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                                        objSrvHTTP.send valores

                                        If objSrvHTTP.Status = 200 Then
                                            response.write objSrvHTTP.responseText
                                        Else
                                            Response.write "Error: (" & objSrvHTTP.Status & ") " & objSrvHTTP.statusText
                                        End If

                                        Set objSrvHTTP = Nothing

                                    ' Opção de pagamento BB Office Banking
                                    ElseIf Session("forma_pagamento") = "Brasil" Then

                                        total = Replace(Replace(raiz_dados_pedido.getAttribute("valor_total"),",",""),".","")
                                        data_inicio = CorrigeData(raiz_dados_pedido.getAttribute("inicio_transacao"))
                                        vencimento = Formatdatetime(DateAdd("d", configuracao.getAttribute("BBDiasdeVencimento"), data_inicio), 2)

                                        BBTipoPagamento = configuracao.getAttribute("BBTipoPagamento")
                                        'Verifica se a variável é vazia e assume um valor padrão
                                        If Trim(BBTipoPagamento) = "" Or IsNull(BBTipoPagamento) Then
                                            BBTipoPagamento = "0"
                                        End If

                                        Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
                                        
                                        'Monta o endereço do sacado
                                        varEnd = Session("logradouro_cobranca") & ", " & Session("numero_cobranca")

                                        'Monta os dados postados à operadora
%>
                                        <FORM action="<%= Application("URLBancoBrasil") %>" method="post" name="brasil" target="vpos"> 
                                            <input type="hidden" name="idConv" value="<%= configuracao.getAttribute("BBConvenio") %>">
                                            <input type="hidden" name="valor" value="<%= total %>">
                                            <input type="hidden" name="refTran" value="<%= CodPedBrasil(Session("codigo_pedido"), configuracao.getAttribute("BBCodCobranca")) %>">
                                            <input type="hidden" name="urlRetorno" value="<%= Replace(Application("URLRecibo"),Application("SSLloja"),"") %>">
                                            <% If Trim(Session("razaosocial_cobranca")) <> "" Then %>
                                            <input type="hidden" name="nome" value="<%= Server.htmlEncode(Session("razaosocial_cobranca")) %>">
                                            <% Else %>
                                            <input type="hidden" name="nome" value="<%= Server.htmlEncode(Session("nome_cobranca")) %>">
                                            <% End If %>                                            
                                            <input type="hidden" name="endereco" value="<%= Server.htmlEncode(varEnd) %>">
                                            <input type="hidden" name="cidade" value="<%= Server.htmlEncode(Session("cidade_cobranca")) %>">
                                            <input type="hidden" name="uf" value="<%= Server.htmlEncode(Session("estado_cobranca")) %>">
                                            <input type="hidden" name="cep" value="<%= Server.htmlEncode(Session("cep_cobranca")) %>">
                                            <input type="hidden" name="dtVenc" value="<%= FormataData(vencimento) %>">
                                            <input type="hidden" name="msgLoja" value="<%= configuracao.getAttribute("BBComentario") %>">
                                            <input type="hidden" name="versao" value="002">
                                            <input type="hidden" name="moeda" value="986">
                                            <input type="hidden" name="convClasse" value="001">
                                            <input type="hidden" name="tpPagamento" value="<%= BBTipoPagamento %>">
                                            <p align="center"><a href="javascript: document.brasil.submit();"><img border=0 src="config/templates/<%=varLang%>/<%=varSkin%>/banner_bb.gif"></a></p>
                                            <p align="center">Caso a janela do banco não seja aberta automaticamente, clique na imagem acima para iniciar a transação!</p>
                                        </FORM>
                                        <SCRIPT LANGUAGE=javascript>
                                        <!--
                                              vpos=window.open('','vpos','toolbar=yes,menubar=yes,resizable=yes,status=no,scrollbars=yes,width=690,height=500');
                                              document.brasil.submit();
                                        //-->
                                        </SCRIPT>
<%
                                    ' Opção de pagamento Itau Shopline
                                    ElseIf Session("forma_pagamento") = "Itau" Then

                                        total = Replace(Replace(raiz_dados_pedido.getAttribute("valor_total"),",",""),".","")
                                        data_inicio = CorrigeData(raiz_dados_pedido.getAttribute("inicio_transacao"))
                                        vencimento = Formatdatetime(DateAdd("d", configuracao.getAttribute("ItauDiasdeVencimento"), data_inicio), 2)

                                        Set objSrvHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
                                        
                                        'Monta o endereço do sacado
                                        varEnd = Session("logradouro_cobranca") & ", " & Session("numero_cobranca")

                                        'Monta os dados postados à operadora
                                        valores = "identificacao=" & configuracao.getAttribute("IdentificacaoLocaweb")
                                        valores = valores & "&modulo=" & configuracao.getAttribute("modulo") 
                                        valores = valores & "&ambiente=" & configuracao.getAttribute("ambiente")
                                        valores = valores & "&operacao=Pagamento"

                                        valores = valores & "&pedido=" & Session("codigo_pedido")
                                        valores = valores & "&valor=" & total
                                        If Trim(Session("razaosocial_cobranca")) = "" Then
                                        valores = valores & "&nome=" & Server.URLEncode(Session("nome_cobranca"))
                                        valores = valores & "&cpfcgc=" & Session("cpf_cobranca")
                                        Else
                                        valores = valores & "&nome=" & Server.URLEncode(Session("razaosocial_cobranca"))
                                        valores = valores & "&cpfcgc=" & Session("cnpj_cobranca")
                                        End If
                                        valores = valores & "&obs=" & Server.URLEncode(configuracao.getAttribute("OBSItau"))
                                        valores = valores & "&endereco=" & Server.URLEncode(varEnd)
                                        valores = valores & "&bairro=" & Server.URLEncode(Session("bairro_cobranca"))
                                        valores = valores & "&cep=" & Session("cep_cobranca")
                                        valores = valores & "&cidade=" & Server.URLEncode(Session("cidade_cobranca"))
                                        valores = valores & "&estado=" & Session("estado_cobranca")
                                        valores = valores & "&vencimento=" & vencimento

                                        objSrvHTTP.open "POST", Application("URLItauShopline"), False

                                        objSrvHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                                        objSrvHTTP.send valores

                                        If objSrvHTTP.Status = 200 Then
                                        response.write objSrvHTTP.responseText
                                        Else
                                        Response.write "Error: (" & objSrvHTTP.Status & ") " & objSrvHTTP.statusText
                                        End If

                                        Set objSrvHTTP = Nothing

                                    ' Opção de pagamento Amex
                                    ElseIf Session("forma_pagamento") = "Amex" Then

                                        ' Formata o numero de parcelas
                                        If Request("dados_adicionais") = "01" Or Request("dados_adicionais") = "" Then
                                            parcelas = "00"
                                        Else
                                            parcelas = Request("dados_adicionais")
                                        End If

                                        ' Verifica a cobraça de juros
                                        If configuracao.getAttribute("juros") = "emissor" Then ' Juros do Emissor
                                            sTipoJuros = "PlanAmex"
                                        Else ' Juros do lojista
                                            sTipoJuros = "PlanN"
                                        End If

                                        total = Replace(Replace(raiz_dados_pedido.getAttribute("valor_total"),",",""),".","")

                                        'Grava os dados iniciais da transação no banco de dados
                                        Call GravaTransacaoInicialAmex(Session("codigo_pedido"),Request("dados_adicionais"),configuracao.getAttribute("AmexTipodeplano"))

                                    %>
                                    <pre>
                                    <form name="amex" method="POST" action="<%= Application("URLAmex") %>">
                                        <input type="hidden" name="identificacao" value="<%= configuracao.getAttribute("IdentificacaoLocaweb") %>">
                                        <input type="hidden" name="modulo" value="<%= configuracao.getAttribute("modulo") %>">
                                        <input type="hidden" name="ambiente" value="<%= configuracao.getAttribute("ambiente") %>">
                                        
                                        <input type="hidden" name="MerchTxnRef" value="<%= Session("codigo_pedido") %>">
                                        <input type="hidden" name="valor" value="<%= total %>">
                                        <input type="hidden" name="OrderInfo" value="<%= Session("codigo_pedido") %>">
                                        <input type="hidden" name="Locale" value="<%= configuracao.getAttribute("AmexLocale") %>">
                                        <% If parcelas <> "00" Then %>
                                        <input type="hidden" name="parcelas" value="<%= parcelas%>">
                                        <input type="hidden" name="PaymentPlan" value="<%= sTipoJuros %>">
                                        <% End If %>
                                        <SCRIPT LANGUAGE=javascript>
                                        <!--
                                              document.amex.submit();
                                        //-->
                                        </SCRIPT>
                                    </form>
                                    </pre>
                                    <%
                                    ' Opção de pagamento Bradesco
                                    ElseIf Session("forma_pagamento") = "Bradesco" Then

                                        'Grava os dados iniciais da transação no banco de dados
                                        Call GravaTransacaoInicialBradesco(Session("codigo_pedido"))

                                        ' Muda para a página de pagamento enviando dados da compra
                                        ' Os dados enviados são obrigatórios
                                        VarMetodoPag = Request("metodo_pag")

                                        ' Redireciona a Scopus conforme a opção de pagamento selecionada
                                        If VarMetodoPag = "TRANSFER" Then
                                            If configuracao.getAttribute("ambiente") = "TESTE" Then
                                                varURLBradescoTransfer = Application("URLTESTEBradescoTransfer")
                                            Else
                                                varURLBradescoTransfer = Application("URLPRODBradescoTransfer")
                                            End If

                                            ' transferencia entre contas
                                            Response.Redirect varURLBradescoTransfer & configuracao.getAttribute("BradescoLoja") & "/prepara_pagto.asp?merchantid=" & configuracao.getAttribute("BradescoLoja") & "&orderid=" & Session("codigo_pedido")
                                        ElseIf VarMetodoPag = "CC" Then
                                            If configuracao.getAttribute("ambiente") = "TESTE" Then
                                                varURLBradescoPagFacil = Application("URLTESTEBradescoPagFacil")
                                            Else
                                                varURLBradescoPagFacil = Application("URLPRODBradescoPagFacil")
                                            End If

                                            ' pagamento facil
                                            Response.Redirect varURLBradescoPagFacil & configuracao.getAttribute("BradescoLoja") & "/prepara_pagto.asp?merchantid=" & configuracao.getAttribute("BradescoLoja") & "&orderid=" & Session("codigo_pedido")
                                        ElseIf VarMetodoPag = "FINANCIAMENTO" Then
                                            If configuracao.getAttribute("ambiente") = "TESTE" Then
                                                varURLBradescoFinanciamento = Application("URLTESTEBradescoFinanciamento")
                                            Else
                                                varURLBradescoFinanciamento = Application("URLPRODBradescoFinanciamento")
                                            End If

                                            ' financiamento
                                            Response.Redirect varURLBradescoFinanciamento & configuracao.getAttribute("BradescoLoja") & "/prepara_pagto.asp?merchantid=" & configuracao.getAttribute("BradescoLoja") & "&orderid=" & Session("codigo_pedido")
                                        End If

                                    ' Opção de pagamento ABNCDC
                                    ElseIf Session("forma_pagamento") = "ABNCDC" Then

                                    'Resgata os valores adicionais e associa as respectivas variáveis
                                    VARdados_adicionais = split(Request("dados_adicionais"),"|")
                                    VARabn_formapgto = VARdados_adicionais(0)
                                    VARabn_garantia  = VARdados_adicionais(1)
                                    VARabn_entrada   = VARdados_adicionais(2)
                                    VARabn_vencto    = VARdados_adicionais(3)

                                    %>
                                    <form method="POST" name="frmABNCDC" action="<%= Application("URLABNCDC")%>">
                                        <input name="VAR01" type="hidden" value="<%= configuracao.getAttribute("VAR01")%>">
                                        <input name="VAR02" type="hidden" value="<%= configuracao.getAttribute("VAR02")%>">
                                        <input name="VAR03" type="hidden" value="<%= Application("URLRecibo") & "?codigo_pedido=" & Session("codigo_pedido") %>">
                                        <input name="VAR04" type="hidden" value="<%= Session("codigo_pedido")%>">
                                        <input name="VAR05" type="hidden" value="<%= VARabn_formapgto%>">
                                        <% If Session("razaosocial_cobranca") <> "" And Session("cnpj_cobranca") <> "" Then %>
                                            <input name="VAR06" type="hidden" value="J">
                                            <input name="VAR07" type="hidden" value="<%= Session("razaosocial_cobranca")%>">
                                            <input name="VAR09" type="hidden" value="<%= Session("cnpj_cobranca")%>">
                                        <% Else %>
                                            <input name="VAR06" type="hidden" value="F">
                                            <input name="VAR07" type="hidden" value="<%= Session("nome_cobranca")%>">
                                            <input name="VAR09" type="hidden" value="<%= Session("cpf_cobranca")%>">
                                        <% End If %>
                                        <input name="VAR08" type="hidden" value="<%= Session("user_id")%>">
                                        <input name="VAR21" type="hidden" value="<%= configuracao.getAttribute("VAR21")%>">
                                        <% If VARabn_vencto <> "" Then %>
                                            <input name="VAR23" type="hidden" value="<%= VARabn_vencto%>">
                                        <% End If %>
                                        <input name="VAR22" type="hidden" value="<%= raiz_dados_pedido.getAttribute("valor_total")%>">
                                        <input name="VAR26" type="hidden" value="<%= VARabn_garantia%>">
                                        <input name="VAR27" type="hidden" value="Simulação de financiamento de compra">
                                        <% If VARabn_entrada <> "" Then %>
                                            <input name="VAR28" type="hidden" value="<%= VARabn_entrada%>">
                                        <% End If %>
                                    </form>
                                    <SCRIPT LANGUAGE=javascript>
                                    <!--
                                        document.frmABNCDC.submit();
                                    //-->
                                    </SCRIPT>

									<%
                                    ' Opção de pagamento PagSeguro                                   
                                    Elseif Session("forma_pagamento") = "PagSeguro" Then
                                    'Call Session_Usuario_Transacao(Conexao,Session("codigo_pedido"))
                                    
                                    If Trim(Session("razaosocial_cobranca")) <> "" Then
                                      NomeCliente = Session("razaosocial_cobranca")
                                    Else
                                      NomeCliente = Session("nome_cobranca")                                    
                                    End IF

                                    If Trim(configuracao.getAttribute("tipo_frete")) = "sedex" Then 
                                       tipo_de_frete = "SD"
                                    Else
                                       tipo_de_frete = "EN"
                                    End If

                                    Set ProdutosItens = raiz_dados_pedido.selectNodes("produto")

                                    %>
                                    <form target="_self" action="<%= Application("URLPagSeguro")%>" method="post" name="frmPagSeguro">
                                        <input type="hidden" name="email_cobranca" value="<%= configuracao.getAttribute("email_cobranca") %>">
                                        <input type="hidden" name="tipo"  value="CP">
                                        <input type="hidden" name="moeda" value="BRL">

                                        <input type="hidden" name="cliente_nome"   value="<%= NomeCliente %>">
                                        <input type="hidden" name="cliente_cep"    value="<%= raiz_dados_pedido.getAttribute("cep_frete") %>">
                                        <input type="hidden" name="cliente_end"    value="<%= Session("logradouro_cobranca") %>">
                                        <input type="hidden" name="cliente_num"    value="<%= Session("numero_cobranca") %>">
                                        <input type="hidden" name="cliente_compl"  value="<%= Session("complemento_cobranca") %>">
                                        <input type="hidden" name="cliente_bairro" value="<%= Session("bairro_cobranca") %>">
                                        <input type="hidden" name="cliente_cidade" value="<%= Session("cidade_cobranca") %>">
                                        <input type="hidden" name="cliente_uf"     value="<%= Session("estado_cobranca") %>">
                                        <input type="hidden" name="cliente_pais"   value="BRA">
                                        <input type="hidden" name="cliente_ddd"    value="<%= Session("ddd_cobranca") %>">
                                        <input type="hidden" name="cliente_tel"    value="<%= Session("telefone_cobranca") %>">
                                        <input type="hidden" name="cliente_email"  value="<%= Session("user_id") %>">

                                        <input type="hidden" name="ref_transacao"  value="<%= Session("codigo_pedido") %>">
                                        <input type="hidden" name="tipo_frete"     value="<%= tipo_de_frete %>">

                                        <%
                                        Cont = 0
                                        For Each Node In ProdutosItens
                                        Cont = Cont + 1
                                        %>
                                        <input type="hidden" name="item_id_<%= CStr(Cont) %>"    value="<%= mid(Node.getAttribute("codigo_produto"), 1, Instr(Node.getAttribute("codigo_produto"),"_")-1) %>">
                                        <input type="hidden" name="item_descr_<%= CStr(Cont) %>" value="<%= Node.getAttribute("nome_produto") %>">
                                        <input type="hidden" name="item_quant_<%= CStr(Cont) %>" value="<%= Node.getAttribute("quantidade_produto") %>">
                                        <input type="hidden" name="item_valor_<%= CStr(Cont) %>" value="<%= Replace(Node.getAttribute("preco_unitario"),",","") %>">
                                        <input type="hidden" name="item_frete_<%= CStr(Cont) %>" value="<%= Replace(raiz_dados_pedido.getAttribute("valor_frete"),",","") %>">
                                        <input type="hidden" name="item_peso_<%= CStr(Cont) %>"  value="<%= Replace(Node.getAttribute("peso_parcial"),",","") %>">
                                        <%
                                        Next
                                        %>

                                    </form>
                                    <SCRIPT LANGUAGE=javascript>
                                    <!--
                                        document.frmPagSeguro.submit();
                                    //-->
                                    </SCRIPT>

                                    <%
									' Opção de pagamento PAGAMENTO CERTO
									ElseIf Session("forma_pagamento") = "PagamentoCerto" Then

										mensagemAdicional = "<b>" & Application("InitTxtInstrAlterarPedido") & " <a href=""" & Application("URLloja") & "/carrinho.asp?lang=" & varLang & "&mode=changeMeioPagto"">" & Application("InitTxtCliqueAqui") & "</a></b>"

                                        ' Resgata alguns valores padrão
                                        processada = False

                                        total = Replace(Replace(raiz_dados_pedido.getAttribute("valor_total"),",",""),".","")

                                        ' Monta o XML com os dados da transação
                                        xmlTransacao = XmlDados_Transacao(Session("codigo_pedido"),sParamAdicionais)

                                        ' Monta os parâmetros de entrada do Web Service
                                        entrada = "<?xml version=""1.0"" encoding=""utf-8""?>"
                                        entrada = entrada & "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
                                        entrada = entrada & "<soap12:Body>"
                                        entrada = entrada & "   <IniciaTransacao xmlns=""http://www.locaweb.com.br"">"
                                        entrada = entrada & "      <chaveVendedor>" & configuracao.getAttribute("chaveVendedor") & "</chaveVendedor>"
                                        entrada = entrada & "      <urlRetorno>" & Application("URLRecibo") & "</urlRetorno>"
                                        entrada = entrada & "      <xml>" & Server.HTMLEncode(xmlTransacao) & "</xml>"
                                        entrada = entrada & "    </IniciaTransacao>"
                                        entrada = entrada & "  </soap12:Body>"
                                        entrada = entrada & " </soap12:Envelope>"

                                        set objXmlDom = CreateObject("Microsoft.XMLDOM")
                                        set objXmlHttp = CreateObject("Microsoft.XMLHTTP")
                                         
                                        ' Efetua a conexão ao Web Service
                                        objXmlHttp.open "POST", Application("URLWSPagamentoCertoLocaweb"), false
                                        objXmlHttp.setRequestHeader "Man", POST & " " & Application("URLWSPagamentoCertoLocaweb") & " HTTP/1.1"
                                        objXmlHttp.setRequestHeader "MessageType", "CALL"
                                        objXmlHttp.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
                                        objXmlHttp.send(entrada)

                                        ' Resgata o XML de resposta
                                        retorno = objXmlHttp.responsetext

                                        ' Verifica se o processo de registro da transação foi feito com sucesso
                                        If objXmlHttp.Status = 200 Then

                                            ' Trata o retorno do processo
                                            objXmlDom.async = False
                                            objXmlDom.LoadXML(retorno)
                                            xmlRetornoTransacao = objXmlDom.selectSingleNode("soap:Envelope/soap:Body/IniciaTransacaoResponse/IniciaTransacaoResult").text

                                            Set objXmlDom = Nothing
                                            Set objXmlDom = CreateObject("Microsoft.XMLDOM")

                                            ' Trata o retorno de retorno do registro da transação
                                            objXmlDom.async = False
                                            objXmlDom.LoadXML(xmlRetornoTransacao)
                                            
                                            ' Resgata os dados iniciais do retorno da transação
                                            nodeCodRetornoInicio = objXmlDom.selectSingleNode("LocaWeb/Transacao/CodRetorno").text
                                            nodeMensagemRetornoInicio = objXmlDom.selectSingleNode("LocaWeb/Transacao/MensagemRetorno").text

                                            ' Verifica se o registro da transação foi feito com sucesso
                                            If nodeCodRetornoInicio = "0" Then

                                                ' Resgata o id e a mensagem da transação
                                                nodeIdTransacao = objXmlDom.selectSingleNode("LocaWeb/Transacao/IdTransacao").text
                                                nodeCodigoRef = objXmlDom.selectSingleNode("LocaWeb/Transacao/Codigo").text

                                                processada = True

                                            Else

                                                ' Exibe a mensagem de erro
                                                Response.write "<b>Erro: " & nodeMensagemRetornoInicio & "</b>"
                                                Response.write "<br>" & mensagemAdicional

                                            End If

                                        Else

                                            ' Exibe a mensagem de erro
                                            Response.write "<b>Erro: (" & objXmlHttp.Status & ") " & objXmlHttp.statusText & "</b>"
                                            Response.write "<br>" & mensagemAdicional
                                         
                                        End If
                                         
                                        Set objXmlHttp = Nothing
                                        Set objXmlDom = Nothing

                                        'Grava os dados iniciais da transação no banco de dados
                                        Call GravaTransacaoInicialPagamentoCerto(Session("codigo_pedido"),nodeIdTransacao,nodeCodigoRef,Now(),sModulo,sTipoModulo,nodeCodRetornoInicio,nodeMensagemRetornoInicio)

                                        ' Se a transação for processada com sucesso
                                        If processada Then
                                            ' Inicia a transação
                                            Response.write "<script> location.href='" & Application("URLPagamentoCertoLocaweb") & "?tdi=" & nodeIdTransacao &"'; </script>"
                                        End If

                                    ' Opção de pagamento Paggo
                                    ElseIf Session("forma_pagamento") = "Paggo" Then

										If Request("dados_cartao") <> "" Then

											numeroCelular = Request("dados_cartao")
											
											If Request("dados_adicionais") = "01" Or Request("dados_adicionais") = "" Then
												parcelas = "01"
											Else
												parcelas = Request("dados_adicionais")
											End If

											total = Replace(Replace(raiz_dados_pedido.getAttribute("valor_total"),",",""),".","")

											'Grava os dados iniciais da transação no banco de dados
											Call GravaTransacaoInicialPaggo(Session("codigo_pedido"),numeroCelular,Now())

										%>
											<div id="textoPaggo">
											<p align="center"><IMG border="0" src="config/templates/<%=varLang%>/<%=varSkin%>/banner_paggo.gif"></p>
											<p align="center"><b><%= Application("InitTxtAvisoPaggo")%></b></p>
											</div>
											<iframe frameborder="no" width="0" height="0" id="transacaoPaggo" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" name="transacaoPaggo" noresize scrolling="auto" src="dados_compra_paggo.asp?pedido=<%= Session("codigo_pedido")%>"></iframe>
										<%
										Else

											' Exibe a mensagem de erro
                                            Response.write "<b>Erro: " & Application("InitTxtProbTransacao") & "</b>"
                                            Response.write "<br>" & Application("InitTxtInstrAlterarPedido") & " <a href=""" & Application("URLloja") & "/carrinho.asp?lang=" & varLang & "&mode=changeMeioPagto"">" & Application("InitTxtCliqueAqui") & "</a>"

										End If
									End If
                                    %>
                                </td>
                            </tr>
                            <%
                            Set raiz_dados_pedido = Nothing
                            'Fecha conexao ao XML do pedido.
                            Call fecha_xmlpedido(session("id_transacao"))

                            'Fecha conexao ao XML dos meios de pagto.
                            Call fecha_ArquivoXML(Application("XMLMeiosPagamentos"),FctobjXML,FctobjRoot) 
                            %>
                        </table>
                    </td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <% IF Session("cep_entrega") <> "" THEN %>
                    <td  height="25" align="center" bgcolor="#EAEAEA" disabled><B><%=Application("InitTxtTitDadosCobranca")%></B></td>
                    <% ELSE %>
                    <td  height="25" align="center" bgcolor="#EAEAEA" disabled><B><%=Application("InitTxtTitDadosCobrancaEntrega")%></B></td>
                    <% END IF %>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td align="center"  valign="top">
                        <table width="100%" border=0 align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                            <tr bgcolor="#EFEFEF">        
                                <td align="center"><%Call Mostra_Endereco("cobranca")%></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <script>
                checkTipCad('cobranca');
                </script>
                <% IF Session("cep_entrega") <> "" THEN %>
                <tr bgcolor="#FFFFFF">
                    <td  height="25" align="center"  bgcolor="#EAEAEA" disabled><B><%=Application("InitTxtTitDadosEntrega")%></B></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td align="center"  valign="top">
                        <table width="100%" border=0 align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                            <tr bgcolor="#EFEFEF">        
                                <td align="center"><%Call Mostra_Endereco("entrega")%></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <script>
                checkTipCad('entrega');
                </script>
                <% END IF %>
            </table>
            </form>
        </td>
        <td valign="top" height="10%" width="10%" class="TBLlatdireita"><!--#INCLUDE FILE="lateral_servicos.asp" --></td>
    </tr>
    <tr>
        <td colspan="3" valign="top"><!--#INCLUDE FILE="rodape.htm" --></td>
    </tr>
</table>