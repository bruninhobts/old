<%
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Loja Exemplo Locaweb
' Versão: 6.5
' Data: 12/09/06
' Arquivo: recibo.asp
' Versão do arquivo: 0.0
' Data da ultima atualização: 13/10/08
'
'-----------------------------------------------------------------------------
' Licença Código Livre: http://comercio.Locaweb.com.br/gpl/gpl.txt
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-

rodape = "no"
navegacaocompra = "fim"
page = "recibo"
passo=4
%>
<!--#INCLUDE FILE="funcoes/funcoes_grava_transacao.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_estrutura_recibo.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_cartao.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_usuario.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_endereco.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_uteis.asp" -->
<!--#INCLUDE FILE="funcoes/funcoes_mail.asp" -->
<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
        <td colspan="3" valign="top" height="30"><!--#INCLUDE FILE="cabecalho.asp" --></td>
    </tr>
    <tr><%If navegacaocompra = "fim" Then%>
        <td valign="top" height="10%" width="10%" class="TBLlatesquerda"><!--#INCLUDE FILE="menu_poscarrinho.asp" --></td>
        <%Else%>
        <td valign="top" height="10%" width="10%" class="TBLlatesquerda"><!--#INCLUDE FILE="menu.asp" --></td>
        <%End If%>
        <td valign="top" height="95%">
            <%
            permissao = "read"
            page = "recibo"
            readonly = "readonly"
            compra = "aprovada"

            ' Condição para alterar a forma de pagamento do pedido
            If Request.Form("mode") = "changeMeioPagto" Then

                ' Caso o arquivo XML do pedido não for localizado haverá um redirecionamento para página de carrinho vazio.
                If Not VerificaExistenciaArquivo(Application("DiretorioPedidos")&session("id_transacao")&".xml") Then
                    Response.redirect ("carrinho_vazio.asp")
                    Response.end
                End If

                ' Ajusta a sessão da opção de pagamento
                Session("forma_pagamento") = Request.Form("newMeioPagto")
                                
                ' Remonta o valor total do pedido
                currTotalPedido = Cdbl(pegaValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","valor_subtotal")) + Cdbl(pegaValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","valor_frete"))

                ' Atualiza o XML do pedido com os novos dados
                Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","valor_total",FormatNumber(currTotalPedido,2))
                Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","num_parcelas","01")
                Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","tipo_taxa_adicional","")
                Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","taxa_adicional","0")
                Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","forma_pagamento","Boleto")

                ' Formata o valor total do pedido
                VARtotal = replace(currTotalPedido,".","")
                VARtotal = replace(VARtotal,",",".")
                
                ' Atualiza o banco de dados com os novos dados
                Conexao.Execute("UPDATE Pedidos SET boleto_emitido = 1, tipo_taxa_adicional = '', taxa_adicional = 0, num_parcelas = 1, total = "& VARtotal &" WHERE codigo_pedido = " & Request.Form("codigo_pedido"))
            End If

            ' ***********************  AMEX **********************
            If Request("MerchTxnRef") <> "" Then
                ' Captura os parâmetros de retorno
                MerchTxnRef = request("MerchTxnRef")
                OrderInfo = request("OrderInfo")
                BatchNo = request("BatchNo")
                ReceiptNo = request("ReceiptNo")
                AuthorizeId = request("AuthorizeId")
                TransactionNo = request("TransactionNo")
                AcqResponseCode = request("AcqResponseCode")
                TxnResponseCode = request("TxnResponseCode")
                Message = request("Message")
                CSCResultCode = request("CSCResultCode")
                CSCRequestCode = request("CSCRequestCode")
                AcqCSCRespCode = request("AcqCSCRespCode")

                ' Caso a referência da transação seja nula, resgata do campo livre
                If MerchTxnRef = "" Or MerchTxnRef = "No Value Returned" Then
                    MerchTxnRef = OrderInfo
                End If

                'Grava os dados da transação Amex no banco de dados
                Call GravaTransacaoFinalAmex(MerchTxnRef,BatchNo,ReceiptNo,AuthorizeId,TransactionNo,AcqResponseCode,TxnResponseCode,Message,CSCResultCode,CSCRequestCode,AcqCSCRespCode)

                If TxnResponseCode <> "0" Then
                    compra = "negada"
                    Session("codigo_pedido") = MerchTxnRef
                    cod_erro = TxnResponseCode
                    msg_erro = AMEX_getresponseDescription(TxnResponseCode)
                End If

            End If
            ' ***********************  ITAU **********************
            'Post para descriptografia do DC retornado
            If Request("DC") <> "" Then
                Call Recibo(Conexao,"Itau",compra,cod_erro,msg_erro,identificacao_pedido)
            End If
            ' ***********************  VISANET **********************
            'Captura dados de retorno da Visanet
            If Request("TID") <> "" Then
                'Verifica se a transação é Visa ou VisaElectron
                If Right(Request("TID"),4) = "A001" Then
                    formaPagamento = "VisaElectron"
                Else
                    formaPagamento = "Visa"
                End If                
                Call Recibo(Conexao,formaPagamento,compra,cod_erro,msg_erro,identificacao_pedido)
            End If

            ' ***********************  BANCO DO BRASIL **********************
            'Verificacao inicial
            If Request("refTran") <> "" And Request("RECIBOFIM") <> "1" Then
                Call Recibo(Conexao,"Brasil",compra,cod_erro,msg_erro,identificacao_pedido)
            End If

            ' ***********************  REDECARD **********************
            'Verificacao inicial
            If Request("NR_CARTAO") <> "" Then
                Call Recibo(Conexao,"Mastercard",compra,cod_erro,msg_erro,identificacao_pedido)
            End If

            ' ***********************  BRADESCO **********************
            'Captura dados de retorno do Bradesco
            If Request("merchantid") <> "" Then
                ' Recria a sessão do codigo do pedido
                If Request("numOrder") <> "" Then
                    Session("codigo_pedido") = Request("numOrder")
                ElseIf Request("orderId") <> "" Then
                    Session("codigo_pedido") = Request("orderId")
                End If

                ' Atualiza o banco de dados com o status da transação
                Set RS_Bradesco = CreateObject("ADODB.Recordset")
                Set RS_Bradesco.ActiveConnection = Conexao
                    RS_Bradesco.CursorLocation = 3
                    RS_Bradesco.CursorType = 0
                    RS_Bradesco.LockType =  3

                    RS_Bradesco.Open "SELECT codigo_pedido, cod, errordesc FROM Transacao_Bradesco WHERE codigo_pedido = "&Session("codigo_pedido")&"", Conexao

                    If Request("cod") <> "" Then
                        RS_Bradesco("cod")  = Request("cod")
                    End if
                    If Request("errordesc") <> "" Then
                        RS_Bradesco("errordesc") = Request("errordesc")
                    End If
                    
                    RS_Bradesco.Update
                    RS_Bradesco.Close
                Set RS_Bradesco = nothing

                'Verifica se houve erro na transação
                If Request("errordesc") <> "" Or Request("cod") <> "0" Then
                    compra = "negada"
                    cod_erro = Request("cod")
                    msg_erro = Request("errordesc")

                    If (request("cod")<=-101) And (request("cod")>=-104) Or (request("cod")=-111) Or (request("cod")=-125) Or (request("cod")=-124) Then
                        VarStrAdicional = Application("RecTxtPagtoSujConfirmacao")
                    End If
                End If
            End If

            ' ***********************  ABNCDC **********************
            If Request("RET01") <> "" Then
                ' Erro no processamento da transacao
                If Request("RET01") <> 1 And Request("RET01") <> 2 Then
                    compra = "negada"
                    cod_erro = Request("RET01")
                    msg_erro = ABN_MSG_status(Request("RET01"), Request("RET02"))
                End If
                'Ativa a sessão para o código do pedido
                Session("codigo_pedido") = Request.Querystring("NUMPEDIDO")
            End If

			' ***********************  PAGAMENTO CERTO **********************
            If Request("tdi") <> "" Then
                
                ' Captura os parâmetros de retorno
                idTransacao = request("tdi")

                Call Recibo(Conexao,"PagamentoCerto",compra,cod_erro,msg_erro,identificacao_pedido)

            End If

			' ***********************  PAGGO **********************
            If Session("forma_pagamento") = "Paggo" Then
                
                Call Recibo(Conexao,"Paggo",compra,cod_erro,msg_erro,identificacao_pedido)

            End If

            ' ***********************  PAGSEGURO **********************

	          If Request("TransacaoID") <> "" Then
			  
			  ' Abre o xml para acha o token salvo e enviar para validação
              Call abre_ArquivoXML(Application("XMLMeiosPagamentos"),FctobjXML,FctobjRoot) 
			  ' Procura o nó no xml referente ao xml
              Set configuracao = FctobjRoot.selectSingleNode("configuracao/pagto[@nome_pagto='PagSeguro']")
			  ' Concatena o token com o post do pagseguro
              Ret_pag = Request.Form & "&Comando=validar&Token=" & configuracao.getAttribute("token")
			  ' Pega o email
              Loja_email = configuracao.getAttribute("email_cobranca")
			  ' fecha o xml
              Call fecha_ArquivoXML(Application("XMLMeiosPagamentos"),FctobjXML,FctobjRoot) 

              CodPedido       = Request("Referencia")
              TransacaoID     = Request("TransacaoID")
              DataTransacao   = Request("DataTransacao")
              TipoPagamento   = Request("TipoPagamento")
              StatusTransacao = Request("StatusTransacao")
              CliEmail        = Request("CliEmail")
              VendedorEmail   = Request("VendedorEmail")

              Session("codigo_pedido") = CodPedido

              Set FSO = CreateObject("Scripting.FileSystemObject")
              Set Arquivo = FSO.CreateTextFile(Application("DiretorioDados") & "1_Log_PS_" & Replace(CliEmail,"@","_") & ".txt")
              Arquivo.WriteLine (CStr(Ret_pag))
              Set Arquivo = Nothing
              Set FSO = Nothing

              If (Trim(Loja_email) = Trim(VendedorEmail)) Then 

                 SET objHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")

                 objHttp.OPEN "POST", "https://pagseguro.uol.com.br/Security/NPI/Default.aspx", False
                 objHttp.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
                 objHttp.Send Ret_pag

				 If (objHttp.status = 200) Then

                    IF (objHttp.responseText = "VERIFICADO") Then

                       Call GravaTransacaoPagSeguro(CodPedido, TransacaoID, DataTransacao, TipoPagamento, StatusTransacao, CliEmail)
 
                       If (Instr(1,LCase(StatusTransacao),"aprovado") > 0) Then 
                          'Abre conexao ao BD para captura dos dados do pedido
                          Set RS_Pedido = CreateObject("ADODB.Recordset")
                          Set RS_Pedido.ActiveConnection = Conexao
                          RS_Pedido.CursorLocation = 3
                          RS_Pedido.CursorType = 0
                          RS_Pedido.LockType =  3

                          RS_Pedido.Open "SELECT Pedidos.codigo_pedido, Pedidos.pago FROM Pedidos WHERE codigo_pedido = " & CodPedido & "" , Conexao
                          If Not RS_Pedido.EOF Then
                             RS_Pedido("pago") = "1"
                             RS_Pedido.Update
                          End If
                          RS_Pedido.Close
                          Set RS_Pedido = nothing
                       End If
                    End IF
                 End If
                 SET objHttp = NOTHING
              End If

              If Session("forma_pagamento") = "" Then
                 Response.end
              End If
            End If
            %>
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="10">
                <tr>
                    <td align="center" height="18" valign="middle"><!--#INCLUDE FILE="barra_passoapasso.asp" --></td>
                </tr>
            <% 
            ' ***********************  EM CASO DE COMPRA NEGADA **********************
            If compra = "negada" Then
            %>

                <tr class="FUNDOTABtopico">
                    <td align="center" height="18" valign="middle"><B><span class="TXTTABtopico"><%=Application("RecTxtTitResultadoTrans")%></span></B></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                    <td align="center"  valign="top">
                        <table width="100%" border=0 align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                            <tr>
                                <td><font color='red'><B><%=Application("RecTxtTitMsgErro")%></B></font></td>
                            </tr>
                            <tr>
                                <td><% Response.write Application("RecTxtNumPedido") &": <b>" & Session("codigo_pedido")%></td>
                            </tr>
                            <%If identificacao_pedido <> "" Then%>
                            <tr>
                                <td><% Response.write Application("RecTxtIdentPedido") & ": <b>" & identificacao_pedido%></td>
                            </tr>
                            <%End if%>
                            <tr>
                                <td><% Response.write Application("RecTxtCodigoErro") & ": <b>" & cod_erro%></td>
                            </tr>
                            <tr>
                                <td><% Response.write Application("RecTxtMsgErro") & ": <b>" & msg_erro%></td>
                            </tr>
                            <tr>
                                <td><% Response.write  VarStrAdicional%></td>
                            </tr>
                            
                            <tr>
                                <td>
                                    <b><%=Application("RecTxtInstrAlterarPagto")%> <a href="<%= Application("URLloja")%>/carrinho.asp?lang=<%=varLang%>&mode=changeMeioPagto"><%=Application("RecTxtCliqueAqui")%></a>
                                    <% If pegaValorAtrib(Application("XMLMeiosPagamentos"),"configuracao/pagto[@nome_pagto='Boleto']","disponivel") = "sim" Or pegaValorAtrib(Application("XMLMeiosPagamentos"),"configuracao/pagto[@nome_pagto='CobreBem']","disponivel") = "sim" Then %>
                                        <%=Application("RecTxtInstrFinalizarPagtoBoleto")%> <a onclick="document.formPedido.submit();" style="cursor:pointer;"><%=Application("RecTxtCliqueAqui")%></a></b>
                                    <form method="POST" name="formPedido" action="recibo.asp">
                                        <input type="hidden" name="lang" value="<%=varLang%>">
                                        <input type="hidden" name="codigo_pedido" value="<%=Session("codigo_pedido")%>">
                                        <input type="hidden" name="mode" value="changeMeioPagto">
                                        <input type="hidden" name="newMeioPagto" value="Boleto">
                                    </form>
                                    <% End If %>
                                </td>
                            </tr>
                        </table>
            <% 
            ' ***********************  EM CASO DE COMPRA APROVADA **********************
            ElseIf compra = "aprovada" Then
            %>
                <tr class="FUNDOTABtopico">
                    <td align="center" height="18" valign="middle"><B><span class="TXTTABtopico"><%=Application("RecTxtTitReciboCompra")%><span></B></td>
                </tr>
                <%If pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","EbitAtivo") = "sim" Then%>
                <tr bgcolor="#FFFFFF" align="center" width="468" height="60">
                    <%CodigoEbit=pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","CodigoEbit")%>
                    <form name="formebit" method="get" target="_blank" action="https://www.ebitempresa.com.br/bitrate/pesquisa1.asp">
                    <td align="center" height="18" valign="top">                        
                        <input type=hidden name='empresa' value='<%=CodigoEbit%>'><input type="image" border="0" name="banner" src="https://www.ebitempresa.com.br/bitrate/banners/b<%=CodigoEbit%>.gif"alt="O que voc&ecirc; achou desta loja?" width="468" height="60" target="_blank">                                      
                    </td>
                    </form>     
                </tr>
                <%End if%>
                <tr bgcolor="#FFFFFF">
                    <td align="center"  valign="top">
                        
                        <%
                        'Prazo de entrega conforme a modalida selecionada pelo comprador
                        opcao_frete = Pega_DadoBanco("Pedidos","tipo_frete","codigo_pedido",session("codigo_pedido"))
                        'Forma de entrega SEDEX
                        If opcao_frete = "SEDEX" Then
                            msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaSedex")
                        'Forma de entrega PAC
                        ElseIf opcao_frete = "PAC" Then
                                msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaPAC") 
                        'Forma de entrega E-SEDEX
                        ElseIf opcao_frete = "E-SEDEX" Then
                            msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaESedex")
                        'Forma de entrega DIRECT EXPRESS
                        ElseIf opcao_frete = "DIRECT EXPRESS" Then
                            msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaDirectExpress")
                        'Forma de entrega TRANSPORTADORA A COBRAR
                        ElseIf opcao_frete = "TRANSPORTADORA A COBRAR" Then
                            msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaTransportadora")
                        'Forma de entrega FEDEX
                        ElseIf opcao_frete = "FEDEX" Then
                            msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaFedex")
                        'Forma de entrega RETIRAR NA LOJA
                        ElseIf opcao_frete = "RETIRAR NA LOJA" Then
                            msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaRetirarNaLoja")
                        'Forma de entrega FRETE PERSONALIZADO
                        ElseIf opcao_frete = UCASE(pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","NomeFretePersonalizado")) Then
                            msgPrazoEntrega =  pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","PrazoEntregaFretePersonalizado")
                        End If                        
                        %>

                        <table width="100%" border=0 align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                            <tr><td><B><%=Application("RecTxtNumPedido")%>:</B> <FONT SIZE="2" COLOR="red"><B><%=session("codigo_pedido")%></B></FONT></td></tr>
                            <tr><td><B><%=Application("RecTxtIPUsado")%>:</B> <%=request.ServerVariables("REMOTE_ADDR")%></td></tr>
                            <tr><td><B><%=Application("RecTxtPrazoEntrega")%>:</B>  <%= msgPrazoEntrega%></td></tr>
                            <tr><td><B><%=Application("RecTxtFormaEntrega")%>:</B> <%=opcao_frete%></td></tr>                

                            <tr bgcolor="#EFEFEF">        
                                <td align="left">
            <%
                                    ' ***********************  BOLETO **********************
                                    If Session("forma_pagamento") = "Boleto" Then
            %>
                                        <p align="center"><a class="TextoPadrao"><b><%=Application("RecTxtTiInstrPagto")%>:</b></a></p>
            <%	
                                        Call Recibo(Conexao,Session("forma_pagamento"),compra,cod_erro,msg_erro,identificacao_pedido)
                                    End If

                                    ' ***********************  AMEX **********************
                                    If Request("TxnResponseCode") <> "" Then
                                        
                                        MerchTxnRef = Request("MerchTxnRef")
                                        OrderInfo = Request("OrderInfo")

                                        ' Caso a referência da transação seja nula, resgata do campo livre
                                        If MerchTxnRef = "" Or MerchTxnRef = "No Value Returned" Then
                                            MerchTxnRef = OrderInfo
                                        End If

                                        'Ativa a sessão para o código do pedido
                                        Session("codigo_pedido") = MerchTxnRef
                                        
                                        Call Recibo(Conexao,"Amex",compra,cod_erro,msg_erro,identificacao_pedido)
                                    End If

                                    ' ***********************  VISANET **********************
                                    If Request("tid") <> "" Then
                                    
                                        ' Exibe as váriaveis de retorno
                                        Response.write Application("RecTxtCodigoTrans") & ": " & session("TID") & "<br>"
                                        Response.write Application("RecTxtCodigoResposta") & ": " & session("LR") & "<br>"
                                        Response.write Application("RecTxtCodigoAutorizacao") & ": " & session("ARP") & "<br>"

                                        If session("ARS") <> "" Then
                                            Response.write Application("RecTxtMsgTransacao") & ": " & session("ARS") & "<br>"
                                        End If
                                        If session("AUTHENT") <> "" Then
                                            Response.write Application("RecTxtTipoAutent") & ": " & session("AUTHENT")  & "<br>"
                                        End If

                                        Response.write "<br>"

                                    End If

                                    ' ***********************  REDECARD **********************
                                    If Request("NR_CARTAO") <> "" Then

                                        ' ************** Em caso da transação já ter sido confirmada ***************
                                        If (status = 1) Then
                                            Response.write Application("RecTxtMsgTransJaConfirmada")
                                        End If 

                                        ' ************************** Monta o cupom *********************************
                                        If varLang = "en_UK" Or varLang = "en_US" Then
                                            LANGUAGE = "E" 'Inglês
                                        Else
                                            LANGUAGE = "" 'Português
                                        End If

                                        URLCupom = "https://ecommerce.redecard.com.br/pos_virtual/cupom.asp?DATA=" & request("DATA") & "&TRANSACAO=201&NUMAUTOR=" & request("NUMAUTOR") & "&NUMCV=" & request("NUMCV")
                                        
                                        If LANGUAGE <> "" Then
                                            URLCupom = URLCupom & "&LANGUAGE=" & LANGUAGE
                                        End If
                                        
                                        %>
                                        <SCRIPT LANGUAGE=javascript>
                                        <!--
                                                vpos=window.open('<%= URLCupom %>','vpos','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=auto,resizable=no,copyhistory=no,width=290,height=460');
                                        //-->
                                        </SCRIPT>
            <%
                                    End If

                                    ' ***********************  ITAU **********************
                                    ' **** link para a conexao com o Itaú
                                    If Request("codEmp") <> "" And Request("tipPag") <> "" Then
                                        'Ativa a sessão para o código do pedido
                                        Session("codigo_pedido") = Request("pedido")

                                        'Atualiza alguns dados restantes do final da transação Itaú no banco de dados
                                        Call GravaTransacaoFinalItau(Session("codigo_pedido"),Request("tipPag"))

                                        ' Exibe as váriaveis de retorno
                                        Response.write Application("RecTxtFormaPagto") & ": " & ITAU_TipPag(Request("tipPag"))  & "<br><br>"

                                    End If 

                                    ' ***********************  BANCO DO BRASIL **********************
                                    If Session("forma_pagamento") = "Brasil" Then
                                        ' Exibe a forma de pagamento escolhida
                                        Response.write Application("RecTxtFormaPagto") & ": " & BB_TipPag(Request("tpPagamento")) & "<br><br>"
                                    End If

                                    ' ***********************  UNIBANCO **********************
                                    If Session("forma_pagamento") = "Unibanco" Then
                                        Call Recibo(Conexao,Session("forma_pagamento"),compra,cod_erro,msg_erro,identificacao_pedido)
                                    End If

                                    ' ***********************  DEPOSITO BANCARIO **********************
                                    If Session("forma_pagamento") = "Deposito" Then
                                        Call Recibo(Conexao,Session("forma_pagamento"),compra,cod_erro,msg_erro,identificacao_pedido)
                                    End If

                                    ' ***********************  COBREBEM ECOMMERCE **********************
                                    If Session("forma_pagamento") = "CobreBem" Then
                                        Call Recibo(Conexao,Session("forma_pagamento"),compra,cod_erro,msg_erro,identificacao_pedido)
                                    End If

                                    ' ***********************  BRADESCO **********************
                                    If Request("cod") = "0" Then
                                        Call Recibo(Conexao,"Bradesco",compra,cod_erro,msg_erro,identificacao_pedido)
                                    End If

                                    ' ***********************  ABNCDC **********************
                                    If Request("RET01") <> "" Then
                                        Call Recibo(Conexao,"ABNCDC",compra,cod_erro,msg_erro,identificacao_pedido)
                                    End If

									' ***********************  PAGAMENTO CERTO **********************
                                    If Request("tdi") <> "" Then
                                    
                                        ' Exibe as váriaveis de retorno
                                        Response.write "<b>" & Application("RecTxtCodigoTrans") & ":</b>&nbsp;" & identificacao_pedido & "<br>"
                                        Response.write "<b>" & Application("RecTxtMsgTransacao") & ":</b>&nbsp;" & Pega_DadoBanco("Transacao_PagamentoCerto","msgRetornoPagamento","idTransacao","'"&Request("tdi")&"'")

                                        Response.write "<br>"

										If Session("URLBoleto") <> "" Then

											' URL para geração do boleto
											idTransacao = Pega_DadoBanco("Transacao_PagamentoCerto","idTransacao","codigo_pedido",Session("codigo_pedido"))
											str_Boleto = str_Boleto & "tdi=" & idTransacao
											URLBoleto = Application("URLLocaWebBoletoPagamentoCerto") & "?" & str_Boleto

											Response.write "<p align=""center""><a class=""TextoPadrao""><b>" & Application("RecTxtTiInstrPagto") & ":</b></a>&nbsp;&nbsp;" & Application("FestrTxtParaImprimir") & "<a Onclick=""javascript:JanelaNova('" & URLBoleto & "',700,500);"" style=""cursor:pointer;"">"& Application("FestrTxtCliqueAqui") & "</a> (" & Application("FestrTxtUtilizeImpressora") & ").</p>"

										End If


                                    End If

									' ***********************  PAGGO **********************
                                    If Session("forma_pagamento") = "Paggo" Then
                                    
                                        ' Exibe as váriaveis de retorno
                                        Response.write "<b>" & Application("RecTxtCodigoResposta") & ":</b>&nbsp;" & Pega_DadoBanco("Transacao_Paggo","codRetornoTransacao","codigo_pedido",Session("codigo_pedido")) & "<br>"
										Response.write "<b>" & Application("RecTxtMsgTransacao") & ":</b>&nbsp;" & Pega_DadoBanco("Transacao_Paggo","msgRetornoTransacao","codigo_pedido",Session("codigo_pedido")) & "<br>"
										Response.write "<b>NSU PAGGO:</b>&nbsp;" & Pega_DadoBanco("Transacao_Paggo","nsuPaggo","codigo_pedido",Session("codigo_pedido"))

                                    End If


									' ***********************  PAGSEGURO **********************
                                    If Session("forma_pagamento") = "PagSeguro" Then

	                                   Set RS_PagSeguro = CreateObject("ADODB.Recordset")
                                       Set RS_PagSeguro.ActiveConnection = Conexao
                                       RS_PagSeguro.CursorLocation = 3
                                       RS_PagSeguro.CursorType = 0
                                       RS_PagSeguro.LockType =  1

                                       RS_PagSeguro.Open "SELECT codigo_pedido, transacaoid, datatransacao, tipopagamento, statustransacao, cliemail FROM Transacao_PagSeguro WHERE codigo_pedido = " & Session("codigo_pedido") & "", Conexao

                                       If Not RS_PagSeguro.EOF Then
                                          Response.write "<b>" & Application("CtuscadTxtTitDadosAtualizados") & " PAGSEGURO:</b><br><br>"
                                          Response.write Application("RecTxtCodigoTrans")    & ": <b>" & RS_PagSeguro("transacaoid")     & "</b><br>"
                                          Response.write Application("CtusmospedTxtDataPag") & ": <b>" & RS_PagSeguro("datatransacao")   & "</b><br>"
                                          Response.write Application("CtusmospedTxtTipPag")  & ": <b>" & RS_PagSeguro("tipopagamento")   & "</b><br>"
                                          Response.write "Status: <b>" & RS_PagSeguro("statustransacao") & "</b><br><br>"
                                       End If

                                       RS_PagSeguro.Close
                                       Set RS_PagSeguro = nothing
                                    End If
            %>
                                </td>
                            </tr>
                        </table>
            <%
            ' Resgata dados do Usuário nesta transação
            Call Session_Usuario_Transacao(Conexao,Session("codigo_pedido"))
            %>

                    </td>
                </tr>
            </table>
            <!--#INCLUDE FILE="lista_pedidos.asp" -->
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="10">
                <tr class="FUNDOTABtopico">
                    <% If Session("cep_entrega") <> "" Then %>
                    <td align="center" height="18" valign="middle"><B><span class="TXTTABtopico"><%=Application("RecTxtTitDadosCobranca")%></span></B></td>
                    <% Else %>
                    <td align="center" height="18" valign="middle"><B><span class="TXTTABtopico"><%=Application("RecTxtTitDadosCobrancaEntrega")%></span></B></td>
                    <% End if %>
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
                <tr class="FUNDOTABtopico">
                    <td align="center" height="18" valign="middle"><B><span class="TXTTABtopico"><%=Application("RecTxtTitDadosEntrega")%></span></B></td>
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
            <%
                'Envia o e-mail de notificação de compra
                Call Envia_Email_recibo(Conexao,Session("user_id"),Session("URLBoleto"), Replace(Session("LinhaDigitavel"),"&nbsp;&nbsp;&nbsp;","  "))
                
                'Finaliza arquivo de pedido caso exista
                Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
                If objFSO.FileExists(Application("DiretorioPedidos")&session("id_transacao")&".xml") Then

                    ' Atualiza o status do pedido no XML do pedido
                    Call alteraValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","status_pedido","finalizado")

                    ' Verifica se está ativo o DEBUG
                    If pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","DebugPedido") = "sim" Then

                        If objFSO.FolderExists(Application("DiretorioPedidos")&"debugpedidos") = False Then
                            Set CriaFSO = CreateObject("Scripting.FileSystemObject")
                                CriaFSO.CreateFolder(Application("DiretorioPedidos")&"debugpedidos")
                            Set CriaFSO = Nothing
                        End If

                        objFSO.CopyFile Application("DiretorioPedidos")&session("id_transacao")&".xml",Application("DiretorioPedidos")&"debugpedidos/"&session("id_transacao")&".xml"
                    
                    End If

                End If
                Set objFSO = Nothing

                ' Anula todas as sessions abertas, exceto codigo_pedido e id_transacao
                Call Anula_TodasSessions()

            End If
            %>
            </table>
        </td>
        <td valign="top" height="10%" width="10%" class="TBLlatdireita"><!--#INCLUDE FILE="lateral_servicos.asp" --></td>
    </tr>
    <tr>
        <td colspan="3" valign="top"><!--#INCLUDE FILE="rodape.htm" --></td>
    </tr>
</table>