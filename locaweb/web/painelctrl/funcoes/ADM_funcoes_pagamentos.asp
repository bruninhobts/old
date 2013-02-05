<%
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Loja Exemplo Locaweb
' Versão: 6.5
' Data: 12/09/06
' Arquivo: ADM_funcoes_pagamentos.asp
' Versão do arquivo: 0.0
' Data da ultima atualização: 14/10/08
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

'########################################################################################################
'SUB abre_xmlpagamentos
'   - Abre conexão com o arquivo de XML lista_meiospagamentos.xml
'   - Chamada no arquivo ADM_config_pagamento.asp
'SUB fecha_xmlpagamentos
'   - Fecha conexão com o arquivo de XML lista_meiospagamentos.xml
'   - Chamada no arquivo ADM_config_pagamento.asp
'########################################################################################################

Sub abre_xmlpagamentos(FctobjXML,FctobjRoot) 

    set FctobjXML = Server.CreateObject("Microsoft.XMLDOM")
        FctobjXML.preserveWhiteSpace = False
        FctobjXML.async = False
        FctobjXML.validateOnParse = True
        FctobjXML.resolveExternals = True
        FctobjXML.load (Application("XMLMeiosPagamentos"))
    Set FctobjRoot = FctobjXML.documentElement

End Sub

Sub fecha_xmlpagamentos(FctobjXML,FctobjRoot) 

    If request("acao") = "alterar" Then
        FctobjXML.save(Application("XMLMeiosPagamentos"))
    End if
    set FctobjXML = Nothing
    Set FctobjRoot = Nothing

End Sub

'########################################################################################################
'--> FIM SUB abre_xmlpagamentos e SUB fecha_xmlpagamentos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB altera_xmlpagamentos
' - Altera as informações referente aos meios de pagamentos.
' - Chamada no arquivo ADM_config_pagamento.asp
'########################################################################################################

Sub altera_xmlpagamentos(VarNome_pagto,objRoot,configuracao) 
    Set configuracao = objRoot.selectSingleNode("configuracao/pagto[@nome_pagto='"&VarNome_pagto&"']")

    If request("acao") = "alterar" Then

        FOR EACH count in request.form
            if request.form.key(count)<>"acao" Then
                If request.form.key("DadosDeposito") <> "" Or request.form.key("descricao_pagamento") <> "" Then
                    configuracao.setAttribute request.form.key(count),replace(replace(request.form.item(count),vbcrlf,"<br>"),Chr(34),"&quot;")
                Else                 
                    configuracao.setAttribute request.form.key(count),(request.form.item(count))  
                End if
            end if
        NEXT

    End if
    Set configuracao = Nothing
End Sub

'########################################################################################################
'--> FIM SUB altera_xmlpagamentos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_Combo_juros_parcelado
' - Monta as opções de juros, ref. ao parcelamento da opção de pagamento..
' - Chamada no arquivo ADM_funcoes_pagamento.asp
'########################################################################################################
Sub Cria_Combo_juros_parcelado(nome,opcao,onchange,tipJuros)

    If tipJuros = "Composto" Then
		Dim Valor(2), Tipo(2)
		Valor(1)="lojista"
		Valor(2)="emissor"
		Tipo(1)="Juros do lojista"
		Tipo(2)="Juros do Emissor"
	Else
		ReDim Valor(1), Tipo(1)
		Valor(1)="lojista"
		Tipo(1)="Juros do lojista"
	End If
%>
    <SELECT NAME="<%=nome%>" class="FORMbox" <%= onchange %>>
<% 

    For I=1 to UBound(Valor)
        If opcao = Valor(i) Then    %>
            <OPTION SELECTED VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% Else %>
            <OPTION VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% End If
    Next 
%>
    </SELECT>
<% End Sub
'########################################################################################################
'--> FIM SUB Cria_Combo_juros_parcelado
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_Combo_codigo_pagamentoVisaElectron
' - Monta as opções de código de pagamento para o VBV da Visanet, ref. visa electron.
' - Chamada no arquivo ADM_funcoes_pagamento.asp
'########################################################################################################
Sub Cria_Combo_codigo_pagamentoVisaElectron(nome,opcao)

    Dim Valor(1), Tipo(1)

    Valor(1)="A0"

    Tipo(1)="Pagamento à vista"
%>
    <SELECT NAME="<%=nome%>" class="FORMbox">
<% 

    For I=1 to UBound(Valor)
        If opcao = Valor(i) Then    %>
            <OPTION SELECTED VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% Else %>
            <OPTION VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% End If
    Next 
%>
    </SELECT>
<% End Sub
'########################################################################################################
'--> FIM SUB Cria_Combo_codigo_pagamentoVisaElectron
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_Combo_TipoVisa
' - Monta as opções dos tipos de transações da Visanet a ser usada no site (VBV/MOSET).
' - Chamada no arquivo ADM_funcoes_pagamento.asp
'########################################################################################################
Sub Cria_Combo_TipoVisa(nome,opcao,adicional)

    Dim Valor(2), Tipo(2)

    Valor(1)="VISAVBV"
    Valor(2)="VISAMOSET"

    Tipo(1)="VBV"
    Tipo(2)="MOSET"
%>
    <SELECT NAME="<%=nome%>" class="FORMbox" <%= adicional %>>
<% 

    For I=1 to UBound(Valor)
        If opcao = Valor(i) Then    %>
            <OPTION SELECTED VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% Else %>
            <OPTION VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% End If
    Next 
%>
    </SELECT>
<% End Sub
'########################################################################################################
'--> FIM SUB Cria_Combo_TipoVisa
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_Combo_OpcaoParcela
' - Monta as opções dos tipos de parcelamento para as parcelas.
' - Chamada no arquivo ADM_funcoes_pagamento.asp
'########################################################################################################
Sub Cria_Combo_OpcaoParcela(nome,opcao)

    Dim Valor(3)

    Valor(1)="Desconto"
    Valor(2)="Sem juros"
    Valor(3)="Com juros"
%>
    <SELECT NAME="<%=nome%>" class="FORMbox">
<% 

    For I=1 to UBound(Valor)
        If opcao = Valor(i) Then    %>
            <OPTION SELECTED VALUE="<%= Valor(i) %>"><%= Valor(i) %></OPTION>		
        <% Else %>
            <OPTION VALUE="<%= Valor(i) %>"><%= Valor(i) %></OPTION>		
        <% End If
    Next 
%>
    </SELECT>
<% End Sub
'########################################################################################################
'--> FIM SUB Cria_Combo_OpcaoParcela
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Mostra_formularioPagto
' - Case de formulários de cada forma de pagamento, para configuração.
' - Chamada no arquivo ADM_config_pagamento.asp
'########################################################################################################
Sub Mostra_formularioPagto(VarNome_pagto)

    Select Case VarNome_pagto

    Case "Amex"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dAmexParcelado',this);" style="cursor:pointer;"></td>
        <td>Permite parcelamento?</td>
        <td><%Call Cria_Combo_opcao("permite_parcelamento",configuracao.getAttribute("permite_parcelamento"),"onchange=""define_parcelamento(this.value,document.getElementsByName('juros')[0].options[document.getElementsByName('juros')[0].selectedIndex].text,'parcelamento');""")%></td>
    </tr>
    <tr id="dAmexParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de parcelamento.</td>
    </tr>
    <tr id="tblTipoParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipoParcelado',this);" style="cursor:pointer;"></td>
        <td>Tipo de parcelamento</td>
        <td><%Call Cria_Combo_juros_parcelado("juros",configuracao.getAttribute("juros"),"onchange=""define_parcelamento(document.getElementsByName('permite_parcelamento')[0].options[document.getElementsByName('permite_parcelamento')[0].selectedIndex].value,this.options[this.selectedIndex].text,'tipoParcelamento');""","Composto")%></td>
    </tr>
    <tr id="dTipoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo de parcelamento que será aplicado nas transações parceladas. Sendo "Juros do Emissor" com a taxa de juros do emissor do cartão do comprador e "Juros do Lojista" com a taxa de juros definida pelo lojista.</td>
    </tr>
    <tr id="tblTaxaDesc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercDesc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Desconto</td>
        <td><input type="text" size="5" name="taxa_desconto" value="<%=configuracao.getAttribute("taxa_desconto")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercDesc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de desconto que será aplicada ao valor total do pedido.</td>
    </tr>
    <tr id="tblTaxaAcresc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercAcresc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Acréscimo</td>
        <td><input type="text" size="5" name="taxa_juros" value="<%=configuracao.getAttribute("taxa_juros")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercAcresc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de acréscimo que será aplicado ao valor total do pedido.</td>
    </tr>
    <tr id="tblNumParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dNumParcelas',this);" style="cursor:pointer;"></td>
        <td>Número de Parcelas</td>
        <td><%Call Cria_Combo_Numeros("num_parcelas",configuracao.getAttribute("num_parcelas"),1,12,"onchange=""ajusta_exibeiframe(12,this.options[this.selectedIndex].value,'divparc')""")%>&nbsp;<span Onclick="mostraiframe('tblCondParc');" style="cursor:pointer;"><span id="divCondParc"><u>Clique e defina as condições de parcelamento</u></span></span></td>
    </tr>
    <tr id="dNumParcelas" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número máximo de parcelas permitido. Quando utilizado o tipo de parcelamento "Juros do Lojista" é possível a configuração das ações aplicadas em cada tipo de parcelamento.</td>
    </tr>
    <tr id="tblCondParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dCondParc',this);" style="cursor:pointer;"></td>
        <td height="30">Condições de Parcelamento</td>
        <td>
            <span id="divparc1">01&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc1",configuracao.getAttribute("parc1"))%><br></span>
            <span id="divparc2">02&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc2",configuracao.getAttribute("parc2"))%><br></span>
            <span id="divparc3">03&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc3",configuracao.getAttribute("parc3"))%><br></span>
            <span id="divparc4">04&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc4",configuracao.getAttribute("parc4"))%><br></span>
            <span id="divparc5">05&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc5",configuracao.getAttribute("parc5"))%><br></span>
            <span id="divparc6">06&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc6",configuracao.getAttribute("parc6"))%><br></span>
            <span id="divparc7">07&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc7",configuracao.getAttribute("parc7"))%><br></span>
            <span id="divparc8">08&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc8",configuracao.getAttribute("parc8"))%><br></span>
            <span id="divparc9">09&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc9",configuracao.getAttribute("parc9"))%><br></span>
            <span id="divparc10">10&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc10",configuracao.getAttribute("parc10"))%><br></span>
            <span id="divparc11">11&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc11",configuracao.getAttribute("parc11"))%><br></span>
            <span id="divparc12">12&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc12",configuracao.getAttribute("parc12"))%></span>
        </td>
    </tr>
    <tr id="dCondParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina o tipo de ação efetuado para cada forma de parcelamento. <br>Sendo: <br>- "Desconto": Será aplicado o percentual de desconto definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas. <br>- "Sem Juros": Será dividido o valor total do pedido pelo respectivo número de parcelas, sem acréscimo ou desconto. <br>- "Com Juros": Será aplicado o percentual de acréscimo definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas.</td>
    </tr>
    <tr id="tblValorMinParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dValorMinParc',this);" style="cursor:pointer;"></td>
        <td>Valor mínimo por parcela</td>
        <td><input type="text" size="20" name="valormin_parcela" value="<%=configuracao.getAttribute("valormin_parcela")%>" class="FORMbox" Onblur="fncPreencheValue(this, '0,00')"></td>
    </tr>
    <tr id="dValorMinParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o valor mínimo permitido para cada parcela.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dAmexLocale',this);" style="cursor:pointer;"></td>
        <td>Idioma</td>
        <td><input type="text" size="20" name="AmexLocale" value="<%=configuracao.getAttribute("AmexLocale")%>" class="FORMbox"></td>
    </tr>
    <tr id="dAmexLocale" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Idioma do pais origem da loja, usar os padrões pt_BR, etc...</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "ABNCDC"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVAR01',this);" style="cursor:pointer;"></td>
        <td>Número da Loja?</td>
        <td><input type="text" size="20" name="VAR01" value="<%=configuracao.getAttribute("VAR01")%>" class="FORMbox"></td>
    </tr>
    <tr id="dVAR01" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Numero da Loja junto a financeira.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVAR02',this);" style="cursor:pointer;"></td>
        <td>Número do Serviço?</td>
        <td><input type="text" size="20" name="VAR02" value="<%=configuracao.getAttribute("VAR02")%>" class="FORMbox"></td>
    </tr>
    <tr id="dVAR02" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Número do serviço contratado junto a financeira.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVAR21',this);" style="cursor:pointer;"></td>
        <td>Número da tabela de financiamento?</td>
        <td><input type="text" size="20" name="VAR21" value="<%=configuracao.getAttribute("VAR21")%>" class="FORMbox"></td>
    </tr>
    <tr id="dVAR21" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Número que identifica a tabela de financiamento a ser utilizada.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Brasil"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBBConvenio',this);" style="cursor:pointer;"></td>
        <td>Convênio da Loja (RCB)?</td>
        <td><input type="text" size="20" name="BBConvenio" value="<%=configuracao.getAttribute("BBConvenio")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBBConvenio" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Código de convênio para Meios Eletrônicos de Pagamentos do Banco do Brasil.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBBCodCobranca',this);" style="cursor:pointer;"></td>
        <td>Código de cobrança (CBR) ?</td>
        <td><input type="text" size="20" name="BBCodCobranca" value="<%=configuracao.getAttribute("BBCodCobranca")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBBCodCobranca" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Código de cobrança para emissão de boletos bancários.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBBFormatoRetorno',this);" style="cursor:pointer;"></td>
        <td>Tipo de pagamento disponível</td>
        <td><%Call lista_TipPagBB("BBTipoPagamento",configuracao.getAttribute("BBTipoPagamento"))%></td>
    </tr>
    <tr id="dBBFormatoRetorno" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipos de pagamentos disponíveis no ambiente BB Office Banking para escolha do usuário. A opção "Todas opções" disponibilizará apenas as formas de pagamentos contratadas e liberadas no contrato junto ao Banco. </td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBBDiasdeVencimento',this);" style="cursor:pointer;"></td>
        <td>Dias de vencimento?</td>
        <td><input type="text" size="20" name="BBDiasdeVencimento" value="<%=configuracao.getAttribute("BBDiasdeVencimento")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBBDiasdeVencimento" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o vencimento do boleto. O numero informado aqui será somado ao dia da geração do boleto.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBBComentario',this);" style="cursor:pointer;"></td>
        <td>Linha de instrução</td>
        <td><input type="text" size="20" name="BBComentario" value="<%=configuracao.getAttribute("BBComentario")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBBComentario" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Linha de instrução para o boleto bancário.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Boleto"
%>

    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipoBoleto',this);" style="cursor:pointer;"></td>
        <td>Tipo de boleto</td>
        <td><%Call Cria_Combo_Tipo_Boleto("BoletoTipo",configuracao.getAttribute("BoletoTipo"))%></td>
    </tr>
    <tr id="dTipoBoleto" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo boleto configurado, BoletoGenerico, BoletoItau, BoletoBradesco, BoletoBancoBrasil. Detalhes, veja na documentação da loja.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dExibirBotoes',this);" style="cursor:pointer;"></td>
        <td>Exibir botões?</td>
        <td><%Call MontaCombo_opcaoNum("botoesboleto",configuracao.getAttribute("botoesboleto"))%></td>
    </tr>
    <tr id="dExibirBotoes" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Exibe botões de IMPRIMIR e FECHAR JANELA na página do boleto.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTituloPagina',this);" style="cursor:pointer;"></td>
        <td>Título da página</td>
        <td><input type="text" size="20" name="titulo_boleto" value="<%=configuracao.getAttribute("titulo_boleto")%>" class="FORMbox"></td>
    </tr>
    <tr id="dTituloPagina" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Titulo no navegador na página de exibição do boleto.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dDiasdeVencimento',this);" style="cursor:pointer;"></td>
        <td>Dias de vencimento</td>
        <td><input type="text" size="20" name="DiasdeVencimento" value="<%=configuracao.getAttribute("DiasdeVencimento")%>" class="FORMbox"></td>
    </tr>
    <tr id="dDiasdeVencimento" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define quantos dias terá o  vencimento do boleto após sua emissão. O numero informado aqui será somado ao dia da geração do boleto. Para definir o vencimento do Boleto como "Contra Apresentação" utilize o valor "ca".</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dInstrucoesboleto1',this);" style="cursor:pointer;"></td>
        <td>Linha de instrução 1</td>
        <td><input type="text" size="20" name="instrucoesboleto1" value="<%=configuracao.getAttribute("instrucoesboleto1")%>" class="FORMbox"></td>
    </tr>
    <tr id="dInstrucoesboleto1" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe aqui a 1ª linha de instrução do boleto.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dInstrucoesboleto2',this);" style="cursor:pointer;"></td>
        <td>Linha de instrução 2</td>
        <td><input type="text" size="20" name="instrucoesboleto2" value="<%=configuracao.getAttribute("instrucoesboleto2")%>" class="FORMbox"></td>
    </tr>
    <tr id="dInstrucoesboleto2" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe aqui a 2ª linha de instrução do boleto.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dInstrucoesboleto3',this);" style="cursor:pointer;"></td>
        <td>Linha de instrução 3</td>
        <td><input type="text" size="20" name="instrucoesboleto3" value="<%=configuracao.getAttribute("instrucoesboleto3")%>" class="FORMbox"></td>
    </tr>
    <tr id="dInstrucoesboleto3" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe aqui a 3ª linha de instrução do boleto.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dInstrucoesboleto4',this);" style="cursor:pointer;"></td>
        <td>Linha de instrução 4</td>
        <td><input type="text" size="20" name="instrucoesboleto4" value="<%=configuracao.getAttribute("instrucoesboleto4")%>" class="FORMbox"></td>
    </tr>
    <tr id="dInstrucoesboleto4" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe aqui a 4ª linha de instrução do boleto.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dInstrucoesboleto5',this);" style="cursor:pointer;"></td>
        <td>Linha de instrução 5</td>
        <td><input type="text" size="20" name="instrucoesboleto5" value="<%=configuracao.getAttribute("instrucoesboleto5")%>" class="FORMbox"></td>
    </tr>
    <tr id="dInstrucoesboleto5" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe aqui a 5ª linha de instrução do boleto.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Bradesco"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoShopFacil',this);" style="cursor:pointer;"></td>
        <td>Afiliada ao Bradesco ShopFácil</td>
        <td><%Call MontaCombo_opcaoNum("BradescoShopFacil",configuracao.getAttribute("BradescoShopFacil"))%></td>
    </tr>
    <tr id="dBradescoShopFacil" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Indica se a loja é afiliada ou não ao shopping shopfacil.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoPagFacil',this);" style="cursor:pointer;"></td>
        <td>Ativar Bradesco Pagamento Fácil - Cartões?</td>
        <td><%Call MontaCombo_opcaoNum("BradescoPagFacil",configuracao.getAttribute("BradescoPagFacil"))%></td>
    </tr>
    <tr id="dBradescoPagFacil" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de Pagamento Facil Bradesco.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoTransfer',this);" style="cursor:pointer;"></td>
        <td>Ativar Tranferência entre contas Bradesco?</td>
        <td><%Call MontaCombo_opcaoNum("BradescoTransfer",configuracao.getAttribute("BradescoTransfer"))%></td>
    </tr>
    <tr id="dBradescoTransfer" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de pagamento por transferência entre contas Bradesco.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoFinanciamento',this);" style="cursor:pointer;"></td>
        <td>Ativar Financiamento Bradesco?</td>
        <td><%Call MontaCombo_opcaoNum("BradescoFinanciamento",configuracao.getAttribute("BradescoFinanciamento"))%></td>
    </tr>
    <tr id="dBradescoFinanciamento" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de pagamento de financiamento para clientes Bradesco.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoLoja',this);" style="cursor:pointer;"></td>
        <td>Código da Loja?</td>
        <td><input type="text" size="20" name="BradescoLoja" value="<%=configuracao.getAttribute("BradescoLoja")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoLoja" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Código da loja no sistema BradescoNet.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoRazaoSocial',this);" style="cursor:pointer;"></td>
        <td>Razão Social?</td>
        <td><input type="text" size="20" name="BradescoRazaoSocial" value="<%=configuracao.getAttribute("BradescoRazaoSocial")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoRazaoSocial" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Razão Social conforme cadastrado junto ao banco.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoAgencia',this);" style="cursor:pointer;"></td>
        <td>Código da Agência?</td>
        <td><input type="text" size="20" name="BradescoAgencia" value="<%=configuracao.getAttribute("BradescoAgencia")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoAgencia" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Código da agência bancária.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoCodigoCedente',this);" style="cursor:pointer;"></td>
        <td>Código de Cedente?</td>
        <td><input type="text" size="20" name="BradescoCodigoCedente" value="<%=configuracao.getAttribute("BradescoCodigoCedente")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoCodigoCedente" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Código do cedente junto ao banco.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoAssinaturaBoleto',this);" style="cursor:pointer;"></td>
        <td>Assinatura Boleto</td>
        <td><input type="text" size="20" name="BradescoAssinaturaBoleto" value="<%=configuracao.getAttribute("BradescoAssinaturaBoleto")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoAssinaturaBoleto" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Assinatura para geração de boleto fornecido pelo banco.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoAssinaturaTransfer',this);" style="cursor:pointer;"></td>
        <td>Assinatura Transferência?</td>
        <td><input type="text" size="20" name="BradescoAssinaturaTransfer" value="<%=configuracao.getAttribute("BradescoAssinaturaTransfer")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoAssinaturaTransfer" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Assinatura para transferência bancária fornecido pelo banco.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoTaxaBoleto',this);" style="cursor:hand;"></td>
        <td width="250">Cobrar taxa de emissão do boleto?</td>
        <td width="350" align="left"><%Call Cria_Combo_opcao("BradescoTaxaBoleto",configuracao.getAttribute("BradescoTaxaBoleto"),"")%></td>
    </tr>
    <tr id="dBradescoTaxaBoleto" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a cobrança da taxa de boleto.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoValorTaxaBoleto',this);" style="cursor:hand;"></td>
        <td width="250">Valor da taxa de emissão do boleto:</td>
        <td width="350" align="left"><input type="text" size="10" name="BradescoValorTaxaBoleto" value="<%=configuracao.getAttribute("BradescoValorTaxaBoleto")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoValorTaxaBoleto" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Valor a ser adicionado ao total da compra como taxa de boleto.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPrazoMaxBradescoDebito',this);" style="cursor:pointer;"></td>
        <td>Prazo máximo para debito</td>
        <td><input type="text" size="20" name="PrazoMaxBradescoDebito" value="<%=configuracao.getAttribute("PrazoMaxBradescoDebito")%>" class="FORMbox"></td>
    </tr>
    <tr id="dPrazoMaxBradescoDebito" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Prazo máximo para debito do valor em conta.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dBradescoDiasdeVencimento',this);" style="cursor:pointer;"></td>
        <td>Dias de vencimento?</td>
        <td><input type="text" size="20" name="BradescoDiasdeVencimento" value="<%=configuracao.getAttribute("BradescoDiasdeVencimento")%>" class="FORMbox"></td>
    </tr>
    <tr id="dBradescoDiasdeVencimento" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o vencimento do boleto. O numero informado aqui será somado ao dia da geração do boleto.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
    <script>
        //Seta as mascaras nos inputs
        var decimalSeparator = ",";
        var groupSeparator = ".";

        var numParserValor = new NumberParser(2, decimalSeparator, groupSeparator, true);
        numParserValor.currencySymbol = ""
        numParserValor.useCurrency = true;
        numParserValor.currencyInside = true;
        var numMaskValor = new NumberMask(numParserValor, "BradescoValorTaxaBoleto", 6);
    </script>
<%
    Case "CobreBem"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_4',this);" style="cursor:pointer;"></td>
        <td>UsuarioBoleto</td>
        <td><input type="text" size="20" name="UsuarioBoleto" value="<%=configuracao.getAttribute("UsuarioBoleto")%>" class="FORMbox"></td>
    </tr>
    <tr id="dconfpag_4" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Este parâmetro identifica a conta corrente de cobrança a ser utilizada para a geração do boleto..</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_5',this);" style="cursor:pointer;"></td>
        <td>CSID</td>
        <td><input type="text" size="20" name="CSID" value="<%=configuracao.getAttribute("CSID")%>" class="FORMbox"></td>
    </tr>
    <tr id="dconfpag_5" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Este parâmetro identifica o usuário que administra a conta corrente de cobrança a ser utilizada para a<br> geração do boleto. Esta informação será fornecidade pela LocaWeb quando solicitar a configuração do boleto<br> CobreBem.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dDiasdeVencimento',this);" style="cursor:pointer;"></td>
        <td>Dias de vencimento</td>
        <td><input type="text" size="20" name="DiasdeVencimento" value="<%=configuracao.getAttribute("DiasdeVencimento")%>" class="FORMbox"></td>
    </tr>
    <tr id="dDiasdeVencimento" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define quantos dias terá o  vencimento do boleto após sua emissão. O numero informado aqui será somado ao dia da geração do boleto.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_6',this);" style="cursor:pointer;"></td>
        <td colspan="2" height="30">Instruções para o Caixa.</td>
    </tr>
    <tr id="dconfpag_6" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe neste parâmetro o código HTML que será exibido nas instruções para o caixa na ficha de compensação do boleto.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="InstrucoesCaixaCedente" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("InstrucoesCaixaCedente")%></textarea></td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Deposito"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dDepositoCorrentista',this);" style="cursor:pointer;"></td>
        <td>Nome do Correntista?</td>
        <td><input type="text" size="20" name="DepositoCorrentista" value="<%=configuracao.getAttribute("DepositoCorrentista")%>" class="FORMbox"></td>
    </tr>
    <tr id="dDepositoCorrentista" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Nome do favorecido.</td>
    </tr>

    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dDepositoEnvioCompEmail',this);" style="cursor:pointer;"></td>
        <td>Email: </td>
        <td><input type="text" size="20" name="DepositoEnvioCompEmail" value="<%=configuracao.getAttribute("DepositoEnvioCompEmail")%>" class="FORMbox"></td>
    </tr>
    <tr id="dDepositoEnvioCompEmail" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> E-mail para onde deve ser enviado o comprovante de pagamento.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dDepositoEnvioCompFax',this);" style="cursor:pointer;"></td>
        <td>Fax: </td>
        <td><input type="text" size="20" name="DepositoEnvioCompFax" value="<%=configuracao.getAttribute("DepositoEnvioCompFax")%>" class="FORMbox"></td>
    </tr>
    <tr id="dDepositoEnvioCompFax" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> FAX para onde deve ser enviado o comprovante de pagamento..</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30"><b>Informe abaixo os dados para depósito</b></td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="DadosDeposito" ROWS="10" COLS="100%" class="FORMbox"><%=Replace(configuracao.getAttribute("DadosDeposito"),"<br>",vbCrLf)%></textarea></td>
    </tr>
    <tr id="dDadosDeposito" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe os dados da conta bancária.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30"><b>Breve descrição desta opção de pagamento</b></td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Finasa"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_3',this);" style="cursor:pointer;"></td>
        <td>Loja</td>
        <td><input type="text" size="20" name="loja" value="<%=configuracao.getAttribute("loja")%>" class="FORMbox"></td>
    </tr>
    <tr id="dconfpag_3" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Número da loja junto à financeira.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_4',this);" style="cursor:pointer;"></td>
        <td>Filial</td>
        <td><input type="text" size="20" name="filial" value="<%=configuracao.getAttribute("filial")%>" class="FORMbox"></td>
    </tr>
    <tr id="dconfpag_4" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Número da filial da loja junto à financeira.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_5',this);" style="cursor:pointer;"></td>
        <td>Senha</td>
        <td><input type="text" size="20" name="senha" value="<%=configuracao.getAttribute("senha")%>" class="FORMbox"></td>
    </tr>
    <tr id="dconfpag_5" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Senha de acesso ao sistema da financeira.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_6',this);" style="cursor:pointer;"></td>
        <td>Action</td>
        <td><input type="text" size="20" name="action" value="<%=configuracao.getAttribute("action")%>" class="FORMbox"></td>
    </tr>
    <tr id="dconfpag_6" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ação para as transações.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_7',this);" style="cursor:pointer;"></td>
        <td>Tipo Usuário</td>
        <td><input type="text" size="20" name="tipoUsuario" value="<%=configuracao.getAttribute("tipoUsuario")%>" class="FORMbox"></td>
    </tr>
    <tr id="dconfpag_7" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo de usuário para as transações. Valor fixo "simulação" para teste.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Itau"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dOBSItau',this);" style="cursor:pointer;"></td>
        <td>Linha de instrução</td>
        <td><input type="text" size="20" name="OBSItau" value="<%=configuracao.getAttribute("OBSItau")%>" class="FORMbox"></td>
    </tr>
    <tr id="dOBSItau" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Linha de instrução no boleto bancário.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dItauDiasdeVencimento',this);" style="cursor:pointer;"></td>
        <td>Dias de vencimento?</td>
        <td><input type="text" size="20" name="ItauDiasdeVencimento" value="<%=configuracao.getAttribute("ItauDiasdeVencimento")%>" class="FORMbox"></td>
    </tr>
    <tr id="dItauDiasdeVencimento" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o vencimento do boleto. O numero informado aqui será somado ao dia da geração do boleto.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Mastercard"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dnum_afiliacao',this);" style="cursor:pointer;"></td>
        <td>Número de Afiliação</td>
        <td><input type="text" size="20" name="RedeCardFiliacao" value="<%=configuracao.getAttribute("RedeCardFiliacao")%>" class="FORMbox"></td>
    </tr>
    <tr id="dnum_afiliacao" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número de afiliação Redecard de seu estabelecimento.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dpermite_cartoesestrangeiros',this);" style="cursor:pointer;"></td>
        <td>Permite Cartões estrangeiros?</td>
        <td><%Call Cria_Combo_opcao("permite_cartoesestrangeiros",configuracao.getAttribute("permite_cartoesestrangeiros"),"")%></td>
    </tr>
    <tr id="dpermite_cartoesestrangeiros" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"><img src="images/hr_l.gif">Define se o site aceitará compras realizadas com cartões Mastercard emitidos fora do Brasil de abrangência internacional.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dRedeCardParcelado',this);" style="cursor:pointer;"></td>
        <td>Permite Parcelamento?</td>
        <td><%Call Cria_Combo_opcao("permite_parcelamento",configuracao.getAttribute("permite_parcelamento"),"onchange=""define_parcelamento(this.value,document.getElementsByName('juros')[0].options[document.getElementsByName('juros')[0].selectedIndex].text,'parcelamento');""")%></td>
    </tr>
    <tr id="dRedeCardParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de parcelamento.</td>
    </tr>
    <tr class="Linha2Tabela" id="tblTipoParc" style="display:none;" >
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipoParcelado',this);" style="cursor:pointer;"></td>
        <td>Tipo de parcelamento</td>
        <td><%Call Cria_Combo_juros_parcelado("juros",configuracao.getAttribute("juros"),"onchange=""define_parcelamento(document.getElementsByName('permite_parcelamento')[0].options[document.getElementsByName('permite_parcelamento')[0].selectedIndex].value,this.options[this.selectedIndex].text,'tipoParcelamento');""","Composto")%></td>
    </tr>
    <tr id="dTipoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo de parcelamento que será aplicado nas transações parceladas. Sendo "Juros do Emissor" com a taxa de juros do emissor do cartão do comprador e "Juros do Lojista" com a taxa de juros definida pelo lojista.</td>
    </tr>
    <tr id="tblTaxaDesc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercDesc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Desconto</td>
        <td><input type="text" size="5" name="taxa_desconto" value="<%=configuracao.getAttribute("taxa_desconto")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercDesc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de desconto que será aplicada ao valor total do pedido.</td>
    </tr>
    <tr id="tblTaxaAcresc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercAcresc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Acréscimo</td>
        <td><input type="text" size="5" name="taxa_juros" value="<%=configuracao.getAttribute("taxa_juros")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercAcresc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de acréscimo que será aplicado ao valor total do pedido.</td>
    </tr>
    <tr id="tblNumParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dNumParcelas',this);" style="cursor:pointer;"></td>
        <td>Número de Parcelas</td>
        <td><%Call Cria_Combo_Numeros("num_parcelas",configuracao.getAttribute("num_parcelas"),1,12,"onchange=""ajusta_exibeiframe(12,this.options[this.selectedIndex].value,'divparc')""")%>&nbsp;<span Onclick="mostraiframe('tblCondParc');" style="cursor:pointer;"><span id="divCondParc"><u>Clique e defina as condições de parcelamento</u></span></span></td>
    </tr>
    <tr id="dNumParcelas" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número máximo de parcelas permitido. Quando utilizado o tipo de parcelamento "Juros do Lojista" é possível a configuração das ações aplicadas em cada tipo de parcelamento.</td>
    </tr>
    <tr id="tblCondParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dCondParc',this);" style="cursor:pointer;"></td>
        <td height="30">Condições de Parcelamento</td>
        <td>
            <span id="divparc1">01&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc1",configuracao.getAttribute("parc1"))%><br></span>
            <span id="divparc2">02&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc2",configuracao.getAttribute("parc2"))%><br></span>
            <span id="divparc3">03&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc3",configuracao.getAttribute("parc3"))%><br></span>
            <span id="divparc4">04&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc4",configuracao.getAttribute("parc4"))%><br></span>
            <span id="divparc5">05&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc5",configuracao.getAttribute("parc5"))%><br></span>
            <span id="divparc6">06&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc6",configuracao.getAttribute("parc6"))%><br></span>
            <span id="divparc7">07&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc7",configuracao.getAttribute("parc7"))%><br></span>
            <span id="divparc8">08&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc8",configuracao.getAttribute("parc8"))%><br></span>
            <span id="divparc9">09&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc9",configuracao.getAttribute("parc9"))%><br></span>
            <span id="divparc10">10&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc10",configuracao.getAttribute("parc10"))%><br></span>
            <span id="divparc11">11&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc11",configuracao.getAttribute("parc11"))%><br></span>
            <span id="divparc12">12&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc12",configuracao.getAttribute("parc12"))%></span>
        </td>
    </tr>
    <tr id="dCondParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina o tipo de ação efetuado para cada forma de parcelamento. <br>Sendo: <br>- "Desconto": Será aplicado o percentual de desconto definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas. <br>- "Sem Juros": Será dividido o valor total do pedido pelo respectivo número de parcelas, sem acréscimo ou desconto. <br>- "Com Juros": Será aplicado o percentual de acréscimo definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas.</td>
    </tr>
    <tr id="tblValorMinParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dValorMinParc',this);" style="cursor:pointer;"></td>
        <td>Valor mínimo por parcela</td>
        <td><input type="text" size="20" name="valormin_parcela" value="<%=configuracao.getAttribute("valormin_parcela")%>" class="FORMbox" Onblur="fncPreencheValue(this, '0,00')"></td>
    </tr>
    <tr id="dValorMinParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o valor mínimo permitido para cada parcela.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dRedeCardAVS',this);" style="cursor:pointer;"></td>
        <td>Ativar AVS?</td>
        <td><%Call MontaCombo_opcaoNum("RedeCardAVS",configuracao.getAttribute("RedeCardAVS"))%></td>
    </tr>
    <tr id="dRedeCardAVS" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de AVS no site.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Diners"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dnum_afiliacao',this);" style="cursor:pointer;"></td>
        <td>Número de Afiliação</td>
        <td><input type="text" size="20" name="RedeCardFiliacao" value="<%=configuracao.getAttribute("RedeCardFiliacao")%>" class="FORMbox"></td>
    </tr>
    <tr id="dnum_afiliacao" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número de afiliação Redecard de seu estabelecimento.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dpermite_cartoesestrangeiros',this);" style="cursor:pointer;"></td>
        <td>Permite Cartões estrangeiros?</td>
        <td><%Call Cria_Combo_opcao("permite_cartoesestrangeiros",configuracao.getAttribute("permite_cartoesestrangeiros"),"")%></td>
    </tr>
    <tr id="dpermite_cartoesestrangeiros" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"><img src="images/hr_l.gif">Define se o site aceitará compras realizadas com cartões Diners emitidos fora do Brasil de abrangência internacional.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dRedeCardParcelado',this);" style="cursor:pointer;"></td>
        <td>Permite Parcelamento?</td>
        <td><%Call Cria_Combo_opcao("permite_parcelamento",configuracao.getAttribute("permite_parcelamento"),"onchange=""define_parcelamento(this.value,document.getElementsByName('juros')[0].options[document.getElementsByName('juros')[0].selectedIndex].text,'parcelamento');""")%></td>
    </tr>
    <tr id="dRedeCardParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de parcelamento.</td>
    </tr>
    <tr id="tblTipoParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipoParcelado',this);" style="cursor:pointer;"></td>
        <td>Tipo de parcelamento</td>
        <td><%Call Cria_Combo_juros_parcelado("juros",configuracao.getAttribute("juros"),"onchange=""define_parcelamento(document.getElementsByName('permite_parcelamento')[0].options[document.getElementsByName('permite_parcelamento')[0].selectedIndex].value,this.options[this.selectedIndex].text,'tipoParcelamento');""","Composto")%></td>
    </tr>
    <tr id="dTipoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo de parcelamento que será aplicado nas transações parceladas. Sendo "Juros do Emissor" com a taxa de juros do emissor do cartão do comprador e "Juros do Lojista" com a taxa de juros definida pelo lojista.</td>
    </tr>
    <tr id="tblTaxaDesc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercDesc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Desconto</td>
        <td><input type="text" size="5" name="taxa_desconto" value="<%=configuracao.getAttribute("taxa_desconto")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercDesc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de desconto que será aplicada ao valor total do pedido.</td>
    </tr>
    <tr id="tblTaxaAcresc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercAcresc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Acréscimo</td>
        <td><input type="text" size="5" name="taxa_juros" value="<%=configuracao.getAttribute("taxa_juros")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercAcresc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de acréscimo que será aplicado ao valor total do pedido.</td>
    </tr>
    <tr id="tblNumParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dNumParcelas',this);" style="cursor:pointer;"></td>
        <td>Número de Parcelas</td>
        <td><%Call Cria_Combo_Numeros("num_parcelas",configuracao.getAttribute("num_parcelas"),1,12,"onchange=""ajusta_exibeiframe(12,this.options[this.selectedIndex].value,'divparc')""")%>&nbsp;<span Onclick="mostraiframe('tblCondParc');" style="cursor:pointer;"><span id="divCondParc"><u>Clique e defina as condições de parcelamento</u></span></span></td>
    </tr>
    <tr id="dNumParcelas" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número máximo de parcelas permitido. Quando utilizado o tipo de parcelamento "Juros do Lojista" é possível a configuração das ações aplicadas em cada tipo de parcelamento.</td>
    </tr>
    <tr id="tblCondParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dCondParc',this);" style="cursor:pointer;"></td>
        <td height="30">Condições de Parcelamento</td>
        <td>
            <span id="divparc1">01&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc1",configuracao.getAttribute("parc1"))%><br></span>
            <span id="divparc2">02&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc2",configuracao.getAttribute("parc2"))%><br></span>
            <span id="divparc3">03&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc3",configuracao.getAttribute("parc3"))%><br></span>
            <span id="divparc4">04&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc4",configuracao.getAttribute("parc4"))%><br></span>
            <span id="divparc5">05&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc5",configuracao.getAttribute("parc5"))%><br></span>
            <span id="divparc6">06&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc6",configuracao.getAttribute("parc6"))%><br></span>
            <span id="divparc7">07&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc7",configuracao.getAttribute("parc7"))%><br></span>
            <span id="divparc8">08&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc8",configuracao.getAttribute("parc8"))%><br></span>
            <span id="divparc9">09&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc9",configuracao.getAttribute("parc9"))%><br></span>
            <span id="divparc10">10&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc10",configuracao.getAttribute("parc10"))%><br></span>
            <span id="divparc11">11&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc11",configuracao.getAttribute("parc11"))%><br></span>
            <span id="divparc12">12&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc12",configuracao.getAttribute("parc12"))%></span>
        </td>
    </tr>
    <tr id="dCondParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina o tipo de ação efetuado para cada forma de parcelamento. <br>Sendo: <br>- "Desconto": Será aplicado o percentual de desconto definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas. <br>- "Sem Juros": Será dividido o valor total do pedido pelo respectivo número de parcelas, sem acréscimo ou desconto. <br>- "Com Juros": Será aplicado o percentual de acréscimo definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas.</td>
    </tr>
    <tr id="tblValorMinParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dValorMinParc',this);" style="cursor:pointer;"></td>
        <td>Valor mínimo por parcela</td>
        <td><input type="text" size="20" name="valormin_parcela" value="<%=configuracao.getAttribute("valormin_parcela")%>" class="FORMbox" Onblur="fncPreencheValue(this, '0,00')"></td>
    </tr>
    <tr id="dValorMinParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o valor mínimo permitido para cada parcela.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dRedeCardAVS',this);" style="cursor:pointer;"></td>
        <td>Ativar AVS?</td>
        <td><%Call MontaCombo_opcaoNum("RedeCardAVS",configuracao.getAttribute("RedeCardAVS"))%></td>
    </tr>
    <tr id="dRedeCardAVS" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de AVS no site.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Visa"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisanetID',this);" style="cursor:pointer;"></td>
        <td>Número de Afiliação</td>
        <td><input type="text" size="20" name="VisanetID" value="<%=configuracao.getAttribute("VisanetID")%>" class="FORMbox"></td>
    </tr>
    <tr id="dVisanetID" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número de afiliação Visanet de seu estabelecimento.</td>
    </tr>
    <tr class="Linha2Tabela" id="tblVisanetAuthentType">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisanetAuthentType',this);" style="cursor:pointer;"></td>
        <td>Usar autenticação bancária?</td>
        <td><%Call MontaCombo_opcaoNum("VisanetAuthentType",configuracao.getAttribute("VisanetAuthentType"))%></td>
    </tr>
    <tr id="dVisanetAuthentType" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define se a transação para ser aprovada deve ser autenticada pelo Banco. Válido atualmente apenas para o Bradesco, demais bancos as transações ocorrem normalmente.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisanetParcelado',this);" style="cursor:pointer;"></td>
        <td width="200">Permite Parcelamento?</td>
        <td><%Call Cria_Combo_opcao("permite_parcelamento",configuracao.getAttribute("permite_parcelamento"),"onchange=""define_parcelamento(this.value,document.getElementsByName('juros')[0].options[document.getElementsByName('juros')[0].selectedIndex].text,'parcelamento');""")%></td>
    </tr>
    <tr id="dVisanetParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de parcelamento.</td>
    </tr>
    <tr id="tblTipoParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipoParcelado',this);" style="cursor:pointer;"></td>
        <td>Tipo de parcelamento</td>
        <td><%Call Cria_Combo_juros_parcelado("juros",configuracao.getAttribute("juros"),"onchange=""define_parcelamento(document.getElementsByName('permite_parcelamento')[0].options[document.getElementsByName('permite_parcelamento')[0].selectedIndex].value,this.options[this.selectedIndex].text,'tipoParcelamento');""","Composto")%></td>
    </tr>
    <tr id="dTipoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo de parcelamento que será aplicado nas transações parceladas. Sendo "Juros do Emissor" com a taxa de juros do emissor do cartão do comprador e "Juros do Lojista" com a taxa de juros definida pelo lojista.</td>
    </tr>
    <tr id="tblTaxaDesc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercDesc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Desconto</td>
        <td><input type="text" size="5" name="taxa_desconto" value="<%=configuracao.getAttribute("taxa_desconto")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercDesc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de desconto que será aplicada ao valor total do pedido.</td>
    </tr>
    <tr id="tblTaxaAcresc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercAcresc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Acréscimo</td>
        <td><input type="text" size="5" name="taxa_juros" value="<%=configuracao.getAttribute("taxa_juros")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercAcresc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de acréscimo que será aplicado ao valor total do pedido.</td>
    </tr>
    <tr id="tblNumParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dNumParcelas',this);" style="cursor:pointer;"></td>
        <td>Número de Parcelas</td>
        <td><%Call Cria_Combo_Numeros("num_parcelas",configuracao.getAttribute("num_parcelas"),1,12,"onchange=""ajusta_exibeiframe(12,this.options[this.selectedIndex].value,'divparc')""")%>&nbsp;<span Onclick="mostraiframe('tblCondParc');" style="cursor:pointer;"><span id="divCondParc"><u>Clique e defina as condições de parcelamento</u></span></span></td>
    </tr>
    <tr id="dNumParcelas" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número máximo de parcelas permitido. Quando utilizado o tipo de parcelamento "Juros do Lojista" é possível a configuração das ações aplicadas em cada tipo de parcelamento.</td>
    </tr>
    <tr id="tblCondParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dCondParc',this);" style="cursor:pointer;"></td>
        <td height="30">Condições de Parcelamento</td>
        <td>
            <span id="divparc1">01&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc1",configuracao.getAttribute("parc1"))%><br></span>
            <span id="divparc2">02&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc2",configuracao.getAttribute("parc2"))%><br></span>
            <span id="divparc3">03&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc3",configuracao.getAttribute("parc3"))%><br></span>
            <span id="divparc4">04&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc4",configuracao.getAttribute("parc4"))%><br></span>
            <span id="divparc5">05&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc5",configuracao.getAttribute("parc5"))%><br></span>
            <span id="divparc6">06&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc6",configuracao.getAttribute("parc6"))%><br></span>
            <span id="divparc7">07&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc7",configuracao.getAttribute("parc7"))%><br></span>
            <span id="divparc8">08&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc8",configuracao.getAttribute("parc8"))%><br></span>
            <span id="divparc9">09&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc9",configuracao.getAttribute("parc9"))%><br></span>
            <span id="divparc10">10&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc10",configuracao.getAttribute("parc10"))%><br></span>
            <span id="divparc11">11&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc11",configuracao.getAttribute("parc11"))%><br></span>
            <span id="divparc12">12&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc12",configuracao.getAttribute("parc12"))%></span>
        </td>
    </tr>
    <tr id="dCondParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina o tipo de ação efetuado para cada forma de parcelamento. <br>Sendo: <br>- "Desconto": Será aplicado o percentual de desconto definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas. <br>- "Sem Juros": Será dividido o valor total do pedido pelo respectivo número de parcelas, sem acréscimo ou desconto. <br>- "Com Juros": Será aplicado o percentual de acréscimo definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas.</td>
    </tr>
    <tr id="tblValorMinParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dValorMinParc',this);" style="cursor:pointer;"></td>
        <td>Valor mínimo por parcela</td>
        <td><input type="text" size="20" name="valormin_parcela" value="<%=configuracao.getAttribute("valormin_parcela")%>" class="FORMbox" Onblur="fncPreencheValue(this, '0,00')"></td>
    </tr>
    <tr id="dValorMinParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o valor mínimo permitido para cada parcela.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisanetTipo',this);" style="cursor:pointer;"></td>
        <td>Tipo Visanet?</td>
        <td><%Call Cria_Combo_TipoVisa("modulo",configuracao.getAttribute("modulo"),"onchange='verificaTipoVisa(this.value);'")%></td>
    </tr>
    <tr id="dVisanetTipo" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Atualmente a Visanet possui as soluções MOSET e VBV para transações na internet, selecione a opção contratada junto a operadora.</td>
    </tr>
    <tr class="Linha1Tabela" id="tblVisaNetAntiPopup">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisaNetAntiPopup',this);" style="cursor:pointer;"></td>
        <td>Antipoup</td>
        <td><%Call MontaCombo_opcaoNum("VisaNetAntiPopup",configuracao.getAttribute("VisaNetAntiPopup"))%></td>
    </tr>
    <tr id="dVisaNetAntiPopup" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define se a janela da Visa para captura dos dados do cartão será aberta automáticamente ou através de um clique num botão.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
	<script>
		verificaTipoVisa(document.getElementById('modulo').options[document.getElementById('modulo').selectedIndex].value);
	</script>
<%
    Case "VisaElectron"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisanetID',this);" style="cursor:pointer;"></td>
        <td>Número de Afiliação</td>
        <td><input type="text" size="20" name="VisanetID" value="<%=configuracao.getAttribute("VisanetID")%>" class="FORMbox"></td>
    </tr>
    <tr id="dVisanetID" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número de afiliação Visanet de seu estabelecimento.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisanetCodPagamento',this);" style="cursor:pointer;"></td>
        <td>Código de pagamento?</td>
        <td><%Call Cria_Combo_codigo_pagamentoVisaElectron("VisanetCodPagamento",configuracao.getAttribute("VisanetCodPagamento"))%></td>
    </tr>
    <tr id="dVisanetCodPagamento" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define se a janela da Visa para captura dos dados do cartão será aberta automáticamente ou através de um clique num botão.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dVisaNetAntiPopup',this);" style="cursor:pointer;"></td>
        <td>Antipoup</td>
        <td><%Call MontaCombo_opcaoNum("VisaNetAntiPopup",configuracao.getAttribute("VisaNetAntiPopup"))%></td>
    </tr>
    <tr id="dVisaNetAntiPopup" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define se a janela da Visa para captura dos dados do cartão será aberta automáticamente ou através de um clique num botão.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Unibanco"
%>
    <tr class="Linha1Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha1Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "PagamentoCerto"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dChaveVendedor',this);" style="cursor:pointer;"></td>
        <td>Chave do Vendedor</td>
        <td><input type="text" size="20" name="chaveVendedor" value="<%=configuracao.getAttribute("chaveVendedor")%>" class="FORMbox"></td>
    </tr>
    <tr id="dChaveVendedor" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Informe sua chave de vendedor (no formato Guide) junto ao Pagamento Certo.</td>
    </tr>
	<tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "Paggo"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPaggoParcelado',this);" style="cursor:pointer;"></td>
        <td>Permite parcelamento?</td>
        <td><%Call Cria_Combo_opcao("permite_parcelamento",configuracao.getAttribute("permite_parcelamento"),"onchange=""define_parcelamento(this.value,document.getElementsByName('juros')[0].options[document.getElementsByName('juros')[0].selectedIndex].text,'parcelamento');""")%></td>
    </tr>
    <tr id="dPaggoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de parcelamento.</td>
    </tr>
	<tr id="tblTipoParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipoParcelado',this);" style="cursor:pointer;"></td>
        <td>Tipo de parcelamento</td>
        <td><%Call Cria_Combo_juros_parcelado("juros",configuracao.getAttribute("juros"),"onchange=""define_parcelamento(document.getElementsByName('permite_parcelamento')[0].options[document.getElementsByName('permite_parcelamento')[0].selectedIndex].value,this.options[this.selectedIndex].text,'tipoParcelamento');""","Simples")%></td>
    </tr>
    <tr id="dTipoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo de parcelamento que será aplicado nas transações parceladas. Sendo "Juros do Emissor" com a taxa de juros do emissor do cartão do comprador e "Juros do Lojista" com a taxa de juros definida pelo lojista.</td>
    </tr>
    <tr id="tblTaxaDesc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercDesc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Desconto</td>
        <td><input type="text" size="5" name="taxa_desconto" value="<%=configuracao.getAttribute("taxa_desconto")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercDesc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de desconto que será aplicada ao valor total do pedido.</td>
    </tr>
    <tr id="tblTaxaAcresc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercAcresc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Acréscimo</td>
        <td><input type="text" size="5" name="taxa_juros" value="<%=configuracao.getAttribute("taxa_juros")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercAcresc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de acréscimo que será aplicado ao valor total do pedido.</td>
    </tr>
    <tr id="tblNumParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dNumParcelas',this);" style="cursor:pointer;"></td>
        <td>Número de Parcelas</td>
        <td><%Call Cria_Combo_Numeros("num_parcelas",configuracao.getAttribute("num_parcelas"),1,12,"onchange=""ajusta_exibeiframe(12,this.options[this.selectedIndex].value,'divparc')""")%>&nbsp;<span Onclick="mostraiframe('tblCondParc');" style="cursor:pointer;"><span id="divCondParc"><u>Clique e defina as condições de parcelamento</u></span></span></td>
    </tr>
    <tr id="dNumParcelas" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número máximo de parcelas permitido. Quando utilizado o tipo de parcelamento "Juros do Lojista" é possível a configuração das ações aplicadas em cada tipo de parcelamento.</td>
    </tr>
    <tr id="tblCondParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dCondParc',this);" style="cursor:pointer;"></td>
        <td height="30">Condições de Parcelamento</td>
        <td>
            <span id="divparc1">01&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc1",configuracao.getAttribute("parc1"))%><br></span>
            <span id="divparc2">02&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc2",configuracao.getAttribute("parc2"))%><br></span>
            <span id="divparc3">03&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc3",configuracao.getAttribute("parc3"))%><br></span>
            <span id="divparc4">04&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc4",configuracao.getAttribute("parc4"))%><br></span>
            <span id="divparc5">05&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc5",configuracao.getAttribute("parc5"))%><br></span>
            <span id="divparc6">06&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc6",configuracao.getAttribute("parc6"))%><br></span>
            <span id="divparc7">07&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc7",configuracao.getAttribute("parc7"))%><br></span>
            <span id="divparc8">08&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc8",configuracao.getAttribute("parc8"))%><br></span>
            <span id="divparc9">09&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc9",configuracao.getAttribute("parc9"))%><br></span>
            <span id="divparc10">10&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc10",configuracao.getAttribute("parc10"))%><br></span>
            <span id="divparc11">11&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc11",configuracao.getAttribute("parc11"))%><br></span>
            <span id="divparc12">12&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc12",configuracao.getAttribute("parc12"))%></span>
        </td>
    </tr>
    <tr id="dCondParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina o tipo de ação efetuado para cada forma de parcelamento. <br>Sendo: <br>- "Desconto": Será aplicado o percentual de desconto definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas. <br>- "Sem Juros": Será dividido o valor total do pedido pelo respectivo número de parcelas, sem acréscimo ou desconto. <br>- "Com Juros": Será aplicado o percentual de acréscimo definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas.</td>
    </tr>
    <tr id="tblValorMinParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dValorMinParc',this);" style="cursor:pointer;"></td>
        <td>Valor mínimo por parcela</td>
        <td><input type="text" size="20" name="valormin_parcela" value="<%=configuracao.getAttribute("valormin_parcela")%>" class="FORMbox" Onblur="fncPreencheValue(this, '0,00')"></td>
    </tr>
    <tr id="dValorMinParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o valor mínimo permitido para cada parcela.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
    Case "PagSeguro"
%>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dEmailCobranca',this);" style="cursor:pointer;"></td>
        <td>E-mail Cobrança</td>
        <td><input type="text" size="30" name="email_cobranca" value="<%=configuracao.getAttribute("email_cobranca")%>" class="FORMbox"  Onblur="fncPreencheValue(this, 0)">&nbsp;</td>
    </tr>
	<tr id="dEmailCobranca" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Substitua o e-mail suporte@lojamodelo.com.br pelo seu e-mail cadastrado no PagSeguro.</td>
	</tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dtoken',this);" style="cursor:pointer;"></td>
        <td>Token:</td>
        <td><input type="text" size="46" name="token" value="<%=configuracao.getAttribute("token")%>" class="FORMbox"></td>
    </tr>
    <tr id="dtoken" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Para retorno automático à loja após o Pagamento com os detalhes da transação. É necessário configurar no site PagSeguro no menu <b><i>Meus Dados - Retorno Automático</i></b>.</td>
    </tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipo',this);" style="cursor:pointer;"></td>
        <td>Tipo</td>
        <td><% Call MontaCombo_opcaoTipo("tipo",configuracao.getAttribute("tipo")) %></td>
    </tr>
	<tr id="dTipo" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Este é o valor que informa que você está usando o carrinho PagSeguro. Não é necessário alteração. Para usar carrinho próprio, o valor é "CP".</td>
	</tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dMoeda',this);" style="cursor:pointer;"></td>
        <td>Moeda</td>
        <td><input type="text" size="20" name="moeda" value="<%=configuracao.getAttribute("moeda")%>" class="FORMbox"  Onblur="fncPreencheValue(this, 0)">&nbsp;</td>
    </tr>
	<tr id="dMoeda" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif">Não é necessário alteração. Por enquanto, o PagSeguro aceita apenas pagamento em moeda brasileira (Real).</td>
	</tr>
    <tr class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPaggoParcelado',this);" style="cursor:pointer;"></td>
        <td>Permite parcelamento?</td>
        <td><%Call Cria_Combo_opcao("permite_parcelamento",configuracao.getAttribute("permite_parcelamento"),"onchange=""define_parcelamento(this.value,document.getElementsByName('juros')[0].options[document.getElementsByName('juros')[0].selectedIndex].text,'parcelamento');""")%></td>
    </tr>
    <tr id="dPaggoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Ativa ou desativa a opção de parcelamento.</td>
    </tr>
	<tr id="tblTipoParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dTipoParcelado',this);" style="cursor:pointer;"></td>
        <td>Tipo de parcelamento</td>
        <td><%Call Cria_Combo_juros_parcelado("juros",configuracao.getAttribute("juros"),"onchange=""define_parcelamento(document.getElementsByName('permite_parcelamento')[0].options[document.getElementsByName('permite_parcelamento')[0].selectedIndex].value,this.options[this.selectedIndex].text,'tipoParcelamento');""","Simples")%></td>
    </tr>
    <tr id="dTipoParcelado" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Tipo de parcelamento que será aplicado nas transações parceladas. Sendo "Juros do Emissor" com a taxa de juros do emissor do cartão do comprador e "Juros do Lojista" com a taxa de juros definida pelo lojista.</td>
    </tr>
    <tr id="tblTaxaDesc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercDesc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Desconto</td>
        <td><input type="text" size="5" name="taxa_desconto" value="<%=configuracao.getAttribute("taxa_desconto")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercDesc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de desconto que será aplicada ao valor total do pedido.</td>
    </tr>
    <tr id="tblTaxaAcresc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dPercAcresc',this);" style="cursor:pointer;"></td>
        <td>Percentual de Acréscimo</td>
        <td><input type="text" size="5" name="taxa_juros" value="<%=configuracao.getAttribute("taxa_juros")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d.]*/gi,"");' Onblur="fncPreencheValue(this, 0)">&nbsp;%</td>
    </tr>
    <tr id="dPercAcresc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina a porcentagem de acréscimo que será aplicado ao valor total do pedido.</td>
    </tr>
    <tr id="tblNumParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dNumParcelas',this);" style="cursor:pointer;"></td>
        <td>Número de Parcelas</td>
        <td><%Call Cria_Combo_Numeros("num_parcelas",configuracao.getAttribute("num_parcelas"),1,12,"onchange=""ajusta_exibeiframe(12,this.options[this.selectedIndex].value,'divparc')""")%>&nbsp;<span Onclick="mostraiframe('tblCondParc');" style="cursor:pointer;"><span id="divCondParc"><u>Clique e defina as condições de parcelamento</u></span></span></td>
    </tr>
    <tr id="dNumParcelas" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define o número máximo de parcelas permitido. Quando utilizado o tipo de parcelamento "Juros do Lojista" é possível a configuração das ações aplicadas em cada tipo de parcelamento.</td>
    </tr>
    <tr id="tblCondParc" style="display:none;" class="Linha2Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dCondParc',this);" style="cursor:pointer;"></td>
        <td height="30">Condições de Parcelamento</td>
        <td>
            <span id="divparc1">01&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc1",configuracao.getAttribute("parc1"))%><br></span>
            <span id="divparc2">02&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc2",configuracao.getAttribute("parc2"))%><br></span>
            <span id="divparc3">03&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc3",configuracao.getAttribute("parc3"))%><br></span>
            <span id="divparc4">04&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc4",configuracao.getAttribute("parc4"))%><br></span>
            <span id="divparc5">05&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc5",configuracao.getAttribute("parc5"))%><br></span>
            <span id="divparc6">06&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc6",configuracao.getAttribute("parc6"))%><br></span>
            <span id="divparc7">07&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc7",configuracao.getAttribute("parc7"))%><br></span>
            <span id="divparc8">08&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc8",configuracao.getAttribute("parc8"))%><br></span>
            <span id="divparc9">09&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc9",configuracao.getAttribute("parc9"))%><br></span>
            <span id="divparc10">10&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc10",configuracao.getAttribute("parc10"))%><br></span>
            <span id="divparc11">11&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc11",configuracao.getAttribute("parc11"))%><br></span>
            <span id="divparc12">12&nbsp;X&nbsp;<%Call Cria_Combo_OpcaoParcela("parc12",configuracao.getAttribute("parc12"))%></span>
        </td>
    </tr>
    <tr id="dCondParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Defina o tipo de ação efetuado para cada forma de parcelamento. <br>Sendo: <br>- "Desconto": Será aplicado o percentual de desconto definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas. <br>- "Sem Juros": Será dividido o valor total do pedido pelo respectivo número de parcelas, sem acréscimo ou desconto. <br>- "Com Juros": Será aplicado o percentual de acréscimo definido anteriormente ao valor total do pedido e depois dividido pelo respectivo número de parcelas.</td>
    </tr>
    <tr id="tblValorMinParc" style="display:none;" class="Linha1Tabela">
        <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dValorMinParc',this);" style="cursor:pointer;"></td>
        <td>Valor mínimo por parcela</td>
        <td><input type="text" size="20" name="valormin_parcela" value="<%=configuracao.getAttribute("valormin_parcela")%>" class="FORMbox" Onblur="fncPreencheValue(this, '0,00')"></td>
    </tr>
    <tr id="dValorMinParc" style="display:none;"> 
        <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif"> Define qual será o valor mínimo permitido para cada parcela.</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3" height="30">Breve descrição desta opção de pagamento</td>
    </tr>
    <tr class="Linha2Tabela">
        <td colspan="3"><textarea name="descricao_pagamento" ROWS="10" COLS="100%" class="FORMbox"><%=configuracao.getAttribute("descricao_pagamento")%></textarea></TD>
    </tr>
<%
End Select
End Sub
'########################################################################################################
'--> FIM SUB Mostra_formularioPagto
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB menu_ConfiguracaoPagamentos
' - 
' - 
'########################################################################################################
Sub menu_ConfiguracaoPagamentos()
'Chamada de Sub para conexão com o arquivo XML.
Call abre_ArquivoXML(Application("XMLMeiosPagamentos"),VarobjXML,VarobjRoot)

Set configuracao = VarobjRoot.selectSingleNode("configuracao")
    Set itens = configuracao.getElementsByTagName("pagto[@disponivel='sim']") 
%>
        <select name="Sel_pagto" class="FORMbox" onChange="MM_jumpMenu('parent',this,0)">
            <option value="">Selecione</option>
<%
           n_itens = itens.length
           for i = 0 to (n_itens - 1)
           Set pagto = itens.item(i)
%>
            <option value="ADM_config_pagamento.asp?nome_pagto=<%=pagto.getAttribute("nome_pagto")%>"><%=pagto.getAttribute("nome_pagto")%></option>	
<%
            next
%>
        </select>
<%
        Set pagto = Nothing
    Set configuracao = Nothing
Set itens = Nothing
'Chamada de Sub para fechamento da conexão com o arquivo XML.
Call fecha_xmlpagamentos(VarobjXML,VarobjRoot)

End Sub
'########################################################################################################
'--> FIM SUB menu_ConfiguracaoPagamentos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB menu_MeiosPagamentos
' - 
' - 
'########################################################################################################
Sub menu_MeiosPagamentos(VarFinalidade)
'Chamada de Sub para conexão com o arquivo XML.
Call abre_ArquivoXML(Application("XMLMeiosPagamentos"),VarobjXML,VarobjRoot)

If VarFinalidade = "Configuracao" Then
    URLRedirect = "ADM_config_pagamento.asp"
ElseIf VarFinalidade = "Listar Pedidos" Then
    URLRedirect = "ADM_lista_pedido.asp"
End If

Set configuracao = VarobjRoot.selectSingleNode("configuracao")
    If  VarFinalidade = "Configuracao" Then
        Set itens = configuracao.getElementsByTagName("pagto")
    Else
        Set itens = configuracao.getElementsByTagName("pagto[@disponivel='sim']") 
    End If
%>
        <select style="WIDTH: 160px;" name="Sel_pagto" class="FORMbox" onChange="MM_jumpMenu('parent',this,0)">
            <option value="">Selecione o tipo pagto</option>
            <%If VarFinalidade = "Listar Pedidos" Then%>
            <option value="ADM_lista_pedido.asp?nome_pagto=Todos">Todos</option>            
<%          End if

           n_itens = itens.length
           for i = 0 to (n_itens - 1)
           Set pagto = itens.item(i)
%>
            <option value="<%=URLRedirect%>?nome_pagto=<%=pagto.getAttribute("nome_pagto")%>"><%=pagto.getAttribute("nome_visualizacao")%></option>	
<%
            next
%>
        </select>
<%
        Set pagto = Nothing
    Set configuracao = Nothing
Set itens = Nothing
'Chamada de Sub para fechamento da conexão com o arquivo XML.
Call fecha_xmlpagamentos(VarobjXML,VarobjRoot)
End Sub
'########################################################################################################
'--> FIM SUB menu_MeiosPagamentos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB MeiosPagamentos_Aceito
' - 
' - 
'########################################################################################################
Function MeiosPagamentos_Aceito()
'Chamada de Sub para conexão com o arquivo XML.
Call abre_ArquivoXML(Application("XMLMeiosPagamentos"),VarobjXML,VarobjRoot)


Set configuracao = VarobjRoot.selectSingleNode("configuracao")
    Set itens = configuracao.getElementsByTagName("pagto[@disponivel='sim']") 
 
           n_itens = itens.length
           for i = 0 to (n_itens - 1)
           Set pagto = itens.item(i)

           MeiosPagamentos_Aceito = MeiosPagamentos_Aceito & pagto.getAttribute("nome_visualizacao") & ","

            next

        Set pagto = Nothing
    Set configuracao = Nothing
Set itens = Nothing
'Chamada de Sub para fechamento da conexão com o arquivo XML.
Call fecha_xmlpagamentos(VarobjXML,VarobjRoot)
End Function
'########################################################################################################
'--> FIM SUB MeiosPagamentos_Aceito
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB lista_TipPagBB
' - Função para listagem dos tipos de pagamentos disponíveis para o BB Office Banking
'########################################################################################################
Sub lista_TipPagBB(nome,opcao)

    Dim Valor(2), Tipo(2)

    Valor(1)="0"
    Valor(2)="3"

    Tipo(1)="Todas opções"
    Tipo(2)="Débito em Conta Corrente"
%>
    <SELECT NAME="<%= nome%>" class="FORMbox">
<% 

    For I=1 to 2
        If opcao = Valor(i) Then    %>
            <OPTION SELECTED VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% Else %>
            <OPTION VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% End If
    Next 
%>
    </SELECT>
<% End Sub
'########################################################################################################
'--> FIM SUB lista_TipPagBB
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_Combo_Tipo_Boleto
' - Função para listagem dos tipos de boletos bancários
'########################################################################################################
Sub Cria_Combo_Tipo_Boleto(nome,opcao)

    Dim Valor(4), Tipo(4)

    Valor(1)="BoletoGenerico"
    Valor(2)="BoletoItau"
    Valor(3)="BoletoBradesco"
    Valor(4)="BoletoBancoBrasil"

    Tipo(1)="Boleto Genérico LocaWeb"
    Tipo(2)="Boleto Itaú ShopLine"
    Tipo(3)="Boleto Pagamento Fácil Bradesco"
    Tipo(4)="Boleto BB Office Banking"
%>
    <SELECT NAME="<%= nome%>" class="FORMbox">
<% 

    For I=1 to 4
        If opcao = Valor(i) Then    %>
            <OPTION SELECTED VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% Else %>
            <OPTION VALUE="<%= Valor(i) %>"><%= Tipo(i) %></OPTION>		
        <% End If
    Next 
%>
    </SELECT>
<% End Sub
'########################################################################################################
'--> FIM SUB Cria_Combo_Tipo_Boleto
'########################################################################################################
%>