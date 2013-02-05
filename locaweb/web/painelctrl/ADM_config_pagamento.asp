<%
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Loja Exemplo Locaweb
' Versão: 6.5
' Data: 12/09/06
' Arquivo: ADM_config_pagamento.asp
' Versão do arquivo: 0.0
' Data da ultima atualização: 08/10/08
'
'-----------------------------------------------------------------------------
' Licença Código Livre: http://comercio.Locaweb.com.br/gpl/gpl.txt
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

' Esta página só pode ser acessada se o visitante já se autenticou
checa_senha()

'Verifica se o perfil de usuário permite acesso a esta página
Call checa_perfil_admin(""&ADMMeioPagto&"")
%>
<HTML>
<HEAD>
<TITLE> <%=Application("NomeLoja")%> </TITLE>
</HEAD>
<body bgcolor="#EFEFEF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td width="1" bgcolor="#CCCCCC"><img src="images/regua1x1.gif" height="1"></td>
        <td valign="top">
			<!--##################################################################################-->
            <table width="778" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr>
					<td colspan="2" height="45"><!--#INCLUDE FILE="ADM_layout_inicio.asp"--></td>
				</tr>
                <tr>
					<td colspan="2" height="1" bgcolor="#CCCCCC"><img src="images/regua1x1.gif" height="1"></td>
				</tr>
				<tr>
					<td align="center" width="180" valign="top" bgcolor="#F9F9F9"><!--#INCLUDE FILE="ADM_menu.asp"--></td>
					<td width="596" bgcolor="#FFFFFF" valign="top" style="padding:10px;">
                    <!--##################################################################################-->
                        <%
                        nome_pagto = request("nome_pagto")
                        Call abre_xmlpagamentos(VarobjXML,VarobjRoot)
                        Call altera_xmlpagamentos(request("nome_pagto"),VarobjRoot,configuracao)
                        Set configuracao = VarobjRoot.selectSingleNode("configuracao/pagto[@nome_pagto='"&nome_pagto&"']")
                        %>
                        <span class="TituloPage">&#8226; Configuração <%=configuracao.getAttribute("nome_visualizacao")%></span>
                        <div align="right"><a href="ADM_lista_Formaspagamentos.asp" class="TextoPageLink">Ir para relação de formas de pagamentos</a></div>
                        <br>
                        <br>
						<form method="post" action="ADM_config_pagamento.asp">
                        <table width="558" border="0" cellpadding="4" cellspacing="1" class="BordaTabela" align="center" valign="middle">
                        <input type="hidden" name="acao" value="alterar">
                        <input type="hidden" name="nome_pagto" value="<%=request("nome_pagto")%>">
                        <%If request("acao") = "alterar" Then%>
                            <tr class="TituloTabela">
                                <td colspan="3" align="center" bgcolor="#FFFFFF">
                                    <table border="0" width="100%" cellpadding="0" cellspacing="0" bgcolor="#FBEDED">
                                        <tr> 
                                            <td align="center" height="30"><font color="#FF0000"><B>Dados alterados com sucesso.</B></font></td>
                                        </tr> 
                                    </table>
                                </td>
                            </tr>
                        <%End If%>
                            <tr class="Linha2Tabela">
                                <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dconfpag_1',this);" style="cursor:hand;"></td>
                                <td width="250">Disponível no site?</td>
                                <td width="350" align="left"><%Call Cria_Combo_opcao("disponivel",configuracao.getAttribute("disponivel"),"")%></td>
                            </tr>
                            <tr id="dconfpag_1" style="display:none;"> 
                                <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif">  Ativa ou desativa a opção de pagamento no site.</td>
                            </tr>
                        <%If nome_pagto = "Visa" Or nome_pagto = "VisaElectron" Or nome_pagto = "Diners" Or nome_pagto = "Mastercard" Or nome_pagto = "Amex" Or nome_pagto = "Itau" Or nome_pagto = "Unibanco" Or nome_pagto = "Boleto" Or nome_pagto = "Paggo"   Then%>
                            <tr class="Linha1Tabela">
                                <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dIdentificacaoLocaweb',this);" style="cursor:hand;"></td>
                                <td width="250">Identificação LocaWeb?</td>
                                <td width="350" align="left"><input type="text" size="30" name="IdentificacaoLocaweb" value="<%=configuracao.getAttribute("IdentificacaoLocaweb")%>" class="FORMbox" onKeyUp='this.value=this.value.replace(/[^\d]*/gi,"");' maxlength="15"></td>
                            </tr>
                            <tr id="dIdentificacaoLocaweb" style="display:none;"> 
                                <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif">  Número de identificação do serviço de Comércio Eletrônico contratado junto a LocaWeb. Este número pode ser obtido em seu <a href="http://painel.locaweb.com.br" target="_blank" class="TextoPageLink">Painel de Controle</a> no serviço Comércio Eletrônico.</td>
                            </tr>
                        <%End If%>
                        <%If nome_pagto = "Visa" Or nome_pagto = "VisaElectron" Or nome_pagto = "Diners" Or nome_pagto = "Mastercard" Or nome_pagto = "Amex" Or nome_pagto = "Itau" Or nome_pagto = "Unibanco" Or nome_pagto = "Boleto" Or nome_pagto = "Bradesco" Or nome_pagto = "Paggo" Or nome_pagto = "PagSeguro" Then%>
                            <tr class="Linha2Tabela">
                                <td width="16"><IMG SRC="images/duvida.gif" WIDTH="16" HEIGHT="16" BORDER="0" ALT="" onClick="mostrahelp('dUsuarioLocaweb',this);" style="cursor:hand;"></td>
                                <td width="250">Ambiente</td>
                                <td width="350" align="left"><%Call MontaCombo_opcaoAmb("ambiente",configuracao.getAttribute("ambiente"))%></td>
                            </tr>
                            <tr id="dUsuarioLocaweb" style="display:none;"> 
                                <td align="left" valign="top" colspan="3" bgcolor="#FFF9F9"> <img src="images/hr_l.gif">  Define o ambiente da configuração no Comércio Eletrônico LocaWeb, entre "TESTE" ou "PRODUÇÃO".</td>
                            </tr>
                        <%End If%>
                            <% Call Mostra_formularioPagto(request("nome_pagto"))%>
                            <tr class="Linha3Tabela">
                                <td colspan="3" align="right"><input type="submit" value="Aplicar alterações" class="bttn4"></td>
                            </tr>
                        <%
                        Call fecha_xmlpagamentos(VarobjXML,VarobjRoot)
                        %>
                        </table></form>
                        <%If nome_pagto = "Visa" Or nome_pagto = "Diners" Or nome_pagto = "Mastercard" Or nome_pagto = "Amex" Or nome_pagto = "Paggo" Or nome_pagto = "PagSeguro" Then%>
                        <script>
                            // Executa as definições de parcelamento
                            define_parcelamento(document.getElementsByName('permite_parcelamento')[0].options[document.getElementsByName('permite_parcelamento')[0].selectedIndex].value,document.getElementsByName('juros')[0].options[document.getElementsByName('juros')[0].selectedIndex].text,'parcelamento');
                            ajusta_exibeiframe(12,document.getElementsByName('num_parcelas')[0].options[document.getElementsByName('num_parcelas')[0].selectedIndex].value,'divparc')

                            //Seta as mascaras nos inputs
                            var decimalSeparator = ",";
                            var groupSeparator = ".";

                            var numParserValor = new NumberParser(2, decimalSeparator, groupSeparator, true);
                            numParserValor.currencySymbol = ""
                            numParserValor.useCurrency = true;
                            numParserValor.currencyInside = true;
                            var numMaskValor = new NumberMask(numParserValor, "valormin_parcela", 6);
                        </script>
                        <%End If%>
                    <!--##################################################################################-->
                    </td>
				</tr>
				<tr>
					<td colspan="2" height="1" bgcolor="#CCCCCC"><img src="images/regua1x1.gif" height="1"></td>
				</tr>
                <tr>
					<td align="center" height="20" colspan="2" bgcolor="#F2F2F2"><!--#INCLUDE FILE="ADM_layout_termino.asp"--></td>
				</tr>
			</table>
            <!--##################################################################################-->
        </td>
    <td width="1" bgcolor="#CCCCCC"><img src="images/regua1x1.gif" height="1"></td>
  </tr>
</table>

</BODY>
</HTML>