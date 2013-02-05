<%
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Loja Exemplo Locaweb 
' Versão: 6.5
' Data: 12/09/06
' Arquivo: funcoes_config_loja.asp
' Versão do arquivo: 0.0
' Data da ultima atualização: 15/10/08
'
'-----------------------------------------------------------------------------
' Licença Código Livre: http://comercio.locaweb.com.br/gpl/gpl.txt
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

'########################################################################################################
'SUB identifica_caminhos
'   - Esta SUB identifica caminhos e URLs a serem usadas na loja
'   - Ela é chamada no arquivo cabecalho.asp
'########################################################################################################
Sub identifica_caminhos()

    REM Identifica o diretório onde se encontra os arquivos da loja.
    VarCaminhoArq = request.servervariables("PATH_INFO")
    ArrCaminhoArq = Split(VarCaminhoArq,"/")
    For I = 0 to (Ubound(ArrCaminhoArq) - 1)
        VarUrl_Loja = VarUrl_Loja &"/"& ArrCaminhoArq(I)
    Next

    REM Identifica o usuário.
    REM O usuário é o mesmo que o de FTP da 765254 .
    VarCaminhoFis = request.servervariables("PATH_TRANSLATED")
    ArrCaminhoFis = split(VarCaminhoFis,"\")
    PathFis = Ubound(ArrCaminhoFis) - 1
    VarUsuario = ArrCaminhoFis(2)

    REM Caminho fisico da aplicação.
    VarCaminhoApp = request.servervariables("APPL_PHYSICAL_PATH")
    ArrCaminhoApp = split(VarCaminhoApp,"\")
    PathApp = Ubound(arrcaminhoApp) - 1

    REM Verifica se o diretório de aplicação é o mesmo onde se encontra os arquivos da loja.
    REM Esta condição é usada para identificar se o diretório está devidamente configurado com aplicação.
    If UCase(ArrCaminhoFis(PathFis)) <> UCase(ArrCaminhoApp(PathApp)) Then
        Session("caminhoApp") = "diferente"
        session("PathFis") = ArrCaminhoFis(PathFis)
    End If

    REM Verifica qual o drive que está instalada a loja
    VarDrive = MID(Request.ServerVariables("APPL_PHYSICAL_PATH"),1,Instr(Request.ServerVariables("APPL_PHYSICAL_PATH"),"\"))

    REM Application("Loja")= O mesmo que o usuário de FTP na 765254 
    Application("Loja") = VarUsuario
	
	REM Verifica se a hospedagem é tipo Plesk
    If LCase(ArrCaminhoFis(1)) = "vhosts" Then

		REM Tipo de hospedagem
		Application("TipoHospedagem") = "LWPlesk"

		REM Caminho dos diretórios da loja    
		If LCase(ArrCaminhoFis(PathFis)) <> "httpdocs" Then 
			If ArrCaminhoFis(PathFis) = "painelctrl" Then
				Application("DiretorioLoja") = VarDrive & "vhosts\"&VarUsuario&"\httpdocs\"&ArrCaminhoFis(PathFis-1) & "\"
				If Right(VarUsuario,3) <> "httpdocs" Then
					Application("DiretorioLoja") = Replace(LCase(Application("DiretorioLoja")),"httpdocs\httpdocs\","httpdocs\")
				Else
					Application("DiretorioLoja") = Replace(LCase(Application("DiretorioLoja")),"httpdocs\httpdocs\httpdocs\","httpdocs\httpdocs\")
				End If
				Application("nomeConfiguracao") = ArrCaminhoFis(PathFis-1)
			Else
				Application("DiretorioLoja") = VarDrive & "vhosts\"&VarUsuario&"\httpdocs\"&ArrCaminhoFis(PathFis) & "\"
				Application("nomeConfiguracao") = ArrCaminhoFis(PathFis)
			End If
		Else
			Application("DiretorioLoja") = VarDrive & "vhosts\"&VarUsuario&"\httpdocs\"
			Application("nomeConfiguracao") = "vhosts"
		End If

		Application("DiretorioDados") = VarDrive & "vhosts\"&VarUsuario&"\private\dadosloja_"&Application("nomeConfiguracao") & "\"

		If LCase(ArrCaminhoFis(PathFis)) <> "httpdocs" Then
			Application("URLADMloja") = "http://" & request.servervariables("SERVER_NAME") & Replace(VarUrl_Loja,"//","/") & "/painelctrl"
		Else
			Application("URLADMloja") = "http://" & request.servervariables("SERVER_NAME") & Replace(VarUrl_Loja,"//","/") & "painelctrl"
		End If

		REM Definição de URLs que serão utilizadas na loja.
		If InStr(Application("SSLloja"),"https") Then
			
			If InStr(Application("SSLloja"),"https://ssl") Then
				If LCase(Application("nomeConfiguracao")) = "httpdocs" Then
					Application("URLSiteSeguro") = Application("SSLloja") &"/"&varUsuario&""
					Application("URLAdmSeguro") = Application("SSLloja") &"/"&varUsuario&""
				Else
					Application("URLSiteSeguro") = Application("SSLloja") &"/"&varUsuario&"/"&ArrCaminhoFis(PathFis)
					Application("URLAdmSeguro") = Application("SSLloja") &"/"&varUsuario&"/"&ArrCaminhoFis(PathFis-1)
				End if
			Else
				If LCase(Application("nomeConfiguracao")) = "httpdocs" Then
					Application("URLSiteSeguro") = Application("SSLloja")
					Application("URLAdmSeguro") = Application("SSLloja")
				Else
					Application("URLSiteSeguro") = Application("SSLloja") &"/"&ArrCaminhoFis(PathFis)
					Application("URLAdmSeguro") = Application("SSLloja") &"/"&ArrCaminhoFis(PathFis-1)
				End if
			End If
			Application("URLAdministracao") = Application("URLAdmSeguro") & "/painelctrl/"
		Else
			Application("URLSiteSeguro") = Application("URLloja")
			Application("URLAdministracao") = Application("URLSiteSeguro") & "/painelctrl/"
		End If

	REM Verifica se a hospedagem é tipo convencional
    ElseIf LCase(ArrCaminhoFis(1)) = "home" Then

		REM Tipo de hospedagem
		Application("TipoHospedagem") = "LWConvencional"

		REM Caminho dos diretórios da loja    
		If LCase(ArrCaminhoFis(PathFis)) <> "web" Then
			If ArrCaminhoFis(PathFis) = "painelctrl" Then
				Application("DiretorioLoja") = VarDrive & "home\"&VarUsuario&"\web\"&ArrCaminhoFis(PathFis-1) & "\"
				If Right(VarUsuario,3) <> "web" Then
					Application("DiretorioLoja") = Replace(LCase(Application("DiretorioLoja")),"web\web\","web\")
				Else
					Application("DiretorioLoja") = Replace(LCase(Application("DiretorioLoja")),"web\web\web\","web\web\")
				End If
				Application("nomeConfiguracao") = ArrCaminhoFis(PathFis-1)
			Else
				Application("DiretorioLoja") = VarDrive & "home\"&VarUsuario&"\web\"&ArrCaminhoFis(PathFis) & "\"
				Application("nomeConfiguracao") = ArrCaminhoFis(PathFis)
			End If
		Else
			Application("DiretorioLoja") = VarDrive & "home\"&VarUsuario&"\web\"
			Application("nomeConfiguracao") = "web"
		End If

		Application("DiretorioDados") = VarDrive & "home\"&VarUsuario&"\dados\dadosloja_"&Application("nomeConfiguracao") & "\"

		If LCase(ArrCaminhoFis(PathFis)) <> "web" Then
			Application("URLADMloja") = "http://" & request.servervariables("SERVER_NAME") & Replace(VarUrl_Loja,"//","/") & "/painelctrl"
		Else
			Application("URLADMloja") = "http://" & request.servervariables("SERVER_NAME") & Replace(VarUrl_Loja,"//","/") & "painelctrl"
		End If

		REM Definição de URLs que serão utilizadas na loja.
		If InStr(Application("SSLloja"),"https") Then
			
			If InStr(Application("SSLloja"),"https://ssl") Then
				If LCase(Application("nomeConfiguracao")) = "web" Then
					Application("URLSiteSeguro") = Application("SSLloja") &"/"&varUsuario&""
					Application("URLAdmSeguro") = Application("SSLloja") &"/"&varUsuario&""
				Else
					Application("URLSiteSeguro") = Application("SSLloja") &"/"&varUsuario&"/"&ArrCaminhoFis(PathFis)
					Application("URLAdmSeguro") = Application("SSLloja") &"/"&varUsuario&"/"&ArrCaminhoFis(PathFis-1)
				End If
			Else
				If LCase(Application("nomeConfiguracao")) = "web" Then
					Application("URLSiteSeguro") = Application("SSLloja")
					Application("URLAdmSeguro") = Application("SSLloja")
				Else
					Application("URLSiteSeguro") = Application("SSLloja") &"/"&ArrCaminhoFis(PathFis)
					Application("URLAdmSeguro") = Application("SSLloja") &"/"&ArrCaminhoFis(PathFis-1)
				End if
			End If
			Application("URLAdministracao") = Application("URLAdmSeguro") & "/painelctrl/"
		Else
			Application("URLSiteSeguro") = Application("URLloja")
			Application("URLAdministracao") = Application("URLSiteSeguro") & "/painelctrl/"
		End If

	REM Verifica se a hospedagem é indefinida
	Else

		' Notificação de formato de hospedagem indefinido
		Response.write "Erro: Não foi possível prosseguir com a inicialização da loja. Formato de hospedagem não identificado. Necessário ajustes no arquivo de inicialização da loja (funcoes_config_loja.asp) para prosseguir."
		Response.end

	End If

    If Application("TipoHospedagem") <> "LWPlesk" Then

		'Cria o diretório dadosloja_nomedaloja em DADOS para gravação dos arquivos de gerenciamento da loja
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Application("DiretorioDados")) = False Then
			objFSO.CreateFolder(Application("DiretorioDados"))
		End If
		Set objFSO = Nothing

	End If

    Application("DiretorioAtualizacaoProdutos") = Application("DiretorioDados") &"atualizacao_produtos\"
    Application("DiretorioConfig") = Application("DiretorioLoja") &"config\"
    Application("DiretorioPedidos") = Application("DiretorioDados") &"pedidos_loja\" 
    Application("DiretorioResultsVBV") =  Application("DiretorioDados") & "resultsVBV\"
    Application("LogsADM") = Application("DiretorioDados") & "LogsADM\"
    Application("DiretorioImagensConteudo") = Application("DiretorioConfig") & "imagens_conteudo"

    REM XMLMeiosPagamentos = Arquivo de configuração dos meios de pagamentos.
    REM XMLMeiosPagamentosTemp = Caminhos dos arquivos de configuração da loja.
    REM XMLArquivoConfiguracao = Arquivos de configuração da loja (Nome da Loja, Caminho Do banco, tipo de banco, SSL)
    REM XMLEstruturaDadosSQL = Arquivo que contem a estrutura de dados SQL SERVER da loja. Este arquivo é executado na configuração da loja.
	REM XMLEstruturaDadosMySQL = Arquivo que contem a estrutura de dados MYSQL da loja. Este arquivo é executado na configuração da loja.
    Application("XMLMeiosPagamentos") = Application("DiretorioDados") & "meiosPagamentos.xml"
    Application("XMLMeiosPagamentosTemp") = Application("DiretorioConfig") & "meiosPagamentos.xml"
    Application("XMLArquivoConfiguracao") = Application("DiretorioDados") &"configuracaoLoja.xml" 
    Application("XMLEstruturaDadosSQL") = Application("DiretorioConfig") & "instalador\estruturadadosSQLServer.xml"
	Application("XMLEstruturaDadosMySQL") = Application("DiretorioConfig") & "instalador\estruturadadosMySQL.xml"

    'Campos utilizado na consulta da negativação de dados.
    Application("CamposNegativados") = "user_ID,ip_cliente,razaosocial_cobranca,cnpj_cobranca,inscricaoestadual_cobranca,cpf_cobranca,rg_cobranca,telefone_cobranca,razaosocial_entrega,cnpj_entrega,inscricaoestadual_entrega,rg_entrega,telefone_entrega,email_entrega"

    'Carrega a configuração geral da loja.
    'Na Function Carrega_Configuracao outras Applications serão criadas para alimentar a loja.
    Call Carrega_Configuracao()
    
    'Define qual o idioma default usado na Loja.
    If Request("lang") <> "" Then
        varLangTemp = Request("lang")
    Else
        If session("varLangUser") <> "" Then
            varLangTemp = session("varLangUser")
        Else
            varLangTemp = Application("IdiomaDefault") 
        End if
    End If
    Application("varLang") = varLangTemp

    varLang = Application("varLang")

    REM DiretorioTemplates = Diretorio onde consta o layout da loja
    REM DiretorioConfiguracao = Diretorio onde consta a configuração de idiomas da loja
    REM XMLTextosAdicionais = Arquivo que contem textos de instrução e informação Do site (Como comprar, Termos de Uso, etc...)
    Application("DiretorioTemplates") = Application("DiretorioLoja")&"config\templates\"
    Application("DiretorioConfiguracao") = Application("DiretorioLoja")&"config\templates\"& varLang & "\Configuracao\"
    Application("XMLTextosAdicionais") = Application("DiretorioConfiguracao") & "textosadicionais.xml"
    
	Application("URLWSPagamentoCertoLocaweb") = "https://www.pagamentocerto.com.br/vendedor/vendedor.asmx"
    Application("URLPagamentoCertoLocaweb") = "https://www.pagamentocerto.com.br/pagamento/pagamento.aspx"
	Application("URLLocaWebBoletoPagamentoCerto") = "https://www.pagamentocerto.com.br/pagamento/ReemissaoBoleto.aspx"
	Application("URLWebServiceCorreiosLocaweb") = "https://comercio.locaweb.com.br/correios/frete.asmx"
    Application("URLRecibo") = Application("URLSiteSeguro") & "/recibo.asp"
    Application("URLLogoLoja") = Application("URLloja") & "/config/imagens_conteudo/padrao/logo.gif"
    Application("URLLogoLojaSSL") = Application("URLSiteSeguro") & "/config/imagens_conteudo/padrao/logo.gif"
    Application("URLDirectExpresCalculo") = "http://www.directlog.com.br/frete/pega_frete.asp"
    Application("URLTESTEFEDEX") = "gatewaybeta.fedex.com"
    Application("URLPRODFEDEX") = "gateway.fedex.com"
    Application("URLLocaWebBoleto") = "http://comercio.locaweb.com.br/comercio.comp"
    Application("URLLocaWebBoletoCobreBem") = "https://comercio.locaweb.com.br/cgi-bin/cobrebemecommerce.exe"
    Application("URLBancoBrasil") = "https://www16.bancodobrasil.com.br/site/mpag/"
    Application("URLBancoBrasilCaptura") = "https://www11.bb.com.br/site/mpag/REC3.jsp"
    Application("URLRecebeDadosVisaMOSET") = "https://comercio.locaweb.com.br/locawebce/comercio.aspx"
    Application("URLRecebeDadosVisaVBV") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLVisanetXMLVBV") = "https://comercio.locaweb.com.br/visavbv/results/"
    Application("URLVisanetCaptura") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLRedecard") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLRedeCardAdmin") = "http://www.redecard.com.br"
    Application("URLRedecardConfirma") = "http://ecommerce.redecard.com.br/pos_virtual/confirma.asp"
    Application("URLRedeCardCupom") = "https://ecommerce.redecard.com.br/pos_virtual/cupom.asp"
    Application("URLTESTEBradescoPagFacil") = "http://mupteste.comercioeletronico.com.br/sepsapplet/"
    Application("URLTESTEBradescotransfer") = "http://mupteste.comercioeletronico.com.br/sepstransfer/"
    Application("URLTESTEBradescoFinanciamento") = "http://mupteste.comercioeletronico.com.br/sepsfinanciamento/"
    Application("URLTESTEBradescoBoleto") = "http://mupteste.comercioeletronico.com.br/sepsboleto/"
    Application("URLPRODBradescoPagFacil") = "https://mup.comercioeletronico.com.br/sepsapplet/"
    Application("URLPRODBradescotransfer") = "https://mup.comercioeletronico.com.br/sepstransfer/"
    Application("URLPRODBradescoFinanciamento") = "https://mup.comercioeletronico.com.br/sepsfinanciamento/"
    Application("URLPRODBradescoBoleto") = "https://mup.comercioeletronico.com.br/sepsboleto/"
    Application("URLItauShopline") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLItauConsulta") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLUnibanco") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLUnibancoConsulta") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLAmex") = "https://comercio.locaweb.com.br/comercio.comp"
    Application("URLAmexCaptura") = "http://comercio.locaweb.com.br/comercio.comp"
    Application("URLABNCDC") = "https://wwws.aymorefinanciamentos.com.br/scripts/flv.dll/Simula?Pagina=simula_completa"
    Application("URLABNCDCSimulador") = "https://wwws.aymorefinanciamentos.com.br/scripts/flv.dll/Simula?Pagina=simula_simples"
    Application("URLABNCDCconsulta") = "https://wwws.aymorefinanciamentos.com.br"
	Application("URLPaggo") = "https://comercio.locaweb.com.br/locawebce/comercio.aspx"
	Application("URLPagSeguro") = "https://pagseguro.uol.com.br/security/webpagamentos/webpagto.aspx"
    Application("URLClearSale") = "http://comercio.locaweb.com.br/LocaWebCE/comercio.aspx"

    Application("URL_Senha_Admin") = "default.asp"

End Sub
'########################################################################################################
'--> FIM FUNCTION identifica_caminhos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Carrega_Configuracao
' - Carrega as configurações gerais e de formas de entrega da Loja.
' - Chamada no arquivo cabecalho.asp
'########################################################################################################
Function Carrega_Configuracao()
    
    'Procura o arquivo de configuração da Loja
    Set FSO = CreateObject ("Scripting.FileSystemObject")
        If FSO.fileExists(Application("XMLArquivoConfiguracao")) Then
            erro = 0
        Else
            erro = 1
        End If
    Set FSO = Nothing

    'Redireciona para o instalador caso o arquivo de configuração não seja localizado.
    If erro <> 0 Then
        Response.redirect "config/instalador"
        Response.end
    End If

    Set varObjXML = CreateObject("Microsoft.XMLDOM")
        varObjXML.preserveWhiteSpace = False
        varObjXML.async = False
        varObjXML.validateOnParse = True
        varObjXML.resolveExternals = True
        varObjXML.load (Application("XMLArquivoConfiguracao"))
        Set varObjRoot = varObjXML.documentElement
        'Define a raiz do arquivo XML.
        Set objRaiz = varObjRoot.selectSingleNode("dados")
            'Cria o objeto FSO para ler os applications de configuração da Loja.
            Set fs = Server.CreateObject("Scripting.FileSystemObject")
            'Caminho do arquivo TXT de configuração.
            caminho = Application("DiretorioDados")&"camposconfigxml.txt"
                'Abre o arquivo TXT de configuração.
                Set txt = fs.OpenTextFile(caminho, 1,0)
                    'Loop para leitura do arquivo TXT de configuração. Cada linha refere-se a um Application de configuração.
                    While (Not txt.AtEndOfStream) And response.isclientconnected()
						'Leitura de cada linha do arquivo
						linha_txt = txt.readline
						'Define um array contendo todas as linhas do arquivo.
						arrayx = split(linha_txt)
                        Set objNote = objRaiz.getElementsByTagName("configuracao_dados") 
                            Set ValorAtrib = objNote.item(i)
                                'Cria as applications a partir das variávies do arquivo camposconfigxml.txt;
                                'Atrbui os valores do arquivo configxml.xml.
                                application(arrayx(0)) = ValorAtrib.getAttribute(arrayx(0))
                                'Liberando a linha abaixo, será listado todo o arquivo de configuração da loja.
                                'response.write arrayx(0) & ": " & application(arrayx(0))  & "<br>"
                            Set ValorAtrib = Nothing
                        Set objNote = Nothing
                    Wend
                'Fecha o arquivo TXT de configuração
                txt.close 
                'Libera o objeto da memória.
                Set txt  = Nothing
        'Libera o objeto da memória.
        Set objRaiz  = Nothing
    ' Fecha o arquivo de configuração XML da loja.
    set varObjXML = Nothing
    Set varObjRoot = Nothing
End Function
'########################################################################################################
'--> FIM FUNCTION Carrega_Configuracao
'########################################################################################################
%>