<%
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Loja Exemplo Locaweb 
' Versão: 6.5
' Data: 08/02/07
' Arquivo: funcoes.asp
' Versão do arquivo: 0.0
' Data da ultima atualização: 23/10/08
'
'-----------------------------------------------------------------------------
' Licença Código Livre: http://comercio.locaweb.com.br/gpl/gpl.txt
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
'########################################################################################################
'SUB abre_conexao
'   - Abre conexão com o banco de dados
'   - Chamada no arquivo layout_inicio.asp 
'SUB fecha_conexao
'   - Fecha conexão com o banco de dados
'   - Chamada no arquivo layout_termino.asp 
'########################################################################################################

Sub abre_conexao(conexao)
    'Define o objeto de conexão ao banco de dados
    Set Conexao = CreateObject("ADODB.Connection")
    'Chama a funcão de carregamento do arquivo XML
    Call abre_ArquivoXML(Application("XMLArquivoConfiguracao"),FctobjXML,FctobjRoot)
    'Verifica a existência do arquivo XML
    existe_configuracao = FctobjXML.load(Application("XMLArquivoConfiguracao"))
        'Define o objeto de raiz do documento   
        Set FctobjRoot = FctobjXML.documentElement
            'Se o arquivo XML existir haverá a leitura do mesmo
            If existe_configuracao = True Then
                'Define o objeto de leitura dos NÓS
                Set configuracao = FctobjRoot.selectSingleNode("dados/configuracao_dados")
                    Application("TipoBanco") = configuracao.getAttribute("TipoBanco")
                    If Application("TipoBanco") = "mssql" Then
                        ' Verifica se foi específicado uma base SQL
                        If configuracao.getAttribute("BaseBD") <> "" Then
                            baseMssql = configuracao.getAttribute("UsuarioBD")
                        ' Caso contrário define a base com o mesmo nome do usuário SQL
                        Else
                            baseMssql = configuracao.getAttribute("BaseBD")
                        End If
                        Application("StringConexaoBanco") = "PROVIDER=SQLOLEDB.1;DATA SOURCE=" & configuracao.getAttribute("EnderecoBD") & ";DATABASE=" & baseMssql & " ;UID=" & configuracao.getAttribute("UsuarioBD") & " ;PWD=" & configuracao.getAttribute("SenhaBD") & ";"
                    ElseIf Application("TipoBanco") = "mysql" Then
                        ' Verifica se foi específicado uma base MySQL
                        If configuracao.getAttribute("BaseBD") <> "" Then
                            baseMysql = configuracao.getAttribute("BaseBD")
                        ' Caso contrário define a base com o mesmo nome do usuário MySQL
                        Else
                            baseMysql = configuracao.getAttribute("UsuarioBD")
                        End If
                        Application("StringConexaoBanco") = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & configuracao.getAttribute("EnderecoBD") & ";PORT=3306;DATABASE=" & baseMysql & ";USER=" & configuracao.getAttribute("UsuarioBD") & ";PASSWORD=" & configuracao.getAttribute("SenhaBD") & ";OPTION=3;"
                    End If
                    Application("NomeLoja") = configuracao.getAttribute("NomeLoja")
                'Destroi o objeto de leitura do nó
                Set configuracao =Nothing
            End If
        'Destrói o objeto de raiz do documento
        Set FctobjRoot = Nothing
    'Abre o banco de dados
    Conexao.open Application("StringConexaoBanco")
End sub

sub fecha_conexao
    'Fecha conexão com o banco de dados
    Conexao.Close
    'Destrói o objeto de conexão
    Set Conexao=nothing
End Sub

'########################################################################################################
'--> FIM SUB abre_conexao
'--> FIM SUB fecha_conexao
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB menu_categorias
' - Lista Categorias e Subcategorias cadastradas
' - Chamada no arquivo CATEGORIAS.ASP
'########################################################################################################
Sub menu_categorias()

    queryORDERBY = "ORDER BY " & Application("OrdemCategoria")
    
    'Query de consulta a tabela Categorias
    Query_categorias = "SELECT codigo_chave, codigo_categoria, nome_categoria, descricao_categoria, sigla_idioma FROM Categorias WHERE sigla_idioma = '"&varLang&"' " & queryORDERBY
    'Cria objeto para consultar as categorias
    Set RS_Categorias = Server.CreateObject("ADODB.Recordset")
    'Conexão ADO
    Set RS_Categorias.ActiveConnection = Conexao
    RS_Categorias.CursorLocation = 3
    RS_Categorias.CursorType = 0
    RS_Categorias.LockType =  1
    'Executa a query
    RS_Categorias.Open Query_categorias
    If  RS_Categorias.EOF then
%>
    <table width="172" cellpadding="0" cellspacing="0">
        <tr>
            <td style="padding-top:2px;padding-bottom:2px;padding-left:10px;padding-right:10px" class="MNlatesquerda"><%= Application("MenuTxtCatVazio")%></td>
        </tr>
    </table>
<%
    Else
        'Loop dos registros encontrados na tebela categorias
		DO UNTIL  RS_Categorias.EOF
            'Atribui o calor do codigo_categoria a session("codigo_categoria")
			session("codigo_categ") = RS_Categorias("codigo_categoria")
            'Query de consulta a tabela subcategorias
			Query_subcategorias = "SELECT codigo_chave, codigo_subcategoria, codigo_categoria, nome_subcategoria, descricao_subcategoria, sigla_idioma FROM Subcategorias WHERE codigo_categoria = " & session("codigo_categ") & " AND sigla_idioma = '"&varLang&"' ORDER BY nome_subcategoria"
            'Cria objeto para consultar Subcategorias
            Set RS_SubCategorias = CreateObject("ADODB.Recordset")
            'Conexão ADO
            Set RS_SubCategorias.ActiveConnection = Conexao
            RS_SubCategorias.CursorLocation = 3
            RS_SubCategorias.CursorType = 0
            RS_SubCategorias.LockType =  1
            'Executa a query de consulta
            RS_SubCategorias.Open Query_subcategorias
%>
                <table width="172" cellpadding="0" cellspacing="0">
<%
                'Verifica se existe subcategoria 
                'Caso exista a navegação será a partir da subcategoria 
                If not RS_SubCategorias.EOF Then
                    ExisteCategoria="sim"

%>
                    <tr>
                        <td style="padding-top:2px;padding-bottom:2px;padding-left:10px;padding-right:10px"><img src="config/templates/<%=varLang%>/<%=varSkin%>/seta.gif" width="6" height="7" border="0">&nbsp;&nbsp;<a href="produtos.asp?lang=<%=varLang%>&tipo_busca=categoria&codigo_categoria=<%=RS_Categorias("codigo_categoria")%>" onclick="mostraDados('<%=RS_Categorias("codigo_categoria")%>');" class="MNlatesquerda"><%= RS_Categorias("nome_categoria")%></font></td>
                    </tr>
<%                  
                'Não existindo subcategoria a navegação será a partir da categoria
                Else
%>
                    <tr>
                        <td style="padding-top:2px;padding-bottom:2px;padding-left:10px;padding-right:10px"><img src="config/templates/<%=varLang%>/<%=varSkin%>/seta.gif" width="6" height="7" border="0">&nbsp;&nbsp;<a href="produtos.asp?lang=<%=varLang%>&tipo_busca=categoria&codigo_categoria=<%=RS_Categorias("codigo_categoria")%>" <%If CDbl(RS_Categorias("codigo_categoria")) = CDbl(request("codigo_categoria")) And request("codigo_subcategoria") = "" Then %>class="MNlatesquerdaAtivo"<%Else%>class="MNlatesquerda"<%End If%>><%=fontcolor%><%= RS_Categorias("nome_categoria")%></font></td>
                    </tr>
<%                  
                'Fim da verificação de existência de subcategoria
                End If
                If ExisteCategoria = "sim" then%>
                    <%If CDbl(RS_Categorias("codigo_categoria")) = CDbl(request("codigo_categoria")) Then%> 
                    <tr id="<%=RS_Categorias("codigo_categoria")%>" style="display:'';">
                    <%Else%>
                    <tr id="<%=RS_Categorias("codigo_categoria")%>" style="display:none;">
                    <%End if%>
                        <td align="left" style="padding-top:2px;padding-bottom:2px;padding-left:10px;padding-right:10px">
<%
                        'Loop dos registros encontrados na tebela categorias
                        DO UNTIL  RS_SubCategorias.EOF 
                            If CDbl(RS_SubCategorias("codigo_subcategoria")) = CDbl(request("codigo_subcategoria")) Then
                                Response.write "&nbsp;&nbsp;<span class='MNlatesquerdaAtivo'>-</span> <a href='produtos.asp?lang="&varLang&"&tipo_busca=subcategoria&codigo_categoria="&RS_Categorias("codigo_categoria")&"&codigo_subcategoria="&RS_SubCategorias("codigo_subcategoria")&"' class='MNlatesquerdaAtivo'>" & RS_SubCategorias("nome_subcategoria") & "</a></span><br>"
                            Else
                                Response.write "&nbsp;&nbsp;<span class='MNlatesquerda'>-</span> <a href='produtos.asp?lang="&varLang&"&tipo_busca=subcategoria&codigo_categoria="&RS_Categorias("codigo_categoria")&"&codigo_subcategoria="&RS_SubCategorias("codigo_subcategoria")&"' class='MNlatesquerda'>"& fontcolor & RS_SubCategorias("nome_subcategoria") & "</a></span><br>"
                            End if
                        RS_SubCategorias.MoveNext
                        LOOP
%>
                        </td>
                    </tr>
<%
                End If
%>
                    <tr>
                        <td><img src="config/templates/<%=varLang%>/<%=varSkin%>/regua1x1.gif" height="1"></td>
                    </tr>
                </table>
<%
            'Fecha e libera da memória o objeto de Recordset de consulta a tabela Subcategorias    
			RS_SubCategorias.Close
			Set RS_SubCategorias = Nothing

		RS_Categorias.MoveNext
		Loop
    End if    
    'Fecha e libera da memória o objeto de Recordset de consulta a tabela categorias 		
	RS_Categorias.Close
	Set RS_Categorias = Nothing
End sub
'########################################################################################################
'--> FIM SUB menu_categorias
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB menu_servicos
' - Lista textos adicionais exibido na lateraln direita da loja
' - Chamado no arquivo lateral_servicos.asp
'########################################################################################################
Sub menu_servicos(fctLang,fctPosicao,fctClass)

If fctPosicao = "vertical" Then
    AdicionalCHR = "<br><br>"
    AdicionalIMG = "<img src='config/templates/"&varLang&"/"&varSkin&"/seta.gif' width='6' height='7' border='0'>&nbsp;&nbsp;"
ElseIf fctPosicao = "horizontal" Then
    AdicionalCHR = " | "
End if

ArquivoTextosAdicionais = Application("DiretorioTemplates") & fctLang & "\" & "configuracao\textosadicionais.xml"
'Abre o arquivo XML: XMLTextosAdicionais
Call abre_ArquivoXML(ArquivoTextosAdicionais,FctobjXML,FctobjRoot)
    'Seta objeto para o nó
    Set menu = FctobjRoot.selectSingleNode("configuracao")
    'Seta objeto para os atributos.
        Set itens = menu.getElementsByTagName("infos[@ativo='sim']") 
        'Verifica o númeto de itens
            n_itens = itens.length
            'Loop para captura dos itens
            for i = 0 to (n_itens - 1)
                Set pagto = itens.item(i)
                    Response.write AdicionalIMG&"<a href='infos.asp?lang="&varLang&"&codigo_texto="&pagto.getAttribute("codigo_texto")&"' class='"&fctClass&"'>" & pagto.getAttribute("titulo") & "</a>" & AdicionalCHR
                'Libera objeto da memória
                Set pagto = Nothing
            Next
        'Libera objeto da memória
        Set itens = Nothing
    'Libera objeto da memória
    Set menu = Nothing

End Sub

'########################################################################################################
'--> FIM SUB menu_categorias
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB exibe_produtos
' - Lista os RS_Produto cadastrados com as opções COMPRA/DESCRIÇÃO/INDICAR
' - Chamada no arquivo RS_Produto.ASP
'########################################################################################################
Sub exibe_produtos()

'Query_produtos monta uma string para consulta no banco de dados

' Define a ordem de exibição dos produtos
If Request("orderby") <> "" Then
    queryORDERBY = "ORDER BY Produtos." & Request("orderby")
Else
    queryORDERBY = "ORDER BY Produtos." & Application("OrdemProduto")
End If

' Oculta a possibilidade de ordem dos produtos
exibeOrdemProd = False

'Se a busca partir de uma categoria
If request("tipo_busca") = "categoria" Then

	Query_produtos = "SELECT Produtos.codigo_produto,  Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE Produtos.codigo_categoria = "&request("codigo_categoria")&" AND Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 " & queryORDERBY

	' Habilita a possibilidade de ordem dos produtos
	exibeOrdemProd = True

'Se a busca partir de uma subcategoria
ElseIf request("tipo_busca") = "subcategoria" Then

	Query_produtos = "SELECT Produtos.codigo_produto,  Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE Produtos.codigo_subcategoria = "&request("codigo_subcategoria")&"  AND Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 " & queryORDERBY

	' Habilita a possibilidade de ordem dos produtos
	exibeOrdemProd = True

'Se a busca partir de uma marca
ElseIf request("tipo_busca") = "marca" Then

	Query_produtos = "SELECT Produtos.codigo_produto,  Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE Produtos.codigo_marca = "&request("codigo_marca")&"  AND Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 " & queryORDERBY

	' Habilita a possibilidade de ordem dos produtos
	exibeOrdemProd = True

'Se a busca partir de uma pesquisa por palavra chave
ElseIf request("tipo_busca") = "palavra" Then

	'Request da palavra pesquisada
	produto = request("produto")
    
    'Caso não exista uma categoria para esta consulta
    If request("codigo_categoria") <> "000" Then
        
		Query_produtos = "SELECT Produtos.codigo_produto,  Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE (nome_produto like '%" & produto & "%' OR descricao_produto LIKE '%" & produto & "%') AND Produtos.codigo_categoria = "&request("codigo_categoria")&"  AND Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 " & queryORDERBY

    'Caso exista uma categoria para esta consulta
    Else

        Query_produtos = "SELECT Produtos.codigo_produto,  Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE (nome_produto like '%" & produto & "%' OR descricao_produto LIKE '%" & produto & "%')  AND Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 " & queryORDERBY

	End If
    
'Se a busca é de apenas produtos em ofertas
ElseIf request("tipo_busca") = "ofertas" Then

	Query_produtos = "SELECT Produtos.codigo_produto,  Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria  WHERE Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 " & queryORDERBY

	' Habilita a possibilidade de ordem dos produtos
	exibeOrdemProd = True

'Se não for busca a consulta ao banco de da página inicial ou outra página especifíca
Else
    'Exibição dos produtos na página incial    
    If page="default" Then

        'Caso o banco de dados seja MSSQL
        If Application("TipoBanco") = "mssql" Then

			Query_produtos = "SELECT Produtos.codigo_produto, Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE destaque_vitrine = 1 AND Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang& "' AND Produtos.disponivel = 1 ORDER BY NewId()"
        
        'Caso o banco de dados seja MYSQL
        ElseIf Application("TipoBanco") = "mysql" Then

			Query_produtos = "SELECT Produtos.codigo_produto, Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE destaque_vitrine = 1 AND Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 ORDER BY RAND()"

        End If
        
    'Exibição dos produtos em página que não seja inicial e nem derivada de uma pesquisa  
    Else

	    Query_produtos = "SELECT Produtos.codigo_produto,  Produtos.codigo_categoria, Produtos.codigo_subcategoria, Produtos.codigo_marca, Produtos.codigo_produto_loja, Produtos.nome_produto, Produtos.descricao_produto, Produtos.autor, Produtos.codigo_isbn, Produtos.tamanhos, Produtos.cores, Produtos.preco_base, Produtos.preco_unitario, Produtos.desconto, Produtos.moeda, Produtos.quantidade_produto, Produtos.img_produtoPQN, Produtos.img_produtoGRD, Produtos.img_produto_adic01PQN, Produtos.img_produto_adic01GRD, Produtos.img_produto_adic02PQN, Produtos.img_produto_adic02GRD, Produtos.img_produto_adic03PQN, Produtos.img_produto_adic03GRD, Produtos.peso, Produtos.destaque_vitrine, Produtos.promocao, Produtos.data_inicio, Produtos.data_fim, Produtos.disponivel,  Produtos.sigla_idioma FROM Categorias INNER JOIN Produtos ON Categorias.codigo_categoria = Produtos.codigo_categoria WHERE Produtos.sigla_idioma = '"&varLang&"' AND Categorias.sigla_idioma = '"&varLang&"' AND Produtos.disponivel = 1 " & queryORDERBY

		' Habilita a possibilidade de ordem dos produtos
		exibeOrdemProd = True

    End If

End If

'Captura o número da página atual se existir
If Request.QueryString("PN") = "" Then
    'Se a captura for vazia, será atribuido o valor inicial 
    PaginaCorrente = 1
Else
    'Existindo captura, o valor é atribuido à página corrente
    PaginaCorrente = Request.QueryString("PN")
End If

'Cpatura o número para exibição de produtos por página
If page="default" Then
    'Se página exibida for a inicial (vitrine), será atribuido o valor do application("ProdutosVitrine")
    VarIntervalo = Application("ProdutosVitrine")
ElseIf Request.QueryString("FctIntervalo")= "" Then
    'Se a captura for vazia, será atribuido o valor do application("produtosporpagina")
    VarIntervalo = Application("ProdutosPorPagina")
Else
    'Existindo captura, o valor será atribuido ao intervalo de produtos exibidos por páginas
    VarIntervalo = Request.QueryString("FctIntervalo")
End If

'Cria o objeto RS_Produto de recordset
Set RS_Produto = CreateObject("ADODB.Recordset")
Set RS_Produto.ActiveConnection = Conexao
RS_Produto.CursorLocation = 3
RS_Produto.CursorType = 0
RS_Produto.LockType =  1
RS_Produto.CacheSize = VarIntervalo

'Havendo necessidade descomente a linha abaixo para saber qual query está sendo executada.
'Response.write Query_produtos

'Executa da Query de consulta
RS_Produto.Open Query_produtos
'Define o o número de produtos exibidos na página
RS_Produto.PageSize = CInt(VarIntervalo)
'Define o número total de páginas
VarTotalPaginas= RS_Produto.PageCount
'Define o número total de produtos
Var_TotalRegistros = RS_Produto.recordcount

'Formuário de envio do produto para o carrinho.
%>
<style>
#go
{
margin: 0;
padding: 1em 0 0 0;
max-height: 1300px;
}
#go:after
{
content: ".";
display: block;
line-height: 1px;
font-size: 1px;
clear: both;
}
ul#gop
{
list-style: none;
padding: 0;
margin: 0;
width: 100%;
min-height: 100px;
}
ul#gop li
{
display: block;
float: left;
width: 100%;
margin-left: 2px;
padding: 1em 0 0 0;
text-align: center;
}
ul#gop li a
{
display: block;
width: 100%;
padding: 0;
}
</style>

<%
If VerificaExistenciaDado("codigo_categoria","Categorias","codigo_categoria",Request("codigo_categoria")) And Request("codigo_categoria") <> "" Then

    'Captura a subcategoria se existir
    Set RS_Categoria = Server.CreateObject("ADODB.Recordset")
    RS_Categoria.CursorLocation = 3
    RS_Categoria.CursorType = 0
    RS_Categoria.LockType =  1
        RS_Categoria.Open "SELECT codigo_categoria, nome_categoria FROM Categorias WHERE sigla_idioma='"&varLang&"' AND codigo_categoria = "&Request("codigo_categoria")&"" , Conexao
        nome_categoria = RS_Categoria("nome_categoria")
    'Fecha e libera da memória o objeto de Recordset
    RS_Categoria.close
    Set RS_Categoria = Nothing

    txtExibicao = "<a href='produtos.asp?lang="&varLang&"&tipo_busca=categoria&codigo_categoria="&Request("codigo_categoria")&"' class='TXTproduto'><b>" & nome_categoria & "</b></a>"

End If


If VerificaExistenciaDado("codigo_subcategoria","Subcategorias","codigo_subcategoria",Request("codigo_subcategoria")) And Request("codigo_subcategoria") <> "" Then

    'Captura a subcategoria se existir
    Set RS_Subcategoria = Server.CreateObject("ADODB.Recordset")
    RS_Subcategoria.CursorLocation = 3
    RS_Subcategoria.CursorType = 0
    RS_Subcategoria.LockType =  1
        RS_Subcategoria.Open "SELECT codigo_subcategoria, nome_subcategoria FROM Subcategorias WHERE sigla_idioma='"&varLang&"' AND codigo_subcategoria = "&Request("codigo_subcategoria")&"" , Conexao
        nome_subcategoria = RS_Subcategoria("nome_subcategoria")
    'Fecha e libera da memória o objeto de Recordset
    RS_Subcategoria.close
    Set RS_Subcategoria = Nothing

    txtExibicao = txtExibicao & " > <a href='produtos.asp?lang="&varLang&"&tipo_busca=subcategoria&codigo_categoria="&Request("codigo_categoria")&"&codigo_subcategoria="&Request("codigo_subcategoria")&"' class='TXTproduto'><b>" & nome_subcategoria & "</b></a>"

End If


If VerificaExistenciaDado("codigo_marca","Marcas","codigo_marca",Request("codigo_marca")) And Request("tipo_busca") = "marca" Then

    'Captura a marca se existir
    Set RS_Marca = Server.CreateObject("ADODB.Recordset")
    RS_Marca.CursorLocation = 3
    RS_Marca.CursorType = 0
    RS_Marca.LockType =  1
        RS_Marca.Open "SELECT codigo_marca, nome_marca FROM Marcas WHERE codigo_marca = "&Request("codigo_marca")&"" , Conexao
        nome_marca = RS_Marca("nome_marca")
    'Fecha e libera da memória o objeto de Recordset
    RS_Marca.close
    Set RS_Marca = Nothing

    txtExibicao = "<a href='produtos.asp?lang="&varLang&"&tipo_busca=marca&codigo_marca="&Request("codigo_marca")&"' class='TXTproduto'><b>" & nome_marca & "</b></a>"

End if

If request("tipo_busca") = "palavra" Then

    If request("codigo_categoria") <> "000" Then
        categoriaPesq = nome_categoria
    Else
        categoriaPesq = Application("BoxSelPesquisarTodas")
    End If

    txtExibicao = Application("MiddleTxtPesquisaPor") & "&nbsp;&nbsp;<i>" & produto & "</i>&nbsp;&nbsp;" & Application("MiddleTxtPesquisaPorComp") & "&nbsp;&nbsp;<i>" & categoriaPesq & "</i>"

End if
%>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" align="top" valign="top">
<%
If txtExibicao <> "" Or exibeOrdemProd = True Then
%>
    <tr>
        <td height="5%">
            <table width="100%" height="5%" border="0" cellpadding="1" cellspacing="4" align="center" valign="top" bgcolor="#F7F7F7">
                <tr>
                    <td width="50%"><b>&nbsp;<%= txtExibicao %></b></td>
                    <%
                    ' Verifica se está ativo a exibição da ordem dos produtos
                    If exibeOrdemProd = True And NOT RS_Produto.Eof Then
                    %>
                    <form method="GET" name="formOrderBy" action="">
                    <%
                    ' Resgata todos os parâmetros que estão na string para nova postagem
                    FOR EACH count in request.querystring
                        If request.querystring.item(count) <> "" And request.querystring.key(count) <> "orderby" Then
                    %>
                            <INPUT TYPE="hidden" NAME="<%=request.querystring.key(count)%>" VALUE="<%=request.querystring.item(count)%>">
                    <%
                        End if
                    NEXT
                    %>        
                    <td width="50%" align="right"><%= Application("MiddleTxtOrdemPor") %>&nbsp;<% Call MontaCombo_orderby("orderby",Request("orderby"),"parcial") %></td>
                    </form>
                    <%
                    End If
                    %>
                </tr>
            </table>
        </td>
    </tr>
<%
End if
%>
    <tr>
        <td height="90%">
<%
'Se a consulta retornar vazia imprime a informação de nada encontrada
If RS_Produto.Eof Then 
%>
    <table width="100%" height="90%" border="0" cellpadding="0" cellspacing="10" align="center" valign="top">
        <tr>
            <td style="padding-top:2px;padding-bottom:2px;padding-left:10px;padding-right:10px"><%= Application("FuncTxtNaoEncontrado")%></td>
        </tr>
<%
'Se houver resultad para a consulta os produtos serão listados
Else
%>
    <table width="100%" height="90%" border="0" cellpadding="0" cellspacing="10" align="center" valign="top">
        <tr>
<%
    'Página atual
    RS_Produto.AbsolutePage = CInt(PaginaCorrente)

    'Zera os contadores    
    Coluna = 0	
    Contador = 0
    'Loop para exibição dos produtos
    Do Until RS_Produto.AbsolutePage <> CInt(PaginaCorrente) OR RS_Produto.EOF

        'Converte para R$ caso o cadastro seja em outra moeda
        Set DadosCambio = Server.CreateObject("ADODB.Recordset")
        DadosCambio.CursorLocation = 3
        DadosCambio.CursorType = 0
        DadosCambio.LockType =  1
            DadosCambio.Open "SELECT simbolo_moeda, valor_moeda FROM IdiomaseCambios WHERE sigla_idioma='"&varLang&"'" , Conexao
            vlcambio = DadosCambio("valor_moeda")
            abvmoeda = DadosCambio("simbolo_moeda")
            Session("abvmoeda") = abvmoeda
        'Fecha e libera da memória o objeto de Recordset
        DadosCambio.close
        Set DadosCambio = Nothing

        valor_produto = FormatNumber(RS_Produto("preco_unitario")/(vlcambio),2)

%>
            <td valign="top" align="center" height="25%" width="25%">
                <div id="go">
                <ul id="gop">
                <li>
                <a href="produtos_descricao.asp?lang=<%=varLang%>&codigo_produto=<%=RS_Produto("codigo_produto")%>"><img src="<%= RS_Produto("img_produtoPQN") %>" border="0"></a>
                <a href="produtos_descricao.asp?lang=<%=varLang%>&codigo_produto=<%=RS_Produto("codigo_produto")%>" class="TXTproduto"><%= RS_Produto("nome_produto") %><br>
                <%If RS_Produto("quantidade_produto") > 0 Then%>
                    <%If RS_Produto("preco_unitario") > 0 then %>
                        <span class="TXTproduto">
                        <%If pegavalor_promocao(codigo_produto,RS_Produto) < FormatNumber(valor_produto) Then%>
                            <b><%If pegavalor_promocao(codigo_produto,RS_Produto) <> "" Then%><%=Application("MiddleTxtDe")%> <s><%=abvmoeda%>&nbsp;&nbsp;<%= FormatNumber(valor_produto) %></s><%Else%><%=abvmoeda%>&nbsp;&nbsp;<%= FormatNumber(valor_produto) %><%End if%></b><br>
                        
                            <%If pegavalor_promocao(codigo_produto,RS_Produto) <> "" Then%>
                                <b><%=Application("MiddleTxtPor")%> <%=abvmoeda%>&nbsp;&nbsp;<%=FormatNumber(pegavalor_promocao(valor_produto,RS_Produto)) %></b><br>
                                <% Response.write Application("MiddleTxtValido") & "&nbsp;" & RS_Produto("data_fim")%>
                            <%End if%>
                        <%Else%>
                            <%If valor_produto <> "" Then%>
                                <b><%=abvmoeda%>&nbsp;&nbsp;<%= FormatNumber(valor_produto) %><b>
                            <%End if%>
                        <%End if%>
                    <%End if%>
                <%Else%>
                    <b><%=Application("MiddleTxtNaoDisponivel")%></b>
                <%End if%>
                </a>
                <br>
                <br>
                </li>
            </td>
<%
        'Define o numero de Produtos por linha
        Coluna=Coluna+1
        If Coluna >=4 Then
            Coluna=0
%>
        </tr>
        <tr> 
<%          'Atualiza o contador
            Contador = Contador + 1
            RS_Produto.MoveNext    
        Else
            'Atualiza o contador
            Contador = Contador + 1
            RS_Produto.MoveNext    
        End If
    Loop 

End If
%>
        </tr>
    </table>
<%
' FctStrAdicional monta uma string com as variavies relacionados ao produto/busca.
' - As variaveis (CODIGO_SUBCATEGORIA/NOME_SUBCATEGORIA/CODIGO_CATEGORIA/NOME_CATEGORIA/PROCURA).

VarStrAdicional = VarStrAdicional & "&lang=" & varLang

If Request("codigo_subcategoria") <> "" Then
    VarStrAdicional = VarStrAdicional & "&codigo_subcategoria=" & Request("codigo_subcategoria")
End If

If Request("nome_subcategoria") <> "" Then
    VarStrAdicional = VarStrAdicional & "&nome_subcategoria=" & Request("nome_subcategoria")
End If 

If Request("codigo_categoria") <> "" Then
    VarStrAdicional = VarStrAdicional & "&codigo_categoria=" & Request("codigo_categoria")
End If

If Request("nome_categoria") <> "" Then
    VarStrAdicional = VarStrAdicional & "&nome_categoria=" & Request("nome_categoria")
End If

If Request("tipo_busca") <> "" Then
    VarStrAdicional = VarStrAdicional & "&tipo_busca=" & Request("tipo_busca")
End If

If Request("produto") <> "" Then
    VarStrAdicional = VarStrAdicional & "&produto=" & Request("produto")
End If

If Request("procura") <> "" Then
    VarStrAdicional = VarStrAdicional & "&procura=" & Request("procura")
End If

If Request("codigo_marca") <> "" Then
    VarStrAdicional = VarStrAdicional & "&codigo_marca=" & Request("codigo_marca")
End If

If Request("orderby") <> "" Then
    VarStrAdicional = VarStrAdicional & "&orderby=" & Request("orderby")
End If
%>
        </td>
    </tr>
    <tr height="5%">
        <td align="right"><img src="config/templates/<%=varLang%>/<%=varSkin%>/regua1x1.gif" height="20">
        <% If rodape <> "no" Then
               'Chama a páginação
               Call paginacao(VarStrAdicional,VarTotalPaginas,VarIntervalo)
           End if
        %>        
        </td>
    </tr>
</table>
<form method="POST" name="produto" action="carrinho.asp">
    <input type="hidden" name="codigo_produto" value="">
    <input type="hidden" name="codigo_categoria" value="">
    <input type="hidden" name="ato" value="FIM">
    <input type="hidden" name="mode" value="comprar">
</form>
<%
'Fecha e libera da memória o objeto de Recordset
RS_Produto.Close
Set RS_Produto = Nothing

End Sub

'########################################################################################################
'--> FIM SUB exibe_produtos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION paginacao
' - Monta os links de páginação.
' - Usado onde ha necessidade de paginação
'########################################################################################################

Function paginacao(FctStrAdicional,FctTotalPaginas,FctIntervalo)

    'Define qual é a página corrente
    If Request.QueryString("PN") = "" THEN
        PaginaCorrente = 1
    Else
        PaginaCorrente = Request.QueryString("PN")
    End If
    
    'Captura o endereço da página corrente
    var_url = Request.serverVariables("SCRIPT_NAME")
    arrayx = split(var_url,"/") 

    'Monta URL da paginação a partir do endereço capturado
    do while I < ubound(arrayx) 
    I = I + 1 

        If len(trim(VarURLPaginacao)) = 0 then 
            VarURLPaginacao = arrayx(i) 
        End If 

        If arrayx(I) > VarURLPaginacao then 
            VarURLPaginacao = arrayx(I) 
            lngIndexMaiorValor = I 
            exit do
        End If 
    loop 

    If RIGHT(VarURLPaginacao,4) <> ".asp" Then
        VarURLPaginacao = ""
    End If

    Flag = INT(FctTotalPaginas / FctIntervalo) 
    Flag1 =  INT(PaginaCorrente / FctIntervalo)
    PI = Flag1 * FctIntervalo

    If PI = 0 THEN
        PI = 1
    End If

    PF = PI + FctIntervalo - 1
    If CInt(Flag1) >= CInt(1) THEN
        Response.Write "<a href="&VarURLPaginacao&"?PN=" &  PI - 1 & "&FctIntervalo=" & FctIntervalo & FctStrAdicional & "  ><B>"&Application("MiddleTxtAnterior")&"</B>&nbsp;.</a>"
    End If

    If (PaginaCorrente - 1) >= "1" Then
        Response.Write "<a href="&VarURLPaginacao&"?PN=" &  Request("PN") - 1  & "&FctIntervalo=" & FctIntervalo & FctStrAdicional & "  ><B>"&Application("MiddleTxtAnterior")&"</B></a>&nbsp;."
    End If

    FOR I = PI TO PF
      If CInt(I) <= CInt(FctTotalPaginas) THEN
         If CInt(PaginaCorrente) = CInt(I) THEN
            response.write "<b>" & I & "</b>&nbsp;.&nbsp;"
         Else
            response.write "<a href="&VarURLPaginacao&"?PN=" & I & "&FctIntervalo=" & FctIntervalo & FctStrAdicional & "  ><b>" & I & "</b></a>&nbsp;.&nbsp;"
            FctPaginaAtual = Cint(PaginaCorrente)
         End If
      End If
    NEXT 

    If CDbl(PaginaCorrente) < CDbl(FctTotalPaginas) Then
        Response.Write "<a href="&VarURLPaginacao&"?PN=" &  FctPaginaAtual + 1 & "&FctIntervalo=" & FctIntervalo & FctStrAdicional &"  ><B>"&Application("MiddleTxtProxima")&"</B></a>&nbsp;"
    End If

    If (CInt(Flag1) < CInt(Flag)) THEN
        If CInt(PF) <> CInt(FctTotalPaginas) THEN
         Response.Write "<a href="&VarURLPaginacao&"?PN=" &  PF + 1 & "&FctIntervalo=" & FctIntervalo & FctStrAdicional &"  ><B>"&Application("MiddleTxtProxima")&"</B>&nbsp;</a>"
        End If
    End If

End Function

'########################################################################################################
'--> FIM SUB paginacao
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Gerar_id_transacao
'Esta Funcão gera um ID único da transação a partir do Session.sessionID.
'O ID unico define o nome do arquivo XML de cada pedido e serve como identificação do pedido como id_transacao.
'########################################################################################################

function Gerar_id_transacao(VarSession_userid)

    If not isnumeric(VarSession_userid) then
        response.write("idloja deve ser numérico")
        exit function
    End If

    id_usuario=mid(VarSession_userid,4,6)

    hora=right("00" & hour(time),2)
    minuto=right("00" & minute(time),2)
    segundo=right("00" & second(time),2)

    hhmmssd=hora&minuto&segundo&proximo

    d0=DateSerial (year(date), "1", "1")
    datajuliana=right("000" & (Date - d0 + 1),3)
    
    'Define o ID da transação
    Gerar_id_transacao=id_usuario&datajuliana&hhmmssd&loja

End function

'Adicional de segurança para evitar que o ID da transação seja repetido.
function proximo

    If application("dc")=9 then
	    application("dc")=0
    Else
        application("dc")=application("dc")+1
    End If
    proximo=application("dc")
 
End Function

'########################################################################################################
'--> FIM SUB Gerar_id_transacao
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION deleteItem_xmlpedidos
' - Esta FUNCTION apaga produtos
' - Ação definida no botão EXCLUIR ou atualização de produtos com quantidade 0 no carrinho
' - Esta FUNCTION é chamada no arquivo CARRINHO.ASP e FUNCOES.ASP na FUNCTION altera_xmlpedidos
'########################################################################################################

Function deleteItem_xmlpedidos(FctId_transacao, codigo_produto, FctAdicional)

'Abre o arquixo XML postado na variavel FctId_transacao
Call abre_xmlpedido(FctId_transacao, objXML, objRoot) 
codigo_produto_real = codigo_produto
'Limpa o código do produto
'O código bruto do produto inclui codigos de COR e TAMANHO
'codigo_produto = Split(codigo_produto,"_")
	set base=objRoot.selectsinglenode("dados_pedido")
		set objnode = base.selectsinglenode("produto[@codigo_produto='"&codigo_produto_real&"']")

        peso_total_temp = base.getAttribute("peso_total") 
        peso_total_produto_temp = objnode.getAttribute("peso_parcial") 
        peso_temp =   cdbl(peso_total_temp) - cdbl(peso_total_produto_temp)
        base.setAttribute "peso_total",formatNumber(peso_temp) 

        valor_subtotalTemp = base.getAttribute("valor_subtotal") 
        valor_parcialTemp = objnode.getAttribute("total_parcial")

        'Caso o produto é retirado pelo BOTÃO excluir é subtraído do Valor_SubTotal o Valor_parcial
        If FctAdicional = "atualiza_quantidades" Then
        valor_subtotalTemp =  valor_subtotalTemp - valor_parcialTemp
        End If

        base.setAttribute "valor_subtotal",formatNumber(valor_subtotalTemp)

        'Salva o arquivo com as alterações iniciais
        'Essa operação é necessária devido para o cálculo de frete
        objXML.save(Application("DiretorioPedidos")&FctId_transacao&".xml")
        
        If base.getAttribute("opcao_frete") <> "0" Then
            varNovoFrete = Atualiza_CEP(base.getAttribute("cep_frete"),base.getAttribute("pais_frete"),base.getAttribute("peso_total"),base.getAttribute("opcao_frete"))
            base.setAttribute "valor_frete",formatNumber(varNovoFrete)  
        End if

        valor_freteTemp = base.getAttribute("valor_frete")
        
        valor_totalTemp = cdbl(valor_subtotalTemp) + cdbl(valor_freteTemp)
        base.setAttribute "valor_total",formatNumber(valor_totalTemp) 

		base.removechild(objnode)
        
        'Salva o arquivo com as alterações finais (cálculo do frete)
		objXML.save(Application("DiretorioPedidos")&FctId_transacao&".xml")

		set objnode = Nothing
	set base = Nothing

'Fecha arquivo XML
Call fecha_xmlpedido(FctId_transacao) 

End Function

'########################################################################################################
'--> FIM SUB deleteItem_xmlpedidos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION altera_xmlpedidos
' - Esta Function aplica alterações globais no arquivo XML  
' - As alterações são: Inserção do frete, Alteração de quantidade
' - Produtos com quantidade = 0 são retirados pela FUNCTION deleteItem_xmlpedidos
' - Esta FUNCTION é chamada no arquivo carrinho.asp
'########################################################################################################
Function altera_xmlpedidos(FctId_transacao)

'Abre o arquixo XML postado na variavel FctId_transacao
Call abre_xmlpedido(FctId_transacao, objXML, objRoot) 

codigo_produtoTemp = Replace(Request("codigo_produto")," ","")
VarCodigo_produto = split(codigo_produtoTemp,",")
quantidade_produto = split(Request("quantidade_produto"),",")

'Define objeto de consulta ao nó
Set objAtualizaPedido = objRoot.selectSingleNode("dados_pedido[@id_transacao="&FctId_transacao&"]")

    ' Resgata o peso atual do pedido
    peso_totalAntigo = objAtualizaPedido.getAttribute("peso_total")

    ' Zero o peso no XML do pedido para recalculo
    objAtualizaPedido.setAttribute "peso_total","0"
    objXML.save (Application("DiretorioPedidos")&FctId_transacao&".xml")

    'Loop para os produtos inseridos no arquivo XML
    For i=0 to Ubound(VarCodigo_produto)

    If Instr(VarCodigo_produto(i),"_") <> 0 Then
        Set objAtualizaProduto = objRoot.SelectSingleNode("dados_pedido/produto[@codigo_produto='"&VarCodigo_produto(i)&"']")
    Else
        Set objAtualizaProduto = objRoot.SelectSingleNode("dados_pedido/produto[@codigo_produto="&VarCodigo_produto(i)&"]")
    End If

            Codigo_produtoRealTemp = Split(VarCodigo_produto(i),"_")
    
            If Trim(quantidade_produto(i)) =  "0" Then
            Session("conta_item") = Session("conta_item") & VarCodigo_produto(i) & ","

                If pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","Estoque") = "sim" Then
                    Call Atualiza_Estoque(Codigo_produtoRealTemp(0),"delete",objAtualizaProduto.getAttribute("quantidade_produto"),"0")
                End If

            Else

                If pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","Estoque") = "sim" Then
                        Call Atualiza_Estoque(Codigo_produtoRealTemp(0),"update",objAtualizaProduto.getAttribute("quantidade_produto"),Trim(quantidade_produto(i)))
                End If 

                    objAtualizaProduto.setAttribute "quantidade_produto",Trim(quantidade_produto(i))
                    peso_parcialTemp = objAtualizaProduto.getAttribute("peso_parcial")

                    'Altera o valor do campo peso_parcial
                    peso_unitarioTemp = objAtualizaProduto.getAttribute("peso_unitario")
                    peso_unitarioTemp = peso_unitarioTemp * quantidade_produto(i)
                    objAtualizaProduto.setAttribute "peso_parcial",formatNumber(peso_unitarioTemp)  

                    'Altera o valor do campo peso_total
                    peso_totalTemp = objAtualizaPedido.getAttribute("peso_total") 
                    peso_totalTemp = peso_totalTemp + peso_unitarioTemp
                    objAtualizaPedido.setAttribute "peso_total",formatNumber(peso_totalTemp)

                    'Altera o valor do campo total_parcial
                    preco_unitarioTemp = objAtualizaProduto.getAttribute("preco_unitario") 
                    preco_unitarioTemp = cdbl(preco_unitarioTemp * quantidade_produto(i))
                    objAtualizaProduto.setAttribute "total_parcial",formatNumber(preco_unitarioTemp)  

                    'Calcula o valor que será inserido nos campos Valor_subtotal/Valor_total
                    VarSomaValorSubtotal = preco_unitarioTemp + VarSomaValorSubtotal
                    session("ResultValorSubtotal") = VarSomaValorSubtotal

                    'Calcula o valor que será inserido nos campos Peso_parcial/Peso_Total
                    session("ResultTotalPeso") = quantidade_produto(i) * peso_totalTemp

                    'Salva o arquivo com as alterações
                    objXML.save (Application("DiretorioPedidos")&FctId_transacao&".xml")

            End If

        Set objAtualizaProduto = Nothing 
    Next

    'Define as variaveis como vazio para garantir o re-uso
    VarSomaValorSubtotal    = ""
    valor_subtotalTemp      = ""

    valor_subtotalTemp = objAtualizaPedido.getAttribute("valor_subtotal") 
    VarSomaValorSubtotal = session("ResultValorSubtotal")

    If VarSomaValorSubtotal <> "" Then
        objAtualizaPedido.setAttribute "valor_subtotal",formatNumber(VarSomaValorSubtotal)
    Else
        objAtualizaPedido.setAttribute "valor_subtotal",formatNumber(0)
    End If

    If request("pais_frete") <> "" Then 
        objAtualizaPedido.setAttribute "pais_frete",request("pais_frete")
    End if
    If request("cep_frete") <> "" Then
        objAtualizaPedido.setAttribute "cep_frete",Replace(request("cep_frete"),"-","")
    End if
    If request("opcao_frete") <> "" Then 
        objAtualizaPedido.setAttribute "opcao_frete",request("opcao_frete")
    End if

    'Salva o arquivo com as alterações iniciais
    'Essa operação é necessária devido para o cálculo de frete
    objXML.save (Application("DiretorioPedidos")&FctId_transacao&".xml")
    
    If (objAtualizaPedido.getAttribute("opcao_frete") <> "0") Then
        varNovoFrete = Atualiza_CEP(objAtualizaPedido.getAttribute("cep_frete"),objAtualizaPedido.getAttribute("pais_frete"),objAtualizaPedido.getAttribute("peso_total"),objAtualizaPedido.getAttribute("opcao_frete"))
        objAtualizaPedido.setAttribute "valor_frete",formatNumber(varNovoFrete)
    End if

    If Session("opcao_frete") = empty Then
        objAtualizaPedido.setAttribute "opcao_frete","0"
    End If

    valor_freteTemp = objAtualizaPedido.getAttribute("valor_frete")

    If request("frete") <> "" And CDBL(peso_totalAntigo) = CDBL(objAtualizaPedido.getAttribute("peso_total")) Then
        frete = request("frete")
        
        'Verifica se o frete é gratuíto
        If Session("msgErroFrete") = "" And frete = "0,00" Then
            Session("Frete_gratuito") = "ativo"
        Else
            Session("Frete_gratuito") = empty
        End If
    
    else
        frete = valor_freteTemp
    end if

    objAtualizaPedido.setAttribute "valor_frete",formatNumber(frete)
    objAtualizaPedido.setAttribute "valor_total",formatNumber(VarSomaValorSubtotal + frete)

    'Salva o arquivo com as alterações finais (cálculo do frete)
    objXML.save (Application("DiretorioPedidos")&FctId_transacao&".xml")

    'Define as sessions como vázio para garantir o re-uso
    session("ResultValorSubtotal")  = ""
    session("ResultTotalPeso")      = ""

Set objAtualizaPedido = Nothing

If Session("conta_item") <> "" Then
	conta_item_ = MID(session("conta_item"),1,LEN(session("conta_item"))-1)
	conta_item = split(conta_item_,",")
	Session("conta_item") = ""
		For z=0 to Ubound(conta_item)
			Call deleteItem_xmlpedidos(FctId_transacao, conta_item(z),adicional)
		Next
        conta_item = ""
End If

'Fecha arquivo XML
Call fecha_xmlpedido(FctId_transacao) 

End Function 

'########################################################################################################
'--> FIM SUB altera_xmlpedidos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Altera_dados_pedidos
' - Esta function é chamada para inserir um novo produto no arquivo XML depois que o arquivo XML foi criado
' - Esta Function é chamada na SUB Cria_pedidoTemp
'########################################################################################################
Function Altera_dados_pedidos(FctExiste_produto,FctQtd)

            quantidade_produtoTemp = FctExiste_produto.getAttribute("quantidade_produto") + FctQtd
			FctExiste_produto.setAttribute "quantidade_produto",quantidade_produtoTemp

			preco_unitarioTemp = FctExiste_produto.getAttribute("preco_unitario") 
			peso_unitarioU = FctExiste_produto.getAttribute("peso_unitario") 
			total_soma = quantidade_produtoTemp * preco_unitarioTemp

            peso_unitarioTemp = peso_unitarioU * quantidade_produtoTemp

			FctExiste_produto.setAttribute "total_parcial",formatNumber(total_soma) 
			FctExiste_produto.setAttribute "total_parcial",formatNumber(total_soma) 


End Function

'########################################################################################################
'--> FIM SUB Altera_dados_pedidos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB abre_xmlpedido
' - Abre o arquivo XML da transação
'SUB fecha_xmlpedido
' - o arquivo XML da transação
'########################################################################################################
Sub abre_xmlpedido(FctId_transacao, objXML, objRoot) 

    set objXML = CreateObject("Microsoft.XMLDOM")
        objXML.preserveWhiteSpace = False
        objXML.async = False
        objXML.validateOnParse = True
        objXML.resolveExternals = True
        objXML.load (Application("DiretorioPedidos")&FctId_transacao&".xml")
    Set objRoot = objXML.documentElement

End Sub

Sub fecha_xmlpedido(FctId_transacao) 

 set objXML = Nothing
 Set objRoot = Nothing

End Sub

'########################################################################################################
'--> FIM SUB abre_xmlpedido
'--> FIM SUB fecha_xmlpedido
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_pedidoTemp
' - Esta FUNCTION cria o arquivo dos pedidos e insere nele os produtos colocados no carrinho
' - Esta FUNCTION é chamada no arquivo CARRINHO.ASP
'########################################################################################################

Sub Cria_pedidoTemp(FctId_transacao,FctCodigo_produto,FctCodigo_categoria,Fctnome_produto,Fctquantidade_produto,Fctpreco_unitario,FctDesconto,FctPeso,FctCor,FctTamanho,FctLang,FctSimbolo_Moeda,FctValor_Cambio)

set docxml=CreateObject("microsoft.xmldom")
'Verifica se ja existe o arquivo XML da transação
existe_pedidoTemp = docxml.load(Application("DiretorioPedidos")&session("id_transacao")&".xml")
    
    'Se existir chama o arquivo
	If existe_pedidoTemp=True then
		set pedido=docxml.documentElement
		Set dados_pedido = pedido.SelectSingleNode("dados_pedido")
    'Se não existir, cria-se.        
	Else
		set pedido=docxml.createElement("pedido")
		docxml.appendchild(pedido)
		set dados_pedido=docxml.createElement("dados_pedido")
		pedido.appendchild(dados_pedido)
		dados_pedido.SetAttribute "siglaidioma",session("requestIdioma")
		dados_pedido.SetAttribute "simbolo_moeda",FctSimbolo_Moeda
		dados_pedido.SetAttribute "valor_cambio",FctValor_Cambio
		dados_pedido.SetAttribute "sigla_idioma",session("requestIdioma")
		dados_pedido.SetAttribute "id_transacao",FctId_transacao
		dados_pedido.SetAttribute "inicio_transacao",NOW
		dados_pedido.SetAttribute "valor_frete",0
		dados_pedido.SetAttribute "valor_subtotal",0
		dados_pedido.SetAttribute "valor_total",0

		dados_pedido.SetAttribute "peso_total",0
        dados_pedido.SetAttribute "opcao_frete",0
		dados_pedido.SetAttribute "pais_frete",""
        dados_pedido.SetAttribute "cep_frete",0
        dados_pedido.SetAttribute "forma_pagamento",0
        dados_pedido.SetAttribute "tipo_taxa_adicional",""
        dados_pedido.SetAttribute "taxa_adicional",0
        dados_pedido.SetAttribute "num_parcelas",1
		dados_pedido.SetAttribute "logado",0
		dados_pedido.SetAttribute "user_id",0
        dados_pedido.SetAttribute "ip_usado", request.ServerVariables("REMOTE_ADDR")
        dados_pedido.SetAttribute "status_pedido","pendente"
	End If
    'Verifica se o produto foi inserido no carrinho anteriomente.
	set existe_produto = pedido.SelectSingleNode("dados_pedido/produto[@codigo_produto='"&FctCodigo_produto&"']")
    'Altera a quantida do produto
	If Not existe_produto Is Nothing Then

	    'Atribui valor de quantidade para o produto se postado    
	    If request("quantidade_produto") <> "" Then
	        qtd = request("quantidade_produto")
        'Se não postado o valor será de 1 (uma unidade)
	    Else
	        qtd = 1
	    End If

        'Chama funcão para alterar a quantidade produtos, caso a chamado ao carrinho seja inclusão de um mesmo produto.
	    Call Altera_dados_pedidos(existe_produto,qtd)

    'Insere um novo produto
	Else

        'Grava valores de atributos no arquivo XML
		set produto=docxml.createelement("produto")
		produto.SetAttribute "codigo_produto",FctCodigo_produto
		produto.SetAttribute "codigo_categoria",FctCodigo_categoria
		produto.SetAttribute "codigo_cor",FctCor
		produto.SetAttribute "codigo_tamanho",FctTamanho
		produto.SetAttribute "codigo_produto",FctCodigo_produto
		produto.SetAttribute "nome_produto",Fctnome_produto
		produto.SetAttribute "preco_unitario",formatNumber(Fctpreco_unitario)
		produto.SetAttribute "desconto",FctDesconto
		produto.SetAttribute "peso_unitario",formatNumber(FctPeso,3)
		produto.SetAttribute "peso_parcial",formatNumber(FctPeso,3)
		produto.SetAttribute "total_parcial",formatNumber(Fctpreco_unitario)
		produto.SetAttribute "quantidade_produto",Fctquantidade_produto
		dados_pedido.appendchild(produto)
		set produto=Nothing
		
	End If

    Set objAtualizaPedido = pedido.selectSingleNode("dados_pedido[@id_transacao="&FctId_transacao&"]")
        valor_subtotalTemp = objAtualizaPedido.getAttribute("valor_subtotal") 
        peso_unit = objAtualizaPedido.getAttribute("peso_total") 
        ResultTotalPeso = FctPeso + peso_unit
        objAtualizaPedido.setAttribute "peso_total",formatNumber(ResultTotalPeso,3)  
        If valor_subtotalTemp = "" Then
            valor_subtotalTemp = "0"
        End if
        VarSomaValorSubtotal = FormatNumber(CDbl(Fctpreco_unitario) + CDbl(valor_subtotalTemp))
        valor_total_finalTemp = FormatNumber(CDbl(VarSomaValorSubtotal))
        objAtualizaPedido.setAttribute "valor_subtotal",VarSomaValorSubtotal

        
        'Salva o arquivo com as alterações iniciais
        'Essa operação é necessária devido para o cálculo de frete
        docxml.save(Application("DiretorioPedidos")&session("id_transacao")&".xml")

        If objAtualizaPedido.getAttribute("opcao_frete") <> "0" Then
            varNovoFrete = Atualiza_CEP(objAtualizaPedido.getAttribute("cep_frete"),objAtualizaPedido.getAttribute("pais_frete"),objAtualizaPedido.getAttribute("peso_total"),objAtualizaPedido.getAttribute("opcao_frete"))
            objAtualizaPedido.setAttribute "valor_frete",formatNumber(varNovoFrete)  
        End if

        valor_freteTemp0 = dados_pedido.getAttribute("valor_frete")
        valor_total_final = FormatNumber(CDbl(valor_freteTemp0) + CDbl(valor_total_finalTemp))

        objAtualizaPedido.setAttribute "valor_total",valor_total_final
    Set objAtualizaPedido = Nothing
'Salva o arquivo com as alterações finais (cálculo do frete)
docxml.save(Application("DiretorioPedidos")&session("id_transacao")&".xml")
set docxml = Nothing
End sub
'########################################################################################################
'--> FIM SUB Cria_pedidoTemp
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION CarregaGrava_dados_pedido
' - Esta FUNCTION captura os dados do formulário de cadastro e do arquivo XML para gravação no banco de dados
' - Esta FUNCTION é chamada no arquivo INICIA_TRANSACAO.ASP
'########################################################################################################
Function CarregaGrava_dados_pedido(FctId_transacao, objXML, objRoot, FctAdicional)

    If FctAdicional <> "alterar" And FctAdicional <> "novocadastro" Then
        'Abre arquivo XML da transação.    
        Call abre_xmlpedido(FctId_transacao, objXML, objRoot)
        'Define a raiz do documento
        Set raiz = objXML.documentElement
        Set raiz_dados_pedido = raiz.selectSingleNode("dados_pedido[@id_transacao="&FctId_transacao&"]")
        Set raiz_dados_produto = objXML.getElementsByTagName("dados_pedido[@id_transacao="&FctId_transacao&"]/produto") 
        'Define o número de atributos registrados
        n_raiz_dados_produto = raiz_dados_produto.length
        'Captura os valos dos atribuitos
        VARinicio_transacao = CorrigeDataHora(raiz_dados_pedido.getAttribute("inicio_transacao"))
        VARpeso_total       = raiz_dados_pedido.getAttribute("peso_total") 
        VARopcao_frete      = raiz_dados_pedido.getAttribute("opcao_frete") 
        VARforma_pagamento  = raiz_dados_pedido.getAttribute("forma_pagamento") 
        VARcep_frete        = raiz_dados_pedido.getAttribute("cep_frete") 
        VARboleto_tipo      = raiz_dados_pedido.getAttribute("BoletoTipo") 
        VARsigla_idioma     = raiz_dados_pedido.getAttribute("sigla_idioma")

        VARtipo_taxa_adicional = raiz_dados_pedido.getAttribute("tipo_taxa_adicional")
        VARtaxa_adicional      = raiz_dados_pedido.getAttribute("taxa_adicional")
        VARnum_parcelas        = raiz_dados_pedido.getAttribute("num_parcelas")

        ' Retira virgula e repõe por ponto, para evitar problemas no UPDATE

        VARvalor_subtotal   = replace(raiz_dados_pedido.getAttribute("valor_subtotal"),".","")
        VARvalor_subtotal   = replace(VARvalor_subtotal,",",".")

        VARvalor_frete      = replace(raiz_dados_pedido.getAttribute("valor_frete"),".","")
        VARvalor_frete      = replace(VARvalor_frete,",",".")

        VARvalor_total      = replace(raiz_dados_pedido.getAttribute("valor_total"),".","")
        VARvalor_total      = replace(VARvalor_total,",",".")

        VARpeso_total       = replace(raiz_dados_pedido.getAttribute("peso_total"),".","")
        VARpeso_total       = replace(VARpeso_total,",",".")

        
        'Verifica se os dados de entrega são os mesmos de cobrança
        If Request("cep_entrega") <> "" Then
            cod_coluna = "_entrega"
        Else
            cod_coluna = "_cobranca"
        End If
    Else
        'Verifica se os dados de entrega são os mesmos de cobrança
        If Request("cep_entrega") <> "" Then
            cod_coluna = "_entrega"
        Else
            cod_coluna = "_cobranca"
        End If
    End If

    ' Verifica se a sessão do usuário não está vazia
    If Session("user_ID") = "" Or IsNull(Session("user_ID")) Or Request("nome_cobranca") = "" Or Request("logradouro_cobranca") = "" Then
        Response.redirect "carrinho_vazio.asp?msg=probcad"
        Response.End
    End If

    'Cria objeto de consulta a tabela de usuários
    Set RS_Usuario = CreateObject("ADODB.Recordset")
    Set RS_Usuario.ActiveConnection = Conexao
    RS_Usuario.CursorLocation = 3
    RS_Usuario.CursorType = 0
    RS_Usuario.LockType =  3

    RS_Usuario.Open "SELECT user_id, chave, razaosocial_cobranca, cnpj_cobranca, inscricaoestadual_cobranca, nome_cobranca, cpf_cobranca, rg_cobranca, data_nascimento_cobranca, logradouro_cobranca, numero_cobranca, complemento_cobranca, bairro_cobranca, cidade_cobranca, estado_cobranca, cep_cobranca, pais_cobranca, ddd_cobranca, telefone_cobranca, razaosocial_entrega, cnpj_entrega, inscricaoestadual_entrega, nome_entrega, cpf_entrega, rg_entrega, data_nascimento_entrega, logradouro_entrega, numero_entrega, complemento_entrega, bairro_entrega, cidade_entrega, estado_entrega, cep_entrega, pais_entrega, ddd_entrega, telefone_entrega, email_entrega, cookieID, data_criacao FROM Usuarios WHERE user_id = '" & session("user_ID") & "'", Conexao
    
    'Se o usuário não existir é feito um novo cadastro
    If RS_Usuario.EOF Then
        
        'Novo registro       
        RS_Usuario.Addnew
		RS_Usuario("user_id") = Session("user_ID")
		RS_Usuario("chave") = Request("senha1")

    End If

    'Dados Usuário Cobrança
    RS_Usuario("razaosocial_cobranca") = Request("razaosocial_cobranca")
    RS_Usuario("cnpj_cobranca") = Request("cnpj_cobranca")
    RS_Usuario("inscricaoestadual_cobranca") = Request("inscricaoestadual_cobranca")
    RS_Usuario("nome_cobranca") = Request("nome_cobranca")
    If Request("data_nascimento_cobranca") <> "" Then
        RS_Usuario("data_nascimento_cobranca") = Request("data_nascimento_cobranca")
    End If
    RS_Usuario("cpf_cobranca") = Request("cpf_cobranca")
    RS_Usuario("rg_cobranca") = Request("rg_cobranca")
    RS_Usuario("logradouro_cobranca") = Request("logradouro_cobranca")
    RS_Usuario("numero_cobranca") = Request("numero_cobranca")
    RS_Usuario("complemento_cobranca") = Request("complemento_cobranca")
    RS_Usuario("bairro_cobranca") = Request("bairro_cobranca")
    RS_Usuario("cep_cobranca") = Request("cep_cobranca")
    RS_Usuario("cidade_cobranca") = Request("cidade_cobranca")
    RS_Usuario("estado_cobranca") = Request("estado_cobranca")
    RS_Usuario("pais_cobranca") = Cria_Combo_Paises(Request("pais_cobranca"),"codifica","")
    RS_Usuario("ddd_cobranca") = Request("ddd_cobranca")
    RS_Usuario("telefone_cobranca") = Request("telefone_cobranca")
    
    'Dados Usuário Entrega
    RS_Usuario("razaosocial_entrega") = Request("razaosocial" & cod_coluna)
    RS_Usuario("cnpj_entrega") = Request("cnpj" & cod_coluna)
    RS_Usuario("inscricaoestadual_entrega") = Request("inscricaoestadual" & cod_coluna)
    RS_Usuario("nome_entrega") = Request("nome" & cod_coluna)
    If Request("data_nascimento" & cod_coluna) <> "" Then
        RS_Usuario("data_nascimento_entrega") = Request("data_nascimento" & cod_coluna)
    End If
    RS_Usuario("cpf_entrega") = Request("cpf" & cod_coluna)
    RS_Usuario("rg_entrega") = Request("rg" & cod_coluna)
    RS_Usuario("logradouro_entrega") = Request("logradouro" & cod_coluna)
    RS_Usuario("numero_entrega") = Request("numero" & cod_coluna)
    RS_Usuario("complemento_entrega") = Request("complemento" & cod_coluna)
    RS_Usuario("bairro_entrega") = Request("bairro" & cod_coluna)
    RS_Usuario("cep_entrega") = Request("cep" & cod_coluna)
    RS_Usuario("cidade_entrega") = Request("cidade" & cod_coluna)
    RS_Usuario("estado_entrega") = Request("estado" & cod_coluna)
    RS_Usuario("pais_entrega") = Cria_Combo_Paises(Request("pais" & cod_coluna),"codifica","")
    RS_Usuario("ddd_entrega") = Request("ddd" & cod_coluna)
    RS_Usuario("telefone_entrega") = Request("telefone" & cod_coluna)
    If cod_coluna = "_entrega" Then
        RS_Usuario("email_entrega") = Request("email_entrega")
    Else
        RS_Usuario("email_entrega") = Session("user_ID")
    End If
    
    'Autoriza o recebimento de newsletter?
    If Session("autorizo_newsletter") = "1" Then
        'Checa se o e-mail não está cadastrado
        If Not VerificaExistenciaDado("email","Newsletter","email","'"&Session("user_ID")&"'") Then
            id_unico = session("id_transacao")&CALCMD5(email)
            'Cadastra o e-mail
            Set InsertNewsletter = Server.CreateObject("adodb.recordset")
            InsertNewsletter.Open "SELECT Newsletter.* FROM Newsletter",conexao,3,3
            InsertNewsletter.AddNew
            InsertNewsletter("nome") = Request("nome_cobranca")
            InsertNewsletter("email") = Session("user_ID")
            InsertNewsletter("id_unico") = id_unico
            InsertNewsletter("autorizo_newsletter") = "0"
            InsertNewsletter("ip_usado") = request.ServerVariables("REMOTE_ADDR")
            InsertNewsletter("data_cadastro") = date
            InsertNewsletter.Update
            'Dispara e-mail solicitando a confirmação do cadastro
            Call Envia_mail_confirmacao(Session("user_ID"),Request("nome_cobranca"),id_unico)
        End If
    End If

    'Insere os dados na tabela.
    RS_Usuario.Update

    'Fecha e libera da memória o objeto de Recordset
    RS_Usuario.Close
    Set RS_Usuario = Nothing

    If FctAdicional <> "alterar" And FctAdicional <> "novocadastro" Then

        If session("registrado") <> FctId_transacao & "_ok" Then   

            If session("resgistroPedido") = "" Then
            
                'Cria objeto de consulta a tabela de pedidos
                Set RS_Pedido = CreateObject("ADODB.Recordset")
                Set RS_Pedido.ActiveConnection = Conexao
                RS_Pedido.CursorLocation = 3
                RS_Pedido.CursorType = 0
                RS_Pedido.LockType =  3

                RS_Pedido.Open "SELECT Pedidos.* FROM Pedidos", Conexao
                'Novo registro
                RS_Pedido.Addnew

                'Verifica o número do pedido
                Set RS_PedidoNovo = Server.CreateObject("ADODB.Recordset")
                RS_PedidoNovo.CursorLocation = 3
                RS_PedidoNovo.CursorType = 0
                RS_PedidoNovo.LockType = 3

                RS_PedidoNovo.Open "SELECT MAX(codigo_pedido) AS novo_codigo_pedido FROM Pedidos" , Conexao
                'Checa se existem categorias no banco de dados 
                If IsNull(RS_PedidoNovo("novo_codigo_pedido")) Then
                    novo_codigo_pedido = Application("NumPedidoInicial")
                Else 
                    novo_codigo_pedido = RS_PedidoNovo("novo_codigo_pedido") + 1
                End If

                'Dados Pedido
                RS_Pedido("codigo_pedido") = novo_codigo_pedido
                RS_Pedido("data_pedido_inicio") = VARinicio_transacao
                RS_Pedido("data_pedido") = Now
                RS_Pedido("sessionID") = FctId_transacao
                RS_Pedido("user_ID") = Session("user_ID")
                RS_Pedido("ip_cliente") = request.ServerVariables("REMOTE_ADDR")
                
                'Dados Pedido Cobrança
                RS_Pedido("razaosocial_cobranca") = Request("razaosocial_cobranca")
                RS_Pedido("cnpj_cobranca") = Request("cnpj_cobranca")
                RS_Pedido("inscricaoestadual_cobranca") = Request("inscricaoestadual_cobranca")
                RS_Pedido("nome_cobranca") = Request("nome_cobranca")
                If Request("data_nascimento_cobranca") <> "" Then
                    RS_Pedido("data_nascimento_cobranca") = Request("data_nascimento_cobranca")
                End If
                RS_Pedido("cpf_cobranca") = Request("cpf_cobranca")
                RS_Pedido("rg_cobranca") = Request("rg_cobranca")
                RS_Pedido("logradouro_cobranca") = Request("logradouro_cobranca")
                RS_Pedido("numero_cobranca") = Request("numero_cobranca")
                RS_Pedido("complemento_cobranca") = Request("complemento_cobranca")
                RS_Pedido("bairro_cobranca") = Request("bairro_cobranca")
                RS_Pedido("cep_cobranca") = Request("cep_cobranca")
                RS_Pedido("cidade_cobranca") = Request("cidade_cobranca")
                RS_Pedido("estado_cobranca") = Request("estado_cobranca")
                RS_Pedido("pais_cobranca") = Cria_Combo_Paises(Request("pais_cobranca"),"codifica","")
                RS_Pedido("ddd_cobranca") = Request("ddd_cobranca")
                RS_Pedido("telefone_cobranca") = Request("telefone_cobranca")
                RS_Pedido("instrucoes") = Request("instrucoes")
            
                'Dados Pedido Entrega
                RS_Pedido("razaosocial_entrega") = Request("razaosocial" & cod_coluna)
                RS_Pedido("cnpj_entrega") = Request("cnpj" & cod_coluna)
                RS_Pedido("inscricaoestadual_entrega") = Request("inscricaoestadual" & cod_coluna)
                RS_Pedido("nome_entrega") = Request("nome" & cod_coluna)
                If Request("data_nascimento" & cod_coluna) <> "" Then
                    RS_Pedido("data_nascimento_entrega") = Request("data_nascimento" & cod_coluna)
                End If
                RS_Pedido("cpf_entrega") = Request("cpf" & cod_coluna)
                RS_Pedido("rg_entrega") = Request("rg" & cod_coluna)
                RS_Pedido("logradouro_entrega") = Request("logradouro" & cod_coluna)
                RS_Pedido("numero_entrega") = Request("numero" & cod_coluna)
                RS_Pedido("complemento_entrega") = Request("complemento" & cod_coluna)
                RS_Pedido("bairro_entrega") = Request("bairro" & cod_coluna)
                RS_Pedido("cep_entrega") = Request("cep" & cod_coluna)
                RS_Pedido("cidade_entrega") = Request("cidade" & cod_coluna)
                RS_Pedido("estado_entrega") = Request("estado" & cod_coluna)
                RS_Pedido("pais_entrega") = Cria_Combo_Paises(Request("pais" & cod_coluna),"codifica","")
                RS_Pedido("ddd_entrega") = Request("ddd" & cod_coluna)
                RS_Pedido("telefone_entrega") = Request("telefone" & cod_coluna)
                If cod_coluna = "_entrega" Then
                    RS_Pedido("email_entrega") = Request("email_entrega")
                Else
                    RS_Pedido("email_entrega") = Session("user_ID")
                End If

                'Dados Pedido
                RS_Pedido("subtotal") = VARvalor_subtotal
                RS_Pedido("taxa_envio") = VARvalor_frete
                RS_Pedido("total") = VARvalor_total
                RS_Pedido("tipo_frete") = VARopcao_frete
                RS_Pedido("codigo_frete") = "0"
                RS_Pedido("peso_total") = VARpeso_total
                RS_Pedido("forma_pagamento") = VARforma_pagamento
                RS_Pedido("tipo_taxa_adicional") = VARtipo_taxa_adicional
                RS_Pedido("taxa_adicional") = VARtaxa_adicional
                RS_Pedido("num_parcelas") = VARnum_parcelas
                RS_Pedido("boleto_tipo") = VARboleto_tipo
                RS_Pedido("sigla_idioma") = VARsigla_idioma
                RS_Pedido("cartao_encrypt") = ""
                RS_Pedido("atendido") = "0"
                RS_Pedido("pago") = "0"
                RS_Pedido("falha") = "0"
                RS_Pedido("num_remessa") = ""

                'Insere os dados na tabela.
                RS_Pedido.Update

                'Fecha e libera da memória o objeto de Recordset
                RS_Pedido.Close
                Set RS_Pedido = Nothing
                
                'Cria sessão de verificação de que os dados foram registrados
                session("resgistroPedido") = "Concluido"

            End If

            If session("resgistroPedidoItem") = "" Then

                'Cria objeto de consulta a tabela de Pedido_item
                Set RS_Pedido_Temp = CreateObject("ADODB.Recordset")
                Set RS_Pedido_Temp.ActiveConnection = Conexao
                RS_Pedido_Temp.CursorLocation = 3
                RS_Pedido_Temp.CursorType = 0
                RS_Pedido_Temp.LockType =  1

                If Application("TipoBanco") = "mysql" Then
					QueryPedidos = "SELECT codigo_pedido FROM Pedidos WHERE sessionID = '" & FctId_transacao & "' ORDER BY codigo_pedido DESC LIMIT 0,1"
				Else
					QueryPedidos = "SELECT TOP 1 codigo_pedido FROM Pedidos WHERE sessionID = '" & FctId_transacao & "' ORDER BY codigo_pedido DESC"
				End If

				RS_Pedido_Temp.Open QueryPedidos, Conexao

                VARcodigo_pedido = RS_Pedido_Temp("codigo_pedido")

                RS_Pedido_Temp.Close
                Set RS_Pedido_Temp = Nothing

                Set RS_Pedido_item = CreateObject("ADODB.Recordset")
                Set RS_Pedido_item.ActiveConnection = Conexao
                RS_Pedido_item.CursorLocation = 3
                RS_Pedido_item.CursorType = 0
                RS_Pedido_item.LockType =  3

                RS_Pedido_item.Open "SELECT Pedido_item.* FROM Pedido_item", Conexao

                'Loop para inserção dos itens do pedido na tabela Pedido_item 
                For i = 0 To (n_raiz_dados_produto - 1)

                    'Novo registro
                    RS_Pedido_item.Addnew

                    Set dados = raiz_dados_produto.item(i)

                    If InStr(dados.getAttribute("codigo_produto"),"_") <> 0 Then
                        tempCodigo_produto1 = Split(dados.getAttribute("codigo_produto"),"_")
                        tempCodigo_produto = tempCodigo_produto1(0)
                    Else
                        tempCodigo_produto = dados.getAttribute("codigo_produto")
                    End If
                
                    RS_Pedido_item("codigo_pedido") = VARcodigo_pedido
                    RS_Pedido_item("codigo_produto") = tempCodigo_produto
                    RS_Pedido_item("codigo_categoria") = dados.getAttribute("codigo_categoria")
                    RS_Pedido_item("codigo_cor") = dados.getAttribute("codigo_cor")
                    RS_Pedido_item("codigo_tamanho") = dados.getAttribute("codigo_tamanho")
                    RS_Pedido_item("peso") = replace(dados.getAttribute("peso_parcial"),",",".")
                    RS_Pedido_item("nome_produto") = dados.getAttribute("nome_produto")

                    ' Retira virgula e repõe por ponto, para evitar problemas no UPDATE
                    VARvalor_unitario       = replace(dados.getAttribute("preco_unitario"),".","")
                    VARvalor_unitario       = replace(VARvalor_unitario,",",".")

                    RS_Pedido_item("preco_unitario") = VARvalor_unitario
                    RS_Pedido_item("quantidade") = dados.getAttribute("quantidade_produto")
                    Set dados = Nothing

                    'Insere os dados na tabela.
                    RS_Pedido_item.Update
                Next
            
                'Fecha e libera da memória o objeto de Recordset
                RS_Pedido_item.Close
                Set RS_Pedido_item = Nothing

                'Cria sessão de verificação de que os dados foram registrados
                session("resgistroPedidoItem") = "Concluido"
            
            End If

            'Cria sessão de verificação de que os dados foram registrados
            session("registrado") = FctId_transacao & "_ok"
        
        End If

        'Cria sessão codigo_pedido
        Session("codigo_pedido") = VARcodigo_pedido
        
        'Fecha arquivo XML
        Call fecha_xmlpedido(FctId_transacao) 

        'Libera objetos da memória
        Set raiz_dados_produto = Nothing
        Set raiz_dados_pedido = Nothing
        Set raiz = Nothing

    End If

End Function
'########################################################################################################
'--> FIM FUNCTION CarregaGrava_dados_pedido
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Grava_Session
' - Esta FUNCTION captura os dados do formulário de cadastro e do arquivo XML para gravação no banco de dados
' - Esta FUNCTION é chamada no arquivo INICIA_TRANSACAO.ASP
'########################################################################################################
Sub Grava_Session()
        'Grava sessões a partir dos dados postados
        Session("razaosocial_cobranca")        = Request("razaosocial_cobranca")
        Session("cnpj_cobranca")               = Request("cnpj_cobranca")
        Session("inscricaoestadual_cobranca")  = Request("inscricaoestadual_cobranca")
        Session("nome_cobranca")               = Request("nome_cobranca")
        Session("data_nascimento_cobranca")    = Request("data_nascimento_cobranca")
        Session("cpf_cobranca")                = Request("cpf_cobranca")
        Session("rg_cobranca")                 = Request("rg_cobranca")
        Session("logradouro_cobranca")         = Request("logradouro_cobranca")
        Session("numero_cobranca")             = Request("numero_cobranca")
        Session("complemento_cobranca")        = Request("complemento_cobranca")
        Session("bairro_cobranca")             = Request("bairro_cobranca")
        Session("cep_cobranca")                = Request("cep_cobranca")
        Session("cidade_cobranca")             = Request("cidade_cobranca")
        Session("estado_cobranca")             = Request("estado_cobranca")
        Session("pais_cobranca")               = Request("pais_cobranca")
        Session("ddd_cobranca")                = Request("ddd_cobranca")
        Session("telefone_cobranca")           = Request("telefone_cobranca")
        Session("instrucoes")                  = Request("instrucoes")
        Session("razaosocial_entrega")         = Request("razaosocial_entrega")
        Session("cnpj_entrega")                = Request("cnpj_entrega")
        Session("inscricaoestadual_entrega")   = Request("inscricaoestadual_entrega")
        Session("nome_entrega")                = Request("nome_entrega")
        Session("data_nascimento_entrega")     = Request("data_nascimento_entrega")
        Session("cpf_entrega")                 = Request("cpf_entrega")
        Session("rg_entrega")                  = Request("rg_entrega")
        Session("logradouro_entrega")          = Request("logradouro_entrega")
        Session("numero_entrega")              = Request("numero_entrega")
        Session("complemento_entrega")         = Request("complemento_entrega")
        Session("bairro_entrega")              = Request("bairro_entrega")
        Session("cep_entrega")                 = Request("cep_entrega")
        Session("cidade_entrega")              = Request("cidade_entrega")
        Session("estado_entrega")              = Request("estado_entrega")
        Session("pais_entrega")                = Request("pais_entrega")
        Session("ddd_entrega")                 = Request("ddd_entrega")
        Session("telefone_entrega")            = Request("telefone_entrega")
        Session("email_entrega")               = Request("email_entrega")
        If Request("autorizo_newsletter") <> "" Then
        Session("autorizo_newsletter")         = Request("autorizo_newsletter")
        Else
        Session("autorizo_newsletter")         = "0"
        End if

End Sub
'########################################################################################################
'--> FIM FUNCTION Grava_Session
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB abre_ArquivoXML
'   - 
'   - 
'SUB fecha_ArquivoXML
'   - 
'   - 
'########################################################################################################

Sub abre_ArquivoXML(FctArquivo,FctobjXML,FctobjRoot) 
    'Cira objeto para abertura do XML    
    set FctobjXML = CreateObject("Microsoft.XMLDOM")
        FctobjXML.preserveWhiteSpace = False
        FctobjXML.async = False
        FctobjXML.validateOnParse = True
        FctobjXML.resolveExternals = True
        FctobjXML.load (FctArquivo)
    Set FctobjRoot = FctobjXML.documentElement

End Sub

Sub fecha_ArquivoXML(FctArquivo,FctobjXML,FctobjRoot) 
    'Fecha arquivo de XML
    If request("acao") = "alterar" Then
        FctobjXML.save(FctArquivo)
    End If
    'Libera objetos da memória
    set FctobjXML = Nothing
    Set FctobjRoot = Nothing

End Sub
'########################################################################################################
'--> FIM SUB abre_ArquivoXML 
'--> FIM SUB fecha_ArquivoXML 
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION retorna_NomeVisPagamento
'   - Retorna o nome de visualização do meio de pagamento escolhido
'########################################################################################################

Function retorna_NomeVisPagamento(FctMeioPag) 

    'Abre conexao ao XML dos meios de pagto.
    Call abre_ArquivoXML(Application("XMLMeiosPagamentos"),VarobjXML,VarobjRoot)
    Set configuracao = VarobjRoot.selectSingleNode("configuracao/pagto[@nome_pagto='"&FctMeioPag&"']")

    VarNomeVisPag = configuracao.getAttribute("nome_visualizacao")

    'Fecha conexao ao XML dos meios de pagto.
    Call fecha_ArquivoXML(Application("XMLMeiosPagamentos"),VarobjXML,VarobjRoot)

    retorna_NomeVisPagamento = VarNomeVisPag

End Function

'########################################################################################################
'--> FIM FUNCTION retorna_NomeVisPagamento
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION previsualiza_pagto
'   - Monta o formulário de pré-visulização do meio de pagamento
'########################################################################################################

Function previsualiza_pagto() 

    varTotal = pegaValorAtrib(Application("DiretorioPedidos")&session("id_transacao")&".xml","dados_pedido","valor_total")

    'Abre conexao ao XML dos meios de pagto.
    Call abre_ArquivoXML(Application("XMLMeiosPagamentos"),VarobjXML,VarobjRoot)
    
    'Pre-visualizacao do meio de pagto ABNCDC
    Set configuracao = VarobjRootPag.selectSingleNode("configuracao/pagto[@nome_pagto='ABNCDC']")
%>
        <form method="POST" name="abnfinanc" target="vpos" action="<%= Application("URLABNCDCSimulador")%>">
            <input name="VAR01" type="hidden" value="<%= configuracao.getAttribute("VAR01")%>">
            <input name="VAR02" type="hidden" value="02">
            <input name="VAR21" type="hidden" value="<%= configuracao.getAttribute("VAR21")%>">
            <input name="VAR22" type="hidden" value="<%= varTotal%>">
            <input name="VAR27" type="hidden" value="Simulação de financiamento de compra">
        </form>
<%

    'Fecha conexao ao XML dos meios de pagto.
    Call fecha_ArquivoXML(Application("XMLMeiosPagamentos"),VarobjXML,VarobjRoot)

End Function
'########################################################################################################
'--> FIM FUNCTION previsualiza_pagto
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Checa_TodasSessions
'  - A Function abaixo, quando usada, exibe todas as sessions ativas facilitando algum debug necessário
'########################################################################################################
Function Checa_TodasSessions()
    'Loop para listar todas as sessões ativas
    For Each TodasSessions in Session.Contents
        Response.Write TodasSessions & " = " & Session.Contents(TodasSessions) & "<BR>"
    Next

End Function
'########################################################################################################
'--> FIM FUNCTION Checa_TodasSessions
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Checa_TodosRequest
' - A Function abaixo, quando usada, exibe todas os dados enviados para uma página por POST ou GET.
'########################################################################################################
Function Checa_TodosRequest() 
    'Imprimir na tela todos os valores postados por GET e/ou POST
	Response.Write "<b>Form</b><BR>"
	Response.Write Replace(Request.Form(),"&","<BR>")
	Response.Write "<BR><BR>"
	Response.Write "<b>QueryString</b><BR>"
	Response.Write Replace(Request.QueryString(),"&","<BR>")

End Function
'########################################################################################################
'--> FIM FUNCTION Checa_TodosRequest
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Formata_Texto
' - Formata textos substituindo caracteres.
'########################################################################################################
Function Formata_Texto(texto)
'Formatar texto
texto = replace(texto,"'","&rsquo;")
texto = replace(texto,chr(13),"<BR>")

Formata_Texto = texto

End Function
'########################################################################################################
'--> FIM FUNCTION Formata_Texto
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Desformata_Texto
' - Formata textos substituindo caracteres.
'########################################################################################################
Function Desformata_Texto(texto)
'Retornando os caracteres originais
texto = replace(texto,"&rsquo;","'")
texto = replace(texto,"<BR>",chr(13))

Desformata_Texto = texto

End Function
'########################################################################################################
'--> FIM FUNCTION Desformata_Texto
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION pegavalor_promocao
' -
'########################################################################################################
Function pegavalor_promocao(FctValor_produto,RS_Produto)
        'Define a data atual
        dataAtual = date()
        'Se a data registrada for válida    
        If RS_Produto("data_inicio") <> "00:00:00"  Then
        If RS_Produto("data_inicio") <= Date then

            'Tempo decorrido a partir da data cadastrada como data final da promoção.
            TempoCorrente = DateDiff("d", date(),RS_Produto("data_fim")) 
            If TempoCorrente >= 0 Then
                TempoValido = DateDiff("d", RS_Produto("data_inicio"), RS_Produto("data_fim")) 
                If TempoValido >= 0 Then
                    'Se existir algum desconto para o produto
                    If RS_Produto("desconto") <> "0" Then
                        FctValor_produto_promo = FormatNumber(FctValor_produto * (RS_Produto("desconto")/100))
                        FctValor_produto_promo = FctValor_produto - FctValor_produto_promo
                        pegavalor_promocao = FctValor_produto_promo
                    End if
                End if
            End if 
            End if 
        End if   

End Function
'########################################################################################################
'--> FIM FUNCTION pegavalor_promocao
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Cria_Combo_Categoria
' -
'########################################################################################################
Function Cria_Combo_Categoria(VarCodigo_categoria,VarAdicional)

    'Cria objeto de consulta a tabela de Categorias
    Set RS_Categorias = Server.CreateObject("ADODB.Recordset")
    RS_Categorias.CursorLocation = 3
    RS_Categorias.CursorType = 0
    RS_Categorias.LockType = 1

    RS_Categorias.Open "SELECT codigo_categoria, nome_categoria FROM Categorias WHERE sigla_idioma = '"&varLang&"' ORDER BY nome_categoria ", Conexao
    'Se não existir categorias cadastradas
    If RS_Categorias.Eof Then
%>
        <span class="MNlatesquerda"><%= Application("MenuTxtCatVazio")%></span>
<%
    'Existindo categorias, serão listadas.
    Else
%>
        <SELECT NAME="codigo_categoria" class="LCNlatesquerda" <%=Action%> <%=VarAdicional%>>
            <OPTION value=""><%=Application("BoxSelSelecione")%></OPTION>
<%
            While Not RS_Categorias.EOF
            If CDbl(VarCodigo_categoria) = CDbl(RS_Categorias("codigo_categoria")) Then
                SELECTED = "SELECTED"
            Else
                SELECTED = ""
            End if
%>
            <OPTION value="<%= RS_Categorias("codigo_categoria")%>" <%=SELECTED%>><%=RS_Categorias("nome_categoria")%></OPTION>
<%
            RS_Categorias.MoveNext
            Wend
%>	
            <OPTION value="000"><%=Application("BoxSelPesquisarTodas")%></OPTION>

        </SELECT>
<%
    End If

    'Fecha e libera da memória o objeto de Recordset
    RS_Categorias.Close
    Set RS_Categorias = Nothing
End Function
'########################################################################################################
'--> FIM FUNCTION Cria_Combo_Categoria
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Mostra_frete
' - Exibe o box para calculo do frete na página de carrinho
'########################################################################################################
Sub Mostra_frete()
%>
<table width="100%"  border="0" cellspacing="0" cellpadding="10">
    <tr>
        <form name="frm" method="post" action="" Onsubmit="return false;">
        <input type="hidden" name="act" value="inserir">
        <input type="hidden" name="textocep" value="<%= Application("MiddleTxtTitFrete") %>">
        <input type="hidden" name="pesofrete" value="<%= pesofrete %>">
        <td>
            <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
                <% If Session("msgErroFrete") <> "" Then %>
                <tr>                    
                    <td bgcolor="#FFFF00" valign="bottom" align="center" height="20"><%= Replace(Application("MiddleTxtRecalcFrete"),"varUltimaOpcaoFrete",Session("ultima_opcao_frete"))%><br><B><%= Session("msgErroFrete")%></B><br><%= Application("MiddleTxtSelNovoFrete") %>                    
                    </td>
                </tr>
                <% 
                End If 
                
                If Session("cep_frete") <> "" Then
                    atualCEP = Session("cep_frete")
                End If
                
                If atualCEP = "" Or atualCEP = "0" Then
                    atualCEP = Session("cep_entrega")
                End If

                If atualCEP = "" Or atualCEP = "0" Then
                    atualCEP = Session("cep_cobranca")
                End If
                %>
                <tr>                    
                    <td bgcolor="#F5F5F5" valign="bottom"><B><%=Application("MiddleTxtSelPais")%>:</B>&nbsp;&nbsp;<%Call Cria_Combo_Paises(Session("pais_frete"),"paises","onchange=""travaCepFrete(this.options[this.selectedIndex].value,document.getElementById('cep'))""")%>&nbsp;&nbsp;&nbsp;<B><%=Application("MiddleTxtInformeCEP")%>:</B>&nbsp;&nbsp;&nbsp;<input type="text" name="cep" id="cep" size="11" maxlength="8" class="FORMbox" onKeyUp='ajustaCepFrete(document.getElementById("paises").options[document.getElementById("paises").selectedIndex].value,this);' value="<%= atualCEP %>">&nbsp;&nbsp;&nbsp;<input name="calc_frete" class="bttn4" type="button" value="<%=Application("BttCalcultarFrete")%>" <%If pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","Estoque") = "sim" Then%>Onclick="atualiza_carrinho('frete');"<%Else%>Onclick="valida_pesquisar_cep();"<%End If%>></td>
                </tr>
                <tr>                    
                    <td bgcolor="#FFFFFF">
                        <table id="freteTable" border="0" width="100%">
                            <%
                            If Session("valor_frete") > FormatNumber("0") Or Session("frete_gratuito") = "ativo" Then
                                'Converte o valor real do frete para a moeda utilizada na compra
                                varValorFrete = FormatNumber(Session("valor_frete")*FatorCambio(Session("Valor_Cambio")))
                                                                
                                'Trata o texto de info sobre a opção de frete escolhida conforme os dados do frete
                                VarTxtInformacoesFrete = Replace(Application("MiddleTxtInformacoesFrete"),"varOpcaoFreteExib",Session("opcao_frete"))
                                VarTxtInformacoesFrete = Replace(VarTxtInformacoesFrete,"varValorFreteExib",varValorFrete)
                                VarTxtInformacoesFrete = Replace(VarTxtInformacoesFrete,"varCEPFreteExib",Session("cep_frete"))
                                VarTxtInformacoesFrete = Replace(VarTxtInformacoesFrete,"varPesoTotalExib",Session("peso_total"))
                            %>
                            <tr>
                                <td><%=VarTxtInformacoesFrete%></td>
                                <%
                                LogoFrete = Replace(Session("opcao_frete")," ","")
                                LogoFrete = Replace(LogoFrete,"-","")

                                ' Verifica se a imagem da opção de frete escolhida existe
                                If VerificaExistenciaArquivo(Application("DiretorioLoja") & "\config\imagens_conteudo\padrao\"&LogoFrete&"_logo.gif") Then
                                %>
                                <td align="left" align="right" valign="middle"><img src="config/imagens_conteudo/padrao/<%=LogoFrete%>_logo.gif"></td>
                                <%
                                End If
                                %>
                            </tr>
                            <% End If %>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
        </form>
    </tr>
</table>
<%
If Application("disponivelfedex") <> "sim" Then
%>
<script>
    document.frm.paises.disabled = true;
</script>
<%
End If

End Sub
'########################################################################################################
'--> FIM SUB Mostra_frete
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Cria_Combo_Paises
'   - Monsta a relação de paises
'   - Relação utilizada para calculo do frete pelo FEDEX
'########################################################################################################
Function Cria_Combo_Paises(pais_sel,tipo,onchange)
    
    Dim Pais(195), Sigla(195)

    Pais(1)="Albania"
    Pais(2)="Algeria"
    Pais(3)="American Samoa"
    Pais(4)="Andorra"
    Pais(5)="Angola"
    Pais(6)="Anguilla"
    Pais(7)="Antigua"
    Pais(8)="Argentina"
    Pais(9)="Armenia"
    Pais(10)="Aruba"
    Pais(11)="Australia"
    Pais(12)="Austria"
    Pais(13)="Azerbaijan"
    Pais(14)="Bahamas"
    Pais(15)="Bahrain"
    Pais(16)="Bangladesh"
    Pais(17)="Barbados"
    Pais(18)="Belarus"
    Pais(19)="Belgium"
    Pais(20)="Belize"
    Pais(21)="Benin"
    Pais(22)="Bermuda"
    Pais(23)="Bhutan"
    Pais(24)="Bolivia"
    Pais(25)="Botswana"
    Pais(26)="Brasil"
    Pais(27)="British Virgin Islands"
    Pais(28)="Brunei"
    Pais(29)="Bulgaria"
    Pais(30)="Burkina Faso"
    Pais(31)="Burundi"
    Pais(32)="Cambodia"
    Pais(33)="Cameroon"
    Pais(34)="Canada"
    Pais(35)="Cape Verde"
    Pais(36)="Cayman Islands"
    Pais(37)="Chad"
    Pais(38)="Chile"
    Pais(39)="China"
    Pais(40)="Colombia"
    Pais(41)="Congo"
    Pais(42)="Cook Islands"
    Pais(43)="Costa Rica"
    Pais(44)="Croatia"
    Pais(45)="Cyprus"
    Pais(46)="Czech Republic"
    Pais(47)="Denmark"
    Pais(48)="Djiibouti"
    Pais(49)="Dominica"
    Pais(50)="Dominican Republic"
    Pais(51)="Ecuador"
    Pais(52)="Egypt"
    Pais(53)="El Salvador"
    Pais(54)="Equatorial Guinea"
    Pais(55)="Eritrea"
    Pais(56)="Estonia"
    Pais(57)="Ethiopia"
    Pais(58)="Faeroe Islands"
    Pais(59)="Fiji"
    Pais(60)="Finland"
    Pais(61)="France"
    Pais(62)="French Guiana"
    Pais(63)="French Polynesia"
    Pais(64)="Gabon"
    Pais(65)="Gambia"
    Pais(66)="Georgia"
    Pais(67)="Germany"
    Pais(68)="Ghana"
    Pais(69)="Gibraltar"
    Pais(70)="Greece"
    Pais(71)="Greenland"
    Pais(72)="Grenada"
    Pais(73)="Guadeloupe"
    Pais(74)="Guam"
    Pais(75)="Guatemala"
    Pais(76)="Guinea"
    Pais(77)="Guyana"
    Pais(78)="Haiti"
    Pais(79)="Honduras"
    Pais(80)="Hong Kong"
    Pais(81)="Hungary"
    Pais(82)="Iceland"
    Pais(83)="India"
    Pais(84)="Indonesia"
    Pais(85)="Iraq Republic"
    Pais(86)="Ireland"
    Pais(87)="Israel"
    Pais(88)="Italy"
    Pais(89)="Ivory Coast"
    Pais(90)="Jamaica"
    Pais(91)="Japan"
    Pais(92)="Jordan"
    Pais(93)="Kazakhstan"
    Pais(94)="Kenya"
    Pais(95)="Kuwait"
    Pais(96)="Kyrgyzstan"
    Pais(97)="Latvia"
    Pais(98)="Lebanon"
    Pais(99)="Lesotho"
    Pais(100)="Liberia"
    Pais(101)="Liechtenstein"
    Pais(102)="Lithuania"
    Pais(103)="Luxembourg"
    Pais(104)="Macau"
    Pais(105)="Macedonia"
    Pais(106)="Madagascar"
    Pais(107)="Malawi"
    Pais(108)="Malaysia"
    Pais(109)="Maldives"
    Pais(110)="Mali"
    Pais(111)="Malta"
    Pais(112)="Marshall Islands"
    Pais(113)="Martinique"
    Pais(114)="Mauritania"
    Pais(115)="Mauritius"
    Pais(116)="Mexico"
    Pais(117)="Micronesia"
    Pais(118)="Moldova"
    Pais(119)="Monaco"
    Pais(120)="Montserrat"
    Pais(121)="Morocco"
    Pais(122)="Mozambique"
    Pais(123)="Namibia"
    Pais(124)="Nepal"
    Pais(125)="Netherlands Antilles"
    Pais(126)="Netherlands"
    Pais(127)="New Caledonia"
    Pais(128)="New Zealand"
    Pais(129)="Nicaragua"
    Pais(130)="Niger"
    Pais(131)="Nigeria"
    Pais(132)="Norway"
    Pais(133)="Oman"
    Pais(134)="Pakistan"
    Pais(135)="Palau"
    Pais(136)="Panama"
    Pais(137)="Papua New Guinea"
    Pais(138)="Paraguay"
    Pais(139)="Peru"
    Pais(140)="Philippines"
    Pais(141)="Poland"
    Pais(142)="Portugal"
    Pais(143)="Puerto Rico"
    Pais(144)="Qatar"
    Pais(145)="Reunion"
    Pais(146)="Romania"
    Pais(147)="Russian Federation"
    Pais(148)="Rwanda"
    Pais(149)="Saipan"
    Pais(150)="Saudi Arabia"
    Pais(151)="Scotland"
    Pais(152)="Senegal"
    Pais(153)="Seychelles"
    Pais(154)="Singapore"
    Pais(155)="Slovak Republic"
    Pais(156)="Slovenia"
    Pais(157)="South Africa"
    Pais(158)="South Korea"
    Pais(159)="Spain"
    Pais(160)="Sri Lanka"
    Pais(161)="St. Kitts and Nevis"
    Pais(162)="St. Lucia"
    Pais(163)="St. Vincent"
    Pais(164)="Suriname"
    Pais(165)="Swaziland"
    Pais(166)="Sweden"
    Pais(167)="Switzerland"
    Pais(168)="Syria"
    Pais(169)="Taiwan"
    Pais(170)="Tanzania"
    Pais(171)="Thailand"
    Pais(172)="Togo"
    Pais(173)="Trinidad and Tobago"
    Pais(174)="Tunisia"
    Pais(175)="Turkey"
    Pais(176)="Turks and Caicos Islands"
    Pais(177)="U S Virgin Islands"
    Pais(178)="Uganda"
    Pais(179)="Ukraine"
    Pais(180)="United Arab Emirates"
    Pais(181)="United Kingdom"
    Pais(182)="United States"
    Pais(183)="Uruguay"
    Pais(184)="Uzbekistan"
    Pais(185)="Vanuatu"
    Pais(186)="Vatican City"
    Pais(187)="Venezuela"
    Pais(188)="Vietnam"
    Pais(189)="Wales"
    Pais(190)="Wallis &amp; Futuna"
    Pais(191)="Yemen"
    Pais(192)="Servia &amp; Montenegro"
    Pais(193)="Zaire"
    Pais(194)="Zambia"
    Pais(195)="Zimbabwe"

    Sigla(1)="AL"
    Sigla(2)="DZ"
    Sigla(3)="AS"
    Sigla(4)="AD"
    Sigla(5)="AO"
    Sigla(6)="AI"
    Sigla(7)="AG"
    Sigla(8)="AR"
    Sigla(9)="AM"
    Sigla(10)="AW"
    Sigla(11)="AU"
    Sigla(12)="AT"
    Sigla(13)="AZ"
    Sigla(14)="BS"
    Sigla(15)="BH"
    Sigla(16)="BD"
    Sigla(17)="BB"
    Sigla(18)="BY"
    Sigla(19)="BE"
    Sigla(20)="BZ"
    Sigla(21)="BJ"
    Sigla(22)="BM"
    Sigla(23)="BT"
    Sigla(24)="BO"
    Sigla(25)="BW"
    Sigla(26)="BR"
    Sigla(27)="VG"
    Sigla(28)="BN"
    Sigla(29)="BG"
    Sigla(30)="BF"
    Sigla(31)="BI"
    Sigla(32)="KH"
    Sigla(33)="CM"
    Sigla(34)="CA"
    Sigla(35)="CV"
    Sigla(36)="KY"
    Sigla(37)="TD"
    Sigla(38)="CL"
    Sigla(39)="CN"
    Sigla(40)="CO"
    Sigla(41)="CG"
    Sigla(42)="CK"
    Sigla(43)="CR"
    Sigla(44)="HR"
    Sigla(45)="CY"
    Sigla(46)="CZ"
    Sigla(47)="DK"
    Sigla(48)="DJ"
    Sigla(49)="DM"
    Sigla(50)="DO"
    Sigla(51)="EC"
    Sigla(52)="EG"
    Sigla(53)="SV"
    Sigla(54)="GQ"
    Sigla(55)="ER"
    Sigla(56)="EE"
    Sigla(57)="ET"
    Sigla(58)="FO"
    Sigla(59)="FJ"
    Sigla(60)="FI"
    Sigla(61)="FR"
    Sigla(62)="GF"
    Sigla(63)="PF"
    Sigla(64)="GA"
    Sigla(65)="GM"
    Sigla(66)="GE"
    Sigla(67)="DE"
    Sigla(68)="GH"
    Sigla(69)="GI"
    Sigla(70)="GR"
    Sigla(71)="GL"
    Sigla(72)="GD"
    Sigla(73)="GP"
    Sigla(74)="GU"
    Sigla(75)="GT"
    Sigla(76)="GN"
    Sigla(77)="GY"
    Sigla(78)="HT"
    Sigla(79)="HN"
    Sigla(80)="HK"
    Sigla(81)="HU"
    Sigla(82)="IS"
    Sigla(83)="IN"
    Sigla(84)="ID"
    Sigla(85)="IQ"
    Sigla(86)="IE"
    Sigla(87)="IL"
    Sigla(88)="IT"
    Sigla(89)="CI"
    Sigla(90)="JM"
    Sigla(91)="JP"
    Sigla(92)="JO"
    Sigla(93)="KZ"
    Sigla(94)="KE"
    Sigla(95)="KW"
    Sigla(96)="KG"
    Sigla(97)="LV"
    Sigla(98)="LB"
    Sigla(99)="LS"
    Sigla(100)="LR"
    Sigla(101)="LI"
    Sigla(102)="LT"
    Sigla(103)="LU"
    Sigla(104)="MO"
    Sigla(105)="MK"
    Sigla(106)="MG"
    Sigla(107)="MW"
    Sigla(108)="MY"
    Sigla(109)="MV"
    Sigla(110)="ML"
    Sigla(111)="MT"
    Sigla(112)="MH"
    Sigla(113)="MQ"
    Sigla(114)="MR"
    Sigla(115)="MU"
    Sigla(116)="MX"
    Sigla(117)="FM"
    Sigla(118)="MD"
    Sigla(119)="MC"
    Sigla(120)="MS"
    Sigla(121)="MA"
    Sigla(122)="MZ"
    Sigla(123)="NA"
    Sigla(124)="NP"
    Sigla(125)="AN"
    Sigla(126)="NL"
    Sigla(127)="NC"
    Sigla(128)="NZ"
    Sigla(129)="NI"
    Sigla(130)="NE"
    Sigla(131)="NG"
    Sigla(132)="NO"
    Sigla(133)="OM"
    Sigla(134)="PK"
    Sigla(135)="PW"
    Sigla(136)="PA"
    Sigla(137)="PG"
    Sigla(138)="PY"
    Sigla(139)="PE"
    Sigla(140)="PH"
    Sigla(141)="PL"
    Sigla(142)="PT"
    Sigla(143)="PR"
    Sigla(144)="QA"
    Sigla(145)="RE"
    Sigla(146)="RO"
    Sigla(147)="RU"
    Sigla(148)="RW"
    Sigla(149)="MP"
    Sigla(150)="SA"
    Sigla(151)="GB"
    Sigla(152)="SN"
    Sigla(153)="SC"
    Sigla(154)="SG"
    Sigla(155)="SK"
    Sigla(156)="SI"
    Sigla(157)="ZA"
    Sigla(158)="KR"
    Sigla(159)="ES"
    Sigla(160)="LK"
    Sigla(161)="KN"
    Sigla(162)="LC"
    Sigla(163)="VC"
    Sigla(164)="SR"
    Sigla(165)="SZ"
    Sigla(166)="SE"
    Sigla(167)="CH"
    Sigla(168)="SY"
    Sigla(169)="TW"
    Sigla(170)="TZ"
    Sigla(171)="TH"
    Sigla(172)="TG"
    Sigla(173)="TT"
    Sigla(174)="TN"
    Sigla(175)="TR"
    Sigla(176)="TC"
    Sigla(177)="VI"
    Sigla(178)="UG"
    Sigla(179)="UA"
    Sigla(180)="AE"
    Sigla(181)="GB"
    Sigla(182)="US"
    Sigla(183)="UY"
    Sigla(184)="UZ"
    Sigla(185)="VU"
    Sigla(186)="VA"
    Sigla(187)="VE"
    Sigla(188)="VN"
    Sigla(189)="UK"
    Sigla(190)="WF"
    Sigla(191)="YE"
    Sigla(192)="YU"
    Sigla(193)="ZR"
    Sigla(194)="ZM"
    Sigla(195)="ZW"

'Codifica o nome do pais para a respectiva sigla
If tipo = "codifica" Then

    For I=1 to 195
        If pais_sel = Pais(i) Then
            pais_sel = Sigla(i)
            Exit For
        End If
    Next

    Cria_Combo_Paises = pais_sel

'Codifica a sigla do pais para o respectivo nome
ElseIf tipo = "decodifica" Then

    For I=1 to 195
        If pais_sel = Sigla(i) Then
            pais_sel = Pais(i)
            Exit For
        End If
    Next

    Cria_Combo_Paises = pais_sel

' Lista os paises cadastrados
Else

    'Define o pais padrão do dropmenu
    If (pais_sel = "0") Then
        pais_sel = "BR"
    End If

    If (pais_sel = "") Then
        pais_sel = "BR"
    End If

%>
<select size="1" name="<%= tipo%>" id="<%= tipo%>" tabindex="27" class="FORMbox" <%= onchange %>>
<% 
	For I=1 to 195
    If pais_sel = Sigla(i) Then    %>
        <option value="<%= Sigla(i) %>" SELECTED><%= Pais(i) %></option>		
    <% Else %>
        <option value="<%= Sigla(i) %>"><%= Pais(i) %></option>		
    <% End If
    Next 
%>
</select>
<%
End If

End Function
'########################################################################################################
'--> FIM SUB Cria_Combo_Paises
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB pegaValorAtrib
' - Captura um valor especifico de um atributo.
'########################################################################################################

Function pegaValorAtrib(fctArquivo,fctNode,fctAtrib) 
    'Abre arquivo XML
    Call abre_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot)
        If right(fctArquivo,5) = "\.xml" Then
            Response.redirect "carrinho_vazio.asp?refereRecibo=ok"
        Else        
            Set configuracao = FctobjRoot.selectSingleNode(fctNode)
            'Captura valor do atributo desejado.
            pegaValorAtrib = configuracao.getAttribute(fctAtrib)

        End if
    Call fecha_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot) 
End Function

'########################################################################################################
'--> FIM SUB pegaValorAtrib
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB pegaValorNode
' - Captura um valor especifico de um nó.
'########################################################################################################

Function pegaValorNode(fctArquivo,fctAtrib) 
    'Abre arquivo XML
    Call abre_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot)
        Set configuracao = FctobjRoot.selectSingleNode("configuracao/infos[@codigo_texto='"&codigo_texto&"']")
            'Captura valor do nó desejado.
            pegaValorNode = configuracao.text
    Call fecha_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot) 
End Function

'########################################################################################################
'--> FIM SUB pegaValorNode
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_Combo_opcao
' - Monta as opções de SIM e NÃO para os campos de configuração dos meios de pagamentos.
' - Chamada no arquivo ADM_config_pagamento.asp
'########################################################################################################
Sub Cria_Combo_opcao(opcao,valor,onchange)
%>
    <SELECT NAME="<%=opcao%>" class="FORMbox" <%= onchange%>>
	<%
    If (valor = "sim") Then %>
        <OPTION SELECTED VALUE="sim"><%= Application("MiddleTxtSim")%></OPTION>		
        <OPTION VALUE="não"><%= Application("MiddleTxtNao")%></OPTION>		
    <% Else %>
        <OPTION VALUE="sim"><%= Application("MiddleTxtSim")%></OPTION>		
        <OPTION SELECTED VALUE="não"><%= Application("MiddleTxtNao")%></OPTION>		
    <% End If %>
    </SELECT>
<%
End Sub
'########################################################################################################
'--> FIM SUB Cria_Combo_opcao
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB MontaCombo_orderby
' - Monta as opções de ordem da consulta
'########################################################################################################
Sub MontaCombo_orderby(opcao,valorSel,adicional)

    Dim tipoORDERBY(6), valorORDERBY(6)

    tipoORDERBY(1)=Application("MiddleTxtNomeCrescente")
    tipoORDERBY(2)=Application("MiddleTxtNomeDecrescente")
    tipoORDERBY(3)=Application("MiddleTxtPrecoCrescente")
    tipoORDERBY(4)=Application("MiddleTxtPrecoDecrescente")
    tipoORDERBY(5)=Application("MiddleTxtCodigoCrescente")
    tipoORDERBY(6)=Application("MiddleTxtCodigoDecrescente")
    
    valorORDERBY(1)="nome_produto" ' Nome do produto crescente
    valorORDERBY(2)="nome_produto DESC" ' Nome do produto decrescente
    valorORDERBY(3)="preco_unitario" ' Valor do produto crescente
    valorORDERBY(4)="preco_unitario DESC" ' Valor do produto decrescente
    valorORDERBY(5)="codigo_produto" ' Código do produto crescente
    valorORDERBY(6)="codigo_produto DESC" ' Código do produto decrescente

%>
<select size="1" name="<%= opcao%>" tabindex="27" class="FORMbox" Onchange="javascript: document.formOrderBy.submit();">
<% 
	'Define o tipo de combo das opções
    If adicional = "parcial" Then
        tipoCombo = 4
    Else
        tipoCombo = 6
    End if

    For I=1 to tipoCombo
    
    If valorSel = valorORDERBY(i) Then    %>
        <option value="<%= valorORDERBY(i) %>" SELECTED><%= tipoORDERBY(i) %></option>		
    <% Else %>
        <option value="<%= valorORDERBY(i) %>"><%= tipoORDERBY(i) %></option>		
    <% End If
    Next 
%>
</select>
<%

End Sub
'########################################################################################################
'--> FIM SUB MontaCombo_orderby
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB MontaCombo_CATorderby
' - Monta as opções de ordem da consulta para categoria
'########################################################################################################
Sub MontaCombo_CATorderby(opcao,valorSel,adicional)

    Dim tipoORDERBY(4), valorORDERBY(4)

    tipoORDERBY(1)="Nome Crescente"
    tipoORDERBY(2)="Nome Decrescente"
    tipoORDERBY(3)="Código Crescente"
    tipoORDERBY(4)="Código Decrescente"
    
    valorORDERBY(1)="nome_categoria" ' Nome da categoria crescente
    valorORDERBY(2)="nome_categoria DESC" ' Nome da categoria decrescente
    valorORDERBY(3)="codigo_categoria" ' Código da categoria crescente
    valorORDERBY(4)="codigo_categoria DESC" ' Código da categoria decrescente

%>
<select size="1" name="<%= opcao%>" tabindex="27" class="FORMbox" Onchange="javascript: document.formOrderBy.submit();">
<% 
    tipoCombo = 4

    For I=1 to tipoCombo
    
    If valorSel = valorORDERBY(i) Then    %>
        <option value="<%= valorORDERBY(i) %>" SELECTED><%= tipoORDERBY(i) %></option>		
    <% Else %>
        <option value="<%= valorORDERBY(i) %>"><%= tipoORDERBY(i) %></option>		
    <% End If
    Next 
%>
</select>
<%

End Sub
'########################################################################################################
'--> FIM SUB MontaCombo_CATorderby
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB MontaCombo_opcaoNum
' - Monta as opções de SIM e NÃO para os campos de configuração dos meios de pagamentos.
' - Chamada no arquivo ADM_config_pagamento.asp
'########################################################################################################
Sub MontaCombo_opcaoNum(opcao,valor)
%>
    <SELECT NAME="<%=opcao%>" class="FORMbox" >
	<%
    If (valor = "1") Then %>
        <OPTION SELECTED VALUE="1"><%= Application("MiddleTxtSim")%></OPTION>		
        <OPTION VALUE="0"><%= Application("MiddleTxtNao")%></OPTION>		
    <% Else %>
        <OPTION VALUE="1"><%= Application("MiddleTxtSim")%></OPTION>		
        <OPTION SELECTED VALUE="0"><%= Application("MiddleTxtNao")%></OPTION>		
    <% End If %>
    </SELECT>
<%
End Sub
'########################################################################################################
'--> FIM SUB MontaCombo_opcaoNum
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB MontaCombo_opcaoAmb
' - Monta as opções de SIM e NÃO para os campos de configuração dos meios de pagamentos.
' - Chamada no arquivo ADM_config_pagamento.asp
'########################################################################################################
Sub MontaCombo_opcaoAmb(opcao,valor)
%>
    <SELECT NAME="<%=opcao%>" class="FORMbox">
	<%
    If (valor = "TESTE") Then %>
        <OPTION SELECTED VALUE="TESTE">TESTE</OPTION>		
        <OPTION VALUE="PRODUCAO">PRODUÇÃO</OPTION>		
    <% Else %>
        <OPTION VALUE="TESTE">TESTE</OPTION>		
        <OPTION SELECTED VALUE="PRODUCAO">PRODUÇÃO</OPTION>		
    <% End If %>
    </SELECT>
<%
End Sub
'########################################################################################################
'--> FIM SUB MontaCombo_opcaoAmb
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB MontaCombo_opcaoTipo
' - Monta as opções de CP e CBR para os campos de configuração dos meios de pagamentos.
' - Chamada no arquivo ADM_funcoes_pagamentos.asp
'########################################################################################################
Sub MontaCombo_opcaoTipo(opcao,valor)
%>
    <select name="<%=opcao%>" class="formbox">
	<%
    if (valor = "CP") then %>
        <option selected value="CP">CP</option>		
        <option value="CBR">CBR</option>		
    <% else %>
        <option value="CBR">CBR</option>		
        <option selected value="CP">CP</option>		
    <% end if %>
    </select>
<%
End Sub
'########################################################################################################
'--> FIM SUB MontaCombo_opcaoAmb
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Lista_Cores
' - Lista as cores disponiveis para cada produto
' - Função chamada na página de descrição do produto - produtos_descricao.asp
'########################################################################################################
Function Lista_Cores(cores)
    'Cria objeto de consulta a tabela de cores
    Set RS_Cores = Server.CreateObject("ADODB.Recordset")
    RS_Cores.CursorLocation = 3
    RS_Cores.CursorType = 0
    RS_Cores.LockType = 3
    RS_Cores.Open "SELECT codigo_cor, nome_cor, url_imagem FROM Cores ORDER BY nome_cor", Conexao
    Coluna = 0
    If Not RS_Cores.Eof Then

            While Not RS_Cores.EOF

			If cores <> "" Then 
                Vetor = Split(cores, ",") 
                For I = 0 To Ubound(Vetor) 
                    If CDbl(Vetor(I)) = CDbl(RS_Cores("codigo_cor")) Then
%>
                        <input type="radio" name="nome_cor" value="<%=RS_Cores("codigo_cor")%>" <%If I = 0 Then response.write "checked" End If%>><%If RS_Cores("url_imagem") <> "" Then%><img src="<%=RS_Cores("url_imagem")%>" alt="<%=RS_Cores("nome_cor")%>" border="1" bordercolor="#330000"><%Else%><%=RS_Cores("nome_cor")%><%End If%><img src="config/templates/<%=varLang%>/<%=varSkin%>/regua1x1.gif" height="3" width="5">
<%
                        Coluna=Coluna+1
                    End If
                Next 
			End If 	

            'Monta o numero de cores por linha
            If Coluna >=3 Then
            Coluna=0
%>
            <br>
<%          End If
            RS_Cores.MoveNext
            Wend

    End If
    'Fecha e libera da memória o objeto de Recordset
    RS_Cores.Close
    Set RS_Cores = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION Lista_Cores
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Lista_Tamanhos
' - Lista as tamanhos disponiveis para cada produto
' - Função chamada na página de descrição do produto - produtos_descricao.asp
'########################################################################################################
Function Lista_Tamanhos(tamanhos)
    'Cria objeto de consulta a tabela de tamanhos
    Set RS_Tamanhos = Server.CreateObject("ADODB.Recordset")
    RS_Tamanhos.CursorLocation = 3
    RS_Tamanhos.CursorType = 0
    RS_Tamanhos.LockType = 3
    RS_Tamanhos.Open "SELECT codigo_tamanho, nome_tamanho FROM Tamanhos ORDER BY nome_tamanho", Conexao
    Coluna = 0	

    If RS_Tamanhos.Eof Then

    Else
            While Not RS_Tamanhos.EOF

            If tamanhos <> "" Then 
                Vetor = Split(tamanhos, ",") 
                For I = 0 To Ubound(Vetor) 
                    If CDbl(Vetor(I)) = CDbl(RS_Tamanhos("codigo_tamanho")) Then
%>
                        <input type="radio" name="nome_tamanho" value="<%=RS_Tamanhos("codigo_tamanho")%>" <%If I = 0 Then response.write "checked" End If%>><%=RS_Tamanhos("nome_tamanho")%><img src="config/templates/<%=varLang%>/<%=varSkin%>/regua1x1.gif" height="3" width="5">
<%
                        Coluna=Coluna+1
                    End If
                Next 
			End If 	

            'Monta o numero de produtos por linha
            If Coluna >=4 Then
            Coluna=0
%>
            <br>
<%          End If
            RS_Tamanhos.MoveNext
            Wend

    End If
   'Fecha e libera da memória o objeto de Recordset
    RS_Tamanhos.Close
    Set RS_Tamanhos = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION Lista_Tamanhos
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Pega_Cor
' - Capturar uma cor especifica mediante código passado
'########################################################################################################
Function Pega_Cor(fctCodigo_cor)

    'Cria objeto de consulta a tabela de Cores
    Set RS_Cor = Server.CreateObject("ADODB.Recordset")
    RS_Cor.CursorLocation = 3
    RS_Cor.CursorType = 0
    RS_Cor.LockType = 3

    If fctCodigo_cor <> "" Then
        RS_Cor.Open "SELECT url_imagem, nome_cor FROM Cores WHERE codigo_cor = "&fctCodigo_cor&"", Conexao
    Else
        RS_Cor.Open "SELECT url_imagem, nome_cor FROM Cores", Conexao
    End If
    
    If Not RS_Cor.Eof Then
            If RS_Cor("url_imagem") <> "" Then
                Pega_Cor = "<img src="&RS_Cor("url_imagem")&" border='1' bordercolor='#000000'>"
            Else
                Pega_Cor = RS_Cor("nome_cor")
            End if

        End If
    'Fecha e libera da memória o objeto de Recordset
    RS_Cor.Close
    Set RS_Cor = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION Pega_Cor
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Pega_Tamanho
' - Captura um tamanho especifica mediante código passado
'########################################################################################################
Function Pega_Tamanho(fctCodigo_tam)

    'Cria objeto de consulta a tabela de Tamahos
    Set RS_Tamanho = Server.CreateObject("ADODB.Recordset")
    RS_Tamanho.CursorLocation = 3
    RS_Tamanho.CursorType = 0
    RS_Tamanho.LockType = 3

    If fctCodigo_tam <> "" Then
        RS_Tamanho.Open "SELECT nome_tamanho FROM Tamanhos WHERE codigo_tamanho = "&fctCodigo_tam&"", Conexao
    Else
        RS_Tamanho.Open "SELECT nome_tamanho FROM Tamanhos", Conexao
    End If
    
        If Not RS_Tamanho.Eof Then
            Pega_Tamanho = RS_Tamanho("nome_tamanho")
        Else
            Pega_Tamanho = Empty
        End If
    'Fecha e libera da memória o objeto de Recordset
    RS_Tamanho.Close
    Set RS_Tamanho = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION Pega_Tamanho
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION VerificaExistenciaDado
' - Verifica se uma informação específica se encontra registrada no banco de dados
'########################################################################################################

Function VerificaExistenciaDado(varCampo,varTabela,varCampoCondicao,varCondicao) 

    'Cria objeto de consulta a tabela definida
    Set RS_Verifica = Server.CreateObject("ADODB.Recordset")
    RS_Verifica.CursorLocation = 3
    RS_Verifica.CursorType = 0
    RS_Verifica.LockType = 3
    
    'Monta a query conforme os parametros passados
    If varCondicao <> "" Then
        sql_Verifica = "SELECT "&varCampo&" FROM "&varTabela&" WHERE "&varCampoCondicao&" = " &varCondicao& ""
    Else
        sql_Verifica = "SELECT "&varCampo&" FROM "&varTabela&""
    End If
     
    'Executa a query    
    RS_Verifica.Open sql_Verifica, Conexao
	
    ' Verica se consulta retornou algum registro	
    If Not RS_Verifica.EOF then 
        VerificaExistenciaDado = True ' Existe
    Else
        VerificaExistenciaDado = False ' Não Existe
    End If

    'Fecha e Libera recordset da memória
    RS_Verifica.Close				
    Set RS_Verifica = Nothing

End Function

'########################################################################################################
'--> FIM FUNCTION VerificaExistenciaDado
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Sigla_Idioma
' - Captura a sigla de idioma configurado
' - Caso seja necessário a configuração de um novo idioma, inserí-lo na função abaixo após o cadastro na loja.
'########################################################################################################

Function Sigla_IdiomaPais(ilvarLang)

    Dim varIdiomaPais(4), SiglaIdiomaPais(4)

    varIdiomaPais(1)="pt_BR"
    varIdiomaPais(2)="en_US"
    varIdiomaPais(3)="es_ES"
    varIdiomaPais(4)="en_UK"


    SiglaIdiomaPais(1)="Português"
    SiglaIdiomaPais(2)="Inglês Americano"
    SiglaIdiomaPais(3)="Espanhol"
    SiglaIdiomaPais(4)="Inglês Britânico"

	For I=1 to 4
        If ilvarLang = varIdiomaPais(i) Then 
            Sigla_IdiomaPais = SiglaIdiomaPais(i)
        End If
    Next 

    ilvarLang = ""


End Function

'########################################################################################################
'--> FIM FUNCTION Sigla_Idioma
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION VerificaExistenciaArquivo
' - Verifica a existência de arquivo/pasta específico mediante parametros postados
'########################################################################################################


'Verifica a existência de arquivo/pasta
Function VerificaExistenciaArquivo(varArquivo) 
    'Cria objeto    
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

        If objFSO.FileExists(varArquivo) Then
            VerificaExistenciaArquivo = True
        Else
            VerificaExistenciaArquivo = False
        End If
    'Libera objeto da memória
    Set objFSO = Nothing
    
End Function

'########################################################################################################
'--> FIM FUNCTION VerificaExistenciaArquivo
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION FatorCambio
'-Transforma em reais o calor do cambio.
'########################################################################################################

'Para não termos problemas com valores dizimais o calculo do cambio sempre será feito a partir de 1,00 e arredondado a 2 digitos de casa decinais.
Function FatorCambio(FctValor_Cambio)
    
    'O valor em Reais será atribuido ao cambio, caso variável vazia ou com valor igual a ZERO.
    If FctValor_Cambio = "" Or FctValor_Cambio = "0" Then
        FctValor_Cambio = 1
    End If
    
    FatorCambioTemp = 1 / FctValor_Cambio

    FatorCambio = FatorCambioTemp
End Function

'########################################################################################################
'--> FIM FUNCTION FatorCambio
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION IdentificaItemNegativado
' - Função que altera o nome do campo de dados em nome legivel
'########################################################################################################

Function IdentificaItemNegativado(varCampo)

    Dim varNomeCampo(16), varItem(16)

    varNomeCampo(1)="user_ID"
    varNomeCampo(2)="ip_cliente"
    varNomeCampo(3)="razaosocial_cobranca"
    varNomeCampo(4)="cnpj_cobranca"
    varNomeCampo(5)="inscricaoestadual_cobranca"
    varNomeCampo(6)="nome_cobranca"
    varNomeCampo(7)="cpf_cobranca"
    varNomeCampo(8)="rg_cobranca"
    varNomeCampo(9)="telefone_cobranca"
    varNomeCampo(10)="razaosocial_entrega"
    varNomeCampo(11)="cnpj_entrega"
    varNomeCampo(12)="inscricaoestadual_entrega"
    varNomeCampo(13)="rg_entrega"
    varNomeCampo(14)="cpf_entrega"
    varNomeCampo(15)="telefone_entrega"
    varNomeCampo(16)="email_entrega"
    


    varItem(1)="E-mail do Usuário"
    varItem(2)="IP do Cliente"
    varItem(3)="Razão Social - Cobrança"
    varItem(4)="CNPJ - Cobrança"
    varItem(5)="Inscrição Estadual - Cobrança"
    varItem(6)="Nome - Cobranca"
    varItem(7)="CPF - Cobranca"
    varItem(8)="RG - Cobranca"
    varItem(9)="Telefone - Entrega"
    varItem(10)="Razão Social - Entrega"
    varItem(11)="CNPJ - Entrega"
    varItem(12)="Inscrição Estadual - Cobrança"
    varItem(13)="CPF - Entrega"
    varItem(14)="RG - Entrega"
    varItem(15)="Telefone - Entrega"
    varItem(16)="E-mail - Entrega"


	For I=1 to 16
        If varCampo = varNomeCampo(i) Then 
            IdentificaItemNegativado = varItem(i)
        End If
    Next 

End Function

'########################################################################################################
'--> FIM FUNCTION IdentificaItemNegativado
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Pega_DadoBanco
'- Captura um registro no banco de dados mediante parametros enviados.
'########################################################################################################

Function Pega_DadoBanco(fctTabela,fctCampoConsultado,fctCampoCondicao,fctCondicao)
    
    Set RS_Pega_Dado = Server.CreateObject("ADODB.Recordset")
    RS_Pega_Dado.CursorLocation = 3
    RS_Pega_Dado.CursorType = 0
    RS_Pega_Dado.LockType = 3

    'Monta query conforme parametros postados.
    Query_Pega_DadoBanco = "SELECT "&fctCampoConsultado&" FROM "&fctTabela&" WHERE "&fctCampoCondicao&" = "&fctCondicao&""
    'Executa query de consulta.
    'response.write Query_Pega_DadoBanco
    RS_Pega_Dado.Open Query_Pega_DadoBanco, Conexao
        If Not RS_Pega_Dado.Eof Then
            Pega_DadoBanco = RS_Pega_Dado(fctCampoConsultado)
        Else
            Pega_DadoBanco = ""
        End If
    
    'Fecha e Libera recordset da memória
    RS_Pega_Dado.Close
    Set RS_Pega_Dado = Nothing

End Function

'########################################################################################################
'--> FIM FUNCTION Pega_DadoBanco
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Atualiza_Estoque
'- Captura um registro no banco de dados mediante parametros enviados.
'########################################################################################################

Function Atualiza_Estoque(FctCodigo_produto,FctAcao,FctQte_Atual,FctQte_Nova)

    If pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","Estoque") = "sim" Then

        If Instr(FctCodigo_produto,"_") <> 0 Then
            Tempcodigo_produto1 = Split(FctCodigo_produto,"_")
            Tempcodigo_produto = Tempcodigo_produto1(0)
        Else
            Tempcodigo_produto = FctCodigo_produto
        End if


        'Monta a string certa para consulta conforme formatação do código do pedido
        'O código do produto pode ser númerico ou alfanumérico
        'A formatação depende da existência de cor e/ou tamanho
        If Instr(FctCodigo_produto,"_") <> 0 Then
            varStrPegaQte = "dados_pedido/produto[@codigo_produto='"&FctCodigo_produto&"']"
        Else
            varStrPegaQte = "dados_pedido/produto[@codigo_produto="&FctCodigo_produto&"]"
        End If

        Set RS_EstoqueProduto = CreateObject("ADODB.Recordset")
        Set RS_EstoqueProduto.ActiveConnection = Conexao
        RS_EstoqueProduto.CursorLocation = 3
        RS_EstoqueProduto.CursorType = 0
        RS_EstoqueProduto.LockType =  3

        RS_EstoqueProduto.Open "SELECT codigo_produto, quantidade_produto FROM Produtos WHERE codigo_produto=" & Tempcodigo_produto &"", Conexao

        If Not RS_EstoqueProduto.EOF Then

            If CDbl(Tempcodigo_produto) <> CDbl(Session("codigo_produtoT")) Then

                If FctAcao = "delete" Then
                    'Soma a quantidade de produtos disponível com a quantidade
                    Resultado_qteTemp = CDbl(RS_EstoqueProduto("quantidade_produto")) + CDbl(FctQte_Atual)
                    Resultado_qte = Resultado_qteTemp - CDbl(FctQte_Nova)
                ElseIf FctAcao = "update" Then
                    Resultado_qteTemp = CDbl(RS_EstoqueProduto("quantidade_produto")) + CDbl(FctQte_Atual)
                    Resultado_qte = Resultado_qteTemp - CDbl(FctQte_Nova)
                ElseIf FctAcao = "novo" Then
                    Resultado_qte = CDbl(RS_EstoqueProduto("quantidade_produto")) - CDbl(FctQte_Atual)
                End if    

                Conexao.Execute "UPDATE Produtos SET quantidade_produto="&Resultado_qte&" WHERE codigo_produto=" & Tempcodigo_produto &""

                Resultado_qte = ""
                Resultado_qteTemp = "" 

                Tempcodigo_produto = Session("codigo_produtoT")

            End If

        End If
        
        Set RS_EstoqueProduto = Nothing


    End If

End Function
'########################################################################################################
'--> FIM FUNCTION Atualiza_Estoque
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Encriptor
'- Encripta ou desencripta um dado.
'########################################################################################################

Function Encriptor(FctDadoEncr,FctAcao)

    Set oEncryptor = Server.Createobject("Dynu.Encrypt") 

        If FctAcao = "encriptar" then
            'Decriptando o valor da string:
            Encriptor = oEncryptor.Encrypt(FctDadoEncr, pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","ChaveCripto")) 
        ElseIf FctAcao = "decriptar" then
            'Decriptando o valor da string:
            Encriptor = oEncryptor.Decrypt(FctDadoEncr, pegaValorAtrib(Application("XMLArquivoConfiguracao"),"dados/configuracao_dados","ChaveCripto")) 
        End If

    Set oEncryptor = Nothing

End Function

'########################################################################################################
'--> FIM FUNCTION Encriptor
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION PreparaData
' - Função para tratamento de datas
'########################################################################################################
Function PreparaData(data) 
    if Day(data) <= 9 AND len(Day(data)) <=2 then 
        dia = "0" & Day(data) 
    else 
        dia = Day(data) 
    end if 
    if month(data) <= 9 AND len(Month(data)) <=2 then 
        mes = "0" & month(data) 
    else 
        mes = month(data) 
    end if 
    if Year(data) <= 9 AND len(Year(data)) <=2 then 
        ano = Left(Year(Now),2) & Year(data) 
    else 
        ano = Year(data) 
    end if 
        PreparaData = dia & "/" & mes & "/" & ano
End Function
'########################################################################################################
'--> FIM FUNCTION PreparaData
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION PreparaDataBD
' - Função para tratamento de datas para Querys em banco de dados
'########################################################################################################
Function PreparaDataBD(data) 
    if Day(data) <= 9 AND len(Day(data)) <=2 then 
        dia = "0" & Day(data) 
    else 
        dia = Day(data) 
    end if 
    if month(data) < 9 AND len(Month(data)) <=2 then 
        mes = "0" & month(data) 
    else 
        mes = month(data) 
    end if 
    if Year(data) <= 9 AND len(Year(data)) <=2 then 
        ano = Left(Year(Now),2) & Year(data) 
    else 
        ano = Year(data) 
    end if 
        PreparaDataBD = ano & "-" & mes & "-" & dia
End function
'########################################################################################################
'--> FIM FUNCTION PreparaDataBD
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION CorrigeData
' - Função para correção de datas de formato MM/DD/AAAA para DD/MM/AAAA
'########################################################################################################
Function CorrigeData(data) 
    
    if Day(data) <= 9 AND len(Day(data)) <=2 then 
        mes = "0" & Day(data) 
    else 
        mes = Day(data) 
    end if 
    
    if month(data) < 9 AND len(Month(data)) <=2 then 
        dia = "0" & month(data) 
    else 
        dia = month(data) 
    end if 
    
    if Year(data) <= 9 AND len(Year(data)) <=2 then 
        ano = Left(Year(Now),2) & Year(data) 
    else 
        ano = Year(data) 
    end if 

    CorrigeData = dia & "/" & mes & "/" & ano

End Function
'########################################################################################################
'--> FIM FUNCTION CorrigeData
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION CorrigeDataHora
' - Função para correção de datas de formato MM/DD/AAAA para DD/MM/AAAA
'########################################################################################################
Function CorrigeDataHora(datahora) 
   
    if Day(datahora) <= 9 AND len(Day(datahora)) <=2 then 
        mes = "0" & Day(datahora) 
    else 
        mes = Day(datahora) 
    end if 
    
    if month(datahora) < 9 AND len(Month(datahora)) <=2 then 
        dia = "0" & month(datahora) 
    else 
        dia = month(datahora) 
    end if 
    
    if Year(datahora) <= 9 AND len(Year(datahora)) <=2 then 
        ano = Left(Year(Now),2) & Year(datahora) 
    else 
        ano = Year(datahora) 
    end if

        if Hour(datahora) < 9 AND len(Hour(datahora)) <=2 then 
        hora = "0" & Hour(datahora) 
    else 
        hora = Hour(datahora) 
    end if

    if Minute(datahora) < 9 AND len(Minute(datahora)) <=2 then 
        minuto = "0" & Minute(datahora) 
    else 
        minuto = Minute(datahora) 
    end if

    if Second(datahora) < 9 AND len(Second(datahora)) <=2 then 
        segundo = "0" & Second(datahora) 
    else 
        segundo = Second(datahora) 
    end if

    CorrigeDataHora = dia & "/" & mes & "/" & ano & " " & hora & ":" & minuto & ":" & segundo

End Function
'########################################################################################################
'--> FIM FUNCTION CorrigeDataHora
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB PreparaDado
' - Função para tratamento de dados
'########################################################################################################
Sub PreparaDado(Fctvalor,FctTipo,Ret1,Ret2,Ret3,Ret4,Ret5)

    If FctTipo = "CPF" Then
        Ret1 = mid(Fctvalor,1,3)
        Ret2 = mid(Fctvalor,4,3)
        Ret3 = mid(Fctvalor,7,3)
        Ret4 = mid(Fctvalor,10,2)
    ElseIf FctTipo = "RG" Then
        Ret1 = mid(Fctvalor,1,3)
        Ret2 = mid(Fctvalor,4,3)
        Ret3 = mid(Fctvalor,7,3)
        Ret4 = mid(Fctvalor,10,1)
    ElseIf FctTipo = "CNPJ" Then
        Ret1 = mid(Fctvalor,1,2)
        Ret2 = mid(Fctvalor,3,3)
        Ret3 = mid(Fctvalor,6,3)
        Ret4 = mid(Fctvalor,9,4)
        Ret5 = mid(Fctvalor,13,2)
    ElseIf FctTipo = "DATA_NASCIMENTO" Then
        Ret1 = mid(Fctvalor,1,2)
        Ret2 = mid(Fctvalor,3,2)
        Ret3 = mid(Fctvalor,5,4)
    ElseIf FctTipo = "CEP" Then
        Ret1 = mid(Fctvalor,1,5)
        Ret2 = mid(Fctvalor,6,3)
    End If

End Sub
'########################################################################################################
'--> FIM SUB PreparaDado
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION FormataDado
' - Função para formatação de dados
'########################################################################################################
Function FormataDado(Fctvalor,FctTipo)

    If FctTipo = "CPF" Then
        Ret1 = mid(Fctvalor,1,3)
        Ret2 = mid(Fctvalor,4,3)
        Ret3 = mid(Fctvalor,7,3)
        Ret4 = mid(Fctvalor,10,2)
        FormataDado = Ret1 & "." & Ret2 & "." & Ret3 & "-" &  Ret4
    ElseIf FctTipo = "RG" Then
        Ret1 = mid(Fctvalor,1,3)
        Ret2 = mid(Fctvalor,4,3)
        Ret3 = mid(Fctvalor,7,3)
        Ret4 = mid(Fctvalor,10,1)
        FormataDado = Ret1 & "." & Ret2 & "." & Ret3 & "-" &  Ret4
    ElseIf FctTipo = "CNPJ" Then
        Ret1 = mid(Fctvalor,1,2)
        Ret2 = mid(Fctvalor,3,3)
        Ret3 = mid(Fctvalor,6,3)
        Ret4 = mid(Fctvalor,9,4)
        Ret5 = mid(Fctvalor,13,2)
        FormataDado = Ret1 & "." & Ret2 & "." & Ret3 & "/" &  Ret4 & "-" & Ret5
    ElseIf FctTipo = "DATA_NASCIMENTO" Then
        Ret1 = mid(Fctvalor,1,2)
        Ret2 = mid(Fctvalor,3,2)
        Ret3 = mid(Fctvalor,5,4)
        FormataDado = Ret1 & "/" & Ret2 & "/" & Ret3
    ElseIf FctTipo = "CEP" Then
        Ret1 = mid(Fctvalor,1,5)
        Ret2 = mid(Fctvalor,6,3)
        FormataDado = Ret1 & "-" & Ret2
    End If

End Function
'########################################################################################################
'--> FIM FUNCTION FormataDado
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Decodifica_Perfil
' -
'########################################################################################################

Function Decodifica_Perfil(perfil)

    If Instr(perfil,",") <> 0 Then
        varPerfil = split(perfil,",")

        for i=0 to ubound(varPerfil)
            Select Case varPerfil(i)
            Case "ADMPed"
               resultado = resultado & "Administrador Pedidos, "
            Case "ADMProd"
               resultado = resultado & "Administrador Produtos, Categorias e Cambio, "
            Case "ADMMeioPagto"
               resultado = resultado & "Administrador Meios de Pagamento, "
            Case "ADMRelat"
               resultado = resultado & "Administrador Relatorios, "
            Case "ADMText"
               resultado = resultado & "Administrador Textos, "
            Case "ADMTarifas"
               resultado = resultado & "Administrador Tarifas, "
            Case "ADMMailing"
               resultado = resultado & "Administrador Mailing, "
            End Select
        next
        resultado = mid(resultado,1,len(resultado)-2)
    Else
        varPerfil = perfil
        Select Case varPerfil
        Case "ADMPed"
           resultado = "Administrador Pedidos"
        Case "ADMProd"
           resultado = "Administrador Produtos, Categorias e Cambio"
        Case "ADMMeioPagto"
           resultado = "Administrador Meios de Pagamento"
        Case "ADMRelat"
           resultado = "Administrador Relatorios"
        Case "ADMText"
           resultado = "Administrador Textos"
        Case "ADMTarifas"
           resultado = "Administrador Tarifas"
        Case "ADMMailing"
           resultado = "Administrador Mailing"
        Case "ADMGeral"
           resultado = "Administrador"
        End Select                        
    End If

    Decodifica_Perfil = resultado

End Function
'########################################################################################################
'--> FIM FUNCTION Decodifica_Perfil
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION FormataData
' - Função para formatar uma data especifica para DDMMAAAA.
'########################################################################################################
Function FormataData(data)

    ' converte vencimento para formato ddmmaaaa
    dia = DatePart("d", data)
    If (dia < 10) Then
        dia = "0" & dia
    End If
    mes = DatePart("m", data)
    If (mes < 10) Then
        mes = "0" & mes
    End If
    ano = DatePart("yyyy", data)

    FormataData = dia & mes & ano

End Function
'########################################################################################################
'--> FIM FUNCTION FormataData
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB Cria_Combo_Numeros
' - Monta as opções numericas do número inicial até o número máximo definido
' - 
'########################################################################################################
Sub Cria_Combo_Numeros(opcao,valor,minopcao,maxopcao,onchange)

    ' Verifica se o minimo está nulo
    If minopcao = "" Then
        ' Define o minimo como zero
        minopcao = 0
    End If
%>
    <SELECT NAME="<%=opcao%>" class="FORMbox" <%= onchange%>>
	<%
    For N=minopcao To Int(maxopcao)
        If N = Int(valor) Then
    %>
        <option value="<%= N %>" SELECTED><%= N %></option>		
    <%
        Else
    %>
        <option value="<%= N %>"><%= N %></option>		
    <%
        End If
    Next 
    %>
    </SELECT>
<%
End Sub
'########################################################################################################
'--> FIM SUB Cria_Combo_Numeros
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB calculaValorTaxa
' - Calcula o valor definido com um acréscimo/desconto da taxa definida em porcentagem
'########################################################################################################

Function calculaValorTaxa(fctValor,fctTaxa,fctAcao)

    ' Verifica o tipo de ação definido
    If fctAcao = "Desconto" Then ' Desconto
        currValorDesc = fctValor * (Replace(fctTaxa,".",",")/100)
        calculaValorTaxa = FormatNumber(fctValor - currValorDesc)
    ElseIf fctAcao = "Acrescimo" Then ' Acréscimo
        currValorJuros = fctValor * (Replace(fctTaxa,".",",")/100)
        calculaValorTaxa = FormatNumber(fctValor + currValorJuros)
    End If

End Function

'########################################################################################################
'--> FIM SUB calculaValorTaxa
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB alteraValorAtrib
' - Altera um valor especifico de um atributo.
'########################################################################################################

Function alteraValorAtrib(fctArquivo,fctNode,fctAtrib,fctNovoValor) 
    'Abre arquivo XML
    Call abre_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot)
        If right(fctArquivo,5) = "\.xml" Then
            Response.redirect "carrinho_vazio.asp?refereRecibo=ok"
        Else        
            Set configuracao = FctobjRoot.selectSingleNode(fctNode)
            'Altera valor do atributo desejado.
            configuracao.setAttribute fctAtrib,""&Trim(fctNovoValor)&""
            FctobjXML.save(fctArquivo)
            Set configuracao = Nothing
        End if
    Call fecha_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot) 
End Function

'########################################################################################################
'--> FIM SUB alteraValorAtrib
'########################################################################################################
'========================================================================================================
'########################################################################################################
'SUB alterValorNode
' - Altera um valor especifico de um nó.
'########################################################################################################

Function alterValorNode(fctArquivo,fctNode,fctNovoValor) 
    'Abre arquivo XML
    Call abre_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot)
        Set configuracao = FctobjRoot.selectSingleNode(fctNode)
            'Altera valor do nó desejado.
            configuracao.text = fctNovoValor
        FctobjXML.save(fctArquivo)
        Set configuracao = Nothing
    Call fecha_ArquivoXML(fctArquivo,FctobjXML,FctobjRoot) 
End Function

'########################################################################################################
'--> FIM SUB alterValorNode
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION BinaryToString
' - Tratamento dos acentos
'########################################################################################################
Function BinaryToString(strBinary) 
	Dim intCount, xBinaryToString 
	xBinaryToString ="" 
		For intCount = 1 to LenB(strBinary) 
			xBinaryToString = xBinaryToString & chr(AscB(MidB(strBinary,intCount,1))) 
		Next 
	BinaryToString = xBinaryToString 
End Function 
'########################################################################################################
'--> FIM FUNCTION BinaryToString
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION ChecaValorArray
' - Verifica se um determinado valor consta em um determinado array. Retorno: True ou False
'########################################################################################################
Function ChecaValorArray(strArray,strValor)

resultChec = False

vetorArray = Split(strArray, ",")

For nValorArray = 0 To Ubound(vetorArray)

    If vetorArray(nValorArray) = strValor Then
        resultChec = True
        Exit For
    End If

Next

ChecaValorArray = resultChec

End Function 
'########################################################################################################
'--> FIM FUNCTION ChecaValorArray
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Anula_TodasSessions
'  - A Function abaixo, quando usada, anula todas as sessions ativas exceto a "codigo_pedido" e "id_transacao"
'########################################################################################################
Function Anula_TodasSessions()
    'Loop para listar todas as sessões ativas
    For Each TodasSessions in Session.Contents
        If TodasSessions <> "codigo_pedido" And TodasSessions <> "id_transacao" And TodasSessions <> "forma_pagamento" Then
            Session.Contents(TodasSessions) = EMPTY
        End if
    Next

End Function
'########################################################################################################
'--> FIM FUNCTION Anula_TodasSessions
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION Dados_Transacao
' - 
'########################################################################################################
Function Dados_Transacao(nCodigoPedido,sIdentificacao,sModulo,sOperacao,sAmbiente,sDescPedido,sParamAdicionais)

Set RS_Dados = Server.CreateObject("ADODB.Recordset")
RS_Dados.CursorLocation = 3
RS_Dados.CursorType = 0
RS_Dados.LockType =  1

RS_Dados.Open "SELECT Pedido_item.codigo_produto, Pedido_item.codigo_categoria, Pedido_item.codigo_cor, Pedido_item.codigo_tamanho, Pedido_item.quantidade, Pedido_item.preco_unitario, Pedido_item.peso, Pedido_item.nome_produto, Pedido_item.sigla_moeda, Pedido_item.valor_moeda, Pedidos.codigo_pedido, Pedidos.data_pedido_inicio, Pedidos.data_pedido, Pedidos.user_ID, Pedidos.sessionID, Pedidos.ip_cliente, Pedidos.razaosocial_cobranca, Pedidos.cnpj_cobranca, Pedidos.inscricaoestadual_cobranca, Pedidos.nome_cobranca, Pedidos.cpf_cobranca, Pedidos.rg_cobranca, Pedidos.data_nascimento_cobranca, Pedidos.logradouro_cobranca, Pedidos.numero_cobranca, Pedidos.complemento_cobranca, Pedidos.bairro_cobranca, Pedidos.cidade_cobranca, Pedidos.estado_cobranca, Pedidos.cep_cobranca, Pedidos.pais_cobranca, Pedidos.ddd_cobranca, Pedidos.telefone_cobranca, Pedidos.razaosocial_entrega, Pedidos.cnpj_entrega, Pedidos.inscricaoestadual_entrega, Pedidos.nome_entrega, Pedidos.cpf_entrega, Pedidos.rg_entrega, Pedidos.data_nascimento_entrega, Pedidos.logradouro_entrega, Pedidos.numero_entrega, Pedidos.complemento_entrega, Pedidos.bairro_entrega, Pedidos.cidade_entrega, Pedidos.estado_entrega, Pedidos.cep_entrega, Pedidos.pais_entrega, Pedidos.ddd_entrega, Pedidos.telefone_entrega, Pedidos.email_entrega, Pedidos.subtotal, Pedidos.taxa_envio, Pedidos.tipo_taxa_adicional, Pedidos.taxa_adicional, Pedidos.total, Pedidos.tipo_frete, Pedidos.codigo_frete, Pedidos.peso_total, Pedidos.forma_pagamento, Pedidos.num_parcelas, Pedidos.cartao_encrypt, Pedidos.tipo_cartao, Pedidos.instrucoes, Pedidos.atendido, Pedidos.pago, Pedidos.falha, Pedidos.cancelado, Pedidos.devolvido, Pedidos.fraude, Pedidos.num_remessa, Pedidos.boleto_emitido, Pedidos.boleto_tipo, Pedidos.sigla_idioma, Pedidos.sigla_moeda, Pedidos.valor_moeda FROM Pedido_item INNER JOIN Pedidos ON Pedido_item.codigo_pedido = Pedidos.codigo_pedido WHERE Pedidos.codigo_pedido = " & nCodigoPedido, Conexao

If Not RS_Dados.EOF Then

    arrayDado = LCase("identificacao=") & sIdentificacao
    arrayDado = arrayDado & LCase("&modulo=") & sModulo
    arrayDado = arrayDado & LCase("&ambiente=") & sAmbiente
    arrayDado = arrayDado & LCase("&operacao=") & sOperacao

    arrayDado = arrayDado & LCase("&codPedido=") & nCodigoPedido
    arrayDado = arrayDado & LCase("&valorTotal=") & RS_Dados("total")
    arrayDado = arrayDado & LCase("&dataPedido=") & RS_Dados("data_pedido")
    arrayDado = arrayDado & LCase("&tipoFrete=") & RS_Dados("tipo_frete")
    arrayDado = arrayDado & LCase("&formaPagamento=") & RS_Dados("forma_pagamento")
    arrayDado = arrayDado & LCase("&bandeira=") & RS_Dados("tipo_cartao")
    arrayDado = arrayDado & LCase("&ipConsumidor=") & RS_Dados("ip_cliente")
    arrayDado = arrayDado & LCase("&descPedido=") & sDescPedido

    ' Resgata o número de parcelas do pedido
    If RS_Dados("forma_pagamento") = "Bradesco" Then
        nQtdParcelas = Pega_DadoBanco("Transacao_Bradesco","numParcelas","codigo_pedido",nCodigoPedido)
    ElseIf RS_Dados("forma_pagamento") = "Amex" Then
        nQtdParcelas = Pega_DadoBanco("Transacao_Amex","num_parcelas","codigo_pedido",nCodigoPedido)
    ElseIf RS_Dados("forma_pagamento") = "Mastercard" Or RS_Dados("forma_pagamento") = "Diners" Then
        nQtdParcelas = Pega_DadoBanco("Transacao_Redecard","num_parcelas","codigo_pedido",nCodigoPedido)
    ElseIf RS_Dados("forma_pagamento") = "Visa" Then
        nQtdParcelas = Pega_DadoBanco("Transacao_Visanet","num_parcelas","codigo_pedido",nCodigoPedido)
    ElseIf RS_Dados("forma_pagamento") = "ABNCDC" Then
        nQtdParcelas = Pega_DadoBanco("Transacao_Abncdc","qtd_parcelas","codigo_pedido",nCodigoPedido)
    ElseIf RS_Dados("forma_pagamento") = "Finasa" Then
        nQtdParcelas = Pega_DadoBanco("Transacao_Finasa","qtd_parcelas","codigo_pedido",nCodigoPedido)
    Else
        nQtdParcelas = 1
    End If
    ' Se for vazio define como 1
    If nQtdParcelas = "" Then
        nQtdParcelas = 1
    End If
    arrayDado = arrayDado & LCase("&qtdeParcelas=") & nQtdParcelas

    arrayDado = arrayDado & LCase("&codConsumidorCobranca=") & RS_Dados("user_ID")

    If RS_Dados("razaosocial_cobranca") <> "" And RS_Dados("cnpj_cobranca") <> "" Then
    arrayDado = arrayDado & LCase("&tipoPessoaCobranca=PJ")
    arrayDado = arrayDado & LCase("&nomeCobranca=") & RS_Dados("razaosocial_cobranca")
    arrayDado = arrayDado & LCase("&cpfCnpjCobranca=") & RS_Dados("cnpj_cobranca")
    Else
    arrayDado = arrayDado & LCase("&tipoPessoaCobranca=PF")
    arrayDado = arrayDado & LCase("&nomeCobranca=") & RS_Dados("nome_cobranca")
    arrayDado = arrayDado & LCase("&cpfCnpjCobranca=") & RS_Dados("cpf_cobranca")
    End If

    arrayDado = arrayDado & LCase("&dataNascimentoCobranca=") & FormataDado(RS_Dados("data_nascimento_cobranca"),"DATA_NASCIMENTO")
    arrayDado = arrayDado & LCase("&sexoCobranca=")
    arrayDado = arrayDado & LCase("&enderecoCobranca=") & RS_Dados("logradouro_cobranca")
    arrayDado = arrayDado & LCase("&numeroEndCobranca=") & RS_Dados("numero_cobranca")
    arrayDado = arrayDado & LCase("&complementoEndCobranca=") & RS_Dados("complemento_cobranca")
    arrayDado = arrayDado & LCase("&bairroCobranca=") & RS_Dados("bairro_cobranca")
    arrayDado = arrayDado & LCase("&cidadeCobranca=") & RS_Dados("cidade_cobranca")
    arrayDado = arrayDado & LCase("&cepCobranca=") & RS_Dados("cep_cobranca")
    arrayDado = arrayDado & LCase("&ufCobranca=") & RS_Dados("estado_cobranca")
    arrayDado = arrayDado & LCase("&paisCobranca=") & RS_Dados("pais_cobranca")
    arrayDado = arrayDado & LCase("&tipoEnderecoCobranca=")
    arrayDado = arrayDado & LCase("&ddd1Cobranca=") & RS_Dados("ddd_cobranca")
    arrayDado = arrayDado & LCase("&fone1Cobranca=") & RS_Dados("telefone_cobranca")
    arrayDado = arrayDado & LCase("&ddd2Cobranca=")
    arrayDado = arrayDado & LCase("&fone2Cobranca=")
    arrayDado = arrayDado & LCase("&emailCobranca=") & RS_Dados("user_ID")

    arrayDado = arrayDado & LCase("&codConsumidorEntrega=") & RS_Dados("user_ID")

    If RS_Dados("razaosocial_entrega") <> "" And RS_Dados("cnpj_entrega") <> "" Then
    arrayDado = arrayDado & LCase("&tipoPessoaEntrega=PJ")
    arrayDado = arrayDado & LCase("&nomeEntrega=") & RS_Dados("razaosocial_entrega")
    arrayDado = arrayDado & LCase("&cpfCnpjEntrega=") & RS_Dados("cnpj_entrega")
    Else
    arrayDado = arrayDado & LCase("&tipoPessoaEntrega=PF")
    arrayDado = arrayDado & LCase("&nomeEntrega=") & RS_Dados("nome_entrega")
    arrayDado = arrayDado & LCase("&cpfCnpjEntrega=") & RS_Dados("cpf_entrega")
    End If

    arrayDado = arrayDado & LCase("&dataNascimentoEntrega=") & FormataDado(RS_Dados("data_nascimento_entrega"),"DATA_NASCIMENTO")
    arrayDado = arrayDado & LCase("&sexoEntrega=")
    arrayDado = arrayDado & LCase("&enderecoEntrega=") & RS_Dados("logradouro_entrega")
    arrayDado = arrayDado & LCase("&numeroEndEntrega=") & RS_Dados("numero_entrega")
    arrayDado = arrayDado & LCase("&complementoEndEntrega=") & RS_Dados("complemento_entrega")
    arrayDado = arrayDado & LCase("&bairroEntrega=") & RS_Dados("bairro_entrega")
    arrayDado = arrayDado & LCase("&cidadeEntrega=") & RS_Dados("cidade_entrega")
    arrayDado = arrayDado & LCase("&cepEntrega=") & RS_Dados("cep_entrega")
    arrayDado = arrayDado & LCase("&ufEntrega=") & RS_Dados("estado_entrega")
    arrayDado = arrayDado & LCase("&paisEntrega=") & RS_Dados("pais_entrega")
    arrayDado = arrayDado & LCase("&tipoEnderecoEntrega=")
    arrayDado = arrayDado & LCase("&ddd1Entrega=") & RS_Dados("ddd_entrega")
    arrayDado = arrayDado & LCase("&fone1Entrega=") & RS_Dados("telefone_entrega")
    arrayDado = arrayDado & LCase("&ddd2Entrega=")
    arrayDado = arrayDado & LCase("&fone2Entrega=")
    arrayDado = arrayDado & LCase("&emailEntrega=") & RS_Dados("email_entrega")

    nContItem = 1

    While Not RS_Dados.EOF
        arrayDado = arrayDado & LCase("&codItem") & nContItem & "=" & RS_Dados("codigo_produto")
        arrayDado = arrayDado & LCase("&descItem") & nContItem & "=" & RS_Dados("nome_produto")
        arrayDado = arrayDado & LCase("&qtdeItem") & nContItem & "=" & RS_Dados("quantidade")
        arrayDado = arrayDado & LCase("&valorUnitItem") & nContItem & "=" & RS_Dados("preco_unitario")
        
        nContItem = nContItem + 1
        RS_Dados.MoveNext()
    Wend

    arrayDado = arrayDado & LCase("&qtdeItens=") & (nContItem-1)

    If Trim(sParamAdicionais) <> "" Then
        arrayDado = arrayDado & sParamAdicionais
    End If

Else
    arrayDado = ""
End If

RS_Dados.Close
Set RS_Dados = Nothing

' Retorna os dados da transação
Dados_Transacao = arrayDado

End Function
'########################################################################################################
'--> FIM FUNCTION Dados_Transacao
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION IsValidUTF8
' - Verifica se a string é valida no padrão UTF8
'########################################################################################################
Function IsValidUTF8(s)
  dim i
  dim c
  dim n

  IsValidUTF8 = false
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      select case n
      case 1
        exit function
      case 2
        if (c and &HE0) <> &HC0 then
          exit function
        end if
      case 3
        if (c and &HF0) <> &HE0 then
          exit function
        end if
      case 4
        if (c and &HF8) <> &HF0 then
          exit function
        end if
      case else
        exit function
      end select
      i = i + n
    else
      i = i + 1
    end if
  loop
  IsValidUTF8 = true 
End Function
'########################################################################################################
'--> FIM FUNCTION IsValidUTF8
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION DecodeUTF8
' - Decodes a UTF-8 string to the Windows character set.
' - Non-convertable characters are replace by an upside down question mark.
' - Returns: A Windows string
'########################################################################################################
Function DecodeUTF8(s)
  dim i
  dim c
  dim n

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      if n = 2 and ((c and &HE0) = &HC0) then
        c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
      else
        c = 191 
      end if
      s = left(s,i-1) + chr(c) + mid(s,i+n)
    end if
    i = i + 1
  loop
  DecodeUTF8 = s 
End  Function
'########################################################################################################
'--> FIM FUNCTION DecodeUTF8
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION EncodeUTF8
' - Encodes a Windows string in UTF-8
' - Returns: A UTF-8 encoded string
'########################################################################################################
Function EncodeUTF8(s)
  dim i
  dim c

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c >= &H80 then
      s = left(s,i-1) + chr(&HC2 + ((c and &H40) / &H40)) + chr(c and &HBF) + mid(s,i+1)
      i = i + 1
    end if
    i = i + 1
  loop
  EncodeUTF8 = s 
End Function
'########################################################################################################
'--> FIM FUNCTION EncodeUTF8
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION URLDecode
'  - Decodifica uma string pela codificação utilizada na postagem via GET
'########################################################################################################
Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If
	
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")
	
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
	
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
	
    URLDecode = sOutput
End Function
'########################################################################################################
'--> FIM FUNCTION URLDecode
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION XmlDados_Transacao
'  - Função para montar o XML dos dados da transação
'########################################################################################################
Function XmlDados_Transacao(nCodigoPedido,sParamAdicionais)

Set RS_Dados = Server.CreateObject("ADODB.Recordset")
RS_Dados.CursorLocation = 3
RS_Dados.CursorType = 0
RS_Dados.LockType =  1

RS_Dados.Open "SELECT Pedido_item.codigo_produto, Pedido_item.codigo_categoria, Pedido_item.codigo_cor, Pedido_item.codigo_tamanho, Pedido_item.quantidade, Pedido_item.preco_unitario, Pedido_item.peso, Pedido_item.nome_produto, Pedido_item.sigla_moeda, Pedido_item.valor_moeda, Pedidos.codigo_pedido, Pedidos.data_pedido_inicio, Pedidos.data_pedido, Pedidos.user_ID, Pedidos.sessionID, Pedidos.ip_cliente, Pedidos.razaosocial_cobranca, Pedidos.cnpj_cobranca, Pedidos.inscricaoestadual_cobranca, Pedidos.nome_cobranca, Pedidos.cpf_cobranca, Pedidos.rg_cobranca, Pedidos.data_nascimento_cobranca, Pedidos.logradouro_cobranca, Pedidos.numero_cobranca, Pedidos.complemento_cobranca, Pedidos.bairro_cobranca, Pedidos.cidade_cobranca, Pedidos.estado_cobranca, Pedidos.cep_cobranca, Pedidos.pais_cobranca, Pedidos.ddd_cobranca, Pedidos.telefone_cobranca, Pedidos.razaosocial_entrega, Pedidos.cnpj_entrega, Pedidos.inscricaoestadual_entrega, Pedidos.nome_entrega, Pedidos.cpf_entrega, Pedidos.rg_entrega, Pedidos.data_nascimento_entrega, Pedidos.logradouro_entrega, Pedidos.numero_entrega, Pedidos.complemento_entrega, Pedidos.bairro_entrega, Pedidos.cidade_entrega, Pedidos.estado_entrega, Pedidos.cep_entrega, Pedidos.pais_entrega, Pedidos.ddd_entrega, Pedidos.telefone_entrega, Pedidos.email_entrega, Pedidos.subtotal, Pedidos.taxa_envio, Pedidos.tipo_taxa_adicional, Pedidos.taxa_adicional, Pedidos.total, Pedidos.tipo_frete, Pedidos.codigo_frete, Pedidos.peso_total, Pedidos.forma_pagamento, Pedidos.num_parcelas, Pedidos.cartao_encrypt, Pedidos.tipo_cartao, Pedidos.instrucoes, Pedidos.atendido, Pedidos.pago, Pedidos.falha, Pedidos.cancelado, Pedidos.devolvido, Pedidos.fraude, Pedidos.num_remessa, Pedidos.boleto_emitido, Pedidos.boleto_tipo, Pedidos.sigla_idioma, Pedidos.sigla_moeda, Pedidos.valor_moeda FROM Pedido_item INNER JOIN Pedidos ON Pedido_item.codigo_pedido = Pedidos.codigo_pedido WHERE Pedidos.codigo_pedido = " & nCodigoPedido, Conexao

If Not RS_Dados.EOF Then

' Armazena os dados cadastrais
sEndereco = RS_Dados("logradouro_cobranca")
sNumero = RS_Dados("numero_cobranca")
sBairro = RS_Dados("bairro_cobranca")
sCidade = RS_Dados("cidade_cobranca")
sCep = RS_Dados("cep_cobranca")
sEstado = RS_Dados("estado_cobranca")

' Cabeçalho
arrayDado = "<?xml version=""1.0"" encoding=""utf-8"" ?>"
arrayDado = arrayDado & "<LocaWeb>"

    ' Dados do comprador
    arrayDado = arrayDado & "<Comprador>"
        arrayDado = arrayDado & "<Nome>" & RS_Dados("nome_cobranca") & "</Nome>"
        arrayDado = arrayDado & "<Email>" & RS_Dados("user_ID") & "</Email>"
        arrayDado = arrayDado & "<Cpf>" & RS_Dados("cpf_cobranca") & "</Cpf>"
        If RS_Dados("rg_cobranca") <> "" Then
            arrayDado = arrayDado & "<Rg>" & RS_Dados("rg_cobranca") & "</Rg>"
        End If
        If RS_Dados("ddd_cobranca") <> "" Then
            arrayDado = arrayDado & "<Ddd>" & RS_Dados("ddd_cobranca") & "</Ddd>"
        End If
        If RS_Dados("telefone_cobranca") <> "" Then
            arrayDado = arrayDado & "<Telefone>" & RS_Dados("telefone_cobranca") & "</Telefone>"
        End If
        If RS_Dados("razaosocial_cobranca") <> "" And RS_Dados("cnpj_cobranca") <> "" Then
            arrayDado = arrayDado & "<TipoPessoa>Juridica</TipoPessoa>"
            arrayDado = arrayDado & "<RazaoSocial>" & RS_Dados("razaosocial_cobranca") & "</RazaoSocial>"
            arrayDado = arrayDado & "<Cnpj>" & RS_Dados("cnpj_cobranca") & "</Cnpj>"
        Else
            arrayDado = arrayDado & "<TipoPessoa>Fisica</TipoPessoa>"
        End If
    arrayDado = arrayDado & "</Comprador>"

    ' Dados do pedido
    arrayDado = arrayDado & "<Pedido>"

        ' Resgata o número de parcelas do pedido
        If IsNull(RS_Dados("num_parcelas")) And RS_Dados("num_parcelas") <> "" Then
            nQtdParcelas = RS_Dados("num_parcelas")
        Else
            nQtdParcelas = 1
        End If
        ' Se for vazio define como 1
        If nQtdParcelas = "" Then
            nQtdParcelas = 1
        End If


        ' Verifica se há acrescimo
        If RS_Dados("tipo_taxa_adicional") = "Sem juros" And RS_Dados("tipo_taxa_adicional") = "Juros do Emissor" Then
            sAcrescPedido = RS_Dados("taxa_adicional")
        End If

        ' Verifica se há desconto
        If RS_Dados("tipo_taxa_adicional") = "Desconto" Then
            sDescPedido = RS_Dados("taxa_adicional")
        End If

        ' Dados Gerais
        arrayDado = arrayDado & "<Numero>" & nCodigoPedido & "</Numero>"
        arrayDado = arrayDado & "<ValorSubTotal>" & FormataValor_Transacao(RS_Dados("subtotal")) & "</ValorSubTotal>"
        arrayDado = arrayDado & "<ValorFrete>" & FormataValor_Transacao(RS_Dados("taxa_envio")) & "</ValorFrete>"
        arrayDado = arrayDado & "<ValorAcrescimo>" & FormataValor_Transacao(sAcrescPedido) & "</ValorAcrescimo>"
        arrayDado = arrayDado & "<ValorDesconto>" & FormataValor_Transacao(sDescPedido) & "</ValorDesconto>"
        arrayDado = arrayDado & "<ValorTotal>" & FormataValor_Transacao(RS_Dados("total")) & "</ValorTotal>"

        ' Itens do pedido
        arrayDado = arrayDado & "<Itens>"

            nContItem = 1

            While Not RS_Dados.EOF
                nValorTotal = RS_Dados("preco_unitario") * RS_Dados("quantidade")

                arrayDado = arrayDado & "<Item>"
                    arrayDado = arrayDado & "<CodProduto>" & RS_Dados("codigo_produto") & "</CodProduto>"
                    arrayDado = arrayDado & "<DescProduto>" & RS_Dados("nome_produto") & "</DescProduto>"
                    arrayDado = arrayDado & "<Quantidade>" & RS_Dados("quantidade") & "</Quantidade>"
                    arrayDado = arrayDado & "<ValorUnitario>" & FormataValor_Transacao(RS_Dados("preco_unitario")) & "</ValorUnitario>"
                    arrayDado = arrayDado & "<ValorTotal>" & FormataValor_Transacao(nValorTotal) & "</ValorTotal>"
                arrayDado = arrayDado & "</Item>"

                nContItem = nContItem + 1
                RS_Dados.MoveNext()
            Wend
        
        arrayDado = arrayDado & "</Itens>"

        ' Dados de cobrança
        arrayDado = arrayDado & "<Cobranca>"
            arrayDado = arrayDado & "<Endereco>" & sEndereco & "</Endereco>"
            arrayDado = arrayDado & "<Numero>" & sNumero & "</Numero>"
            arrayDado = arrayDado & "<Bairro>" & sBairro & "</Bairro>"
            arrayDado = arrayDado & "<Cidade>" & sCidade & "</Cidade>"
            arrayDado = arrayDado & "<Cep>" & sCep & "</Cep>"
            arrayDado = arrayDado & "<Estado>" & sEstado & "</Estado>"
        arrayDado = arrayDado & "</Cobranca>"

    arrayDado = arrayDado & "</Pedido>"

    If Trim(sParamAdicionais) <> "" Then
        arrayDado = arrayDado & sParamAdicionais
    End If

arrayDado = arrayDado & "</LocaWeb>"

Else
    arrayDado = ""
End If

RS_Dados.Close
Set RS_Dados = Nothing

' Retorna os dados da transação
XmlDados_Transacao = arrayDado

End Function
'########################################################################################################
'--> FIM FUNCTION XmlDados_Transacao
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION FormataValor_Transacao
'  - Função para formatar os valores das transações
'########################################################################################################
Function FormataValor_Transacao(fctValor)

    ' Verifica se o valor está vazio
    If fctValor = "" Or IsNull(fctValor) Then
        nValor = 0
    End If

    ' Formata com casas decimais
    nValor = FormatNumber(fctValor,2)

    ' Retira a formatação
    nValor = Replace(nValor,".","")
    nValor = Replace(nValor,",","")

    ' Retorna o valor formatado
    FormataValor_Transacao = nValor

End Function
'########################################################################################################
'--> FIM FUNCTION FormataValor_Transacao
'########################################################################################################
%>