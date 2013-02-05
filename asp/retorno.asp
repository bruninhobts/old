<%
response.ContentType = "text/HTML"
response.Charset = "ISO-8859-1"

Dim TOKEN

'TOKEN = "cole aqui o token do vendedor"



timeout = 20  'Timeout em segundos

function notificationPost()
	
	postdata = "Comando=validar&Token=" & TOKEN
	
	For each x In Request.Form
		valued = clearStr(request.Form(x))
		postdata = postdata & "&" & x & "=" & valued
	Next
	
	notificationPost = verify(postdata)
	
end Function

function clearStr(str)
	
	str = replace(str, "'","\'")
	clearStr = str
	
end function

function verify(data)
	
	strUrl = "https://pagseguro.uol.com.br/pagseguro-ws/checkout/NPI.jhtml"
	
	Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP")
	
	xmlHttp.Open "POST", strUrl, False
	xmlHttp.setRequestHeader "User-Agent", "asp httprequest"
	xmlHttp.setRequestHeader "content-type", "application/x-www-form-urlencoded"
	xmlHttp.setRequestHeader "content-length", Len(data)
	xmlHttp.Send(data)
	
	retorno = xmlHttp.responseText
	
	xmlHttp.abort()
	
	set xmlHttp = Nothing
	
	verify = retorno

end function

if Request.Form.count > 0 then
	
	result = notificationPost()
	
	if Request.Form("TransacaoID") <> empty then
		transacaoID = Request.Form("TransacaoID")
	Else
		transacaoID = ""
	end If
	
	if result = "VERIFICADO" then
		'O post foi validado pelo PagSeguro.
	elseif result = "FALSO" then
		'O post não foi validado pelo PagSeguro.
	else
		'Erro na integração com o PagSeguro.
	end if
	
else
	' POST não recebido, indica que a requisição é o retorno do Checkout PagSeguro.
	' No término do checkout o usuário é redirecionado para este bloco.
	%>
	
    <h3>Obrigado por efetuar a compra.</h3>

    <%
end if
%>
