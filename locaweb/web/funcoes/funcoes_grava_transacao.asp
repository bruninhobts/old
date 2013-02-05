<%
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#
' Loja Exemplo Locaweb 
' Versão: 6.5
' Data: 12/09/06
' Arquivo: funcoes_grava_transacao.asp
' Versão do arquivo: 0.0
' Data da ultima atualização: 21/10/08
'
'-----------------------------------------------------------------------------
' Licença Código Livre: http://comercio.locaweb.com.br/gpl/gpl.txt
'-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#

'##########################################################################################################
'FUNCTION GravaTransacaoInicialVisa
' - Grava na tabela Transacao_Visanet no incio da transação
'##########################################################################################################
Function GravaTransacaoInicialVisa(CODPEDIDO,TID,PRICE,TIPOCARTAO,NUMPARCELAS,JUROS,METODO,AMBIENTE,IDENTIFICACAOLOCAWEB)

Set RS_Visanet = CreateObject("ADODB.Recordset")
Set RS_Visanet.ActiveConnection = Conexao
    RS_Visanet.CursorLocation = 3
    RS_Visanet.CursorType = 0
    RS_Visanet.LockType =  3

    RS_Visanet.Open "SELECT codigo_pedido, price, tid, lr, arp, free, pan, bank, ars, authent, tipo_cartao, num_parcelas, juros, captura, msg_captura, metodo, ambiente, identificacaoLocaweb FROM Transacao_Visanet WHERE codigo_pedido = "&CODPEDIDO&"", Conexao

    If RS_Visanet.EOF Then
        'Insere um novo registro   
        RS_Visanet.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Visanet("codigo_pedido")          = CODPEDIDO
    End if
    If TID <> "" Then
        RS_Visanet("tid")                    = TID
    End If
    If PRICE <> "" Then
        RS_Visanet("price")                  = PRICE
    End If
    If TIPOCARTAO <> "" Then
        RS_Visanet("tipo_cartao")            = TIPOCARTAO
    End If
    If NUMPARCELAS <> "" Then
        RS_Visanet("num_parcelas")           = NUMPARCELAS
    End If
    If JUROS <> "" Then
        RS_Visanet("juros")                  = JUROS
    End If
    If METODO <> "" Then
        RS_Visanet("metodo")                 = METODO
    End If
	If AMBIENTE <> "" Then
        RS_Visanet("ambiente")               = AMBIENTE
    End If
	If IDENTIFICACAOLOCAWEB <> "" Then
        RS_Visanet("identificacaoLocaweb")   = IDENTIFICACAOLOCAWEB
    End If


    RS_Visanet.Update
    RS_Visanet.Close
Set RS_Visanet = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoInicialVisa
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoFinalVisa
' - Grava na tabela Transacao_Visanet o final da transação
'##########################################################################################################
Function GravaTransacaoFinalVisa(ORDERID,PRICE,TID,LR,ARP,FREE,PAN,BANK,ARS,AUTHENT,IDREQLOCAWEB)

Set RS_Visanet = CreateObject("ADODB.Recordset")
Set RS_Visanet.ActiveConnection = Conexao
    RS_Visanet.CursorLocation = 3
    RS_Visanet.CursorType = 0
    RS_Visanet.LockType =  3

    RS_Visanet.Open "SELECT codigo_pedido, idReqLocaWeb, price, tid, lr, arp, free, pan, bank, ars, authent, tipo_cartao, num_parcelas, juros, captura, msg_captura, metodo FROM Transacao_Visanet WHERE codigo_pedido = "&ORDERID&"", Conexao

    If RS_Visanet.EOF Then
        'Insere um novo registro   
        RS_Visanet.Addnew
    End If

    If ORDERID <> "" Then
        RS_Visanet("codigo_pedido") = ORDERID
    End If
	If IDREQLOCAWEB <> "" Then
        RS_Visanet("idReqLocaWeb") = IDREQLOCAWEB
    End If
    If PRICE <> "" Then
        RS_Visanet("price")         = PRICE
    End If
    If TID <> "" Then
        RS_Visanet("tid")           = TID
    End If
    If LR <> "" Then
        RS_Visanet("lr")            = LR
    End if
    If ARP <> "" Then
        RS_Visanet("arp")           = ARP
    End If
    If FREE <> "" Then
        RS_Visanet("free")          = FREE
    End If
    If PAN <> "" Then
        RS_Visanet("pan")           = PAN
    End If
    If BANK <> "" Then
        RS_Visanet("bank")          = BANK
    End If 
    If ARS <> "" Then
        RS_Visanet("ars")           = ARS
    End If
    If AUTHENT <> "" Then
        RS_Visanet("authent")       = AUTHENT
    End If

    RS_Visanet.Update
    RS_Visanet.Close
Set RS_Visanet = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoFinalVisa
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoInicialRedecard
' - Grava na tabela Transacao_Redecard no incio da transação
'##########################################################################################################
Function GravaTransacaoInicialRedecard(CODPEDIDO,TIPOCARTAO,NUMPARCELAS,JUROS)

Set RS_Redecard = CreateObject("ADODB.Recordset")
Set RS_Redecard.ActiveConnection = Conexao
    RS_Redecard.CursorLocation = 3
    RS_Redecard.CursorType = 0
    RS_Redecard.LockType =  3

    RS_Redecard.Open "SELECT codigo_pedido, CODRET, MSGRET, NUMAUTOR, NUMSQN, NUMCV, NUMAUTENT, NR_CARTAO, ORIGEM_BIN, PAX1, RESPAVS, MSGAVS, tipo_cartao, num_parcelas, juros, CODRET_confirmacao, MSGRET_confirmacao FROM Transacao_Redecard WHERE codigo_pedido = "&CODPEDIDO&"", Conexao

    If RS_Redecard.EOF Then
        'Insere um novo registro   
        RS_Redecard.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Redecard("codigo_pedido") = CODPEDIDO
    End if
    If TIPOCARTAO <> "" Then
        RS_Redecard("tipo_cartao") = TIPOCARTAO
    End If
    If NUMPARCELAS <> "" Then
        RS_Redecard("num_parcelas") = NUMPARCELAS
    End If
    If JUROS <> "" Then
        RS_Redecard("juros") = JUROS
    End If

    RS_Redecard.Update
    RS_Redecard.Close
Set RS_Redecard = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoInicialRedecard
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoFinalRedecard
' - Grava na tabela Transacao_Redecard o final da transação
'##########################################################################################################
Function GravaTransacaoFinalRedecard(CODPEDIDO,CODRET,MSGRET,NUMAUTOR,NUMSQN,NUMCV,NUMAUTENT,NR_CARTAO,ORIGEM_BIN,PAX1,RESPAVS,MSGAVS,CODRETconf,MSGRETconf)

Set RS_Redecard = CreateObject("ADODB.Recordset")
Set RS_Redecard.ActiveConnection = Conexao
    RS_Redecard.CursorLocation = 3
    RS_Redecard.CursorType = 0
    RS_Redecard.LockType =  3

    RS_Redecard.Open "SELECT codigo_pedido, CODRET, MSGRET, NUMAUTOR, NUMSQN, NUMCV, NUMAUTENT, NR_CARTAO, ORIGEM_BIN, PAX1, RESPAVS, MSGAVS, tipo_cartao, num_parcelas, juros, CODRET_confirmacao, MSGRET_confirmacao FROM Transacao_Redecard WHERE codigo_pedido = "&CODPEDIDO&"", Conexao

    If RS_Redecard.EOF Then
        'Insere um novo registro   
        RS_Redecard.Addnew
    End If

    If CODRET <> "" Then
        RS_Redecard("CODRET") = CODRET
    End if
    If MSGRET <> "" Then
        RS_Redecard("MSGRET") = MSGRET
    End if
    If NUMAUTOR <> "" Then
        RS_Redecard("NUMAUTOR") = NUMAUTOR
    End if
    If NUMSQN <> "" Then
        RS_Redecard("NUMSQN") = NUMSQN
    End if
    If NUMCV <> "" Then
        RS_Redecard("NUMCV") = NUMCV
    End if
    If NUMAUTENT <> "" Then
        RS_Redecard("NUMAUTENT") = NUMAUTENT
    End if
    If NR_CARTAO <> "" Then
        RS_Redecard("NR_CARTAO") = NR_CARTAO
    End if
    If ORIGEM_BIN <> "" Then
        RS_Redecard("ORIGEM_BIN") = ORIGEM_BIN
    End if
    If PAX1 <> "" Then
        RS_Redecard("PAX1") = PAX1
    End if
    If RESPAVS <> "" Then
        RS_Redecard("RESPAVS") = RESPAVS
    End if
    If MSGAVS <> "" Then
        RS_Redecard("MSGAVS") = MSGAVS
    End if
    If CODRETconf <> "" Then
        RS_Redecard("CODRET_confirmacao") = CODRETconf
    End if
    If MSGRETconf <> "" Then
        RS_Redecard("MSGRET_confirmacao") = MSGRETconf
    End if

    RS_Redecard.Update
    RS_Redecard.Close
Set RS_Redecard = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoFinalRedecard
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoFinalItau
' - Grava na tabela Transacao_Itau o resultado da transação 
'##########################################################################################################
Function GravaTransacaoFinalItau(CODPEDIDO,TIPPAG)

Set RS_Itau = CreateObject("ADODB.Recordset")
Set RS_Itau.ActiveConnection = Conexao
    RS_Itau.CursorLocation = 3
    RS_Itau.CursorType = 0
    RS_Itau.LockType =  3

    RS_Itau.Open "SELECT codigo_pedido, tipPag, sitPag, dtPag, codAut, numId, compVend, tipCart FROM Transacao_Itau WHERE codigo_pedido = "& CODPEDIDO, Conexao

    If RS_Itau.EOF Then
        'Insere um novo registro   
        RS_Itau.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Itau("codigo_pedido") = CODPEDIDO
    End If

    If TIPPAG <> "" Then
        RS_Itau("tipPag") = TIPPAG
    End If

    RS_Itau.Update
    RS_Itau.Close
Set RS_Itau = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoFinalItau
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoInicialAmex
' - Grava na tabela Transacao_Amex no incio da transação
'##########################################################################################################
Function GravaTransacaoInicialAmex(CODPEDIDO,NUMPARCELAS,PLANTYPE)

Set RS_Amex = CreateObject("ADODB.Recordset")
Set RS_Amex.ActiveConnection = Conexao
    RS_Amex.CursorLocation = 3
    RS_Amex.CursorType = 0
    RS_Amex.LockType =  3

    RS_Amex.Open "SELECT codigo_pedido, num_parcelas, plantype, status_captura FROM Transacao_Amex WHERE codigo_pedido = "&CODPEDIDO&"", Conexao

    If RS_Amex.EOF Then
        'Insere um novo registro   
        RS_Amex.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Amex("codigo_pedido") = CODPEDIDO
    End if
    If NUMPARCELAS <> "" Then
        RS_Amex("num_parcelas") = NUMPARCELAS
    End If
    If PLANTYPE <> "" Then
        RS_Amex("PlanType") = PLANTYPE
    End If

    RS_Amex.Update
    RS_Amex.Close
Set RS_Amex = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoInicialAmex
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoFinalAmex
' - Grava na tabela Transacao_Amex o resultado da transação 
'##########################################################################################################
Function GravaTransacaoFinalAmex(MerchTxnRef,BatchNo,ReceiptNo,AuthorizeId,TransactionNo,AcqResponseCode,TxnResponseCode,Message,CSCResultCode,CSCRequestCode,AcqCSCRespCode)

Set RS_Amex = CreateObject("ADODB.Recordset")
Set RS_Amex.ActiveConnection = Conexao
    RS_Amex.CursorLocation = 3
    RS_Amex.CursorType = 0
    RS_Amex.LockType =  3

    RS_Amex.Open "SELECT codigo_pedido, BatchNo, ReceiptNo, AuthorizeId, TransactionNo, AcqResponseCode, TxnResponseCode, Message, CSCResultCode, CSCRequestCode, AcqCSCRespCode FROM Transacao_Amex WHERE codigo_pedido = "& MerchTxnRef, Conexao

    If RS_Amex.EOF Then
        'Insere um novo registro   
        RS_Amex.Addnew
    End If

    If MerchTxnRef <> "" Then
        RS_Amex("codigo_pedido") = MerchTxnRef
    End If

    If BatchNo <> "" Then
        RS_Amex("BatchNo") = BatchNo
    End If
    
    If ReceiptNo <> "" Then
        RS_Amex("ReceiptNo") = ReceiptNo
    End If

    If AuthorizeId <> "" Then
        RS_Amex("AuthorizeId") = AuthorizeId
    End If

    If TransactionNo <> "" Then
        RS_Amex("TransactionNo") = TransactionNo
    End If

    If AcqResponseCode <> "" Then
        RS_Amex("AcqResponseCode") = AcqResponseCode
    End If

    If TxnResponseCode <> "" Then
        RS_Amex("TxnResponseCode") = TxnResponseCode
    End If

    If Message <> "" Then
        RS_Amex("Message") = Message
    End If

    If CSCResultCode <> "" Then
        RS_Amex("CSCResultCode") = CSCResultCode
    End If

    If CSCRequestCode <> "" Then
        RS_Amex("CSCRequestCode") = CSCRequestCode
    End If

    If AcqCSCRespCode <> "" Then
        RS_Amex("AcqCSCRespCode") = AcqCSCRespCode
    End If

    RS_Amex.Update
    RS_Amex.Close
Set RS_Amex = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoFinalRedecard
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoInicialBradesco
' - Grava na tabela Transacao_Bradesco no incio da transação
'##########################################################################################################
Function GravaTransacaoInicialBradesco(CODPEDIDO)

Set RS_Bradesco = CreateObject("ADODB.Recordset")
Set RS_Bradesco.ActiveConnection = Conexao
    RS_Bradesco.CursorLocation = 3
    RS_Bradesco.CursorType = 0
    RS_Bradesco.LockType =  3

    RS_Bradesco.Open "SELECT codigo_pedido FROM Transacao_Bradesco WHERE codigo_pedido = "&CODPEDIDO&"", Conexao

    If RS_Bradesco.EOF Then
        'Insere um novo registro   
        RS_Bradesco.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Bradesco("codigo_pedido") = CODPEDIDO
    End if

    RS_Bradesco.Update
    RS_Bradesco.Close
Set RS_Bradesco = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoInicialBradesco
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoFinalBrasil
' - Grava na tabela Transacao_Brasil os dados da transação 
'##########################################################################################################
Function GravaTransacaoFinalBrasil(CODPEDIDO,REFTRAN,VALOR,TPPAGAMENTO)

Set RS_BB = CreateObject("ADODB.Recordset")
Set RS_BB.ActiveConnection = Conexao
    RS_BB.CursorLocation = 3
    RS_BB.CursorType = 0
    RS_BB.LockType =  3

    RS_BB.Open "SELECT codigo_pedido, valor, refTran, tpPagamento FROM Transacao_Brasil WHERE codigo_pedido = "& CODPEDIDO, Conexao

    If RS_BB.EOF Then
        'Insere um novo registro   
        RS_BB.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_BB("codigo_pedido") = CODPEDIDO
    End If

    If TPPAGAMENTO <> "" Then
        RS_BB("tpPagamento") = TPPAGAMENTO
    End If

    If REFTRAN <> "" Then
        RS_BB("refTran") = REFTRAN
    End If

    If VALOR <> "" Then
        RS_BB("valor") = VALOR
    End If

    RS_BB.Update
    RS_BB.Close
Set RS_BB = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoFinalBrasil
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoUnibanco
' - Grava na tabela Transacao_Unibanco no incio da transação
'########################################################################################################
Function GravaTransacaoUnibanco(CODPEDIDO,CODPARCEIRO,SESSAOPARCEIRO)

Set RS_Unibanco = CreateObject("ADODB.Recordset")
Set RS_Unibanco.ActiveConnection = Conexao
    RS_Unibanco.CursorLocation = 3
    RS_Unibanco.CursorType = 0
    RS_Unibanco.LockType =  3

    RS_Unibanco.Open "SELECT codigo_pedido, codigo_parceiro, sessao_parceiro FROM Transacao_Unibanco WHERE codigo_pedido = "&CODPEDIDO&"", Conexao

    If RS_Unibanco.EOF Then
        'Insere um novo registro   
        RS_Unibanco.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Unibanco("codigo_pedido")    = CODPEDIDO
    End if
    If CODPARCEIRO <> "" Then
        RS_Unibanco("codigo_parceiro")  = CODPARCEIRO
    End If
    If SESSAOPARCEIRO <> "" Then
        RS_Unibanco("sessao_parceiro")  = SESSAOPARCEIRO
    End If

    RS_Unibanco.Update
    RS_Unibanco.Close
Set RS_Unibanco = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoUnibanco
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoPagSeguro
' - Grava na tabela transacao_pagseguro no fim da transação
'########################################################################################################
Function GravaTransacaoPagSeguro(CODPEDIDO, TransacaoID, DataTransacao, TipoPagamento, StatusTransacao, CliEmail)

Set RS_PagSeguro = CreateObject("ADODB.Recordset")
Set RS_PagSeguro.ActiveConnection = Conexao
    RS_PagSeguro.CursorLocation = 3
    RS_PagSeguro.CursorType = 0
    RS_PagSeguro.LockType =  3

    RS_PagSeguro.Open "SELECT codigo_pedido, transacaoid, datatransacao, tipopagamento, statustransacao, cliemail FROM Transacao_PagSeguro WHERE codigo_pedido = "&CODPEDIDO&"", Conexao

    If RS_PagSeguro.EOF Then
        'Insere um novo registro   
        RS_PagSeguro.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_PagSeguro("codigo_pedido")   = CODPEDIDO
    End if
    If TransacaoID <> "" Then
        RS_PagSeguro("transacaoid")     = TransacaoID
    End If
    If CStr(DataTransacao) <> "" Then
        RS_PagSeguro("datatransacao")   = CStr(DataTransacao)
    End If
    If TipoPagamento <> "" Then
        RS_PagSeguro("tipopagamento")   = TipoPagamento
    End If
    If StatusTransacao <> "" Then
        RS_PagSeguro("statustransacao") = StatusTransacao
    End If
    If CliEmail <> "" Then
        RS_PagSeguro("cliemail")        = CliEmail
    End If

    RS_PagSeguro.Update
    RS_PagSeguro.Close
Set RS_PagSeguro = nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoPagSeguro
'########################################################################################################

'########################################################################################################
'FUNCTION GravaTransacaoInicialPagamentoCerto
' - Grava na tabela transacao_pagamentocerto no incio da transação
'########################################################################################################
Function GravaTransacaoInicialPagamentoCerto(CODPEDIDO,idTransacao,codigoTransacao,data,modulo,tipoModulo,codRetornoInicio,msgRetornoInicio)

Set RS_PagCerto = CreateObject("ADODB.Recordset")
Set RS_PagCerto.ActiveConnection = Conexao
    RS_PagCerto.CursorLocation = 3
    RS_PagCerto.CursorType = 0
    RS_PagCerto.LockType =  3

    strSQL = "SELECT codigo_pedido, idTransacao, codigo, data, modulo, tipo, codRetornoInicioTransac, msgRetornoInicioTransac FROM Transacao_PagamentoCerto WHERE codigo_pedido = "&CODPEDIDO&""

    RS_PagCerto.Open strSQL, Conexao

    If RS_PagCerto.EOF Then
        'Insere um novo registro   
        RS_PagCerto.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_PagCerto("codigo_pedido") = CODPEDIDO
    End If

    If idTransacao <> "" Then
        RS_PagCerto("idTransacao") = idTransacao
    End If
    
    If codigoTransacao <> "" Then
        RS_PagCerto("codigo") = codigoTransacao
    End If
    
    If data <> "" Then
        RS_PagCerto("data") = data
    End If
    
    If modulo <> "" Then
        RS_PagCerto("modulo") = modulo
    End If
    
    If tipoModulo <> "" Then
        RS_PagCerto("tipo") = tipoModulo
    End If

    If codRetornoInicio <> "" Then
        RS_PagCerto("codRetornoInicioTransac") = codRetornoInicio
    End If
    
    If msgRetornoInicio <> "" Then
        RS_PagCerto("msgRetornoInicioTransac") = msgRetornoInicio
    End if

    RS_PagCerto.Update
    RS_PagCerto.Close
Set RS_PagCerto = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoInicialPagamentoCerto
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoFinalPagamentoCerto
' - Grava na tabela transacao_pagamentocerto no final da transação
'########################################################################################################
Function GravaTransacaoFinalPagamentoCerto(CODPEDIDO,modulo,tipoModulo,idTransacao,codigoTransacao,dataTransacao,codRetornoConsulta,msgRetornoConsulta,processadoPagamento,msgRetornoPagamento)

Set RS_PagCerto = CreateObject("ADODB.Recordset")
Set RS_PagCerto.ActiveConnection = Conexao
    RS_PagCerto.CursorLocation = 3
    RS_PagCerto.CursorType = 0
    RS_PagCerto.LockType =  3

    strSQL = "SELECT codigo_pedido, idTransacao, codigo, data, modulo, tipo, codRetornoConsultaTransac, msgRetornoConsultaTransac, processadoPagamento, msgRetornoPagamento FROM Transacao_PagamentoCerto WHERE codigo_pedido = "&CODPEDIDO&""

    RS_PagCerto.Open strSQL, Conexao

    If RS_PagCerto.EOF Then
        'Insere um novo registro   
        RS_PagCerto.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_PagCerto("codigo_pedido") = CODPEDIDO
    End If

    If modulo <> "" Then
        RS_PagCerto("modulo") = modulo
    End If

    If tipoModulo <> "" Then
        RS_PagCerto("tipo") = tipoModulo
    End If

    If idTransacao <> "" Then
        RS_PagCerto("idTransacao") = idTransacao
    End If
    
    If codigoTransacao <> "" Then
        RS_PagCerto("codigo") = codigoTransacao
    End If
    
    If dataTransacao <> "" Then
        RS_PagCerto("data") = dataTransacao
    End If
    
    If codRetornoConsulta <> "" Then
        RS_PagCerto("codRetornoConsultaTransac") = codRetornoConsulta
    End If
    
    If msgRetornoConsulta <> "" Then
        RS_PagCerto("msgRetornoConsultaTransac") = msgRetornoConsulta
    End If

	If processadoPagamento <> "" Then
        RS_PagCerto("processadoPagamento") = processadoPagamento
    End If
    
    If msgRetornoPagamento <> "" Then
        RS_PagCerto("msgRetornoPagamento") = msgRetornoPagamento
    End If

    RS_PagCerto.Update
    RS_PagCerto.Close
Set RS_PagCerto = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoFinalPagamentoCerto
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoInicialPaggo
' - Grava na tabela transacao_paggo no incio da transação
'########################################################################################################
Function GravaTransacaoInicialPaggo(CODPEDIDO,numeroCelular,data)

Set RS_Paggo = CreateObject("ADODB.Recordset")
Set RS_Paggo.ActiveConnection = Conexao
    RS_Paggo.CursorLocation = 3
    RS_Paggo.CursorType = 0
    RS_Paggo.LockType =  3

    strSQL = "SELECT codigo_pedido, numeroCelular, data FROM Transacao_Paggo WHERE codigo_pedido = "&CODPEDIDO&""

    RS_Paggo.Open strSQL, Conexao

    If RS_Paggo.EOF Then
        'Insere um novo registro   
        RS_Paggo.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Paggo("codigo_pedido") = CODPEDIDO
    End If

	If numeroCelular <> "" Then
        RS_Paggo("numeroCelular") = Trim(numeroCelular)
    End If

    If data <> "" Then
        RS_Paggo("data") = data
    End If

    RS_Paggo.Update
    RS_Paggo.Close
Set RS_Paggo = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoInicialPaggo
'########################################################################################################
'========================================================================================================
'########################################################################################################
'FUNCTION GravaTransacaoFinalPaggo
' - Grava na tabela transacao_paggo no final da transação
'########################################################################################################
Function GravaTransacaoFinalPaggo(CODPEDIDO, merchantIdentification, idReqLocaWeb, codRetornoTransacao, msgRetornoTransacao, nsuPaggo)

Call abre_conexao(Conexao)

Set RS_Paggo = CreateObject("ADODB.Recordset")
Set RS_Paggo.ActiveConnection = Conexao
    RS_Paggo.CursorLocation = 3
    RS_Paggo.CursorType = 0
    RS_Paggo.LockType =  3

    strSQL = "SELECT codigo_pedido, merchantIdentification, idReqLocaWeb, codRetornoTransacao, msgRetornoTransacao, nsuPaggo FROM Transacao_Paggo WHERE codigo_pedido = "&CODPEDIDO&""

    RS_Paggo.Open strSQL, Conexao

    If RS_Paggo.EOF Then
        'Insere um novo registro   
        RS_Paggo.Addnew
    End If

    If CODPEDIDO <> "" Then
        RS_Paggo("codigo_pedido") = CODPEDIDO
    End If

    If merchantIdentification <> "" Then
        RS_Paggo("merchantIdentification") = merchantIdentification
    End If

    If idReqLocaWeb <> "" Then
        RS_Paggo("idReqLocaWeb") = idReqLocaWeb
    End If

	If codRetornoTransacao <> "" Then
        RS_Paggo("codRetornoTransacao") = codRetornoTransacao
    End If

	If msgRetornoTransacao <> "" Then
        RS_Paggo("msgRetornoTransacao") = msgRetornoTransacao
    End If

	If nsuPaggo <> "" Then
        RS_Paggo("nsuPaggo") = nsuPaggo
    End If

    RS_Paggo.Update
    RS_Paggo.Close
Set RS_Paggo = Nothing

End Function
'########################################################################################################
'--> FIM FUNCTION GravaTransacaoFinalPaggo
'########################################################################################################
%>