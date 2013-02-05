<?php
/*
+--------------------------------------------------------------------------
|   CubeCart v3
|   ========================================
|   by Alistair Brookbanks
|	CubeCart is a Trade Mark of Devellion Limited
|   Copyright Devellion Limited 2005 - 2006. All rights reserved.
|   Devellion Limited,
|   5 Bridge Street,
|   Bishops Stortford,
|   HERTFORDSHIRE.
|   CM23 2JU
|   UNITED KINGDOM
|   http://www.devellion.com
|	UK Private Limited Company No. 5323904
|   ========================================
|   Web: http://www.cubecart.com
|   Date: Tuesday, 17th July 2007
|   Email: sales (at) cubecart (dot) com
|	License Type: CubeCart is NOT Open Source Software and Limitations Apply 
|   Licence Info: http://www.cubecart.com/site/faq/license.php
+--------------------------------------------------------------------------
|	retorno.php
|   ========================================
|	Retorno do gateway de pagamento PagSeguro
+--------------------------------------------------------------------------
*/


include("../../../includes/ini.inc.php");
include("../../../includes/global.inc.php");
require_once("../../../classes/db.inc.php");
$db = new db();
include_once("../../../includes/functions.inc.php");
$config = fetchDbConfig("config");
include_once("../../../language/".$config['defaultLang']."/lang.inc.php");
include("../../../includes/currencyVars.inc.php");

$module = fetchDbConfig("PagSeguro");

require_once("biblioteca_pagseguro_v0.21/retorno.php");

function retorno_automatico ( $VendedorEmail, $TransacaoID, 
    $Referencia, $TipoFrete, $ValorFrete, $Anotacao, $DataTransacao,
    $TipoPagamento, $StatusTransacao, $CliNome, $CliEmail, 
    $CliEndereco, $CliNumero, $CliComplemento, $CliBairro, $CliCidade,
    $CliEstado, $CliCEP, $CliTelefone, $produtos, $NumItens) {

    $pago=false;$debugErr='';

    global $module, $db, $glob;
    $summary = $db->select("SELECT prod_total, comments FROM ".$glob['dbprefix']."CubeCart_order_sum WHERE cart_order_id = ".$db->mySQLsafe($Referencia));

    if('Completo'==$StatusTransacao || 'Aprovado'==$StatusTransacao){

        $checado=true;
	    // checa email
	    if($VendedorEmail!==trim($module['email'])){
            $checado=false;
            $debugErr='Email não confere: '.$VendedorEmail.' - '.trim($module['email']);
        }
        if($VendedorEmail=='email_cobranca'){
            $Anotacao='Pagamento gerado pelo ambiente de testes do PagSeguro.';
            $checado=true;
        }
	
	    // checa valor
        $valorTotal=0;
        for($i=1;$i<=$NumItens;$i++){
            $valorI=str_replace(',','.',$_POST['ProdValor_'.$i]);
            $freteI=str_replace(',','.',$_POST['ProdFrete_'.$i]);
            $valorTotal+=$valorI+$freteI;
        }
	    if(floatval($valorTotal)!==floatval($summary[0]['prod_total'])){
            $checado=false;
            $debugErr='Valor não confere: '.$valorTotal.' - '.$summary[0]['prod_total'];
        }
	
	    // se estiver OK, marca como pago
	    if($checado){
		    $cart_order_id = $Referencia;
		    include("../../../includes/orderSuccess.inc.php");
            $pago=true;
	    }
    }

    $updateComment['comments'] = "'" . ( empty($summary[0]['comments']) ? '' : $summary[0]['comments']."\r\n\r\n" ) . "Tipo de Pagamento: ".utf8_decode($TipoPagamento)." \r\nStatus da Transação: ".utf8_decode($StatusTransacao)." - ".date("d/m/Y H:i:s")."\r\nAnotação: ".utf8_decode($Anotacao)."'";
	$update = $db->update($glob['dbprefix']."CubeCart_order_sum", $updateComment,"cart_order_id=".$db->mySQLSafe($Referencia));

    echo $pago?'OK':'FAIL '.$debugErr;
}

$dest=$glob['storeURL'].'/cart.php?act=viewOrders';
header("Location: ".$dest);

?>
