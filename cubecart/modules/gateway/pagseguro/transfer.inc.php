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
|	transfer.php
|   ========================================
|	Funções para o gateway de pagamento PagSeguro
+--------------------------------------------------------------------------
*/

$module = fetchDbConfig("PagSeguro");

require_once("biblioteca_pagseguro_v0.21/pgs.php");
$formAction = "https://pagseguro.uol.com.br/security/webpagamentos/webpagto.aspx";
$formMethod = "post";
$formTarget = "_self";

function repeatVars(){
    return FALSE;
}

function fixedVars(){
	global $glob, $db, $module, $basket, $ccUserData, $cart_order_id, $config, $GLOBALS, $pgs;
    $curr=$db->select("SELECT * FROM ".$glob['dbprefix']."CubeCart_currencies ORDER BY name ASC");
    $fator=1;
    foreach($curr as $moeda){
        if($moeda["active"]==1 && $moeda["code"]==$config['defaultCurrency']){
            $fator=$moeda["value"];
            break;
        }
    }
	$amount = sprintf("%.2f",$basket['subTotal']*$fator+$basket['tax']*$fator);

    $pgs=new pgs(array('email_cobranca'=>$module['email'], 'ref_transacao'=>$cart_order_id));
    $pgs->adicionar(array(
        array(
          "descricao"=> "Cart Order No: ".$cart_order_id,
          "valor"=>$amount,
          "quantidade"=>1,
          "id"=>base64_encode($cart_order_id),
          "frete"=>intval($basket['shipCost']*100*$fator),
        )
    ));
    $num=null;
    $arrend=split(' ',$ccUserData[0]['add_1']);
    array_reverse(&$arrend);
    foreach($arrend as $part) if(is_numeric($part)){
        $num=$part;
        break;
    }
    $ddd=null;
    $tel=$ccUserData[0]['phone'];
    $arrtel=split(' ',str_replace('-',' ',$ccUserData[0]['phone']));
    if(count($arrtel)>1){
        if(strlen($arrtel[0])==2){
            $ddd=array_shift($arrtel);
            $tel=join(' ',$arrtel);
        }
    }
    $pgs->cliente(
      array (
       'nome'   => $ccUserData[0]['firstName'].' '.$ccUserData[0]['lastName'],
       'email'    => $ccUserData[0]['email'],
       'cep'    => $ccUserData[0]['postcode'],
       'num'    => $num,
       'compl'  => $ccUserData[0]['add_2'],
       'ddd'    => $ddd,
       'tel'    => $tel,
      )
    );
    $form=$pgs->mostra(
        array (
          'print'       => false,
          'open_form'   => false,
          'close_form'  => false,
          'show_submit' => false,
          'img_button'  => false,
          'bnt_submit'  => false,
        )
    );
	return $form;
}

function success(){
	global $db, $glob, $module, $basket;
	return false;
}

?>
