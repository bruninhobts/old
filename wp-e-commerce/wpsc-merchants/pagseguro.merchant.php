<?php
$_GET["sessionid"] = $_GET["sessionid"]=="" ? $_SESSION["pagseguro_id"] : $_GET["sessionid"];
require_once ("PagSeguroLibrary/PagSeguroLibrary.php");

/**
	* WP eCommerce Test Merchant Gateway
	* This is the file for the test merchant gateway
	*
	* @package wp-e-comemrce
	* @since 3.7.6
	* @subpackage wpsc-merchants
*/
$nzshpcrt_gateways[$num] = array(
	'name' => 'pagseguro',
	'api_version' => 2.0,
	'class_name' => 'wpsc_merchant_pagseguro',
	'has_recurring_billing' => true,
	'display_name' => 'PagSeguro',	
	'wp_admin_cannot_cancel' => false,
	'submit_function' => 'submit_pagseguro',
	'requirements' => array(
		 /// so that you can restrict merchant modules to PHP 5, if you use PHP 5 features
		///'php_version' => 5.0,
	),	
	'form' => 'form_pagseguro',
    'supported_iso_codes' => array(
        'BR' => 'BRA'
    ),	
	// this may be legacy, not yet decided
	'internalname' => 'wpsc_merchant_pagseguro',
);
#'image' => WPSC_URL . '/images/pagseguro.gif',

class wpsc_merchant_pagseguro extends wpsc_merchant {
	
	var $name = 'PagSeguro';
	
	function submit() {
	    global $wpdb, $wpsc_cart, $wpsc_gateways;
        
        #echo '<pre>'; print_r($this->cart_data); echo '</pre>'; die('fim');
        
	    // Instantiate a new payment request
	    $paymentRequest = new PaymentRequest();	    
	    
	    // Sets the currency
	    $paymentRequest->setCurrency("BRL");	    
	    
	    $pagseguro_freight = $wpsc_cart->selected_shipping_method == $wpsc_gateways['wpsc_merchant_pagseguro']['name'];

	    $freight = 0;
	    
	    if ( $pagseguro_freight )
	    {
            $freight = sprintf('%01.2f', $this->cart_data['cart_tax'] + $this->cart_data['base_shipping']);
        }

        foreach($wpsc_cart->cart_items as $item) 
        {
            $paymentRequest->addItem(
                $item->product_id,
                $item->product_name, 
                $item->quantity, 
                $item->unit_price,
                intval(round($item->weight * 453.59237)),
                $freight
            );
        }
        
	    // Sets a reference code for this payment request, it's useful to identify this payment in future notifications.
	    $paymentRequest->setReference($this->cart_data['session_id']);        		
        
        $collected_data = array();

        $checkout_form_sql = "SELECT id, unique_name FROM `".WPSC_TABLE_CHECKOUT_FORMS."`";
        $checkout_form = $wpdb->get_results($checkout_form_sql, ARRAY_A) ;

        // Pega a referência dos campos de formulário definido pelo usuário                        
        foreach($checkout_form as $item) {
            $collected_data[$item['unique_name']] = $item['id'];
        }

        // Pega os dados do post
        $_client = $_POST["collected_data"];

        list($prefix, $phone)   = splitPhone($_client[if_isset($collected_data['billingphone'])]);

        $street = explode(',', $_client[if_isset($collected_data['billingaddress'])]);      

        $street = array_slice(array_merge($street, array("", "", "", "")), 0, 4); 
        
        list($address, $number, $complement, $neighborhood) = $street;    

        // Sets your customer information.
        $paymentRequest->setSender(
        
            $_client[if_isset($collected_data['billingfirstname'])] . ' ' . $_client[if_isset($collected_data['billinglastname'])],
        
            $_client[if_isset($collected_data['billingemail'])], $prefix, $phone
        
        );

#        $shipping  = get_option('pagseguro_shipping_configs');

        // Get the freight data
        if ($pagseguro_freight)
        {

            $freight_type = strtoupper($wpsc_cart->selected_shipping_option);
            
            $FREIGHT_CODE = 0;
            
            if ($freight_type=='PAC') {
	            
	            $FREIGHT_CODE = ShippingType::getCodeByType('PAC');
	            
	        }elseif($freight_type=='SEDEX') {
	
	            $FREIGHT_CODE = ShippingType::getCodeByType('SEDEX');
	            
	        }

	        if( $FREIGHT_CODE > 0 ) { 

	            $paymentRequest->setShippingType($FREIGHT_CODE);
	            
	            $paymentRequest->setShippingAddress(
	                $_client[if_isset($collected_data['shippingpostcode'])], 
	                $address, 
	                $number, 
	                $complement, 
	                $neighborhood,
	                $_client[if_isset(utf8_decode($collected_data['shippingcity']))], 
	                $_client[if_isset(utf8_decode($collected_data['shippingstate']))], 
	                $wpsc_gateways['wpsc_merchant_pagseguro']['supported_iso_codes'][$_client[if_isset($collected_data['shippingcountry'])]]
	            );
	            
	        }
        }
        $extra = 0;
        if ( $this->cart_data['has_discounts'] ) {
             $extra = $this->cart_data['cart_discount_value'] * (-1);
        }

        $paymentRequest->setExtraAmount( $extra );

        #$paymentRequest->setRedirectUrl($this->cart_data['transaction_results_url']);
        $paymentRequest->setRedirectUrl('http://homologacao.visie.com.br/bibliotecas/pagseguro/opencart1505/notification.php');
        
        // Pegando as configurações definidas no admin do módulo
	    $email = get_option('pagseguro_email');
	    $token = get_option("pagseguro_token");	    
        
	    /**
	     * Você pode utilizar o método getData para capturar as credenciais
	     * do usuário (email e token)
         * $email = PagSeguroConfig::getData('credentials', 'email');
         * $token = PagSeguroConfig::getData('credentials', 'token');
	     */
	    try {

		    /**
             * #### Crendenciais ##### 
             * Se desejar, utilize as credenciais pré-definidas no arquivo de configurações
             * $credentials = PagSeguroConfig::getAccountCredentials();
		     */		
	        $credentials = new AccountCredentials($email, $token);

		    if ($geteway_url = $paymentRequest->register($credentials)) {

		        $_SESSION["pagseguro_id"] = $sessionid;

		        $wpsc_cart->empty_cart();

		        wp_redirect($geteway_url);

		        exit();
		    }

	    } catch (PagSeguroServiceException $e) {

		    die($e->getMessage());

	    }

	}
}

/**
 * Separa o prefixo do telefone
 * @return array() - Array contendo o prefixo e o telefone
 */
function splitPhone($phone)
{

    $phone = preg_replace('/[a-w]+.*/', '', $phone);
    
    $numbers = preg_replace('/\D/', '', $phone);
    
    $telephone = substr($numbers, sizeof($numbers) - 9);
    
    $prefix = substr($numbers, sizeof($numbers) - 11, 2);
    
    return array($prefix, $telephone);
    
}
    
function if_isset(&$a, $b = '')
{
    
    return isset($a) ? $a : $b;

}

function submit_pagseguro() 
{
    if($_POST['pagseguro_email'] != null) {
        
        update_option('pagseguro_email', $_POST['pagseguro_email']);
    
    }
    
    if($_POST['pagseguro_token'] != null) {
     
        update_option('pagseguro_token', $_POST['pagseguro_token']);
    
    }
    return true;
}

/**
 * form_pagseguro
 *
 * Exibe o formulário de configuração do método de pagamento, dados do pagseguro
 * @return string Html do formulário
 *
 */
function form_pagseguro() 
{
    $output = "<tr>\n\r";
    $output .= "<tr>\n\r";
    $output .= "	<td colspan='2'>\n\r";
    $output .= "<strong>".TXT_WPSC_PAYMENT_INSTRUCTIONS_DESCR.":</strong><br />\n\r";
    $output .= "Email vendedor <input type=\"text\" name=\"pagseguro_email\" value=\"" . get_option('pagseguro_email') . "\"/><br/>\n\r";
    $output .= "TOKEN <input type=\"text\" name=\"pagseguro_token\" value=\"" . get_option('pagseguro_token') . "\"/><br/>\n\r";
    $output .= "<em>".TXT_WPSC_PAYMENT_INSTRUCTIONS_BELOW_DESCR."</em>\n\r";
    $output .= "	</td>\n\r";
    $output .= "</tr>\n\r";
    return $output;
}

/**
 * transact_url()
 *
 * Verifica o post do pagseguro e atualiza o pedido com o status da transação
 *
 */
function transact_url()
{
    if(!function_exists("automatic_return")) {

        function automatic_return ($transactionStatus, $reference)
        {
            global $wpdb;
            
            switch($transactionStatus) 
            {
                case "3":case "4":
                
                    $sql = "UPDATE `".WPSC_TABLE_PURCHASE_LOGS."` SET `processed`= '2' WHERE `sessionid`=" . $reference;
                
                    $wpdb->query($sql);
                
                default:
                
                    break;
            }
            
        }
        
        require_once("pagseguro/notification.php");
        
    }
    
}

/**
 * pagseguro_return()
 *
 * Sensível ao carregamento da pág. de retorno (transaction_results), executa o 
 * transact_url caso tenha recebido um post
 *
 */
function pagseguro_return() {

    if ($_SERVER['REQUEST_METHOD']=='POST' and $_POST) {

        if( get_option('transact_url')=="http://".$_SERVER["SERVER_NAME"].$_SERVER["REQUEST_URI"]){ transact_url();}

    }
}

add_action('init', 'pagseguro_return');


