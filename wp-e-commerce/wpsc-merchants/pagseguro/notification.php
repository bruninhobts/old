<?php
 
// Library PagSeguro
require_once ('../PagSeguroLibrary/PagSeguroLibrary.php');

class NotificationListener extends Controller {

    public static function main() {
        
    	$code = self::verifyData($_POST['notificationCode']);
    	$type = self::verifyData($_POST['notificationtype']);
    	
    	if ( $code && $type ) {
			
    		$notificationType = new NotificationType($type);
    		$strType = $notificationType->getTypeFromValue();

			switch($strType) {
				
				case 'TRANSACTION':
					self::TransactionNotification($code);
					break;
				
				default:
					LogPagSeguro::error("Tipo de notificação não reconhecido [".$notificationType->getValue()."]");
					
			}

			self::saveLog($strType);
			
		} else {
			
			LogPagSeguro::error("Os parâmetros de notificação (notificationCode e notificationType) não foram recebidos.");
			
			self::saveLog();
			
		}
		
    }
    
    /**
     * TransactionNotification Envia as credenciais do usuário para a Api e
     * retorna os dados da transação a partir do notificationCode
     * @notificationCode string Identificador único da transação
     */
    private static function TransactionNotification($notificationCode) {
		
    	/*
    	* #### Crendenciais #####
    	* Se desejar, utilize as credenciais pré-definidas no arquivo de configurações
    	* $credentials = PagSeguroConfig::getAccountCredentials();
    	*/
    	
        // Pegando as configurações definidas no admin do módulo
        $config = self::getConfig();
        
    	$credentials = new AccountCredentials($config['email'], $config['token']);
    	
    	try {
    		
    		$transaction = NotificationService::checkTransaction($credentials, $notificationCode);
    		
    		self::validateTransaction($transaction);
    		
    	} catch (PagSeguroServiceException $e) {

    		die($e->getMessage());

    	}
    	
    }
  
    private static function validateTransaction(Transaction $transaction) {
        global $wpdb;

        $transactionStatus = $transaction->getStatus();
        
        $reference = $transaction->getReference();
        
        if ($reference && function_exists( 'automatic_return' )) 
        {
            automatic_return( $transactionStatus, $reference );    
        }
            
    }
  
    /**
     * verifyData - Corrige os dados enviados via post
     * @data string Dados enviados via post
     */
    private static function verifyData($data){
    
        return isset($data) && trim($data) !== "" ? trim($data) : null;
    
    }  
    
    /**
     * getConfig - Retorna as configurações definidas para as credenciais
     * @return array Array contendo as credenciais do usuário
     */
    private static function getConfig() {
        global $db;
        
        $config = array();
                
        // Settings
        $query = $db->query("SELECT value FROM " . DB_PREFIX . "setting s where s.key='pagseguro_mail'");
        $config['email'] = $query->row['value'];

        $query = $db->query("SELECT value FROM " . DB_PREFIX . "setting s where s.key='pagseguro_token'");
        $config['token'] = $query->row['value'];
    
        return $config;
    
    }
    
    private static function saveLog($strType = null) {
        #LogPagSeguro::getHtml();
    }
	
}
NotificationListener::main();
?>
