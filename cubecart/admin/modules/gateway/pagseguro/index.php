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
|	index.php
|   ========================================
|	Configure PagSeguro
+--------------------------------------------------------------------------
*/
include("../../../../includes/ini.inc.php");
include("../../../../includes/global.inc.php");
require_once("../../../../classes/db.inc.php");
$db = new db();
include_once("../../../../includes/functions.inc.php");
$config = fetchDbConfig("config");

include_once("../../../../language/".$config['defaultLang']."/lang.inc.php");
$enableSSl = 1;
include_once("../../../../includes/sslSwitch.inc.php");
include("../../../includes/auth.inc.php");
include("../../../includes/header.inc.php");

if(permission("settings","read")==FALSE){
	header("Location: ".$GLOBALS['rootRel']."admin/401.php");
	exit;
}

if(isset($_POST['module'])){
	include("../../status.php");
	include("../../../includes/functions.inc.php");
	$module = fetchDbConfig($_GET['folder']);
	$msg = writeDbConf($_POST['module'], $_GET['folder'], $module);
	
}
$module = fetchDbConfig($_GET['folder']);
?>

<p><a href="https://pagseguro.uol.com.br/" title="PagSeguro"><img src="logo.gif" alt="PagSeguro" border="0" /></a></p>
<?php 
if(isset($msg)){ 
	echo stripslashes($msg); 
} 
?>
<p class="copyText">Este é o gateway de pagamento PagSeguro para CubeCart. </p>

<form action="<?php echo $GLOBALS['rootRel'];?>admin/modules/<?php echo $_GET['module']; ?>/<?php echo $_GET['folder']; ?>/index.php?module=<?php echo $_GET['module']; ?>&amp;folder=<?php echo $_GET['folder']; ?>" method="post" enctype="multipart/form-data">
<table border="0" cellspacing="0" cellpadding="3" class="mainTable">
  <tr>
    <td colspan="2" class="tdTitle">Configurações </td>
  </tr>
  <tr>
    <td align="left" class="tdText"><label for="moduleStatus"><strong>Status:</strong></label></td>
    <td class="tdText">
	<select name="module[status]" id="moduleStatus">
		<option value="1" <?php if($module['status']==1) echo "selected='selected'"; ?>>Ativado</option>
		<option value="0" <?php if($module['status']==0) echo "selected='selected'"; ?>>Desativado</option>
    </select>	</td>
  </tr>
   <tr>
  	<td align="left" class="tdText"><label for="moduleDesc"><strong>Descrição:</strong></label>	</td>
    <td class="tdText"><input type="text" name="module[desc]" id="moduleDesc" value="<?php echo $module['desc']; ?>" class="textbox" size="30" /></td>
  </tr>
  <tr>
  <td align="left" class="tdText"><label for="moduleEmail"><strong>Seu email cadastrado no <a href="https://pagseguro.uol.com.br">PagSeguro:</a></strong></label></td>
    <td class="tdText"><input type="text" name="module[email]" id="moduleEmail" value="<?php echo $module['email']; ?>" class="textbox" size="30" /></td>
  </tr>
  <tr>
  <td align="left" class="tdText"><label for="moduleToken"><strong>Seu TOKEN gerado pelo <a href="https://pagseguro.uol.com.br">PagSeguro:</a></strong></label></td>
    <td class="tdText"><input type="text" name="module[token]" id="moduleToken" value="<?php echo $module['token']; ?>" class="textbox" size="30" /></td>
  </tr>
    <tr>
    <td align="right" class="tdText">&nbsp;</td>
    <td class="tdText"><input type="submit" class="submit" value="Edit Config" /></td>
  </tr>
</table>
</form>
<?php include("../../../includes/footer.inc.php"); ?>
