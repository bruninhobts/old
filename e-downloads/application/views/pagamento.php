<?php
header('cache-control: NO-CACHE');
?><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr" lang="pt-BR">
<head>
    <title></title>
</head><body>
<?php echo $form_pgs; ?>
</form>
<div id="geral">Aguarde... redirecionando para o PagSeguro...</div>
<script language="javascript" type="text/javascript">
    document.forms[0].target='_parent';
    document.forms[0].submit();
</script>
</body>
</html>
