<?php
// $Id: form_latin1.tpl.php,v 1.2 2010/09/28 15:42:11 wundo Exp $

/**
 * @file form_latin1.tpl.php
 *
 * Theme implementation to display a Pay via PagSeguro Form, inside an iframe
 * Since PagSeguro uses the iso-8859-1 (latin1) charset, the form must be submitted
 *  using this charset, so the only reasonable solution is to render the form inside an iframe
 *
 * Available variables:
 * - $title: the page title
 * - $form: the rendered HTML form with pagseguro vars as hidden fields and a button to POST
 *
 * @see template_preprocess()
 * @see template_preprocess_node()
 */
?><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
	"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="pt-br" lang="pt-br">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
	<title><?php print $title; ?></title>	
</head>
<body><?php print $form; ?></body>
</html>
