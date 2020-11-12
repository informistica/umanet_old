<%
'	Author: Danilo Cicognani
'	Script: example1.asp
'	Version: 1.00
'	Date: 15/11/2007
'	Use: Sample script for upload.asp: this sample uploads a file
'	Copyright (c) 2007 Danilo Cicognani
'	https://www.ciconet.it
'	You can use, modify and redistribute this file, but you must mantain this copyright notice

filePath = Server.MapPath(".") & "\file"	'I suppose your files will be saved in a "file" folder which is child of
filePath=Replace(filePath,"\","/")
' the current folder
%>
<html>
<head>
<title>Sample script for upload.asp: this sample uploads a file</title>
<link rel="stylesheet" type="text/css" href="../../stile.css">

<script>
	function uploadWindow(form, field, path, prev) {
		var upload = window.open('upload.asp?field=' + form + '.' + field + '&path=' + path + (prev != '' ? '&prev=' + prev : ''), 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=500,height=200');
		upload.focus();
	}
</script>
</head>
<body>
<div class="contenuti_login">
<form name="theform" method="post" action="save.asp">
	<!-- I suppose you have a save.asp script that you will use to save all data in this form -->
	Title: <input name="title" /><br/>	<!-- Maybe you will need a title field? -->
	File:&nbsp; <input readonly="readonly" name="file" />
	<input type="button" value="Scegli..." onClick="uploadWindow(this.form.name, 'file', '<%= Server.URLEncode(filePath) %>', this.form.file.value);" /><br/>
	Text: <textarea name="text"></textarea><br/>	<!-- Maybe you will need a text field? -->
	<!-- Add any other field you will need... -->
	<input type="submit" value="Save" />
</form>
</div>
</body>