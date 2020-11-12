<%
'	Author: Danilo Cicognani
'	Script: example2.asp
'	Version: 2.00
'	Date: 13/08/2009
'	Use: Sample script for upload.asp, resize.asp and upload.aspx: this sample uploads and resizes images, also creating thumbnails
'	Copyright (c) 2007-2009 Danilo Cicognani
'	https://www.ciconet.it
'	You can use, modify and redistribute this file, but you must mantain this copyright notice

imgPath = Server.MapPath(".") & "\img"	'I suppose your images will be saved in an "img" folder which is child of the current folder
imgPath=Replace(imgPath,"\","/")
thumbPath = Server.MapPath(".") & "\thumb"	'I suppose your images will be saved in a "thumb" folder which is child of the current folder
thumbPath=Replace(thumbPath,"\","/")
%>
<!-- #include File="resizecheck.asp" -->
<%
Dim pageUpload
pageUpload = "upload.asp"
If Not CheckResizeLib Then
	pageUpload = "upload.aspx"
End if
%>
<html>
<head>
<title>Sample script of upload.asp, resize.asp and upload.aspx: this sample upload and resize images, also creating thumbnails</title>
<script>
	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=500,height=200');
		upload.focus();
	}
</script>
</head>
<body>
<form name="theform" method="post" action="save.asp">
	<!-- I suppose you have a save.asp script that you will use to save all data in this form -->
	Title: <input name="title" /><br/>	<!-- Maybe you will need a title field? -->
	Image:&nbsp; <input readonly="readonly" name="img" /><input type="hidden" name="thumb" />
	<input type="button" value="Scegli..." onClick="uploadImgWindow(this.form.name, 'img', 'thumb', '<%= Server.URLEncode(imgPath) %>', '<%= Server.URLEncode(thumbPath) %>', this.form.img.value, 700, 700, 120, 92);" /><br/>
	<!-- I resize the image with the max with of 191 and the max height of 144.
	I create a thumbnail with the max with of 120 and the max height of 92.-->
	Text: <textarea name="text"></textarea><br/>	<!-- Maybe you will need a text field? -->
	<!-- Add any other field you will need... -->
	<input type="submit" value="Save" />
</form>
</body>