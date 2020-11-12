<%
'	Author: Danilo Cicognani
'	Script: upload.asp
'	Version: 1.04
'	Data: 13/08/2009
'	Use: Popup window for the upload of files
'	Copyright (c) 2007-2009 Danilo Cicognani
'	https://www.ciconet.it
'	You can use, modify and redistribute this file, but you must mantain this copyright notice
'	This library uses ServerObjects ASPImage (https://www.serverobjects.com/products.htm) or Microsoft Office Web Components (https://www.microsoft.com/downloads/details.aspx?FamilyID=982b0359-0a86-4fb2-a7ee-5f3a499515dd) if you don't have these library installed you should user upload.aspx, provided that you have ASP.NET installed (see example2.asp for more information)
%>
<!--#include file="resize.asp" -->

<%
'Encode JavaScript strings
Function jsEncode(text)
	jsEncode = Replace(text, "\", "\\")
	jsEncode = Replace(jsEncode, "'", "\'")
End Function

'Convert a String to Binary
Function string2bin(String)
	Dim I, B
	For I=1 to len(String)
		B = B & ChrB(Asc(Mid(String,I,1)))
	Next
	string2bin = B
End Function

'Convert a Binary to String
Function bin2string(bin)
	For i = 1 To lenB(bin)
		bin2string  = bin2string  & chr(ascB(midB(bin, i, 1)))
	Next
End Function

' This function converts multibyte string to real binary data (VT_UI1 | VT_ARRAY)
Function bin2array(bin)
	Dim RS, LMultiByte, Binary
	Const adLongVarBinary = 205
	' Using recordset
	Set RS = CreateObject("ADODB.Recordset")
	LMultiByte = LenB(bin)
	if LMultiByte>0 then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
		RS("mBinary").AppendChunk bin & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	End If
	bin2array = Binary
End Function

Server.ScriptTimeout = 600	'10 minutes
'Read data
ReceivedBytes = Request.TotalBytes
If ReceivedBytes > 0 Then
	ChunkReadSize =  &H10000 '64 kB
	BytesRead = 0
	Set DataStream = createobject("ADODB.Stream")
	DataStream.Open
	DataStream.Type = 1 'Binary
	Do While BytesRead < ReceivedBytes
		'Read chunk of data
		PartSize = ChunkReadSize
		if PartSize + BytesRead > ReceivedBytes Then PartSize = ReceivedBytes - BytesRead
		DataPart = Request.BinaryRead(PartSize)
		BytesRead = BytesRead + PartSize
		DataStream.Write DataPart
	Loop
	DataStream.Position = 0
	ReceivedData = DataStream.Read
	Set DataStream = Nothing
	Upload = False
	Boundary = MidB(ReceivedData, 1, InstrB(ReceivedData, string2bin("" & vbCrLf)) - 1)
	pos = InstrB(ReceivedData, string2bin("" & vbCrLf)) + len("" & vbCrLf)

	lenCrLf = Len("" & vbCrLf)

	Do While pos < ReceivedBytes
		StartPos = InstrB(pos, ReceivedData, string2bin("" & vbCrLf & "" & vbCrLf))
		If StartPos - pos + lenCrLf + lenCrLf <= 0 Then
			Exit Do
		End If
		Header = bin2string(MidB(ReceivedData, pos, StartPos - pos + lenCrLf + lenCrLf))
		pos = InstrB(pos, ReceivedData, Boundary) - 1
		StartPos = StartPos + lenCrLf + lenCrLf
		FileContent = MidB(ReceivedData, StartPos, pos - StartPos - lenCrLf + 1)
		pos = pos + LenB(Boundary)

		' Get the fields if they are compiled
		if instr(Header, "field") > 0 then
			field = bin2string(FileContent)
		end if
		if instr(Header, "path") > 0 then
			path = bin2string(FileContent)
		end if
		if instr(Header, "prev") > 0 then
			prev = bin2string(FileContent)
		end if
		if instr(Header, "thumbField") > 0 then
			thumbField = bin2string(FileContent)
		end if
		if instr(Header, "thumbPath") > 0 then
			thumbPath = bin2string(FileContent)
		end if
		if instr(Header, "thumbWidth") > 0 then
			thumbWidth = bin2string(FileContent)
		end if
		if instr(Header, "thumbHeight") > 0 then
			thumbHeight = bin2string(FileContent)
		end if
		if instr(Header, "imgWidth") > 0 then
			imgWidth = bin2string(FileContent)
		end if
		if instr(Header, "imgHeight") > 0 then
			imgHeight = bin2string(FileContent)
		end if

		' Get the file to upload (if present) and write it to the server
		If Instr(Header, "upload") > 0 then
			i = Instr(Header, "filename=")
			j = Instr(i + 10, Header, chr(34))
			UploadName = mid(Header, i + 10, j - i - 10)
			i = instrRev(UploadName, "\")
			If i <> 0 then
				FileName = mid(UploadName, i + 1)
			Else
				FileName = UploadName
			End If
			Session("FileName")=FileName
			If FileName <> "" then
				Set binaryStream = createobject("ADODB.Stream")
				Set FSO = CreateObject("Scripting.FileSystemObject")
				Upload = true
				j = InStrRev(FileName, ".")
				FileNameWithoutExtension = left(FileName, j - 1)
				Extension = right(FileName, Len(FileName) - j + 1)
				toResize = False
				FileNameWithoutExtensionJpg = FileNameWithoutExtension
				If CInt(imgWidth) > 0 Then
					toResize = true
				End If
				If FileName <> prev Then
					j = 1
					Do While FSO.FileExists(path & "/" & FileName) Or (toResize And FSO.FileExists(path & "/" & FileNameWithoutExtensionJpg & ".jpg"))
						FileName = FileNameWithoutExtension & j & Extension
						FileNameWithoutExtensionJpg = FileNameWithoutExtension & j
						j = j + 1
					Loop
					If j > 1 Then
						FileNameWithoutExtension = FileNameWithoutExtension & (j - 1)
					End If
				End If
				binaryStream.Type = 1 'Binary
				binaryStream.Open
				if lenb(FileContent) > 0 then binaryStream.Write bin2array(FileContent)
				binaryStream.SaveToFile path & "/" & FileName, 2 'Overwrite
				If CInt(thumbWidth) > 0 Then
					'I must create the thumbnail
					ResizeImage path & "/" & FileName, thumbPath & "/" & FileNameWithoutExtension & ".jpg", "JPG", CInt(thumbWidth), CInt(thumbHeight)
					ThumbFile = FileNameWithoutExtension & ".jpg"
				End If
				If CInt(imgWidth) > 0 Then
					'I must resize the original image
					ResizeImage path & "/" & FileName, path & "/" & FileNameWithoutExtension & ".jpg", "JPG", CInt(imgWidth), CInt(imgHeight)
					If LCase(Extension) <> ".jpg" Then
						'Delete the original file
						FSO.DeleteFile path & "/" & FileName
					End If
					'FileName = FileNameWithoutExtension & ".jpg"
					FileName = FileName 
				End If
				Set binaryStream = Nothing
				Set FSO = Nothing
			End If
		End If
	Loop

	If Upload = True Then
	'session("uploaded")=true
	 
%>
<html>
<head>
	<title>Upload</title>
	<meta name="author" content="Danilo Cicognani" />
	<meta name="robots" content="noindex,nofollow"/>
	<script language="javascript">
		window.opener.document.<%= field %>.value = '<%= jsEncode(FileName) %>';
		<% If CInt(thumbWidth) > 0 Then %>
		window.opener.document.<%= thumbField %>.value = '<%= jsEncode(ThumbFile) %>';
		<% End If %>
		window.close();
	</script>
</head>
<body>
</body>
</html>
<%

 
	End If
Else
	field = Request("field")
	path = Request("path")
	prev = Request("prev")
	thumbField = Request("thumbField")
	thumbPath = Request("thumbPath")
	thumbWidth = Request("thumbWidth")
	thumbHeight = Request("thumbHeight")
	imgWidth = Request("imgWidth")
	imgHeight = Request("imgHeight")
End If

If Upload = False then
%>
<html>
<head>
	<title>Upload</title>
	<meta name="author" content="Danilo Cicognani" />
	<meta name="robots" content="noindex,nofollow"/>
    <link rel="stylesheet" type="text/css" href="../../stile.css">
    <script language="javascript" type="text/javascript" >
 function validate2() {
	var stringa=upload.upload.value;
	//alert(stringa); 
	if ((stringa.search(".jpg") == -1) && (stringa.search(".JPG") == -1) )
	{
	   alert("L'immagine deve essere in formato .jpg");
	 //  upload.upload.setfocus();
	   return 0;
	}
 	else
	{
	    document.upload.action = "upload.asp";
		document.upload.submit();		
   }
}


 </script>
</head>
<body>
<div class="contenuti_login">
	<h1>Upload</h1>
	<form name="upload"  action="upload.asp" method="post" enctype="multipart/form-data">
		<input type="hidden" name="field" value="<%= field %>"/>
		<input type="hidden" name="path" value="<%= path %>"/>
		<input type="hidden" name="prev" value="<%= prev %>"/>
		<input type="hidden" name="thumbField" value="<%= thumbField %>"/>
		<input type="hidden" name="thumbPath" value="<%= thumbPath %>"/>
		<input type="hidden" name="thumbWidth" value="<%= thumbWidth %>"/>
		<input type="hidden" name="thumbHeight" value="<%= thumbHeight %>"/>
		<input type="hidden" name="imgWidth" value="<%= imgWidth %>"/>
		<input type="hidden" name="imgHeight" value="<%= imgHeight %>"/>
		<input type="file" name="upload"/>	 
		<input type="submit" name="submit" value="Carica"/>
   <!-- <input type="button" name="submit2" value="Carica2" onClick="validate2();"/>-->
	</form>
 </div>
</body>
</html>
<%
End If
'Example 1: this is an example of a JavaScript function to open the upload window for the upload of one file.
'"form" is the form object
'"field" is the name of the field where the uploaded file name should be displayed
'"path" is the absolute path where the uploaded file should be saved
'"prev" is the previous name of the file if any: if there is already a file with the the same name of the uploaded one and is not the previous file name, the new file will be renamed adding a number
'	function uploadWindow(form, field, path, prev) {
'		var upload = window.open('upload.asp?field=' + form + '.' + field + '&path=' + path + (prev != '' ? '&prev=' + prev : ''), 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=350,height=200');
'		upload.focus();
'	}
'
'Example 2: this is an example of a JavaScript function to open the upload window for the upload and resize of one image
'"form" is the form object
'"imgField" is the name of the field where the uploaded file name should be displayed
'"thumbField" is the name of the field where the created thumbnail name should be displayed (can be "")
'"imgPath" is the absolute path where the uploaded file should be saved
'"thumbPath" is the absolute path where the thumbnail file should be saved
'"prev" is the previous name of the file if any: if there is already a file with the the same name 
'"imgWidth" is the requested width for the uploaded image, if the image is bigger it will be resized (if is 0 the image will not be resized)
'"imgHeight" is the requested height for the uploaded image, if the image is bigger it will be resized
'"thumbWidth" is the requested width for the thumbnail image, if the image is bigger it will be resized (if is 0 the thumbnail will not be created)
'"thumbHeight" is the requested height for the thumbnail image, if the image is bigger it will be resized
'	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
'		var upload = window.open('upload.asp?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=350,height=200');
'		upload.focus();
'	}
%>