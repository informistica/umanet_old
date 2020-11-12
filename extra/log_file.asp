<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<%
				dim objFSO,objCreatedFile
				Const ForReading = 1, ForWriting = 2, ForAppending = 8
				Dim sRead, sReadLine, sReadAll, objTextFile
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
					url="C:\Inetpub\umanetroot\expo2015Server\log.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
				
				
				
Response.Redirect("inserisci_test.asp?Cartella=Cartella&Modulo=Modulo") 

%>
<body>
</body>
</html>
