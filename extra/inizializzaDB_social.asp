<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento senza titolo</title>
</head>
<%
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
Dim QuerySQL
%>
<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<body>
<%

   				   
	' in realtà non server perchè la stored procedure per calcolare è PS è cambiata e quindi InizializzaDB non serve per tutti i social ma ne basta 1 quella Id_Social=0			  
	
	QuerySQL="Select * from Allievi;"
	set rsTabella=ConnessioneDB.Execute(QuerySQL)
	
	 
	 i=0
	 Messaggio="InizializzaDB"
	 Id_Classe="6COM"
	 DataTest="12/12/2112"
	 
	 do while not rsTabella.eof 
		for j=1 to 2 
		Id_Social=j
	QuerySQL="INSERT INTO FORUM_MESSAGES (comments,CodiceAllievo,Id_Classe,Punti,DatePosted,Id_Social) SELECT '" &Messaggio & "','" & rsTabella("CodiceAllievo") & "','" & Id_Classe & "',0,'"&DataTest &"',"&Id_Social&";"
 '
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\log_ago.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
   			ConnessioneDB.Execute QuerySQL 
			response.write(QuerySQL)
    next 
 rsTabella.movenext()
loop

%>
</body>
</html>
