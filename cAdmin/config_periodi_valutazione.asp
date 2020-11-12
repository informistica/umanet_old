<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
Dim ConnessioneDB, rsTabella, QuerySQL,Privato,Valutato,Classe   
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
  Id_Classe = Request.QueryString("Id_Classe")
  Classe=Request.Form("TxtClasse")
  DataCla=Request.Form("TxtDataCla") ' periodo precedente
  DataCla2=Request.Form("TxtDataVal") ' nuovo periodo
 
  'cancella=Request.QueryString("cancella")
  'idperiodo=Request.QueryString("idperiodo")
 ' data=Request.Form("data")
 
  
  ' dim objFSO,objCreatedFile
'   Const ForReading = 1, ForWriting = 2, ForAppending = 8
' 	Dim sRead, sReadLine, sReadAll, objTextFile
' 	Set objFSO = CreateObject("Scripting.FileSystemObject")  
'	url="C:\Inetpub\umanetroot\anno_2012-2013\logsetting1.txt"
'	Set objCreatedFile = objFSO.CreateTextFile(url, True)
'	objCreatedFile.WriteLine(QuerySQL)
'	objCreatedFile.Close
	
	if DataCla2<>"" then ' se sono chiamata per creare un nuovo periodo
	 
		     QuerySQL="insert into [dbo].[3PERIODI] (ID_Classe,Data) select '" & Id_Classe &"','" & DataCla2 & "';"
			 ConnessioneDB.Execute QuerySQL
			 '  response.Write(QuerySQL)
	        ' response.Redirect "../genera_grafico.asp?Id_Classe="&Id_Classe&"&DataCla2="&DataCla2&"&DataCla="&DataCla&"&PS=1"	 
           
	else ' aggiorno il perido iniziale
	
	      QuerySQL="SELECT * FROM [dbo].[3PERIODI] Where ID_Classe='"& ID_Classe &"' order by Data;"
	 
	 
	    Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	    i=1
		do while not rsTabella.eof
		 ' response.write(cint(Request.form("CheckPeriodo")) &"---"& i &"<br>")
		  if cint(Request.form("CheckPeriodo"))=i then
		    setIniziale=1
		  else
		    setIniziale=0
		  end if
		  
		  
		' QuerySQL="UPDATE 3PERIODI SET Iniziale = " & cint(Request.form("txtDataInizioCla"&i)) &_
		   QuerySQL="UPDATE [dbo].[3PERIODI] SET Iniziale = " & setIniziale &_
		 " WHERE ID_Periodo="&rsTabella("ID_Periodo") &";"
		  ConnessioneDB.Execute QuerySQL
		  response.write(QuerySQL&"<br>")
		i=i+1
	    rsTabella.movenext()
		loop  
	
	end if  
	 if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
		 end if 
	
 
%>
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Periodi valutazione</title>
</head>

<body>
</body>
</html>
