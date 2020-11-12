<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
</head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<body>

<% 
'on error resume next
  CodiceAllievo=request.querystring("CodiceAllievo")
  Dim ConnessioneDB,rsTabella
  Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
    
	 
	
	
   </body>
<center>
<p>
 
</p>
<div class="citazioni" ><div> <span style="font-style: normal">

<b><font size="3">GESTIONE</font>&nbsp;</b> </span></div>
<hr>
<%
 dim objFSO,objCreatedFile
 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Dim sRead, sReadLine, sReadAll, objTextFile
 

 
	
	 QuerySQL="select * from 4PERIODI_CLASSIFICA where CodiceAllievo='" &CodiceAllievo &"';"
	 
	
				'response.write("Valore=" & cint(Request.Form("cb"&i)))  
			'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logAAAA.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine("Valore=" & cint(Request.Form("cbPS")))
'				objCreatedFile.Close
'	
 
	
	 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
   
        	

				
		Dim FileObject,riga
	'	Set FileObject=CreateObject("Scripting.FileSystemObject")
	'	esiste=FileObject.FileExists("Data.xml")
	'	Response.Write("<p>il file esiste? "&esiste&"</p>")
		'Set FileObject=Nothing
		
		
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
	   
		url=Server.MapPath(homesito & "/Grafici")& "/data_"&CodiceAllievo &".xml"   
		Set objCreatedFile = objFSO.CreateTextFile(url, True)
		' Write a line with a newline character.
		gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
		
		riga="<chart caption='VOTI NEI VARI PERIODI' subcaption='" & CodiceAllievo &"' xAxisName='Studente' yAxisName='Punti' numberPrefix='P.'>"
		objCreatedFile.WriteLine(riga)
		 do while not rsTabella.EOF 
			riga="<set label='" &rsTabella("Dal") &"-" &rsTabella("Al") &"' value='" &rsTabella.fields("Vv") & "'/>"
			objCreatedFile.WriteLine(riga)
			rsTabella.movenext
		   loop
		
		riga="</chart>"
		objCreatedFile.WriteLine(riga)
		
		
		' objCreatedFile.WriteLine(rsTabella(0))
		
		
		'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
		objCreatedFile.Close
		
		rsTabella.close()
		
		
		If Err.Number = 0 Then
			Response.redirect "../Grafici/genera.asp?CodiceAllievo="&CodiceAllievo
		Else
			Response.Write Err.Description 
			Err.Number = 0
		End If
 
 
	 
	
%>
<h4 style="text-align: center"><i><a href="../../home.asp" >Vai all'HomePage</a> </h4>
</center>
</html>