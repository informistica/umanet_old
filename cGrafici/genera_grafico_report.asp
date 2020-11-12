<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
</head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<body>

<% 

Id_Eser=request.querystring("ID_ESER")


  Dim ConnessioneDB,ConnessioneDB1, rsTabella,rsTabella1, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione
   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 
	%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   </body>
<center>
<p>
 
</p>
<div class="citazioni" ><div> <span style="font-style: normal">

<b><font size="3">GESTIONE</font>&nbsp;</b> </span></div>
<hr>
<%
'                            		
 QuerySQL="SELECT * FROM [2REPORT_CREDITI] Where ID_Esercitazione="& Cint(Id_Eser) &" order by Crediti desc "
	response.write(QuerySQL)

 Set rsTabella = ConnessioneDB.Execute(QuerySQL)

Dim FileObject,riga
Set FileObject=CreateObject("Scripting.FileSystemObject")
esiste=FileObject.FileExists("data.xml")
'Response.Write("<p>il file esiste? "&esiste&"</p>")
Set FileObject=Nothing

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

 
url=Server.MapPath(homesito & "/script/cGrafici/Grafici")& "/data.xml"   
		url=Replace(url,"\","/")
Set objCreatedFile = objFSO.CreateTextFile(url, True)
response.write(url)
' Write a line with a newline character.
gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())

riga="<chart caption='PUNTEGGI DEGLI STUDENTI' subcaption='" & gira_data &"' xAxisName='Studente' yAxisName='Punti' numberPrefix='P.'>"
objCreatedFile.WriteLine(riga)
 do while not rsTabella.EOF 
    cognome=rsTabella.fields("Cognome")
    cognome=Replace(cognome," ","")
    

    riga="<set label='" &cognome &"' value='" &rsTabella.fields("Crediti") & "'/>"
	objCreatedFile.WriteLine(riga)
    rsTabella.movenext
   loop

riga="</chart>"
objCreatedFile.WriteLine(riga)

' objCreatedFile.WriteLine(rsTabella(0))

 



'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
objCreatedFile.Close

On Error Resume Next
If Err.Number = 0 Then

Response.redirect "Grafici/genera.html"
Else
Response.Write Err.Description 
Err.Number = 0
End If

%>
<h4 style="text-align: center"><i><a href="../../home.asp" >Vai all'HomePage</a> </h4>
</center>
</html>