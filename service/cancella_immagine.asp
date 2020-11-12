<%@ Language=VBScript %>
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	<style>
	<!--
	 li.MsoNormal
		{mso-style-parent:"";
		margin-bottom:.0001pt;
		font-size:12.0pt;
		font-family:"Times New Roman";
		margin-left:0cm; margin-right:0cm; margin-top:0cm}
	-->
	</style>
    <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi cancellare i dati degli altri studenti!")

location.href="studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>"
//location.href=window.history.back();
 }
 </script>
</head>


   <% Response.Buffer=True 
   Dim ConnessioneDB,  rsTabella, QuerySQL,urlimg
   
   urlProvenienza = Request.ServerVariables("HTTP_REFERER")
   ID = Request.QueryString("CodiceFrase")
   
   
    urlimg=request.querystring("urlimg")
	urldb=request.querystring("urldb")
	CodiceAllievo=request.querystring("CodiceAllievo")
	by_Domande= request.querystring("by_Domande")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
   
 
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%


if by_Domande<>"" then ' se devo eliminare un immagine di domanda
    QuerySQL ="DELETE   FROM Domande_Img WHERE Url ='" &urldb&"';"
else
     QuerySQL ="DELETE   FROM Frasi_Img WHERE Url ='" &urldb&"' AND Id_Frase='"&ID&"';"
end if

 ' Set objFSO = CreateObject("Scripting.FileSystemObject")  
'  	url2="C:\Inetpub\umanetroot\anno_2012-2013\logCancellaDoma.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
'				objCreatedFile.WriteLine(QuerySQL & "---" & urlimg)
'				objCreatedFile.Close 

	 ConnessioneDB.Execute(QuerySQL)
'response.write(QuerySQL)

'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'response.write(urldb)
'objFSO.DeleteFile urlimg
'response.write("<br>Immagine="&urlimg)
On Error Resume Next
If Err.Number = 0 Then

 ' if Request.ServerVariables("HTTP_REFERER") <>"" then 
		' response.Redirect request.serverVariables("HTTP_REFERER") 
 ' end if 
Else
'Response.Write Err.Description 
Err.Number = 0
End If





   %>
	</font>   
						 
						</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	<%else%>
    
   <BODY onLoad="showText();">
	
	<%end if%>
	
	<script>alert('Cancellazione effettuata correttamente'); window.location.href = "<%=urlProvenienza%>"</script>

	</body>
	</html>
	