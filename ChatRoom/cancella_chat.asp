<%@ Language=VBScript %>

  <% Response.Buffer=True 
   Dim ConnessioneDB,  rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
    
   
   'Apertura della connessione al database  
   ' Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	ID_Chat=request.querystring("ID_Chat")
	nome=request.querystring("nome")
	 
	 
       %> 
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	 
    <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi modificare i dati degli altri studenti!")

location.href="ShowChat.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID_Chat=<%=ID_Chat%>"
//location.href=window.history.back();
 }
 </script>
</head>


   
 <%  Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%>
	<!--#include file = "../service/controllo_sessione.asp"-->
     <!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
     <!-- #include file = "../var_globali.inc" -->

<%  
                           
   
 
   

 
 
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
<body>
    <div id="container">
<div class="contenuti_forum">
	<font color=#FF0000 size="4">
 
<%

 
		 
QuerySQL ="delete * from CHAT_SESSION WHERE ID_Chat =" &ID_Chat&";"
				 ConnessioneDB.Execute(QuerySQL)	 
		        dim objFSO,objCreatedFile
				Const ForReading = 1, ForWriting = 2, ForAppending = 8
				Dim sRead, sReadLine, sReadAll, objTextFile
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				'url="C:\Inetpub\umanetroot\anno_2012-2013\logAggiorna.txt"
				url1=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Session("cartella") & "/Chatlog/" & nome 
				url1=Replace(url1,"\","/")
				objFSO.deletefile url1
				 
'				
			' response.write(QuerySQL &"<br>")
			 ' conn.Execute(QuerySQL)
		 
	 

'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA

 
'response.write(url)
 
On Error Resume Next
If Err.Number = 0 Then
		'Response.Write "Aggiornamento avvenuta! "
		%>
		<script>
location.href="ShowChat.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID_Chat=<%=ID_Chat%>&cartella=<%=session("cartella")%>"

		</script>
		<%'Response.Redirect "showChat.asp"
Else
		Response.Write Err.Description 
		Err.Number = 0
End If 
%>
	<center><br><br><font size="3">
<!--#include file = "footer.inc"-->
</center>
<!--#include file = "database_cleanup.inc"-->
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	<%else%>
    
   <BODY onLoad="showText();">
	
	<%end if%>
	</body>
	</html>
	