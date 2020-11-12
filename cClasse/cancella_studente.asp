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

location.href="quaderno.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>"
//location.href=window.history.back();
 }
 </script>
</head>


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
  
   CodiceAllievo=Request.QueryString("CodiceAllievo")
   
   'Apertura della connessione al database  
    
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")   
	 
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
     
        	  

<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
   
   
   
  
 
if  (Session("Admin")=true) then  %>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%

     QuerySQL ="DELETE  FROM Domande WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM Nodi WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM Frasi WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM M_Desideri WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM M_Navigazione WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM M_Topolino WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM Risultati WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM Risultati1 WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM FORUM_MESSAGES WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
 QuerySQL ="DELETE  FROM FILE_FORUM WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
     QuerySQL ="DELETE  FROM stud_as_classe WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM ALLIEVI WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	
	 


 response.Redirect "classifica.asp?id_classe="&Session("Id_Classe")&"&classe="&Session("Classe")
 
On Error Resume Next
If Err.Number = 0 Then

Response.Write "Cancellazione avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If





   %>
	</font>   
	 
 
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
		<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="home_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Apprendimento... </a></h3> 

		<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h3> 
						 
						</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	<%else%>
    
   <BODY onLoad="showText();">
	
	<%end if%>
	</body>
	</html>
	