<!-- calcola_risultato_MODBC3.asp -->
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../stile.css">
	 
   
  
   
</head>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
      
<% Response.Buffer=True %>


 
<%   Dim ConnessioneDB, rsTabella, QuerySQL
 
   
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	 
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   
     <!-- #include file = "../var_globali.inc" -->
     <%  
 
 
 
Sessione=Request.Form("txtSessione")
Tipo= Request.Form("txtTipo")
Data=Request.Form("txtData")
 
'Id_Mod=Request.Form("txtId_Mod")
'Id_Arg=Request.Form("txtId_Arg")

Id_Arg=Session("ID_ParSel") 
Id_Mod=Session("IdMod")


'Id_SottoArg="Provasottoarg"
Convalidata=0
 
   QuerySQL="INSERT INTO [2SESSIONI_QUIZ] (Titolo,Data,Id_Classe,TipoQuiz,Id_Mod,Id_Arg,Convalidata)  SELECT '" & Sessione & "','" & Data   & "', '" & Session("Id_Classe") &"'," &Tipo & ", '" & Id_Mod & "', '" & Id_Arg & "', " & Convalidata & ";"   
 
 response.write(QuerySQL)
   ConnessioneDB.Execute (QuerySQL) 
   

	'On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento dell'avviso avvenuto! "
		Session("IdxSel")=""
		Session("IdxSelPar")=""
		Session("PosPar")=""
		Session("ID_ParSel")="" 
		Session("Id_Mod")=""
	Else
		Response.Write Err.Description 
		Err.Number = 0
	End If
    
	
	 
	
	 
	
	
	
	 
	 
	  if Request.ServerVariables("HTTPS_REFERER") <>"" then 
									response.Redirect request.serverVariables("HTTP_REFERER") 
								end if 

   %>
	 


 
 
    
 
		  
          
<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../cClasse/home_app.asp?id_classe=<%=Id_Classe%>&divid=<%=divid%>"> Torna alla pagina Apprendimento... </a></h3> 
	
  
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	