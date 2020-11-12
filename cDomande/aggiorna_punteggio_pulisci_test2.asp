<!-- calcola_risultato_MODBC3.asp -->
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
</head>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
   <% Response.Buffer=True 
  ' on error resume next
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   
   
 
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	 
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   
    <!-- #include file = "../var_globali.inc" -->
<%  
                     'Lettura dei dati memorizzati nei cookie. 
   
 
  classe=Request.QueryString("classe")
  id_classe=Request.QueryString("id_classe")
  xQuiz=Request.QueryString("xQuiz")
  CodiceTest=Request.QueryString("CodiceTest")
   SessioneQuiz=Request.QueryString("SessioneQuiz")
  
  tipoTest=clng(Request.QueryString("tipoTest"))
  
  if tipoTest=0 then
     tipoDesc="(V/F)"
  else if tipoTest=1 then 
          tipoDesc="(Singola)"
	    else
		   tipoDesc="(Multipla)"
		end if
  end if		   	  
  DataTest=Request.QueryString("DataTest")
  CodiceTest=Request.QueryString("CodiceTest")
  numRec=Request.Form("txtnumRec")
  
' inserisco l'esercitazione 
'se l'esercitazione si riferisce alla convalida di un quiz mi vado a prendere, a partire dal codice del quiz, il titola del modulo da mettere nella campo descrizione, altrimenti prendo il valore del campo txtVerifica
 
	Titolo="Quiz : " & Request.Querystring("TitoloTest")    
    Titolo = Replace(Titolo, Chr(34), Chr(96))
    Titolo=Titolo&tipoDesc
  if  strcomp("on",Request.Form("cbScrutini"))=0 then ' se non lo devo registrare per la media dello scrutinio ma solo per la classifica
    Scrutini=1
  else
    Scrutini=0
  end if
  response.write("VF="&Request.Form("VF"))
  TipoVoto=Request.Form("txtTipoVoto")
  Classifica=1
  QuerySQL="INSERT INTO [2ESERCITAZIONI_SINGOLI] (Descrizione,Data,Id_Classe,Scrutini,Classifica,TipoVoto) SELECT '" & Titolo  & "','" & DataTest & "','" & id_classe & "'," & Scrutini & "," & Classifica & ",'" & TipoVoto & "';"
	 response.Write(QuerySQL)
	ConnessioneDB.Execute QuerySQL 
	
' convalido la sessione
   QuerySQL ="UPDATE [2SESSIONI_QUIZ] SET Convalidata = 1  WHERE ID_Sessione =" &SessioneQuiz&";"
	ConnessioneDB.Execute QuerySQL 	 
   response.Write(QuerySQL)
  
    
'	url="C:\Inetpub\wwwroot\anno_2012-2013_2\logCrediti.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
   'ConnessioneDB.Execute QuerySQL 
   
     QuerySQL="SELECT MAX([ID_Esercitazione]) "&_
   " FROM [2ESERCITAZIONI_SINGOLI];" 
	Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
	ID_ESER=rsTabella1(0) 
	response.Write("<br>"&QuerySQL)

 
for i=1 to numRec
     allievo=Request.form("txtStud"&i)
	 punti=round(fix(Request.form("txtPunti"&i))/10)
     QuerySQL="INSERT INTO [2CREDITI] (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & ID_ESER & "','" & allievo & "','" & punti & "';"
	 ConnessioneDB.Execute(QuerySQL)
	 response.Write("<br>"&QuerySQL)
next
 'response.write("NumRec="&numRec)
 ConnessioneDB.close()
'Response.Redirect "studente_domande.asp?cla="&cla
'end if 
  %>
  
	</font>   
	 
		      <h4><a href="../cClasse/classifica.asp?Id_Classe=<%=Session("Id_Classe")%>">Aggiornamento avvenuto ... torna alla classifica ...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
		 
			 
			
			 
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	