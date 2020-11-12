<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
<link rel="stylesheet" type="text/css" href="../../stile.css">
</head>

<body >
<div class="citazioni" ><div> <span style="font-style: normal">
<%
Dim Id_Eser
Dim ConnessioneDB,rsTabella,QuerySQL
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 
%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<% 


' inizio parte nuova
  AggiungiReport=Request("AggiungiReport") 
  Cancella=Request("Cancella") 
  id_classe=Request("id_classe")
  xQuiz=Request.QueryString("xQuiz")
  CodiceTest=Request.QueryString("CodiceTest")
  DataCla=Request.QueryString("DataCla")
 
  if  (Request.Form("txtData")="Data") and (xQuiz="") then
  ' if 3>4 then
          Response.Redirect "../cClasse/classifica.asp?id_classe="&id_classe
  else
  tipoTest=Request.QueryString("tipoTest")
  DataTest=Request.QueryString("DataTest")
  CodiceTest=Request.QueryString("CodiceTest")
  TitoloTest=Request.QueryString("TitoloTest")
  
' fine parametri nuovi
if (Cancella<>"") then
 ID_ESER=Request.QueryString("ID_ESER")
  QuerySQL ="Delete * from [2ESERCITAZIONI_SINGOLI] where ID_Esercitazione="&ID_ESER & ";"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)

else

   if AggiungiReport="" then ' aggiorna esistente

		Id_Eser=Request("Id_Eser1")
		numStud=Request("numStud")
		if numStud<>"" then
		   numStud=cint(Request("numStud"))  
		else
		 numStud=0
		end if
		txtData=Request.form("txtData")
		response.write(txtData & "=?" & DataTest)
		if (strcomp(txtData,DataTest)<>0) then
		' aggiorno la data
		    QuerySQL ="UPDATE [2ESERCITAZIONI_SINGOLI] SET Data = '" & txtData & "' WHERE ID_Esercitazione =" &Id_Eser &";"
			response.write(QuerySQL)
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)	
		end if
		for i=0 to numStud-1	 
			QuerySQL ="UPDATE [2CREDITI] SET Crediti = " & Request("Punti"&i) & " WHERE Id_Esercitazione =" &Id_Eser &" and Id_Stud='" & Request("CodiceAllievo"&i) & "';"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			response.write("<br>"&QuerySQL)
		next 
	else
	 
			' parte nuova
		'response.write("ciao2")	
			
			' inserisco l'esercitazione 
		'se l'esercitazione si riferisce alla convalida di un quiz mi vado a prendere, a partire dal codice del quiz, il titola del modulo da mettere nella campo descrizione, altrimenti prendo il valore del campo txtVerifica
			IF xQuiz<>"" then
				Titolo=Request.Querystring("TitoloTest")  
				Data=Request.Querystring("DataTest")  
				
			  else
				Titolo=Request("txtVerifica") 
				Data=Request("txtData")
			end if   
			Titolo = Replace(Titolo, Chr(34), "'")
			Titolo=  Replace(Titolo,"'",Chr(96))
			
			if  strcomp("on",Request.Form("cbScrutini"))=0 then
			  response.write("cbScrutini=on")
			  Scrutini=1
			else
			  response.write("cbScrutini<>on")
			  Scrutini=0
			end if
			Classifica=1
			
			TipoVoto=Request("txtTipoVoto")
			if  strcomp("on",Request.Form("cbClassifica"))=0 then' lo devo registrare solo per lo scrutnino per la media dello scrutinio ma solo per la classifica
			   response.write("cbClassifica=on")
			  Scrutini=1
			  Classifica=0
			end if
			
			QuerySQL="INSERT INTO [2ESERCITAZIONI_SINGOLI] (Descrizione,Data,Id_Classe,Scrutini,Classifica,TipoVoto) SELECT '" & Titolo  & "','" & Data & "','" & id_classe & "'," & Scrutini & "," & Classifica & ",'" & TipoVoto & "';"
			
		
		'	url="C:\Inetpub\wwwroot\anno_2012-2013_2\logCrediti.txt"
		'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
		'				objCreatedFile.WriteLine(QuerySQL)
		'				objCreatedFile.Close
		  response.write(QuerySQL)
		   ConnessioneDB.Execute QuerySQL 
		   
		   'prelevo il codice dell'esercitazione appena inserita
		   QuerySQL="SELECT MAX([ID_Esercitazione]) "&_
		   " FROM [2ESERCITAZIONI_SINGOLI];" 
			Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
			ID_ESER=rsTabella1(0) 
			IF xQuiz="" THEN ' se non sono stato chiamato dal bottone per convalidarer il quiz allora prendo i dati dal form altrimenti dalla query sui risultati del quiz
			   'per ogni studente inserisco il suo punteggio prelevato dal form
				i=0
			  QuerySQL="Select * from Allievi where Id_Classe='" & id_classe & "' and Attivo=1 order by Cognome asc;" 
			  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			  do while not rsTabella.eof 
				 allievo=rsTabella.fields("CodiceAllievo")
				 QuerySQL="INSERT INTO [2CREDITI] (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & ID_ESER & "','" & allievo & "','" & Request("Punti"&i) & "';"
				 ConnessioneDB.Execute(QuerySQL)
				 response.write("<br>"&QuerySQL)
				 i=i+1
				 rsTabella.movenext
			   loop
			   rsTabella.close
			ELSE
			
			   if tipoTest=0 then '
				  QuerySQL="SELECT Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Risultati.Risultato, Risultati.Ora, Risultati.CodiceTest,Risultati.Risultato*8/100 as [PUNTI] " &_
			" FROM Allievi INNER JOIN Risultati ON Allievi.CodiceAllievo = Risultati.CodiceAllievo " &_
			" WHERE  Risultati.Data=#" & Request.QueryString("DataTest")& "# AND Risultati.CodiceTest='"&  Request.QueryString("CodiceTest") & "'" &_
			" ORDER BY Risultati.Risultato DESC , Risultati.Ora; "
				else
				   QuerySQL="SELECT Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Risultati1.Risultato, Risultati1.Ora, Risultati1.CodiceTest,Risultati1.Risultato*8/100 as [PUNTI] " &_
			" FROM Allievi INNER JOIN Risultati1 ON Allievi.CodiceAllievo = Risultati1.CodiceAllievo " &_
			" WHERE  Risultati1.Data=#" & Request.QueryString("DataTest")& "# AND Risultati1.CodiceTest='"&  Request.QueryString("CodiceTest") & "'" &_
			" ORDER BY Risultati1.Risultato DESC , Risultati1.Ora; "
				end if
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			   ' prelevo i dati da inserire dalla query sui risultati
			   do while not rsTabella.eof 
				 allievo=rsTabella.fields("CodiceAllievo")
				 punti=round(fix(rsTabella.fields("PUNTI")))
				 QuerySQL="INSERT INTO 2CREDITI (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & ID_ESER & "','" & allievo & "','" & punti & "';"
				 response.write("<br>"&QuerySQL)
				 ConnessioneDB.Execute(QuerySQL)
				 rsTabella.movenext
			   loop
			   rsTabella.close
		END IF 
		 
		
		'Response.Redirect "studente_domande.asp?cla="&id_classe
		'end if 
		 
		  
			
			
			
			
			
			end if 
    end if
end if ' if (Cancella<>"") then

	ConnessioneDB.close
     Response.Redirect "../cClasse/classifica.asp?id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")
	
        '  if Request.ServerVariables("HTTP_REFERER") <>"" then 
		'					response.Redirect request.serverVariables("HTTP_REFERER") 
		' end if 

 %>
</body>
</html>
