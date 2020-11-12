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
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo ,xQuiz ' vale 1 se sono stato chiamato dal bottone per la convalida del quiz, in tal caso inserisco il quiz nelle attività e prendo i dati non dal request.form ma dalla query sui risultati del quiz
   dim objFSO,objCreatedFile

   Dim periodi() ' vettore delle date per il calcolo della classifica, più avanti farò il redim
   Dim vetstud(35) ' massimo numero di studenti possibile
   vetstud(0)="?"

 Const ForReading = 1, ForWriting = 2, ForAppending = 8
 Dim sRead, sReadLine, sReadAll, objTextFile
 Set objFSO = CreateObject("Scripting.FileSystemObject")  
 
   
   
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	 Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
    <!-- #include file = "../var_globali.inc" -->
<%  
                     'Lettura dei dati memorizzati nei cookie. 
   
'DataCla=request.form("txtData") 
'DataCla2=request.form("txtData2")
DataCla=request.Querystring("DataCla") 
DataCla2=request.Querystring("DataCla2")
Session("DataCla")=DataCla
Session("DataCla2")=DataCla2 ' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
DataClaq=request.QueryString("DataClaq") 
DataClaq2=request.QueryString("DataClaq2")
Session("DataClaq")=DataClaq
Session("DataClaq2")=DataClaq2
 on error resume next
  classe=Request.QueryString("classe")
  id_classe=Request.QueryString("id_classe")
  xQuiz=Request.QueryString("xQuiz")
  CodiceTest=Request.QueryString("CodiceTest")
  DataCla=Request.QueryString("DataCla")
  if  (Request.Form("txtData")="Data") and (xQuiz="") then
  ' if 3>4 then
          Response.Redirect "studente_domande.asp?id_classe="&id_classe
  else
  tipoTest=Request.QueryString("tipoTest")
  DataTest=Request.QueryString("DataTest")
  CodiceTest=Request.QueryString("CodiceTest")
  TitoloTest=Request.QueryString("TitoloTest")
	
		' verifico se esitono già le tabelle ed in caso le elimino
	QuerySQL="select count(*) FROM T_PUNTEGGI_STUDENTI_DOMANDE;"
    Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
'   ' carico il vettore delle date di valutazione
	if err.number = 0 then
'	' non c'è errore quindi la tabella esiste quindi le cancello tutte
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_DOMANDE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_FRASI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_NODI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MTOPOLINO")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MNAVIGAZIONE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MDESIDERI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_METAFORE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_CREDITI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_FORUM")
   end if

' INIZIO FILE DI INCLUSIONE COMUNE

''ricreo le tabelle che servono per calcolare il punteggio
QuerySQL="SELECT Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe, Sum(Domande.Voto) AS PD INTO TPUNTEGGI_STUDENTI_DOMANDE " &_
" FROM Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud" &_
" WHERE Allievi.Id_Classe='"& id_classe& "'" &_
" AND(Domande.Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"# " &_
" AND Domande.Data<=#" & mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4)  &"#" &_
" OR Domande.Data=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_
" GROUP BY Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe;"
				
			
 
  
  
   	'url="C:\Inetpub\umanetroot\Anno_2010-2011\logDomande.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close	
' 
 
 
 
 
 
   Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	'Creo TFrasi
	  DataCla2=left(periodi(indice_periodo2),10)
		QuerySQL="SELECT DISTINCTROW Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome,  Sum(Frasi.Voto) AS PUNTI, Allievi.Id_Classe INTO TPUNTEGGI_STUDENTI_FRASI " &_
	" FROM Allievi LEFT JOIN Frasi ON Allievi.CodiceAllievo = Frasi.Id_Stud " &_
	" WHERE Allievi.Id_Classe='"& id_classe& "'" &_
	" AND  (Frasi.Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"# " &_
	" AND Frasi.Data<=#" & mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4)  &"#" &_
	" OR Frasi.Data=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_

	" GROUP BY Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe;"
 
	
	
		'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logFrasi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close	
'    
'	
	
	
	
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	'Creo TNodi
	  'DataCla2=left(periodi(indice_periodo2),10)
	QuerySQL="SELECT DISTINCTROW Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Sum(Nodi.Voto) AS PUNTI, Allievi.Id_Classe INTO TPUNTEGGI_STUDENTI_NODI " &_
" FROM Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud " &_
" WHERE Allievi.Id_Classe='"& id_classe& "'" &_
" AND (Nodi.Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
" AND Nodi.Data<=#" & mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4)  &"#" &_
" OR Nodi.Data=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_

" GROUP BY Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe;"
  
'		url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logNodi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close	



      Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	
	
	
	
	
	'Creo TMetafore
	
	' Creo Topolino
 	 'DataCla2=left(periodi(indice_periodo2),10)
	QuerySQL="SELECT DISTINCTROW Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Sum(M_Topolino.Voto) AS PUNTI, Allievi.Id_Classe INTO TPUNTEGGI_STUDENTI_MTOPOLINO " &_
" FROM Allievi INNER JOIN M_Topolino ON Allievi.CodiceAllievo=M_Topolino.Id_Stud " &_
" WHERE Allievi.Id_Classe='"& id_classe& "'" &_
" AND (M_Topolino.Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"# " &_
" AND  M_Topolino.Data<=#" & mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4)  &"#" &_
" OR M_Topolino.Data=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_
" GROUP BY Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe;"
   
'	
'		url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logMetafore.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close	
'    
	
	
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	' Creo Navigazione
 	' DataCla2=left(periodi(indice_periodo2),10)
	QuerySQL="SELECT DISTINCTROW Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Sum(M_Navigazione.Voto) AS PUNTI, Allievi.Id_Classe INTO TPUNTEGGI_STUDENTI_MNAVIGAZIONE " &_
" FROM Allievi INNER JOIN M_Navigazione ON Allievi.CodiceAllievo=M_Navigazione.Id_Stud " &_
" WHERE Allievi.Id_Classe='"& id_classe& "'" &_
" AND (M_Navigazione.Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"# " &_
" AND M_Navigazione.Data<=#" & mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4)  &"#" &_
" OR M_Navigazione.Data=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_
" GROUP BY Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe;"
  
	
		
    
	
	
	
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
  QuerySQL="SELECT DISTINCTROW Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Sum(M_Desideri.Voto) AS PUNTI, Allievi.Id_Classe INTO TPUNTEGGI_STUDENTI_MDESIDERI " &_
" FROM Allievi INNER JOIN M_Desideri ON Allievi.CodiceAllievo=M_Desideri.Id_Stud " &_
" WHERE Allievi.Id_Classe='"& id_classe& "'" &_
" AND (M_Desideri.Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"# " &_
" AND M_Desideri.Data<=#" & mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4)  &"#" &_
" OR M_Desideri.Data=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_
" GROUP BY Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe;"
	
	'	url="C:\Inetpub\umanetroot\Anno_2012-2013_2\logDESIDERI.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close	
'
'    
	
	
	
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	' CREO RIEPILOGO METAFORE TPUNTEGGI_STUDENTI_METAFORE, prendo i dati da tutte le metafore
	QuerySQL="SELECT TPUNTEGGI_STUDENTI_MTOPOLINO.CodiceAllievo, TPUNTEGGI_STUDENTI_MTOPOLINO.Nome, " &_
" TPUNTEGGI_STUDENTI_MTOPOLINO.Cognome, TPUNTEGGI_STUDENTI_MTOPOLINO.PUNTI+ TPUNTEGGI_STUDENTI_MNAVIGAZIONE.PUNTI + TPUNTEGGI_STUDENTI_MDESIDERI.PUNTI as [PM] INTO TPUNTEGGI_STUDENTI_METAFORE " &_
" FROM TPUNTEGGI_STUDENTI_MTOPOLINO,TPUNTEGGI_STUDENTI_MNAVIGAZIONE, TPUNTEGGI_STUDENTI_MDESIDERI  " &_ 
" WHERE TPUNTEGGI_STUDENTI_MTOPOLINO.CodiceAllievo = TPUNTEGGI_STUDENTI_MNAVIGAZIONE.CodiceAllievo "&_ 
" AND TPUNTEGGI_STUDENTI_MNAVIGAZIONE.CodiceAllievo=TPUNTEGGI_STUDENTI_MDESIDERI.CodiceAllievo;"
Dim queryM
queryM=QuerySQL
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	
	
	
	'Creo TCrediti
	
	  'DataCla2=left(periodi(indice_periodo2),10)
	  QuerySQL="SELECT DISTINCTROW Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Sum([2CREDITI].Crediti) AS Punti, Allievi.Id_Classe INTO TPUNTEGGI_STUDENTI_CREDITI " &_
" FROM Allievi INNER JOIN (2ESERCITAZIONI_SINGOLI INNER JOIN 2CREDITI ON [2ESERCITAZIONI_SINGOLI].ID_Esercitazione=[2CREDITI].Id_Esercitazione) ON Allievi.CodiceAllievo=[2CREDITI].Id_Stud " &_
" WHERE Allievi.Id_Classe='"& id_classe& "'" &_
" AND ([2ESERCITAZIONI_SINGOLI].Data>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
" AND [2ESERCITAZIONI_SINGOLI].Data<=#" & mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4)  &"#" &_
" OR [2ESERCITAZIONI_SINGOLI].Data=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_
" GROUP BY Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Allievi.Id_Classe; " 

 
	 
	 		'	url="C:\Inetpub\umanetroot\anno_2012-2013\logCrediti.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
	 
	 
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	'Creo TForum  prima prelevo il recordset dal db forum e poi creo la tabella
	
	 ' DataCla2=left(periodi(indice_periodo2),10)
	  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(Punti) AS PuntiForum,  Id_Classe,DatePosted  " &_
" INTO TPUNTEGGI_STUDENTI_FORUM0 FROM FORUM_MESSAGES" &_
" WHERE Id_Classe='"& id_classe& "'" &_
" AND (DatePosted>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
" AND DatePosted<=#" & 1+ CDate(mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4))  &"#" &_
" OR DatePosted=#" & mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4)  &"#)" &_
" GROUP BY  CodiceAllievo,  Id_Classe,DatePosted; " 

	  Set rsTabella = ConnessioneDB1.Execute(QuerySQL) ' eseguo sul db forum
 
	  ' url="C:\Inetpub\umanetroot\anno_2012-2013\logForumPuntiPrima.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
'				
  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(PuntiForum),  Id_Classe " &_
" FROM TPUNTEGGI_STUDENTI_FORUM0" &_
" GROUP BY  CodiceAllievo,  Id_Classe; " 

	  Set rsTabella = ConnessioneDB1.Execute(QuerySQL) ' eseguo sul db forum

 'url="C:\Inetpub\umanetroot\anno_2012-2013\logForumPuntiSeconda.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close

				
	
      k=1	  	 
	  do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_FORUM (CodiceAllievo, PuntiForum, Id_Classe) SELECT '" & rsTabella(0) & "'," & rsTabella(1) & ",'" & rsTabella(2) &"';"
		ConnessioneDB.Execute QuerySQL2 
			'   url="C:\Inetpub\umanetroot\anno_2012-2013\logForumPunti"&k&".txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL2)
'				objCreatedFile.Close
	     rsTabella.movenext
		k=k+1
	  loop
	  set rsTabella=nothing
	  
	  connessioneDB1.execute("Drop Table TPUNTEGGI_STUDENTI_FORUM0")

	 ' url="C:\Inetpub\umanetroot\anno_2012-2013\logClassifica.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL2)
'				objCreatedFile.Close	
	  
	  
' devo calcolare il max sulla nuova classifica	
    QuerySQL="SELECT MAX([TOT]) AS [MAX] FROM PUNTEGGI_STUDENTI_DATA WHERE PUNTEGGI_STUDENTI_DATA.Id_Classe='" & id_classe & "'"
	
	 
	end if 
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	max=rsTabella(0) 
	if max=0 then
	  max=1
	end if
	 'QuerySQL="SELECT Cognome, Nome, CodiceAllievo FROM Allievi WHERE Classe='" & d & "' ORDER BY Allievi.Cognome" 
								'			0						1						2							3						4				5							6
	' se il campo data è settatto devo calcolare la classifica dalla data specificata
	
	'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logMetafore2.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
	
	if (DataCla<>"") then 
	   QuerySQL="SELECT PUNTEGGI_STUDENTI_DATA.PD, PUNTEGGI_STUDENTI_DATA.Cognome, PUNTEGGI_STUDENTI_DATA.Nome, PUNTEGGI_STUDENTI_DATA.CodiceAllievo,PUNTEGGI_STUDENTI_DATA.Crediti,PUNTEGGI_STUDENTI_DATA.TOT,PUNTEGGI_STUDENTI_DATA.PN,PUNTEGGI_STUDENTI_DATA.PF,PUNTEGGI_STUDENTI_DATA.PM" &_
	" FROM PUNTEGGI_STUDENTI_DATA " &_
	" WHERE PUNTEGGI_STUDENTI_DATA.Id_Classe='" & id_classe & "'" &_
	" ORDER BY PUNTEGGI_STUDENTI_DATA.TOT DESC"
	
	
				'dim objFSO,objCreatedFile
				'Const ForReading = 1, ForWriting = 2, ForAppending = 8
				'Dim sRead, sReadLine, sReadAll, objTextFile
			'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\log.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close

 
	
	 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	
	
	 else
	 ' qua non dovrebbe più entrarci 
	' QuerySQL="SELECT PUNTEGGI_STUDENTI.PUNTI, PUNTEGGI_STUDENTI.Cognome, PUNTEGGI_STUDENTI.Nome, PUNTEGGI_STUDENTI.CodiceAllievo,PUNTEGGI_STUDENTI.Crediti,PUNTEGGI_STUDENTI.TOT,PUNTEGGI_STUDENTI.PN,PUNTEGGI_STUDENTI.PF" &_
'	" FROM PUNTEGGI_STUDENTI " &_
'	" WHERE PUNTEGGI_STUDENTI.Classe='" & d & "'" &_
'	" ORDER BY PUNTEGGI_STUDENTI.TOT DESC, PUNTEGGI_STUDENTI.PUNTI DESC"
'	  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	end if 
  
' inserisco l'esercitazione 
'se l'esercitazione si riferisce alla convalida di un quiz mi vado a prendere, a partire dal codice del quiz, il titola del modulo da mettere nella campo descrizione, altrimenti prendo il valore del campo txtVerifica
IF xQuiz<>"" then
	Titolo=Request.Querystring("TitoloTest")  
	Data=Request.Querystring("DataTest")  
	
  else
    Titolo=Request.Form("txtVerifica") 
	Data=Request.Form("txtData")
end if   
    Titolo = Replace(Titolo, Chr(34), "'")
	Titolo=  Replace(Titolo,"'","''")
	
	if  strcomp("on",Request.Form("cbScrutini"))=0 then
	  response.write("cbScrutini=on")
      Scrutini=1
    else
	  response.write("cbScrutini<>on")
      Scrutini=0
    end if
	Classifica=1
	
	TipoVoto=Request.Form("txtTipoVoto")
	if  strcomp("on",Request.Form("cbClassifica"))=0 then' lo devo registrare solo per lo scrutnino per la media dello scrutinio ma solo per la classifica
	   response.write("cbClassifica=on")
      Scrutini=1
	  Classifica=0
    end if
      
	
	
	
    QuerySQL="INSERT INTO 2ESERCITAZIONI_SINGOLI (Descrizione,Data,Id_Classe,Scrutini,Classifica,TipoVoto) SELECT '" & Titolo  & "','" & Data & "','" & id_classe & "'," & Scrutini & "," & Classifica & ",'" & TipoVoto & "';"
	

'	url="C:\Inetpub\umanetroot\anno_2012-2013_2\logCrediti.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
  response.write(QuerySQL)
   ConnessioneDB.Execute QuerySQL 
   
   'prelevo il codice dell'esercitazione appena inserita
   QuerySQL="SELECT MAX([ID_Esercitazione]) "&_
   " FROM 2ESERCITAZIONI_SINGOLI;" 
	Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
	ID_ESER=rsTabella1(0) 
IF xQuiz="" THEN ' se non sono stato chiamato dal bottone per convalidarer il quiz allora prendo i dati dal form altrimenti dalla query sui risultati del quiz
   'per ogni studente inserisco il suo punteggio prelevato dal form
    i=0
  do while not rsTabella.eof 
     allievo=rsTabella.fields("CodiceAllievo")
     QuerySQL="INSERT INTO 2CREDITI (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & ID_ESER & "','" & allievo & "','" & Request.Form("" & i & "") & "';"
	 ConnessioneDB.Execute(QuerySQL)
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
	 
	 ConnessioneDB.Execute(QuerySQL)
     rsTabella.movenext
   loop
   rsTabella.close
END IF 
if (DataCla<>"") then 
' dopo aver caricato la classifica cancello le tabelle create
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_DOMANDE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_FRASI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_NODI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_CREDITI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MTOPOLINO")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MNAVIGAZIONE")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_MDESIDERI")
	ConnessioneDB.Execute("Drop table TPUNTEGGI_STUDENTI_METAFORE")
    ConnessioneDB.Execute("Delete * From TPUNTEGGI_STUDENTI_FORUM")
end if 
ConnessioneDB.close()
Response.Redirect "studente_domande.asp?cla="&cla
'end if 
  %>
  
	</font>   
	 
		      <h4><a href="studente_domande.asp?cla=<%=cla%>">Aggiornamento avvenuto ... continua ...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
			<%
			 
			
			 if Session("stato")=0 then %>
			<a href="../../ECDL/home_ecdl_ver.asp">Torna all'Home Page Verifica</a> 
			<% else %>
			    <a href="../../U-ECDL/home_uecdl_ver.asp">Torna all'Home Page Verifica</a> 
			<% end if %>    
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	