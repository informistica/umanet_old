
	
	
      
	
	<%
	
 
	' se il campo data è settato devo calcolare la classifica dalla data specificata
	' per fare ciò devo creare le tabelle da cui la query finale preleverà i dati : 
	'TPUNTEGGI_STUDENTI : Frasi,Nodi,Crediti
	'response.write("numpoeridoi="&numPeriodi)
	if (DataCla<>"") then 
	
	' devo cercare nel vettore periodi DataCla, una volta trovata se ha un successivo lo assegno a DataCla1 
	' che mi serve per delimitare l'intervallo di tempo che mi interessa calcolare, se non c'è perchè DataCla è
	' l'ultimo periodo non faccio niente e lascio la query cosi com'è!
		for i=0 to numPeriodi 
			if left(DataCla,10) = left(periodi(i),10) then
				indice_periodo=i 
			end if	
			' se scelgo inizio a/s allora devo mostrare tutti i punnteggi, quindi faccio in modo di eseguire il ramo else dell'if che 
			'lascia tutto inalterato
			if left(DataCla,10)=left(inizio_anno,10) then
				indice_periodo=numPeriodi-1
			end if
			 if left(DataCla2,10)= left(periodi(i),10) then
				indice_periodo2=i
				
			end if
			'response.write(left(DataCla2,10) & "-" & left(periodi(i),10) & "<br>" )
	      next 
	'DataCla="09/09/2013"
	
	
		'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
			'		url="C:\Inetpub\umanetroot\expo2015Server\1_cla_46.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL="riga 66 prims tutto"
'				objCreatedFile.WriteLine(indice_periodo & " " & indice_periodo2)
'				objCreatedFile.Close
	
	
'''	' verifico se esitono già le tabelle ed in caso le elimino

'	QuerySQL="select count(*) FROM T_PUNTEGGI_STUDENTI_DOMANDE;"
'   Set rsTabella = ConnessioneDB0.Execute(QuerySQL)
'  ' response.write(QuerySQL &"<br>"&err.number ) 
'''''''   ' carico il vettore delle date di valutazione
'	if err.number = 0 then
'''''''	' non c'è errore quindi la tabella esiste quindi le cancello tutte
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_DOMANDE")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_FRASI")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_NODI")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_MTOPOLINO")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_MNAVIGAZIONE")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_MDESIDERI")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_METAFORE")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_CREDITI")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_FORUM")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_DIARIO")
'	ConnessioneDB0.Execute("Drop table TPUNTEGGI_STUDENTI_LAVAGNA") ' da implementare
'	end if
'	
	'Creo TDomande	
	DataCla2=left(periodi(indice_periodo2),10)	 
'DataCla2="11/11/2013"
' per il problema per lo zero omesso da iis7
%>


     

<%

function formattaDataCla(DataCla)
  giornoD=DatePart("d",DataCla)
 if len(giornoD)=1 then
    giornoD= "0" & giornoD
 end if
 meseD=DatePart("m",DataCla)
  if len(meseD)=1 then
    meseD= "0" & meseD
 end if
 annoD=DatePart("yyyy",DataCla)
 formattaDataCla=meseD&"/"&giornoD&"/"&annoD
end function

'DataCla=formattaDataCla(DataCla)
	
	
	   
		QuerySQL="SELECT DISTINCT CodiceAllievo, Nome, Cognome, Id_Classe, Sum(Voto) AS PD " &_
" FROM Allievi INNER JOIN Domande ON CodiceAllievo = Id_Stud" &_
" WHERE Id_Classe='"& id_classe& "'" &_
 " and (Data>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataCla2) &"', 104))"&_
	  " OR (Data= CONVERT(DATETIME,'" &CDATE(DataClaFine) &"', 104))"&_
" GROUP BY CodiceAllievo, Nome, Cognome, Id_Classe;"

 
				url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
					url="C:\Inetpub\umanetroot\expo2015Server\log1.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'QuerySQL="riga 66 prims tutto"
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close

          





Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
   
    do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_DOMANDE  (CodiceAllievo, Nome, Cognome,Id_Classe,PD) SELECT '" & rsTabella(0) & "','" & rsTabella(1) & "','" & rsTabella(2) & "','" & rsTabella(3) & "'," & rsTabella(4) &";"
		
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabella.movenext
	  loop
	  set rsTabella=nothing
	  
 
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\110.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'			
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
   
	'Creo TFrasi
	  DataCla2=left(periodi(indice_periodo2),10)
		QuerySQL="SELECT DISTINCT CodiceAllievo, Nome, Cognome,  Sum(Voto) AS PUNTI, Id_Classe " &_
	" FROM Allievi LEFT JOIN Frasi ON CodiceAllievo = Id_Stud " &_
	" WHERE Id_Classe='"& id_classe& "'" &_
	 " and (Data>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataCla2) &"', 104))"&_
	  " OR (Data= CONVERT(DATETIME,'" &CDATE(DataClaFine) &"', 104))"&_
	" GROUP BY CodiceAllievo, Nome, Cognome, Id_Classe;"
 
	
	'response.write("<br>"&QuerySQL)		
	 

		 
				url="C:\Inetpub\umanetroot\expo2015Server\log2.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
 	
	
	
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	 do while not (rsTabella.eof) 'and k<40
	 
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_FRASI  (CodiceAllievo, Nome, Cognome,Punti,Id_Classe) SELECT '" & rsTabella(0) & "','" & rsTabella(1) & "','" & rsTabella(2) & "'," & rsTabella(3) & ",'" & rsTabella(4) &"';"
		objCreatedFile.WriteLine(QuerySQL2&"<br>")
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabella.movenext
	  loop
	  set rsTabella=nothing
	
	'objCreatedFile.Close
	
	'url="C:\Inetpub\umanetroot\expo2015Server\log183.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
	
	'Creo TNodi
	  DataCla2=left(periodi(indice_periodo2),10)
	QuerySQL="SELECT  DISTINCT CodiceAllievo, Nome, Cognome, Sum(Voto) AS PUNTI, Id_Classe " &_
" FROM Allievi INNER JOIN Nodi ON CodiceAllievo=Nodi.Id_Stud " &_
" WHERE Id_Classe='"& id_classe& "'" &_
 " and (Data>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataCla2) &"', 104))"&_
	  " OR (Data= CONVERT(DATETIME,'" &CDATE(DataClaFine) &"', 104))"&_
" GROUP BY CodiceAllievo, Nome, Cognome, Id_Classe;"
	
	url="C:\Inetpub\umanetroot\expo2015Server\log3.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
	

'response.write("<br>"&QuerySQL)	
      Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	 do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_NODI  (CodiceAllievo, Nome, Cognome,Punti,Id_Classe) SELECT '" & rsTabella(0) & "','" & rsTabella(1) & "','" & rsTabella(2) & "'," & rsTabella(3) & ",'" & rsTabella(4) &"';"
		
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabella.movenext
	  loop
	  set rsTabella=nothing
	
	
	
	
	'Creo TMetafore
	
	' Creo Topolino
 	 DataCla2=left(periodi(indice_periodo2),10)
	QuerySQL="SELECT  DISTINCT CodiceAllievo, Nome, Cognome, Sum(Voto) AS PUNTI, Id_Classe " &_
" FROM Allievi INNER JOIN M_Topolino ON CodiceAllievo=M_Topolino.Id_Stud " &_
" WHERE Id_Classe='"& id_classe& "'" &_
 " and (Data>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataCla2) &"', 104))"&_
	  " OR (Data= CONVERT(DATETIME,'" &CDATE(DataClaFine) &"', 104))"&_
" GROUP BY CodiceAllievo, Nome, Cognome, Id_Classe;"
   
url="C:\Inetpub\umanetroot\expo2015Server\log4.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
	
	'response.write("<br>"&QuerySQL)	
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	
	
	 do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_MTOPOLINO   (CodiceAllievo, Nome, Cognome,Punti,Id_Classe) SELECT '" & rsTabella(0) & "','" & rsTabella(1) & "','" & rsTabella(2) & "'," & rsTabella(3) & ",'" & rsTabella(4) &"';"
		
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabella.movenext
	  loop
	  set rsTabella=nothing
	
	
	' Creo Navigazione
 	 DataCla2=left(periodi(indice_periodo2),10)
	QuerySQL="SELECT DISTINCT CodiceAllievo, Nome, Cognome, Sum(Voto) AS PUNTI, Id_Classe " &_
" FROM Allievi INNER JOIN M_Navigazione ON CodiceAllievo=M_Navigazione.Id_Stud " &_
" WHERE Id_Classe='"& id_classe& "'" &_
 " and (Data>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataCla2) &"', 104))"&_
	  " OR (Data= CONVERT(DATETIME,'" &CDATE(DataClaFine) &"', 104))"&_
" GROUP BY CodiceAllievo, Nome, Cognome, Id_Classe;"
  
  
  url="C:\Inetpub\umanetroot\expo2015Server\log5.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
	
  
 ' response.write("<br>"&QuerySQL)	
   Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
  
	         
			
				''k=0
  do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_MNAVIGAZIONE    (CodiceAllievo, Nome, Cognome,Punti,Id_Classe) SELECT '" & rsTabella(0) & "','" & rsTabella(1) & "','" & rsTabella(2) & "'," & rsTabella(3) & ",'" & rsTabella(4) &"';"
		'objCreatedFile.WriteLine(QuerySQL2)
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabella.movenext
'k=k+1
	  loop
	  set rsTabella=nothing
	
  
     'url="C:\Inetpub\umanetroot\expo2015Server\log271.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
  
  QuerySQL="SELECT DISTINCT CodiceAllievo, Nome, Cognome, Sum(Voto) AS PUNTI, Id_Classe " &_
" FROM Allievi INNER JOIN M_Desideri ON CodiceAllievo=M_Desideri.Id_Stud " &_
" WHERE Id_Classe='"& id_classe& "'" &_
 " and (Data>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataCla2) &"', 104))"&_
	  " OR (Data= CONVERT(DATETIME,'" &CDATE(DataClaFine) &"', 104))"&_
" GROUP BY CodiceAllievo, Nome, Cognome, Id_Classe;"
	
	url="C:\Inetpub\umanetroot\expo2015Server\log6.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
''    
	
	
	'response.write("<br>cazzo"&QuerySQL)	
	  Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
	finito=0	
	k=1
	do while not (rsTabella1.eof) 
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_MDESIDERI     (CodiceAllievo, Nome, Cognome,Punti,Id_Classe) SELECT '" & rsTabella1(0) & "','" & rsTabella1(1) & "','" & rsTabella1(2) & "'," & rsTabella1(3) & ",'" & rsTabella1(4) &"';"	
		ConnessioneDB0.Execute QuerySQL2 
	'	objCreatedFile.WriteLine(QuerySQL2)
	     rsTabella1.movenext
		 
	  loop
	  set rsTabella1=nothing

	'objCreatedFile.Close
	
	' CREO RIEPILOGO METAFORE TPUNTEGGI_STUDENTI_METAFORE, prendo i dati da tutte le metafore
	QuerySQL="SELECT distinct(TPUNTEGGI_STUDENTI_MTOPOLINO.CodiceAllievo), TPUNTEGGI_STUDENTI_MTOPOLINO.Nome, " &_
" TPUNTEGGI_STUDENTI_MTOPOLINO.Cognome, TPUNTEGGI_STUDENTI_MTOPOLINO.PUNTI+ TPUNTEGGI_STUDENTI_MNAVIGAZIONE.PUNTI + TPUNTEGGI_STUDENTI_MDESIDERI.PUNTI as [PM] " &_
" FROM TPUNTEGGI_STUDENTI_MTOPOLINO,TPUNTEGGI_STUDENTI_MNAVIGAZIONE, TPUNTEGGI_STUDENTI_MDESIDERI  " &_ 
" WHERE TPUNTEGGI_STUDENTI_MTOPOLINO.CodiceAllievo = TPUNTEGGI_STUDENTI_MNAVIGAZIONE.CodiceAllievo "&_ 
" AND TPUNTEGGI_STUDENTI_MNAVIGAZIONE.CodiceAllievo=TPUNTEGGI_STUDENTI_MDESIDERI.CodiceAllievo;"
Dim queryM
queryM=QuerySQL


url="C:\Inetpub\umanetroot\expo2015Server\log7.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
'response.write("<br>"&QuerySQL)	
Set rsTabella1 = ConnessioneDB0.Execute(QuerySQL) 
	
	'url="C:\Inetpub\umanetroot\anno_2013-2014\logMetafora.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)	
''    
rsTabella1.movefirst
	do while not (rsTabella1.eof) 'and k<40
		'objCreatedFile.WriteLine("INSERT INTO TPUNTEGGI_STUDENTI_METAFORE(CodiceAllievo, Nome, Cognome,PM) SELECT '" & rsTabella1(0) & "','" & rsTabella1(1) & "','" & rsTabella1(2) & "'," & rsTabella1(3) &";")
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_METAFORE(CodiceAllievo, Nome, Cognome,PM) SELECT '" & rsTabella1(0) & "','" & rsTabella1(1) & "','" & rsTabella1(2) & "'," & rsTabella1(3) &";"
		 
		'objCreatedFile.WriteLine(QuerySQL2)		
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabella1.movenext
	  loop
	  set rsTabella1=nothing
	'objCreatedFile.Close
	
	
	
	'Creo TCrediti
	
	  DataCla2=left(periodi(indice_periodo2),10)
	  QuerySQL="SELECT DISTINCT  CodiceAllievo, Nome,  Cognome, Sum([dbo].[2CREDITI].Crediti) AS Punti,  [dbo].[Allievi].Id_Classe" &_
" FROM Allievi INNER JOIN ([dbo].[2ESERCITAZIONI_SINGOLI] INNER JOIN [dbo].[2CREDITI] ON [dbo].[2ESERCITAZIONI_SINGOLI].ID_Esercitazione=[dbo].[2CREDITI].Id_Esercitazione) ON  CodiceAllievo=Id_Stud " &_
" WHERE [dbo].[Allievi].Id_Classe='"& id_classe& "'" &_
 " and (Data>= CONVERT(DATETIME,'" &DataCla  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataCla2) &"', 104))"&_
	  " OR (Data= CONVERT(DATETIME,'" &CDATE(DataClaFine) &"', 104))"&_
" GROUP BY CodiceAllievo, Nome, Cognome, [dbo].[Allievi].Id_Classe; " 


url="C:\Inetpub\umanetroot\expo2015Server\log8.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
'response.write("<br><br>"&QuerySQL)	
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	  
	  do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_CREDITI   (CodiceAllievo, Nome, Cognome,Punti,Id_Classe) SELECT '" & rsTabella(0) & "','" & rsTabella(1) & "','" & rsTabella(2) & "'," & rsTabella(3) & ",'" & rsTabella(4) &"';"
		
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabella.movenext
	  loop
	  set rsTabella=nothing
	  
	 ' url="C:\Inetpub\umanetroot\expo2015Server\log357.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL2)
'				objCreatedFile.Close
	' fin qui OK
	
	'Creo TForum  prima prelevo il recordset dal db forum e poi creo la tabella
	 
	 
	 
	 
	 
	 
	  DataCla2=left(periodi(indice_periodo2),10)
	 ' response.write("<br>periodi(indice_periodo2)="&periodi(indice_periodo2))
	  
	  
	 ' DataCla= mid(DataCla,4,2) &"/" &left(DataCla,2)&"/"& right(DataCla,4) 
	  'DataCla2=1+ CDate(mid(DataCla2,4,2)&"/" &left(DataCla2,2)&"/"& right(DataCla2,4))
	  'prova ad invertite gg e mm per renderela come la queruy sul dbcopiatestonline  quindi
	   DataCla= left(DataCla,2) &"/" &mid(DataCla,4,2) &"/"& right(DataCla,4) 
	   DataCla2=1+ CDate(left(DataCla2,2)&"/" &mid(DataCla2,4,2)&"/"& right(DataCla2,4))
	   DataClaFine="12/12/2112"
	  
	  'response.write("<br>DataCla2="&DataCla2)
	 ' DataClaFine2=mid(DataClaFine,4,2)&"/" &left(DataClaFine,2)&"/"& right(DataClaFine,4) 
	  'response.write("<br> DataClaFine2="&DataClaFine2)
	  
	   ' DataCla="09/05/2012"
	   'DataCla2="12/15/2013"
'	  
'	 
'	 
'	  DataCla2="20/12/2013"
	  
	   
	  
	  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(Punti) AS PuntiForum,  Id_Classe,DatePosted  " &_
" INTO TPUNTEGGI_STUDENTI_FORUM0 FROM FORUM_MESSAGES" &_
" WHERE Id_Classe='"& id_classe& "'" &_
" AND (DatePosted>=#" & DataCla &"#" &_
" AND DatePosted<=#"  & cdate(formattaDataCla(DataCla2)) &"# OR DatePosted=#" & DataClaFine  &"#)" &_
" GROUP BY  CodiceAllievo,  Id_Classe,DatePosted; " 
'response.write(QuerySQL)
'" AND DatePosted<=#" &CDate(1+cint(mid(DataCla2,4,2))&"/" &left(DataCla2,2)&"/"& right(DataCla2,4))  &"#" &_ 

'response.write("<br><br>"&QuerySQL)	
url="C:\Inetpub\umanetroot\expo2015Server\log9.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close

'" AND (DatePosted>=#" &left(DataCla,2) &"/" &mid(DataCla,4,2)&"/"& right(DataCla,4)  &"#" &_
'" AND DatePosted<=#" & 1+ CDate(left(DataCla2,2)&"/" &mid(DataCla2,4,2)&"/"& right(DataCla2,4))  &"#" &_ 

' " AND DatePosted<=#" & 1+ CDate(left(DataCla2,2)&"/" &mid(DataCla2,4,2)&"/"& right(DataCla2,4))  &"#" &_
'" AND (DatePosted>=#" & mid(DataCla,4,2)&"/" &left(DataCla,2)&"/"& right(DataCla,4)  &"#" &_
'" AND DatePosted<=#01/12/2013#" &_
	  Set rsTabella = ConnessioneDB1.Execute(QuerySQL) ' eseguo sul db forum
 
	 '  url="C:\Inetpub\umanetroot\expo2015Server\logForumPuntiPrima.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
'			response.write("<br>"& QuerySQL)

  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(PuntiForum),  Id_Classe " &_
" FROM TPUNTEGGI_STUDENTI_FORUM0" &_
" GROUP BY  CodiceAllievo,  Id_Classe; " 
'response.write("<br>"&QuerySQL)	
	  Set rsTabella = ConnessioneDB1.Execute(QuerySQL) ' eseguo sul db forum

 url="C:\Inetpub\umanetroot\expo2015Server\log9_1.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
'	
      k=1	  	
	  ' QuerySQL3="pippo"
	 '  url="C:\Inetpub\umanetroot\anno_2013-2014\log9_2.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL3)
'				objCreatedFile.Close
				
	  do while not (rsTabella.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_FORUM (CodiceAllievo, PuntiForum, Id_Classe) SELECT '" & rsTabella(0) & "'," & rsTabella(1) & ",'" & rsTabella(2) &"';"
		
		ConnessioneDB0.Execute QuerySQL2 
		 ' QuerySQL2="pippo"
			  ' url="C:\Inetpub\umanetroot\anno_2013-2014\logForum2Punti"&k&".txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL2)
'				objCreatedFile.Close
	     rsTabella.movenext
		k=k+1
	  loop
	  set rsTabella=nothing
	  
	  connessioneDB1.execute("Drop Table TPUNTEGGI_STUDENTI_FORUM0")



' ripeto la stessa logica per il LAVAGNA
 
	 
	  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(Punti) AS PuntiForum,  Id_Classe " &_
" INTO TPUNTEGGI_STUDENTI_LAVAGNA0 FROM FORUM_MESSAGES" &_
" WHERE Id_Classe='"& id_classe& "'" &_
" AND (DatePosted>=#"&DataCla&"#" &_
" AND DatePosted<=#" &cdate(formattaDataCla(DataCla2))  &"#" &_ 
 " OR DatePosted=#" & DataClaFine   &"#)" &_
" GROUP BY  CodiceAllievo,  Id_Classe; " 

url="C:\Inetpub\umanetroot\expo2015Server\log10.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
'' 
'response.write("<br><br>"&QuerySQL)	
	  'Set rsTabella = ConnessioneDB2.Execute(QuerySQL) ' eseguo sul db lavagna
	  ConnessioneDB2.Execute(QuerySQL) ' eseguo sul db lavagna
  

  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(PuntiForum),  Id_Classe " &_
" FROM TPUNTEGGI_STUDENTI_LAVAGNA0" &_
" GROUP BY  CodiceAllievo,  Id_Classe; " 

url="C:\Inetpub\umanetroot\expo2015Server\log11.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close

	  Set rsTabella = ConnessioneDB2.Execute(QuerySQL) ' eseguo sul db lavagna

  
       	  	 k=0
	  do while not (rsTabella.eof)  
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_LAVAGNA (CodiceAllievo, PuntiForum, Id_Classe) SELECT '" & rsTabella(0) & "'," & rsTabella(1) & ",'" & rsTabella(2) &"';"
		ConnessioneDB0.Execute QuerySQL2 
		objCreatedFile.WriteLine(k&")"&QuerySQL2)
	     rsTabella.movenext	 
		 k=k+1	 
	  loop
	  set rsTabella=nothing
	  
	 
	 
	   connessioneDB2.execute("Drop Table TPUNTEGGI_STUDENTI_LAVAGNA0")
' NON CAPISCO PERCHE IL LOG HA 72 INSERT MENTRE LA TABELLA NE HA IL DOPPIO

' ripeto la stessa logica per il DIARIO

'Creo TDiario  prima prelevo il recordset dal db diario e poi creo la tabella	 
	' gg=CDate(mid(DataCla2,4,2)
	 
	 	' url="C:\Inetpub\umanetroot\expo2015Server\log520.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				
				
				
	  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(Punti) AS PuntiForum, Id_Classe  " &_
" INTO TPUNTEGGI_STUDENTI_DIARIO0 FROM FORUM_MESSAGES" &_
" WHERE Id_Classe='"& id_classe& "'" &_
" AND (DatePosted>=#"&DataCla&"#" &_
" AND DatePosted<=#" &cdate(formattaDataCla(DataCla2))  &"#" &_ 
 " OR DatePosted=#" & DataClaFine  &"#)" &_
" GROUP BY  CodiceAllievo,Id_Classe; " 

url="C:\Inetpub\umanetroot\expo2015Server\log12.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
	  ConnessioneDB3.Execute(QuerySQL) ' eseguo sul db diario
 'response.write("<br><br>"&QuerySQL)	 

  QuerySQL="SELECT DISTINCTROW CodiceAllievo, Sum(PuntiForum),Id_Classe  " &_
" FROM TPUNTEGGI_STUDENTI_DIARIO0" &_
" GROUP BY  CodiceAllievo,Id_Classe; " 


 

	  Set rsTabellaD = ConnessioneDB3.Execute(QuerySQL) ' eseguo sul db diario


url="C:\Inetpub\umanetroot\expo2015Server\log12_1.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close

  
      k=1
	  rsTabellaD.movefirst	  	 
	  do while not (rsTabellaD.eof) 'and k<40
		QuerySQL2="INSERT INTO TPUNTEGGI_STUDENTI_DIARIO (CodiceAllievo, PuntiForum, Id_Classe) SELECT '" & rsTabellaD(0) & "'," & rsTabellaD(1) & ",'" & rsTabellaD(2) &"';"
		
		'objCreatedFile.WriteLine(k&")"&QuerySQL2)
		
		ConnessioneDB0.Execute QuerySQL2 
	     rsTabellaD.movenext
		k=k+1
	  loop
	  set rsTabellaD=nothing
	  
	  
	  'objCreatedFile.Close
	  
	  connessioneDB3.execute("Drop Table TPUNTEGGI_STUDENTI_DIARIO0")



	 'url="C:\Inetpub\umanetroot\expo2015Server\log570.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL2)
'				objCreatedFile.Close
	  
	  
' devo calcolare il max sulla nuova classifica	
	' calcolo su tutti i punti inclusi i punti social
	if cint(PS)=1 then
   			 QuerySQL="SELECT MAX([TOT]) AS [MAX] FROM PUNTEGGI_STUDENTI_DATA WHERE PUNTEGGI_STUDENTI_DATA.Id_Classe='" & id_classe & "'"
	else
	      QuerySQL="SELECT MAX([TOT]) AS [MAX] FROM PUNTEGGI_STUDENTI_DATA_SINT WHERE PUNTEGGI_STUDENTI_DATA_SINT.Id_Classe='" & id_classe & "'"
	end if
	
else
	   ' qua non ci entrerà più perchè se non viene selezionata la data la impongo uguale a inizio_anno definita in var_globali.inc
	  ' QuerySQL="SELECT MAX([TOT]) AS [MAX] FROM PUNTEGGI_STUDENTI WHERE PUNTEGGI_STUDENTI.Classe='" & d & "'"
	end if 
	
	
		 url="C:\Inetpub\umanetroot\expo2015Server\log14.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
'	
	
	Set rsTabella = ConnessioneDB0.Execute(QuerySQL) 
	max=rsTabella(0) 
	if max=0 then
	  max=1
	end if 
	 'QuerySQL="SELECT Cognome, Nome, CodiceAllievo FROM Allievi WHERE Classe='" & d & "' ORDER BY Allievi.Cognome" 
								'			0						1						2							3						4				5							6
	' se il campo data è settatto devo calcolare la classifica dalla data specificata
	
'url="C:\Inetpub\umanetroot\anno_2013-2014\log13.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
	
	if (DataCla<>"") then 
	
	if cint(PS)=0 then ' se non devo mostrare i Punti Social faccio query parziale
	    QuerySQL="SELECT PUNTEGGI_STUDENTI_DATA_SINT.PD, PUNTEGGI_STUDENTI_DATA_SINT.Cognome, PUNTEGGI_STUDENTI_DATA_SINT.Nome, PUNTEGGI_STUDENTI_DATA_SINT.CodiceAllievo,PUNTEGGI_STUDENTI_DATA_SINT.Crediti,PUNTEGGI_STUDENTI_DATA_SINT.TOT,PUNTEGGI_STUDENTI_DATA_SINT.PN,PUNTEGGI_STUDENTI_DATA_SINT.PF,PUNTEGGI_STUDENTI_DATA_SINT.PM " &_
	" FROM PUNTEGGI_STUDENTI_DATA_SINT " &_
	" WHERE PUNTEGGI_STUDENTI_DATA_SINT.Id_Classe='" & id_classe & "'" &_
	" ORDER BY PUNTEGGI_STUDENTI_DATA_SINT.TOT DESC"
	else
	
	   QuerySQL="SELECT PUNTEGGI_STUDENTI_DATA.PD, PUNTEGGI_STUDENTI_DATA.Cognome, PUNTEGGI_STUDENTI_DATA.Nome, PUNTEGGI_STUDENTI_DATA.CodiceAllievo,PUNTEGGI_STUDENTI_DATA.Crediti,PUNTEGGI_STUDENTI_DATA.TOT,PUNTEGGI_STUDENTI_DATA.PN,PUNTEGGI_STUDENTI_DATA.PF,PUNTEGGI_STUDENTI_DATA.PM,PUNTEGGI_STUDENTI_DATA.PuntiForum, PUNTEGGI_STUDENTI_DATA.PuntiDiario,PUNTEGGI_STUDENTI_DATA.PuntiLavagna" &_
	" FROM PUNTEGGI_STUDENTI_DATA " &_
	" WHERE PUNTEGGI_STUDENTI_DATA.Id_Classe='" & id_classe & "'" &_
	" ORDER BY PUNTEGGI_STUDENTI_DATA.TOT DESC"

'QuerySQL="SELECT PUNTEGGI_STUDENTI_DATA.PD, PUNTEGGI_STUDENTI_DATA.Cognome, PUNTEGGI_STUDENTI_DATA.Nome, PUNTEGGI_STUDENTI_DATA.CodiceAllievo,PUNTEGGI_STUDENTI_DATA.Crediti,PUNTEGGI_STUDENTI_DATA.TOT,PUNTEGGI_STUDENTI_DATA.PN,PUNTEGGI_STUDENTI_DATA.PF,PUNTEGGI_STUDENTI_DATA.PM,PUNTEGGI_STUDENTI_DATA.PuntiForum, PUNTEGGI_STUDENTI_DATA.PuntiDiario" &_
'	" FROM PUNTEGGI_STUDENTI_DATA " &_
'	" WHERE PUNTEGGI_STUDENTI_DATA.Id_Classe='" & id_classe & "'" &_
'	" ORDER BY PUNTEGGI_STUDENTI_DATA.TOT DESC"
'	
'	
	end if
	
		url="C:\Inetpub\umanetroot\expo2015Server\log15.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
 
 
 
 'querySQL="Select * From Allievi where Id_Classe='"&id_classe &"';"
	
	 Set rsTabella = ConnessioneDB0.Execute(QuerySQL)
	 
	 '********* QUA AGGIUNGERO' IL CODICE PER SALVARE LA CLASSIFICA QUANDO INIZIA UN NUOVO PERIODO
	
	
	 else
	 ' qua non dovrebbe più entrarci 
	' QuerySQL="SELECT PUNTEGGI_STUDENTI.PUNTI, PUNTEGGI_STUDENTI.Cognome, PUNTEGGI_STUDENTI.Nome, PUNTEGGI_STUDENTI.CodiceAllievo,PUNTEGGI_STUDENTI.Crediti,PUNTEGGI_STUDENTI.TOT,PUNTEGGI_STUDENTI.PN,PUNTEGGI_STUDENTI.PF" &_
'	" FROM PUNTEGGI_STUDENTI " &_
'	" WHERE PUNTEGGI_STUDENTI.Classe='" & d & "'" &_
'	" ORDER BY PUNTEGGI_STUDENTI.TOT DESC, PUNTEGGI_STUDENTI.PUNTI DESC"
'	  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	end if 
	
	%>   
	<!-- <h4 style="text-align: center"><i><a href="../home.asp" >Vai all'HomePage</a></i> </h4> -->
<!--<span class="sottotitolo"><a href="../home_app.asp?id_classe=<%=Session("id_classe")%>&divid=<%=Session("divid")%>" >Vai all'HomePage</a></span><br><br>-->
   
 
<%
	
	'
'	Domanda="?"
'		R1="?"
'		R2="?"
'		R3="?"
'	R4="?"
'		Chi="?"
'		Cosa="?"
'		Dove="?"
'		Quando="?"
'		Come="?"
'		Perche="?"
'		Quindi="?"
'		RE=1
'		Spiegazione="?"
		CodiceAllievo=id
		'CodiceTest="6Classe"
'		Modulo="6C" 
'		DataTest="12/12/2112"  ' NB DOPO QUESTA DATA SE ESISTEREMO ANCORA DOVRO' METTERE UNA DATA SEMPRE MAGGIORE DELL'A/S in CORSO IN MODO DA FAR FUNZIONARE IL LEFT JOIN  NELLE QUERY PER LE CLASSIFICHE CON DATA VARIABILE
'	'	 dovrò inserire aNCHE I CREDITI INIZIZIALI !
'		Cartella="?"
		Voto=0
'		In_Quiz=0
'		Topolino="?"
'        Autista="?"
		
'response.write(QuerySQL)		 
		 
	
	  			'   SoggettoC =  "?"
'				   DomandaC =  "?"
'				   MotivazioneC = "?" 
'				   DesiderioC = "?"
'				   BisognoC="?"
'				   SoggettoS =  "?"
'				   RispostaS = "?"    
'				   MotivazioneS = "?"    
'				   DesiderioS =  "?" 
'				   BisognoS="?"
'				   TipoEvento = 1 
'				   TolleranzaC = 3 
'	
	%>
	