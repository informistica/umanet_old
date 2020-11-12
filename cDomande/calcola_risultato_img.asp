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
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,Stato
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest 
   Dim i,RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
    
%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
   
                        'Lettura dei dati memorizzati nei cookie. 
   CodiceTest = Request.Cookies("Dati")("CodiceTest")
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")

Function url_img(cartella,nome_img)	 
	 url_img="../img_quiz" & "/" & cartella &"/" & nome_img&".jpg"
 	 'url_img=replace(url_img,"/","\")
End Function 

Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
   DataTest = gira_data()
   
   Stato=Request.QueryString("Stato")
   Modulo=Request.QueryString("Modulo")
   CodiceTest=Request.QueryString("CodiceTest")
   'Definizione query SQL per contare il numero di domande del test.
   
'   QuerySQL="SELECT count(*)" &_
'             "FROM Domande INNER JOIN " &_
'             "(Test INNER JOIN ComposizioneTest ON " &_
'             "Test.CodiceTest = ComposizioneTest.CodiceTest) " &_
'             "ON Domande.CodiceDomanda = ComposizioneTest.CodiceDomanda " &_
'             "WHERE Test.CodiceTest='" & CodiceTest & "';"
 
'Definzione codice SQl della query per ricercare le domande del modulo
 
	
	if (Stato=0) then 
 'Definzione codice SQl della query per ricercare le domande del paragrafo 
   QuerySQL="SELECT count(*) " &_
             "FROM Domande1 " &_
             "WHERE Domande1.Id_Arg='" & CodiceTest & "';"
    'Assegna alla variabile il risultato della query prodotta utilizzando il metodo Execute(stringa della query) dell'oggetto connessione
else 
'Definzione codice SQl della query per ricercare le domande del modulo
QuerySQL="SELECT count(*) " &_
             "FROM Domande1 " &_
             "WHERE Domande1.Id_Mod='" & Modulo & "';"
end if   
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
NumDom=rsTabella(0).value 'Assegno a NumDom numero delle domande


if (Stato=0) then 
   'Definizione query SQL per la lettura delle risposte esatta nel test scelto.
   QuerySQL="SELECT Domande1.CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria " &_
             "FROM Domande1 " &_
             "WHERE Domande1.Id_Arg='" & CodiceTest & "' order by Domande1.CodiceDomanda asc;"
else
  QuerySQL="SELECT Domande1.CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria " &_
             "FROM Domande1 " &_
             "WHERE Domande1.Id_Mod='" & Modulo & "' order by Domande1.CodiceDomanda asc;"
end if  




 
   

    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    
  'Calcolo del numero di risposte esatte.  
  i=1
  inbianco=0
  RisposteKO = 0   			  'contatore delle risposte esatte
  RisposteOK = 0   			  'contatore delle risposte errate
  ReDim RispDate(NumDom+1)    'dimensionamento dell'array dinamico che tiene traccia delle risposte date
  ReDim RispEsatte(NumDom+1)  'dimensionamento dell'array dinamico che tiene traccia delle risposte esatte
  ReDim RispDate1(NumDom+1)    'dimensionamento dell'array dinamico che tiene traccia delle risposte date
  ReDim RispEsatte1(NumDom+1) 
 ReDim Errori(NumDom+1) 	  'dimensionamento dell'array dinamico che tiene traccia degli errori
  
  Do While not(rsTabella.EOF) ' per ogni risposta confronta la risposta esatta con quella data dall'utente
     RispostaEsatta=rsTabella.Fields("RispostaEsatta") 'legge dal risultato e memorizza in una variabile d'appoggio la risposta esatta
     'legge il valore associato all'oggetto avente per nome il numero contenuto nella variabile i, in base al valore ricava la risposta data  
     SELECT CASE Request.Form("" & i & "")
     CASE "1"
       RispostaData=1
     CASE "2"
       RispostaData=2
     CASE "3"
       RispostaData=3
     CASE "4"
       RispostaData=4
     CASE ELSE
     	RispostaData=0
     END SELECT  
     dim a
     a=1
     
     RispDate(i) = RispostaData     ' memorizza nel vettore risposte date il valore della risposta data (i)
    ' Response.write(rsTabella.Fields.Count)
    ' Response.write(rsTabella.Fields(a).value)
    ' Response.write(rsTabella.Fields(a+1).value)
     'Response.write(rsTabella.Fields(1).value)
     'Response.write(rsTabella.Fields(2).value)
     'Response.write(rsTabella.Fields(3).value)
     IF (RispostaData=0) THEN
        RispDate1(i)= "IN BIANCO"
        inbianco=inbianco+1
     ELSE 
	   IF (RispostaData=4) THEN
	      RispDate1(i)= "NESSUNA DELLE PRECEDENTI"
	   end if 
       RispDate1(i) =  rsTabella.Fields(1+RispostaData).value
       'Response.Write(rsTabella.Fields(1+RispostaData).value)
     END IF
     
     RispEsatte(i) = RispostaEsatta ' memorizza nel vettore risposte esatte il valore della risposta esatta (i)
     RispEsatte1(i) = rsTabella.Fields(1+RispostaEsatta).value ' *****si blocca qua non gli piace l'assegnazione
     IF (RispostaEsatta=RispostaData) THEN  ' se sono uguali incrementa il numero delle risposte ok e pone a 0 l'elemento i del vettore errori 
           RisposteOK = RisposteOK +1
           Errori(i)=0 				'0 = domanda i esatta
     ELSE       					'1 = domanda i errata  
           Errori(i)=1				'se sono diversi incrementa il numero delle risposte ko e pone a 1 l'elemento i del vettore errori 
           RisposteKO = RisposteKO +1   
     END IF
     i = i + 1						' incrementa i
     rsTabella.MoveNext 			' passa alla prossima domanda
   Loop 
   
   'Calcolo della percentuale di domande corrette. 
    Risultato = (RisposteOK/(i-1))*100
    Risultato_relativo = (RisposteOK/(i-inbianco-1))*100
   
    
    DataTest=date()
   'Esecuzione della query per inserire il risultato del test nella tabella Risulati
   QuerySQL="  INSERT INTO Risultati (CodiceAllievo, CodiceTest, Data,Ora,Risultato) SELECT '" & CodiceAllievo & "','" & CodiceTest & "', '" & DataTest & "', '" & FormatDateTime(now, 4) & "','" & Round(Risultato,0) & "';"
   ConnessioneDB.Execute QuerySQL 
  
   'Stampa del risultato all'utente
   Response.Write("<H3>Risultato assoluto del test : " & Round(Risultato,0) & "% - Voto = " &  Round(Risultato,0)*8/100 &"</H3>")
   %>
	</font>   
	<font color=#3333FF size="3">
   <%
   Response.Write("<H4>Su un totale di " & NumDom & " domande ci sono " & RisposteOK & " risposte corrette e " & RisposteKO & " risposte errate <BR>")
   %>
	</font>   
	<font color=#FF0000 size="3">
   <%  
   Response.Write("<H3>Risultato relativo del test : " & Round(Risultato_relativo,0) & "% - Voto = " &  Round(Risultato_relativo,0)*8/100 &"</H3>")
    %>
	</font>   
	<font color=#3333FF size="3">
   <%
   Response.Write("<H4>Su un totale di " & NumDom-inbianco & " domande risposte ci sono " & RisposteOK & " risposte corrette e " & NumDom-inbianco-RisposteOK & " risposte errate <BR>")

  %>
   <!-- stampa la tabella per offire l'opportunità di visualizzare le correzioni -->
   	</font>
   	<p>     
   		
		
	  <table border=1> 
		<tr>
			<td><font color=#FF0000><b>Domanda</b> </td>
			<td><font color=#FF0000><b>Codice</b> </td>
			<td><font color=#FF0000><b>Quesito</b> </td>
			<td><font color=#FF0000><b>Risposta Data</b> </td>
			<td><font color=#FF0000><b>Risposta Esatta</b> </td>
		</tr>
		</font>
		<%	rsTabella.Movefirst ' torna all'inizio delle domande
		   	    i=1
				Do While Not rsTabella.EOF %>
		  <tr>
			<% if Errori(i)=1 then %>  <!-- se la risposta è errata usa il colore rosso -->
				
		   
     			
		    <td valign=top><font color="red"> <b><%=i%></b></font></td>
		    <td valign=top><font color="red"><b><%=rsTabella.Fields("CodiceDomanda")%></b></font></td>
		    <td valign=top><font color="red"><b><%=rsTabella.Fields("Quesito")%></b></font></td>
		    <td valign=top> <img src="<%=url_img(Modulo,RispDate1(i))%>"><BR> </td>
		    <td valign=top> <img src="<%=url_img(Modulo,RispEsatte1(i))%>"><BR></td>				
				  
		    <% else %>      <!-- se la risposta è correta usa il colore verde -->
	
		    <td valign=top><b><font color="green"><%=i%></b></font></td>
		    <td valign=top><b><font color="green"><%=rsTabella.Fields("CodiceDomanda")%></b></font></td>
		    <td valign=top><b><font color="green"><%=rsTabella.Fields("Quesito")%></b></font></td>
		     <td valign=top> <font color="green"><img src="<%=url_img(Modulo,RispDate1(i))%>"><BR> </td>
		    <td valign=top> <font color="green"><img src="<%=url_img(Modulo,RispEsatte1(i))%>"><BR></td>				
				  			
				  
			<%End if%>
		  </tr>
		  <% rsTabella.movenext
				i=i+1
				Loop %>	
		</table> 
			</p>
	  </div>
			<div id=piede_pagina>
			<p><p>
		
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h3> 	


			 
			</div>
			</div>
			
	</body>
	</html>
	