<!-- esegui_test_MODBC3.asp -->

<%@ Language=VBScript %>
<%Function url_img(cartella,nome_img)
	 
	 url_img="../img_quiz" & "/" & cartella &"/" & nome_img&".jpg"
 	 'url_img=replace(url_img,"/","\")
End Function %>
<% Response.Buffer=True %>

<HTML>
<HEAD>
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

<TITLE>ESEGUI TEST</TITLE>
</HEAD>
<BODY> 

 <%  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query
 Dim ConnessioneDB, rsTabella, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,stato

    StringaConnessione= Request.Cookies("Dati")("StrConn")   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
   Stato=Request.QueryString("Stato") 
   Modulo=Request.QueryString("Modulo") 
   'Raccolta dei dati digitati dall'utente e salvati nel cookie
   TitoloTest=Request.Cookies("Dati")("TitoloTest")
'   CodiceTest = Request.Cookies("Dati")("CodiceTest")
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   CodiceTest = Request.QueryString("CodiceTest") 

%>
<div id="container">


<div class="contenuti_test" >
<p align="center"><b><font face="Verdana" size="4" color=#FF0000>ESEGUI TEST </font></b></p>
<p align="center"><font size="4"><b><%Response.write (TitoloTest) %></b></font></p> <!-- stampa il titolo del test -->
<%  
 
 if (Stato=0) then 
 'Definzione codice SQl della query per ricercare le domande del paragrafo 
 ' mi serve anche il titolo del paragrafo per ricostruire il nome del file che contine la domanda plus
  QuerySQL="SELECT Domande1.CodiceDomanda,Domande1.Quesito, Domande1.Risposta1, Domande1.Risposta2, Domande1.Risposta3, Domande1.Risposta4, Domande1.RispostaEsatta, Moduli.ID_Mod " &_
" FROM Moduli INNER JOIN Domande1 ON Moduli.ID_Mod = Domande1.Id_Mod " &_
   " WHERE Domande1.Id_Arg='" & CodiceTest & "' order by Domande1.CodiceDomanda asc;"
   
 else
 
 
'Definzione codice SQl della query per ricercare le domande del modulo
 QuerySQL="SELECT Domande1.CodiceDomanda,Domande1.Quesito, Domande1.Risposta1, Domande1.Risposta2, Domande1.Risposta3, Domande1.Risposta4, Domande1.RispostaEsatta, Moduli.ID_Mod " &_
" FROM Moduli INNER JOIN Domande1 ON Moduli.ID_Mod = Domande1.Id_Mod " &_
   " WHERE Domande1.Id_Mod='" & Modulo & "'order by Domande1.CodiceDomanda asc  ;"
 
 end if
 
  		'dim objFSO,objCreatedFile
		'		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'		Dim sRead, sReadLine, sReadAll, objTextFile
		'		Set objFSO = CreateObject("Scripting.FileSystemObject")
		'		url="C:\Inetpub\umanetroot\anno_2009-2010\ECDL\database\log3.txt"
		'		Set objCreatedFile = objFSO.CreateTextFile(url, True)
		'		objCreatedFile.WriteLine(QuerySQL)
		'		objCreatedFile.Close 
	
	
	
	
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 
'Creazione di una pagina HTML dinamica con i test. 
'Le domande sono individuate da un nome del tipo NAME=i, dove i e' il numero
'della domanda. Il test e' indipendente dal numero di domande memorizzato.
'Dopo la compilazione del test, la pagina richiama calcola_risultato.asp
'che effettua il calcolo del risultato raggiunto.      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then %>
  <H4>Test non ancora disponibile!<h4>
  <p><h5><a href="javascript:history.back()"onMouseOver="window.status='Indietro';return true;" onMouseOut="window.status=''">Indietro</a>
</H5>

<% Else %>
<FORM METHOD="POST" ACTION="calcola_risultato_img.asp?Stato=<%=Stato%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>">
  <%i=1 'inizializza la variabile i (contatore delle domande)
  ' utilizzo la funzione url_img per creare il percorso dell'immagine partendo dal modulo e considerando che nella cartella img_quiz porrò le cartelle per le immagini dei quiz
  Do until rsTabella.EOF  %>
      <FIELDSET><LEGEND><B> <%=i & ") "%><%=rsTabella.Fields("Quesito")%></B></LEGEND>
      <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="1">
      <img src="<%=url_img(Modulo,rsTabella.Fields("Risposta1"))%>" style="border: 1px dotted #4F6A98;"><BR>
      <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="2">
      <img src="<%=url_img(Modulo,rsTabella.Fields("Risposta2"))%>" style="border: 1px dotted #4F6A98;"><BR>
      <INPUT TYPE="RADIO" NAME="<%=i%>"  VALUE="3">
      <img src="<%=url_img(Modulo,rsTabella.Fields("Risposta3"))%>" style="border: 1px dotted #4F6A98;"><BR>
      <INPUT TYPE="RADIO" NAME="<%=i%>"  VALUE="4">
       <%response.write(rsTabella.Fields("Risposta4"))%>  <BR>
    </FIELDSET>
	 <%  
       i = i+ 1 
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande%>
   <% Loop %>
   <P>
      <INPUT TYPE="SUBMIT" NAME="submit" VALUE="Invia le risposte del test"> <!-- crea il bottone per inviare le riposte alla pagina che calcola il risultato -->
   </P>
   </FORM>
<% End If %>
<% rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
   ConnessioneDB.Close : Set ConnessioneDB = Nothing %>
  

  </div>
  </div>
</BODY>
</HTML>