<!-- esegui_test_MODBC3.asp -->

<%@ Language=VBScript %>
 
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

<TITLE>MESCOLA TEST</TITLE>
</HEAD>
<BODY> 

 <%  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query
 Dim ConnessioneDB, rsTabella, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,stato,CodiceDomanda
 Dim R()
    
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
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
 
%>
<div id="container">


<div class="contenuti_test" >
<p align="center"><b><font face="Verdana" size="4" color=#FF0000>ESEGUI TEST </font></b></p>
<p align="center"><font size="4"><b><%Response.write (TitoloTest) %></b></font></p> <!-- stampa il titolo del test -->

<%  
if (Stato=0) then 

  if CodiceSottopar<>"" then
	        QuerySQL="SELECT Domande.*, Paragrafi.Titolo " &_
   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
   " WHERE Domande.Id_Arg='" & CodiceTest & "' and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' order by Domande.CodiceDomanda asc;"
   
	   else
 
        QuerySQL="SELECT Domande.*, Paragrafi.Titolo " &_
   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
   " WHERE Domande.Id_Arg='" & CodiceTest & "' order by Domande.CodiceDomanda asc;"
   
	end if


 'Definzione codice SQl della query per ricercare le domande del paragrafo 
 ' mi serve anche il titolo del paragrafo per ricostruire il nome del file che contine la domanda plus
 
   
    
else 
'Definzione codice SQl della query per ricercare le domande del modulo
 QuerySQL="SELECT Domande.*, Paragrafi.Titolo" &_
   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
   " WHERE Domande.Id_Mod='" & Modulo & "' order by Domande.CodiceDomanda asc;"
 
end if    
    
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
  
  <% randomize()
  redim R(6)
  Do until rsTabella.EOF   ' esegue un ciclo e ad ogni iterazione crea un quiz (con 4 valori possibili) avente per nome il numero contenuto nella variabile i 
         
		R(0)=0
		for i=1 to 4 ' carico nel vettore le 4 risposte
		   'response.Write(rsTabella.Fields(1+i))
		   R(i)=rsTabella.Fields(1+i)
		next 
		CodiceDomanda=rsTabella.Fields(0)
		RE=rsTabella.Fields(6) ' prelevo il numero della risposta esatta
		appo=R(RE) ' leggo la risposta esatta che andrà spostata
		do ' genero il numero casuale per sapere dove spostare la domanda esatta
		  rand1=left((rnd()*5),1)
		loop until (rand1>0) and (rand1<5) and R(rand1)<>"altro"
		'rand1=left((rand*5),1)
		appo1=R(rand1) ' prelevo il valore contenuto nella nuova posizione della risposta esatta
		R(rand1)=appo   ' aggiorno il campo che contiene la nuova risposta esatta
		    ' scrivo il valore prelevato 
		R(RE)=appo1
		RE=rand1   ' aggiorno il valore della risposta esatta con il nuovo valore generato
		for i=1 to 4 '
		   R(i) = Replace(R(i), Chr(34), "'")
		   R(i)=  Replace(R(i),"'",Chr(96))
        next
 QuerySQL ="UPDATE Domande SET Risposta1 = '" &R(1)& "', Risposta2='" &R(2)&"', Risposta3 ='" &R(3)& "',Risposta4 = '" &R(4)& "', RispostaEsatta = '" &RE& "'  WHERE Domande.CodiceDomanda= "&CodiceDomanda & ";"
' dim objFSO,objCreatedFile
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Dim sRead, sReadLine, sReadAll, objTextFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'url="C:\Inetpub\umanetroot\Anno_2009-2010\log1.txt"
'Set objCreatedFile = objFSO.CreateTextFile(url, True)
'objCreatedFile.WriteLine(QuerySQL)
'objCreatedFile.Close
		
	    ConnessioneDB.Execute(QuerySQL)  
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande 
     Loop %>
   <P>
       
<% End If %>
<% rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
   ConnessioneDB.Close : Set ConnessioneDB = Nothing %>
  
 <H4>Test mescolato!<h4>
  </div>
  </div>
</BODY>
</HTML>