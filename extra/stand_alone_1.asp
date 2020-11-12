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
   Dim Risultato_relativo 
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   MO=Request.QueryString("MO")
   Modulo=MO
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
   
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   Cartella=Request.QueryString("Cartella")
  
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Capitolo=Request.QueryString("Capitolo")
   ID=Request.QueryString("CodiceFrase")
   

Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
d=request.querystring("cla")
cod=request.querystring("cod")
   
  homesito="/anno_2010-2011_ITC/ECDL"     
url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Replace(url,"\","/")
 




     QuerySQL ="SELECT CodiceDomanda FROM DOMANDE WHERE In_Quiz = 2 and Id_Mod='3COM_4';"
     Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	 
        do while not rsTabella.eof  
			QuerySQL ="UPDATE Domande2 SET In_Quiz=0  WHERE CodiceDomanda =" &rsTabella.fields("CodiceDomanda")&";"
            ConnessioneDB.Execute(QuerySQL) 
	        response.write("<br>" & QuerySQL)
			rsTabella.movenext
		loop
        rsTabella.close()
		%>
	</font>   
	 
		
      <h4><a href="../cClasse/studente_domande.asp?cla=<%=d%>">Continua ...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Home Verifica... </a></h3> 
						</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	