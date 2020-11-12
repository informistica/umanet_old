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
function showText() {window.alert("Non puoi spostare i dati degli altri studenti!")

location.href="studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>"
//location.href=window.history.back();
 }
 </script>
</head>





   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Sposta=Request.QueryString("Sposta") ' serve quando richiamo la pagina da se stessa
   ParagrafoNew=Request.QueryString("ParagrafoNew")
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
     
   
    
   
ID=Request.QueryString("CodiceDomanda")    
Cartella=Request.QueryString("Cartella")
Paragrafo=Request.QueryString("Paragrafo")
Titolo=Request.QueryString("Titolo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
CodiceAllievo=request.querystring("cod")
   
if sposta<>"" then ' else alla fine faccio sciegliere il nuovo paragrafo 
  
	url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	url=Replace(url,"\","/")
	urlNew=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&  "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Titolo&"_"&ID&".txt" 
	urlNew=Replace(urlNew,"\","/")
	 if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
	<body>
		<div id="container">
		<div class="risultati_test" >
		<font color=#FF0000 size="4">
	
	<%
		 
			QuerySQL ="UPDATE Domande SET Domande.Id_Arg ='"&ParagrafoNew&"' WHERE Domande.CodiceDomanda=" &ID&";"
			ConnessioneDB.Execute(QuerySQL)
	
	
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA
	
	Dim objFSO,objCreatedFile,OggFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	 
	'Create the FSO.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
    set OggFile = objFSO.GetFile (url)
    OggFile.Copy urlNew,true
     
	'CANCELLA LA VECCHIA VERSIONE DEL FILE11
	'response.write(url)
	objFSO.DeleteFile url
	response.write("url="&url)
	response.write("<br>urlnew="&urlNew)
	On Error Resume Next
	If Err.Number = 0 Then
		 
			Response.Write "Spostamento avvenuto! "
	 
	
	Else
	Response.Write Err.Description 
	Err.Number = 0
	End If
	
	
	
	
	
	   %>
		</font>   
		 
		 
		  <h4><a href="../cClasse/studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&cod=<%=CodiceAllievo%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>">Continua ...</a></h4>
		<p>&nbsp;</p>
		<div id=piede_pagina>
				<p><p>
				
			<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
	<h3><a href="../../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h3> 
							</div>
	 <!-- se il login è corretto richima la pagina per inserire le domande del test -->
	
		<%else%>
		
	   <BODY onLoad="showText();">
		
		<%end if%>
 <%else 
 
   QuerySQL="Select * from CARTELLA_MODULO_PARAGRAFI where  ID_Mod ='" & Modulo &"';"
  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
  ' response.write(QuerySQL)%>
  <div align="center">
   <table id="zebra" width="auto" align="center"><tr><th>Scegli nuovo paragrafo</th></tr>
<%
   If rsTabella.BOF=True And rsTabella.EOF=True Then 
    response.write("Nessun paragrafo per modulo il " & Modulo)
   else
	   Do until rsTabella.EOF %>
		 <tr><td><a href="8_sposta_domanda.asp?Sposta=1&Titolo=<%=rsTabella("Titolo")%>&ParagrafoNew=<%=rsTabella("ID_Paragrafo")%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&CodiceDomanda=<%=ID%>"><%=rsTabella(3)%></a></td></tr>
	   
	   <%rsTabella.movenext()
		loop
   end if	
    
 
 
 
 end if%>       
     </table>
     </div>   
    </body>
    </html>
	